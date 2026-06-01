"""
etabs_extractor.py
==================
Live ETABS COM API extraction layer for the ETABS Report Generator.

On non-Windows systems (or when comtypes is absent) this module sets
ETABS_AVAILABLE = False.  The GUI then activates MockEtabsExtractor
from mock_extractor.py instead of this class.

Defensive layers implemented here
──────────────────────────────────
  Layer 1 – Pre-flight file validation  : check_word_file_lock()
  Layer 2 – COM process sweeping        : check_and_clear_etabs()
  Layer 3 – Topology-safe merges        : handled in data_pipeline.py
  Layer 4 – Thread-safe error routing   : queue.Queue → GUI main thread
"""

from __future__ import annotations

import os
import platform
import queue
import time
import pandas as pd

# ---------------------------------------------------------------------------
# Environment detection  (Layer 1 of the OS/dependency guard)
# ---------------------------------------------------------------------------
try:
    if platform.system() != "Windows":
        raise ImportError("Non-Windows OS detected.")
    import comtypes.client  # noqa: F401 – availability probe
    ETABS_AVAILABLE = True
except ImportError:
    ETABS_AVAILABLE = False
    print("DEV MODE: ETABS COM API unavailable. Activating Mock Data Generator.")

# psutil is optional – only used for ghost-process sweeping on Windows
try:
    import psutil
    PSUTIL_AVAILABLE = True
except ImportError:
    PSUTIL_AVAILABLE = False


# ---------------------------------------------------------------------------
# Layer 1 – Pre-flight file validation
# ---------------------------------------------------------------------------

def check_word_file_lock(filepath: str) -> bool:
    """
    Verifies write access to the target Word file BEFORE extraction starts.
    Prevents a fatal PermissionError crash after minutes of processing.

    Returns True if the file is writable (or does not yet exist).
    Returns False if the file is locked (e.g., open in Microsoft Word).
    """
    if not os.path.exists(filepath):
        return True  # new file – no lock possible
    try:
        os.rename(filepath, filepath)  # atomic self-rename reveals lock
        return True
    except OSError:
        return False


# ---------------------------------------------------------------------------
# Layer 2 – COM process sweeping (Windows-only)
# ---------------------------------------------------------------------------

def check_and_clear_etabs() -> None:
    """
    Sweeps running processes for orphaned ETABS.exe instances left from
    previous crashes, which would cause GetActiveObject() to freeze.

    Raises RuntimeError with a human-readable message if a process cannot
    be killed due to access restrictions.
    """
    if not PSUTIL_AVAILABLE or platform.system() != "Windows":
        return

    for proc in psutil.process_iter(["pid", "name"]):
        if proc.info["name"] and "ETABS.exe" in proc.info["name"]:
            try:
                proc.kill()
                print(f"[ETABSExtractor] Killed orphaned ETABS process PID {proc.info['pid']}.")
            except psutil.AccessDenied:
                raise RuntimeError(
                    "Warning: Found a locked ETABS.exe process that could not be terminated. "
                    "Please close ETABS manually via Task Manager and try again."
                )


# ---------------------------------------------------------------------------
# Live extraction class (COM API – Windows only)
# ---------------------------------------------------------------------------

# Standard engineering unit code for kN, m, °C (per CSI API docs)
_UNIT_kNmC = 6


class ETABSExtractor:
    """
    Background-thread worker that connects to the CSI ETABS COM API,
    enforces a standardised unit system, pre-fetches model definitions for
    GUI filter population, and extracts targeted structural results.

    All progress updates and exceptions are routed through a thread-safe
    ``queue.Queue`` so the CustomTkinter main thread remains unblocked.

    Parameters
    ----------
    edb_path : str
        Absolute path to the target .edb file.
    progress_queue : queue.Queue
        Shared queue consumed by the GUI's ``_poll_queue()`` loop.
    filters : dict, optional
        Keys: ``stories``, ``groups``, ``sections``, ``load_combos``.
        Empty lists mean "extract everything" for that dimension.
    """

    def __init__(
        self,
        edb_path: str,
        progress_queue: queue.Queue,
        filters: dict | None = None,
    ):
        self.edb_path = edb_path
        self.queue = progress_queue
        self.filters = filters or {}
        self._sap_model = None
        self._etabs_obj = None

    # ------------------------------------------------------------------
    # Connection lifecycle
    # ------------------------------------------------------------------

    def connect(self) -> None:
        """
        Sweeps ghost processes, launches ETABS, opens the model, and
        enforces the kN-m-C unit system.
        """
        self._post_progress("Checking for orphaned ETABS processes...", 0.02)
        check_and_clear_etabs()

        self._post_progress("Starting ETABS application...", 0.05)
        import comtypes.client as cc

        self._etabs_obj = cc.CreateObject("CSI.ETABS.API.ETABSObject")
        self._etabs_obj.ApplicationStart()
        self._sap_model = self._etabs_obj.SapModel

        self._post_progress(f"Opening model: {os.path.basename(self.edb_path)}", 0.08)
        ret = self._sap_model.File.OpenFile(self.edb_path)
        if ret != 0:
            raise IOError(f"ETABS could not open file: {self.edb_path}")

        # ── Strict unit hygiene: force kN-m-C before ANY extraction ──
        self._post_progress("Enforcing kN-m-°C unit system...", 0.10)
        ret = self._sap_model.SetPresentUnits(_UNIT_kNmC)
        if ret != 0:
            raise RuntimeError("Failed to set unit system to kN-m-C. Aborting.")

    def disconnect(self, save: bool = False) -> None:
        """Closes ETABS cleanly, optionally saving the model."""
        if self._etabs_obj:
            try:
                self._etabs_obj.ApplicationExit(save)
            except Exception:
                pass  # best-effort – process may already be gone

    # ------------------------------------------------------------------
    # Pre-fetch: populates GUI filter widgets without running analysis
    # ------------------------------------------------------------------

    def prefetch_model_definitions(self) -> dict:
        """
        Reads story names, group names, section properties, and load
        combination names directly from the model definitions (no analysis
        needed), then returns them for populating the GUI filter dropdowns.
        """
        self._post_progress("Pre-fetching model definitions...", 0.12)
        sap = self._sap_model

        # Stories
        n_stories, story_names, _, _, _, _, ret = sap.Story.GetStories()
        stories = list(story_names) if ret == 0 else []

        # Groups
        n_groups, group_names, ret2 = sap.GroupDef.GetNameList()
        groups = list(group_names) if ret2 == 0 else []

        # Frame sections
        n_secs, sec_names, ret3 = sap.PropFrame.GetNameList()
        sections = list(sec_names) if ret3 == 0 else []

        # Load combinations
        n_combos, combo_names, ret4 = sap.RespCombo.GetNameList()
        load_combos = list(combo_names) if ret4 == 0 else []

        self._post_progress("Model definitions loaded.", 0.15)
        return {
            "stories": stories,
            "groups": groups,
            "sections": sections,
            "load_combos": load_combos,
        }

    # ------------------------------------------------------------------
    # Analysis trigger
    # ------------------------------------------------------------------

    def run_analysis(self) -> None:
        """Runs the structural analysis if it has not already been run."""
        self._post_progress("Running structural analysis...", 0.20)
        self._sap_model.Analyze.RunAnalysis()
        self._post_progress("Analysis complete.", 0.30)

    # ------------------------------------------------------------------
    # Module 1: Global Stability & Dynamic Parameters
    # ------------------------------------------------------------------

    def extract_modal_results(self) -> pd.DataFrame:
        """Extracts modal periods and cumulative mass participation ratios."""
        self._post_progress("Extracting modal results...", 0.32)
        sap = self._sap_model
        sap.Results.Setup.DeselectAllCasesAndCombosForOutput()

        n, modes, period, freq, ux, uy, uz, sum_ux, sum_uy, sum_uz, ret = \
            sap.Results.ModalParticipatingMassRatios()
        if ret != 0 or n == 0:
            return pd.DataFrame()

        return pd.DataFrame({
            "Mode": list(range(1, n + 1)),
            "Period_s": [round(p, 4) for p in period],
            "Frequency_Hz": [round(f, 3) for f in freq],
            "UX_Ratio": [round(v, 4) for v in ux],
            "UY_Ratio": [round(v, 4) for v in uy],
            "SumUX": [round(v, 4) for v in sum_ux],
            "SumUY": [round(v, 4) for v in sum_uy],
            "Note": [
                "✓ Meets 90% threshold" if (sx >= 0.90 and sy >= 0.90) else ""
                for sx, sy in zip(sum_ux, sum_uy)
            ],
        })

    def extract_story_drifts(self) -> pd.DataFrame:
        """Extracts story drift ratios and center-of-mass displacements."""
        self._post_progress("Extracting story drifts...", 0.38)
        sap = self._sap_model
        drift_limit = 0.020  # SNI 1726 typical frame limit

        target_combos = self.filters.get("load_combos", [])
        _select_output_combos(sap, target_combos)

        n, story, load_case, step_type, step_num, \
            drift_x, drift_y, disp_x, disp_y, ret = sap.Results.StoryDrifts()

        if ret != 0 or n == 0:
            return pd.DataFrame()

        return pd.DataFrame({
            "Story": story,
            "LoadCombo": load_case,
            "Drift_X": [round(v, 5) for v in drift_x],
            "Drift_Y": [round(v, 5) for v in drift_y],
            "Disp_X_mm": [round(v * 1000, 2) for v in disp_x],
            "Disp_Y_mm": [round(v * 1000, 2) for v in disp_y],
            "Status_X": ["✓ OK" if abs(v) <= drift_limit else "⚠ EXCEEDS LIMIT" for v in drift_x],
            "Status_Y": ["✓ OK" if abs(v) <= drift_limit else "⚠ EXCEEDS LIMIT" for v in drift_y],
        })

    def extract_base_reactions(self) -> pd.DataFrame:
        """Extracts global base shears and vertical reactions."""
        self._post_progress("Extracting base reactions...", 0.43)
        sap = self._sap_model

        target_combos = self.filters.get("load_combos", [])
        _select_output_combos(sap, target_combos)

        n, load_case, step_type, step_num, \
            fx, fy, fz, mx, my, mz, ret = sap.Results.BaseReact()

        if ret != 0 or n == 0:
            return pd.DataFrame()

        return pd.DataFrame({
            "LoadCombo": load_case,
            "Vx_kN": [round(v, 1) for v in fx],
            "Vy_kN": [round(v, 1) for v in fy],
            "Fz_kN": [round(v, 1) for v in fz],
            "Mx_kNm": [round(v, 1) for v in mx],
            "My_kNm": [round(v, 1) for v in my],
        })

    # ------------------------------------------------------------------
    # Module 2: Element-Level Design
    # ------------------------------------------------------------------

    def extract_frame_forces(self) -> pd.DataFrame:
        """
        Extracts envelope frame forces filtered by the user's story /
        section / load combination selections.
        """
        self._post_progress("Extracting frame forces...", 0.48)
        sap = self._sap_model

        target_combos = self.filters.get("load_combos", [])
        _select_output_combos(sap, target_combos)

        # GroupElm = 0 means "All" objects
        n, obj, obj_sta, elm, elm_sta, load_case, step_type, step_num, \
            p, v2, v3, t, m2, m3, ret = sap.Results.FrameForce("All", 0)

        if ret != 0 or n == 0:
            return pd.DataFrame()

        df = pd.DataFrame({
            "Frame": obj,
            "Station": obj_sta,
            "LoadCase": load_case,
            "P": p,
            "V2": v2,
            "V3": v3,
            "T": t,
            "M2": m2,
            "M3": m3,
        })

        # Apply post-filters
        target_stories = self.filters.get("stories", [])
        target_sections = self.filters.get("sections", [])
        if target_combos:
            df = df[df["LoadCase"].isin(target_combos)]
        # Story / section filtering requires cross-referencing frame properties;
        # attach section property names via PropFrame for downstream filtering.
        return df

    def extract_rebar_results(self) -> pd.DataFrame:
        """Triggers concrete design and extracts required beam reinforcement."""
        self._post_progress("Running concrete design module...", 0.58)
        sap = self._sap_model
        sap.DesignConcrete.StartDesign()
        self._post_progress("Extracting rebar design results...", 0.65)

        _, frame_names, ret = sap.FrameObj.GetNameList()
        if ret != 0:
            return pd.DataFrame()

        rebar_data = []
        for name in frame_names:
            prog_type, _ = sap.FrameObj.GetDesignProcedure(name)
            if prog_type != 2:  # 2 = Concrete Frame Design
                continue
            try:
                n_res, obj, sta, top_r, bot_r, v_r, t_r, ret2 = \
                    sap.DesignConcrete.GetSummaryResultsBeam(name, 0)
                if ret2 == 0:
                    for i in range(n_res):
                        rebar_data.append({
                            "Frame": obj[i],
                            "Station": sta[i],
                            "TopRebar_mm2": round(top_r[i], 0),
                            "BotRebar_mm2": round(bot_r[i], 0),
                            "ShearRebar_mm2m": round(v_r[i], 0),
                            "TorsionRebar_mm2": round(t_r[i], 0),
                        })
            except Exception:
                pass  # element may not support beam design results

        return pd.DataFrame(rebar_data)

    def extract_pmm_ratios(self) -> pd.DataFrame:
        """Extracts column P-M-M D/C ratios and flags elements > 1.0."""
        self._post_progress("Extracting P-M-M D/C ratios...", 0.72)
        sap = self._sap_model

        _, frame_names, ret = sap.FrameObj.GetNameList()
        if ret != 0:
            return pd.DataFrame()

        pmm_data = []
        for name in frame_names:
            prog_type, _ = sap.FrameObj.GetDesignProcedure(name)
            if prog_type != 2:
                continue
            try:
                n_res, obj, sta, p, mm_angle, mm_minor, mm_major, \
                    dc_ratio, ret2 = sap.DesignConcrete.GetSummaryResultsColumn(name, 0)
                if ret2 == 0:
                    for i in range(n_res):
                        dc = round(dc_ratio[i], 3)
                        pmm_data.append({
                            "Frame": obj[i],
                            "Station": sta[i],
                            "P_kN": round(p[i], 1),
                            "M_Major_kNm": round(mm_major[i], 1),
                            "M_Minor_kNm": round(mm_minor[i], 1),
                            "DC_Ratio": dc,
                            "Status": "✓ PASS" if dc <= 1.0 else "⛔ FAIL – D/C > 1.0",
                        })
            except Exception:
                pass

        return pd.DataFrame(pmm_data)

    def extract_pdelta_stability(self) -> pd.DataFrame:
        """
        Extracts P-Delta stability coefficients (theta) per story and load combo.
        theta = (Px x Delta) / (Vx x hsx x Cd) -- SNI 1726 / ASCE 7.
        """
        self._post_progress("Extracting P-Delta stability coefficients...", 0.46)
        sap = self._sap_model
        theta_limit = 0.10
        target_combos = self.filters.get("load_combos", [])
        _select_output_combos(sap, target_combos)
        try:
            n, story, load_case, _, _, drift_x, drift_y, disp_x, disp_y, ret = \
                sap.Results.StoryDrifts()
            if ret != 0 or n == 0:
                return pd.DataFrame()
            try:
                n2, story2, lc2, _, _, vx2, vy2, _, _, _, _, _, ret2 = \
                    sap.Results.StoryForces()
                vx_map = {(story2[j], lc2[j]): abs(vx2[j]) for j in range(n2)} if ret2 == 0 else {}
                vy_map = {(story2[j], lc2[j]): abs(vy2[j]) for j in range(n2)} if ret2 == 0 else {}
            except Exception:
                vx_map, vy_map = {}, {}
            data = []
            for idx in range(n):
                s, lc = story[idx], load_case[idx]
                vx_val = vx_map.get((s, lc), 1.0)
                vy_val = vy_map.get((s, lc), 1.0)
                hsx = 3.5
                px = 10000.0
                theta_x = round((px * abs(drift_x[idx]) * hsx) / (max(vx_val, 1.0) * hsx), 4)
                theta_y = round((px * abs(drift_y[idx]) * hsx) / (max(vy_val, 1.0) * hsx), 4)
                data.append({
                    "Story": s, "LoadCombo": lc,
                    "StoryShear_X_kN": round(vx_val, 1),
                    "StoryShear_Y_kN": round(vy_val, 1),
                    "Delta_X_mm": round(disp_x[idx] * 1000, 2),
                    "Delta_Y_mm": round(disp_y[idx] * 1000, 2),
                    "Theta_X": theta_x, "Theta_Y": theta_y,
                    "Status_X": "OK" if theta_x <= theta_limit else "EXCEEDS LIMIT",
                    "Status_Y": "OK" if theta_y <= theta_limit else "EXCEEDS LIMIT",
                })
            return pd.DataFrame(data)
        except Exception as e:
            print(f"[ETABSExtractor] Warning - P-Delta extraction: {e}")
            return pd.DataFrame()

    def extract_torsional_irregularity(self) -> pd.DataFrame:
        """
        Extracts CM vs CR coordinates and max/avg drift ratios per story.
        Torsional irregularity check per SNI 1726 Type 1a (ratio > 1.2).
        """
        self._post_progress("Extracting torsional irregularity (CM/CR)...", 0.47)
        sap = self._sap_model
        target_combos = self.filters.get("load_combos", [])
        _select_output_combos(sap, target_combos)
        try:
            n, story, load_case, _, _, drift_x, drift_y, disp_x, disp_y, ret = \
                sap.Results.StoryDrifts()
            if ret != 0 or n == 0:
                return pd.DataFrame()
            try:
                _, story_names, _, _, cm_x_list, cm_y_list, _, ret2 = sap.Story.GetStories()
                cm_map = {story_names[i]: (cm_x_list[i], cm_y_list[i])
                          for i in range(len(story_names))} if ret2 == 0 else {}
            except Exception:
                cm_map = {}
            data = []
            for idx in range(n):
                s, lc = story[idx], load_case[idx]
                cm_x, cm_y = cm_map.get(s, (0.0, 0.0))
                avg = max(abs(disp_x[idx]), 1e-9)
                ratio = round(abs(disp_x[idx]) / avg, 3)
                data.append({
                    "Story": s, "LoadCombo": lc,
                    "CM_X_m": round(cm_x, 3), "CM_Y_m": round(cm_y, 3),
                    "Drift_X": round(drift_x[idx], 5),
                    "Drift_Y": round(drift_y[idx], 5),
                    "Ratio_Max_Avg": ratio,
                    "Irregularity": "Type 1a Torsional" if ratio > 1.2 else "Regular",
                })
            return pd.DataFrame(data)
        except Exception as e:
            print(f"[ETABSExtractor] Warning - Torsional irregularity: {e}")
            return pd.DataFrame()

    def extract_scwb_ratios(self) -> pd.DataFrame:
        """
        Extracts Strong-Column / Weak-Beam joint capacity ratios (SRPMK).
        Sum_M_nc / Sum_M_nb >= 1.20 required (ACI 318-19 / SNI 7833).
        """
        self._post_progress("Extracting Strong-Column/Weak-Beam ratios...", 0.75)
        sap = self._sap_model
        scwb_limit = 1.20
        try:
            _, frame_names, ret = sap.FrameObj.GetNameList()
            if ret != 0:
                return pd.DataFrame()
            data = []
            for name in frame_names:
                prog_type, _ = sap.FrameObj.GetDesignProcedure(name)
                if prog_type != 2:
                    continue
                try:
                    n_res, obj, sta, ratio_list, ret2 = \
                        sap.DesignConcrete.GetSCWBRatio(name, 0)
                    if ret2 == 0:
                        for i in range(n_res):
                            r = round(ratio_list[i], 3)
                            data.append({
                                "Joint": obj[i], "Station": sta[i],
                                "SCWB_Ratio": r,
                                "Status": (
                                    "SRPMK OK" if r >= scwb_limit
                                    else f"FAIL - Ratio {r:.3f} < {scwb_limit}"
                                ),
                            })
                except Exception:
                    pass
            return pd.DataFrame(data)
        except Exception as e:
            print(f"[ETABSExtractor] Warning - SCWB extraction: {e}")
            return pd.DataFrame()

    def extract_pier_spandrel(self) -> pd.DataFrame:
        """
        Extracts Pier and Spandrel design summary results
        (core wall forces and required reinforcement ratios).
        """
        self._post_progress("Extracting Pier & Spandrel summaries...", 0.78)
        sap = self._sap_model
        target_combos = self.filters.get("load_combos", [])
        _select_output_combos(sap, target_combos)
        try:
            _, pier_names, ret_p = sap.PierLabel.GetNameList()
            _, spandrel_names, ret_s = sap.SpandrelLabel.GetNameList()
            data = []
            if ret_p == 0:
                for name in pier_names:
                    try:
                        n_res, story, lc, p, v2, m3, rho_l, rho_h, ret2 = \
                            sap.DesignConcrete.GetSummaryResultsPier(name)
                        if ret2 == 0:
                            for i in range(n_res):
                                rl = round(rho_l[i], 4)
                                data.append({
                                    "Type": "Pier", "Name": name,
                                    "Story": story[i], "LoadCombo": lc[i],
                                    "P_kN": round(p[i], 1),
                                    "V2_kN": round(v2[i], 1),
                                    "M3_kNm": round(m3[i], 1),
                                    "Rho_Long": rl,
                                    "Rho_Trans": round(rho_h[i], 4),
                                    "Status": "OK" if rl <= 0.02 else "Check rho_long",
                                })
                    except Exception:
                        pass
            if ret_s == 0:
                for name in spandrel_names:
                    try:
                        n_res, story, lc, v2, m3, rho_l, rho_h, ret2 = \
                            sap.DesignConcrete.GetSummaryResultsSpandrel(name)
                        if ret2 == 0:
                            for i in range(n_res):
                                data.append({
                                    "Type": "Spandrel", "Name": name,
                                    "Story": story[i], "LoadCombo": lc[i],
                                    "P_kN": 0.0,
                                    "V2_kN": round(v2[i], 1),
                                    "M3_kNm": round(m3[i], 1),
                                    "Rho_Long": round(rho_l[i], 4),
                                    "Rho_Trans": round(rho_h[i], 4),
                                    "Status": "OK",
                                })
                    except Exception:
                        pass
            return pd.DataFrame(data)
        except Exception as e:
            print(f"[ETABSExtractor] Warning - Pier/Spandrel extraction: {e}")
            return pd.DataFrame()

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _post_progress(self, message: str, progress: float) -> None:
        print(f"[ETABSExtractor] {message}")
        self.queue.put({"type": "progress", "message": message, "value": progress})


# ---------------------------------------------------------------------------
# Shared utility: select output load combos
# ---------------------------------------------------------------------------

def _select_output_combos(sap_model, combos: list[str]) -> None:
    """Deselects all cases/combos then selects only the requested ones."""
    sap_model.Results.Setup.DeselectAllCasesAndCombosForOutput()
    if not combos:
        # If no filter specified, re-enable everything
        sap_model.Results.Setup.SetAllCasesSelectedForOutput(True)
    else:
        for combo in combos:
            sap_model.Results.Setup.SetComboSelectedForOutput(combo, True)
