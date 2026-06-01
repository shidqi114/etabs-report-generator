"""
mock_extractor.py
=================
Cross-platform mock data stub for the ETABS Report Generator.
Activated automatically when running on macOS/Linux or when the
ETABS COM API (comtypes) is not available.

Generates structurally plausible engineering data that mirrors the
exact column schema of the real ETABSExtractor, enabling full UI/UX
and Word template development without a live ETABS environment.
"""

import pandas as pd
import numpy as np
import random
import time
import queue

# ---------------------------------------------------------------------------
# Default model definitions (mirrors what a real .edb would define)
# ---------------------------------------------------------------------------
MOCK_STORIES = ["Story 1", "Story 2", "Story 3", "Story 4", "Story 5", "Roof"]
MOCK_GROUPS = ["Core_Walls", "Typical_Floor_Beams", "Perimeter_Columns", "Transfer_Beams"]
MOCK_SECTIONS = ["B-30x60", "B-40x80", "C-60x60", "C-80x80", "Wall-300", "Wall-200"]
MOCK_COMBOS = [
    "1.4D",
    "1.2D+1.6L",
    "1.2D+1.0Ex+0.3Ey",
    "1.2D+1.0Ey+0.3Ex",
    "0.9D+1.0Ex",
    "ENV-ALL",
]


class MockEtabsExtractor:
    """
    High-fidelity mock data generator that produces realistic structural
    engineering datasets compatible with data_pipeline.py and report_generator.py.

    Parameters
    ----------
    target_stories : list[str]
        Story names to include in generated data.
    load_combos : list[str]
        Load combination names to include in generated data.
    target_sections : list[str], optional
        Section names to filter generated frames.
    progress_queue : queue.Queue, optional
        Thread-safe queue for routing progress updates back to the GUI.
    """

    def __init__(
        self,
        target_stories=None,
        load_combos=None,
        target_sections=None,
        progress_queue: queue.Queue = None,
    ):
        self.stories = target_stories or MOCK_STORIES
        self.combos = load_combos or MOCK_COMBOS
        self.sections = target_sections or MOCK_SECTIONS[:2]  # default: only beams
        self.beam_sections = [s for s in self.sections if s.startswith("B-")]
        if not self.beam_sections:
            self.beam_sections = ["B-30x60", "B-40x80"]
        self.col_sections = [s for s in self.sections if s.startswith("C-")]
        if not self.col_sections:
            self.col_sections = ["C-60x60", "C-80x80"]
        self.queue = progress_queue

    # ------------------------------------------------------------------
    # Pre-fetch: mimics ETABSExtractor.prefetch_model_definitions()
    # ------------------------------------------------------------------
    def prefetch_model_definitions(self) -> dict:
        """Returns mock model definitions for populating GUI filter widgets."""
        self._post_progress("DEV MODE: Loading mock model definitions...", 0.05)
        time.sleep(0.6)  # simulate COM latency
        return {
            "stories": MOCK_STORIES,
            "groups": MOCK_GROUPS,
            "sections": MOCK_SECTIONS,
            "load_combos": MOCK_COMBOS,
        }

    # ------------------------------------------------------------------
    # Module 1: Global Stability & Dynamic Parameters
    # ------------------------------------------------------------------
    def extract_modal_results(self) -> pd.DataFrame:
        """Generates mock modal periods and mass participation ratios (> 90% in X/Y)."""
        self._post_progress("DEV MODE: Generating modal results...", 0.15)
        data = []
        cumulative_ux, cumulative_uy = 0.0, 0.0
        for mode in range(1, 13):
            ux_inc = random.uniform(0.05, 0.20) if mode <= 5 else random.uniform(0.01, 0.05)
            uy_inc = random.uniform(0.05, 0.20) if mode <= 5 else random.uniform(0.01, 0.05)
            cumulative_ux = min(cumulative_ux + ux_inc, 1.0)
            cumulative_uy = min(cumulative_uy + uy_inc, 1.0)
            period = round(2.5 / mode + random.uniform(-0.05, 0.05), 4)
            data.append(
                {
                    "Mode": mode,
                    "Period_s": max(period, 0.05),
                    "Frequency_Hz": round(1.0 / max(period, 0.05), 3),
                    "UX_Ratio": round(ux_inc, 4),
                    "UY_Ratio": round(uy_inc, 4),
                    "SumUX": round(cumulative_ux, 4),
                    "SumUY": round(cumulative_uy, 4),
                    "Note": "✓ Meets 90% threshold" if mode == 6 else "",
                }
            )
        return pd.DataFrame(data)

    def extract_story_drifts(self) -> pd.DataFrame:
        """Generates mock story drift ratios (SNI 1726 limit: 0.020 for typical frames)."""
        self._post_progress("DEV MODE: Generating story drift data...", 0.25)
        data = []
        drift_limit = 0.020
        for story in self.stories:
            for combo in self.combos:
                drift_x = round(random.uniform(0.001, 0.025), 5)
                drift_y = round(random.uniform(0.001, 0.025), 5)
                disp_x = round(random.uniform(5.0, 80.0), 2)
                disp_y = round(random.uniform(5.0, 80.0), 2)
                data.append(
                    {
                        "Story": story,
                        "LoadCombo": combo,
                        "Drift_X": drift_x,
                        "Drift_Y": drift_y,
                        "Disp_X_mm": disp_x,
                        "Disp_Y_mm": disp_y,
                        "Status_X": "✓ OK" if drift_x <= drift_limit else "⚠ EXCEEDS LIMIT",
                        "Status_Y": "✓ OK" if drift_y <= drift_limit else "⚠ EXCEEDS LIMIT",
                    }
                )
        return pd.DataFrame(data)

    def extract_base_reactions(self) -> pd.DataFrame:
        """Generates mock base shears (Vx, Vy) and axial loads (Fz)."""
        self._post_progress("DEV MODE: Generating base reactions...", 0.30)
        data = []
        for combo in self.combos:
            data.append(
                {
                    "LoadCombo": combo,
                    "Vx_kN": round(random.uniform(1500, 4500), 1),
                    "Vy_kN": round(random.uniform(1500, 4500), 1),
                    "Fz_kN": round(random.uniform(15000, 45000), 1),
                    "Mx_kNm": round(random.uniform(500, 3000), 1),
                    "My_kNm": round(random.uniform(500, 3000), 1),
                }
            )
        return pd.DataFrame(data)

    # ------------------------------------------------------------------
    # Module 2: Element-Level Optimization & Design
    # ------------------------------------------------------------------
    def extract_frame_forces(self, is_optimized_model: bool = False) -> pd.DataFrame:
        """
        Generates mock frame forces for all combinations and stories.
        Applies a reduction factor for the 'optimized' (Retrofit) model
        to produce meaningful variance in comparison tests.
        """
        self._post_progress(
            "DEV MODE: Generating frame forces"
            + (" (optimized model)..." if is_optimized_model else "..."),
            0.40,
        )
        time.sleep(0.4)
        variance_factor = 0.85 if is_optimized_model else 1.0
        data = []
        for combo in self.combos:
            for story in self.stories:
                for i in range(1, 6):
                    section = random.choice(self.beam_sections)
                    data.append(
                        {
                            "Story": story,
                            "Frame": f"B{i}-{story.replace(' ', '')}",
                            "SectionProperty": section,
                            "LoadCase": combo,
                            "Station": "0.0",
                            "P": round(random.uniform(-10, 10) * variance_factor, 2),
                            "V2": round(random.uniform(50, 150) * variance_factor, 2),
                            "V3": round(random.uniform(0, 5) * variance_factor, 2),
                            "T": round(random.uniform(0, 2) * variance_factor, 3),
                            "M2": round(random.uniform(0, 10) * variance_factor, 2),
                            "M3": round(random.uniform(100, 350) * variance_factor, 2),
                            "Status": "✓ OK" if random.random() > 0.1 else "⚠ FAIL",
                        }
                    )
        return pd.DataFrame(data)

    def extract_rebar_results(self) -> pd.DataFrame:
        """Generates mock required reinforcement areas for concrete beam design."""
        self._post_progress("DEV MODE: Generating rebar design results...", 0.50)
        data = []
        for story in self.stories:
            for i in range(1, 6):
                section = random.choice(self.beam_sections)
                for station in ["0.0", "L/2", "L"]:
                    top = round(random.uniform(300, 1500), 0)
                    bot = round(random.uniform(200, 800), 0)
                    data.append(
                        {
                            "Story": story,
                            "Frame": f"B{i}-{story.replace(' ', '')}",
                            "SectionProperty": section,
                            "Station": station,
                            "TopRebar_mm2": top,
                            "BotRebar_mm2": bot,
                            "ShearRebar_mm2m": round(random.uniform(100, 600), 0),
                            "TorsionRebar_mm2": round(random.uniform(0, 200), 0),
                        }
                    )
        return pd.DataFrame(data)

    def extract_pmm_ratios(self) -> pd.DataFrame:
        """Generates mock P-M-M Demand/Capacity ratios for columns. Flags D/C > 1.0."""
        self._post_progress("DEV MODE: Generating P-M-M D/C ratios...", 0.60)
        data = []
        for story in self.stories:
            for i in range(1, 5):
                section = random.choice(self.col_sections)
                dc_ratio = round(random.uniform(0.40, 1.20), 3)
                data.append(
                    {
                        "Story": story,
                        "Frame": f"C{i}-{story.replace(' ', '')}",
                        "SectionProperty": section,
                        "P_kN": round(random.uniform(-5000, -500), 1),
                        "M2_kNm": round(random.uniform(10, 200), 1),
                        "M3_kNm": round(random.uniform(10, 200), 1),
                        "DC_Ratio": dc_ratio,
                        "Status": "✓ PASS" if dc_ratio <= 1.0 else "⛔ FAIL – D/C > 1.0",
                    }
                )
        return pd.DataFrame(data)

    def extract_pdelta_stability(self) -> pd.DataFrame:
        """
        Generates mock P-Delta stability coefficients (θ) per story and combo.
        SNI 1726 limit: θ ≤ 0.10 (typical frame) or θ ≤ θ_max ≤ 0.25.
        """
        self._post_progress("DEV MODE: Generating P-Delta stability coefficients...", 0.33)
        data = []
        theta_limit = 0.10
        for story in self.stories:
            # Story weight decreases toward roof
            story_weight = round(random.uniform(3000, 12000) / (self.stories.index(story) + 1), 0)
            for combo in self.combos:
                vx = round(random.uniform(500, 3000), 1)
                vy = round(random.uniform(500, 3000), 1)
                hx = round(random.uniform(3.2, 4.5), 2)  # story height (m)
                delta_x = round(random.uniform(2.0, 15.0), 2)  # inelastic drift (mm)
                delta_y = round(random.uniform(2.0, 15.0), 2)
                # θ = (Px × Δ) / (Vx × hx × Cd)  — simplified
                theta_x = round((story_weight * delta_x / 1000) / (vx * hx), 4)
                theta_y = round((story_weight * delta_y / 1000) / (vy * hx), 4)
                data.append({
                    "Story": story,
                    "LoadCombo": combo,
                    "StoryWeight_kN": story_weight,
                    "StoryShear_X_kN": vx,
                    "StoryShear_Y_kN": vy,
                    "StoryHeight_m": hx,
                    "Delta_X_mm": delta_x,
                    "Delta_Y_mm": delta_y,
                    "Theta_X": theta_x,
                    "Theta_Y": theta_y,
                    "Status_X": "✓ OK" if theta_x <= theta_limit else "⚠ EXCEEDS θ LIMIT",
                    "Status_Y": "✓ OK" if theta_y <= theta_limit else "⚠ EXCEEDS θ LIMIT",
                })
        return pd.DataFrame(data)

    def extract_torsional_irregularity(self) -> pd.DataFrame:
        """
        Generates mock torsional irregularity check (CM vs CR) per story.
        Flags stories where max_drift / avg_drift > 1.2 (SNI 1726 Type 1a).
        """
        self._post_progress("DEV MODE: Generating torsional irregularity data...", 0.36)
        data = []
        for story in self.stories:
            for combo in self.combos:
                cm_x = round(random.uniform(8.0, 12.0), 3)
                cm_y = round(random.uniform(8.0, 12.0), 3)
                cr_x = round(cm_x + random.uniform(-1.5, 1.5), 3)
                cr_y = round(cm_y + random.uniform(-1.5, 1.5), 3)
                avg_drift = round(random.uniform(0.003, 0.018), 5)
                max_drift = round(avg_drift * random.uniform(0.90, 1.35), 5)
                ratio = round(max_drift / avg_drift, 3)
                data.append({
                    "Story": story,
                    "LoadCombo": combo,
                    "CM_X_m": cm_x,
                    "CM_Y_m": cm_y,
                    "CR_X_m": cr_x,
                    "CR_Y_m": cr_y,
                    "Ecc_X_m": round(abs(cm_x - cr_x), 3),
                    "Ecc_Y_m": round(abs(cm_y - cr_y), 3),
                    "Max_Drift": max_drift,
                    "Avg_Drift": avg_drift,
                    "Ratio_Max_Avg": ratio,
                    "Irregularity": "⚠ Type 1a Torsional" if ratio > 1.2 else "✓ Regular",
                })
        return pd.DataFrame(data)

    def extract_scwb_ratios(self) -> pd.DataFrame:
        """
        Generates mock Strong-Column / Weak-Beam joint capacity ratios (SRPMK).
        SNI 7833 / ACI 318 requires ΣM_nc / ΣM_nb ≥ 1.20.
        """
        self._post_progress("DEV MODE: Generating Strong-Column/Weak-Beam ratios...", 0.63)
        data = []
        scwb_limit = 1.20
        for story in self.stories:
            for joint_i in range(1, 5):
                m_nc = round(random.uniform(300, 900), 1)   # Sum column capacity (kNm)
                m_nb = round(random.uniform(200, 700), 1)   # Sum beam capacity (kNm)
                ratio = round(m_nc / m_nb, 3)
                data.append({
                    "Story": story,
                    "Joint": f"J{joint_i}-{story.replace(' ', '')}",
                    "Sum_M_nc_kNm": m_nc,
                    "Sum_M_nb_kNm": m_nb,
                    "SCWB_Ratio": ratio,
                    "Status": (
                        "✓ SRPMK OK" if ratio >= scwb_limit
                        else f"⛔ FAIL – Ratio {ratio:.3f} < {scwb_limit}"
                    ),
                })
        return pd.DataFrame(data)

    def extract_pier_spandrel(self) -> pd.DataFrame:
        """
        Generates mock Pier and Spandrel summary forces and required
        reinforcement ratios for core walls and coupling beams.
        """
        self._post_progress("DEV MODE: Generating Pier & Spandrel summaries...", 0.68)
        pier_names = ["PIER-CoreN", "PIER-CoreS", "PIER-CoreE", "PIER-CoreW"]
        spandrel_names = ["SPANDREL-L1", "SPANDREL-L2"]
        data = []
        for story in self.stories:
            for combo in self.combos:
                for pier in pier_names:
                    p = round(random.uniform(-8000, -1000), 1)
                    v2 = round(random.uniform(100, 800), 1)
                    m3 = round(random.uniform(200, 2000), 1)
                    rho_long = round(random.uniform(0.005, 0.025), 4)
                    rho_horiz = round(random.uniform(0.0025, 0.010), 4)
                    data.append({
                        "Type": "Pier",
                        "Name": pier,
                        "Story": story,
                        "LoadCombo": combo,
                        "P_kN": p,
                        "V2_kN": v2,
                        "M3_kNm": m3,
                        "Rho_Long": rho_long,
                        "Rho_Trans": rho_horiz,
                        "Status": "✓ OK" if rho_long <= 0.02 else "⚠ Check ρ_long",
                    })
                for spandrel in spandrel_names:
                    v2 = round(random.uniform(50, 400), 1)
                    m3 = round(random.uniform(100, 800), 1)
                    data.append({
                        "Type": "Spandrel",
                        "Name": spandrel,
                        "Story": story,
                        "LoadCombo": combo,
                        "P_kN": 0.0,
                        "V2_kN": v2,
                        "M3_kNm": m3,
                        "Rho_Long": round(random.uniform(0.004, 0.015), 4),
                        "Rho_Trans": round(random.uniform(0.0025, 0.008), 4),
                        "Status": "✓ OK",
                    })
        return pd.DataFrame(data)

    # ------------------------------------------------------------------
    # Topology mismatch test set (for safe_model_merge validation)
    # ------------------------------------------------------------------
    def generate_comparison_test_set(self) -> tuple[pd.DataFrame, pd.DataFrame]:
        """
        Intentionally breaks topology by removing 3 elements from Model B
        to exercise the outer-merge / topology-mismatch handling in data_pipeline.py.
        """
        df_model_a = self.extract_frame_forces(is_optimized_model=False)
        df_model_b = self.extract_frame_forces(is_optimized_model=True)
        # Simulate elements being deleted in the Retrofit model
        df_model_b = df_model_b.iloc[3:].reset_index(drop=True)
        return df_model_a, df_model_b

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------
    def _post_progress(self, message: str, progress: float):
        """Routes progress updates to the GUI queue (if connected) or stdout."""
        print(message)
        if self.queue:
            self.queue.put({"type": "progress", "message": message, "value": progress})

