"""
app.py
======
CustomTkinter GUI for the ETABS Report Generator.

Architecture
────────────
  • All long-running tasks (prefetch, extraction, report) run on daemon
    background threads to keep the GUI frame rate stable.
  • Thread → GUI communication uses a queue.Queue polled by _poll_queue()
    via self.after(), following the thread-safe Tkinter pattern.
  • On non-Windows / no-comtypes environments ETABS_AVAILABLE is False and
    MockEtabsExtractor is used transparently.

Palette (from PROMPT_GUIDE.md)
───────────────────────────────
  BG_MAIN  = #0F172A   (Slate 900)
  BG_CARD  = #1E293B   (Slate 800)
  BORDER   = #334155   (Slate 700)
  BLUE     = #3B82F6   (hover #2563EB)
  GREEN    = #10B981   (hover #059669)
  AMBER    = #F59E0B
  TEXT     = #F8FAFC   (Slate 50)
  MUTED    = #94A3B8   (Slate 400)
"""

from __future__ import annotations

import os
os.environ["TK_SILENCE_DEPRECATION"] = "1"
import queue
import threading
import time
from tkinter import filedialog, messagebox

import customtkinter as ctk

import data_pipeline
import report_generator
from etabs_extractor import ETABS_AVAILABLE, check_word_file_lock

if ETABS_AVAILABLE:
    from etabs_extractor import ETABSExtractor
else:
    from mock_extractor import MockEtabsExtractor  # type: ignore[import]

# ── Theme ─────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

BG_MAIN = "#0F172A"
BG_CARD = "#1E293B"
BORDER  = "#334155"
BLUE    = "#3B82F6"
BLUE_H  = "#2563EB"
GREEN   = "#10B981"
AMBER   = "#F59E0B"
TEXT    = "#F8FAFC"
MUTED   = "#94A3B8"
RED     = "#EF4444"

FONT_TITLE  = ("Helvetica Neue", 16, "bold")
FONT_HDR    = ("Helvetica Neue", 13, "bold")
FONT_BODY   = ("Helvetica Neue", 12)
FONT_MUTED  = ("Helvetica Neue", 11)

# ── Helper: labelled card frame ────────────────────────────────────────────

def _card(parent, title: str = "", **kwargs) -> ctk.CTkFrame:
    frame = ctk.CTkFrame(
        parent,
        fg_color=BG_CARD,
        border_width=1,
        border_color=BORDER,
        corner_radius=10,
        **kwargs,
    )
    if title:
        lbl = ctk.CTkLabel(frame, text=title, font=FONT_HDR, text_color=TEXT)
        lbl.grid(row=0, column=0, columnspan=4, padx=15, pady=(14, 4), sticky="w")
    return frame


# ═══════════════════════════════════════════════════════════════════════════
# Main Application Window
# ═══════════════════════════════════════════════════════════════════════════

class ETABSReportApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("ETABS Report Generator")
        self.geometry("860x820")
        self.minsize(760, 680)
        self.configure(fg_color=BG_MAIN)

        # Shared state
        self._q: queue.Queue = queue.Queue()
        self._poll_id = None
        self._model_defs: dict = {}          # populated by prefetch
        self._filter_vars: dict[str, list] = {
            "stories": [],
            "groups": [],
            "sections": [],
            "load_combos": [],
        }

        # ── Layout ──────────────────────────────────────────────────────
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=0)  # header banner
        self.grid_rowconfigure(1, weight=0)  # info card
        self.grid_rowconfigure(2, weight=0)  # file card
        self.grid_rowconfigure(3, weight=1)  # filter panel (expandable)
        self.grid_rowconfigure(4, weight=0)  # module checklist
        self.grid_rowconfigure(5, weight=0)  # generate + status

        self._build_header()
        self._build_info_card()
        self._build_file_card()
        self._build_filter_panel()
        self._build_module_checklist()
        self._build_action_bar()

        # Dev-mode banner
        if not ETABS_AVAILABLE:
            self._post_status(
                "⚡ DEV MODE — ETABS unavailable. Mock Data Generator active.", AMBER
            )

        # Force window rendering update safely after mapping to prevent macOS blank/black screen bug
        self._update_id = self.after(100, self._safe_update)

    # ──────────────────────────────────────────────────────────────────
    # UI Builders
    # ──────────────────────────────────────────────────────────────────

    def _build_header(self):
        header = ctk.CTkFrame(self, fg_color="#0B1120", corner_radius=0)
        header.grid(row=0, column=0, sticky="ew", padx=0, pady=0)
        header.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(
            header,
            text="⚙  ETABS Report Generator",
            font=("Helvetica Neue", 20, "bold"),
            text_color=TEXT,
        )
        title.grid(row=0, column=0, padx=24, pady=(16, 4), sticky="w")

        sub = ctk.CTkLabel(
            header,
            text="Automated structural verification report builder  ·  CSI ETABS COM API",
            font=FONT_MUTED,
            text_color=MUTED,
        )
        sub.grid(row=1, column=0, padx=24, pady=(0, 14), sticky="w")

        mode_badge = ctk.CTkLabel(
            header,
            text="LIVE" if ETABS_AVAILABLE else "DEV MODE",
            font=("Helvetica Neue", 11, "bold"),
            text_color="#0F172A",
            fg_color=GREEN if ETABS_AVAILABLE else AMBER,
            corner_radius=6,
            width=80,
            height=24,
        )
        mode_badge.grid(row=0, column=1, rowspan=2, padx=20, pady=10, sticky="e")

    def _build_info_card(self):
        card = _card(self, title="📋  Project Information")
        card.grid(row=1, column=0, padx=20, pady=(16, 6), sticky="ew")
        card.grid_columnconfigure(1, weight=1)
        card.grid_columnconfigure(3, weight=1)

        _lbl(card, "Project Name").grid(row=1, column=0, padx=(15, 8), pady=8, sticky="w")
        self.proj_entry = _entry(card)
        self.proj_entry.grid(row=1, column=1, padx=8, pady=8, sticky="ew")

        _lbl(card, "Engineer Name").grid(row=1, column=2, padx=(16, 8), pady=8, sticky="w")
        self.eng_entry = _entry(card)
        self.eng_entry.grid(row=1, column=3, padx=(8, 15), pady=8, sticky="ew")

    def _build_file_card(self):
        card = _card(self, title="📁  Model & Template Files")
        card.grid(row=2, column=0, padx=20, pady=6, sticky="ew")
        card.grid_columnconfigure(1, weight=1)
        card.grid_columnconfigure(3, weight=1)

        # Primary EDB
        self.edb_path = ctk.StringVar()
        _lbl(card, "ETABS Model A (.edb)").grid(row=1, column=0, padx=(15, 8), pady=8, sticky="w")
        self.edb_entry = _entry(card, self.edb_path)
        self.edb_entry.grid(row=1, column=1, padx=8, pady=8, sticky="ew")
        _btn(card, "Browse", command=self._browse_edb).grid(row=1, column=2, padx=8, pady=8)

        # Secondary EDB (comparison)
        self.edb_b_path = ctk.StringVar()
        _lbl(card, "ETABS Model B (.edb)\n(comparison, optional)", size=11).grid(
            row=2, column=0, padx=(15, 8), pady=8, sticky="w"
        )
        self.edb_b_entry = _entry(card, self.edb_b_path, placeholder="Leave empty for single-model report")
        self.edb_b_entry.grid(row=2, column=1, padx=8, pady=8, sticky="ew")
        _btn(card, "Browse", command=self._browse_edb_b).grid(row=2, column=2, padx=8, pady=8)

        # Word template
        self.tpl_path = ctk.StringVar()
        _lbl(card, "Word Template (.docx)").grid(row=3, column=0, padx=(15, 8), pady=8, sticky="w")
        self.tpl_entry = _entry(card, self.tpl_path)
        self.tpl_entry.grid(row=3, column=1, padx=8, pady=8, sticky="ew")
        _btn(card, "Browse", command=self._browse_template).grid(row=3, column=2, padx=8, pady=8)

        # Output path display
        self.output_path = ctk.StringVar(value="Output: auto-generated alongside template")
        _lbl(card, "Output File", color=MUTED, size=11).grid(
            row=4, column=0, padx=(15, 8), pady=(4, 12), sticky="w"
        )
        ctk.CTkLabel(
            card, textvariable=self.output_path,
            font=FONT_MUTED, text_color=GREEN, anchor="w"
        ).grid(row=4, column=1, columnspan=2, padx=8, pady=(4, 12), sticky="ew")

        self.tpl_path.trace_add("write", self._update_output_label)

    def _build_filter_panel(self):
        outer = _card(self, title="🎯  Precision Extraction Filters")
        outer.grid(row=3, column=0, padx=20, pady=6, sticky="nsew")
        outer.grid_columnconfigure((0, 1, 2, 3), weight=1)
        outer.grid_rowconfigure(3, weight=1)

        hint = ctk.CTkLabel(
            outer,
            text="Select a Model A file to auto-populate filters from the .edb definitions.",
            font=FONT_MUTED,
            text_color=MUTED,
        )
        hint.grid(row=1, column=0, columnspan=4, padx=15, pady=(0, 8), sticky="w")
        self._filter_hint = hint

        # Four filter scrollable frames
        self._story_checks: dict[str, ctk.CTkCheckBox] = {}
        self._combo_checks: dict[str, ctk.CTkCheckBox] = {}
        self._section_checks: dict[str, ctk.CTkCheckBox] = {}
        self._group_checks: dict[str, ctk.CTkCheckBox] = {}

        self._story_scroll  = self._build_filter_column(outer, "Stories",        col=0)
        self._combo_scroll  = self._build_filter_column(outer, "Load Combos",    col=1)
        self._section_scroll= self._build_filter_column(outer, "Sections",       col=2)
        self._group_scroll  = self._build_filter_column(outer, "Groups",         col=3)

    def _build_filter_column(self, parent, label: str, col: int) -> ctk.CTkScrollableFrame:
        hdr = ctk.CTkLabel(parent, text=label, font=FONT_BODY, text_color=MUTED)
        hdr.grid(row=2, column=col, padx=(12, 4), pady=(0, 2), sticky="w")

        scroll = ctk.CTkScrollableFrame(
            parent,
            fg_color="#141E30",
            border_width=1,
            border_color=BORDER,
            corner_radius=8,
            height=140,
        )
        scroll.grid(row=3, column=col, padx=(12, 8), pady=(0, 14), sticky="nsew")
        return scroll

    def _build_module_checklist(self):
        card = _card(self, title="📦  Extraction Modules")
        card.grid(row=4, column=0, padx=20, pady=6, sticky="ew")
        card.grid_columnconfigure((0, 1, 2), weight=1)

        # ── Module 1: Global Stability & Dynamic Parameters ─────────────
        # Row 0 = card title ("📦 Extraction Modules") — DO NOT use row 0
        _lbl(card, "Module 1 · Global Stability", color=MUTED, size=11).grid(
            row=1, column=0, padx=15, pady=(10, 2), sticky="w"
        )
        self._mod1_modal = _check(card, "Modal Periods & Mass Participation")
        self._mod1_modal.grid(row=2, column=0, padx=15, pady=3, sticky="w")
        self._mod1_drift = _check(card, "Story Drifts & Displacements")
        self._mod1_drift.grid(row=3, column=0, padx=15, pady=3, sticky="w")
        self._mod1_base = _check(card, "Base Reactions (Vx, Vy, Fz)")
        self._mod1_base.grid(row=4, column=0, padx=15, pady=3, sticky="w")
        self._mod1_pdelta = _check(card, "P-Delta Stability Coeff. (θ)")
        self._mod1_pdelta.grid(row=5, column=0, padx=15, pady=3, sticky="w")
        self._mod1_torsion = _check(card, "Torsional Irregularity (CM/CR)")
        self._mod1_torsion.grid(row=6, column=0, padx=15, pady=(3, 14), sticky="w")

        # ── Module 2: Element-Level Optimization & Design ────────────────
        _lbl(card, "Module 2 · Element Design", color=MUTED, size=11).grid(
            row=1, column=1, padx=15, pady=(10, 2), sticky="w"
        )
        self._mod2_forces = _check(card, "Frame Internal Forces (P, V, M)")
        self._mod2_forces.grid(row=2, column=1, padx=15, pady=3, sticky="w")
        self._mod2_rebar = _check(card, "Required Rebar (Beam Design)")
        self._mod2_rebar.grid(row=3, column=1, padx=15, pady=3, sticky="w")
        self._mod2_pmm = _check(card, "Column P-M-M D/C Ratios")
        self._mod2_pmm.grid(row=4, column=1, padx=15, pady=3, sticky="w")
        self._mod2_scwb = _check(card, "Strong-Column/Weak-Beam (SRPMK)")
        self._mod2_scwb.grid(row=5, column=1, padx=15, pady=3, sticky="w")
        self._mod2_pier = _check(card, "Pier & Spandrel Summaries")
        self._mod2_pier.grid(row=6, column=1, padx=15, pady=(3, 14), sticky="w")

        # ── Module 3: Visual Documentation (Windows/ETABS only) ─────────
        _lbl(card, "Module 3 · Visual Documentation", color=MUTED, size=11).grid(
            row=1, column=2, padx=15, pady=(10, 2), sticky="w"
        )
        self._mod3_geometry = _check(card, "Model Geometry (3D Extruded)", default=False)
        self._mod3_geometry.grid(row=2, column=2, padx=15, pady=3, sticky="w")
        self._mod3_loading = _check(card, "Applied Loading (3D + Plan)", default=False)
        self._mod3_loading.grid(row=3, column=2, padx=15, pady=3, sticky="w")
        self._mod3_deform = _check(card, "Deformed Shapes (Mode 1 & 2)", default=False)
        self._mod3_deform.grid(row=4, column=2, padx=15, pady=3, sticky="w")
        self._mod3_dc = _check(card, "D/C Ratio Color Map (3D)", default=False)
        self._mod3_dc.grid(row=5, column=2, padx=15, pady=3, sticky="w")
        self._mod3_foundation = _check(card, "Foundation Reactions (Fz Plan)", default=False)
        self._mod3_foundation.grid(row=6, column=2, padx=15, pady=(3, 14), sticky="w")

    def _build_action_bar(self):
        bar = ctk.CTkFrame(self, fg_color=BG_MAIN)
        bar.grid(row=5, column=0, padx=20, pady=(8, 20), sticky="ew")
        bar.grid_columnconfigure(0, weight=1)

        self.progress_bar = ctk.CTkProgressBar(bar, height=8, corner_radius=4)
        self.progress_bar.grid(row=0, column=0, columnspan=2, padx=0, pady=(0, 8), sticky="ew")
        self.progress_bar.set(0)

        self.status_lbl = ctk.CTkLabel(
            bar, text="Ready.", font=FONT_MUTED, text_color=MUTED, anchor="w"
        )
        self.status_lbl.grid(row=1, column=0, padx=0, pady=0, sticky="w")

        self.gen_btn = ctk.CTkButton(
            bar,
            text="⚡  Generate Report",
            font=FONT_TITLE,
            fg_color=BLUE,
            hover_color=BLUE_H,
            height=46,
            corner_radius=10,
            command=self._start_generation,
        )
        self.gen_btn.grid(row=0, column=2, rowspan=2, padx=(20, 0), pady=0, sticky="e")

    # ──────────────────────────────────────────────────────────────────
    # File browsing
    # ──────────────────────────────────────────────────────────────────

    def _browse_edb(self):
        path = filedialog.askopenfilename(
            title="Select Primary ETABS Model",
            filetypes=[("ETABS Files", "*.edb"), ("All Files", "*.*")],
        )
        if path:
            self.edb_path.set(path)
            # Trigger background pre-fetch to populate filter dropdowns
            threading.Thread(target=self._prefetch_worker, args=(path,), daemon=True).start()
            self._start_polling()

    def _browse_edb_b(self):
        path = filedialog.askopenfilename(
            title="Select Comparison ETABS Model (optional)",
            filetypes=[("ETABS Files", "*.edb"), ("All Files", "*.*")],
        )
        if path:
            self.edb_b_path.set(path)

    def _browse_template(self):
        path = filedialog.askopenfilename(
            title="Select Word Template",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
        )
        if path:
            self.tpl_path.set(path)

    def _update_output_label(self, *_):
        tpl = self.tpl_path.get()
        if tpl:
            out = os.path.splitext(tpl)[0] + "_Generated.docx"
            self.output_path.set(f"Output → {os.path.basename(out)}")

    # ──────────────────────────────────────────────────────────────────
    # Pre-fetch worker (populates filter checkboxes)
    # ──────────────────────────────────────────────────────────────────

    def _prefetch_worker(self, edb_path: str):
        try:
            if ETABS_AVAILABLE:
                from etabs_extractor import ETABSExtractor
                ex = ETABSExtractor(edb_path, self._q)
                ex.connect()
                defs = ex.prefetch_model_definitions()
                ex.disconnect()
            else:
                from mock_extractor import MockEtabsExtractor
                ex = MockEtabsExtractor(progress_queue=self._q)
                defs = ex.prefetch_model_definitions()

            self._q.put({"type": "prefetch_done", "defs": defs})
        except Exception as exc:
            self._q.put({"type": "error", "message": str(exc)})

    def _populate_filters(self, defs: dict):
        """Called on main thread after prefetch completes."""
        self._model_defs = defs
        self._filter_hint.configure(
            text=f"✓ Loaded: {len(defs.get('stories', []))} stories · "
                 f"{len(defs.get('load_combos', []))} combos · "
                 f"{len(defs.get('sections', []))} sections · "
                 f"{len(defs.get('groups', []))} groups",
            text_color=GREEN,
        )

        self._fill_scroll(self._story_scroll,   self._story_checks,   defs.get("stories", []))
        self._fill_scroll(self._combo_scroll,   self._combo_checks,   defs.get("load_combos", []))
        self._fill_scroll(self._section_scroll, self._section_checks, defs.get("sections", []))
        self._fill_scroll(self._group_scroll,   self._group_checks,   defs.get("groups", []))

    def _fill_scroll(
        self,
        scroll: ctk.CTkScrollableFrame,
        checks: dict,
        items: list[str],
    ):
        # Clear existing
        for w in scroll.winfo_children():
            w.destroy()
        checks.clear()

        for item in items:
            var = ctk.BooleanVar(value=True)
            cb = ctk.CTkCheckBox(
                scroll,
                text=item,
                variable=var,
                font=FONT_MUTED,
                text_color=TEXT,
                border_color=BORDER,
                checkmark_color="#0F172A",
                fg_color=BLUE,
            )
            cb.pack(anchor="w", padx=6, pady=2)
            checks[item] = cb  # store widget; read .get() on run

    # ──────────────────────────────────────────────────────────────────
    # Report generation
    # ──────────────────────────────────────────────────────────────────

    def _start_generation(self):
        if not self.tpl_path.get():
            messagebox.showwarning("Missing Template", "Please select a Word template (.docx).")
            return

        if not self.edb_path.get() and ETABS_AVAILABLE:
            messagebox.showwarning("Missing Model", "Please select an ETABS model (.edb).")
            return

        # Pre-flight output file lock check
        tpl = self.tpl_path.get()
        output = os.path.splitext(tpl)[0] + "_Generated.docx" if tpl else "output.docx"
        if not check_word_file_lock(output):
            messagebox.showerror(
                "File Locked",
                f"The output file is open in another application:\n{output}\n\n"
                "Please close it and try again.",
            )
            return

        # Extract all parameters from widgets safely on the main thread
        edb_a = self.edb_path.get()
        edb_b = self.edb_b_path.get()
        proj  = self.proj_entry.get() or "Untitled Project"
        eng   = self.eng_entry.get()  or "—"

        filters = {
            "stories":     [k for k, v in self._story_checks.items()   if v.get()],
            "load_combos": [k for k, v in self._combo_checks.items()   if v.get()],
            "sections":    [k for k, v in self._section_checks.items() if v.get()],
            "groups":      [k for k, v in self._group_checks.items()   if v.get()],
        }

        modules = {
            "mod1_modal":   self._mod1_modal.get(),
            "mod1_drift":   self._mod1_drift.get(),
            "mod1_base":    self._mod1_base.get(),
            "mod1_pdelta":  self._mod1_pdelta.get(),
            "mod1_torsion": self._mod1_torsion.get(),
            "mod2_forces":  self._mod2_forces.get(),
            "mod2_rebar":   self._mod2_rebar.get(),
            "mod2_pmm":     self._mod2_pmm.get(),
            "mod2_scwb":    self._mod2_scwb.get(),
            "mod2_pier":    self._mod2_pier.get(),
        }

        self.gen_btn.configure(state="disabled", text="Generating…")
        self._start_polling()

        threading.Thread(
            target=self._generate_worker,
            args=(output, edb_a, edb_b, tpl, proj, eng, filters, modules),
            daemon=True
        ).start()

    def _generate_worker(
        self,
        output_path: str,
        edb_a: str,
        edb_b: str,
        tpl: str,
        proj: str,
        eng: str,
        filters: dict[str, list],
        modules: dict[str, bool],
    ):
        try:
            is_comparison = bool(edb_b)

            # ── Choose extractor ──────────────────────────────────────
            if ETABS_AVAILABLE:
                ex_a = ETABSExtractor(edb_a, self._q, filters=filters)
                ex_a.connect()
                ex_a.run_analysis()
            else:
                ex_a = MockEtabsExtractor(
                    target_stories=filters["stories"] or None,
                    load_combos=filters["load_combos"] or None,
                    target_sections=filters["sections"] or None,
                    progress_queue=self._q,
                )

            # ── Module 1: Global Stability & Dynamic Parameters ───────
            df_modal   = ex_a.extract_modal_results()          if modules["mod1_modal"]   else None
            df_drifts  = ex_a.extract_story_drifts()           if modules["mod1_drift"]   else None
            df_base    = ex_a.extract_base_reactions()         if modules["mod1_base"]    else None
            df_pdelta  = ex_a.extract_pdelta_stability()       if modules["mod1_pdelta"]  else None
            df_torsion = ex_a.extract_torsional_irregularity() if modules["mod1_torsion"] else None

            # ── Module 2: Element-Level Design ────────────────────────
            df_forces = ex_a.extract_frame_forces()     if modules["mod2_forces"] else None
            df_rebar  = ex_a.extract_rebar_results()    if modules["mod2_rebar"]  else None
            df_pmm    = ex_a.extract_pmm_ratios()       if modules["mod2_pmm"]    else None
            df_scwb   = ex_a.extract_scwb_ratios()      if modules["mod2_scwb"]   else None
            df_pier   = ex_a.extract_pier_spandrel()    if modules["mod2_pier"]   else None

            # ── Comparison model (Model B) ────────────────────────────
            df_forces_b = None
            if is_comparison and df_forces is not None:
                if ETABS_AVAILABLE:
                    ex_b = ETABSExtractor(edb_b, self._q, filters=filters)
                    ex_b.connect()
                    ex_b.run_analysis()
                    df_forces_b = ex_b.extract_frame_forces()
                    ex_b.disconnect()
                else:
                    ex_b = MockEtabsExtractor(
                        target_stories=filters["stories"] or None,
                        load_combos=filters["load_combos"] or None,
                        progress_queue=self._q,
                    )
                    df_forces_b = ex_b.extract_frame_forces(is_optimized_model=True)

            if ETABS_AVAILABLE:
                ex_a.disconnect()

            # ── Assemble context ──────────────────────────────────────
            self._post_progress("Assembling report context...", 0.85)
            context = data_pipeline.build_report_context(
                project_name=proj,
                engineer_name=eng,
                df_forces=df_forces,
                df_rebar=df_rebar,
                df_modal=df_modal,
                df_drifts=df_drifts,
                df_base_reactions=df_base,
                df_pmm=df_pmm,
                df_pdelta=df_pdelta,
                df_torsion=df_torsion,
                df_scwb=df_scwb,
                df_pier=df_pier,
                df_forces_b=df_forces_b,
                is_comparison_mode=is_comparison,
            )

            # ── Generate Word document ────────────────────────────────
            self._post_progress("Rendering Word template...", 0.90)
            report_generator.generate_report(context, tpl, output_path)

            self._q.put({"type": "done", "output_path": output_path})

        except Exception as exc:
            self._q.put({"type": "error", "message": str(exc)})

    # ──────────────────────────────────────────────────────────────────
    # Queue polling (thread → main-thread bridge)
    # ──────────────────────────────────────────────────────────────────

    def _start_polling(self):
        if self._poll_id is None:
            self._poll_queue()

    def _poll_queue(self):
        try:
            while True:
                msg = self._q.get_nowait()
                self._handle_message(msg)
        except queue.Empty:
            pass
        except Exception:
            pass
        finally:
            try:
                if self.winfo_exists():
                    self._poll_id = self.after(100, self._poll_queue)
            except Exception:
                self._poll_id = None

    def _safe_update(self):
        try:
            if self.winfo_exists():
                self.update()
        except Exception:
            pass

    def destroy(self):
        if hasattr(self, "_update_id") and self._update_id is not None:
            try:
                self.after_cancel(self._update_id)
            except Exception:
                pass
            self._update_id = None
        if hasattr(self, "_poll_id") and self._poll_id is not None:
            try:
                self.after_cancel(self._poll_id)
            except Exception:
                pass
            self._poll_id = None
        super().destroy()

    def _handle_message(self, msg: dict):
        mtype = msg.get("type")

        if mtype == "progress":
            self.progress_bar.set(msg.get("value", 0))
            self.status_lbl.configure(text=msg.get("message", ""), text_color=MUTED)

        elif mtype == "prefetch_done":
            self._populate_filters(msg["defs"])
            self._post_status("✓ Model definitions loaded — filters populated.", GREEN)

        elif mtype == "done":
            out = msg.get("output_path", "")
            self.progress_bar.set(1.0)
            self._post_status(f"✅  Report generated → {os.path.basename(out)}", GREEN)
            self.gen_btn.configure(state="normal", text="⚡  Generate Report")
            messagebox.showinfo(
                "Report Complete",
                f"Report generated successfully!\n\nSaved to:\n{out}",
            )

        elif mtype == "error":
            self.progress_bar.set(0)
            err = msg.get("message", "Unknown error.")
            self._post_status(f"⛔  Error: {err}", RED)
            self.gen_btn.configure(state="normal", text="⚡  Generate Report")
            messagebox.showerror("Extraction Error", err)

    # ──────────────────────────────────────────────────────────────────
    # Safe status helpers (called from main thread only)
    # ──────────────────────────────────────────────────────────────────

    def _post_progress(self, text: str, value: float):
        self._q.put({"type": "progress", "message": text, "value": value})

    def _post_status(self, text: str, color: str = MUTED):
        self.status_lbl.configure(text=text, text_color=color)


# ═══════════════════════════════════════════════════════════════════════════
# Small widget factories
# ═══════════════════════════════════════════════════════════════════════════

def _lbl(parent, text: str, color: str = TEXT, size: int = 12) -> ctk.CTkLabel:
    return ctk.CTkLabel(
        parent, text=text,
        font=("Helvetica Neue", size),
        text_color=color,
    )


def _entry(parent, textvariable=None, placeholder: str = "") -> ctk.CTkEntry:
    kwargs = dict(
        fg_color="#141E30",
        border_color=BORDER,
        text_color=TEXT,
        placeholder_text_color=MUTED,
        font=FONT_BODY,
        height=34,
        corner_radius=6,
    )
    if textvariable is not None:
        kwargs["textvariable"] = textvariable
    if placeholder:
        kwargs["placeholder_text"] = placeholder
    return ctk.CTkEntry(parent, **kwargs)


def _btn(parent, text: str, command=None) -> ctk.CTkButton:
    return ctk.CTkButton(
        parent,
        text=text,
        command=command,
        fg_color=BG_CARD,
        hover_color="#263248",
        border_width=1,
        border_color=BORDER,
        text_color=TEXT,
        font=FONT_BODY,
        height=34,
        width=90,
        corner_radius=6,
    )


def _check(parent, text: str, default: bool = True) -> ctk.CTkCheckBox:
    var = ctk.BooleanVar(value=default)
    cb = ctk.CTkCheckBox(
        parent,
        text=text,
        variable=var,
        font=FONT_BODY,
        text_color=TEXT,
        border_color=BORDER,
        checkmark_color="#0F172A",
        fg_color=BLUE,
    )
    # Monkey-patch .get() → var.get() for convenience
    cb.get = var.get  # type: ignore[attr-defined]
    return cb


# ═══════════════════════════════════════════════════════════════════════════
# Entry point
# ═══════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app = ETABSReportApp()
    app.mainloop()
