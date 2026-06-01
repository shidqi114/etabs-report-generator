"""
data_pipeline.py
================
Data processing layer for the ETABS Report Generator.

Responsibilities:
  1. safe_model_merge()  – Outer-join two model DataFrames, compute variance (Δ),
                           and fill topology mismatches gracefully (no NaN errors).
  2. group_by_combo()    – Reshape a flat DataFrame into a nested dict keyed by
                           Load Combination, ready for Jinja2 / docxtpl rendering.
  3. build_report_context() – Assembles the full context dictionary passed to
                              report_generator.generate_report().

No COM API calls occur here.  All functions operate purely on pandas DataFrames.
"""

from __future__ import annotations

import numpy as np
import pandas as pd
from typing import Any


# ---------------------------------------------------------------------------
# 1. Topology-safe model merge (for multi-model comparison)
# ---------------------------------------------------------------------------

def safe_model_merge(
    df_baseline: pd.DataFrame,
    df_optimized: pd.DataFrame,
    merge_keys: list[str] | None = None,
    force_columns: list[str] | None = None,
) -> pd.DataFrame:
    """
    Merges two model DataFrames with an outer join so topology changes
    (added / deleted elements) are preserved rather than silently dropped.

    Variance (Δ) is computed for all numeric force columns where BOTH models
    have a value.  Missing entries are flagged with human-readable strings
    instead of NaN so that Jinja2 templates can render them cleanly.

    Parameters
    ----------
    df_baseline : pd.DataFrame
        Extraction results from Model A (e.g., "Existing" structure).
    df_optimized : pd.DataFrame
        Extraction results from Model B (e.g., "Retrofit" structure).
    merge_keys : list[str]
        Columns that uniquely identify a result row across both models.
        Defaults to ['Frame', 'Station', 'LoadCase'].
    force_columns : list[str]
        Numeric force columns for which Δ is computed.
        Defaults to ['P', 'V2', 'V3', 'T', 'M2', 'M3'].

    Returns
    -------
    pd.DataFrame
        Merged DataFrame with _A / _B suffixed columns plus Variance_<col> columns.
    """
    if merge_keys is None:
        merge_keys = ["Frame", "Station", "LoadCase"]
    if force_columns is None:
        force_columns = ["P", "V2", "V3", "T", "M2", "M3"]

    # Outer merge to capture topology mismatches
    df_merged = pd.merge(
        df_baseline,
        df_optimized,
        on=merge_keys,
        how="outer",
        suffixes=("_A", "_B"),
    )

    # Compute variance for each force column only when both sides are present
    for col in force_columns:
        col_a = f"{col}_A"
        col_b = f"{col}_B"
        if col_a in df_merged.columns and col_b in df_merged.columns:
            df_merged[f"Variance_{col}"] = np.where(
                df_merged[col_a].notna() & df_merged[col_b].notna(),
                (df_merged[col_b] - df_merged[col_a]).round(3),
                "Topology Changed",
            )

    # Fill remaining NaNs with descriptive strings for clean Jinja rendering
    for col in force_columns:
        col_a = f"{col}_A"
        col_b = f"{col}_B"
        if col_a in df_merged.columns:
            df_merged[col_a] = df_merged[col_a].fillna("Not in Model A")
        if col_b in df_merged.columns:
            df_merged[col_b] = df_merged[col_b].fillna("Not in Model B")

    # Pass-through non-force NaNs in non-numeric categorical columns
    for col in merge_keys:
        if col in df_merged.columns:
            df_merged[col] = df_merged[col].fillna("—")

    return df_merged


# ---------------------------------------------------------------------------
# 2. Group flat DataFrame by Load Combination for Jinja2 table loops
# ---------------------------------------------------------------------------

def group_by_combo(
    df: pd.DataFrame,
    combo_col: str = "LoadCase",
    round_decimals: int = 3,
) -> list[dict[str, Any]]:
    """
    Transforms a flat extraction DataFrame into a list of dicts, one per
    Load Combination, suitable for the docxtpl ``{% for lc in combinations_list %}``
    template loop.

    Each dict has the shape::

        {
            "combo_name": "1.2D+1.0Ex+0.3Ey",
            "forces_data": [
                {"Frame": "B1-Story1", "M3": 245.3, "Status": "✓ OK", ...},
                ...
            ]
        }

    Parameters
    ----------
    df : pd.DataFrame
        Flat extraction result (from ETABSExtractor or MockEtabsExtractor).
    combo_col : str
        Column name that holds the Load Combination label.
    round_decimals : int
        Decimal places to round numeric values before serialisation.

    Returns
    -------
    list[dict]
        Ready-to-render list for the Jinja2 ``combinations_list`` variable.
    """
    if df.empty or combo_col not in df.columns:
        return []

    # Round numeric columns for clean template output
    numeric_cols = df.select_dtypes(include=[np.number]).columns
    df = df.copy()
    df[numeric_cols] = df[numeric_cols].round(round_decimals)

    result = []
    for combo_name, group_df in df.groupby(combo_col, sort=False):
        result.append(
            {
                "combo_name": combo_name,
                "forces_data": group_df.to_dict("records"),
            }
        )
    return result


# ---------------------------------------------------------------------------
# 3. Build the full docxtpl context dictionary
# ---------------------------------------------------------------------------

def build_report_context(
    project_name: str,
    engineer_name: str,
    *,
    df_forces: pd.DataFrame | None = None,
    df_rebar: pd.DataFrame | None = None,
    df_modal: pd.DataFrame | None = None,
    df_drifts: pd.DataFrame | None = None,
    df_base_reactions: pd.DataFrame | None = None,
    df_pmm: pd.DataFrame | None = None,
    # Module 1 additions
    df_pdelta: pd.DataFrame | None = None,
    df_torsion: pd.DataFrame | None = None,
    # Module 2 additions
    df_scwb: pd.DataFrame | None = None,
    df_pier: pd.DataFrame | None = None,
    # Multi-model comparison
    df_forces_b: pd.DataFrame | None = None,
    is_comparison_mode: bool = False,
) -> dict[str, Any]:
    """
    Assembles every template variable into one context dict.

    The dict is passed directly to ``DocxTemplate.render()``.

    Parameters
    ----------
    project_name, engineer_name : str
        Header metadata injected into the cover page.
    df_forces : pd.DataFrame
        Frame forces (Module 2).  If ``is_comparison_mode`` is True and
        ``df_forces_b`` is also provided, runs safe_model_merge() and
        populates ``combinations_comparison_list`` instead.
    df_forces_b : pd.DataFrame
        Retrofit / optimised model forces for comparison.
    is_comparison_mode : bool
        Activates multi-model comparison rendering path.
    All other DataFrames map to Module 1 / Module 2 tables.

    Returns
    -------
    dict
        Full context ready for docxtpl.
    """
    ctx: dict[str, Any] = {
        "project_name": project_name,
        "engineer_name": engineer_name,
        "is_comparison_mode": is_comparison_mode,
    }

    # ── Module 1: Global Stability ────────────────────────────────────
    ctx["modal_list"] = _df_to_records(df_modal)
    ctx["drifts_list"] = _df_to_records(df_drifts)
    ctx["base_reactions_list"] = _df_to_records(df_base_reactions)

    # Group drifts by combo for per-combination story tables
    ctx["drifts_by_combo"] = group_by_combo(
        df_drifts, combo_col="LoadCombo"
    ) if df_drifts is not None and not df_drifts.empty else []

    # P-Delta stability coefficients
    ctx["pdelta_list"] = _df_to_records(df_pdelta)
    ctx["pdelta_by_combo"] = group_by_combo(
        df_pdelta, combo_col="LoadCombo"
    ) if df_pdelta is not None and not df_pdelta.empty else []

    # Torsional irregularity (CM vs CR)
    ctx["torsion_list"] = _df_to_records(df_torsion)
    if df_torsion is not None and not df_torsion.empty and "Irregularity" in df_torsion.columns:
        irregular = df_torsion[df_torsion["Irregularity"].str.contains("Torsional", na=False)]
        ctx["torsion_irregular_count"] = int(len(irregular))
        ctx["torsion_irregular_list"] = _df_to_records(irregular)
    else:
        ctx["torsion_irregular_count"] = 0
        ctx["torsion_irregular_list"] = []

    # ── Module 2: Element Design ────────────────────────────────────
    ctx["rebar_list"] = _df_to_records(df_rebar)
    ctx["pmm_list"] = _df_to_records(df_pmm)

    # Flag any D/C failures for the executive summary
    if df_pmm is not None and not df_pmm.empty and "DC_Ratio" in df_pmm.columns:
        failed = df_pmm[df_pmm["DC_Ratio"] > 1.0]
        ctx["pmm_failures_count"] = int(len(failed))
        ctx["pmm_failures_list"] = _df_to_records(failed)
    else:
        ctx["pmm_failures_count"] = 0
        ctx["pmm_failures_list"] = []

    # Strong-Column / Weak-Beam ratios
    ctx["scwb_list"] = _df_to_records(df_scwb)
    if df_scwb is not None and not df_scwb.empty and "Status" in df_scwb.columns:
        scwb_fail = df_scwb[df_scwb["Status"].str.contains("FAIL", na=False)]
        ctx["scwb_failures_count"] = int(len(scwb_fail))
        ctx["scwb_failures_list"] = _df_to_records(scwb_fail)
    else:
        ctx["scwb_failures_count"] = 0
        ctx["scwb_failures_list"] = []

    # Pier & Spandrel summaries
    ctx["pier_spandrel_list"] = _df_to_records(df_pier)
    ctx["pier_spandrel_by_combo"] = group_by_combo(
        df_pier, combo_col="LoadCombo"
    ) if df_pier is not None and not df_pier.empty else []

    # ── Frame Forces: single-model vs comparison ───────────────────────
    if is_comparison_mode and df_forces is not None and df_forces_b is not None:
        # Build comparison table grouped by combo
        df_merged = safe_model_merge(df_forces, df_forces_b)
        ctx["combinations_comparison_list"] = group_by_combo(df_merged)
        ctx["combinations_list"] = []  # unused in comparison mode
    else:
        ctx["combinations_list"] = group_by_combo(df_forces) if df_forces is not None else []
        ctx["combinations_comparison_list"] = []

    return ctx


# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------

def _df_to_records(df: pd.DataFrame | None, max_rows: int | None = None) -> list[dict]:
    """Safely converts a DataFrame to a list of dicts, rounding floats."""
    if df is None or df.empty:
        return []
    out = df.copy()
    numeric_cols = out.select_dtypes(include=[np.number]).columns
    out[numeric_cols] = out[numeric_cols].round(3)
    if max_rows:
        out = out.head(max_rows)
    return out.to_dict("records")
