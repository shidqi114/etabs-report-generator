# ETABS Report Generator 📊🏗️

An advanced, GUI-driven automation tool for senior structural engineers. It connects directly to the **CSI ETABS COM API**, extracts targeted verification metrics across multiple modules, performs multi-model variance analysis, and generates formatted, code-compliant Microsoft Word reports using Jinja2 templating — fully testable on macOS via a high-fidelity Mock Data Engine.

---

## 🌟 Overview & Engineering Value

Standard structural reporting involves massive data dumps and repetitive formatting. This software shifts the paradigm from simple "data extraction" to **automated structural verification**.

| Capability | Detail |
| :--- | :--- |
| **Precision Extraction** | Pulls exact verification metrics (drifts, mass participation, D/C ratios) rather than raw forces |
| **Multi-Model Comparison** | Analyzes two `.edb` files simultaneously; computes Δ variance for value engineering |
| **Strict Unit Hygiene** | Forces the ETABS API into kN-m-°C before any extraction to guarantee calculation safety |
| **Dynamic Word Templating** | `docxtpl` + Jinja2 generates isolated tables per Load Combination |
| **Auto TOC Update** | Word COM API finalizes page numbers and Table of Contents after rendering |
| **Cross-Platform Dev** | Mock Data Engine activates automatically on macOS/Linux — no ETABS required |

---

## 🛠️ Architecture & System Workflow

```
┌───────────────────────────────────────────────────────────────┐
│                         app.py  (GUI)                         │
│  CustomTkinter · Dark Slate theme · queue.Queue thread bridge │
└────┬───────────────────────────────────────┬──────────────────┘
     │ Background Thread                      │ Background Thread
     ▼                                        ▼
┌──────────────────┐                 ┌─────────────────────┐
│ etabs_extractor  │  (Windows)      │   mock_extractor    │  (macOS/Linux)
│ ETABSExtractor   │  ←─── OR ──→   │ MockEtabsExtractor  │
│ COM API + psutil │                 │ Plausible stub data │
└────────┬─────────┘                 └──────────┬──────────┘
         │ pandas DataFrames                    │
         └──────────────┬───────────────────────┘
                        ▼
              ┌──────────────────┐
              │  data_pipeline   │
              │  safe_model_merge│  outer join + Δ variance
              │  group_by_combo  │  flat → combo-keyed dicts
              │  build_report_context │
              └────────┬─────────┘
                       ▼
            ┌────────────────────┐
            │  report_generator  │
            │  DocxTemplate.render()      │
            │  update_word_toc() │  Word COM (Windows) / skip (macOS)
            └────────┬───────────┘
                     ▼
         ┌─────────────────────────┐
         │  *_Generated.docx       │
         │  (TOC & page numbers    │
         │   auto-refreshed)       │
         └─────────────────────────┘
```

### Component Breakdown

| File | Responsibility |
| :--- | :--- |
| **`app.py`** | Premium CustomTkinter GUI. Precision filter dropdowns auto-populated by pre-fetch. Module checklist, second EDB slot for comparison mode, thread-safe `queue.Queue` bridge. |
| **`etabs_extractor.py`** | Live COM layer. OS detection, ghost-process sweep (`psutil`), kN-m-°C unit enforcement, `ETABSExtractor` class with prefetch + 3-module extraction methods. |
| **`mock_extractor.py`** | Cross-platform dev stub. `MockEtabsExtractor` produces structurally plausible data matching the exact schema of the real extractor. |
| **`data_pipeline.py`** | Pure-pandas processing. Outer-join model merge, variance calculation, topology-mismatch handling, combo grouping, full context assembly. |
| **`report_generator.py`** | Document construction. `docxtpl` rendering + `update_word_toc()` via Word COM for automatic TOC refresh. |
| **`create_template.py`** | Helper: appends Jinja-tagged tables to an existing master `.docx`. |

---

## 🛑 Defensive Error-Handling Architecture (4 Layers)

| Layer | Where | What it prevents |
| :--- | :--- | :--- |
| **1. Pre-flight File Lock** | `check_word_file_lock()` in `etabs_extractor.py` | Fatal `PermissionError` when target `.docx` is open in Word |
| **2. COM Process Sweep** | `check_and_clear_etabs()` with `psutil` | Frozen `GetActiveObject()` from orphaned `ETABS.exe` instances |
| **3. Topology Mismatch** | `safe_model_merge()` in `data_pipeline.py` | `NaN` calculation errors from deleted/added elements between models |
| **4. Thread-Safe Routing** | `queue.Queue` + `self.after()` in `app.py` | GUI freezes from widget updates called off the main thread |

---

## 🎯 Precision Extraction Filters (Pre-Fetch)

On selecting an `.edb` file, a background thread prefetches model definitions and auto-populates four filter columns:

* **Target Stories** — e.g., `Story 1`, `Roof`
* **Target Groups** — e.g., `Core_Walls`, `Typical_Floor_Beams`
* **Target Sections** — e.g., `B-40x80`, `Wall-300`
* **Target Load Combos** — e.g., `ENV-ALL`, `1.2D+1.0Ex+0.3Ey`

Unchecked items are excluded from the extraction, reducing memory usage and API round-trips.

---

## 📋 Modular Extraction Checklist

### Module 1: Global Stability & Dynamic Parameters
*(Code compliance — e.g., SNI 1726)*
- ☑ Modal Periods & Frequencies
- ☑ Mass Participation Ratios (verifies > 90% in X/Y)
- ☑ Story Drifts & Displacements
- ☑ Base Reactions (Vx, Vy, Fz)

### Module 2: Element-Level Design
*(Member sizing, demand/capacity, reinforcement)*
- ☑ Frame Internal Forces (P, V2, V3, T, M2, M3)
- ☑ Required Rebar (longitudinal + transverse for beams)
- ☑ Column P-M-M Interaction D/C Ratios (auto-flags > 1.0)

### Module 3: Visual Documentation
*(Image capture — live Windows/ETABS only)*
- ☐ Model Geometry (3D extruded)
- ☐ Deformed Shapes (Mode 1 & 2)
- ☐ D/C Ratio Color Map (3D)

---

## 📝 Word Template Tag Schema

### Single Placeholders
| Tag | Type | Description |
| :--- | :--- | :--- |
| `{{ project_name }}` | String | Project name (cover page) |
| `{{ engineer_name }}` | String | Analyzing engineer's name |
| `{{ is_comparison_mode }}` | Boolean | True when two models are compared |
| `{{ pmm_failures_count }}` | Integer | Number of columns with D/C > 1.0 |

### Table Loop Variables
| Variable | Loop Tag | Description |
| :--- | :--- | :--- |
| `modal_list` | `{% tr for item in modal_list %}` | Modal periods + mass participation |
| `drifts_list` | `{% tr for item in drifts_list %}` | Flat story drift table |
| `drifts_by_combo` | `{% for lc in drifts_by_combo %}` | Drifts grouped by Load Combo |
| `base_reactions_list` | `{% tr for item in base_reactions_list %}` | Base shears and vertical reactions |
| `combinations_list` | `{% for lc in combinations_list %}` | Forces grouped by Load Combo |
| `combinations_comparison_list` | `{% for lc in combinations_comparison_list %}` | Merged Model A vs B with Δ |
| `rebar_list` | `{% tr for item in rebar_list %}` | Required beam reinforcement |
| `pmm_list` | `{% tr for item in pmm_list %}` | Column D/C ratios (all elements) |
| `pmm_failures_list` | `{% tr for item in pmm_failures_list %}` | Only failing columns (D/C > 1.0) |

### Multi-Model Comparison Template Example
```jinja
{% for lc in combinations_comparison_list %}
**Load Combination:** {{ lc.combo_name }}

| Frame | M3 (Existing) | M3 (Retrofit) | Variance (Δ) |
| :--- | :--- | :--- | :--- |
{% tr for item in lc.forces_data %}{{ item.Frame }} | {{ item.M3_A }} | {{ item.M3_B }} | {{ item.Variance_M3 }}{% endtr %}

{% endfor %}
```

> **Important**: The Table of Contents in `Report_Template_Tagged.docx` must be a **native Word automatic TOC** (References → Table of Contents → Automatic Table). The `update_word_toc()` function uses Word field codes to trigger the refresh — a plain-text list will not work.

---

## 🚀 Installation & Getting Started

### Prerequisites
- Python 3.8+
- *(For live ETABS)* Microsoft Windows + CSI ETABS installed

### Setup

```bash
# 1. Create and activate virtual environment
python -m venv venv
source venv/bin/activate          # macOS/Linux
# venv\Scripts\activate           # Windows

# 2. Install dependencies
pip install customtkinter pandas docxtpl python-docx numpy

# Windows-only (for live ETABS + TOC update):
# pip install comtypes psutil pywin32
```

### Running

```bash
# Optional: pre-tag your existing master document
python create_template.py

# Launch the GUI
python app.py
```

**On macOS / without ETABS**: The app starts in ⚡ DEV MODE. Select any `.docx` template, click **Generate Report** — the Mock Data Engine activates automatically and produces a fully rendered output document.

**On Windows with ETABS**: Select the `.edb` model (triggers prefetch → filter population), configure module checkboxes, optionally add a second Model B for comparison, then Generate.

---

## 🔄 Changelog

### v2.0 — Current
- **New**: `mock_extractor.py` — structurally plausible stub data for all 3 modules
- **New**: `data_pipeline.py` — `safe_model_merge()`, `group_by_combo()`, `build_report_context()`
- **Upgraded**: `etabs_extractor.py` — OS detection, 4-layer error handling, `ETABSExtractor` class with prefetch + full module coverage
- **Upgraded**: `report_generator.py` — `update_word_toc()` via Word COM, new context-dict API
- **Upgraded**: `app.py` — filter panel, module checklist, comparison mode, thread-safe queue bridge, premium Slate palette

### v1.0 — Initial Release
- Basic beam forces + rebar extraction
- Simple dummy data fallback on macOS
- Single-file template generation
