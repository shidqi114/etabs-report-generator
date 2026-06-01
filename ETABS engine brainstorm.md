# ETABS Automated Report & Comparison Engine 📊🏗️

An advanced, GUI-driven automation tool designed for senior structural engineers. It connects directly to the **CSI ETABS COM API** to extract targeted analytical data, perform cross-model variance calculations, and dynamically generate formatted, code-compliant Microsoft Word structural reports using Jinja2 templating.

---

## 🌟 Overview & Engineering Value

Standard structural reporting involves massive data dumps and repetitive formatting. This software shifts the paradigm from simple "data extraction" to **automated structural verification**. 

**Core Capabilities:**
1. **Precision Extraction**: Pulls exact verification metrics (drifts, mass participation, D/C ratios) rather than just raw forces.
2. **Multi-Model Comparison**: Analyzes multiple `.edb` files simultaneously to compute variance ($\Delta$) for value engineering and sensitivity analysis.
3. **Strict Unit Hygiene**: Forces the ETABS API into a standardized unit system (e.g., kN-m-C) prior to extraction to guarantee calculation safety.
4. **Dynamic Word Templating**: Utilizes `docxtpl` to generate distinct, beautifully formatted tables for specific load combinations without breaking company standards.

---

## 🛠️ System Architecture

The application is built on a decoupled, modular Python architecture to prevent API bottlenecks and UI freezing during heavy execution.

* **`app.py`**: A fully responsive CustomTkinter GUI featuring dynamic multi-select filters populated directly from the `.edb` file definitions.
* **`etabs_extractor.py`**: The background thread (`threading.Thread`) handling low-level COM API communication, unit overrides, and data fetching.
* **`data_pipeline.py`**: Processes flat Pandas DataFrames, calculates variance ($\Delta$) for multi-model runs, and groups data hierarchically by Load Combination.
* **`report_generator.py`**: Integrates the processed Python dictionaries into a tagged Word document (`.docx`), generating looped tables and injecting automated 3D screenshots.

---

## 🎯 Precision Filtering System (The Pre-Fetch)

To prevent memory overload and ensure the extracted data is highly targeted (essential for pushing payloads into downstream simulation wrappers like OpenSees or BIM models), the GUI automatically pre-fetches model definitions. 

Users can strictly filter extraction by:
* **Target Stories** (e.g., `Story 1`, `Roof`)
* **Target Groups** (e.g., `Core_Walls`, `Typical_Floor_Beams`)
* **Target Frame/Shell Sections** (e.g., `B-40x80`, `Wall-300`)
* **Target Load Combinations** (e.g., `ENV-ALL`, `1.2D+1.0Ex+0.3Ey`)

---

## 📋 Modular Extraction Checklist

The software categorizes extraction into three core engineering modules. Users can toggle specific parameters based on the report phase.

### Module 1: Global Stability & Dynamic Parameters
*Code compliance verification (e.g., SNI 1726).*
* **Modal Periods & Frequencies**: Time periods for primary modes.
* **Mass Participation Ratios**: Verifies >90% participation in X/Y directions.
* **Story Drifts & Displacements**: Maximum drift ratios and center of mass displacements.
* **Base Reactions ($V_x, V_y, F_z$)**: Global shears and axial loads for foundations.
* **P-Delta Stability Coefficient ($\theta$)**: Story shears and weights for secondary moment checks.
* **Torsional Irregularity**: CM vs. CR coordinates and max/avg drift ratios.

### Module 2: Element-Level Optimization & Design
*Member sizing, demand/capacity, and reinforcement detailing.*
* **Frame Internal Forces**: Envelope forces ($P, V_2, V_3, M_2, M_3$).
* **Required Rebar**: Longitudinal and transverse areas for beams and columns.
* **Column P-M-M Interaction**: Extracts D/C ratios and automatically flags elements $> 1.0$.
* **Strong-Column/Weak-Beam**: Joint capacity ratios for Special Moment Frames (SRPMK).
* **Pier & Spandrel Summaries**: Core wall forces and required reinforcement ratios.

### Module 3: Visual Documentation
*Automated API image capturing (`cOAPI.View.CaptureImage`) for the report appendix.*
* **Model Geometry**: 3D extruded views of frames and shells.
* **Applied Loading**: 3D and Plan views of assigned distributed frame and area shell loads.
* **Deformed Shapes**: Mode 1 and Mode 2 translational verification.
* **Design Verification**: 3D view color-coded by frame D/C ratios (Pass/Fail).
* **Foundation Reactions**: Plan view displaying $F_z$ node text.

---

## 📝 Dynamic Word Templating Engine

The data pipeline transitions the flat Pandas extractions into nested dictionaries, allowing the Word template to generate isolated tables per Load Combination.

### Multi-Model Comparison Example
When comparing an "Existing" model vs. a "Retrofit" model, the tool calculates the absolute and percentage variance before sending the payload to Word.

**Word Template Tagging Structure (`docxtpl`):**
```jinja
{% for lc in combinations_list %}
**Load Combination:** {{ lc.combo_name }}

| Frame | M3 (Existing) | M3 (Retrofit) | Variance ($\Delta$) | Pass/Fail |
| :--- | :--- | :--- | :--- | :--- |
| {% tr for item in lc.forces_data %} {{ item.Frame }} | {{ item.M3_A }} | {{ item.M3_B }} | {{ item.Variance }} | {{ item.Status }} {% endtr %} |

{% endfor %}