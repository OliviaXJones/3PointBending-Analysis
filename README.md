# 3-Point Bending Analysis Pipeline

Automated pipeline for analyzing, graphing, and exporting force-displacement data from the Biomomentum Mach-1 mechanical testing system. Supports multiple studies, genotyping workflows, and publication-ready overlay graphs.

---

## Entry Points

| Interface | File | Use case |
|---|---|---|
| Desktop GUI | `Main_Dashboard.py` | Primary workflow hub — study management, analysis, overlay graphs |
| Web app | `streamlit_app.py` | Cloud-based single-study analysis (Google Cloud Run + GCS) |
| Batch CLI | `Fz_Displacement_3PB_3.0.py` | Quick one-off batch analysis of any folder, no study setup needed |

---

## Desktop Dashboard — Workflow Modes

Launch with `python Main_Dashboard.py`. Select a mode from the top dropdown.

### FKBP5 Heat
Hardcoded genotyping pipeline for the FKBP5 Heat study.

**Mouse code format:** `W.12.F23` → Genotype . Age(wks) . Sex+ID

| Code | Genotype |
|---|---|
| W / 2W | Wildtype |
| M / 2M | Mutant |
| H / 2H | Heterozygous (2x = Pre-heat lineage) |

Outputs: individual PNG plots, per-session Excel summary, Tibia + Femur master Excel sync, genotype + lineage CSVs, anatomical diameter CSVs.

### FKBP5 New
Same pipeline as FKBP5 Heat with a different code scheme and **no lineage folder** — genotype-only CSVs.

**Mouse code format:** `Z.4.M2` → Genotype . Age(wks) . Sex+ID

| Code | Genotype |
|---|---|
| W | Wildtype |
| Z | Heterozygous |
| X | Mutant |

### Custom Studies (Single Study)
Configurable template for any cohort. Studies are defined in `studies_config.json` and can be added, edited, or removed via the **+ Add Study / Edit Study / Remove Study** buttons.

Each study stores: raw data folder, measurements Excel, output folder, sex, age, bone type, and a group prefix map (e.g. `CV → Control + Medigel`). Supports Male, Female, and Mixed cohorts with per-group sex assignment.

### Overlay Graphs
Click the **Overlay Graphs** button (top-right of the management row) to enter overlay mode. Select a study from the dropdown to auto-fill the scan folder, then:

1. Click **Scan .txt Files** — recursively finds all `.txt` files including group subfolders
2. Use the **Search** box to filter by filename or label
3. Check the files to include, set a **Mouse Code / Label** and **Color** per file
4. Set a **Graph Title** and **Save Folder**
5. Click **Generate Overlay**

Outputs a 300 DPI PNG with toe-adjusted x-axis (all curves aligned at x = 0), stiffness slope lines, max load markers, and failure point markers.

---

## Shared Analysis Core

All analysis logic lives in `bending_core.py`:

- `read_bending_txt` — parses Mach-1 `.txt` files (handles `<DATA>` blocks, column aliasing, sign inversion)
- `dominant_linear_region` — finds the longest R² ≥ 0.995 window for stiffness calculation

**Constants (consistent across all workflows):**

| Constant | Value | Purpose |
|---|---|---|
| `TOE_LOAD_FRACTION` | 0.05 | 5% of max load — defines toe region cutoff |
| `LINEAR_WINDOW_POINTS` | 90 | Minimum points for stiffness fit |
| `MIN_R2` | 0.995 | Minimum R² for linear region |
| `DISPLACEMENT_LIMIT` | 1.75 mm | Caps peak search range |
| `NEAR_ZERO_THRESHOLD` | 0.5 N | Failure trigger (load near zero) |
| `DROP_THRESHOLD_FRACTION` | 0.80 | 20% load drop — secondary failure trigger |

---

## Study Configuration

Studies are stored in `studies_config.json`. Example entry:

```json
"IFS+SHP099+Medigel 2026": {
    "raw_data_root": "E:\\...\\IFS+SHP099+Medigel_LFemur_051226",
    "measurement_file": "E:\\...\\IFS+SHP099+Medigel_LFemur_051226.xlsx",
    "output_folder": "E:\\...\\IFS+SHP099+Medigel_CSVFiles",
    "sex": "Male",
    "age": "21 Weeks",
    "bone": "Femur",
    "group_map": {
        "CV": "Control + Medigel",
        "PV": "IFS + Medigel",
        "PS": "IFS + SHP Medigel"
    }
}
```

---

## Web App

The Streamlit web app (`streamlit_app.py`) runs the Single Study workflow from a browser. Deployed on Google Cloud Run with a GCS bucket for master file persistence.

- Upload a `.zip` of raw `.txt` files + measurement Excel
- Edit the group map inline
- Download results as a `.zip`
- Master files accumulate across sessions via GCS

**Deployment:** requires `streamlit_app.py`, `SingleStudy_workflow.py`, and `bending_core.py` in the container (see `Dockerfile`).

---

## File Structure

```
Main_Dashboard.py          — PyQt6 GUI hub
streamlit_app.py           — Streamlit cloud interface
bending_core.py            — Shared analysis primitives
FKBP5Heat_workflow.py      — FKBP5 Heat genotyping pipeline
FKBP5New_workflow.py       — FKBP5 New genotyping pipeline (W/Z/X)
SingleStudy_workflow.py    — Configurable single-study pipeline
OverlayGraphs_workflow.py  — Overlay graph generation
Fz_Displacement_3PB_3.0.py — Standalone batch analyzer (no study setup)
studies_config.json        — Study definitions
Dockerfile                 — Cloud Run container spec
```
