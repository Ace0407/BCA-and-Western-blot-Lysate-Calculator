# BCA Protein Assay Calculator

Reads Pierce BCA plate reader data, fits a linear standard curve, calculates unknown protein concentrations, generates Western blot loading volumes, and exports a formatted Excel workbook.

Available as both a **browser app** (no install required) and a **command-line tool**.

---

## For lab users — browser app

### Open the app

> **URL:** *(add your Streamlit Community Cloud URL here after deployment)*

No install needed. Works in Chrome, Firefox, or Safari on any computer.

### How to use

1. **Upload** your plate reader file — SoftMax Pro `.txt`, `.xls`, `.xlsx`, or `.csv` all work directly (no manual conversion needed).
2. **Confirm the plate grid** looks correct.
3. **Check the standards table** — pre-filled with Pierce BCA kit defaults (2000 → 25 µg/mL, triplicates A–C). Edit only if your layout differs.
4. **Enter your unknown sample wells** — e.g. `D1 E1` for a duplicate, or add multiple rows for multiple samples.
5. **Click Calculate**.
6. **Download** the Excel file or the standard curve PNG.

The sidebar lets you change the CV threshold (default 15%) and the lysate:buffer ratio (default 5:1).

---

## Command-line tool

### Setup

Python 3.10 or newer. Install dependencies:

```bash
pip install -r requirements.txt
```

### Run

```bash
cd /path/to/BCA_Program_Claude
python3 bca_calculator.py "/path/to/your_plate_file.xls"
```

The tool accepts `.txt` (SoftMax Pro), `.xls`, `.xlsx`, and `.csv` — no pre-conversion needed.

Optional flags:

| Flag | Default | Description |
|---|---|---|
| `--output`, `-o` | `<input>_results.xlsx` | Custom Excel output path |
| `--cv-threshold` | `15.0` | % CV above which replicates are flagged |

### What the CLI asks (in order)

1. **Run name** — a short label for the output folder
2. **Blank row question** — whether there is an empty row between your standards and samples on the plate (informational only, does not affect calculations)
3. **Sample wells** — enter as a block range, e.g. `D1:E5`
4. **Name samples?** — `y` to name each sample, `n` to auto-name (Sample_1, Sample_2, …)

Everything else is fixed:
- Blank wells: **A9, B9, C9**
- Standards: **A1:C8** → 2000, 1500, 1000, 750, 500, 250, 125, 25 µg/mL
- Mixing ratio: **5:1** (lysate:sample buffer)
- WB loading target: **10 µg/lane**

---

## Input file formats

### SoftMax Pro exports (.txt or .xls)

SoftMax Pro files are UTF-16 encoded tab-separated text. The tool detects and parses them automatically — pass the file directly with no conversion.

### Plate grid layout

The most common spreadsheet export. The first column is the row letter (A–H) and the remaining columns are numbered 1–12:

```
    1       2       3  …
A   0.0985  0.0978  0.0991
B   0.1373  0.1381  0.1366
…
```

### Tabular list

A two-column table with headers `Well` and `Absorbance` (case-insensitive).

---

## Block range notation

Enter wells as a rectangular block: `D1:F5` expands to all wells in rows D–F, columns 1–5.

Same-row ranges (`A1:A8`) and single wells (`D1`) are also accepted.

---

## How concentrations are calculated

1. **Blank average** — mean absorbance of the 0 µg/mL standard wells (A9, B9, C9).
2. **Blank correction** — subtract the blank average from every individual replicate, then average the corrected replicates per group.
3. **Standard curve** — linear regression following Excel FORECAST direction:
   `Conc = slope × Adj.Abs + intercept`
4. **Unknown concentration in well** — same formula applied to each unknown.
5. **Dilution correction** — `final_conc = CIW ÷ correction_factor`
   where `correction_factor = (lysate_parts + buffer_parts) / lysate_parts` (1.2 for the default 5:1 ratio).
6. **CV%** — `(std_dev / mean) × 100` on raw replicates (single replicates report 0%).

---

## Western blot loading volumes

Uses the neat (pre-dilution) concentration and the 5:1 lysate:sample buffer protocol:

```
lysate_vol       = target_µg / (neat_conc_µg_mL / 1000)
sample_buffer_vol = lysate_vol / lysate_parts
total_vol        = lysate_vol + sample_buffer_vol
```

For example, a sample at 2065 µg/mL targeting 10 µg:
- Lysate: 4.84 µL
- Sample buffer: 0.97 µL
- Total: 5.81 µL

---

## QC flags

| Flag | Meaning |
|---|---|
| `BELOW CURVE RANGE` | Concentration < lowest standard. Dilute less or add a lower standard. |
| `ABOVE CURVE RANGE` | Concentration > highest standard. Dilute more. |
| `NEAR UPPER CURVE LIMIT (>1500 µg/mL)` | Pierce manual notes the curve may deviate from linear above 1500 µg/mL. |
| `HIGH CV (xx.x%)` | Replicate variability exceeds the CV threshold. Check pipetting. |
| `OK` | All checks passed. |

---

## Excel output

Four sheets:

| Sheet | Contents |
|---|---|
| **Summary** | Input file name, blank stats, curve equation, R², protocol settings |
| **Standards** | Per-standard: concentration, raw absorbance, blank-corrected absorbance, CV% |
| **Unknowns** | Per-sample: wells, corrected absorbance, concentration, CV%, QC flags |
| **Western_Loading** | Per-sample: neat concentration, loading volumes, WB flags |

Flagged cells are highlighted in red; passing cells in green.

---

## Project structure

```
BCA_Program_Claude/
├── bca_calculator.py   # All calculation, parsing, and export logic (CLI entry point)
├── streamlit_app.py    # Browser front-end — imports from bca_calculator.py
├── requirements.txt    # Python dependencies
└── README.md           # This file
```

---

## Deploying the browser app (Streamlit Community Cloud)

1. Push this repository to GitHub.
2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with your GitHub account.
3. Click **New app**.
4. Select your repository and set:
   - **Branch:** `main` (or whichever branch you pushed to)
   - **Main file path:** `streamlit_app.py`
5. Click **Deploy**. Streamlit installs `requirements.txt` automatically.
6. Copy the URL and share it with your lab.

The app is free to host on Streamlit Community Cloud for public repositories.

---

## Running locally

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

The app opens at `http://localhost:8501`.

---

## Troubleshooting

**"Could not read file"** — Make sure the file is a real SoftMax Pro export or a standard plate grid Excel/CSV. Password-protected Excel files are not supported.

**"Standard wells not found in plate data"** — The wells you entered do not have absorbance data. Check the plate grid display after uploading.

**Low R²** — One or more standards may be outliers (pipetting error, evaporation, bubble in well). Check the standards table and plate grid for suspicious values.

**Wells showing `—` in the plate grid** — Those positions have no data in the file. Verify the plate was exported completely.
