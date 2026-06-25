"""BCA Protein Assay Calculator — Streamlit web front-end.

All calculation logic lives in bca_calculator.py (CLI tool).
This module is the web UI only — it imports and calls those validated functions.
The CLI remains fully functional and independent.
"""

import os
import tempfile

import numpy as np
import pandas as pd
import streamlit as st

# bca_calculator.py sets matplotlib.use("Agg") at module level,
# so importing it here locks in the non-interactive backend before any plt call.
from bca_calculator import (
    FIT_LINEAR,
    PLATE_ROWS,
    _validate_wells,
    calculate_concentrations,
    calculate_western_loading,
    compute_blank_corrected,
    export_to_excel,
    fit_linear_curve,
    generate_standard_curve_plot,
    load_plate_data,
)

# ── Page config ───────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="BCA Protein Assay Calculator",
    page_icon="🧬",
    layout="wide",
)

# ── Default Pierce BCA kit standard series ────────────────────────────────────

_PIERCE_STANDARDS = pd.DataFrame([
    {"Wells": "A1 B1 C1", "Concentration (µg/mL)": 2000},
    {"Wells": "A2 B2 C2", "Concentration (µg/mL)": 1500},
    {"Wells": "A3 B3 C3", "Concentration (µg/mL)": 1000},
    {"Wells": "A4 B4 C4", "Concentration (µg/mL)":  750},
    {"Wells": "A5 B5 C5", "Concentration (µg/mL)":  500},
    {"Wells": "A6 B6 C6", "Concentration (µg/mL)":  250},
    {"Wells": "A7 B7 C7", "Concentration (µg/mL)":  125},
    {"Wells": "A8 B8 C8", "Concentration (µg/mL)":   25},
])

_EMPTY_UNKNOWNS = pd.DataFrame(
    {"Wells": [""], "Sample name": [""], "Target µg": [10.0]}
)

# ── Helper: load uploaded file via temp path ──────────────────────────────────

@st.cache_data(show_spinner=False)
def _load_well_data(file_bytes: bytes, filename: str) -> dict:
    """Save upload to a temp file, call load_plate_data, clean up."""
    suffix = os.path.splitext(filename)[1]
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    try:
        return load_plate_data(tmp_path)
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


def _plate_grid_df(well_data: dict) -> pd.DataFrame:
    """Build a display DataFrame for the 96-well plate (rows A–H, cols 1–12)."""
    data = {}
    for row in PLATE_ROWS:
        data[row] = {
            col: (f"{well_data[f'{row}{col}']:.4f}" if f"{row}{col}" in well_data else "—")
            for col in range(1, 13)
        }
    return pd.DataFrame(data, index=range(1, 13)).T


def _make_plot_bytes(standards_df, slope, intercept, r_squared) -> bytes:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
        tmp_path = tmp.name
    try:
        generate_standard_curve_plot(
            standards_df, FIT_LINEAR, slope, intercept, r_squared, tmp_path
        )
        with open(tmp_path, "rb") as f:
            return f.read()
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


def _make_excel_bytes(
    blank_avg, blank_cv, blank_wells, standards_df, unknowns_results,
    slope, intercept, r_squared, cv_threshold, input_filename,
    western_df, western_params,
) -> bytes:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp_path = tmp.name
    try:
        export_to_excel(
            tmp_path, blank_avg, blank_cv, blank_wells,
            standards_df, unknowns_results,
            FIT_LINEAR, slope, intercept, r_squared, cv_threshold, input_filename,
            western_df=western_df, western_params=western_params,
        )
        with open(tmp_path, "rb") as f:
            return f.read()
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


def _style_flag_col(df: pd.DataFrame, col: str):
    """Return a pandas Styler with colour-coded flag column."""
    def _color(val):
        v = str(val)
        if v == "OK":
            return "background-color: #d4edda; color: #155724"
        if "ABOVE" in v or "BELOW" in v or "RANGE" in v:
            return "background-color: #fff3cd; color: #856404"
        if v and v != "—":
            return "background-color: #f8d7da; color: #721c24"
        return ""

    try:
        return df.style.map(_color, subset=[col])
    except AttributeError:
        return df.style.applymap(_color, subset=[col])


# ══════════════════════════════════════════════════════════════════════════════
# UI layout
# ══════════════════════════════════════════════════════════════════════════════

st.title("🧬 BCA Protein Assay Calculator")
st.caption(
    "Upload your plate reader file, fill in well assignments, and click **Calculate**. "
    "Downloads match the CLI output exactly — all calculations run from the same validated code."
)

# ── Sidebar: global settings ──────────────────────────────────────────────────

with st.sidebar:
    st.header("⚙️ Settings")
    cv_threshold = st.number_input(
        "CV flag threshold (%)",
        min_value=1.0, max_value=50.0, value=15.0, step=1.0,
        help="Replicates with CV above this value are flagged HIGH CV.",
    )
    st.divider()
    st.subheader("Sample prep protocol")
    lysate_parts = st.number_input(
        "Lysate parts (lysate : 1 buffer)",
        min_value=1, max_value=50, value=5, step=1,
        help="5 means a 5:1 lysate-to-sample-buffer ratio.",
    )
    buffer_parts = 1
    dilution_factor = (lysate_parts + buffer_parts) / lysate_parts
    st.caption(f"Correction factor: ÷ {dilution_factor:.4f}   ({lysate_parts}:{buffer_parts} ratio)")
    st.divider()
    with st.expander("ℹ️ About the math"):
        st.markdown(
            "**Standard curve** follows Excel FORECAST direction:  \n"
            "`Conc = slope × Adj.Abs + intercept`\n\n"
            "**Blank correction:** average of 0 µg/mL BSA standard wells "
            "(the last point in the Pierce kit).  \n\n"
            "**Dilution correction:** `final_conc = CIW ÷ correction_factor`  \n\n"
            "**WB loading:**  \n"
            "`lysate = target_µg / (neat_conc / 1000)`  \n"
            "`sample_buffer = lysate / lysate_parts`"
        )

# ── Step 1: File upload ───────────────────────────────────────────────────────

st.header("1 · Upload plate file")
uploaded = st.file_uploader(
    "SoftMax Pro export (.txt, .xls), Excel (.xlsx), or CSV (.csv)",
    type=["txt", "xls", "xlsx", "csv"],
    label_visibility="collapsed",
)

well_data = None
if uploaded is not None:
    with st.spinner("Reading plate file…"):
        try:
            well_data = _load_well_data(uploaded.getvalue(), uploaded.name)
            st.success(f"Loaded **{len(well_data)}** wells from **{uploaded.name}**")
        except Exception as exc:
            st.error(f"Could not read file: {exc}")
            well_data = None

    if well_data:
        with st.expander("📋 Plate layout (absorbance values)", expanded=True):
            grid = _plate_grid_df(well_data)
            st.dataframe(grid, use_container_width=True)

# ── Step 2: Well assignment ───────────────────────────────────────────────────

if well_data is not None:
    st.header("2 · Well assignment")

    # Blank wells
    blank_raw = st.text_input(
        "0 µg/mL standard wells (blank reference)",
        value="A9 B9 C9",
        help=(
            "Wells with your 0 µg/mL BSA standard (last point in Pierce kit). "
            "Formats: 'A9 B9 C9', 'A9:C9', 'A9,B9,C9'."
        ),
    )

    st.markdown("---")

    # Standards table
    st.subheader("BSA standards")
    st.caption(
        "Edit wells or concentrations. Add / remove rows with the ＋ / ✕ controls. "
        "Pre-filled with Pierce BCA kit defaults."
    )
    std_table = st.data_editor(
        _PIERCE_STANDARDS,
        num_rows="dynamic",
        use_container_width=True,
        key="std_editor",
        column_config={
            "Wells": st.column_config.TextColumn(
                "Wells", help="Space-, comma-, or semicolon-separated well IDs, e.g. A1 B1 C1"
            ),
            "Concentration (µg/mL)": st.column_config.NumberColumn(
                "Conc (µg/mL)", min_value=0, step=25
            ),
        },
    )

    st.markdown("---")

    # Unknowns table
    st.subheader("Unknown samples")
    st.caption(
        "Leave **Sample name** blank to auto-name (Sample_1, Sample_2, …).  "
        "**Target µg** can differ per row for Western blot loading volumes."
    )
    unk_table = st.data_editor(
        _EMPTY_UNKNOWNS,
        num_rows="dynamic",
        use_container_width=True,
        key="unk_editor",
        column_config={
            "Wells": st.column_config.TextColumn(
                "Wells", help="Well IDs for this sample, e.g. D1 E1"
            ),
            "Sample name": st.column_config.TextColumn("Sample name (optional)"),
            "Target µg": st.column_config.NumberColumn(
                "Target µg", min_value=0.1, default=10.0, step=1.0
            ),
        },
    )

    # ── Step 3: Calculate ─────────────────────────────────────────────────────

    st.markdown("---")
    st.header("3 · Calculate")
    calc_btn = st.button("▶  Calculate", type="primary")

    if calc_btn:
        errors: list[str] = []

        # Validate blank wells
        blank_wells, bad_blank = _validate_wells(blank_raw, well_data)
        if bad_blank:
            errors.append(f"Blank wells not found in plate data: {', '.join(bad_blank)}")
        if not blank_wells:
            errors.append("No valid blank wells entered.")

        # Parse standards
        standard_defs: list[dict] = []
        for _, row in std_table.dropna(subset=["Wells"]).iterrows():
            raw_w = str(row["Wells"]).strip()
            if not raw_w:
                continue
            conc = row.get("Concentration (µg/mL)")
            if conc is None or (isinstance(conc, float) and np.isnan(conc)):
                errors.append(f"Standard wells '{raw_w}' has no concentration.")
                continue
            valid_w, bad_w = _validate_wells(raw_w, well_data)
            if bad_w:
                errors.append(f"Standard wells not found in plate: {', '.join(bad_w)}")
            if valid_w:
                standard_defs.append({"wells": valid_w, "concentration": float(conc)})

        if len(standard_defs) < 2:
            errors.append("At least 2 standard concentrations are required for linear regression.")

        # Parse unknowns
        unknown_defs: list[dict] = []
        sample_counter = 1
        target_map: dict[str, float] = {}
        for _, row in unk_table.dropna(subset=["Wells"]).iterrows():
            raw_w = str(row["Wells"]).strip()
            if not raw_w:
                continue
            valid_w, bad_w = _validate_wells(raw_w, well_data)
            if bad_w:
                errors.append(f"Unknown wells not found in plate: {', '.join(bad_w)}")
            if valid_w:
                name = str(row.get("Sample name", "") or "").strip()
                if not name:
                    name = f"Sample_{sample_counter}"
                target_ug = float(row.get("Target µg") or 10.0)
                unknown_defs.append({
                    "wells":   valid_w,
                    "name":    name,
                    "dilution": dilution_factor,
                })
                target_map[name] = target_ug
                sample_counter += 1

        if not unknown_defs:
            errors.append("No valid unknown samples entered.")

        if errors:
            for msg in errors:
                st.error(msg)
        else:
            with st.spinner("Running BCA calculations…"):
                try:
                    # ── Core calculations (all in bca_calculator.py) ───────────
                    blank_avg, blank_cv, standards_df, unknowns_df = compute_blank_corrected(
                        well_data, blank_wells, standard_defs, unknown_defs
                    )

                    slope, intercept, r_squared = fit_linear_curve(standards_df)

                    std_min = standards_df["concentration_ug_mL"].min()
                    std_max = standards_df["concentration_ug_mL"].max()

                    unknowns_results = calculate_concentrations(
                        unknowns_df, FIT_LINEAR, standards_df,
                        slope, intercept, std_min, std_max, cv_threshold,
                    )

                    # ── Western blot loading (per-sample target µg) ────────────
                    # Call calculate_western_loading once per sample so each can
                    # use its own target_ug without changing the function signature.
                    wb_parts: list[pd.DataFrame] = []
                    for udef in unknown_defs:
                        sname = udef["name"]
                        sub   = unknowns_results[unknowns_results["sample_name"] == sname]
                        params = {
                            "target_ug":    target_map.get(sname, 10.0),
                            "max_well_vol": None,
                            "lysate_parts": lysate_parts,
                            "buffer_parts": buffer_parts,
                        }
                        wb_parts.append(calculate_western_loading(sub, params))

                    western_df = pd.concat(wb_parts).reset_index(drop=True)

                    western_params = {
                        "target_ug":    10.0,   # Excel header note only; actual targets above
                        "max_well_vol": None,
                        "lysate_parts": lysate_parts,
                        "buffer_parts": buffer_parts,
                    }

                    # ── Export ─────────────────────────────────────────────────
                    plot_bytes  = _make_plot_bytes(standards_df, slope, intercept, r_squared)
                    excel_bytes = _make_excel_bytes(
                        blank_avg, blank_cv, blank_wells,
                        standards_df, unknowns_results,
                        slope, intercept, r_squared,
                        cv_threshold, uploaded.name,
                        western_df, western_params,
                    )

                    # Store in session state so results survive widget interactions.
                    st.session_state["bca_results"] = {
                        "blank_avg":       blank_avg,
                        "blank_cv":        blank_cv,
                        "standards_df":    standards_df,
                        "unknowns_results": unknowns_results,
                        "western_df":      western_df,
                        "slope":           slope,
                        "intercept":       intercept,
                        "r_squared":       r_squared,
                        "plot_bytes":      plot_bytes,
                        "excel_bytes":     excel_bytes,
                        "base_name":       os.path.splitext(uploaded.name)[0],
                    }

                except Exception as exc:
                    import traceback
                    st.error(f"Calculation failed: {exc}")
                    with st.expander("Details"):
                        st.code(traceback.format_exc())

# ── Step 4: Results ───────────────────────────────────────────────────────────

if "bca_results" in st.session_state:
    res = st.session_state["bca_results"]
    slope      = res["slope"]
    intercept  = res["intercept"]
    r_squared  = res["r_squared"]
    blank_avg  = res["blank_avg"]
    blank_cv   = res["blank_cv"]

    st.header("4 · Results")

    # Curve summary metrics
    sign_str = "+" if intercept >= 0 else "−"
    col_r2, col_slope, col_int = st.columns(3)
    col_r2.metric(
        "R²",
        f"{r_squared:.5f}",
        delta="⚠ below 0.99" if r_squared < 0.99 else "✓ good fit",
        delta_color="inverse" if r_squared < 0.99 else "normal",
    )
    col_slope.metric("Slope", f"{slope:.4f}")
    col_int.metric("Intercept", f"{intercept:.4f}")

    st.caption(
        f"Equation: **Conc = {slope:.4f} × Adj.Abs {sign_str} {abs(intercept):.4f}**  "
        f"&nbsp;|&nbsp; "
        f"0 µg/mL blank: mean **{blank_avg:.5f}**  CV **{blank_cv:.2f}%**"
    )

    if r_squared < 0.99:
        st.warning("⚠️ R² < 0.99 — standard curve fit is poor. Check for outlier wells or mis-assigned concentrations.")

    # Standard curve plot
    st.subheader("Standard curve")
    st.image(res["plot_bytes"])

    # Concentrations table
    st.subheader("Sample concentrations")
    conc_df = res["unknowns_results"][[
        "sample_name", "wells", "n_replicates",
        "blank_corrected_abs", "final_concentration_ug_mL", "cv_pct", "qc_flags",
    ]].copy()
    conc_df.columns = [
        "Sample", "Wells", "N reps",
        "Adj. Abs", "Final Conc (µg/mL)", "CV%", "Flags",
    ]
    conc_df["Final Conc (µg/mL)"] = conc_df["Final Conc (µg/mL)"].round(3)
    conc_df["Adj. Abs"]           = conc_df["Adj. Abs"].round(5)
    conc_df["CV%"]                = conc_df["CV%"].round(2)
    st.dataframe(_style_flag_col(conc_df, "Flags"), use_container_width=True, hide_index=True)

    # Western blot loading table
    st.subheader("Western blot loading volumes")
    wb_df = res["western_df"][[
        "sample_name", "neat_conc_ug_mL", "conc_in_sample_buffer_ug_mL",
        "target_ug", "lysate_vol_uL", "sample_buffer_vol_uL", "total_load_vol_uL", "wb_flags",
    ]].copy()
    wb_df.columns = [
        "Sample", "Neat Conc (µg/mL)", "Buf Conc (µg/mL)",
        "Target (µg)", "Lysate (µL)", "Sample Buffer (µL)", "Total (µL)", "Flags",
    ]
    st.dataframe(_style_flag_col(wb_df, "Flags"), use_container_width=True, hide_index=True)

    # Downloads
    st.subheader("Downloads")
    base = res["base_name"]
    col_xl, col_png = st.columns(2)
    with col_xl:
        st.download_button(
            label="⬇️  Download Excel results",
            data=res["excel_bytes"],
            file_name=f"{base}_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col_png:
        st.download_button(
            label="⬇️  Download standard curve PNG",
            data=res["plot_bytes"],
            file_name=f"{base}_standard_curve.png",
            mime="image/png",
            use_container_width=True,
        )
