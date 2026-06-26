"""BCA Protein Assay Calculator — Streamlit web front-end.

All calculation logic lives in bca_calculator.py (CLI tool).
This module is the web UI only — it imports and calls those validated functions.
The CLI remains fully functional and independent.
"""

import hashlib
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

# ── Fixed plate layout ────────────────────────────────────────────────────────
# Standards always occupy rows A–C, columns 1–8 (triplicates).
# Blank (0 µg/mL) always occupies rows A–C, column 9.
# These assignments are pre-set and non-editable by the user.

PIERCE_CONCS: dict[int, float] = {
    1: 2000, 2: 1500, 3: 1000, 4: 750,
    5: 500,  6: 250,  7: 125,  8: 25,
}

FIXED_ASSIGNMENTS: dict[str, str] = {}
for _col, _conc in PIERCE_CONCS.items():
    for _r in ("A", "B", "C"):
        FIXED_ASSIGNMENTS[f"{_r}{_col}"] = f"std_{_col}"
for _r in ("A", "B", "C"):
    FIXED_ASSIGNMENTS[f"{_r}9"] = "blank"

# ── Color palettes ────────────────────────────────────────────────────────────

BLANK_COLOR      = "#1565C0"
UNASSIGNED_COLOR = "#2d2d4e"

# Colorblind-safe palette (Wong 2011) — one per standard concentration (8 total)
STD_COLORS: dict[int, str] = {
    1: "#E69F00",
    2: "#56B4E9",
    3: "#009E73",
    4: "#F0E442",
    5: "#0072B2",
    6: "#D55E00",
    7: "#CC79A7",
    8: "#44AA99",
}

# 20 unknowns — visually distinct from standard palette
UNK_COLORS: list[str] = [
    "#FF6B6B", "#A8E6CF", "#FFD93D", "#6C5CE7", "#FD79A8",
    "#FDCB6E", "#74B9FF", "#00CEC9", "#E17055", "#81ECEC",
    "#636E72", "#B2BEC3", "#00B894", "#6D4C41", "#F48FB1",
    "#80CBC4", "#FFB74D", "#CE93D8", "#EF9A9A", "#90CAF9",
]


def _role_color(role: str) -> str:
    if not role:
        return UNASSIGNED_COLOR
    if role == "blank":
        return BLANK_COLOR
    if role.startswith("std_"):
        n = int(role.split("_")[1])
        return STD_COLORS.get(n, "#888888")
    if role.startswith("unk_"):
        n = int(role.split("_")[1])
        return UNK_COLORS[(n - 1) % len(UNK_COLORS)]
    return UNASSIGNED_COLOR


# ── Session state initialisation ──────────────────────────────────────────────

def _init_state() -> None:
    if "well_roles" not in st.session_state:
        # Standards and blank are fixed; user only assigns unknowns
        st.session_state.well_roles = dict(FIXED_ASSIGNMENTS)
    if "unk_count" not in st.session_state:
        st.session_state.unk_count = 1
        st.session_state["unk_name_1"]   = ""
        st.session_state["unk_target_1"] = 10.0
    if "active_role" not in st.session_state:
        st.session_state.active_role = "unk_1"
    if "last_file_hash" not in st.session_state:
        st.session_state.last_file_hash = ""

_init_state()


# ── Callbacks ─────────────────────────────────────────────────────────────────

def _pair_well(well: str) -> str | None:
    """Return the well one row below in the same column, or None if out of bounds / fixed."""
    col_str = well[1:]
    row_idx = PLATE_ROWS.index(well[0])
    if row_idx + 1 >= len(PLATE_ROWS):
        return None
    candidate = f"{PLATE_ROWS[row_idx + 1]}{col_str}"
    return None if candidate in FIXED_ASSIGNMENTS else candidate


def _is_lower_duplicate(well: str) -> bool:
    """True when this well was auto-assigned as the lower half of a duplicate pair."""
    col_str = well[1:]
    row_idx = PLATE_ROWS.index(well[0])
    if row_idx == 0:
        return False
    well_above = f"{PLATE_ROWS[row_idx - 1]}{col_str}"
    if well_above in FIXED_ASSIGNMENTS:
        return False
    role_above = st.session_state.well_roles.get(well_above, "")
    role_here  = st.session_state.well_roles.get(well, "")
    return bool(role_here and role_here.startswith("unk_") and role_here == role_above)


def _assign_well(well: str) -> None:
    if well in FIXED_ASSIGNMENTS:
        return
    # Lower-duplicate wells are controlled only by clicking the primary (upper) well
    if _is_lower_duplicate(well):
        return

    active   = st.session_state.get("active_role", "unk_1")
    current  = st.session_state.well_roles.get(well, "")
    pair     = _pair_well(well)

    if current == active:
        # Unassign primary and its auto-paired duplicate
        st.session_state.well_roles.pop(well, None)
        if pair and st.session_state.well_roles.get(pair) == active:
            st.session_state.well_roles.pop(pair, None)
    else:
        # Assign primary and automatically assign the row below as its duplicate
        st.session_state.well_roles[well] = active
        if pair:
            st.session_state.well_roles[pair] = active


def _add_unknown() -> None:
    n = st.session_state.unk_count + 1
    if n <= 20:
        st.session_state.unk_count = n
        if f"unk_name_{n}" not in st.session_state:
            st.session_state[f"unk_name_{n}"]   = ""
        if f"unk_target_{n}" not in st.session_state:
            st.session_state[f"unk_target_{n}"] = 10.0
        st.session_state.active_role = f"unk_{n}"


# ── File / calc helpers ───────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def _load_well_data(file_bytes: bytes, filename: str) -> dict:
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


def _style_flags(df: pd.DataFrame, col: str):
    def _color(val):
        v = str(val)
        if v == "OK":
            return "background-color:#d4edda;color:#155724"
        if "ABOVE" in v or "BELOW" in v or "RANGE" in v:
            return "background-color:#fff3cd;color:#856404"
        if v and v not in ("—", "nan"):
            return "background-color:#f8d7da;color:#721c24"
        return ""
    try:
        return df.style.map(_color, subset=[col])
    except AttributeError:
        return df.style.applymap(_color, subset=[col])


# ── Role label helper ─────────────────────────────────────────────────────────

def _role_label(r: str) -> str:
    if r == "blank":
        return "Blank  (0 µg/mL)"
    if r.startswith("std_"):
        n    = int(r.split("_")[1])
        conc = PIERCE_CONCS.get(n, 0)
        return f"Standard {n}  ({int(conc)} µg/mL)"
    n    = int(r.split("_")[1])
    name = (st.session_state.get(f"unk_name_{n}", "") or f"Sample_{n}").strip() or f"Sample_{n}"
    return f"Unknown {n}:  {name}"


# ── Plate CSS ─────────────────────────────────────────────────────────────────

def _plate_css() -> str:
    """
    Inject global + per-well CSS.

    Global rule targets any button preceded by a .wm-* marker div and makes
    it circular. Per-well rules set background color; fixed wells get
    pointer-events:none so they can't be clicked (and no hover scale).
    font-size is 6px with white-space:nowrap so 3-char labels (e.g. A10)
    always fit cleanly inside the 44 px circle.
    """
    active = st.session_state.get("active_role", "")

    global_css = """<style>
/* ── Dark background for every plate row ─────────────────────────────── */
div[data-testid="stHorizontalBlock"]:has([class^="wm-"]) {
    background: #16213e !important;
    padding: 3px 6px !important;
    margin: 1px 0 !important;
}
div[data-testid="stHorizontalBlock"]:has(.plate-col-hdr) {
    background: #0f3460 !important;
    border-radius: 8px 8px 0 0 !important;
    padding: 4px 6px 2px !important;
    margin-bottom: 0 !important;
}
div[data-testid="stHorizontalBlock"]:has(.plate-last-row) {
    border-radius: 0 0 8px 8px !important;
}

/* ── Circular well buttons ───────────────────────────────────────────── */
div[data-testid="element-container"]:has([class^="wm-"])
+ div[data-testid="element-container"] > div[data-testid="stButton"] > button {
    border-radius: 50% !important;
    width: 44px !important;
    height: 44px !important;
    min-height: 44px !important;
    padding: 0 !important;
    font-size: 6px !important;
    font-weight: 700 !important;
    white-space: nowrap !important;
    overflow: hidden !important;
    color: rgba(255,255,255,0.9) !important;
    line-height: 1 !important;
    margin: 0 auto !important;
    display: flex !important;
    align-items: center !important;
    justify-content: center !important;
    border: 2px solid rgba(255,255,255,0.12) !important;
    transition: transform 0.05s ease, filter 0.05s ease !important;
    cursor: pointer !important;
}
div[data-testid="element-container"]:has([class^="wm-"])
+ div[data-testid="element-container"] > div[data-testid="stButton"] > button:hover {
    transform: scale(1.2) !important;
    filter: brightness(1.3) !important;
    z-index: 10 !important;
    position: relative !important;
}
</style>"""

    color_rules = ["<style>"]
    for row in PLATE_ROWS:
        for col_num in range(1, 13):
            well          = f"{row}{col_num}"
            role          = st.session_state.well_roles.get(well, "")
            color         = _role_color(role)
            is_fixed      = well in FIXED_ASSIGNMENTS
            is_auto_pair  = (not is_fixed) and _is_lower_duplicate(well)
            is_locked     = is_fixed or is_auto_pair

            # White ring only for primary (non-locked) wells matching the active role
            ring = (
                "box-shadow: 0 0 0 2px white, 0 0 0 4px rgba(255,255,255,0.35) !important;"
                "transform: scale(1.06) !important;"
                if (role and role == active and not is_locked)
                else ""
            )
            lock = "pointer-events: none !important; cursor: default !important;" if is_locked else ""

            color_rules.append(
                f"div[data-testid='element-container']:has(.wm-{well})"
                f" + div[data-testid='element-container']"
                f" > div[data-testid='stButton'] > button"
                f" {{ background-color: {color} !important; {ring} {lock} }}"
            )
            if is_locked:
                color_rules.append(
                    f"div[data-testid='element-container']:has(.wm-{well})"
                    f" + div[data-testid='element-container']"
                    f" > div[data-testid='stButton'] > button:hover"
                    f" {{ transform: none !important; filter: none !important; }}"
                )
    color_rules.append("</style>")
    return global_css + "\n" + "\n".join(color_rules)


# ══════════════════════════════════════════════════════════════════════════════
# Main UI
# ══════════════════════════════════════════════════════════════════════════════

st.title("🧬 BCA Protein Assay Calculator")
st.caption(
    "Upload your plate reader file, click sample wells on the plate to assign them, "
    "then click **Calculate**. Standards (A–C, cols 1–8) and blank (A–C, col 9) "
    "are pre-assigned automatically."
)

# ── Sidebar ───────────────────────────────────────────────────────────────────

with st.sidebar:
    st.header("⚙️ Settings")
    cv_threshold = st.number_input(
        "CV flag threshold (%)", min_value=1.0, max_value=50.0,
        value=15.0, step=1.0,
        help="Replicates with CV above this are flagged HIGH CV.",
    )
    st.divider()
    st.subheader("Sample prep protocol")
    lysate_parts = st.number_input(
        "Lysate parts (lysate : 1 buffer)",
        min_value=1, max_value=50, value=5, step=1,
        help="5 = 5:1 lysate-to-sample-buffer ratio.",
    )
    buffer_parts    = 1
    dilution_factor = (lysate_parts + buffer_parts) / lysate_parts
    st.caption(f"Correction factor: ÷{dilution_factor:.4f}   ({lysate_parts}:{buffer_parts})")
    st.divider()
    with st.expander("ℹ️ About the math"):
        st.markdown(
            "**Curve (FORECAST):** `Conc = slope × Adj.Abs + intercept`  \n"
            "**Blank:** avg of 0 µg/mL BSA standard wells  \n"
            "**Dilution correction:** `final = CIW ÷ correction_factor`  \n"
            "**WB loading:** `lysate = target / (neat_conc / 1000)`  \n\n"
            "All numeric values are stored at full float64 precision. "
            "Excel cells use 9 decimal places."
        )
    st.divider()
    st.subheader("Fixed standard layout")
    st.caption(
        "Pierce BCA kit defaults — pre-assigned, not editable in the app. "
        "Edit `PIERCE_CONCS` in `streamlit_app.py` if your kit differs."
    )
    for col_idx, conc in PIERCE_CONCS.items():
        c = STD_COLORS[col_idx]
        st.markdown(
            f'<div style="display:flex;align-items:center;gap:6px;font-size:12px;margin:2px 0">'
            f'<span style="width:10px;height:10px;border-radius:50%;background:{c};'
            f'flex-shrink:0"></span>'
            f'Col {col_idx} (A{col_idx}/B{col_idx}/C{col_idx}) — {int(conc)} µg/mL</div>',
            unsafe_allow_html=True,
        )
    st.markdown(
        f'<div style="display:flex;align-items:center;gap:6px;font-size:12px;margin:2px 0">'
        f'<span style="width:10px;height:10px;border-radius:50%;background:{BLANK_COLOR};'
        f'flex-shrink:0"></span>'
        f'Col 9 (A9/B9/C9) — 0 µg/mL (blank)</div>',
        unsafe_allow_html=True,
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
    fhash = hashlib.md5(uploaded.getvalue()).hexdigest()[:12]
    if st.session_state.last_file_hash != fhash:
        st.session_state.last_file_hash = fhash
        # Reset only user-assigned wells; keep the fixed standard/blank assignments
        st.session_state.well_roles = dict(FIXED_ASSIGNMENTS)
        st.session_state.pop("bca_results", None)

    with st.spinner("Reading plate file…"):
        try:
            well_data = _load_well_data(uploaded.getvalue(), uploaded.name)
            st.success(f"Loaded **{len(well_data)}** wells from **{uploaded.name}**")
        except Exception as exc:
            st.error(f"Could not read file: {exc}")

    if well_data:
        with st.expander("📋 Plate layout (absorbance values)", expanded=False):
            st.dataframe(_plate_grid_df(well_data), use_container_width=True)

# ── Step 2: Well assignment ───────────────────────────────────────────────────

if well_data is not None:
    st.header("2 · Assign sample wells")
    st.info(
        "**Samples run in duplicate.** Click wells in **one row only** — "
        "each click assigns that well and the one directly below as a duplicate pair. "
        "Click an assigned well again to unassign both.",
        icon="ℹ️",
    )

    n_unk        = st.session_state.unk_count
    role_options = [f"unk_{i}" for i in range(1, n_unk + 1)]

    # Guard: reset to first unknown if active role no longer exists
    if st.session_state.get("active_role") not in role_options:
        st.session_state.active_role = "unk_1"

    left_col, right_col = st.columns([1, 2.9])

    # ── Left: unknown picker + settings ──────────────────────────────────────
    with left_col:
        st.subheader("Active sample")
        st.caption(
            "Select a sample, click wells in its **first** row. "
            "The row below is auto-assigned as the duplicate. "
            "Standards and blank wells are locked."
        )

        active_role = st.selectbox(
            "Sample",
            options=role_options,
            format_func=_role_label,
            key="active_role",
            label_visibility="collapsed",
        )

        st.markdown("---")
        n      = int(active_role.split("_")[1])
        n_wells = sum(1 for r in st.session_state.well_roles.values() if r == active_role)
        st.text_input(f"Sample {n} name (optional)", key=f"unk_name_{n}")
        st.number_input(
            "Target µg / lane",
            key=f"unk_target_{n}",
            min_value=0.1,
            step=1.0,
        )
        color_badge = UNK_COLORS[(n - 1) % len(UNK_COLORS)]
        st.markdown(
            f'<span style="display:inline-flex;align-items:center;gap:6px;font-size:12px">'
            f'<span style="width:12px;height:12px;border-radius:50%;background:{color_badge};'
            f'border:1px solid rgba(0,0,0,0.2)"></span>'
            f'{"✓ " + str(n_wells) + " well" + ("s" if n_wells != 1 else "") + " assigned" if n_wells else "No wells assigned yet"}'
            f'</span>',
            unsafe_allow_html=True,
        )

        st.markdown("---")
        st.button(
            "+ Add sample",
            on_click=_add_unknown,
            disabled=(n_unk >= 20),
            use_container_width=True,
        )

        # Assignment summary — unknowns only (fixed wells shown as one summary line)
        st.markdown("**Well assignments**")
        fixed_count = sum(1 for w in FIXED_ASSIGNMENTS if w in well_data)
        st.markdown(
            f'<div style="font-size:12px;color:#888;margin:2px 0">'
            f'Standards + blank: {fixed_count} wells (fixed)</div>',
            unsafe_allow_html=True,
        )
        any_unk = False
        for role_key in role_options:
            cnt = sum(1 for r in st.session_state.well_roles.values() if r == role_key)
            if not cnt:
                continue
            any_unk = True
            c   = _role_color(role_key)
            lbl = _role_label(role_key)
            st.markdown(
                f'<div style="display:flex;align-items:center;gap:7px;'
                f'margin:2px 0;font-size:12px">'
                f'<span style="width:11px;height:11px;border-radius:50%;'
                f'background:{c};flex-shrink:0;border:1px solid rgba(0,0,0,0.15)"></span>'
                f'{lbl} — {cnt} well{"s" if cnt > 1 else ""}</div>',
                unsafe_allow_html=True,
            )
        if not any_unk:
            st.caption("No sample wells assigned yet — click wells on the plate.")

    # ── Right: 96-well plate ──────────────────────────────────────────────────
    with right_col:
        st.subheader("96-well plate")
        st.caption(
            f"Active: **{_role_label(active_role)}** · "
            "Click one row of wells — the row below is auto-assigned as the duplicate. "
            "Standards and blank wells are locked."
        )

        # Inject all CSS before any buttons render
        st.markdown(_plate_css(), unsafe_allow_html=True)

        # Column-number header row
        hdr = st.columns([0.6] + [1] * 12)
        with hdr[0]:
            st.markdown('<span class="plate-col-hdr"></span>', unsafe_allow_html=True)
        for j, col_num in enumerate(range(1, 13)):
            with hdr[j + 1]:
                st.markdown(
                    f'<p class="plate-col-hdr" style="text-align:center;font-size:9px;'
                    f'color:#8899bb;margin:0;padding:2px 0">{col_num}</p>',
                    unsafe_allow_html=True,
                )

        # Plate rows A–H
        for row in PLATE_ROWS:
            row_cols = st.columns([0.6] + [1] * 12)
            with row_cols[0]:
                extra = 'class="plate-last-row"' if row == "H" else ""
                st.markdown(
                    f'<p {extra} style="font-size:12px;color:#8899bb;text-align:right;'
                    f'padding:0 6px 0 0;margin:0;line-height:44px">{row}</p>',
                    unsafe_allow_html=True,
                )
            for j, col_num in enumerate(range(1, 13)):
                well = f"{row}{col_num}"
                with row_cols[j + 1]:
                    # Marker div: CSS :has(.wm-{well}) anchor for per-well styling
                    st.markdown(f'<div class="wm-{well}"></div>', unsafe_allow_html=True)
                    st.button(
                        well,                        # label is always "A1"…"H12"
                        key=f"plate_{well}",
                        on_click=_assign_well,
                        args=(well,),
                    )

        # Colour legend
        legend_items = [
            ("Unassigned", UNASSIGNED_COLOR),
            ("Blank (col 9)", BLANK_COLOR),
        ]
        for col_idx, conc in PIERCE_CONCS.items():
            legend_items.append((f"Std {col_idx} ({int(conc)})", STD_COLORS[col_idx]))
        for i in range(1, n_unk + 1):
            name = (st.session_state.get(f"unk_name_{i}", "") or f"Sample {i}")[:14]
            legend_items.append((name, UNK_COLORS[(i - 1) % len(UNK_COLORS)]))

        badges = "".join(
            f'<span style="display:inline-flex;align-items:center;margin:3px 8px 3px 0;'
            f'font-size:11px;color:#555">'
            f'<span style="width:13px;height:13px;border-radius:50%;background:{c};'
            f'display:inline-block;margin-right:5px;border:1px solid rgba(0,0,0,0.12)"></span>'
            f'{lbl}</span>'
            for lbl, c in legend_items
        )
        st.markdown(
            f'<div style="margin-top:10px;line-height:2;padding:6px 2px">{badges}</div>',
            unsafe_allow_html=True,
        )

    # ── Step 3: Calculate ─────────────────────────────────────────────────────
    st.markdown("---")
    st.header("3 · Calculate")
    calc_btn = st.button("▶  Calculate", type="primary")

    if calc_btn:
        errors: list[str] = []

        # Blank and standard wells are always the fixed layout
        blank_wells = [w for w in ("A9", "B9", "C9") if w in well_data]
        if not blank_wells:
            errors.append(
                "Blank wells (A9, B9, C9) not found in plate data. "
                "Check that the file contains data for those positions."
            )

        standard_defs: list[dict] = []
        for col_idx, conc in PIERCE_CONCS.items():
            wells = [w for w in (f"A{col_idx}", f"B{col_idx}", f"C{col_idx}") if w in well_data]
            if wells:
                standard_defs.append({"wells": wells, "concentration": float(conc)})
        if len(standard_defs) < 2:
            errors.append(
                "Fewer than 2 standard concentrations found in the plate data. "
                "Verify the file contains absorbance values for columns 1–8, rows A–C."
            )

        unknown_defs: list[dict]   = []
        target_map:   dict[str, float] = {}
        for i in range(1, st.session_state.unk_count + 1):
            wells = sorted(
                w for w, r in st.session_state.well_roles.items()
                if r == f"unk_{i}" and w in well_data
            )
            if not wells:
                continue
            name      = (st.session_state.get(f"unk_name_{i}", "") or f"Sample_{i}").strip() or f"Sample_{i}"
            target_ug = float(st.session_state.get(f"unk_target_{i}", 10.0) or 10.0)
            unknown_defs.append({"wells": wells, "name": name, "dilution": dilution_factor})
            target_map[name] = target_ug
        if not unknown_defs:
            errors.append("No sample wells assigned. Click wells on the plate to assign samples.")

        if errors:
            for msg in errors:
                st.error(msg)
        else:
            with st.spinner("Running BCA calculations…"):
                try:
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

                    # WB loading — one call per sample so each uses its own target µg
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
                        "target_ug":    10.0,
                        "max_well_vol": None,
                        "lysate_parts": lysate_parts,
                        "buffer_parts": buffer_parts,
                    }

                    plot_bytes  = _make_plot_bytes(standards_df, slope, intercept, r_squared)
                    excel_bytes = _make_excel_bytes(
                        blank_avg, blank_cv, blank_wells,
                        standards_df, unknowns_results,
                        slope, intercept, r_squared,
                        cv_threshold, uploaded.name,
                        western_df, western_params,
                    )

                    st.session_state["bca_results"] = {
                        "blank_avg":        blank_avg,
                        "blank_cv":         blank_cv,
                        "standards_df":     standards_df,
                        "unknowns_results": unknowns_results,
                        "western_df":       western_df,
                        "slope":            slope,
                        "intercept":        intercept,
                        "r_squared":        r_squared,
                        "plot_bytes":       plot_bytes,
                        "excel_bytes":      excel_bytes,
                        "base_name":        os.path.splitext(uploaded.name)[0],
                    }

                except Exception as exc:
                    import traceback
                    st.error(f"Calculation failed: {exc}")
                    with st.expander("Details"):
                        st.code(traceback.format_exc())

# ── Step 4: Results ───────────────────────────────────────────────────────────

if "bca_results" in st.session_state:
    res       = st.session_state["bca_results"]
    slope     = res["slope"]
    intercept = res["intercept"]
    r_squared = res["r_squared"]

    st.header("4 · Results")

    sign_str = "+" if intercept >= 0 else "−"
    c1, c2, c3 = st.columns(3)
    c1.metric(
        "R²", f"{r_squared:.5f}",
        delta="⚠ below 0.99" if r_squared < 0.99 else "✓ good fit",
        delta_color="inverse" if r_squared < 0.99 else "normal",
    )
    c2.metric("Slope",     f"{slope:.4f}")
    c3.metric("Intercept", f"{intercept:.4f}")
    st.caption(
        f"Equation: **Conc = {slope:.4f} × Adj.Abs {sign_str} {abs(intercept):.4f}**  "
        f"| blank mean **{res['blank_avg']:.5f}**  CV **{res['blank_cv']:.2f}%**"
    )
    if r_squared < 0.99:
        st.warning("⚠️ R² < 0.99 — check for outlier standards or mis-assigned concentrations.")

    st.subheader("Standard curve")
    st.image(res["plot_bytes"])

    st.subheader("Sample concentrations")
    conc_df = res["unknowns_results"][[
        "sample_name", "wells", "n_replicates",
        "blank_corrected_abs", "final_concentration_ug_mL", "cv_pct", "qc_flags",
    ]].copy()
    conc_df.columns = ["Sample", "Wells", "N", "Adj. Abs", "Final Conc (µg/mL)", "CV%", "Flags"]
    conc_df["Final Conc (µg/mL)"] = conc_df["Final Conc (µg/mL)"].round(4)
    conc_df["Adj. Abs"]           = conc_df["Adj. Abs"].round(6)
    conc_df["CV%"]                = conc_df["CV%"].round(2)
    st.dataframe(_style_flags(conc_df, "Flags"), use_container_width=True, hide_index=True)

    st.subheader("Western blot loading volumes")
    wb_df = res["western_df"][[
        "sample_name", "neat_conc_ug_mL", "conc_in_sample_buffer_ug_mL",
        "target_ug", "lysate_vol_uL", "sample_buffer_vol_uL", "total_load_vol_uL", "wb_flags",
    ]].copy()
    wb_df.columns = [
        "Sample", "Neat Conc (µg/mL)", "Buf Conc (µg/mL)",
        "Target (µg)", "Lysate (µL)", "Sample Buffer (µL)", "Total (µL)", "Flags",
    ]
    for col in ["Neat Conc (µg/mL)", "Buf Conc (µg/mL)", "Lysate (µL)", "Sample Buffer (µL)", "Total (µL)"]:
        wb_df[col] = wb_df[col].round(4)
    st.dataframe(_style_flags(wb_df, "Flags"), use_container_width=True, hide_index=True)

    st.subheader("Downloads")
    base = res["base_name"]
    ca, cb = st.columns(2)
    with ca:
        st.download_button(
            "⬇️  Download Excel results",
            res["excel_bytes"],
            f"{base}_results.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with cb:
        st.download_button(
            "⬇️  Download standard curve PNG",
            res["plot_bytes"],
            f"{base}_standard_curve.png",
            "image/png",
            use_container_width=True,
        )
