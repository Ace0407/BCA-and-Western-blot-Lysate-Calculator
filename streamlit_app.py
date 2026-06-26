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

# ── Color palettes ────────────────────────────────────────────────────────────

PIERCE_DEFAULTS = {1: 2000, 2: 1500, 3: 1000, 4: 750, 5: 500, 6: 250, 7: 125, 8: 25}

BLANK_COLOR     = "#1565C0"   # rich blue
UNASSIGNED_COLOR = "#2d2d4e"  # dark navy-grey

# 9 standards: colorblind-safe palette (Wong 2011 + extensions)
STD_COLORS = {
    1: "#E69F00",
    2: "#56B4E9",
    3: "#009E73",
    4: "#F0E442",
    5: "#0072B2",
    6: "#D55E00",
    7: "#CC79A7",
    8: "#44AA99",
    9: "#882255",
}

# 20 unknowns: visually distinct from standard palette
UNK_COLORS = [
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
        st.session_state.well_roles = {}
    if "std_count" not in st.session_state:
        st.session_state.std_count = 8
        for i in range(1, 9):
            st.session_state[f"std_conc_{i}"] = float(PIERCE_DEFAULTS[i])
    if "unk_count" not in st.session_state:
        st.session_state.unk_count = 1
        st.session_state["unk_name_1"]   = ""
        st.session_state["unk_target_1"] = 10.0
    if "active_role" not in st.session_state:
        st.session_state.active_role = "blank"
    if "last_file_hash" not in st.session_state:
        st.session_state.last_file_hash = ""

_init_state()


# ── Click / add callbacks ─────────────────────────────────────────────────────

def _assign_well(well: str) -> None:
    active = st.session_state.get("active_role", "blank")
    if st.session_state.well_roles.get(well) == active:
        st.session_state.well_roles.pop(well, None)   # same role → unassign
    else:
        st.session_state.well_roles[well] = active     # assign current role


def _add_standard() -> None:
    n = st.session_state.std_count + 1
    if n <= 9:
        st.session_state.std_count = n
        if f"std_conc_{n}" not in st.session_state:
            st.session_state[f"std_conc_{n}"] = float(PIERCE_DEFAULTS.get(n, 0.0))
        st.session_state.active_role = f"std_{n}"


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


# ── Plate CSS ─────────────────────────────────────────────────────────────────

def _plate_css() -> str:
    """
    Single <style> block injected once before the plate grid is rendered.
    Global rule: all buttons preceded by a .wm-* marker become circular.
    Per-well rules: set each button's background color and optional ring for
    wells that belong to the currently active role.
    """
    active = st.session_state.get("active_role", "")
    parts = ["""<style>
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

/* ── Circular well buttons (only those preceded by a .wm-* marker) ───── */
div[data-testid="element-container"]:has([class^="wm-"])
+ div[data-testid="element-container"] > div[data-testid="stButton"] > button {
    border-radius: 50% !important;
    width: 44px !important;
    height: 44px !important;
    min-height: 44px !important;
    padding: 0 !important;
    font-size: 7px !important;
    font-weight: 700 !important;
    color: rgba(255,255,255,0.85) !important;
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
</style>"""]

    # Per-well color rules
    color_rules = ["<style>"]
    for row in PLATE_ROWS:
        for col_num in range(1, 13):
            well   = f"{row}{col_num}"
            role   = st.session_state.well_roles.get(well, "")
            color  = _role_color(role)
            # White ring when this well belongs to the active role
            ring   = ("box-shadow: 0 0 0 2px white, 0 0 0 4px rgba(255,255,255,0.35) !important;"
                      "transform: scale(1.06) !important;") if (role and role == active) else ""
            color_rules.append(
                f"div[data-testid='element-container']:has(.wm-{well})"
                f" + div[data-testid='element-container']"
                f" > div[data-testid='stButton'] > button"
                f" {{ background-color: {color} !important; {ring} }}"
            )
    color_rules.append("</style>")
    parts.append("\n".join(color_rules))
    return "\n".join(parts)


# ── Role label helper ─────────────────────────────────────────────────────────

def _role_label(r: str) -> str:
    if r == "blank":
        return "Blank  (0 µg/mL)"
    if r.startswith("std_"):
        n    = int(r.split("_")[1])
        conc = st.session_state.get(f"std_conc_{n}", PIERCE_DEFAULTS.get(n, "?"))
        conc_str = str(int(conc)) if isinstance(conc, float) and conc.is_integer() else str(conc)
        return f"Standard {n}  ({conc_str} µg/mL)"
    n    = int(r.split("_")[1])
    name = (st.session_state.get(f"unk_name_{n}", "") or f"Sample_{n}").strip() or f"Sample_{n}"
    return f"Unknown {n}:  {name}"


# ══════════════════════════════════════════════════════════════════════════════
# Main UI
# ══════════════════════════════════════════════════════════════════════════════

st.title("🧬 BCA Protein Assay Calculator")
st.caption(
    "Upload your plate reader file, use the plate picker to assign wells, "
    "then click **Calculate**. Downloads match CLI output exactly — all "
    "calculations run from the same validated code."
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
    buffer_parts   = 1
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

# ── Step 1: File upload ───────────────────────────────────────────────────────

st.header("1 · Upload plate file")
uploaded = st.file_uploader(
    "SoftMax Pro export (.txt, .xls), Excel (.xlsx), or CSV (.csv)",
    type=["txt", "xls", "xlsx", "csv"],
    label_visibility="collapsed",
)

well_data = None
if uploaded is not None:
    # Clear well assignments when a new file is loaded
    fhash = hashlib.md5(uploaded.getvalue()).hexdigest()[:12]
    if st.session_state.last_file_hash != fhash:
        st.session_state.last_file_hash = fhash
        st.session_state.well_roles = {}
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
    st.header("2 · Well assignment")

    n_std = st.session_state.std_count
    n_unk = st.session_state.unk_count
    role_options = (
        ["blank"]
        + [f"std_{i}" for i in range(1, n_std + 1)]
        + [f"unk_{i}" for i in range(1, n_unk + 1)]
    )

    # Guard: if a previously selected role was removed, reset to blank
    if st.session_state.get("active_role") not in role_options:
        st.session_state.active_role = "blank"

    left_col, right_col = st.columns([1, 2.9])

    # ── Left: role picker + settings ─────────────────────────────────────────
    with left_col:
        st.subheader("Active role")
        st.caption("Pick a role, then click wells on the plate to assign them. "
                   "Click an assigned well again to unassign it.")

        active_role = st.selectbox(
            "Role",
            options=role_options,
            format_func=_role_label,
            key="active_role",
            label_visibility="collapsed",
        )

        # Role-specific settings
        st.markdown("---")
        if active_role == "blank":
            n_blank = sum(1 for r in st.session_state.well_roles.values() if r == "blank")
            if n_blank:
                st.success(f"✓ {n_blank} blank well{'s' if n_blank > 1 else ''} assigned")
            else:
                st.info("Click wells to mark the 0 µg/mL blank reference.")

        elif active_role.startswith("std_"):
            n = int(active_role.split("_")[1])
            n_wells = sum(1 for r in st.session_state.well_roles.values() if r == active_role)
            st.number_input(
                f"Standard {n} concentration (µg/mL)",
                key=f"std_conc_{n}",
                min_value=0.0,
                step=25.0,
            )
            color_badge = STD_COLORS.get(n, "#888")
            st.markdown(
                f'<span style="display:inline-flex;align-items:center;gap:6px;font-size:12px">'
                f'<span style="width:12px;height:12px;border-radius:50%;background:{color_badge};'
                f'border:1px solid rgba(0,0,0,0.2)"></span>'
                f'{"✓ " + str(n_wells) + " wells" if n_wells else "No wells yet"}'
                f'</span>',
                unsafe_allow_html=True,
            )

        elif active_role.startswith("unk_"):
            n = int(active_role.split("_")[1])
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
                f'{"✓ " + str(n_wells) + " wells" if n_wells else "No wells yet"}'
                f'</span>',
                unsafe_allow_html=True,
            )

        st.markdown("---")
        ca, cb = st.columns(2)
        with ca:
            st.button("+ Standard", on_click=_add_standard,
                      disabled=(n_std >= 9), use_container_width=True)
        with cb:
            st.button("+ Unknown",  on_click=_add_unknown,
                      disabled=(n_unk >= 20), use_container_width=True)

        # Assignment summary
        st.markdown("**Assigned wells**")
        counts: dict[str, int] = {}
        for r in st.session_state.well_roles.values():
            counts[r] = counts.get(r, 0) + 1

        if not counts:
            st.caption("None yet — click wells on the plate.")
        else:
            for role_key in (["blank"]
                             + [f"std_{i}" for i in range(1, n_std + 1)]
                             + [f"unk_{i}" for i in range(1, n_unk + 1)]):
                cnt = counts.get(role_key, 0)
                if not cnt:
                    continue
                c = _role_color(role_key)
                lbl = _role_label(role_key)
                st.markdown(
                    f'<div style="display:flex;align-items:center;gap:7px;'
                    f'margin:2px 0;font-size:12px">'
                    f'<span style="width:11px;height:11px;border-radius:50%;'
                    f'background:{c};flex-shrink:0;border:1px solid rgba(0,0,0,0.15)"></span>'
                    f'{lbl} — {cnt} well{"s" if cnt > 1 else ""}</div>',
                    unsafe_allow_html=True,
                )

    # ── Right: 96-well plate ──────────────────────────────────────────────────
    with right_col:
        st.subheader("96-well plate")
        st.caption(
            f"Active role: **{_role_label(active_role)}** — "
            "click a well to assign, click again to unassign."
        )

        # Inject all CSS in one pass before any buttons render
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
        for row_idx, row in enumerate(PLATE_ROWS):
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
                    # Marker div: provides the :has(.wm-{well}) CSS anchor
                    st.markdown(f'<div class="wm-{well}"></div>', unsafe_allow_html=True)
                    st.button(
                        well,
                        key=f"plate_{well}",
                        on_click=_assign_well,
                        args=(well,),
                    )

        # Colour legend
        legend_items = [("Unassigned", UNASSIGNED_COLOR), ("Blank", BLANK_COLOR)]
        for i in range(1, n_std + 1):
            conc = st.session_state.get(f"std_conc_{i}", PIERCE_DEFAULTS.get(i, "?"))
            cstr = str(int(conc)) if isinstance(conc, float) and conc.is_integer() else str(conc)
            legend_items.append((f"Std {i} ({cstr})", STD_COLORS[i]))
        for i in range(1, n_unk + 1):
            name = (st.session_state.get(f"unk_name_{i}", "") or f"Sample_{i}")[:14]
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

        blank_wells = sorted(
            w for w, r in st.session_state.well_roles.items()
            if r == "blank" and w in well_data
        )
        if not blank_wells:
            errors.append("No blank wells assigned. Assign at least one 0 µg/mL well.")

        standard_defs: list[dict] = []
        for i in range(1, st.session_state.std_count + 1):
            wells = sorted(
                w for w, r in st.session_state.well_roles.items()
                if r == f"std_{i}" and w in well_data
            )
            if not wells:
                continue
            conc = float(st.session_state.get(f"std_conc_{i}", 0) or 0)
            if conc <= 0:
                errors.append(f"Standard {i} has no concentration set.")
                continue
            standard_defs.append({"wells": wells, "concentration": conc})
        if len(standard_defs) < 2:
            errors.append("At least 2 standards with assigned wells are required.")

        unknown_defs: list[dict]  = []
        target_map:   dict[str, float] = {}
        for i in range(1, st.session_state.unk_count + 1):
            wells = sorted(
                w for w, r in st.session_state.well_roles.items()
                if r == f"unk_{i}" and w in well_data
            )
            if not wells:
                continue
            name = (st.session_state.get(f"unk_name_{i}", "") or f"Sample_{i}").strip() or f"Sample_{i}"
            target_ug = float(st.session_state.get(f"unk_target_{i}", 10.0) or 10.0)
            unknown_defs.append({"wells": wells, "name": name, "dilution": dilution_factor})
            target_map[name] = target_ug
        if not unknown_defs:
            errors.append("No unknown samples assigned. Assign at least one unknown well.")

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

                    # WB loading — one call per sample so each can have its own target µg
                    wb_parts: list[pd.DataFrame] = []
                    for udef in unknown_defs:
                        sname = udef["name"]
                        sub = unknowns_results[unknowns_results["sample_name"] == sname]
                        params = {
                            "target_ug":    target_map.get(sname, 10.0),
                            "max_well_vol": None,
                            "lysate_parts": lysate_parts,
                            "buffer_parts": buffer_parts,
                        }
                        wb_parts.append(calculate_western_loading(sub, params))
                    western_df = pd.concat(wb_parts).reset_index(drop=True)
                    western_params = {
                        "target_ug":    10.0,    # Excel header; actual per-sample targets above
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
    c1.metric("R²", f"{r_squared:.5f}",
              delta="⚠ below 0.99" if r_squared < 0.99 else "✓ good fit",
              delta_color="inverse" if r_squared < 0.99 else "normal")
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
