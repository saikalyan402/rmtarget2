from __future__ import annotations

# PATCH VERSION: overlap-fix-v2 + AUM-current-performance
 
import inspect
import io
from html import escape
from typing import Any, Dict, List, Optional, Sequence, Tuple
 
import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook as openpyxl_load_workbook
 
 
 
# =============================================================================
# 1. APPLICATION CONFIGURATION  (management assumptions only - no business data)
# =============================================================================
 
APP_TITLE = "Sales Performance Command Center"
APP_SUBTITLE = "FY27 · Current reality → scenario planning → target achievement"
 
# --- Financial-year timeline -------------------------------------------------
MONTHS_COMPLETED = 3        # April, May, June are complete
MONTHS_REMAINING = 9        # July .. March
MONTHS_JUL_JAN = 7          # July .. January
MONTHS_FEB_MAR = 2          # February, March
 
FUTURE_MONTHS: List[str] = [
    "July", "August", "September", "October",
    "November", "December", "January", "February", "March",
]
MOMENTUM_MONTHS = FUTURE_MONTHS[:MONTHS_JUL_JAN]      # July .. January
LEAKAGE_MONTHS = FUTURE_MONTHS[MONTHS_JUL_JAN:]       # February, March
MONTH_DATES = pd.date_range("2026-07-01", periods=len(FUTURE_MONTHS), freq="MS")
 
# --- Revenue assumptions (basis points) --------------------------------------
REVENUE_BPS: Dict[str, float] = {"Equity": 60.0, "Debt": 20.0, "Liquid": 10.0}
REVENUE_RATE: Dict[str, float] = {k: v / 10000.0 for k, v in REVENUE_BPS.items()}
 
ASSETS: List[str] = ["Equity", "Debt", "Liquid"]
SALES_TYPES: List[str] = ["GS", "NS"]
SALES_LABEL: Dict[str, str] = {"GS": "Gross Sales", "NS": "Net Sales"}
VERTICALS: List[str] = ["Retail", "DHNI", "VRM"]
 
# Revenue is always measured on one basis so Gross and Net are never added.
REVENUE_BASIS = "NS"
 
# --- Scenario assumptions -----------------------------------------------------
S1_RUNRATE_UPLIFT = 0.20            # +20% on current run rate
S2_EQUITY_TARGET = 1.00             # 100% of Equity FY target by January
S2_OVERALL_TARGET = 0.75            # 75% of overall FY target by January
S3_TARGET = 1.00                    # 100% of FY target by January
S3_DEFAULT_DIP = 0.20               # default Feb-Mar run-rate dip
S4_TARGET = 1.20                    # 120% of FY target by March
S5_EQUITY_TARGET = 1.20             # 120% Equity by March
S5_OVERALL_TARGET = 1.00            # 100% overall by March
S6_SEGMENT_TARGETS: Dict[str, float] = {
    "Digital": 1.40,
    "Retail B30": 1.25,
    "Others": 1.15,
}
S7_DEFAULT_JAN_TARGET = 1.00        # January milestone = 100% of FY target
S7_DEFAULT_MAR_TARGET = 1.00        # March outcome = 100% of FY target
S7_DEFAULT_LEAKAGE = 0.20           # Feb-Mar AUM leakage / run-rate pressure
 
SEGMENT_ORDER: List[str] = ["Digital", "Retail B30", "Others"]
 
# --- Scenario navigator definitions -------------------------------------------
# "short" and "thesis" are presentation copy only; every calculation key
# ("kind", milestones, multipliers) is unchanged.
SCENARIOS: Dict[int, Dict[str, str]] = {
    1: {
        "label": "Scenario 1 · +20% Run-Rate Push",
        "name": "+20% Run-Rate Push",
        "short": "Run Rate",
        "kind": "runrate",
        "thesis": "Lift the current pace by a fifth and let the year compound.",
        "explanation": (
            "Increase the current Apr-Jun monthly run rate by 20% from July onward "
            "and measure the resulting March achievement."
        ),
        "milestone": "March 2027 · run rate lifted 20% for the remaining 9 months",
    },
    2: {
        "label": "Scenario 2 · 75% Overall by Jan + 100% Equity",
        "name": "75% Overall by January + 100% Equity",
        "short": "Jan Milestone",
        "kind": "jan_target",
        "thesis": "Finish Equity early, carry three quarters of the book by January.",
        "explanation": (
            "Reach 100% of the Equity FY target and 75% of the overall FY target by January. "
            "The residual requirement is allocated to Debt and Liquid in FY-target proportion."
        ),
        "milestone": "January 2027 · Equity 100% of FY target, portfolio 75% of FY target",
    },
    3: {
        "label": "Scenario 3 · 100% by Jan, Then Feb-Mar Dip",
        "name": "100% by January, then Feb-Mar dip",
        "short": "Jan + Leakage",
        "kind": "jan_target",
        "thesis": "Land the full year by January, then absorb the closing dip.",
        "explanation": (
            "Reach 100% of the FY target by January, followed by a configurable "
            "February-March run-rate decline."
        ),
        "milestone": "January 2027 · 100% of FY target, then a reduced Feb-Mar run rate",
    },
    4: {
        "label": "Scenario 4 · 120% by March",
        "name": "120% by March",
        "short": "March Target",
        "kind": "march_target",
        "thesis": "Hold one pace for nine months and close the year at 120%.",
        "explanation": (
            "Determine the monthly run rate required to finish March at 120% of the FY target."
        ),
        "milestone": "March 2027 · 120% of FY target",
    },
    5: {
        "label": "Scenario 5 · 120% Equity + 100% Overall",
        "name": "120% Equity + 100% Overall by March",
        "short": "Equity + Overall",
        "kind": "march_target",
        "thesis": "Push Equity beyond target while the portfolio still lands at 100%.",
        "explanation": (
            "Reach 120% of the Equity FY target and 100% of the overall FY target by March, "
            "with Debt and Liquid balancing the remaining requirement."
        ),
        "milestone": "March 2027 · Equity 120% of FY target, portfolio 100% of FY target",
    },
    6: {
        "label": "Scenario 6 · Digital 140% + B30 125% + Others 115%",
        "name": "Digital 140% + Retail B30 125% + Others 115%",
        "short": "Segment",
        "kind": "march_target",
        "thesis": "Let the fastest segments carry a heavier share of the year.",
        "explanation": (
            "Model differentiated performance where Digital achieves 140%, Retail B30 achieves "
            "125% and Others achieve 115% of their respective FY targets."
        ),
        "milestone": "March 2027 · differentiated achievement by business segment",
    },
    7: {
        "label": "Scenario 7 · Momentum Build-Up to March 2027",
        "name": "Momentum Build-Up to March 2027",
        "short": "Momentum",
        "kind": "momentum",
        "thesis": "Build momentum now. Create a January buffer. Protect March.",
        "explanation": (
            "Build progressive month-on-month momentum from July 2026 to reach the January 2027 "
            "milestone, create sufficient buffer to absorb Feb-Mar run-rate leakage, and protect the "
            "March 2027 target."
        ),
        "milestone": "January 2027 milestone → Feb-Mar leakage absorbed → March 2027 target held",
    },
    8: {
        "label": "Scenario 8 · Channel Growth & Target Simulator",
        "name": "Channel Growth & Target Simulator",
        "short": "Channel Simulator",
        "kind": "channel_simulator",
        "thesis": "Nine channels, nine dials. Set the pace channel by channel.",
        "explanation": (
            "Independently adjust monthly growth, January 2027 target achievement and March 2027 "
            "target achievement for Digital, VRM, EM, B30, T30, T8, DHNI, Retail and Institutional."
        ),
        "milestone": "January 2027 target → February/March leakage → March 2027 target",
    },
    9: {
        "label": "Scenario 9 · Channel Mix Optimiser",
        "name": "Channel Mix Optimiser",
        "short": "Optimiser",
        "kind": "channel_optimizer",
        "thesis": "Find the minimum growth each channel must deliver. Nothing more.",
        "explanation": (
            "Find the minimum channel growth trajectory required to achieve a selected portfolio March "
            "ambition, while preserving the January milestone and leakage assumption."
        ),
        "milestone": "Portfolio March ambition optimised across nine channels",
    },
}
SCENARIO_ORDER: List[int] = [1, 2, 3, 4, 5, 6, 7, 8, 9]
 
# Scenario 8/9 channel planning universe.  Each channel has independent
# momentum, January 2027 target and March 2027 target controls.
CHANNELS: List[str] = [
    "Digital", "VRM", "EM", "B30", "T30", "T8", "DHNI", "Retail", "Institutional"
]
S8_DEFAULT_GROWTH: Dict[str, float] = {c: 0.05 for c in CHANNELS}
S8_DEFAULT_JAN_TARGET: Dict[str, float] = {c: 1.00 for c in CHANNELS}
S8_DEFAULT_MAR_TARGET: Dict[str, float] = {c: 1.00 for c in CHANNELS}
S8_DEFAULT_LEAKAGE = 0.20
 
# --- Workbook contract --------------------------------------------------------
SHEET_ALIASES: Dict[str, List[str]] = {
    "Retail": ["RM Retail Sales", "RM Retail", "Retail Sales"],
    "DHNI": ["RM DHNI", "DHNI", "RM D-HNI"],
    "VRM": ["VRM", "RM VRM", "VRM Sales"],
}
 
# Presentation/dashboard sheet. It is deliberately NOT included in
# SHEET_ALIASES because it is not an employee-level calculation sheet.
FINAL_SHEET_ALIASES: List[str] = ["FINAL", "Final", "Final Dashboard"]
 
COLUMN_SPEC: Dict[Tuple[str, str], Dict[str, List[str]]] = {
    ("GS", "Equity"): {
        "fy": ["FY 26 TGT EQ", "Equity GS Targets"],
        "ytd_tgt": ["YTD June EQ TGT", "Q1 Equity GS Targets"],
        "ach": ["Equity GS Ach YTD June", "Equity GS Actuals"],
    },
    ("GS", "Debt"): {
        "fy": ["FY 26 TGT DT", "Debt GS Targets"],
        "ytd_tgt": ["YTD June DT TGT", "Q1 Debt GS Targets"],
        "ach": ["Debt GS Ach", "Debt GS Actuals"],
    },
    ("GS", "Liquid"): {
        "fy": ["FY 26 TGT LIQ", "Liquid GS Targets"],
        "ytd_tgt": ["YTD June LIQ TGT", "Q1 Liquid GS Targets"],
        "ach": ["Liquid GS Ach", "Liquid GS Actuals"],
    },
    ("NS", "Equity"): {
        "fy": ["FY 26 TGT EQ NS", "Equity Net Targets"],
        "ytd_tgt": ["YTD June EQ NS TGT", "Q1 Equity Net Targets"],
        "ach": ["Equity NS Ach YTD June", "Equity Net Actuals"],
    },
    ("NS", "Debt"): {
        "fy": ["FY 26 TGT DT NS", "Debt Net Targets"],
        "ytd_tgt": ["YTD June DT NS TGT", "Q1 Debt Net Targets"],
        "ach": ["Debt NS Ach", "Debt Net Actuals"],
    },
    ("NS", "Liquid"): {
        "fy": ["FY 26 TGT LIQ NS", "Liquid Net Targets"],
        "ytd_tgt": ["YTD June LIQ NS TGT", "Q1 Liquid Net Targets"],
        "ach": ["Liquid NS Ach", "Liquid Net Actuals"],
    },
}
 
META_ALIASES: Dict[str, List[str]] = {
    "Employee Name": ["Employee Name", "Emp Name", "Name"],
    "Emp Code": ["Emp Code", "Employee Code"],
    "ADID": ["ADID", "AD ID"],
    "Status": ["Status", "Employee Status"],
    "Type": ["Type", "Employment Type", "Functional Designation"],
    "ZONE": ["ZONE", "Zone"],
    "REGION": ["REGION", "Region"],
    "EM City": ["EM City", "City", "Location"],
    "MKT TYPE": ["MKT TYPE", "Market Type", "Mkt Type"],
}
META_FIELDS: List[str] = list(META_ALIASES.keys())
 
# -----------------------------------------------------------------------------
# SEGMENT CLASSIFICATION CONFIGURATION (Scenario 6)
# Edit this block to change how business segments are identified. The scenario
# calculation engine reads the resulting mapping and never needs to change.
# -----------------------------------------------------------------------------
SEGMENT_RULES: Dict[str, Dict[str, Any]] = {
    "Digital": {
        "search_columns": ["MKT TYPE", "Type", "REGION", "ZONE", "EM City", "Status"],
        "keywords": ["digital", "online", "d2c", "e-com", "ecom", "virtual", "vrm",
                     "inside sales", "web"],
    },
    "Retail B30": {
        "search_columns": ["MKT TYPE", "REGION", "Type"],
        "keywords": ["b30"],
    },
}
FALLBACK_SEGMENT = "Others"
 
 
# =============================================================================
# 2. GENERIC HELPERS
# =============================================================================
 
class WorkbookError(Exception):
    """Raised when the uploaded workbook does not satisfy the data contract."""
 
 
def normalize_column_name(column: Any) -> str:
    """Collapse non-breaking spaces, repeated spaces and stray padding."""
    return " ".join(str(column).replace("\u00a0", " ").strip().split())
 
 
def _norm_key(column: Any) -> str:
    return normalize_column_name(column).casefold()
 
 
def _squash_key(column: Any) -> str:
    return "".join(ch for ch in _norm_key(column) if ch.isalnum())
 
 
def normalize_frame(frame: pd.DataFrame) -> pd.DataFrame:
    """Return a copy of the frame with normalised column labels."""
    out = frame.copy()
    out.columns = [normalize_column_name(c) for c in out.columns]
    return out
 
 
def clean_numeric(series: Optional[pd.Series]) -> pd.Series:
    """Coerce a column to numeric, tolerating text-formatted numbers."""
    if series is None:
        return pd.Series(dtype="float64")
    if series.dtype == object:
        series = (
            series.astype(str)
            .str.replace("\u00a0", " ", regex=False)
            .str.replace(",", "", regex=False)
            .str.replace("\u20b9", "", regex=False)
            .str.replace("%", "", regex=False)
            .str.strip()
            .replace({"": np.nan, "-": np.nan, "nan": np.nan, "NA": np.nan, "N/A": np.nan})
        )
    return pd.to_numeric(series, errors="coerce")
 
 
def as_text(series: pd.Series) -> pd.Series:
    """Coerce any column to clean text, mapping every missing marker to ''."""
    filled = series.where(series.notna(), "")
    return (
        filled.astype(str)
        .str.replace("\u00a0", " ", regex=False)
        .str.strip()
        .replace({"nan": "", "NaN": "", "NaT": "", "None": "", "<NA>": ""})
    )
 
 
def text_column(frame: pd.DataFrame, column: str) -> pd.Series:
    """Safe accessor for a metadata column that may be absent."""
    if column not in frame.columns:
        return pd.Series([""] * len(frame), index=frame.index, dtype=object)
    return as_text(frame[column])
 
 
def safe_div(numerator: Optional[float], denominator: Optional[float]) -> Optional[float]:
    """Division that never raises and never returns inf/nan."""
    try:
        if numerator is None or denominator is None:
            return None
        n = float(numerator)
        d = float(denominator)
        if not np.isfinite(n) or not np.isfinite(d) or abs(d) < 1e-12:
            return None
        result = n / d
        return result if np.isfinite(result) else None
    except (TypeError, ValueError):
        return None
 
 
def _num(value: Any) -> Optional[float]:
    """Normalise a possibly-missing numeric to float or None."""
    if value is None:
        return None
    try:
        f = float(value)
    except (TypeError, ValueError):
        return None
    return f if np.isfinite(f) else None
 
 
def _z(value: Any) -> float:
    """Numeric value with missing treated as zero (for summation)."""
    v = _num(value)
    return 0.0 if v is None else v
 
 
def _ssum(series: pd.Series) -> Optional[float]:
    """Sum that returns None when every entry is missing."""
    total = series.sum(min_count=1)
    return _num(total)
 
 
# --- Display formatting -------------------------------------------------------
 
NA_TEXT = "0"
 
 
def fmt_cr(value: Any, decimals: int = 0) -> str:
    v = _num(value)
    if v is None:
        return NA_TEXT
    return f"\u20b9 {v:,.{decimals}f} Cr"
 
 
def fmt_cr_signed(value: Any, decimals: int = 0) -> str:
    v = _num(value)
    if v is None:
        return NA_TEXT
    return f"\u20b9 {v:+,.{decimals}f} Cr"
 
 
def fmt_pct(value: Any, decimals: int = 1) -> str:
    v = _num(value)
    if v is None:
        return NA_TEXT
    return f"{v * 100:,.{decimals}f}%"
 
 
def fmt_pct_signed(value: Any, decimals: int = 1) -> str:
    v = _num(value)
    if v is None:
        return NA_TEXT
    return f"{v * 100:+,.{decimals}f}%"
 
 
def fmt_pts(value: Any, decimals: int = 1) -> str:
    """Percentage-point delta, e.g. +18.4 pts."""
    v = _num(value)
    if v is None:
        return NA_TEXT
    return f"{v * 100:+,.{decimals}f} pts"
 
 
def fmt_num(value: Any, decimals: int = 0) -> str:
    v = _num(value)
    if v is None:
        return NA_TEXT
    return f"{v:,.{decimals}f}"
 
 
FORMATTERS = {
    "cr": fmt_cr,
    "cr1": lambda v: fmt_cr(v, 1),
    "cr_signed": fmt_cr_signed,
    "cr1_signed": lambda v: fmt_cr_signed(v, 1),
    "pct": fmt_pct,
    "pct_signed": fmt_pct_signed,
    "pts": fmt_pts,
    "num": fmt_num,
    "txt": lambda v: NA_TEXT if v is None or (isinstance(v, float) and not np.isfinite(v)) else str(v),
}
 
# Formats whose values are read as numbers (right aligned in glass tables).
NUMERIC_FORMATS = {"cr", "cr1", "cr_signed", "cr1_signed", "pct", "pct_signed", "pts", "num"}
SIGNED_FORMATS = {"cr_signed", "cr1_signed", "pct_signed", "pts"}
 
 
def format_table(frame: pd.DataFrame, formats: Dict[str, str]) -> pd.DataFrame:
    """Return a display-ready copy of a numeric frame using the given formats."""
    out = pd.DataFrame(index=frame.index)
    for column in frame.columns:
        kind = formats.get(column, "txt")
        formatter = FORMATTERS.get(kind, FORMATTERS["txt"])
        out[column] = [formatter(v) for v in frame[column]]
    return out
 
 
# =============================================================================
# 3. WORKBOOK LOADING, VALIDATION & CLEANING
# =============================================================================
 
def _build_column_index(frame: pd.DataFrame) -> Dict[str, int]:
    """Map normalised column keys to positional index (first occurrence wins)."""
    index: Dict[str, int] = {}
    for position, column in enumerate(frame.columns):
        for key in (_norm_key(column), _squash_key(column)):
            if key and key not in index:
                index[key] = position
    return index
 
 
def _resolve_column(index: Dict[str, int], aliases: Sequence[str]) -> Optional[int]:
    for alias in aliases:
        for key in (_norm_key(alias), _squash_key(alias)):
            if key in index:
                return index[key]
    return None
 
 
def _expected_header_keys() -> set:
    keys = set()
    for spec in COLUMN_SPEC.values():
        for aliases in spec.values():
            for alias in aliases:
                keys.add(_norm_key(alias))
    for aliases in META_ALIASES.values():
        for alias in aliases:
            keys.add(_norm_key(alias))
    return keys
 
 
def _detect_header_row(raw: pd.DataFrame, max_scan: int = 15) -> int:
    """Find the row that actually holds the column headers."""
    expected = _expected_header_keys()
    best_row, best_score = 0, -1
    for row in range(min(max_scan, len(raw))):
        values = {_norm_key(v) for v in raw.iloc[row].tolist() if str(v).strip().lower() != "nan"}
        score = len(values & expected)
        if score > best_score:
            best_row, best_score = row, score
    return best_row if best_score >= 4 else 0
 
 
def _match_sheet(available: Sequence[str], aliases: Sequence[str]) -> Optional[str]:
    normalised = {_norm_key(s): s for s in available}
    for alias in aliases:
        key = _norm_key(alias)
        if key in normalised:
            return normalised[key]
    for alias in aliases:
        key = _norm_key(alias)
        for sheet_key, sheet in normalised.items():
            if key in sheet_key:
                return sheet
    return None
 
 
def validate_frame(frame: pd.DataFrame, index: Dict[str, int], sheet_label: str) -> List[str]:
    """Return a list of human-readable descriptions of missing required columns."""
    missing: List[str] = []
    for (sales, asset), spec in COLUMN_SPEC.items():
        for role, aliases in spec.items():
            if _resolve_column(index, aliases) is None:
                role_label = {
                    "fy": "FY target",
                    "ytd_tgt": "YTD June target",
                    "ach": "YTD June achievement",
                }[role]
                missing.append(
                    f"{sheet_label} · {SALES_LABEL[sales]} · {asset} {role_label} "
                    f"(expected column '{aliases[0]}')"
                )
    return missing
 
 
def _extract_records(frame: pd.DataFrame, vertical: str) -> pd.DataFrame:
    """Turn one workbook sheet into a tidy per-employee record frame."""
    index = _build_column_index(frame)
    records = pd.DataFrame(index=frame.index)
    records["Vertical"] = vertical
 
    for field, aliases in META_ALIASES.items():
        position = _resolve_column(index, aliases)
        if position is None:
            records[field] = ""
        else:
            records[field] = as_text(frame.iloc[:, position]).to_numpy()
 
    for (sales, asset), spec in COLUMN_SPEC.items():
        for role, aliases in spec.items():
            position = _resolve_column(index, aliases)
            series = frame.iloc[:, position] if position is not None else None
            records[f"{sales}_{asset}_{role}"] = clean_numeric(series).to_numpy()
 
    return records
 
 
def _clean_records(records: pd.DataFrame) -> pd.DataFrame:
    """Drop non-employee rows while preserving legitimate negative values."""
    names = text_column(records, "Employee Name").str.casefold()
    invalid = names.isin({"", "nan", "none", "total", "grand total", "sum", "subtotal"})
    numeric_columns = [c for c in records.columns if c.split("_")[0] in SALES_TYPES]
    empty_rows = records[numeric_columns].isna().all(axis=1)
    cleaned = records.loc[~(invalid | empty_rows)].copy()
    cleaned[numeric_columns] = cleaned[numeric_columns].fillna(0.0)
    return cleaned.reset_index(drop=True)
 
 
@st.cache_data(show_spinner=False)
def load_workbook(payload: bytes) -> pd.DataFrame:
    """Read, validate and clean the workbook into a single tidy record frame."""
    try:
        excel = pd.ExcelFile(io.BytesIO(payload), engine="openpyxl")
    except Exception as exc:  # pragma: no cover - defensive
        raise WorkbookError(
            "The file could not be opened as an Excel workbook. "
            "Please upload a valid .xlsx file."
        ) from exc
 
    available = list(excel.sheet_names)
    resolved: Dict[str, str] = {}
    for vertical, aliases in SHEET_ALIASES.items():
        sheet = _match_sheet(available, aliases)
        if sheet is not None:
            resolved[vertical] = sheet
 
    missing_sheets = [v for v in SHEET_ALIASES if v not in resolved]
    if missing_sheets:
        wanted = ", ".join(f"'{SHEET_ALIASES[v][0]}'" for v in missing_sheets)
        raise WorkbookError(
            f"The workbook is missing the required calculation sheet(s): {wanted}. "
            "Please upload the standard RM scorecard workbook."
        )
 
    frames: List[pd.DataFrame] = []
    problems: List[str] = []
    for vertical, sheet in resolved.items():
        raw = pd.read_excel(excel, sheet_name=sheet, header=None, nrows=20)
        header_row = _detect_header_row(raw)
        frame = normalize_frame(pd.read_excel(excel, sheet_name=sheet, header=header_row))
        index = _build_column_index(frame)
        problems.extend(validate_frame(frame, index, vertical))
        if not problems:
            frames.append(_extract_records(frame, vertical))
 
    if problems:
        raise WorkbookError(
            "The workbook is missing required columns:\n\n- " + "\n- ".join(problems[:12])
            + ("\n- \u2026" if len(problems) > 12 else "")
        )
 
    records = _clean_records(pd.concat(frames, ignore_index=True))
    if records.empty:
        raise WorkbookError("No employee records were found in the calculation sheets.")
    return records
 
 
# =============================================================================
# 3A. FINAL SHEET - MANAGEMENT SOURCE LAYER
# =============================================================================
 
def _find_final_sheet_name(sheet_names: Sequence[str]) -> Optional[str]:
    """Resolve the sixth workbook sheet named FINAL."""
    return _match_sheet(sheet_names, FINAL_SHEET_ALIASES)
 
 
def _excel_rgb(color: Any) -> Optional[str]:
    """Convert an openpyxl RGB colour to a CSS hex colour when possible."""
    try:
        if color is None or color.type != "rgb" or not color.rgb:
            return None
        raw = str(color.rgb)
        rgb = raw[-6:]
        if len(rgb) == 6 and all(ch in "0123456789abcdefABCDEF" for ch in rgb):
            return f"#{rgb}"
    except Exception:
        return None
    return None
 
 
def _display_excel_value(value: Any, number_format: str = "") -> str:
    """Format Excel values for the FINAL dashboard without exposing formulas."""
    if value is None:
        return ""
 
    if isinstance(value, (pd.Timestamp,)):
        return value.strftime("%d-%b-%Y")
 
    # Date/datetime objects from openpyxl.
    if hasattr(value, "strftime") and not isinstance(value, str):
        try:
            return value.strftime("%d-%b-%Y")
        except Exception:
            pass
 
    if isinstance(value, (int, float, np.integer, np.floating)) and not isinstance(value, bool):
        numeric = float(value)
        if not np.isfinite(numeric):
            return ""
 
        fmt = str(number_format or "")
        if "%" in fmt:
            # Excel stores percentages as fractions.
            decimals = 0
            if "." in fmt.split("%")[0]:
                decimals = len(fmt.split("%")[0].split(".")[-1].replace("0", "0"))
            decimals = min(max(decimals, 0), 2)
            return f"{numeric * 100:,.{decimals}f}%"
 
        # Respect accounting-style parentheses when possible.
        negative_parentheses = numeric < 0 and "(" in fmt and ")" in fmt
        abs_value = abs(numeric)
 
        if abs(abs_value - round(abs_value)) < 1e-9:
            rendered = f"{abs_value:,.0f}"
        else:
            rendered = f"{abs_value:,.2f}".rstrip("0").rstrip(".")
 
        if numeric < 0:
            return f"({rendered})" if negative_parentheses else f"-{rendered}"
        return rendered
 
    return str(value)
 
 
@st.cache_data(show_spinner=False)
def load_final_sheet_frame(payload: bytes) -> pd.DataFrame:
    """
    Read the FINAL sheet as raw cells for a fallback / inspection table.
 
    The sheet intentionally has no single header row, so header=None is used.
    """
    excel = pd.ExcelFile(io.BytesIO(payload), engine="openpyxl")
    sheet = _find_final_sheet_name(excel.sheet_names)
    if sheet is None:
        raise WorkbookError(
            "The workbook does not contain the sixth sheet 'FINAL'. "
            "Please upload the workbook that contains Summary, Summary-Achievement, "
            "RM Retail Sales, RM DHNI, VRM and FINAL."
        )
 
    frame = pd.read_excel(excel, sheet_name=sheet, header=None)
    frame = frame.dropna(axis=0, how="all").dropna(axis=1, how="all")
    return frame.reset_index(drop=True)
 
 
@st.cache_data(show_spinner=False)
def build_final_sheet_html(payload: bytes) -> str:
    """
    Render the Excel FINAL sheet into a scrollable dark-glass HTML table.
 
    Merged headings, alignment and the workbook's own fills are preserved,
    using cached formula results (data_only=True).
    """
    workbook = openpyxl_load_workbook(io.BytesIO(payload), data_only=True)
    sheet_name = _find_final_sheet_name(workbook.sheetnames)
    if sheet_name is None:
        raise WorkbookError(
            "The workbook does not contain the sixth sheet 'FINAL'. "
            "Please upload the workbook that contains the FINAL dashboard sheet."
        )
 
    ws = workbook[sheet_name]
 
    # Find the real content boundary from non-empty values rather than Excel's
    # formatted used-range, which can extend thousands of blank rows/columns.
    non_empty = [
        (cell.row, cell.column)
        for row in ws.iter_rows()
        for cell in row
        if cell.value not in (None, "")
    ]
    if not non_empty:
        return "<div class='glass-note'>The FINAL sheet is empty.</div>"
 
    min_row = min(r for r, _ in non_empty)
    max_row = max(r for r, _ in non_empty)
    min_col = min(c for _, c in non_empty)
    max_col = max(c for _, c in non_empty)
 
    # Merged-cell lookup.
    merge_anchor: Dict[Tuple[int, int], Tuple[int, int]] = {}
    merge_covered: set = set()
    for merged in ws.merged_cells.ranges:
        # Only consider merged ranges intersecting the content boundary.
        if (
            merged.max_row < min_row or merged.min_row > max_row
            or merged.max_col < min_col or merged.min_col > max_col
        ):
            continue
        anchor = (merged.min_row, merged.min_col)
        merge_anchor[anchor] = (
            merged.max_row - merged.min_row + 1,
            merged.max_col - merged.min_col + 1,
        )
        for rr in range(merged.min_row, merged.max_row + 1):
            for cc in range(merged.min_col, merged.max_col + 1):
                if (rr, cc) != anchor:
                    merge_covered.add((rr, cc))
 
    html_parts = [
        "<div class='final-sheet-scroll'><table class='final-sheet-table'>"
    ]
 
    for row_idx in range(min_row, max_row + 1):
        row_values = [
            ws.cell(row=row_idx, column=col_idx).value
            for col_idx in range(min_col, max_col + 1)
        ]
 
        # Keep dashboard spacing, but collapse very large blank areas to a
        # single slim spacer row.
        if all(v in (None, "") for v in row_values):
            html_parts.append(
                "<tr><td colspan='{}' class='final-spacer'></td></tr>".format(
                    max_col - min_col + 1
                )
            )
            continue
 
        html_parts.append("<tr>")
        for col_idx in range(min_col, max_col + 1):
            if (row_idx, col_idx) in merge_covered:
                continue
 
            cell = ws.cell(row=row_idx, column=col_idx)
            rowspan, colspan = merge_anchor.get((row_idx, col_idx), (1, 1))
 
            value = _display_excel_value(cell.value, cell.number_format)
            fill_color = _excel_rgb(cell.fill.fgColor)
            font_color = _excel_rgb(cell.font.color)
 
            styles: List[str] = []
            if fill_color and fill_color.lower() not in {"#000000", "#ffffff"}:
                # Workbook fills are light; keep them, but darken the ink so the
                # cell stays readable inside the dark command centre.
                styles.append(f"background:{fill_color}")
                styles.append("color:#0B0D12")
            elif font_color and font_color.lower() in {"#ff0000", "#c00000", "#e60000"}:
                styles.append("color:#FF8A8A")
 
            if cell.font.bold:
                styles.append("font-weight:650")
            if cell.font.italic:
                styles.append("font-style:italic")
 
            horizontal = getattr(cell.alignment, "horizontal", None)
            if horizontal in {"center", "centerContinuous"}:
                styles.append("text-align:center")
            elif horizontal == "right":
                styles.append("text-align:right")
 
            attrs = []
            if rowspan > 1:
                attrs.append(f"rowspan='{rowspan}'")
            if colspan > 1:
                attrs.append(f"colspan='{colspan}'")
 
            html_parts.append(
                f"<td {' '.join(attrs)} style=\"{';'.join(styles)}\">{escape(value)}</td>"
            )
        html_parts.append("</tr>")
 
    html_parts.append("</table></div>")
    return "".join(html_parts)
 
 
FINAL_METRIC_ROWS: List[str] = [
    "Overall", "Equity", "Debt", "Liquid",
    "Retail", "DHNI", "VRM", "Insti", "Digital",
    "Alternatives", "Passives",
]
 
FINAL_ASSET_ROWS: List[str] = ["Equity", "Debt", "Liquid"]
FINAL_CHANNEL_ROWS: List[str] = [
    "Retail", "DHNI", "VRM", "Insti", "Digital", "Alternatives", "Passives",
]
 
 
def _final_key(value: Any) -> str:
    return " ".join(str(value or "").replace("\u00a0", " ").strip().split()).lower()
 
 
def _final_number(value: Any) -> Optional[float]:
    """Convert FINAL-sheet numeric / accounting text into float."""
    if value is None:
        return None
    if isinstance(value, (int, float, np.integer, np.floating)) and not isinstance(value, bool):
        numeric = float(value)
        return numeric if np.isfinite(numeric) else None
 
    raw = str(value).strip()
    if not raw or raw.lower() in {"-", "\u2014", "na", "n/a", "none", "nan", "#div/0!"}:
        return None
 
    negative = raw.startswith("(") and raw.endswith(")")
    cleaned = (
        raw.replace(",", "")
        .replace("\u20b9", "")
        .replace("%", "")
        .replace("(", "")
        .replace(")", "")
        .strip()
    )
    try:
        numeric = float(cleaned)
    except ValueError:
        return None
    if negative:
        numeric = -numeric
    return numeric
 
 
def _final_known_label(value: Any) -> Optional[str]:
    key = _final_key(value)
    aliases = {
        "overall": "Overall",
        "equity": "Equity",
        "debt": "Debt",
        "liquid": "Liquid",
        "retail": "Retail",
        "dhni": "DHNI",
        "vrm": "VRM",
        "insti": "Insti",
        "institutional": "Insti",
        "digital": "Digital",
        "alternatives": "Alternatives",
        "alternate": "Alternatives",
        "passives": "Passives",
        "passive": "Passives",
    }
    return aliases.get(key)
 
 
def _scan_final_sheet(ws: Any) -> Tuple[int, int]:
    """Cap the scan to the management-dashboard area, not formatted blank Excel space."""
    return min(max(ws.max_row, 1), 320), min(max(ws.max_column, 1), 180)
 
 
def _find_final_cells(ws: Any, wanted: str) -> List[Tuple[int, int]]:
    wanted_key = _final_key(wanted)
    max_row, max_col = _scan_final_sheet(ws)
    found: List[Tuple[int, int]] = []
    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            if _final_key(ws.cell(row=row, column=col).value) == wanted_key:
                found.append((row, col))
    return found
 
 
def _find_header_near(
    ws: Any,
    title_position: Tuple[int, int],
    header_name: str,
    row_window: int = 8,
    col_before: int = 4,
    col_after: int = 24,
) -> Optional[Tuple[int, int]]:
    title_row, title_col = title_position
    max_row, max_col = _scan_final_sheet(ws)
    wanted = _final_key(header_name)
    for row in range(title_row, min(title_row + row_window, max_row) + 1):
        start_col = max(1, title_col - col_before)
        end_col = min(max_col, title_col + col_after)
        for col in range(start_col, end_col + 1):
            if _final_key(ws.cell(row=row, column=col).value) == wanted:
                return row, col
    return None
 
 
def _augment_final_runrate(frame: pd.DataFrame, months_done: int) -> pd.DataFrame:
    """Recreate the run-rate formulas shown in FINAL from Target and YTD."""
    if frame.empty:
        return frame
 
    months = max(int(months_done), 1)
    out = frame.copy()
    out["FY27 Target"] = pd.to_numeric(out["FY27 Target"], errors="coerce")
    out["YTD"] = pd.to_numeric(out["YTD"], errors="coerce")
 
    out["Achievement %"] = np.where(
        (out["FY27 Target"] > 0) & (out["YTD"] >= 0),
        out["YTD"] / out["FY27 Target"],
        np.nan,
    )
    out["Current RR"] = out["YTD"] / months
 
    # This matches the FINAL workbook:
    # 154,757 / 12 = 12,896; 20,699 / 12 = 1,725.
    out["Required RR to Target"] = out["FY27 Target"] / 12.0
 
    # Annualise the current Apr-Jun run rate to a full 12-month FY.
    out["Estimated FY @ Current RR"] = out["Current RR"] * 12.0
    out["Projected FY %"] = np.where(
        (out["FY27 Target"] > 0) & (out["Estimated FY @ Current RR"] >= 0),
        out["Estimated FY @ Current RR"] / out["FY27 Target"],
        np.nan,
    )
    return out
 
 
def _parse_final_sales_block(ws: Any, title: str, months_done: int) -> pd.DataFrame:
    """Parse the NET SALES / GROSS SALES run-rate block on the FINAL sheet."""
    positions = _find_final_cells(ws, title)
    for position in positions:
        header = _find_header_near(ws, position, "FY27 Target")
        if header is None:
            continue
 
        header_row, target_col = header
 
        # In the FINAL block the row label sits immediately to the left of FY27 Target.
        label_col = max(1, target_col - 1)
        ytd_col = target_col + 1
 
        rows: List[Dict[str, Any]] = []
        seen: set = set()
        max_row, _ = _scan_final_sheet(ws)
 
        for row in range(header_row + 1, min(header_row + 28, max_row) + 1):
            label = _final_known_label(ws.cell(row=row, column=label_col).value)
            if label is None or label in seen:
                continue
 
            target = _final_number(ws.cell(row=row, column=target_col).value)
            ytd = _final_number(ws.cell(row=row, column=ytd_col).value)
 
            # Ignore title/spacer rows accidentally matching a label.
            if target is None and ytd is None:
                continue
 
            rows.append({"Metric": label, "FY27 Target": target, "YTD": ytd})
            seen.add(label)
 
        if rows:
            frame = pd.DataFrame(rows).set_index("Metric")
            order = [label for label in FINAL_METRIC_ROWS if label in frame.index]
            return _augment_final_runrate(
                frame.loc[order].reset_index(), months_done
            ).set_index("Metric")
 
    return pd.DataFrame()
 
 
def _parse_final_aum_block(ws: Any) -> pd.DataFrame:
    """Parse Target / Current AUM from the FINAL management sheet."""
    positions = _find_final_cells(ws, "AUM")
    for position in positions:
        title_row, title_col = position
        max_row, max_col = _scan_final_sheet(ws)
 
        header_row = None
        target_col = None
        current_col = None
 
        for row in range(title_row, min(title_row + 6, max_row) + 1):
            for col in range(max(1, title_col - 4), min(max_col, title_col + 8) + 1):
                if _final_key(ws.cell(row=row, column=col).value) == "target":
                    # Find Current on the same header row.
                    for cc in range(col + 1, min(max_col, col + 5) + 1):
                        if _final_key(ws.cell(row=row, column=cc).value) == "current":
                            header_row = row
                            target_col = col
                            current_col = cc
                            break
                if header_row is not None:
                    break
            if header_row is not None:
                break
 
        if header_row is None or target_col is None or current_col is None:
            continue
 
        label_col = max(1, target_col - 1)
        rows: List[Dict[str, Any]] = []
        seen: set = set()
 
        for row in range(header_row + 1, min(header_row + 28, max_row) + 1):
            label = _final_known_label(ws.cell(row=row, column=label_col).value)
            if label is None or label in seen:
                continue
            target = _final_number(ws.cell(row=row, column=target_col).value)
            current = _final_number(ws.cell(row=row, column=current_col).value)
            if target is None and current is None:
                continue
            rows.append({"Metric": label, "Target": target, "Current": current})
            seen.add(label)
 
        if rows:
            frame = pd.DataFrame(rows).set_index("Metric")
            order = [label for label in FINAL_METRIC_ROWS if label in frame.index]
            frame = frame.loc[order].copy()
            frame["Achievement %"] = np.where(
                (frame["Target"] > 0) & (frame["Current"] >= 0),
                frame["Current"] / frame["Target"],
                np.nan,
            )
            frame["Gap to Target"] = frame["Target"] - frame["Current"]
            return frame
 
    return pd.DataFrame()
 
 
def _parse_months_done(ws: Any) -> int:
    positions = (
        _find_final_cells(ws, "#months done")
        + _find_final_cells(ws, "months done")
        + _find_final_cells(ws, "# months done")
    )
    max_row, max_col = _scan_final_sheet(ws)
 
    for row, col in positions:
        # Search immediately around / below the label for the highlighted count.
        for rr in range(row, min(row + 4, max_row) + 1):
            for cc in range(max(1, col - 2), min(max_col, col + 4) + 1):
                value = _final_number(ws.cell(row=rr, column=cc).value)
                if value is not None and 1 <= value <= 12:
                    return int(round(value))
    return MONTHS_COMPLETED
 
 
@st.cache_data(show_spinner=False)
def parse_final_dashboard_metrics(payload: bytes) -> Dict[str, Any]:
    """Return structured management metrics from the workbook's FINAL sheet."""
    workbook = openpyxl_load_workbook(io.BytesIO(payload), data_only=True)
    sheet_name = _find_final_sheet_name(workbook.sheetnames)
    if sheet_name is None:
        raise WorkbookError(
            "The workbook is missing the required sixth sheet 'FINAL'. "
            "Please upload the workbook containing FINAL."
        )
 
    ws = workbook[sheet_name]
    months_done = _parse_months_done(ws)
 
    gs = _parse_final_sales_block(ws, "GROSS SALES", months_done)
    ns = _parse_final_sales_block(ws, "NET SALES", months_done)
    aum = _parse_final_aum_block(ws)
 
    return {
        "sheet_name": sheet_name,
        "months_done": months_done,
        "GS": gs,
        "NS": ns,
        "AUM": aum,
    }
 
 
# =============================================================================
# 4. SEGMENT & CHANNEL IDENTIFICATION (Scenarios 6, 8, 9)
# =============================================================================
 
def identify_segments(records: pd.DataFrame) -> Dict[str, Dict[str, Any]]:
    """Suggest a column + values that identify each configured business segment."""
    suggestions: Dict[str, Dict[str, Any]] = {}
    for segment, rule in SEGMENT_RULES.items():
        for column in rule["search_columns"]:
            if column not in records.columns:
                continue
            values = sorted({v for v in text_column(records, column) if v.strip()})
            matches = [
                v for v in values
                if any(keyword in v.casefold() for keyword in rule["keywords"])
            ]
            if matches:
                suggestions[segment] = {"column": column, "values": matches}
                break
    return suggestions
 
 
def map_business_segments(
    records: pd.DataFrame,
    mapping: Dict[str, Dict[str, Any]],
) -> pd.DataFrame:
    """Assign every record to Digital / Retail B30 / Others using the mapping."""
    out = records.copy()
    out["Segment"] = FALLBACK_SEGMENT
    # Reverse priority so that the first segment in SEGMENT_ORDER wins.
    for segment in reversed([s for s in SEGMENT_ORDER if s != FALLBACK_SEGMENT]):
        rule = mapping.get(segment)
        if not rule:
            continue
        column, values = rule.get("column"), set(rule.get("values") or [])
        if not column or column not in out.columns or not values:
            continue
        mask = text_column(out, column).isin(values)
        out.loc[mask, "Segment"] = segment
    return out
 
 
def segment_diagnostics(records: pd.DataFrame) -> Dict[str, int]:
    counts = records["Segment"].value_counts().to_dict()
    return {segment: int(counts.get(segment, 0)) for segment in SEGMENT_ORDER}
 
 
CHANNEL_KEYWORDS: Dict[str, List[str]] = {
    "Digital": ["digital", "online", "d2c", "e-com", "ecom", "virtual", "web"],
    "VRM": ["vrm", "virtual relationship", "virtual rm"],
    "EM": ["em", "emerging market", "em city"],
    "B30": ["b30", "b-30", "b 30"],
    "T30": ["t30", "t-30", "t 30"],
    "T8": ["t8", "t-8", "t 8"],
    "DHNI": ["dhni", "d-hni", "hni", "wealth"],
    "Retail": ["retail", "rm retail"],
    "Institutional": ["insti", "institutional", "institution", "institutional sales"],
}
 
 
def _channel_text_score(row: pd.Series, channel: str) -> int:
    values = " ".join(str(row.get(c, "")) for c in META_FIELDS).casefold()
    return sum(1 for keyword in CHANNEL_KEYWORDS[channel] if keyword in values)
 
 
def identify_channels(records: pd.DataFrame) -> Dict[str, Dict[str, Any]]:
    """Suggest channel mappings from workbook metadata without hard-coding one schema."""
    suggestions: Dict[str, Dict[str, Any]] = {}
    usable = [
        f for f in META_FIELDS
        if f in records.columns and text_column(records, f).ne("").any()
    ]
    for channel in CHANNELS:
        best = (0, None, [])
        for column in usable:
            values = sorted({v for v in text_column(records, column) if v.strip()})
            matches = [v for v in values if any(k in v.casefold() for k in CHANNEL_KEYWORDS[channel])]
            score = len(matches)
            if score > best[0]:
                best = (score, column, matches)
        if best[1] and best[2]:
            suggestions[channel] = {"column": best[1], "values": best[2]}
        elif channel == "VRM" and "Vertical" in records.columns:
            suggestions[channel] = {"column": "Vertical", "values": ["VRM"]}
        elif channel == "DHNI" and "Vertical" in records.columns:
            suggestions[channel] = {"column": "Vertical", "values": ["DHNI"]}
        elif channel == "Retail" and "Vertical" in records.columns:
            suggestions[channel] = {"column": "Vertical", "values": ["Retail"]}
    return suggestions
 
 
def map_business_channels(
    records: pd.DataFrame,
    mapping: Dict[str, Dict[str, Any]],
) -> pd.DataFrame:
    """Assign every record to one of the nine Scenario 8 planning channels."""
    out = records.copy()
    out["Channel"] = "Unclassified"
    # More specific channels are applied first; explicit mapping wins.
    for channel in CHANNELS:
        rule = mapping.get(channel)
        if not rule:
            continue
        column, values = rule.get("column"), set(rule.get("values") or [])
        if not column or column not in out.columns or not values:
            continue
        mask = text_column(out, column).isin(values)
        out.loc[mask, "Channel"] = channel
    # Use the existing vertical as a safe fallback for the three explicit RM populations.
    if "Vertical" in out.columns:
        for channel, vertical in (("VRM", "VRM"), ("DHNI", "DHNI"), ("Retail", "Retail")):
            mask = (out["Channel"] == "Unclassified") & (out["Vertical"] == vertical)
            out.loc[mask, "Channel"] = channel
    return out
 
 
# =============================================================================
# 5. BASE GRID & CURRENT-STATE STATISTICS
# =============================================================================
 
def build_base_grid(records: pd.DataFrame) -> pd.DataFrame:
    """Aggregate records to the finest analytical grain used by the engine."""
    rows: List[Dict[str, Any]] = []
    grouped = records.groupby(["Vertical", "Segment", "Channel"], dropna=False)
    for (vertical, segment, channel), block in grouped:
        for sales in SALES_TYPES:
            for asset in ASSETS:
                rows.append({
                    "Vertical": vertical,
                    "Segment": segment,
                    "Channel": channel,
                    "Sales": sales,
                    "Asset": asset,
                    "fy_target": float(block[f"{sales}_{asset}_fy"].sum()),
                    "ytd_target": float(block[f"{sales}_{asset}_ytd_tgt"].sum()),
                    "ytd_ach": float(block[f"{sales}_{asset}_ach"].sum()),
                })
    return pd.DataFrame(rows)
 
 
def filter_grid(
    grid: pd.DataFrame,
    sales: Optional[str] = None,
    asset: Optional[str] = None,
    vertical: Optional[str] = None,
    segment: Optional[str] = None,
    channel: Optional[str] = None,
) -> pd.DataFrame:
    mask = pd.Series(True, index=grid.index)
    if sales is not None:
        mask &= grid["Sales"] == sales
    if asset is not None:
        mask &= grid["Asset"] == asset
    if vertical is not None:
        mask &= grid["Vertical"] == vertical
    if segment is not None:
        mask &= grid["Segment"] == segment
    if channel is not None and "Channel" in grid.columns:
        mask &= grid["Channel"] == channel
    return grid.loc[mask]
 
 
def current_asset_stats(fy_target: float, ytd_target: float, ytd_ach: float) -> Dict[str, Any]:
    """Baseline statistics for one asset / group. Never scenario dependent."""
    fy_target = _z(fy_target)
    ytd_target = _z(ytd_target)
    ytd_ach = _z(ytd_ach)
    current_rr = ytd_ach / MONTHS_COMPLETED if MONTHS_COMPLETED else None
    current_march = ytd_ach + (current_rr or 0.0) * MONTHS_REMAINING
    return {
        "fy_target": fy_target,
        "ytd_target": ytd_target,
        "ytd_ach": ytd_ach,
        "current_rr": current_rr,
        "ytd_ach_pct": safe_div(ytd_ach, ytd_target),
        "fy_completed_pct": safe_div(ytd_ach, fy_target),
        "current_march": current_march,
        "current_march_pct": safe_div(current_march, fy_target),
    }
 
 
def summarize_current(grid: pd.DataFrame, **filters: Any) -> Dict[str, Any]:
    """Baseline statistics for an arbitrary slice of the base grid."""
    subset = filter_grid(grid, **filters)
    return current_asset_stats(
        subset["fy_target"].sum(),
        subset["ytd_target"].sum(),
        subset["ytd_ach"].sum(),
    )
 
 
# =============================================================================
# 6. SCENARIO ENGINE - SCENARIOS 1 TO 6
# =============================================================================
 
def _blank_cell(stats: Dict[str, Any]) -> Dict[str, Any]:
    cell = dict(stats)
    cell.update({
        "scen_rr": None, "rr_change_pct": None, "feb_mar_rr": None,
        "jan_required": None, "jan_amount": None, "jan_pct": None,
        "jan_buffer": None, "jan_buffer_pct": None,
        "march_required": None, "march_amount": None, "march_pct": None,
        "milestone_pct": None, "incremental_sales": None,
        "headroom_amt": None, "headroom_pct": None,
        "momentum_g": None, "feasible": None, "binding": None,
        "trajectory": None, "note": "",
    })
    return cell
 
 
def compute_cell(
    fy_target: float,
    ytd_target: float,
    ytd_ach: float,
    kind: str,
    multiplier: Optional[float] = None,
    uplift: Optional[float] = None,
    dip: float = 0.0,
) -> Dict[str, Any]:
    """Scenario mathematics for one asset / group (scenarios 1-6)."""
    stats = current_asset_stats(fy_target, ytd_target, ytd_ach)
    cell = _blank_cell(stats)
    ach = stats["ytd_ach"]
    current_rr = stats["current_rr"] or 0.0
 
    if kind == "runrate":
        scen_rr = current_rr * (1.0 + (uplift or 0.0))
        jan_amount = ach + scen_rr * MONTHS_JUL_JAN
        cell.update({
            "scen_rr": scen_rr,
            "feb_mar_rr": scen_rr,
            "jan_amount": jan_amount,
            "march_amount": ach + scen_rr * MONTHS_REMAINING,
            "milestone_pct": None,
        })
 
    elif kind == "jan_target":
        required = max(_z(multiplier) * stats["fy_target"], ach)
        scen_rr = max(required - ach, 0.0) / MONTHS_JUL_JAN
        jan_amount = ach + scen_rr * MONTHS_JUL_JAN
        feb_mar_rr = scen_rr * (1.0 - dip)
        cell.update({
            "scen_rr": scen_rr,
            "feb_mar_rr": feb_mar_rr,
            "jan_required": required,
            "jan_amount": jan_amount,
            "march_amount": jan_amount + feb_mar_rr * MONTHS_FEB_MAR,
            "milestone_pct": multiplier,
        })
 
    elif kind == "march_target":
        required = max(_z(multiplier) * stats["fy_target"], ach)
        scen_rr = max(required - ach, 0.0) / MONTHS_REMAINING
        cell.update({
            "scen_rr": scen_rr,
            "feb_mar_rr": scen_rr,
            "march_required": required,
            "jan_amount": ach + scen_rr * MONTHS_JUL_JAN,
            "march_amount": ach + scen_rr * MONTHS_REMAINING,
            "milestone_pct": multiplier,
        })
 
    else:  # pragma: no cover - guarded by the scenario registry
        raise ValueError(f"Unknown scenario kind: {kind}")
 
    return _finalise_cell(cell)
 
 
def _finalise_cell(cell: Dict[str, Any]) -> Dict[str, Any]:
    """Recompute all derived ratios from the absolute amounts in the cell."""
    fy_target = cell.get("fy_target")
    current_rr = _num(cell.get("current_rr"))
    scen_rr = _num(cell.get("scen_rr"))
 
    cell["ytd_ach_pct"] = safe_div(cell.get("ytd_ach"), cell.get("ytd_target"))
    cell["fy_completed_pct"] = safe_div(cell.get("ytd_ach"), fy_target)
    cell["current_march_pct"] = safe_div(cell.get("current_march"), fy_target)
    cell["jan_pct"] = safe_div(cell.get("jan_amount"), fy_target)
    cell["march_pct"] = safe_div(cell.get("march_amount"), fy_target)
 
    if current_rr is not None and current_rr > 0 and scen_rr is not None:
        cell["rr_change_pct"] = (scen_rr / current_rr) - 1.0
    else:
        cell["rr_change_pct"] = None
 
    march_amount = _num(cell.get("march_amount"))
    current_march = _num(cell.get("current_march"))
    if march_amount is not None and current_march is not None:
        cell["incremental_sales"] = march_amount - current_march
 
    jan_required = _num(cell.get("jan_required"))
    jan_amount = _num(cell.get("jan_amount"))
    if jan_required is not None and jan_amount is not None:
        cell["jan_buffer"] = jan_amount - jan_required
        cell["jan_buffer_pct"] = safe_div(cell["jan_buffer"], jan_required)
 
    march_required = _num(cell.get("march_required"))
    if march_required is not None and march_amount is not None:
        cell["headroom_amt"] = march_amount - march_required
        march_pct = cell.get("march_pct")
        required_pct = safe_div(march_required, fy_target)
        if march_pct is not None and required_pct is not None:
            cell["headroom_pct"] = march_pct - required_pct
        cell["feasible"] = cell["headroom_amt"] >= -1e-6
 
    if cell.get("milestone_pct") is None:
        cell["milestone_pct"] = cell.get("march_pct")
    return cell
 
 
def scenario_multipliers(grid: pd.DataFrame, scenario_id: int) -> Dict[Tuple[str, str, str], float]:
    """
    Derive the per-asset FY-target multiplier for the selected scenario.
 
    Key is (sales type, asset, segment); segment is '*' unless the scenario
    differentiates by business segment.
    """
    multipliers: Dict[Tuple[str, str, str], float] = {}
 
    if scenario_id in (1, 7):
        return multipliers
 
    if scenario_id == 3:
        for sales in SALES_TYPES:
            for asset in ASSETS:
                multipliers[(sales, asset, "*")] = S3_TARGET
        return multipliers
 
    if scenario_id == 4:
        for sales in SALES_TYPES:
            for asset in ASSETS:
                multipliers[(sales, asset, "*")] = S4_TARGET
        return multipliers
 
    if scenario_id == 6:
        for sales in SALES_TYPES:
            for asset in ASSETS:
                for segment in SEGMENT_ORDER:
                    multipliers[(sales, asset, segment)] = S6_SEGMENT_TARGETS.get(segment, 1.0)
        return multipliers
 
    # Scenarios 2 and 5 balance Debt and Liquid around a fixed Equity ambition.
    equity_mult = S2_EQUITY_TARGET if scenario_id == 2 else S5_EQUITY_TARGET
    overall_mult = S2_OVERALL_TARGET if scenario_id == 2 else S5_OVERALL_TARGET
 
    for sales in SALES_TYPES:
        targets = {
            asset: float(filter_grid(grid, sales=sales, asset=asset)["fy_target"].sum())
            for asset in ASSETS
        }
        total_target = sum(targets.values())
        required_overall = overall_mult * total_target
        required_equity = equity_mult * targets["Equity"]
        remaining = max(required_overall - required_equity, 0.0)
        denominator = targets["Debt"] + targets["Liquid"]
        share = safe_div(remaining, denominator)
        balance_mult = 0.0 if share is None else share
        multipliers[(sales, "Equity", "*")] = equity_mult
        multipliers[(sales, "Debt", "*")] = balance_mult
        multipliers[(sales, "Liquid", "*")] = balance_mult
 
    return multipliers
 
 
def _multiplier_for(
    multipliers: Dict[Tuple[str, str, str], float],
    sales: str,
    asset: str,
    segment: str,
) -> Optional[float]:
    if (sales, asset, segment) in multipliers:
        return multipliers[(sales, asset, segment)]
    return multipliers.get((sales, asset, "*"))
 
 
def apply_scenario_grid(
    grid: pd.DataFrame,
    scenario_id: int,
    params: Dict[str, Any],
    multipliers: Dict[Tuple[str, str, str], float],
) -> pd.DataFrame:
    """Evaluate scenarios 1-6 over every cell of the base grid."""
    kind = SCENARIOS[scenario_id]["kind"]
    dip = float(params.get("dip", 0.0)) if scenario_id == 3 else 0.0
    uplift = S1_RUNRATE_UPLIFT if scenario_id == 1 else None
 
    results: List[Dict[str, Any]] = []
    for row in grid.to_dict("records"):
        multiplier = _multiplier_for(multipliers, row["Sales"], row["Asset"], row["Segment"])
        cell = compute_cell(
            row["fy_target"], row["ytd_target"], row["ytd_ach"],
            kind=kind, multiplier=multiplier, uplift=uplift, dip=dip,
        )
        cell.update({
            "Vertical": row["Vertical"], "Segment": row["Segment"],
            "Channel": row.get("Channel", "Unclassified"),
            "Sales": row["Sales"], "Asset": row["Asset"],
        })
        results.append(cell)
    return pd.DataFrame(results)
 
 
SUMMABLE_FIELDS = [
    "fy_target", "ytd_target", "ytd_ach", "current_rr", "current_march",
    "scen_rr", "feb_mar_rr", "jan_required", "jan_amount",
    "march_required", "march_amount",
]
 
 
def summarize_cells(subset: pd.DataFrame) -> Dict[str, Any]:
    """Aggregate scenario cells additively and rebuild every derived ratio."""
    cell: Dict[str, Any] = {}
    for field in SUMMABLE_FIELDS:
        cell[field] = _ssum(subset[field]) if field in subset.columns else None
    milestones = subset["milestone_pct"].dropna().unique() if "milestone_pct" in subset else []
    cell["milestone_pct"] = float(milestones[0]) if len(milestones) == 1 else None
    cell["note"] = ""
    cell["momentum_g"] = None
    cell["trajectory"] = None
    cell["binding"] = None
    cell["feasible"] = None
    return _finalise_cell(cell)
 
 
# --- Named scenario entry points (thin wrappers over the shared engine) -------
 
def _scenario_frame(grid: pd.DataFrame, scenario_id: int, params: Dict[str, Any]) -> pd.DataFrame:
    return apply_scenario_grid(grid, scenario_id, params, scenario_multipliers(grid, scenario_id))
 
 
def calculate_scenario_1(grid: pd.DataFrame, params: Dict[str, Any]) -> pd.DataFrame:
    """+20% run-rate push from July onward."""
    return _scenario_frame(grid, 1, params)
 
 
def calculate_scenario_2(grid: pd.DataFrame, params: Dict[str, Any]) -> pd.DataFrame:
    """100% Equity and 75% overall FY target by January."""
    return _scenario_frame(grid, 2, params)
 
 
def calculate_scenario_3(grid: pd.DataFrame, params: Dict[str, Any]) -> pd.DataFrame:
    """100% of FY target by January, then a configurable Feb-Mar dip."""
    return _scenario_frame(grid, 3, params)
 
 
def calculate_scenario_4(grid: pd.DataFrame, params: Dict[str, Any]) -> pd.DataFrame:
    """120% of FY target by March."""
    return _scenario_frame(grid, 4, params)
 
 
def calculate_scenario_5(grid: pd.DataFrame, params: Dict[str, Any]) -> pd.DataFrame:
    """120% Equity and 100% overall FY target by March."""
    return _scenario_frame(grid, 5, params)
 
 
def calculate_scenario_6(grid: pd.DataFrame, params: Dict[str, Any]) -> pd.DataFrame:
    """Digital 140%, Retail B30 125%, Others 115% of their FY targets."""
    return _scenario_frame(grid, 6, params)
 
 
SCENARIO_FUNCTIONS = {
    1: calculate_scenario_1,
    2: calculate_scenario_2,
    3: calculate_scenario_3,
    4: calculate_scenario_4,
    5: calculate_scenario_5,
    6: calculate_scenario_6,
}
 
 
# =============================================================================
# 7. SCENARIO 7 - MOMENTUM ENGINE
# =============================================================================
 
def _momentum_sum(growth: float, months: int, tail_factors: Sequence[float]) -> float:
    """Sum of compounding monthly run-rate multiples, including leakage tail."""
    factor = 1.0 + growth
    build = sum(factor ** k for k in range(1, months + 1))
    tail = sum(f * factor ** months for f in tail_factors)
    return build + tail
 
 
def solve_momentum_rate(
    current_rr: float,
    required_amount: float,
    months: int = MONTHS_JUL_JAN,
    tail_factors: Sequence[float] = (),
    upper: float = 3.0,
) -> Optional[float]:
    """
    Back-solve the minimum month-on-month growth rate g such that the
    compounding trajectory delivers the required incremental amount.
 
    Returns 0.0 when no additional momentum is needed and None when the
    requirement cannot be met within the search bounds.
    """
    rr = _num(current_rr)
    need = _num(required_amount)
    if rr is None or need is None or rr <= 0:
        return None
    if need <= 0:
        return 0.0
 
    def shortfall(growth: float) -> float:
        return rr * _momentum_sum(growth, months, tail_factors) - need
 
    if shortfall(0.0) >= 0:
        return 0.0
    if shortfall(upper) < 0:
        return None
 
    low, high = 0.0, upper
    for _ in range(240):
        mid = (low + high) / 2.0
        if shortfall(mid) >= 0:
            high = mid
        else:
            low = mid
    return high
 
 
def calculate_momentum_trajectory(
    current_rr: float,
    growth: Optional[float],
    leakage: float,
    flat_rate: Optional[float] = None,
) -> List[float]:
    """Monthly run rates for July -> March (momentum build, then leakage)."""
    if growth is None:
        base = _z(flat_rate)
        build = [base] * MONTHS_JUL_JAN
    else:
        rr = _z(current_rr)
        build = [rr * (1.0 + growth) ** k for k in range(1, MONTHS_JUL_JAN + 1)]
    january_rr = build[-1] if build else 0.0
    february_rr = january_rr * (1.0 - leakage)
    march_rr = february_rr * (1.0 - leakage)
    return build + [february_rr, march_rr]
 
 
def calculate_leakage_impact(january_rr: float, leakage: float) -> Dict[str, float]:
    """February and March run rates after AUM leakage / run-rate pressure."""
    february_rr = _z(january_rr) * (1.0 - leakage)
    march_rr = february_rr * (1.0 - leakage)
    return {
        "february_rr": february_rr,
        "march_rr": march_rr,
        "feb_mar_sales": february_rr + march_rr,
    }
 
 
def calculate_momentum_headroom(
    march_amount: float,
    march_required: float,
    fy_target: float,
    march_target_pct: float,
) -> Dict[str, Optional[float]]:
    """Scenario achievement versus the March ambition, in Cr and in points."""
    headroom_amt = _z(march_amount) - _z(march_required)
    achieved_pct = safe_div(march_amount, fy_target)
    headroom_pct = None if achieved_pct is None else achieved_pct - march_target_pct
    return {"headroom_amt": headroom_amt, "headroom_pct": headroom_pct}
 
 
def calculate_scenario_7(
    fy_target: float,
    ytd_target: float,
    ytd_ach: float,
    params: Dict[str, Any],
) -> Dict[str, Any]:
    """
    Momentum build-up model.
 
    Solves for the month-on-month growth rate that simultaneously satisfies the
    January milestone and, after Feb-Mar leakage, the March ambition.
    """
    jan_target_pct = float(params.get("jan_target", S7_DEFAULT_JAN_TARGET))
    mar_target_pct = float(params.get("mar_target", S7_DEFAULT_MAR_TARGET))
    leakage = float(params.get("leakage", S7_DEFAULT_LEAKAGE))
 
    stats = current_asset_stats(fy_target, ytd_target, ytd_ach)
    cell = _blank_cell(stats)
    ach = stats["ytd_ach"]
    current_rr = stats["current_rr"] or 0.0
 
    jan_required = max(jan_target_pct * stats["fy_target"], ach)
    mar_required = max(mar_target_pct * stats["fy_target"], ach)
    tail = ((1.0 - leakage), (1.0 - leakage) ** 2)
 
    growth_jan = solve_momentum_rate(current_rr, jan_required - ach, MONTHS_JUL_JAN)
    growth_mar = solve_momentum_rate(current_rr, mar_required - ach, MONTHS_JUL_JAN, tail)
 
    note = ""
    binding = None
    flat_rate = None
 
    if current_rr <= 0:
        # Momentum compounding is undefined on a non-positive run rate:
        # fall back to the flat run rate required to hold both milestones.
        flat_jan = max(jan_required - ach, 0.0) / MONTHS_JUL_JAN
        denominator = MONTHS_JUL_JAN + tail[0] + tail[1]
        flat_mar = max(mar_required - ach, 0.0) / denominator
        flat_rate = max(flat_jan, flat_mar)
        growth = None
        binding = "January" if flat_jan >= flat_mar else "March"
        note = (
            "Current run rate is not positive, so compounding momentum cannot be applied. "
            "The flat monthly run rate required to hold the milestones is shown instead."
        )
    elif growth_jan is None and growth_mar is None:
        growth = None
        flat_rate = max(mar_required - ach, 0.0) / (MONTHS_JUL_JAN + tail[0] + tail[1])
        binding = "March"
        note = "The requirement exceeds the momentum search range; a flat required run rate is shown."
    else:
        candidates = [g for g in (growth_jan, growth_mar) if g is not None]
        growth = max(candidates)
        binding = "March" if (growth_mar is not None and growth == growth_mar
                              and (growth_jan is None or growth_mar >= growth_jan)) else "January"
 
    trajectory = calculate_momentum_trajectory(current_rr, growth, leakage, flat_rate)
    build_phase = trajectory[:MONTHS_JUL_JAN]
    jan_amount = ach + sum(build_phase)
    january_rr = build_phase[-1] if build_phase else 0.0
    leak = calculate_leakage_impact(january_rr, leakage)
    march_amount = jan_amount + leak["feb_mar_sales"]
 
    headroom = calculate_momentum_headroom(
        march_amount, mar_required, stats["fy_target"], mar_target_pct
    )
    shortfall = max(mar_required - march_amount, 0.0)
    denominator = 1.0 + tail[0] + tail[1]
    additional_jan_rr = shortfall / denominator if denominator else None
 
    cell.update({
        "scen_rr": january_rr,
        "feb_mar_rr": leak["february_rr"],
        "march_rr": leak["march_rr"],
        "jan_required": jan_required,
        "jan_amount": jan_amount,
        "march_required": mar_required,
        "march_amount": march_amount,
        "milestone_pct": jan_target_pct,
        "march_target_pct": mar_target_pct,
        "momentum_g": growth,
        "flat_rate": flat_rate,
        "leakage": leakage,
        "trajectory": trajectory,
        "binding": binding,
        "note": note,
        "additional_march_sales": shortfall,
        "additional_jan_rr": additional_jan_rr,
        "avg_scen_rr": (sum(trajectory) / len(trajectory)) if trajectory else None,
    })
    cell = _finalise_cell(cell)
    cell["headroom_amt"] = headroom["headroom_amt"]
    cell["headroom_pct"] = headroom["headroom_pct"]
    cell["feasible"] = shortfall <= 1e-6
    # Momentum run rate versus the current flat run rate, measured on the
    # January exit rate (the pace the business must be running at by then).
    cell["rr_change_pct"] = (
        (january_rr / current_rr) - 1.0 if current_rr and current_rr > 0 else None
    )
    return cell
 
 
# =============================================================================
# 7A. SCENARIO 8/9 - CHANNEL SIMULATOR & MIX OPTIMISER
# =============================================================================
 
def _s8_channel_params(params: Dict[str, Any], channel: str) -> Tuple[float, float, float, float]:
    growth = float(params.get("channel_growth", {}).get(channel, S8_DEFAULT_GROWTH.get(channel, 0.05)))
    jan_target = float(
        params.get("channel_jan_target", {}).get(channel, S8_DEFAULT_JAN_TARGET.get(channel, 1.0))
    )
    mar_target = float(
        params.get("channel_mar_target", {}).get(channel, S8_DEFAULT_MAR_TARGET.get(channel, 1.0))
    )
    leakage = float(params.get("leakage", S8_DEFAULT_LEAKAGE))
    return growth, jan_target, mar_target, leakage
 
 
def calculate_scenario_8_cell(
    fy_target: float,
    ytd_target: float,
    ytd_ach: float,
    channel: str,
    params: Dict[str, Any],
) -> Dict[str, Any]:
    """Fixed-growth channel trajectory with independent Jan-2027 and Mar-2027 targets."""
    stats = current_asset_stats(fy_target, ytd_target, ytd_ach)
    cell = _blank_cell(stats)
    growth, jan_target_pct, mar_target_pct, leakage = _s8_channel_params(params, channel)
    current_rr = _z(stats["current_rr"])
    trajectory = calculate_momentum_trajectory(current_rr, growth, leakage)
    jan_amount = _z(stats["ytd_ach"]) + sum(trajectory[:MONTHS_JUL_JAN])
    march_amount = jan_amount + sum(trajectory[MONTHS_JUL_JAN:])
    jan_required = max(jan_target_pct * stats["fy_target"], _z(stats["ytd_ach"]))
    march_required = max(mar_target_pct * stats["fy_target"], _z(stats["ytd_ach"]))
    jan_gap = jan_amount - jan_required
    march_gap = march_amount - march_required
    cell.update({
        "channel": channel,
        "scen_rr": trajectory[MONTHS_JUL_JAN - 1] if trajectory else current_rr,
        "feb_mar_rr": trajectory[MONTHS_JUL_JAN] if len(trajectory) > MONTHS_JUL_JAN else None,
        "march_rr": trajectory[-1] if trajectory else None,
        "jan_required": jan_required, "jan_amount": jan_amount,
        "march_required": march_required, "march_amount": march_amount,
        "milestone_pct": jan_target_pct, "march_target_pct": mar_target_pct,
        "momentum_g": growth, "leakage": leakage, "trajectory": trajectory,
        "jan_buffer": jan_gap, "jan_buffer_pct": safe_div(jan_gap, jan_required),
        "headroom_amt": march_gap, "headroom_pct": safe_div(march_gap, march_required),
        "feasible": jan_gap >= -1e-6 and march_gap >= -1e-6,
        "binding": "January" if jan_gap < 0 else ("March" if march_gap < 0 else "None"),
        "additional_march_sales": max(-march_gap, 0.0),
        "additional_jan_rr": max(-jan_gap, 0.0) / MONTHS_JUL_JAN,
    })
    return _finalise_cell(cell)
 
 
def calculate_scenario_8_grid(grid: pd.DataFrame, params: Dict[str, Any]) -> pd.DataFrame:
    rows = []
    for row in grid.to_dict("records"):
        channel = row.get("Channel", "Unclassified")
        cell = calculate_scenario_8_cell(
            row["fy_target"], row["ytd_target"], row["ytd_ach"], channel, params
        )
        cell.update({k: row[k] for k in ("Vertical", "Segment", "Channel", "Sales", "Asset")})
        rows.append(cell)
    return pd.DataFrame(rows)
 
 
def calculate_scenario_9_grid(grid: pd.DataFrame, params: Dict[str, Any]) -> pd.DataFrame:
    """Optimiser: solve minimum MoM growth per channel to meet its Jan/Mar targets."""
    rows = []
    for row in grid.to_dict("records"):
        channel = row.get("Channel", "Unclassified")
        _, jan_target, mar_target, leakage = _s8_channel_params(params, channel)
        stats = current_asset_stats(row["fy_target"], row["ytd_target"], row["ytd_ach"])
        ach = _z(stats["ytd_ach"])
        rr = _z(stats["current_rr"])
        jan_req = max(jan_target * stats["fy_target"], ach)
        mar_req = max(mar_target * stats["fy_target"], ach)
        gj = solve_momentum_rate(rr, jan_req - ach, MONTHS_JUL_JAN)
        gm = solve_momentum_rate(rr, mar_req - ach, MONTHS_JUL_JAN, ((1 - leakage), (1 - leakage) ** 2))
        growth = max([g for g in (gj, gm) if g is not None], default=0.0)
        cell = calculate_scenario_8_cell(
            row["fy_target"], row["ytd_target"], row["ytd_ach"], channel,
            {**params, "channel_growth": {**params.get("channel_growth", {}), channel: growth}},
        )
        cell.update({k: row[k] for k in ("Vertical", "Segment", "Channel", "Sales", "Asset")})
        cell["optimized_growth"] = growth
        rows.append(cell)
    return pd.DataFrame(rows)
 
 
# =============================================================================
# 8. SCENARIO MODEL - ONE INTERFACE FOR EVERY VIEW
# =============================================================================
 
class ScenarioModel:
    """Evaluates the selected scenario for any slice of the business."""
 
    def __init__(self, scenario_id: int, grid: pd.DataFrame, params: Dict[str, Any]):
        self.scenario_id = scenario_id
        self.meta = SCENARIOS[scenario_id]
        self.grid = grid
        self.params = params
        self.multipliers = scenario_multipliers(grid, scenario_id)
        self._cache: Dict[Tuple, Dict[str, Any]] = {}
        if scenario_id == 7:
            self.scenario_grid = None
        elif scenario_id == 8:
            self.scenario_grid = calculate_scenario_8_grid(grid, params)
        elif scenario_id == 9:
            self.scenario_grid = calculate_scenario_9_grid(grid, params)
        else:
            self.scenario_grid = SCENARIO_FUNCTIONS[scenario_id](grid, params)
 
    # -- core accessor --------------------------------------------------------
    def cell(
        self,
        sales: str,
        asset: Optional[str] = None,
        vertical: Optional[str] = None,
        segment: Optional[str] = None,
        channel: Optional[str] = None,
    ) -> Dict[str, Any]:
        key = (sales, asset, vertical, segment, channel)
        if key in self._cache:
            return self._cache[key]
 
        if self.scenario_id == 7:
            subset = filter_grid(self.grid, sales=sales, asset=asset,
                                 vertical=vertical, segment=segment, channel=channel)
            cell = calculate_scenario_7(
                subset["fy_target"].sum(),
                subset["ytd_target"].sum(),
                subset["ytd_ach"].sum(),
                self.params,
            )
        else:
            frame = self.scenario_grid
            mask = frame["Sales"] == sales
            if asset is not None:
                mask &= frame["Asset"] == asset
            if vertical is not None:
                mask &= frame["Vertical"] == vertical
            if segment is not None:
                mask &= frame["Segment"] == segment
            if channel is not None and "Channel" in frame.columns:
                mask &= frame["Channel"] == channel
            cell = summarize_cells(frame.loc[mask])
 
        self._cache[key] = cell
        return cell
 
    # -- convenience views ----------------------------------------------------
    def assets(self, sales: str, **filters: Any) -> Dict[str, Dict[str, Any]]:
        return {asset: self.cell(sales, asset=asset, **filters) for asset in ASSETS}
 
    def baseline(self, sales: str, **filters: Any) -> Dict[str, Any]:
        return summarize_current(self.grid, sales=sales, **filters)
 
    def implied_milestones(self, sales: str) -> Dict[str, Optional[float]]:
        return {
            asset: _multiplier_for(self.multipliers, sales, asset, "*")
            for asset in ASSETS
        }
 
    def available_segments(self) -> List[str]:
        present = set(self.grid["Segment"].unique())
        return [s for s in SEGMENT_ORDER if s in present]
 
    def available_verticals(self) -> List[str]:
        present = set(self.grid["Vertical"].unique())
        return [v for v in VERTICALS if v in present]
 
    def available_channels(self) -> List[str]:
        if "Channel" not in self.grid.columns:
            return []
        present = set(self.grid["Channel"].unique())
        return [c for c in CHANNELS if c in present]
 
 
# =============================================================================
# 9. REVENUE ENGINE
# =============================================================================
 
def calculate_revenue(cells_by_asset: Dict[str, Dict[str, Any]], field: str) -> Dict[str, Any]:
    """Asset-class revenue from a set of scenario cells. No blended rate."""
    by_asset: Dict[str, Optional[float]] = {}
    sales_by_asset: Dict[str, Optional[float]] = {}
    total = 0.0
    for asset in ASSETS:
        amount = _num(cells_by_asset.get(asset, {}).get(field))
        sales_by_asset[asset] = amount
        revenue = None if amount is None else amount * REVENUE_RATE[asset]
        by_asset[asset] = revenue
        total += 0.0 if revenue is None else revenue
    return {"by_asset": by_asset, "sales_by_asset": sales_by_asset, "total": total}
 
 
def calculate_baseline_revenue(cells_by_asset: Dict[str, Dict[str, Any]]) -> Dict[str, Any]:
    """Baseline revenue on the current-run-rate March projection (scenario independent)."""
    return calculate_revenue(cells_by_asset, "current_march")
 
 
def calculate_incremental_revenue(
    scenario_revenue: Dict[str, Any],
    baseline_revenue: Dict[str, Any],
) -> Dict[str, Any]:
    incremental = {
        asset: (
            None
            if scenario_revenue["by_asset"].get(asset) is None
            or baseline_revenue["by_asset"].get(asset) is None
            else scenario_revenue["by_asset"][asset] - baseline_revenue["by_asset"][asset]
        )
        for asset in ASSETS
    }
    total = scenario_revenue["total"] - baseline_revenue["total"]
    uplift = safe_div(scenario_revenue["total"], baseline_revenue["total"])
    return {
        "by_asset": incremental,
        "total": total,
        "uplift_pct": None if uplift is None else uplift - 1.0,
        "contribution": {
            asset: safe_div(scenario_revenue["by_asset"].get(asset), scenario_revenue["total"])
            for asset in ASSETS
        },
    }
 
 
def revenue_bundle(model: ScenarioModel, basis: str, **filters: Any) -> Dict[str, Any]:
    """Baseline / scenario / incremental revenue for any slice of the business."""
    cells = model.assets(basis, **filters)
    baseline = calculate_baseline_revenue(cells)
    scenario = calculate_revenue(cells, "march_amount")
    january = calculate_revenue(cells, "jan_amount")
    incremental = calculate_incremental_revenue(scenario, baseline)
    return {
        "cells": cells,
        "baseline": baseline,
        "scenario": scenario,
        "january": january,
        "incremental": incremental,
    }
 
 
# =============================================================================
# 10. FINAL <-> MODEL BRIDGE
# =============================================================================
 
def _model_metric_baseline(model: "ScenarioModel", sales: str, label: str) -> Dict[str, Any]:
    if label == "Overall":
        return model.baseline(sales)
    if label in ASSETS:
        return model.baseline(sales, asset=label)
    if label in VERTICALS:
        return model.baseline(sales, vertical=label)
    return {}
 
 
def _model_metric_cell(model: "ScenarioModel", sales: str, label: str) -> Optional[Dict[str, Any]]:
    if label == "Overall":
        return model.cell(sales)
    if label in ASSETS:
        return model.cell(sales, asset=label)
    if label in VERTICALS:
        return model.cell(sales, vertical=label)
    return None
 
 
def final_sales_metrics(
    final_metrics: Dict[str, Any],
    model: "ScenarioModel",
    sales: str,
) -> pd.DataFrame:
    """
    Use FINAL Target/YTD as the visible source of truth.
    Fill modelled Overall/asset/vertical rows from the scenario data only if
    the FINAL parser could not locate them.
    """
    parsed = final_metrics.get(sales)
    if isinstance(parsed, pd.DataFrame) and not parsed.empty:
        frame = parsed.copy()
    else:
        frame = pd.DataFrame()
 
    months_done = int(final_metrics.get("months_done", MONTHS_COMPLETED))
    needed = ["Overall", *ASSETS, *VERTICALS]
 
    fallback_rows: List[Dict[str, Any]] = []
    for label in needed:
        if not frame.empty and label in frame.index:
            continue
        base = _model_metric_baseline(model, sales, label)
        if not base:
            continue
        fallback_rows.append(
            {
                "Metric": label,
                "FY27 Target": base.get("fy_target"),
                "YTD": base.get("ytd_ach"),
            }
        )
 
    if fallback_rows:
        fallback = _augment_final_runrate(pd.DataFrame(fallback_rows), months_done).set_index("Metric")
        if frame.empty:
            frame = fallback
        else:
            frame = pd.concat([frame, fallback], axis=0)
 
    if frame.empty:
        return frame
 
    # Preserve management ordering and avoid accidental duplicate rows.
    frame = frame.loc[~frame.index.duplicated(keep="first")].copy()
    order = [label for label in FINAL_METRIC_ROWS if label in frame.index]
    return frame.loc[order]
 
 
def build_final_scenario_comparison(
    final_metrics: Dict[str, Any],
    model: "ScenarioModel",
    sales: str,
) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """Scenario comparison expressed on the same Target/YTD/run-rate metrics as FINAL."""
    current = final_sales_metrics(final_metrics, model, sales)
    rows: List[Dict[str, Any]] = []
 
    for label in current.index.tolist():
        source = current.loc[label]
        cell = _model_metric_cell(model, sales, label)
 
        final_target = _num(source.get("FY27 Target"))
        current_rr = _num(source.get("Current RR"))
        current_projection = _num(source.get("Estimated FY @ Current RR"))
        current_pct = _num(source.get("Projected FY %"))
 
        model_base = _model_metric_baseline(model, sales, label)
        model_target = _num(model_base.get("fy_target")) if model_base else None
 
        scenario_pct = _num(cell.get("march_pct")) if cell is not None else None
 
        scenario_amount = None
        scenario_rr = None
        rr_change = None
        delta_pp = None
 
        if cell is not None:
            # Anchor the scenario outcome to the FINAL FY27 target so the
            # management comparison uses one common metric base.
            scenario_amount = (
                final_target * scenario_pct
                if final_target is not None and scenario_pct is not None
                else _num(cell.get("march_amount"))
            )
 
            scenario_rr = _num(cell.get("scen_rr"))
            if (
                scenario_rr is not None
                and final_target is not None
                and model_target is not None
                and model_target != 0
            ):
                scenario_rr = scenario_rr * final_target / model_target
 
            if current_rr is not None and scenario_rr is not None and current_rr != 0:
                rr_change = scenario_rr / current_rr - 1.0
 
            if current_pct is not None and scenario_pct is not None:
                delta_pp = scenario_pct - current_pct
 
        rows.append(
            {
                "Metric": label,
                "FY27 Target": final_target,
                "YTD": _num(source.get("YTD")),
                "Current RR": current_rr,
                "Required RR to Target": _num(source.get("Required RR to Target")),
                "Current FY Estimate": current_projection,
                "Current Projected %": current_pct,
                "Scenario / Required RR": scenario_rr,
                "Run Rate Change %": rr_change,
                "Scenario March Estimate": scenario_amount,
                "Scenario March %": scenario_pct,
                "Scenario \u0394 pp": delta_pp,
            }
        )
 
    formats = {
        "Metric": "txt",
        "FY27 Target": "cr",
        "YTD": "cr",
        "Current RR": "cr",
        "Required RR to Target": "cr",
        "Current FY Estimate": "cr",
        "Current Projected %": "pct",
        "Scenario / Required RR": "cr",
        "Run Rate Change %": "pct_signed",
        "Scenario March Estimate": "cr",
        "Scenario March %": "pct",
        "Scenario \u0394 pp": "pts",
    }
    return pd.DataFrame(rows), formats
 
 
# =============================================================================
# 11. TABLE BUILDERS
# =============================================================================
 
def build_current_overview(model: ScenarioModel) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """Baseline table - identical for every scenario."""
    rows = []
    for sales in SALES_TYPES:
        base = model.baseline(sales)
        rows.append({
            "Sales": SALES_LABEL[sales],
            "FY Target": base["fy_target"],
            "YTD June Target": base["ytd_target"],
            "YTD June Achievement": base["ytd_ach"],
            "Target Achieved %": base["ytd_ach_pct"],
            "FY Target Completed %": base["fy_completed_pct"],
            "Current Run Rate": base["current_rr"],
            "Current March Projection": base["current_march"],
            "Current March Projection %": base["current_march_pct"],
        })
    formats = {
        "Sales": "txt", "FY Target": "cr", "YTD June Target": "cr",
        "YTD June Achievement": "cr", "Target Achieved %": "pct",
        "FY Target Completed %": "pct", "Current Run Rate": "cr",
        "Current March Projection": "cr", "Current March Projection %": "pct",
    }
    return pd.DataFrame(rows), formats
 
 
def summarize_scenario(model: ScenarioModel, sales: str) -> Dict[str, Any]:
    """Headline current-versus-scenario numbers for one sales basis."""
    cell = model.cell(sales)
    return {
        "Sales": SALES_LABEL[sales],
        "Current Run Rate": cell["current_rr"],
        "Scenario Run Rate": cell["scen_rr"],
        "Run Rate Change %": cell["rr_change_pct"],
        "Jan Achievement": cell["jan_amount"],
        "Jan Achievement %": cell["jan_pct"],
        "Feb-Mar Run Rate": cell["feb_mar_rr"],
        "Current March Projection %": cell["current_march_pct"],
        "Scenario March Achievement": cell["march_amount"],
        "Scenario March Achievement %": cell["march_pct"],
        "Incremental Sales": cell["incremental_sales"],
    }
 
 
def build_comparison(model: ScenarioModel) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """Current versus selected scenario, for Gross Sales and Net Sales."""
    rows = [summarize_scenario(model, sales) for sales in SALES_TYPES]
    formats = {
        "Sales": "txt", "Current Run Rate": "cr", "Scenario Run Rate": "cr",
        "Run Rate Change %": "pct_signed", "Jan Achievement": "cr",
        "Jan Achievement %": "pct", "Feb-Mar Run Rate": "cr",
        "Current March Projection %": "pct", "Scenario March Achievement": "cr",
        "Scenario March Achievement %": "pct", "Incremental Sales": "cr_signed",
    }
    return pd.DataFrame(rows), formats
 
 
def build_revenue_impact(model: ScenarioModel, basis: str) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """Revenue by asset class - never a blended rate."""
    bundle = revenue_bundle(model, basis)
    rows = []
    for asset in ASSETS:
        rows.append({
            "Asset Class": asset,
            "Scenario Sales": bundle["scenario"]["sales_by_asset"][asset],
            "Revenue Rate": f"{REVENUE_BPS[asset]:.0f} bps",
            "Scenario Revenue": bundle["scenario"]["by_asset"][asset],
            "Baseline Revenue": bundle["baseline"]["by_asset"][asset],
            "Incremental Revenue": bundle["incremental"]["by_asset"][asset],
            "Revenue Contribution %": bundle["incremental"]["contribution"][asset],
        })
    rows.append({
        "Asset Class": "Total",
        "Scenario Sales": sum(_z(v) for v in bundle["scenario"]["sales_by_asset"].values()),
        "Revenue Rate": "\u2014",
        "Scenario Revenue": bundle["scenario"]["total"],
        "Baseline Revenue": bundle["baseline"]["total"],
        "Incremental Revenue": bundle["incremental"]["total"],
        "Revenue Contribution %": 1.0 if bundle["scenario"]["total"] else None,
    })
    formats = {
        "Asset Class": "txt", "Scenario Sales": "cr", "Revenue Rate": "txt",
        "Scenario Revenue": "cr1", "Baseline Revenue": "cr1",
        "Incremental Revenue": "cr1_signed", "Revenue Contribution %": "pct",
    }
    return pd.DataFrame(rows), formats
 
 
def build_vertical_summary(model: ScenarioModel) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """Retail, DHNI and VRM, for Gross Sales and Net Sales."""
    rows = []
    for vertical in model.available_verticals():
        for sales in SALES_TYPES:
            cell = model.cell(sales, vertical=vertical)
            bundle = revenue_bundle(model, sales, vertical=vertical)
            rows.append({
                "Vertical": vertical,
                "Sales": SALES_LABEL[sales],
                "FY Target": cell["fy_target"],
                "YTD Achievement": cell["ytd_ach"],
                "Target Achieved %": cell["ytd_ach_pct"],
                "Current Run Rate": cell["current_rr"],
                "Scenario Run Rate": cell["scen_rr"],
                "Run Rate Change %": cell["rr_change_pct"],
                "Current March Projection %": cell["current_march_pct"],
                "Scenario Milestone %": cell["milestone_pct"],
                "Scenario March Projection %": cell["march_pct"],
                "Scenario Revenue": bundle["scenario"]["total"],
                "Incremental Revenue": bundle["incremental"]["total"],
            })
    formats = {
        "Vertical": "txt", "Sales": "txt", "FY Target": "cr", "YTD Achievement": "cr",
        "Target Achieved %": "pct", "Current Run Rate": "cr", "Scenario Run Rate": "cr",
        "Run Rate Change %": "pct_signed", "Current March Projection %": "pct",
        "Scenario Milestone %": "pct", "Scenario March Projection %": "pct",
        "Scenario Revenue": "cr1", "Incremental Revenue": "cr1_signed",
    }
    return pd.DataFrame(rows), formats
 
 
def build_asset_breakdown(model: ScenarioModel, sales: str) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """Equity / Debt / Liquid split within each vertical."""
    rows = []
    include_dip = model.scenario_id in (3, 7)
    for vertical in model.available_verticals():
        for asset in ASSETS:
            cell = model.cell(sales, asset=asset, vertical=vertical)
            revenue = _z(cell["march_amount"]) * REVENUE_RATE[asset]
            baseline_revenue = _z(cell["current_march"]) * REVENUE_RATE[asset]
            row = {
                "Vertical": vertical,
                "Asset": asset,
                "FY Target": cell["fy_target"],
                "YTD Achievement": cell["ytd_ach"],
                "Target Achieved %": cell["ytd_ach_pct"],
                "Current Run Rate": cell["current_rr"],
                "Scenario Run Rate": cell["scen_rr"],
                "Run Rate Change %": cell["rr_change_pct"],
            }
            if include_dip:
                row["Feb-Mar Run Rate"] = cell["feb_mar_rr"]
            row.update({
                "Current March Projection %": cell["current_march_pct"],
                "Scenario Milestone %": cell["milestone_pct"],
                "Scenario March Projection %": cell["march_pct"],
                "Scenario Revenue": revenue,
                "Incremental Revenue": revenue - baseline_revenue,
            })
            rows.append(row)
    formats = {
        "Vertical": "txt", "Asset": "txt", "FY Target": "cr", "YTD Achievement": "cr",
        "Target Achieved %": "pct", "Current Run Rate": "cr", "Scenario Run Rate": "cr",
        "Run Rate Change %": "pct_signed", "Feb-Mar Run Rate": "cr",
        "Current March Projection %": "pct", "Scenario Milestone %": "pct",
        "Scenario March Projection %": "pct", "Scenario Revenue": "cr1",
        "Incremental Revenue": "cr1_signed",
    }
    return pd.DataFrame(rows), formats
 
 
def build_segment_scenario_analysis(
    model: ScenarioModel, sales: str
) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """Scenario 6 - differentiated performance by business segment."""
    rows = []
    for segment in model.available_segments():
        cell = model.cell(sales, segment=segment)
        rows.append({
            "Segment": segment,
            "FY Target": cell["fy_target"],
            "YTD Achievement": cell["ytd_ach"],
            "Current Run Rate": cell["current_rr"],
            "Current March Projection": cell["current_march"],
            "Scenario Achievement %": cell["milestone_pct"],
            "Scenario Target Amount": cell["march_required"],
            "Scenario Required Run Rate": cell["scen_rr"],
            "Run Rate Uplift %": cell["rr_change_pct"],
            "Incremental Amount": cell["incremental_sales"],
        })
    overall = model.cell(sales)
    rows.append({
        "Segment": "Overall",
        "FY Target": overall["fy_target"],
        "YTD Achievement": overall["ytd_ach"],
        "Current Run Rate": overall["current_rr"],
        "Current March Projection": overall["current_march"],
        "Scenario Achievement %": overall["march_pct"],
        "Scenario Target Amount": overall["march_required"],
        "Scenario Required Run Rate": overall["scen_rr"],
        "Run Rate Uplift %": overall["rr_change_pct"],
        "Incremental Amount": overall["incremental_sales"],
    })
    formats = {
        "Segment": "txt", "FY Target": "cr", "YTD Achievement": "cr",
        "Current Run Rate": "cr", "Current March Projection": "cr",
        "Scenario Achievement %": "pct", "Scenario Target Amount": "cr",
        "Scenario Required Run Rate": "cr", "Run Rate Uplift %": "pct_signed",
        "Incremental Amount": "cr_signed",
    }
    return pd.DataFrame(rows), formats
 
 
def build_momentum_analysis(cell: Dict[str, Any]) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """Scenario 7 - month-by-month momentum trajectory."""
    trajectory = cell.get("trajectory") or []
    current_rr = _num(cell.get("current_rr"))
    growth = cell.get("momentum_g")
    leakage = _z(cell.get("leakage"))
    fy_target = cell.get("fy_target")
    cumulative = _z(cell.get("ytd_ach"))
 
    rows = []
    previous = current_rr
    for position, month in enumerate(FUTURE_MONTHS):
        run_rate = trajectory[position] if position < len(trajectory) else None
        cumulative += _z(run_rate)
        if position < MONTHS_JUL_JAN:
            mom = growth if growth is not None else safe_div(run_rate, previous)
            if growth is None and mom is not None:
                mom = mom - 1.0
            phase = "Momentum build-up"
        else:
            mom = -leakage
            phase = "Feb-Mar leakage"
        rows.append({
            "Month": month,
            "Phase": phase,
            "Current Run Rate": current_rr,
            "Required Scenario Run Rate": run_rate,
            "MoM Growth": mom,
            "Cumulative Achievement": cumulative,
            "Achievement %": safe_div(cumulative, fy_target),
        })
        previous = run_rate
    formats = {
        "Month": "txt", "Phase": "txt", "Current Run Rate": "cr",
        "Required Scenario Run Rate": "cr", "MoM Growth": "pct_signed",
        "Cumulative Achievement": "cr", "Achievement %": "pct",
    }
    return pd.DataFrame(rows), formats
 
 
def build_momentum_by_group(
    model: ScenarioModel, sales: str, dimension: str
) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """Scenario 7 applied independently to asset classes or to verticals."""
    if dimension == "asset":
        keys = [("Asset", asset, {"asset": asset}) for asset in ASSETS]
        first_column = "Asset"
    else:
        keys = [("Vertical", v, {"vertical": v}) for v in model.available_verticals()]
        first_column = "Vertical"
 
    rows = []
    for _, name, filters in keys:
        cell = model.cell(sales, **filters)
        rows.append({
            first_column: name,
            "Current Run Rate": cell["current_rr"],
            "Required MoM Momentum": cell["momentum_g"],
            "January Achievement": cell["jan_amount"],
            "January Achievement %": cell["jan_pct"],
            "Feb-Mar Leakage": cell.get("leakage"),
            "March Achievement": cell["march_amount"],
            "March Achievement %": cell["march_pct"],
            "Headroom / Shortfall": cell["headroom_amt"],
        })
    formats = {
        first_column: "txt", "Current Run Rate": "cr",
        "Required MoM Momentum": "pct_signed", "January Achievement": "cr",
        "January Achievement %": "pct", "Feb-Mar Leakage": "pct",
        "March Achievement": "cr", "March Achievement %": "pct",
        "Headroom / Shortfall": "cr_signed",
    }
    return pd.DataFrame(rows), formats
 
 
def build_monthly_revenue(model: ScenarioModel, basis: str) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """Scenario 7 - monthly revenue implied by the momentum trajectory."""
    trajectories = {
        asset: (model.cell(basis, asset=asset).get("trajectory") or [])
        for asset in ASSETS
    }
    rows = []
    for position, month in enumerate(FUTURE_MONTHS):
        row = {"Month": month}
        total = 0.0
        for asset in ASSETS:
            series = trajectories[asset]
            sales_amount = series[position] if position < len(series) else 0.0
            revenue = _z(sales_amount) * REVENUE_RATE[asset]
            row[f"{asset} Revenue"] = revenue
            total += revenue
        row["Total Revenue"] = total
        rows.append(row)
    formats = {"Month": "txt", "Total Revenue": "cr1"}
    for asset in ASSETS:
        formats[f"{asset} Revenue"] = "cr1"
    return pd.DataFrame(rows), formats
 
 
def build_leakage_sensitivity(
    model: ScenarioModel, basis: str
) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """Scenario 7 - how the March outcome moves with the leakage assumption."""
    subset = filter_grid(model.grid, sales=basis)
    fy_target = subset["fy_target"].sum()
    ytd_target = subset["ytd_target"].sum()
    ytd_ach = subset["ytd_ach"].sum()
    rows = []
    for leakage in (0.0, 0.10, 0.20, 0.30):
        params = dict(model.params)
        params["leakage"] = leakage
        cell = calculate_scenario_7(fy_target, ytd_target, ytd_ach, params)
        rows.append({
            "Feb-Mar Leakage": leakage,
            "Required MoM Momentum": cell["momentum_g"],
            "January Achievement %": cell["jan_pct"],
            "March Achievement %": cell["march_pct"],
            "Headroom / Shortfall": cell["headroom_amt"],
        })
    formats = {
        "Feb-Mar Leakage": "pct", "Required MoM Momentum": "pct_signed",
        "January Achievement %": "pct", "March Achievement %": "pct",
        "Headroom / Shortfall": "cr_signed",
    }
    return pd.DataFrame(rows), formats
 
 
def build_channel_scenario_analysis(
    model: "ScenarioModel", basis: str
) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """Scenario 8/9 - one row per planning channel."""
    rows = []
    frame = model.scenario_grid
    if frame is None or "Channel" not in getattr(frame, "columns", []):
        return pd.DataFrame(), {}
    for channel in CHANNELS:
        subset = frame[(frame["Sales"] == basis) & (frame["Channel"] == channel)]
        if subset.empty:
            continue
        cell = summarize_cells(subset)
        growth, jan_target, mar_target, leakage = _s8_channel_params(model.params, channel)
        if model.scenario_id == 9 and "optimized_growth" in subset.columns:
            solved = subset["optimized_growth"].dropna()
            if not solved.empty:
                growth = float(solved.max())
        rows.append({
            "Channel": channel, "MoM Growth": growth,
            "Jan 2027 Target": jan_target, "Jan Achievement": cell.get("jan_pct"),
            "Jan Gap / Headroom": cell.get("jan_buffer"),
            "Mar 2027 Target": mar_target, "Mar Achievement": cell.get("march_pct"),
            "Mar Gap / Headroom": cell.get("headroom_amt"),
            "Current Run Rate": cell.get("current_rr"),
            "Jan Exit Run Rate": cell.get("scen_rr"),
            "March Incremental Sales": cell.get("incremental_sales"),
        })
    formats = {
        "Channel": "txt", "MoM Growth": "pct_signed", "Jan 2027 Target": "pct",
        "Jan Achievement": "pct", "Jan Gap / Headroom": "cr_signed",
        "Mar 2027 Target": "pct", "Mar Achievement": "pct", "Mar Gap / Headroom": "cr_signed",
        "Current Run Rate": "cr", "Jan Exit Run Rate": "cr",
        "March Incremental Sales": "cr_signed",
    }
    return pd.DataFrame(rows), formats
 
 
def build_scenario_guide(model: ScenarioModel, basis: str) -> pd.DataFrame:
    rows = []
    for scenario_id in SCENARIO_ORDER:
        meta = SCENARIOS[scenario_id]
        rows.append({
            "Scenario": meta["label"],
            "Description": meta["explanation"],
            "Milestone": meta["milestone"],
            "Selected": "Yes" if scenario_id == model.scenario_id else "",
        })
    rows.append({
        "Scenario": "Revenue methodology",
        "Description": (
            "Revenue is estimated at asset-class level using 60 bps for Equity, 20 bps for Debt "
            f"and 10 bps for Liquid, applied to {SALES_LABEL[basis]} only so that Gross Sales and "
            "Net Sales revenue are never double counted."
        ),
        "Milestone": "",
        "Selected": "",
    })
    rows.append({
        "Scenario": "Timeline assumption",
        "Description": (
            "April, May and June are complete. Three months completed, nine months remaining "
            "(July-January is seven months, February-March is two months). The current run rate "
            "is YTD achievement divided by three."
        ),
        "Milestone": "",
        "Selected": "",
    })
    if model.scenario_id == 3:
        rows.append({
            "Scenario": "Scenario 3 setting",
            "Description": f"Feb-Mar run-rate dip: {fmt_pct(model.params.get('dip', S3_DEFAULT_DIP))}",
            "Milestone": "", "Selected": "",
        })
    if model.scenario_id == 7:
        rows.append({
            "Scenario": "Scenario 7 settings",
            "Description": (
                f"January target: {fmt_pct(model.params.get('jan_target', S7_DEFAULT_JAN_TARGET))} · "
                f"March target: {fmt_pct(model.params.get('mar_target', S7_DEFAULT_MAR_TARGET))} · "
                f"Feb-Mar leakage: {fmt_pct(model.params.get('leakage', S7_DEFAULT_LEAKAGE))}"
            ),
            "Milestone": "", "Selected": "",
        })
    return pd.DataFrame(rows)
 
 
def scenario_default_params(scenario_id: int) -> Dict[str, Any]:
    """Defaults exactly from the existing scenario engine configuration."""
    return {
        "dip": S3_DEFAULT_DIP,
        "jan_target": S7_DEFAULT_JAN_TARGET,
        "mar_target": S7_DEFAULT_MAR_TARGET,
        "leakage": S7_DEFAULT_LEAKAGE,
        "channel_growth": dict(S8_DEFAULT_GROWTH),
        "channel_jan_target": dict(S8_DEFAULT_JAN_TARGET),
        "channel_mar_target": dict(S8_DEFAULT_MAR_TARGET),
        "optimizer_target": 1.20,
        "channel_mapping": {},
    }
 
 
def build_all_scenario_matrix(
    filtered_grid: pd.DataFrame,
    selected_scenario_id: int,
    selected_params: Dict[str, Any],
    sales_key: str,
    asset: str = "All",
) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """Calculate Scenarios 1-9 on the same selected cut."""
    rows: List[Dict[str, Any]] = []
    formats = {
        "Scenario": "txt", "Strategy": "txt", "FY Target": "cr",
        "Current YTD": "cr", "Current March Projection": "cr",
        "Scenario March Estimate": "cr", "Scenario March %": "pct",
        "Required Run Rate": "cr", "Run Rate Change %": "pct_signed",
        "Incremental Sales": "cr_signed", "Headroom / Gap": "cr_signed",
    }
    if filtered_grid.empty:
        return pd.DataFrame(), formats
 
    for sid in SCENARIO_ORDER:
        params = scenario_default_params(sid)
        if sid == selected_scenario_id:
            # The live controls are authoritative for the selected scenario.
            params.update(selected_params)
        try:
            model = ScenarioModel(sid, filtered_grid, params)
            cell = model.cell(sales_key, asset=None if asset == "All" else asset)
            rows.append({
                "Scenario": f"{sid:02d} · {SCENARIOS[sid]['short']}",
                "Strategy": SCENARIOS[sid]["name"],
                "FY Target": cell.get("fy_target"),
                "Current YTD": cell.get("ytd_ach"),
                "Current March Projection": cell.get("current_march"),
                "Scenario March Estimate": cell.get("march_amount"),
                "Scenario March %": cell.get("march_pct"),
                "Required Run Rate": cell.get("scen_rr"),
                "Run Rate Change %": cell.get("rr_change_pct"),
                "Incremental Sales": cell.get("incremental_sales"),
                "Headroom / Gap": cell.get("headroom_amt"),
            })
        except Exception:  # pragma: no cover - one scenario must never break the view
            rows.append({
                "Scenario": f"{sid:02d} · {SCENARIOS[sid]['short']}",
                "Strategy": SCENARIOS[sid]["name"],
            })
    return pd.DataFrame(rows), formats
 
 
# =============================================================================
# 12. EXCEL EXPORT
# =============================================================================
 
def _round_frame(frame: pd.DataFrame) -> pd.DataFrame:
    out = frame.copy()
    for column in out.columns:
        if pd.api.types.is_numeric_dtype(out[column]):
            out[column] = out[column].astype(float).round(4)
    return out
 
 
def make_export_excel(model: ScenarioModel, basis: str) -> bytes:
    """Build the management export workbook for the selected scenario."""
    sheets: List[Tuple[str, pd.DataFrame]] = []
    sheets.append(("Scenario Guide", build_scenario_guide(model, basis)))
    sheets.append(("Current Baseline", build_current_overview(model)[0]))
    sheets.append(("Current vs Scenario", build_comparison(model)[0]))
    sheets.append(("Revenue Impact", build_revenue_impact(model, basis)[0]))
    sheets.append(("Retail-DHNI-VRM Summary", build_vertical_summary(model)[0]))
    sheets.append(("Gross Sales Breakdown", build_asset_breakdown(model, "GS")[0]))
    sheets.append(("Net Sales Breakdown", build_asset_breakdown(model, "NS")[0]))
 
    segment_model = model if model.scenario_id == 6 else ScenarioModel(6, model.grid, model.params)
    segment_frames = []
    for sales in SALES_TYPES:
        frame = build_segment_scenario_analysis(segment_model, sales)[0]
        frame.insert(0, "Sales", SALES_LABEL[sales])
        segment_frames.append(frame)
    sheets.append(("Scenario 6 Segments", pd.concat(segment_frames, ignore_index=True)))
 
    momentum_model = model if model.scenario_id == 7 else ScenarioModel(7, model.grid, model.params)
    momentum_frames = []
    for sales in SALES_TYPES:
        overall = momentum_model.cell(sales)
        frame = build_momentum_analysis(overall)[0]
        frame.insert(0, "Sales", SALES_LABEL[sales])
        momentum_frames.append(frame)
    for sales in SALES_TYPES:
        for dimension in ("asset", "vertical"):
            frame = build_momentum_by_group(momentum_model, sales, dimension)[0]
            frame.insert(0, "Sales", SALES_LABEL[sales])
            momentum_frames.append(frame)
    sheets.append(("Scenario 7 Momentum", pd.concat(momentum_frames, ignore_index=True)))
    sheets.append(("S7 Monthly Revenue", build_monthly_revenue(momentum_model, basis)[0]))
 
    if model.scenario_id in (8, 9):
        channel_frame = build_channel_scenario_analysis(model, basis)[0]
        if not channel_frame.empty:
            sheets.append((f"Scenario {model.scenario_id} Channels", channel_frame))
 
    all_scenarios = build_all_scenario_matrix(model.grid, model.scenario_id, model.params, basis)[0]
    if not all_scenarios.empty:
        sheets.append(("All Scenarios", all_scenarios))
 
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for name, frame in sheets:
            _round_frame(frame).to_excel(writer, sheet_name=name[:31], index=False)
        _style_workbook(writer)
    return buffer.getvalue()
 
 
def _style_workbook(writer: Any) -> None:
    """Bold headers, frozen top row and sensible column widths."""
    try:
        from openpyxl.styles import Alignment, Font, PatternFill
    except Exception:  # pragma: no cover - openpyxl always present in the stack
        return
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="12141A")
    for worksheet in writer.book.worksheets:
        worksheet.freeze_panes = "A2"
        for cell in worksheet[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(vertical="center", wrap_text=True)
        for column_cells in worksheet.columns:
            longest = 0
            letter = column_cells[0].column_letter
            for cell in column_cells:
                value = "" if cell.value is None else str(cell.value)
                longest = max(longest, min(len(value), 60))
            worksheet.column_dimensions[letter].width = max(12, min(longest + 3, 62))
 
 
# =============================================================================
# 13. RM PERFORMANCE ANALYTICS (segmentation & contribution)
# =============================================================================
 
ACHIEVEMENT_BANDS: List[Tuple[str, float, Optional[float]]] = [
    ("100% and above", 1.00, None),
    ("90% - 100%", 0.90, 1.00),
    ("75% - 90%", 0.75, 0.90),
    ("50% - 75%", 0.50, 0.75),
    ("30% - 50%", 0.30, 0.50),
    ("Less than 30%", float("-inf"), 0.30),
]
ACHIEVEMENT_BAND_ORDER: List[str] = [item[0] for item in ACHIEVEMENT_BANDS]
 
 
def achievement_band(value: Any) -> str:
    """Map YTD-target achievement to the management bands used for RMs."""
    ratio = _num(value)
    # Undefined achievement is treated and displayed as 0.
    ratio = 0.0 if ratio is None else ratio
 
    if ratio >= 1.00:
        return "100% and above"
    if ratio >= 0.90:
        return "90% - 100%"
    if ratio >= 0.75:
        return "75% - 90%"
    if ratio >= 0.50:
        return "50% - 75%"
    if ratio >= 0.30:
        return "30% - 50%"
    return "Less than 30%"
 
 
def _rm_identity_columns(records: pd.DataFrame) -> List[str]:
    preferred = [
        "Employee Name", "Emp Code", "ADID", "ZONE", "REGION",
        "EM City", "MKT TYPE", "Type", "Status",
    ]
    return [column for column in preferred if column in records.columns]
 
 
def build_rm_performance_detail(
    records: pd.DataFrame,
    vertical: str,
    sales: str,
    final_target: Optional[float] = None,
) -> pd.DataFrame:
    """
    Build one row per RM.
 
    Banding is based on overall YTD achievement / overall YTD target across
    Equity + Debt + Liquid. Run-rate projection annualises the first three
    completed months.
    """
    subset = records.loc[records["Vertical"] == vertical].copy()
    if subset.empty:
        return pd.DataFrame()
 
    identity = _rm_identity_columns(subset)
    out = subset[identity].copy()
 
    fy_cols = [f"{sales}_{asset}_fy" for asset in ASSETS]
    ytd_target_cols = [f"{sales}_{asset}_ytd_tgt" for asset in ASSETS]
    ach_cols = [f"{sales}_{asset}_ach" for asset in ASSETS]
 
    for columns in (fy_cols, ytd_target_cols, ach_cols):
        for column in columns:
            if column not in subset.columns:
                subset[column] = 0.0
 
    fy_matrix = subset[fy_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0)
    ytd_target_matrix = subset[ytd_target_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0)
    ach_matrix = subset[ach_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0)
 
    out["FY Target"] = fy_matrix.sum(axis=1)
    out["YTD Target"] = ytd_target_matrix.sum(axis=1)
    out["YTD Achievement"] = ach_matrix.sum(axis=1)
 
    out["YTD Achievement %"] = np.where(
        out["YTD Target"] > 0,
        out["YTD Achievement"] / out["YTD Target"],
        0.0,
    )
 
    out["Achievement Category"] = out["YTD Achievement %"].map(achievement_band)
    out["Current Run Rate"] = out["YTD Achievement"] / max(MONTHS_COMPLETED, 1)
    out["Estimated FY @ Current RR"] = out["Current Run Rate"] * 12.0
    out["Projected FY Achievement %"] = np.where(
        out["FY Target"] > 0,
        out["Estimated FY @ Current RR"] / out["FY Target"],
        0.0,
    )
 
    denominator = _num(final_target)
    if denominator is None or denominator <= 0:
        denominator = _num(out["FY Target"].sum()) or 0.0
 
    out["Contribution to Overall Target %"] = np.where(
        denominator > 0,
        out["Estimated FY @ Current RR"] / denominator,
        0.0,
    )
 
    # Asset-level achieved percentages are useful when drilling into an RM.
    for asset in ASSETS:
        fy = pd.to_numeric(subset[f"{sales}_{asset}_fy"], errors="coerce").fillna(0.0)
        ytd_target = pd.to_numeric(subset[f"{sales}_{asset}_ytd_tgt"], errors="coerce").fillna(0.0)
        ach = pd.to_numeric(subset[f"{sales}_{asset}_ach"], errors="coerce").fillna(0.0)
 
        out[f"{asset} YTD %"] = np.where(ytd_target > 0, ach / ytd_target, 0.0)
        out[f"{asset} Current RR"] = ach / max(MONTHS_COMPLETED, 1)
        out[f"{asset} Projected FY %"] = np.where(
            fy > 0,
            (ach / max(MONTHS_COMPLETED, 1) * 12.0) / fy,
            0.0,
        )
 
    out = out.sort_values(
        ["YTD Achievement %", "YTD Achievement"],
        ascending=[False, False],
    ).reset_index(drop=True)
 
    return out
 
 
def build_category_contribution(
    detail: pd.DataFrame,
    final_target: Optional[float] = None,
) -> pd.DataFrame:
    """
    Aggregate RM bands and quantify how many percentage points each band is
    projected to contribute to the FINAL FY target at the current run rate.
    """
    if detail.empty:
        return pd.DataFrame()
 
    target_denominator = _num(final_target)
    if target_denominator is None or target_denominator <= 0:
        target_denominator = _num(detail["FY Target"].sum()) or 0.0
 
    total_projected = _num(detail["Estimated FY @ Current RR"].sum()) or 0.0
 
    work = detail.copy()
    work["_rm"] = 1
 
    grouped = (
        work.groupby("Achievement Category", dropna=False)
        .agg(
            **{
                "RM Count": ("_rm", "sum"),
                "FY Target": ("FY Target", "sum"),
                "YTD Target": ("YTD Target", "sum"),
                "YTD Achievement": ("YTD Achievement", "sum"),
                "Current Run Rate": ("Current Run Rate", "sum"),
                "Estimated FY @ Current RR": ("Estimated FY @ Current RR", "sum"),
            }
        )
        .reindex(ACHIEVEMENT_BAND_ORDER, fill_value=0)
        .reset_index()
    )
 
    grouped["Current YTD Achievement %"] = np.where(
        grouped["YTD Target"] > 0,
        grouped["YTD Achievement"] / grouped["YTD Target"],
        0.0,
    )
    grouped["Category Projected FY %"] = np.where(
        grouped["FY Target"] > 0,
        grouped["Estimated FY @ Current RR"] / grouped["FY Target"],
        0.0,
    )
    grouped["Contribution to Overall Target %"] = np.where(
        target_denominator > 0,
        grouped["Estimated FY @ Current RR"] / target_denominator,
        0.0,
    )
    grouped["Share of Projected Sales %"] = np.where(
        total_projected != 0,
        grouped["Estimated FY @ Current RR"] / total_projected,
        0.0,
    )
 
    return grouped
 
 
def _final_vertical_target(
    final_metrics: Dict[str, Any],
    sales: str,
    vertical: str,
) -> Optional[float]:
    frame = final_metrics.get(sales)
    if not isinstance(frame, pd.DataFrame) or frame.empty or vertical not in frame.index:
        return None
    return _num(frame.loc[vertical].get("FY27 Target"))
 
 
def _final_vertical_ytd(
    final_metrics: Dict[str, Any],
    sales: str,
    vertical: str,
) -> Optional[float]:
    frame = final_metrics.get(sales)
    if not isinstance(frame, pd.DataFrame) or frame.empty or vertical not in frame.index:
        return None
    return _num(frame.loc[vertical].get("YTD"))
 
 
def make_rm_segmentation_export(
    records: pd.DataFrame,
    final_metrics: Dict[str, Any],
) -> bytes:
    """Downloadable workbook covering every vertical and both sales bases."""
    output = io.BytesIO()
 
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for vertical in VERTICALS:
            for sales in SALES_TYPES:
                final_target = _final_vertical_target(final_metrics, sales, vertical)
                detail = build_rm_performance_detail(records, vertical, sales, final_target)
                contribution = build_category_contribution(detail, final_target)
                stars = (
                    detail.sort_values(
                        ["YTD Achievement %", "YTD Achievement"],
                        ascending=[False, False],
                    ).head(10).copy()
                    if not detail.empty else pd.DataFrame()
                )
 
                prefix = f"{vertical}-{sales}"
                if detail.empty:
                    detail = pd.DataFrame({"Note": [f"No {vertical} records for {SALES_LABEL[sales]}."]})
                if contribution.empty:
                    contribution = pd.DataFrame({"Note": ["No banded contribution available."]})
                if stars.empty:
                    stars = pd.DataFrame({"Note": ["No ranked RMs available."]})
 
                detail.to_excel(writer, sheet_name=f"{prefix}-RM"[:31], index=False)
                contribution.to_excel(writer, sheet_name=f"{prefix}-Bands"[:31], index=False)
                stars.to_excel(writer, sheet_name=f"{prefix}-Stars"[:31], index=False)
 
        try:
            from openpyxl.styles import Font
        except Exception:  # pragma: no cover
            Font = None
 
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            if Font is not None:
                for cell in ws[1]:
                    cell.font = Font(bold=True)
            for cells in ws.columns:
                width = max((len(str(cell.value or "")) for cell in cells[:150]), default=10)
                ws.column_dimensions[cells[0].column_letter].width = min(max(width + 2, 12), 36)
 
    output.seek(0)
    return output.getvalue()
 
 
# =============================================================================
# 14. DESIGN SYSTEM - LIQUID GLASS THEME, COMPONENTS & PLOTLY TEMPLATE
# =============================================================================
 
INK = "#F5F5F7"
INK_SOFT = "#A1A4AD"
INK_MUTED = "#6D7079"
GOLD = "#D8B76A"
GOLD_SOFT = "#C9AA65"
GREEN = "#63D99A"
RED = "#FF6B6B"
GRID_LINE = "rgba(255,255,255,0.07)"
FONT_STACK = 'Inter, -apple-system, BlinkMacSystemFont, "SF Pro Display", "Helvetica Neue", sans-serif'
 
GLASS_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
 
:root {
    --background: #050608;
    --glass-1: rgba(255,255,255,0.075);
    --glass-2: rgba(255,255,255,0.045);
    --glass-3: rgba(255,255,255,0.025);
    --border: rgba(255,255,255,0.11);
    --border-strong: rgba(255,255,255,0.18);
    --text: #F5F5F7;
    --secondary: #A1A4AD;
    --tertiary: #6D7079;
    --gold: #D8B76A;
    --gold-soft: #C9AA65;
    --gold-glow: rgba(216,183,106,0.16);
    --green: #63D99A;
    --red: #FF6B6B;
    --radius-lg: 24px;
    --radius-md: 18px;
    --radius-sm: 12px;
}
 
/* ---------- canvas ---------- */
.stApp, [data-testid="stAppViewContainer"] {
    background:
        radial-gradient(circle at 10% 0%, rgba(255,255,255,0.035), transparent 30%),
        radial-gradient(circle at 90% 8%, rgba(214,179,106,0.055), transparent 26%),
        radial-gradient(circle at 50% 100%, rgba(90,90,120,0.045), transparent 36%),
        #050608;
    color: var(--text);
    font-family: Inter, -apple-system, BlinkMacSystemFont, "SF Pro Display", "Helvetica Neue", sans-serif;
}
[data-testid="stHeader"] { background: transparent; }
#MainMenu, footer { visibility: hidden; }
 
.block-container, [data-testid="stMainBlockContainer"] {
    padding-top: 1.4rem;
    padding-bottom: 4rem;
    max-width: 1560px;
}
 
.stApp, .stApp p, .stApp span, .stApp li, .stApp label,
.stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5 {
    color: var(--text);
    font-family: Inter, -apple-system, BlinkMacSystemFont, "SF Pro Display", "Helvetica Neue", sans-serif;
    -webkit-font-smoothing: antialiased;
}
a, a:visited { color: var(--gold-soft); }
 
/* ---------- glass surfaces ---------- */
.glass-card, .glass-panel, .glass-kpi {
    background: linear-gradient(135deg, rgba(255,255,255,0.085), rgba(255,255,255,0.035));
    border: 1px solid var(--border);
    border-radius: var(--radius-lg);
    backdrop-filter: blur(24px) saturate(140%);
    -webkit-backdrop-filter: blur(24px) saturate(140%);
    box-shadow: 0 20px 60px rgba(0,0,0,0.28), inset 0 1px 0 rgba(255,255,255,0.08);
    transition: border-color 180ms ease, transform 180ms ease, box-shadow 180ms ease;
}
.glass-card:hover, .glass-kpi:hover {
    border-color: var(--border-strong);
    transform: translateY(-1px);
}
.glass-panel {
    background: linear-gradient(135deg, rgba(255,255,255,0.05), rgba(255,255,255,0.022));
    border-radius: var(--radius-md);
    box-shadow: 0 14px 40px rgba(0,0,0,0.22), inset 0 1px 0 rgba(255,255,255,0.05);
    padding: 18px 20px;
}
.glass-card { padding: 22px 24px; }
.glass-kpi { padding: 16px 18px; border-radius: var(--radius-md); }
 
/* ---------- executive header ---------- */
.exec-header {
    display: flex;
    align-items: flex-end;
    justify-content: space-between;
    gap: 24px;
    flex-wrap: wrap;
    padding: 6px 2px 18px 2px;
    border-bottom: 1px solid var(--border);
    margin-bottom: 26px;
}
.exec-mark {
    font-size: 0.66rem;
    letter-spacing: 0.30em;
    text-transform: uppercase;
    color: var(--gold);
    font-weight: 600;
    margin-bottom: 10px;
}
.exec-title {
    font-size: 2.15rem;
    font-weight: 600;
    letter-spacing: -0.028em;
    line-height: 1.05;
    margin: 0;
}
.exec-sub {
    color: var(--secondary);
    font-size: 0.9rem;
    font-weight: 400;
    margin-top: 8px;
    letter-spacing: 0.005em;
}
.exec-status { display: flex; gap: 10px; flex-wrap: wrap; align-items: stretch; }
.status-chip {
    background: var(--glass-2);
    border: 1px solid var(--border);
    border-radius: 14px;
    padding: 9px 14px;
    min-width: 104px;
    backdrop-filter: blur(18px);
    -webkit-backdrop-filter: blur(18px);
}
.status-chip .k {
    font-size: 0.6rem;
    letter-spacing: 0.16em;
    text-transform: uppercase;
    color: var(--tertiary);
    font-weight: 600;
}
.status-chip .v {
    font-size: 0.92rem;
    font-weight: 600;
    margin-top: 3px;
    color: var(--text);
}
.status-live .v { color: var(--green); }
.status-dot {
    display: inline-block;
    width: 7px; height: 7px;
    border-radius: 50%;
    background: var(--green);
    margin-right: 7px;
    box-shadow: 0 0 0 3px rgba(99,217,154,0.14);
}
 
/* ---------- section headers ---------- */
.glass-section {
    display: flex;
    align-items: baseline;
    gap: 14px;
    margin: 40px 0 6px 0;
    padding-top: 6px;
}
.glass-section .idx {
    font-size: 0.7rem;
    font-weight: 600;
    color: var(--gold);
    letter-spacing: 0.18em;
    font-variant-numeric: tabular-nums;
}
.glass-section .ttl {
    font-size: 1.22rem;
    font-weight: 600;
    letter-spacing: -0.015em;
    color: var(--text);
}
.glass-section .sub {
    font-size: 0.82rem;
    color: var(--tertiary);
    font-weight: 400;
}
.section-rule {
    height: 1px;
    background: linear-gradient(90deg, rgba(216,183,106,0.42), rgba(255,255,255,0.05) 42%, transparent);
    margin: 10px 0 18px 0;
}
 
/* ---------- metric typography ---------- */
.metric-label {
    font-size: 0.63rem;
    letter-spacing: 0.17em;
    text-transform: uppercase;
    color: var(--tertiary);
    font-weight: 600;
}
.metric-hero {
    font-size: 2.5rem;
    font-weight: 600;
    letter-spacing: -0.035em;
    line-height: 1.03;
    margin: 10px 0 2px 0;
    font-variant-numeric: tabular-nums;
}
.metric-hero.gold { color: var(--gold); }
.metric-value {
    font-size: 1.42rem;
    font-weight: 600;
    letter-spacing: -0.02em;
    margin-top: 6px;
    font-variant-numeric: tabular-nums;
}
.metric-secondary {
    font-size: 0.78rem;
    color: var(--secondary);
    margin-top: 5px;
    line-height: 1.5;
}
.metric-delta {
    font-size: 0.78rem;
    font-weight: 600;
    margin-top: 6px;
    font-variant-numeric: tabular-nums;
}
.metric-delta.pos { color: var(--green); }
.metric-delta.neg { color: var(--red); }
.metric-delta.flat { color: var(--secondary); font-weight: 500; }
.metric-delta.gold { color: var(--gold); }
 
.kpi-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 12px;
    margin: 4px 0 8px 0;
}
.hero-grid {
    display: grid;
    grid-template-columns: 1.15fr 1fr 1fr;
    gap: 16px;
    margin-bottom: 8px;
}
.hero-grid.no-aum {
    grid-template-columns: repeat(2, minmax(0, 1fr));
}
.trio-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
    gap: 14px;
}
 
/* ---------- text-overlap safety (layout/colors unchanged) ---------- */
.glass-card, .glass-panel, .glass-kpi,
.kpi-grid > *, .hero-grid > *, .trio-grid > *,
.kpi-head, .kpi-row, .progress-legend, .exec-header > *, .exec-status > * {
    min-width: 0;
}
.metric-label, .metric-hero, .metric-value, .metric-secondary, .metric-delta,
.kpi-row .k, .kpi-row .v, .scenario-hero .title, .scenario-hero .thesis,
.scenario-hero .detail, .scenario-hero .milestone, .glass-note, .glass-callout,
.exec-title, .exec-sub, .status-chip .v {
    overflow-wrap: anywhere;
    word-break: normal;
}
.kpi-row .k { flex: 1 1 auto; }
.kpi-row .v { flex: 0 1 52%; text-align: right; }
.progress-legend { gap: 10px; flex-wrap: wrap; }
.progress-legend span:last-child { margin-left: auto; text-align: right; }
div[role="radiogroup"] > label { max-width: 100%; }
div[role="radiogroup"] > label div[data-testid="stMarkdownContainer"] p {
    white-space: normal !important;
    overflow-wrap: anywhere;
}
.stTabs [data-baseweb="tab"] {
    min-width: 0;
    white-space: normal;
    text-align: center;
}

/* ---------- Streamlit text/icon overlap hardening ---------- */
/*
   IMPORTANT: On some corporate machines the Material Symbols web-font is
   blocked. Streamlit then paints the icon name itself (for example
   "keyboard_arrow_right") inside sidebar expanders. The long fallback text
   sits on top of labels such as "Segment mapping" / "Channel mapping".

   This patch does not redesign the sidebar. It removes only the broken icon
   glyph and gives the expander label its own protected column.
*/
[data-testid="stExpander"] summary {
    min-width: 0 !important;
}

/* Catch current and older Streamlit Material-icon wrappers. */
[data-testid="stExpander"] summary [data-testid="stExpanderToggleIcon"],
[data-testid="stExpander"] summary [data-testid="stIconMaterial"],
[data-testid="stExpander"] summary [aria-hidden="true"],
[data-testid="stExpander"] summary span[class*="material-symbol"],
[data-testid="stExpander"] summary span[class*="material-icon"] {
    font-size: 0 !important;
    line-height: 0 !important;
    letter-spacing: 0 !important;
    text-indent: -9999px !important;
    color: transparent !important;
    overflow: hidden !important;
    white-space: nowrap !important;
    max-width: 18px !important;
}

/* Strong fallback for the exact sidebar overlap shown in the screenshot.
   Streamlit places the expander toggle before the label. If the toggle loses
   its icon font, constrain that first control to a tiny fixed slot so its
   literal fallback text can never cross into the label. */
[data-testid="stSidebar"] [data-testid="stExpander"] summary {
    display: grid !important;
    grid-template-columns: 18px minmax(0, 1fr) !important;
    align-items: center !important;
    column-gap: 8px !important;
    width: 100% !important;
    overflow: hidden !important;
}
[data-testid="stSidebar"] [data-testid="stExpander"] summary > * {
    min-width: 0 !important;
}
[data-testid="stSidebar"] [data-testid="stExpander"] summary > :first-child {
    width: 18px !important;
    min-width: 18px !important;
    max-width: 18px !important;
    height: 18px !important;
    overflow: hidden !important;
    font-size: 0 !important;
    line-height: 0 !important;
    letter-spacing: 0 !important;
    text-indent: -9999px !important;
    color: transparent !important;
    white-space: nowrap !important;
}
[data-testid="stSidebar"] [data-testid="stExpander"] summary > :first-child * {
    font-size: 0 !important;
    line-height: 0 !important;
    letter-spacing: 0 !important;
    text-indent: -9999px !important;
    color: transparent !important;
    overflow: hidden !important;
    white-space: nowrap !important;
}
/* Replace the missing Material glyph with a font-independent chevron. */
[data-testid="stSidebar"] [data-testid="stExpander"] summary > :first-child::after {
    content: "›";
    display: block !important;
    width: 18px !important;
    height: 18px !important;
    text-align: center !important;
    text-indent: 0 !important;
    font-family: Arial, sans-serif !important;
    font-size: 17px !important;
    font-weight: 600 !important;
    line-height: 18px !important;
    color: var(--secondary) !important;
    transform-origin: 50% 50%;
}
[data-testid="stSidebar"] [data-testid="stExpander"] details[open] summary > :first-child::after {
    transform: rotate(90deg);
}

/* Keep the actual expander title in its own column and allow wrapping. */
[data-testid="stSidebar"] [data-testid="stExpander"] summary p,
[data-testid="stSidebar"] [data-testid="stExpander"] summary div[data-testid="stMarkdownContainer"] {
    min-width: 0 !important;
    margin: 0 !important;
    white-space: normal !important;
    overflow-wrap: anywhere !important;
    word-break: normal !important;
    line-height: 1.3 !important;
}

[data-testid="stSidebar"] [data-testid="stWidgetLabel"],
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"],
[data-testid="stSidebar"] [data-testid="stCaptionContainer"],
[data-testid="stMetricLabel"],
[data-testid="stMetricValue"] {
    min-width: 0 !important;
}

[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p,
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
[data-testid="stSidebar"] [data-testid="stCaptionContainer"] p,
[data-testid="stMetricLabel"] p,
[data-testid="stMetricValue"] > div {
    white-space: normal !important;
    overflow-wrap: anywhere !important;
    word-break: normal !important;
    line-height: 1.3 !important;
}

.stTabs [data-baseweb="tab-list"] {
    flex-wrap: wrap !important;
}

/* ---------- kpi card internals ---------- */
.kpi-head {
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 10px;
}
.kpi-tag {
    font-size: 0.6rem;
    letter-spacing: 0.14em;
    text-transform: uppercase;
    color: var(--tertiary);
    border: 1px solid var(--border);
    border-radius: 999px;
    padding: 3px 9px;
}
.kpi-rows { margin-top: 16px; border-top: 1px solid rgba(255,255,255,0.07); }
.kpi-row {
    display: flex;
    justify-content: space-between;
    align-items: baseline;
    gap: 12px;
    padding: 8px 0;
    border-bottom: 1px solid rgba(255,255,255,0.045);
    font-size: 0.83rem;
}
.kpi-row:last-child { border-bottom: none; }
.kpi-row .k { color: var(--secondary); }
.kpi-row .v { font-weight: 600; font-variant-numeric: tabular-nums; }
.kpi-row .v.gold { color: var(--gold); }
.kpi-row .v.pos { color: var(--green); }
.kpi-row .v.neg { color: var(--red); }
 
/* ---------- progress ---------- */
.progress {
    position: relative;
    height: 6px;
    border-radius: 999px;
    background: rgba(255,255,255,0.08);
    margin: 16px 0 8px 0;
    overflow: visible;
}
.progress-fill {
    position: absolute;
    left: 0; top: 0; bottom: 0;
    border-radius: 999px;
    background: linear-gradient(90deg, rgba(216,183,106,0.55), var(--gold));
    box-shadow: 0 0 18px rgba(216,183,106,0.28);
}
.progress-fill.neutral {
    background: linear-gradient(90deg, rgba(255,255,255,0.22), rgba(255,255,255,0.42));
    box-shadow: none;
}
.progress-marker {
    position: absolute;
    top: -4px;
    width: 2px;
    height: 14px;
    background: var(--text);
    opacity: 0.85;
    border-radius: 2px;
}
.progress-legend {
    display: flex;
    justify-content: space-between;
    font-size: 0.72rem;
    color: var(--tertiary);
    font-variant-numeric: tabular-nums;
}
 
/* ---------- notes, callouts, pills ---------- */
.glass-note {
    color: var(--tertiary);
    font-size: 0.78rem;
    line-height: 1.6;
    margin: 8px 0 14px 0;
}
.glass-callout {
    background: var(--glass-3);
    border: 1px solid var(--border);
    border-left: 2px solid var(--gold);
    border-radius: var(--radius-sm);
    padding: 13px 16px;
    font-size: 0.85rem;
    line-height: 1.62;
    color: var(--secondary);
    margin: 10px 0 14px 0;
}
.glass-callout b, .glass-callout strong { color: var(--text); font-weight: 600; }
.glass-callout.ok { border-left-color: var(--green); }
.glass-callout.warn { border-left-color: var(--red); }
.tag-ok { color: var(--green); font-weight: 600; letter-spacing: 0.04em; }
.tag-warn { color: var(--red); font-weight: 600; letter-spacing: 0.04em; }
.inline-pill {
    display: inline-block;
    border: 1px solid var(--border);
    background: var(--glass-3);
    color: var(--secondary);
    border-radius: 999px;
    padding: 4px 11px;
    font-size: 0.7rem;
    font-weight: 500;
    margin: 3px 5px 3px 0;
}
.inline-pill.gold { color: var(--gold); border-color: rgba(216,183,106,0.35); }
 
/* ---------- scenario hero ---------- */
.scenario-hero {
    background:
        radial-gradient(circle at 88% -30%, rgba(216,183,106,0.16), transparent 55%),
        linear-gradient(135deg, rgba(255,255,255,0.085), rgba(255,255,255,0.03));
    border: 1px solid rgba(216,183,106,0.24);
    border-radius: var(--radius-lg);
    padding: 26px 28px;
    backdrop-filter: blur(26px) saturate(140%);
    -webkit-backdrop-filter: blur(26px) saturate(140%);
    box-shadow: 0 24px 70px rgba(0,0,0,0.34), inset 0 1px 0 rgba(255,255,255,0.09);
    margin-bottom: 16px;
}
.scenario-hero .eyebrow {
    font-size: 0.64rem;
    letter-spacing: 0.24em;
    text-transform: uppercase;
    color: var(--gold);
    font-weight: 600;
}
.scenario-hero .title {
    font-size: 1.72rem;
    font-weight: 600;
    letter-spacing: -0.028em;
    margin: 10px 0 6px 0;
}
.scenario-hero .thesis {
    font-size: 1.0rem;
    color: var(--text);
    opacity: 0.9;
    font-weight: 400;
    margin-bottom: 8px;
}
.scenario-hero .detail {
    font-size: 0.84rem;
    color: var(--secondary);
    line-height: 1.62;
    max-width: 940px;
}
.scenario-hero .milestone {
    margin-top: 14px;
    font-size: 0.76rem;
    color: var(--gold-soft);
    letter-spacing: 0.02em;
}
.stage-card { padding: 18px 20px; }
.stage-card .stage {
    font-size: 0.62rem;
    letter-spacing: 0.2em;
    text-transform: uppercase;
    color: var(--tertiary);
    font-weight: 600;
}
.stage-card.now { border-top: 2px solid rgba(255,255,255,0.22); }
.stage-card.jan { border-top: 2px solid rgba(216,183,106,0.55); }
.stage-card.mar { border-top: 2px solid rgba(99,217,154,0.45); }
 
/* ---------- glass tables ---------- */
.glass-table-wrap {
    border: 1px solid var(--border);
    border-radius: var(--radius-md);
    background: linear-gradient(135deg, rgba(255,255,255,0.045), rgba(255,255,255,0.018));
    backdrop-filter: blur(18px);
    -webkit-backdrop-filter: blur(18px);
    overflow: auto;
    margin: 6px 0 14px 0;
    max-height: 560px;
}
table.glass-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 0.8rem;
    font-variant-numeric: tabular-nums;
}
table.glass-table thead th {
    position: sticky;
    top: 0;
    z-index: 2;
    background: rgba(16,17,22,0.94);
    backdrop-filter: blur(14px);
    color: var(--tertiary);
    font-size: 0.62rem;
    letter-spacing: 0.13em;
    text-transform: uppercase;
    font-weight: 600;
    text-align: left;
    padding: 11px 14px;
    border-bottom: 1px solid var(--border);
    white-space: nowrap;
}
table.glass-table thead th.num { text-align: right; }
table.glass-table tbody td {
    padding: 10px 14px;
    border-bottom: 1px solid rgba(255,255,255,0.045);
    color: var(--text);
    white-space: nowrap;
}
table.glass-table tbody td.num { text-align: right; }
table.glass-table tbody td.pos { color: var(--green); }
table.glass-table tbody td.neg { color: var(--red); }
table.glass-table tbody tr:hover td { background: rgba(255,255,255,0.035); }
table.glass-table tbody tr.total td { font-weight: 600; color: var(--gold); }
table.glass-table tbody tr:last-child td { border-bottom: none; }
 
/* ---------- FINAL workbook reference ---------- */
.final-sheet-scroll {
    overflow: auto;
    max-height: 70vh;
    border: 1px solid var(--border);
    border-radius: var(--radius-md);
    background: rgba(255,255,255,0.02);
    padding: 4px;
}
table.final-sheet-table {
    border-collapse: collapse;
    width: max-content;
    min-width: 100%;
    font-size: 11.5px;
    color: var(--text);
}
table.final-sheet-table td {
    border: 1px solid rgba(255,255,255,0.07);
    padding: 5px 8px;
    min-width: 74px;
    white-space: nowrap;
    vertical-align: middle;
}
td.final-spacer { height: 7px; border: none !important; }
 
/* ---------- streamlit widgets ---------- */
div[role="radiogroup"] { gap: 8px !important; flex-wrap: wrap; }
div[role="radiogroup"] > label {
    background: var(--glass-2);
    border: 1px solid var(--border);
    border-radius: 999px;
    padding: 8px 15px;
    margin: 0 !important;
    transition: all 160ms ease;
    backdrop-filter: blur(14px);
    -webkit-backdrop-filter: blur(14px);
}
div[role="radiogroup"] > label > div:first-child { display: none !important; }
div[role="radiogroup"] > label div[data-testid="stMarkdownContainer"] p {
    font-size: 0.78rem !important;
    color: var(--secondary) !important;
    font-weight: 500;
    margin: 0 !important;
}
div[role="radiogroup"] > label:hover { border-color: var(--border-strong); }
div[role="radiogroup"] > label:has(input:checked) {
    border-color: rgba(216,183,106,0.55);
    background: linear-gradient(135deg, rgba(216,183,106,0.16), rgba(255,255,255,0.05));
    box-shadow: 0 8px 26px rgba(216,183,106,0.14), inset 0 1px 0 rgba(255,255,255,0.10);
}
div[role="radiogroup"] > label:has(input:checked) div[data-testid="stMarkdownContainer"] p {
    color: var(--text) !important;
    font-weight: 600;
}
 
[data-baseweb="select"] > div, [data-baseweb="input"], .stTextInput input, .stNumberInput input {
    background: var(--glass-2) !important;
    border-color: var(--border) !important;
    color: var(--text) !important;
    border-radius: 12px !important;
}
[data-baseweb="select"] svg { fill: var(--secondary); }
[data-baseweb="popover"] div[role="listbox"], [data-baseweb="menu"] {
    background: rgba(16,17,22,0.97) !important;
    border: 1px solid var(--border) !important;
    border-radius: 14px !important;
}
[data-baseweb="menu"] li:hover { background: rgba(255,255,255,0.06) !important; }
[data-baseweb="tag"] {
    background: rgba(216,183,106,0.16) !important;
    border: 1px solid rgba(216,183,106,0.3) !important;
    color: var(--text) !important;
}
 
.stSlider [data-baseweb="slider"] div[role="slider"] {
    background: var(--gold) !important;
    box-shadow: 0 0 0 4px rgba(216,183,106,0.16) !important;
    transition: transform 140ms ease;
}
.stSlider [data-baseweb="slider"] div[role="slider"]:active { transform: scale(1.1); }
.stSlider [data-testid="stTickBar"], .stSlider [data-testid="stTickBarMin"],
.stSlider [data-testid="stTickBarMax"] { color: var(--tertiary) !important; }
 
.stButton > button, .stDownloadButton > button, .stFormSubmitButton > button {
    background: linear-gradient(135deg, rgba(216,183,106,0.9), rgba(201,170,101,0.78)) !important;
    color: #10120F !important;
    border: 1px solid rgba(216,183,106,0.5) !important;
    border-radius: 14px !important;
    font-weight: 600 !important;
    letter-spacing: 0.01em;
    padding: 8px 18px !important;
    transition: transform 160ms ease, box-shadow 160ms ease;
    box-shadow: 0 8px 24px rgba(216,183,106,0.16);
}
.stButton > button:hover, .stDownloadButton > button:hover {
    transform: translateY(-1px);
    box-shadow: 0 14px 34px rgba(216,183,106,0.24);
}
.stButton > button[kind="secondary"] {
    background: var(--glass-2) !important;
    color: var(--text) !important;
    border: 1px solid var(--border) !important;
    box-shadow: none;
}
 
.stTabs [data-baseweb="tab-list"] {
    gap: 6px;
    border-bottom: 1px solid var(--border);
    background: transparent;
}
.stTabs [data-baseweb="tab"] {
    background: var(--glass-3);
    border: 1px solid transparent;
    border-radius: 12px 12px 0 0;
    color: var(--tertiary);
    padding: 8px 16px;
    font-size: 0.8rem;
}
.stTabs [aria-selected="true"] {
    color: var(--text) !important;
    background: var(--glass-1);
    border-color: var(--border);
    border-bottom-color: transparent;
    font-weight: 600;
}
.stTabs [data-baseweb="tab-highlight"] { background: var(--gold) !important; }
 
[data-testid="stExpander"] {
    border: 1px solid var(--border) !important;
    border-radius: var(--radius-md) !important;
    background: var(--glass-3) !important;
    backdrop-filter: blur(16px);
    -webkit-backdrop-filter: blur(16px);
}
[data-testid="stExpander"] summary { color: var(--secondary) !important; font-size: 0.84rem; }
[data-testid="stExpander"] summary:hover { color: var(--text) !important; }
 
[data-testid="stSidebar"] {
    background: rgba(10,11,14,0.86) !important;
    border-right: 1px solid var(--border);
    backdrop-filter: blur(26px) saturate(140%);
    -webkit-backdrop-filter: blur(26px) saturate(140%);
}
[data-testid="stSidebar"] * { color: var(--text); }
[data-testid="stSidebar"] [data-testid="stCaptionContainer"] p { color: var(--tertiary) !important; }
.sidebar-mark {
    font-size: 0.6rem;
    letter-spacing: 0.26em;
    text-transform: uppercase;
    color: var(--gold);
    font-weight: 600;
    margin-bottom: 4px;
}
.sidebar-title {
    font-size: 0.62rem;
    letter-spacing: 0.16em;
    text-transform: uppercase;
    color: var(--tertiary);
    font-weight: 600;
    margin: 18px 0 6px 0;
}
 
[data-testid="stDataFrame"] {
    border: 1px solid var(--border);
    border-radius: var(--radius-md);
    overflow: hidden;
}
[data-testid="stFileUploaderDropzone"] {
    background: var(--glass-2) !important;
    border: 1px dashed rgba(255,255,255,0.18) !important;
    border-radius: var(--radius-md) !important;
    color: var(--secondary) !important;
}
[data-testid="stFileUploaderDropzone"]:hover { border-color: rgba(216,183,106,0.45) !important; }
 
[data-testid="stAlert"] {
    background: var(--glass-2) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--radius-sm) !important;
    color: var(--text) !important;
}
hr { border-color: var(--border) !important; }
[data-testid="stMetricValue"] { color: var(--text) !important; }
 
/* ---------- responsive ---------- */
@media (max-width: 1180px) {
    .hero-grid { grid-template-columns: 1fr 1fr; }
    .exec-title { font-size: 1.85rem; }
}
@media (max-width: 760px) {
    .hero-grid, .hero-grid.no-aum { grid-template-columns: 1fr; }
    .kpi-grid { grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); }
    .metric-hero { font-size: 2.05rem; }
    .exec-header { flex-direction: column; align-items: flex-start; }
    div[role="radiogroup"] { flex-wrap: nowrap; overflow-x: auto; padding-bottom: 6px; }
}
@media (prefers-reduced-motion: reduce) {
    * { animation-duration: 0.01ms !important; transition-duration: 0.01ms !important; }
    .glass-card:hover, .glass-kpi:hover, .stButton > button:hover { transform: none; }
}
</style>
"""
 
 
def inject_theme() -> None:
    """Inject the liquid-glass design system exactly once per run."""
    st.markdown(GLASS_CSS, unsafe_allow_html=True)
 
 
# --- Small HTML helpers -------------------------------------------------------
 
def _pct_width(value: Any, cap: float = 100.0) -> float:
    v = _num(value)
    if v is None:
        return 0.0
    return float(min(max(v * 100.0, 0.0), cap))
 
 
def _tone_for(value: Any, invert: bool = False) -> str:
    v = _num(value)
    if v is None:
        return "flat"
    if abs(v) < 1e-12:
        return "flat"
    positive = v > 0
    if invert:
        positive = not positive
    return "pos" if positive else "neg"
 
 
def section_header(index: str, title: str, subtitle: str = "") -> None:
    sub = f"<span class='sub'>{escape(subtitle)}</span>" if subtitle else ""
    st.markdown(
        f"<div class='glass-section'><span class='idx'>{escape(index)}</span>"
        f"<span class='ttl'>{escape(title)}</span>{sub}</div>"
        "<div class='section-rule'></div>",
        unsafe_allow_html=True,
    )
 
 
def glass_note(text: str) -> None:
    st.markdown(f"<div class='glass-note'>{text}</div>", unsafe_allow_html=True)
 
 
def glass_callout(text: str, tone: str = "") -> None:
    css = "glass-callout" + (f" {tone}" if tone else "")
    st.markdown(f"<div class='{css}'>{text}</div>", unsafe_allow_html=True)
 
 
def progress_html(
    achieved_pct: Any,
    marker_pct: Any = None,
    left_label: str = "",
    right_label: str = "",
    neutral: bool = False,
) -> str:
    fill = _pct_width(achieved_pct)
    marker = _pct_width(marker_pct) if marker_pct is not None else None
    marker_html = (
        f"<div class='progress-marker' style='left:{marker:.1f}%'></div>"
        if marker is not None else ""
    )
    fill_class = "progress-fill neutral" if neutral else "progress-fill"
    legend = ""
    if left_label or right_label:
        legend = (
            "<div class='progress-legend'>"
            f"<span>{escape(left_label)}</span><span>{escape(right_label)}</span></div>"
        )
    return (
        f"<div class='progress'><div class='{fill_class}' style='width:{fill:.1f}%'></div>"
        f"{marker_html}</div>{legend}"
    )
 
 
def kpi_tile_html(
    label: str,
    value: str,
    delta: str = "",
    delta_tone: str = "flat",
    secondary: str = "",
) -> str:
    delta_html = (
        f"<div class='metric-delta {delta_tone}'>{escape(delta)}</div>" if delta else ""
    )
    secondary_html = (
        f"<div class='metric-secondary'>{escape(secondary)}</div>" if secondary else ""
    )
    return (
        "<div class='glass-kpi'>"
        f"<div class='metric-label'>{escape(label)}</div>"
        f"<div class='metric-value'>{escape(value)}</div>"
        f"{delta_html}{secondary_html}</div>"
    )
 
 
def kpi_strip(tiles: Sequence[Dict[str, str]]) -> None:
    """Render a responsive row of glass KPI tiles from one markdown call."""
    if not tiles:
        return
    cards = "".join(
        kpi_tile_html(
            tile.get("label", ""),
            tile.get("value", NA_TEXT),
            tile.get("delta", ""),
            tile.get("tone", "flat"),
            tile.get("secondary", ""),
        )
        for tile in tiles
    )
    st.markdown(f"<div class='kpi-grid'>{cards}</div>", unsafe_allow_html=True)
 
 
def _dataframe_kwargs() -> Dict[str, Any]:
    try:
        parameters = inspect.signature(st.dataframe).parameters
    except (TypeError, ValueError):
        return {}
    if "width" in parameters and parameters["width"].default == "stretch":
        return {}
    if "use_container_width" in parameters:
        return {"use_container_width": True}
    return {}
 
 
def render_glass_table(
    frame: pd.DataFrame,
    formats: Optional[Dict[str, str]] = None,
    total_rows: Sequence[str] = (),
    max_html_rows: int = 240,
    empty_message: str = "No rows for this selection.",
) -> None:
    """Compact glass table: sticky header, right-aligned numbers, signed colour."""
    if frame is None or frame.empty:
        glass_note(escape(empty_message))
        return
 
    if len(frame) > max_html_rows:  # very large frames stay virtualised
        display = format_table(frame, formats) if formats else frame
        try:
            st.dataframe(display, hide_index=True, **_dataframe_kwargs())
        except TypeError:  # pragma: no cover - very old Streamlit
            st.dataframe(display)
        return
 
    formats = formats or {}
    display = format_table(frame, formats) if formats else frame.astype(str)
 
    header_cells = []
    for column in display.columns:
        kind = formats.get(column, "txt")
        css = " class='num'" if kind in NUMERIC_FORMATS else ""
        header_cells.append(f"<th{css}>{escape(str(column))}</th>")
 
    first_column = display.columns[0] if len(display.columns) else None
    body_rows = []
    for position in range(len(display)):
        raw_first = str(frame.iloc[position][first_column]) if first_column is not None else ""
        row_class = " class='total'" if raw_first in total_rows else ""
        cells = []
        for column in display.columns:
            kind = formats.get(column, "txt")
            classes = []
            if kind in NUMERIC_FORMATS:
                classes.append("num")
            if kind in SIGNED_FORMATS:
                classes.append(_tone_for(frame.iloc[position][column]))
            css = f" class='{' '.join(classes)}'" if classes else ""
            cells.append(f"<td{css}>{escape(str(display.iloc[position][column]))}</td>")
        body_rows.append(f"<tr{row_class}>{''.join(cells)}</tr>")
 
    st.markdown(
        "<div class='glass-table-wrap'><table class='glass-table'>"
        f"<thead><tr>{''.join(header_cells)}</tr></thead>"
        f"<tbody>{''.join(body_rows)}</tbody></table></div>",
        unsafe_allow_html=True,
    )
 
 
def reset_stale_selection(key: str, options: Sequence[Any]) -> None:
    """Drop a stored widget value that is no longer offered, so nothing errors."""
    if key in st.session_state and st.session_state[key] not in options:
        st.session_state.pop(key, None)
 
 
def rerun() -> None:
    handler = getattr(st, "rerun", None) or getattr(st, "experimental_rerun", None)
    if handler is not None:
        handler()
 
 
# =============================================================================
# 15. EXECUTIVE COMPONENTS - CURRENT REALITY
# =============================================================================
 
def render_apple_header(
    final_metrics: Dict[str, Any],
    records: pd.DataFrame,
    page_label: str,
) -> None:
    """01 · Minimal executive header with workbook status on the right."""
    months_done = int(final_metrics.get("months_done", MONTHS_COMPLETED))
    sheet = str(final_metrics.get("sheet_name") or "FINAL")
    rm_count = len(records)
 
    st.markdown(
        "<div class='exec-header'>"
        "<div>"
        f"<div class='exec-mark'>{escape(page_label)}</div>"
        f"<div class='exec-title'>{escape(APP_TITLE)}</div>"
        f"<div class='exec-sub'>{escape(APP_SUBTITLE)}</div>"
        "</div>"
        "<div class='exec-status'>"
        "<div class='status-chip status-live'><div class='k'>Workbook</div>"
        f"<div class='v'><span class='status-dot'></span>{escape(sheet)}</div></div>"
        "<div class='status-chip'><div class='k'>Data period</div>"
        "<div class='v'>Apr–Jun FY27</div></div>"
        "<div class='status-chip'><div class='k'>Months done</div>"
        f"<div class='v'>{months_done} of 12</div></div>"
        "<div class='status-chip'><div class='k'>RMs in scope</div>"
        f"<div class='v'>{rm_count:,}</div></div>"
        "</div></div>",
        unsafe_allow_html=True,
    )
 
 
def _kpi_row_html(label: str, value: str, css: str = "") -> str:
    value_class = f"v {css}".strip()
    return (
        f"<div class='kpi-row'><span class='k'>{escape(label)}</span>"
        f"<span class='{value_class}'>{escape(value)}</span></div>"
    )
 
 
def render_aum_hero(aum: pd.DataFrame) -> str:
    """AUM card shown only in Current Performance Metrics."""
    if not isinstance(aum, pd.DataFrame) or aum.empty or "Overall" not in aum.index:
        return (
            "<div class='glass-card'><div class='metric-label'>Assets under management</div>"
            "<div class='metric-hero'>—</div>"
            "<div class='metric-secondary'>AUM could not be located on the FINAL sheet.</div></div>"
        )

    row = aum.loc["Overall"]
    achievement = _num(row.get("Achievement %"))
    gap = _num(row.get("Gap to Target"))

    rows = "".join([
        _kpi_row_html("Target", fmt_cr(row.get("Target"))),
        _kpi_row_html(
            "Gap to target",
            fmt_cr_signed(None if gap is None else -gap),
            "neg" if (gap or 0) > 0 else "pos",
        ),
        _kpi_row_html("Achieved", fmt_pct(achievement), "gold"),
    ])

    return (
        "<div class='glass-card'>"
        "<div class='kpi-head'><span class='metric-label'>Assets under management</span>"
        "<span class='kpi-tag'>FINAL</span></div>"
        f"<div class='metric-hero gold'>{escape(fmt_cr(row.get('Current')))}</div>"
        "<div class='metric-label'>Current AUM</div>"
        + progress_html(
            achievement,
            marker_pct=1.0,
            left_label=f"{fmt_pct(achievement)} of target",
            right_label=f"target {fmt_cr(row.get('Target'))}",
        )
        + f"<div class='kpi-rows'>{rows}</div></div>"
    )


def render_sales_kpi_card(title: str, row: pd.Series) -> str:
    """Gross / Net sales card: YTD is the hero, pace and projection support it."""
    if row is None or len(row) == 0:
        return (
            f"<div class='glass-card'><div class='metric-label'>{escape(title)}</div>"
            "<div class='metric-hero'>—</div>"
            "<div class='metric-secondary'>Metrics could not be located on the FINAL sheet.</div></div>"
        )
 
    achievement = _num(row.get("Achievement %"))
    projected = _num(row.get("Projected FY %"))
    current_rr = _num(row.get("Current RR"))
    required_rr = _num(row.get("Required RR to Target"))
    pace_gap = None
    if current_rr is not None and required_rr is not None and required_rr != 0:
        pace_gap = current_rr / required_rr - 1.0
 
    rows = "".join([
        _kpi_row_html("FY27 target", fmt_cr(row.get("FY27 Target"))),
        _kpi_row_html("Current run rate", fmt_cr(current_rr)),
        _kpi_row_html("Required run rate", fmt_cr(required_rr), "gold"),
        _kpi_row_html(
            "Pace vs required",
            fmt_pct_signed(pace_gap),
            _tone_for(pace_gap),
        ),
        _kpi_row_html(
            "Projected FY",
            fmt_pct(projected),
            "pos" if (projected or 0) >= 1 else "neg",
        ),
    ])
 
    return (
        "<div class='glass-card'>"
        f"<div class='kpi-head'><span class='metric-label'>{escape(title)}</span>"
        "<span class='kpi-tag'>FY27</span></div>"
        f"<div class='metric-hero'>{escape(fmt_cr(row.get('YTD')))}</div>"
        "<div class='metric-label'>Year to date</div>"
        + progress_html(
            achievement,
            marker_pct=projected,
            left_label=f"{fmt_pct(achievement)} achieved",
            right_label=f"projected {fmt_pct(projected)}",
        )
        + f"<div class='kpi-rows'>{rows}</div></div>"
    )
 
 
def _overall_row(frame: Any) -> pd.Series:
    if isinstance(frame, pd.DataFrame) and not frame.empty and "Overall" in frame.index:
        return frame.loc["Overall"]
    return pd.Series(dtype=float)
 
 
def render_current_performance(final_metrics: Dict[str, Any], model: ScenarioModel) -> None:
    """02 · Current Performance Metrics · FINAL - AUM, Gross and Net at one glance."""
    section_header(
        "02",
        "Current Performance Metrics · FINAL",
        "Where the business stands before any scenario is applied",
    )

    gs = final_sales_metrics(final_metrics, model, "GS")
    ns = final_sales_metrics(final_metrics, model, "NS")
    aum = final_metrics.get("AUM")
    aum_frame = aum if isinstance(aum, pd.DataFrame) else pd.DataFrame()

    # Current Performance Metrics must always contain these three cards in this
    # order: AUM, Gross Sales, Net Sales. AUM remains excluded from the separate
    # Business Drivers section, as requested.
    aum_card = render_aum_hero(aum_frame)
    gross_card = render_sales_kpi_card("Gross Sales", _overall_row(gs))
    net_card = render_sales_kpi_card("Net Sales", _overall_row(ns))
    cards = aum_card + gross_card + net_card
    st.markdown(f"<div class='hero-grid'>{cards}</div>", unsafe_allow_html=True)

    gs_row, ns_row = _overall_row(gs), _overall_row(ns)
    aum_row = _overall_row(aum_frame)
    tiles = [
        {
            "label": "AUM gap to target",
            "value": fmt_cr(aum_row.get("Gap to Target")),
            "secondary": "Target less current AUM",
        },
        {
            "label": "Gross sales shortfall",
            "value": fmt_cr(
                None if _num(gs_row.get("FY27 Target")) is None
                else _z(gs_row.get("FY27 Target")) - _z(gs_row.get("YTD"))
            ),
            "secondary": "Still to book by March 2027",
        },
        {
            "label": "Net sales shortfall",
            "value": fmt_cr(
                None if _num(ns_row.get("FY27 Target")) is None
                else _z(ns_row.get("FY27 Target")) - _z(ns_row.get("YTD"))
            ),
            "secondary": "Still to book by March 2027",
        },
        {
            "label": "Net projected FY",
            "value": fmt_pct(ns_row.get("Projected FY %")),
            "delta": fmt_pts(
                None if _num(ns_row.get("Projected FY %")) is None
                else _num(ns_row.get("Projected FY %")) - 1.0
            ),
            "tone": _tone_for(
                None if _num(ns_row.get("Projected FY %")) is None
                else _num(ns_row.get("Projected FY %")) - 1.0
            ),
            "secondary": "At the current run rate",
        },
    ]
    kpi_strip(tiles)
    glass_note(
        "AUM, Gross Sales and Net Sales are read directly from the workbook's "
        "<b>FINAL</b> sheet. Achievement, run rate and projected FY are derived from those "
        "same Target and YTD values."
    )

def render_business_driver_selector(
    final_metrics: Dict[str, Any],
    model: ScenarioModel,
) -> None:
    """03 · Business drivers - shown only by Asset Class and Channel."""
    section_header(
        "03",
        "Business drivers",
        "Performance is segregated only by asset class and channel",
    )

    st.markdown("<div class='metric-label'>Sales basis</div>", unsafe_allow_html=True)
    basis_label = st.radio(
        "Sales basis",
        [SALES_LABEL["GS"], SALES_LABEL["NS"]],
        index=1,
        horizontal=True,
        key="driver_basis",
        label_visibility="collapsed",
    )
    basis = "GS" if basis_label == SALES_LABEL["GS"] else "NS"
    st.session_state["display_basis"] = basis

    frame = final_sales_metrics(final_metrics, model, basis)
    if frame.empty:
        glass_note("Driver metrics could not be located on the FINAL sheet.")
        return

    overall = _overall_row(frame)
    projected = _num(overall.get("Projected FY %"))
    kpi_strip([
        {"label": "FY27 target", "value": fmt_cr(overall.get("FY27 Target"))},
        {
            "label": "YTD",
            "value": fmt_cr(overall.get("YTD")),
            "delta": fmt_pct(overall.get("Achievement %")),
            "tone": "gold",
            "secondary": "of FY27 target booked",
        },
        {"label": "Current run rate", "value": fmt_cr(overall.get("Current RR"))},
        {
            "label": "Required run rate",
            "value": fmt_cr(overall.get("Required RR to Target")),
            "secondary": "Target ÷ 12 months",
        },
        {
            "label": "Projected FY",
            "value": fmt_pct(projected),
            "delta": fmt_pts(None if projected is None else projected - 1.0),
            "tone": _tone_for(None if projected is None else projected - 1.0),
        },
    ])

    display_columns = [
        "FY27 Target", "YTD", "Achievement %", "Current RR",
        "Required RR to Target", "Estimated FY @ Current RR", "Projected FY %",
    ]
    formats = {
        "Scope": "txt", "FY27 Target": "cr", "YTD": "cr", "Achievement %": "pct",
        "Current RR": "cr", "Required RR to Target": "cr",
        "Estimated FY @ Current RR": "cr", "Projected FY %": "pct",
    }

    asset_frame = frame.loc[[i for i in FINAL_ASSET_ROWS if i in frame.index]]
    channel_frame = frame.loc[[i for i in FINAL_CHANNEL_ROWS if i in frame.index]]

    tabs = st.tabs(["Asset class", "Channel"])
    with tabs[0]:
        if asset_frame.empty:
            glass_note("Asset-class metrics could not be located on the FINAL sheet.")
        else:
            render_glass_table(
                asset_frame.reset_index().rename(columns={"Metric": "Scope"})[["Scope", *display_columns]],
                formats,
            )
    with tabs[1]:
        if channel_frame.empty:
            glass_note("Channel metrics could not be located on the FINAL sheet.")
        else:
            render_glass_table(
                channel_frame.reset_index().rename(columns={"Metric": "Scope"})[["Scope", *display_columns]],
                formats,
            )

# =============================================================================
# 16. ANALYSIS SCOPE & CURRENT RUN RATE
# =============================================================================
 
LOCATION_FIELD_ALIASES = ["MKT TYPE", "Market Type", "Mkt Type"]
 
 
def _location_column(records: pd.DataFrame) -> Optional[str]:
    for c in LOCATION_FIELD_ALIASES:
        if c in records.columns:
            return c
    return None
 
 
def _location_options(records: pd.DataFrame, channel: str) -> List[str]:
    col = _location_column(records)
    if col is None:
        return ["All"]
    work = records
    if channel != "All" and "Vertical" in work.columns:
        work = work.loc[work["Vertical"] == channel]
    values = sorted({str(v).strip() for v in work[col].dropna().tolist() if str(v).strip()})
    # Preserve the workbook's own location cuts, including T2/T6/T30/B30/EM.
    return ["All"] + values
 
 
def apply_management_cuts(records: pd.DataFrame, channel: str, location: str) -> pd.DataFrame:
    out = records.copy()
    if channel != "All" and "Vertical" in out.columns:
        out = out.loc[out["Vertical"] == channel].copy()
    col = _location_column(out)
    if location != "All" and col is not None:
        out = out.loc[out[col].astype(str).str.strip() == location].copy()
    return out
 
 
def render_analysis_scope(records: pd.DataFrame) -> Tuple[str, str, str, pd.DataFrame]:
    """
    Scope controls that recalculate the analytical engine.

    Channel and location remain as existing filters. Asset-class results are
    shown together downstream instead of forcing an Equity / Debt / Liquid drill-down.
    """
    st.markdown("<div class='metric-label'>Analysis scope</div>", unsafe_allow_html=True)
    left, middle = st.columns([1, 1])

    with left:
        present = set(records.get("Vertical", pd.Series(dtype=str)).astype(str))
        channel_options = ["All"] + [v for v in VERTICALS if v in present]
        reset_stale_selection("scope_channel", channel_options)
        channel = st.selectbox("Channel", channel_options, index=0, key="scope_channel")
    with middle:
        location_options = _location_options(records, channel)
        reset_stale_selection("scope_location", location_options)
        location = st.selectbox(
            "Location / market type", location_options, index=0, key="scope_location"
        )

    # Asset classes stay together; no Equity / Debt / Liquid selector.
    asset = "All"
    filtered = apply_management_cuts(records, channel, location)

    active = [x for x in (
        channel if channel != "All" else None,
        location if location != "All" else None,
    ) if x]
    scope_text = " · ".join(active) if active else "All business"
    st.markdown(
        f"<div class='glass-note'>Scope <span class='inline-pill gold'>{escape(scope_text)}</span>"
        f"<span class='inline-pill'>{len(filtered):,} RMs</span>"
        "Every scenario below is recalculated on this population.</div>",
        unsafe_allow_html=True,
    )

    if filtered.empty:
        glass_callout(
            "No RM records match this scope. Widen the channel or location selection "
            "to bring the calculation engine back online.",
            tone="warn",
        )
    return channel, location, asset, filtered

def render_current_runrate(grid: pd.DataFrame, basis: str, asset: str) -> None:
    """04 · Current pace, required pace and projected finish. Graph removed."""
    section_header(
        "04",
        "Current run rate & target gap",
        "The pace the business is running at, against the pace the target needs",
    )

    if grid.empty:
        glass_note("No records in scope, so the current run rate cannot be calculated.")
        return

    cell = summarize_current(grid, sales=basis, asset=None if asset == "All" else asset)
    required_rr = _z(cell.get("fy_target")) / 12.0
    current_rr = _num(cell.get("current_rr"))
    pace_gap = None
    if current_rr is not None and required_rr:
        pace_gap = current_rr / required_rr - 1.0
    projected = _num(cell.get("current_march_pct"))
    shortfall = _z(cell.get("fy_target")) - _z(cell.get("current_march"))

    kpi_strip([
        {"label": "FY target", "value": fmt_cr(cell.get("fy_target")),
         "secondary": f"{SALES_LABEL[basis]} · RM calculation sheets"},
        {"label": "YTD achievement", "value": fmt_cr(cell.get("ytd_ach")),
         "delta": fmt_pct(cell.get("ytd_ach_pct")), "tone": "gold",
         "secondary": "against the YTD June target"},
        {"label": "Current run rate", "value": fmt_cr(current_rr),
         "secondary": "YTD ÷ 3 completed months"},
        {"label": "Required run rate", "value": fmt_cr(required_rr),
         "delta": fmt_pct_signed(pace_gap), "tone": _tone_for(pace_gap),
         "secondary": "pace gap on today's run rate"},
        {"label": "Projected March", "value": fmt_pct(projected),
         "delta": fmt_pts(None if projected is None else projected - 1.0),
         "tone": _tone_for(None if projected is None else projected - 1.0),
         "secondary": fmt_cr(cell.get("current_march"))},
        {"label": "Gap at current pace", "value": fmt_cr(shortfall),
         "tone": _tone_for(-shortfall),
         "secondary": "FY target less projected March"},
    ])

    glass_note(
        "The current run rate is the completed Apr–Jun achievement divided by three. "
        "The required run rate is the FY target spread evenly across twelve months."
    )

# =============================================================================
# 17. SCENARIO PLANNING COMPONENTS
# =============================================================================
 
def render_scenario_navigator() -> int:
    """05 · Horizontal glass scenario navigator."""
    section_header(
        "05",
        "Scenario planning",
        "What changes if the organisation changes the trajectory",
    )
    options = [f"{sid:02d}  {SCENARIOS[sid]['short']}" for sid in SCENARIO_ORDER]
    choice = st.radio(
        "Scenario",
        options,
        index=SCENARIO_ORDER.index(st.session_state.get("scenario_id", 1)),
        horizontal=True,
        key="scenario_navigator",
        label_visibility="collapsed",
    )
    scenario_id = SCENARIO_ORDER[options.index(choice)]
    st.session_state["scenario_id"] = scenario_id
    return scenario_id
 
 
def render_scenario_controls(scenario_id: int, base_params: Dict[str, Any]) -> Dict[str, Any]:
    """Only the assumptions the selected scenario actually uses."""
    params = dict(base_params)
 
    if scenario_id == 3:
        with st.expander("Scenario assumptions", expanded=True):
            params["dip"] = st.slider(
                "February–March run-rate dip", 0, 60, int(S3_DEFAULT_DIP * 100), 5,
                format="%d%%", key="s3_dip",
            ) / 100.0
 
    elif scenario_id == 7:
        with st.expander("Scenario assumptions", expanded=True):
            columns = st.columns(3)
            params["jan_target"] = columns[0].slider(
                "January achievement target", 90, 120, int(S7_DEFAULT_JAN_TARGET * 100), 1,
                format="%d%%", key="s7_jan",
            ) / 100.0
            params["mar_target"] = columns[1].slider(
                "March achievement target", 90, 120, int(S7_DEFAULT_MAR_TARGET * 100), 1,
                format="%d%%", key="s7_mar",
            ) / 100.0
            params["leakage"] = columns[2].slider(
                "February–March leakage", 0, 30, int(S7_DEFAULT_LEAKAGE * 100), 1,
                format="%d%%", key="s7_leak",
            ) / 100.0
 
    elif scenario_id == 8:
        with st.expander("Channel assumptions", expanded=False):
            params["leakage"] = st.slider(
                "February–March leakage", 0, 30, int(S8_DEFAULT_LEAKAGE * 100), 1,
                format="%d%%", key="s8_leakage",
            ) / 100.0
            growth = dict(params.get("channel_growth", S8_DEFAULT_GROWTH))
            jan_targets = dict(params.get("channel_jan_target", S8_DEFAULT_JAN_TARGET))
            mar_targets = dict(params.get("channel_mar_target", S8_DEFAULT_MAR_TARGET))
            for channel in CHANNELS:
                st.markdown(
                    f"<div class='metric-label' style='margin-top:10px'>{escape(channel)}</div>",
                    unsafe_allow_html=True,
                )
                columns = st.columns(3)
                growth[channel] = columns[0].slider(
                    "Monthly growth", -20, 30, int(S8_DEFAULT_GROWTH[channel] * 100), 1,
                    format="%d%%", key=f"s8_g_{channel}", label_visibility="collapsed",
                ) / 100.0
                jan_targets[channel] = columns[1].slider(
                    "January 2027", 80, 180, int(S8_DEFAULT_JAN_TARGET[channel] * 100), 1,
                    format="%d%%", key=f"s8_j_{channel}", label_visibility="collapsed",
                ) / 100.0
                mar_targets[channel] = columns[2].slider(
                    "March 2027", 80, 200, int(S8_DEFAULT_MAR_TARGET[channel] * 100), 1,
                    format="%d%%", key=f"s8_m_{channel}", label_visibility="collapsed",
                ) / 100.0
            params["channel_growth"] = growth
            params["channel_jan_target"] = jan_targets
            params["channel_mar_target"] = mar_targets
            glass_note(
                "Each row sets monthly growth, the January 2027 target and the "
                "March 2027 target for one channel."
            )
 
    elif scenario_id == 9:
        with st.expander("Optimiser assumptions", expanded=True):
            columns = st.columns(2)
            ambition = columns[0].slider(
                "Portfolio March ambition", 100, 180, 120, 1, format="%d%%", key="s9_target",
            ) / 100.0
            params["optimizer_target"] = ambition
            params["leakage"] = columns[1].slider(
                "February–March leakage", 0, 30, int(S8_DEFAULT_LEAKAGE * 100), 1,
                format="%d%%", key="s9_leakage",
            ) / 100.0
            # The optimiser solves every channel against the selected ambition,
            # holding the January milestone at 100% of the FY target.
            params["channel_mar_target"] = {c: ambition for c in CHANNELS}
            params["channel_jan_target"] = dict(S8_DEFAULT_JAN_TARGET)
 
    return params
 
 
def render_scenario_hero(model: ScenarioModel, basis: str, asset: str) -> Dict[str, Any]:
    """06 · Scenario hero: the thesis, then Current → January → March."""
    meta = model.meta
    section_header(
        "06",
        f"Scenario {model.scenario_id} · {meta['name']}",
        "The selected strategy and what it demands",
    )
 
    st.markdown(
        "<div class='scenario-hero'>"
        f"<div class='eyebrow'>Scenario {model.scenario_id:02d} · {escape(meta['short'])}</div>"
        f"<div class='title'>{escape(meta['name'])}</div>"
        f"<div class='thesis'>{escape(meta['thesis'])}</div>"
        f"<div class='detail'>{escape(meta['explanation'])}</div>"
        f"<div class='milestone'>{escape(meta['milestone'])}</div>"
        "</div>",
        unsafe_allow_html=True,
    )
 
    cell = model.cell(basis, asset=None if asset == "All" else asset)
 
    now_rows = "".join([
        _kpi_row_html("YTD booked", fmt_cr(cell.get("ytd_ach"))),
        _kpi_row_html("Current run rate", fmt_cr(cell.get("current_rr"))),
        _kpi_row_html("Projected March", fmt_pct(cell.get("current_march_pct"))),
    ])
    jan_rows = "".join([
        _kpi_row_html("Milestone", fmt_pct(cell.get("milestone_pct")), "gold"),
        _kpi_row_html("Required run rate", fmt_cr(cell.get("scen_rr"))),
        _kpi_row_html(
            "Buffer vs milestone",
            fmt_cr_signed(cell.get("jan_buffer")),
            _tone_for(cell.get("jan_buffer")),
        ),
    ])
    march_rows = "".join([
        _kpi_row_html("Scenario achievement", fmt_cr(cell.get("march_amount"))),
        _kpi_row_html("Feb–Mar run rate", fmt_cr(cell.get("feb_mar_rr"))),
        _kpi_row_html(
            "Headroom vs ambition",
            fmt_cr_signed(cell.get("headroom_amt")),
            _tone_for(cell.get("headroom_amt")),
        ),
    ])
 
    st.markdown(
        "<div class='trio-grid'>"
        "<div class='glass-card stage-card now'><div class='stage'>Current trajectory</div>"
        f"<div class='metric-hero'>{escape(fmt_pct(cell.get('current_march_pct')))}</div>"
        "<div class='metric-label'>Projected March at today's pace</div>"
        f"<div class='kpi-rows'>{now_rows}</div></div>"
        "<div class='glass-card stage-card jan'><div class='stage'>January 2027 milestone</div>"
        f"<div class='metric-hero gold'>{escape(fmt_pct(cell.get('jan_pct')))}</div>"
        "<div class='metric-label'>Of FY target by January</div>"
        f"<div class='kpi-rows'>{jan_rows}</div></div>"
        "<div class='glass-card stage-card mar'><div class='stage'>March 2027 outcome</div>"
        f"<div class='metric-hero'>{escape(fmt_pct(cell.get('march_pct')))}</div>"
        "<div class='metric-label'>Of FY target by March</div>"
        f"<div class='kpi-rows'>{march_rows}</div></div>"
        "</div>",
        unsafe_allow_html=True,
    )
    return cell
 
 
def render_scenario_comparison(
    final_metrics: Dict[str, Any],
    model: ScenarioModel,
    basis: str,
    asset: str,
) -> None:
    """The management answer to: what changes if I choose this scenario?"""
    cell = model.cell(basis, asset=None if asset == "All" else asset)
 
    revenue_total = None
    if asset == "All":
        bundle = revenue_bundle(model, REVENUE_BASIS)
        revenue_total = bundle["incremental"]["total"]
    else:
        scenario_revenue = _z(cell.get("march_amount")) * REVENUE_RATE.get(asset, 0.0)
        baseline_revenue = _z(cell.get("current_march")) * REVENUE_RATE.get(asset, 0.0)
        revenue_total = scenario_revenue - baseline_revenue
 
    current_pct = _num(cell.get("current_march_pct"))
    scenario_pct = _num(cell.get("march_pct"))
    delta = None if (current_pct is None or scenario_pct is None) else scenario_pct - current_pct
 
    kpi_strip([
        {"label": "Current projected FY", "value": fmt_pct(current_pct),
         "secondary": fmt_cr(cell.get("current_march"))},
        {"label": "Scenario projected FY", "value": fmt_pct(scenario_pct),
         "delta": fmt_pts(delta), "tone": _tone_for(delta),
         "secondary": fmt_cr(cell.get("march_amount"))},
        {"label": "Incremental sales", "value": fmt_cr_signed(cell.get("incremental_sales")),
         "tone": _tone_for(cell.get("incremental_sales")),
         "secondary": "versus the current trajectory"},
        {"label": "Required run rate", "value": fmt_cr(cell.get("scen_rr")),
         "delta": fmt_pct_signed(cell.get("rr_change_pct")),
         "tone": _tone_for(cell.get("rr_change_pct")),
         "secondary": "monthly, from July"},
        {"label": "Revenue impact", "value": fmt_cr_signed(revenue_total, 1),
         "tone": _tone_for(revenue_total),
         "secondary": f"on {SALES_LABEL[REVENUE_BASIS]}, asset-class rates"},
    ])
 
    gs_frame, gs_formats = build_final_scenario_comparison(final_metrics, model, "GS")
    ns_frame, ns_formats = build_final_scenario_comparison(final_metrics, model, "NS")
 
    tabs = st.tabs(["Net Sales · current vs scenario", "Gross Sales · current vs scenario"])
    with tabs[0]:
        render_glass_table(ns_frame, ns_formats, total_rows=("Overall",))
    with tabs[1]:
        render_glass_table(gs_frame, gs_formats, total_rows=("Overall",))
 
    glass_note(
        "Scenario outcomes are anchored to the FINAL FY27 targets so the comparison uses one "
        "common base. Retail, DHNI and VRM carry scenario calculations from their detailed "
        "sheets; channels such as Insti and Digital stay visible as current FINAL metrics only."
    )
 
 
# =============================================================================
# 18. SCENARIO TRAJECTORY, REVENUE & DRIVER COMPONENTS
# =============================================================================
 
def scenario_trajectory(model: ScenarioModel, basis: str, asset: str) -> List[float]:
    """Monthly run rates July -> March for any scenario, on the selected cut."""
    asset_filter = None if asset == "All" else asset
    cell = model.cell(basis, asset=asset_filter)
 
    trajectory = cell.get("trajectory")
    if isinstance(trajectory, (list, tuple)) and len(trajectory) == len(FUTURE_MONTHS):
        return [float(_z(v)) for v in trajectory]
 
    frame = model.scenario_grid
    if frame is not None and "trajectory" in getattr(frame, "columns", []):
        mask = frame["Sales"] == basis
        if asset_filter is not None:
            mask &= frame["Asset"] == asset_filter
        totals = [0.0] * len(FUTURE_MONTHS)
        found = False
        for series in frame.loc[mask, "trajectory"]:
            if isinstance(series, (list, tuple)) and len(series) == len(FUTURE_MONTHS):
                found = True
                for position, value in enumerate(series):
                    totals[position] += _z(value)
        if found:
            return totals
 
    scen_rr = _z(cell.get("scen_rr"))
    feb_mar_rr = _z(cell.get("feb_mar_rr")) if cell.get("feb_mar_rr") is not None else scen_rr
    return [scen_rr] * MONTHS_JUL_JAN + [feb_mar_rr] * MONTHS_FEB_MAR
 
 
def trajectory_cell(model: ScenarioModel, basis: str, asset: str) -> Dict[str, Any]:
    """A scenario cell enriched with a trajectory and the right leakage label."""
    cell = dict(model.cell(basis, asset=None if asset == "All" else asset))
    cell["trajectory"] = scenario_trajectory(model, basis, asset)
    if cell.get("leakage") is None:
        if model.scenario_id == 3:
            cell["leakage"] = float(model.params.get("dip", S3_DEFAULT_DIP))
        elif model.scenario_id in (7, 8, 9):
            cell["leakage"] = float(model.params.get("leakage", S7_DEFAULT_LEAKAGE))
        else:
            cell["leakage"] = 0.0
    return cell
 
 
def render_scenario_trajectory(model: ScenarioModel, basis: str, asset: str) -> None:
    """07 · Monthly progression shown as data only; all graphs are removed."""
    section_header(
        "07",
        "Scenario trajectory",
        "Month by month from July 2026 to March 2027",
    )

    cell = trajectory_cell(model, basis, asset)
    frame, formats = build_momentum_analysis(cell)
    render_glass_table(frame, formats)

    if model.scenario_id == 7:
        render_momentum_detail(model, basis)

def render_momentum_detail(model: ScenarioModel, basis: str) -> None:
    """Scenario 7 specifics: binding milestone, buffer, leakage and sensitivity."""
    cell = model.cell(basis)
    momentum = cell.get("momentum_g")
 
    kpi_strip([
        {"label": "Required MoM momentum",
         "value": fmt_pct(momentum) if momentum is not None else NA_TEXT,
         "tone": "gold",
         "secondary": f"binding milestone: {cell.get('binding')}"},
        {"label": "January achievement", "value": fmt_pct(cell.get("jan_pct")),
         "delta": fmt_cr_signed(cell.get("jan_buffer")),
         "tone": _tone_for(cell.get("jan_buffer")),
         "secondary": "buffer created before leakage"},
        {"label": "Feb–Mar leakage", "value": fmt_pct(cell.get("leakage")),
         "secondary": f"Feb {fmt_cr(cell.get('feb_mar_rr'))} · Mar {fmt_cr(cell.get('march_rr'))}"},
        {"label": "March achievement", "value": fmt_pct(cell.get("march_pct")),
         "delta": fmt_pts(cell.get("headroom_pct")),
         "tone": _tone_for(cell.get("headroom_pct"))},
        {"label": "January exit run rate", "value": fmt_cr(cell.get("scen_rr")),
         "delta": fmt_pct_signed(cell.get("rr_change_pct")),
         "tone": _tone_for(cell.get("rr_change_pct"))},
    ])
 
    if cell.get("feasible"):
        glass_callout(
            "<span class='tag-ok'>Target achievable</span> — the momentum trajectory reaches the "
            f"January milestone and still clears the March ambition after "
            f"{fmt_pct(cell.get('leakage'))} February–March leakage.",
            tone="ok",
        )
    else:
        glass_callout(
            "<span class='tag-warn'>Additional momentum required</span> — March needs a further "
            f"{fmt_cr(cell.get('additional_march_sales'))}, which means lifting the January run "
            f"rate by {fmt_cr(cell.get('additional_jan_rr'))} per month.",
            tone="warn",
        )
    if cell.get("note"):
        glass_note(escape(str(cell["note"])))
 
    momentum_text = fmt_pct(momentum) if momentum is not None else "a flat required"
    outcome = (
        f"{fmt_pct(cell.get('headroom_pct'))} headroom" if _z(cell.get("headroom_amt")) >= 0
        else f"a {fmt_cr(abs(_z(cell.get('headroom_amt'))))} shortfall"
    )
    glass_callout(
        f"<b>What this asks of the organisation:</b> build roughly {momentum_text} month-on-month "
        f"momentum from July through January, lifting the monthly run rate from "
        f"{fmt_cr(cell.get('current_rr'))} to {fmt_cr(cell.get('scen_rr'))} by January to reach the "
        f"{fmt_pct(model.params.get('jan_target', S7_DEFAULT_JAN_TARGET))} milestone. After "
        f"{fmt_pct(cell.get('leakage'))} February–March leakage the trajectory lands at "
        f"{fmt_pct(cell.get('march_pct'))} of the FY target, leaving {outcome}."
    )
 
    tabs = st.tabs(["Asset class", "Retail / DHNI / VRM", "Leakage sensitivity", "Monthly revenue"])
    with tabs[0]:
        for sales in SALES_TYPES:
            st.markdown(
                f"<div class='metric-label'>{SALES_LABEL[sales]}</div>", unsafe_allow_html=True
            )
            frame, formats = build_momentum_by_group(model, sales, "asset")
            render_glass_table(frame, formats)
    with tabs[1]:
        for sales in SALES_TYPES:
            st.markdown(
                f"<div class='metric-label'>{SALES_LABEL[sales]}</div>", unsafe_allow_html=True
            )
            frame, formats = build_momentum_by_group(model, sales, "vertical")
            render_glass_table(frame, formats)
    with tabs[2]:
        frame, formats = build_leakage_sensitivity(model, basis)
        render_glass_table(frame, formats)
        glass_note(
            "Momentum is re-solved at each leakage assumption, so the required July–January "
            "build changes with the February–March pressure."
        )
    with tabs[3]:
        frame, formats = build_monthly_revenue(model, REVENUE_BASIS)
        render_glass_table(frame, formats)
        january_revenue = calculate_revenue(model.assets(REVENUE_BASIS), "jan_amount")
        march_revenue = calculate_revenue(model.assets(REVENUE_BASIS), "march_amount")
        baseline = calculate_baseline_revenue(model.assets(REVENUE_BASIS))
        glass_note(
            f"January scenario revenue {fmt_cr(january_revenue['total'], 1)} · "
            f"March scenario revenue {fmt_cr(march_revenue['total'], 1)} · baseline "
            f"{fmt_cr(baseline['total'], 1)} · incremental "
            f"{fmt_cr_signed(march_revenue['total'] - baseline['total'], 1)}."
        )
 
 
def render_revenue_impact(model: ScenarioModel) -> Dict[str, Any]:
    """08 · Revenue impact as KPI + table only; graph removed."""
    section_header(
        "08",
        "Revenue impact",
        "Equity 60 bps · Debt 20 bps · Liquid 10 bps on Net Sales",
    )

    bundle = revenue_bundle(model, REVENUE_BASIS)
    incremental = bundle["incremental"]

    kpi_strip([
        {"label": "Incremental revenue", "value": fmt_cr_signed(incremental["total"], 1),
         "tone": _tone_for(incremental["total"]),
         "delta": fmt_pct_signed(incremental["uplift_pct"]),
         "secondary": "versus the current run rate"},
        {"label": "Baseline revenue", "value": fmt_cr(bundle["baseline"]["total"], 1),
         "secondary": "current trajectory to March"},
        {"label": "Scenario revenue", "value": fmt_cr(bundle["scenario"]["total"], 1),
         "secondary": "selected scenario to March"},
        {"label": "January revenue", "value": fmt_cr(bundle["january"]["total"], 1),
         "secondary": "booked by the January milestone"},
    ])

    frame, formats = build_revenue_impact(model, REVENUE_BASIS)
    render_glass_table(frame, formats, total_rows=("Total",))

    parts = " + ".join(
        f"{asset} {fmt_cr_signed(incremental['by_asset'][asset], 1)}" for asset in ASSETS
    )
    contribution = " · ".join(
        f"{asset} {fmt_pct(incremental['contribution'][asset])}" for asset in ASSETS
    )
    glass_callout(
        f"<b>Revenue bridge:</b> baseline {fmt_cr(bundle['baseline']['total'], 1)} "
        f"+ {parts} = scenario {fmt_cr(bundle['scenario']['total'], 1)}.<br>"
        f"<b>Scenario revenue mix:</b> {contribution}."
    )
    return bundle

def render_all_scenarios(
    grid: pd.DataFrame,
    scenario_id: int,
    params: Dict[str, Any],
    basis: str,
    asset: str,
) -> None:
    """Every scenario evaluated on the selected scope, table only."""
    with st.expander("Compare all nine scenarios on this scope", expanded=False):
        frame, formats = build_all_scenario_matrix(grid, scenario_id, params, basis, asset)
        if frame.empty:
            glass_note("No scenario output is available for this scope.")
            return
        render_glass_table(
            frame, formats,
            total_rows=(f"{scenario_id:02d} · {SCENARIOS[scenario_id]['short']}",),
        )
        glass_note(
            "Scenarios 1–9 are calculated on the same scope and sales basis. The selected "
            "scenario uses the live assumptions above; the others use their configured defaults."
        )

# =============================================================================
# 19. SCENARIO-SPECIFIC DRIVERS
# =============================================================================
 
def render_segment_section(model: ScenarioModel, basis: str, counts: Dict[str, int]) -> None:
    """Scenario 6 · differentiated performance by business segment."""
    unmapped = [s for s in ("Digital", "Retail B30") if counts.get(s, 0) == 0]
    if unmapped:
        missing = " and ".join(unmapped)
        glass_callout(
            f"<b>Segment validation:</b> no records could be classified as {missing} from the "
            "workbook metadata (MKT TYPE, Type, ZONE, REGION, EM City). Those records remain in "
            f"Others, so the {missing} uplift is not applied. Point the classification at the "
            "right column under <i>Segment mapping</i> in the sidebar.",
            tone="warn",
        )
 
    present = " · ".join(
        f"{segment} {S6_SEGMENT_TARGETS[segment]:.0%} of FY target ({counts.get(segment, 0)} RMs)"
        for segment in SEGMENT_ORDER
    )
    glass_note(f"Scenario assumption — {present}.")
 
    tabs = st.tabs([SALES_LABEL["NS"], SALES_LABEL["GS"]])
    for tab, sales in zip(tabs, ["NS", "GS"]):
        with tab:
            frame, formats = build_segment_scenario_analysis(model, sales)
            render_glass_table(frame, formats, total_rows=("Overall",))
 
    overall = model.cell(basis)
    lines = []
    for segment in model.available_segments():
        cell = model.cell(basis, segment=segment)
        uplift = (
            fmt_pct_signed(cell["rr_change_pct"]) if cell["rr_change_pct"] is not None else NA_TEXT
        )
        lines.append(
            f"<b>{segment}</b> moves from {fmt_pct(cell['current_march_pct'])} to "
            f"{fmt_pct(cell['march_pct'])} of FY target, needing {fmt_cr(cell['scen_rr'])} per month "
            f"({uplift} run-rate uplift)"
        )
    glass_callout(
        f"On {SALES_LABEL[basis]}: " + "; ".join(lines) + ". Overall March achievement moves from "
        f"{fmt_pct(overall['current_march_pct'])} to {fmt_pct(overall['march_pct'])}, an improvement "
        f"of {fmt_cr_signed(overall['incremental_sales'])}."
    )
 
 
def render_channel_simulator(model: ScenarioModel, basis: str) -> None:
    """Scenario 8 · executive channel simulator, table only."""
    frame, formats = build_channel_scenario_analysis(model, basis)
    if frame.empty:
        glass_callout(
            "No mapped channel data is available. Use <i>Channel mapping</i> in the sidebar to "
            "classify the workbook into the nine planning channels.",
            tone="warn",
        )
        return

    jan_gap = _z(frame["Jan Gap / Headroom"].sum())
    mar_gap = _z(frame["Mar Gap / Headroom"].sum())
    incremental = _z(frame["March Incremental Sales"].sum())

    kpi_strip([
        {"label": "January portfolio headroom", "value": fmt_cr_signed(jan_gap),
         "tone": _tone_for(jan_gap), "secondary": "above or below the January target"},
        {"label": "March portfolio headroom", "value": fmt_cr_signed(mar_gap),
         "tone": _tone_for(mar_gap), "secondary": "above or below the March target"},
        {"label": "Incremental March sales", "value": fmt_cr_signed(incremental),
         "tone": _tone_for(incremental), "secondary": "versus the current projection"},
        {"label": "Channels in play", "value": f"{len(frame)} of {len(CHANNELS)}",
         "secondary": "mapped from workbook metadata"},
    ])

    render_glass_table(frame, formats)

    on_track = (frame["Jan Gap / Headroom"] >= 0).all() and (frame["Mar Gap / Headroom"] >= 0).all()
    if on_track:
        glass_callout(
            "<span class='tag-ok'>All channels on track</span> — the selected growth assumptions "
            "clear both the January and March targets after leakage.",
            tone="ok",
        )
    else:
        misses = frame.loc[
            (frame["Jan Gap / Headroom"] < 0) | (frame["Mar Gap / Headroom"] < 0), "Channel"
        ].tolist()
        glass_callout(
            "<span class='tag-warn'>Channel gap</span> — review " + ", ".join(misses)
            + ". Raise monthly growth or reset the target for those channels.",
            tone="warn",
        )

def render_channel_optimizer(model: ScenarioModel, basis: str) -> None:
    """Scenario 9 · minimum growth required, table only."""
    ambition = float(model.params.get("optimizer_target", 1.20))
    frame, formats = build_channel_scenario_analysis(model, basis)
    if frame.empty:
        glass_callout(
            "No mapped channel data is available. Map the workbook into channels in the sidebar "
            "to run the optimiser.",
            tone="warn",
        )
        return

    solved = [(_num(v) or 0.0) for v in frame["MoM Growth"]]
    hardest = frame.iloc[int(np.argmax(solved))] if solved else None
    portfolio = model.cell(basis)

    kpi_strip([
        {"label": "March ambition", "value": fmt_pct(ambition), "tone": "gold",
         "secondary": "portfolio target set by management"},
        {"label": "Minimum required growth", "value": fmt_pct(max(solved) if solved else None),
         "secondary": f"hardest channel: {hardest['Channel']}" if hardest is not None else ""},
        {"label": "Average required growth",
         "value": fmt_pct(float(np.mean(solved)) if solved else None),
         "secondary": "across mapped channels"},
        {"label": "January milestone", "value": fmt_pct(portfolio.get("jan_pct")),
         "secondary": "held at 100% of FY target"},
        {"label": "Feb–Mar leakage",
         "value": fmt_pct(model.params.get("leakage", S8_DEFAULT_LEAKAGE)),
         "secondary": "applied after January"},
    ])

    display = frame.copy()
    display["Portfolio ambition"] = ambition
    display_formats = dict(formats)
    display_formats["Portfolio ambition"] = "pct"
    render_glass_table(display, display_formats)

    glass_callout(
        "The optimiser solves the minimum compounding trajectory each channel must run from July "
        f"to hold the January milestone and still land the {fmt_pct(ambition)} March ambition after "
        f"{fmt_pct(model.params.get('leakage', S8_DEFAULT_LEAKAGE))} leakage. Channels already at "
        "or above the requirement solve to zero additional growth."
    )

def render_scenario_drivers(
    model: ScenarioModel,
    basis: str,
    segment_counts: Dict[str, int],
) -> None:
    """09 · Scenario data segregated only into Channel and Asset Class views."""
    section_header(
        "09",
        "Scenario drivers",
        "Where the required delivery actually sits",
    )

    if model.scenario_id == 6:
        render_segment_section(model, basis, segment_counts)
    elif model.scenario_id == 8:
        render_channel_simulator(model, basis)
    elif model.scenario_id == 9:
        render_channel_optimizer(model, basis)

    tabs = st.tabs(["Channel", "Asset class"])
    with tabs[0]:
        frame, formats = build_vertical_summary(model)
        if frame.empty:
            glass_note("No channel data is available for this scope.")
        else:
            render_glass_table(frame, formats)

    with tabs[1]:
        rows: List[Dict[str, Any]] = []
        for sales in ("NS", "GS"):
            for asset_name in ASSETS:
                cell = model.cell(sales, asset=asset_name)
                rows.append({
                    "Sales": SALES_LABEL[sales],
                    "Asset Class": asset_name,
                    "FY Target": cell.get("fy_target"),
                    "YTD Achievement": cell.get("ytd_ach"),
                    "Target Achieved %": cell.get("ytd_ach_pct"),
                    "Current Run Rate": cell.get("current_rr"),
                    "Scenario Run Rate": cell.get("scen_rr"),
                    "Run Rate Change %": cell.get("rr_change_pct"),
                    "Current March Projection %": cell.get("current_march_pct"),
                    "Scenario March Projection %": cell.get("march_pct"),
                    "Incremental Sales": cell.get("incremental_sales"),
                })
        asset_frame = pd.DataFrame(rows)
        render_glass_table(
            asset_frame,
            {
                "Sales": "txt", "Asset Class": "txt", "FY Target": "cr",
                "YTD Achievement": "cr", "Target Achieved %": "pct",
                "Current Run Rate": "cr", "Scenario Run Rate": "cr",
                "Run Rate Change %": "pct_signed", "Current March Projection %": "pct",
                "Scenario March Projection %": "pct", "Incremental Sales": "cr_signed",
            },
        )

def render_detail_tables(model: ScenarioModel) -> None:
    """10 · Detailed analytical tables, supporting the cards above."""
    section_header("10", "Detailed analytical tables", "The full numbers behind every card")
 
    tabs = st.tabs(["Current baseline", "Current vs scenario", "Scenario guide"])
    with tabs[0]:
        frame, formats = build_current_overview(model)
        render_glass_table(frame, formats)
    with tabs[1]:
        frame, formats = build_comparison(model)
        render_glass_table(frame, formats)
    with tabs[2]:
        guide = build_scenario_guide(model, REVENUE_BASIS)
        render_glass_table(guide, {c: "txt" for c in guide.columns})
 
 
def render_final_reference(payload: bytes) -> None:
    """11 · The workbook's own FINAL sheet, kept as the audit surface."""
    section_header("11", "FINAL workbook reference", "The uploaded source of truth, unaltered")
 
    with st.expander("Open the FINAL sheet", expanded=False):
        try:
            st.markdown(build_final_sheet_html(payload), unsafe_allow_html=True)
        except WorkbookError as error:
            st.error(str(error))
        except Exception:  # pragma: no cover - defensive
            st.error("The FINAL sheet could not be rendered from this workbook.")
 
    with st.expander("FINAL sheet as raw data", expanded=False):
        try:
            raw = load_final_sheet_frame(payload)
            st.dataframe(raw, hide_index=True, height=560, **_dataframe_kwargs())
        except Exception:  # pragma: no cover - defensive
            st.error("The FINAL sheet could not be read as raw data.")
 
 
def render_export(model: ScenarioModel, payload: bytes) -> None:
    """12 · Export the selected scenario and the source workbook."""
    section_header("12", "Export", "Take the analysis into the management pack")
 
    left, right = st.columns(2)
    with left:
        try:
            export_payload = make_export_excel(model, REVENUE_BASIS)
            st.download_button(
                f"Download scenario {model.scenario_id} analysis",
                data=export_payload,
                file_name=f"scenario_{model.scenario_id}_analysis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception:  # pragma: no cover - defensive
            st.warning("The scenario export could not be generated for this scope.")
    with right:
        st.download_button(
            "Download the uploaded workbook",
            data=payload,
            file_name="sales_command_center_source.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
 
    glass_note(
        "The scenario workbook contains the scenario guide, current baseline, current-vs-scenario "
        "comparison, revenue impact, Retail/DHNI/VRM summary, Gross and Net breakdowns, the "
        "Scenario 6 segment view, the Scenario 7 momentum trajectory and monthly revenue, the "
        "channel matrix where relevant, and all nine scenarios side by side."
    )
 
 
# =============================================================================
# 20. SIDEBAR - DATA MAPPING & NAVIGATION
# =============================================================================
 
def render_segment_controls(records: pd.DataFrame) -> Dict[str, Any]:
    """Configurable segment classification for Scenario 6."""
    suggestions = identify_segments(records)
    mapping: Dict[str, Any] = dict(st.session_state.get("segment_mapping") or suggestions)
 
    usable_columns = [
        field for field in META_FIELDS
        if field in records.columns and text_column(records, field).ne("").any()
    ]
 
    with st.sidebar.expander("Segment mapping", expanded=False):
        st.caption(
            "Digital, Retail B30 and Others are derived from workbook metadata. "
            "Everything unmatched falls into Others."
        )
        for segment in ("Digital", "Retail B30"):
            options = ["Not mapped"] + usable_columns
            current = mapping.get(segment, {}).get("column", "Not mapped")
            index = options.index(current) if current in options else 0
            column = st.selectbox(
                f"{segment} identified by", options, index=index, key=f"seg_col_{segment}"
            )
            if column == "Not mapped":
                mapping.pop(segment, None)
                continue
            values = sorted({v for v in text_column(records, column) if v.strip()})
            preset = [v for v in mapping.get(segment, {}).get("values", []) if v in values]
            chosen = st.multiselect(
                f"{segment} values", values, default=preset, key=f"seg_vals_{segment}_{column}"
            )
            if chosen:
                mapping[segment] = {"column": column, "values": list(chosen)}
            else:
                mapping.pop(segment, None)
 
    st.session_state["segment_mapping"] = mapping
    return mapping
 
 
def render_channel_controls(records: pd.DataFrame) -> Dict[str, Any]:
    """Channel mapping for the Scenario 8 / 9 planning universe."""
    suggestions = identify_channels(records)
    mapping: Dict[str, Any] = dict(st.session_state.get("channel_mapping") or suggestions)
    usable_columns = [
        f for f in META_FIELDS + ["Vertical"]
        if f in records.columns and text_column(records, f).ne("").any()
    ]
 
    with st.sidebar.expander("Channel mapping", expanded=False):
        st.caption(
            "Map workbook metadata to Digital, VRM, EM, B30, T30, T8, DHNI, Retail and "
            "Institutional. Unmapped rows stay unclassified."
        )
        for channel in CHANNELS:
            options = ["Not mapped"] + usable_columns
            current = mapping.get(channel, {}).get("column", "Not mapped")
            index = options.index(current) if current in options else 0
            column = st.selectbox(channel, options, index=index, key=f"ch_col_{channel}")
            if column == "Not mapped":
                mapping.pop(channel, None)
                continue
            values = sorted({v for v in text_column(records, column) if v.strip()})
            preset = [v for v in mapping.get(channel, {}).get("values", []) if v in values]
            chosen = st.multiselect(
                f"{channel} values", values, default=preset, key=f"ch_vals_{channel}_{column}"
            )
            if chosen:
                mapping[channel] = {"column": column, "values": list(chosen)}
            else:
                mapping.pop(channel, None)
 
    st.session_state["channel_mapping"] = mapping
    return mapping
 
 
def render_sidebar(records: pd.DataFrame) -> Tuple[str, Dict[str, Any], Dict[str, Any]]:
    """Quiet utility rail: page, data mapping, assumptions and workbook actions."""
    sidebar = st.sidebar
    sidebar.markdown("<div class='sidebar-mark'>Command Center</div>", unsafe_allow_html=True)
    sidebar.markdown("<div class='sidebar-title'>View</div>", unsafe_allow_html=True)
    page = sidebar.radio(
        "View",
        ["Executive command center", "RM performance"],
        index=0,
        key="application_page_selector",
        label_visibility="collapsed",
    )
 
    sidebar.markdown("<div class='sidebar-title'>Data mapping</div>", unsafe_allow_html=True)
    segment_mapping = render_segment_controls(records)
    channel_mapping = render_channel_controls(records)
 
    sidebar.markdown("<div class='sidebar-title'>Assumptions</div>", unsafe_allow_html=True)
    with sidebar.expander("Model assumptions", expanded=False):
        st.caption(
            "Revenue basis: Net Sales only. Rates: Equity 60 bps · Debt 20 bps · Liquid 10 bps. "
            "Timeline: April–June complete, nine months remaining, July–January is seven months "
            "and February–March is two months. The current run rate is YTD ÷ 3."
        )
 
    sidebar.markdown("<div class='sidebar-title'>Workbook</div>", unsafe_allow_html=True)
    if sidebar.button("Use another workbook", key="reset_workbook_button"):
        reset_workbook()
 
    return page, segment_mapping, channel_mapping
 
 
# =============================================================================
# 21. UPLOAD EXPERIENCE
# =============================================================================
 
def render_upload_screen() -> None:
    st.markdown(
        "<div class='exec-header'><div>"
        "<div class='exec-mark'>FY27 · Executive management view</div>"
        f"<div class='exec-title'>{escape(APP_TITLE)}</div>"
        f"<div class='exec-sub'>{escape(APP_SUBTITLE)}</div>"
        "</div></div>",
        unsafe_allow_html=True,
    )
 
    left, right = st.columns([1.35, 1])
    with left:
        st.markdown(
            "<div class='glass-card'>"
            "<div class='metric-label'>Start here</div>"
            "<div class='metric-hero'>Upload the RM scorecard</div>"
            "<div class='metric-secondary'>The workbook stays in this session only. "
            "Nothing is stored once the tab is closed.</div></div>",
            unsafe_allow_html=True,
        )
        uploaded = st.file_uploader(
            "Workbook", type=["xlsx", "xlsm"], label_visibility="collapsed"
        )
    with right:
        st.markdown(
            "<div class='glass-panel'>"
            "<div class='metric-label'>The workbook needs</div>"
            "<div class='metric-secondary'>"
            "<span class='inline-pill gold'>FINAL</span>"
            "<span class='inline-pill'>RM Retail Sales</span>"
            "<span class='inline-pill'>RM DHNI</span>"
            "<span class='inline-pill'>VRM</span><br><br>"
            "FINAL supplies Gross Sales and Net Sales targets. The three RM sheets supply "
            "employee-level targets and achievement used for the Asset Class and Channel views."
            "</div></div>",
            unsafe_allow_html=True,
        )
 
    if uploaded is not None:
        payload = uploaded.getvalue()
        try:
            load_workbook(payload)
            parse_final_dashboard_metrics(payload)
        except WorkbookError as error:
            st.error(str(error))
        except Exception:  # pragma: no cover - never show a traceback
            st.error(
                "That file could not be read as the RM scorecard workbook. Check that it is a "
                "valid Excel file containing RM Retail Sales, RM DHNI, VRM and FINAL."
            )
        else:
            st.session_state["workbook"] = payload
            rerun()
 
 
# =============================================================================
# 22. EXECUTIVE COMMAND CENTER - SCREEN FLOW
# =============================================================================
 
def render_command_center(records: pd.DataFrame, payload: bytes) -> None:
    """The full management journey, in the order an executive reads it."""
    final_metrics = parse_final_dashboard_metrics(payload)
 
    render_apple_header(final_metrics, records, "FY27 · Executive management view")
 
    full_grid = build_base_grid(records)
    base_params = scenario_default_params(1)
    base_model = ScenarioModel(1, full_grid, base_params)
 
    # 02 · Current reality, straight from FINAL - before any scenario exists.
    render_current_performance(final_metrics, base_model)
 
    # 03 · Drivers behind that reality.
    render_business_driver_selector(final_metrics, base_model)
    basis = st.session_state.get("display_basis", "NS")
 
    # 04 · Scope + the pace question.
    section_header(
        "04",
        "Current run rate & target gap",
        "The pace the business is running at, against the pace the target needs",
    )
    channel, location, asset, scoped_records = render_analysis_scope(records)
    scoped_grid = build_base_grid(scoped_records) if not scoped_records.empty else pd.DataFrame()
 
    if scoped_grid.empty:
        render_final_reference(payload)
        return
 
    scoped_cell = summarize_current(
        scoped_grid, sales=basis, asset=None if asset == "All" else asset
    )
    required_rr = _z(scoped_cell.get("fy_target")) / 12.0
    current_rr = _num(scoped_cell.get("current_rr"))
    pace_gap = (current_rr / required_rr - 1.0) if (current_rr is not None and required_rr) else None
    projected = _num(scoped_cell.get("current_march_pct"))
    shortfall = _z(scoped_cell.get("fy_target")) - _z(scoped_cell.get("current_march"))
 
    kpi_strip([
        {"label": "FY target", "value": fmt_cr(scoped_cell.get("fy_target")),
         "secondary": f"{SALES_LABEL[basis]} · RM calculation sheets"},
        {"label": "YTD achievement", "value": fmt_cr(scoped_cell.get("ytd_ach")),
         "delta": fmt_pct(scoped_cell.get("ytd_ach_pct")), "tone": "gold",
         "secondary": "against the YTD June target"},
        {"label": "Current run rate", "value": fmt_cr(current_rr),
         "secondary": "YTD ÷ 3 completed months"},
        {"label": "Required run rate", "value": fmt_cr(required_rr),
         "delta": fmt_pct_signed(pace_gap), "tone": _tone_for(pace_gap),
         "secondary": "pace gap on today's run rate"},
        {"label": "Projected March", "value": fmt_pct(projected),
         "delta": fmt_pts(None if projected is None else projected - 1.0),
         "tone": _tone_for(None if projected is None else projected - 1.0),
         "secondary": fmt_cr(scoped_cell.get("current_march"))},
        {"label": "Gap at current pace", "value": fmt_cr(shortfall),
         "tone": _tone_for(-shortfall),
         "secondary": "FY target less projected March"},
    ])
 
    glass_note(
        "The current run rate is completed Apr–Jun achievement divided by three. The required "
        "run rate is the FY target spread evenly across twelve months."
    )
 
    # 05 · Scenario planning.
    scenario_id = render_scenario_navigator()
    params = render_scenario_controls(scenario_id, scenario_default_params(scenario_id))
    params["channel_mapping"] = st.session_state.get("channel_mapping", {})
 
    try:
        model = ScenarioModel(scenario_id, scoped_grid, params)
    except Exception:  # pragma: no cover - defensive
        st.error("This scenario could not be calculated on the current scope. Widen the scope or "
                 "select another scenario.")
        render_final_reference(payload)
        return
 
    # 06 · Selected scenario.
    render_scenario_hero(model, basis, asset)
    render_scenario_comparison(final_metrics, model, basis, asset)
    if channel != "All" or location != "All":
        glass_note(
            "The scope filter applies to the RM calculation sheets. FINAL targets in the table "
            "above remain the published portfolio targets, so scenario percentages are the "
            "comparable figures while a filter is active."
        )
    render_all_scenarios(scoped_grid, scenario_id, params, basis, asset)
 
    # 07 · Trajectory.
    render_scenario_trajectory(model, basis, asset)
 
    # 08 · Revenue.
    render_revenue_impact(model)
 
    # 09 · Drivers behind the scenario.
    render_scenario_drivers(model, basis, segment_diagnostics(scoped_records))
 
    # 10-12 · Tables, source workbook, export.
    render_detail_tables(model)
    render_final_reference(payload)
    render_export(model, payload)
 
 
# =============================================================================
# 23. RM PERFORMANCE SEGMENTATION PAGE
# =============================================================================
 
def _clean_filter_values(series: pd.Series) -> List[str]:
    """Return sorted non-empty values for RM filter dropdowns."""
    cleaned = series.astype(str).str.replace("\u00a0", " ", regex=False).str.strip()
    cleaned = cleaned[
        cleaned.ne("")
        & cleaned.str.lower().ne("nan")
        & cleaned.str.lower().ne("none")
    ]
    return sorted(cleaned.drop_duplicates().tolist(), key=lambda value: value.lower())
 
 
def _apply_exact_text_filter(frame: pd.DataFrame, column: str, selected: str) -> pd.DataFrame:
    """Apply one exact dropdown filter while treating 'All' as no filter."""
    if selected == "All" or column not in frame.columns:
        return frame
    values = frame[column].astype(str).str.replace("\u00a0", " ", regex=False).str.strip()
    return frame.loc[values == selected].copy()
 
 
def render_retail_rm_filters(records: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """Cascading Retail filters: ZONE -> REGION -> MKT TYPE."""
    retail = records.loc[records["Vertical"] == "Retail"].copy()
    selections: Dict[str, str] = {"ZONE": "All", "REGION": "All", "MKT TYPE": "All"}
 
    st.markdown("<div class='metric-label'>Retail filters</div>", unsafe_allow_html=True)
    columns = st.columns(3)
 
    zone_options = ["All"] + (
        _clean_filter_values(retail["ZONE"]) if "ZONE" in retail.columns else []
    )
    with columns[0]:
        reset_stale_selection("retail_rm_zone_filter", zone_options)
        selections["ZONE"] = st.selectbox("Zone", zone_options, index=0, key="retail_rm_zone_filter")
    filtered_retail = _apply_exact_text_filter(retail, "ZONE", selections["ZONE"])
 
    region_options = ["All"] + (
        _clean_filter_values(filtered_retail["REGION"]) if "REGION" in filtered_retail.columns else []
    )
    with columns[1]:
        selections["REGION"] = st.selectbox(
            "Region", region_options, index=0,
            key=f"retail_rm_region_filter_{selections['ZONE']}",
        )
    filtered_retail = _apply_exact_text_filter(filtered_retail, "REGION", selections["REGION"])
 
    market_options = ["All"] + (
        _clean_filter_values(filtered_retail["MKT TYPE"]) if "MKT TYPE" in filtered_retail.columns else []
    )
    with columns[2]:
        selections["MKT TYPE"] = st.selectbox(
            "Market type", market_options, index=0,
            key=f"retail_rm_market_filter_{selections['ZONE']}_{selections['REGION']}",
        )
    filtered_retail = _apply_exact_text_filter(filtered_retail, "MKT TYPE", selections["MKT TYPE"])
 
    non_retail = records.loc[records["Vertical"] != "Retail"].copy()
    filtered_records = pd.concat([filtered_retail, non_retail], ignore_index=True, sort=False)
 
    active = [f"{column}: {value}" for column, value in selections.items() if value != "All"]
    active_text = " · ".join(active) if active else "All Retail RMs"
    st.markdown(
        f"<div class='glass-note'><span class='inline-pill gold'>{len(filtered_retail):,} Retail RMs"
        f"</span><span class='inline-pill'>{escape(active_text)}</span></div>",
        unsafe_allow_html=True,
    )
    return filtered_records, selections
 
 
def _star_card(rank: int, row: pd.Series) -> str:
    return (
        "<div class='glass-card stage-card jan'>"
        f"<div class='stage'>Star {rank:02d}</div>"
        f"<div class='metric-value'>{escape(str(row.get('Employee Name', 'RM')))}</div>"
        f"<div class='metric-hero gold' style='font-size:1.7rem'>"
        f"{escape(fmt_pct(row.get('YTD Achievement %')))}</div>"
        "<div class='metric-label'>YTD achievement</div>"
        "<div class='kpi-rows'>"
        + _kpi_row_html("Current run rate", fmt_cr(row.get("Current Run Rate")))
        + _kpi_row_html("Projected FY", fmt_pct(row.get("Projected FY Achievement %")))
        + _kpi_row_html(
            "Contribution", fmt_pct(row.get("Contribution to Overall Target %")), "gold"
        )
        + "</div></div>"
    )
 
 
def render_stars_of_month(detail: pd.DataFrame) -> None:
    section_header("05", "Stars of the month", "Ranked on YTD achievement against YTD target")
 
    if detail.empty:
        glass_note("No RM data is available for this selection.")
        return
 
    ranked = detail.sort_values(
        ["YTD Achievement %", "YTD Achievement"], ascending=[False, False]
    ).reset_index(drop=True)
 
    cards = "".join(_star_card(i + 1, row) for i, (_, row) in enumerate(ranked.head(3).iterrows()))
    st.markdown(f"<div class='trio-grid'>{cards}</div>", unsafe_allow_html=True)
 
    glass_note(
        "Stars are ranked from the fields the scorecard actually provides: overall YTD achievement "
        "against YTD target, with the achievement amount as the tie-breaker. The workbook has no "
        "standalone monthly actual, so no separate monthly score is invented."
    )
 
    columns = [
        c for c in [
            "Employee Name", "Emp Code", "ADID", "REGION", "EM City",
            "Achievement Category", "YTD Achievement %", "Current Run Rate",
            "Projected FY Achievement %", "Contribution to Overall Target %",
        ] if c in ranked.columns
    ]
    render_glass_table(
        ranked.head(10)[columns],
        {
            "Employee Name": "txt", "Emp Code": "txt", "ADID": "txt", "REGION": "txt",
            "EM City": "txt", "Achievement Category": "txt", "YTD Achievement %": "pct",
            "Current Run Rate": "cr", "Projected FY Achievement %": "pct",
            "Contribution to Overall Target %": "pct",
        },
    )
 
 
def render_rm_sales_segmentation(
    records: pd.DataFrame,
    final_metrics: Dict[str, Any],
    vertical: str,
    sales: str,
) -> None:
    final_target = _final_vertical_target(final_metrics, sales, vertical)
    detail = build_rm_performance_detail(records, vertical, sales, final_target)
    if detail.empty:
        glass_note(f"No RM records are available for {vertical} · {SALES_LABEL[sales]}.")
        return
 
    contribution = build_category_contribution(detail, final_target)
 
    projected_total = _num(detail["Estimated FY @ Current RR"].sum()) or 0.0
    ytd_total = _num(detail["YTD Achievement"].sum()) or 0.0
    ytd_target_total = _num(detail["YTD Target"].sum()) or 0.0
    rm_count = len(detail)
    ytd_pct = ytd_total / ytd_target_total if ytd_target_total > 0 else 0.0
    denominator = final_target if (final_target or 0) > 0 else (_num(detail["FY Target"].sum()) or 1.0)
    projected_pct = projected_total / denominator
    high_performers = int(
        detail["Achievement Category"].isin(["100% and above", "90% - 100%"]).sum()
    )
 
    kpi_strip([
        {"label": "RMs in view", "value": fmt_num(rm_count),
         "secondary": f"{vertical} · {SALES_LABEL[sales]}"},
        {"label": "YTD vs YTD target", "value": fmt_pct(ytd_pct),
         "secondary": fmt_cr(ytd_total)},
        {"label": "Projected FY achievement", "value": fmt_pct(projected_pct),
         "delta": fmt_pts(projected_pct - 1.0), "tone": _tone_for(projected_pct - 1.0),
         "secondary": fmt_cr(projected_total)},
        {"label": "RMs at 90% or better", "value": fmt_num(high_performers),
         "secondary": fmt_pct(high_performers / rm_count if rm_count else 0.0)},
        {"label": "FINAL FY27 target", "value": fmt_cr(final_target),
         "secondary": "published channel target"},
    ])
 
    section_header("02", "Achievement bands", "How the RM population is distributed")
 
    counts = dict(zip(contribution["Achievement Category"], contribution["RM Count"]))
    kpi_strip([
        {"label": band, "value": fmt_num(counts.get(band, 0)), "secondary": "RMs"}
        for band in ACHIEVEMENT_BAND_ORDER
    ])
 
    section_header("03", "Run-rate contribution to target", "If each band holds its current pace")
    glass_note(
        "Each band's RMs are annualised at their current run rate and divided by the FY27 target "
        "from FINAL for the same channel and sales basis. The result is how many percentage points "
        "of the overall target each band is on track to deliver."
    )
    render_glass_table(
        contribution,
        {
            "Achievement Category": "txt", "RM Count": "num", "FY Target": "cr",
            "YTD Target": "cr", "YTD Achievement": "cr", "Current YTD Achievement %": "pct",
            "Current Run Rate": "cr", "Estimated FY @ Current RR": "cr",
            "Category Projected FY %": "pct", "Contribution to Overall Target %": "pct",
            "Share of Projected Sales %": "pct",
        },
    )
 
    total_contribution = _num(contribution["Contribution to Overall Target %"].sum()) or 0.0
    glass_callout(
        f"<b>{vertical} · {SALES_LABEL[sales]}:</b> at today's RM run rates the six bands together "
        f"are on track to deliver <b>{fmt_pct(total_contribution)}</b> of the FINAL FY27 target"
        + (f" ({fmt_cr(final_target)})." if final_target is not None else ".")
    )
 
    section_header("04", "RM drill-down", "Every RM inside the selected band")
    band = st.radio(
        "Achievement band", ACHIEVEMENT_BAND_ORDER, index=0, horizontal=True,
        key=f"rm_band_{vertical}_{sales}", label_visibility="collapsed",
    )
    rows = detail.loc[detail["Achievement Category"] == band].copy()
    if rows.empty:
        glass_note(f"No RMs fall in {band} for this selection.")
    else:
        columns = [
            c for c in [
                "Employee Name", "Emp Code", "ADID", "ZONE", "REGION", "EM City",
                "Achievement Category", "FY Target", "YTD Target", "YTD Achievement",
                "YTD Achievement %", "Current Run Rate", "Estimated FY @ Current RR",
                "Projected FY Achievement %", "Contribution to Overall Target %",
            ] if c in rows.columns
        ]
        render_glass_table(
            rows[columns],
            {
                "Employee Name": "txt", "Emp Code": "txt", "ADID": "txt", "ZONE": "txt",
                "REGION": "txt", "EM City": "txt", "Achievement Category": "txt",
                "FY Target": "cr", "YTD Target": "cr", "YTD Achievement": "cr",
                "YTD Achievement %": "pct", "Current Run Rate": "cr",
                "Estimated FY @ Current RR": "cr", "Projected FY Achievement %": "pct",
                "Contribution to Overall Target %": "pct",
            },
        )
 
    render_stars_of_month(detail)
 
 
def render_rm_segmentation_page(records: pd.DataFrame, payload: bytes) -> None:
    """RM achievement segmentation and contribution analysis."""
    final_metrics = parse_final_dashboard_metrics(payload)
    render_apple_header(final_metrics, records, "FY27 · RM performance")
 
    section_header(
        "01",
        "RM performance segmentation",
        "Retail · DHNI · VRM, banded by achievement against the YTD target",
    )
    st.markdown(
        "".join(
            f"<span class='inline-pill{' gold' if band.startswith('100') else ''}'>{escape(band)}</span>"
            for band in ACHIEVEMENT_BAND_ORDER
        ),
        unsafe_allow_html=True,
    )
 
    st.markdown("<div class='metric-label'>Channel</div>", unsafe_allow_html=True)
    vertical = st.radio(
        "Channel", VERTICALS, index=0, horizontal=True,
        key="rm_seg_vertical", label_visibility="collapsed",
    )
 
    page_records = records
    selections: Dict[str, str] = {"ZONE": "All", "REGION": "All", "MKT TYPE": "All"}
    if vertical == "Retail":
        page_records, selections = render_retail_rm_filters(records)
 
    tabs = st.tabs([SALES_LABEL["NS"], SALES_LABEL["GS"]])
    for tab, sales in zip(tabs, ["NS", "GS"]):
        with tab:
            render_rm_sales_segmentation(page_records, final_metrics, vertical, sales)
 
    section_header("06", "Export", "The full RM segmentation pack")
    try:
        export_payload = make_rm_segmentation_export(page_records, final_metrics)
        st.download_button(
            "Download RM segmentation workbook",
            data=export_payload,
            file_name="rm_performance_segmentation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        if vertical == "Retail":
            active = [f"{k}: {v}" for k, v in selections.items() if v != "All"]
            st.caption(
                "Retail sheets follow the filters above"
                + (f" ({' · '.join(active)})." if active else " (all Retail RMs).")
                + " DHNI and VRM remain unfiltered."
            )
    except Exception:  # pragma: no cover - defensive
        st.warning("The RM segmentation export could not be generated for this selection.")
 
    glass_note("Undefined numeric outputs are shown as 0.")
 
 
# =============================================================================
# 24. APPLICATION ENTRY POINT
# =============================================================================
 
def reset_workbook() -> None:
    for key in ("workbook", "segment_mapping", "channel_mapping", "application_page_selector"):
        st.session_state.pop(key, None)
    rerun()
 
 
def main() -> None:
    st.set_page_config(
        page_title=APP_TITLE,
        page_icon="\u25c6",
        layout="wide",
        initial_sidebar_state="collapsed",
    )
    inject_theme()
 
    if "workbook" not in st.session_state:
        render_upload_screen()
        return
 
    payload = st.session_state["workbook"]
 
    try:
        records = load_workbook(payload)
    except WorkbookError as error:
        st.error(str(error))
        if st.button("Use another workbook"):
            reset_workbook()
        return
    except Exception:  # pragma: no cover - never surface a traceback
        st.error(
            "This workbook could not be read. Upload the standard RM scorecard containing "
            "RM Retail Sales, RM DHNI, VRM and FINAL."
        )
        if st.button("Use another workbook"):
            reset_workbook()
        return
 
    page, segment_mapping, channel_mapping = render_sidebar(records)
    records = map_business_segments(records, segment_mapping)
    records = map_business_channels(records, channel_mapping)
 
    try:
        if page == "RM performance":
            render_rm_segmentation_page(records, payload)
        else:
            render_command_center(records, payload)
    except WorkbookError as error:
        st.error(str(error))
    except Exception:  # pragma: no cover - never surface a traceback
        st.error(
            "This view could not be prepared from the uploaded workbook. Select another scenario "
            "or scope, or upload a workbook that matches the standard RM scorecard format."
        )
 
 
if __name__ == "__main__":
    main()
