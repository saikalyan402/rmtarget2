"""
theme.py — Sales Performance Command Center · presentation layer
=================================================================

Single source of truth for the dashboard's visual system.

CONTRACT WITH THE APPLICATION
-----------------------------
This module is *display only*. It contains no business logic and it never
touches a numeric value.

  * Every component helper below accepts **already-formatted strings**
    produced by the app's own formatters (`fmt_cr`, `fmt_pct_signed`,
    `format_table`, ...). It renders them verbatim.
  * Nothing here rounds, abbreviates, converts, recalculates or renames.
    If `fmt_cr` returns "₹ 1,284 Cr", this module renders "₹ 1,284 Cr".
  * `NA_TEXT` ("0") passes through unchanged.

The only deliberate exception is *typographic*: numeric text is rendered with
tabular figures so digits align in columns. That changes glyph width, not value.

USAGE
-----
    import theme

    st.set_page_config(**theme.PAGE_CONFIG)
    theme.inject()                       # once, immediately after set_page_config

    theme.page_header(
        eyebrow="FY27 · EXECUTIVE MANAGEMENT VIEW",
        title=APP_TITLE,
        subtitle=APP_SUBTITLE,
    )

    theme.kpi_card(
        label="Net Sales",
        value=fmt_cr(cell["ytd_ach"]),          # formatted by the app
        delta=fmt_pct_signed(cell["rr_change_pct"]),
        delta_state=theme.state_of(cell["rr_change_pct"]),
        context="vs target",
        footnote_label="Target",
        footnote_value=fmt_cr(cell["fy_target"]),
        emphasis="primary",
    )
"""

from __future__ import annotations

import inspect
from contextlib import contextmanager
from html import escape
from typing import Any, Dict, Iterable, List, Optional, Sequence

import streamlit as st

__all__ = [
    "PAGE_CONFIG", "inject", "PLOTLY_TEMPLATE", "apply_chart_theme",
    "page_header", "section_header", "glass", "kpi_card", "kpi_row",
    "delta_pill", "status_badge", "sheet_chips", "upload_intro",
    "upload_success", "upload_error", "glass_table", "chart_frame",
    "empty_state", "utility_bar", "state_of", "GOOGLE_MARK",
    # colour constants re-exported for the Plotly / logic layer
    "INK", "INK_SOFT", "INK_MUTED", "GOLD", "GOLD_SOFT", "GREEN", "RED",
    "AMBER", "NEUTRAL", "GRID_LINE", "CHART_SEQUENCE",
]


def _supports_parameter(func: Any, parameter: str) -> bool:
    """Return True when a callable exposes a named parameter.

    Streamlit adds optional UI parameters over time.  Inspecting the installed
    version keeps this presentation module compatible across local and cloud
    deployments without pinning the application to one exact Streamlit build.
    """
    try:
        return parameter in inspect.signature(func).parameters
    except (TypeError, ValueError):
        return False


# =============================================================================
# 1. DESIGN TOKENS
# =============================================================================
#
# Accent decision, stated so it is a choice rather than a default:
#
# The brief suggests blue / indigo / cyan. This dashboard already carries a
# muted gold (#D8B76A) as its accent, and gold is the more defensible choice
# for "quiet luxury + enterprise" — a saturated cyan on near-black is the exact
# signature of the crypto/gaming dashboards the brief rules out. So gold stays
# as the single accent, and the cool tones are spent where the brief actually
# wants them: as extremely low-opacity ambient light in the background, and as
# the neutral semantic state. One accent, one warm/cool tension, no rainbow.
#
# The existing constant NAMES are preserved verbatim so this module is a
# drop-in for any code that already imports INK / GOLD / GRID_LINE etc.

# --- Surfaces -----------------------------------------------------------------
BG_BASE = "#080A0D"          # page floor
BG_RAISE = "#0B0D10"         # ambient mid-tone
BG_LIFT = "#0F1115"          # highest ambient tone

# --- Ink ----------------------------------------------------------------------
INK = "#F5F7FA"              # primary text
INK_SOFT = "#A8ADB7"         # secondary text
INK_MUTED = "#707681"        # metadata, axis labels
INK_FAINT = "#4C525C"        # disabled, hairlines with text

# --- Accent -------------------------------------------------------------------
GOLD = "#D8B76A"
GOLD_SOFT = "#C9AA65"
GOLD_DIM = "rgba(216, 183, 106, 0.14)"
GOLD_EDGE = "rgba(216, 183, 106, 0.32)"

# --- Semantic -----------------------------------------------------------------
GREEN = "#63D99A"
GREEN_DIM = "rgba(99, 217, 154, 0.12)"
RED = "#FF6B6B"
RED_DIM = "rgba(255, 107, 107, 0.12)"
AMBER = "#E0B252"            # muted, never neon
AMBER_DIM = "rgba(224, 178, 82, 0.12)"
NEUTRAL = "#7C8899"          # slate-blue, the cool counterweight to gold
NEUTRAL_DIM = "rgba(124, 136, 153, 0.12)"

# --- Glass --------------------------------------------------------------------
GLASS_1_BG = "rgba(255, 255, 255, 0.045)"    # primary surfaces
GLASS_2_BG = "rgba(255, 255, 255, 0.028)"    # secondary surfaces
GLASS_3_BG = "rgba(255, 255, 255, 0.055)"    # interactive controls
GLASS_1_EDGE = "rgba(255, 255, 255, 0.085)"
GLASS_2_EDGE = "rgba(255, 255, 255, 0.055)"
GLASS_3_EDGE = "rgba(255, 255, 255, 0.10)"

# --- Charts -------------------------------------------------------------------
GRID_LINE = "rgba(255, 255, 255, 0.055)"
AXIS_LINE = "rgba(255, 255, 255, 0.10)"

# Ordered so the first series is always the accent, then cool neutrals.
# Deliberately not a rainbow: one warm lead, four cool supports.
CHART_SEQUENCE: List[str] = [
    GOLD,        # lead series
    "#7C8899",   # slate
    "#5E7A93",   # steel
    "#8E7FA8",   # muted violet
    "#5F8C86",   # muted teal
    "#9A8F7A",   # sand
]

# --- Streamlit page config ----------------------------------------------------
PAGE_CONFIG: Dict[str, Any] = {
    "page_title": "Sales Performance Command Center",
    "layout": "wide",
    "initial_sidebar_state": "collapsed",
}


def state_of(value: Any, neutral_band: float = 1e-9) -> str:
    """
    Map a raw numeric to a semantic state name for styling purposes only.

    Returns "positive" / "negative" / "neutral". The *value* is never
    displayed from here — pass the app-formatted string separately.
    """
    try:
        if value is None:
            return "neutral"
        v = float(value)
    except (TypeError, ValueError):
        return "neutral"
    if v > neutral_band:
        return "positive"
    if v < -neutral_band:
        return "negative"
    return "neutral"


# =============================================================================
# 2. STYLESHEET
# =============================================================================

_CSS = f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

:root {{
  /* --- surfaces --- */
  --bg-base:        {BG_BASE};
  --bg-raise:       {BG_RAISE};
  --bg-lift:        {BG_LIFT};

  /* --- glass --- */
  --glass-1:        {GLASS_1_BG};
  --glass-2:        {GLASS_2_BG};
  --glass-3:        {GLASS_3_BG};
  --glass-1-edge:   {GLASS_1_EDGE};
  --glass-2-edge:   {GLASS_2_EDGE};
  --glass-3-edge:   {GLASS_3_EDGE};

  /* --- ink --- */
  --text-primary:   {INK};
  --text-secondary: {INK_SOFT};
  --text-muted:     {INK_MUTED};
  --text-faint:     {INK_FAINT};

  /* --- accent --- */
  --accent:         {GOLD};
  --accent-soft:    {GOLD_SOFT};
  --accent-dim:     {GOLD_DIM};
  --accent-edge:    {GOLD_EDGE};

  /* --- semantic --- */
  --success:        {GREEN};
  --success-dim:    {GREEN_DIM};
  --danger:         {RED};
  --danger-dim:     {RED_DIM};
  --warning:        {AMBER};
  --warning-dim:    {AMBER_DIM};
  --neutral:        {NEUTRAL};
  --neutral-dim:    {NEUTRAL_DIM};

  /* --- radius --- */
  --radius-sm:      11px;   /* controls  */
  --radius-md:      18px;   /* cards     */
  --radius-lg:      24px;   /* sections  */
  --radius-pill:    999px;

  /* --- blur --- */
  --blur-sm:        10px;
  --blur-md:        18px;
  --blur-lg:        26px;

  /* --- shadow: diffused, layered, never plastic --- */
  --shadow-1:       0 1px 2px rgba(0,0,0,0.20);
  --shadow-2:       0 6px 20px -8px rgba(0,0,0,0.42);
  --shadow-3:       0 18px 48px -20px rgba(0,0,0,0.58);

  /* --- spacing scale (strict) --- */
  --space-1: 4px;  --space-2: 8px;  --space-3: 12px; --space-4: 16px;
  --space-5: 24px; --space-6: 32px; --space-7: 48px; --space-8: 64px;

  /* --- type --- */
  --font: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;

  /* --- layout --- */
  --page-max: 1440px;
  --grid-gap: 24px;

  --ease: cubic-bezier(0.32, 0.72, 0, 1);
}}

/* ===========================================================================
   BACKGROUND — quiet, with three barely-there ambient pools
   =========================================================================== */

html, body, [data-testid="stAppViewContainer"] {{
  background: var(--bg-base);
  font-family: var(--font);
  color: var(--text-primary);
  -webkit-font-smoothing: antialiased;
}}

[data-testid="stAppViewContainer"]::before {{
  content: "";
  position: fixed;
  inset: 0;
  pointer-events: none;
  z-index: 0;
  background:
    radial-gradient(900px 620px at 12% -8%,  rgba(94, 122, 147, 0.10), transparent 62%),
    radial-gradient(760px 520px at 88% 4%,   rgba(142, 127, 168, 0.075), transparent 60%),
    radial-gradient(1200px 760px at 50% 108%, rgba(216, 183, 106, 0.045), transparent 66%),
    linear-gradient(180deg, var(--bg-raise) 0%, var(--bg-base) 46%, var(--bg-lift) 100%);
}}
[data-testid="stAppViewContainer"] > * {{ position: relative; z-index: 1; }}

[data-testid="stHeader"] {{ background: transparent; }}
[data-testid="stToolbar"] {{ right: var(--space-3); }}
#MainMenu, footer {{ visibility: hidden; }}

/* Controlled content width — an application, not a stretched page */
.block-container {{
  max-width: var(--page-max) !important;
  padding: var(--space-7) var(--space-6) var(--space-8) !important;
}}
@media (min-width: 1600px) {{
  .block-container {{ padding-left: var(--space-8) !important;
                      padding-right: var(--space-8) !important; }}
}}
@media (max-width: 900px) {{
  .block-container {{ padding: var(--space-6) var(--space-5) var(--space-7) !important; }}
}}

/* ===========================================================================
   TYPOGRAPHY — one family, one scale. Flat type against glass containers.
   =========================================================================== */

.sp-eyebrow {{
  font-size: 11px;
  font-weight: 600;
  letter-spacing: 0.16em;
  text-transform: uppercase;
  color: var(--text-muted);
  margin: 0 0 var(--space-3) 0;
}}
.sp-title {{
  font-size: 40px;
  font-weight: 650;
  line-height: 1.08;
  letter-spacing: -0.024em;
  color: var(--text-primary);
  margin: 0 0 var(--space-3) 0;
}}
.sp-subtitle {{
  font-size: 15px;
  font-weight: 400;
  line-height: 1.55;
  color: var(--text-secondary);
  margin: 0;
  max-width: 68ch;
}}
.sp-section-title {{
  font-size: 22px;
  font-weight: 600;
  letter-spacing: -0.014em;
  color: var(--text-primary);
  margin: 0;
}}
.sp-section-desc {{
  font-size: 13px;
  font-weight: 400;
  line-height: 1.6;
  color: var(--text-muted);
  margin: var(--space-2) 0 0 0;
  max-width: 76ch;
}}
.sp-card-title {{
  font-size: 17px;
  font-weight: 600;
  letter-spacing: -0.008em;
  color: var(--text-primary);
  margin: 0;
}}
.sp-label {{
  font-size: 11px;
  font-weight: 600;
  letter-spacing: 0.11em;
  text-transform: uppercase;
  color: var(--text-muted);
  margin: 0;
}}
.sp-body {{ font-size: 14px; line-height: 1.62; color: var(--text-secondary); }}
.sp-meta {{ font-size: 12px; line-height: 1.5;  color: var(--text-muted); }}

@media (max-width: 900px) {{
  .sp-title {{ font-size: 30px; }}
  .sp-section-title {{ font-size: 19px; }}
}}

/* Numeric typography — tabular figures so columns of digits align.
   Affects glyph width only, never the value. */
.sp-num, .sp-kpi-value, .sp-glass-table td.num, .sp-pill {{
  font-variant-numeric: tabular-nums;
  font-feature-settings: "tnum" 1, "cv05" 1;
}}

/* Streamlit's own markdown headings inherit the scale */
.stMarkdown h1 {{ font-size: 32px; font-weight: 650; letter-spacing: -0.02em; }}
.stMarkdown h2 {{ font-size: 22px; font-weight: 600; letter-spacing: -0.014em; }}
.stMarkdown h3 {{ font-size: 17px; font-weight: 600; }}
.stMarkdown p  {{ font-size: 14px; line-height: 1.62; color: var(--text-secondary); }}

/* ===========================================================================
   HEADER
   =========================================================================== */

.sp-header {{ margin: 0 0 var(--space-7) 0; }}      /* no glass around the header */

.sp-header-row {{
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
  gap: var(--space-6);
  flex-wrap: wrap;
}}
.sp-header-text {{ flex: 1 1 480px; min-width: 0; }}

/* Utility actions sit apart from the title and read as secondary */
.sp-utility {{ display: flex; gap: var(--space-2); flex: 0 0 auto; padding-top: var(--space-2); }}
.sp-utility a {{
  display: inline-flex;
  align-items: center;
  gap: var(--space-2);
  height: 36px;
  padding: 0 var(--space-3);
  border-radius: var(--radius-sm);
  background: var(--glass-3);
  border: 1px solid var(--glass-3-edge);
  backdrop-filter: blur(var(--blur-sm));
  -webkit-backdrop-filter: blur(var(--blur-sm));
  color: var(--text-secondary);
  font-size: 13px;
  font-weight: 500;
  text-decoration: none;
  transition: background 170ms var(--ease), color 170ms var(--ease),
              border-color 170ms var(--ease);
}}
.sp-utility a:hover {{
  background: rgba(255,255,255,0.085);
  border-color: rgba(255,255,255,0.16);
  color: var(--text-primary);
}}
.sp-utility a:focus-visible {{ outline: 2px solid var(--accent-edge); outline-offset: 2px; }}
.sp-utility svg {{ width: 16px; height: 16px; flex: 0 0 16px; }}

/* ===========================================================================
   SECTION HEADER
   =========================================================================== */

.sp-section {{
  display: flex;
  align-items: baseline;
  gap: var(--space-3);
  margin: var(--space-7) 0 var(--space-5) 0;
}}
.sp-section:first-child {{ margin-top: 0; }}
.sp-section-index {{
  font-size: 11px;
  font-weight: 600;
  letter-spacing: 0.12em;
  color: var(--text-faint);
  flex: 0 0 auto;
  padding-top: 4px;
}}
.sp-section-rule {{
  flex: 1 1 auto;
  height: 1px;
  background: linear-gradient(90deg, var(--glass-1-edge), transparent);
  margin-bottom: 6px;
}}

/* ===========================================================================
   GLASS LEVELS
   =========================================================================== */

.sp-glass {{
  position: relative;
  backdrop-filter: blur(var(--blur-md));
  -webkit-backdrop-filter: blur(var(--blur-md));
  transition: background 180ms var(--ease), border-color 180ms var(--ease);
}}

/* LEVEL 01 — major sections, main analytics, upload card, main table */
.sp-glass.l1 {{
  background: var(--glass-1);
  border: 1px solid var(--glass-1-edge);
  border-radius: var(--radius-lg);
  padding: var(--space-6);
  box-shadow: var(--shadow-3);
}}
/* the single hairline highlight that makes the surface read as glass */
.sp-glass.l1::before {{
  content: "";
  position: absolute;
  inset: 0 0 auto 0;
  height: 1px;
  border-radius: var(--radius-lg) var(--radius-lg) 0 0;
  background: linear-gradient(90deg, transparent, rgba(255,255,255,0.14), transparent);
  pointer-events: none;
}}

/* LEVEL 02 — KPI cards, supporting analytics, insights */
.sp-glass.l2 {{
  background: var(--glass-2);
  border: 1px solid var(--glass-2-edge);
  border-radius: var(--radius-md);
  padding: var(--space-5);
  box-shadow: var(--shadow-2);
}}

/* LEVEL 03 — interactive: buttons, chips, filters, tabs */
.sp-glass.l3 {{
  background: var(--glass-3);
  border: 1px solid var(--glass-3-edge);
  border-radius: var(--radius-sm);
  padding: var(--space-2) var(--space-3);
  backdrop-filter: blur(var(--blur-sm));
  -webkit-backdrop-filter: blur(var(--blur-sm));
  box-shadow: var(--shadow-1);
}}

/* Tertiary — supporting detail, minimal treatment */
.sp-glass.l3-flat {{
  background: transparent;
  border: 1px solid var(--glass-2-edge);
  border-radius: var(--radius-md);
  padding: var(--space-4) var(--space-5);
  box-shadow: none;
}}

/* Streamlit containers created with st.container(key="...") pick up glass
   via the generated .st-key-<key> class. */
[class*="st-key-glass1"] {{
  background: var(--glass-1);
  border: 1px solid var(--glass-1-edge);
  border-radius: var(--radius-lg);
  padding: var(--space-6);
  box-shadow: var(--shadow-3);
  backdrop-filter: blur(var(--blur-md));
  -webkit-backdrop-filter: blur(var(--blur-md));
}}
[class*="st-key-glass2"] {{
  background: var(--glass-2);
  border: 1px solid var(--glass-2-edge);
  border-radius: var(--radius-md);
  padding: var(--space-5);
  box-shadow: var(--shadow-2);
  backdrop-filter: blur(var(--blur-md));
  -webkit-backdrop-filter: blur(var(--blur-md));
}}

/* ===========================================================================
   KPI CARDS — label ↓ value ↓ growth ↓ context, always in that order
   =========================================================================== */

.sp-kpi {{ display: flex; flex-direction: column; height: 100%; }}

.sp-kpi-label {{
  font-size: 11px;
  font-weight: 600;
  letter-spacing: 0.11em;
  text-transform: uppercase;
  color: var(--text-muted);
  margin: 0 0 var(--space-3) 0;
}}
.sp-kpi-value {{
  font-size: 32px;
  font-weight: 620;
  line-height: 1.12;
  letter-spacing: -0.022em;
  color: var(--text-primary);
  margin: 0;
  overflow-wrap: anywhere;   /* long ₹ strings wrap rather than clip */
}}
.sp-kpi-growth {{
  display: flex;
  align-items: center;
  gap: var(--space-2);
  margin-top: var(--space-3);
  flex-wrap: wrap;
}}
.sp-kpi-context {{ font-size: 12px; color: var(--text-muted); }}
.sp-kpi-foot {{
  display: flex;
  align-items: baseline;
  justify-content: space-between;
  gap: var(--space-3);
  margin-top: var(--space-4);
  padding-top: var(--space-3);
  border-top: 1px solid var(--glass-2-edge);
}}
.sp-kpi-foot .k {{ font-size: 11px; letter-spacing: 0.09em; text-transform: uppercase;
                   color: var(--text-muted); }}
.sp-kpi-foot .v {{ font-size: 14px; font-weight: 550; color: var(--text-secondary);
                   font-variant-numeric: tabular-nums; }}

/* Primary card — strongest presence, larger value, more padding */
.sp-kpi-card.primary {{
  background: var(--glass-1);
  border: 1px solid var(--glass-1-edge);
  border-radius: var(--radius-lg);
  padding: var(--space-6);
  box-shadow: var(--shadow-3);
}}
.sp-kpi-card.primary .sp-kpi-value {{ font-size: 36px; }}
.sp-kpi-card.primary::after {{
  content: "";
  position: absolute;
  left: var(--space-6); top: 0;
  width: 40px; height: 2px;
  background: var(--accent);
  border-radius: 0 0 2px 2px;
  opacity: 0.85;
}}

/* Secondary card — restrained */
.sp-kpi-card.secondary {{
  background: var(--glass-2);
  border: 1px solid var(--glass-2-edge);
  border-radius: var(--radius-md);
  padding: var(--space-5);
  box-shadow: var(--shadow-2);
}}
.sp-kpi-card.secondary .sp-kpi-value {{ font-size: 26px; }}

/* Supporting card — minimal */
.sp-kpi-card.supporting {{
  background: transparent;
  border: 1px solid var(--glass-2-edge);
  border-radius: var(--radius-md);
  padding: var(--space-4) var(--space-5);
  box-shadow: none;
}}
.sp-kpi-card.supporting .sp-kpi-value {{ font-size: 21px; font-weight: 600; }}
.sp-kpi-card.supporting .sp-kpi-label {{ margin-bottom: var(--space-2); }}

.sp-kpi-card {{
  position: relative;
  height: 100%;
  backdrop-filter: blur(var(--blur-md));
  -webkit-backdrop-filter: blur(var(--blur-md));
}}

@media (max-width: 900px) {{
  .sp-kpi-card.primary .sp-kpi-value {{ font-size: 28px; }}
  .sp-kpi-card.secondary .sp-kpi-value {{ font-size: 22px; }}
}}

/* ===========================================================================
   SEMANTIC PILLS — colour + glyph + label, never colour alone
   =========================================================================== */

.sp-pill {{
  display: inline-flex;
  align-items: center;
  gap: 5px;
  height: 24px;
  padding: 0 9px;
  border-radius: var(--radius-pill);
  font-size: 12.5px;
  font-weight: 600;
  letter-spacing: -0.004em;
  border: 1px solid transparent;
  white-space: nowrap;
}}
.sp-pill .g {{ font-size: 11px; line-height: 1; }}   /* the ↑ ↓ → glyph */

.sp-pill.positive {{ color: var(--success); background: var(--success-dim);
                     border-color: rgba(99,217,154,0.24); }}
.sp-pill.negative {{ color: var(--danger);  background: var(--danger-dim);
                     border-color: rgba(255,107,107,0.24); }}
.sp-pill.warning  {{ color: var(--warning); background: var(--warning-dim);
                     border-color: rgba(224,178,82,0.24); }}
.sp-pill.neutral  {{ color: var(--neutral); background: var(--neutral-dim);
                     border-color: rgba(124,136,153,0.22); }}

/* Status badges */
.sp-badge {{
  display: inline-flex;
  align-items: center;
  gap: var(--space-2);
  height: 26px;
  padding: 0 10px;
  border-radius: var(--radius-pill);
  font-size: 11.5px;
  font-weight: 600;
  letter-spacing: 0.05em;
  text-transform: uppercase;
  border: 1px solid transparent;
}}
.sp-badge::before {{ content: ""; width: 6px; height: 6px; border-radius: 50%;
                     background: currentColor; flex: 0 0 6px; }}
.sp-badge.positive {{ color: var(--success); background: var(--success-dim);
                      border-color: rgba(99,217,154,0.22); }}
.sp-badge.negative {{ color: var(--danger);  background: var(--danger-dim);
                      border-color: rgba(255,107,107,0.22); }}
.sp-badge.warning  {{ color: var(--warning); background: var(--warning-dim);
                      border-color: rgba(224,178,82,0.22); }}
.sp-badge.neutral  {{ color: var(--neutral); background: var(--neutral-dim);
                      border-color: rgba(124,136,153,0.20); }}
.sp-badge.accent   {{ color: var(--accent);  background: var(--accent-dim);
                      border-color: var(--accent-edge); }}

/* ===========================================================================
   SHEET CHIPS
   =========================================================================== */

.sp-chips {{ display: flex; flex-wrap: wrap; gap: var(--space-2); margin: 0; }}
.sp-chip {{
  display: inline-flex;
  align-items: center;
  gap: var(--space-2);
  height: 32px;
  padding: 0 var(--space-3);
  border-radius: var(--radius-sm);
  font-size: 13px;
  font-weight: 500;
  background: var(--glass-3);
  border: 1px solid var(--glass-3-edge);
  color: var(--text-secondary);
  backdrop-filter: blur(var(--blur-sm));
  -webkit-backdrop-filter: blur(var(--blur-sm));
  transition: background 170ms var(--ease), border-color 170ms var(--ease);
}}
.sp-chip:hover {{ background: rgba(255,255,255,0.08); border-color: rgba(255,255,255,0.16); }}

/* FINAL is the primary sheet — stronger emphasis than the rest */
.sp-chip.primary {{
  color: var(--accent);
  background: var(--accent-dim);
  border-color: var(--accent-edge);
  font-weight: 600;
}}
.sp-chip.found::before   {{ content: "✓"; font-size: 11px; color: var(--success); }}
.sp-chip.missing {{ color: var(--text-faint); border-style: dashed; }}
.sp-chip.missing::before {{ content: "—"; font-size: 11px; }}

/* ===========================================================================
   BUTTONS — three levels, never all styled alike
   =========================================================================== */

.stButton > button,
.stDownloadButton > button,
[data-testid="stFormSubmitButton"] > button {{
  font-family: var(--font) !important;
  height: 42px;
  min-height: 42px;
  padding: 0 var(--space-4) !important;
  border-radius: var(--radius-sm) !important;
  font-size: 14px !important;
  font-weight: 550 !important;
  letter-spacing: -0.003em;
  transition: background 170ms var(--ease), border-color 170ms var(--ease),
              color 170ms var(--ease), transform 170ms var(--ease) !important;
}}

/* Secondary (Streamlit default) — medium emphasis */
.stButton > button[kind="secondary"],
.stButton > button:not([kind="primary"]),
[data-testid="stBaseButton-secondary"] {{
  background: var(--glass-3) !important;
  border: 1px solid var(--glass-3-edge) !important;
  color: var(--text-primary) !important;
  backdrop-filter: blur(var(--blur-sm));
  -webkit-backdrop-filter: blur(var(--blur-sm));
  box-shadow: var(--shadow-1) !important;
}}
.stButton > button:not([kind="primary"]):hover,
[data-testid="stBaseButton-secondary"]:hover {{
  background: rgba(255,255,255,0.085) !important;
  border-color: rgba(255,255,255,0.17) !important;
}}

/* Primary CTA — highest priority, the only filled accent button */
.stButton > button[kind="primary"],
[data-testid="stBaseButton-primary"],
.stDownloadButton > button {{
  background: linear-gradient(180deg, {GOLD} 0%, {GOLD_SOFT} 100%) !important;
  border: 1px solid rgba(216,183,106,0.55) !important;
  color: #14110A !important;
  font-weight: 620 !important;
  box-shadow: 0 6px 18px -8px rgba(216,183,106,0.42) !important;
}}
.stButton > button[kind="primary"]:hover,
[data-testid="stBaseButton-primary"]:hover,
.stDownloadButton > button:hover {{
  filter: brightness(1.07);
  transform: translateY(-1px);
}}

/* Tertiary — minimal, for "View details →" style actions.
   Apply with: st.button("View details →", type="tertiary")   [Streamlit ≥1.42]
   or wrap in st.container(key="tertiary-actions"). */
.stButton > button[kind="tertiary"],
[class*="st-key-tertiary"] .stButton > button {{
  background: transparent !important;
  border: 1px solid transparent !important;
  color: var(--text-secondary) !important;
  padding: 0 var(--space-2) !important;
  box-shadow: none !important;
  font-weight: 500 !important;
}}
.stButton > button[kind="tertiary"]:hover,
[class*="st-key-tertiary"] .stButton > button:hover {{
  color: var(--accent) !important;
  background: transparent !important;
}}

/* States */
.stButton > button:focus-visible,
.stDownloadButton > button:focus-visible {{
  outline: 2px solid var(--accent-edge) !important;
  outline-offset: 2px;
}}
.stButton > button:active {{ transform: translateY(0) scale(0.995); }}
.stButton > button:disabled {{ opacity: 0.42 !important; transform: none !important;
                               filter: none !important; cursor: not-allowed; }}

/* ===========================================================================
   FILE UPLOADER — one component, one action, no duplicated text
   Pass label_visibility="collapsed" so the surrounding card owns the copy.
   =========================================================================== */

[data-testid="stFileUploader"] {{ width: 100%; }}
[data-testid="stFileUploader"] > label {{ display: none !important; }}

[data-testid="stFileUploaderDropzone"] {{
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: var(--space-3);
  min-height: 224px;
  padding: var(--space-7) var(--space-5) !important;
  background: var(--glass-2) !important;
  border: 1px dashed rgba(255,255,255,0.16) !important;
  border-radius: var(--radius-md) !important;
  backdrop-filter: blur(var(--blur-md));
  -webkit-backdrop-filter: blur(var(--blur-md));
  transition: background 180ms var(--ease), border-color 180ms var(--ease);
}}
[data-testid="stFileUploaderDropzone"]:hover {{
  background: rgba(255,255,255,0.05) !important;
  border-color: var(--accent-edge) !important;
}}

/* Suppress Streamlit's native instruction text and icon, then supply our own
   copy once. This is the fix for the duplicated upload label. */
[data-testid="stFileUploaderDropzoneInstructions"] > * {{ display: none !important; }}
[data-testid="stFileUploaderDropzoneInstructions"] {{
  display: flex !important;
  flex-direction: column;
  align-items: center;
  gap: var(--space-2);
  color: var(--text-secondary);
}}
[data-testid="stFileUploaderDropzoneInstructions"]::before {{
  /* upload glyph, same stroke family as the rest of the icon set */
  content: "";
  width: 40px; height: 40px;
  margin-bottom: var(--space-2);
  background-color: {INK_SOFT};
  -webkit-mask: var(--sp-upload-icon) center / 40px 40px no-repeat;
  mask: var(--sp-upload-icon) center / 40px 40px no-repeat;
  opacity: 0.85;
}}
[data-testid="stFileUploaderDropzoneInstructions"]::after {{
  content: "Drop your RM scorecard here";
  font-size: 15px;
  font-weight: 550;
  color: var(--text-primary);
  letter-spacing: -0.006em;
}}

/* The single upload action, relabelled */
[data-testid="stFileUploaderDropzone"] button {{
  height: 42px !important;
  padding: 0 var(--space-5) !important;
  border-radius: var(--radius-sm) !important;
  background: var(--glass-3) !important;
  border: 1px solid rgba(255,255,255,0.16) !important;
  color: var(--text-primary) !important;
  font-size: 14px !important;
  font-weight: 550 !important;
  overflow: hidden;
  position: relative;
  text-indent: -9999px;   /* hide "Browse files" */
  white-space: nowrap;
}}
[data-testid="stFileUploaderDropzone"] button::after {{
  content: "Choose Excel file";
  position: absolute;
  inset: 0;
  display: flex;
  align-items: center;
  justify-content: center;
  text-indent: 0;
}}
[data-testid="stFileUploaderDropzone"] button:hover {{
  background: rgba(255,255,255,0.10) !important;
  border-color: var(--accent-edge) !important;
}}
[data-testid="stFileUploaderDropzone"] small {{
  font-size: 11.5px !important;
  color: var(--text-muted) !important;
  letter-spacing: 0.03em;
}}

/* Uploaded-file row */
[data-testid="stFileUploaderFile"] {{
  background: var(--glass-2);
  border: 1px solid var(--glass-2-edge);
  border-radius: var(--radius-sm);
  padding: var(--space-3) var(--space-4);
  margin-top: var(--space-3);
}}
[data-testid="stFileUploaderFileName"] {{
  font-size: 14px; font-weight: 550; color: var(--text-primary);
}}

/* ===========================================================================
   INPUTS · FILTERS — grouped, consistent, interactive glass
   =========================================================================== */

.sp-filters {{
  display: flex;
  flex-wrap: wrap;
  align-items: flex-end;
  gap: var(--space-4);
  padding: var(--space-4) var(--space-5);
  background: var(--glass-2);
  border: 1px solid var(--glass-2-edge);
  border-radius: var(--radius-md);
  backdrop-filter: blur(var(--blur-md));
  -webkit-backdrop-filter: blur(var(--blur-md));
  margin-bottom: var(--space-5);
}}

[data-baseweb="select"] > div,
[data-testid="stTextInput"] input,
[data-testid="stNumberInput"] input,
[data-testid="stDateInput"] input {{
  background: var(--glass-3) !important;
  border: 1px solid var(--glass-3-edge) !important;
  border-radius: var(--radius-sm) !important;
  color: var(--text-primary) !important;
  min-height: 42px;
  font-family: var(--font) !important;
  font-size: 14px !important;
  transition: border-color 170ms var(--ease), background 170ms var(--ease);
}}
[data-baseweb="select"] > div:hover {{ border-color: rgba(255,255,255,0.17) !important; }}
[data-baseweb="select"] > div:focus-within,
[data-testid="stTextInput"] input:focus,
[data-testid="stNumberInput"] input:focus {{
  border-color: var(--accent-edge) !important;
  box-shadow: 0 0 0 3px rgba(216,183,106,0.10) !important;
}}
[data-baseweb="popover"] [role="listbox"] {{
  background: rgba(15,17,21,0.96) !important;
  border: 1px solid var(--glass-1-edge) !important;
  border-radius: var(--radius-sm) !important;
  backdrop-filter: blur(var(--blur-lg));
  -webkit-backdrop-filter: blur(var(--blur-lg));
  box-shadow: var(--shadow-3);
}}
[role="option"]:hover {{ background: rgba(255,255,255,0.06) !important; }}
[aria-selected="true"][role="option"] {{ color: var(--accent) !important; }}

/* Widget labels use the metadata scale */
[data-testid="stWidgetLabel"] label,
.stSlider label, .stSelectbox label, .stRadio label > div:first-child {{
  font-size: 11px !important;
  font-weight: 600 !important;
  letter-spacing: 0.1em !important;
  text-transform: uppercase;
  color: var(--text-muted) !important;
}}

/* Sliders — scenario assumption controls */
[data-testid="stSlider"] [data-baseweb="slider"] div[role="slider"] {{
  background: {GOLD} !important;
  border: 2px solid rgba(20,17,10,0.9) !important;
  box-shadow: 0 2px 8px -2px rgba(0,0,0,0.5) !important;
}}
[data-testid="stSlider"] [data-testid="stSliderTickBarMin"],
[data-testid="stSlider"] [data-testid="stSliderTickBarMax"] {{
  font-size: 11px; color: var(--text-faint);
  font-variant-numeric: tabular-nums;
}}
[data-testid="stSlider"] [data-testid="stThumbValue"] {{
  color: var(--accent) !important;
  font-size: 12.5px !important;
  font-weight: 600 !important;
  font-variant-numeric: tabular-nums;
}}

/* Checkbox / radio / toggle accents */
[data-testid="stCheckbox"] svg,
[data-baseweb="checkbox"] svg {{ color: var(--accent); }}

/* ===========================================================================
   TABS — interactive glass, clear active state
   =========================================================================== */

.stTabs [data-baseweb="tab-list"] {{
  gap: var(--space-1);
  padding: var(--space-1);
  background: var(--glass-2);
  border: 1px solid var(--glass-2-edge);
  border-radius: var(--radius-sm);
  backdrop-filter: blur(var(--blur-sm));
  -webkit-backdrop-filter: blur(var(--blur-sm));
  overflow-x: auto;
  scrollbar-width: none;
}}
.stTabs [data-baseweb="tab-list"]::-webkit-scrollbar {{ display: none; }}
.stTabs [data-baseweb="tab"] {{
  height: 38px;
  padding: 0 var(--space-4);
  border-radius: 8px;
  background: transparent;
  color: var(--text-muted);
  font-size: 13.5px;
  font-weight: 550;
  letter-spacing: -0.003em;
  white-space: nowrap;
  transition: background 170ms var(--ease), color 170ms var(--ease);
}}
.stTabs [data-baseweb="tab"]:hover {{ color: var(--text-secondary);
                                      background: rgba(255,255,255,0.04); }}
.stTabs [aria-selected="true"] {{
  background: rgba(255,255,255,0.075) !important;
  color: var(--text-primary) !important;
  box-shadow: inset 0 0 0 1px rgba(255,255,255,0.08);
}}
.stTabs [data-baseweb="tab-highlight"],
.stTabs [data-baseweb="tab-border"] {{ display: none; }}
.stTabs [data-baseweb="tab-panel"] {{ padding-top: var(--space-6); }}

/* ===========================================================================
   TABLES — integrated into the glass, strong numeric alignment
   =========================================================================== */

.sp-table-wrap {{
  background: var(--glass-1);
  border: 1px solid var(--glass-1-edge);
  border-radius: var(--radius-lg);
  padding: var(--space-2);
  box-shadow: var(--shadow-3);
  backdrop-filter: blur(var(--blur-md));
  -webkit-backdrop-filter: blur(var(--blur-md));
  overflow: auto;
  max-height: 620px;
}}
.sp-glass-table {{
  width: 100%;
  border-collapse: separate;
  border-spacing: 0;
  font-size: 13.5px;
}}
.sp-glass-table thead th {{
  position: sticky;
  top: 0;
  z-index: 2;
  background: rgba(15,17,21,0.94);
  backdrop-filter: blur(var(--blur-md));
  -webkit-backdrop-filter: blur(var(--blur-md));
  padding: var(--space-3) var(--space-4);
  text-align: left;
  font-size: 11px;
  font-weight: 600;
  letter-spacing: 0.09em;
  text-transform: uppercase;
  color: var(--text-muted);
  border-bottom: 1px solid var(--glass-1-edge);
  white-space: nowrap;
}}
.sp-glass-table thead th.num {{ text-align: right; }}
.sp-glass-table tbody td {{
  padding: var(--space-3) var(--space-4);
  border-bottom: 1px solid rgba(255,255,255,0.035);
  color: var(--text-secondary);
  white-space: nowrap;
}}
.sp-glass-table tbody td.num {{
  text-align: right;
  color: var(--text-primary);
  font-weight: 500;
}}
.sp-glass-table tbody td:first-child {{
  color: var(--text-primary);
  font-weight: 550;
  position: sticky;
  left: 0;
  background: rgba(11,13,16,0.90);
  backdrop-filter: blur(var(--blur-sm));
  -webkit-backdrop-filter: blur(var(--blur-sm));
}}
.sp-glass-table tbody tr:hover td {{ background: rgba(255,255,255,0.032); }}
.sp-glass-table tbody tr:hover td:first-child {{ background: rgba(20,22,27,0.94); }}
.sp-glass-table tbody tr:last-child td {{ border-bottom: none; }}
.sp-glass-table td.pos {{ color: var(--success); }}
.sp-glass-table td.neg {{ color: var(--danger); }}
.sp-glass-table tbody tr.total td {{
  border-top: 1px solid var(--glass-1-edge);
  font-weight: 620;
  color: var(--text-primary);
}}

/* st.dataframe, when the native widget is used instead */
[data-testid="stDataFrame"], [data-testid="stTable"] {{
  border: 1px solid var(--glass-1-edge) !important;
  border-radius: var(--radius-md) !important;
  overflow: hidden;
  background: var(--glass-2) !important;
}}
[data-testid="stDataFrame"] [role="columnheader"] {{
  background: rgba(15,17,21,0.9) !important;
  color: var(--text-muted) !important;
  font-size: 11px !important;
  letter-spacing: 0.08em;
  text-transform: uppercase;
}}

/* ---- FINAL sheet passthrough ----------------------------------------------
   These class names are emitted by build_final_sheet_html() in the logic
   layer. Styling them here keeps the workbook mirror inside the design system
   without touching that function. */
.final-sheet-scroll {{
  overflow: auto;
  max-height: 700px;
  background: var(--glass-1);
  border: 1px solid var(--glass-1-edge);
  border-radius: var(--radius-lg);
  padding: var(--space-2);
  box-shadow: var(--shadow-3);
  backdrop-filter: blur(var(--blur-md));
  -webkit-backdrop-filter: blur(var(--blur-md));
}}
.final-sheet-table {{
  border-collapse: separate;
  border-spacing: 0;
  font-size: 12.5px;
  font-variant-numeric: tabular-nums;
  color: var(--text-secondary);
}}
.final-sheet-table td {{
  padding: 7px var(--space-3);
  border-bottom: 1px solid rgba(255,255,255,0.035);
  border-right: 1px solid rgba(255,255,255,0.022);
  white-space: nowrap;
}}
.final-sheet-table td:last-child {{ border-right: none; }}
.final-sheet-table tr:hover td {{ background: rgba(255,255,255,0.028); }}
.final-spacer {{ height: 10px; border: none !important; background: transparent !important; }}
.glass-note {{
  padding: var(--space-5);
  background: var(--glass-2);
  border: 1px solid var(--glass-2-edge);
  border-radius: var(--radius-md);
  color: var(--text-muted);
  font-size: 13.5px;
}}

/* ===========================================================================
   CHART CONTAINERS — hierarchy by size, consistent chrome
   =========================================================================== */

.sp-chart {{
  background: var(--glass-2);
  border: 1px solid var(--glass-2-edge);
  border-radius: var(--radius-md);
  padding: var(--space-5);
  box-shadow: var(--shadow-2);
  backdrop-filter: blur(var(--blur-md));
  -webkit-backdrop-filter: blur(var(--blur-md));
}}
.sp-chart.primary {{
  background: var(--glass-1);
  border-color: var(--glass-1-edge);
  border-radius: var(--radius-lg);
  padding: var(--space-6);
  box-shadow: var(--shadow-3);
}}
.sp-chart-head {{
  display: flex;
  align-items: baseline;
  justify-content: space-between;
  gap: var(--space-4);
  margin-bottom: var(--space-5);
  flex-wrap: wrap;
}}
.sp-chart .js-plotly-plot,
.sp-chart-body .js-plotly-plot {{ background: transparent !important; }}

/* ===========================================================================
   MESSAGES · EMPTY · LOADING
   =========================================================================== */

.sp-notice {{
  display: flex;
  gap: var(--space-4);
  padding: var(--space-5);
  border-radius: var(--radius-md);
  border: 1px solid var(--glass-2-edge);
  background: var(--glass-2);
  backdrop-filter: blur(var(--blur-md));
  -webkit-backdrop-filter: blur(var(--blur-md));
}}
.sp-notice .rail {{ width: 3px; border-radius: 3px; flex: 0 0 3px; }}
.sp-notice.error   {{ border-color: rgba(255,107,107,0.26); background: var(--danger-dim); }}
.sp-notice.error   .rail {{ background: var(--danger); }}
.sp-notice.success {{ border-color: rgba(99,217,154,0.24); background: var(--success-dim); }}
.sp-notice.success .rail {{ background: var(--success); }}
.sp-notice.warning {{ border-color: rgba(224,178,82,0.24); background: var(--warning-dim); }}
.sp-notice.warning .rail {{ background: var(--warning); }}
.sp-notice h4 {{ margin: 0 0 var(--space-2) 0; font-size: 15px; font-weight: 600;
                 color: var(--text-primary); }}
.sp-notice p  {{ margin: 0; font-size: 13.5px; line-height: 1.6; color: var(--text-secondary); }}
.sp-notice-actions {{ margin-top: var(--space-4); display: flex; gap: var(--space-3);
                      flex-wrap: wrap; }}

/* An empty dashboard should look intentional, not broken */
.sp-empty {{
  display: flex;
  flex-direction: column;
  align-items: center;
  text-align: center;
  gap: var(--space-3);
  padding: var(--space-8) var(--space-6);
  border: 1px dashed rgba(255,255,255,0.10);
  border-radius: var(--radius-lg);
  background: var(--glass-2);
}}
.sp-empty h4 {{ margin: 0; font-size: 17px; font-weight: 600; color: var(--text-primary); }}
.sp-empty p  {{ margin: 0; font-size: 13.5px; color: var(--text-muted); max-width: 46ch; }}

/* Loading feedback */
[data-testid="stSpinner"] > div {{ border-top-color: var(--accent) !important; }}
[data-testid="stSpinner"] p {{ color: var(--text-secondary) !important; font-size: 13.5px; }}
.stProgress > div > div > div > div {{
  background: linear-gradient(90deg, {GOLD_SOFT}, {GOLD}) !important;
}}

/* Streamlit's native alerts, brought into the system */
[data-testid="stAlert"] {{
  border-radius: var(--radius-md) !important;
  border: 1px solid var(--glass-2-edge) !important;
  background: var(--glass-2) !important;
  backdrop-filter: blur(var(--blur-md));
  -webkit-backdrop-filter: blur(var(--blur-md));
  font-size: 13.5px;
}}

/* Expanders */
[data-testid="stExpander"] {{
  border: 1px solid var(--glass-2-edge) !important;
  border-radius: var(--radius-md) !important;
  background: var(--glass-2) !important;
  backdrop-filter: blur(var(--blur-md));
  -webkit-backdrop-filter: blur(var(--blur-md));
}}
[data-testid="stExpander"] summary {{ font-size: 14px; font-weight: 550;
                                      color: var(--text-primary); }}

/* Dividers and metrics */
hr, [data-testid="stDivider"] hr {{
  border: none;
  height: 1px;
  background: linear-gradient(90deg, transparent, var(--glass-1-edge) 12%,
              var(--glass-1-edge) 88%, transparent);
  margin: var(--space-6) 0 !important;
}}
[data-testid="stMetricValue"] {{
  font-size: 28px !important;
  font-weight: 620 !important;
  color: var(--text-primary) !important;
  font-variant-numeric: tabular-nums;
}}
[data-testid="stMetricLabel"] {{
  font-size: 11px !important;
  letter-spacing: 0.1em;
  text-transform: uppercase;
  color: var(--text-muted) !important;
}}

/* Grid gaps — the horizontal rhythm between columns */
[data-testid="stHorizontalBlock"] {{ gap: var(--grid-gap); }}
@media (max-width: 1100px) {{
  [data-testid="stHorizontalBlock"] {{ gap: var(--space-4); }}
}}

/* Scrollbars */
::-webkit-scrollbar {{ width: 10px; height: 10px; }}
::-webkit-scrollbar-track {{ background: transparent; }}
::-webkit-scrollbar-thumb {{
  background: rgba(255,255,255,0.10);
  border-radius: var(--radius-pill);
  border: 2px solid transparent;
  background-clip: content-box;
}}
::-webkit-scrollbar-thumb:hover {{ background: rgba(255,255,255,0.18);
                                   background-clip: content-box; }}

/* Accessibility floor */
:focus-visible {{ outline: 2px solid var(--accent-edge); outline-offset: 2px; }}
@media (prefers-reduced-motion: reduce) {{
  *, *::before, *::after {{
    animation-duration: 0.01ms !important;
    transition-duration: 0.01ms !important;
  }}
}}

/* Icon masks — one stroke family, 1.6px weight, 24px grid */
:root {{
  --sp-upload-icon: url("data:image/svg+xml;utf8,\
<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' \
stroke='%23fff' stroke-width='1.6' stroke-linecap='round' stroke-linejoin='round'>\
<path d='M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4'/>\
<polyline points='17 8 12 3 7 8'/><line x1='12' y1='3' x2='12' y2='15'/></svg>");
}}
</style>
"""


def inject() -> None:
    """Inject the stylesheet. Call once, right after st.set_page_config."""
    st.markdown(_CSS, unsafe_allow_html=True)


# =============================================================================
# 3. ICONS
# =============================================================================

#: The real Google brand mark — four colours, correct geometry.
#: Use this wherever Google is represented. Never an AI sparkle or chat glyph.
GOOGLE_MARK = (
    "<svg viewBox='0 0 48 48' xmlns='http://www.w3.org/2000/svg' aria-hidden='true'>"
    "<path fill='#4285F4' d='M45.12 24.5c0-1.56-.14-3.06-.4-4.5H24v8.51h11.84"
    "c-.51 2.75-2.06 5.08-4.39 6.64v5.52h7.11c4.16-3.83 6.56-9.47 6.56-16.17z'/>"
    "<path fill='#34A853' d='M24 46c5.94 0 10.92-1.97 14.56-5.33l-7.11-5.52"
    "c-1.97 1.32-4.49 2.1-7.45 2.1-5.73 0-10.58-3.87-12.31-9.07H4.34v5.7"
    "C7.96 41.07 15.4 46 24 46z'/>"
    "<path fill='#FBBC05' d='M11.69 28.18C11.25 26.86 11 25.45 11 24"
    "s.25-2.86.69-4.18v-5.7H4.34C2.85 17.09 2 20.45 2 24c0 3.55.85 6.91 2.34 9.88"
    "l7.35-5.7z'/>"
    "<path fill='#EA4335' d='M24 10.75c3.23 0 6.13 1.11 8.41 3.29l6.31-6.31"
    "C34.91 4.18 29.93 2 24 2 15.4 2 7.96 6.93 4.34 14.12l7.35 5.7"
    "c1.73-5.2 6.58-9.07 12.31-9.07z'/></svg>"
)

_GITHUB_MARK = (
    "<svg viewBox='0 0 24 24' fill='currentColor' aria-hidden='true'>"
    "<path d='M12 2C6.48 2 2 6.58 2 12.25c0 4.53 2.87 8.37 6.84 9.73.5.1.68-.22.68-.48"
    "v-1.7c-2.78.62-3.37-1.37-3.37-1.37-.45-1.18-1.11-1.5-1.11-1.5-.91-.63.07-.62.07-.62"
    "1 .07 1.53 1.06 1.53 1.06.89 1.57 2.34 1.12 2.91.85.09-.66.35-1.12.63-1.38"
    "-2.22-.26-4.56-1.14-4.56-5.07 0-1.12.39-2.03 1.03-2.75-.1-.26-.45-1.3.1-2.71"
    "0 0 .84-.28 2.75 1.05a9.4 9.4 0 0 1 5.01 0c1.91-1.33 2.75-1.05 2.75-1.05"
    ".55 1.41.2 2.45.1 2.71.64.72 1.03 1.63 1.03 2.75 0 3.94-2.34 4.81-4.57 5.06"
    ".36.32.68.94.68 1.9v2.82c0 .27.18.59.69.48A10.11 10.11 0 0 0 22 12.25"
    "C22 6.58 17.52 2 12 2z'/></svg>"
)

_ARROWS = {"positive": "↑", "negative": "↓", "neutral": "→", "warning": "!"}


def _e(value: Any) -> str:
    """Escape for HTML. Values arrive pre-formatted; they are never re-derived."""
    return escape("" if value is None else str(value), quote=True)


# =============================================================================
# 4. COMPONENTS
# =============================================================================

def page_header(
    eyebrow: str,
    title: str,
    subtitle: str = "",
    utility_links: Optional[Sequence[Dict[str, str]]] = None,
) -> None:
    """
    Page header. Flat type, no glass container — the contrast between flat
    typography and glass surfaces is what creates the hierarchy.

    utility_links: [{"label": "Fork", "url": "...", "icon": "github"}]
    """
    links_html = ""
    if utility_links:
        parts = []
        for link in utility_links:
            icon = link.get("icon", "")
            glyph = _GITHUB_MARK if icon == "github" else (
                GOOGLE_MARK if icon == "google" else ""
            )
            parts.append(
                f"<a href='{_e(link.get('url', '#'))}' target='_blank' rel='noopener'>"
                f"{glyph}<span>{_e(link.get('label', ''))}</span></a>"
            )
        links_html = f"<div class='sp-utility'>{''.join(parts)}</div>"

    subtitle_html = f"<p class='sp-subtitle'>{_e(subtitle)}</p>" if subtitle else ""

    st.markdown(
        f"""<div class='sp-header'><div class='sp-header-row'>
        <div class='sp-header-text'>
          <p class='sp-eyebrow'>{_e(eyebrow)}</p>
          <h1 class='sp-title'>{_e(title)}</h1>
          {subtitle_html}
        </div>{links_html}</div></div>""",
        unsafe_allow_html=True,
    )


def section_header(title: str, description: str = "", index: str = "") -> None:
    """
    Section heading with an optional index.

    Only pass `index` when the sections genuinely form a sequence (the upload
    flow does; a set of parallel analytics tabs does not).
    """
    index_html = f"<span class='sp-section-index'>{_e(index)}</span>" if index else ""
    st.markdown(
        f"""<div class='sp-section'>{index_html}
        <div style='flex:0 1 auto;min-width:0;'>
          <h2 class='sp-section-title'>{_e(title)}</h2>
          {f"<p class='sp-section-desc'>{_e(description)}</p>" if description else ""}
        </div>
        <div class='sp-section-rule'></div></div>""",
        unsafe_allow_html=True,
    )


@contextmanager
def glass(level: int = 1, key: Optional[str] = None):
    """Container with a glass surface that can hold real Streamlit widgets.

    The helper adapts to the Streamlit version installed by the deployment.
    Newer versions receive a stable key so the CSS can target the container;
    older versions fall back to a normal container without passing unsupported
    keyword arguments.
    """
    level = max(1, min(int(level), 2))
    suffix = f"-{key}" if key else ""
    container_key = f"glass{level}{suffix}"

    kwargs: Dict[str, Any] = {}
    if _supports_parameter(st.container, "key"):
        kwargs["key"] = container_key
    elif _supports_parameter(st.container, "border"):
        # Older builds may support border but not key.
        kwargs["border"] = True

    with st.container(**kwargs):
        yield


def delta_pill(text: str, state: str = "neutral", show_glyph: bool = True) -> str:
    """
    Semantic delta pill as an HTML string.

    `text` must be the app-formatted value (e.g. fmt_pct_signed(...) output);
    it is rendered exactly as given. Colour is never the only signal — a
    directional glyph is included.
    """
    state = state if state in _ARROWS else "neutral"
    glyph = f"<span class='g'>{_ARROWS[state]}</span>" if show_glyph else ""
    return f"<span class='sp-pill {state}'>{glyph}{_e(text)}</span>"


def status_badge(text: str, state: str = "neutral") -> str:
    """Status badge as an HTML string. Valid states: positive, negative,
    warning, neutral, accent."""
    allowed = {"positive", "negative", "warning", "neutral", "accent"}
    state = state if state in allowed else "neutral"
    return f"<span class='sp-badge {state}'>{_e(text)}</span>"


def kpi_card(
    label: str,
    value: str,
    delta: Optional[str] = None,
    delta_state: str = "neutral",
    context: Optional[str] = None,
    footnote_label: Optional[str] = None,
    footnote_value: Optional[str] = None,
    emphasis: str = "secondary",
    render: bool = True,
) -> str:
    """
    KPI card with a fixed internal order: label ↓ value ↓ growth ↓ context.

    `value`, `delta` and `footnote_value` must already be formatted by the
    application's own formatters. This function renders them verbatim.

    emphasis: "primary" | "secondary" | "supporting"
    """
    emphasis = emphasis if emphasis in {"primary", "secondary", "supporting"} else "secondary"

    growth = ""
    if delta or context:
        pill = delta_pill(delta, delta_state) if delta else ""
        ctx = f"<span class='sp-kpi-context'>{_e(context)}</span>" if context else ""
        growth = f"<div class='sp-kpi-growth'>{pill}{ctx}</div>"

    foot = ""
    if footnote_label or footnote_value:
        foot = (
            f"<div class='sp-kpi-foot'>"
            f"<span class='k'>{_e(footnote_label or '')}</span>"
            f"<span class='v'>{_e(footnote_value or '')}</span></div>"
        )

    html = (
        f"<div class='sp-kpi-card {emphasis}'><div class='sp-kpi'>"
        f"<p class='sp-kpi-label'>{_e(label)}</p>"
        f"<p class='sp-kpi-value'>{_e(value)}</p>"
        f"{growth}{foot}</div></div>"
    )
    if render:
        st.markdown(html, unsafe_allow_html=True)
    return html


def kpi_row(cards: Sequence[Dict[str, Any]], widths: Optional[Sequence[float]] = None) -> None:
    """
    Lay KPI cards on the 12-column grid.

    Pass `widths` to give the primary card more space than its neighbours —
    equal widths are what make every card look equally important.

        theme.kpi_row([primary, sec, sec], widths=[2, 1, 1])
    """
    if not cards:
        return
    ratios = list(widths) if widths else [1] * len(cards)
    if _supports_parameter(st.columns, "gap"):
        columns = st.columns(ratios, gap="medium")
    else:
        columns = st.columns(ratios)
    for column, card in zip(columns, cards):
        with column:
            kpi_card(**card)


def utility_bar(items: Sequence[Dict[str, str]]) -> None:
    """Standalone compact utility actions (icon + label), secondary by design."""
    parts = []
    for item in items:
        icon = item.get("icon", "")
        glyph = _GITHUB_MARK if icon == "github" else (GOOGLE_MARK if icon == "google" else "")
        parts.append(
            f"<a href='{_e(item.get('url', '#'))}' target='_blank' rel='noopener'>"
            f"{glyph}<span>{_e(item.get('label', ''))}</span></a>"
        )
    st.markdown(f"<div class='sp-utility'>{''.join(parts)}</div>", unsafe_allow_html=True)


def sheet_chips(
    sheets: Sequence[str],
    primary: Optional[str] = None,
    found: Optional[Iterable[str]] = None,
) -> None:
    """
    Render workbook sheet names as chips.

    Sheet names are printed exactly as supplied — pass the app's own
    SHEET_ALIASES / FINAL_SHEET_ALIASES values, never a retyped copy.

    primary: the sheet that carries the most weight (e.g. "FINAL").
    found:   sheets detected in the uploaded workbook, from real app state.
    """
    found_set = set(found) if found is not None else None
    parts = []
    for sheet in sheets:
        classes = ["sp-chip"]
        if primary is not None and sheet == primary:
            classes.append("primary")
        if found_set is not None:
            classes.append("found" if sheet in found_set else "missing")
        parts.append(f"<span class='{' '.join(classes)}'>{_e(sheet)}</span>")
    st.markdown(f"<div class='sp-chips'>{''.join(parts)}</div>", unsafe_allow_html=True)


def upload_intro(
    step: str = "STEP 01 · GET STARTED",
    title: str = "Upload the RM scorecard",
    description: str = (
        "Upload the Excel workbook required to initialize the "
        "Sales Performance Command Center."
    ),
) -> None:
    """
    Copy block that sits above st.file_uploader inside the same glass card,
    so the upload region reads as one coherent component.

        with theme.glass(level=1, key="upload"):
            theme.upload_intro()
            file = st.file_uploader("RM scorecard", type=["xlsx", "xlsm"],
                                    label_visibility="collapsed")
    """
    st.markdown(
        f"""<p class='sp-eyebrow'>{_e(step)}</p>
        <h2 class='sp-section-title' style='margin-bottom:var(--space-3);'>{_e(title)}</h2>
        <p class='sp-body' style='margin:0 0 var(--space-5) 0;max-width:56ch;'>
          {_e(description)}</p>""",
        unsafe_allow_html=True,
    )


def upload_success(filename: str, detail: str, render: bool = True) -> str:
    """
    Success state after a workbook is accepted.

    `filename` and `detail` must come from real application state — the
    uploaded file's own name and the actual sheet/validation result. Do not
    pass a placeholder.
    """
    html = (
        f"<div class='sp-notice success'><div class='rail'></div><div>"
        f"<h4>Workbook ready</h4><p>{_e(filename)} · {_e(detail)}</p>"
        f"</div></div>"
    )
    if render:
        st.markdown(html, unsafe_allow_html=True)
    return html


def upload_error(
    heading: str = "Workbook structure incomplete",
    detail: str = "",
    render: bool = True,
) -> str:
    """
    Error state. Distinct but not alarming: red accent, clear heading, short
    explanation, and a corrective action rendered by the caller as a button.

    `detail` should be the user-facing part of the existing WorkbookError
    message. The validation logic itself is unchanged.
    """
    html = (
        f"<div class='sp-notice error'><div class='rail'></div><div>"
        f"<h4>{_e(heading)}</h4><p>{_e(detail)}</p></div></div>"
    )
    if render:
        st.markdown(html, unsafe_allow_html=True)
    return html


def empty_state(title: str, description: str, render: bool = True) -> str:
    """An empty region that looks deliberately empty rather than broken."""
    html = (
        f"<div class='sp-empty'><h4>{_e(title)}</h4>"
        f"<p>{_e(description)}</p></div>"
    )
    if render:
        st.markdown(html, unsafe_allow_html=True)
    return html


def glass_table(
    display_frame,
    formats: Optional[Dict[str, str]] = None,
    numeric_formats: Optional[Iterable[str]] = None,
    signed_formats: Optional[Iterable[str]] = None,
    total_rows: Optional[Iterable[Any]] = None,
    render: bool = True,
) -> str:
    """
    Render a **pre-formatted** frame as a glass table.

    Pass the output of the app's own `format_table(frame, formats)` — every
    cell is already a display string and is emitted verbatim. Supply the
    original `formats` dict plus the app's `NUMERIC_FORMATS` / `SIGNED_FORMATS`
    sets so columns can be right-aligned and signed values tinted.

        display = format_table(frame, formats)
        theme.glass_table(display, formats,
                          numeric_formats=NUMERIC_FORMATS,
                          signed_formats=SIGNED_FORMATS)

    total_rows: values in the first column that should read as a total row
                (e.g. ["Total", "Overall"]).
    """
    formats = formats or {}
    numeric = set(numeric_formats or ())
    signed = set(signed_formats or ())
    totals = {str(t) for t in (total_rows or ())}

    columns = list(display_frame.columns)
    numeric_columns = {c for c in columns if formats.get(c) in numeric}
    signed_columns = {c for c in columns if formats.get(c) in signed}

    head = "".join(
        f"<th class='{'num' if c in numeric_columns else ''}'>{_e(c)}</th>"
        for c in columns
    )

    body_rows: List[str] = []
    for _, row in display_frame.iterrows():
        first = str(row[columns[0]]) if columns else ""
        row_class = " class='total'" if first in totals else ""
        cells: List[str] = []
        for column in columns:
            text = "" if row[column] is None else str(row[column])
            classes: List[str] = []
            if column in numeric_columns:
                classes.append("num")
            # Tint is read off the rendered sign only. No value is recomputed.
            if column in signed_columns:
                stripped = text.lstrip("₹ \u20b9")
                if stripped.startswith("+"):
                    classes.append("pos")
                elif stripped.startswith("-"):
                    classes.append("neg")
            attr = f" class='{' '.join(classes)}'" if classes else ""
            cells.append(f"<td{attr}>{_e(text)}</td>")
        body_rows.append(f"<tr{row_class}>{''.join(cells)}</tr>")

    html = (
        "<div class='sp-table-wrap'><table class='sp-glass-table'>"
        f"<thead><tr>{head}</tr></thead>"
        f"<tbody>{''.join(body_rows)}</tbody></table></div>"
    )
    if render:
        st.markdown(html, unsafe_allow_html=True)
    return html


@contextmanager
def chart_frame(
    title: str,
    subtitle: str = "",
    badge: Optional[str] = None,
    emphasis: str = "secondary",
):
    """
    Chart container. Use emphasis="primary" for the one chart that matters
    most on the page and "secondary"/"supporting" for the rest — equal-sized
    charts are what flatten the editorial hierarchy.

        with theme.chart_frame("Momentum trajectory", emphasis="primary"):
            st.plotly_chart(fig, use_container_width=True)
    """
    css_class = "sp-chart primary" if emphasis == "primary" else "sp-chart"
    badge_html = badge or ""
    st.markdown(
        f"""<div class='{css_class}'><div class='sp-chart-head'>
        <div><h3 class='sp-card-title'>{_e(title)}</h3>
        {f"<p class='sp-meta' style='margin-top:var(--space-2);'>{_e(subtitle)}</p>"
          if subtitle else ""}</div>{badge_html}</div>
        <div class='sp-chart-body'>""",
        unsafe_allow_html=True,
    )
    try:
        yield
    finally:
        st.markdown("</div></div>", unsafe_allow_html=True)


# =============================================================================
# 5. PLOTLY TEMPLATE
# =============================================================================

#: Plotly layout template matching the design system. Chart *data* is never
#: touched — this controls chrome, type and colour only.
PLOTLY_TEMPLATE: Dict[str, Any] = {
    "layout": {
        "colorway": CHART_SEQUENCE,
        "paper_bgcolor": "rgba(0,0,0,0)",
        "plot_bgcolor": "rgba(0,0,0,0)",
        "font": {"family": "Inter, -apple-system, 'Segoe UI', sans-serif",
                 "size": 12, "color": INK_SOFT},
        "title": {"font": {"size": 16, "color": INK}, "x": 0, "xanchor": "left"},
        "margin": {"l": 8, "r": 8, "t": 8, "b": 8},
        "xaxis": {
            "gridcolor": GRID_LINE, "zeroline": False,
            "linecolor": AXIS_LINE, "tickcolor": AXIS_LINE,
            "tickfont": {"size": 11.5, "color": INK_MUTED},
            "title": {"font": {"size": 11, "color": INK_MUTED}},
            "automargin": True,
        },
        "yaxis": {
            "gridcolor": GRID_LINE, "zeroline": False,
            "linecolor": "rgba(0,0,0,0)", "tickcolor": "rgba(0,0,0,0)",
            "tickfont": {"size": 11.5, "color": INK_MUTED},
            "title": {"font": {"size": 11, "color": INK_MUTED}},
            "automargin": True,
        },
        "legend": {
            "orientation": "h", "yanchor": "bottom", "y": 1.02,
            "xanchor": "left", "x": 0,
            "font": {"size": 11.5, "color": INK_MUTED},
            "bgcolor": "rgba(0,0,0,0)",
        },
        "hoverlabel": {
            "bgcolor": "rgba(15,17,21,0.95)",
            "bordercolor": GLASS_1_EDGE,
            "font": {"family": "Inter, sans-serif", "size": 12.5, "color": INK},
        },
        "hovermode": "x unified",
        "separators": ".,",   # keeps Indian-style grouping from Python formatting intact
    }
}


def apply_chart_theme(fig, height: Optional[int] = None, show_legend: Optional[bool] = None):
    """Apply the design system to an existing Plotly figure, in place.

    Only layout properties are modified. Traces, values, axis ranges derived
    from data and hover text are left untouched.

    A tiny compatibility fallback is included because Streamlit deployments
    can resolve a different Plotly minor version than a local environment.
    """
    layout = dict(PLOTLY_TEMPLATE["layout"])
    if height is not None:
        layout["height"] = height
    if show_legend is not None:
        layout["showlegend"] = show_legend

    try:
        fig.update_layout(**layout)
    except (TypeError, ValueError):
        # Keep the chart functional even if an older Plotly build rejects one
        # of the non-essential cosmetic options.
        safe_layout = dict(layout)
        safe_layout.pop("separators", None)
        title = safe_layout.get("title")
        if isinstance(title, dict):
            title = dict(title)
            font = title.get("font")
            if isinstance(font, dict):
                font = {k: v for k, v in font.items() if k in {"family", "size", "color"}}
                title["font"] = font
            safe_layout["title"] = title
        fig.update_layout(**safe_layout)

    return fig

