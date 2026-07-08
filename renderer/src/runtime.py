"""Template Runtime — rendering primitives only.

This module knows how to:
- Open a template, drop its sample slides, expose its layouts and theme.
- Add a slide for a given layout name, with a 'Title Only' fallback.
- Mutate placeholders: set title text, fill bullets, fill a KPI card, fill an
  insight card, bind a chart graphicFrame, bind a table graphicFrame.
- Suppress inherited bullets, shrink long titles to fit, theme charts using
  the template's accent colors.

It does NOT know about:
- Semantic intents, slide types, content meaning.
- Which layout a piece of content "wants" to live in.
- How an author phrases a request.

That's the compiler's job (see ``compiler.py``).
"""
from __future__ import annotations

import functools
import logging
import math
import os
import re
from copy import deepcopy
from dataclasses import dataclass
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_MARKER_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE, PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.util import Inches, Pt


EMU_PER_INCH = 914400
# Document author stamped on every generated deck so output isn't attributed to
# the template's creator (the template carries its author through to outputs).
DOC_AUTHOR = "CS PresentationGPT"
EMU_PER_PT   = 12700


@functools.lru_cache(maxsize=1)
def _treemap_assets():
    """Load and cache the native-treemap chartEx template parts once (avoids
    re-reading three files on every treemap)."""
    base = os.path.join(os.path.dirname(__file__), "assets", "treemap_chartex")
    return {
        "template": open(os.path.join(base, "chartEx_template.xml"), encoding="utf-8").read(),
        "colors": open(os.path.join(base, "colors1.xml"), "rb").read(),
        "style": open(os.path.join(base, "style1.xml"), "rb").read(),
    }

logger = logging.getLogger("runtime")

# Runtime version — stamp on builds for traceability across teams/decks.
__version__ = "2.6.18"
if not logger.handlers:
    h = logging.StreamHandler()
    h.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))
    logger.addHandler(h)
logger.setLevel(logging.INFO)

# OOXML namespaces
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
NSMAP = {"p": P_NS, "a": A_NS}


# ---------------------------------------------------------------------------
# Theme tokens
# ---------------------------------------------------------------------------
DEFAULT_THEME = {
    "accent1": "000F47", "accent2": "82BAFF", "accent3": "CEECFF",
    "accent4": "7B7974", "accent5": "B9B6B1", "accent6": "EBE7E2",
    "dk1": "000F47", "lt1": "FFFFFF",
    "majorLatin": "Marsh Serif", "minorLatin": "Noto Sans",
}
# ---------------------------------------------------------------------------
# OW chart palette and styling constants
# ---------------------------------------------------------------------------
# Series 1–6 cycle through the six theme accents; series 7–15 are explicit
# overrides (orange, yellow, green, purple families). Extracted directly
# from the OW chart templates (Bar.crtx, Column.crtx, etc.) and the
# Chart_Templates_-_20260223.pptx reference deck.
OW_CHART_PALETTE = (
    "000F47",  # 1: navy           (accent1)
    "82BAFF",  # 2: sky blue       (accent2)
    "CEECFF",  # 3: light blue     (accent3)
    "7B7974",  # 4: warm grey      (accent4)
    "B9B6B1",  # 5: light grey     (accent5)
    "EBE7E2",  # 6: greige         (accent6)
    "CB7E03",  # 7: dark orange
    "FFBF00",  # 8: amber          (= BRAND_COLORS["text_highlight"])
    "FFD98A",  # 9: light yellow
    "2F7500",  # 10: dark green
    "6ABF30",  # 11: mid green
    "B0DC92",  # 12: light green
    "8F20DE",  # 13: purple
    "DEB1FF",  # 14: light purple
    "F5E8FF",  # 15: very light purple
)

# Chart style numerics — extracted from the .crtx files
CHART_GRIDLINE_COLOR_HEX = "D1CEC9"     # major gridline color (warm light grey)
CHART_LINE_WEIGHT_EMU    = 9525         # 0.75 pt — axis lines and gridlines
CHART_LINE_SERIES_EMU    = 28575        # 2.25 pt — line-chart series stroke
CHART_LABEL_INSETS_EMU   = (38100, 19050, 38100, 19050)  # l/t/r/b on dLbl bodies
CHART_DOUGHNUT_HOLE_PCT  = 75           # thin doughnut, per OW spec

# Named table style defined in ow_default.pptx's tableStyles.xml. OW Table 1 is
# the clean house style: bold header row with a bottom rule, no fills, dark
# text, no banding. Applying it means we must NOT hard-code per-cell fills.
OW_TABLE_STYLE_1 = "{839DD9DD-9E6C-4910-8AC0-68ADFF6A6AFC}"

# Bar/column gap & overlap by grouping (matches the OW reference exactly)
CHART_BAR_GEOMETRY = {
    ("col", "clustered"): {"gapWidth": 220, "overlap": -50},
    ("col", "stacked"):   {"gapWidth": 219, "overlap": 100},
    ("col", "percentStacked"): {"gapWidth": 219, "overlap": 100},
    ("bar", "clustered"): {"gapWidth": 200, "overlap": -25},
    ("bar", "stacked"):   {"gapWidth": 200, "overlap": 100},
    ("bar", "percentStacked"): {"gapWidth": 200, "overlap": 100},
}


def _hex_is_dark(hex_str: str) -> bool:
    """sRGB relative luminance test — True if a colored fill needs WHITE text overlaid."""
    h = hex_str.replace("#", "")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    luma = 0.299 * r + 0.587 * g + 0.114 * b
    return luma < 110


# Top 10 chart types + scatter
CHART_TYPE_MAP = {
    "column_clustered":   XL_CHART_TYPE.COLUMN_CLUSTERED,
    "column":             XL_CHART_TYPE.COLUMN_CLUSTERED,
    "column_stacked":     XL_CHART_TYPE.COLUMN_STACKED,
    "stacked_column":     XL_CHART_TYPE.COLUMN_STACKED,
    "column_stacked_100": XL_CHART_TYPE.COLUMN_STACKED_100,
    "column_100":         XL_CHART_TYPE.COLUMN_STACKED_100,
    "bar_clustered":      XL_CHART_TYPE.BAR_CLUSTERED,
    "bar":                XL_CHART_TYPE.BAR_CLUSTERED,
    "bar_stacked":        XL_CHART_TYPE.BAR_STACKED,
    "stacked_bar":        XL_CHART_TYPE.BAR_STACKED,
    "line":               XL_CHART_TYPE.LINE,
    "line_markers":       XL_CHART_TYPE.LINE_MARKERS,
    "line_with_markers":  XL_CHART_TYPE.LINE_MARKERS,
    "pie":                XL_CHART_TYPE.PIE,
    "doughnut":           XL_CHART_TYPE.DOUGHNUT,
    "donut":              XL_CHART_TYPE.DOUGHNUT,
    "area_stacked":       XL_CHART_TYPE.AREA_STACKED,
    "stacked_area":       XL_CHART_TYPE.AREA_STACKED,
    "area":               XL_CHART_TYPE.AREA_STACKED,
    "scatter":            XL_CHART_TYPE.XY_SCATTER,
    "xy_scatter":         XL_CHART_TYPE.XY_SCATTER,
}
PIE_LIKE = {XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT}
XY_TYPES = {XL_CHART_TYPE.XY_SCATTER}
# Bar/column families — the only series type that carries <c:invertIfNegative>.
BAR_LIKE = {
    XL_CHART_TYPE.COLUMN_CLUSTERED, XL_CHART_TYPE.COLUMN_STACKED,
    XL_CHART_TYPE.COLUMN_STACKED_100, XL_CHART_TYPE.BAR_CLUSTERED,
    XL_CHART_TYPE.BAR_STACKED, XL_CHART_TYPE.BAR_STACKED_100,
}


def resolve_chart_type(name: str) -> XL_CHART_TYPE:
    if not name:
        return XL_CHART_TYPE.COLUMN_CLUSTERED
    canonical = name.strip().lower().replace("-", "_").replace(" ", "_")
    xl_type = CHART_TYPE_MAP.get(canonical)
    if xl_type is None:
        logger.warning("Unknown chart type %r; defaulting to column_clustered.", name)
        return XL_CHART_TYPE.COLUMN_CLUSTERED
    return xl_type


class Tokens:
    """Central design system for the deck engine — the single place that defines
    every recurring size and spacing, the way CSS variables / design tokens work
    in a front-end. Renderers reference these names instead of hard-coding numbers,
    so the look stays consistent and one change here propagates everywhere.

      - TYPE_*  font sizes (points)
      - SPACE_* layout spacing/geometry (inches)
      - PARA_*  paragraph space-after (points)
      - INSET_* text-box margins (points)

    Colours live in BRAND_COLORS; the OW master's per-level list styles in
    BULLET_LEVELS. Treat this class as the deck's stylesheet.
    """
    # -- Type scale (points) -------------------------------------------------
    TYPE_BODY        = 12   # body / bullet text (matches the master body default)
    TYPE_HEADING     = 12   # bold sub-heading inside a body/pane
    TYPE_KPI         = 26   # dashboard / KPI value (big display number)
    TYPE_KPI_COMPACT = 26   # KPI value when a chart shares the cell
    TYPE_MARKER      = 14   # numbered marker (cycle number, timeline year)
    TYPE_EYEBROW     = 11   # eyebrow / small label / image caption
    TYPE_FOOTNOTE    = 8    # footnote / source line
    TYPE_INSIGHTS    = 18   # insight / "signal" callout body (Marsh Serif min)
    TYPE_CONCLUSION  = 18   # bottom-pinned key-takeaway band body
    TYPE_QUOTE_MAX   = 54   # pull-quote upper bound (large editorial statement)
    TYPE_QUOTE_MIN   = 18   # pull-quote lower bound
    TYPE_MAJOR_MIN   = 18   # Marsh Serif (major font) hard minimum
    TYPE_AUTOFIT_MIN = 10   # shrink-to-fit floor for body/heading text
    TYPE_DIAGRAM_MIN = 8    # tighter floor for dense diagrams (org chart, steps)
    TYPE_TABLE_HEADER = 12  # table header row (== body; bold differentiates)
    TYPE_AXIS        = 12   # chart axis tick labels
    TYPE_CONTENTS    = 16   # table-of-contents row

    # -- Card components (cards.py) ------------------------------------------
    # (eyebrow uses TYPE_EYEBROW=11; heading/delta/body use TYPE_BODY=12)
    TYPE_CARD_VALUE      = 40   # big stat / KPI number on a card
    TYPE_CARD_QUOTE      = 22   # quote-card body text
    TYPE_CARD_QUOTE_MARK = 60   # oversized opening quote glyph
    TYPE_CARD_CAPTION    = 10   # card caption / sub-footnote
    LINE_HEIGHT          = 1.30 # line-height factor for card text measurement

    # -- Spacing / geometry (inches) -----------------------------------------
    SPACE_GAP        = 0.25     # tight gap: stacked cards/callouts, and the NARROW
                                # grid gutter (used only when 0.5" breaches a floor)
    SPACE_GAP_LG     = 0.50     # DEFAULT column/row gutter for grids and two-pane splits
    SPACE_CONTENT_X  = 0.5      # content-area left/right margin
    SPACE_CONTENT_Y  = 1.54     # content-area top (clears a two-line title)
    SPACE_CONTENT_W  = 12.333
    SPACE_CONTENT_H  = 5.06
    SPACE_CONTENT_BOTTOM = 0.20 # gap above the conclusion/footnote
    SPACE_CONCL_H    = 0.55     # conclusion single-line height (matches PowerPoint's rendered cell for 18pt text + insets, so placement math and rendered box agree)
    SPACE_CONCL_LINE = 0.30     # conclusion per-wrapped-line height
    SPACE_SHAPE_MARGIN = 0.15
    SPACE_ICON         = 0.5      # default icon size on cards/columns
    SPACE_ICON_INSET   = 0.2      # icon inset from card top/left
    GRID_MIN_COL_W     = 1.8      # smallest readable column (12pt heading)
    GRID_MIN_ROW_H     = 1.3      # smallest card row (icon + heading + body)
    GRID_COLS_MAX      = 4        # working column max for the smallest-grid solver   # in-shape text margin (steps, conclusion L/R)
    SPACE_TABLE_ROW  = 0.30     # data-table body row height (min; grows on wrap)
    SPACE_TABLE_HEADER = 0.34   # data-table header row height
    CHEVRON_H        = 1.0      # HARD chevron/pentagon shape height — never scaled
    SPACE_AXIS_LABEL_H = 0.32   # matrix axis-label box height
    SPACE_TL_LABEL_H = 0.5      # timeline label box height
    SPACE_HEADING_H  = 0.30     # canvas-block heading-row height
    QUOTE_SHAPE_W    = 10.0     # pull-quote textbox width
    QUOTE_INDENT     = 0.25     # pull-quote left + hanging indent

    # -- Card components: spacing / geometry (inches) ------------------------
    SPACE_CARD_PAD      = 0.22   # card internal padding
    SPACE_CARD_ACCENT_H = 0.07   # optional top accent bar height
    SPACE_CARD_VIZ_GAP  = 0.16   # breathing room under an in-card chart
    SPACE_CARD_ICON     = 0.46   # icon box on icon_card
    VIZ_H_SPARKLINE     = 0.78   # native line micro-chart height
    VIZ_H_BAR           = 0.85   # native column micro-chart height
    VIZ_H_TRAFFIC       = 0.20   # traffic-light row height

    # -- Paragraph space-after (points) --------------------------------------
    PARA_BULLET      = 4
    PARA_PROSE       = 6
    PARA_HEADING     = 2
    PARA_SUBHEAD     = 8
    PARA_QUOTE_GAP   = 12       # pull-quote paragraph spacing (Pt)

    # -- Text-box insets (points) --------------------------------------------
    INSET_X          = 2
    INSET_R          = 5
    INSET_TOP        = 0

    # -- Strokes / outlines (points) -----------------------------------------
    STROKE_W         = 0.75  # all lines & borders (uniform 0.75pt)
    STROKE_W_HEAVY   = 0.75  # emphasis frame (matrix border, org boxes)
    STROKE_W_MED     = 0.75  # connectors (cycle arrows)
    STROKE_HAIRLINE  = 0.75  # thin rule / divider / footnote line
    STROKE_AXIS      = 0.75  # axis line (matrix, timeline)
    STROKE_RULE      = 0.75  # accent rule (kept at the uniform 0.75pt)
    STROKE_PANEL     = 1.5   # key-message panel rule (the sole heavier rule)


@dataclass(frozen=True)
class ContentArea:
    x: float = Tokens.SPACE_CONTENT_X
    y: float = Tokens.SPACE_CONTENT_Y   # matches template body placeholder top
    w: float = Tokens.SPACE_CONTENT_W
    h: float = Tokens.SPACE_CONTENT_H

    @property
    def max_y(self) -> float:
        return self.y + self.h


FALLBACK_CONTENT = ContentArea()


# ---- Alignment design variables -------------------------------------------
# One place to tune every managed alignment. Tables: header row bottom-left,
# body rows top-left. Conclusion & Footnote: bottom-left. ``*_ANCHOR`` are
# python-pptx anchors; ``*_ANCHOR_XML`` are raw OOXML tokens ("t"/"b"/"ctr")
# used where tcPr/bodyPr is written directly.
class Align:
    # Raw OOXML enum bindings — the single place these enums are named.
    LEFT   = PP_ALIGN.LEFT
    CENTER = PP_ALIGN.CENTER
    RIGHT  = PP_ALIGN.RIGHT
    TOP    = MSO_ANCHOR.TOP
    MIDDLE = MSO_ANCHOR.MIDDLE
    BOTTOM = MSO_ANCHOR.BOTTOM
    # Semantic roles (managed alignments).
    TABLE_HEADER_ANCHOR   = BOTTOM
    TABLE_BODY_ANCHOR     = TOP
    TABLE_H               = LEFT
    CONCLUSION_ANCHOR_XML = "b"
    FOOTNOTE_ANCHOR_XML   = "b"
    FOOTNOTE_H            = LEFT

# ---- bottom adornment geometry (Conclusion + Footnote) ---------------------
# Footer baseline: where the bottom edge of both Conclusion and Footnote sits.
# Anchored 0.60" from the slide bottom (slide is 7.5" tall).
FOOTER_BASELINE_Y = 6.90

# Conclusion height. A single line of 18pt text plus the amber-rule treatment
# fits in CONCLUSION_HEIGHT; this is the single-line floor. Longer text is
# estimated per wrapped line (CONCLUSION_LINE_HEIGHT) so the box grows UPWARD
# from its pinned bottom edge — the cell is bottom-anchored, matching the
# footnote's behaviour — instead of overflowing or growing downward.
CONCLUSION_HEIGHT = Tokens.SPACE_CONCL_H
CONCLUSION_LINE_HEIGHT = Tokens.SPACE_CONCL_LINE
# When False, the conclusion uses a FIXED single-line height (no upward growth
# with wrapped line count). Flip to True to restore line-count-aware sizing.
ADJUST_CONCLUSION_HEIGHT = False

# Vertical gap between the Conclusion's bottom edge and the Footnote's top edge
# when both are present. Held constant regardless of how tall the footnote is.
CONCLUSION_FOOTNOTE_GAP = 0.15

# Per-line height for footnote text at 8pt (with default line spacing). Used
# both to estimate footnote shape height and to place the conclusion above it.
FOOTNOTE_LINE_HEIGHT = 0.135


def _wrapped_line_count(text: str, *, size_pt: float, kind: str,
                        width_in: float, chars_per_line: int) -> int:
    """Number of lines `text` wraps to. Uses real font metrics
    (textmetrics.wrapped_lines) when a face is available — exact and size-aware —
    and falls back to a size-calibrated character heuristic otherwise. Shared by
    the footnote and conclusion height estimates so the rule lives in one place."""
    if not text:
        return 0
    try:
        import textmetrics as _tm
        if _tm.available():
            return sum(_tm.wrapped_lines(seg, width_in, size_pt=size_pt, kind=kind)
                       for seg in text.split("\n"))
    except Exception:
        pass
    return sum(max(1, math.ceil(len(seg) / chars_per_line))
               for seg in text.split("\n"))

# Margin between the slide content area and the topmost bottom-adornment.
# Keeps the chart/bullets/etc. from butting right against the conclusion.
CONTENT_BOTTOM_MARGIN = Tokens.SPACE_CONTENT_BOTTOM

# Standard grid spacing between adjacent elements (columns, steps, ovals, cards).
# Every renderer uses these so spacing can't drift between visuals.
GRID_GAP = Tokens.SPACE_GAP
GRID_GAP_LG = Tokens.SPACE_GAP_LG

# Auto-fit floor. When label/body text would render too narrowly, spill outside
# its shape, or collide with neighbouring text, renderers shrink the font down to
# this size — but never below it. If text still doesn't fit at the floor, it is
# left as-is for manual adjustment in PowerPoint (a deliberate, visible signal
# rather than unreadably tiny text).
AUTOFIT_MIN_PT = Tokens.TYPE_AUTOFIT_MIN


def fit_point_size(text: str, box_w_in: float, box_h_in: float, max_pt: float,
                   *, min_pt: float = AUTOFIT_MIN_PT, pad_in: float = 0.12,
                   char_w_frac: float = 0.55, line_h_frac: float = 1.25) -> int:
    """Largest integer point size in ``[min_pt, max_pt]`` at which ``text`` wraps
    inside a ``box_w_in`` × ``box_h_in`` box without breaking a word or exceeding
    the height. Estimation-based (glyph width ≈ ``char_w_frac`` × pt), so it is a
    safe approximation, not exact metrics. Never returns below ``min_pt`` — at the
    floor, overflow is accepted and left for manual adjustment."""
    text = (text or "").strip()
    if not text:
        return int(max_pt)
    usable_w = max(0.15, box_w_in - pad_in)
    usable_h = max(0.15, box_h_in - pad_in)
    words = text.split()
    longest = max((len(w) for w in words), default=1)
    pt = int(max_pt)
    while pt > min_pt:
        char_w = (char_w_frac * pt) / 72.0
        line_h = (line_h_frac * pt) / 72.0
        cpl = max(1, int(usable_w / char_w))
        if longest <= cpl:                      # longest word fits without breaking
            line_len, lines = 0, 1
            for w in words:
                add = len(w) + (1 if line_len else 0)
                if line_len + add <= cpl:
                    line_len += add
                else:
                    lines += 1
                    line_len = len(w)
            if lines * line_h <= usable_h:
                return pt
        pt -= 1
    return int(min_pt)


def resolve_token_ref(ref: str):
    """Resolve a design-token reference string to its live value.

    Supported reference namespaces (the same syntax the component layer and the
    linter use, so there is exactly one resolver):

        {palette.midnightblue} -> PALETTE["midnightblue"]    -> "000F47"
        {BRAND_COLORS.x}      -> BRAND_COLORS["x"]          (may itself be an accentN)
        {BRAND_RULES.x}       -> BRAND_RULES["x"]
        {Tokens.TYPE_KPI}     -> Tokens.TYPE_KPI            -> 26

    Returns the resolved value, or raises KeyError/AttributeError with the
    offending reference so the linter can report a broken-ref finding. A bare
    (non-braced) string is returned unchanged — it is already a literal.
    """
    if not (isinstance(ref, str) and ref.startswith("{") and ref.endswith("}")):
        return ref
    body = ref[1:-1]
    ns, _, key = body.partition(".")
    if not key:
        raise KeyError(f"malformed token reference {ref!r}")
    if ns == "palette":
        return PALETTE[key]
    if ns == "BRAND_COLORS":
        return BRAND_COLORS[key]
    if ns == "BRAND_RULES":
        return BRAND_RULES[key]
    if ns == "Tokens":
        return getattr(Tokens, key)
    raise KeyError(f"unknown token namespace in reference {ref!r}")


def resolve_color_ref(ref: str) -> str:
    """Resolve a colour reference all the way to a 6-hex string. Colour roles may
    chain (BRAND_COLORS['subtle_highlight'] == 'accent3' -> DEFAULT_THEME['accent3']),
    so this follows accentN / dk1 / lt1 indirections to a literal hex."""
    val = resolve_token_ref(ref)
    seen = set()
    while isinstance(val, str) and val in DEFAULT_THEME and val not in seen:
        seen.add(val)
        val = DEFAULT_THEME[val]
    return str(val).lstrip("#").upper()


def _relative_luminance(hex_str: str) -> float:
    """WCAG relative luminance of a 6-hex colour, for contrast ratios."""
    h = hex_str.lstrip("#")
    def chan(c):
        c = c / 255.0
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    return 0.2126 * chan(r) + 0.7152 * chan(g) + 0.0722 * chan(b)


def contrast_ratio(hex_a: str, hex_b: str) -> float:
    """WCAG contrast ratio between two 6-hex colours (1.0–21.0)."""
    la, lb = _relative_luminance(hex_a), _relative_luminance(hex_b)
    hi, lo = max(la, lb), min(la, lb)
    return round((hi + 0.05) / (lo + 0.05), 2)


def design_system_markdown() -> str:
    """Render the design-system reference from the LIVE tokens, so the document is
    generated from code and can never drift. Printed by
    ``generate_deck.py --list-tokens``. Token tables are generated; the prose
    under each section is merged from the human-authored ``design_prose`` module,
    and the component inventory from ``design_components``."""
    T = Tokens
    try:
        import design_prose as _prose
        SECTION_PROSE = _prose.SECTION_PROSE
        COMPONENT_PROSE = _prose.COMPONENT_PROSE
    except Exception:
        SECTION_PROSE, COMPONENT_PROSE = {}, {}

    def prose(key: str) -> str:
        p = SECTION_PROSE.get(key, "").strip()
        return f"\n{p}\n" if p else ""

    def rows(prefix, unit):
        out = []
        for k in sorted(vars(T)):
            if k.startswith(prefix):
                out.append(f"| `{k}` | {getattr(T, k)}{unit} |")
        return "\n".join(out)

    bl = "\n".join(
        f"| {lvl} | {t.role} | {t.size_pt} | "
        f"{'Marsh Serif' if t.font_kind == 'major' else 'Noto Sans'}"
        f"{' bold' if t.bold else ''} | "
        f"{t.bullet_char or '—'} | {t.mar_l_in:g}\" / {t.indent_in:g}\" | "
        f"{t.space_before:g}/{t.space_after:g}pt |"
        for lvl, t in sorted(BULLET_LEVELS.items()))
    # Colour tables: raw OW palette (explicit hex, top) + semantic roles resolved
    # to literal hex with provenance (Via) and a folded-in usage line.
    try:
        import design_prose as _cp
        _COLOR_USAGE = _cp.COLOR_USAGE
        _PALETTE_USAGE = _cp.PALETTE_USAGE
    except Exception:
        _COLOR_USAGE, _PALETTE_USAGE = {}, {}

    palette_tbl = "\n".join(
        f"| `{name}` | `#{hexv.upper()}` | {_PALETTE_USAGE.get(name, '')} |"
        for name, hexv in PALETTE.items())

    def _role_row(role, raw):
        try:
            lit = "#" + resolve_color_ref(f"{{BRAND_COLORS.{role}}}")
        except Exception:
            lit = str(raw)
        via = f"`{raw}`" if raw in DEFAULT_THEME else "—"
        return f"| `{role}` | `{lit}` | {via} | {_COLOR_USAGE.get(role, '')} |"

    colors = "\n".join(_role_row(k, v) for k, v in BRAND_COLORS.items())
    rules = "\n".join(f"| `{k}` | `{v}` |" for k, v in BRAND_RULES.items())
    chart = (
        f"| `CHART_LINE_WEIGHT_EMU` | {CHART_LINE_WEIGHT_EMU} EMU ({CHART_LINE_WEIGHT_EMU/EMU_PER_PT:g} pt) |\n"
        f"| `CHART_LINE_SERIES_EMU` | {CHART_LINE_SERIES_EMU} EMU ({CHART_LINE_SERIES_EMU/EMU_PER_PT:g} pt) |\n"
        f"| `CHART_DOUGHNUT_HOLE_PCT` | {CHART_DOUGHNUT_HOLE_PCT}% |\n"
        f"| `CHART_GRIDLINE_COLOR_HEX` | `{CHART_GRIDLINE_COLOR_HEX}` |")
    align = "\n".join(
        f"| `{k}` | `{getattr(Align, k)}` |"
        for k in vars(Align) if not k.startswith("_"))

    # ---- Component inventory (generated from design_components) -------------
    def _component_section() -> str:
        try:
            import design_components as _dc
        except Exception:
            return ""
        blocks = []
        for name in _dc.component_names():
            spec = _dc.COMPONENTS[name]
            slot_rows = "\n".join(
                f"| `{s['name']}` | {s['type']} | {'yes' if s['required'] else 'no'} | {s['note']} |"
                for s in spec["slots"])
            tok_rows = []
            for role, ref in spec["tokens"].items():
                try:
                    val = resolve_token_ref(ref)
                    if isinstance(val, str) and (val in DEFAULT_THEME or len(val) == 6):
                        val = f"#{resolve_color_ref(ref)}"
                except Exception as exc:                       # pragma: no cover
                    val = f"<broken: {exc}>"
                tok_rows.append(f"| `{role}` | `{ref}` | `{val}` |")
            tok_block = "\n".join(tok_rows)
            ov = spec.get("overrides") or {}
            ov_line = (", ".join(f"`{k}`" for k in ov) if ov else "—")
            # contrast read-out for declared pairs
            cpairs = []
            for bg_role, tx_role in spec.get("contrast_pairs", []):
                try:
                    bg = resolve_color_ref(spec["tokens"][bg_role])
                    tx = resolve_color_ref(spec["tokens"][tx_role])
                    cr = contrast_ratio(bg, tx)
                    verdict = "AA" if cr >= 4.5 else ("AA-large" if cr >= 3.0 else "FAIL")
                    cpairs.append(f"`{tx_role}` on `{bg_role}` = {cr}:1 ({verdict})")
                except Exception:
                    pass
            cp_line = ("; ".join(cpairs) if cpairs else "—")
            rationale = COMPONENT_PROSE.get(name, "").strip()
            blocks.append(
                f"### `{name}`\n"
                f"{spec['summary']}\n\n"
                f"{rationale + chr(10) + chr(10) if rationale else ''}"
                f"**Slots**\n\n| Slot | Type | Required | Note |\n| --- | --- | --- | --- |\n{slot_rows}\n\n"
                f"**Bound tokens**\n\n| Role | Reference | Resolves to |\n| --- | --- | --- |\n{tok_block}\n\n"
                f"**Overrides:** {ov_line}  \n"
                f"**Contrast (WCAG):** {cp_line}\n")
        if not blocks:
            return ""
        return ("\n## Components — block inventory (generated)\n"
                "Each recurring block declares the content slots it consumes and "
                "the design tokens it binds. Bindings are *references*, resolved "
                "live; the linter checks them for broken refs, orphaned palette "
                "roles, and WCAG contrast. The prose says when to reach for each.\n\n"
                + "\n".join(blocks))

    components_md = _component_section()

    return f"""# OW Presentation GPT — Design System

**Runtime v{__version__}. Generated from code** (`Tokens`, `BRAND_COLORS`,
`BRAND_RULES`, `BULLET_LEVELS`). Do not hand-edit — regenerate with
`generate_deck.py --list-tokens`. This is the single source of truth for every
quantitative design fact; knowledge files reference it rather than restating it.

Fonts and the colour palette come from the `ow_default.pptx` theme (Marsh Serif
major / Noto Sans minor; midnight blue `#000F47`, gold `#FFBF00`, light blue
`#CEECFF`, cream `#F7F3EE`, warm greys `#7B7974` / `#B9B6B1`). Everything below is
enforced by the engine — the GPT never sets these.
{prose("overview")}
## Colours
The raw OW palette is the literal colour set; the semantic roles below reference
into it so there is one place to change a colour. Reach for a *role*, not a hex.

**OW palette (literal)**

| Name | Hex | Use for |
| --- | --- | --- |
{palette_tbl}

**Semantic roles — BRAND_COLORS**

| Role | Resolves to | Via | Use for |
| --- | --- | --- | --- |
{colors}
{prose("colors")}
## Type scale — TYPE_* (points)
| Token | Value |
| --- | --- |
{rows("TYPE_", "")}

Two floors: Marsh Serif (major) never below `TYPE_MAJOR_MIN`; auto-fit body/label
text never below `TYPE_AUTOFIT_MIN`.
{prose("type")}
## Spacing & geometry — SPACE_* (inches)
| Token | Value |
| --- | --- |
{rows("SPACE_", '"')}
{prose("space")}
## Paragraph space-after — PARA_* (points)
| Token | Value |
| --- | --- |
{rows("PARA_", "")}
{prose("para")}
## Text-box insets — INSET_* (points)
| Token | Value |
| --- | --- |
{rows("INSET_", "")}
{prose("inset")}
## Strokes / outlines — STROKE_* (points)
| Token | Value |
| --- | --- |
{rows("STROKE_", "")}

A weight of `0` means no outline; this is the one place to change shape borders.
{prose("stroke")}
## Charts — CHART_* (chart-specific constants)
| Constant | Value |
| --- | --- |
{chart}
{prose("chart")}
## Body list styles — BULLET_LEVELS (from the OW master)
| Level | Role | Size | Font | Glyph | Margin / hang | Space b/a |
| --- | --- | --- | --- | --- | --- | --- |
{bl}
{prose("bullets")}
## Alignment — Align (managed anchors + horizontal)
Table header row is bottom-left; body rows top-left; Conclusion and Footnote
bottom-left. These are the only alignments set from variables; diagram-internal
alignments remain renderer-local.

| Variable | Value |
| --- | --- |
{align}
{prose("align")}
## Hard rules — BRAND_RULES
| Rule | Value |
| --- | --- |
{rules}
{prose("rules")}
{components_md}
## Changing the system
Edit the value in `Tokens` / `BRAND_COLORS` / `BRAND_RULES`, run `pytest`
(`test_tokens_are_single_source_of_truth` + the invariant tests guard it), rebuild
the runtime zip, and regenerate this file with `--list-tokens`. Component slots
and bindings live in `design_components.py`; their rationale in `design_prose.py`.
"""


def iter_bullets(items, level: int = 1):
    """Walk a (possibly nested) bullet list, yielding ``(level, text)`` for every
    item depth-first. An item is a plain string, or a dict with ``text``/``label``
    and nested ``bullets``/``sub``. This is the single source of truth for the
    bullet tree — fill_bullets, draw_columns, and the compiler's density estimate
    all walk it through here so nesting behaviour stays consistent."""
    for it in items or []:
        if isinstance(it, dict):
            text = it.get("text", it.get("label", ""))
            subs = it.get("bullets") or it.get("sub") or []
        else:
            text, subs = str(it), []
        yield level, text
        if subs:
            yield from iter_bullets(subs, level + 1)


# =========================================================================
# Brand spec — codifies hard rules and the OW master's text style table.
# Centralized so strategies and runtime primitives reference one source of
# truth instead of typing magic numbers/colors everywhere.
# =========================================================================

# Hard rules (from the brand owner; never violated by code in this module).
BRAND_RULES = {
    # The "primary" / major Latin font (Marsh Serif). NEVER bold; can be
    # regular, italic, or underlined for emphasis.
    "primary_font_never_bold": True,
    # Minimum size for the primary (Marsh Serif) font anywhere. Marsh Serif
    # text is never rendered below this.
    "primary_font_min_pt": Tokens.TYPE_MAJOR_MIN,
    # Default size for callout/insight card body text (a "such shapes" default).
    "insight_body_pt": Tokens.TYPE_INSIGHTS,
}

# Default colors. Strategies reach for these by role rather than by hex
# value so we only have one place to change the palette.
BRAND_COLORS = {
    # Default fill for any decorative shape (chevrons, pyramid bands, org
    # boxes, cycle ovals) when not specifically highlighted.
    "default_fill":      "F7F3EE",   # warm cream
    # "Subtle highlight" — emphasis without dominating the slide.
    "subtle_highlight":  "accent3",  # CEECFF (light blue) on OW
    # "Strong highlight" — the primary attention-grab.
    "strong_highlight":  "accent1",  # 000F47 (navy) on OW
    # Inline text-highlight (used to mark a word/phrase, not a shape).
    "text_highlight":    "FFBF00",   # amber, same as Conclusion's left rule
}

# Semantic chart fills — charts need distinct colours beyond the text palette.
# Centralised here (not inline) so the chart language stays governable.
CHART_COLORS = {"up": "5BA150", "down": "C0504D"}      # waterfall delta up / down
TREEMAP_PALETTE = ["000F47", "82BAFF", "CEECFF", "7B7974", "B9B6B1"]

# Locale codes (LCID) used to prefix chart number formats so PowerPoint renders
# thousand separators per the deck's language: en-US -> "1,000", de-DE -> "1.000",
# fr-FR -> "1 000". Set via deck_spec["language"] (default "en-US"); consumed in
# _build_ser_dLbls to prefix "[$-{LCID}]" onto the format code.
LCID_MAP = {
    "en-US": "409", "en-GB": "809",
    "de-DE": "407", "de-AT": "C07", "de-CH": "807",
    "fr-FR": "40C", "fr-CH": "100C",
    "it-IT": "410",
    "es-ES": "C0A", "es-MX": "80A",
    "nl-NL": "413",
    "pt-PT": "816", "pt-BR": "416",
    "pl-PL": "415",
    "sv-SE": "41D",
    "da-DK": "406",
    "no-NO": "414",
    "fi-FI": "40B",
    "cs-CZ": "405",
    "ja-JP": "411",
    "zh-CN": "804", "zh-TW": "404",
    "ko-KR": "412",
}

# Stopword-frequency fingerprints for auto-detecting deck language when the
# deck spec doesn't set one explicitly. Words are the most frequent function
# words per language — high signal, low false-positive rate on any content
# longer than a headline. English is the fallback for numbers-only decks or
# ties. Detection is DETERMINISTIC (unlike langdetect, which is stochastic).
_LANG_STOPWORDS = {
    "en-US": {"the", "and", "of", "to", "in", "is", "that", "for", "with",
              "this", "on", "are", "as", "at", "be", "by", "it", "from"},
    "de-DE": {"der", "die", "das", "und", "ist", "für", "fuer", "mit", "nicht",
              "auf", "ein", "eine", "den", "dem", "des", "im", "zu", "von",
              "sich", "aber", "auch", "sind", "werden", "wird", "nach"},
    "fr-FR": {"le", "la", "les", "des", "et", "une", "un", "à", "en", "dans",
              "pour", "avec", "par", "sur", "est", "sont", "ne", "pas",
              "que", "qui", "au", "aux", "du"},
    "it-IT": {"il", "la", "di", "e", "un", "una", "per", "con", "in", "non",
              "del", "che", "gli", "lo", "le", "sono", "è", "come", "ma",
              "alla", "dei", "delle", "della", "nel", "nella"},
    "es-ES": {"el", "la", "de", "y", "un", "una", "en", "es", "con", "para",
              "por", "del", "no", "se", "los", "las", "que", "al", "su",
              "sus", "como", "más", "pero", "también"},
    "nl-NL": {"de", "het", "van", "en", "een", "is", "in", "met", "voor",
              "dat", "op", "niet", "aan", "ook", "zijn", "worden", "wordt",
              "door", "naar", "over", "bij", "uit"},
    "pt-PT": {"o", "a", "de", "e", "um", "uma", "para", "com", "em", "do",
              "da", "dos", "das", "que", "é", "não", "se", "por", "no",
              "na", "como", "os", "as"},
    "pl-PL": {"i", "w", "na", "z", "do", "że", "to", "jest", "nie", "jak",
              "ale", "o", "po", "przez", "się", "są", "być", "który",
              "która", "które", "za", "od"},
    "sv-SE": {"och", "att", "det", "som", "är", "med", "för", "av", "en",
              "ett", "den", "till", "på", "de", "har", "inte", "vi",
              "kan", "eller", "inom"},
}


def detect_language(text: str) -> tuple[str, int, str | None]:
    """Deterministic stopword-frequency language detection.

    Returns ``(lang, hits, runner_up)`` where:
      - ``lang`` is the BCP-47 code with the most stopword hits (default en-US)
      - ``hits`` is the winner's raw hit count
      - ``runner_up`` is the second-place BCP-47 code, or None if the winner
        beats every other language by a clear margin

    A confidence threshold is enforced elsewhere: callers should fall back to
    ``en-US`` when ``hits < 3`` or when the winner ties the runner-up, since
    both indicate a mostly-numeric or too-short input.
    """
    tokens = [w.lower() for w in re.findall(r"[A-Za-zÄÖÜäöüßÀ-ÿ]+", text or "")]
    if not tokens:
        return "en-US", 0, None
    scores = {lang: sum(1 for t in tokens if t in bag)
              for lang, bag in _LANG_STOPWORDS.items()}
    ranked = sorted(scores.items(), key=lambda kv: (-kv[1], kv[0]))
    winner, hits = ranked[0]
    runner = ranked[1][0] if len(ranked) > 1 and ranked[1][1] > 0 else None
    return winner, hits, runner

# Raw OW palette — the single source for the literal hex values used by the
# component and infographic drawers (which must not redefine their own copies).
# Names are distinct by role so two different greys can't collide under one name.
PALETTE = {
    "midnightblue": "000F47",   # primary text / strong fill
    "cream":        "F7F3EE",   # default decorative fill
    "skyblue":      "82BAFF",   # mid blue
    "lblue":        "CEECFF",   # subtle highlight
    "gold":         "FFBF00",   # text highlight / accent
    "grey":         "7B7974",   # secondary text
    "lgrey":        "B9B6B1",   # mid grey — hairlines and dividers
    "pale":         "EBE7E2",   # pale grey — neutral track / fill
    "white":        "FFFFFF",
}
FONT_MAJOR = "Marsh Serif"   # primary (never bold)
FONT_MINOR = "Noto Sans"

# Component block kinds usable inside `compose` regions / graphic panes.
_TITLE_KEEP = {  # common proper nouns kept capitalised when sentence-casing a title
    "China", "India", "France", "Germany", "Japan", "Singapore", "Brazil",
    "Canada", "Europe", "European", "America", "American", "Africa", "Asia",
    "Australia", "Korea", "Mexico", "Spain", "Italy", "Russia", "Saudi",
    "Arabia", "Kingdom", "States", "United", "UK", "US", "USA", "EU", "UAE",
    "Linky", "Oliver", "Wyman", "Marsh", "January", "February", "March",
    "April", "May", "June", "July", "August", "September", "October",
    "November", "December",
}
_PUNCT = "\"'()[]{}.,:;!?\u2014-"


def _is_titlecase(core: str) -> bool:
    return bool(core) and core[:1].isupper() and core[1:].islower() and core.isalpha()


def _preserve_title_token(w: str) -> bool:
    core = w.strip(_PUNCT)
    if not core:
        return True
    if core.isupper():
        return True                                   # acronym (AI, GDP, WEF)
    if any(ch.isdigit() for ch in core):
        return True                                   # 5G, 4Cs, Q3, 2030
    if any(ch.isupper() for ch in core[1:]):
        return True                                   # internal caps (iPhone)
    return core in _TITLE_KEEP                         # known proper noun


def normalize_chart_parts(pptx_path: str) -> None:
    """Post-build sweep over every chart part (path-independent enforcement):
    force all axis tick marks to ``none`` (the OW 'no tick marks ever' rule) and
    drop the redundant value-axis scale on charts whose bars already carry data
    labels. Bulletproof because it runs on the finished file, after every
    chart-building path."""
    import zipfile, re, os
    try:
        zin = zipfile.ZipFile(pptx_path, "r")
    except Exception:
        return

    def _hide_value_axis(xml: str) -> str:
        def fix(m):
            block = m.group(0)
            if "<c:delete" in block:
                block = re.sub(r'<c:delete[^>]*/>', '<c:delete val="1"/>', block, count=1)
            else:
                block = block.replace('</c:scaling>',
                                      '</c:scaling><c:delete val="1"/>', 1)
            return block
        return re.sub(r'<c:valAx>.*?</c:valAx>', fix, xml, flags=re.S)

    items = []
    for info in zin.infolist():
        data = zin.read(info.filename)
        if re.match(r"ppt/charts/chart\d+\.xml$", info.filename):
            xml = data.decode("utf-8", "ignore")
            xml = re.sub(r'(majorTickMark val=")\w+(")', r"\1none\2", xml)
            xml = re.sub(r'(minorTickMark val=")\w+(")', r"\1none\2", xml)
            if re.search(r'<c:showVal val="1"', xml):
                xml = _hide_value_axis(xml)
            data = xml.encode("utf-8")
        items.append((info, data))
    zin.close()
    tmp = pptx_path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for info, data in items:
            zout.writestr(info, data)
    os.replace(tmp, pptx_path)


def solve_grid(n_cells, cols=None, rows=None, cols_max=4):
    """Smallest grid for n_cells. Rows are minimised first (the slide is ~2.44:1
    wide, so a row costs more than a column): M is the smallest count with
    ceil(n/M) <= cols_max, then N = ceil(n/M). Explicit cols/rows override."""
    import math
    if cols and rows:
        return int(cols), int(rows)
    if cols:
        return int(cols), max(1, math.ceil(n_cells / int(cols)))
    if rows:
        return max(1, math.ceil(n_cells / int(rows))), int(rows)
    M = 1
    while math.ceil(n_cells / M) > cols_max:
        M += 1
    return math.ceil(n_cells / M), M


def resolve_grid_dims(cells, cols=None, rows=None):
    """Resolve a grid's final (N, M) exactly the way ``draw_grid`` lays it out,
    plus the author-declared (decl_N, decl_M). Single source of the grid geometry
    used by both the runtime (to place cells) and the gate (to validate them)."""
    def _cs(c):
        return max(1, int(c.get("colspan", 1) or 1))

    def _rs(c):
        return max(1, int(c.get("rowspan", 1) or 1))

    eff = sum(_cs(c) * _rs(c) for c in cells) or 1
    decl_N = int(cols) if cols else None
    decl_M = int(rows) if rows else None
    N, M = solve_grid(eff, cols, rows, Tokens.GRID_COLS_MAX)
    N = max(N, max((int(c.get("col", 0)) + _cs(c) for c in cells), default=1))
    M = max(M, max((int(c.get("row", 0)) + _rs(c) for c in cells), default=1))
    return N, M, decl_N, decl_M


def to_sentence_case(text: str) -> str:
    """Enforce sentence case on a title: first word capitalised, the rest
    lowercased — but acronyms, numbers, internal-caps and known proper nouns are
    preserved. Titles already in sentence case are left untouched (so correctly
    capitalised proper nouns are never harmed)."""
    if not text or not text.strip():
        return text
    words = text.split()
    mid = words[1:]
    tc = [w for w in mid if _is_titlecase(w.strip(_PUNCT))]
    if mid and len(tc) / len(mid) < 0.6:              # already sentence-ish
        if words[0][:1].islower():
            words[0] = words[0][:1].upper() + words[0][1:]
        return " ".join(words)
    out = []
    for i, w in enumerate(words):
        if i == 0:
            out.append(w[:1].upper() + w[1:] if w[:1].islower() else w)
        elif _preserve_title_token(w):
            out.append(w)
        elif _is_titlecase(w.strip(_PUNCT)):
            out.append(w[:1].lower() + w[1:])
        else:
            out.append(w)
    return " ".join(out)


_COMPONENT_BLOCK_KINDS = frozenset({
    "stat_callout", "stat_cards", "stats", "icon_rows", "icon_list",
    "compare", "sticker", "infographic",
})

# Card components (cards.py) — self-measuring blocks that HUG their content and
# equalize within a row under the autolayout solver.
_CARD_BLOCK_KINDS = frozenset({
    "kpi_card", "quote_card", "stat_card", "icon_card",
})

# Block registry metadata — minimum sensible size (inches) per block kind. The
# composer's layout grammar warns when a region falls below these, so an
# on-grid-but-cramped slide is flagged rather than silently shipped.
BLOCK_MIN_SIZE = {
    "text": (1.4, 0.5), "chart": (3.0, 1.8), "waterfall": (3.2, 2.0),
    "treemap": (2.8, 2.0), "cycle": (2.4, 1.8), "pyramid": (2.4, 1.8),
    "process": (3.6, 1.0), "org_chart": (2.8, 1.8), "matrix": (2.8, 2.6),
    "timeline": (3.6, 1.0), "growing_steps": (2.8, 1.4),
    "kpi": (1.2, 1.0), "table": (2.5, 0.8), "image": (1.5, 1.0),
    "rule_text": (1.4, 0.6), "kpi_callout": (1.2, 0.7), "key_message": (2.4, 2.0),
    "kpi_card": (1.6, 1.2), "quote_card": (2.4, 1.6), "stat_card": (1.4, 1.0),
    "icon_card": (1.8, 0.7),
}

# Per-level text styles, extracted from the OW master's <p:bodyStyle>.
# Levels are 0-indexed (matching python-pptx and slide-level <a:p lvl=...>).
# The OOXML levels are 1-indexed (lvl1pPr ... lvl9pPr) — convert by adding 1.
#
# When a paragraph lives in a CONTENT placeholder, setting its level here is
# enough — the master's lstStyle drives size/font/color/bullet for free.
# When a paragraph lives in a free TEXT BOX (no placeholder inheritance),
# strategies must apply these properties explicitly. ``apply_text_level``
# below does this.
@dataclass(frozen=True)
class TextLevel:
    role: str          # human-readable: "Body", "Bullet 1", "Heading", etc.
    size_pt: float     # font size in points
    font_kind: str     # "minor" (body, +mn-lt) or "major" (primary, +mj-lt)
    bold: bool         # weight; major font is *never* True per BRAND_RULES
    bullet_char: Optional[str]  # None = no bullet; otherwise the glyph
    mar_l_in: float    # left margin (inches) — used for hanging indents
    indent_in: float   # paragraph indent (inches; negative = hanging)
    space_before: float = 0.0   # paragraph space-before (pt)
    space_after: float = 0.0    # paragraph space-after (pt)


BULLET_LEVELS: Dict[int, TextLevel] = {
    0: TextLevel("Body",            12, "minor", False, None, 0.000,  0.000),
    1: TextLevel("Bullet 1",        12, "minor", False, "•",  0.197, -0.197, space_after=Tokens.PARA_BULLET),
    2: TextLevel("Bullet 2",        12, "minor", False, "–",  0.394, -0.197, space_after=Tokens.PARA_BULLET),
    3: TextLevel("Bullet 3",        12, "minor", False, "-",  0.591, -0.197, space_after=Tokens.PARA_BULLET),
    4: TextLevel("Heading",         12, "minor", True,  None, 0.000,  0.000, space_after=Tokens.PARA_SUBHEAD),  # only level that's bold
    5: TextLevel("Subheading",      12, "minor", False, None, 0.000,  0.000),
    6: TextLevel("Small primary",   26, "major", False, None, 0.000,  0.000),
    7: TextLevel("Medium primary",  40, "major", False, None, 0.000,  0.000),
    8: TextLevel("Large primary",   54, "major", False, None, 0.000,  0.000),
}


def load_theme_from_template(prs: Presentation) -> Dict[str, str]:
    theme = dict(DEFAULT_THEME)
    master = prs.slide_master
    theme_part = None
    for rel in master.part.rels.values():
        if rel.reltype.endswith("/theme"):
            theme_part = rel.target_part
            break
    if theme_part is None:
        return theme
    try:
        theme_xml = etree.fromstring(theme_part.blob)
    except etree.XMLSyntaxError:
        return theme
    clr = theme_xml.find(".//a:clrScheme", {"a": A_NS})
    if clr is not None:
        for child in clr:
            tag = etree.QName(child).localname
            if tag in {"dk1", "lt1", "accent1", "accent2", "accent3",
                       "accent4", "accent5", "accent6"}:
                hex_v = _read_color(child)
                if hex_v:
                    theme[tag] = hex_v
    fonts = theme_xml.find(".//a:fontScheme", {"a": A_NS})
    if fonts is not None:
        for key, child_name in (("majorLatin", "majorFont"),
                                ("minorLatin", "minorFont")):
            block = fonts.find(f"a:{child_name}", {"a": A_NS})
            if block is None:
                continue
            latin = block.find("a:latin", {"a": A_NS})
            if latin is not None and latin.get("typeface"):
                theme[key] = latin.get("typeface")
    return theme


def _read_color(entry: etree._Element) -> Optional[str]:
    for child in entry:
        if etree.QName(child).localname == "srgbClr":
            return child.get("val")
        if etree.QName(child).localname == "sysClr":
            return child.get("lastClr") or child.get("val")
    return None


# ---------------------------------------------------------------------------
# Runtime
# ---------------------------------------------------------------------------

def _tm_layout(sizes, x, y, dx, dy):
    """Lay a run of areas along the shorter side of the (x,y,dx,dy) box."""
    if dx >= dy:
        width = sum(sizes) / dy
        out, yy = [], y
        for s in sizes:
            out.append({"x": x, "y": yy, "dx": width, "dy": s / width}); yy += s / width
        return out
    height = sum(sizes) / dx
    out, xx = [], x
    for s in sizes:
        out.append({"x": xx, "y": y, "dx": s / height, "dy": height}); xx += s / height
    return out


def _tm_leftover(sizes, x, y, dx, dy):
    if dx >= dy:
        width = sum(sizes) / dy
        return (x + width, y, dx - width, dy)
    height = sum(sizes) / dx
    return (x, y + height, dx, dy - height)


def _tm_worst(sizes, x, y, dx, dy):
    rects = _tm_layout(sizes, x, y, dx, dy)
    return max(max(r["dx"] / r["dy"], r["dy"] / r["dx"]) for r in rects)


def squarify_treemap(values, x, y, dx, dy):
    """Squarified treemap (Bruls et al.): return a rect {x,y,dx,dy} per value,
    sized proportionally to fill the (x,y,dx,dy) box with near-square tiles.
    ``values`` should be positive and sorted descending for best results."""
    total = float(sum(values))
    if total <= 0 or dx <= 0 or dy <= 0:
        return []
    area = dx * dy
    sizes = [float(v) * area / total for v in values]

    def _rec(sizes, x, y, dx, dy):
        if not sizes:
            return []
        if len(sizes) == 1:
            return _tm_layout(sizes, x, y, dx, dy)
        i = 1
        while (i < len(sizes)
               and _tm_worst(sizes[:i], x, y, dx, dy)
               >= _tm_worst(sizes[:i + 1], x, y, dx, dy)):
            i += 1
        cur, rest = sizes[:i], sizes[i:]
        lx, ly, ldx, ldy = _tm_leftover(cur, x, y, dx, dy)
        return _tm_layout(cur, x, y, dx, dy) + _rec(rest, lx, ly, ldx, ldy)

    return _rec(sizes, x, y, dx, dy)


def _is_light_hex(hexc: str) -> bool:
    """True if a hex colour is light enough to need dark text on top."""
    h = hexc.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return (0.299 * r + 0.587 * g + 0.114 * b) / 255.0 > 0.6


class TemplateRuntime:
    """Rendering primitives over a PPTX template."""

    def __init__(self, template_path: str):
        self.prs = Presentation(template_path)
        self.theme = load_theme_from_template(self.prs)
        self._remove_existing_slides()
        self._patch_template_malformations()
        self.layouts = {layout.name: layout for layout in self.prs.slide_layouts}
        if "Title Only" not in self.layouts:
            raise RuntimeError("Template is missing the required 'Title Only' fallback layout.")
        # Per-slide override for the content area. Set when a slide has
        # Conclusion or Footnote adornments so canvas strategies (chevrons,
        # pyramid, cycle, org chart) draw within the reduced area instead
        # of overlapping the bottom adornments.
        self._content_area_override: Optional[ContentArea] = None
        # Deck language (BCP-47). Drives chart number formatting via LCID
        # prefix so "1,000" or "1.000" or "1 000" matches the audience's
        # locale. Set from deck_spec["language"] by the compiler; default en-US.
        self.language: str = "en-US"

    @property
    def _number_lcid(self) -> str:
        """LCID (hex string, no leading zeros) for the current language, used
        as the ``[$-{lcid}]`` prefix in chart number-format codes. Falls back
        to en-US (409) for unknown languages."""
        return LCID_MAP.get(self.language, LCID_MAP["en-US"])

    def _patch_template_malformations(self) -> None:
        """Walk every slide layout and master and add ``lang="en-US"`` to
        any bare ``<a:endParaRPr/>`` element. PowerPoint's schema
        validator rejects the bare form and triggers a "repair" prompt
        on open. Many real-world templates (including ow_default.pptx)
        ship with the bare form because the authoring tool elided the
        attribute; we fix it once at load time so every saved deck is
        clean.
        """
        targets = list(self.prs.slide_masters)
        for master in self.prs.slide_masters:
            targets.extend(master.slide_layouts)
        patched = 0
        for part in targets:
            xml = part.element
            for epr in xml.iter(f"{{{A_NS}}}endParaRPr"):
                if epr.get("lang") is None:
                    epr.set("lang", "en-US")
                    patched += 1
        if patched:
            logger.debug("Patched %d empty <a:endParaRPr/> in template parts", patched)

    def set_content_area_override(self, area: Optional[ContentArea]) -> None:
        self._content_area_override = area

    def content_area(self) -> ContentArea:
        """The effective content area for canvas strategies. Equal to the
        global FALLBACK_CONTENT unless overridden for the current slide."""
        return self._content_area_override or FALLBACK_CONTENT

    def _remove_existing_slides(self) -> None:
        xml_slides = self.prs.slides._sldIdLst  # noqa: SLF001
        for sld_id in list(xml_slides):
            self.prs.part.drop_rel(sld_id.rId)
            xml_slides.remove(sld_id)

    # ---- theme helpers --------------------------------------------------

    def rgb(self, token: str) -> RGBColor:
        v = self.theme.get(token, DEFAULT_THEME.get(token, "000000")).replace("#", "")
        return RGBColor(int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16))

    @staticmethod
    def rgb_hex(hex_str: str) -> RGBColor:
        """Resolve a literal 6-digit hex color (e.g. 'F7F3EE') to an RGBColor.
        Used for brand colors that aren't theme tokens, like the default fill."""
        v = hex_str.replace("#", "")
        return RGBColor(int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16))

    def resolve_color(self, token_or_hex: str) -> RGBColor:
        """Accept either a theme token ('accent1', 'dk1', 'lt1') or a 6-digit
        hex string ('F7F3EE') and return the corresponding RGBColor.
        Used by strategies that pull colors from BRAND_COLORS where values
        can be either form."""
        if len(token_or_hex) == 6 and all(c in "0123456789abcdefABCDEF" for c in token_or_hex):
            return self.rgb_hex(token_or_hex)
        return self.rgb(token_or_hex)

    # ---- slide construction --------------------------------------------

    def add_slide(self, layout_name: str):
        if layout_name not in self.layouts:
            logger.warning("Layout %r missing; falling back to 'Title Only'.", layout_name)
            layout_name = "Title Only"
        return self.prs.slides.add_slide(self.layouts[layout_name])

    def save(self, path: str) -> None:
        # Stamp the document author so generated decks are attributed to the GPT,
        # not the template's creator (whose name the template carries through to
        # every output). Overwrite both creator and last-modified-by. PowerPoint
        # will update last-modified-by to the human once they open and save it.
        try:
            cp = self.prs.core_properties
            cp.author = DOC_AUTHOR
            cp.last_modified_by = DOC_AUTHOR
        except Exception as exc:                          # pragma: no cover
            logger.debug("Could not set document author: %s", exc)
        # Save normally, then run a post-pass over the resulting ZIP to
        # patch any bare ``<a:endParaRPr/>`` in any XML part. PowerPoint's
        # schema validator rejects the bare form (it requires ``lang``)
        # and surfaces the file as needing repair on open. Going through
        # the ZIP catches malformations in handout / notes masters and
        # other parts that python-pptx doesn't expose as element trees.
        self.prs.save(path)
        self._postprocess_pptx_for_pptx_validator(path)
        # Single enforcement point for the chart rules (no tick marks ever; drop
        # the value-axis scale when bars are data-labelled). Runs on every save so
        # it covers every chart-building path, not just generate_deck.
        normalize_chart_parts(path)

    @staticmethod
    def _postprocess_pptx_for_pptx_validator(path: str) -> None:
        """Patch any bare ``<a:endParaRPr/>`` to ``<a:endParaRPr lang="en-US"/>``
        in every .xml file inside the saved ``.pptx``. Idempotent and
        purely string-replace based — no XML parsing overhead since the
        replacement is unambiguous (the bare form has no attributes)."""
        import shutil
        import tempfile
        import zipfile
        bare = b"<a:endParaRPr/>"
        fixed = b'<a:endParaRPr lang="en-US"/>'
        # Same with namespace-prefix variants seen in some files.
        bare2 = b"<a:endParaRPr />"

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
            tmp_path = tmp.name
        try:
            with zipfile.ZipFile(path, "r") as zin, \
                 zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename.endswith(".xml") and (
                        bare in data or bare2 in data
                    ):
                        data = data.replace(bare, fixed).replace(bare2, fixed)
                    zout.writestr(item, data)
            shutil.move(tmp_path, path)
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    # ---- placeholder discovery -----------------------------------------

    @staticmethod
    def placeholder_type(shape) -> Optional[PP_PLACEHOLDER]:
        try:
            return shape.placeholder_format.type
        except Exception:
            return None

    @staticmethod
    def is_title(shape, slide) -> bool:
        return shape is slide.shapes.title or TemplateRuntime.placeholder_type(shape) == PP_PLACEHOLDER.TITLE

    def content_placeholders(self, slide) -> List:
        """Non-title content placeholders, sorted (top, left). Inherited
        geometry is None-safe via or-zero coercion."""
        out = []
        for shape in slide.placeholders:
            if self.is_title(shape, slide):
                continue
            t = self.placeholder_type(shape)
            if t in {PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT,
                     PP_PLACEHOLDER.PICTURE, PP_PLACEHOLDER.TABLE,
                     PP_PLACEHOLDER.CHART, PP_PLACEHOLDER.SLIDE_IMAGE,
                     PP_PLACEHOLDER.SUBTITLE}:
                out.append(shape)
        out.sort(key=lambda s: (s.top or 0, s.left or 0))
        return out

    # ---- text fillers --------------------------------------------------

    def set_title(self, slide, text: str) -> None:
        if slide.shapes.title:
            slide.shapes.title.text = to_sentence_case(text)
            try:
                slide.shapes.title.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            except Exception:
                pass

    @staticmethod
    def suppress_bullets(text_frame) -> None:
        """Override inherited body bullets with explicit <a:buNone/>."""
        for p in text_frame.paragraphs:
            p_el = p._p  # noqa: SLF001
            pPr = p_el.find(f"{{{A_NS}}}pPr")
            if pPr is None:
                pPr = etree.SubElement(p_el, f"{{{A_NS}}}pPr")
                p_el.insert(0, pPr)
            for ch in list(pPr):
                if etree.QName(ch).localname in {"buChar", "buAutoNum", "buNone", "buFont"}:
                    pPr.remove(ch)
            etree.SubElement(pPr, f"{{{A_NS}}}buNone")
            pPr.set("marL", "0"); pPr.set("indent", "0")

    def fill_bullets(self, placeholder, heading: str, bullets: Iterable,
                      level: int = 1) -> None:
        """Fill a content placeholder with a bold heading + a (possibly nested)
        bulleted list.

        Each item may be a plain string (a bullet at ``level``) or a dict
        ``{"text": "...", "bullets": [...]}`` whose nested ``bullets`` render one
        indent level deeper. The OW master defines per-level styles in
        <p:bodyStyle> — level 4 (Heading) is bold; levels 1/2/3 carry the
        '•' / '–' / '-' glyphs and hanging indents — so we only set ``<a:p
        lvl="N">`` and the master cascades size/font/colour/bullet for free.
        Depth is capped at level 3 (the deepest glyph the master defines).
        """
        tf = placeholder.text_frame
        tf.clear(); tf.word_wrap = True
        used_first = False
        if heading:
            tf.paragraphs[0].text = heading
            tf.paragraphs[0].level = 4   # Heading style (12pt, bold)
            used_first = True

        for lvl, text in iter_bullets(bullets, level):
            p = tf.paragraphs[0] if not used_first else tf.add_paragraph()
            used_first = True
            p.text = text
            p.level = min(lvl, 3)

    def fill_prose(self, placeholder, heading: str,
                   paragraphs: Iterable[str]) -> None:
        """Fill a content placeholder with an optional bold heading followed by
        editorial prose paragraphs: complete sentences, no bullet glyphs, flush
        left, with a little space between paragraphs. Used by the editorial
        writing style so body copy reads like a short report section rather than
        a bulleted list. Size/font come from the master's body level; bullets
        are suppressed and 6pt space-before separates paragraphs.
        """
        tf = placeholder.text_frame
        tf.clear(); tf.word_wrap = True
        paras = [str(p) for p in paragraphs if str(p).strip()]
        first = tf.paragraphs[0]
        if heading:
            first.text = heading
            first.level = 4            # bold heading per OW master
            for body in paras:
                p = tf.add_paragraph(); p.text = body; p.level = 1
        elif paras:
            first.text = paras[0]; first.level = 1
            for body in paras[1:]:
                p = tf.add_paragraph(); p.text = body; p.level = 1
        # Drop glyphs + hanging indent so paragraphs sit flush left.
        self.suppress_bullets(tf)
        # Add space-before to every paragraph after the first for readability.
        # Insert spcBef at index 0 of pPr so it precedes <a:buNone/> (schema
        # order: spcBef before bullet elements) and never trips the validator.
        for idx, p in enumerate(tf.paragraphs):
            if idx == 0:
                continue
            pPr = p._p.find(f"{{{A_NS}}}pPr")  # noqa: SLF001
            if pPr is None:
                pPr = etree.SubElement(p._p, f"{{{A_NS}}}pPr")
                p._p.insert(0, pPr)
            for ch in list(pPr):
                if etree.QName(ch).localname == "spcBef":
                    pPr.remove(ch)
            spc = pPr.makeelement(f"{{{A_NS}}}spcBef", {})
            pts = etree.SubElement(spc, f"{{{A_NS}}}spcPts")
            pts.set("val", "600")      # 6pt
            pPr.insert(0, spc)

    def fill_plain_text(self, placeholder, text: str, *,
                         level: Optional[int] = None,
                         size: Optional[int] = None) -> None:
        """Plain text in a content placeholder. If ``level`` is set, defer
        to the master's per-level style (size, font, weight). Otherwise
        the master's default level applies.

        Note: the ``bold`` parameter is intentionally removed. Bold should
        come from the level styling, not from caller-side overrides — this
        prevents accidental "primary-font + bold" violations.
        """
        tf = placeholder.text_frame
        tf.clear(); tf.word_wrap = True
        tf.paragraphs[0].text = text
        if level is not None:
            tf.paragraphs[0].level = level
        if size:
            for run in tf.paragraphs[0].runs:
                run.font.size = Pt(size)
        self.suppress_bullets(tf)

    # ---- text-level helpers (per BRAND_SPEC) -------------------------

    @staticmethod
    def _set_font_kind(rpr, font_kind: str) -> None:
        """Point a run's latin/ea/cs typeface at the theme major (+mj-lt =
        Marsh Serif) or minor (+mn-lt = Noto Sans) font. Single source of truth
        for the theme-font reference used by every free-text-box run."""
        ref = "+mj-lt" if font_kind == "major" else "+mn-lt"
        for tag in ("latin", "ea", "cs"):
            for ex in rpr.findall(f"{{{A_NS}}}{tag}"):
                rpr.remove(ex)
            etree.SubElement(rpr, f"{{{A_NS}}}{tag}", typeface=ref)

    def _box_margins(self, tf) -> None:
        """Set a shape's internal text margins to 0.15" on all four sides — the
        rule for any shape that carries a fill colour or an outline."""
        m = Inches(Tokens.SPACE_SHAPE_MARGIN)
        tf.margin_top = m; tf.margin_bottom = m
        tf.margin_left = m; tf.margin_right = m

    def _apply_arrowhead(self, line) -> None:
        """Apply the standard OW arrowhead (arrow style 2 — a filled triangle) to
        a connector's tail end. One definition, used by every arrow."""
        sp_pr = line._element.spPr
        ln_el = sp_pr.find(f"{{{A_NS}}}ln")
        if ln_el is None:
            ln_el = etree.SubElement(sp_pr, f"{{{A_NS}}}ln")
        for existing in ln_el.findall(f"{{{A_NS}}}tailEnd"):
            ln_el.remove(existing)
        tail = etree.SubElement(ln_el, f"{{{A_NS}}}tailEnd")
        tail.set("type", "triangle"); tail.set("w", "med"); tail.set("len", "med")

    def set_font_color(self, font, token: str = "dk1") -> None:
        """Set a run's font colour. Navy (Text 1 / #000F47) is applied as a THEME
        LINK (MSO_THEME_COLOR.TEXT_1) so it follows the template theme rather than
        being baked in as a literal RGB; any other token resolves to its RGB."""
        try:
            is_text1 = str(self.rgb(token)).upper() == DEFAULT_THEME["dk1"]
        except Exception:
            is_text1 = False
        if is_text1:
            font.color.theme_color = MSO_THEME_COLOR.TEXT_1
        else:
            font.color.rgb = self.rgb(token)

    def apply_text_level(self, paragraph, level: int, *,
                          color_token: Optional[str] = None) -> None:
        """Apply a ``BULLET_LEVELS`` entry to a paragraph that lives outside
        a content placeholder (free text box, table cell, shape).

        Sets size, font (major vs minor), bold, and bullet glyph + indent
        explicitly, since the master's lstStyle only cascades into
        placeholders. Use ``color_token`` to override the default text
        color (defaults to ``dk1``).

        Enforces ``BRAND_RULES["primary_font_never_bold"]``: if the level's
        font is the major Latin font, bold is forced False regardless of
        what the level table claims.
        """
        spec = BULLET_LEVELS.get(level)
        if spec is None:
            return
        is_primary_font = (spec.font_kind == "major")
        bold = spec.bold and not (is_primary_font and BRAND_RULES["primary_font_never_bold"])

        p_el = paragraph._p  # noqa: SLF001
        pPr = p_el.find(f"{{{A_NS}}}pPr")
        if pPr is None:
            pPr = etree.SubElement(p_el, f"{{{A_NS}}}pPr")
            p_el.insert(0, pPr)
        # Indent / margin for hanging bullets
        if spec.mar_l_in or spec.indent_in:
            pPr.set("marL",   str(int(spec.mar_l_in   * EMU_PER_INCH)))
            pPr.set("indent", str(int(spec.indent_in * EMU_PER_INCH)))
        if spec.space_before:
            paragraph.space_before = Pt(spec.space_before)
        if spec.space_after:
            paragraph.space_after = Pt(spec.space_after)
        # Bullet glyph
        for ch in list(pPr):
            if etree.QName(ch).localname in {"buChar", "buAutoNum",
                                              "buNone", "buFont"}:
                pPr.remove(ch)
        if spec.bullet_char:
            buFont = etree.SubElement(pPr, f"{{{A_NS}}}buFont")
            buFont.set("typeface", "Arial")
            buChar = etree.SubElement(pPr, f"{{{A_NS}}}buChar")
            buChar.set("char", spec.bullet_char)
        else:
            etree.SubElement(pPr, f"{{{A_NS}}}buNone")

        # Run-level styling
        for run in paragraph.runs:
            run.font.size = Pt(spec.size_pt)
            run.font.bold = bold
            self.set_font_color(run.font, color_token or "dk1")
            # Force the right font kind via theme reference
            rpr = run._r.find(f"{{{A_NS}}}rPr")  # noqa: SLF001
            if rpr is None:
                rpr = etree.SubElement(run._r, f"{{{A_NS}}}rPr")
                run._r.insert(0, rpr)
            for tag in ("latin", "ea", "cs"):
                for existing in rpr.findall(f"{{{A_NS}}}{tag}"):
                    rpr.remove(existing)
            self._set_font_kind(rpr, spec.font_kind)

    def _set_run_font(self, run, kind: str = "minor", size_pt: float = 12,
                      color_token: str = "dk1", bold: bool = False) -> None:
        """Style a run in a free text box with the theme major/minor font at an
        arbitrary size. ``kind='major'`` uses Marsh Serif (forced non-bold per
        the brand rule); ``kind='minor'`` uses Noto Sans."""
        if kind == "major":
            bold = False  # primary font is never bold
        run.font.size = Pt(size_pt)
        run.font.bold = bold
        self.set_font_color(run.font, color_token)
        rpr = run._r.find(f"{{{A_NS}}}rPr")  # noqa: SLF001
        if rpr is None:
            rpr = etree.SubElement(run._r, f"{{{A_NS}}}rPr"); run._r.insert(0, rpr)
        self._set_font_kind(rpr, kind)

    def _run(self, paragraph, text, *, size=None, color="dk1",
             bold=False, kind="minor"):
        """Add a run to ``paragraph`` and style it in one call: size (defaults to
        ``Tokens.TYPE_BODY``), colour token, bold, and theme font kind. Returns
        the run. Collapses the add_run + size + colour + bold idiom."""
        run = paragraph.add_run()
        run.text = "" if text is None else str(text)
        self._set_run_font(run, kind=kind,
                           size_pt=(Tokens.TYPE_BODY if size is None else size),
                           color_token=color, bold=bold)
        return run

    # ---- placeholder card insets --------------------------------------
    #
    # Filled callout placeholders need explicit body insets — the master
    # default (~0.1") makes the text crash into the placeholder edges and
    # into the chart legend on the adjacent placeholder. The reference
    # uses 0.2" / ~5 mm on all four sides.
    CARD_INSET_EMU = 183600   # 0.2 inch = 5.08 mm

    @classmethod
    def _set_body_insets(cls, text_frame,
                         l: Optional[int] = None, t: Optional[int] = None,
                         r: Optional[int] = None, b: Optional[int] = None) -> None:
        """Force explicit l/t/r/b insets on a placeholder's text frame.

        Without this, PowerPoint inherits the master's default (~0.1") and
        the text brushes against the adjacent chart's legend. Used by all
        ``_style_card`` callers.
        """
        l = cls.CARD_INSET_EMU if l is None else l
        t = cls.CARD_INSET_EMU if t is None else t
        r = cls.CARD_INSET_EMU if r is None else r
        b = cls.CARD_INSET_EMU if b is None else b
        body_pr = text_frame._txBody.find(f"{{{A_NS}}}bodyPr")  # noqa: SLF001
        if body_pr is None:
            return
        body_pr.set("lIns", str(l)); body_pr.set("tIns", str(t))
        body_pr.set("rIns", str(r)); body_pr.set("bIns", str(b))

    def fill_kpi(self, placeholder, label: str, value: str, delta: str = "") -> None:
        """Fill a placeholder as a KPI card (cream fill, dark text).

        Layout
        ------
            label       → level 5 (Subheading; minor font, not bold)
                          + explicit 11pt to match the eyebrow size
            value       → level 6 (Small primary; major font, not bold)
                          + explicit 34pt for display weight
                          (level 6's master size is 26pt; KPI cards want
                          bigger for at-a-glance reading)
            delta       → level 5 (Subheading; minor font, not bold)
                          + explicit 12pt

        The major font on the value run is *never bold* per
        ``BRAND_RULES["primary_font_never_bold"]`` — the visual weight comes
        from size and serif glyph shape, not from a bold weight.
        """
        self._style_card(placeholder)   # cream fill, no border
        tf = placeholder.text_frame
        tf.clear(); tf.word_wrap = True
        self._set_body_insets(tf)

        # ---- label (lvl 5: Subheading, 12pt minor, not bold) ----
        p1 = tf.paragraphs[0]; p1.level = 5
        r1 = p1.add_run(); r1.text = label
        r1.font.size = Pt(Tokens.TYPE_BODY)
        r1.font.bold = False
        p1.space_after = Pt(Tokens.PARA_PROSE)        # 6pt below the label

        # ---- value (lvl 6: Small primary, 26pt major, not bold) ----
        p2 = tf.add_paragraph(); p2.level = 6
        r2 = p2.add_run(); r2.text = value
        r2.font.size = Pt(Tokens.TYPE_KPI)            # explicit display size
        # NO bold (brand rule); NO color override (inherits dk1 from level)

        # ---- delta (lvl 5: Subheading, 12pt minor) ----
        if delta:
            p3 = tf.add_paragraph(); p3.level = 5
            p3.space_before = Pt(Tokens.PARA_PROSE)   # 6pt above the body text
            r3 = p3.add_run(); r3.text = delta
            r3.font.size = Pt(Tokens.TYPE_BODY)
            r3.font.bold = False

        self.suppress_bullets(tf)

    # --- card-grid drawing ---------------------------------------------
    #
    # When KPI cards (or any other "card-shaped" content) don't fit in
    # the available placeholder layouts — e.g. 8 metrics on a 4-column
    # layout — the OW preference is NOT to paginate across two slides
    # with "(continued)" headers, but to draw all cards as real shapes
    # on a Title Only layout (see the 8_cards.pptx reference, slide 3).
    #
    # `draw_card_grid` produces a visually-identical grid by drawing
    # rectangles with the same fill / insets / typography as `fill_kpi`,
    # but as free shapes rather than placeholder mutations. The shape
    # text frame doesn't inherit master typography, so we set every
    # paragraph's font kind / size / color explicitly via `apply_text_level`.

    @staticmethod
    def _grid_candidates(n: int):
        """Candidate (rows, cols) grids for ``n`` cards, in preference order:
        the aesthetic default first (wider, fewer rows), then progressively
        TALLER fallbacks (more rows → more height per card). A fit search walks
        this list and takes the first layout where the content fits, so long-text
        cards drop from a 1×4 strip to a 2×2 (double the height) automatically."""
        table = {
            1: [(1, 1)],
            2: [(1, 2), (2, 1)],
            3: [(1, 3), (3, 1)],
            4: [(1, 4), (2, 2), (4, 1)],
            5: [(2, 3), (1, 5)],
            6: [(2, 3), (3, 2)],
            7: [(2, 4), (3, 3)],
            8: [(2, 4), (4, 2)],
        }
        if n in table:
            return table[n]
        if n <= 12:
            return [(3, 4), (4, 3)]
        return [(4, 4)]

    def _card_content_height(self, item: Dict[str, Any], cw: float) -> float:
        """Estimate the vertical inches a card's CONTENT needs at width ``cw``
        (text wraps inside the card insets). Used by ``_pick_card_grid`` to choose
        a grid where the cards fit instead of one fixed by count alone."""
        inset = self.INSET_IN
        tw = max(0.3, cw - 2 * inset)
        h = 2 * inset
        cpl = max(8, int(tw / 0.085))      # ~chars per line fallback (textmetrics used when present)

        def add(text, size_pt):
            nonlocal h
            if text is None or str(text) == "":
                return
            n_lines = _wrapped_line_count(str(text), size_pt=size_pt, kind="minor",
                                          width_in=tw, chars_per_line=cpl)
            h += n_lines * (size_pt * 1.2 / 72.0) + 0.04   # line height + para gap

        add(item.get("label"), Tokens.TYPE_BODY)
        if item.get("value"):
            h += Tokens.TYPE_KPI * 1.2 / 72.0 + 0.04        # big KPI value line
        add(item.get("delta"), Tokens.TYPE_BODY)
        add(item.get("heading"), Tokens.TYPE_HEADING)
        add(item.get("subheading"), Tokens.TYPE_BODY)
        body = item.get("body")
        if isinstance(body, dict):
            for b in (body.get("bullets") or []):
                add(b, Tokens.TYPE_BODY)
            for pgr in (body.get("paragraphs") or []):
                add(pgr, Tokens.TYPE_BODY)
        for b in (item.get("bullets") or []):
            add(b, Tokens.TYPE_BODY)
        if isinstance(item.get("chart"), dict) and item.get("chart"):
            h += 1.0                                         # minimum readable chart band
        return h

    def _pick_card_grid(self, n: int, items=None, area=None, gutter=None):
        """Pick (rows, cols) for ``n`` cards. Fit-aware when ``items`` + ``area``
        are given: it walks the candidate grids (``_grid_candidates``) from the
        preferred wide layout to taller fallbacks and returns the FIRST where
        every card's measured content fits its cell. If none fit cleanly it
        returns the candidate with the most slack (least overrun). With no
        content/area it falls back to the count-based default."""
        cands = self._grid_candidates(n)
        if n <= 0:
            return (0, 0)
        if not items or area is None:
            return cands[0]
        g = gutter if gutter is not None else GRID_GAP
        best = None
        for (r, c) in cands:
            cw = (area.w - (c - 1) * g) / c
            ch = (area.h - (r - 1) * g) / r
            if cw <= 0 or ch <= 0:
                continue
            worst_need = max((self._card_content_height(it, cw)
                              for it in items[:r * c]), default=0.0)
            slack = ch - worst_need
            if best is None or slack > best[1]:
                best = ((r, c), slack)
            if slack >= 0:                       # all cards fit in this candidate
                return (r, c)
        return best[0] if best else cands[0]

    @staticmethod
    def _parse_grid_spec(spec: Any) -> Optional[Tuple[int, int]]:
        """Parse a grid override and return ``(rows, cols)``.

        **Convention: the string form is ``"ROWSxCOLS"`` (matrix/numpy).**
        This matches how every language describes 2-D arrays and how the
        runtime stores grids internally. So "1x4" = 1 row of 4 cells
        (horizontal strip), "2x3" = 2 rows × 3 cols (wide rectangle).
        Tuple/list inputs are taken as ``(rows, cols)`` at face value.

        Accepts:
          - ``"1x4"`` → 1 row × 4 cols → ``(1, 4)`` (horizontal strip)
          - ``"2x3"`` → 2 rows × 3 cols → ``(2, 3)`` (wide grid)
          - ``"3x2"`` → 3 rows × 2 cols → ``(3, 2)`` (tall grid)
          - ``(rows, cols)`` tuples or lists
          - ``None`` → returns ``None`` (caller falls back to auto-pick)

        Returns ``None`` on unparsable input, with a warning logged.
        """
        if spec is None:
            return None
        if isinstance(spec, (tuple, list)) and len(spec) == 2:
            try:
                r, c = int(spec[0]), int(spec[1])
            except (TypeError, ValueError):
                logger.warning("Invalid grid spec %r — falling back to auto.", spec)
                return None
            return (r, c) if r > 0 and c > 0 else None
        if isinstance(spec, str):
            s = spec.lower().replace(" ", "")
            if "x" in s:
                a, b = s.split("x", 1)
                try:
                    rows, cols = int(a), int(b)   # ROWSxCOLS convention
                except ValueError:
                    logger.warning("Invalid grid spec %r — falling back to auto.", spec)
                    return None
                return (rows, cols) if rows > 0 and cols > 0 else None
        logger.warning("Unrecognized grid spec %r — falling back to auto.", spec)
        return None

    def draw_card_grid(self, slide, items: Sequence[Dict[str, Any]],
                       *, area: Optional[ContentArea] = None,
                       gutter: float = Tokens.SPACE_GAP_LG,
                       grid: Any = None) -> None:
        """Draw a grid of cards on a Title Only slide.

        Each item is a dict; the cell's rendering depends on what's in it:

          - ``label`` / ``value`` / ``delta`` only → KPI card (cream fill,
            stacked text, no border) — same look as ``fill_kpi``
          - any of those plus ``chart`` → composition cell: text on top,
            native chart at column width along the bottom
          - ``chart`` only (no text fields) → chart-only cell

        Parameters
        ----------
        slide
            Title Only slide to draw on (the compiler picks the layout;
            the runtime doesn't validate it).
        items
            Sequence of cell dicts (see above for accepted keys).
        area
            Content area to fill. Defaults to the runtime's current
            content area (already shrunk for footnote/conclusion).
        gutter
            Spacing between cards in inches. Default 0.5" — matches the
            8_cards.pptx reference.
        grid
            Optional explicit grid override:
              - ``"2x2"`` / ``"3x4"`` (string)
              - ``(rows, cols)`` (tuple)
              - ``None`` → auto-pick from item count via ``_pick_card_grid``
            Useful when authors want a specific shape (e.g. a 1x4 row of
            charts vs the auto-picked 2x2). Excess items beyond the grid
            capacity are silently dropped.
        """
        n = len(items)
        if n == 0:
            return
        a = area or self.content_area()
        rows, cols = self._parse_grid_spec(grid) or self._pick_card_grid(n, items, a, gutter)
        card_w = (a.w - (cols - 1) * gutter) / cols
        card_h = (a.h - (rows - 1) * gutter) / rows
        if card_w <= 0 or card_h <= 0:
            logger.warning("draw_card_grid: card area too small (w=%.2f h=%.2f); skipping.",
                            card_w, card_h)
            return

        for i, item in enumerate(items):
            if i >= rows * cols:
                break
            r, c = divmod(i, cols)
            cx = a.x + c * (card_w + gutter)
            cy = a.y + r * (card_h + gutter)
            self._draw_grid_cell(slide, cx, cy, card_w, card_h, item)

    # ---- grid-cell renderers (text-only / composition / chart-only) ---

    # Vertical space the text region claims when the cell hosts both text
    # and a chart. ~1.2" is enough for label (11pt) + value (28pt) +
    # optional delta (12pt) at the inset-conscious sizes used in the
    # composition layout. Used by ``_draw_composition_cell``.
    COMPOSITION_TEXT_H = 1.2     # inches
    COMPOSITION_GAP    = 0.10    # inches between text region and chart
    INSET_IN           = 0.20    # inches — matches CARD_INSET_EMU

    def _draw_grid_cell(self, slide, cx: float, cy: float,
                        cw: float, ch: float, item: Dict[str, Any]) -> None:
        """Dispatch a single grid cell to the right renderer based on
        what's in the item dict."""
        has_text = bool(item.get("label") or item.get("value") or item.get("delta"))
        chart_spec = item.get("chart")
        has_chart = isinstance(chart_spec, dict) and bool(chart_spec)

        if has_chart and has_text:
            self._draw_composition_cell(slide, cx, cy, cw, ch, item)
        elif has_chart:
            self._draw_chart_only_cell(slide, cx, cy, cw, ch, item)
        else:
            self._draw_kpi_cell(slide, cx, cy, cw, ch, item)

    def _draw_kpi_cell(self, slide, cx: float, cy: float,
                       cw: float, ch: float, item: Dict[str, Any]) -> None:
        """Cream-filled card with stacked label / value / delta — the v3.3
        KPI layout. Used when the cell has no chart. Optional ``icon`` on
        the item renders as a small navy icon at the top-left inset; the
        label/value/delta text shifts down to make room."""
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(cx), Inches(cy), Inches(cw), Inches(ch),
        )
        self._style_card(shape)
        tf = shape.text_frame
        tf.clear(); tf.word_wrap = True
        self._set_body_insets(tf)
        body_pr = tf._txBody.find(f"{{{A_NS}}}bodyPr")  # noqa: SLF001
        if body_pr is not None and body_pr.get("anchor") is not None:
            del body_pr.attrib["anchor"]

        # Optional icon at the card's top-left. Rendered before text so the
        # text-frame margin can shift the text down and clear the icon.
        icon_name = item.get("icon") if isinstance(item, dict) else None
        if icon_name:
            try:
                import icons as _ic
                icon_side = min(Tokens.SPACE_ICON, cw * 0.32, ch * 0.32)
                icon_inset = Tokens.SPACE_ICON_INSET
                icon_color = str(item.get("icon_color", "000F47")).lstrip("#")
                _ic.draw_icon(slide, str(icon_name),
                              cx + icon_inset, cy + icon_inset,
                              icon_side, icon_side, color_hex=icon_color)
                # Push label/value/delta below the icon.
                tf.margin_top = Inches(icon_inset + icon_side + 0.08)
            except Exception as exc:
                logger.warning("KPI icon %r skipped: %s", icon_name, exc)

        label = str(item.get("label", ""))
        value = str(item.get("value", ""))
        delta = str(item.get("delta", "")).strip()

        p1 = tf.paragraphs[0]
        if label:
            r1 = p1.add_run(); r1.text = label
            self.apply_text_level(p1, 5)
            for run in p1.runs:
                run.font.size = Pt(Tokens.TYPE_BODY)
            p1.alignment = Align.LEFT

        if value:
            p2 = tf.add_paragraph()
            r2 = p2.add_run(); r2.text = value
            self.apply_text_level(p2, 6)
            for run in p2.runs:
                run.font.size = Pt(Tokens.TYPE_KPI)
            p2.alignment = Align.LEFT

        if delta:
            p3 = tf.add_paragraph()
            r3 = p3.add_run(); r3.text = delta
            self.apply_text_level(p3, 5)
            p3.alignment = Align.LEFT

        self.suppress_bullets(tf)

    def _draw_composition_cell(self, slide, cx: float, cy: float,
                               cw: float, ch: float, item: Dict[str, Any]) -> None:
        """Composition cell: text region on top, native chart on bottom.

        Layout
        ------
            +----------------------+  ← cy
            |  LABEL  (lvl 5, 11pt)|     ↑ COMPOSITION_TEXT_H (~1.2")
            |  Value  (lvl 6, 28pt)|     │
            |  delta  (lvl 5, 12pt)|     ↓
            |  ─── (gap)           |
            |                      |
            |   chart fills        |     chart_h = ch − inset×2 − text_h − gap
            |   the remainder      |
            |                      |
            +----------------------+  ← cy + ch

        Chart width = cell width − 2×inset (matches the column width minus
        breathing room). The chart's plot area is transparent; if the cell
        has a card background the cream shows through, which is by design.
        """
        # Card background (cream fill, no border) — same as the KPI cell.
        bg = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(cx), Inches(cy), Inches(cw), Inches(ch),
        )
        self._style_card(bg)
        # The background shape's text frame is unused — clear it so a stray
        # default paragraph doesn't render on top of our text box.
        bg.text_frame.text = ""

        inset = self.INSET_IN
        text_x = cx + inset
        text_y = cy + inset
        text_w = cw - 2 * inset
        text_h = self.COMPOSITION_TEXT_H

        # ---- text region ----
        # Free text box (not the rectangle's own text frame) so it stays
        # sized to the text region and doesn't claim the whole card.
        tb = slide.shapes.add_textbox(
            Inches(text_x), Inches(text_y), Inches(text_w), Inches(text_h),
        )
        tf = tb.text_frame
        tf.word_wrap = True
        # Tight insets — the card already has 0.2" inset; the text box adds
        # only token padding so it sits flush.
        self._set_body_insets(tf, l=0, t=0, r=0, b=0)

        label = str(item.get("label", ""))
        value = str(item.get("value", ""))
        delta = str(item.get("delta", "")).strip()

        first_p = True
        if label:
            p = tf.paragraphs[0] if first_p else tf.add_paragraph()
            first_p = False
            r = p.add_run(); r.text = label
            self.apply_text_level(p, 5)
            for run in p.runs:
                run.font.size = Pt(Tokens.TYPE_BODY)
            p.alignment = Align.LEFT

        if value:
            p = tf.paragraphs[0] if first_p else tf.add_paragraph()
            first_p = False
            r = p.add_run(); r.text = value
            self.apply_text_level(p, 6)
            for run in p.runs:
                # KPI value size (TYPE_KPI_COMPACT) — held equal to the
                # standalone KPI (26pt) so chart-sharing cells match the
                # sample reference; split the token if they must diverge.
                run.font.size = Pt(Tokens.TYPE_KPI_COMPACT)
            p.alignment = Align.LEFT

        if delta:
            p = tf.paragraphs[0] if first_p else tf.add_paragraph()
            first_p = False
            r = p.add_run(); r.text = delta
            self.apply_text_level(p, 5)
            p.alignment = Align.LEFT

        self.suppress_bullets(tf)

        # ---- chart region ----
        chart_x = text_x
        chart_y = text_y + text_h + self.COMPOSITION_GAP
        chart_w = text_w
        chart_h = ch - 2 * inset - text_h - self.COMPOSITION_GAP
        if chart_h < 0.5:
            logger.warning("Composition cell too short for a chart (h=%.2f); skipping chart.", chart_h)
            return

        self._add_floating_chart(
            slide, item["chart"],
            cx=chart_x, cy=chart_y, cw=chart_w, ch=chart_h,
            in_cell=True,
        )

    def _draw_chart_caption(self, slide, cx: float, cy: float, cw: float,
                            heading: str, subheading: str = "") -> float:
        """Draw a chart's heading (+ optional subheading) as a TEXT SHAPE above
        the chart, in the house Heading/Subheading style (level 4 bold / level 5
        regular) — never an embedded chart title. Returns the height consumed so
        the caller can shrink the chart area beneath it."""
        if not heading and not subheading:
            return 0.0
        sub_h = 0.22 if subheading else 0.0
        h = Tokens.SPACE_HEADING_H + sub_h
        tb = slide.shapes.add_textbox(Inches(cx), Inches(cy), Inches(cw), Inches(h))
        tf = tb.text_frame
        tf.clear(); tf.word_wrap = True
        tf.margin_left = tf.margin_right = 0
        tf.margin_top = tf.margin_bottom = 0
        if heading:
            self._run(tf.paragraphs[0], heading,
                      size=Tokens.TYPE_HEADING, bold=True)        # Heading (lvl 4)
        if subheading:
            p = tf.add_paragraph() if heading else tf.paragraphs[0]
            self._run(p, subheading, size=Tokens.TYPE_BODY, bold=False)  # Subheading (lvl 5)
        self.suppress_bullets(tf)
        return h + Tokens.SPACE_GAP

    def _draw_chart_only_cell(self, slide, cx: float, cy: float,
                              cw: float, ch: float, item: Dict[str, Any]) -> None:
        """Cell holding a chart, optionally with a heading/subheading text shape
        above it (house Heading/Subheading style — not an embedded chart title).
        The chart fills the remaining cell with the standard inset."""
        inset = self.INSET_IN
        heading = item.get("heading") or item.get("title") or item.get("caption") or ""
        subheading = item.get("subheading") or ""
        cap_h = self._draw_chart_caption(
            slide, cx + inset, cy + inset, cw - 2 * inset, heading, subheading)
        self._add_floating_chart(
            slide, item["chart"],
            cx=cx + inset, cy=cy + inset + cap_h,
            cw=cw - 2 * inset, ch=ch - 2 * inset - cap_h,
            in_cell=True,
        )

    def _add_floating_chart(self, slide, chart_spec: Dict[str, Any],
                            *, cx: float, cy: float, cw: float, ch: float,
                            in_cell: bool = False) -> None:
        """Add a native chart at the given geometry and apply OW theming.

        ``in_cell=True`` tightens the chart for small contexts: legend off
        (the cell's label provides the legend), data labels off (they
        clutter small plots, and at narrow widths the labels wrap to two
        lines because they don't fit). For full-slide charts this method
        is bypassed in favor of ``insert_chart_floating``.
        """
        chart_type = chart_spec.get("type", "column_clustered")
        xl_type = resolve_chart_type(chart_type)
        categories = chart_spec.get("categories", [])
        series_in = chart_spec.get("series", [])
        normalized = []
        for s in series_in:
            if "points" in s:
                normalized.append((s.get("name", ""),
                                    [(float(p[0]), float(p[1])) for p in s["points"]]))
            else:
                normalized.append((s.get("name", ""), list(s.get("values", []))))

        chart_data = self._build_chart_data(xl_type, categories, normalized)
        chart_shape = slide.shapes.add_chart(
            xl_type,
            Inches(cx), Inches(cy), Inches(cw), Inches(ch),
            chart_data,
        )
        chart = chart_shape.chart
        self._theme_chart(chart, xl_type)

        if in_cell:
            # Legend: keep it for MULTI-series charts (the reader needs to know
            # which colour is which series). For a single series it is redundant —
            # the caption carries the context — so strip it. Done via chartSpace
            # XML directly; python-pptx's ``has_legend`` setter is unreliable
            # after ``_theme_chart`` mutates the legend element.
            chart_el = chart._chartSpace  # noqa: SLF001
            if len(normalized) <= 1:
                for legend in chart_el.findall(f".//{self._CNS}legend"):
                    legend.getparent().remove(legend)

            # Keep DATA LABELS on — they put the value on each bar/point, which is
            # the information a dashboard chart must carry. (Previously stripped to
            # reduce clutter; that left charts reading as floating bars with no
            # numbers. The value on the mark is worth the density.)

            # Hide value-axis tick labels — with data labels on the marks, the
            # axis ranges (e.g. "795 800 805 810") are redundant noise. Keep
            # category-axis labels (those carry context like "Q1, Q2..").
            for valAx in chart_el.findall(f".//{self._CNS}valAx"):
                tickLblPos = valAx.find(f"{self._CNS}tickLblPos")
                if tickLblPos is None:
                    tickLblPos = etree.SubElement(valAx, f"{self._CNS}tickLblPos")
                tickLblPos.set("val", "none")

    def fill_insight(self, placeholder, title: str, text: str) -> None:
        """Fill a callout placeholder with eyebrow label + body line.

        Layout
        ------
            title (eyebrow) → level 5 (Subheading; minor font, not bold)
                              + explicit 11pt
            text  (body)    → level 6 (Small primary; major font, not bold)
                              — size and font come from the master via
                              level inheritance; no run-level overrides

        The body line is rendered in the major (primary) Latin font at
        ``BRAND_RULES["insight_body_pt"]`` (18pt) — the "strong short statement"
        treatment from the brand spec, a pull-quote feel without bold weight (the
        major font is never bold per ``BRAND_RULES``). 18pt is also the minimum
        size for the Marsh Serif font (``primary_font_min_pt``). For longer prose
        where the major font would feel too heavy, use the bullet/plain-text
        rendering paths instead of the insight card.
        """
        self._style_card(placeholder)   # cream fill, no border
        tf = placeholder.text_frame
        tf.clear(); tf.word_wrap = True
        self._set_body_insets(tf)

        # ---- eyebrow label (lvl 5) ----
        p1 = tf.paragraphs[0]; p1.level = 5
        r1 = p1.add_run(); r1.text = title.upper()
        r1.font.size = Pt(Tokens.TYPE_BODY)
        r1.font.bold = False

        # ---- body line (lvl 6 → major font; sized to the brand minimum 18pt) ----
        p2 = tf.add_paragraph(); p2.level = 6
        r2 = p2.add_run(); r2.text = text
        r2.font.size = Pt(BRAND_RULES["insight_body_pt"])  # 18pt (Marsh Serif min)
        p2.space_before = Pt(Tokens.PARA_PROSE)
        # No bold (brand rule: primary font never bold)
        # No color override (inherits dk1 from level)

        self.suppress_bullets(tf)

    def _style_card(self, placeholder,
                    fill: str = "default_fill",
                    line: Optional[str] = None) -> None:
        """Style a placeholder as a filled callout card.

        Parameters
        ----------
        fill
            Either a theme token (``"accent1"``) or a 6-digit hex string,
            or one of the ``BRAND_COLORS`` keys (``"default_fill"``,
            ``"strong_highlight"``, etc.). Default ``"default_fill"`` is
            the warm cream (``F7F3EE``) used for callout cards across the
            deck. Pass ``"strong_highlight"`` (= ``accent1`` navy) on the
            specific card where you want the legacy dark look.
        line
            Same accepted forms as ``fill``. ``None`` (default) means *no
            border* — written as ``<a:ln><a:noFill/></a:ln>``. This matches
            the reference; the card reads as a panel, not a framed box.
        """
        # ---- fill ----
        try:
            placeholder.fill.solid()
            if fill in BRAND_COLORS:
                placeholder.fill.fore_color.rgb = self.resolve_color(BRAND_COLORS[fill])
            else:
                placeholder.fill.fore_color.rgb = self.resolve_color(fill)
        except Exception as exc:
            logger.debug("Could not set card fill (%r): %s", fill, exc)

        # ---- line ----
        if line is None:
            # noFill on the line — matches reference's <a:ln><a:noFill/></a:ln>
            try:
                placeholder.line.fill.background()
            except Exception:
                # Fallback: write the OOXML directly
                spPr = placeholder._element.spPr  # noqa: SLF001
                ln = spPr.find(f"{{{A_NS}}}ln")
                if ln is None:
                    ln = etree.SubElement(spPr, f"{{{A_NS}}}ln")
                for child in list(ln):
                    ln.remove(child)
                etree.SubElement(ln, f"{{{A_NS}}}noFill")
        else:
            try:
                if line in BRAND_COLORS:
                    placeholder.line.color.rgb = self.resolve_color(BRAND_COLORS[line])
                else:
                    placeholder.line.color.rgb = self.resolve_color(line)
            except Exception as exc:
                logger.debug("Could not set card line (%r): %s", line, exc)

    # ---- chart binding -------------------------------------------------

    def insert_chart(self, slide, placeholder, categories, series, chart_type="column_clustered"):
        if any(v is None for v in (placeholder.left, placeholder.top,
                                    placeholder.width, placeholder.height)):
            raise ValueError("Cannot bind chart: placeholder geometry unresolved.")
        xl_type = resolve_chart_type(chart_type)
        data = self._build_chart_data(xl_type, categories, series)
        chart_shape = slide.shapes.add_chart(
            xl_type,
            placeholder.left, placeholder.top, placeholder.width, placeholder.height,
            data,
        )
        self._theme_chart(chart_shape.chart, xl_type)
        self._bind_to_placeholder(chart_shape, placeholder)

    def insert_chart_floating(self, slide, categories, series, chart_type="column_clustered"):
        sw_in = self.prs.slide_width / EMU_PER_INCH
        w = min(FALLBACK_CONTENT.w, sw_in - 2 * FALLBACK_CONTENT.x)
        xl_type = resolve_chart_type(chart_type)
        data = self._build_chart_data(xl_type, categories, series)
        chart_shape = slide.shapes.add_chart(
            xl_type,
            Inches(FALLBACK_CONTENT.x), Inches(FALLBACK_CONTENT.y),
            Inches(w), Inches(FALLBACK_CONTENT.h),
            data,
        )
        self._theme_chart(chart_shape.chart, xl_type)

    def add_waterfall(self, slide, spec, *, cx, cy, cw, ch) -> None:
        """Waterfall (bridge) chart, simulated with a stacked column/bar chart:
        a transparent 'base' series floats the visible delta bars. Indices in
        ``spec['totals']`` (default first + last) are absolute totals drawn from
        zero in navy; other points are deltas — green up / red down for the
        vertical orientation, light-blue for the horizontal one. Data labels show
        each bar's magnitude. No native waterfall type exists in the OOXML the
        library writes, hence the invisible-base technique.
        """
        from pptx.enum.chart import XL_LABEL_POSITION
        horizontal = str(spec.get("orientation", "vertical")).lower().startswith("h")
        cats = [str(c) for c in spec.get("categories", [])]
        vals = [float(v) for v in spec.get("values", [])]
        if not cats or not vals:
            return
        totals = set(spec.get("totals", [0, len(vals) - 1]))
        base, bar, kinds = [], [], []
        running = 0.0
        for i, v in enumerate(vals):
            if i in totals:
                base.append(0.0); bar.append(abs(v)); running = v; kinds.append("total")
            elif v >= 0:
                base.append(running); bar.append(v); running += v; kinds.append("up")
            else:
                running += v; base.append(running); bar.append(-v); kinds.append("down")

        xl = XL_CHART_TYPE.BAR_STACKED if horizontal else XL_CHART_TYPE.COLUMN_STACKED
        data = self._build_chart_data(xl, cats, [("base", base), ("value", bar)])
        chart = slide.shapes.add_chart(xl, Inches(cx), Inches(cy),
                                       Inches(cw), Inches(ch), data).chart
        chart.has_legend = False
        chart.has_title = False

        for ax in (chart.category_axis, chart.value_axis):
            try:
                ax.tick_labels.font.size = Pt(Tokens.TYPE_AXIS)
                self.set_font_color(ax.tick_labels.font, "dk1")
                ax.has_major_gridlines = False
                ax.has_minor_gridlines = False
            except Exception:
                pass
        # Transparent base series.
        bs = chart.series[0]
        bs.format.fill.background()
        try:
            bs.format.line.fill.background()
        except Exception:
            pass

        # Visible series: per-point fill + per-point label colour.
        palette = {"total": DEFAULT_THEME["dk1"],
                   "up":   DEFAULT_THEME["accent2"] if horizontal else CHART_COLORS["up"],
                   "down": DEFAULT_THEME["accent2"] if horizontal else CHART_COLORS["down"]}
        vs = chart.series[1]
        try:
            vs.format.line.fill.background()
        except Exception:
            pass
        vs.data_labels.show_value = True
        vs.data_labels.number_format = f"[$-{self._number_lcid}]#,##0"
        vs.data_labels.number_format_is_linked = False
        try:
            vs.data_labels.position = XL_LABEL_POSITION.CENTER
        except Exception:
            pass
        for i, pt in enumerate(vs.points):
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = self.rgb_hex(palette[kinds[i]])
            # White label on dark bars (navy total, or any vertical delta);
            # navy label on the light-blue horizontal deltas.
            on_dark = kinds[i] == "total" or not horizontal
            try:
                pt.data_label.font.size = Pt(Tokens.TYPE_BODY)
                self.set_font_color(pt.data_label.font, "lt1" if on_dark else "dk1")
            except Exception:
                pass

    def treemap(self, slide, items, *, x, y, w, h) -> None:
        """Render a treemap as a native PowerPoint chartEx chart (editable, themed);
        falls back to the drawn squarified treemap if injection fails."""
        try:
            if self.add_native_treemap(slide, items, x=x, y=y, w=w, h=h):
                return
        except Exception as exc:                       # pragma: no cover
            logger.warning("native treemap failed (%s); using drawn fallback", exc)
        self.draw_treemap(slide, items, x=x, y=y, w=w, h=h)

    def add_native_treemap(self, slide, items, *, x, y, w, h,
                           series_name="Series 1") -> bool:
        """Insert a native PowerPoint treemap as a cx:chartex graphicFrame,
        reusing a known-good template's chart/colors/style parts and swapping in
        the data (a flat list of {label, value}). The embedded workbook is
        regenerated to match. Returns True on success."""
        import io, re, html as _html
        from pptx.opc.package import Part
        from pptx.opc.packuri import PackURI
        from pptx.oxml import parse_xml
        import openpyxl

        rows = [(str(it.get("label", "")), float(it.get("value", 0) or 0))
                for it in items if isinstance(it, dict)]
        rows = [(l, v) for l, v in rows if v > 0]
        if not rows:
            return False
        n = len(rows)
        _assets = _treemap_assets()
        tmpl = _assets["template"]

        cat_pts = "".join(f'<cx:pt idx="{i}">{_html.escape(l)}</cx:pt>'
                          for i, (l, _) in enumerate(rows))
        size_pts = "".join(f'<cx:pt idx="{i}">{v:g}</cx:pt>'
                           for i, (_, v) in enumerate(rows))
        # chartEx requires the points wrapped in <cx:lvl ptCount> and a
        # <cx:externalData> link to the embedded workbook. Omitting either
        # makes PowerPoint reject the chart and force a document repair.
        chart_data = (
            '<cx:chartData>'
            '<cx:externalData r:id="rId1" cx:autoUpdate="0"/>'
            '<cx:data id="0">'
            f'<cx:strDim type="cat"><cx:f>Sheet1!$A$2:$A${n + 1}</cx:f>'
            f'<cx:lvl ptCount="{n}">{cat_pts}</cx:lvl></cx:strDim>'
            f'<cx:numDim type="size"><cx:f>Sheet1!$B$2:$B${n + 1}</cx:f>'
            f'<cx:lvl ptCount="{n}" formatCode="General">{size_pts}</cx:lvl></cx:numDim>'
            '</cx:data></cx:chartData>'
        )
        chartex_xml = re.sub(r'<cx:chartData>.*?</cx:chartData>', chart_data,
                             tmpl, count=1, flags=re.S).replace("Sheet1!$D$1", "Sheet1!$B$1")

        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"
        ws["A1"] = "Category"; ws["B1"] = series_name
        for i, (l, v) in enumerate(rows, start=2):
            ws[f"A{i}"] = l; ws[f"B{i}"] = v
        buf = io.BytesIO(); wb.save(buf); xlsx_bytes = buf.getvalue()
        colors_bytes = _assets["colors"]
        style_bytes = _assets["style"]

        package = slide.part.package
        cx_pn = package.next_partname("/ppt/charts/chartEx%d.xml")
        idx = cx_pn.idx or 1
        col_part = Part(PackURI(f"/ppt/charts/colors{idx}.xml"),
                        "application/vnd.ms-office.chartcolorstyle+xml", package, colors_bytes)
        sty_part = Part(PackURI(f"/ppt/charts/style{idx}.xml"),
                        "application/vnd.ms-office.chartstyle+xml", package, style_bytes)
        xls_pn = package.next_partname("/ppt/embeddings/Microsoft_Excel_Worksheet%d.xlsx")
        xls_part = Part(xls_pn, "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet", package, xlsx_bytes)
        cx_part = Part(cx_pn, "application/vnd.ms-office.chartex+xml", package,
                       chartex_xml.encode("utf-8"))
        # Relate the embedded workbook FIRST so it takes rId1 (the id the
        # chartData's <cx:externalData> points at); colors/style are found by
        # relationship type, so their ids don't matter.
        cx_part.relate_to(xls_part, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package")
        cx_part.relate_to(col_part, "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle")
        cx_part.relate_to(sty_part, "http://schemas.microsoft.com/office/2011/relationships/chartStyle")
        rId = slide.part.relate_to(cx_part, "http://schemas.microsoft.com/office/2014/relationships/chartEx")

        EMU = EMU_PER_INCH
        existing = [int(el.get("id")) for el in slide.shapes._spTree.iter()
                    if el.tag.endswith("}cNvPr") and (el.get("id") or "").isdigit()]
        sid = (max(existing) if existing else 1) + 1
        gf_xml = (
            '<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<p:nvGraphicFramePr><p:cNvPr id="{sid}" name="Treemap"/>'
            '<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>'
            f'<p:xfrm><a:off x="{int(x * EMU)}" y="{int(y * EMU)}"/>'
            f'<a:ext cx="{int(w * EMU)}" cy="{int(h * EMU)}"/></p:xfrm>'
            '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex">'
            '<cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            f'r:id="{rId}"/></a:graphicData></a:graphic></p:graphicFrame>'
        )
        slide.shapes._spTree.append(parse_xml(gf_xml))
        return True

    def draw_treemap(self, slide, items, *, x, y, w, h) -> None:
        """Single-level treemap: rectangles sized by value (squarified layout),
        filled from the OW palette with white seams (the touching-shape rule),
        each labelled with its name and value. Values must be positive; tiles too
        small for text are left unlabelled.
        """
        rows = []
        for it in items:
            if isinstance(it, dict):
                rows.append((str(it.get("label", "")), float(it.get("value", 0) or 0)))
            else:
                rows.append((str(it), 1.0))
        rows = [(lbl, v) for lbl, v in rows if v > 0]
        if not rows:
            return
        rows.sort(key=lambda r: r[1], reverse=True)
        rects = squarify_treemap([v for _, v in rows], x, y, w, h)
        palette = TREEMAP_PALETTE

        def _fmt(v):
            return str(int(v)) if float(v).is_integer() else f"{v:g}"

        for i, (rect, (label, val)) in enumerate(zip(rects, rows)):
            hexc = palette[i % len(palette)]
            rdx, rdy = max(0.05, rect["dx"]), max(0.05, rect["dy"])
            shp = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(rect["x"]), Inches(rect["y"]), Inches(rdx), Inches(rdy))
            shp.fill.solid(); shp.fill.fore_color.rgb = self.rgb_hex(hexc)
            shp.line.color.rgb = self.rgb("lt1")        # white seams (touching rule)
            shp.line.width = Pt(Tokens.STROKE_W)
            shp.shadow.inherit = False
            tf = shp.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = Align.TOP
            self._box_margins(tf)
            if rdx < 0.7 or rdy < 0.4:               # too small to label cleanly
                continue
            dark_text = _is_light_hex(hexc)
            tcol = "dk1" if dark_text else "lt1"
            fs = int(fit_point_size(label, max(0.3, rdx - 0.16),
                                    max(0.2, rdy - 0.1), 14, min_pt=9))
            p = tf.paragraphs[0]; p.alignment = Align.LEFT
            self._run(p, label, size=fs, bold=True, color=tcol)
            p2 = tf.add_paragraph(); p2.alignment = Align.LEFT
            self._run(p2, _fmt(val), size=max(9, fs - 2), color=tcol)

    @staticmethod
    def _build_chart_data(xl_type, categories, series):
        if xl_type in XY_TYPES:
            data = XyChartData()
            for name, points in series:
                ser = data.add_series(name)
                for x, y in points:
                    ser.add_data_point(x, y)
            return data
        data = CategoryChartData()
        data.categories = list(categories)
        for name, values in series:
            data.add_series(name, tuple(values))
        return data

    # ------------------------------------------------------------------
    # OW chart styling — applies the full spec extracted from the
    # .crtx templates (palette, axes, gridlines, data labels, legend,
    # bar geometry, line stroke, doughnut hole). All set via lxml
    # because python-pptx exposes only a fraction of these knobs.
    # ------------------------------------------------------------------
    C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    _CNS = f"{{{C_NS}}}"
    _ANS = f"{{{A_NS}}}"

    def _theme_chart(self, chart, xl_type) -> None:
        """Apply the OW chart style: palette, axes, gridlines, labels, legend.

        Pure OOXML manipulation via lxml on the chart's XML tree. python-pptx
        is used only for high-level toggles (legend on/off, has_data_labels)
        because the full spec requires nodes python-pptx doesn't expose.
        """
        n_series = len(list(chart.series))

        # 0. Suppress the auto-title that python-pptx adds for single-series
        #    charts (it surfaces the series name above the plot, which fights
        #    the slide's own title). The OW reference always sets this to "1".
        chart_el = chart._chartSpace  # noqa: SLF001
        atd = chart_el.find(f".//{self._CNS}autoTitleDeleted")
        if atd is not None:
            atd.set("val", "1")

        # 1. Series fills (and lines, for line charts).
        self._apply_chart_palette(chart, xl_type)

        # 2. Legend on the bottom for multi-series charts and pie/doughnut.
        if xl_type in PIE_LIKE or n_series > 1:
            chart.has_legend = True
        else:
            chart.has_legend = False

        # 3. Data labels: show value, position by chart type, 12pt minor font.
        self._apply_chart_data_labels(chart, xl_type)

        # 4. Axis styling — gridlines, axis lines, tick label fonts.
        self._apply_chart_axes(chart, xl_type)

        # 5. Legend font.
        self._apply_chart_legend(chart)

        # 6. Per-chart-type knobs (gapWidth/overlap, line width, hole size).
        self._apply_chart_specifics(chart, xl_type)

        # Axis tick labels: 12pt by default.
        for _ax in ("category_axis", "value_axis"):
            try:
                getattr(chart, _ax).tick_labels.font.size = Pt(Tokens.TYPE_AXIS)
            except Exception:
                pass

        # 7. Markers for line_with_markers (unchanged from prior behavior).
        # Markers OFF by default for all line charts (cleaner OW look).
        if xl_type in (XL_CHART_TYPE.LINE, XL_CHART_TYPE.LINE_MARKERS):
            for ser in chart.series:
                try:
                    ser.marker.style = XL_MARKER_STYLE.NONE
                except Exception:
                    pass

        # 8. Patch empty <a:endParaRPr/> elements. python-pptx's default
        #    chart-space txPr block emits this without a `lang` attribute
        #    for `areaChart` and `doughnutChart` types; PowerPoint flags
        #    the bare `<a:endParaRPr/>` as schema-invalid and triggers a
        #    repair prompt on open. Adding `lang="en-US"` matches what
        #    PowerPoint itself writes when saving a fresh chart and
        #    silences the validator without changing rendering.
        for epr in chart_el.iter(f"{self._ANS}endParaRPr"):
            if epr.get("lang") is None:
                epr.set("lang", "en-US")

    # ---- chart styling helpers ---------------------------------------

    def _apply_chart_palette(self, chart, xl_type) -> None:
        """Cycle every series (or pie point) through OW_CHART_PALETTE."""
        if xl_type in PIE_LIKE:
            # Pie/doughnut: colors live on each *point*, not the series.
            try:
                ser = chart.series[0]
                for i, point in enumerate(ser.points):
                    hex_v = OW_CHART_PALETTE[i % len(OW_CHART_PALETTE)]
                    point.format.fill.solid()
                    point.format.fill.fore_color.rgb = self.rgb_hex(hex_v)
            except Exception as exc:
                logger.debug("Could not theme pie points: %s", exc)
            return

        # Bar/column/line/area: color is on the series.
        for i, ser in enumerate(chart.series):
            hex_v = OW_CHART_PALETTE[i % len(OW_CHART_PALETTE)]
            try:
                if xl_type in (XL_CHART_TYPE.LINE, XL_CHART_TYPE.LINE_MARKERS):
                    # For lines, color the line stroke only (no fill).
                    ser.format.line.color.rgb = self.rgb_hex(hex_v)
                else:
                    ser.format.fill.solid()
                    ser.format.fill.fore_color.rgb = self.rgb_hex(hex_v)
                    # Hide the per-bar outline so adjacent bars merge cleanly
                    # in stacked charts.
                    self._set_ser_no_line(ser)
                    # Keep negative bars SOLID-filled. PowerPoint defaults
                    # <c:invertIfNegative> to true, which renders negative bars
                    # hollow (white fill, coloured outline); force it off so an
                    # all-negative or mixed series fills like any other.
                    if xl_type in BAR_LIKE:
                        self._set_invert_if_negative(ser, False)
            except Exception as exc:
                logger.debug("Could not theme series %d: %s", i, exc)

    def _set_invert_if_negative(self, ser, value: bool) -> None:
        """Set `<c:invertIfNegative val="0|1"/>` on a bar/column series, placed
        in its schema-correct slot (immediately after `<c:spPr>`)."""
        el = ser._element  # noqa: SLF001
        existing = el.find(f"{self._CNS}invertIfNegative")
        if existing is not None:
            existing.set("val", "1" if value else "0")
            return
        node = etree.Element(f"{self._CNS}invertIfNegative")
        node.set("val", "1" if value else "0")
        spPr = el.find(f"{self._CNS}spPr")
        if spPr is not None:
            spPr.addnext(node)          # CT_BarSer: invertIfNegative follows spPr
        else:
            # no spPr — fall back to after tx/order/idx (whichever is last)
            anchor = None
            for tag in ("tx", "order", "idx"):
                found = el.find(f"{self._CNS}{tag}")
                if found is not None:
                    anchor = found
            if anchor is not None:
                anchor.addnext(node)
            else:
                el.insert(0, node)

    def _set_ser_no_line(self, ser) -> None:
        """Force `<a:ln><a:noFill/></a:ln>` on a series' spPr."""
        spPr = ser._element.find(f"{self._CNS}spPr")  # noqa: SLF001
        if spPr is None:
            return
        # Remove any existing <a:ln> and append a fresh noFill one.
        for ln in spPr.findall(f"{self._ANS}ln"):
            spPr.remove(ln)
        ln = etree.SubElement(spPr, f"{self._ANS}ln")
        etree.SubElement(ln, f"{self._ANS}noFill")

    def _apply_chart_data_labels(self, chart, xl_type) -> None:
        """Turn on data labels and position them per OW spec.

        Per OOXML schema, ``dLblPos`` values are chart-type-specific.
        Using an invalid value (e.g. ``outEnd`` on an ``areaChart``)
        causes PowerPoint to flag the file as malformed and prompt to
        repair. Valid values per chart family:

        - Column / Bar clustered → ``ctr``, ``inBase``, ``inEnd``, ``outEnd``
        - Column / Bar stacked   → ``ctr``, ``inBase``, ``inEnd`` (no outEnd)
        - Pie / Doughnut         → ``bestFit``, ``b``, ``ctr``, ``inEnd``, ``outEnd``, ``r``, ``t``
        - Line / Scatter         → ``b``, ``ctr``, ``l``, ``r``, ``t``
        - Area                   → ``ctr``, ``b``, ``t``, ``r``, ``l``  (no outEnd!)
        """
        if xl_type in (XL_CHART_TYPE.LINE, XL_CHART_TYPE.LINE_MARKERS):
            return  # line charts are noisy with labels; OW spec leaves them off

        try:
            chart.plots[0].has_data_labels = True
        except AttributeError:
            # python-pptx's `CT_ScatterChart` has no `dLbls` attribute, so
            # `has_data_labels = True` raises AttributeError for XY/scatter.
            # Skip cleanly — data labels are unusual on scatter anyway — but
            # log a warning so the author sees that this chart's numeric
            # labels (if any) will not be localized to the deck's language.
            if xl_type in XY_TYPES:
                logger.warning("Scatter/XY chart: data-label setup skipped "
                               "(python-pptx limitation); any labels will "
                               "render in the client's locale, not the deck's.")
            return
        except Exception:
            return

        # Determine position label from the chart type.
        is_stacked = xl_type in (
            XL_CHART_TYPE.COLUMN_STACKED, XL_CHART_TYPE.COLUMN_STACKED_100,
            XL_CHART_TYPE.BAR_STACKED,    XL_CHART_TYPE.BAR_STACKED_100,
        )
        is_area = xl_type in (
            XL_CHART_TYPE.AREA, XL_CHART_TYPE.AREA_STACKED,
            XL_CHART_TYPE.AREA_STACKED_100,
        )
        is_pie = xl_type in PIE_LIKE
        is_doughnut = xl_type == XL_CHART_TYPE.DOUGHNUT
        is_xy = xl_type in XY_TYPES
        if is_doughnut:
            # Doughnut charts do NOT accept a <c:dLblPos> element at all
            # (unlike pie, which accepts ``ctr``/``bestFit``/...). Emitting
            # one — even ``ctr`` — makes PowerPoint flag the file as corrupt
            # and prompt to repair on open. Suppress the position element;
            # labels still render (centered) without it.
            pos_val = None
        elif is_xy:
            # XY/scatter rejects ``outEnd``; valid values are b/ctr/l/r/t.
            pos_val = "r"
        elif is_pie or is_stacked or is_area:
            # ``ctr`` is valid for every chart type that supports labels and
            # is the right visual choice for stacked / area / pie geometry.
            pos_val = "ctr"
        else:
            pos_val = "outEnd"

        # We need to mutate per-series <c:dLbls> blocks because the OW
        # template puts position + text style on each series, not at the
        # plot level. python-pptx won't help us build these.
        for i, ser in enumerate(chart.series):
            ser_el = ser._element  # noqa: SLF001
            # Determine label color: white on dark fills (stacked only),
            # otherwise inherit dk1 / tx1.
            label_color_hex = None
            if is_stacked:
                bar_hex = OW_CHART_PALETTE[i % len(OW_CHART_PALETTE)]
                label_color_hex = DEFAULT_THEME["lt1"] if _hex_is_dark(bar_hex) else None
            self._build_ser_dLbls(ser_el, pos_val, label_color_hex)

    def _build_ser_dLbls(self, ser_el, pos_val: Optional[str],
                         label_color_hex: Optional[str]) -> None:
        """Replace (or create) a series' <c:dLbls> block with the OW style."""
        # Remove any existing dLbls and rebuild — clean slate.
        for existing in ser_el.findall(f"{self._CNS}dLbls"):
            ser_el.remove(existing)

        dLbls = etree.SubElement(ser_el, f"{self._CNS}dLbls")
        # The dLbls element must come BEFORE <c:cat>, <c:val> in the schema.
        # Move it into position.
        cat = ser_el.find(f"{self._CNS}cat")
        val = ser_el.find(f"{self._CNS}val")
        anchor = cat if cat is not None else val
        if anchor is not None:
            ser_el.remove(dLbls)
            anchor.addprevious(dLbls)

        # txPr — text properties for the label
        txPr = etree.SubElement(dLbls, f"{self._CNS}txPr")
        l, t, r, b = CHART_LABEL_INSETS_EMU
        bodyPr = etree.SubElement(txPr, f"{self._ANS}bodyPr",
                                  rot="0", spcFirstLastPara="1",
                                  vertOverflow="ellipsis", vert="horz",
                                  wrap="square",
                                  lIns=str(l), tIns=str(t),
                                  rIns=str(r), bIns=str(b),
                                  anchor="ctr", anchorCtr="1")
        etree.SubElement(bodyPr, f"{self._ANS}spAutoFit")
        etree.SubElement(txPr, f"{self._ANS}lstStyle")
        p = etree.SubElement(txPr, f"{self._ANS}p")
        pPr = etree.SubElement(p, f"{self._ANS}pPr")
        defRPr = etree.SubElement(pPr, f"{self._ANS}defRPr",
                                   sz=str(Tokens.TYPE_BODY * 100), b="0",
                                   i="0", u="none", strike="noStrike",
                                   kern="1200", baseline="0")
        if label_color_hex is None:
            # Inherit theme's tx1 color.
            sFill = etree.SubElement(defRPr, f"{self._ANS}solidFill")
            etree.SubElement(sFill, f"{self._ANS}schemeClr", val="tx1")
        else:
            sFill = etree.SubElement(defRPr, f"{self._ANS}solidFill")
            etree.SubElement(sFill, f"{self._ANS}srgbClr", val=label_color_hex)
        etree.SubElement(defRPr, f"{self._ANS}latin", typeface="+mn-lt")
        etree.SubElement(defRPr, f"{self._ANS}ea",    typeface="+mn-ea")
        etree.SubElement(defRPr, f"{self._ANS}cs",    typeface="+mn-cs")

        # dLblPos — omitted entirely when pos_val is None (doughnut charts,
        # which reject the element). dLbls must precede <c:cat>/<c:val>, and
        # within dLbls, dLblPos precedes the show* flags — both preserved here.
        if pos_val is not None:
            etree.SubElement(dLbls, f"{self._CNS}dLblPos", val=pos_val)
        # numFmt — thousand separator on all data labels, locale-aware so
        # PowerPoint renders 1,000 (en-US), 1.000 (de-DE), 1 000 (fr-FR)
        # per the deck's language.
        etree.SubElement(dLbls, f"{self._CNS}numFmt",
                         formatCode=f"[$-{self._number_lcid}]#,##0",
                         sourceLinked="0")
        # show flags — value only
        for tag, val in (("showLegendKey", "0"), ("showVal", "1"),
                         ("showCatName",   "0"), ("showSerName", "0"),
                         ("showPercent",   "0"), ("showBubbleSize", "0")):
            etree.SubElement(dLbls, f"{self._CNS}{tag}", val=val)

    def _apply_chart_axes(self, chart, xl_type) -> None:
        """Style cat-axis (visible thin line, no gridlines) and val-axis
        (no line, light major gridlines). Sets 12pt minor font on tick labels.
        """
        if xl_type in PIE_LIKE:
            return  # no axes on pie/doughnut

        chart_el = chart._chartSpace  # noqa: SLF001
        cat_axes = chart_el.findall(f".//{self._CNS}catAx")
        val_axes = chart_el.findall(f".//{self._CNS}valAx")
        date_axes = chart_el.findall(f".//{self._CNS}dateAx")

        for ax in cat_axes + date_axes:
            self._style_axis(ax, role="cat")
        for ax in val_axes:
            self._style_axis(ax, role="val")
            # Locale-aware numFmt on the value axis so tick labels match the
            # data-label format (both show "1,000" in en-US, "1.000" in de-DE).
            # Without this, PowerPoint falls back to "General" and tick labels
            # render in the client's locale — inconsistent with the deck's.
            self._set_axis_numfmt(ax)

    def _set_axis_numfmt(self, ax_el) -> None:
        """Ensure the axis has a <c:numFmt> element with the deck's LCID
        prefix. Replaces any existing numFmt; positioned per CT_ValAx schema
        (after axId/scaling/delete/axPos, before majorGridlines/title/etc.)."""
        code = f"[$-{self._number_lcid}]#,##0"
        for existing in ax_el.findall(f"{self._CNS}numFmt"):
            ax_el.remove(existing)
        numFmt = etree.Element(f"{self._CNS}numFmt")
        numFmt.set("formatCode", code)
        numFmt.set("sourceLinked", "0")
        # Insert after axPos (or the earliest of the required predecessors).
        anchor = None
        for tag in ("axPos", "delete", "scaling", "axId"):
            found = ax_el.find(f"{self._CNS}{tag}")
            if found is not None:
                anchor = found
                break
        if anchor is not None:
            anchor.addnext(numFmt)
        else:
            ax_el.insert(0, numFmt)

    def _style_axis(self, ax_el, role: str) -> None:
        """Apply OW spec to a single axis element (line + gridlines). Tick marks
        and value-axis visibility are owned by ``normalize_chart_parts`` (run on
        save), so they are deliberately not set here."""
        # axis line (spPr/ln)
        spPr = ax_el.find(f"{self._CNS}spPr")
        if spPr is None:
            spPr = etree.SubElement(ax_el, f"{self._CNS}spPr")

        # Category/date axis: pin tick labels to the LOW end (bottom) so they
        # never ride up to the axis crossing. With negative values the cat axis
        # crosses at zero (top of the plot), and the default "nextTo" would put
        # the labels there, overlapping the bars. "low" keeps them at the
        # bottom for every chart. (Positioned before spPr per CT_CatAx schema.)
        if role == "cat":
            tlp = ax_el.find(f"{self._CNS}tickLblPos")
            if tlp is None:
                tlp = etree.Element(f"{self._CNS}tickLblPos")
                spPr.addprevious(tlp)
            tlp.set("val", "low")
        # remove existing ln + fills
        for child in list(spPr):
            tag = etree.QName(child).localname
            if tag in ("noFill", "solidFill", "ln"):
                spPr.remove(child)
        # spPr always has noFill (transparent area for the axis itself)
        etree.SubElement(spPr, f"{self._ANS}noFill")
        ln = etree.SubElement(spPr, f"{self._ANS}ln")
        if role == "val":
            # value axis: no line at all (the gridlines do the visual work)
            etree.SubElement(ln, f"{self._ANS}noFill")
        else:
            # category axis: thin tx1 line
            ln.set("w", str(CHART_LINE_WEIGHT_EMU))
            ln.set("cap", "flat"); ln.set("cmpd", "sng"); ln.set("algn", "ctr")
            sFill = etree.SubElement(ln, f"{self._ANS}solidFill")
            etree.SubElement(sFill, f"{self._ANS}schemeClr", val="tx1")
            etree.SubElement(ln, f"{self._ANS}round")

        # majorGridlines on val axis only
        existing_grid = ax_el.find(f"{self._CNS}majorGridlines")
        if role == "val":
            if existing_grid is None:
                existing_grid = etree.SubElement(ax_el, f"{self._CNS}majorGridlines")
            # rebuild the gridline spPr/ln
            for child in list(existing_grid):
                existing_grid.remove(child)
            grid_spPr = etree.SubElement(existing_grid, f"{self._CNS}spPr")
            grid_ln = etree.SubElement(grid_spPr, f"{self._ANS}ln",
                                        w=str(CHART_LINE_WEIGHT_EMU),
                                        cap="flat", cmpd="sng", algn="ctr")
            grid_fill = etree.SubElement(grid_ln, f"{self._ANS}solidFill")
            etree.SubElement(grid_fill, f"{self._ANS}srgbClr",
                              val=CHART_GRIDLINE_COLOR_HEX)
            etree.SubElement(grid_ln, f"{self._ANS}round")
            self._reorder_axis_majorGridlines(ax_el, existing_grid)
        else:
            # cat axis: ensure no gridlines
            if existing_grid is not None:
                ax_el.remove(existing_grid)

        # txPr — tick label font (12pt minor, not bold)
        existing_tx = ax_el.find(f"{self._CNS}txPr")
        if existing_tx is not None:
            ax_el.remove(existing_tx)
        txPr = etree.SubElement(ax_el, f"{self._CNS}txPr")
        etree.SubElement(txPr, f"{self._ANS}bodyPr",
                          rot="-60000000", spcFirstLastPara="1",
                          vertOverflow="ellipsis", vert="horz",
                          wrap="square", anchor="ctr", anchorCtr="1")
        etree.SubElement(txPr, f"{self._ANS}lstStyle")
        p = etree.SubElement(txPr, f"{self._ANS}p")
        pPr = etree.SubElement(p, f"{self._ANS}pPr")
        defRPr = etree.SubElement(pPr, f"{self._ANS}defRPr",
                                   sz=str(Tokens.TYPE_BODY * 100), b="0",
                                   i="0", u="none", strike="noStrike",
                                   kern="1200", baseline="0")
        sFill = etree.SubElement(defRPr, f"{self._ANS}solidFill")
        etree.SubElement(sFill, f"{self._ANS}schemeClr", val="tx1")
        etree.SubElement(defRPr, f"{self._ANS}latin", typeface="+mn-lt")
        etree.SubElement(defRPr, f"{self._ANS}ea",    typeface="+mn-ea")
        etree.SubElement(defRPr, f"{self._ANS}cs",    typeface="+mn-cs")

    def _reorder_axis_majorGridlines(self, ax_el, grid_el) -> None:
        """Move <c:majorGridlines> to its canonical position (after <c:axPos>)."""
        ax_el.remove(grid_el)
        # Insert after axPos if present, otherwise after the first 4 elements.
        ax_pos = ax_el.find(f"{self._CNS}axPos")
        if ax_pos is not None:
            ax_pos.addnext(grid_el)
        else:
            ax_el.append(grid_el)

    def _apply_chart_legend(self, chart) -> None:
        """Set legend position=bottom, no overlay, 12pt minor font."""
        if not chart.has_legend:
            return
        legend_el = chart._chartSpace.find(f".//{self._CNS}legend")  # noqa: SLF001
        if legend_el is None:
            return
        # legendPos
        pos = legend_el.find(f"{self._CNS}legendPos")
        if pos is None:
            pos = etree.SubElement(legend_el, f"{self._CNS}legendPos")
            legend_el.insert(0, pos)
        pos.set("val", "b")
        # overlay
        ov = legend_el.find(f"{self._CNS}overlay")
        if ov is None:
            ov = etree.SubElement(legend_el, f"{self._CNS}overlay")
        ov.set("val", "0")
        # txPr — 12pt minor
        existing_tx = legend_el.find(f"{self._CNS}txPr")
        if existing_tx is not None:
            legend_el.remove(existing_tx)
        txPr = etree.SubElement(legend_el, f"{self._CNS}txPr")
        etree.SubElement(txPr, f"{self._ANS}bodyPr",
                          rot="0", spcFirstLastPara="1",
                          vertOverflow="ellipsis", vert="horz",
                          wrap="square", anchor="ctr", anchorCtr="1")
        etree.SubElement(txPr, f"{self._ANS}lstStyle")
        p = etree.SubElement(txPr, f"{self._ANS}p")
        pPr = etree.SubElement(p, f"{self._ANS}pPr")
        defRPr = etree.SubElement(pPr, f"{self._ANS}defRPr",
                                   sz=str(Tokens.TYPE_BODY * 100), b="0",
                                   i="0", u="none", strike="noStrike",
                                   kern="1200", baseline="0")
        sFill = etree.SubElement(defRPr, f"{self._ANS}solidFill")
        etree.SubElement(sFill, f"{self._ANS}schemeClr", val="tx1")
        etree.SubElement(defRPr, f"{self._ANS}latin", typeface="+mn-lt")
        etree.SubElement(defRPr, f"{self._ANS}ea",    typeface="+mn-ea")
        etree.SubElement(defRPr, f"{self._ANS}cs",    typeface="+mn-cs")

    def _apply_chart_specifics(self, chart, xl_type) -> None:
        """Per-chart-type knobs: gapWidth/overlap (bar), holeSize (doughnut),
        line width (line)."""
        chart_el = chart._chartSpace  # noqa: SLF001

        # Bar/column charts: gapWidth + overlap by (barDir, grouping).
        # CT_BarChart requires gapWidth/overlap to appear AFTER ser/dLbls but
        # BEFORE serLines/axId. python-pptx does not pre-create these elements,
        # so a naive append lands them after <c:axId> — invalid ordering that
        # can trigger a repair prompt. Insert before the first axId instead.
        for bar_chart in chart_el.findall(f".//{self._CNS}barChart"):
            barDir   = bar_chart.find(f"{self._CNS}barDir")
            grouping = bar_chart.find(f"{self._CNS}grouping")
            key = (barDir.get("val") if barDir is not None else "col",
                   grouping.get("val") if grouping is not None else "clustered")
            geom = CHART_BAR_GEOMETRY.get(key, {"gapWidth": 220, "overlap": 0})
            first_axid = bar_chart.find(f"{self._CNS}axId")
            for tag, val in geom.items():
                el = bar_chart.find(f"{self._CNS}{tag}")
                if el is None:
                    el = bar_chart.makeelement(f"{self._CNS}{tag}", {})
                    if first_axid is not None:
                        first_axid.addprevious(el)
                    else:
                        bar_chart.append(el)
                el.set("val", str(val))

        # Doughnut: holeSize=75
        for d_chart in chart_el.findall(f".//{self._CNS}doughnutChart"):
            hole = d_chart.find(f"{self._CNS}holeSize")
            if hole is None:
                hole = etree.SubElement(d_chart, f"{self._CNS}holeSize")
            hole.set("val", str(CHART_DOUGHNUT_HOLE_PCT))

        # Line: 2.25pt stroke on every series
        if xl_type in (XL_CHART_TYPE.LINE, XL_CHART_TYPE.LINE_MARKERS):
            for ser in chart.series:
                ser_el = ser._element  # noqa: SLF001
                spPr = ser_el.find(f"{self._CNS}spPr")
                if spPr is None:
                    continue
                ln = spPr.find(f"{self._ANS}ln")
                if ln is None:
                    ln = etree.SubElement(spPr, f"{self._ANS}ln")
                ln.set("w", str(CHART_LINE_SERIES_EMU))
                ln.set("cap", "rnd")

    # ---- table binding -------------------------------------------------

    def insert_table(self, slide, placeholder, headers, rows):
        if any(v is None for v in (placeholder.left, placeholder.top,
                                    placeholder.width, placeholder.height)):
            raise ValueError("Cannot bind table: placeholder geometry unresolved.")
        if not headers or not rows:
            raise ValueError("insert_table requires headers and at least one row.")
        cols_n = len(headers); rows_n = len(rows) + 1
        table_shape = slide.shapes.add_table(
            rows_n, cols_n,
            placeholder.left, placeholder.top, placeholder.width, placeholder.height,
        )
        self._fill_table(table_shape.table, headers, rows)
        self._bind_to_placeholder(table_shape, placeholder)

    def _fill_table(self, table, headers, rows) -> None:
        """Apply the OW Table 1 style and fill cells: header row bold and
        bottom-anchored, body rows regular and top-anchored, compact heights.
        Shared by the placeholder ``insert_table`` and the canvas ``draw_table``."""
        cols_n = len(headers)
        self._apply_table_style(table, OW_TABLE_STYLE_1, first_row=True, band_row=False)
        for c, h in enumerate(headers):
            cell = table.cell(0, c)
            self._style_cell(cell, str(h), color="dk1", bold=True, size=Pt(Tokens.TYPE_TABLE_HEADER))
            cell.text_frame.vertical_anchor = Align.TABLE_HEADER_ANCHOR
            cell.text_frame.paragraphs[0].alignment = Align.TABLE_H
        table.rows[0].height = Inches(Tokens.SPACE_TABLE_HEADER)
        for r, row in enumerate(rows, start=1):
            for c in range(cols_n):
                value = row[c] if c < len(row) else ""
                cell = table.cell(r, c)
                self._style_cell(cell, str(value), color="dk1", bold=False, size=Pt(Tokens.TYPE_BODY))
                cell.text_frame.vertical_anchor = Align.TABLE_BODY_ANCHOR
                cell.text_frame.paragraphs[0].alignment = Align.TABLE_H
            table.rows[r].height = Inches(Tokens.SPACE_TABLE_ROW)

    def draw_table(self, slide, x, y, w, h, headers, rows) -> None:
        """Canvas table block: a data table at explicit geometry (the composer's
        table region), styled identically to the placeholder ``insert_table``."""
        if not headers or not rows:
            return
        ts = slide.shapes.add_table(len(rows) + 1, len(headers),
                                    Inches(x), Inches(y), Inches(w), Inches(h))
        self._fill_table(ts.table, headers, rows)

    def draw_kpi(self, slide, x, y, w, h, *, heading="", kpis=None,
                 value=None, label="", delta="") -> None:
        """Canvas KPI block: an optional heading plus one or more stacked KPIs
        (big navy value, small upper-case label, optional delta). Accepts a
        single value/label/delta or a ``kpis`` list of {value,label,delta}."""
        if kpis is None:
            kpis = [{"label": label, "value": value or "", "delta": delta}]
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame; tf.clear(); tf.word_wrap = True
        tf.vertical_anchor = Align.TOP
        first = True
        if heading:
            p = tf.paragraphs[0]
            self._run(p, heading, size=Tokens.TYPE_HEADING, bold=True)
            p.space_after = Pt(Tokens.PARA_SUBHEAD)
            first = False
        for kpi in kpis:
            p = tf.paragraphs[0] if first else tf.add_paragraph(); first = False
            p.space_before = Pt(Tokens.PARA_PROSE)
            self._run(p, kpi.get("value", ""), size=Tokens.TYPE_KPI)
            if kpi.get("label"):
                pl = tf.add_paragraph()
                self._run(pl, str(kpi["label"]), size=Tokens.TYPE_BODY)
            if kpi.get("delta"):
                pd = tf.add_paragraph()
                self._run(pd, kpi["delta"], size=Tokens.TYPE_BODY)
        self.suppress_bullets(tf)

    def draw_image_placeholder(self, slide, x, y, w, h, *, heading="",
                               captions=None) -> None:
        """Canvas image block: an optional heading plus one or more captioned
        placeholder panels (light-blue fill). There is no image pipeline — these
        mark where pictures go, named by their caption, for manual drop-in."""
        captions = captions or ["Image"]
        top = y
        if heading:
            tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(Tokens.SPACE_HEADING_H))
            tf = tb.text_frame; tf.clear()
            self._run(tf.paragraphs[0], heading, size=Tokens.TYPE_HEADING, bold=True)
            top = y + Tokens.SPACE_HEADING_H + Tokens.SPACE_GAP
        n = len(captions); gap = Tokens.SPACE_GAP
        cell_h = max(0.4, (y + h - top - gap * (n - 1)) / n)
        for i, cap in enumerate(captions):
            cy = top + i * (cell_h + gap)
            rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                          Inches(x), Inches(cy), Inches(w), Inches(cell_h))
            rect.fill.solid(); rect.fill.fore_color.rgb = self.rgb("accent3")
            rect.line.fill.background(); rect.shadow.inherit = False
            tf = rect.text_frame; tf.word_wrap = True
            tf.vertical_anchor = Align.MIDDLE
            p = tf.paragraphs[0]; p.alignment = Align.CENTER
            self._run(p, cap, size=Tokens.TYPE_EYEBROW)

    def _apply_table_style(self, table, style_guid: str, *,
                           first_row: bool = True, band_row: bool = False) -> None:
        """Point the table at a named style from the template's tableStyles.xml
        and set the part-emphasis flags. ``tableStyleId`` must be the last child
        of ``tblPr`` per the schema; python-pptx already creates one, so we
        update it in place (or append if missing)."""
        tbl = table._tbl  # noqa: SLF001
        tblPr = tbl.find(f"{{{A_NS}}}tblPr")
        if tblPr is None:
            tblPr = tbl.makeelement(f"{{{A_NS}}}tblPr", {})
            tbl.insert(0, tblPr)
        tblPr.set("firstRow", "1" if first_row else "0")
        tblPr.set("bandRow",  "1" if band_row else "0")
        for attr in ("lastRow", "firstCol", "lastCol", "bandCol"):
            tblPr.set(attr, "0")
        sid = tblPr.find(f"{{{A_NS}}}tableStyleId")
        if sid is None:
            sid = etree.SubElement(tblPr, f"{{{A_NS}}}tableStyleId")
        sid.text = style_guid

    def _style_cell(self, cell, text, *, color, bold, size, fill=None):
        # Fill is left to the named table style unless one is explicitly passed.
        if fill is not None:
            cell.fill.solid(); cell.fill.fore_color.rgb = self.rgb(fill)
        tf = cell.text_frame; tf.clear(); tf.word_wrap = True
        run = tf.paragraphs[0].add_run()
        run.text = text; run.font.bold = bold; run.font.size = size
        self.set_font_color(run.font, color)

    # ---- image insertion (PICTURE placeholders) -----------------------

    def insert_picture(self, placeholder, image_path: str):
        """Replace a PICTURE placeholder with an image from disk.

        Unlike charts and tables, PICTURE placeholders support native
        ``insert_picture()`` in python-pptx: the resulting <p:pic> element
        carries the placeholder's <p:ph> marker automatically. No OOXML
        surgery needed.
        """
        ph_type = self.placeholder_type(placeholder)
        if ph_type != PP_PLACEHOLDER.PICTURE:
            raise ValueError(
                f"insert_picture requires a PICTURE placeholder; got {ph_type}."
            )
        return placeholder.insert_picture(image_path)

    # ---- engine-composed graphics (Title Only canvas) -----------------

    @staticmethod
    def _chevron_geometry(n, y, w, h, has_bullets):
        """Compute the chevron band geometry. Returns (shape_h, band_y, chev_w,
        stride, custom_adj, adj_value, bullet_y, bullet_h).

        The chevron shape height is a HARD 1.0" everywhere (Tokens.CHEVRON_H) —
        with or without bullets, on any layout — so process flows are visually
        consistent across the deck. arrow_depth is capped so the tip stays clean.
        """
        shape_h = Tokens.CHEVRON_H                     # hard 1.0", never scaled
        if has_bullets:
            bullet_y = y + shape_h + GRID_GAP
            bullet_h = max(0.5, h - shape_h - GRID_GAP)
        else:
            bullet_y = None
            bullet_h = 0
        band_y = y if has_bullets else y + max(0.0, (h - shape_h) / 2.0)
        natural = shape_h / 2.0
        arrow_depth = min(natural, 0.4)   # shallower arrow -> wider text rectangle
        custom_adj = arrow_depth < natural
        adj_value = int(arrow_depth / shape_h * 100000) if custom_adj else 50000
        overlap = max(0.0, arrow_depth - GRID_GAP)
        chev_w = (w + overlap * (n - 1)) / n
        stride = chev_w - overlap
        return (shape_h, band_y, chev_w, stride, overlap,
                custom_adj, adj_value, bullet_y, bullet_h)

    def draw_process_chevrons(self, slide, x: float, y: float, w: float, h: float,
                               steps: Sequence[Dict]) -> None:
        """Render N process shapes in a horizontal row, with optional bullet
        columns beneath each shape.

        Visual conventions (matching the OW reference):
          - First shape is a PENTAGON (flat left, pointed right) — marks the
            start of the flow.
          - Subsequent shapes are CHEVRONs (notched left, pointed right) — they
            nest into the previous shape, conveying continuity.
          - Shapes alternate between accent1 (navy) and accent2 (light blue).
          - If any step provides ``heading`` or ``bullets``, the band height
            collapses to a fixed shape height and a bullet area is drawn
            beneath, with one column per step. Columns are stride-aligned so
            each column starts at its corresponding shape's left edge.

        Per-step JSON shape:
          {"label": "Discover", "heading": "Phase 1", "bullets": ["…", "…"]}
        ``heading`` and ``bullets`` are both optional; ``label`` lives on the
        chevron itself.
        """
        n = len(steps)
        if n == 0:
            return

        has_bullets = any(
            isinstance(s, dict) and (s.get("heading") or s.get("bullets"))
            for s in steps
        )

        # Band + optional bullet-area geometry (extracted for clarity).
        (shape_h, band_y, chev_w, stride, overlap, custom_adj, adj_value,
         bullet_y, bullet_h) = self._chevron_geometry(n, y, w, h, has_bullets)

        for i, step in enumerate(steps):
            cx = x + i * stride
            shape_type = MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON
            shape = slide.shapes.add_shape(
                shape_type,
                Inches(cx), Inches(band_y), Inches(chev_w), Inches(shape_h),
            )
            # If we capped arrow_depth, set the shape's adj value explicitly.
            if custom_adj:
                sp_pr = shape._element.spPr  # noqa: SLF001
                prst_geom = sp_pr.find(f"{{{A_NS}}}prstGeom")
                if prst_geom is not None:
                    av_lst = prst_geom.find(f"{{{A_NS}}}avLst")
                    if av_lst is None:
                        av_lst = etree.SubElement(prst_geom, f"{{{A_NS}}}avLst")
                    # Clear any inherited adj
                    for gd in list(av_lst):
                        av_lst.remove(gd)
                    etree.SubElement(av_lst, f"{{{A_NS}}}gd",
                                      name="adj", fmla=f"val {adj_value}")
            # Per BRAND_COLORS, the default fill for any decorative shape is
            # F7F3EE (warm cream). Process steps are equal-weighted by
            # default — no alternating colors. The pentagon-vs-chevron shape
            # itself carries the "start of flow" cue.
            shape.fill.solid()
            shape.fill.fore_color.rgb = self.rgb_hex(BRAND_COLORS["default_fill"])
            shape.line.fill.background()

            label = step.get("label", "") if isinstance(step, dict) else str(step)
            tf = shape.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = Align.MIDDLE
            self._box_margins(tf)
            p = tf.paragraphs[0]
            p.alignment = Align.CENTER
            r = p.add_run()
            r.text = label
            r.font.bold = True   # body font; OK to be bold per brand rules
            r.font.size = Pt(fit_point_size(label, max(0.4, chev_w - overlap),
                                            shape_h, Tokens.TYPE_BODY, min_pt=AUTOFIT_MIN_PT))
            self.set_font_color(r.font, "dk1")  # dark text on cream fill

            # Bullet column directly underneath this shape (if any step has bullets,
            # all columns get rendered — empty columns stay empty for visual
            # alignment). Column width is stride minus a fixed 0.25" gap so
            # adjacent columns sit close without touching.
            if has_bullets:
                col_w = max(0.5, stride - 0.25)
                self._draw_chevron_bullet_column(
                    slide,
                    cx, bullet_y, col_w, bullet_h,
                    step if isinstance(step, dict) else {},
                )

    def _draw_chevron_bullet_column(self, slide, x, y, w, h, step):
        """One bullet column under a chevron: bold heading + bulleted list,
        sized exactly per the BRAND_SPEC bullet level table.

        Uses level 4 (Heading) for the heading and level 1 (Bullet 1) for
        each bullet — same as ``fill_bullets`` does in placeholders, but
        applied explicitly here because text boxes don't inherit master
        styles. ``apply_text_level`` enforces the 'primary font never
        bold' rule on the way through."""
        heading = step.get("heading", "")
        bullets = step.get("bullets", [])
        if not heading and not bullets:
            return
        box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_top = Pt(0); tf.margin_bottom = Pt(0)
        tf.margin_left = Pt(0); tf.margin_right = Pt(0)
        first = True
        if heading:
            p = tf.paragraphs[0]
            r = p.add_run()
            r.text = heading
            self.apply_text_level(p, level=4)  # Heading: 12pt, body, bold
            first = False
        for bullet in bullets:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            r = p.add_run()
            r.text = bullet
            self.apply_text_level(p, level=1)  # Bullet 1: 12pt, body, '•'

    def draw_org_chart(self, slide, x: float, y: float, w: float, h: float,
                        root: Dict, reports: Sequence[Dict] = ()) -> None:
        """Render a hierarchical org chart with arbitrary depth.

        The tree is described recursively — each node is a dict with
        ``name``, ``title``, and an optional ``children`` list of more
        nodes. Backward compat: a flat ``reports`` list is accepted at
        the top level and merged into ``root.children``.

        Layout strategy
        ---------------
        Two-pass leaves-proportional placement (Reingold–Tilford style):

          1. **Measure** — count the leaves descending from each subtree.
          2. **Place** — each subtree is allocated a horizontal slot
             proportional to its leaf count; nodes center within their
             slot. Vertical bands are equal-height.

        Connectors use a "bus line" pattern: parent drops to a horizontal
        bus halfway between levels, the bus spans all children, and each
        child rises from its top-center to the bus. This produces the
        clean orthogonal look readers expect from an org chart.

        Box sizes (width and height) and font sizes scale with the tree's
        depth and breadth so 4–5 level charts with 15+ nodes still fit in
        the standard content area without overlap.

        Parameters
        ----------
        root
            Root node dict. May contain ``children`` for subordinates.
        reports
            Optional flat list of direct reports — used only for
            backward compatibility with the v3.0 two-level shape. If
            ``root`` already has ``children``, this is ignored.
        """
        # ---- merge legacy `reports` into the tree ----
        if reports and not root.get("children"):
            root = dict(root)            # shallow copy so we don't mutate caller's dict
            root["children"] = list(reports)

        # ---- analyze tree ----
        depth   = self._tree_depth(root)
        n_leaves = max(1, self._count_leaves(root))

        if depth <= 0:
            return

        # ---- vertical layout ----
        # Each level gets equal vertical space. Within a band, the box
        # claims ~55% and the connector zone claims ~45%. As depth grows
        # we still cap box height so dense trees don't have boxes too
        # tall to fit text-friendly fonts.
        band_h = h / depth
        box_h  = min(1.05, max(0.55, band_h * 0.55))
        # Horizontal box width is bounded both by the per-leaf slot
        # (so wide trees pack tight) and by a sensible max so 4-leaf
        # trees don't get absurdly wide boxes.
        per_leaf_w = w / n_leaves
        box_w_max  = min(2.6, max(1.2, per_leaf_w * 0.88))

        # Pick fonts based on the smaller box. Smaller boxes → smaller text.
        name_pt, title_pt = self._org_font_sizes(box_h, box_w_max)

        # ---- place every node ----
        placements: List[Dict[str, Any]] = []
        self._place_org_node(
            node=root, x_left=x, x_right=x + w, y=y + 0.05,
            level=0, depth=depth, band_h=band_h, box_h=box_h,
            box_w_max=box_w_max, placements=placements,
        )

        # ---- draw boxes ----
        for p in placements:
            level = p["level"]
            # Root gets the strong navy treatment; deeper levels get the
            # cream card (equal-weighted in the visual hierarchy).
            if level == 0:
                fill, text_color = "accent1", "lt1"
            else:
                fill, text_color = BRAND_COLORS["default_fill"], "dk1"
            self._draw_org_box(
                slide,
                p["x"], p["y"], p["w"], p["h"],
                p["node"].get("name", ""),
                p["node"].get("title", ""),
                fill=fill, text_color=text_color,
                name_pt=name_pt, title_pt=title_pt,
            )

        # ---- draw connectors (bus-line style) ----
        # Group placements by parent so we can draw one bus per parent.
        by_id = {id(p["node"]): p for p in placements}
        for p in placements:
            parent_node = p["node"]
            children = parent_node.get("children") or []
            if not children:
                continue
            child_pts = [by_id.get(id(c)) for c in children if id(c) in by_id]
            child_pts = [cp for cp in child_pts if cp is not None]
            if not child_pts:
                continue
            self._draw_org_bus(slide, p, child_pts)

    # ---- helpers --------------------------------------------------------

    @classmethod
    def _tree_depth(cls, node: Dict) -> int:
        """Number of levels in the tree (root = level 1)."""
        children = node.get("children") or []
        if not children:
            return 1
        return 1 + max(cls._tree_depth(c) for c in children)

    @classmethod
    def _count_leaves(cls, node: Dict) -> int:
        """Number of leaves under this subtree (a node with no children
        counts as a leaf)."""
        children = node.get("children") or []
        if not children:
            return 1
        return sum(cls._count_leaves(c) for c in children)

    @staticmethod
    @staticmethod
    def _snap_font(fs: float) -> int:
        """Snap a fitted size down to the consistent ladder: 12, 10, or 8pt."""
        return 12 if fs >= 12 else (10 if fs >= 10 else 8)

    @staticmethod
    def _org_font_sizes(box_h: float, box_w: float) -> Tuple[int, int]:
        """Pick (name_pt, title_pt) so the text fits the box without
        truncation at the typical name/title lengths."""
        if box_h >= 0.80 and box_w >= 1.5:
            s = 12
        elif box_h >= 0.62:
            s = 10
        else:
            s = 8
        return (s, s)

    def _place_org_node(self, *, node: Dict,
                        x_left: float, x_right: float, y: float,
                        level: int, depth: int, band_h: float, box_h: float,
                        box_w_max: float,
                        placements: List[Dict[str, Any]]) -> None:
        """Recursively place ``node`` and its descendants. Each subtree
        gets a horizontal slot proportional to its leaf count; the node
        sits centered within its slot, and its children are placed in
        the row below using the same proportional allocation."""
        slot_w = x_right - x_left
        own_w  = min(box_w_max, slot_w * 0.85)
        cx     = (x_left + x_right) / 2
        placements.append({
            "node": node,
            "x": cx - own_w / 2,
            "y": y,
            "w": own_w,
            "h": box_h,
            "level": level,
        })

        children = node.get("children") or []
        if not children:
            return

        total_leaves = sum(self._count_leaves(c) for c in children)
        cursor = x_left
        for c in children:
            c_leaves = self._count_leaves(c)
            c_w = slot_w * c_leaves / total_leaves
            self._place_org_node(
                node=c, x_left=cursor, x_right=cursor + c_w,
                y=y + band_h, level=level + 1,
                depth=depth, band_h=band_h, box_h=box_h,
                box_w_max=box_w_max, placements=placements,
            )
            cursor += c_w

    def _draw_org_bus(self, slide, parent: Dict[str, Any],
                      children: List[Dict[str, Any]]) -> None:
        """Draw the bus-line connector pattern from one parent to its
        children: parent drops to a horizontal bus halfway between
        levels, then each child rises from the bus to its top-center.

        Single-child parents skip the bus and use a straight drop."""
        parent_cx = parent["x"] + parent["w"] / 2
        parent_bot_y = parent["y"] + parent["h"]
        child_top_y  = children[0]["y"]
        bus_y = (parent_bot_y + child_top_y) / 2

        # Parent → bus
        self._draw_connector_line(slide, parent_cx, parent_bot_y, parent_cx, bus_y)

        if len(children) == 1:
            # Single child: extend the line straight to it (no bus needed).
            child_cx = children[0]["x"] + children[0]["w"] / 2
            # If perfectly aligned, one line; otherwise dog-leg via bus.
            if abs(child_cx - parent_cx) < 0.01:
                self._draw_connector_line(slide, parent_cx, bus_y, child_cx, child_top_y)
            else:
                self._draw_connector_line(slide, parent_cx, bus_y, child_cx, bus_y)
                self._draw_connector_line(slide, child_cx, bus_y, child_cx, child_top_y)
            return

        # Bus span — leftmost to rightmost child center
        child_cxs = [c["x"] + c["w"] / 2 for c in children]
        leftmost  = min(child_cxs)
        rightmost = max(child_cxs)
        self._draw_connector_line(slide, leftmost, bus_y, rightmost, bus_y)

        # Bus → each child's top-center
        for cp, cx in zip(children, child_cxs):
            self._draw_connector_line(slide, cx, bus_y, cx, cp["y"])

    def _draw_org_box(self, slide, x, y, w, h, name, title, *,
                      fill, text_color,
                      name_pt: int = 12, title_pt: int = 10):
        box = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(x), Inches(y), Inches(w), Inches(h),
        )
        box.fill.solid()
        box.fill.fore_color.rgb = self.resolve_color(fill)
        # No outline on org boxes (per design): the fill alone separates them.
        box.line.fill.background()
        tf = box.text_frame
        tf.word_wrap = True
        # Tighter margins as boxes shrink so text doesn't get clipped.
        self._box_margins(tf)
        # Name
        p1 = tf.paragraphs[0]
        p1.alignment = Align.CENTER
        r1 = p1.add_run()
        r1.text = name
        r1.font.bold = True
        r1.font.size = Pt(name_pt)
        self.set_font_color(r1.font, text_color)
        # Title
        if title:
            p2 = tf.add_paragraph()
            p2.alignment = Align.CENTER
            r2 = p2.add_run()
            r2.text = title
            r2.font.size = Pt(title_pt)
            self.set_font_color(r2.font, text_color)

    @staticmethod
    def _ensure_nonzero_extent(shape, min_emu: int = 1) -> None:
        """Clamp the shape's ``<a:ext>`` cx/cy to at least ``min_emu``.

        PowerPoint flags shapes with ``cx=0`` or ``cy=0`` as malformed and
        triggers a "repair" prompt on open, even though python-pptx
        happily produces zero-extent connectors when an
        ``add_connector(STRAIGHT, x1, y1, x2, y2)`` call has matching
        x or y coordinates (a purely horizontal or vertical line).

        Bumping the collapsed axis to 1 EMU (~0.0003pt — well below
        rendering precision) is invisible to viewers but satisfies the
        validator. Called for every connector we emit."""
        ext = shape._element.spPr.find(f"{{{A_NS}}}xfrm/{{{A_NS}}}ext")  # noqa: SLF001
        if ext is None:
            return
        if int(ext.get("cx", 0)) <= 0:
            ext.set("cx", str(min_emu))
        if int(ext.get("cy", 0)) <= 0:
            ext.set("cy", str(min_emu))

    def _draw_connector_line(self, slide, x1, y1, x2, y2) -> None:
        line = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT,
            Inches(x1), Inches(y1), Inches(x2), Inches(y2),
        )
        line.line.color.rgb = self.rgb("accent4")
        line.line.width = Pt(Tokens.STROKE_W_MED)
        self._ensure_nonzero_extent(line)

    def _render_diagram_label(self, paragraph, label: str, size,
                              color: str = "dk1") -> None:
        """Diagram-label formatting used by pyramid, cycle, and similar
        labelled-shape drawers. When the label follows a ``"Head: body"``
        pattern (colon then space), the head becomes a bold run and the body
        a regular run — mirrors the OW-house convention that the diagram
        label is a bolded head with a descriptive tail. Labels without the
        pattern render as a single regular-weight run."""
        if not label:
            return
        head, sep, tail = label.partition(": ")
        if sep and tail:                              # "Head: body" form
            r = paragraph.add_run(); r.text = head + ": "
            r.font.bold = True
            r.font.size = Pt(size); self.set_font_color(r.font, color)
            r = paragraph.add_run(); r.text = tail
            r.font.bold = False
            r.font.size = Pt(size); self.set_font_color(r.font, color)
        else:                                          # plain label
            r = paragraph.add_run(); r.text = label
            r.font.bold = False
            r.font.size = Pt(size); self.set_font_color(r.font, color)

    def draw_pyramid(self, slide, x: float, y: float, w: float, h: float,
                      levels: Sequence[Dict]) -> None:
        """Render a basic pyramid: stacked trapezoid bands, narrow at top.

        ``levels`` is ordered top-to-bottom (top is the narrowest, peak
        of the pyramid; bottom is the widest, base of the pyramid).

        Each level is drawn as a freeform polygon with four explicit corners
        — this gives exact alignment between adjacent levels' edges, which
        the TRAPEZOID preset doesn't reliably reproduce across renderers
        (renderers disagree about how the ``adj`` value
        maps to the inset).
        """
        n = len(levels)
        if n == 0:
            return
        band_h = h / n

        for i, level in enumerate(levels):
            # Width progression: bottom-of-pyramid at level n-1 has width w.
            # Level i (0 = top) has bottom width w*(i+1)/n and top width w*i/n.
            bot_w = w * (i + 1) / n
            top_w = w * i / n
            level_y_top = y + i * band_h
            level_y_bot = y + (i + 1) * band_h

            # Four corners (clockwise from top-left), centered horizontally.
            cx = x + w / 2  # canvas horizontal center
            tl = (cx - top_w / 2, level_y_top)
            tr = (cx + top_w / 2, level_y_top)
            br = (cx + bot_w / 2, level_y_bot)
            bl = (cx - bot_w / 2, level_y_bot)

            # Build the polygon. python-pptx's freeform builder takes a
            # starting point and then a series of line segments to subsequent
            # points; closing brings us back to the start.
            EMU = EMU_PER_INCH
            builder = slide.shapes.build_freeform(
                tl[0] * EMU, tl[1] * EMU, scale=1.0,
            )
            # If the top has zero width (level 0), the top-left and top-right
            # collapse to a single apex point — the polygon becomes a triangle.
            if top_w > 0:
                builder.add_line_segments([
                    (tr[0] * EMU, tr[1] * EMU),
                    (br[0] * EMU, br[1] * EMU),
                    (bl[0] * EMU, bl[1] * EMU),
                ], close=True)
            else:
                # Triangle: from apex (tl == tr) → bottom-right → bottom-left → close
                builder.add_line_segments([
                    (br[0] * EMU, br[1] * EMU),
                    (bl[0] * EMU, bl[1] * EMU),
                ], close=True)
            shape = builder.convert_to_shape()

            # Per BRAND_COLORS, the default fill for any decorative shape is
            # F7F3EE (warm cream). Pyramid levels are visually differentiated
            # by their decreasing widths and thin separating borders, not by
            # color. If a level needs to be highlighted (e.g., the apex),
            # callers can mark it explicitly in a future spec extension.
            shape.fill.solid()
            shape.fill.fore_color.rgb = self.rgb_hex(BRAND_COLORS["default_fill"])
            # Touching bands: outline = background colour (white) so the
            # seams read as clean separations, not grey rules.
            shape.line.color.rgb = self.rgb("lt1")
            shape.line.width = Pt(Tokens.STROKE_W)

            # Put the label INSIDE the band shape (not a separate text layer)
            # so it travels with the trapezoid/triangle. A filled shape gets
            # 0.15" text margins on all sides. If the label follows a
            # "Head: body" pattern, the head renders bold and the body regular
            # via _render_diagram_label (OW-house convention).
            label = level.get("label", "") if isinstance(level, dict) else str(level)
            if label:
                tf = shape.text_frame
                tf.word_wrap = True
                self._box_margins(tf)
                tf.vertical_anchor = Align.MIDDLE
                p = tf.paragraphs[0]
                p.alignment = Align.CENTER
                self._render_diagram_label(p, label, Tokens.TYPE_HEADING, "dk1")

    def draw_cycle(self, slide, x: float, y: float, w: float, h: float,
                    items: Sequence[Dict]) -> None:
        """Render a basic cycle: N labeled ovals arranged on a virtual circle,
        with arrow connectors flowing item 0 → 1 → … → N-1 → 0.

        Arrows are straight connector lines between the perimeters of adjacent
        ovals, with arrowheads added directly at the OOXML level (python-pptx
        doesn't expose tail-end arrow properties through its high-level API).
        """
        n = len(items)
        if n == 0:
            return

        cx = x + w / 2
        cy = y + h / 2

        gap = GRID_GAP         # minimum clear space between adjacent ovals
        aspect = 1.0          # perfect circles (height = width)

        if n == 1:
            oval_w = min(2.4, w * 0.4)
            oval_h = oval_w * aspect
            radius = 0.0
            positions = [(cx - oval_w / 2, cy - oval_h / 2)]
        else:
            sin_term = math.sin(math.pi / n)
            # Largest oval that sits on a circle with >= `gap` between neighbours
            # (2*R*sin(pi/n) >= oval_w + gap) AND fits the content box; then use the
            # largest radius the box allows so spacing is >= gap and the area is used.
            best = None
            ow = 2.4
            while ow >= 0.7:
                oh = ow * aspect
                r_gap = (ow + gap) / (2 * sin_term)        # min radius, no overlap
                r_fit = min((h - oh) / 2, (w - ow) / 2)    # max radius that fits
                if r_gap <= r_fit:
                    best = (ow, oh, r_fit)                 # r_fit >= r_gap → gap >= 0.25"
                    break
                ow -= 0.05
            if best is None:                               # tiny area fallback
                ow = 0.7; oh = ow * aspect
                best = (ow, oh, max(0.1, min((h - oh) / 2, (w - ow) / 2)))
            oval_w, oval_h, radius = best
            positions = []
            for i in range(n):
                theta = 2 * math.pi * i / n - math.pi / 2  # start at top, clockwise
                ix = cx + radius * math.cos(theta) - oval_w / 2
                iy = cy + radius * math.sin(theta) - oval_h / 2
                positions.append((ix, iy))

        # Size the label font so the longest single word fits the oval interior
        # width (wrap only at spaces, never mid-word).
        labels = [(it.get("label", "") if isinstance(it, dict) else str(it)) for it in items]
        longest_word = max((len(wd) for lb in labels for wd in lb.split()), default=4)
        interior_w = max(0.4, oval_w - 0.18)
        label_fs = int(max(AUTOFIT_MIN_PT, min(12, interior_w / (max(longest_word, 1) * 0.013))))

        # Draw arrows first so ovals render on top.
        for i in range(n):
            j = (i + 1) % n
            cx_i = positions[i][0] + oval_w / 2
            cy_i = positions[i][1] + oval_h / 2
            cx_j = positions[j][0] + oval_w / 2
            cy_j = positions[j][1] + oval_h / 2
            self._draw_cycle_arrow(slide, cx_i, cy_i, cx_j, cy_j,
                                    oval_w / 2, oval_h / 2)

        # Now the labeled ovals
        for i, (ix, iy) in enumerate(positions):
            shape = slide.shapes.add_shape(
                MSO_SHAPE.OVAL,
                Inches(ix), Inches(iy), Inches(oval_w), Inches(oval_h),
            )
            # Per BRAND_COLORS, default fill is F7F3EE. Cycle items are
            # equal-weighted; the arrows convey flow direction, not color.
            shape.fill.solid()
            shape.fill.fore_color.rgb = self.rgb_hex(BRAND_COLORS["default_fill"])
            shape.line.fill.background()   # no outline
            tf = shape.text_frame
            tf.word_wrap = True
            self._box_margins(tf)
            p = tf.paragraphs[0]
            p.alignment = Align.CENTER
            label = items[i].get("label", "") if isinstance(items[i], dict) else str(items[i])
            self._render_diagram_label(p, label, label_fs, "dk1")

    def _draw_cycle_arrow(self, slide, x1, y1, x2, y2, half_w, half_h) -> None:
        """Connector line with arrowhead, from the perimeter of one oval to
        the perimeter of the next. The line stops short of each oval so the
        arrowhead is visible against open space."""
        dx = x2 - x1
        dy = y2 - y1
        length = math.sqrt(dx * dx + dy * dy)
        if length < 0.01:
            return
        # Approximate the oval as a circle of radius = average of half-axes.
        # Good enough for the visual; perimeter math on an ellipse isn't worth
        # the extra code here.
        r = (half_w + half_h) / 2
        ux = dx / length
        uy = dy / length
        sx = x1 + ux * r
        sy = y1 + uy * r
        ex = x2 - ux * r
        ey = y2 - uy * r

        line = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT,
            Inches(sx), Inches(sy), Inches(ex), Inches(ey),
        )
        line.line.color.rgb = self.rgb("accent1")   # navy arrows
        line.line.width = Pt(Tokens.STROKE_AXIS)
        self._ensure_nonzero_extent(line)

        self._apply_arrowhead(line)   # arrow style 2 (filled triangle)

    # ---- adaptive canvas renderers: matrix / timeline / growing steps ----

    def _thin_line(self, slide, x1, y1, x2, y2, *, color="accent4",
                   width_pt=Tokens.STROKE_HAIRLINE) -> None:
        """A straight connector line in inches, themed."""
        ln = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT,
            Inches(x1), Inches(y1), Inches(x2), Inches(y2),
        )
        ln.line.color.rgb = self.resolve_color(color)
        ln.line.width = Pt(width_pt)
        self._ensure_nonzero_extent(ln)

    def draw_matrix(self, slide, x: float, y: float, w: float, h: float,
                    spec) -> None:
        """Render an N×M matrix with two graded axes and any number of items
        placed by normalized coordinates. ``rows``/``cols`` are arbitrary
        (2×2, 3×3, 5×5, …); item count is unlimited.

        spec = {
          "x_axis": "Effort", "y_axis": "Impact",      # axis titles (optional)
          "rows": 3, "cols": 3,                          # grid (default 2×2)
          "highlight_cell": [row, col],                  # optional, 0-indexed, row 0 = top
          "items": [{"label": "A", "x": 0.8, "y": 0.7,   # x,y in 0..1 (0=left/bottom)
                     "comment": "…"}, …]                  # or {"row":r,"col":c}
        }
        """
        if isinstance(spec, list):
            spec = {"items": spec}
        if not isinstance(spec, dict):
            return
        items = spec.get("items", []) or []
        rows = max(1, int(spec.get("rows", 2)))
        cols = max(1, int(spec.get("cols", 2)))
        x_axis = spec.get("x_axis", "")
        y_axis = spec.get("y_axis", "")

        left_gutter = 0.5 if y_axis else 0.1
        bottom_gutter = 0.45 if x_axis else 0.1
        px, py = x + left_gutter, y
        pw, ph = w - left_gutter, h - bottom_gutter

        # Optional highlighted cell (drawn first, under the grid).
        hi = spec.get("highlight_cell")
        if isinstance(hi, (list, tuple)) and len(hi) == 2:
            hr, hc = int(hi[0]), int(hi[1])
            if 0 <= hr < rows and 0 <= hc < cols:
                cw, ch = pw / cols, ph / rows
                cell = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE,
                    Inches(px + hc * cw), Inches(py + hr * ch),
                    Inches(cw), Inches(ch))
                cell.fill.solid(); cell.fill.fore_color.rgb = self.rgb("accent3")
                cell.line.fill.background()

        # Plot border.
        border = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(px), Inches(py), Inches(pw), Inches(ph))
        border.fill.background()
        border.line.color.rgb = self.rgb("accent4"); border.line.width = Pt(Tokens.STROKE_W_HEAVY)
        # Internal grid lines.
        for c in range(1, cols):
            gx = px + pw * c / cols
            self._thin_line(slide, gx, py, gx, py + ph)
        for r in range(1, rows):
            gy = py + ph * r / rows
            self._thin_line(slide, px, gy, px + pw, gy)

        # Axis titles.
        if x_axis:
            tb = slide.shapes.add_textbox(
                Inches(px), Inches(py + ph + 0.06), Inches(pw), Inches(Tokens.SPACE_AXIS_LABEL_H))
            tf = tb.text_frame; tf.word_wrap = True
            p = tf.paragraphs[0]; p.alignment = Align.LEFT
            self._run(p, f"{x_axis}  →", size=Tokens.TYPE_AUTOFIT_MIN, bold=True)
        if y_axis:
            cx_l = x + left_gutter * 0.42
            cy_l = py + ph / 2
            tb = slide.shapes.add_textbox(
                Inches(cx_l - ph / 2), Inches(cy_l - 0.16),
                Inches(ph), Inches(Tokens.SPACE_AXIS_LABEL_H))
            tb.rotation = 270
            tf = tb.text_frame; tf.word_wrap = False
            p = tf.paragraphs[0]; p.alignment = Align.LEFT
            self._run(p, f"{y_axis}  →", size=Tokens.TYPE_AUTOFIT_MIN, bold=True)

        # Items. Place by normalized x,y (0..1); fall back to row/col centers.
        fs = AUTOFIT_MIN_PT  # autofit floor (was 8-10 by count)
        dot = 0.14
        for it in items:
            if not isinstance(it, dict):
                continue
            if "x" in it and "y" in it:
                xn = max(0.0, min(1.0, float(it["x"])))
                yn = max(0.0, min(1.0, float(it["y"])))
            elif "row" in it and "col" in it:
                xn = (int(it["col"]) + 0.5) / cols
                yn = (rows - 0.5 - int(it["row"])) / rows
            else:
                xn = yn = 0.5
            ix = px + xn * pw
            iy = py + (1 - yn) * ph
            d = slide.shapes.add_shape(
                MSO_SHAPE.OVAL, Inches(ix - dot / 2), Inches(iy - dot / 2),
                Inches(dot), Inches(dot))
            d.fill.solid(); d.fill.fore_color.rgb = self.rgb("accent1")
            d.line.fill.background()
            label = it.get("label", "")
            comment = it.get("comment", "")
            if label or comment:
                lw = 1.7
                lb = slide.shapes.add_textbox(
                    Inches(ix + dot * 0.7), Inches(iy - 0.16),
                    Inches(lw), Inches(Tokens.SPACE_TL_LABEL_H))
                tf = lb.text_frame; tf.word_wrap = True
                tf.margin_top = Pt(0); tf.margin_bottom = Pt(0)
                p = tf.paragraphs[0]
                if label:
                    self._run(p, label, size=fs, bold=True)
                if comment:
                    p2 = tf.add_paragraph()
                    self._run(p2, comment, size=max(7, fs - 2))

    def draw_timeline(self, slide, x: float, y: float, w: float, h: float,
                      milestones) -> None:
        """Horizontal timeline with any number of milestones, alternating above
        and below the axis. Each milestone: {date|label, title, description?}."""
        if isinstance(milestones, dict):
            milestones = milestones.get("milestones", [])
        n = len(milestones)
        if n == 0:
            return
        axis_y = y + h / 2
        self._thin_line(slide, x, axis_y, x + w, axis_y,
                        color="accent1", width_pt=Tokens.STROKE_AXIS)
        seg = w / n
        # House body size is 12pt; keep it for normal timelines (≤7 milestones)
        # and reduce only when very dense, where 12pt would overflow the segment.
        fs_title = Tokens.TYPE_BODY if n <= 7 else 11   # 12pt default; floor 11 when dense
        fs_meta = fs_title                        # body == heading/sub-heading size (OW spec)
        block_h = h / 2 - 0.22
        for i, m in enumerate(milestones):
            mx = x + seg * (i + 0.5)
            dd = 0.16
            dot = slide.shapes.add_shape(
                MSO_SHAPE.OVAL, Inches(mx - dd / 2), Inches(axis_y - dd / 2),
                Inches(dd), Inches(dd))
            dot.fill.solid(); dot.fill.fore_color.rgb = self.rgb("accent1")
            dot.line.fill.background()
            above = (i % 2 == 0)
            bw = seg * 0.92
            bx = mx - bw / 2
            if above:
                by = y
                self._thin_line(slide, mx, by + block_h, mx, axis_y)
            else:
                by = axis_y + 0.22
                self._thin_line(slide, mx, axis_y, mx, by)
            icon = (m.get("icon") if isinstance(m, dict) else None)
            tf_x, tf_y, tf_w, tf_h = bx, by, bw, block_h
            force_top = False
            if icon:                                  # optional fine-line icon, left-aligned
                isz = min(0.30, bw * 0.30)
                try:
                    import icons as _ic
                    _ic.draw_icon(slide, str(icon), bx, by, isz, isz, color_hex="000F47")
                    tf_y = by + isz + 0.06
                    tf_h = max(0.2, block_h - isz - 0.06)
                    force_top = True
                except Exception:
                    pass
            tb = slide.shapes.add_textbox(
                Inches(tf_x), Inches(tf_y), Inches(tf_w), Inches(tf_h))
            tf = tb.text_frame; tf.word_wrap = True
            tf.vertical_anchor = Align.TOP if (force_top or not above) else Align.BOTTOM
            date = (m.get("date") or m.get("label", "")) if isinstance(m, dict) else str(m)
            title = m.get("title", "") if isinstance(m, dict) else ""
            desc = m.get("description", "") if isinstance(m, dict) else ""
            p = tf.paragraphs[0]; p.alignment = Align.LEFT       # date = sub-heading
            if date:
                self._run(p, date, size=fs_meta, bold=True, color="accent1")
            if title:
                p2 = tf.add_paragraph(); p2.alignment = Align.LEFT  # title = body
                p2.space_before = Pt(3)                             # default spacing
                self._run(p2, title, size=fs_title)
            if desc:
                p3 = tf.add_paragraph(); p3.alignment = Align.LEFT  # LEFT (was centre)
                p3.space_before = Pt(2)                             # default spacing
                self._run(p3, desc, size=fs_meta)

    def draw_growing_steps(self, slide, x: float, y: float, w: float, h: float,
                           steps) -> None:
        """Ascending 'growing steps' (staircase) with any number of steps. Each
        step is taller than the last; the final step is the target state and is
        highlighted. Step: {label, description?}."""
        if isinstance(steps, dict):
            steps = steps.get("steps", [])
        n = len(steps)
        if n == 0:
            return
        gap = GRID_GAP         # grid spacing between steps
        step_w = (w - gap * (n - 1)) / n
        min_frac = 0.32
        # Cap by count, then shrink-to-fit the narrowest box (the shortest, first
        # step) so labels don't wrap mid-word or spill — floored at AUTOFIT_MIN_PT.
        longest_label = max(
            (str(s.get("label", "")) if isinstance(s, dict) else str(s) for s in steps),
            key=len, default="")
        shortest_bar_h = h * (min_frac if n > 1 else 1.0)
        # Fit the label, then snap to the consistent ladder {12, 10, 8}.
        fs = self._snap_font(fit_point_size(longest_label, step_w, shortest_bar_h,
                                            Tokens.TYPE_BODY, min_pt=Tokens.TYPE_DIAGRAM_MIN))
        for i, s in enumerate(steps):
            frac = min_frac + (1 - min_frac) * (i / (n - 1)) if n > 1 else 1.0
            bar_h = h * frac
            bx = x + i * (step_w + gap)
            by = y + (h - bar_h)
            rect = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(bx), Inches(by),
                Inches(step_w), Inches(bar_h))
            rect.fill.solid()
            rect.fill.fore_color.rgb = self.rgb_hex(BRAND_COLORS["default_fill"])  # all cream
            rect.line.fill.background()   # no outline (touching-grid style)
            label = s.get("label", "") if isinstance(s, dict) else str(s)
            desc = s.get("description", "") if isinstance(s, dict) else ""
            tf = rect.text_frame; tf.word_wrap = True
            tf.vertical_anchor = Align.TOP
            m = Inches(Tokens.SPACE_SHAPE_MARGIN)
            tf.margin_top = m; tf.margin_bottom = m; tf.margin_left = m; tf.margin_right = m
            p = tf.paragraphs[0]; p.alignment = Align.LEFT
            rn = p.add_run(); rn.text = label
            rn.font.bold = True; rn.font.size = Pt(fs)
            self.set_font_color(rn.font, "dk1")
            if desc:
                p2 = tf.add_paragraph(); p2.alignment = Align.LEFT
                rn = p2.add_run(); rn.text = desc
                rn.font.size = Pt(fs)
                self.set_font_color(rn.font, "dk1")

    def draw_contents(self, slide, sections, *, x: float, y: float,
                      w: float, h: float) -> None:
        """Table of contents as a NATIVE border-less 3-column table: section
        number | section name | page number. All cells top-aligned, 16pt, Text 1
        navy, regular weight. A section may set ``number`` ("" to omit, e.g. an
        Appendix) and ``page``. Matches the OW Contents layout."""
        n = len(sections)
        if n == 0:
            return
        fs = Tokens.TYPE_CONTENTS
        num_w, page_w = 0.6, 0.9
        name_w = max(1.0, w - num_w - page_w)
        row_h = GRID_GAP_LG if n * GRID_GAP_LG <= h else max(GRID_GAP, h / n)
        gf = slide.shapes.add_table(n, 3, Inches(x), Inches(y),
                                    Inches(w), Inches(row_h * n))
        table = gf.table
        # Border-less, fill-less: the "No Style, No Grid" table style.
        self._apply_table_style(table, "{2D5ABB26-0587-4C30-8999-92F81FD0307C}",
                                first_row=False, band_row=False)
        table.columns[0].width = Inches(num_w)
        table.columns[1].width = Inches(name_w)
        table.columns[2].width = Inches(page_w)

        def _cell(cell, text, align):
            cell.fill.background()
            cell.vertical_anchor = Align.TOP
            cell.margin_top = Pt(0); cell.margin_bottom = Pt(0)
            cell.margin_left = Pt(0); cell.margin_right = Pt(0)
            tf = cell.text_frame; tf.word_wrap = True
            p = tf.paragraphs[0]; p.alignment = align
            self._run(p, text, size=fs)

        for i, sec in enumerate(sections):
            if isinstance(sec, dict):
                title = sec.get("title", "")
                page = sec.get("page", "")
                num = sec.get("number")
                if num is None:
                    num = f"{i + 1:02d}"
            else:
                title, page, num = str(sec), "", f"{i + 1:02d}"
            _cell(table.cell(i, 0), num, Align.LEFT)
            _cell(table.cell(i, 1), title, Align.LEFT)
            _cell(table.cell(i, 2), page, Align.RIGHT)
            table.rows[i].height = Inches(row_h)

    def _hanging_indent(self, paragraph, *, left_in: float, hang_in: float) -> None:
        """Set a paragraph's left indent (Before text) and a hanging indent."""
        pPr = paragraph._p.get_or_add_pPr()
        pPr.set("marL", str(int(left_in * EMU_PER_INCH)))
        pPr.set("indent", str(int(-hang_in * EMU_PER_INCH)))

    def draw_quote(self, slide, x: float, y: float, w: float, h: float,
                   quote: str, attribution="") -> None:
        """Big editorial pull-quote in ONE textbox (shape width 10"): the quote in
        Marsh Serif (54pt, 12pt space after), then the Name / Title / Location
        attribution in Noto Sans \u2014 all left-aligned with a 0.25" left + hanging
        indent, in the same textbox, Text 1 navy. Sits on a blank layout."""
        q = (quote or "").strip()
        if not q:
            return
        if not (q.startswith('"') or q.startswith("\u201c")):
            q = f"\u201c{q}\u201d"
        if isinstance(attribution, dict):
            attr_lines = [attribution.get("name"),
                          attribution.get("title") or attribution.get("role"),
                          attribution.get("location")]
        elif isinstance(attribution, str) and attribution.strip():
            attr_lines = [attribution.strip().lstrip("\u2014\u2013- ").strip()]
        else:
            attr_lines = []
        attr_lines = [str(a) for a in attr_lines if a]

        width = min(w, Tokens.QUOTE_SHAPE_W)          # shape width 10"
        nq = len(q)
        size = (Tokens.TYPE_QUOTE_MAX if nq <= 110 else 44 if nq <= 180
                else 36 if nq <= 280 else 28 if nq <= 420 else Tokens.TYPE_QUOTE_MIN)
        qb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(width), Inches(h))
        tf = qb.text_frame; tf.word_wrap = True
        tf.vertical_anchor = Align.MIDDLE
        tf.margin_left = Pt(0); tf.margin_right = Pt(0)
        tf.margin_top = Pt(0); tf.margin_bottom = Pt(0)
        p = tf.paragraphs[0]; p.alignment = Align.LEFT
        p.space_after = Pt(Tokens.PARA_QUOTE_GAP)
        self._hanging_indent(p, left_in=Tokens.QUOTE_INDENT, hang_in=Tokens.QUOTE_INDENT)
        r = p.add_run(); r.text = q
        self._set_run_font(r, kind="major", size_pt=size, color_token="dk1")
        for i, line in enumerate(attr_lines):
            pa = tf.add_paragraph(); pa.alignment = Align.LEFT
            self._hanging_indent(pa, left_in=Tokens.QUOTE_INDENT, hang_in=0)
            if i == 0:
                pa.space_before = Pt(Tokens.PARA_QUOTE_GAP)
            ra = pa.add_run(); ra.text = line
            self._set_run_font(ra, kind="minor", size_pt=Tokens.TYPE_BODY, color_token="dk1")
            ra.font.bold = False  # per brand: attribution regular weight, not bold

    def draw_text_block(self, slide, x: float, y: float, w: float, h: float,
                        heading: str = "", subheading: str = "",
                        bullets=None, paragraphs=None) -> None:
        """A single text pane: optional bold heading + regular subheading, then a
        (nestable) bullet list or prose paragraphs. Used as the text half of a
        graphic-with-text slide; shares the OW body-list formatting via
        apply_text_level so bullets match everywhere."""
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = Align.TOP
        tf.margin_left = Pt(Tokens.INSET_X); tf.margin_right = Pt(Tokens.INSET_R)
        tf.margin_top = Pt(Tokens.INSET_TOP); tf.margin_bottom = Pt(Tokens.INSET_TOP)
        first = True
        if heading:
            p = tf.paragraphs[0]; first = False
            p.space_after = Pt(Tokens.PARA_HEADING)
            r = p.add_run(); r.text = heading
            self.apply_text_level(p, level=4)          # bold 12pt heading
        if subheading:
            p = tf.paragraphs[0] if first else tf.add_paragraph(); first = False
            p.space_after = Pt(Tokens.PARA_SUBHEAD)
            r = p.add_run(); r.text = subheading
            self.apply_text_level(p, level=0)          # regular 12pt
        rows = []
        if bullets:
            rows = [(min(lvl, 3), txt, True) for lvl, txt in iter_bullets(bullets, 1)]
        elif paragraphs:
            rows = [(0, str(pp), False) for pp in paragraphs]
        for lvl, txt, is_bullet in rows:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.space_after = Pt(Tokens.PARA_BULLET if is_bullet else Tokens.PARA_PROSE)
            r = p.add_run(); r.text = txt
            self.apply_text_level(p, level=(lvl if is_bullet else 0))

    def draw_rule_text(self, slide, x: float, y: float, w: float, h: float,
                       *, heading: str = "", paragraphs=None, bullets=None,
                       rule: str = "left") -> None:
        """Text block fenced by a single accent rule on one side (``left`` or
        ``top``), at the uniform STROKE_RULE weight. The recurring "text box with
        left line" motif (matrix comments, timeline callouts). Heading is bold
        12pt; body is regular 12pt or a bullet list — all from the body levels."""
        inset = Tokens.SPACE_SHAPE_MARGIN
        if rule == "top":
            self._thin_line(slide, x, y, x + w, y,
                            color="dk1", width_pt=Tokens.STROKE_RULE)
            tx, ty, tw, th = x, y + inset, w, h - inset
        else:  # left
            self._thin_line(slide, x, y, x, y + h,
                            color="dk1", width_pt=Tokens.STROKE_RULE)
            tx, ty, tw, th = x + inset, y, w - inset, h
        self.draw_text_block(slide, tx, ty, tw, th,
                             heading=heading, bullets=bullets,
                             paragraphs=paragraphs)

    def draw_kpi_callout(self, slide, x: float, y: float, w: float, h: float,
                         *, value: str = "", label: str = "") -> None:
        """A KPI fenced by a top accent rule (STROKE_RULE): the "text box with top
        line" motif. Big 26pt value over a regular 12pt label, both navy — sizes
        and colour come straight from the tokens/theme."""
        self._thin_line(slide, x, y, x + w, y,
                        color="dk1", width_pt=Tokens.STROKE_RULE)
        top = y + Tokens.SPACE_SHAPE_MARGIN
        tb = slide.shapes.add_textbox(Inches(x), Inches(top),
                                      Inches(w), Inches(max(0.4, h - Tokens.SPACE_SHAPE_MARGIN)))
        tf = tb.text_frame; tf.clear(); tf.word_wrap = True
        tf.vertical_anchor = Align.TOP
        self._run(tf.paragraphs[0], value, size=Tokens.TYPE_KPI)
        if label:
            self._run(tf.add_paragraph(), label, size=Tokens.TYPE_BODY)
        self.suppress_bullets(tf)

    def draw_key_message(self, slide, x: float, y: float, w: float, h: float,
                         *, message: str = "", callouts=None) -> None:
        """Cream key-message panel: a tinted card (theme default_fill) holding a
        lead message underlined by the single heavier panel rule (STROKE_PANEL),
        with optional top-rule KPI callouts stacked beneath. Fill comes from
        BRAND_COLORS — never a literal hex."""
        panel = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                       Inches(x), Inches(y), Inches(w), Inches(h))
        panel.fill.solid()
        panel.fill.fore_color.rgb = self.rgb_hex(BRAND_COLORS["default_fill"])
        panel.line.fill.background(); panel.shadow.inherit = False
        pad = Tokens.SPACE_SHAPE_MARGIN
        if message:
            mb = slide.shapes.add_textbox(Inches(x + pad), Inches(y + pad),
                                          Inches(w - 2 * pad), Inches(1.3))
            tf = mb.text_frame; tf.clear(); tf.word_wrap = True
            self._run(tf.paragraphs[0], message, size=Tokens.TYPE_KPI)
            self.suppress_bullets(tf)
            self._thin_line(slide, x + pad, y + 1.55, x + w - pad, y + 1.55,
                            color="dk1", width_pt=Tokens.STROKE_PANEL)
        callouts = callouts or []
        if callouts:
            top = y + 1.85
            ch = max(0.6, (y + h - top - Tokens.SPACE_GAP * (len(callouts) - 1)) / len(callouts))
            for i, c in enumerate(callouts):
                cy = top + i * (ch + Tokens.SPACE_GAP)
                self.draw_kpi_callout(slide, x + pad, cy, w - 2 * pad, ch,
                                      value=c.get("value", ""), label=c.get("label", ""))

    def render_block(self, slide, kind: str, data: dict,
                     *, x: float, y: float, w: float, h: float) -> None:
        """Unified block dispatcher — the single entry point that renders any
        block kind into the box (x, y, w, h): text, chart, waterfall, treemap, or
        any diagram. Both composite panes (``draw_visual``) and the composer call
        this, so every block speaks the same render(slide, area, data) contract."""
        kind = (kind or "").strip()
        data = data or {}
        if kind in ("text", "bullets"):
            self.draw_text_block(slide, x, y, w, h,
                                 heading=data.get("heading", ""),
                                 subheading=data.get("subheading", ""),
                                 bullets=data.get("bullets"),
                                 paragraphs=data.get("paragraphs"))
        elif kind == "chart":
            spec = data.get("chart", data)
            if str(spec.get("type", "")).lower() == "treemap":
                cats = spec.get("categories", [])
                vals = (spec.get("series") or [{}])[0].get("values", [])
                items = spec.get("items") or [{"label": c, "value": v}
                                              for c, v in zip(cats, vals)]
                self.treemap(slide, items, x=x, y=y, w=w, h=h)
            else:
                self._add_floating_chart(slide, spec,
                                         cx=x, cy=y, cw=w, ch=h, in_cell=True)
        elif kind == "waterfall":
            self.add_waterfall(slide, data.get("waterfall", data),
                               cx=x, cy=y, cw=w, ch=h)
        elif kind == "treemap":
            tm = data.get("treemap", data)
            if tm.get("render") == "drawn":
                self.draw_treemap(slide, tm.get("items", []), x=x, y=y, w=w, h=h)
            else:
                self.treemap(slide, tm.get("items", []), x=x, y=y, w=w, h=h)
        elif kind == "kpi":
            self.draw_kpi(slide, x, y, w, h, heading=data.get("heading", ""),
                          kpis=data.get("kpis"), value=data.get("value"),
                          label=data.get("label", ""), delta=data.get("delta", ""))
        elif kind == "table":
            self.draw_table(slide, x, y, w, h,
                            data.get("headers", []), data.get("rows", []))
        elif kind == "image":
            reg = getattr(self, "_image_registry", None)
            src = data.get("source") or data.get("path") or data.get("prompt")
            if reg is not None and src:
                try:
                    import images as _img
                    sid = data.get("source")
                    if not sid:
                        sid = f"_compose_{data.get('name', id(data))}"
                        if reg.get(sid) is None:
                            reg.register(sid, data.get("path"), data.get("prompt"))
                    elif reg.get(sid) is None:
                        reg.register(sid, data.get("path"), data.get("prompt"))
                    _img.place_image(slide, reg, sid, x, y, w, h,
                                     aspect=data.get("aspect"), crop=data.get("crop"),
                                     name=data.get("name", sid),
                                     allow_generation=bool(getattr(self, "_allow_image_gen", False)))
                    return
                except Exception:
                    pass
            caps = data.get("captions") or ([data["caption"]] if data.get("caption") else None)
            self.draw_image_placeholder(slide, x, y, w, h,
                                        heading=data.get("heading", ""), captions=caps)
        elif kind == "growing_steps":
            self.draw_growing_steps(slide, x, y, w, h, data.get("steps", []))
        elif kind == "pyramid":
            self.draw_pyramid(slide, x, y, w, h, data.get("levels", []))
        elif kind == "cycle":
            self.draw_cycle(slide, x, y, w, h, data.get("items", []))
        elif kind in ("process", "chevrons"):
            self.draw_process_chevrons(slide, x, y, w, h, data.get("steps", []))
        elif kind == "org_chart":
            oc = data.get("org_chart", data)
            self.draw_org_chart(slide, x, y, w, h,
                                oc.get("root", {}), oc.get("reports", []))
        elif kind == "matrix":
            self.draw_matrix(slide, x, y, w, h, data.get("matrix", data))
        elif kind == "timeline":
            self.draw_timeline(slide, x, y, w, h, data.get("milestones", []))
        elif kind in ("rule_text", "text_rule"):
            self.draw_rule_text(slide, x, y, w, h,
                                heading=data.get("heading", ""),
                                paragraphs=data.get("paragraphs"),
                                bullets=data.get("bullets"),
                                rule=data.get("rule", "left"))
        elif kind == "kpi_callout":
            self.draw_kpi_callout(slide, x, y, w, h,
                                  value=data.get("value", ""),
                                  label=data.get("label", ""))
        elif kind == "key_message":
            self.draw_key_message(slide, x, y, w, h,
                                  message=data.get("message", ""),
                                  callouts=data.get("callouts"))
        elif kind in _CARD_BLOCK_KINDS:
            import cards as _cards
            _cards.render_card(slide, kind, data, x, y, w, h)
        elif kind in _COMPONENT_BLOCK_KINDS:
            self._render_component_block(slide, kind, data, x, y, w, h)
        else:
            logger.warning("render_block: unknown block kind %r", kind)

    def _render_component_block(self, slide, kind, data, x, y, w, h) -> None:
        """Render an OW component (stat_callout/icon_rows/banner/label_stack/
        compare/bracket_callout/sticker) into a compose region or pane."""
        import components as _co
        from types import SimpleNamespace as _NS
        area = _NS(x=x, y=y, w=w, h=h)
        if kind in ("stat_callout", "stat_cards", "stats"):
            _co.draw_stat_callouts(slide, area, data.get("stats", []))
        elif kind in ("icon_rows", "icon_list"):
            _co.draw_icon_rows(slide, area, data.get("rows", []))
        elif kind == "compare":
            _co.draw_compare(slide, area, data.get("pros", []), data.get("cons", []),
                             pros_heading=data.get("pros_heading", ""),
                             cons_heading=data.get("cons_heading", ""))
        elif kind == "sticker":
            _co.draw_sticker(slide, area, data.get("text", ""),
                             note=data.get("note", ""))
        elif kind == "infographic":
            import infographics as _ig
            _ig.draw_infographic(slide, area, data.get("infographic", data))

    def draw_grid(self, slide, spec) -> None:
        """Flexible grid with spanning. ``spec`` =
        ``{cols?, rows?, cells:[{block, data, col?, row?, colspan?, rowspan?}]}``.

        Columns and rows are equal tracks that flex to fill the content area on a
        fixed gutter — 0.5" by default, 0.25" when a track would fall below its
        readable floor (GRID_MIN_COL_W / GRID_MIN_ROW_H). A cell spans
        ``colspan x rowspan`` tracks. When cols/rows are omitted they are solved
        for the smallest grid (rows minimised first). Cells without an explicit
        ``col``/``row`` auto-flow into the first free slot (row-major)."""
        if not isinstance(spec, dict):
            return
        cells = [c for c in spec.get("cells", []) if isinstance(c, dict)]
        if not cells:
            return
        area = self.content_area()

        def cs_of(c):
            return max(1, int(c.get("colspan", 1) or 1))

        def rs_of(c):
            return max(1, int(c.get("rowspan", 1) or 1))

        N, M, _decl_n, _decl_m = resolve_grid_dims(cells, spec.get("cols"),
                                                   spec.get("rows"))

        def col_w(g):
            return (area.w - (N - 1) * g) / N

        def row_h(g):
            return (area.h - (M - 1) * g) / M

        LG, SM = Tokens.SPACE_GAP_LG, Tokens.SPACE_GAP
        cg = LG if col_w(LG) >= Tokens.GRID_MIN_COL_W else SM
        rg = LG if row_h(LG) >= Tokens.GRID_MIN_ROW_H else SM
        cw, rh = col_w(cg), row_h(rg)

        occ = [[False] * N for _ in range(M)]

        def fits(c, r, cs, rs):
            if c + cs > N or r + rs > M:
                return False
            return all(not occ[rr][cc]
                       for rr in range(r, r + rs) for cc in range(c, c + cs))

        def mark(c, r, cs, rs):
            for rr in range(r, r + rs):
                for cc in range(c, c + cs):
                    occ[rr][cc] = True

        def find(cs, rs):
            for r in range(M):
                for c in range(N):
                    if fits(c, r, cs, rs):
                        return c, r
            return None

        for cell in cells:
            cs, rs = cs_of(cell), rs_of(cell)
            if "col" in cell and "row" in cell:
                c, r = int(cell["col"]), int(cell["row"])
            else:
                pos = find(cs, rs)
                if pos is None:
                    logger.warning("grid: no room for a %dx%d cell; skipped", cs, rs)
                    continue
                c, r = pos
            mark(c, r, cs, rs)
            x = area.x + c * (cw + cg)
            y = area.y + r * (rh + rg)
            w = cs * cw + (cs - 1) * cg
            h = rs * rh + (rs - 1) * rg
            self.render_block(slide, cell.get("block", "text"),
                              cell.get("data", cell), x=x, y=y, w=w, h=h)

    def draw_composition(self, slide, rows) -> None:
        """Composer: place heterogeneous blocks into a grid of regions and render
        each through ``render_block``. ``rows`` is a list of
        ``{height?, regions:[{block, data, width?}]}``; width/height are relative
        weights. The grammar snaps regions to the content area, inserts token-sized
        gaps, caps a row at 4 regions, and warns when a region falls below the
        block's minimum sensible size (the conformance hook)."""
        area = self.content_area()
        gap = Tokens.SPACE_GAP_LG          # default gutter 0.5" (grid)
        rows = [r for r in rows if isinstance(r, dict) and r.get("regions")]
        if not rows:
            return
        h_weights = [float(r.get("height", 1.0)) for r in rows]
        tot_h = sum(h_weights) or 1.0
        avail_h = area.h - gap * (len(rows) - 1)
        y = area.y
        for row, hw in zip(rows, h_weights):
            rh = avail_h * (hw / tot_h)
            regions = row["regions"][:4]                      # grammar: <=4 per row
            w_weights = [float(reg.get("width", 1.0)) for reg in regions]
            tot_w = sum(w_weights) or 1.0
            avail_w = area.w - gap * (len(regions) - 1)
            x = area.x
            for reg, ww in zip(regions, w_weights):
                rw = avail_w * (ww / tot_w)
                kind = reg.get("block", "")
                min_w, min_h = BLOCK_MIN_SIZE.get(kind, (1.0, 0.5))
                if rw < min_w or rh < min_h:
                    logger.warning(
                        "compose: region %r is %.2f x %.2f in, below the %.2f x %.2f "
                        "minimum for this block — it may render cramped.",
                        kind, rw, rh, min_w, min_h)
                self.render_block(slide, kind, reg.get("data", reg),
                                  x=x, y=y, w=rw, h=rh)
                x += rw + gap
            y += rh + gap

    def draw_visual(self, slide, kind: str, data: dict,
                    x: float, y: float, w: float, h: float) -> None:
        """Back-compat alias: a graphic pane is just one block. Delegates to the
        unified dispatcher, so composite panes can now hold charts/waterfalls too."""
        self.render_block(slide, kind, data, x=x, y=y, w=w, h=h)

    def draw_columns(self, slide, x: float, y: float, w: float, h: float,
                     columns) -> None:
        """Render 2–5 parallel headed text columns — the standard multi-column
        layout — separated by a >=0.25" gap (no vertical rules), for parallel
        for parallel categories (strengths/weaknesses, pillars, profile facets).
        Each column: {heading, bullets:[...]} or {heading, paragraphs:[...]} or
        {heading, text:"..."}.
        """
        n = len(columns)
        if n == 0:
            return
        gap = GRID_GAP
        col_w = (w - gap * (n - 1)) / n
        for i, col in enumerate(columns):
            cx = x + i * (col_w + gap)
            if isinstance(col, dict):
                heading = col.get("heading", "") or col.get("label", "")
                bullets = col.get("bullets")
                paragraphs = col.get("paragraphs")
                text = col.get("text")
            else:
                heading, bullets, paragraphs, text = "", None, None, str(col)
            tb = slide.shapes.add_textbox(Inches(cx), Inches(y),
                                          Inches(col_w), Inches(h))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = Align.TOP
            tf.margin_left = Pt(Tokens.INSET_X); tf.margin_right = Pt(Tokens.INSET_R)
            tf.margin_top = Pt(0); tf.margin_bottom = Pt(0)
            first = True
            if heading:
                p = tf.paragraphs[0]; first = False
                p.space_after = Pt(Tokens.PARA_PROSE)
                r = p.add_run(); r.text = heading
                self.apply_text_level(p, level=4)   # Heading: 12pt bold, no bullet
            # Bullets (possibly nested) carry the glyph + hanging indent at their
            # level (1=•, 2=–, 3=-); prose paragraphs render with no glyph.
            rows = []
            if bullets:
                rows = [(min(lvl, 3), txt, True) for lvl, txt in iter_bullets(bullets, 1)]
            elif paragraphs:
                rows = [(0, str(b), False) for b in paragraphs]
            elif text:
                rows = [(0, str(text), False)]
            for lvl, txt, is_bullet in rows:
                p = tf.paragraphs[0] if first else tf.add_paragraph()
                first = False
                p.space_after = Pt(Tokens.PARA_BULLET if is_bullet else Tokens.PARA_PROSE)
                r = p.add_run(); r.text = txt
                self.apply_text_level(p, level=(lvl if is_bullet else 0))

    # ---- bottom adornments: Conclusion + Footnote ---------------------

    @staticmethod
    def estimate_footnote_height(text: str) -> float:
        """Estimate footnote shape height in inches from the number of lines the
        source wraps to across the 12.33" column — measured with real font metrics
        when available, so a long one-line source is counted as the 2–3 lines it
        actually wraps to (not a single line, which let the conclusion above it
        collide with the wrapped footnote). Each line is ~0.135" at 8pt."""
        if not text:
            return 0.0
        lines = _wrapped_line_count(
            text, size_pt=Tokens.TYPE_FOOTNOTE, kind="minor",
            width_in=FALLBACK_CONTENT.w, chars_per_line=250)
        return max(FOOTNOTE_LINE_HEIGHT, lines * FOOTNOTE_LINE_HEIGHT)

    def add_footnote(self, slide, text: str, baseline_y: float = FOOTER_BASELINE_Y):
        """Add a Footnote text box at the slide's footer baseline.

        Bottom-anchored, auto-fit, zero internal margins, 8pt minor-font text.
        The shape's top y is computed so its bottom sits exactly at
        ``baseline_y``; height grows upward as the text adds lines.

        Tags the shape with name="Footnote" so it's identifiable in PowerPoint
        and survives a round-trip through editing.
        """
        height = self.estimate_footnote_height(text)
        if height <= 0:
            return None
        top_y = baseline_y - height
        # Set the box to exactly the estimated height so the shape's bottom
        # sits at the footer baseline. spAutoFit will let PowerPoint refine
        # this on open if our line-height estimate is slightly off.
        box = slide.shapes.add_textbox(
            Inches(FALLBACK_CONTENT.x), Inches(top_y),
            Inches(FALLBACK_CONTENT.w), Inches(height),
        )
        # Tag the shape
        nv = box._element.find(f".//{{{P_NS}}}cNvPr")
        if nv is not None:
            nv.set("name", "Footnote")

        # bodyPr: zero margins, bottom-anchored, autofit, wrap within box width.
        # python-pptx's add_textbox defaults to wrap="none" which collapses the
        # box width to the longest line; we want a fixed 12.33" box that wraps.
        tx_body = box.text_frame._txBody  # noqa: SLF001
        body_pr = tx_body.find(f"{{{A_NS}}}bodyPr")
        if body_pr is None:
            body_pr = etree.SubElement(tx_body, f"{{{A_NS}}}bodyPr")
            tx_body.insert(0, body_pr)
        # Strip wrap if set, defaulting to square (the OOXML default)
        if "wrap" in body_pr.attrib:
            del body_pr.attrib["wrap"]
        for k in ("lIns", "tIns", "rIns", "bIns"):
            body_pr.set(k, "0")
        body_pr.set("rtlCol", "0")
        body_pr.set("anchor", Align.FOOTNOTE_ANCHOR_XML)
        # Replace any sizing children with spAutoFit
        for ch in list(body_pr):
            tag = etree.QName(ch).localname
            if tag in {"normAutofit", "spAutoFit", "noAutofit"}:
                body_pr.remove(ch)
        etree.SubElement(body_pr, f"{{{A_NS}}}spAutoFit")

        # Write the text. Multi-line: split on \n and emit one paragraph each.
        tf = box.text_frame
        tf.text = ""  # reset to a single empty paragraph
        for i, line in enumerate(text.split("\n")):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = Align.FOOTNOTE_H  # explicit; defaults can be inherited as center
            # Strip default run that python-pptx may have created
            for r in list(p.runs):
                p._p.remove(r._r)  # noqa: SLF001
            run = p.add_run()
            run.text = line
            run.font.size = Pt(Tokens.TYPE_FOOTNOTE)
            self.set_font_color(run.font, "dk1")
            # Force minor-font references; theme provides the actual typeface.
            rpr = run._r.find(f"{{{A_NS}}}rPr")  # noqa: SLF001
            if rpr is None:
                rpr = etree.SubElement(run._r, f"{{{A_NS}}}rPr")
                run._r.insert(0, rpr)
            for tag in ("ea", "cs"):
                for existing in rpr.findall(f"{{{A_NS}}}{tag}"):
                    rpr.remove(existing)
            etree.SubElement(rpr, f"{{{A_NS}}}ea", typeface="+mn-lt")
            etree.SubElement(rpr, f"{{{A_NS}}}cs", typeface="+mn-lt")
        return box

    @staticmethod
    def estimate_conclusion_height(text: str) -> float:
        """Estimate conclusion box height in inches from the number of lines the
        text wraps to (real font metrics when available). The box bottom is pinned
        just above the footnote and the cell text is bottom-anchored, so the band
        grows UPWARD from its pinned bottom; an exact line count keeps that bottom
        edge stable. Single-line floor is CONCLUSION_HEIGHT."""
        if not text:
            return 0.0
        if not ADJUST_CONCLUSION_HEIGHT:
            return CONCLUSION_HEIGHT          # fixed single-line height (adjustment off)
        # 18pt major-font text in a cell inset by SPACE_SHAPE_MARGIN each side.
        usable = FALLBACK_CONTENT.w - 2 * Tokens.SPACE_SHAPE_MARGIN
        lines = _wrapped_line_count(
            text, size_pt=Tokens.TYPE_CONCLUSION, kind="major",
            width_in=usable, chars_per_line=120)
        return max(CONCLUSION_HEIGHT, lines * CONCLUSION_LINE_HEIGHT)

    def add_conclusion(self, slide, text: str, has_footnote: bool = False,
                        footnote_height: float = 0.0,
                        baseline_y: float = FOOTER_BASELINE_Y):
        """Add a Conclusion: a 1×1 table with an amber left rule, dark bottom
        rule, transparent top rule, no right rule. Text is 18pt major-font,
        bottom-anchored within the cell.

        Position is computed so the conclusion's bottom sits 0.14" above the
        footnote's top when a footnote is present, or at the footer baseline
        when alone. Width matches the slide's content column (12.33").

        Tags the shape with name="Conclusion".
        """
        if not text:
            return None
        # Compute top y
        if has_footnote:
            footnote_top = baseline_y - footnote_height
            bottom_y = footnote_top - CONCLUSION_FOOTNOTE_GAP
        else:
            bottom_y = baseline_y
        # Height grows with the text so the bottom-anchored box expands UPWARD
        # from its pinned bottom edge (bottom_y), rather than overflowing.
        height = self.estimate_conclusion_height(text)
        top_y = bottom_y - height

        # Add a 1×1 table
        gf = slide.shapes.add_table(
            rows=1, cols=1,
            left=Inches(FALLBACK_CONTENT.x), top=Inches(top_y),
            width=Inches(FALLBACK_CONTENT.w), height=Inches(height),
        )
        # Pin the single row to the computed height so it doesn't auto-grow
        # downward when text wraps.
        try:
            gf.table.rows[0].height = Inches(height)
        except Exception:
            pass
        # Tag the graphicFrame
        nv = gf._element.find(f".//{{{P_NS}}}cNvPr")
        if nv is not None:
            nv.set("name", "Conclusion")

        # Style the cell: borders + bottom anchor + text
        cell = gf.table.cell(0, 0)
        tc = cell._tc  # noqa: SLF001
        tc_pr = tc.find(f"{{{A_NS}}}tcPr")
        if tc_pr is None:
            tc_pr = etree.SubElement(tc, f"{{{A_NS}}}tcPr")
        tc_pr.set("anchor", Align.CONCLUSION_ANCHOR_XML)
        cell.margin_top = 0
        cell.margin_bottom = 0
        cell.margin_left = Inches(Tokens.SPACE_SHAPE_MARGIN)
        cell.margin_right = Inches(Tokens.SPACE_SHAPE_MARGIN)

        # Strip any inherited borders, then build the four we care about.
        for ch in list(tc_pr):
            if etree.QName(ch).localname in {"lnL", "lnT", "lnR", "lnB",
                                              "fill", "solidFill", "noFill",
                                              "gradFill", "blipFill", "pattFill"}:
                tc_pr.remove(ch)
        # Borders must be added before fill in OOXML schema order... actually
        # tcPr's child order is: lnL, lnR, lnT, lnB, lnTlBr, lnBlToTr, cell3D,
        # fill (one of), headers, ... Strict order matters for some renderers.
        # We add borders first, then fill.
        self._add_cell_border(tc_pr, "lnL", width_emu=int(Tokens.STROKE_W * EMU_PER_PT),
                              color_kind="srgb", color_val=BRAND_COLORS["text_highlight"])
        self._add_cell_border(tc_pr, "lnT", width_emu=int(Tokens.STROKE_W * EMU_PER_PT),
                              color_kind="scheme", color_val="tx1", alpha="0")
        # No bottom rule — explicit noFill line so nothing is drawn or inherited.
        self._add_cell_border(tc_pr, "lnB", width_emu=int(Tokens.STROKE_W * EMU_PER_PT), color_kind="none",
                              color_val="")
        # No right border — match the reference

        # Override the fill so the inherited table style's header fill (navy
        # in OW) doesn't paint over our text. The reference has no explicit
        # fill but its table style happens to be one without a header fill;
        # we make our intent explicit.
        etree.SubElement(tc_pr, f"{{{A_NS}}}noFill")

        # Cell text
        tf = cell.text_frame
        tf.text = ""  # reset
        p = tf.paragraphs[0]
        # Force alignment on the paragraph (reference has no pPr so it
        # inherits left, but our table style might have inherited centered).
        p.alignment = Align.LEFT
        for r in list(p.runs):
            p._p.remove(r._r)  # noqa: SLF001
        run = p.add_run()
        run.text = text
        run.font.bold = False
        run.font.italic = False
        run.font.size = Pt(Tokens.TYPE_CONCLUSION)
        self.set_font_color(run.font, "dk1")
        # Force major-font (heading) references
        rpr = run._r.find(f"{{{A_NS}}}rPr")  # noqa: SLF001
        if rpr is None:
            rpr = etree.SubElement(run._r, f"{{{A_NS}}}rPr")
            run._r.insert(0, rpr)
        for tag in ("latin", "ea", "cs"):
            for existing in rpr.findall(f"{{{A_NS}}}{tag}"):
                rpr.remove(existing)
        etree.SubElement(rpr, f"{{{A_NS}}}latin", typeface="+mj-lt")
        etree.SubElement(rpr, f"{{{A_NS}}}ea",    typeface="+mj-lt")
        etree.SubElement(rpr, f"{{{A_NS}}}cs",    typeface="+mj-lt")
        return gf

    @staticmethod
    def _add_cell_border(tc_pr, tag: str, width_emu: int, color_kind: str,
                          color_val: str, alpha: Optional[str] = None) -> None:
        """Build one of <a:lnL>/<a:lnT>/<a:lnR>/<a:lnB>. ``color_kind="none"``
        produces a <a:noFill/> line (i.e. no visible border)."""
        ln = etree.SubElement(tc_pr, f"{{{A_NS}}}{tag}",
                               w=str(width_emu), cap="flat", cmpd="sng", algn="ctr")
        if color_kind == "none":
            etree.SubElement(ln, f"{{{A_NS}}}noFill")
            return
        sf = etree.SubElement(ln, f"{{{A_NS}}}solidFill")
        if color_kind == "srgb":
            etree.SubElement(sf, f"{{{A_NS}}}srgbClr", val=color_val)
        else:
            sc = etree.SubElement(sf, f"{{{A_NS}}}schemeClr", val=color_val)
            if alpha is not None:
                etree.SubElement(sc, f"{{{A_NS}}}alpha", val=alpha)
        etree.SubElement(ln, f"{{{A_NS}}}prstDash", val="solid")
        etree.SubElement(ln, f"{{{A_NS}}}round")

    # ---- low-level OOXML binding ---------------------------------------

    def _bind_to_placeholder(self, graphic_shape, placeholder) -> None:
        ph_el = placeholder._element  # noqa: SLF001
        gf_el = graphic_shape._element  # noqa: SLF001
        old_ph = ph_el.find(".//p:nvPr/p:ph", NSMAP)
        gf_nvpr = gf_el.find(".//p:nvPr", NSMAP)
        if old_ph is None:
            raise RuntimeError("Original placeholder has no <p:ph>; cannot bind.")
        if gf_nvpr is None:
            raise RuntimeError("graphicFrame has no <p:nvPr>; namespace failed.")
        for existing in gf_nvpr.findall("p:ph", NSMAP):
            gf_nvpr.remove(existing)
        gf_nvpr.append(deepcopy(old_ph))
        old_cnvpr = ph_el.find(".//p:cNvPr", NSMAP)
        gf_cnvpr = gf_el.find(".//p:cNvPr", NSMAP)
        if old_cnvpr is not None and gf_cnvpr is not None and "name" in old_cnvpr.attrib:
            gf_cnvpr.set("name", old_cnvpr.attrib["name"])
        if gf_el.find(".//p:nvPr/p:ph", NSMAP) is None:
            raise RuntimeError("Placeholder binding verification failed; aborting.")
        ph_el.getparent().remove(ph_el)
