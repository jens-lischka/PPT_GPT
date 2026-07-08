"""Semantic Slide Compiler.

Translates ``{intent, content}`` into a CompilationPlan that the runtime can
execute. The compiler:

  1. Looks up the intent in the INTENTS registry.
  2. Scores every layout in the template against the intent's slot signature
     and preferred-layout hints; picks the highest-scoring layout.
  3. Walks the intent's slot list and pairs each filled slot with a
     placeholder target (a role hint like ``primary`` / ``slot_3``) plus
     a strategy (a small render function from STRATEGIES).
  4. Emits a CompilationPlan (a pure data record) describing what to do.

The compiler invariants:

  - It never generates geometry. Geometry comes from the template's layouts.
  - It never embeds slide-type-specific code. Intents are *data*, scoring is
    a generic function, plan execution dispatches via the strategy registry.
  - Adding a new intent is registry data. Adding a new render kind is one
    function plus one entry in STRATEGIES.
"""
from __future__ import annotations

import logging
import math as _m
from dataclasses import dataclass
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

from pptx.enum.shapes import PP_PLACEHOLDER

from runtime import (TemplateRuntime, EMU_PER_INCH, iter_bullets, GRID_GAP_LG,
                     Tokens, LCID_MAP, detect_language)

logger = logging.getLogger("compiler")
if not logger.handlers:
    h = logging.StreamHandler()
    h.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))
    logger.addHandler(h)
logger.setLevel(logging.INFO)


# ===========================================================================
# Adornment policy
#
# Hard rule: these layouts must NEVER carry a Conclusion or a Footnote, no
# matter what the deck spec provides. Cover/structural/legal pages stay clean.
# Matched case-insensitively on the layout name; common spelling variants are
# included so the rule holds across template naming.
# ===========================================================================
NO_ADORNMENT_LAYOUTS = frozenset({
    "title slide",
    "section",
    "title slide with picture",
    "contents", "table of contents", "agenda",
    "confidentiality",
    "qualification", "qualifications",
    "backcover", "back cover", "back-cover",
})


# ===========================================================================
# Content kinds — the vocabulary used by both intents (what they consume) and
# layouts (what they support). Each strategy renders exactly one content kind.
# ===========================================================================

STRATEGY_TO_KIND: Dict[str, str] = {
    "title":             "text",
    "subtitle":          "text",
    "plain_text":        "text",
    "big_text":          "text",
    "huge_text":         "text",
    "bullets":           "bullets",
    "kpi_card":          "kpi",
    "insight_card":      "insight",
    "chart_native":      "chart",
    "chart_floating":    "chart",
    "table_native":      "table",
    "image":             "image",
    "process_chevrons":  "process",
    "org_chart":         "org_chart",
    "pyramid":           "pyramid",
    "cycle":             "cycle",
    "matrix":            "matrix",
    "timeline":          "timeline",
    "growing_steps":     "growing_steps",
    "contents":          "contents",
    "columns":           "columns",
    "quote":             "quote",
    "graphic_text":      "graphic_text",
    "waterfall":         "waterfall",
    "treemap":           "treemap",
    "compose":           "compose",
    "stat_callout":      "stat_callout",
    "icon_rows":         "icon_rows",
    "compare":           "compare",
    "sticker":           "sticker",
    "infographic":       "infographic",
}


# ===========================================================================
# Layout Capability Registry
#
# Each layout in the OW template advertises which content *kinds* it is
# designed to host. This is a contract between template authors and the
# compiler:
#
#   - A layout's ``supports`` set is the *closed* list of kinds the compiler
#     may route to it. An intent that needs a kind not in ``supports`` will
#     be hard-rejected at scoring time, even if the layout has enough slots.
#   - The set is intentionally small (text / bullets / chart / table / kpi /
#     insight / image). Adding a new kind = adding a strategy + a mapping
#     in STRATEGY_TO_KIND + adding it to the relevant layouts' supports.
#
# Layouts not present in the registry default to a "permissive" capability
# (anything goes), so the compiler still works on unfamiliar templates.
# ===========================================================================

@dataclass(frozen=True)
class LayoutCapability:
    supports: frozenset
    notes: str = ""


LAYOUT_CAPABILITIES: Dict[str, LayoutCapability] = {
    "Title Slide":         LayoutCapability(
        supports=frozenset({"text"}),
        notes="Cover slide; title + short subtitle text."),
    "Title Slide with Picture": LayoutCapability(
        supports=frozenset({"text", "image"}),
        notes="Cover slide variant with a wide picture band."),
    "Title and Content":   LayoutCapability(
        supports=frozenset({"text", "bullets", "chart", "table", "image"}),
        notes="Single full-width content slot; the workhorse layout."),
    "Contents":            LayoutCapability(supports=frozenset({"contents", "text"})),
    "Title Only":          LayoutCapability(
        supports=frozenset({"text", "process", "org_chart", "pyramid", "cycle", "matrix", "timeline", "growing_steps", "contents", "columns", "quote", "graphic_text", "waterfall", "treemap", "compose", "stat_callout", "icon_rows", "compare", "sticker", "infographic"}),
        notes="Universal fallback + canvas for engine-composed graphics."),
    "2 columns":           LayoutCapability(
        supports=frozenset({"text", "bullets"}),
        notes="Two equal columns of text content."),
    "3 columns":           LayoutCapability(
        supports=frozenset({"text", "bullets", "kpi"}),
        notes="Three equal columns; supports KPI grids of three."),
    "4 columns":           LayoutCapability(
        supports=frozenset({"text", "bullets", "kpi"}),
        notes="Four equal columns; ideal for KPI dashboards."),
    "2 columns 1/3 split": LayoutCapability(
        supports=frozenset({"text", "bullets", "chart", "table", "insight"}),
        notes="Narrow left + wide right; insight on left, chart/table on right."),
    "2 columns 2/3 split": LayoutCapability(
        supports=frozenset({"text", "bullets", "chart", "table", "insight"}),
        notes="Wide left + narrow right; chart/table on left, insight on right."),
    "Section":             LayoutCapability(
        supports=frozenset({"text"}),
        notes="Section divider; two BODY placeholders, no TITLE."),
    "CV - bio with photo": LayoutCapability(
        supports=frozenset({"text", "bullets", "image"}),
        notes="Person bio: name, role, headshot, and bio paragraph."),
    "Big portrait photo on right": LayoutCapability(
        supports=frozenset({"text", "bullets", "image"}),
        notes="Wide portrait on the right with text content on the left."),
}


def _required_kinds(intent: "Intent") -> frozenset:
    """Set of content kinds this intent must place. Optional slots are not
    counted — the compiler only hard-rejects layouts on required kinds."""
    return frozenset(
        STRATEGY_TO_KIND[s.strategy]
        for s in intent.slots
        if s.required and s.strategy in STRATEGY_TO_KIND
    )


def _layout_supports(layout_name: str, kinds: frozenset) -> bool:
    """True iff the layout supports every kind in ``kinds``. Layouts not in
    the registry default to permissive (return True), so unknown templates
    still work."""
    cap = LAYOUT_CAPABILITIES.get(layout_name)
    if cap is None:
        return True
    return kinds.issubset(cap.supports)


# ===========================================================================
# Intent data model
# ===========================================================================

@dataclass(frozen=True)
class IntentSlot:
    """One slot in an intent's signature.

    ``target`` selects which placeholder receives the slot's content. Two
    role families are supported:

      Position-based (top-to-bottom, left-to-right):
        - ``"title"``      → the layout's TITLE/CENTER_TITLE placeholder
        - ``"primary"``    → first content placeholder (top-left)
        - ``"secondary"``  → second content placeholder
        - ``"tertiary"``   → third
        - ``"quaternary"`` → fourth
        - ``"slot_N"``     → Nth content placeholder (0-indexed)
        - ``"*"``          → variadic; fan out content list across all slots

      Area-based (largest first):
        - ``"largest"``        → biggest content slot by area
        - ``"second_largest"`` → next biggest

    Use position-based for grids and ordered layouts (dashboard, section).
    Use area-based when the slot's *size* matters semantically (chart in the
    bigger slot, callout in the smaller).

    ``strategy`` names the leaf renderer; see STRATEGIES below.
    """
    name: str
    strategy: str
    target: str
    required: bool = True


@dataclass(frozen=True)
class Intent:
    """A semantic communication intent."""
    name: str
    description: str
    slots: Tuple[IntentSlot, ...]
    preferred_layouts: Tuple[str, ...] = ()
    fallback_layouts: Tuple[str, ...] = ("Title Only",)


# ---------------------------------------------------------------------------
# Intent registry — phase 8 starter set.
# ---------------------------------------------------------------------------
INTENTS: Dict[str, Intent] = {
    "introduce_topic": Intent(
        name="introduce_topic",
        description="Cover slide with a title, optional subtitle, and optional cover image.",
        slots=(
            IntentSlot("title",    strategy="title",    target="title"),
            IntentSlot("subtitle", strategy="subtitle", target="primary", required=False),
            IntentSlot("image",    strategy="image",    target="picture", required=False),
        ),
        preferred_layouts=("Title Slide with Picture", "Title Slide"),
        fallback_layouts=("Title Only",),
    ),

    "explain": Intent(
        name="explain",
        description="A title and a single body of text or bullets.",
        slots=(
            IntentSlot("title", strategy="title",   target="title"),
            IntentSlot("body",  strategy="bullets", target="primary"),
        ),
        preferred_layouts=("Title and Content",),
        fallback_layouts=("Title Only",),
    ),

    "compare_two_options": Intent(
        name="compare_two_options",
        description="Side-by-side comparison of two columns of text.",
        slots=(
            IntentSlot("title", strategy="title",   target="title"),
            IntentSlot("left",  strategy="bullets", target="primary"),
            IntentSlot("right", strategy="bullets", target="secondary"),
        ),
        preferred_layouts=("2 columns",),
        fallback_layouts=("Title and Content", "Title Only"),
    ),

    "show_trend_with_key_message": Intent(
        name="show_trend_with_key_message",
        description="A chart in the larger slot, an insight callout in the smaller.",
        slots=(
            IntentSlot("title",   strategy="title",         target="title"),
            IntentSlot("chart",   strategy="chart_native",  target="largest"),
            IntentSlot("insight", strategy="insight_card",  target="second_largest"),
        ),
        preferred_layouts=("2 columns 2/3 split", "2 columns 1/3 split"),
        fallback_layouts=("Title and Content", "Title Only"),
    ),

    "show_waterfall": Intent(
        name="show_waterfall",
        description="Waterfall / bridge chart: absolute totals from zero, positive/negative deltas floating on a transparent base. content {title, waterfall:{orientation:'vertical'|'horizontal', categories:[...], values:[...], totals:[indices of absolute totals], heading?, subheading?}}.",
        slots=(
            IntentSlot("title",     strategy="title",     target="title"),
            IntentSlot("waterfall", strategy="waterfall", target="canvas"),
        ),
        preferred_layouts=("Title Only",),
        fallback_layouts=("Title Only",),
    ),

    "show_treemap": Intent(
        name="show_treemap",
        description="Treemap: rectangles sized by value (squarified), labelled with name + value, white seams. content {title, treemap:{items:[{label, value}], heading?, subheading?}}. Use for part-to-whole composition across many categories.",
        slots=(
            IntentSlot("title",   strategy="title",   target="title"),
            IntentSlot("treemap", strategy="treemap", target="canvas"),
        ),
        preferred_layouts=("Title Only",),
        fallback_layouts=("Title Only",),
    ),

    "compose": Intent(
        name="compose",
        description="Open composition (escape hatch): place heterogeneous blocks into a grid of regions. block kinds: text, kpi, table, image, chart (any type: column/bar/line/pie/doughnut/area/scatter), waterfall, treemap, or any diagram (cycle/pyramid/process/org_chart/matrix/timeline/growing_steps). content {title, compose:{rows:[{height?, regions:[{block, data, width?}]}]}} or {title, compose:{regions:[...]}} for a single row. For a flexible grid with spanning use content {title, compose:{grid:{cols?, rows?, cells:[{block, data, col?, row?, colspan?, rowspan?}]}}} — equal tracks flex to fill on a 0.5\" gutter (0.25\" when narrow), cells span colspan x rowspan, cols/rows are solved for the smallest grid (rows minimised first) when omitted. Use ONLY when no preset intent fits; the fit router prefers a named preset whenever one matches.",
        slots=(
            IntentSlot("title",   strategy="title",   target="title"),
            IntentSlot("compose", strategy="compose", target="canvas"),
        ),
        preferred_layouts=("Title Only",),
        fallback_layouts=("Title Only",),
    ),

    "stat_callout": Intent(
        name="stat_callout",
        description="Big-number stat cards: each {value, heading, text} on a cream card. content {title, stats:[...]} (1-4).",
        slots=(
            IntentSlot("title", strategy="title",        target="title"),
            IntentSlot("stats", strategy="stat_callout", target="canvas"),
        ),
        preferred_layouts=("Title Only",), fallback_layouts=("Title Only",),
    ),
    "icon_rows": Intent(
        name="icon_rows",
        description="Stacked rows of fine-line icon + heading + text. content {title, rows:[{icon, heading, text}]} (1-5).",
        slots=(
            IntentSlot("title", strategy="title",     target="title"),
            IntentSlot("rows",  strategy="icon_rows", target="canvas"),
        ),
        preferred_layouts=("Title Only",), fallback_layouts=("Title Only",),
    ),
    "compare": Intent(
        name="compare",
        description="Two-column pros/cons with check (green)/cross (red). content {title, compare:{pros:[...], cons:[...], pros_heading?, cons_heading?}}.",
        slots=(
            IntentSlot("title",   strategy="title",   target="title"),
            IntentSlot("compare", strategy="compare", target="canvas"),
        ),
        preferred_layouts=("Title Only",), fallback_layouts=("Title Only",),
    ),
    "sticker": Intent(
        name="sticker",
        description="An amber circular highlight with short text + optional note. content {title, sticker:{text, note?}}.",
        slots=(
            IntentSlot("title",   strategy="title",   target="title"),
            IntentSlot("sticker", strategy="sticker", target="canvas"),
        ),
        preferred_layouts=("Title Only",), fallback_layouts=("Title Only",),
    ),
    "infographic": Intent(
        name="infographic",
        description="Native-freeform infographic. content {title, infographic:{type:funnel|gauge|venn|heatmap, ...}}. funnel:{stages:[{label,value}]}; gauge:{value,label,suffix?}; venn:{sets:[{label}]}; heatmap:{values:[[..]], x_labels?, y_labels?, steps?, legend_labels?, description?}.",
        slots=(
            IntentSlot("title",       strategy="title",       target="title"),
            IntentSlot("infographic", strategy="infographic", target="canvas"),
        ),
        preferred_layouts=("Title Only",), fallback_layouts=("Title Only",),
    ),

    "dashboard": Intent(
        name="dashboard",
        description="Grid of KPI cards (variadic).",
        slots=(
            IntentSlot("title",   strategy="title",    target="title"),
            IntentSlot("metrics", strategy="kpi_card", target="*"),
        ),
        preferred_layouts=("4 columns", "3 columns"),
        fallback_layouts=("2 columns", "Title and Content"),
    ),

    "section_divider": Intent(
        name="section_divider",
        description="Section break with an optional ordinal number.",
        slots=(
            # The OW Section layout has no TITLE — the section title text
            # lives in the first BODY placeholder ('SectionTitle'), and the
            # ordinal lives in the second ('SectionNumber'). target="primary"
            # works on both Section (BODY[0]) and the Title Only fallback
            # (degraded to title placeholder).
            IntentSlot("title",  strategy="big_text",  target="primary"),
            IntentSlot("number", strategy="huge_text", target="secondary", required=False),
        ),
        preferred_layouts=("Section",),
        fallback_layouts=("Title Only",),
    ),

    "show_data": Intent(
        name="show_data",
        description="Tabular data presentation.",
        slots=(
            IntentSlot("title", strategy="title",        target="title"),
            IntentSlot("table", strategy="table_native", target="primary"),
        ),
        preferred_layouts=("Title and Content",),
        fallback_layouts=("Title Only",),
    ),

    "summary": Intent(
        name="summary",
        description="Executive summary — title plus closing body content.",
        slots=(
            IntentSlot("title", strategy="title",   target="title"),
            IntentSlot("body",  strategy="bullets", target="primary"),
        ),
        preferred_layouts=("Title and Content",),
        fallback_layouts=("Title Only",),
    ),

    "introduce_person": Intent(
        name="introduce_person",
        description="Person bio: name in title, role text, headshot, and bio paragraph.",
        slots=(
            IntentSlot("name",  strategy="title",      target="title"),
            IntentSlot("role",  strategy="plain_text", target="body",    required=False),
            IntentSlot("photo", strategy="image",      target="picture", required=False),
            IntentSlot("bio",   strategy="bullets",    target="object",  required=False),
        ),
        preferred_layouts=("CV - bio with photo",),
        fallback_layouts=("Title and Content", "Title Only"),
    ),

    "show_process": Intent(
        name="show_process",
        description="Linear process flow rendered as horizontal chevrons.",
        slots=(
            IntentSlot("title", strategy="title",            target="title"),
            IntentSlot("steps", strategy="process_chevrons", target="canvas"),
        ),
        preferred_layouts=("Title Only",),
        fallback_layouts=("Title Only",),
    ),

    "show_org_chart": Intent(
        name="show_org_chart",
        description="Two-level organizational chart: root + direct reports.",
        slots=(
            IntentSlot("title",     strategy="title",     target="title"),
            IntentSlot("org_chart", strategy="org_chart", target="canvas"),
        ),
        preferred_layouts=("Title Only",),
        fallback_layouts=("Title Only",),
    ),

    "show_pyramid": Intent(
        name="show_pyramid",
        description="Basic pyramid: stacked trapezoid bands, narrow at top.",
        slots=(
            IntentSlot("title",  strategy="title",   target="title"),
            IntentSlot("levels", strategy="pyramid", target="canvas"),
        ),
        preferred_layouts=("Title Only",),
        fallback_layouts=("Title Only",),
    ),

    "show_cycle": Intent(
        name="show_cycle",
        description="Cyclical/iterative relationship: labeled ovals on a circle with arrows. Use only for genuine loops, not to list parallel topics.",
        slots=(
            IntentSlot("title", strategy="title", target="title"),
            IntentSlot("items", strategy="cycle", target="canvas"),
        ),
        preferred_layouts=("Title Only",),
        fallback_layouts=("Title Only",),
    ),

    "show_matrix": Intent(
        name="show_matrix",
        description="Matrix grid (any rows×cols) with items placed across two axes.",
        slots=(
            IntentSlot("title",  strategy="title",  target="title"),
            IntentSlot("matrix", strategy="matrix", target="canvas"),
        ),
        preferred_layouts=("Title Only",),
        fallback_layouts=("Title Only",),
    ),

    "show_timeline": Intent(
        name="show_timeline",
        description="Horizontal timeline with any number of milestones.",
        slots=(
            IntentSlot("title",      strategy="title",    target="title"),
            IntentSlot("milestones", strategy="timeline", target="canvas"),
        ),
        preferred_layouts=("Title Only",),
        fallback_layouts=("Title Only",),
    ),

    "show_growing_steps": Intent(
        name="show_growing_steps",
        description="Ascending growing steps (staircase) with any number of steps.",
        slots=(
            IntentSlot("title", strategy="title",         target="title"),
            IntentSlot("steps", strategy="growing_steps", target="canvas"),
        ),
        preferred_layouts=("Title Only",),
        fallback_layouts=("Title Only",),
    ),

    "show_contents": Intent(
        name="show_contents",
        description="Table of contents (native 3-column table: number | section name | page). content {sections:[{title, page?, number?}]}; number auto-fills 01,02,... (set \"\" to omit, e.g. Appendix). Always on the Contents layout; one per deck.",
        slots=(
            IntentSlot("title",    strategy="title",    target="title"),
            IntentSlot("sections", strategy="contents", target="canvas"),
        ),
        preferred_layouts=("Contents",),
        fallback_layouts=("Contents", "Title Only"),
    ),

    "show_columns": Intent(
        name="show_columns",
        description="2-5 parallel headed text columns (strengths/weaknesses, pillars, profile facets). Use for parallel categories instead of one flat bulleted body. 2-4 columns render in the template's native column layouts; 5 falls back to a manual canvas layout.",
        slots=(
            IntentSlot("title",   strategy="title",   target="title"),
            IntentSlot("columns", strategy="bullets", target="*"),
        ),
        preferred_layouts=("2 columns", "3 columns", "4 columns"),
        fallback_layouts=("Title Only",),
    ),

    "show_graphic_with_text": Intent(
        name="show_graphic_with_text",
        description="Composite: a graphic (growing_steps/pyramid/cycle/org_chart/process/matrix/timeline) in one pane and supporting text in the other. content {panel:{graphic:{kind,...}, text:{heading,subheading,bullets|paragraphs}, side:'graphic_left'|'graphic_right', split:'equal'(default)|'graphic_2_3'|'text_2_3'}}. Equal columns by default on the 0.5-inch grid.",
        slots=(
            IntentSlot("title", strategy="title", target="title"),
            IntentSlot("panel", strategy="graphic_text", target="canvas"),
        ),
        preferred_layouts=("Title Only",),
        fallback_layouts=("Title Only",),
    ),

    "show_quote": Intent(
        name="show_quote",
        description="Big editorial pull-quote (voice of a customer/leader/expert). content {quote:{text, attribution}} or {quote:'...'}. Use for emphasis or a reflective pause in a narrative, not for data; do not bury verbatim quotes in slide titles.",
        slots=(
            IntentSlot("quote", strategy="quote", target="canvas"),
        ),
        preferred_layouts=("Blank",),
        fallback_layouts=("Blank",),
    ),
}


# ===========================================================================
# Layout shape analysis (template introspection)
# ===========================================================================

@dataclass(frozen=True)
class ContentSlotShape:
    ordinal: int
    ph_idx: int
    ph_type: str
    geometry: Tuple[float, float, float, float]  # x, y, w, h in inches


@dataclass(frozen=True)
class LayoutShape:
    name: str
    has_title: bool
    content_slots: Tuple[ContentSlotShape, ...]


def analyze_template(runtime: TemplateRuntime) -> Dict[str, LayoutShape]:
    """Inspect each layout in the template, sort its content placeholders
    top-to-bottom / left-to-right, and produce a LayoutShape record.

    Footer / date / slide-number placeholders are excluded; only content-bearing
    slots remain. The compiler reasons over LayoutShape, never over the live
    Presentation object.
    """
    shapes: Dict[str, LayoutShape] = {}
    SKIP = {PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER,
            PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.HEADER}

    for layout in runtime.prs.slide_layouts:
        has_title = False
        content_phs = []
        for ph in layout.placeholders:
            try:
                ph_type = ph.placeholder_format.type
            except Exception:
                continue
            if ph_type in {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE}:
                has_title = True
                continue
            if ph_type in SKIP:
                continue
            x = (ph.left or 0) / EMU_PER_INCH
            y = (ph.top  or 0) / EMU_PER_INCH
            w = (ph.width  or 0) / EMU_PER_INCH
            h = (ph.height or 0) / EMU_PER_INCH
            content_phs.append((ph.placeholder_format.idx, ph_type.name, x, y, w, h))
        # Stable sort: top-to-bottom, then left-to-right
        content_phs.sort(key=lambda r: (r[3], r[2]))
        slots = tuple(
            ContentSlotShape(ordinal=i, ph_idx=idx, ph_type=ptype, geometry=(x, y, w, h))
            for i, (idx, ptype, x, y, w, h) in enumerate(content_phs)
        )
        shapes[layout.name] = LayoutShape(name=layout.name, has_title=has_title, content_slots=slots)
    return shapes


# ===========================================================================
# Layout scoring & selection
# ===========================================================================

def _content_targets(intent: Intent) -> Tuple[int, bool]:
    """How many content placeholders does this intent consume?

    Returns ``(non_variadic_count, has_variadic)``. Canvas-targeted slots
    don't consume placeholders (they render directly on the slide), so
    they're excluded from the count.
    """
    n = 0
    variadic = False
    for s in intent.slots:
        if s.target == "title":
            continue
        if s.target == "canvas":
            continue
        if s.target == "*":
            variadic = True
            continue
        n += 1
    return n, variadic


def score_layout(layout: LayoutShape, intent: Intent, item_count: int = None) -> int:
    """How well does this layout fit this intent?

    Rules in priority order:
      - Hard filter A: layout must support every required content kind.
        (Inferred from intent's strategies; checked against
        LAYOUT_CAPABILITIES.) If not, score is -infinity.
      - Hard filter B: layout must have enough content slots; -infinity if not.
      - Bonus for being in preferred_layouts (with order acting as tiebreaker).
      - Smaller bonus for being in fallback_layouts.
      - Bonus for exact slot count match.
      - Small penalty per excess content slot (waste).
      - Bonus for title presence when the intent wants a title.
    """
    # Hard filter A — capability support.
    needed_kinds = _required_kinds(intent)
    if not _layout_supports(layout.name, needed_kinds):
        return -10_000

    # Hard filter B — slot count.
    needed, variadic = _content_targets(intent)
    available = len(layout.content_slots)
    if available < needed:
        return -10_000  # incompatible

    score = 0
    # Preferred-list bonus, with order acting as tiebreaker (earlier = higher).
    if layout.name in intent.preferred_layouts:
        idx = intent.preferred_layouts.index(layout.name)
        score += 100 - idx
    if layout.name in intent.fallback_layouts:
        idx = intent.fallback_layouts.index(layout.name)
        score += 25 - idx

    if not variadic:
        if available == needed:
            score += 40
        else:
            score -= (available - needed) * 4
    else:
        # Variadic intents. When we know the item count, strongly prefer the
        # layout whose slot count *matches* (e.g. 3 columns -> "3 columns"); a
        # layout with fewer slots still scores (dashboard overflow splits across
        # slides), one with more is mild waste.
        if item_count is not None:
            if available == item_count:
                score += 60
            elif available > item_count:
                score += max(0, 30 - (available - item_count) * 8)
            else:
                score += min(available, 6) * 6
        else:
            score += min(available, 6) * 6

    intent_wants_title = any(s.target == "title" for s in intent.slots)
    if intent_wants_title and layout.has_title:
        score += 10
    return score


def select_layout(intent: Intent, layouts: Dict[str, LayoutShape], item_count: int = None) -> Tuple[LayoutShape, str]:
    """Pick the highest-scoring layout for this intent.

    Returns ``(layout_shape, reason)`` where reason is a short string for logs.
    Falls back to 'Title Only' if no layout has a positive score.
    """
    scored = [(score_layout(L, intent, item_count), L) for L in layouts.values()]
    scored.sort(key=lambda t: t[0], reverse=True)
    best_score, best = scored[0]
    if best_score <= -10_000:
        if "Title Only" in layouts:
            return layouts["Title Only"], "no compatible layout; fell back to Title Only"
        raise RuntimeError("Template has no 'Title Only' fallback layout.")
    return best, f"score={best_score}"


# ===========================================================================
# Plan emission
# ===========================================================================

# ===========================================================================
# Density evaluation
#
# Given a content payload and the geometry of the slot it would land in,
# estimate whether it overflows. The compiler uses this to decide if a slide
# should be split across multiple plans.
#
# Constants are deliberately rough — the goal is "should we split?", not
# pixel-perfect typesetting. They're tuned to OW body styling (~14pt body,
# ~12pt bullets) at 16:9 dimensions. Templates with tighter or looser
# typography can adjust these as a single source of truth.
# ===========================================================================

LINE_HEIGHT_IN          = 0.45   # body line height incl. paragraph spacing
CHARS_PER_INCH_LINE     = 12.0   # rough cps at OW body sizes
TABLE_FIT_ROW_IN        = 0.32   # realistic per-DATA-row height for fit (render min is 0.30)
TABLE_FIT_SAFETY_IN     = 0.12   # bottom margin + small wrap buffer reserved below the table
TABLE_COL_WIDTH_IN      = 1.4    # comfortable min column width

# Score thresholds:
#   < 1.0 → fits comfortably
#   1.0–1.5 → tight; tighten if you can
#   > 1.5 → split
DENSITY_SPLIT_THRESHOLD = 1.0   # for tables, even slight overflow → split (rows are atomic)
DENSITY_TARGET_PER_PAGE = 0.9   # when splitting, aim for this density per page


@dataclass(frozen=True)
class DensityFit:
    kind: str
    score: float       # ≥0; 1.0 = exactly full
    bottleneck: str    # what's limiting it ('lines', 'rows', 'cols', 'items', 'n/a')
    suggestion: str    # 'fits' | 'tighten' | 'split'
    capacity: int = 0  # for tables: max DATA rows that fit one page (greedy paging)


try:
    import textmetrics as _tm            # Tier-1 real-font measurement
except Exception:                        # pragma: no cover
    _tm = None

# Measured wrapping changes line *counts* and therefore split *decisions*. To
# avoid silently shifting the calibrated production behaviour, it is opt-in:
# set OW_MEASURED_WRAP=1 to drive splits from real-font measurement. The
# BoundHeight-style overflow *verification* (textmetrics.overflow_report) runs
# independently and always uses real metrics.
_USE_MEASURED = bool(_os.environ.get("OW_MEASURED_WRAP")) if (_os := __import__("os")) else False


def _estimate_text_lines(text: str, width_in: float) -> int:
    """How many wrapped lines a string produces in a column ``width_in`` inches
    wide, at OW body sizes. Opt-in real-font wrapping (OW_MEASURED_WRAP) uses
    true glyph advances; the default keeps the calibrated chars-per-inch
    heuristic so existing split decisions are unchanged."""
    if not text:
        return 1
    if _USE_MEASURED and _tm is not None and _tm.available():
        return _tm.wrapped_lines(text, width_in, size_pt=12.0, kind="minor")
    chars_per_line = max(1, int(width_in * CHARS_PER_INCH_LINE))
    return max(1, _m.ceil(len(text) / chars_per_line))


def evaluate_density(kind: str,
                     slot_geom: Tuple[float, float, float, float],
                     payload: Any) -> DensityFit:
    """Estimate overflow score for one piece of content in one slot.

    Slot geometry is ``(x, y, w, h)`` in inches. Score 1.0 = exactly fills
    the slot; >1.0 = overflows by that ratio.
    """
    x, y, w, h = slot_geom
    if w <= 0 or h <= 0:
        return DensityFit(kind=kind, score=0.0, bottleneck="no-geom", suggestion="fits")

    if kind == "bullets":
        max_lines = max(1, int(h / LINE_HEIGHT_IN))
        if not isinstance(payload, dict):
            return DensityFit(kind=kind, score=0.0, bottleneck="n/a", suggestion="fits")
        heading = payload.get("heading", "")
        prose = payload.get("paragraphs")
        if prose is None and payload.get("format") == "prose":
            prose = payload.get("bullets", [])
        if prose is not None:
            # Prose: each paragraph wraps to several lines, plus ~0.4 line of
            # inter-paragraph spacing between them.
            used = ((1 if heading else 0)
                    + sum(_estimate_text_lines(p, w) for p in prose)
                    + max(0, len(prose) - 1) * 0.4)
        else:
            bullets = payload.get("bullets", [])
            used = (1 if heading else 0) + sum(
                _estimate_text_lines(text, w) for _lvl, text in iter_bullets(bullets))
        score = used / max_lines
        return DensityFit(
            kind="bullets", score=score, bottleneck="lines",
            suggestion=("split"   if score > 1.5 else
                        "tighten" if score > 1.0 else
                        "fits"),
        )

    if kind == "table":
        if not isinstance(payload, dict):
            return DensityFit(kind=kind, score=0.0, bottleneck="n/a", suggestion="fits")
        data_rows = len(payload.get("rows", []))
        n_cols = len(payload.get("headers", []))
        # Capacity = how many DATA rows actually fit, from the real render heights
        # (header SPACE_TABLE_HEADER + N body rows of SPACE_TABLE_ROW), minus a
        # small safety reserve. This is what lets a slide carry the MAXIMUM rows
        # rather than under-filling on a conservative estimate.
        usable = max(0.0, h - Tokens.SPACE_TABLE_HEADER - TABLE_FIT_SAFETY_IN)
        row_h = max(Tokens.SPACE_TABLE_ROW, TABLE_FIT_ROW_IN)
        capacity = max(1, int(usable / row_h))
        max_cols = max(1, int(w / TABLE_COL_WIDTH_IN))
        row_score = data_rows / capacity
        col_score = n_cols / max_cols if n_cols else 0.0
        score = max(row_score, col_score)
        bottleneck = "rows" if row_score >= col_score else "cols"
        return DensityFit(
            kind="table", score=score, bottleneck=bottleneck,
            suggestion="split" if score > DENSITY_SPLIT_THRESHOLD else "fits",
            capacity=capacity,
        )

    # Other kinds (chart, kpi, insight, text) don't overflow by content size
    # in interesting ways — chart shapes scale to the slot, KPIs are a
    # fixed-shape card per slot (handled separately as variadic), insight is
    # short by convention.
    return DensityFit(kind=kind, score=0.0, bottleneck="n/a", suggestion="fits")


def split_payload(kind: str, payload: Any, n_chunks: int) -> List[Any]:
    """Chunk a content payload into ``n_chunks`` pieces. Only ``bullets`` and
    ``table`` are splittable in v1; other kinds return [payload] unchanged."""
    if n_chunks <= 1:
        return [payload]

    if kind == "bullets" and isinstance(payload, dict):
        # Prose: split the paragraphs list, preserving the prose form.
        if payload.get("paragraphs"):
            paras = payload["paragraphs"]
            chunk_size = _m.ceil(len(paras) / n_chunks)
            chunks = []
            for i in range(n_chunks):
                sub = paras[i * chunk_size : (i + 1) * chunk_size]
                if not sub:
                    break
                chunks.append({**payload, "paragraphs": sub})
            return chunks or [payload]
        bullets = payload.get("bullets", [])
        if not bullets:
            return [payload]
        chunk_size = _m.ceil(len(bullets) / n_chunks)
        chunks = []
        for i in range(n_chunks):
            sub = bullets[i * chunk_size : (i + 1) * chunk_size]
            if not sub:
                break
            chunks.append({**payload, "bullets": sub})
        return chunks or [payload]

    if kind == "table" and isinstance(payload, dict):
        rows = payload.get("rows", [])
        if not rows:
            return [payload]
        chunk_size = _m.ceil(len(rows) / n_chunks)
        chunks = []
        for i in range(n_chunks):
            sub = rows[i * chunk_size : (i + 1) * chunk_size]
            if not sub:
                break
            chunks.append({**payload, "rows": sub})
        return chunks or [payload]

    return [payload]



@dataclass(frozen=True)
class PlaceholderInstruction:
    """One slot's worth of work: target a placeholder, run a strategy, with payload."""
    target: str       # role hint; resolved to ordinal at execute time
    strategy: str     # name in STRATEGIES
    payload: Any


@dataclass(frozen=True)
class CompilationPlan:
    intent: str
    layout_name: str
    title_text: Optional[str]
    instructions: Tuple[PlaceholderInstruction, ...]
    fallback_used: bool
    notes: Tuple[str, ...]
    # Optional bottom adornments. Independent of intent — apply to any slide.
    conclusion: Optional[str] = None
    footnote: Optional[str] = None

    def describe(self) -> str:
        lines = [f"intent={self.intent} → layout={self.layout_name!r} ({', '.join(self.notes)})"]
        if self.title_text:
            lines.append(f"  title: {self.title_text!r}")
        for instr in self.instructions:
            lines.append(f"  target={instr.target:>10}  strategy={instr.strategy}")
        if self.fallback_used:
            lines.append("  (fallback layout used)")
        return "\n".join(lines)


def _build_base_plan(
    intent: Intent,
    layout_shape: LayoutShape,
    content: Dict[str, Any],
    reason: str,
    fallback_used: bool,
) -> CompilationPlan:
    """Build a single CompilationPlan for the given (intent, layout, content)
    with no overflow handling. Variadic content is fanned out across
    available slots; if the content is longer than the slots can hold, the
    excess is silently dropped — overflow detection is the caller's job.
    """
    title_text: Optional[str] = None
    instructions: List[PlaceholderInstruction] = []
    n_slots = len(layout_shape.content_slots)

    for slot in intent.slots:
        slot_content = content.get(slot.name)
        if slot_content is None:
            if slot.required:
                logger.warning("Intent %r missing required slot %r.", intent.name, slot.name)
            continue

        if slot.target == "title":
            title_text = str(slot_content)
            continue

        if slot.target == "*":
            if not isinstance(slot_content, Sequence) or isinstance(slot_content, (str, bytes)):
                logger.warning("Variadic slot %r expects a list; got %s.",
                               slot.name, type(slot_content).__name__)
                continue
            for i, item in enumerate(slot_content[:n_slots]):
                instructions.append(PlaceholderInstruction(
                    target=f"slot_{i}", strategy=slot.strategy, payload=item,
                ))
            continue

        instructions.append(PlaceholderInstruction(
            target=slot.target, strategy=slot.strategy, payload=slot_content,
        ))

    return CompilationPlan(
        intent=intent.name,
        layout_name=layout_shape.name,
        title_text=title_text,
        instructions=tuple(instructions),
        fallback_used=fallback_used,
        notes=(reason,),
    )


def _split_for_variadic_overflow(
    intent: Intent,
    layout_shape: LayoutShape,
    content: Dict[str, Any],
    reason: str,
    fallback_used: bool,
) -> Optional[List[CompilationPlan]]:
    """If the intent has a variadic slot, decide whether to render in
    placeholders, in a synthetic grid, or split across pages.

    The OW preference is to render rich/overflowing card grids on a single
    Title Only slide rather than paginate across "(continued)" pages.
    The grid path is used when ANY of these are true (in priority order):

      1. ``content["grid"]`` is set — author asked for an explicit shape
         (``"2x2"``, ``"3x3"``, ``"3x2"``, ``"1x4"``, etc.).
      2. Any item has a ``chart`` field — single-text-frame placeholders
         can't host charts; the grid uses free shapes that can.
      3. ``len(items) > n_slots`` — the classic variadic overflow.

    Other variadic intents (long bullet lists, long tables) keep the
    per-page chunking behavior, since pagination is the right answer for
    prose. Returns None if no special routing is needed.
    """
    n_slots = len(layout_shape.content_slots)
    if n_slots == 0:
        return None
    for slot in intent.slots:
        if slot.target != "*":
            continue
        items = content.get(slot.name)
        if not isinstance(items, Sequence) or isinstance(items, (str, bytes)):
            continue

        # Trigger detection — three independent signals.
        explicit_grid = content.get("grid")
        has_rich_cells = any(
            isinstance(it, dict) and isinstance(it.get("chart"), dict) and it["chart"]
            for it in items
        )
        is_overflow = len(items) > n_slots

        if not (explicit_grid or has_rich_cells or is_overflow):
            return None  # fits in placeholders cleanly

        # ----------------------------------------------------------------
        # Strategy A — synthetic grid on Title Only.
        # Used for ``kpi_card`` slots in any of the above conditions.
        # ----------------------------------------------------------------
        if slot.strategy == "kpi_card":
            payload: Dict[str, Any] = {"items": list(items)}
            if explicit_grid:
                payload["grid"] = explicit_grid
            grid_instr = PlaceholderInstruction(
                target="canvas",
                strategy="kpi_card_grid",
                payload=payload,
            )
            title_text = content.get("title")
            # Build a human-readable note explaining why we routed to grid.
            if explicit_grid:
                why = f"explicit grid={explicit_grid!r}"
            elif has_rich_cells:
                why = "rich cells (chart) → cell grid"
            else:
                why = f"variadic overflow ({slot.name}, {len(items)} items)"
            return [CompilationPlan(
                intent=intent.name,
                layout_name="Title Only",
                title_text=str(title_text) if title_text is not None else None,
                instructions=(grid_instr,),
                fallback_used=True,
                notes=(reason, f"{why} → grid on Title Only"),
            )]

        # ----------------------------------------------------------------
        # Strategy B — chunk across pages (existing behavior).
        # Used for bullet lists and any other variadic content that
        # doesn't have a grid strategy. Only fires on classic overflow,
        # not on rich-cell or explicit-grid signals (those don't apply
        # to pure-text variadic slots).
        # ----------------------------------------------------------------
        if not is_overflow:
            return None
        plans = []
        n_chunks = (len(items) + n_slots - 1) // n_slots
        for i in range(n_chunks):
            chunk = list(items[i * n_slots : (i + 1) * n_slots])
            partial = dict(content)
            partial[slot.name] = chunk
            plan = _build_base_plan(intent, layout_shape, partial, reason, fallback_used)
            if i > 0 and plan.title_text:
                plan = CompilationPlan(
                    intent=plan.intent,
                    layout_name=plan.layout_name,
                    title_text=f"{plan.title_text} (continued)",
                    instructions=plan.instructions,
                    fallback_used=plan.fallback_used,
                    notes=plan.notes + (f"split {i+1}/{n_chunks} for variadic overflow ({slot.name})",),
                )
            else:
                plan = CompilationPlan(
                    intent=plan.intent,
                    layout_name=plan.layout_name,
                    title_text=plan.title_text,
                    instructions=plan.instructions,
                    fallback_used=plan.fallback_used,
                    notes=plan.notes + (f"split {i+1}/{n_chunks} for variadic overflow ({slot.name})",),
                )
            plans.append(plan)
        return plans
    return None


def _split_for_density_overflow(
    plan: CompilationPlan,
    layout_shape: LayoutShape,
) -> List[CompilationPlan]:
    """Inspect each instruction's density on its target slot. If any
    splittable instruction overflows (suggestion='split'), chunk that
    instruction's payload and emit one plan per chunk. Other instructions
    are repeated unchanged on every page.

    Only one instruction is split per call — the worst-overflowing one.
    Multi-instruction overflow within a single slide is rare; if it
    happens, the worst gets split and the others are tightened by virtue
    of the smaller payload.
    """
    candidates = []
    for instr in plan.instructions:
        kind = STRATEGY_TO_KIND.get(instr.strategy)
        if kind not in {"bullets", "table"}:
            continue
        ordinal = _resolve_target(instr.target, layout_shape)
        if ordinal is None or ordinal == -1 or ordinal >= len(layout_shape.content_slots):
            continue
        slot_geom = layout_shape.content_slots[ordinal].geometry
        fit = evaluate_density(kind, slot_geom, instr.payload)
        if fit.suggestion == "split":
            candidates.append((instr, fit, kind))
        elif fit.suggestion == "tighten":
            logger.info("Slide is tight (density=%.2f, %s) but within limits; not splitting.",
                        fit.score, fit.bottleneck)

    if not candidates:
        return [plan]

    candidates.sort(key=lambda c: c[1].score, reverse=True)
    target_instr, fit, kind = candidates[0]

    # Tables: greedily fill each slide with the MAXIMUM rows that fit (capacity),
    # then continue the remainder on the next slide — rather than distributing
    # rows evenly across a guessed number of pages. Cut exactly where the table
    # would overflow and duplicate the header on the continuation slide.
    if kind == "table" and fit.bottleneck == "rows" and fit.capacity > 0 \
            and isinstance(target_instr.payload, dict):
        rows = target_instr.payload.get("rows", [])
        cap = fit.capacity
        chunks = [{**target_instr.payload, "rows": rows[i:i + cap]}
                  for i in range(0, len(rows), cap)]
    else:
        n_chunks = max(2, _m.ceil(fit.score / DENSITY_TARGET_PER_PAGE))
        chunks = split_payload(kind, target_instr.payload, n_chunks)
    if len(chunks) <= 1:
        # Couldn't actually split; render with overflow.
        logger.warning("Density overflow detected but payload was not splittable; rendering as-is.")
        return [plan]

    plans = []
    for i, chunk in enumerate(chunks):
        new_instructions = tuple(
            PlaceholderInstruction(instr.target, instr.strategy, chunk)
            if instr is target_instr
            else instr
            for instr in plan.instructions
        )
        title = plan.title_text
        if title and i > 0:
            title = f"{title} (continued)"
        plans.append(CompilationPlan(
            intent=plan.intent,
            layout_name=plan.layout_name,
            title_text=title,
            instructions=new_instructions,
            fallback_used=plan.fallback_used,
            notes=plan.notes + (
                f"split {i+1}/{len(chunks)} for {kind} overflow "
                f"(score={fit.score:.2f}, bottleneck={fit.bottleneck})",
            ),
        ))
    return plans


def compile_slide(slide_spec: Dict[str, Any], layouts: Dict[str, LayoutShape], audience: str = "default") -> List[CompilationPlan]:
    """Translate one ``{intent, content}`` into one or more CompilationPlans.

    Phase 9: a single semantic spec may compile to multiple plans when the
    content overflows the chosen layout. Two split paths exist:

      1. *Variadic overflow* — variadic slot's list (e.g. dashboard.metrics)
         is longer than the layout's slot count. Detected before plan
         construction; chunks the list across pages.
      2. *Density overflow* — bullet count or table size exceeds what fits
         in the slot's geometry. Detected after the base plan is built;
         splits the offending instruction's payload across pages.

    Phase 13: optional ``conclusion`` and ``footnote`` slide-level adornments
    propagate to every plan in the returned list — when a slide splits, all
    splits carry the same bottom adornments.
    """
    intent_name = slide_spec.get("intent")
    if intent_name == "auto":                       # router always in the loop
        slide_spec = route_auto(slide_spec, audience)
        intent_name = slide_spec.get("intent")
    if intent_name not in INTENTS:
        raise ValueError(f"Unknown intent: {intent_name!r}")
    intent = INTENTS[intent_name]
    content = slide_spec.get("content", {})
    # Conclusion/footnote are slide-level adornments, but authors naturally nest
    # them inside `content`. Accept both: hoist a content-level value when no
    # slide-level one is given, and remove it from `content` so it doesn't trip
    # the "unknown content keys" warning (silently dropping a conclusion or a
    # source line is a bad failure — these are exactly what must not be lost).
    if isinstance(content, dict):
        for _k in ("conclusion", "footnote"):
            if _k in content and not slide_spec.get(_k):
                slide_spec[_k] = content[_k]   # copy up; leave in content too
    conclusion = slide_spec.get("conclusion") or None
    footnote   = slide_spec.get("footnote")   or None

    # Cover layout depends on whether a cover image is present:
    #   image  -> "Title Slide with Picture"  (default preferred order)
    #   none   -> plain "Title Slide" (drop the image slot so no empty Picture
    #             placeholder is left on the cover).
    if intent_name == "introduce_topic" and not content.get("image"):
        from dataclasses import replace as _replace
        intent = _replace(
            intent,
            slots=tuple(s for s in intent.slots if s.name != "image"),
            preferred_layouts=("Title Slide", "Title Slide with Picture"),
        )

    # Columns: 2-4 use the template's native column layouts (placeholders, so
    # master styling/geometry come for free); 5+ has no native layout, so fall
    # back to the manual canvas renderer.
    if intent_name == "show_columns" and isinstance(content.get("columns"), list) \
            and len(content["columns"]) > 4:
        from dataclasses import replace as _replace
        intent = _replace(
            intent,
            slots=(IntentSlot("title", strategy="title", target="title"),
                   IntentSlot("columns", strategy="columns", target="canvas")),
            preferred_layouts=("Title Only",), fallback_layouts=("Title Only",))

    # Content-aware layout selection for variadic intents (column/KPI count).
    _vslot = next((s for s in intent.slots if s.target == "*"), None)
    _item_count = None
    if _vslot is not None and isinstance(content.get(_vslot.name), (list, tuple)):
        _item_count = len(content[_vslot.name])

    # All-or-nothing decorative icons: only show item icons if EVERY column/KPI
    # has one that resolves; otherwise drop them all (no lopsided mix).
    if intent_name == "show_columns":
        _enforce_all_or_nothing_icons(content.get("columns"))
    elif intent_name == "dashboard":
        _enforce_all_or_nothing_icons(content.get("metrics"))

    layout, reason = select_layout(intent, layouts, _item_count)
    layout_shape = layouts[layout.name]
    fallback_used = (
        layout.name in intent.fallback_layouts
        and layout.name not in intent.preferred_layouts
    )

    # Path 1: variadic overflow (resolved before plan emission).
    variadic_plans = _split_for_variadic_overflow(
        intent, layout_shape, content, reason, fallback_used,
    )
    if variadic_plans is not None:
        plans = variadic_plans
    else:
        # Path 2: build base plan, then split for density overflow if needed.
        base_plan = _build_base_plan(intent, layout_shape, content, reason, fallback_used)
        plans = _split_for_density_overflow(base_plan, layout_shape)

    # Attach bottom adornments to every plan (replace existing dataclass since
    # CompilationPlan is frozen).
    if conclusion or footnote:
        from dataclasses import replace
        plans = [replace(p, conclusion=conclusion, footnote=footnote) for p in plans]
    return plans


# ===========================================================================
# Strategies — the leaf rendering primitives the compiler dispatches into.
# Each strategy takes (runtime, slide, placeholder, payload). They are
# intentionally small and composable. Adding a new render kind = add an entry.
# ===========================================================================

Strategy = Callable[[TemplateRuntime, Any, Any, Any], None]


def _strategy_subtitle(rt: TemplateRuntime, slide, placeholder, payload):
    rt.fill_plain_text(placeholder, str(payload))


def _strategy_plain_text(rt: TemplateRuntime, slide, placeholder, payload):
    rt.fill_plain_text(placeholder, str(payload))


def _strategy_big_text(rt: TemplateRuntime, slide, placeholder, payload):
    """Section title text. The Section layout's master sets the size and
    bold for the primary BODY placeholder; we just drop in the text and
    let the master cascade. The previous implementation forced bold=True
    here, which would have violated the 'primary font never bold' rule
    if the master used a major-font paragraph level."""
    rt.fill_plain_text(placeholder, str(payload))


def _strategy_huge_text(rt: TemplateRuntime, slide, placeholder, payload):
    """Section number — let the layout style it (the OW master ramps this to ~200pt)."""
    rt.fill_plain_text(placeholder, str(payload))


def _strategy_bullets(rt: TemplateRuntime, slide, placeholder, payload):
    """Body content: a heading + bullets, OR editorial prose paragraphs, OR a
    plain string. Prose is used when the payload carries a ``paragraphs`` list
    (or ``format: "prose"``, in which case ``bullets`` are treated as
    paragraphs) — it renders as report-style prose rather than a bulleted list."""
    if isinstance(payload, str):
        rt.fill_plain_text(placeholder, payload)
        return
    if isinstance(payload, list):                    # bare list -> bullets
        rt.fill_bullets(placeholder, "", payload)
        return
    heading = payload.get("heading", "")
    prose = payload.get("paragraphs")
    if prose is None and payload.get("format") == "prose":
        prose = payload.get("bullets", [])
    if prose:
        rt.fill_prose(placeholder, heading, prose)
        return
    bullets = payload.get("bullets", [])
    rt.fill_bullets(placeholder, heading, bullets)
    _draw_item_icon(rt, slide, placeholder, payload)


def _strategy_kpi_card(rt: TemplateRuntime, slide, placeholder, payload):
    rt.fill_kpi(
        placeholder,
        payload.get("label", ""),
        payload.get("value", ""),
        payload.get("delta", ""),
    )
    _draw_item_icon(rt, slide, placeholder, payload)


def _strategy_insight_card(rt: TemplateRuntime, slide, placeholder, payload):
    rt.fill_insight(
        placeholder,
        payload.get("title", "Insight"),
        payload.get("text", ""),
    )


def _normalize_series(chart_spec: Dict[str, Any]) -> List[Tuple[str, Any]]:
    out = []
    for s in chart_spec.get("series", []):
        if "points" in s:
            out.append((s.get("name", ""), [(float(p[0]), float(p[1])) for p in s["points"]]))
        else:
            out.append((s.get("name", ""), list(s.get("values", []))))
    return out


def _strategy_chart_native(rt: TemplateRuntime, slide, placeholder, payload):
    rt.insert_chart(
        slide, placeholder,
        payload.get("categories", []),
        _normalize_series(payload),
        chart_type=payload.get("type", "column_clustered"),
    )


def _strategy_chart_floating(rt: TemplateRuntime, slide, placeholder, payload):
    rt.insert_chart_floating(
        slide,
        payload.get("categories", []),
        _normalize_series(payload),
        chart_type=payload.get("type", "column_clustered"),
    )


def _strategy_table_native(rt: TemplateRuntime, slide, placeholder, payload):
    rt.insert_table(
        slide, placeholder,
        payload.get("headers", []),
        payload.get("rows", []),
    )


# ----- Engine-composed and image strategies (Phase 10) -----

# Lazy import — only loaded if a strategy that needs PIL is invoked, so
# environments without PIL still work for non-image decks.
_PIL_IMPORTED = False
def _ensure_pil():
    global _PIL_IMPORTED, _Image, _ImageDraw, _ImageFont
    if _PIL_IMPORTED:
        return
    from PIL import Image as _Image_mod, ImageDraw as _ImageDraw_mod, ImageFont as _ImageFont_mod
    _Image = _Image_mod; _ImageDraw = _ImageDraw_mod; _ImageFont = _ImageFont_mod
    _PIL_IMPORTED = True


def _generate_monogram_avatar(label: str, size: int = 600,
                               bg_hex: str = "000F47",
                               fg_hex: str = "FFFFFF",
                               accent_hex: str = "82BAFF") -> str:
    """Render a circular monogram avatar (initials over a colored disk) and
    return its file path. Used as a graceful fallback for ``introduce_person``
    slides that don't supply a real photo path.

    Falls back to a solid-color square + initials if PIL has no usable font.
    """
    _ensure_pil()
    import hashlib, os, tempfile

    parts = [p for p in label.strip().split() if p]
    initials = "".join(p[0].upper() for p in parts[:2]) or "?"

    # Cache per (label, size, palette) so re-runs reuse the same file.
    key = hashlib.sha1(f"{label}|{size}|{bg_hex}|{fg_hex}|{accent_hex}".encode()).hexdigest()[:12]
    out_path = os.path.join(tempfile.gettempdir(), f"avatar_{key}.png")
    if os.path.exists(out_path):
        return out_path

    bg  = tuple(int(bg_hex[i:i+2],  16) for i in (0, 2, 4))
    fg  = tuple(int(fg_hex[i:i+2],  16) for i in (0, 2, 4))
    accent = tuple(int(accent_hex[i:i+2], 16) for i in (0, 2, 4))

    img  = _Image.new("RGB", (size, size), bg)
    draw = _ImageDraw.Draw(img)
    # Subtle accent ring around the disk for visual interest
    pad = size // 14
    draw.ellipse((pad, pad, size - pad, size - pad), fill=bg, outline=accent, width=size // 80)
    inner = size // 9
    draw.ellipse((inner, inner, size - inner, size - inner), fill=bg)

    font = None
    for candidate in (
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/Library/Fonts/Arial Bold.ttf",
        "/Windows/Fonts/arialbd.ttf",
    ):
        try:
            font = _ImageFont.truetype(candidate, int(size * 0.42))
            break
        except (OSError, IOError):
            continue
    if font is None:
        font = _ImageFont.load_default()

    bbox = draw.textbbox((0, 0), initials, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    tx = (size - tw) // 2 - bbox[0]
    ty = (size - th) // 2 - bbox[1]
    draw.text((tx, ty), initials, fill=fg, font=font)

    img.save(out_path, "PNG")
    return out_path


def _placeholder_geom_in(placeholder):
    """(x, y, w, h) inches for a placeholder, or None if unknown."""
    try:
        return (placeholder.left / EMU_PER_INCH, placeholder.top / EMU_PER_INCH,
                placeholder.width / EMU_PER_INCH, placeholder.height / EMU_PER_INCH)
    except Exception:
        return None


def _strategy_image(rt: TemplateRuntime, slide, placeholder, payload):
    """Place an image, always leaving something on the slide.

    Payload (dict, or a bare string treated as ``path``):
      path / source / prompt / label   — where the picture comes from
      aspect ("16:9", "1:1", ...)       — center-crop to fill the box
      crop {left,top,right,bottom}      — explicit crop fractions (overrides aspect)
      name                              — names the shape REPLACE:<name> for swapping

    ``source`` references a deck-level image (``deck_spec['images']``) so one
    photo can be reused across slides with different crops. If no usable image
    resolves, a labelled **swap-target placeholder** is drawn instead — so an
    image slot is never silently empty.
    """
    import images as _img
    reg = getattr(rt, "_image_registry", None)
    if reg is None:
        reg = _img.ImageRegistry()
        rt._image_registry = reg
    allow_gen = bool(getattr(rt, "_allow_image_gen", False))

    if isinstance(payload, str):
        payload = {"path": payload}
    if not isinstance(payload, dict):
        return
    name = payload.get("name") or payload.get("source") or "image"
    aspect = payload.get("aspect")
    crop = payload.get("crop")

    # Resolve to a registry source id.
    source_id = payload.get("source")
    if source_id:
        if reg.get(source_id) is None:                     # not declared deck-level
            reg.register(source_id, payload.get("path"), payload.get("prompt"))
    else:
        source_id = f"_inline_{name}"
        if reg.get(source_id) is None:
            path = payload.get("path")
            if not path and not payload.get("prompt") and payload.get("label"):
                path = _generate_monogram_avatar(str(payload["label"]))  # person fallback
            reg.register(source_id, path, payload.get("prompt"))

    # Clean path: fill the native PICTURE placeholder (auto crop-to-fit) when we
    # have a real asset; otherwise drop a swap-target placeholder in its place.
    if placeholder is not None:
        geom = _placeholder_geom_in(placeholder)
        ready = reg.ensure_ready(source_id, (geom[2] / geom[3]) if geom else None,
                                 allow_generation=allow_gen)
        if ready:
            try:
                pic = rt.insert_picture(placeholder, ready)
                pic.name = f"IMAGE:{name}"
            except Exception:
                pic = None
            if pic is not None:
                try:
                    pic._element.nvPicPr.cNvPr.set(
                        "descr", f"Swappable image '{name}'. Right-click > Change "
                                 f"Picture to replace; crop/position preserved.")
                except Exception:
                    pass
            return
        if geom:                                            # no asset -> placeholder
            x, y, w, h = geom
            _img.place_image(slide, reg, source_id, x, y, w, h,
                             aspect=aspect or (f"{w/h:.2f}"), crop=crop, name=name)
            return
    # Canvas target (no placeholder): full registry path (real or placeholder).
    area = rt.content_area()
    _img.place_image(slide, reg, source_id, area.x, area.y, area.w, area.h,
                     aspect=aspect, crop=crop, name=name, allow_generation=allow_gen)


def _enforce_all_or_nothing_icons(items) -> None:
    """Decorative item icons (``show_columns`` columns, ``dashboard`` KPIs) are
    all-or-nothing per slide: keep them only if EVERY item carries an icon that
    *resolves* to a real glyph. If any item is missing one or names an icon that
    doesn't resolve, strip ``icon`` from every item so the slide never shows a
    lopsided mix of some-with-icon and some-without. Mutates the item dicts in
    place (the render paths read ``icon`` straight off them)."""
    if not isinstance(items, list) or not items:
        return
    try:
        import icons as _ic
    except Exception:
        return

    def _resolves(it) -> bool:
        if not isinstance(it, dict):
            return False
        name = it.get("icon")
        if not name:
            return False
        try:
            return bool(_ic.resolve(str(name)))
        except Exception:
            return False

    if not all(_resolves(it) for it in items):
        for it in items:
            if isinstance(it, dict):
                it.pop("icon", None)


def _draw_item_icon(rt: TemplateRuntime, slide, placeholder, payload):
    """If an item dict carries an ``icon``, draw it at the top of the
    placeholder and push the text down so the two never overlap. No-op for
    non-dict payloads or items without an icon."""
    if not isinstance(payload, dict) or not payload.get("icon"):
        return
    try:
        import icons as _ic
    except Exception:
        return
    geom = _placeholder_geom_in(placeholder)
    if not geom:
        return
    x, y, w, h = geom
    side = min(0.5, w * 0.5, h * 0.4)                 # 0.5" default (OW spec)
    color = str(payload.get("icon_color", "000F47")).lstrip("#")
    ic_x = x + 0.2; ic_y = y + 0.2                    # 0.2" inset from top & left (OW spec)
    try:
        _ic.draw_icon(slide, str(payload["icon"]), ic_x, ic_y, side, side, color_hex=color)
    except Exception as exc:
        logger.warning("icon %r skipped: %s", payload.get("icon"), exc)
        return
    try:                                             # make room for it
        from pptx.util import Inches
        placeholder.text_frame.margin_top = Inches(0.2 + side + 0.2)   # text top 0.9" (OW spec)
        placeholder.text_frame.margin_left = Inches(0.2)
    except Exception:
        pass


def _strategy_process_chevrons(rt: TemplateRuntime, slide, placeholder, payload):
    """Render a horizontal chevron flow on the slide canvas (no placeholder).

    Passes the *full* fallback content area to the runtime; the runtime
    decides whether to allocate a bullet area below the shapes (when any
    step provides bullets) or to use the full height for the chevron band.
    """
    if isinstance(payload, list):
        steps = payload
    elif isinstance(payload, dict):
        steps = payload.get("steps", [])
    else:
        return
    if not steps:
        return
    area = rt.content_area()
    has_bullets = any(
        isinstance(s, dict) and (s.get("heading") or s.get("bullets")) for s in steps
    )
    if has_bullets:
        # Use the full content area; runtime carves out the chevron band at top.
        rt.draw_process_chevrons(
            slide, area.x, area.y, area.w, area.h, steps,
        )
    else:
        # No bullets — the chevron shape is a hard 1.0" (Tokens.CHEVRON_H); pass
        # the full content area and let the runtime centre the 1" shape in it.
        rt.draw_process_chevrons(
            slide, area.x, area.y, area.w, area.h, steps,
        )


def _strategy_kpi_card_grid(rt: TemplateRuntime, slide, placeholder, payload):
    """Draw a grid of cards directly on the slide canvas (no placeholder).

    Used by the dashboard intent for two cases:

      1. The metrics list is longer than any placeholder layout can hold
         (variadic overflow → all items on one Title Only slide).
      2. At least one item carries rich content (``chart``) that doesn't
         fit a single-text-frame placeholder.
      3. The author asked for an explicit grid shape via ``content.grid``.

    Payload accepts two shapes:
      - ``list`` of item dicts — runtime auto-picks the grid
      - ``dict`` with ``items`` (required) and optional ``grid``
        (``"2x2"`` / ``(rows, cols)``) — runtime honors the explicit grid

    Each item dict supports ``label`` / ``value`` / ``delta`` (KPI fields)
    and / or ``chart`` (a chart spec). The runtime decides per-cell
    whether to render a KPI card, a composition cell (text + chart), or
    a chart-only cell."""
    if isinstance(payload, list):
        items, grid = payload, None
    elif isinstance(payload, dict):
        items = payload.get("items", [])
        grid  = payload.get("grid")
    else:
        return
    if not items:
        return
    rt.draw_card_grid(slide, items, grid=grid)


def _strategy_org_chart(rt: TemplateRuntime, slide, placeholder, payload):
    """Render a 2-level org chart on the slide canvas (no placeholder)."""
    if not isinstance(payload, dict):
        return
    root    = payload.get("root", {})
    reports = payload.get("reports", [])
    if not root:
        return
    area = rt.content_area()
    rt.draw_org_chart(
        slide, area.x, area.y, area.w, area.h, root, reports,
    )


def _strategy_pyramid(rt: TemplateRuntime, slide, placeholder, payload):
    """Render a basic pyramid on the slide canvas. Payload is the levels list
    (top to bottom) — either a bare list or a dict containing ``levels``."""
    if isinstance(payload, list):
        levels = payload
    elif isinstance(payload, dict):
        levels = payload.get("levels", [])
    else:
        return
    if not levels:
        return
    area = rt.content_area()
    # Pyramid looks better with a small top/bottom margin and centered horizontally.
    margin_x = area.w * 0.08
    margin_y = 0.2
    rt.draw_pyramid(
        slide,
        area.x + margin_x, area.y + margin_y,
        area.w - 2 * margin_x, area.h - 2 * margin_y,
        levels,
    )


def _strategy_cycle(rt: TemplateRuntime, slide, placeholder, payload):
    """Render a basic cycle on the slide canvas. Payload is the items list
    (in cycle order) — either a bare list or a dict containing ``items``."""
    if isinstance(payload, list):
        items = payload
    elif isinstance(payload, dict):
        items = payload.get("items", [])
    else:
        return
    if not items:
        return
    area = rt.content_area()
    # Square the canvas so the cycle is round, not elliptical.
    side = min(area.w, area.h)
    cx_off = (area.w - side) / 2
    cy_off = (area.h - side) / 2
    rt.draw_cycle(
        slide,
        area.x + cx_off, area.y + cy_off,
        side, side, items,
    )


def _strategy_matrix(rt: TemplateRuntime, slide, placeholder, payload):
    """Render an N×M matrix on the canvas. Payload is the matrix spec dict
    (rows/cols/x_axis/y_axis/items) or a bare items list."""
    if not isinstance(payload, (dict, list)):
        return
    if isinstance(payload, dict) and not payload.get("items") and "items" not in payload:
        # allow {rows,cols,...} with no items, but bail on truly empty
        if not any(k in payload for k in ("rows", "cols", "x_axis", "y_axis")):
            return
    area = rt.content_area()
    mx = area.w * 0.06
    my = 0.15
    rt.draw_matrix(slide, area.x + mx, area.y + my,
                   area.w - 2 * mx, area.h - my - 0.15, payload)


def _strategy_timeline(rt: TemplateRuntime, slide, placeholder, payload):
    """Render a timeline on the canvas. Payload is the milestones list or a
    dict containing ``milestones``."""
    if isinstance(payload, list):
        milestones = payload
    elif isinstance(payload, dict):
        milestones = payload.get("milestones", [])
    else:
        return
    if not milestones:
        return
    area = rt.content_area()
    rt.draw_timeline(slide, area.x, area.y + 0.1, area.w, area.h - 0.2, milestones)


def _strategy_growing_steps(rt: TemplateRuntime, slide, placeholder, payload):
    """Render ascending growing steps on the canvas. Payload is the steps list
    or a dict containing ``steps``."""
    if isinstance(payload, list):
        steps = payload
    elif isinstance(payload, dict):
        steps = payload.get("steps", [])
    else:
        return
    if not steps:
        return
    area = rt.content_area()
    mx = area.w * 0.04
    rt.draw_growing_steps(slide, area.x + mx, area.y + 0.3,
                          area.w - 2 * mx, area.h - 0.5, steps)


def _strategy_contents(rt: TemplateRuntime, slide, placeholder, payload):
    """Render a table of contents on the right of the Contents layout. Payload
    is the sections list (or a dict containing ``sections``)."""
    if isinstance(payload, list):
        sections = payload
    elif isinstance(payload, dict):
        sections = payload.get("sections", [])
    else:
        return
    if not sections:
        return
    # Right-hand column, matching the OW Contents reference (title sits left).
    rt.draw_contents(slide, sections, x=6.67, y=0.5, w=6.17, h=6.4)


DIAGRAM_KINDS = frozenset({"growing_steps", "pyramid", "cycle", "process",
                           "chevrons", "org_chart", "matrix", "timeline"})

# Each preset's region signature (a multiset of block kinds) — used by the fit
# router to decide whether a requested composition maps cleanly onto a preset.
PRESET_SIGNATURES = {
    "explain": ("text",),
    "summary": ("text",),
    "compare_two_options": ("text", "text"),
    "show_columns": ("text", "text", "text"),
    "show_trend_with_key_message": ("chart", "text"),
    "show_graphic_with_text": ("diagram", "text"),
    "dashboard": ("kpi", "kpi", "kpi", "kpi"),
    "show_data": ("table",),
    "show_waterfall": ("waterfall",),
    "show_treemap": ("treemap",),
}

# Snap threshold (tau) by audience: stricter (more preset-bound) for external /
# executive decks, looser (more composer freedom) for internal working sessions.
AUDIENCE_TAU = {"board": 0.90, "executive": 0.85, "client": 0.85,
                "internal": 0.60, "project_team": 0.55, "default": 0.75}


def _composition_signature(regions):
    return ["diagram" if r.get("block", "") in DIAGRAM_KINDS else r.get("block", "")
            for r in regions]


def preset_fit(regions):
    """Return (fit in [0,1], best_preset_name): how cleanly the requested region
    set maps onto the closest preset's signature. 1.0 = exact match."""
    from collections import Counter
    want = Counter(_composition_signature(regions))
    best, best_name = 0.0, None
    for name, psig in PRESET_SIGNATURES.items():
        have = Counter(psig)
        inter = sum((want & have).values())
        score = inter / max(sum(want.values()), sum(have.values()), 1)
        if score > best:
            best, best_name = score, name
    return best, best_name


def route_composition(regions, audience="default"):
    """Decide between snapping to a preset and using the open composer. Returns
    (mode, preset_name_or_None, fit) where mode is 'preset' or 'compose'. tau is
    conditioned on the audience, so the same fit routes differently for a board
    deck vs an internal session."""
    fit, name = preset_fit(regions)
    tau = AUDIENCE_TAU.get(audience, AUDIENCE_TAU["default"])
    return ("preset", name, fit) if fit >= tau else ("compose", None, fit)


# Canonical region proportions per composition-expressible preset. When the
# router snaps a composition to one of these, the composer uses these widths
# (the curated, balanced layout) instead of the author's free widths.
PRESET_LAYOUTS = {
    "show_trend_with_key_message": [2, 1],   # chart wide, commentary narrow
    "compare_two_options": [1, 1],
    "show_graphic_with_text": [1.4, 1],
    "explain": [1], "summary": [1],
    "show_waterfall": [1], "show_treemap": [1], "show_data": [1],
    # show_columns -> equal widths (left unset so the composer splits evenly)
}


def route_auto(slide_spec, audience="default"):
    """Resolve an ``auto`` slide. The router is consulted for the slide's region
    set; on a preset match the composer uses that preset's canonical proportions,
    otherwise the author's free widths stand. Returns a rewritten ``compose`` spec
    so the router is, by construction, in the loop for every auto slide."""
    content = slide_spec.get("content", {}) or {}
    regions = (content.get("regions")
               or (content.get("compose", {}) or {}).get("regions") or [])
    mode, preset, fit = route_composition(regions, audience)
    tau = AUDIENCE_TAU.get(audience, AUDIENCE_TAU["default"])
    widths = PRESET_LAYOUTS.get(preset) if mode == "preset" else None
    out = []
    for i, reg in enumerate(regions):
        r = dict(reg)
        if widths and i < len(widths):
            r["width"] = widths[i]
        out.append(r)
    logger.info("auto route -> %s%s (fit=%.2f, tau=%.2f, audience=%s)",
                mode, f" [{preset}]" if preset else "", fit, tau, audience)
    new_spec = dict(slide_spec)
    new_spec["intent"] = "compose"
    new_spec["content"] = {"title": content.get("title", ""),
                           "compose": {"regions": out}}
    return new_spec


def _strategy_compose(rt: TemplateRuntime, slide, placeholder, payload):
    """Open composition on the canvas. Payload:
    {grid:{cols?, rows?, cells:[{block, data, col?, row?, colspan?, rowspan?}]}}
    for a flexible spanning grid, or
    {rows:[{height?, regions:[{block, data, width?}]}]} / {regions:[...]}."""
    if not isinstance(payload, dict):
        return
    grid = payload.get("grid")
    if isinstance(grid, dict) and grid.get("cells"):
        # Content-measured autolayout: opt in per grid with "autolayout": true,
        # or globally via runtime flag AUTOLAYOUT_DEFAULT. Falls back to the
        # equal-track draw_grid otherwise. Either way the cells render through
        # rt.render_block, so nothing downstream changes.
        use_al = grid.get("autolayout", getattr(rt, "autolayout_default", False))
        if use_al:
            try:
                import autolayout
                res = autolayout.draw_grid_autolayout(rt, slide, grid)
                rt._last_layout = getattr(res, "diagnostics", None)
                return
            except Exception as exc:                         # never block a build
                logger.warning("autolayout failed (%s); equal-track fallback", exc)
        rt.draw_grid(slide, grid)
        return
    rows = payload.get("rows")
    if rows is None:
        regions = payload.get("regions") or []
        if not regions:
            return
        rows = [{"regions": regions}]
    rt.draw_composition(slide, rows)


def _canvas_chart_top(rt, slide, payload):
    """Draw the optional heading/subheading block for a canvas chart and return
    (content_area, chart_top_y). Shared by the waterfall and treemap strategies."""
    area = rt.content_area()
    top = area.y
    if payload.get("heading") or payload.get("subheading"):
        rt.draw_text_block(slide, area.x, area.y, area.w, 0.5,
                           heading=payload.get("heading", "") or "",
                           subheading=payload.get("subheading", "") or "")
        top = area.y + 0.55
    return area, top


def _strategy_treemap(rt: TemplateRuntime, slide, placeholder, payload):
    """Treemap drawn on the canvas (no placeholder). Payload is a dict
    {items:[{label, value}], heading?, subheading?}."""
    if not isinstance(payload, dict):
        return
    items = payload.get("items") or []
    if not items:
        return
    area, top = _canvas_chart_top(rt, slide, payload)
    rt.treemap(slide, items, x=area.x, y=top, w=area.w,
                    h=max(1.0, area.y + area.h - top))


def _strategy_waterfall(rt: TemplateRuntime, slide, placeholder, payload):
    """Waterfall / bridge chart drawn on the canvas (no placeholder). Payload is
    a dict {orientation, categories, values, totals?, heading?, subheading?}."""
    if not isinstance(payload, dict):
        return
    if not payload.get("categories") or not payload.get("values"):
        return
    area, top = _canvas_chart_top(rt, slide, payload)
    rt.add_waterfall(slide, payload, cx=area.x, cy=top, cw=area.w,
                     ch=max(1.0, area.y + area.h - top))


def _strategy_quote(rt: TemplateRuntime, slide, placeholder, payload):
    """Big pull-quote. Payload is a string, or a dict with text/quote +
    attribution/source/by."""
    if isinstance(payload, str):
        text, attribution = payload, ""
    elif isinstance(payload, dict):
        text = payload.get("text") or payload.get("quote") or ""
        attribution = (payload.get("attribution") or payload.get("source")
                       or payload.get("by"))
        if not attribution and any(payload.get(k) for k in ("name", "title", "role", "location")):
            attribution = {"name": payload.get("name"),
                           "title": payload.get("title") or payload.get("role"),
                           "location": payload.get("location")}
        attribution = attribution or ""
    else:
        return
    if not text:
        return
    area = rt.content_area()
    rt.draw_quote(slide, area.x, area.y, area.w, area.h, text, attribution)


def _strategy_graphic_with_text(rt: TemplateRuntime, slide, placeholder, payload):
    """Composite slide: a graphic in one pane, supporting text in the other.
    Payload: {graphic:{kind, ...}, text:{heading, subheading, bullets|paragraphs},
    side:'graphic_left'|'graphic_right'}. The graphic supports the text (text left,
    graphic right) or is supported by the text (graphic left, text right)."""
    if not isinstance(payload, dict):
        return
    graphic = payload.get("graphic") or {}
    text = payload.get("text") or {}
    side = payload.get("side", "graphic_left")
    split = str(payload.get("split", "equal")).lower()
    area = rt.content_area()
    gap = GRID_GAP_LG                                  # 0.5" gutter (grid)
    # Two-pane splits ARE grid layouts: equal = a 2-column grid; 2:1 / 1:2 = a
    # 3-column grid where one pane spans two columns. Computing them from the
    # grid's column width (not an independent 2/3 fraction) keeps a single
    # definition of every split across the whole system.
    col3 = (area.w - 2 * gap) / 3.0                   # one column in a 3-track grid
    if split in ("graphic_2_3", "graphic_major", "2_3"):
        graphic_w = 2 * col3 + gap                    # graphic spans 2 of 3 columns
    elif split in ("text_2_3", "text_major", "1_3"):
        graphic_w = col3                              # graphic is 1 of 3 columns
    else:                                             # equal (default) = 2 columns
        graphic_w = (area.w - gap) / 2.0
    text_w = area.w - graphic_w - gap                 # the other pane fills the rest
    if side == "graphic_right":
        text_x, graphic_x = area.x, area.x + text_w + gap
    else:  # graphic_left (default)
        graphic_x, text_x = area.x, area.x + graphic_w + gap
    rt.draw_visual(slide, graphic.get("kind", ""), graphic,
                   graphic_x, area.y, graphic_w, area.h)
    rt.draw_text_block(slide, text_x, area.y, text_w, area.h,
                       heading=text.get("heading", ""),
                       subheading=text.get("subheading", ""),
                       bullets=text.get("bullets"),
                       paragraphs=text.get("paragraphs"))


def _strategy_columns(rt: TemplateRuntime, slide, placeholder, payload):
    """Render parallel headed text columns on the canvas. Payload is the columns
    list (or a dict containing ``columns``)."""
    if isinstance(payload, list):
        cols = payload
    elif isinstance(payload, dict):
        cols = payload.get("columns", [])
    else:
        return
    if not cols:
        return
    area = rt.content_area()
    rt.draw_columns(slide, area.x, area.y, area.w, area.h, cols)


def _strategy_stat_callout(rt, slide, placeholder, payload):
    stats = payload if isinstance(payload, list) else (payload or {}).get("stats", [])
    import components as _c
    _c.draw_stat_callouts(slide, rt.content_area(), stats)


def _strategy_icon_rows(rt, slide, placeholder, payload):
    rows = payload if isinstance(payload, list) else (payload or {}).get("rows", [])
    import components as _c
    _c.draw_icon_rows(slide, rt.content_area(), rows)


def _strategy_compare(rt, slide, placeholder, payload):
    p = payload if isinstance(payload, dict) else {}
    import components as _c
    _c.draw_compare(slide, rt.content_area(), p.get("pros", []), p.get("cons", []),
                    pros_heading=p.get("pros_heading", ""),
                    cons_heading=p.get("cons_heading", ""))


def _strategy_infographic(rt, slide, placeholder, payload):
    spec = payload if isinstance(payload, dict) else {}
    import infographics as _ig
    _ig.draw_infographic(slide, rt.content_area(), spec)


def _strategy_sticker(rt, slide, placeholder, payload):
    p = payload if isinstance(payload, dict) else {"text": payload}
    import components as _c
    _c.draw_sticker(slide, rt.content_area(), p.get("text", ""), note=p.get("note", ""))


STRATEGIES: Dict[str, Strategy] = {
    "title":             lambda rt, s, p, _: None,  # title is handled by the executor, not a placeholder strategy
    "subtitle":          _strategy_subtitle,
    "plain_text":        _strategy_plain_text,
    "big_text":          _strategy_big_text,
    "huge_text":         _strategy_huge_text,
    "bullets":           _strategy_bullets,
    "kpi_card":          _strategy_kpi_card,
    "kpi_card_grid":     _strategy_kpi_card_grid,
    "insight_card":      _strategy_insight_card,
    "chart_native":      _strategy_chart_native,
    "chart_floating":    _strategy_chart_floating,
    "table_native":      _strategy_table_native,
    "image":             _strategy_image,
    "process_chevrons":  _strategy_process_chevrons,
    "org_chart":         _strategy_org_chart,
    "pyramid":           _strategy_pyramid,
    "cycle":             _strategy_cycle,
    "matrix":            _strategy_matrix,
    "timeline":          _strategy_timeline,
    "growing_steps":     _strategy_growing_steps,
    "contents":          _strategy_contents,
    "columns":           _strategy_columns,
    "quote":             _strategy_quote,
    "graphic_text":      _strategy_graphic_with_text,
    "waterfall":         _strategy_waterfall,
    "treemap":           _strategy_treemap,
    "compose":           _strategy_compose,
    "stat_callout":      _strategy_stat_callout,
    "icon_rows":         _strategy_icon_rows,
    "compare":           _strategy_compare,
    "sticker":           _strategy_sticker,
    "infographic":       _strategy_infographic,
}


# ===========================================================================
# Execution: turn a plan into actual slide mutations
# ===========================================================================

ROLE_TO_ORDINAL = {"primary": 0, "secondary": 1, "tertiary": 2, "quaternary": 3}
AREA_ROLES      = {"largest": 0, "second_largest": 1, "third_largest": 2, "fourth_largest": 3}


def _resolve_target(target: str, layout_shape: LayoutShape) -> Optional[int]:
    """Return the ordinal of the content placeholder this target points to,
    or ``None`` if it should map to the title placeholder instead, or
    ``-1`` if the target cannot be satisfied at all.

    ``slot_N`` and ``primary``/``secondary``/etc. resolve by *position*
    (top-to-bottom, left-to-right). ``largest``/``second_largest`` resolve
    by *area*, which is what most chart-vs-callout intents want.
    ``picture`` resolves by *placeholder type* — the first PICTURE-typed
    content slot in the layout — which lets bio intents target a photo
    placeholder without depending on its position.

    The pseudo-target ``"canvas"`` is handled separately by the executor
    (no placeholder lookup); strategies that draw engine-composed shapes
    like chevrons or org charts use it.
    """
    n = len(layout_shape.content_slots)

    if target.startswith("slot_"):
        try:
            idx = int(target.split("_", 1)[1])
        except ValueError:
            return -1
        return idx if idx < n else -1

    if target in ROLE_TO_ORDINAL:
        idx = ROLE_TO_ORDINAL[target]
        if idx < n:
            return idx
        # Graceful degrade: primary on a content-less layout → title placeholder.
        if idx == 0 and layout_shape.has_title:
            return None
        return -1

    if target in AREA_ROLES:
        rank = AREA_ROLES[target]
        if rank >= n:
            if rank == 0 and layout_shape.has_title:
                return None
            return -1
        ranked = sorted(
            range(n),
            key=lambda i: layout_shape.content_slots[i].geometry[2]
                          * layout_shape.content_slots[i].geometry[3],
            reverse=True,
        )
        return ranked[rank]

    if target in {"picture", "photo", "image"}:
        for i, slot in enumerate(layout_shape.content_slots):
            if slot.ph_type == "PICTURE":
                return i
        return -1

    # Type-based text targets — useful when a layout has heterogeneous
    # placeholder types (e.g. CV bio: BODY for role + OBJECT for content).
    if target == "body":
        for i, slot in enumerate(layout_shape.content_slots):
            if slot.ph_type == "BODY":
                return i
        return -1
    if target == "object":
        for i, slot in enumerate(layout_shape.content_slots):
            if slot.ph_type == "OBJECT":
                return i
        return -1

    return -1


def _sample_deck_text(deck_spec: Dict[str, Any], max_chars: int = 4000) -> str:
    """Concatenate slide titles, bodies, quotes, insight text — the fields
    that actually carry natural-language prose — into one string bounded at
    ``max_chars`` for language detection. Skips numeric-only fields (values,
    deltas, table cells) and non-content metadata."""
    parts: List[str] = []
    used = 0
    def push(s):
        nonlocal used
        if not s or used >= max_chars:
            return
        s = str(s)
        remaining = max_chars - used
        parts.append(s[:remaining]); used += min(len(s), remaining)
    push(deck_spec.get("title"))
    for slide in deck_spec.get("slides", []) or []:
        if used >= max_chars:
            break
        c = slide.get("content") or {}
        if not isinstance(c, dict):
            continue
        for k in ("title", "subtitle", "conclusion", "footnote"):
            push(c.get(k))
        for k in ("body", "insight", "panel", "quote"):
            v = c.get(k)
            if isinstance(v, dict):
                for sub in ("text", "heading", "subheading"):
                    push(v.get(sub))
                for lst in ("paragraphs", "bullets"):
                    for item in (v.get(lst) or []):
                        push(item if isinstance(item, str) else item.get("text") if isinstance(item, dict) else None)
            elif isinstance(v, str):
                push(v)
        # show_columns / dashboard: item headings + bullets (skip pure numbers)
        for k in ("columns", "metrics"):
            for it in (c.get(k) or []):
                if isinstance(it, dict):
                    push(it.get("heading"))
                    for b in (it.get("bullets") or []):
                        push(b if isinstance(b, str) else None)
    return "\n".join(parts)


class SemanticCompiler:
    """The orchestration layer.

    A SemanticCompiler binds a runtime + an analyzed set of layouts. It can
    compile a single slide (returning a CompilationPlan you can inspect) or
    compile-and-execute a deck.
    """

    def __init__(self, runtime: TemplateRuntime):
        self.runtime = runtime
        self.layouts = analyze_template(runtime)

    # ---- compile -------------------------------------------------------

    def compile(self, slide_spec: Dict[str, Any], audience: str = "default") -> List[CompilationPlan]:
        """Compile one semantic slide spec into one or more CompilationPlans.

        A single intent may compile to multiple plans when content overflows
        the chosen layout (e.g. 12 bullets that need to be split across two
        slides). The caller is responsible for executing every plan.
        """
        return compile_slide(slide_spec, self.layouts, audience)

    # ---- execute -------------------------------------------------------

    def _shrink_placeholders_for_adornments(self, slide, new_max_y_in: float) -> None:
        """Adjust content placeholders so their bottom edge doesn't pass below
        ``new_max_y_in`` (inches). Called before strategies run on slides that
        carry Conclusion/Footnote, so charts, tables, and KPI cards bind into
        a reduced area that doesn't overlap the bottom adornments.

        Title placeholders and bottom adornments themselves are skipped.
        Inherited geometry is materialised by reading ``shape.top`` and
        ``shape.height`` (which fall back to the layout's values) and writing
        explicit overrides on the slide.

        IMPORTANT: read all four geometry values BEFORE writing any. Setting
        any of them creates an explicit ``<a:xfrm>`` on the placeholder, and
        subsequent reads return 0 for the other axes (whatever python-pptx
        defaulted them to) instead of the layout's inherited value. Reading
        first, writing second avoids that corruption.
        """
        from pptx.util import Emu
        from pptx.enum.shapes import PP_PLACEHOLDER
        EMU = 914400
        target_bottom_emu = int(new_max_y_in * EMU)
        for ph in slide.placeholders:
            try:
                ph_type = ph.placeholder_format.type
            except Exception:
                continue
            if ph_type == PP_PLACEHOLDER.TITLE:
                continue
            # Snapshot all 4 axes BEFORE any setter call.
            top    = ph.top
            left   = ph.left
            width  = ph.width
            height = ph.height
            if any(v is None for v in (top, left, width, height)):
                continue
            current_bottom = top + height
            if current_bottom <= target_bottom_emu:
                continue
            new_height = target_bottom_emu - top
            if new_height <= 0:
                continue
            # Now write all 4 explicitly so the xfrm is fully populated.
            ph.left   = Emu(left)
            ph.top    = Emu(top)
            ph.width  = Emu(width)
            ph.height = Emu(new_height)

    def execute(self, plan: CompilationPlan) -> None:
        slide = self.runtime.add_slide(plan.layout_name)
        if plan.title_text:
            self.runtime.set_title(slide, plan.title_text)

        layout_shape = self.layouts[plan.layout_name]

        # Hard rule: cover/structural/legal layouts never carry a Conclusion or
        # Footnote, regardless of what the deck spec asked for.
        adornments_allowed = (
            plan.layout_name.strip().lower() not in NO_ADORNMENT_LAYOUTS
        )
        eff_footnote   = plan.footnote   if adornments_allowed else None
        eff_conclusion = plan.conclusion if adornments_allowed else None
        if not adornments_allowed and (plan.footnote or plan.conclusion):
            logger.info(
                "Layout %r disallows adornments; dropping footnote/conclusion.",
                plan.layout_name,
            )

        # Phase 13: if the slide carries Conclusion/Footnote adornments, the
        # canvas strategies need a reduced content area so they don't draw on
        # top of the bottom adornments. Compute the override before running
        # strategies; restore after.
        from runtime import (
            ContentArea, FALLBACK_CONTENT,
            FOOTER_BASELINE_Y, CONCLUSION_FOOTNOTE_GAP,
            CONTENT_BOTTOM_MARGIN,
        )
        footnote_height = 0.0
        if eff_footnote:
            footnote_height = self.runtime.estimate_footnote_height(eff_footnote)
        conclusion_height = 0.0
        if eff_conclusion:
            conclusion_height = self.runtime.estimate_conclusion_height(eff_conclusion)

        if eff_conclusion or eff_footnote:
            # Topmost bottom-adornment edge becomes the new content max y.
            if eff_conclusion and eff_footnote:
                top_edge = (FOOTER_BASELINE_Y - footnote_height
                            - CONCLUSION_FOOTNOTE_GAP - conclusion_height)
            elif eff_conclusion:
                top_edge = FOOTER_BASELINE_Y - conclusion_height
            else:  # footnote only
                top_edge = FOOTER_BASELINE_Y - footnote_height
            new_max_y = top_edge - CONTENT_BOTTOM_MARGIN
            new_h = max(0.5, new_max_y - FALLBACK_CONTENT.y)
            self.runtime.set_content_area_override(ContentArea(
                x=FALLBACK_CONTENT.x, y=FALLBACK_CONTENT.y,
                w=FALLBACK_CONTENT.w, h=new_h,
            ))
            # Also shrink any content placeholders whose inherited geometry
            # would extend below the new max y. Done BEFORE strategies run
            # so charts/tables/cards bind into the reduced area. Title
            # placeholders are skipped — they live above the content band.
            self._shrink_placeholders_for_adornments(slide, new_max_y)

        try:
            # Snapshot the content placeholders BEFORE any chart/table binding
            # destroys them. Each instruction may consume one.
            content_phs = list(self.runtime.content_placeholders(slide))

            for instr in plan.instructions:
                # Canvas instructions render directly on the slide with no
                # placeholder lookup. Used by engine-composed strategies
                # (chevrons, org charts) on Title Only layouts.
                if instr.target == "canvas":
                    strategy = STRATEGIES.get(instr.strategy)
                    if strategy is None:
                        logger.warning("Unknown strategy %r; skipping.", instr.strategy)
                        continue
                    try:
                        strategy(self.runtime, slide, None, instr.payload)
                    except Exception as exc:
                        logger.warning(
                            "Canvas strategy %r failed on layout %r: %s",
                            instr.strategy, plan.layout_name, exc,
                        )
                    continue

                ordinal = _resolve_target(instr.target, layout_shape)
                if ordinal == -1:
                    logger.warning(
                        "Could not resolve target %r in layout %r; skipping %s.",
                        instr.target, plan.layout_name, instr.strategy,
                    )
                    continue

                if ordinal is None:
                    # Title placeholder fallback (used when a primary-targeted slot
                    # lands on a layout that has only a title).
                    if not plan.title_text and slide.shapes.title is not None:
                        self.runtime.set_title(slide, str(instr.payload))
                    else:
                        logger.warning(
                            "Slot %r could not be placed: no content slots and title "
                            "is already set on layout %r.",
                            instr.strategy, plan.layout_name,
                        )
                    continue

                if ordinal >= len(content_phs):
                    logger.warning(
                        "Layout %r exposes %d content slots at runtime but plan asked "
                        "for ordinal %d; skipping.",
                        plan.layout_name, len(content_phs), ordinal,
                    )
                    continue

                placeholder = content_phs[ordinal]
                strategy = STRATEGIES.get(instr.strategy)
                if strategy is None:
                    logger.warning("Unknown strategy %r; skipping.", instr.strategy)
                    continue

                try:
                    strategy(self.runtime, slide, placeholder, instr.payload)
                except Exception as exc:
                    logger.warning(
                        "Strategy %r failed on layout %r ordinal %d: %s",
                        instr.strategy, plan.layout_name, ordinal, exc,
                    )
        finally:
            self.runtime.set_content_area_override(None)

        # Add bottom adornments AFTER strategies so they're drawn on top of
        # any canvas content (in practice they don't overlap because canvas
        # strategies were given a reduced content area).
        if eff_footnote:
            self.runtime.add_footnote(slide, eff_footnote)
        if eff_conclusion:
            self.runtime.add_conclusion(
                slide, eff_conclusion,
                has_footnote=bool(eff_footnote),
                footnote_height=footnote_height,
            )

    # ---- top-level convenience ----------------------------------------

    def build(self, deck_spec: Dict[str, Any]) -> List[CompilationPlan]:
        all_plans: List[CompilationPlan] = []
        # Build report (observability): surfaced by generate_deck so the author
        # and the GPT can see what happened instead of silent failures.
        self.report = {"slides": 0, "intents": {}, "splits": 0, "unknown_keys": []}
        audience = ((deck_spec.get("deck", {}) or {}).get("audience")
                    or deck_spec.get("audience") or "default")
        # Language (BCP-47) drives chart number-format locale: en-US -> 1,000,
        # de-DE -> 1.000, fr-FR -> 1 000. Explicit deck_spec["language"] always
        # wins; when absent, auto-detect from the deck's content (titles + body
        # text) via a stopword-frequency heuristic. Falls back to en-US on ties
        # or too little text (< 3 stopword hits).
        explicit_lang = ((deck_spec.get("deck", {}) or {}).get("language")
                         or deck_spec.get("language"))
        if explicit_lang:
            self.runtime.language = str(explicit_lang)
        else:
            sample = _sample_deck_text(deck_spec)
            lang, hits, runner = detect_language(sample)
            if hits >= 3 and lang in LCID_MAP:
                self.runtime.language = lang
                logger.info("Language auto-detected: %s (hits=%d, runner-up=%s)",
                            lang, hits, runner)
            else:
                self.runtime.language = "en-US"
                if hits > 0:
                    logger.info("Language auto-detect inconclusive "
                                "(hits=%d, top=%s); using en-US", hits, lang)
        for i, slide_spec in enumerate(deck_spec.get("slides", []), 1):
            intent_name = slide_spec.get("intent")
            content = slide_spec.get("content", {}) or {}
            # An image-bearing person slide should never be photo-less: fall back
            # to a monogram avatar derived from the name.
            if (intent_name == "introduce_person" and isinstance(content, dict)
                    and not content.get("photo")):
                content = {**content, "photo": {"label": content.get("name", "")}}
                slide_spec = {**slide_spec, "content": content}
            intent_def = INTENTS.get(intent_name)
            if intent_def and isinstance(content, dict):
                known = {s.name for s in intent_def.slots}
                # conclusion/footnote are universal adornments accepted at content
                # level and hoisted to slide level during compile — not "unknown".
                known |= {"conclusion", "footnote"}
                unknown = [k for k in content if k not in known]
                if unknown:
                    self.report["unknown_keys"].append((i, intent_name, unknown))
                    logger.warning("Slide %d (%s): ignored unknown content keys: %s",
                                   i, intent_name, ", ".join(unknown))
            try:
                plans = self.compile(slide_spec, audience)
            except Exception as exc:
                logger.error("Slide %d failed to compile: %s", i, exc)
                continue
            if len(plans) == 1:
                logger.info("Slide %d — %s", i, plans[0].intent)
            else:
                self.report["splits"] += len(plans) - 1
                logger.info("Slide %d — %s → %d plans (overflow split)",
                            i, plans[0].intent, len(plans))
            for plan in plans:
                for line in plan.describe().splitlines():
                    logger.info("  %s", line)
                self.execute(plan)
                all_plans.append(plan)
                self.report["slides"] += 1
                self.report["intents"][plan.intent] = (
                    self.report["intents"].get(plan.intent, 0) + 1)
        return all_plans

    def build_summary(self) -> str:
        """One-line-per-fact human summary of the last build (for stdout)."""
        from runtime import __version__ as rt_version
        r = getattr(self, "report", None)
        if not r:
            return ""
        lines = [f"runtime v{rt_version}: built {r['slides']} slide(s) "
                 f"from {len(r['intents'])} intent type(s); splits: {r['splits']}"]
        if r["unknown_keys"]:
            lines.append("ignored unknown content keys (check for typos):")
            for idx, intent, keys in r["unknown_keys"]:
                lines.append(f"  - slide {idx} ({intent}): {', '.join(keys)}")
        return "\n".join(lines)


# ===========================================================================
# Introspection helpers (useful for tests and debugging)
# ===========================================================================

def supported_intents() -> List[str]:
    return sorted(INTENTS.keys())


def supported_strategies() -> List[str]:
    return sorted(STRATEGIES.keys())


def supported_content_kinds() -> List[str]:
    """All content kinds known to the registry."""
    seen = set()
    for cap in LAYOUT_CAPABILITIES.values():
        seen |= set(cap.supports)
    seen |= set(STRATEGY_TO_KIND.values())
    return sorted(seen)


def describe_capabilities() -> str:
    """Multi-line string describing every layout's declared capabilities."""
    lines = []
    for name in sorted(LAYOUT_CAPABILITIES.keys()):
        cap = LAYOUT_CAPABILITIES[name]
        supports = ", ".join(sorted(cap.supports)) or "(none)"
        lines.append(f"  {name!r:32}  supports={{{supports}}}")
        if cap.notes:
            lines.append(f"  {' '*32}  ↳ {cap.notes}")
    return "\n".join(lines)
