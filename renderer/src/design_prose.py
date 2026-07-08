"""design_prose.py — the HUMAN-AUTHORED rationale layer of the design system.

This is the one file in the design-system pipeline that a person writes and
reviews. Everything else (token values, component slot schemas) is generated
from the engine; this supplies the *why and when* that generation can't.

Discipline (enforced by the lint rule ``prose-restates-value``):
  - Prose NEVER restates a number. Refer to a token by name (`TYPE_KPI`,
    `SPACE_GAP_LG`); do not write "26pt" or "0.5 inch". If the value changes,
    the generated table updates and this prose stays correct.
  - Each key below maps to a section or component the generator emits. Missing
    keys are fine (the section simply ships without rationale); unknown keys are
    surfaced by the generator so this file can't silently rot.

The generator (``runtime.design_system_markdown``) stitches each block in under
its matching table.
"""
from __future__ import annotations

# --- Per-section rationale (keyed to the generated section headings) --------
SECTION_PROSE = {
    "overview": (
        "This file is the contract between the brand and the engine. The tables "
        "are generated from live code, so they are always true; the prose is "
        "written by hand to say *why* a value exists and *when* to reach for it. "
        "Read the prose to make a judgement; trust the table for the number."
    ),
    "type": (
        "The scale is built around two anchors and two floors. The anchors are "
        "`TYPE_BODY` (everything reads up or down from body) and `TYPE_KPI` (the "
        "display number that has to carry a slide from across a room). The floors "
        "are non-negotiable: `TYPE_MAJOR_MIN` keeps Marsh Serif large enough to "
        "stay editorial rather than fading into body copy, and `TYPE_AUTOFIT_MIN` "
        "is the point below which auto-fit stops shrinking and instead leaves "
        "the overflow visible — a deliberate signal to cut words, never a "
        "licence to render text too small to read. Reach for `TYPE_INSIGHTS` only "
        "for the one-line 'so what'; if every line is an insight, none is."
    ),
    "space": (
        "Spacing is deliberately coarse so slides can't drift into bespoke "
        "gutters. `SPACE_GAP_LG` is the default gutter for grids and two-pane "
        "splits; `SPACE_GAP` is the fallback used *only* when the large gutter "
        "would breach a minimum column width. The content box "
        "(`SPACE_CONTENT_X/Y/W/H`) clears a two-line title on purpose — never "
        "push content above `SPACE_CONTENT_Y` to win space, because that is the "
        "line that stops titles and bodies colliding (a regression we have hit "
        "before). Card internals key off `SPACE_CARD_PAD`; keep them there so "
        "every card breathes identically."
    ),
    "para": (
        "Paragraph spacing does the work that blank lines would otherwise do. "
        "`PARA_BULLET` separates list items; `PARA_PROSE` is the looser setting "
        "for running text; `PARA_SUBHEAD` opens air above a sub-heading so it "
        "reads as a break. Never simulate spacing with empty paragraphs — it "
        "breaks auto-fit's height maths."
    ),
    "inset": (
        "Insets stay near zero on purpose. Text boxes that carry no fill use "
        "zero side margins so the text aligns to the true grid edge; a non-zero "
        "inset there would make a 'left-aligned' block look indented. `INSET_R` "
        "leaves a hair of room on the right so wrapped lines don't kiss a border."
    ),
    "stroke": (
        "There is essentially one stroke weight in this system — "
        "`STROKE_HAIRLINE` (and its aliases) at a uniform hairline — because a "
        "deck of mixed line weights reads as noise. The single sanctioned "
        "exception is `STROKE_PANEL`, the one heavier rule on the key-message "
        "panel, which earns its weight by being rare. A weight of 0 means *no* "
        "border; that is the intended way to remove a box outline, not a "
        "near-zero hairline."
    ),
    "chart": (
        "Chart chrome is quiet and the data is loud — the deliberate contrast "
        "between `CHART_LINE_WEIGHT_EMU` (the thin axis/gridline weight) and "
        "`CHART_LINE_SERIES_EMU` (the heavier data-series stroke) is the whole "
        "point: structure recedes, the line carries. Gridlines use a warm light "
        "grey (`CHART_GRIDLINE_COLOR_HEX`), never black. The doughnut hole is "
        "kept thin (`CHART_DOUGHNUT_HOLE_PCT`) so the ring reads as a figure, "
        "not a pie. Series colour comes from the OW chart palette in order; do "
        "not hand-pick chart colours."
    ),
    "bullets": (
        "The list styles come straight from the OW master, so they inherit "
        "rather than being re-styled per slide. Levels 0–3 are the body/bullet "
        "ladder (glyphs step from • to – to -); levels 6–8 are the Marsh Serif "
        "'primary' display sizes. The rule that matters: the primary font is "
        "never bold (weight comes from size and the serif itself), which is why "
        "the bold flag is suppressed on those levels even if a caller asks for it."
    ),
    "align": (
        "Only the alignments that recur across slides are managed here; "
        "everything inside a diagram stays renderer-local so one timeline tweak "
        "can't shift every table. The managed set encodes the house habits: "
        "table headers sit bottom-left (so a wrapped header grows upward off the "
        "rule), body cells top-left, and both the Conclusion and Footnote pin "
        "bottom-left so they share a baseline."
    ),
    "colors": (
        "Colour is addressed by *role*, never by hex, so the palette has one "
        "place to change. `default_fill` (warm cream) is the resting state for "
        "any decorative shape — reach for it unless a shape is genuinely being "
        "emphasised. `subtle_highlight` (light blue) is emphasis that doesn't "
        "shout; `strong_highlight` (midnight blue) is the primary attention-grab; "
        "`text_highlight` (gold) marks a word or a rule, not a whole shape. The "
        "gold is the same colour as the Conclusion's rule on purpose, so the "
        "eye links 'this is the point' across the deck. If a source deck arrives "
        "in an off-brand colour, rebind to one of these roles — do not carry the "
        "literal through."
    ),
    "rules": (
        "These are owner-set and code-enforced, not preferences. "
        "`primary_font_never_bold` and `primary_font_min_pt` protect the Marsh "
        "Serif voice; `insight_body_pt` sets the floor that makes an insight "
        "callout read as a statement. Code upholds them even when a caller asks "
        "otherwise — that is the point of putting them here rather than in prose."
    ),
}

# --- Per-role colour usage (keyed to BRAND_COLORS roles) --------------------
# One line per semantic role, folded into the colour table so the rationale sits
# beside the value (the co-located-rationale idea, applied to colour).
COLOR_USAGE = {
    "default_fill":     "resting fill for any decorative shape (use unless emphasising)",
    "subtle_highlight": "emphasis that doesn't shout",
    "strong_highlight": "the primary attention-grab",
    "text_highlight":   "mark a word or a rule — not a whole shape",
}

# --- Raw OW palette descriptions (keyed to runtime.PALETTE names) ------------
# The literal OW colour set, surfaced explicitly at the top of the colour
# section. Roles above reference into these.
PALETTE_USAGE = {
    "midnightblue": "primary text / strong fill",
    "cream": "default decorative fill",
    "skyblue": "mid blue (charts, secondary fills)",
    "lblue": "subtle highlight",
    "gold": "text highlight / accent rule",
    "grey":  "secondary text",
    "lgrey": "hairlines and dividers",
    "pale":  "neutral track / pale fill",
    "white": "reversed text / ground",
}

# --- Per-component rationale (keyed to design_components.COMPONENTS names) ---
COMPONENT_PROSE = {
    "kpi_card": (
        "**Use when** a slide needs a small set of headline numbers that each "
        "stand alone — three to four KPIs in a grid, or a single hero figure. "
        "Prefer it over a `dashboard` when the numbers are the message rather "
        "than a supporting scoreboard. The `fill` and `accent` overrides exist "
        "to pick a *different OW role* (e.g. a light-blue emphasis tile), not to "
        "introduce an off-palette colour — a fill that isn't a palette role will "
        "trip the contrast and orphan lints."
    ),
    "stat_card": (
        "**Use when** the figure needs a sentence to land — '€2.4bn, the cost of "
        "doing nothing'. Caption-led, so it carries one stat with context rather "
        "than a bare number. For several bare numbers side by side, use "
        "`kpi_card` instead."
    ),
    "quote_card": (
        "**Use when** a verbatim voice earns its own tile — a customer line, an "
        "expert sentence. One per slide; the oversized mark is the visual, so "
        "keep the quote tight. The light-blue mark is deliberately quiet so the "
        "words, not the punctuation, carry."
    ),
    "icon_card": (
        "The default grid cell: icon, heading, body. **Use when** content is a "
        "set of parallel points (three to four) that each want a glyph and a "
        "line. If the points have no real icon, that is usually a sign the "
        "content is a list, not a card grid — reach for bullets instead of "
        "forcing decorative icons."
    ),
    "insight_panel": (
        "**Use when** a split layout needs its 'so what' stated as a single "
        "sentence beside the evidence. It is the only place Marsh Serif appears "
        "at the insight size with the amber rule, so it should be rare — one per "
        "slide at most, or it stops signalling."
    ),
    "conclusion": (
        "The bottom-pinned takeaway with the amber left rule — the line the "
        "reader should leave with. Keep it to ~one line; it grows upward from a "
        "fixed baseline, so a long conclusion eats the content area rather than "
        "overflowing. Pair it with a `footnote` for the source."
    ),
}
