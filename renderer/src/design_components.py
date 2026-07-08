"""design_components.py — the structured component layer for the design system.

This is the OW analogue of the ``components:`` block in Google's DESIGN.md, but
inverted to fit our architecture: the *renderers* are the source of truth, so a
component entry never re-describes how a block draws. It declares only the
contract the rest of the system needs to reason about a block:

  - ``slots``   — the content shape the block consumes (what the GPT must supply)
  - ``tokens``  — which design tokens / palette roles the block BINDS, so the
                  linter can run orphan + contrast checks against real usage
  - ``overrides``— the per-instance knobs a deck author may pass (e.g. ``fill``)

Every token value here is a *reference* (``{palette.gold}``, ``{Tokens.TYPE_KPI}``,
``{BRAND_COLORS.default_fill}``) — never a literal — exactly so the linter can
resolve it and fail on a broken reference. The resolved value is attached at
generation time for the human-readable doc; it is not stored here.

To add a block to the inventory: add an entry below. The slot schema and the
bound tokens are declarative; the *prose* ("use when…") lives in
``design_prose.py`` keyed by the same name, and the generator stitches the two
together. Keep this list to the blocks that actually recur and/or take
overrides — formalising every one-off design buys nothing but sync cost.
"""
from __future__ import annotations

from typing import Dict, List


# A slot is (name, type, required, note). ``type`` is a coarse author-facing
# kind, not a Python type: "string" | "array<obj>" | "number" | "enum".
def _slot(name: str, type_: str, required: bool, note: str = "",
          **extra) -> dict:
    s = {"name": name, "type": type_, "required": required, "note": note}
    s.update(extra)
    return s


# ---------------------------------------------------------------------------
# Component inventory
#
# tokens: maps a semantic role used by the renderer -> a token REFERENCE.
#   {palette.X}        -> runtime.PALETTE[X]            (raw OW hex)
#   {BRAND_COLORS.X}   -> runtime.BRAND_COLORS[X]       (semantic role -> hex/accent)
#   {Tokens.X}         -> runtime.Tokens.X              (size / spacing token)
# contrast_pairs: list of (background_role, text_role) the linter checks for AA.
# ---------------------------------------------------------------------------
COMPONENTS: Dict[str, dict] = {
    "kpi_card": {
        "renderer": "cards.KpiCard",
        "summary": "A single big-number stat tile: value, label, optional delta.",
        "slots": [
            _slot("value", "string", True, "the headline number, e.g. '+150%'"),
            _slot("label", "string", True, "what the number measures"),
            _slot("delta", "string", False, "movement vs a baseline, e.g. '▲ 12 pts'"),
            _slot("icon", "string", False, "icon keyword shown top-left"),
        ],
        "tokens": {
            "fill":        "{BRAND_COLORS.default_fill}",
            "value_type":  "{Tokens.TYPE_CARD_VALUE}",
            "label_type":  "{Tokens.TYPE_BODY}",
            "caption_type":"{Tokens.TYPE_CARD_CAPTION}",
            "value_color": "{palette.midnightblue}",
            "label_color": "{palette.midnightblue}",
            "pad":         "{Tokens.SPACE_CARD_PAD}",
            "accent_h":    "{Tokens.SPACE_CARD_ACCENT_H}",
        },
        "contrast_pairs": [("fill", "value_color"), ("fill", "label_color")],
        "overrides": {
            "fill":  {"optional": True, "accepts": ["palette-role", "ow-hex"]},
            "accent":{"optional": True, "accepts": ["palette-role", "ow-hex"]},
        },
    },
    "stat_card": {
        "renderer": "cards.StatCallout",
        "summary": "A stat with a one-line callout — like kpi_card but caption-led.",
        "slots": [
            _slot("value", "string", True, "the figure"),
            _slot("label", "string", True, "the callout sentence under the figure"),
            _slot("icon", "string", False, "icon keyword"),
        ],
        "tokens": {
            "fill":        "{BRAND_COLORS.default_fill}",
            "value_type":  "{Tokens.TYPE_CARD_VALUE}",
            "label_type":  "{Tokens.TYPE_BODY}",
            "value_color": "{palette.midnightblue}",
            "label_color": "{palette.midnightblue}",
            "pad":         "{Tokens.SPACE_CARD_PAD}",
        },
        "contrast_pairs": [("fill", "value_color"), ("fill", "label_color")],
        "overrides": {
            "fill":  {"optional": True, "accepts": ["palette-role", "ow-hex"]},
            "accent":{"optional": True, "accepts": ["palette-role", "ow-hex"]},
        },
    },
    "quote_card": {
        "renderer": "cards.QuoteCard",
        "summary": "A pull-quote tile with an oversized opening quote mark.",
        "slots": [
            _slot("quote", "string", True, "the quotation text"),
            _slot("attribution", "string", False, "speaker / source"),
        ],
        "tokens": {
            "fill":        "{BRAND_COLORS.default_fill}",
            "quote_type":  "{Tokens.TYPE_CARD_QUOTE}",
            "mark_type":   "{Tokens.TYPE_CARD_QUOTE_MARK}",
            "caption_type":"{Tokens.TYPE_CARD_CAPTION}",
            "quote_color": "{palette.midnightblue}",
            "mark_color":  "{palette.lblue}",
            "pad":         "{Tokens.SPACE_CARD_PAD}",
        },
        "contrast_pairs": [("fill", "quote_color")],
        "overrides": {
            "fill":  {"optional": True, "accepts": ["palette-role", "ow-hex"]},
        },
    },
    "icon_card": {
        "renderer": "cards.IconTextCard",
        "summary": "Icon + heading + body text card; the default grid cell.",
        "slots": [
            _slot("icon", "string", True, "icon keyword shown top-left"),
            _slot("heading", "string", True, "card heading"),
            _slot("body", "string", False, "supporting line(s)"),
        ],
        "tokens": {
            "fill":         "{BRAND_COLORS.default_fill}",
            "heading_type": "{Tokens.TYPE_HEADING}",
            "body_type":    "{Tokens.TYPE_BODY}",
            "heading_color":"{palette.midnightblue}",
            "body_color":   "{palette.midnightblue}",
            "icon_box":     "{Tokens.SPACE_CARD_ICON}",
            "pad":          "{Tokens.SPACE_CARD_PAD}",
        },
        "contrast_pairs": [("fill", "heading_color"), ("fill", "body_color")],
        "overrides": {
            "fill":  {"optional": True, "accepts": ["palette-role", "ow-hex"]},
        },
    },
    "insight_panel": {
        "renderer": "runtime.draw_insight",
        "summary": "The strong short-statement callout (Marsh Serif, 18pt floor) "
                   "with the amber rule — the 'signal' on split layouts.",
        "slots": [
            _slot("insight", "string", True, "one-sentence signal / 'so what'"),
        ],
        "tokens": {
            "body_type":  "{BRAND_RULES.insight_body_pt}",
            "body_color": "{palette.midnightblue}",
            "rule_color": "{BRAND_COLORS.text_highlight}",
            "rule_w":     "{Tokens.STROKE_PANEL}",
        },
        "contrast_pairs": [],   # text sits on the slide ground, not a fill
        "overrides": {},
    },
    "conclusion": {
        "renderer": "runtime.add_conclusion",
        "summary": "The bottom-pinned key-takeaway line with the amber left rule.",
        "slots": [
            _slot("conclusion", "string", True, "~one-line takeaway"),
        ],
        "tokens": {
            "body_type":  "{Tokens.TYPE_CONCLUSION}",
            "body_color": "{palette.midnightblue}",
            "rule_color": "{BRAND_COLORS.text_highlight}",
            "height":     "{Tokens.SPACE_CONCL_H}",
        },
        "contrast_pairs": [],
        "overrides": {},
    },
}


def component_names() -> List[str]:
    return sorted(COMPONENTS.keys())


def all_bound_token_refs() -> List[str]:
    """Every token reference bound by any component — the input to the
    orphaned-token lint (a palette role nobody binds is a warning)."""
    refs: List[str] = []
    for spec in COMPONENTS.values():
        refs.extend(spec.get("tokens", {}).values())
    return refs
