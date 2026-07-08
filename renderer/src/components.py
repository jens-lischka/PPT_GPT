"""components.py — composite OW components (text + icon + shape merges).

Brand-true canvas drawers used by the stat_callout / icon_rows / banner /
label_stack / compare intents. Everything is a native PowerPoint shape or
textbox (editable, recolourable), styled with the OW theme palette and fonts.
"""
from __future__ import annotations

from typing import Any, Dict, Sequence

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from runtime import Tokens, PALETTE as _P, FONT_MAJOR as MAJOR, FONT_MINOR as MINOR

# OW theme palette (matches ow_default.pptx)
NAVY = _P["midnightblue"]; CREAM = _P["cream"]; SKY = _P["skyblue"]; LBLUE = _P["lblue"]
AMBER = _P["gold"]; GREY = _P["grey"]; LGREY = _P["lgrey"]; WHITE = _P["white"]
GREEN = "3FA535"; RED = "D0021B"


def _rgb(h: str) -> RGBColor:
    return RGBColor.from_string(h.lstrip("#"))


def _rect(slide, x, y, w, h, *, fill=None, line=None, line_pt=1.0,
          shape=MSO_SHAPE.RECTANGLE):
    sp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    sp.shadow.inherit = False
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = _rgb(fill)
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = _rgb(line); sp.line.width = Pt(line_pt)
    return sp


def _lines(tf, lines, align):
    tf.word_wrap = True
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if ln.get("space_before"):
            p.space_before = Pt(ln["space_before"])
        r = p.add_run(); r.text = ln.get("text", "")
        f = r.font
        f.size = Pt(ln.get("size", 12)); f.name = ln.get("font", MINOR)
        f.bold = ln.get("bold", False)
        f.color.rgb = _rgb(ln.get("color", NAVY))


def _textbox(slide, x, y, w, h, lines, *, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = 0            # no fill -> 0 margins (alignment)
    tf.margin_top = tf.margin_bottom = 0
    _lines(tf, lines, align)
    return tb


def _shape_text(sp, lines, *, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE):
    tf = sp.text_frame
    tf.vertical_anchor = anchor
    tf.margin_left = Inches(0.18); tf.margin_right = Inches(0.18)
    tf.margin_top = Pt(2); tf.margin_bottom = Pt(2)
    _lines(tf, lines, align)


def _embed_lines(shape, lines, *, top=0.2, side=0.2, anchor=MSO_ANCHOR.TOP,
                 align=PP_ALIGN.LEFT):
    """Write text INTO a shape's own text frame (not a separate overlaid box)."""
    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Inches(side); tf.margin_right = Inches(side)
    tf.margin_top = Inches(top); tf.margin_bottom = Inches(0.12)
    _lines(tf, lines, align)


# --------------------------------------------------------------------------- #
# 1. Big-number stat callouts  (number + heading + text on cream cards)
# --------------------------------------------------------------------------- #
def draw_stat_callouts(slide, area, stats: Sequence[Dict[str, Any]]) -> None:
    stats = [s for s in stats if isinstance(s, dict)][:8]
    if not stats:
        return
    n = len(stats); gap = Tokens.SPACE_GAP
    cols = n if n <= 4 else (3 if n <= 6 else 4)
    rows = -(-n // cols)
    cw = (area.w - gap * (cols - 1)) / cols
    ch = min((area.h - gap * (rows - 1)) / rows, 2.7)
    fs_val = 36 if ch >= 2.3 else (28 if ch >= 1.7 else 22)
    has_icon = any(s.get("icon") for s in stats)
    PAD = Tokens.SPACE_ICON_INSET; ICON = Tokens.SPACE_ICON
    top = (PAD + ICON + PAD) if has_icon else PAD      # content clears the icon
    for i, s in enumerate(stats):
        cx = area.x + (i % cols) * (cw + gap)
        cy = area.y + (i // cols) * (ch + gap)
        card = _rect(slide, cx, cy, cw, ch, fill=CREAM)
        if s.get("icon"):
            try:
                import icons as _ic
                _ic.draw_icon(slide, str(s["icon"]), cx + PAD, cy + PAD,
                              ICON, ICON, color_hex=NAVY)
            except Exception:
                pass
        # text is embedded INTO the card shape (single text frame), not overlaid boxes
        _embed_lines(card,
                     [{"text": str(s.get("value", "")), "size": fs_val, "font": MAJOR,
                       "color": NAVY},
                      {"text": s.get("heading", ""), "size": Tokens.TYPE_BODY, "bold": True,
                       "color": NAVY, "space_before": 8},
                      {"text": s.get("text", ""), "size": Tokens.TYPE_BODY, "color": NAVY,
                       "space_before": 3}],
                     top=top, side=PAD)


# --------------------------------------------------------------------------- #
# 2. Icon rows  (fine-line icon + heading + text, stacked, divider rules)
# --------------------------------------------------------------------------- #
def draw_icon_rows(slide, area, rows: Sequence[Dict[str, Any]]) -> None:
    rows = [r for r in rows if isinstance(r, dict)][:8]
    if not rows:
        return
    n = len(rows)
    row_h = min(area.h / n, 1.35)
    for i, r in enumerate(rows):
        ry = area.y + i * row_h
        icon_sz = min(Tokens.SPACE_ICON, row_h * 0.55)
        tx = area.x
        if r.get("icon"):
            try:
                import icons as _ic
                _ic.draw_icon(slide, str(r["icon"]), area.x,
                              ry + (row_h - icon_sz) / 2.0, icon_sz, icon_sz,
                              color_hex=NAVY)
                tx = area.x + icon_sz + 0.28
            except Exception:
                tx = area.x
        _textbox(slide, tx, ry, area.w - (tx - area.x), row_h - 0.12,
                 [{"text": r.get("heading", ""), "size": Tokens.TYPE_BODY, "bold": True,
                   "color": NAVY},
                  {"text": r.get("text", ""), "size": Tokens.TYPE_BODY, "color": NAVY,
                   "space_before": 3}], anchor=MSO_ANCHOR.MIDDLE)
        if i < n - 1:  # divider rule
            _rect(slide, area.x, ry + row_h - 0.06, area.w, 0.012, fill=LGREY)
# --------------------------------------------------------------------------- #
# 4. Infographic label stack  (navy / light-blue / amber stacked labels)
# --------------------------------------------------------------------------- #
_LABEL_CYCLE = [(NAVY, WHITE), (LBLUE, NAVY), (AMBER, NAVY)]
_LABEL_MAP = {"midnightblue": (NAVY, WHITE), "blue": (LBLUE, NAVY),
              "lightblue": (LBLUE, NAVY), "skyblue": (SKY, NAVY),
              "gold": (AMBER, NAVY), "cream": (CREAM, NAVY),
              # backward-compat aliases (old palette names)
              "navy": (NAVY, WHITE), "sky": (SKY, NAVY), "amber": (AMBER, NAVY)}
# --------------------------------------------------------------------------- #
# 5. Compare  (✓ pros / ✗ cons, two columns)
# --------------------------------------------------------------------------- #
def _marker_list(slide, x, y, w, h, heading, items, mark, mark_color):
    if heading:
        _textbox(slide, x, y, w, 0.4,
                 [{"text": heading, "size": Tokens.TYPE_BODY, "bold": True, "color": NAVY}])
        y += 0.5
    row_h = 0.5
    for it in items[:6]:
        _textbox(slide, x, y, 0.4, row_h,
                 [{"text": mark, "size": Tokens.TYPE_MARKER, "bold": True, "color": mark_color}],
                 anchor=MSO_ANCHOR.TOP)
        _textbox(slide, x + 0.42, y, w - 0.42, row_h,
                 [{"text": str(it), "size": Tokens.TYPE_BODY, "color": NAVY}],
                 anchor=MSO_ANCHOR.TOP)
        y += row_h


def draw_compare(slide, area, pros, cons, *, pros_heading="", cons_heading=""):
    half = area.w / 2 - 0.3
    _marker_list(slide, area.x, area.y, half, area.h,
                 pros_heading, list(pros or []), "\u2713", GREEN)
    _marker_list(slide, area.x + half + 0.6, area.y, half, area.h,
                 cons_heading, list(cons or []), "\u2717", RED)
def draw_sticker(slide, area, text, note="") -> None:
    """An amber circular 'sticker' highlight with short navy text; optional note."""
    d = min(2.0, area.h * 0.7, area.w * 0.4)
    cx = area.x + (area.w - (d if not note else area.w * 0.9)) / 2.0
    cy = area.y + max(0, (area.h - d) / 2)
    circ = _rect(slide, cx, cy, d, d, fill=AMBER, shape=MSO_SHAPE.OVAL)
    _shape_text(circ, [{"text": str(text), "size": Tokens.TYPE_MARKER, "bold": True, "color": NAVY}],
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    if note:
        _textbox(slide, cx + d + 0.4, cy, area.w - (cx - area.x) - d - 0.4, d,
                 [{"text": str(note), "size": Tokens.TYPE_BODY, "color": GREY}],
                 anchor=MSO_ANCHOR.MIDDLE)
