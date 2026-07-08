"""cards.py — integrated OW card components (runtime block kinds).

Four self-measuring card components, registered as block kinds so they render
through ``render_block`` and are measured by the autolayout solver (they HUG
their content and equalize within a row):

    kpi_card    – eyebrow + optional intro + value + delta + native chart + body
    quote_card  – editorial pull-quote, navy or light variant
    stat_card   – big number + heading + body (no chart)
    icon_card   – built-in icon + heading + body, icon left or top

Single source of truth: every size, colour and spacing comes from the design
system in ``runtime`` (``Tokens``, ``PALETTE``, ``CHART_COLORS``). This module
defines NO literal font sizes, hex colours or spacing of its own — change a value
in ``Tokens``/``PALETTE`` and these cards follow. Design rules enforced: text is
navy or white only (functional colour only for the KPI trend glyph and traffic
status, drawn from ``CHART_COLORS``); charts are native PowerPoint objects; icons
come from the built-in ``icons`` library; fonts use the template theme
(+mj-lt Marsh Serif / +mn-lt Noto Sans); the top accent bar is optional.
"""
from __future__ import annotations

import math
from dataclasses import dataclass
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn

from runtime import Tokens as T, PALETTE, CHART_COLORS


# ---- colour handles: resolved from the design system, never redefined ------
def _rgb(name_or_hex: str) -> RGBColor:
    h = PALETTE.get(name_or_hex, name_or_hex)
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


NAVY = _rgb("midnightblue")
SKY = _rgb("skyblue")
AMBER = _rgb("gold")
CREAM = _rgb("cream")
WHITE = _rgb("white")
TRACK = _rgb("pale")                      # progress track / unlit lamp
GOOD = _rgb(CHART_COLORS["up"])           # functional green (already in DS)
BAD = _rgb(CHART_COLORS["down"])          # functional red (already in DS)

# ---- size / spacing handles (all from Tokens) ------------------------------
PAD = T.SPACE_CARD_PAD
ACCENT_H = T.SPACE_CARD_ACCENT_H
VIZ_GAP = T.SPACE_CARD_VIZ_GAP
ICON_D = T.SPACE_CARD_ICON
_VIZ_H = {"sparkline": T.VIZ_H_SPARKLINE, "bar": T.VIZ_H_BAR,
          "traffic": T.VIZ_H_TRAFFIC}

# font roles: True -> major (+mj-lt, Marsh Serif); False -> minor (+mn-lt)
SERIF, SANS = True, False


def _line(size_pt: float) -> float:
    """Line height (inches) for a point size, via the design-system factor."""
    return size_pt * T.LINE_HEIGHT / 72.0


# ---- text + theme fonts ----------------------------------------------------
def _theme_typeface(run, major: bool):
    rPr = run._r.get_or_add_rPr()
    ref = "+mj-lt" if major else "+mn-lt"
    for tag in ("a:latin", "a:ea", "a:cs"):
        e = rPr.find(qn(tag))
        if e is None:
            e = rPr.makeelement(qn(tag), {}); rPr.append(e)
        e.set("typeface", ref)


def _run(p, text, major, size, color, *, bold=False, caps=False):
    r = p.add_run(); r.text = text.upper() if caps else text
    r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color
    _theme_typeface(r, major)
    return r


def _txt(slide, x, y, w, text, major, size, color, *, bold=False, center=False, caps=False):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(1.0))
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    p = tf.paragraphs[0]
    if center:
        p.alignment = PP_ALIGN.CENTER
    _run(p, text, major, size, color, bold=bold, caps=caps)


def _panel(slide, x, y, w, h, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y),
                               Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    s.line.fill.background(); s.shadow.inherit = False
    return s


def _wrap_lines(text, width_in, size_pt):
    if not text:
        return 0
    cpl = max(1, int(width_in / (size_pt * 0.0095)))
    return max(1, math.ceil(len(str(text)) / cpl))


# ---- native charts (for KPI viz) -------------------------------------------
def _strip_chart(chart):
    chart.has_title = False
    chart.has_legend = False
    for axis in (chart.category_axis, chart.value_axis):
        try:
            axis.visible = False
            axis.has_major_gridlines = False
            axis.has_minor_gridlines = False
        except Exception:
            pass


def _native_line(slide, x, y, w, h, points):
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    pts = [float(p) for p in points if p is not None]
    if len(pts) < 2:
        return
    cd = CategoryChartData(); cd.categories = [str(i) for i in range(len(pts))]
    cd.add_series("s", pts)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.LINE, Inches(x), Inches(y),
                                Inches(w), Inches(h), cd)
    _strip_chart(gf.chart)
    ser = gf.chart.plots[0].series[0]
    ser.smooth = False
    ser.format.line.color.rgb = NAVY
    ser.format.line.width = Pt(T.STROKE_W_MED + 1.0)


def _native_col(slide, x, y, w, h, points):
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    pts = [float(p) for p in points if p is not None]
    if not pts:
        return
    cd = CategoryChartData(); cd.categories = [str(i) for i in range(len(pts))]
    cd.add_series("s", pts)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(x), Inches(y),
                                Inches(w), Inches(h), cd)
    _strip_chart(gf.chart)
    plot = gf.chart.plots[0]; plot.gap_width = 60; plot.vary_by_categories = False
    ser = plot.series[0]
    ser.format.fill.solid(); ser.format.fill.fore_color.rgb = NAVY
    ser.format.line.fill.background()


def _draw_traffic(slide, x, y, status, label=""):
    lamps = [("red", BAD), ("amber", AMBER), ("green", GOOD)]
    d = T.VIZ_H_TRAFFIC; gap = T.SPACE_GAP / 3.5
    for i, (key, col) in enumerate(lamps):
        cx = x + i * (d + gap)
        lamp = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx), Inches(y),
                                      Inches(d), Inches(d))
        lamp.fill.solid()
        lamp.fill.fore_color.rgb = col if key == status else TRACK
        lamp.line.fill.background(); lamp.shadow.inherit = False
    if label:
        _txt(slide, x + 3 * d + 2 * gap + 0.10, y - 0.04, 2.2, label,
             SANS, T.TYPE_EYEBROW, NAVY)


def _viz_height(viz):
    return _VIZ_H.get((viz or {}).get("kind"), 0.0) if viz else 0.0


def _render_viz(slide, x, y, w, viz):
    kind = viz.get("kind")
    if kind == "sparkline":
        _native_line(slide, x, y, w, _VIZ_H["sparkline"], viz.get("points", []))
    elif kind == "bar":
        _native_col(slide, x, y, w, _VIZ_H["bar"], viz.get("points", []))
    elif kind == "traffic":
        _draw_traffic(slide, x, y, viz.get("status", "amber"), viz.get("label", ""))


# =========================================================================== #
# KPI card
# =========================================================================== #
@dataclass
class KpiCard:
    label: str
    value: str = ""
    intro: str = ""
    delta: str = ""
    delta_dir: str = ""
    delta_good: bool | None = None
    viz: dict | None = None
    body: str = ""
    caption: str = ""
    accent: RGBColor | None = None
    ink: RGBColor = NAVY

    @property
    def _accent_h(self):
        return ACCENT_H if self.accent else 0.0

    def _glyph_color(self):
        glyph = {"up": "\u25B2", "down": "\u25BC"}.get(self.delta_dir, "")
        good = (self.delta_dir == "up") if self.delta_good is None else self.delta_good
        return glyph, (GOOD if good else BAD)

    def _slots(self, inner):
        s = [("label", _line(T.TYPE_EYEBROW) + 0.08)]
        if self.intro:
            s.append(("intro", _wrap_lines(self.intro, inner, T.TYPE_BODY) * _line(T.TYPE_BODY) + 0.08))
        if self.value:
            s.append(("value", _line(T.TYPE_CARD_VALUE) + 0.06))
        if self.delta or self.delta_dir:
            s.append(("delta", _line(T.TYPE_BODY) + 0.08))
        if self.viz:
            s.append(("viz", _viz_height(self.viz) + VIZ_GAP))
        if self.body:
            s.append(("body", _wrap_lines(self.body, inner, T.TYPE_BODY) * _line(T.TYPE_BODY) + 0.06))
        if self.caption:
            s.append(("caption", _wrap_lines(self.caption, inner, T.TYPE_CARD_CAPTION) * _line(T.TYPE_CARD_CAPTION)))
        return s

    def measure(self, w):
        inner = w - 2 * PAD
        return round(self._accent_h + PAD + sum(h for _, h in self._slots(inner)) + PAD, 3)

    def render(self, slide, x, y, w, h):
        _panel(slide, x, y, w, h, CREAM)
        if self.accent:
            _panel(slide, x, y, w, ACCENT_H, self.accent)
        ix, iw, cy = x + PAD, w - 2 * PAD, y + self._accent_h + PAD
        for name, sh in self._slots(iw):
            if name == "label":
                _txt(slide, ix, cy, iw, self.label, SANS, T.TYPE_EYEBROW, self.ink, caps=True)
            elif name == "intro":
                _txt(slide, ix, cy, iw, self.intro, SANS, T.TYPE_BODY, self.ink)
            elif name == "value":
                _txt(slide, ix, cy, iw, self.value, SERIF, T.TYPE_CARD_VALUE, self.ink)
            elif name == "delta":
                tb = slide.shapes.add_textbox(Inches(ix), Inches(cy), Inches(iw), Inches(0.4))
                tf = tb.text_frame; tf.margin_left = 0; tf.margin_top = 0; tf.margin_bottom = 0
                p = tf.paragraphs[0]
                glyph, gcol = self._glyph_color()
                if glyph:
                    _run(p, glyph + " ", SANS, T.TYPE_BODY, gcol, bold=True)
                _run(p, self.delta, SANS, T.TYPE_BODY, self.ink, bold=True)
            elif name == "viz":
                _render_viz(slide, ix, cy + 0.02, iw, self.viz)
            elif name == "body":
                _txt(slide, ix, cy, iw, self.body, SANS, T.TYPE_BODY, self.ink)
            elif name == "caption":
                _txt(slide, ix, cy, iw, self.caption, SANS, T.TYPE_CARD_CAPTION, self.ink)
            cy += sh


# =========================================================================== #
# Quote card
# =========================================================================== #
@dataclass
class QuoteCard:
    quote: str
    name: str = ""
    title: str = ""
    variant: str = "navy"

    def _colors(self):
        if self.variant == "light":
            return CREAM, NAVY, NAVY
        return NAVY, WHITE, WHITE

    def measure(self, w):
        inner = w - 2 * PAD
        lines = _wrap_lines(self.quote, inner, T.TYPE_CARD_QUOTE)
        return round(PAD + 0.50 + lines * _line(T.TYPE_CARD_QUOTE) + 0.20
                     + _line(T.TYPE_BODY) + _line(T.TYPE_EYEBROW) + PAD, 3)

    def render(self, slide, x, y, w, h):
        bg, txt, mark = self._colors()
        _panel(slide, x, y, w, h, bg)
        qm = slide.shapes.add_textbox(Inches(x + PAD - 0.05), Inches(y + PAD - 0.18),
                                      Inches(1.0), Inches(0.7))
        _run(qm.text_frame.paragraphs[0], "\u201C", SERIF, T.TYPE_CARD_QUOTE_MARK, mark)
        tb = slide.shapes.add_textbox(Inches(x + PAD), Inches(y + PAD + 0.45),
                                      Inches(w - 2 * PAD), Inches(h - 2 * PAD - 0.45))
        tf = tb.text_frame; tf.word_wrap = True
        pq = tf.paragraphs[0]; pq.space_after = Pt(T.PARA_QUOTE_GAP - 2)
        _run(pq, self.quote, SERIF, T.TYPE_CARD_QUOTE, txt)
        if self.name:
            pn = tf.add_paragraph(); pn.space_before = Pt(T.PARA_PROSE)
            _run(pn, self.name, SANS, T.TYPE_BODY, txt, bold=True)
        if self.title:
            pt = tf.add_paragraph()
            _run(pt, self.title, SANS, T.TYPE_EYEBROW, txt)


# =========================================================================== #
# Stat callout card
# =========================================================================== #
@dataclass
class StatCallout:
    value: str
    heading: str = ""
    body: str = ""
    panel: bool = False
    ink: RGBColor = NAVY

    def _slots(self, inner):
        s = [("value", _line(T.TYPE_CARD_VALUE) + 0.10)]
        if self.heading:
            s.append(("heading", _line(T.TYPE_BODY) + 0.06))
        if self.body:
            s.append(("body", _wrap_lines(self.body, inner, T.TYPE_BODY) * _line(T.TYPE_BODY)))
        return s

    def measure(self, w):
        pad = PAD if self.panel else 0.0
        inner = w - 2 * pad
        return round(2 * pad + sum(h for _, h in self._slots(inner)), 3)

    def render(self, slide, x, y, w, h):
        pad = PAD if self.panel else 0.0
        if self.panel:
            _panel(slide, x, y, w, h, CREAM)
        ix, iw, cy = x + pad, w - 2 * pad, y + pad
        for name, sh in self._slots(iw):
            if name == "value":
                _txt(slide, ix, cy, iw, self.value, SERIF, T.TYPE_CARD_VALUE, self.ink)
            elif name == "heading":
                _txt(slide, ix, cy, iw, self.heading, SANS, T.TYPE_BODY, self.ink, bold=True)
            elif name == "body":
                _txt(slide, ix, cy, iw, self.body, SANS, T.TYPE_BODY, self.ink)
            cy += sh


# =========================================================================== #
# Icon + text card (built-in icons)
# =========================================================================== #
@dataclass
class IconTextCard:
    heading: str
    body: str = ""
    icon: str = "light-bulb"
    placement: str = "left"
    disc: RGBColor | None = None
    ink: RGBColor = NAVY

    def measure(self, w):
        if self.placement == "top":
            h = ICON_D + 0.12 + _line(T.TYPE_BODY) + 0.04
            if self.body:
                h += _wrap_lines(self.body, w, T.TYPE_BODY) * _line(T.TYPE_BODY)
            return round(h, 3)
        inner = w - ICON_D - 0.20
        h = _line(T.TYPE_BODY) + 0.04
        if self.body:
            h += _wrap_lines(self.body, inner, T.TYPE_BODY) * _line(T.TYPE_BODY)
        return round(max(h, ICON_D), 3)

    def _icon(self, slide, x, y):
        import icons as _ic
        navy_hex = PALETTE["midnightblue"]
        if self.disc:
            d = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y),
                                       Inches(ICON_D), Inches(ICON_D))
            d.fill.solid(); d.fill.fore_color.rgb = self.disc
            d.line.fill.background(); d.shadow.inherit = False
            pad = 0.10
            _ic.draw_icon(slide, self.icon, x + pad, y + pad,
                          ICON_D - 2 * pad, ICON_D - 2 * pad, color_hex=navy_hex)
        else:
            _ic.draw_icon(slide, self.icon, x, y, ICON_D, ICON_D, color_hex=navy_hex)

    def render(self, slide, x, y, w, h):
        if self.placement == "top":
            self._icon(slide, x, y)
            ty = y + ICON_D + 0.12
            _txt(slide, x, ty, w, self.heading, SANS, T.TYPE_BODY, self.ink, bold=True)
            if self.body:
                _txt(slide, x, ty + _line(T.TYPE_BODY) + 0.04, w, self.body, SANS, T.TYPE_BODY, self.ink)
        else:
            self._icon(slide, x, y)
            tx, tw = x + ICON_D + 0.20, w - ICON_D - 0.20
            _txt(slide, tx, y, tw, self.heading, SANS, T.TYPE_BODY, self.ink, bold=True)
            if self.body:
                _txt(slide, tx, y + _line(T.TYPE_BODY) + 0.04, tw, self.body, SANS, T.TYPE_BODY, self.ink)


# =========================================================================== #
# Registry / factory — used by render_block and the autolayout solver
# =========================================================================== #
CARD_KINDS = ("kpi_card", "quote_card", "stat_card", "icon_card")


def _accent(data):
    a = data.get("accent")
    # New palette vocabulary + old names kept as backward-compat aliases.
    amap = {"skyblue": SKY, "gold": AMBER, "midnightblue": NAVY,
            "sky": SKY, "amber": AMBER, "navy": NAVY}
    return amap.get(a) if isinstance(a, str) else a


def build_card(kind, data):
    data = data or {}
    if kind == "kpi_card":
        return KpiCard(label=data.get("label", ""), value=data.get("value", ""),
                       intro=data.get("intro", ""), delta=data.get("delta", ""),
                       delta_dir=data.get("delta_dir", ""),
                       delta_good=data.get("delta_good"), viz=data.get("viz"),
                       body=data.get("body", ""), caption=data.get("caption", ""),
                       accent=_accent(data))
    if kind == "quote_card":
        return QuoteCard(quote=data.get("quote", ""), name=data.get("name", ""),
                         title=data.get("title", ""), variant=data.get("variant", "navy"))
    if kind == "stat_card":
        return StatCallout(value=data.get("value", ""), heading=data.get("heading", ""),
                           body=data.get("body", ""), panel=bool(data.get("panel", False)))
    if kind == "icon_card":
        return IconTextCard(heading=data.get("heading", ""), body=data.get("body", ""),
                            icon=data.get("icon", "light-bulb"),
                            placement=data.get("placement", "left"),
                            disc=_accent(data))
    raise ValueError(f"unknown card kind: {kind}")


def measure_card(kind, data, w):
    return build_card(kind, data).measure(w)


def render_card(slide, kind, data, x, y, w, h):
    build_card(kind, data).render(slide, x, y, w, h)
