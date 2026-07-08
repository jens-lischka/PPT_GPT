"""icons.py — native, editable, recolour-to-brand icons.

Converts a filled-path SVG icon into a native PowerPoint freeform
(`<a:custGeom>`) filled with a brand colour — NOT a rasterised PNG. The result
is a real vector shape a designer can recolour, resize, or nudge in PowerPoint
like any other auto-shape.

Bundled set: Heroicons *outline* (MIT, fine single-weight line) under
``assets/icons/``. Both styles render: FILLED svgs become a solid freeform;
LINE svgs (fill:none + stroke) are auto-detected and drawn as a stroked,
no-fill freeform (brand-coloured ``a:ln`` with round caps), so open strokes
like checkmarks and arrows stay open.

Public API:
    resolve(name)                         -> path to the icon's .svg (keyword match)
    list_icons()                          -> available icon names
    draw_icon(slide, name, x,y,w,h, hex)  -> add the icon to a slide, return shape
"""
from __future__ import annotations

import glob
import os
from io import BytesIO
from typing import List, Optional

from pptx.oxml.ns import nsdecls
from pptx.oxml import parse_xml

try:
    from svgelements import SVG, Path, Move, Line, Close, CubicBezier, QuadraticBezier, Arc
    _HAVE_SVGE = True
except Exception:                                    # pragma: no cover
    _HAVE_SVGE = False

EMU_PER_IN = 914400


def _resolve_icon_dir() -> str:
    here = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(here, "assets", "icons"),                       # src/assets/icons (clean-zip)
        os.path.join(os.path.dirname(here), "assets", "icons"),      # ../assets/icons (runtime_bundle)
        os.path.join(here, "icons"),
    ]
    for c in candidates:
        if os.path.isdir(c):
            return c
    return candidates[0]


_ICON_DIR = _resolve_icon_dir()


# --------------------------------------------------------------------------- #
# Resolution
# --------------------------------------------------------------------------- #
def list_icons(icon_dir: str = _ICON_DIR) -> List[str]:
    return sorted(os.path.splitext(os.path.basename(p))[0]
                  for p in glob.glob(os.path.join(icon_dir, "*.svg")))


# --- alias map: friendly / Heroicons-style names -> actual curated file ---
# Lets the GPT use its natural vocabulary (chart-bar, light-bulb,
# device-phone-mobile, user-group, ...) and still hit a real icon.
_ALIASES = {
    "trending": "trend-up",
    "trending-up": "trend-up",
    "trending-down": "trend-down",
    "team": "group",
    "growth-chart": "money-growth",
    'academic-cap': 'paper-diploma',
    'achievement': 'trophy',
    'adjustments-horizontal': 'filter',
    'alert': 'c-warning',
    'approve': 'check',
    'archive-box': 'box',
    'arrow-down-tray': 'download',
    'arrow-path': 'refresh',
    'arrow-trending-down': 'trend-down',
    'arrow-trending-up': 'money-growth',
    'arrow-up-tray': 'upload',
    'audience': 'group',
    'audit': 'document',
    'automation': 'gear',
    'bank': 'digital-bank',
    'banknotes': 'money-11',
    'bar-chart': 'chart',
    'bars-3': 'menu',
    'battery-100': 'battery',
    'bell-alert': 'bell',
    'bolt': 'energy',
    'briefcase': 'briefcase-24',
    'building': 'ny-building',
    'building-library': 'digital-bank',
    'building-office': 'office',
    'building-office-2': 'ny-building',
    'building-storefront': 'store',
    'calendar-days': 'calendar',
    'cancel': 'remove',
    'candlestick': 'candlestick-chart',
    'carbon': 'leaf',
    'care': 'heart',
    'career': 'briefcase-24',
    'cargo': 'box',
    'case': 'briefcase-24',
    'cash': 'money-11',
    'chart-bar': 'chart',
    'chart-pie': 'pie',
    'chat-bubble-left': 'chat',
    'chat-bubble-left-right': 'chat',
    'check-badge': 'check',
    'check-circle': 'check',
    'chip': 'cpu',
    'circle-stack': 'database',
    'city': 'ny-building',
    'client': 'user',
    'climate': 'leaf',
    'clipboard': 'medical-clipboard',
    'clipboard-document': 'medical-clipboard',
    'close': 'remove',
    'cloud-arrow-up': 'cloud',
    'code-bracket': 'code',
    'cog': 'settings',
    'cog-6-tooth': 'settings',
    'cog-8-tooth': 'settings',
    'coin': 'gold-coin',
    'command-line': 'code',
    'communication': 'chat',
    'community': 'group',
    'complete': 'check',
    'compliance': 'shield',
    'computer-desktop': 'laptop',
    'config': 'settings',
    'course': 'book',
    'cpu-chip': 'cpu',
    'creative': 'bulb',
    'credit-card': 'wallet',
    'cube': 'box',
    'currency': 'currency-dollar',
    'customer': 'user',
    'data': 'database',
    'date': 'calendar',
    'deadline': 'calendar',
    'desktop': 'laptop',
    'device': 'laptop',
    'device-phone-mobile': 'phone',
    'device-tablet': 'laptop',
    'discount': 'discount-2',
    'document-check': 'check',
    'document-text': 'document',
    'dollar': 'currency-dollar',
    'done': 'check',
    'eco': 'leaf',
    'education': 'book',
    'employee': 'user',
    'envelope': 'mail',
    'envelope-open': 'mail',
    'environment': 'leaf',
    'euro': 'currency-euro',
    'exclamation-circle': 'c-warning',
    'exclamation-triangle': 'c-warning',
    'files': 'file',
    'find': 'search',
    'finger-print': 'shield',
    'fingerprint': 'shield',
    'fire': 'energy',
    'fitness': 'heart',
    'folder-open': 'folder',
    'gauge': 'speedometer',
    'global': 'globe',
    'globe-alt': 'globe',
    'globe-americas': 'globe',
    'globe-europe-africa': 'globe',
    'goal': 'target',
    'governance': 'scale',
    'graduation': 'paper-diploma',
    'graph': 'chart',
    'green': 'leaf',
    'growth': 'money-growth',
    'guard': 'shield',
    'hand-thumb-up': 'like',
    'health': 'heart',
    'healthcare': 'heart',
    'heart-pulse': 'pulse',
    'history': 'clock',
    'home-modern': 'home',
    'idea': 'bulb',
    'identification': 'user',
    'inbox': 'mail',
    'industry': 'factory',
    'info': 'c-info',
    'information-circle': 'c-info',
    'innovation': 'bulb',
    'internet': 'globe',
    'investment': 'money-growth',
    'invoice': 'receipt',
    'job': 'briefcase-24',
    'justice': 'scale',
    'knowledge': 'book',
    'kpi': 'speedometer',
    'leader': 'user',
    'learning': 'book',
    'legal': 'scale',
    'light-bulb': 'bulb',
    'lightbulb': 'bulb',
    'list-bullet': 'list',
    'location': 'pin',
    'lock-closed': 'lock',
    'lock-open': 'unlocked',
    'logistics': 'truck-front',
    'magnifying-glass': 'search',
    'magnifying-glass-plus': 'search',
    'medical': 'medical-clipboard',
    'message': 'chat',
    'milestone': 'flag',
    'minus-circle': 'remove',
    'mission': 'target',
    'mobile': 'phone',
    'money': 'money-11',
    'nature': 'leaf',
    'newspaper': 'paper',
    'package': 'box',
    'page': 'document',
    'paper-airplane': 'send',
    'partnership': 'handshake',
    'people': 'group',
    'percentage': 'percentage-38',
    'person': 'user',
    'pie-chart': 'pie',
    'plus-circle': 'add',
    'portfolio': 'briefcase-24',
    'power': 'energy',
    'premium': 'diamond',
    'presentation-chart-bar': 'presentation',
    'presentation-chart-line': 'presentation',
    'price': 'tag',
    'prize': 'trophy',
    'process': 'gear',
    'processor': 'cpu',
    'profit': 'money-growth',
    'protection': 'shield',
    'puzzle': 'puzzle-09',
    'puzzle-piece': 'puzzle-09',
    'quality': 'star',
    'question': 'c-question',
    'question-mark-circle': 'c-question',
    'queue-list': 'list',
    'rating': 'star',
    'receipt-percent': 'receipt',
    'recycle': 'recycling',
    'region': 'map',
    'renewable': 'recycling',
    'report': 'document',
    'rocket-launch': 'rocket',
    'safety': 'shield',
    'schedule': 'calendar',
    'server-stack': 'server',
    'setup': 'settings',
    'shield-check': 'shield',
    'shield-exclamation': 'shield',
    'ship': 'boat',
    'shipping': 'truck-front',
    'signal': 'network',
    'smartphone': 'phone',
    'software': 'code',
    'solar': 'sun',
    'solution': 'puzzle-09',
    'sort': 'filter',
    'sparkles': 'star',
    'square-3-stack-3d': 'layers',
    'squares-2x2': 'grid',
    'startup': 'rocket',
    'statistics': 'analytics',
    'strategy': 'target',
    'success': 'check',
    'sustainability': 'leaf',
    'sync': 'refresh',
    'tablet': 'laptop',
    'team': 'group',
    'time': 'clock',
    'tools': 'gear',
    'training': 'book',
    'transport': 'truck-front',
    'tree': 'leaf',
    'trophy-outline': 'trophy',
    'truck': 'truck-front',
    'unlock': 'unlocked',
    'user-circle': 'user',
    'user-group': 'group',
    'user-plus': 'user',
    'users': 'group',
    'verified': 'check',
    'vision': 'target',
    'warehouse': 'factory',
    'warning': 'c-warning',
    'web': 'globe',
    'wellness': 'heart',
    'work': 'briefcase-24',
    'workflow': 'gear',
    'world': 'globe',
    'wrench': 'gear',
    'wrench-screwdriver': 'gear',
    'x-circle': 'remove',
    'x-mark': 'remove',
    'zoom': 'search',
}


def resolve(name: str, icon_dir: str = _ICON_DIR) -> Optional[str]:
    """Resolve an icon name (exact, then keyword/substring) to an .svg path."""
    if not name:
        return None
    name = name.strip().lower()
    exact = os.path.join(icon_dir, f"{name}.svg")
    if os.path.exists(exact):
        return exact
    aliased = _ALIASES.get(name)
    if aliased:
        ap = os.path.join(icon_dir, f"{aliased}.svg")
        if os.path.exists(ap):
            return ap
    cands = list_icons(icon_dir)
    for c in cands:                                   # substring either direction
        if name in c or c in name:
            return os.path.join(icon_dir, f"{c}.svg")
    # token overlap
    toks = set(name.replace("_", "-").split("-"))
    best, best_score = None, 0
    for c in cands:
        s = len(toks & set(c.split("-")))
        if s > best_score:
            best, best_score = c, s
    return os.path.join(icon_dir, f"{best}.svg") if best else None


# --------------------------------------------------------------------------- #
# SVG path -> DrawingML custGeom
# --------------------------------------------------------------------------- #
def _pt(p, w_emu, h_emu):
    x = max(0, min(w_emu, int(round(p.x))))
    y = max(0, min(h_emu, int(round(p.y))))
    return f'<a:pt x="{x}" y="{y}"/>'


def svg_to_custgeom(svg_bytes: bytes, w_emu: int, h_emu: int,
                    close_open: bool = True) -> str:
    """Build a `<a:custGeom>` whose paths are the SVG's outlines, scaled so the
    icon's viewBox maps to a ``w_emu`` x ``h_emu`` box.

    ``close_open`` — when True (filled icons) every subpath is closed so it can
    be filled. When False (line/stroke icons) subpaths close ONLY on an explicit
    ``Z`` so open strokes (checkmarks, arrows) stay open instead of looping back.
    """
    if not _HAVE_SVGE:
        raise RuntimeError("svgelements is required for icon conversion")
    svg = SVG.parse(BytesIO(svg_bytes), width=w_emu, height=h_emu,
                    ppi=EMU_PER_IN)  # parse into EMU user units
    segs_xml: List[str] = []
    for el in svg.elements():
        if not isinstance(el, (Path,)) and not hasattr(el, "segments"):
            continue
        try:
            path = Path(el) if not isinstance(el, Path) else el
            seglist = list(path.segments())
        except Exception:
            continue
        open_sub = False
        for seg in seglist:
            if isinstance(seg, Move):
                if open_sub and close_open:
                    segs_xml.append("<a:close/>")
                segs_xml.append(f"<a:moveTo>{_pt(seg.end, w_emu, h_emu)}</a:moveTo>")
                open_sub = True
            elif isinstance(seg, Line):
                segs_xml.append(f"<a:lnTo>{_pt(seg.end, w_emu, h_emu)}</a:lnTo>")
            elif isinstance(seg, CubicBezier):
                segs_xml.append("<a:cubicBezTo>"
                                + _pt(seg.control1, w_emu, h_emu)
                                + _pt(seg.control2, w_emu, h_emu)
                                + _pt(seg.end, w_emu, h_emu)
                                + "</a:cubicBezTo>")
            elif isinstance(seg, QuadraticBezier):
                segs_xml.append("<a:quadBezTo>"
                                + _pt(seg.control, w_emu, h_emu)
                                + _pt(seg.end, w_emu, h_emu)
                                + "</a:quadBezTo>")
            elif isinstance(seg, Arc):
                for c in seg.as_cubic_curves():
                    segs_xml.append("<a:cubicBezTo>"
                                    + _pt(c.control1, w_emu, h_emu)
                                    + _pt(c.control2, w_emu, h_emu)
                                    + _pt(c.end, w_emu, h_emu)
                                    + "</a:cubicBezTo>")
            elif isinstance(seg, Close):
                segs_xml.append("<a:close/>")
                open_sub = False
        if open_sub and close_open:
            segs_xml.append("<a:close/>")
    body = "".join(segs_xml)
    return (f'<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:rect l="0" t="0" r="{w_emu}" b="{h_emu}"/>'
            f'<a:pathLst><a:path w="{w_emu}" h="{h_emu}">{body}</a:path></a:pathLst></a:custGeom>')


# --------------------------------------------------------------------------- #
# Slide insertion
# --------------------------------------------------------------------------- #
_SHAPE_ID = [1000]


def draw_icon(slide, name_or_path: str, x_in: float, y_in: float,
              w_in: float, h_in: float, color_hex: str = "000F47",
              icon_dir: str = _ICON_DIR):
    """Add ``name`` as a native freeform icon at (x,y) sized w×h inches, filled
    with ``color_hex``. Returns the created shape element. Keeps the icon
    square (centred) within the given box to avoid distortion."""
    path = name_or_path if os.path.exists(name_or_path) else resolve(name_or_path, icon_dir)
    if not path or not os.path.exists(path):
        raise FileNotFoundError(f"icon not found: {name_or_path}")
    with open(path, "rb") as fh:
        svg_bytes = fh.read()

    # Detect stroke/line icons (fill:none + a real stroke) vs filled icons.
    import re as _re
    head = svg_bytes[:600].decode("utf-8", "ignore")
    is_line = ('fill="none"' in head or "fill:none" in head) and \
              bool(_re.search(r'stroke="(?!none)[^"]+"|stroke:(?!none)', head))
    side = min(w_in, h_in)                            # square, centred
    off_x = x_in + (w_in - side) / 2.0
    off_y = y_in + (h_in - side) / 2.0
    w_emu = h_emu = int(round(side * EMU_PER_IN))
    geom = svg_to_custgeom(svg_bytes, w_emu, h_emu, close_open=not is_line)

    if is_line:
        ln_w = 12700   # fixed 1pt outline (OW spec)
        fill_xml = "<a:noFill/>"
        ln_xml = (f'<a:ln w="{ln_w}" cap="rnd"><a:solidFill>'
                  f'<a:srgbClr val="{color_hex}"/></a:solidFill>'
                  f'<a:round/></a:ln>')
    else:
        fill_xml = f'<a:solidFill><a:srgbClr val="{color_hex}"/></a:solidFill>'
        ln_xml = "<a:ln><a:noFill/></a:ln>"

    _SHAPE_ID[0] += 1
    sid = _SHAPE_ID[0]
    nm = os.path.splitext(os.path.basename(path))[0]
    xml = (
        f'<p:sp {nsdecls("p", "a")}>'
        f'<p:nvSpPr><p:cNvPr id="{sid}" name="icon-{nm}"/>'
        f'<p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr>'
        f'<a:xfrm><a:off x="{int(off_x*EMU_PER_IN)}" y="{int(off_y*EMU_PER_IN)}"/>'
        f'<a:ext cx="{w_emu}" cy="{h_emu}"/></a:xfrm>'
        f'{geom}'
        f'{fill_xml}'
        f'{ln_xml}'
        f'</p:spPr>'
        f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'
        f'</p:sp>'
    )
    el = parse_xml(xml)
    slide.shapes._spTree.append(el)
    return el
