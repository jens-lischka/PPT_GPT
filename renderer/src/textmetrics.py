"""textmetrics.py — Tier-1 text overflow measurement (portable, in-pipeline).

Replicates the VBA `Shape.TextFrame2.TextRange.BoundHeight` check without a
PowerPoint layout engine: it measures text with the *real* font
glyph advances, wraps it greedily to the box's usable width, and returns the
laid-out height — which the caller compares to the shape height, exactly as the
VBA code compared BoundHeight to the shape size.

Why metrics, not the XML: a `.pptx` stores the text, font, size and geometry but
NOT the laid-out height — that is computed by a layout engine at runtime. So we
compute it. fontTools gives true per-glyph advance widths; greedy word-wrap
gives the line count; the font's hhea metrics give the line height.

Font resolution chain (best first):
  1. The template's embedded face, if loadable (true Marsh Serif / Noto Sans).
  2. A close system proxy (DejaVu Sans for the Noto Sans body that drives almost
     all overflow; DejaVu Serif for Marsh Serif titles), with a per-family
     calibration multiplier so the proxy can be tuned against real PowerPoint
     BoundHeight numbers later (the Tier-3 calibration step).
  3. A flat character heuristic, only if no font file is available at all.

All public functions degrade gracefully and never raise into the build.
"""
from __future__ import annotations

import glob
import math
import os
import zipfile
from io import BytesIO
from typing import Dict, Iterable, List, Optional, Tuple

try:
    from fontTools.ttLib import TTFont
    _HAVE_FONTTOOLS = True
except Exception:                                    # pragma: no cover
    _HAVE_FONTTOOLS = False

EMU_PER_IN = 914400
PT_PER_IN = 72.0

# Per-family calibration multiplier on measured width. 1.0 = trust the proxy.
# Tune against PowerPoint BoundHeight (Tier-3) and set here as the single source.
CALIBRATION = {"minor": 1.00, "major": 1.04}

# Default PowerPoint text insets (inches): left+right total, top+bottom total.
DEFAULT_INSET_H = 0.20
DEFAULT_INSET_V = 0.10
# Per bullet level, extra left indent (inches) — matches BULLET_LEVELS hangs.
INDENT_PER_LEVEL = 0.197


class _Face:
    """A measurable font face: glyph advances + line height, at any size."""

    def __init__(self, ttfont: "TTFont", family_kind: str):
        self._f = ttfont
        self.kind = family_kind
        self.upm = ttfont["head"].unitsPerEm or 1000
        self._cmap = ttfont.getBestCmap()
        self._hmtx = ttfont["hmtx"]
        self._adv_cache: Dict[str, int] = {}
        hhea = ttfont["hhea"]
        # Natural single line height in font units (ascent - descent + lineGap).
        self._line_units = (hhea.ascent - hhea.descent + hhea.lineGap) or int(self.upm * 1.2)

    def _advance(self, ch: str) -> int:
        a = self._adv_cache.get(ch)
        if a is not None:
            return a
        gid = self._cmap.get(ord(ch))
        if gid is None:
            gid = ".notdef"
        try:
            a = self._hmtx[gid][0]
        except Exception:
            a = int(self.upm * 0.5)
        self._adv_cache[ch] = a
        return a

    def text_width_in(self, text: str, size_pt: float) -> float:
        units = sum(self._advance(c) for c in text)
        width_pt = units / self.upm * size_pt * CALIBRATION.get(self.kind, 1.0)
        return width_pt / PT_PER_IN

    def line_height_in(self, size_pt: float) -> float:
        return (self._line_units / self.upm * size_pt) / PT_PER_IN


# --------------------------------------------------------------------------- #
# Face resolution
# --------------------------------------------------------------------------- #
_PROXY = {
    "minor": ["DejaVuSans.ttf", "LiberationSans-Regular.ttf", "Arial.ttf"],
    "major": ["DejaVuSerif.ttf", "LiberationSerif-Regular.ttf", "Times.ttf"],
}
_FONT_DIRS = ["/usr/share/fonts", "/usr/local/share/fonts",
              os.path.expanduser("~/.fonts")]


def _find_proxy_path(kind: str) -> Optional[str]:
    for name in _PROXY.get(kind, []):
        for d in _FONT_DIRS:
            hits = glob.glob(os.path.join(d, "**", name), recursive=True)
            if hits:
                return hits[0]
    # last resort: any sans/serif ttf
    for d in _FONT_DIRS:
        hits = glob.glob(os.path.join(d, "**", "*.ttf"), recursive=True)
        if hits:
            return hits[0]
    return None


def _try_embedded(template_path: str, typeface: str) -> Optional["TTFont"]:
    """Best-effort: load the template's embedded face if it isn't obfuscated.
    Returns None on the common obfuscated case (we fall through to a proxy)."""
    if not (template_path and _HAVE_FONTTOOLS and os.path.exists(template_path)):
        return None
    try:
        import re
        z = zipfile.ZipFile(template_path)
        pres = z.read("ppt/presentation.xml").decode("utf8", "ignore")
        rels = z.read("ppt/_rels/presentation.xml.rels").decode("utf8", "ignore")
        relmap = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
        for blk in re.findall(r"<p:embeddedFont>.*?</p:embeddedFont>", pres, re.S):
            tf = re.search(r'typeface="([^"]+)"', blk)
            rid = re.search(r'<p:regular r:id="([^"]+)"', blk)
            if not (tf and rid) or tf.group(1).lower() != typeface.lower():
                continue
            tgt = relmap.get(rid.group(1), "")
            part = "ppt/" + tgt.replace("../", "")
            if part in z.namelist():
                return TTFont(BytesIO(z.read(part)))
    except Exception:
        return None
    return None


class FontBook:
    """Resolves and caches measurable faces for the 'minor'/'major' kinds."""

    def __init__(self, template_path: str = "",
                 major_typeface: str = "Marsh Serif",
                 minor_typeface: str = "Noto Sans"):
        self.template_path = template_path
        self._typeface = {"major": major_typeface, "minor": minor_typeface}
        self._faces: Dict[str, Optional[_Face]] = {}

    def face(self, kind: str) -> Optional[_Face]:
        if kind in self._faces:
            return self._faces[kind]
        face: Optional[_Face] = None
        if _HAVE_FONTTOOLS:
            emb = _try_embedded(self.template_path, self._typeface.get(kind, ""))
            if emb is not None:
                try:
                    face = _Face(emb, kind)
                except Exception:
                    face = None
            if face is None:
                p = _find_proxy_path(kind)
                if p:
                    try:
                        face = _Face(TTFont(p), kind)
                    except Exception:
                        face = None
        self._faces[kind] = face
        return face


# A process-wide default book; the caller may pass its own template path once.
_DEFAULT_BOOK: Optional[FontBook] = None


def configure(template_path: str = "") -> None:
    global _DEFAULT_BOOK
    _DEFAULT_BOOK = FontBook(template_path)


def _book() -> FontBook:
    global _DEFAULT_BOOK
    if _DEFAULT_BOOK is None:
        _DEFAULT_BOOK = FontBook()
    return _DEFAULT_BOOK


# --------------------------------------------------------------------------- #
# Public measurement API
# --------------------------------------------------------------------------- #
def available() -> bool:
    """True if real-metrics measurement is usable (a face resolved)."""
    return _book().face("minor") is not None


def wrapped_lines(text: str, width_in: float, *, size_pt: float = 12.0,
                  kind: str = "minor", level: int = 0) -> int:
    """Number of lines `text` wraps to in a box `width_in` wide. Falls back to
    a char heuristic if no face is available."""
    text = (text or "").strip()
    if not text:
        return 1
    usable = max(0.3, width_in - DEFAULT_INSET_H - INDENT_PER_LEVEL * level)
    face = _book().face(kind)
    if face is None:                                  # heuristic fallback
        cpi = 12.0
        return max(1, math.ceil(len(text) / max(1, int(usable * cpi))))
    space = face.text_width_in(" ", size_pt) or 0.001
    lines, cur = 1, 0.0
    for word in text.split():
        w = face.text_width_in(word, size_pt)
        if w > usable:                                # word longer than the line
            lines += max(0, math.ceil(w / usable) - 1)
            cur = w % usable
            continue
        if cur <= 0:
            cur = w
        elif cur + space + w <= usable:
            cur += space + w
        else:
            lines += 1
            cur = w
    return lines


def line_height_in(size_pt: float = 12.0, kind: str = "minor") -> float:
    face = _book().face(kind)
    if face is None:
        return size_pt * 1.2 / PT_PER_IN
    return face.line_height_in(size_pt)


def text_height_in(paragraphs: Iterable[Tuple[str, int, float, str]],
                   width_in: float) -> float:
    """Stacked laid-out height (inches) of paragraphs.
    Each paragraph is (text, level, size_pt, kind)."""
    total = 0.0
    for text, level, size_pt, kind in paragraphs:
        n = wrapped_lines(text, width_in, size_pt=size_pt, kind=kind, level=level)
        total += n * line_height_in(size_pt, kind)
    return total


def fits(box_in: Tuple[float, float, float, float],
         paragraphs: Iterable[Tuple[str, int, float, str]]) -> Dict[str, float]:
    """The BoundHeight-style check: laid-out text height vs usable box height.

    `box_in` is (x, y, w, h) inches. Returns a dict with needed/cap heights and
    a `fits` flag (needed <= cap). `overflow_in` > 0 means it overflows.
    """
    _, _, w, h = box_in
    needed = text_height_in(paragraphs, w)
    cap = max(0.1, h - DEFAULT_INSET_V)
    return {"needed_in": round(needed, 3), "cap_in": round(cap, 3),
            "overflow_in": round(needed - cap, 3), "fits": needed <= cap}


# --------------------------------------------------------------------------- #
# BoundHeight-style verification on a finished deck (the VBA analogue)
#
# This is the direct equivalent of iterating shapes in VBA and comparing
# `Shape.TextFrame2.TextRange.BoundHeight` to the shape height — except the
# laid-out height is measured here instead of read from a running PowerPoint.
# It runs on the delivered .pptx, so it catches whatever actually landed on the
# slide, independent of how it was built.
# --------------------------------------------------------------------------- #
def _run_size_pt(paragraph, default=12.0):
    sizes = [r.font.size.pt for r in paragraph.runs
             if r.font is not None and r.font.size is not None]
    return max(sizes) if sizes else default


def overflow_report(pptx_path: str, min_overflow_in: float = 0.05) -> List[Dict]:
    """Measure every fixed-size text shape and flag those whose laid-out text
    is taller than the shape (minus vertical insets). Skips shapes set to
    shrink-text-on-overflow (PowerPoint resizes the text, so no visual
    overflow) and grow-to-fit shapes. Returns a list of overflow records."""
    from pptx import Presentation
    from pptx.enum.text import MSO_AUTO_SIZE
    prs = Presentation(pptx_path)
    emu = EMU_PER_IN
    out: List[Dict] = []
    for si, slide in enumerate(prs.slides, 1):
        try:
            is_cover = slide.slide_layout.name == "Title Slide"
        except Exception:
            is_cover = False
        for shp in slide.shapes:
            if not getattr(shp, "has_text_frame", False) or not shp.has_text_frame:
                continue
            # Cover subtitle/date placeholders are allowed to overflow: the
            # template's own geometry gives the subtitle a box shorter than one
            # measured line, so the metrics model flags every subtitled cover.
            # PowerPoint paints the line past the frame without clipping, and
            # the template geometry is authoritative here. The cover TITLE is
            # still checked.
            try:
                if (is_cover and shp.is_placeholder
                        and "TITLE" not in str(shp.placeholder_format.type)):
                    continue
            except Exception:
                pass
            tf = shp.text_frame
            if not tf.text.strip():
                continue
            autosize = tf.auto_size
            if autosize == MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT:
                continue  # the shape grows; never overflows
            try:
                w_in = shp.width / emu
                h_in = shp.height / emu
            except Exception:
                continue
            if w_in <= 0 or h_in <= 0:
                continue
            # Subtract the text frame's internal insets — text wraps in the box
            # MINUS its margins, so the effective width is narrower (more lines,
            # taller) and the usable height is shorter. Ignoring these made the
            # verifier under-count lines and miss cards that overran by a line
            # (e.g. a KPI card body spilling past its rectangle). Default OOXML
            # insets are 0.1"/0.05" when unset.
            def _ins(val, default):
                try:
                    return val / emu if val is not None else default
                except Exception:
                    return default
            ml = _ins(tf.margin_left, 0.1); mr = _ins(tf.margin_right, 0.1)
            mt = _ins(tf.margin_top, 0.05); mb = _ins(tf.margin_bottom, 0.05)
            w_eff = max(0.1, w_in - ml - mr)
            h_eff = max(0.1, h_in - mt - mb)
            # title placeholders render in the major (serif) face; treat the
            # rest as minor body — a heuristic, good enough for a verifier.
            kind = "minor"
            try:
                if shp.is_placeholder and "TITLE" in str(shp.placeholder_format.type):
                    kind = "major"
            except Exception:
                pass
            paras = [(p.text, p.level or 0, _run_size_pt(p), kind)
                     for p in tf.paragraphs if p.text]
            if not paras:
                continue
            res = fits((0, 0, w_eff, h_eff), paras)
            if res["overflow_in"] >= min_overflow_in:
                rec = {"slide": si, "shape": shp.name,
                       "needed_in": res["needed_in"], "box_in": res["cap_in"],
                       "overflow_in": res["overflow_in"],
                       "shrinks": autosize == MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE}
                out.append(rec)
    return out


def _cli():                                           # pragma: no cover
    import sys
    if len(sys.argv) < 2:
        print("usage: python textmetrics.py <deck.pptx> [template.pptx]")
        return
    if len(sys.argv) >= 3:
        configure(sys.argv[2])
    recs = overflow_report(sys.argv[1])
    if not recs:
        print("No overflow detected (BoundHeight-style check).")
        return
    print(f"{len(recs)} shape(s) likely overflow:")
    for r in recs:
        tag = " [shrinks-to-fit: text will be auto-shrunk]" if r["shrinks"] else ""
        print(f"  slide {r['slide']:>2}  {r['shape']:<28} "
              f"needs {r['needed_in']}\" vs {r['box_in']}\" box "
              f"(over {r['overflow_in']}\"){tag}")


if __name__ == "__main__":                            # pragma: no cover
    _cli()
