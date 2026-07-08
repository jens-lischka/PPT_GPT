"""images.py — image placement for the OW runtime.

Three things, per requirements:

1. **Swappable placeholders.** Every image is inserted as a named, alt-tagged
   shape ("REPLACE:<name>") so a designer can later right-click -> Change
   Picture (which preserves the box geometry and crop) — or, when no asset
   exists yet, a labelled stand-in panel marks the exact box to fill.

2. **One source, many crops.** A source image is registered once; each
   placement reuses the SAME media part (python-pptx dedupes identical blobs)
   and applies its own non-destructive crop via srcRect, so the same photo can
   be cropped 16:9 on a cover, square for a bio, and a wide strip for a divider
   — and the designer can re-adjust any crop later without re-importing.

3. **Source: upload or generate.** User-supplied files just work. AI generation
   is behind a provider seam (off by default); in environments without an image
   API it falls back to a deterministic brand-tinted stand-in so the pipeline
   is runnable end-to-end, never inventing data — images are decorative only.

Public API:
    reg = ImageRegistry(); reg.register("hero", path_or_None, prompt=...)
    place_image(slide, reg, "hero", x,y,w,h, aspect="16:9", name="cover_hero")
    generate_image(prompt, aspect, out_path, provider=None)
"""
from __future__ import annotations

import hashlib
import os
from dataclasses import dataclass
from typing import Dict, Optional, Tuple

from pptx.util import Inches, Pt

try:
    from PIL import Image, ImageDraw
    _HAVE_PIL = True
except Exception:                                    # pragma: no cover
    _HAVE_PIL = False


# --------------------------------------------------------------------------- #
# Registry
# --------------------------------------------------------------------------- #
@dataclass
class _Source:
    id: str
    path: Optional[str] = None      # resolved image file, or None if not ready
    prompt: Optional[str] = None    # generation prompt, if any
    native_aspect: Optional[float] = None


class ImageRegistry:
    """Holds named source images so one photo can be reused across slides."""

    def __init__(self):
        self._sources: Dict[str, _Source] = {}

    def register(self, source_id: str, path: Optional[str] = None,
                 prompt: Optional[str] = None) -> _Source:
        ar = _native_aspect(path) if path else None
        s = _Source(id=source_id, path=path, prompt=prompt, native_aspect=ar)
        self._sources[source_id] = s
        return s

    def get(self, source_id: str) -> Optional[_Source]:
        return self._sources.get(source_id)

    def ensure_ready(self, source_id: str, aspect: Optional[float] = None,
                     out_dir: str = "/tmp", allow_generation: bool = False) -> Optional[str]:
        """Return a usable path, generating from the prompt if allowed."""
        s = self._sources.get(source_id)
        if not s:
            return None
        if s.path and os.path.exists(s.path):
            return s.path
        if s.prompt and allow_generation:
            out = os.path.join(out_dir, f"gen_{source_id}.png")
            generate_image(s.prompt, aspect or 16 / 9, out)
            s.path = out
            s.native_aspect = _native_aspect(out)
            return out
        return None


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def _native_aspect(path: Optional[str]) -> Optional[float]:
    if not (path and _HAVE_PIL and os.path.exists(path)):
        return None
    try:
        with Image.open(path) as im:
            return im.width / im.height if im.height else None
    except Exception:
        return None


def parse_aspect(aspect) -> Optional[float]:
    if aspect is None:
        return None
    if isinstance(aspect, (int, float)):
        return float(aspect)
    s = str(aspect).strip()
    if ":" in s:
        a, b = s.split(":", 1)
        return float(a) / float(b)
    try:
        return float(s)
    except Exception:
        return None


def _fill_crop(native_ar: float, box_ar: float) -> Tuple[float, float, float, float]:
    """Center-crop fractions (l, t, r, b) so a native-aspect image fills a box
    of aspect ``box_ar`` without distortion."""
    if not native_ar or not box_ar:
        return (0.0, 0.0, 0.0, 0.0)
    if native_ar > box_ar:                            # too wide -> crop sides
        frac = (1 - box_ar / native_ar) / 2.0
        return (frac, 0.0, frac, 0.0)
    frac = (1 - native_ar / box_ar) / 2.0             # too tall -> crop top/bottom
    return (0.0, frac, 0.0, frac)


def _set_name_descr(shape, name: str, descr: str):
    cNvPr = shape._element.nvPicPr.cNvPr if shape.shape_type == 13 else shape._element._nvXxPr.cNvPr
    cNvPr.set("name", name)
    cNvPr.set("descr", descr)


# --------------------------------------------------------------------------- #
# Placement
# --------------------------------------------------------------------------- #
def place_image(slide, registry: "ImageRegistry", source_id: str,
                x_in: float, y_in: float, w_in: float, h_in: float, *,
                aspect=None, crop: Optional[Dict[str, float]] = None,
                name: Optional[str] = None, allow_generation: bool = False,
                out_dir: str = "/tmp"):
    """Place ``source_id`` into the box (x,y,w,h inches). If the source has no
    file yet, draw a labelled swap-target placeholder instead. Returns the
    created shape. ``aspect`` (e.g. "16:9") center-crops to fill the box;
    ``crop`` (fractions l/t/r/b) overrides with an explicit crop."""
    base = name or source_id
    shape_name = f"REPLACE:{base}"            # placeholder name (swap target)
    box_ar = w_in / h_in if h_in else None
    path = registry.ensure_ready(source_id, box_ar, out_dir, allow_generation)

    if not path:                                      # ---- placeholder panel ----
        src = registry.get(source_id)
        desc = (src.prompt if src and src.prompt else
                (name if name and " " in name else ""))
        prompt_text = build_image_prompt(desc, aspect or
                                         (f"{box_ar:.2f}" if box_ar else "4:3"))
        return _draw_placeholder(slide, x_in, y_in, w_in, h_in, shape_name,
                                 prompt_text,
                                 aspect or (f"{box_ar:.2f}" if box_ar else "4:3"))

    pic = slide.shapes.add_picture(path, Inches(x_in), Inches(y_in),
                                   Inches(w_in), Inches(h_in))
    # Non-destructive crop (srcRect) — same media part can be cropped differently
    # on every slide; designers can re-adjust later via Picture Format > Crop.
    if crop:
        pic.crop_left = float(crop.get("left", 0.0))
        pic.crop_top = float(crop.get("top", 0.0))
        pic.crop_right = float(crop.get("right", 0.0))
        pic.crop_bottom = float(crop.get("bottom", 0.0))
    else:
        ar = parse_aspect(aspect) or registry.get(source_id).native_aspect
        if ar and box_ar:
            l, t, r, b = _fill_crop(ar, box_ar)
            pic.crop_left, pic.crop_top, pic.crop_right, pic.crop_bottom = l, t, r, b
    _set_name_descr(pic, f"IMAGE:{base}",
                    f"Swappable image '{name or source_id}'. Right-click > Change "
                    f"Picture to replace; crop and position are preserved.")
    return pic


_PHOTO_RULES = ("Photorealistic, real photography, natural lighting, sharp "
                "focus, no CGI, no illustration, no text or logos")


def build_image_prompt(desc: str, aspect_label: str) -> str:
    """A ready-to-paste ChatGPT prompt that always yields a photoreal image."""
    desc = (desc or "").strip().rstrip(".")
    core = (f"A photograph of {desc}." if desc
            else "[Describe the photo you want here].")
    return f"{core} {_PHOTO_RULES}. Aspect ratio {aspect_label}."


def _draw_placeholder(slide, x, y, w, h, shape_name, prompt_text, aspect_label):
    """A stand-in box showing a copy-paste, photoreal ChatGPT image prompt.
    The shape keeps its ``REPLACE:`` name so designers can still swap it."""
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.dml.color import RGBColor
    sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y),
                                Inches(w), Inches(h))
    sp.fill.solid(); sp.fill.fore_color.rgb = RGBColor(0xEF, 0xEC, 0xE7)  # warm grey
    sp.line.color.rgb = RGBColor(0xB9, 0xB3, 0xAA); sp.line.width = Pt(1)
    sp.shadow.inherit = False
    tf = sp.text_frame; tf.word_wrap = True
    tf.text = "\u25a2  IMAGE PROMPT  (paste into ChatGPT)"
    p2 = tf.add_paragraph(); p2.text = prompt_text
    p3 = tf.add_paragraph(); p3.text = aspect_label
    sizes = (10, 11, 9)
    for i, p in enumerate(tf.paragraphs):
        p.alignment = 1  # left — it's a prompt to read/copy
        for r in p.runs:
            r.font.size = Pt(sizes[i] if i < len(sizes) else 9)
            r.font.color.rgb = RGBColor(0x6B, 0x66, 0x5E)
            if i == 0:
                r.font.bold = True
    sp.name = shape_name
    sp._element.nvSpPr.cNvPr.set(
        "descr", "Image prompt — copy the prompt line into ChatGPT to generate a "
                 "photorealistic image, then replace this box with it.")
    return sp


# --------------------------------------------------------------------------- #
# Generation (provider seam + offline fallback)
# --------------------------------------------------------------------------- #
def generate_image(prompt: str, aspect, out_path: str, provider=None) -> str:
    """Generate an image for a *decorative* slot. With a real provider wired in
    (env key), call it here. Otherwise produce a deterministic brand-tinted
    stand-in so the pipeline runs end-to-end without inventing data."""
    ar = parse_aspect(aspect) or 16 / 9
    w = 1280
    h = int(w / ar)
    if provider is not None and callable(provider):   # real provider seam
        return provider(prompt, (w, h), out_path)
    if not _HAVE_PIL:                                 # pragma: no cover
        open(out_path, "wb").close()
        return out_path
    # Deterministic gradient keyed on the prompt, in OW navy->cream.
    seed = int(hashlib.sha1(prompt.encode()).hexdigest(), 16)
    img = Image.new("RGB", (w, h))
    px = img.load()
    top = (0x00, 0x0F, 0x47)            # navy
    bot = (0xF7, 0xF3, 0xEE)            # cream
    skew = (seed % 40) / 100.0
    for j in range(h):
        f = (j / h) ** (1.0 + skew)
        col = tuple(int(top[k] + (bot[k] - top[k]) * f) for k in range(3))
        for i in range(w):
            px[i, j] = col
    d = ImageDraw.Draw(img)
    d.text((24, h - 40), "[generated placeholder — replace before delivery]",
           fill=(0xFF, 0xBF, 0x00))
    img.save(out_path)
    return out_path
