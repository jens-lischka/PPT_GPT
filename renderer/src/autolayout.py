"""autolayout.py — content-measured grid solver (Figma-style intrinsic sizing).

The stock ``draw_grid`` divides the content area into **equal tracks**: a cell's
height comes from how many rows the grid has, not from what is in the cell. Text
that exceeds its track overflows, and that overflow is only caught afterwards by
the gate (build -> fail -> re-author -> rebuild).

This module promotes the runtime's own ``textmetrics`` measurement from a
post-hoc gate input to a **solver input**. Each cell is measured; rows size to
their content (Figma "Hug contents"); a flexible cell can absorb leftover space
("Fill container"); explicit sizes are honoured ("Fixed"). When the measured
stack still exceeds the slide, a deterministic resolution ladder runs
(tighten gutter -> compress within floors -> flag residual) so overflow is
resolved *in the solver*, not by a rebuild loop.

It is intentionally dependency-light: it imports ``textmetrics`` (already in the
runtime) and takes the token values + a content area as plain inputs, so it can
be unit-tested in isolation and dropped into ``draw_grid`` as a placement source.

Entry point
-----------
``solve_grid_layout(spec, area, tokens, measure=None) -> LayoutResult``
    ``spec``  : ``{cols?, rows?, cells:[{block, data, col?, row?, colspan?,
                rowspan?, resize?, min_h?, fixed_h?}]}`` — the same shape
                ``draw_grid`` already accepts, plus optional autolayout hints.
    ``area``  : anything with ``.x .y .w .h`` (the runtime's ``ContentArea``).
    ``tokens``: the runtime ``Tokens`` class (for gutters/floors/type scale).
    ``measure``: optional ``f(kind, data, width_in) -> float`` height override;
                defaults to the built-in textmetrics-based measurer.

Returns ``LayoutResult`` with one ``PlacedCell`` per input cell (x, y, w, h in
inches, plus the measured intrinsic height and an ``overflow_in`` residual) and
a ``diagnostics`` dict suitable for a ``layout_report.json``.
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional, Tuple

try:                                   # reuse the runtime's calibrated metrics
    import textmetrics as _tm
except Exception:                      # pragma: no cover - allows standalone import
    _tm = None


# --------------------------------------------------------------------------- #
# Resize modes (the Figma primitive)
# --------------------------------------------------------------------------- #
HUG = "hug"      # size to measured content   (default for text/kpi/table)
FILL = "fill"    # absorb leftover row space   (default for chart/diagram blocks)
FIXED = "fixed"  # honour an explicit height   (cell["fixed_h"])

# Blocks that want to grow into available space rather than hug their content.
_FILL_BY_DEFAULT = frozenset({
    "chart", "waterfall", "treemap", "matrix", "cycle", "pyramid",
    "org_chart", "process", "timeline", "growing_steps", "image", "key_message",
})

# Blocks that read as "cards" — a matrix of these usually wants one uniform
# height (used by equalize:"auto").
_CARD_KINDS = frozenset({
    "text", "bullets", "kpi", "kpi_callout", "stat_callout", "sticker",
    "compare", "icon_rows", "rule_text",
})

# Minimum sensible heights (inches) per block kind — the autolayout floor. Kept
# in sync with runtime.BLOCK_MIN_SIZE; passed in by the caller when available.
_DEFAULT_MIN_H = {
    "text": 0.5, "chart": 1.8, "waterfall": 2.0, "treemap": 2.0, "cycle": 1.8,
    "pyramid": 1.8, "process": 1.0, "org_chart": 1.8, "matrix": 2.6,
    "timeline": 1.0, "growing_steps": 1.4, "kpi": 1.0, "table": 0.8,
    "image": 1.0, "rule_text": 0.6, "kpi_callout": 0.7, "key_message": 2.0,
    "kpi_card": 1.2, "quote_card": 1.6, "stat_card": 1.0, "icon_card": 0.7,
}


@dataclass
class PlacedCell:
    cell: dict
    col: int
    row: int
    colspan: int
    rowspan: int
    x: float
    y: float
    w: float
    h: float
    intrinsic_h: float            # measured content height (inches)
    resize: str
    overflow_in: float = 0.0      # >0 means content still exceeds the box


@dataclass
class LayoutResult:
    cells: List[PlacedCell] = field(default_factory=list)
    cols: int = 1
    rows: int = 1
    col_w: float = 0.0
    col_gutter: float = 0.0
    row_gutter: float = 0.0
    row_heights: List[float] = field(default_factory=list)
    diagnostics: dict = field(default_factory=dict)


# --------------------------------------------------------------------------- #
# Measurement — intrinsic height of a block in a column of a given width
# --------------------------------------------------------------------------- #
def _paragraphs_for_text(data: dict, tokens) -> List[Tuple[str, int, float, str]]:
    """Map a text/bullets block to (text, level, size_pt, kind) tuples for
    textmetrics, mirroring how draw_text_block / draw_columns emit runs."""
    body = float(getattr(tokens, "TYPE_BODY", 12))
    paras: List[Tuple[str, int, float, str]] = []
    heading = (data.get("heading") or data.get("label") or "").strip()
    if heading:
        paras.append((heading, 0, body, "minor"))           # 12pt bold heading
    sub = (data.get("subheading") or "").strip()
    if sub:
        paras.append((sub, 0, body, "minor"))
    bullets = data.get("bullets")
    paragraphs = data.get("paragraphs")
    text = data.get("text")
    if bullets and _tm is not None:
        # iter_bullets lives in runtime; replicate its (level, text) walk simply.
        for lvl, txt in _walk_bullets(bullets, 1):
            paras.append((str(txt), min(lvl, 3), body, "minor"))
    elif paragraphs:
        for b in paragraphs:
            paras.append((str(b), 0, body, "minor"))
    elif text:
        paras.append((str(text), 0, body, "minor"))
    return paras


def _walk_bullets(items, level: int):
    """Lightweight stand-in for runtime.iter_bullets: yields (level, text)."""
    for it in items or []:
        if isinstance(it, dict):
            txt = it.get("text", it.get("label", ""))
            if txt:
                yield level, txt
            for child_key in ("children", "bullets", "sub"):
                if it.get(child_key):
                    yield from _walk_bullets(it[child_key], level + 1)
        elif isinstance(it, (list, tuple)):
            yield from _walk_bullets(it, level + 1)
        else:
            yield level, str(it)


def _space_after_in(paras, tokens) -> float:
    """Inter-paragraph space-after (inches), approximating the OW para tokens."""
    if not paras:
        return 0.0
    prose = float(getattr(tokens, "PARA_PROSE", 6))
    bullet = float(getattr(tokens, "PARA_BULLET", 4))
    pts = 0.0
    for (_t, lvl, _s, _k) in paras[:-1]:                     # no space after last
        pts += bullet if lvl >= 1 else prose
    return pts / 72.0


def default_measure(kind: str, data: dict, width_in: float, tokens) -> float:
    """Intrinsic content height (inches) for a block at column width ``width_in``.

    Text blocks are measured with textmetrics (line wrapping + line height +
    space-after). Non-text blocks return their minimum sensible height — they are
    typically FILL blocks, so the exact intrinsic height matters less than the
    floor they must not drop below."""
    kind = (kind or "text").strip()
    min_h = _DEFAULT_MIN_H.get(kind, 0.5)
    if kind in ("kpi_card", "quote_card", "stat_card", "icon_card"):
        try:
            import cards as _cards
            return max(min_h, _cards.measure_card(kind, data, width_in))
        except Exception:
            return min_h
    if kind in ("text", "bullets") and _tm is not None and _tm.available():
        paras = _paragraphs_for_text(data, tokens)
        if not paras:
            return min_h
        stack = _tm.text_height_in(paras, width_in)
        pad = 2 * 0.04                                        # top+bottom inset
        return max(min_h, stack + _space_after_in(paras, tokens) + pad)
    return min_h


# --------------------------------------------------------------------------- #
# Grid topology — faithful to runtime.resolve_grid_dims / draw_grid placement
# --------------------------------------------------------------------------- #
def _cs(c) -> int:
    return max(1, int(c.get("colspan", 1) or 1))


def _rs(c) -> int:
    return max(1, int(c.get("rowspan", 1) or 1))


def _solve_dims(cells, cols, rows, cols_max) -> Tuple[int, int]:
    eff = sum(_cs(c) * _rs(c) for c in cells) or 1
    if cols and rows:
        N, M = int(cols), int(rows)
    elif cols:
        N, M = int(cols), max(1, math.ceil(eff / int(cols)))
    elif rows:
        N, M = max(1, math.ceil(eff / int(rows))), int(rows)
    else:
        M = 1
        while math.ceil(eff / M) > cols_max:
            M += 1
        N = math.ceil(eff / M)
    # never clip explicit placements
    N = max(N, max((int(c.get("col", 0)) + _cs(c) for c in cells), default=1))
    M = max(M, max((int(c.get("row", 0)) + _rs(c) for c in cells), default=1))
    return N, M


def _place(cells, N, M):
    """Assign (col,row) to every cell (auto-flow row-major, honour explicit)."""
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

    placed = []
    for cell in cells:
        cs, rs = _cs(cell), _rs(cell)
        if "col" in cell and "row" in cell:
            c, r = int(cell["col"]), int(cell["row"])
        else:
            pos = find(cs, rs)
            if pos is None:
                continue                                     # no room; skip
            c, r = pos
        mark(c, r, cs, rs)
        placed.append((cell, c, r, cs, rs))
    return placed


def _resize_mode(cell) -> str:
    explicit = (cell.get("resize") or "").strip().lower()
    if explicit in (HUG, FILL, FIXED):
        return explicit
    if cell.get("fixed_h"):
        return FIXED
    kind = (cell.get("block") or "text").strip()
    return FILL if kind in _FILL_BY_DEFAULT else HUG


# --------------------------------------------------------------------------- #
# The solver
# --------------------------------------------------------------------------- #
def solve_grid_layout(spec: dict, area, tokens,
                      measure: Optional[Callable] = None) -> LayoutResult:
    cells = [c for c in (spec.get("cells") or []) if isinstance(c, dict)]
    if not cells:
        return LayoutResult()

    cols_max = int(getattr(tokens, "GRID_COLS_MAX", 4))
    N, M = _solve_dims(cells, spec.get("cols"), spec.get("rows"), cols_max)
    placed = _place(cells, N, M)

    LG = float(getattr(tokens, "SPACE_GAP_LG", 0.5))
    SM = float(getattr(tokens, "SPACE_GAP", 0.25))
    min_col_w = float(getattr(tokens, "GRID_MIN_COL_W", 1.8))
    min_row_h = float(getattr(tokens, "GRID_MIN_ROW_H", 1.3))

    # Columns stay equal tracks (OW keeps columns aligned); pick the gutter the
    # same way draw_grid does, then derive column width.
    col_gut = LG if (area.w - (N - 1) * LG) / N >= min_col_w else SM
    col_w = (area.w - (N - 1) * col_gut) / N

    measure = measure or (lambda k, d, w: default_measure(k, d, w, tokens))

    # ---- 1. Measure every cell at its spanned column width -----------------
    intrinsic: Dict[int, float] = {}
    modes: Dict[int, str] = {}
    for i, (cell, c, r, cs, rs) in enumerate(placed):
        w = cs * col_w + (cs - 1) * col_gut
        mode = _resize_mode(cell)
        modes[i] = mode
        if mode == FIXED and cell.get("fixed_h"):
            intrinsic[i] = float(cell["fixed_h"])
        else:
            intrinsic[i] = float(measure(cell.get("block", "text"),
                                         cell.get("data", cell), w))

    # ---- 1b. Equalize peer elements (consistent card heights) -------------
    # Two consistency controls:
    #   * cell["group"] = "<id>"  -> every cell in the group gets the SAME
    #     height (= the tallest member). Use for "these cards belong together".
    #   * spec["equalize"]:
    #       "rows"  -> all grid ROWS share one height (uniform card matrix)
    #       "all"/"cells" -> every cell shares one height (one big group)
    #       "auto"  -> "rows" when the grid is all card-like blocks & M>1
    # Within a single row, cells already share the row track, so a row of N
    # cards is uniform with no setting at all; these handle the cross-row case.
    eq = (spec.get("equalize") or "").strip().lower()
    if eq == "auto":
        all_cards = all((cell.get("block", "text") in _CARD_KINDS)
                        for cell, *_ in placed)
        eq = "rows" if (M > 1 and all_cards) else ""

    groups: Dict[str, List[int]] = {}
    for i, (cell, *_rest) in enumerate(placed):
        gid = cell.get("group")
        if gid is None and eq in ("all", "cells"):
            gid = "__all__"
        if gid is not None:
            groups.setdefault(str(gid), []).append(i)
    for gid, idxs in groups.items():
        gmax = max(intrinsic[i] for i in idxs)
        for i in idxs:
            intrinsic[i] = gmax            # peers now share the tallest height

    # ---- 2. Per-grid-row required height (HUG) -----------------------------
    # A single-row cell pushes its row to at least its measured height. A cell
    # spanning rs rows distributes its requirement across those rows (so one
    # tall spanning cell doesn't inflate every row it crosses).
    row_need = [min_row_h] * M
    has_fill_in_row = [False] * M
    for i, (cell, c, r, cs, rs) in enumerate(placed):
        if modes[i] == FILL:
            for rr in range(r, r + rs):
                has_fill_in_row[rr] = True
        if rs == 1:
            row_need[r] = max(row_need[r], intrinsic[i])
        else:
            share = intrinsic[i] / rs
            for rr in range(r, r + rs):
                row_need[rr] = max(row_need[rr], share)

    # Uniform-rows equalization: every row track shares the tallest row's need,
    # so a 2x3 (or NxM) card matrix comes out as one consistent height.
    uniform_rows = eq in ("rows", "all")
    if uniform_rows and M > 1:
        rmax = max(row_need)
        row_need = [rmax] * M

    vfill = bool(spec.get("vfill", True))     # justify rows to fill the slide?
    # ---- 3. Distribute slack to FILL rows / resolve overflow ---------------
    row_gut = LG if M <= 1 or (area.h - (M - 1) * LG) / M >= min_row_h else SM
    avail = area.h - (M - 1) * row_gut
    base = sum(row_need)
    diag = {"area_h": round(area.h, 3), "rows": M, "cols": N,
            "row_gutter": row_gut, "col_gutter": col_gut,
            "equalize": eq or None, "uniform_rows": uniform_rows and M > 1,
            "groups": {g: len(v) for g, v in groups.items()} or None,
            "row_need_in": [round(x, 3) for x in row_need],
            "ladder": []}

    if base <= avail:
        # Slack exists. With vfill off, rows keep their hug heights and content
        # sits at the top (preferred for card layouts). Otherwise distribute.
        slack = avail - base
        fill_rows = [r for r in range(M) if has_fill_in_row[r]]
        if not vfill:
            diag["ladder"].append("vfill off: rows hug (top-anchored)")
        elif uniform_rows and M > 1 and slack > 0.001:
            add = slack / M
            row_need = [rn + add for rn in row_need]
            diag["ladder"].append(f"equalize: +{add:.2f}in to all rows (uniform)")
        elif slack > 0.001 and fill_rows:
            add = slack / len(fill_rows)
            for r in fill_rows:
                row_need[r] += add
            diag["ladder"].append(f"fill: +{add:.2f}in to {len(fill_rows)} row(s)")
        elif slack > 0.001:
            # No FILL row: distribute slack evenly so the grid fills the slide.
            add = slack / M
            for r in range(M):
                row_need[r] += add
            diag["ladder"].append(f"justify: +{add:.2f}in to all rows")
        row_h = row_need
    else:
        # Overflow ladder ----------------------------------------------------
        # (a) tighten the row gutter
        if row_gut != SM:
            row_gut = SM
            avail = area.h - (M - 1) * row_gut
            diag["ladder"].append("tighten row gutter 0.50->0.25in")
        if base > avail:
            # (b) proportionally compress rows toward (not below) their floor.
            floor_total = min_row_h * M
            compress = (avail - floor_total) / max(0.001, base - floor_total)
            compress = max(0.0, min(1.0, compress))
            row_h = [min_row_h + (rn - min_row_h) * compress for rn in row_need]
            diag["ladder"].append(f"compress rows x{compress:.2f} toward floor")
            # (c) residual overflow (content genuinely too dense -> split slide)
            residual = sum(row_h) - avail
            if residual > 0.01:
                diag["ladder"].append(
                    f"RESIDUAL {residual:.2f}in: recommend slide split")
                diag["recommend_split"] = True
        else:
            row_h = row_need

    # ---- 4. Emit placed rects (top-left origin, row-major stacking) --------
    row_y = [area.y]
    for r in range(1, M):
        row_y.append(row_y[-1] + row_h[r - 1] + row_gut)

    out: List[PlacedCell] = []
    for i, (cell, c, r, cs, rs) in enumerate(placed):
        x = area.x + c * (col_w + col_gut)
        y = row_y[r]
        w = cs * col_w + (cs - 1) * col_gut
        h = sum(row_h[r:r + rs]) + (rs - 1) * row_gut
        overflow = max(0.0, intrinsic[i] - h) if modes[i] != FILL else 0.0
        out.append(PlacedCell(cell=cell, col=c, row=r, colspan=cs, rowspan=rs,
                              x=round(x, 3), y=round(y, 3),
                              w=round(w, 3), h=round(h, 3),
                              intrinsic_h=round(intrinsic[i], 3),
                              resize=modes[i], overflow_in=round(overflow, 3)))

    diag["row_height_in"] = [round(x, 3) for x in row_h]
    diag["total_h_in"] = round(sum(row_h) + (M - 1) * row_gut, 3)
    return LayoutResult(cells=out, cols=N, rows=M, col_w=round(col_w, 3),
                        col_gutter=col_gut, row_gutter=row_gut,
                        row_heights=[round(x, 3) for x in row_h],
                        diagnostics=diag)


# --------------------------------------------------------------------------- #
# Drop-in for TemplateRuntime.draw_grid
# --------------------------------------------------------------------------- #
def draw_grid_autolayout(runtime, slide, spec) -> "LayoutResult | None":
    """Content-measured replacement for ``runtime.draw_grid``. Solves the layout
    with intrinsic sizing, renders each cell through the runtime's own
    ``render_block``, and returns the LayoutResult (for a layout_report.json)."""
    if not isinstance(spec, dict):
        return None
    from runtime import Tokens                                # token source
    area = runtime.content_area()
    result = solve_grid_layout(spec, area, Tokens)
    for pc in result.cells:
        runtime.render_block(slide, pc.cell.get("block", "text"),
                             pc.cell.get("data", pc.cell),
                             x=pc.x, y=pc.y, w=pc.w, h=pc.h)
    return result
