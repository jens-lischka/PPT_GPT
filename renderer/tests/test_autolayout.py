"""Tests for the content-measured autolayout solver.

These pin the four behaviours that distinguish it from equal tracks:
  - HUG rows size to measured content (a dense cell gets a taller row)
  - FILL blocks absorb leftover vertical space
  - FIXED honours an explicit height
  - the overflow ladder compresses toward the floor and flags residual splits
Geometry stays a pure function of content + tokens, so it is golden-testable.
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

import autolayout as al


class _Tok:                         # minimal token stand-in (matches runtime)
    SPACE_GAP = 0.25
    SPACE_GAP_LG = 0.5
    GRID_MIN_COL_W = 1.8
    GRID_MIN_ROW_H = 1.3
    GRID_COLS_MAX = 4
    TYPE_BODY = 12
    TYPE_AUTOFIT_MIN = 10
    PARA_PROSE = 6
    PARA_BULLET = 4


class _Area:
    x, y, w, h = 0.5, 1.54, 12.333, 5.06


def _spec(cells, cols=None, rows=None):
    s = {"cells": cells}
    if cols:
        s["cols"] = cols
    if rows:
        s["rows"] = rows
    return s


def test_hug_rows_size_to_content():
    # A tall cell (measured 3.0") and a short cell (measured 1.0") in a 1x2 grid:
    # the rows must differ, unlike equal tracks which would give 2 x ~2.28".
    def measure(kind, data, w):
        return data["_need"]
    cells = [
        {"col": 0, "row": 0, "block": "text", "data": {"_need": 3.0}, "resize": "hug"},
        {"col": 0, "row": 1, "block": "text", "data": {"_need": 1.0}, "resize": "hug"},
    ]
    res = al.solve_grid_layout(_spec(cells, cols=1, rows=2), _Area, _Tok, measure)
    assert res.rows == 2
    top, bottom = res.row_heights
    assert top > bottom                       # content drove the rows
    assert abs(top - 3.0) < 0.6               # tall row ~hugs its 3.0" content
    assert bottom >= _Tok.GRID_MIN_ROW_H      # never below the readable floor


def test_fill_absorbs_slack():
    # One hug text row + one fill chart row; the chart must take the leftover.
    def measure(kind, data, w):
        return {"text": 1.2, "chart": 1.8}[kind]
    cells = [
        {"col": 0, "row": 0, "block": "text", "data": {}},      # hug by default
        {"col": 0, "row": 1, "block": "chart", "data": {}},     # fill by default
    ]
    res = al.solve_grid_layout(_spec(cells, cols=1, rows=2), _Area, _Tok, measure)
    chart = next(c for c in res.cells if c.cell["block"] == "chart")
    text = next(c for c in res.cells if c.cell["block"] == "text")
    assert chart.resize == al.FILL
    assert chart.h > chart.intrinsic_h        # grew beyond its minimum
    assert chart.h > text.h                   # absorbed the slack


def test_fixed_height_honoured():
    def measure(kind, data, w):
        return 1.0
    cells = [
        {"col": 0, "row": 0, "block": "text", "data": {}, "resize": "fixed",
         "fixed_h": 2.5},
        {"col": 0, "row": 1, "block": "text", "data": {}},
    ]
    res = al.solve_grid_layout(_spec(cells, cols=1, rows=2), _Area, _Tok, measure)
    fixed = res.cells[0]
    assert fixed.resize == al.FIXED
    assert abs(fixed.intrinsic_h - 2.5) < 0.01


def test_overflow_ladder_compresses():
    # Two cells each needing 4.0" (8.0" total) overflow 5.06": the ladder must
    # tighten the gutter and compress toward the floor — but they DO fit once
    # compressed (2 x floor < area), so no split is recommended.
    def measure(kind, data, w):
        return 4.0
    cells = [
        {"col": 0, "row": 0, "block": "text", "data": {}, "resize": "hug"},
        {"col": 0, "row": 1, "block": "text", "data": {}, "resize": "hug"},
    ]
    res = al.solve_grid_layout(_spec(cells, cols=1, rows=2), _Area, _Tok, measure)
    assert all(h >= _Tok.GRID_MIN_ROW_H - 1e-6 for h in res.row_heights)
    assert any("compress" in step for step in res.diagnostics["ladder"])
    assert res.diagnostics.get("recommend_split") in (None, False)


def test_overflow_ladder_flags_split_when_floor_exceeds_area():
    # Five rows that cannot fit even at the readable floor (5 x 1.3 = 6.5" > area):
    # the solver clamps to the floor and recommends a slide split.
    def measure(kind, data, w):
        return 3.0
    cells = [{"col": 0, "row": i, "block": "text", "data": {}, "resize": "hug"}
             for i in range(5)]
    res = al.solve_grid_layout(_spec(cells, cols=1, rows=5), _Area, _Tok, measure)
    assert res.diagnostics.get("recommend_split") is True
    assert all(h >= _Tok.GRID_MIN_ROW_H - 1e-6 for h in res.row_heights)


def test_pure_function_of_inputs():
    # Same inputs -> identical geometry (reproducible builds / golden tests).
    def measure(kind, data, w):
        return 2.0
    cells = [{"block": "text", "data": {}} for _ in range(4)]
    a = al.solve_grid_layout(_spec(list(cells)), _Area, _Tok, measure)
    b = al.solve_grid_layout(_spec(list(cells)), _Area, _Tok, measure)
    assert a.row_heights == b.row_heights
    assert [(c.x, c.y, c.w, c.h) for c in a.cells] == \
           [(c.x, c.y, c.w, c.h) for c in b.cells]


def test_row_of_cards_is_equal_height_by_default():
    # Three cards in ONE row with different content lengths must still come out
    # the same height — they share the row track. No setting required.
    def measure(kind, data, w):
        return data["_need"]
    cells = [
        {"col": 0, "row": 0, "block": "text", "data": {"_need": 1.0}},
        {"col": 1, "row": 0, "block": "text", "data": {"_need": 2.4}},
        {"col": 2, "row": 0, "block": "text", "data": {"_need": 1.3}},
    ]
    res = al.solve_grid_layout(_spec(cells, cols=3, rows=1), _Area, _Tok, measure)
    heights = {round(c.h, 2) for c in res.cells}
    assert len(heights) == 1                  # all three identical


def test_equalize_rows_makes_card_matrix_uniform():
    # 2 rows x 3 cards; one card in row 0 is tall. Without equalize the rows
    # differ; with equalize:"rows" all six cards share one height.
    def measure(kind, data, w):
        return data["_need"]
    cells = []
    needs = [[1.0, 2.6, 1.1], [1.0, 1.0, 1.0]]
    for r in range(2):
        for c in range(3):
            cells.append({"col": c, "row": r, "block": "text",
                          "data": {"_need": needs[r][c]}})

    plain = al.solve_grid_layout(_spec(list(cells), cols=3, rows=2),
                                 _Area, _Tok, measure)
    assert plain.row_heights[0] > plain.row_heights[1]      # rows differ

    spec = _spec(list(cells), cols=3, rows=2)
    spec["equalize"] = "rows"
    uni = al.solve_grid_layout(spec, _Area, _Tok, measure)
    assert len({round(h, 2) for h in uni.row_heights}) == 1  # rows equal
    assert len({round(c.h, 2) for c in uni.cells}) == 1      # every card equal


def test_peer_group_equalizes_across_rows():
    # Two cards tagged into the same group, in different rows, share a height
    # even though other cells are shorter.
    def measure(kind, data, w):
        return data["_need"]
    cells = [
        {"col": 0, "row": 0, "block": "text", "data": {"_need": 2.4}, "group": "kpi"},
        {"col": 1, "row": 0, "block": "text", "data": {"_need": 1.0}},
        {"col": 0, "row": 1, "block": "text", "data": {"_need": 1.0}, "group": "kpi"},
        {"col": 1, "row": 1, "block": "text", "data": {"_need": 1.0}},
    ]
    res = al.solve_grid_layout(_spec(cells, cols=2, rows=2), _Area, _Tok, measure)
    g = [c for c in res.cells if c.cell.get("group") == "kpi"]
    assert abs(g[0].h - g[1].h) < 0.01        # grouped peers share a height


if __name__ == "__main__":
    import traceback
    fns = [v for k, v in sorted(globals().items()) if k.startswith("test_")]
    passed = 0
    for fn in fns:
        try:
            fn()
            print(f"PASS {fn.__name__}")
            passed += 1
        except Exception:
            print(f"FAIL {fn.__name__}")
            traceback.print_exc()
    print(f"\n{passed}/{len(fns)} passed")
