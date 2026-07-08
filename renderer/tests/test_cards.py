"""Tests for the integrated card components (cards.py) and their registration."""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

import cards
import runtime
import autolayout


def test_registry_matches_runtime():
    assert set(cards.CARD_KINDS) == set(runtime._CARD_BLOCK_KINDS)


def test_build_card_types():
    assert isinstance(cards.build_card("kpi_card", {"label": "X"}), cards.KpiCard)
    assert isinstance(cards.build_card("quote_card", {"quote": "q"}), cards.QuoteCard)
    assert isinstance(cards.build_card("stat_card", {"value": "9%"}), cards.StatCallout)
    assert isinstance(cards.build_card("icon_card", {"heading": "H"}), cards.IconTextCard)


def test_unknown_kind_raises():
    try:
        cards.build_card("nope", {})
        assert False, "expected ValueError"
    except ValueError:
        pass


def test_measure_positive_and_deterministic():
    specs = {
        "kpi_card": {"label": "Cycle time", "value": "11 days",
                     "delta": "38%", "delta_dir": "down", "delta_good": True,
                     "viz": {"kind": "sparkline", "points": [3, 2, 1]},
                     "caption": "last 8 weeks"},
        "quote_card": {"quote": "a fairly long quote that should wrap onto "
                       "several lines in a half-width column", "name": "A B"},
        "stat_card": {"value": "59%", "heading": "Adoption", "body": "lorem ipsum"},
        "icon_card": {"heading": "Single owner", "body": "one lead", "icon": "users"},
    }
    for kind, data in specs.items():
        h1 = cards.measure_card(kind, data, 3.8)
        h2 = cards.measure_card(kind, data, 3.8)
        assert h1 > 0 and h1 == h2


def test_optional_slots_change_height():
    # adding the optional intro slot makes a KPI card taller
    base = cards.measure_card("kpi_card", {"label": "L", "value": "9"}, 3.8)
    taller = cards.measure_card(
        "kpi_card", {"label": "L", "value": "9", "intro": "a descriptive line"}, 3.8)
    assert taller > base


def test_value_is_optional():
    # a paragraph-only KPI card (no big number) still measures
    h = cards.measure_card("kpi_card", {"label": "What changed",
                                        "body": "we consolidated four flows"}, 3.8)
    assert h > 0


def test_solver_measures_cards():
    # the autolayout solver routes card kinds through cards.measure_card
    h = autolayout.default_measure("stat_card",
                                   {"value": "59%", "heading": "Adoption"}, 3.8,
                                   runtime.Tokens)
    assert h > 0


def test_accent_off_by_default():
    # the top bar is opt-in: a KPI card without accent is shorter than with one
    no_bar = cards.measure_card("kpi_card", {"label": "L", "value": "9"}, 3.8)
    with_bar = cards.measure_card("kpi_card",
                                  {"label": "L", "value": "9", "accent": "amber"}, 3.8)
    assert with_bar > no_bar


if __name__ == "__main__":
    import traceback
    fns = [v for k, v in sorted(globals().items()) if k.startswith("test_")]
    p = 0
    for fn in fns:
        try:
            fn(); print("PASS", fn.__name__); p += 1
        except Exception:
            print("FAIL", fn.__name__); traceback.print_exc()
    print(f"{p}/{len(fns)} passed")
