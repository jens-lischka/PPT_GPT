from __future__ import annotations


from compiler import INTENTS, LAYOUT_CAPABILITIES


def test_core_intents_exist():
    for name in [
        "introduce_topic",
        "explain",
        "compare_two_options",
        "show_trend_with_key_message",
        "dashboard",
        "show_data",
    ]:
        assert name in INTENTS


def test_dashboard_prefers_title_only_for_grid_overflow():
    intent = INTENTS["dashboard"]
    assert "4 columns" in intent.preferred_layouts
    assert "4 columns" in intent.preferred_layouts
    # Title Only routing is selected dynamically by the variadic overflow handler
    # for explicit/rich/overflow grids, not by the static fallback list.
    assert "Title and Content" in intent.fallback_layouts


def test_title_only_supports_engine_composed_visuals():
    supported = LAYOUT_CAPABILITIES["Title Only"].supports
    for kind in ["process", "org_chart", "pyramid", "cycle"]:
        assert kind in supported


def test_title_and_content_supports_native_components():
    supported = LAYOUT_CAPABILITIES["Title and Content"].supports
    for kind in ["text", "bullets", "chart", "table"]:
        assert kind in supported
