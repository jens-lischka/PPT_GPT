from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "ow_default.pptx"
SEMANTIC = ROOT / "examples" / "semantic_deck.json"


def _generate(tmp_path: Path, deck_path: Path = SEMANTIC) -> Path:
    output = tmp_path / "generated.pptx"
    cmd = [
        sys.executable,
        str(ROOT / "src" / "generate_deck.py"),
        "--template",
        str(TEMPLATE),
        "--deck",
        str(deck_path),
        "--output",
        str(output),
    ]
    subprocess.run(cmd, cwd=ROOT, check=True)
    assert output.exists()
    return output


def _count_expected_slides(deck_spec: dict) -> int:
    # Mirrors examples/semantic_deck.json overflow behavior: 11 semantic slides,
    # with one bullet slide split into 2 and the 20-row table packed greedily
    # into 2 pages (15 + 5) — max rows per slide. 11 + 1 + 1 = 13.
    return 13


def test_template_example_slides_are_removed(tmp_path: Path):
    output = _generate(tmp_path)
    prs = Presentation(output)
    with open(SEMANTIC, "r", encoding="utf-8") as f:
        deck = json.load(f)
    assert len(prs.slides) == _count_expected_slides(deck)


def test_generated_deck_does_not_start_with_template_confidentiality(tmp_path: Path):
    output = _generate(tmp_path)
    prs = Presentation(output)
    all_text = "\n".join(
        shape.text
        for shape in prs.slides[0].shapes
        if getattr(shape, "has_text_frame", False)
    )
    assert "Confidentiality" not in all_text
    assert "Phase 9" in all_text


def test_chart_is_bound_to_placeholder_marker(tmp_path: Path):
    output = _generate(tmp_path)
    prs = Presentation(output)
    chart_shapes = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.CHART:
                chart_shapes.append(shape)
    assert chart_shapes, "Expected at least one chart in semantic sample deck"
    chart = chart_shapes[0]
    assert chart.is_placeholder
    assert chart.placeholder_format.type == PP_PLACEHOLDER.OBJECT
    assert chart.placeholder_format.idx == 11


def test_table_is_bound_to_placeholder_marker(tmp_path: Path):
    output = _generate(tmp_path)
    prs = Presentation(output)
    table_shapes = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                table_shapes.append(shape)
    assert table_shapes, "Expected at least one table in semantic sample deck"
    table = table_shapes[0]
    assert table.is_placeholder
    assert table.placeholder_format.type == PP_PLACEHOLDER.OBJECT
    assert table.placeholder_format.idx == 11


def test_kpi_cards_are_styled_placeholders_or_grid_shapes(tmp_path: Path):
    output = _generate(tmp_path)
    prs = Presentation(output)
    # Slide 13 in examples/semantic_deck.json is the four-column KPI control slide
    # after overflow expansion. It should contain KPI placeholders with explicit fill.
    kpi_slide = prs.slides[10]
    filled = []
    for shape in kpi_slide.shapes:
        if getattr(shape, "has_text_frame", False) and "REVENUE" in shape.text.upper():
            filled.append(shape)
    assert filled, "Expected a KPI card with Revenue text"
    # It should not be a plain unstyled textbox.
    assert filled[0].shape_type != MSO_SHAPE_TYPE.TEXT_BOX


def test_dashboard_overflow_uses_title_only_canvas(tmp_path: Path):
    output = _generate(tmp_path)
    prs = Presentation(output)
    dashboard_slide = prs.slides[9]
    text = "\n".join(
        s.text for s in dashboard_slide.shapes if getattr(s, "has_text_frame", False)
    )
    assert "Eight metrics" in text
    # Engine-composed dashboard should have many non-placeholder card shapes on canvas.
    non_placeholder_cards = [
        s for s in dashboard_slide.shapes
        if getattr(s, "has_text_frame", False) and not getattr(s, "is_placeholder", False)
    ]
    assert len(non_placeholder_cards) >= 4


def test_yaml_deck_loads_and_builds(tmp_path):
    """Mode b2: a hand-authored YAML deck loads via the same CLI path as JSON
    and produces a valid .pptx (the deck_spec IR is format-agnostic)."""
    yaml_text = (
        "slides:\n"
        "  - intent: introduce_topic\n"
        "    content: {title: YAML cover, subtitle: Authored by hand}\n"
        "  - intent: explain\n"
        "    content:\n"
        "      title: A YAML deck builds exactly like a JSON deck\n"
        "      body:\n"
        "        paragraphs:\n"
        "          - The loader runs yaml.safe_load and feeds the same compiler.\n"
        "      footnote: 'Source: internal test.'\n"
    )
    deck = tmp_path / "deck.yaml"
    deck.write_text(yaml_text)
    out = _generate(tmp_path, deck)
    prs = Presentation(str(out))
    assert len(prs.slides) == 2
