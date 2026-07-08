"""Regression tests for the chart/treemap normalizer and the new gate checks.

Each guards a specific bug that drew slide-specific reviewer feedback:
  - treemap broke PowerPoint (malformed chartEx)           -> test_treemap_chartex_valid
  - charts showed tick marks / redundant value-axis scale  -> test_chart_normalizer_*
  - off-token font sizes                                    -> test_gate_font_size_check
"""
from __future__ import annotations

import re
import sys
import zipfile
import xml.dom.minidom as minidom
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
TEMPLATE = ROOT / "ow_default.pptx"

from compiler import SemanticCompiler                      # noqa: E402
from runtime import TemplateRuntime, normalize_chart_parts  # noqa: E402
import gate as _gate                                       # noqa: E402


def _build(slides, out):
    rt = TemplateRuntime(str(TEMPLATE))
    SemanticCompiler(rt).build({"deck": {"title": "t", "template": "ow_default.pptx"},
                                "slides": slides})
    rt.save(str(out))
    return str(out)


def _chart_xml(path, pattern):
    z = zipfile.ZipFile(path)
    return {n: z.read(n).decode("utf-8", "ignore")
            for n in z.namelist() if re.match(pattern, n)}


def test_treemap_chartex_valid(tmp_path):
    """The chartEx treemap must be schema-valid: <cx:lvl> wrappers + an
    externalData link that resolves to the embedded workbook. This is the exact
    malformation that forced a PowerPoint repair before."""
    out = _build([{"intent": "show_treemap", "content": {
        "title": "Cost mix is broader than licenses",
        "treemap": {"items": [{"label": "Platforms", "value": 22},
                              {"label": "Usage", "value": 15},
                              {"label": "Data prep", "value": 18}]}}}],
        tmp_path / "tm.pptx")
    parts = _chart_xml(out, r"ppt/charts/chartEx\d+\.xml$")
    assert parts, "no chartEx treemap part was written"
    for name, xml in parts.items():
        minidom.parseString(xml)                       # well-formed
        assert "<cx:lvl" in xml, f"{name}: points not wrapped in <cx:lvl>"
        m = re.search(r'externalData r:id="(rId\d+)"', xml)
        assert m, f"{name}: missing externalData workbook link"
        rels = zipfile.ZipFile(out).read(
            f"ppt/charts/_rels/{name.split('/')[-1]}.rels").decode()
        rm = re.search(rf'Id="{m.group(1)}"[^>]*Target="([^"]+)"', rels)
        assert rm and ".xlsx" in rm.group(1), f"{name}: externalData does not resolve"


def test_chart_normalizer_strips_tick_marks(tmp_path):
    out = _build([{"intent": "show_chart", "content": {
        "title": "Bookings are rising across the quarter",
        "chart": {"type": "column_clustered", "categories": ["Q1", "Q2", "Q3"],
                  "series": [{"name": "Bookings", "values": [10, 20, 35]}]}}}],
        tmp_path / "ch.pptx")
    normalize_chart_parts(out)
    for name, xml in _chart_xml(out, r"ppt/charts/chart\d+\.xml$").items():
        marks = set(re.findall(r'(?:major|minor)TickMark val="(\w+)"', xml))
        assert marks <= {"none"}, f"{name}: tick marks present {marks}"


def test_chart_normalizer_hides_value_axis_when_labeled(tmp_path):
    out = _build([{"intent": "show_chart", "content": {
        "title": "Bookings are rising across the quarter",
        "chart": {"type": "column_clustered", "categories": ["Q1", "Q2", "Q3"],
                  "series": [{"name": "Bookings", "values": [10, 20, 35]}]}}}],
        tmp_path / "ch2.pptx")
    normalize_chart_parts(out)
    for name, xml in _chart_xml(out, r"ppt/charts/chart\d+\.xml$").items():
        if "showVal val=\"1\"" in xml:
            vm = re.search(r"<c:valAx>.*?</c:valAx>", xml, re.S)
            assert vm and 'delete val="1"' in vm.group(0), f"{name}: value axis not hidden"


def test_gate_charts_and_treemap_pass(tmp_path):
    out = _build([
        {"intent": "show_chart", "content": {"title": "Bookings rise through the year",
            "chart": {"type": "column_clustered", "categories": ["Q1", "Q2"],
                      "series": [{"name": "s", "values": [1, 2]}]}}},
        {"intent": "show_treemap", "content": {"title": "Cost mix is broad",
            "treemap": {"items": [{"label": "A", "value": 5}, {"label": "B", "value": 3}]}}},
    ], tmp_path / "g.pptx")
    normalize_chart_parts(out)
    assert _gate._check_charts(out)["passed"], "tick-mark gate should pass after normalize"
    assert _gate._check_treemap(out)["passed"], "treemap gate should pass"


def test_gate_font_size_check_flags_off_token(tmp_path):
    """The font-size check is a token guard: a clean deck has no off-token sizes."""
    out = _build([{"intent": "explain", "content": {
        "title": "Synthesis is the product, and it is exposed",
        "body": {"paragraphs": ["Professional services sell judgment at speed."]}}}],
        tmp_path / "f.pptx")
    res = _gate._check_font_sizes(out)
    assert res["severity"] == "warn"
    # explain body uses TYPE_BODY (12pt) -> no sub-18pt off-token sizes expected
    assert not res["fail_items"], f"unexpected off-token sizes: {res['fail_items']}"


def test_voice_linter_flags_banned_terms_and_inflections():
    import gate
    bad = [{"intent": "explain", "content": {
        "title": "Leveraging synergies to transform the business",
        "body": {"paragraphs": ["A best-in-class platform that will optimize ops."]}}}]
    hits = gate._check_voice(bad)["fail_items"]
    assert hits, "should flag banned terms"
    joined = " ".join(hits)
    for stem in ("leverag", "synerg", "transform", "best-in-class", "optimi"):
        assert stem in joined, f"missing {stem}"
    clean = [{"intent": "explain", "content": {
        "title": "Margins improved as freight costs fell",
        "body": {"paragraphs": ["Gross margin rose three points after the reset."]}}}]
    assert gate._check_voice(clean)["fail_items"] == []


def test_design_heuristics_chart_choice_and_table_vs_cards():
    import gate
    pie = [{"intent": "show_data", "content": {"title": "Split", "chart": {
        "type": "pie", "categories": list("ABCDEFGH"),
        "series": [{"name": "s", "values": [1] * 8}]}}}]
    assert any("slices" in x for x in gate._check_design(pie)["fail_items"])
    small = [{"intent": "show_data", "content": {"title": "Two options", "table": {
        "headers": ["Option", "Verdict"],
        "rows": [["Build", "owned but slower"], ["Buy", "faster but locked-in"]]}}}]
    assert any("table" in x for x in gate._check_design(small)["fail_items"])
    # a numeric table of the same size should NOT be flagged
    numeric = [{"intent": "show_data", "content": {"title": "KPIs", "table": {
        "headers": ["Metric", "Value"], "rows": [["Rev", "$4.2m"], ["Margin", "31%"]]}}}]
    assert gate._check_design(numeric)["fail_items"] == []


def test_sources_check_accepts_slide_level_footnote():
    import gate
    # footnote at slide level (canonical) must satisfy the source check
    slide_level = [{"intent": "explain", "footnote": "Source: X.",
                    "content": {"title": "t", "body": {"paragraphs": ["b"]}}}]
    assert gate._check_sources(slide_level)["fail_items"] == []
    # footnote nested in content must also satisfy it
    nested = [{"intent": "explain", "content": {
        "title": "t", "body": {"paragraphs": ["b"]}, "footnote": "Source: X."}}]
    assert gate._check_sources(nested)["fail_items"] == []
    # genuinely missing -> warned
    missing = [{"intent": "explain", "content": {
        "title": "t", "body": {"paragraphs": ["b"]}}}]
    assert gate._check_sources(missing)["fail_items"]


def test_cover_subtitle_exempt_from_overflow(tmp_path):
    """The cover subtitle placeholder is shorter than one measured line, so the
    overflow model used to hard-fail every subtitled cover. Cover subtitle/date
    placeholders are exempt; the cover title and all other slides stay checked."""
    import textmetrics as tm
    out = _build([{"intent": "introduce_topic",
                   "content": {"title": "Q4 commercial review",
                               "subtitle": "Revenue, margin and pipeline"}}],
                 tmp_path / "cover.pptx")
    tm.configure(str(TEMPLATE))
    recs = tm.overflow_report(out)
    assert not [r for r in recs if r["slide"] == 1], \
        f"cover subtitle flagged as overflow: {recs}"


def test_chart_data_labels_have_locale_prefix_default_en_us(tmp_path):
    """Chart data labels get a numFmt with the deck's LCID prefix, so
    thousand separators render per locale. Default en-US -> "[$-409]#,##0"."""
    out = _build([{"intent": "show_trend_with_key_message",
                   "content": {"title": "Revenue by quarter beat plan across the year",
                               "chart": {"type": "column_clustered",
                                         "categories": ["Q1", "Q2", "Q3", "Q4"],
                                         "series": [{"name": "Plan",
                                                     "values": [1200, 1350, 1420, 1580]}]},
                               "insight": {"title": "On track", "text": "Every quarter beat plan."}}}],
                 tmp_path / "d.pptx")
    xmls = _chart_xml(out, r"ppt/charts/chart\d+\.xml")
    assert any("[$-409]#,##0" in x for x in xmls.values()), \
        "Expected [$-409]#,##0 in chart data-label numFmt for default en-US deck"


def test_chart_data_labels_use_de_de_lcid_when_set(tmp_path):
    """Setting deck_spec['language'] = 'de-DE' switches the chart LCID to 407,
    so PowerPoint renders 1.000 instead of 1,000."""
    rt = TemplateRuntime(str(TEMPLATE))
    SemanticCompiler(rt).build({
        "deck": {"title": "t", "template": "ow_default.pptx"},
        "language": "de-DE",
        "slides": [{"intent": "show_trend_with_key_message",
                    "content": {"title": "Umsatz pro Quartal ueber Plan im ganzen Jahr",
                                "chart": {"type": "column_clustered",
                                          "categories": ["Q1", "Q2", "Q3", "Q4"],
                                          "series": [{"name": "Plan",
                                                      "values": [1200, 1350, 1420, 1580]}]},
                                "insight": {"title": "Volltreffer", "text": "Jedes Quartal ueber Plan."}}}],
    })
    out = tmp_path / "d.pptx"
    rt.save(str(out))
    xmls = _chart_xml(str(out), r"ppt/charts/chart\d+\.xml")
    assert any("[$-407]#,##0" in x for x in xmls.values()), \
        "Expected [$-407]#,##0 in chart data-label numFmt for de-DE deck"


def test_kpi_dashboard_renders_icons_when_present(tmp_path):
    """When every KPI item in a dashboard carries an ``icon`` that resolves,
    each KPI cell renders the icon as a native freeform (shape_type FREEFORM)."""
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    out = _build([{"intent": "dashboard",
                   "content": {"title": "Eight KPIs to watch across the transformation",
                               "metrics": [
                                   {"label": "Cost-to-income", "value": "61%",
                                    "delta": "vs 53%", "icon": "percentage"},
                                   {"label": "Fee growth",     "value": "-8%",
                                    "delta": "vs +2%", "icon": "trending-down"},
                                   {"label": "NPS",            "value": "18",
                                    "delta": "vs 34", "icon": "chat"},
                                   {"label": "Digital sales",  "value": "42%",
                                    "delta": "vs 71%", "icon": "chart"},
                                   {"label": "CAC",            "value": "185",
                                    "delta": "vs 95", "icon": "money-time"},
                                   {"label": "Deposit margin", "value": "1.4%",
                                    "delta": "vs 1.9%", "icon": "coins"},
                                   {"label": "Branch productivity","value": "1.1M",
                                    "delta": "vs 1.8M", "icon": "briefcase"},
                                   {"label": "Engagement",     "value": "62",
                                    "delta": "vs 74", "icon": "users"},
                               ]}}], tmp_path / "d.pptx")
    p = Presentation(out)
    freeforms = [sh for sh in p.slides[0].shapes if sh.shape_type == MSO_SHAPE_TYPE.FREEFORM]
    assert len(freeforms) == 8, f"expected 8 icon freeforms, got {len(freeforms)}"


def test_value_axis_numfmt_locale_prefixed(tmp_path):
    """Value-axis tick labels must carry the same LCID prefix as data labels.
    Without this, tick numbers render in the client's locale, not the deck's —
    so a de-DE deck opened on an en-US machine would show axis "12,000" and
    data label "12.000" on the same bar."""
    rt = TemplateRuntime(str(TEMPLATE))
    SemanticCompiler(rt).build({
        "deck": {"title": "t", "template": "ow_default.pptx"},
        "language": "de-DE",
        "slides": [{"intent": "show_trend_with_key_message",
                    "content": {"title": "Umsatz pro Quartal ueber Plan im ganzen Jahr",
                                "chart": {"type": "column_clustered",
                                          "categories": ["Q1","Q2","Q3","Q4"],
                                          "series": [{"name":"Rev","values":[12000,15000,18000,22000]}]},
                                "insight": {"title":"x","text":"y"}}}],
    })
    out = tmp_path / "d.pptx"
    rt.save(str(out))
    xmls = _chart_xml(str(out), r"ppt/charts/chart\d+\.xml")
    xml = list(xmls.values())[0]
    valAx = re.search(r"<c:valAx>.*?</c:valAx>", xml, re.S)
    assert valAx and '[$-407]#,##0' in valAx.group(0), \
        "Expected value-axis numFmt with [$-407]#,##0 (de-DE) — none found"


def test_scatter_chart_does_not_crash_under_localization(tmp_path):
    """python-pptx raises AttributeError on chart.plots[0].has_data_labels for
    XY/scatter charts. The runtime must catch this cleanly — deck build must
    succeed and the chart part must exist — even when the deck sets a language
    (which triggers the locale-aware data-label path)."""
    out = _build([{"intent": "show_trend_with_key_message",
                   "content": {"title": "Deal size scales with tenure across the mid-market",
                               "chart": {"type": "scatter",
                                         "series": [{"name": "Deals",
                                                     "points": [[1,1000],[2,2500],[3,4200],[4,3800]]}]},
                               "insight": {"title":"Trend","text":"Grows."}}}],
                 tmp_path / "d.pptx")
    xmls = _chart_xml(out, r"ppt/charts/chart\d+\.xml")
    assert xmls, "expected a chart part to be written for scatter"
    # Scatter charts must have XY (xVal/yVal) not category (cat/val) structure
    assert "<c:xVal>" in list(xmls.values())[0]


def test_language_auto_detect_english(tmp_path):
    """When no explicit language is set, the compiler auto-detects English
    from deck content and applies [$-409] to chart labels."""
    out = _build([{"intent": "show_trend_with_key_message",
                   "content": {"title": "Revenue grew every quarter across the entire year",
                               "body": {"paragraphs": [
                                   "The plan for this year was to grow revenue in every quarter and we did that in all four quarters with margin to spare."]},
                               "chart": {"type": "column_clustered",
                                         "categories": ["Q1","Q2","Q3","Q4"],
                                         "series": [{"name":"Revenue","values":[1200,1350,1420,1580]}]},
                               "insight": {"title":"x","text":"y"}}}],
                 tmp_path / "d.pptx")
    xml = list(_chart_xml(out, r"ppt/charts/chart\d+\.xml").values())[0]
    assert '[$-409]' in xml, "Auto-detect should have picked en-US -> LCID 409"
    assert '[$-407]' not in xml


def test_language_auto_detect_german(tmp_path):
    """When no explicit language is set and content is German, the compiler
    auto-detects de-DE and applies [$-407] to chart labels."""
    out = _build([{"intent": "show_trend_with_key_message",
                   "content": {"title": "Der Umsatz ist in jedem Quartal ueber dem Plan gewachsen",
                               "body": {"paragraphs": [
                                   "Der Plan für dieses Jahr war es, den Umsatz in jedem Quartal zu steigern, und das ist uns in allen vier Quartalen mit Spielraum gelungen."]},
                               "chart": {"type": "column_clustered",
                                         "categories": ["Q1","Q2","Q3","Q4"],
                                         "series": [{"name":"Umsatz","values":[1200,1350,1420,1580]}]},
                               "insight": {"title":"x","text":"y"}}}],
                 tmp_path / "d.pptx")
    xml = list(_chart_xml(out, r"ppt/charts/chart\d+\.xml").values())[0]
    assert '[$-407]' in xml, "Auto-detect should have picked de-DE -> LCID 407"
    assert '[$-409]' not in xml


def test_explicit_language_overrides_content(tmp_path):
    """Explicit language always wins over content-based detection.
    English-content deck with language='de-DE' must produce [$-407]."""
    rt = TemplateRuntime(str(TEMPLATE))
    SemanticCompiler(rt).build({
        "deck": {"title": "t", "template": "ow_default.pptx"},
        "language": "de-DE",
        "slides": [{"intent": "show_trend_with_key_message",
                    "content": {"title": "Revenue grew every quarter across the entire year",
                                "chart": {"type": "column_clustered",
                                          "categories": ["Q1","Q2","Q3","Q4"],
                                          "series": [{"name":"Revenue","values":[1200,1350,1420,1580]}]},
                                "insight": {"title":"x","text":"y"}}}],
    })
    out = tmp_path / "d.pptx"
    rt.save(str(out))
    xml = list(_chart_xml(str(out), r"ppt/charts/chart\d+\.xml").values())[0]
    assert '[$-407]' in xml, "Explicit de-DE must win over English content"


def test_pyramid_labeled_head_renders_bold_and_body_regular(tmp_path):
    """Diagram labels of the form 'Head: body' split into a bold head run and a
    regular-weight body run. This is the OW-house convention for pyramid,
    cycle, and other labeled-shape diagrams."""
    out = _build([{"intent": "show_pyramid",
                   "content": {"title": "The company culture is built on three layers of increasing specificity",
                               "levels": [
                                   {"label": "Vision: mobility for all across every market we serve"},
                                   {"label": "Way: continuous improvement and respect for people"},
                                   {"label": "System: kanban, jidoka, kaizen, andon, genchi genbutsu"},
                               ]}}], tmp_path / "d.pptx")
    from pptx import Presentation
    p = Presentation(out)
    # Find each pyramid band's textframe and check its two-run structure
    heads_found = 0
    for sh in p.slides[0].shapes:
        if sh.has_text_frame and ":" in sh.text_frame.text:
            para = sh.text_frame.paragraphs[0]
            runs = para.runs
            if len(runs) == 2 and runs[0].font.bold is True and runs[1].font.bold is False:
                heads_found += 1
    assert heads_found == 3, f"expected 3 bold-head/regular-body labels, got {heads_found}"


def test_cycle_labeled_head_renders_bold_and_body_regular(tmp_path):
    """Same head-bold/body-regular pattern applies to cycle ovals."""
    out = _build([{"intent": "show_cycle",
                   "content": {"title": "The improvement cycle runs on four repeating phases every quarter",
                               "items": [
                                   {"label": "Plan: define the standard and the target"},
                                   {"label": "Do: run the improvement at small scale"},
                                   {"label": "Check: measure the actual against the target"},
                                   {"label": "Act: standardise the winner, restart"},
                               ]}}], tmp_path / "d.pptx")
    from pptx import Presentation
    p = Presentation(out)
    heads_found = 0
    for sh in p.slides[0].shapes:
        if sh.has_text_frame and ":" in sh.text_frame.text:
            para = sh.text_frame.paragraphs[0]
            runs = para.runs
            if len(runs) == 2 and runs[0].font.bold is True and runs[1].font.bold is False:
                heads_found += 1
    assert heads_found == 4, f"expected 4 bold-head/regular-body ovals, got {heads_found}"


def test_content_keys_check_flags_silent_drops(tmp_path):
    """When an author passes a content key the strategy doesn't read (typo
    or wrong-shape spec), the runtime silently drops it — no visible slide
    text, but overflow/fonts/sources all pass. This test locks in that the
    new content_keys gate check catches it and surfaces a warn_item so the
    author sees the drop in gate_result.json, not just in scrollback."""
    import json
    from compiler import SemanticCompiler
    from gate import run_gate
    rt = TemplateRuntime(str(TEMPLATE))
    deck = {
        "slides": [
            {"intent": "show_timeline",
             "content": {"title": "A five-milestone timeline of the change programme",
                         # Intent reads 'milestones'; author passed 'items' by mistake.
                         "items": [{"date": "2021", "label": "start"},
                                   {"date": "2025", "label": "target"}]}},
            {"intent": "show_contents",
             "content": {"title": "Agenda covers the five sections of this document",
                         # Reads 'sections'; author passed 'chapters'.
                         "chapters": [{"title": "One"}, {"title": "Two"}]}},
            {"intent": "show_matrix",
             "content": {"title": "The two-by-two shows scale versus margin dynamics",
                         "matrix": {"x_axis": "Scale", "y_axis": "Margin",
                                    "items": [{"label": "A", "row": 0, "col": 1}]},
                         # Extra unknown key — likely author added metadata.
                         "audience_note": "for the exec committee"}},
        ]
    }
    SemanticCompiler(rt).build(deck)
    out = tmp_path / "d.pptx"; rt.save(str(out))
    result = run_gate(deck, str(out), str(TEMPLATE), str(tmp_path / "g.json"))
    ck = result["checks"]["content_keys"]

    # The check must be present and warn-severity — never a hard fail (would
    # break the build for benign extra fields).
    assert ck["severity"] == "warn"
    assert ck["passed"] is True

    # All three unknown keys must be surfaced with the slide index + intent.
    joined = " | ".join(ck["fail_items"])
    assert "slide 1" in joined and "show_timeline" in joined and "'items'" in joined
    assert "slide 2" in joined and "show_contents" in joined and "'chapters'" in joined
    assert "slide 3" in joined and "show_matrix" in joined and "'audience_note'" in joined

    # And the check's warns must appear in the top-level warn_items so a human
    # (or scriptable check) reading gate_result.json sees them without opening
    # the nested per-check dict.
    top_warns = " | ".join(result["warn_items"])
    assert "[content_keys]" in top_warns
    assert top_warns.count("[content_keys]") == 3


def test_content_keys_check_allows_universal_adornments(tmp_path):
    """Common adornments (footnote, source, conclusion, insight, why,
    design_choices, photo) are consumed by the compiler regardless of intent
    and must never trigger the content_keys warn."""
    from compiler import SemanticCompiler
    from gate import run_gate
    rt = TemplateRuntime(str(TEMPLATE))
    deck = {
        "slides": [{
            "intent": "show_columns",
            "content": {
                "title": "Three moves the digital-only banks cannot replicate at scale",
                "columns": [
                    {"heading": "First",  "bullets": ["one", "two"]},
                    {"heading": "Second", "bullets": ["three", "four"]},
                    {"heading": "Third",  "bullets": ["five", "six"]},
                ],
                "footnote": "Source: analysis, 2026",
                "conclusion": "Bundle three moves; the sum beats the parts.",
            }
        }]
    }
    SemanticCompiler(rt).build(deck)
    out = tmp_path / "d.pptx"; rt.save(str(out))
    result = run_gate(deck, str(out), str(TEMPLATE), str(tmp_path / "g.json"))
    ck = result["checks"]["content_keys"]
    assert ck["fail_items"] == [], \
        f"universal adornments must not warn; got: {ck['fail_items']}"
