import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
from pptx import Presentation
import textmetrics as tm
import icons, images

TEMPLATE = os.path.join(os.path.dirname(__file__), "..", "ow_default.pptx")

# ---------- A: overflow measurement ----------
def test_metrics_available_and_overflow():
    tm.configure(TEMPLATE)
    assert tm.available()
    long_ = "word " * 120
    small = tm.fits((0,0,3.0,0.6), [(long_,0,12.0,"minor")])
    big   = tm.fits((0,0,3.0,6.0), [(long_,0,12.0,"minor")])
    assert small["fits"] is False and small["overflow_in"] > 0   # BoundHeight-style catch
    assert big["fits"] is True

def test_wrapped_lines_monotonic():
    a = tm.wrapped_lines("short", 3.0)
    b = tm.wrapped_lines("word " * 60, 3.0)
    assert b > a >= 1

def test_overflow_report_flags_only_real_overflow(tmp_path):
    tm.configure(TEMPLATE)
    prs = Presentation(TEMPLATE); L = prs.slide_layouts[0]; s = prs.slides.add_slide(L)
    long_ = "This sentence is deliberately long enough to wrap several times. " * 4
    from pptx.util import Inches
    from pptx.enum.text import MSO_AUTO_SIZE
    tight = s.shapes.add_textbox(Inches(0.5), Inches(1.0), Inches(3.0), Inches(0.5))
    tight.text_frame.word_wrap = True; tight.text_frame.auto_size = MSO_AUTO_SIZE.NONE
    tight.text_frame.text = long_; tight.name = "tight_box"
    roomy = s.shapes.add_textbox(Inches(0.5), Inches(3.0), Inches(3.0), Inches(5.0))
    roomy.text_frame.word_wrap = True; roomy.text_frame.auto_size = MSO_AUTO_SIZE.NONE
    roomy.text_frame.text = long_; roomy.name = "roomy_box"
    out = str(tmp_path/"o.pptx"); prs.save(out)
    recs = tm.overflow_report(out)
    flagged = {r["shape"] for r in recs}
    assert "tight_box" in flagged and "roomy_box" not in flagged

# ---------- B: icons ----------
def test_icon_resolve_and_geom():
    p = icons.resolve("trending")          # keyword -> arrow-trending-up
    assert p and p.endswith(".svg")
    geom = icons.svg_to_custgeom(open(icons.resolve("chart-bar"),"rb").read(), 914400, 914400)
    assert "<a:custGeom>" in geom and "<a:path " in geom and "moveTo" in geom

def test_draw_icon_adds_shape():
    prs = Presentation(TEMPLATE); L = prs.slide_layouts[0]
    s = prs.slides.add_slide(L); before = len(s.shapes._spTree)
    icons.draw_icon(s, "shield-check", 1,1,1,1, "000F47")
    assert len(s.shapes._spTree) == before + 1

# ---------- C: images ----------
def test_aspect_and_fill_crop():
    assert abs(images.parse_aspect("16:9") - 16/9) < 1e-6
    l,t,r,b = images._fill_crop(2.0, 1.0)   # wide image into square box -> crop sides
    assert l > 0 and r > 0 and t == 0 and b == 0

def test_one_source_many_crops_dedupes(tmp_path):
    img = str(tmp_path/"src.png"); images.generate_image("x","16:9",img)
    prs = Presentation(TEMPLATE); L = prs.slide_layouts[0]; s = prs.slides.add_slide(L)
    reg = images.ImageRegistry(); reg.register("h", img)
    images.place_image(s, reg, "h", 0,0,4,2, aspect="16:9", name="a")
    images.place_image(s, reg, "h", 4,0,2,2, aspect="1:1",  name="b")
    out = str(tmp_path/"o.pptx"); prs.save(out)
    import zipfile
    pngs = [n for n in zipfile.ZipFile(out).namelist() if n.startswith("ppt/media/") and n.endswith(".png")]
    assert len(pngs) == 1                  # one blob reused for both crops

def test_placeholder_when_no_asset():
    prs = Presentation(TEMPLATE); L = prs.slide_layouts[0]; s = prs.slides.add_slide(L)
    reg = images.ImageRegistry(); reg.register("missing", None)
    shp = images.place_image(s, reg, "missing", 0,0,4,2, aspect="16:9", name="todo")
    assert shp.name.startswith("REPLACE:")


# ---------- wiring: icons/images consumed by the compiler ----------
def _build(deck):
    import runtime, compiler, images
    rt = runtime.TemplateRuntime(TEMPLATE)
    rt._image_registry = images.ImageRegistry()
    for sid, spec in (deck.get("images") or {}).items():
        rt._image_registry.register(sid, spec.get("path") if isinstance(spec, dict) else spec)
    compiler.SemanticCompiler(rt).build(deck)
    return rt

def _names(slide):
    return [s.name for s in slide.shapes]

def test_column_icon_is_rendered():
    rt = _build({"slides": [{"intent": "show_columns", "content": {
        "title": "T", "columns": [
            {"heading": "A", "icon": "chart-bar", "bullets": ["x"]},
            {"heading": "B", "icon": "rocket-launch", "bullets": ["y"]}]}}]})
    names = sum((_names(s) for s in rt.prs.slides), [])
    assert any(n.startswith("icon-") for n in names)

def test_person_without_photo_still_gets_an_image():
    rt = _build({"slides": [{"intent": "introduce_person",
                             "content": {"name": "Jane Roe", "bio": ["x"]}}]})
    # a picture shape (avatar) lands on the slide
    pics = [s for sl in rt.prs.slides for s in sl.shapes
            if s._element.tag.endswith('}pic')]   # a <p:pic> landed (avatar)
    assert pics

def test_compose_image_with_source_renders_picture(tmp_path):
    import images
    p = str(tmp_path/"s.png"); images.generate_image("x", "1:1", p)
    rt = _build({"images": {"hero": {"path": p}},
                 "slides": [{"intent": "compose", "content": {"title": "T", "compose": {"regions": [
                     {"block": "image", "data": {"source": "hero", "aspect": "1:1", "name": "h"}},
                     {"block": "text", "data": {"bullets": ["x"]}}]}}}]})
    pics = [s for sl in rt.prs.slides for s in sl.shapes
            if s._element.tag.endswith('}pic')]
    assert pics


# ---------- machine-readable build gate ----------
def _gate_build(tmp_path, deck):
    import json, sys, subprocess, os
    src = os.path.join(os.path.dirname(__file__), "..", "src")
    dj = str(tmp_path/"d.json"); open(dj,"w").write(json.dumps(deck))
    out = str(tmp_path/"d.pptx")
    subprocess.run([sys.executable, f"{src}/generate_deck.py", "--template", TEMPLATE,
                    "--deck", dj, "--output", out], capture_output=True, text=True)
    return json.load(open(tmp_path/"gate_result.json"))

def test_gate_passes_good_deck(tmp_path):
    res = _gate_build(tmp_path, {"slides":[
        {"intent":"show_columns","content":{"title":"Digital now drives most of new demand","columns":[
            {"heading":"A","icon":"chart-bar","bullets":["x"]},
            {"heading":"B","icon":"globe-alt","bullets":["y"]}]}}]})
    assert res["passed"] is True

def test_gate_icon_coverage_all_or_nothing(tmp_path):
    # Render-level: the build enforces all-or-nothing — a spec with icons on
    # only some columns renders with NONE (no lopsided mix).
    import os
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    def n_icons(deck):
        import json, sys, subprocess
        src = os.path.join(os.path.dirname(__file__), "..", "src")
        dj = str(tmp_path / "d.json"); open(dj, "w").write(json.dumps(deck))
        out = str(tmp_path / "d.pptx")
        subprocess.run([sys.executable, f"{src}/generate_deck.py", "--template", TEMPLATE,
                        "--deck", dj, "--output", out], capture_output=True, text=True)
        return sum(1 for s in Presentation(out).slides for sh in s.shapes
                   if sh.shape_type == MSO_SHAPE_TYPE.FREEFORM)
    cols = lambda items: {"slides": [{"intent": "show_columns",
        "content": {"title": "Three forces reshape the market now", "columns": items}}]}
    assert n_icons(cols([{"heading": "A", "icon": "growth", "bullets": ["x"]},
                         {"heading": "B", "icon": "shield", "bullets": ["y"]}])) == 2
    assert n_icons(cols([{"heading": "A", "icon": "growth", "bullets": ["x"]},
                         {"heading": "B", "bullets": ["y"]}])) == 0   # lopsided -> none

    # Gate-logic level: lopsided authoring warns (soft); all-or-none passes.
    import sys
    sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
    import gate
    sl = lambda items: [{"intent": "show_columns", "content": {"title": "T", "columns": items}}]
    lop = gate._check_icon_coverage(sl([{"heading": "A", "icon": "growth"}, {"heading": "B"}]))
    assert lop["severity"] == "warn" and not lop["passed"]
    assert gate._check_icon_coverage(sl([{"heading": "A", "icon": "growth"},
                                         {"heading": "B", "icon": "shield"}]))["passed"]
    assert gate._check_icon_coverage(sl([{"heading": "A"}, {"heading": "B"}]))["passed"]

def test_gate_fails_label_title(tmp_path):
    res = _gate_build(tmp_path, {"slides":[
        {"intent":"show_columns","content":{"title":"Overview","columns":[
            {"heading":"A","icon":"bolt","bullets":["x"]}]}}]})
    assert res["passed"] is False
    assert any("action_titles" in f for f in res["fail_items"])

def test_gate_writes_machine_readable_bool(tmp_path):
    res = _gate_build(tmp_path, {"slides":[{"intent":"introduce_topic","content":{"title":"Cover"}}]})
    assert isinstance(res["passed"], bool) and "overall_score" in res and "verdict" in res


# ---------- new composite components ----------
def _shape_count(pptx):
    from pptx import Presentation
    return [len(s.shapes._spTree) for s in Presentation(pptx).slides]

def test_components_build_and_pass_gate(tmp_path):
    deck={"slides":[
      {"intent":"stat_callout","content":{"title":"Three numbers that frame the case",
        "stats":[{"value":"50%","heading":"A","text":"x"},{"value":"3x","heading":"B","text":"y"}]}},
      {"intent":"icon_rows","content":{"title":"Four capabilities power the platform",
        "rows":[{"icon":"bolt","heading":"Fast","text":"x"},{"icon":"clock","heading":"On time","text":"y"}]}},
      {"intent":"banner","content":{"title":"The sequence runs left to right here",
        "banners":[{"text":"One","direction":"right","style":"navy"},{"text":"Two","style":"amber"}]}},
      {"intent":"label_stack","content":{"title":"Three tiers of the segmentation model",
        "labels":[{"text":"Tier 1"},{"text":"Tier 2"}]}},
      {"intent":"compare","content":{"title":"Build versus buy at a glance now",
        "compare":{"pros":["a","b"],"cons":["c"]}}},
    ]}
    res=_gate_build(tmp_path, deck)
    assert res["passed"] is True
    # every slide drew shapes onto the canvas
    assert all(c>1 for c in _shape_count(tmp_path/"d.pptx"))

def test_component_intents_registered():
    import compiler
    for name in ("stat_callout","icon_rows","compare"):
        assert name in compiler.INTENTS
        assert name in compiler.STRATEGIES
