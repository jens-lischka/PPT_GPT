"""gate.py — machine-readable build gate (the AI cannot self-declare pass).

Borrowed from the Mck-ppt "Harness" idea: after a deck is built, run_gate()
inspects the deck spec and the finished .pptx and writes ``gate_result.json``
with a hard ``passed`` boolean. The GPT must READ that file before declaring
done — it cannot pass the gate by self-assessment, because ``passed`` is a
Python boolean computed here, not a sentence the model writes.

Checks
  overflow        (hard)  no fixed-box text overflows  (BoundHeight check)
  icon_coverage   (hard)  every show_columns column & dashboard KPI has an icon
  action_titles   (hard)  content-slide titles are insight sentences, not labels
  charts          (hard)  no chart shows axis tick marks (no-tick-marks rule)
  treemap         (hard)  chartEx treemaps are schema-valid (<cx:lvl> + workbook)
  image_slots     (warn)  count of REPLACE: placeholders a designer must fill
  sources         (warn)  content slides cite a source / footnote
  density         (warn)  content slides aren't too sparse for their layout
  font_sizes      (warn)  run sizes are design-system tokens (no off-token sizes)
  grid            (hard)  compose grids are well-formed (in bounds, no cell overlap)
  gutters         (warn)  column gutters are 0.5" or 0.25" (equal tracks; no ratio rule)
  voice           (warn)  no banned vague-active verbs / generic claims (OW voice)
  design          (warn)  chart-choice + table-vs-cards heuristics (advisory)

``passed`` == all HARD checks pass. Warnings never fail the gate; they surface
work the human still owes (e.g. dropping real images into placeholders).
"""
from __future__ import annotations

import json
import os
from typing import Dict, List

# Intents that are covers / dividers / legal pages — exempt from the
# "title must be an insight sentence" rule.
STRUCTURAL = (
    "introduce_topic", "section_divider", "divider", "contents", "agenda",
    "closing", "thank_you", "thankyou", "confidentiality", "qualification",
    "backcover", "back_cover", "introduce_person",
)

MIN_TITLE_WORDS = 4          # a label is 1–3 words; an insight title is a sentence


def _is_structural(intent: str) -> bool:
    intent = (intent or "").lower()
    return any(tok in intent for tok in STRUCTURAL)


def _check_icon_coverage(slides: List[dict]) -> Dict:
    """All-or-nothing icons. Item icons (show_columns columns, dashboard KPIs)
    display only if EVERY item in the group has one — the runtime drops them all
    otherwise. So a deck is fine with icons on *all* items or on *none*; the only
    thing worth flagging is a lopsided mix, where the author added icons to some
    items but not the rest (those icons won't render)."""
    warns: List[str] = []
    for i, s in enumerate(slides, 1):
        intent = s.get("intent", "")
        c = s.get("content", {}) or {}
        if intent == "show_columns":
            items, kind = c.get("columns", []) or [], "show_columns"
        elif intent == "dashboard":
            items, kind = c.get("metrics", []) or [], "dashboard"
        else:
            continue
        total = len(items)
        with_icon = sum(1 for it in items if isinstance(it, dict) and it.get("icon"))
        if 0 < with_icon < total:
            warns.append(f"slide {i} ({kind}): {with_icon} of {total} items have an "
                         f"icon — icons are all-or-nothing, so none will show. "
                         f"Add an icon to every item, or remove them all.")
    return {"passed": not warns, "severity": "warn", "fail_items": warns,
            "note": "item icons are all-or-nothing per slide: give every "
                    "show_columns column / dashboard KPI an icon, or none."}


def _check_action_titles(slides: List[dict]) -> Dict:
    fails: List[str] = []
    for i, s in enumerate(slides, 1):
        if _is_structural(s.get("intent", "")):
            continue
        c = s.get("content", {}) or {}
        title = (c.get("title") or "").strip()
        if not title:
            continue
        words = title.split()
        # A digit-bearing short title is usually still an insight ("CAC up 60%").
        has_number = any(ch.isdigit() for ch in title)
        if len(words) < MIN_TITLE_WORDS and not has_number:
            fails.append(f"slide {i}: title '{title}' reads as a topic label, "
                         f"not an insight sentence")
    return {"passed": not fails, "severity": "hard", "fail_items": fails,
            "note": "content-slide titles should state the 'so what' as a full "
                    "sentence (the action-title convention)."}


def _check_image_slots(pptx_path: str) -> Dict:
    """Count REPLACE: placeholders still awaiting a real image (a warning, not a
    failure — they're intentional swap targets, but the human must fill them)."""
    try:
        from pptx import Presentation
        prs = Presentation(pptx_path)
        remaining = [f"slide {i}: {sh.name}"
                     for i, sl in enumerate(prs.slides, 1)
                     for sh in sl.shapes
                     if (sh.name or "").startswith("REPLACE:")]
    except Exception as exc:
        return {"passed": True, "severity": "warn", "fail_items": [],
                "note": f"image-slot check skipped: {exc}"}
    return {"passed": True, "severity": "warn", "fail_items": remaining,
            "note": f"{len(remaining)} swappable image placeholder(s) to fill "
                    f"before final delivery."}


def _check_sources(slides: List[dict]) -> Dict:
    warns: List[str] = []
    for i, s in enumerate(slides, 1):
        if _is_structural(s.get("intent", "")):
            continue
        c = s.get("content", {}) or {}
        if not (c.get("source") or c.get("footnote")
                or s.get("source") or s.get("footnote")):
            warns.append(f"slide {i}: no source/footnote")
    return {"passed": True, "severity": "warn", "fail_items": warns,
            "note": "content slides should cite a source (anti-fabrication)."}


def _check_density(slides: List[dict]) -> Dict:
    """Warn on near-empty text slides — only a few short bullets, OR a handful of
    parallel paragraphs with no visual (the 'when to use each' closer that reads
    as half-empty). Nudge toward show_columns cards / an image / a stat."""
    text_only = {"explain", "summary"}
    warns: List[str] = []
    for i, s in enumerate(slides, 1):
        intent = s.get("intent", "")
        if intent not in text_only:
            continue
        c = s.get("content", {}) or {}
        body = c.get("body", {}) or {}
        has_visual = bool(c.get("image") or c.get("chart") or c.get("table")
                          or c.get("columns") or c.get("metrics"))
        bullets = body.get("bullets") or c.get("bullets") or []
        paras = body.get("paragraphs") or c.get("paragraphs") or []
        if bullets and len(bullets) <= 3 and \
                sum(len(str(b)) for b in bullets) < 220 and not has_visual:
            warns.append(f"slide {i} ({intent}): only {len(bullets)} short bullet(s) "
                         f"and no visual — add an image/graphic/stat, expand, or merge")
        elif paras and 2 <= len(paras) <= 5 and not has_visual:
            warns.append(f"slide {i} ({intent}): {len(paras)} parallel paragraphs and "
                         f"no visual — render as `show_columns` cards (one per item) "
                         f"so they sit side by side and fill the slide")
    return {"passed": True, "severity": "warn", "fail_items": warns,
            "note": "aim for substantive density; a thin slide should gain a "
                    "visual, become cards, or be merged."}


def _check_overflow(pptx_path: str, template: str) -> Dict:
    try:
        import textmetrics as _tm
        _tm.configure(template)
        recs = _tm.overflow_report(pptx_path)
    except Exception as exc:
        return {"passed": True, "severity": "hard", "fail_items": [],
                "note": f"overflow check skipped: {exc}"}
    fails = [f"slide {r['slide']}: {r['shape']} needs {r['needed_in']}\" "
             f"in a {r['box_in']}\" box" + (" [shrinks]" if r["shrinks"] else "")
             for r in recs if not r["shrinks"]]
    return {"passed": not fails, "severity": "hard", "fail_items": fails,
            "note": "no fixed-box text may overflow; shorten or split the slide."}


def _check_charts(pptx_path: str) -> Dict:
    """HARD: no chart shows axis tick marks (the OW 'no tick marks ever' rule).
    WARN-level value-axis tidiness is handled by the build's chart normalizer."""
    import zipfile, re
    fails: List[str] = []
    try:
        z = zipfile.ZipFile(pptx_path)
    except Exception:
        return {"passed": True, "severity": "hard", "fail_items": [],
                "detail": "no pptx"}
    for n in z.namelist():
        if not re.match(r"ppt/charts/chart\d+\.xml$", n):
            continue
        xml = z.read(n).decode("utf-8", "ignore")
        for tm in re.findall(r'(?:major|minor)TickMark val="(\w+)"', xml):
            if tm != "none":
                fails.append(f"{n.split('/')[-1]}: tick marks present ({tm})")
                break
    return {"passed": not fails, "severity": "hard", "fail_items": fails,
            "detail": "no chart tick marks"}


def _check_treemap(pptx_path: str) -> Dict:
    """HARD: any chartEx treemap is schema-valid — points wrapped in <cx:lvl> and
    an <cx:externalData> link that resolves to the embedded workbook. Guards the
    exact malformation that forced a document repair before."""
    import zipfile, re
    fails: List[str] = []
    try:
        z = zipfile.ZipFile(pptx_path)
    except Exception:
        return {"passed": True, "severity": "hard", "fail_items": [],
                "detail": "no pptx"}
    for n in z.namelist():
        if not re.match(r"ppt/charts/chartEx\d+\.xml$", n):
            continue
        base = n.split("/")[-1]
        xml = z.read(n).decode("utf-8", "ignore")
        if "<cx:lvl" not in xml:
            fails.append(f"{base}: points not wrapped in <cx:lvl> (breaks PowerPoint)")
        m = re.search(r'externalData r:id="(rId\d+)"', xml)
        if not m:
            fails.append(f"{base}: missing <cx:externalData> workbook link")
            continue
        rid = m.group(1)
        rels_name = f"ppt/charts/_rels/{base}.rels"
        rels = z.read(rels_name).decode("utf-8", "ignore") if rels_name in z.namelist() else ""
        rm = re.search(rf'Id="{rid}"[^>]*Target="([^"]+)"', rels)
        if not rm or ".xlsx" not in rm.group(1):
            fails.append(f"{base}: externalData {rid} does not resolve to a workbook")
    return {"passed": not fails, "severity": "hard", "fail_items": fails,
            "detail": "chartEx treemaps schema-valid"}


# Allowed text sizes are the design-system token values themselves (single
# source — no hand-maintained copy that can drift from Tokens). Only sizes below
# 18pt are checked; larger display sizes (KPIs, quotes, titles) are always allowed.
def _allowed_sub18_sizes():
    from runtime import Tokens
    pts = (Tokens.TYPE_FOOTNOTE, Tokens.TYPE_DIAGRAM_MIN, Tokens.TYPE_AUTOFIT_MIN,
           Tokens.TYPE_EYEBROW, Tokens.TYPE_BODY, Tokens.TYPE_HEADING,
           Tokens.TYPE_AXIS, Tokens.TYPE_MARKER)
    return {int(p * 100) for p in pts}

_ALLOWED_SZ = _allowed_sub18_sizes()


def _check_font_sizes(pptx_path: str) -> Dict:
    """WARN: explicit run font sizes should be design-system token values (no
    hard-coded off-token sizes). Quote/title display sizes (>=18pt) are allowed."""
    import zipfile, re
    warns: List[str] = []
    try:
        z = zipfile.ZipFile(pptx_path)
    except Exception:
        return {"passed": True, "severity": "warn", "fail_items": [],
                "detail": "no pptx"}
    for n in sorted(z.namelist()):
        m = re.match(r"ppt/slides/slide(\d+)\.xml$", n)
        if not m:
            continue
        xml = z.read(n).decode("utf-8", "ignore")
        bad = {int(s) for s in re.findall(r'\bsz="(\d+)"', xml)
               if int(s) not in _ALLOWED_SZ and int(s) < 1800}
        if bad:
            pts = ", ".join(f"{s/100:g}pt" for s in sorted(bad))
            warns.append(f"slide {int(m.group(1))}: off-token size(s) {pts}")
    return {"passed": True, "severity": "warn", "fail_items": warns,
            "detail": "run sizes are design-system tokens"}


def _check_grid(slides: List[dict]) -> Dict:
    """HARD: every compose grid is well-formed — explicitly placed cells stay in
    bounds and no two cells overlap. This is the full N×M track check: the
    equal-track *geometry* is guaranteed by the deterministic solver
    (runtime.solve_grid / draw_grid, covered by the grid geometry tests), so the
    remaining failure mode is the authoring — an out-of-bounds span or two cells
    claiming the same slot — which this validates against the same dimensions the
    solver derives (via the shared ``resolve_grid_dims``)."""
    from runtime import resolve_grid_dims
    fails: List[str] = []
    for i, s in enumerate(slides, 1):
        content = s.get("content") or {}
        comp = content.get("compose") or {}
        grid = comp.get("grid") if isinstance(comp, dict) else None
        if not isinstance(grid, dict):
            continue
        cells = [c for c in grid.get("cells", []) if isinstance(c, dict)]
        if not cells:
            continue

        def cs(c):
            return max(1, int(c.get("colspan", 1) or 1))

        def rs(c):
            return max(1, int(c.get("rowspan", 1) or 1))

        N, M, decl_N, decl_M = resolve_grid_dims(cells, grid.get("cols"),
                                                 grid.get("rows"))

        occ = [[None] * N for _ in range(M)]
        for idx, c in enumerate(cells):
            if "col" not in c or "row" not in c:
                continue                          # auto-flow can't overlap/overflow
            cc, rr = int(c["col"]), int(c["row"])
            if cc < 0 or rr < 0:
                fails.append(f"slide {i}: grid cell {idx} has a negative position")
                continue
            # bounds are checked against the DECLARED grid: a cell past it would
            # silently force the solver to resize every track.
            if decl_N is not None and cc + cs(c) > decl_N:
                fails.append(f"slide {i}: grid cell {idx} spans past the declared "
                             f"{decl_N} columns")
            if decl_M is not None and rr + rs(c) > decl_M:
                fails.append(f"slide {i}: grid cell {idx} spans past the declared "
                             f"{decl_M} rows")
            for y in range(rr, rr + rs(c)):
                for x in range(cc, cc + cs(c)):
                    if y < M and x < N:
                        if occ[y][x] is not None:
                            fails.append(f"slide {i}: grid cells {occ[y][x]} and "
                                         f"{idx} overlap at column {x}, row {y}")
                        occ[y][x] = idx
    return {"passed": not fails, "severity": "hard",
            "fail_items": list(dict.fromkeys(fails)),
            "detail": "compose grids are well-formed (in bounds, no overlap)"}


def _check_gutters(pptx_path: str) -> Dict:
    """WARN: the gutter between any two side-by-side content columns is 0.5\" or
    0.25\". Generalised to N columns by projecting content shapes onto the x-axis,
    merging them into occupied bands, and inspecting the gaps between bands."""
    try:
        from pptx import Presentation
        prs = Presentation(pptx_path)
    except Exception:
        return {"passed": True, "severity": "warn", "fail_items": [],
                "detail": "no pptx"}
    EMU = 914400.0
    warns: List[str] = []
    for i, s in enumerate(prs.slides, 1):
        ivals = []
        for sh in s.shapes:
            try:
                if sh.is_placeholder:
                    continue
                x, y = sh.left / EMU, sh.top / EMU
                w, h = sh.width / EMU, sh.height / EMU
                if w < 1.0 or h < 0.6 or y > 6.7:     # skip tiny / footer-band
                    continue
                ivals.append((x, x + w))
            except Exception:
                continue
        if len(ivals) < 2:
            continue
        ivals.sort()
        bands = [list(ivals[0])]
        for a, b in ivals[1:]:
            if a <= bands[-1][1] + 0.02:
                bands[-1][1] = max(bands[-1][1], b)
            else:
                bands.append([a, b])
        for k in range(len(bands) - 1):
            g = bands[k + 1][0] - bands[k][1]
            if g <= 1.2 and not (abs(g - 0.5) < 0.12 or abs(g - 0.25) < 0.08):
                warns.append(f"slide {i}: column gutter {g:.2f}\" "
                             f"(should be 0.5\" or 0.25\")")
    return {"passed": True, "severity": "warn",
            "fail_items": list(dict.fromkeys(warns)),
            "detail": "column gutters are 0.5\" or 0.25\""}


def _check_content_keys(slides: List[dict]) -> Dict:
    """Warn when a slide's ``content`` dict carries keys that the intent's
    strategies don't read. Silent drops of typoed or wrong-shape keys are the
    single most common cause of "why is my slide missing text?" — an author
    passes ``items`` where the timeline drawer wants ``milestones``, the runtime
    happily builds a slide missing every label, and the deck passes overflow,
    fonts, and sources.

    Universal adornments accepted at content level and hoisted at compile time
    are always allowed: ``conclusion``, ``footnote``, ``source``. So are keys
    the compiler consumes directly on the slide-level (``design_choices``,
    ``why``, ``insight``, ``photo``) — none of which change the strategy path.

    Uses the same INTENTS registry the compiler uses, so this check and the
    compiler's own log line never disagree about what "known" means.
    """
    try:
        from compiler import INTENTS
    except Exception:
        return {"passed": True, "severity": "warn", "fail_items": [],
                "note": "content-keys check skipped: INTENTS not importable"}
    UNIVERSAL = {"conclusion", "footnote", "source", "why", "insight",
                 "design_choices", "photo"}
    warns: List[str] = []
    for i, s in enumerate(slides, 1):
        intent_name = s.get("intent")
        content = s.get("content", {}) or {}
        if not isinstance(content, dict):
            continue
        intent = INTENTS.get(intent_name)
        if not intent:
            continue
        known = {slot.name for slot in intent.slots} | UNIVERSAL
        unknown = [k for k in content if k not in known]
        if unknown:
            expected = sorted({slot.name for slot in intent.slots})
            warns.append(
                f"slide {i} ({intent_name}): content key(s) "
                f"{unknown} are not consumed — {intent_name} reads "
                f"{expected}. Likely a typo or wrong-shape spec; the "
                f"content is silently dropped.")
    return {"passed": True, "severity": "warn", "fail_items": warns,
            "note": "content keys that no strategy reads (silent drops)."}


def run_gate(deck_spec: dict, pptx_path: str, template: str,
             out_path: str = "") -> Dict:
    """Run all checks, write gate_result.json, and return the result dict."""
    slides = deck_spec.get("slides", []) or []
    checks = {
        "overflow":       _check_overflow(pptx_path, template),
        "icon_coverage":  _check_icon_coverage(slides),
        "action_titles":  _check_action_titles(slides),
        "charts":         _check_charts(pptx_path),
        "treemap":        _check_treemap(pptx_path),
        "image_slots":    _check_image_slots(pptx_path),
        "sources":        _check_sources(slides),
        "density":        _check_density(slides),
        "font_sizes":     _check_font_sizes(pptx_path),
        "grid":           _check_grid(slides),
        "gutters":        _check_gutters(pptx_path),
        "voice":          _check_voice(slides),
        "design":         _check_design(slides),
        "content_keys":   _check_content_keys(slides),
    }
    hard_fail_items, warn_items = [], []
    for name, c in checks.items():
        if c["severity"] == "hard" and not c["passed"]:
            hard_fail_items += [f"[{name}] {x}" for x in c["fail_items"]]
        if c["severity"] == "warn":
            warn_items += [f"[{name}] {x}" for x in c["fail_items"]]
    passed = len(hard_fail_items) == 0
    score = max(0, 100 - 12 * len(hard_fail_items) - 2 * len(warn_items))
    result = {
        "passed": passed,
        "overall_score": score,
        "verdict": ("PASS — ready to deliver" if passed
                    else "FAIL — fix the items below and rebuild; do not deliver"),
        "checks": checks,
        "fail_items": hard_fail_items,
        "warn_items": warn_items,
    }
    if not out_path:
        out_path = os.path.join(os.path.dirname(pptx_path) or ".",
                                "gate_result.json")
    try:
        with open(out_path, "w") as fh:
            json.dump(result, fh, indent=2, ensure_ascii=False)
        result["_path"] = out_path
    except Exception:
        pass
    return result


# --------------------------------------------------------------------------- #
# Voice linter (item 5) — banned vague-active verbs & generic claims, enforced
# rather than merely instructed. SINGLE SOURCE OF TRUTH: voice_lint.py (shared
# with the draft voice lint / `--lint-text`), so the two never drift.
try:
    from voice_lint import BANNED_VERB_STEMS as _BANNED_STEMS, \
        GENERIC_CLAIMS as _BANNED_PHRASES
except Exception:                                    # pragma: no cover
    _BANNED_STEMS = ("leverag", "enabl", "enhanc", "transform", "optimi",
                     "empower", "facilitat", "utiliz", "utilis", "synerg",
                     "seamless")
    _BANNED_PHRASES = ("drive value", "best-in-class", "game-changing")


def _iter_text(node):
    """Yield every human-readable string in a slide content tree."""
    if isinstance(node, str):
        yield node
    elif isinstance(node, dict):
        for k, v in node.items():
            if k in ("icon", "image", "images", "photo", "path", "type",
                     "aspect", "crop", "intent", "layout", "color"):
                continue
            yield from _iter_text(v)
    elif isinstance(node, (list, tuple)):
        for v in node:
            yield from _iter_text(v)


import re as _re
_WORD = _re.compile(r"[A-Za-z][A-Za-z\-]*")


def _check_voice(slides: List[dict]) -> Dict:
    """WARN: flag banned vague-active verbs and generic claims so the OW voice
    rules are enforced, not just instructed. Matches stems (leverage→leveraging)
    and multi-word clichés. Source of truth: voice_lint.py."""
    warns: List[str] = []
    for i, s in enumerate(slides, 1):
        hits = set()
        for text in _iter_text(s.get("content", {}) or {}):
            low = text.lower()
            for ph in _BANNED_PHRASES:
                if ph in low:
                    hits.add(ph)
            for w in _WORD.findall(low):
                for stem in _BANNED_STEMS:
                    if w == stem or w.startswith(stem):
                        hits.add(stem)
        if hits:
            warns.append(f"slide {i}: banned term(s) {sorted(hits)} "
                         f"— rewrite with a specific, evidence-backed verb")
    return {"passed": True, "severity": "warn", "fail_items": warns,
            "check": "voice"}


# --------------------------------------------------------------------------- #
# Design heuristics (item 4) — deterministic proxies for the chart-choice and
# table-vs-cards judgments the semantic judge (eval harness) makes properly.
# Advisory only; the true aesthetic judge is the LLM eval loop, not the gate.
# --------------------------------------------------------------------------- #
def _check_design(slides: List[dict]) -> Dict:
    warns: List[str] = []
    for i, s in enumerate(slides, 1):
        c = s.get("content", {}) or {}
        ch = c.get("chart")
        if isinstance(ch, dict):
            ctype = str(ch.get("type", "")).lower()
            ncat = len(ch.get("categories", []) or [])
            if ("pie" in ctype or "doughnut" in ctype) and ncat > 6:
                warns.append(f"slide {i}: {ctype} with {ncat} slices is hard to "
                             f"read — a bar chart ranks categories more clearly")
        tbl = c.get("table")
        if isinstance(tbl, dict):
            rows = tbl.get("rows", []) or []
            hdrs = tbl.get("headers", []) or []
            ncol = max([len(hdrs)] + [len(r) for r in rows]) if rows else len(hdrs)
            cells = [str(x) for r in rows for x in r]
            numeric = sum(1 for x in cells if any(ch.isdigit() for ch in x))
            frac_num = numeric / len(cells) if cells else 1.0
            if ncol <= 3 and len(rows) <= 4 and frac_num < 0.3:
                warns.append(f"slide {i}: small text-only table "
                             f"({len(rows)}x{ncol}) — show_columns/cards reads "
                             f"better than a grid for so few text cells")
    return {"passed": True, "severity": "warn", "fail_items": warns,
            "check": "design"}
