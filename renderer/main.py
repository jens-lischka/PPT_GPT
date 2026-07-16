#!/usr/bin/env python3
"""Renderer service — thin FastAPI wrapper around the unchanged OW deck runtime.

Endpoints:
  POST /render        full deck_spec  -> {pptx_base64, gate_result, slide_count}
  POST /render-slide  single slide    -> {pptx_base64, gate_result} (1-slide deck)
  GET  /healthz       liveness + runtime version

The runtime (src/) is byte-identical to the GPT package's
deck_system_runtime_clean.zip v2.6.18 — do NOT fork it here. When the runtime
evolves, replace src/ wholesale from the package and bump the pin in
CLAUDE.md.
"""
from __future__ import annotations

import base64
import json
import logging
import os
import re
import sys
import tempfile
from pathlib import Path
from typing import Any

from fastapi import FastAPI, HTTPException
from fastapi.responses import JSONResponse, Response
from fastapi.middleware.gzip import GZipMiddleware
from starlette.requests import Request
from pydantic import BaseModel

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE / "src"))

from runtime import TemplateRuntime, __version__ as RUNTIME_VERSION  # noqa: E402
from compiler import SemanticCompiler, INTENTS as _INTENTS           # noqa: E402
from gate import run_gate                                           # noqa: E402

import jobs                                                         # noqa: E402

TEMPLATE = str(HERE / "ow_default.pptx")

app = FastAPI(title="OW deck renderer", version=RUNTIME_VERSION)

# Compress the ~1 MB base64 pptx JSON on the wire (added first = innermost, so
# the CORS/error middleware still stamps headers on the compressed response).
app.add_middleware(GZipMiddleware, minimum_size=2048)

# CORS + error handling in one middleware. The stock CORSMiddleware does NOT
# attach Access-Control-Allow-Origin to responses produced by unhandled
# exceptions (they bypass it), so a server-side 500 shows up in the browser as a
# misleading "No 'Access-Control-Allow-Origin' header" instead of the real
# error. Here we (a) answer preflight, (b) catch every exception and return it as
# JSON, and (c) stamp CORS headers on EVERY response — so the pane always sees
# the true error text. ALLOWED_ORIGINS ("*" default for the test phase) echoes a
# specific origin when configured.
_allowed = os.environ.get("ALLOWED_ORIGINS", "*").strip()
_ORIGINS = None if _allowed == "*" else {o.strip() for o in _allowed.split(",") if o.strip()}


def _cors(origin: str | None) -> dict[str, str]:
    if _ORIGINS is None:
        allow = "*"
    elif origin and origin in _ORIGINS:
        allow = origin
    else:
        allow = next(iter(_ORIGINS))
    return {
        "Access-Control-Allow-Origin": allow,
        "Access-Control-Allow-Methods": "POST, GET, OPTIONS",
        "Access-Control-Allow-Headers": "authorization, content-type",
        "Vary": "Origin",
    }


@app.middleware("http")
async def cors_and_errors(request: Request, call_next):
    origin = request.headers.get("origin")
    if request.method == "OPTIONS":
        return Response(status_code=204, headers=_cors(origin))
    try:
        resp = await call_next(request)
    except Exception as exc:                                  # noqa: BLE001
        resp = JSONResponse(status_code=500,
                            content={"error": f"{type(exc).__name__}: {exc}"})
    for k, v in _cors(origin).items():
        resp.headers[k] = v
    return resp


class RenderRequest(BaseModel):
    deck_spec: dict[str, Any]


class RenderSlideRequest(BaseModel):
    slide_spec: dict[str, Any]          # one element of deck_spec["slides"]
    language: str | None = None         # carried over from the deck level
    audience: str | None = None


class AgentRequest(BaseModel):
    """Single-service orchestration: Claude drafts the spec, then we render.
    Mirrors the Supabase edge function's contract so the task pane can point at
    either backend unchanged."""
    mode: str                           # storyline|generate|generate_slide|edit_slide
    prompt: str | None = None
    storyline: Any = None               # object or list of slide outlines
    slide_spec: dict[str, Any] | None = None
    command: str | None = None
    language: str | None = None
    audience: str | None = None
    background: bool = False            # True → return {job_id}; poll /agent/jobs/{id}
    slide_pptx_base64: str | None = None  # edit_slide: export of the LIVE slide
    slide_state: dict[str, Any] | None = None  # shape_ops: live shape inventory


# ---------------------------------------------------------------------------
# Convenience macros — a small, forgiving layer over the strict runtime schema.
# The model emits an EASY shape; Python assembles the error-prone geometry (e.g.
# a compose grid). This kills whole classes of "blank slide" schema mistakes
# without the model ever writing compose-grid mechanics. Runtime src/ untouched.
# ---------------------------------------------------------------------------
def _expand_macros(deck_spec: dict) -> dict:
    slides = deck_spec.get("slides")
    if not isinstance(slides, list):
        return deck_spec
    if not any(isinstance(s, dict) and s.get("intent") in _MACROS for s in slides):
        return deck_spec
    out = dict(deck_spec)
    out["slides"] = [_expand_slide(s) for s in slides]
    return out


def _expand_slide(slide: dict) -> dict:
    if not isinstance(slide, dict) or slide.get("intent") not in _MACROS:
        return slide
    return _MACROS[slide["intent"]](slide.get("content") or {})


def _macro_columns_layout(content: dict) -> dict:
    """columns_layout → compose. Uses the ROWS form with weighted heights so a
    heading is a thin band directly above its visual — the old grid form gave
    the heading a full equal-height row, leaving a huge gap below it. Rows keep
    columns aligned because every row carries one region per column (empty
    text placeholders pad columns that lack a heading/body)."""
    cols = [c for c in (content.get("columns") or []) if isinstance(c, dict)]

    def _visual(col: dict) -> dict | None:
        if col.get("chart"):
            return {"block": "chart", "data": {"chart": col["chart"]}}
        if col.get("table"):
            return {"block": "table", "data": col["table"]}
        if col.get("kpi"):
            return {"block": "kpi", "data": col["kpi"]}
        if col.get("image"):
            return {"block": "image", "data": col["image"]}
        if col.get("waterfall"):
            return {"block": "waterfall", "data": {"waterfall": col["waterfall"]}}
        if col.get("treemap"):
            return {"block": "treemap", "data": {"treemap": col["treemap"]}}
        return None

    def _body(col: dict) -> dict:
        if col.get("bullets"):
            return {"block": "text", "data": {"bullets": col["bullets"]}}
        if col.get("text"):
            txt = col["text"]
            return {"block": "text",
                    "data": {"paragraphs": [txt] if isinstance(txt, str) else txt}}
        return {"block": "text", "data": {}}          # alignment placeholder

    has_head = any(c.get("heading") for c in cols)
    has_body = any(c.get("bullets") or c.get("text") for c in cols)
    rows: list[dict] = []
    if has_head:
        rows.append({"height": 0.13, "regions": [
            {"block": "text",
             "data": {"heading": c["heading"]} if c.get("heading") else {}}
            for c in cols]})
    rows.append({"height": 0.59 if (has_head or has_body) else 1.0, "regions": [
        (_visual(c) or {"block": "text", "data": {}}) for c in cols]})
    if has_body:
        rows.append({"height": 0.28, "regions": [_body(c) for c in cols]})

    new_content: dict[str, Any] = {"title": content.get("title", "")}
    if content.get("footnote"):
        new_content["footnote"] = content["footnote"]
    new_content["compose"] = {"rows": rows}
    return {"intent": "compose", "content": new_content}


_MACROS = {"columns_layout": _macro_columns_layout}


# ---------------------------------------------------------------------------
# Intent normalization + compile-error capture.
#
# The runtime compiler SKIPS a slide whose intent is unknown or whose content
# blows up (logger.error + continue) — the deck renders with fewer slides, the
# gate can still pass, and the pane's uuid↔spec index mapping silently
# misaligns (later edits then hit the wrong spec). Two defenses, both outside
# the untouched runtime:
#   1. _normalize_intents: map the model's usual near-misses (e.g. "quote",
#      "show_trend", "timeline") onto real registry names before compiling.
#   2. a logging handler on the "compiler" logger records every
#      "Slide N failed to compile" during build, so dropped slides become
#      actionable feedback (repair / block) instead of a vanished slide.
# ---------------------------------------------------------------------------
_INTENT_ALIASES = {
    "quote": "show_quote",
    "trend": "show_trend_with_key_message",
    "show_trend": "show_trend_with_key_message",
    "show_chart": "show_trend_with_key_message",
    "chart": "show_trend_with_key_message",
    "timeline": "show_timeline",
    "process": "show_process",
    "table": "show_data",
    "data": "show_data",
    "columns": "show_columns",
    "matrix": "show_matrix",
    "org_chart": "show_org_chart",
    "waterfall": "show_waterfall",
    "treemap": "show_treemap",
    "pyramid": "show_pyramid",
    "cycle": "show_cycle",
    "contents": "show_contents",
    "toc": "show_contents",
    "agenda": "show_contents",
    "kpi": "dashboard",
    "kpis": "dashboard",
    "metrics": "dashboard",
    "stats": "stat_callout",
    "stat": "stat_callout",
    "person": "introduce_person",
    "cover": "introduce_topic",
    "title_slide": "introduce_topic",
    "divider": "section_divider",
    "section": "section_divider",
    "pros_cons": "compare",
    "comparison": "compare_two_options",
    "column_layout": "columns_layout",
    "multi_column": "columns_layout",
}


def _canon_intent(name: Any) -> str | None:
    """Best-effort map of a model-emitted intent onto the registry (or macro)
    name. Returns None when there is no defensible match."""
    if isinstance(name, str) and (name in _INTENTS or name in _MACROS):
        return name
    n = str(name or "").strip().lower().replace("-", "_").replace(" ", "_")
    if n in _INTENTS or n in _MACROS:
        return n
    if n in _INTENT_ALIASES:
        return _INTENT_ALIASES[n]
    if "show_" + n in _INTENTS:
        return "show_" + n
    if n.startswith("show_") and n[5:] in _INTENTS:
        return n[5:]
    return None


def _normalize_intents(deck_spec: dict) -> tuple[dict, list[tuple[int, str]]]:
    """Canonicalize slide intents in a copy of deck_spec. Returns the new spec
    and the 1-based (index, bad_name) list of slides that stay unknown."""
    slides = deck_spec.get("slides")
    if not isinstance(slides, list):
        return deck_spec, []
    unknown: list[tuple[int, str]] = []
    new_slides = []
    for i, s in enumerate(slides, 1):
        if isinstance(s, dict):
            canon = _canon_intent(s.get("intent"))
            if canon is None:
                unknown.append((i, str(s.get("intent"))))
            elif canon != s.get("intent"):
                s = {**s, "intent": canon}
            s = _fix_content_shapes(s)
        new_slides.append(s)
    if not unknown and all(a is b for a, b in zip(new_slides, slides)):
        return deck_spec, []
    return {**deck_spec, "slides": new_slides}, unknown


def _fix_content_shapes(slide: dict) -> dict:
    """Forgiving-layer fixups for shapes the model plausibly gets wrong and the
    runtime then silently degrades on. Currently: a plain-string 'insight' on
    show_trend_with_key_message breaks the insight-card strategy (needs
    {title?, text}) — the slide would render WITHOUT its colored callout."""
    if not isinstance(slide, dict):
        return slide
    c = slide.get("content")
    if (slide.get("intent") == "show_trend_with_key_message"
            and isinstance(c, dict) and isinstance(c.get("insight"), str)):
        return {**slide, "content": {**c, "insight": {"text": c["insight"]}}}
    return slide


class _CompileErrorCapture(logging.Handler):
    """Collects the compiler's 'Slide N failed to compile: …' errors so the
    caller learns WHICH slides were silently dropped."""
    def __init__(self) -> None:
        super().__init__(level=logging.ERROR)
        self.items: list[tuple[int, str]] = []

    def emit(self, record: logging.LogRecord) -> None:
        m = re.match(r"Slide (\d+) failed to compile: (.*)", record.getMessage())
        if m:
            self.items.append((int(m.group(1)), m.group(2)))


def _render(deck_spec: dict) -> dict:
    deck_spec, unknown = _normalize_intents(deck_spec)
    if unknown:
        names = ", ".join(f"slide {i}: {n!r}" for i, n in unknown)
        raise ValueError(
            f"unknown intent(s) — {names}. Valid intents: "
            + ", ".join(sorted(set(_INTENTS) | set(_MACROS))))
    deck_spec = _expand_macros(deck_spec)          # convenience macros → real intents
    capture = _CompileErrorCapture()
    compiler_logger = logging.getLogger("compiler")
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "deck.pptx"
        gate_path = Path(td) / "gate_result.json"
        rt = TemplateRuntime(TEMPLATE)
        compiler_logger.addHandler(capture)
        try:
            SemanticCompiler(rt).build(deck_spec)
        finally:
            compiler_logger.removeHandler(capture)
        rt.save(str(out))
        gate = run_gate(deck_spec, str(out), TEMPLATE, str(gate_path))
        gate.pop("_path", None)
        raw = out.read_bytes()
        analysis = _analyze(raw)                             # ONE parse; reused downstream
        return {
            "pptx_base64": base64.b64encode(raw).decode(),
            "gate_result": gate,
            "slide_count": len(analysis),
            "runtime_version": RUNTIME_VERSION,
            "_analysis": analysis,
            # 1-based SPEC indices of slides the compiler dropped (+ why) —
            # valid even when rendered count no longer matches the spec.
            "_compile_errors": capture.items,
        }


def _analyze(pptx: bytes | str) -> list[dict]:
    """Parse the pptx ONCE into a per-slide inventory (charts/tables/pics + a
    body-content count). Everything that used to re-open the file reads this."""
    import io
    from pptx import Presentation
    raw = pptx if isinstance(pptx, (bytes, bytearray)) else base64.b64decode(pptx)
    out = []
    for s in Presentation(io.BytesIO(raw)).slides:
        charts = tables = pics = body = 0
        for sh in s.shapes:
            c = sh.has_chart
            t = sh.has_table
            p = sh.shape_type == 13
            charts += c; tables += t; pics += p
            center = ((sh.top or 0) + (sh.height or 0) / 2) / 914400
            if c or t or p or (1.4 < center < 6.2):
                body += 1
        out.append({"charts": charts, "tables": tables, "pics": pics, "body": body})
    return out


def _as_analysis(x) -> list[dict]:
    return x if isinstance(x, list) else _analyze(x)


def _empty_slides(analysis) -> list[int]:
    """1-based indices of slides with NO body content (the empty-canvas gap the
    gate misses). Accepts a precomputed analysis list or a pptx base64 string."""
    return [i for i, a in enumerate(_as_analysis(analysis), 1) if a["body"] == 0]


def _inspect_slide(pptx_b64: str) -> dict | None:
    """Parse ONE exported slide (pptx base64 from Office.js exportAsBase64)
    into a compact inventory of what is ACTUALLY on it — the ground truth for
    edits, including any manual changes made after generation."""
    import io
    from pptx import Presentation
    try:
        prs = Presentation(io.BytesIO(base64.b64decode(pptx_b64)))
        slides = list(prs.slides)
    except Exception as exc:                                  # noqa: BLE001
        print(f"[inspect] unparseable slide export: {exc}", flush=True)
        return None
    if not slides:
        return None
    s = slides[0]
    inv: dict[str, Any] = {"texts": [], "tables": [], "charts": [],
                           "pictures": 0, "other_graphics": 0}
    for sh in s.shapes:
        try:
            if sh.has_text_frame:
                t = sh.text_frame.text.strip()
                if t:
                    inv["texts"].append(t[:400])
            if sh.has_table:
                rows = [[c.text.strip() for c in r.cells] for r in sh.table.rows]
                inv["tables"].append(rows[:20])
            elif sh.has_chart:
                ch = sh.chart
                try:
                    ctype = str(ch.chart_type).split(" ")[0]
                except Exception:                             # noqa: BLE001
                    ctype = "unknown"
                cats: list = []
                series: list = []
                try:
                    plot = ch.plots[0]
                    cats = [str(c) for c in plot.categories]
                    series = [{"name": sr.name,
                               "values": [None if v is None else round(v, 4)
                                          for v in sr.values]}
                              for sr in plot.series]
                except Exception:                             # noqa: BLE001
                    pass
                inv["charts"].append({"type": ctype, "categories": cats,
                                      "series": series})
            elif sh.shape_type == 6 or str(sh.shape_type).startswith("GRAPHIC"):
                inv["other_graphics"] += 1                    # chartEx/diagrams
            if sh.shape_type == 13:
                inv["pictures"] += 1
        except Exception:                                     # noqa: BLE001
            continue          # a single exotic shape must not kill the edit
    return inv


def _slide_requests(slide: dict, kind: str) -> bool:
    """Does this slide spec ask for a chart / table? (scan the JSON + intent)."""
    try:
        blob = json.dumps(slide)
    except Exception:                                        # noqa: BLE001
        blob = str(slide)
    intent = slide.get("intent") if isinstance(slide, dict) else None
    # Match the KEY ("chart":), not the bare string — icon values like
    # "icon": "chart" or prose containing the word must not count as a
    # requested visual (a false positive here loops the repair pass).
    if kind == "chart":
        return '"chart":' in blob or intent == "show_trend_with_key_message"
    if kind == "table":
        return '"table":' in blob or intent == "show_data"
    return False


def _shortfall_slides(deck_spec: dict, analysis) -> list[int]:
    """1-based indices where the spec asked for a chart/table but NONE rendered."""
    a = _as_analysis(analysis)
    slides = deck_spec.get("slides") or []
    if len(a) != len(slides):                                # overflow split → unmappable
        return []
    out = []
    for i, (spec, inv) in enumerate(zip(slides, a), 1):
        if _slide_requests(spec, "chart") and inv["charts"] == 0:
            out.append(i)
        elif _slide_requests(spec, "table") and inv["tables"] == 0:
            out.append(i)
    return out


def _bad_slides(deck_spec: dict, analysis, compile_errors=None) -> list[int]:
    """Slides that are blank, missing a requested visual, OR silently dropped
    by the compiler — none of these may ship."""
    a = _as_analysis(analysis)
    dropped = {i for i, _ in (compile_errors or [])}
    return sorted(set(_empty_slides(a)) | set(_shortfall_slides(deck_spec, a)) | dropped)


def _gate_feedback(gate: dict) -> str:
    """Render the gate's failing items as a fix list for the repair prompt.
    Hard fails block; warns only lower the score — list hard first, clearly."""
    hard, warn = [], []
    for name, chk in (gate.get("checks") or {}).items():
        for item in chk.get("fail_items", []):
            (hard if chk.get("severity") == "hard" and not chk.get("passed")
             else warn).append(f"[{name}] {item}")
    lines = []
    if hard:
        lines.append("MUST-FIX (these block the deck):")
        lines += [f"  - {x}" for x in hard]
    if warn:
        lines.append("SHOULD-FIX (raise quality):")
        lines += [f"  - {x}" for x in warn]
    return "\n".join(lines) or "(no specific items reported)"


@app.get("/healthz")
def healthz():
    return {"ok": True, "runtime_version": RUNTIME_VERSION}


@app.get("/selftest")
def selftest():
    """Browser-openable diagnostic (a GET, so it bypasses the pane, CORS
    preflight, and any POST filtering). Reports whether the API key is present
    and whether a minimal Claude call succeeds — returns the real error text."""
    key = os.environ.get("ANTHROPIC_API_KEY")
    info: dict[str, Any] = {
        "runtime_version": RUNTIME_VERSION,
        "key_present": bool(key),
        "key_prefix": (key[:8] + "…") if key else None,
        "model": os.environ.get("ANTHROPIC_MODEL", "claude-opus-4-8"),
        "allowed_origins": os.environ.get("ALLOWED_ORIGINS", "*"),
    }
    try:
        import agent as ag
        out = ag.call_claude('Reply with exactly {"pong": true} as JSON.',
                             "ping", max_tokens=64, thinking=False)
        info["claude"] = {"ok": True, "sample": out}
    except Exception as exc:                                  # noqa: BLE001
        import traceback
        traceback.print_exc()
        info["claude"] = {"ok": False, "error": f"{type(exc).__name__}: {exc}"}
    return info


@app.post("/render")
def render(req: RenderRequest):
    try:
        return _render(req.deck_spec)
    except Exception as exc:                                  # noqa: BLE001
        raise HTTPException(status_code=422, detail=f"{type(exc).__name__}: {exc}")


@app.post("/render-slide")
def render_slide(req: RenderSlideRequest):
    """Render a single slide as a 1-slide deck — the regenerate-and-replace
    path for selection-aware edits. Deck-level fields that affect rendering
    (language for chart locale, audience for composition routing) must be
    passed through or the slide will render with defaults."""
    deck: dict[str, Any] = {"slides": [req.slide_spec]}
    if req.language:
        deck["language"] = req.language
    if req.audience:
        deck["audience"] = req.audience
    try:
        return _render(deck)
    except Exception as exc:                                  # noqa: BLE001
        raise HTTPException(status_code=422, detail=f"{type(exc).__name__}: {exc}")


def _render_single(slide_spec: dict, language: str | None, audience: str | None) -> dict:
    deck: dict[str, Any] = {"slides": [slide_spec]}
    if language:
        deck["language"] = language
    if audience:
        deck["audience"] = audience
    return _render(deck)


def _one_slide(ag, system: str, user: str, language, audience):
    """Render one slide; if it comes out blank (empty-canvas gap), rebuild once
    with an explicit instruction. Returns (slide_spec, render_result)."""
    def _bad(slide, result):
        """Blank/missing-visual from the spec, PLUS command-aware: if the request
        text explicitly asked for a chart/table but none rendered, that's a fail
        even when the model quietly produced a text-only slide. Reuses the single
        parse in result['_analysis'] — no re-open of the pptx."""
        a = result["_analysis"]
        bad = set(_bad_slides({"slides": [slide]}, a, result.get("_compile_errors")))
        inv = a[0] if a else {"charts": 0, "tables": 0}
        low = user.lower()
        if any(w in low for w in ("chart", "graph", "plot")) and inv["charts"] == 0:
            bad.add(1)
        if "table" in low and inv["tables"] == 0:
            bad.add(1)
        return sorted(bad)

    jobs.progress("drafting the slide (fast model)")
    slide = ag.call_claude(system, user, fast=True)     # fast happy path
    jobs.progress("rendering")
    render_error = None
    try:
        result = _render_single(slide, language, audience)
        bad = _bad(slide, result)
    except ValueError as exc:           # unknown intent / compile-level reject
        render_error, result, bad = str(exc), None, [1]
    if bad:
        jobs.progress("first draft rejected — retrying with the strong model")
        print("[/agent] one-slide blank/missing-visual — retrying (strong model)", flush=True)
        reason = (f"it used an invalid intent ({render_error})" if render_error else
                  "it rendered with MISSING content — either an empty canvas, or a "
                  "chart/table the request asked for did not appear (you omitted it "
                  "or its data shape was wrong)")
        retry = (user + f"\n\nYour previous slide was REJECTED: {reason}. "
                 "Include EVERY element the request asks for. For multi-"
                 "column-with-visuals use the columns_layout intent; for a chart use "
                 '{"chart":{"type":"column|bar|line|pie|area","categories":[…],'
                 '"series":[{"name":"…","values":[…]}]}}. Return the corrected slide.')
        slide = ag.call_claude(system, retry, fast=False)   # escalate to strong
        jobs.progress("rendering the corrected slide")
        result = _render_single(slide, language, audience)  # 2nd failure → raise to caller
        bad = _bad(slide, result)
    return slide, result, bad


# ---------------------------------------------------------------------------
# Slide-by-slide generation: fix the structure (storyline), then fill each
# slide in its own focused, parallel call. Smaller per-call schema surface →
# far fewer mistakes; failing slides are rebuilt individually.
# ---------------------------------------------------------------------------
def _resolve_storyline(ag, prompt, storyline, language, audience):
    items = None
    if isinstance(storyline, list):
        items = storyline
    elif isinstance(storyline, dict):
        items = storyline.get("storyline") or storyline.get("slides")
        language = language or storyline.get("language")
        audience = audience or storyline.get("audience")
    if not items:
        user = f"Brief: {prompt}\n"
        if language:
            user += f"language: {language}\n"
        if audience:
            user += f"audience: {audience}\n"
        s = ag.call_claude(ag.SYSTEM_STORYLINE, user, fast=True)
        items = (s.get("storyline") if isinstance(s, dict) else s) or []
        if isinstance(s, dict):
            language = language or s.get("language")
            audience = audience or s.get("audience")
    norm = [{"intent": it.get("intent", "explain"), "title": it.get("title", ""),
             "note": it.get("note", "")}
            for it in items if isinstance(it, dict)]
    return norm, language, audience


def _fill_one(ag, item, prompt, language, audience, fix=None):
    user = ("Produce EXACTLY ONE slide for this outline entry.\n"
            f"intent: {item['intent']}\n"
            f"action title (use it; you may refine wording): {item['title']}\n"
            f"what it should contain: {item.get('note', '')}\n"
            f"deck brief, for context: {prompt}\n")
    if language:
        user += f"language: {language}\n"
    if audience:
        user += f"audience: {audience}\n"
    if fix:
        user += f"\nYour previous version of THIS slide was rejected:\n{fix}\nReturn a corrected slide."
    # fast model on the first pass; escalate to the strong model on repair.
    return ag.call_claude(ag.SYSTEM_GENERATE_SLIDE, user, fast=(fix is None))


def _fill_all(ag, items, prompt, language, audience, fixes=None, indices=None):
    """Fill slides in parallel. indices (0-based) limits which slides to fill —
    used by repair to rebuild ONLY the failing ones. Others stay None."""
    import concurrent.futures as cf
    fixes = fixes or {}
    todo = list(indices) if indices is not None else list(range(len(items)))
    out: list = [None] * len(items)
    if not todo:
        return out
    with cf.ThreadPoolExecutor(max_workers=min(6, len(todo))) as ex:
        futs = {ex.submit(_fill_one, ag, items[i], prompt, language, audience, fixes.get(i)): i
                for i in todo}
        for f in cf.as_completed(futs):
            out[futs[f]] = f.result()
    return out


def _failing_indices(result, deck_spec) -> dict:
    """Map 0-based SPEC slide index -> feedback string. Gate hard fails, blank
    slides, and missing visuals carry RENDERED indices, so they are only used
    when rendered count == spec count (no splits/drops shifting positions).
    Compiler-dropped slides carry SPEC indices and are always usable."""
    gate = result["gate_result"]
    fb: dict[int, list[str]] = {}
    aligned = result["slide_count"] == len(deck_spec.get("slides") or [])
    if aligned:
        for name, chk in (gate.get("checks") or {}).items():
            if chk.get("severity") == "hard" and not chk.get("passed"):
                for item in chk.get("fail_items", []):
                    m = re.search(r"slide (\d+)", item)
                    if m:
                        fb.setdefault(int(m.group(1)) - 1, []).append(f"[{name}] {item}")
        a = result["_analysis"]
        for i in _empty_slides(a):
            fb.setdefault(i - 1, []).append(
                "[blank] rendered EMPTY — the content keys were wrong for the intent; "
                "rebuild with a valid structure (columns_layout for multi-column).")
        for i in _shortfall_slides(deck_spec, a):
            fb.setdefault(i - 1, []).append(
                "[missing visual] a chart/table you specified did NOT render — the data "
                'shape was wrong. Use chart {type, categories, series:[{name, values}]}.')
    for i, msg in result.get("_compile_errors") or []:
        fb.setdefault(i - 1, []).append(
            f"[dropped] the slide failed to compile and was DROPPED: {msg}. "
            "Fix the intent/content structure and return a valid slide.")
    return {i: "\n".join(v) for i, v in fb.items() if i >= 0}


def _generate_deck(ag, prompt, storyline, language, audience):
    """Two-stage deck build: storyline → parallel per-slide content → render.
    On failure, rebuild only the failing slides (when indices are mappable) or
    fall back to one whole-deck repair."""
    jobs.progress("resolving the storyline")
    items, language, audience = _resolve_storyline(ag, prompt, storyline, language, audience)
    if not items:                                    # storyline empty → single shot
        jobs.progress("drafting the whole deck in one pass")
        spec = ag.call_claude(ag.SYSTEM_GENERATE, f"Brief: {prompt}\n")
        return spec, _render(spec), 1

    jobs.progress(f"writing {len(items)} slides in parallel")
    slides = [s for s in _fill_all(ag, items, prompt, language, audience) if isinstance(s, dict)]
    deck: dict[str, Any] = {"slides": slides}
    if language:
        deck["language"] = language
    if audience:
        deck["audience"] = audience

    # Pre-render intent check: a slide with an unmappable intent would fail the
    # whole render — rebuild just those slides with explicit feedback instead.
    _, unknown = _normalize_intents(deck)
    if unknown:
        jobs.progress(f"fixing invalid intents on {len(unknown)} slide(s)")
        fixes = {i - 1: (f"intent {name!r} does not exist. Choose one of: "
                         + ", ".join(sorted(set(_INTENTS) | set(_MACROS))))
                 for i, name in unknown if 0 <= i - 1 < len(slides)}
        fixed = _fill_all(ag, items, prompt, language, audience,
                          fixes=fixes, indices=list(fixes))
        for i in fixes:
            if isinstance(fixed[i], dict):
                slides[i] = fixed[i]
        deck["slides"] = slides

    jobs.progress("rendering the deck")
    result = _render(deck)
    attempts = 1

    fails = {i: v for i, v in _failing_indices(result, deck).items() if 0 <= i < len(slides)}
    if fails:                                        # indices mappable → surgical repair
        jobs.progress(f"repairing {len(fails)} slide(s)")
        print(f"[generate] per-slide repair of slides {sorted(k + 1 for k in fails)}", flush=True)
        fixed = _fill_all(ag, items, prompt, language, audience,
                          fixes=fails, indices=list(fails))
        for i in fails:
            if isinstance(fixed[i], dict):
                slides[i] = fixed[i]
        deck["slides"] = slides
        jobs.progress("re-rendering the repaired deck")
        result = _render(deck)
        attempts = 2
    elif (result["slide_count"] != len(slides)
          and (not result["gate_result"]["passed"]
               or _bad_slides(deck, result["_analysis"], result.get("_compile_errors")))):
        # overflow split shifted indices and something is wrong → whole-deck repair
        jobs.progress("deck rejected — one whole-deck repair pass (strong model)")
        gate = result["gate_result"]
        bad = _bad_slides(deck, result["_analysis"], result.get("_compile_errors"))
        fb = _gate_feedback(gate)
        if bad:
            fb += "\nBLANK / MISSING-VISUAL SLIDES: " + ", ".join(map(str, bad))
        repair = (f"Brief: {prompt}\n" + (f"language: {language}\n" if language else "")
                  + "\nYou previously produced this deck:\n"
                  + json.dumps(deck, ensure_ascii=False)
                  + "\n\nIt was REJECTED. Fix EVERY item and return the COMPLETE "
                  "corrected deck_spec:\n" + fb)
        deck = ag.call_claude(ag.SYSTEM_GENERATE, repair)
        jobs.progress("re-rendering the repaired deck")
        result = _render(deck)
        attempts = 2
    if result["slide_count"] == 0:
        raise ValueError("the deck rendered with ZERO slides — generation failed; retry")
    return deck, result, attempts


class _BadRequest(ValueError):
    """User-fixable request problem → 400, not 502."""


# ---------------------------------------------------------------------------
# shape_ops validation — the model plans, but the server enforces the physics:
# whitelisted ops, required fields, coordinates clamped to the slide, sane
# sizes and colors. Invalid ops are dropped (and counted), never "fixed" into
# something the model didn't say.
# ---------------------------------------------------------------------------
_SLIDE_W, _SLIDE_H = 13.334, 7.5
_HEX = re.compile(r"^[0-9a-fA-F]{6}$")
_OP_SPECS: dict[str, set[str]] = {          # op -> required fields
    "move": {"id"}, "fill": {"id", "color"}, "no_fill": {"id"},
    "line": {"id"}, "no_line": {"id"}, "font": {"id"}, "text": {"id", "text"},
    "delete": {"id"},
    "add_textbox": {"text", "left", "top", "width", "height"},
    "add_shape": {"shape", "left", "top", "width", "height"},
    "group": {"ids"},
}


def _validate_ops(ops: Any) -> tuple[list[dict], int]:
    """Return (valid_ops, dropped_count)."""
    if not isinstance(ops, list):
        return [], 0
    out: list[dict] = []
    dropped = 0
    for op in ops[:60]:                                       # hard cap
        if not (isinstance(op, dict) and op.get("op") in _OP_SPECS
                and _OP_SPECS[op["op"]] <= set(op)):
            dropped += 1
            continue
        o = dict(op)
        ok = True
        for k in ("left", "top", "width", "height"):
            if k in o:
                try:
                    v = float(o[k])
                except (TypeError, ValueError):
                    ok = False
                    break
                lim = _SLIDE_W if k in ("left", "width") else _SLIDE_H
                o[k] = round(min(max(v, 0.0 if k in ("left", "top") else 0.05),
                                 lim), 3)
        for k in ("color", "fill", "line_color", "font_color"):
            if k in o and not (isinstance(o[k], str) and _HEX.match(o[k].lstrip("#"))):
                del o[k]
            elif k in o:
                o[k] = o[k].lstrip("#").upper()
        if "size_pt" in o:
            try:
                o["size_pt"] = min(max(float(o["size_pt"]), 6.0), 96.0)
            except (TypeError, ValueError):
                del o["size_pt"]
        if "weight_pt" in o:
            try:
                o["weight_pt"] = min(max(float(o["weight_pt"]), 0.25), 12.0)
            except (TypeError, ValueError):
                del o["weight_pt"]
        if o["op"] == "add_shape" and o.get("shape") not in (
                "rectangle", "rounded_rectangle", "oval", "line"):
            ok = False
        if o["op"] == "group" and not (isinstance(o.get("ids"), list)
                                       and len(o["ids"]) >= 2):
            ok = False
        if not (_OP_SPECS[o["op"]] <= set(o)):    # required field cleaned away
            ok = False
        if ok:
            out.append(o)
        else:
            dropped += 1
    return out, dropped


def _agent_impl(ag, req: AgentRequest) -> dict:
    """The actual /agent work — runs synchronously OR inside a background job."""
    if req.mode == "storyline":
        jobs.progress("drafting the storyline")
        user = f"Brief: {req.prompt}\n"
        if req.language:
            user += f"language: {req.language}\n"
        if req.audience:
            user += f"audience: {req.audience}\n"
        return {"storyline": ag.call_claude(ag.SYSTEM_STORYLINE, user, fast=True)}

    if req.mode == "generate":
        spec, result, attempts = _generate_deck(
            ag, req.prompt, req.storyline, req.language, req.audience)
        bad = _bad_slides(spec, result["_analysis"], result.get("_compile_errors"))
        print(f"[generate] done attempts={attempts} "
              f"passed={result['gate_result']['passed']} "
              f"score={result['gate_result']['overall_score']} bad={bad}", flush=True)
        result.pop("_analysis", None)
        result.pop("_compile_errors", None)
        return {"spec": spec, **result, "attempts": attempts, "empty_slides": bad}

    if req.mode == "generate_slide":
        user = f"Request: {req.prompt}\n"
        if req.language:
            user += f"language: {req.language}\n"
        if req.audience:
            user += f"audience: {req.audience}\n"
        slide, result, bad = _one_slide(ag, ag.SYSTEM_GENERATE_SLIDE, user,
                                        req.language, req.audience)
        result.pop("_analysis", None)
        result.pop("_compile_errors", None)
        return {"spec": slide, **result, "empty_slides": bad}

    if req.mode == "edit_slide":
        if not req.command:
            raise _BadRequest("edit_slide requires a command")
        # Ground truth first: what is ACTUALLY on the slide right now (catches
        # manual edits made after generation — changed data, pasted charts).
        inventory = None
        if req.slide_pptx_base64:
            jobs.progress("reading the actual slide contents")
            inventory = _inspect_slide(req.slide_pptx_base64)
        if not req.slide_spec and inventory is None:
            raise _BadRequest("edit_slide requires slide_spec and/or a readable "
                              "slide_pptx_base64 export")
        user = ""
        if req.slide_spec:
            user += f"Stored spec (may be STALE):\n{json.dumps(req.slide_spec, ensure_ascii=False)}\n\n"
        if inventory is not None:
            user += ("Actual slide contents RIGHT NOW (ground truth — preserve "
                     "manual changes):\n"
                     + json.dumps(inventory, ensure_ascii=False) + "\n\n")
        user += f"Edit command: {req.command}\n"
        if req.language:
            user += f"language: {req.language}\n"
        if req.audience:
            user += f"audience: {req.audience}\n"
        slide, result, bad = _one_slide(ag, ag.SYSTEM_EDIT_SLIDE, user,
                                        req.language, req.audience)
        result.pop("_analysis", None)
        result.pop("_compile_errors", None)
        return {"spec": slide, **result, "empty_slides": bad}

    if req.mode == "shape_ops":
        if not req.command or not isinstance(req.slide_state, dict):
            raise _BadRequest("shape_ops requires command and slide_state")
        jobs.progress("planning shape operations")
        user = ("Slide inventory:\n"
                + json.dumps(req.slide_state, ensure_ascii=False)
                + f"\n\nCommand: {req.command}\n")
        out = ag.call_claude(ag.SYSTEM_SHAPE_OPS, user, fast=True)
        if not isinstance(out, dict) or "ops" not in out:
            jobs.progress("plan unusable — retrying with the strong model")
            out = ag.call_claude(ag.SYSTEM_SHAPE_OPS, user
                                 + "\nReturn EXACTLY {\"note\": string, \"ops\": [...]}.",
                                 fast=False)
        ops, dropped = _validate_ops(out.get("ops") if isinstance(out, dict) else None)
        note = (out.get("note") or "") if isinstance(out, dict) else ""
        if dropped:
            note = (note + f" ({dropped} invalid op(s) dropped)").strip()
        return {"ops": ops, "note": note}

    raise _BadRequest(f"unknown mode: {req.mode}")


@app.post("/agent")
def agent(req: AgentRequest):
    """Claude-driven generation/edit collapsed into the renderer service (test
    deployment). Same request/response contract as the Supabase edge function.
    With background=true, returns {job_id} immediately — poll /agent/jobs/{id}.
    This is the reliable path: a full build can outlive the host's HTTP window."""
    try:
        import agent as ag  # lazy import: only /agent needs httpx + the API key
        if req.background:
            return {"job_id": jobs.submit(lambda: _agent_impl(ag, req))}
        return _agent_impl(ag, req)
    except _BadRequest as exc:
        return JSONResponse(status_code=400, content={"error": str(exc)})
    except Exception as exc:                                  # noqa: BLE001
        import traceback
        traceback.print_exc()                                # surfaces in Render logs
        print(f"[/agent] {type(exc).__name__}: {exc}", flush=True)
        return JSONResponse(status_code=502,
                            content={"error": f"{type(exc).__name__}: {exc}"})


@app.get("/agent/jobs/{job_id}")
def agent_job(job_id: str):
    """Poll a background job. Cheap + fast — each poll also keeps the free-tier
    instance awake for the duration of a build."""
    job = jobs.get(job_id)
    if job is None:
        return JSONResponse(status_code=404, content={
            "error": "unknown job — the service likely restarted mid-build; retry"})
    return jobs.public_view(job)
