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
import os
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
from compiler import SemanticCompiler                               # noqa: E402
from gate import run_gate                                           # noqa: E402

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
    """columns_layout → compose grid. Each column becomes a grid column with up
    to three stacked cells on fixed rows so columns align: heading (row 0), a
    visual (row 1), body text (row 2)."""
    cols = content.get("columns") or []
    cells: list[dict] = []
    for c, col in enumerate(cols):
        if not isinstance(col, dict):
            continue
        if col.get("heading"):
            cells.append({"block": "text", "data": {"heading": col["heading"]},
                          "col": c, "row": 0})
        # one visual per column (first match wins)
        if col.get("chart"):
            cells.append({"block": "chart", "data": {"chart": col["chart"]}, "col": c, "row": 1})
        elif col.get("table"):
            cells.append({"block": "table", "data": col["table"], "col": c, "row": 1})
        elif col.get("kpi"):
            cells.append({"block": "kpi", "data": col["kpi"], "col": c, "row": 1})
        elif col.get("image"):
            cells.append({"block": "image", "data": col["image"], "col": c, "row": 1})
        elif col.get("waterfall"):
            cells.append({"block": "waterfall", "data": {"waterfall": col["waterfall"]}, "col": c, "row": 1})
        elif col.get("treemap"):
            cells.append({"block": "treemap", "data": {"treemap": col["treemap"]}, "col": c, "row": 1})
        # body text
        if col.get("bullets"):
            cells.append({"block": "text", "data": {"bullets": col["bullets"]}, "col": c, "row": 2})
        elif col.get("text"):
            txt = col["text"]
            cells.append({"block": "text",
                          "data": {"paragraphs": [txt] if isinstance(txt, str) else txt},
                          "col": c, "row": 2})
    new_content: dict[str, Any] = {"title": content.get("title", "")}
    if content.get("footnote"):
        new_content["footnote"] = content["footnote"]
    new_content["compose"] = {"grid": {"cols": max(1, len(cols)), "cells": cells}}
    return {"intent": "compose", "content": new_content}


_MACROS = {"columns_layout": _macro_columns_layout}


def _render(deck_spec: dict) -> dict:
    deck_spec = _expand_macros(deck_spec)          # convenience macros → real intents
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "deck.pptx"
        gate_path = Path(td) / "gate_result.json"
        rt = TemplateRuntime(TEMPLATE)
        SemanticCompiler(rt).build(deck_spec)
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


def _slide_requests(slide: dict, kind: str) -> bool:
    """Does this slide spec ask for a chart / table? (scan the JSON + intent)."""
    try:
        blob = json.dumps(slide)
    except Exception:                                        # noqa: BLE001
        blob = str(slide)
    intent = slide.get("intent") if isinstance(slide, dict) else None
    if kind == "chart":
        return '"chart"' in blob or intent == "show_trend_with_key_message"
    if kind == "table":
        return '"table"' in blob or intent == "show_data"
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


def _bad_slides(deck_spec: dict, analysis) -> list[int]:
    """Slides that are blank OR missing a requested visual — must not ship."""
    a = _as_analysis(analysis)
    return sorted(set(_empty_slides(a)) | set(_shortfall_slides(deck_spec, a)))


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
        bad = set(_bad_slides({"slides": [slide]}, a))
        inv = a[0] if a else {"charts": 0, "tables": 0}
        low = user.lower()
        if any(w in low for w in ("chart", "graph", "plot")) and inv["charts"] == 0:
            bad.add(1)
        if "table" in low and inv["tables"] == 0:
            bad.add(1)
        return sorted(bad)

    slide = ag.call_claude(system, user, fast=True)     # fast happy path
    result = _render_single(slide, language, audience)
    bad = _bad(slide, result)
    if bad:
        print("[/agent] one-slide blank/missing-visual — retrying (strong model)", flush=True)
        retry = (user + "\n\nYour previous slide was REJECTED: it rendered with "
                 "MISSING content — either an empty canvas, or a chart/table the "
                 "request asked for did not appear (you omitted it or its data shape "
                 "was wrong). Include EVERY element the request asks for. For multi-"
                 "column-with-visuals use the columns_layout intent; for a chart use "
                 '{"chart":{"type":"column|bar|line|pie|area","categories":[…],'
                 '"series":[{"name":"…","values":[…]}]}}. Return the corrected slide.')
        slide = ag.call_claude(system, retry, fast=False)   # escalate to strong
        result = _render_single(slide, language, audience)
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
    """Map 0-based slide index -> feedback string, from the gate's hard fails
    (which name 'slide N'), blank-slide detection, and missing-visual detection."""
    import re
    gate = result["gate_result"]
    fb: dict[int, list[str]] = {}
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
    return {i: "\n".join(v) for i, v in fb.items() if i >= 0}


def _generate_deck(ag, prompt, storyline, language, audience):
    """Two-stage deck build: storyline → parallel per-slide content → render.
    On failure, rebuild only the failing slides (when indices are mappable) or
    fall back to one whole-deck repair."""
    items, language, audience = _resolve_storyline(ag, prompt, storyline, language, audience)
    if not items:                                    # storyline empty → single shot
        spec = ag.call_claude(ag.SYSTEM_GENERATE, f"Brief: {prompt}\n")
        return spec, _render(spec), 1

    slides = [s for s in _fill_all(ag, items, prompt, language, audience) if isinstance(s, dict)]
    deck: dict[str, Any] = {"slides": slides}
    if language:
        deck["language"] = language
    if audience:
        deck["audience"] = audience
    result = _render(deck)
    attempts = 1

    if result["slide_count"] == len(slides):         # no overflow split → per-slide repair
        fails = {i: v for i, v in _failing_indices(result, deck).items() if 0 <= i < len(slides)}
        if fails:
            print(f"[generate] per-slide repair of slides {sorted(k + 1 for k in fails)}", flush=True)
            fixed = _fill_all(ag, items, prompt, language, audience,
                              fixes=fails, indices=list(fails))
            for i in fails:
                if isinstance(fixed[i], dict):
                    slides[i] = fixed[i]
            deck["slides"] = slides
            result = _render(deck)
            attempts = 2
    else:                                            # split happened → whole-deck repair
        gate = result["gate_result"]
        bad = _bad_slides(deck, result["_analysis"])
        if not gate["passed"] or bad:
            fb = _gate_feedback(gate)
            if bad:
                fb += "\nBLANK / MISSING-VISUAL SLIDES: " + ", ".join(map(str, bad))
            repair = (f"Brief: {prompt}\n" + (f"language: {language}\n" if language else "")
                      + "\nYou previously produced this deck:\n"
                      + json.dumps(deck, ensure_ascii=False)
                      + "\n\nIt was REJECTED. Fix EVERY item and return the COMPLETE "
                      "corrected deck_spec:\n" + fb)
            deck = ag.call_claude(ag.SYSTEM_GENERATE, repair)
            result = _render(deck)
            attempts = 2
    return deck, result, attempts


@app.post("/agent")
def agent(req: AgentRequest):
    """Claude-driven generation/edit collapsed into the renderer service (test
    deployment). Same request/response contract as the Supabase edge function."""
    try:
        import agent as ag  # lazy import: only /agent needs httpx + the API key
        if req.mode == "storyline":
            user = f"Brief: {req.prompt}\n"
            if req.language:
                user += f"language: {req.language}\n"
            if req.audience:
                user += f"audience: {req.audience}\n"
            return {"storyline": ag.call_claude(ag.SYSTEM_STORYLINE, user, fast=True)}

        if req.mode == "generate":
            spec, result, attempts = _generate_deck(
                ag, req.prompt, req.storyline, req.language, req.audience)
            bad = _bad_slides(spec, result["_analysis"])
            print(f"[generate] done attempts={attempts} "
                  f"passed={result['gate_result']['passed']} "
                  f"score={result['gate_result']['overall_score']} bad={bad}", flush=True)
            result.pop("_analysis", None)
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
            return {"spec": slide, **result, "empty_slides": bad}

        if req.mode == "edit_slide":
            if not req.slide_spec or not req.command:
                return JSONResponse(status_code=400,
                    content={"error": "edit_slide requires slide_spec and command"})
            user = (f"Existing slide spec:\n{req.slide_spec}\n\n"
                    f"Edit command: {req.command}\n")
            if req.language:
                user += f"language: {req.language}\n"
            if req.audience:
                user += f"audience: {req.audience}\n"
            slide, result, bad = _one_slide(ag, ag.SYSTEM_EDIT_SLIDE, user,
                                             req.language, req.audience)
            result.pop("_analysis", None)
            return {"spec": slide, **result, "empty_slides": bad}

        return JSONResponse(status_code=400, content={"error": f"unknown mode: {req.mode}"})
    except Exception as exc:                                  # noqa: BLE001
        import traceback
        traceback.print_exc()                                # surfaces in Render logs
        print(f"[/agent] {type(exc).__name__}: {exc}", flush=True)
        return JSONResponse(status_code=502,
                            content={"error": f"{type(exc).__name__}: {exc}"})
