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
from starlette.requests import Request
from pydantic import BaseModel

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE / "src"))

from runtime import TemplateRuntime, __version__ as RUNTIME_VERSION  # noqa: E402
from compiler import SemanticCompiler                               # noqa: E402
from gate import run_gate                                           # noqa: E402

TEMPLATE = str(HERE / "ow_default.pptx")

app = FastAPI(title="OW deck renderer", version=RUNTIME_VERSION)

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


def _render(deck_spec: dict) -> dict:
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "deck.pptx"
        gate_path = Path(td) / "gate_result.json"
        rt = TemplateRuntime(TEMPLATE)
        SemanticCompiler(rt).build(deck_spec)
        rt.save(str(out))
        gate = run_gate(deck_spec, str(out), TEMPLATE, str(gate_path))
        gate.pop("_path", None)
        return {
            "pptx_base64": base64.b64encode(out.read_bytes()).decode(),
            "gate_result": gate,
            "slide_count": _slide_count(out),
            "runtime_version": RUNTIME_VERSION,
        }


def _slide_count(pptx_path: Path) -> int:
    from pptx import Presentation
    return len(Presentation(str(pptx_path)).slides)


def _empty_slides(pptx_b64: str) -> list[int]:
    """1-based indices of slides that rendered with NO body content — catches
    the empty-canvas gap (a compose/canvas spec the renderer silently dropped
    because of wrong keys) that the deterministic gate does not flag."""
    import io
    from pptx import Presentation
    prs = Presentation(io.BytesIO(base64.b64decode(pptx_b64)))
    empty = []
    for i, s in enumerate(prs.slides, 1):
        n = 0
        for sh in s.shapes:
            center = ((sh.top or 0) + (sh.height or 0) / 2) / 914400
            rich = sh.has_chart or sh.has_table or sh.shape_type == 13
            if rich or (1.4 < center < 6.2):
                n += 1
        if n == 0:
            empty.append(i)
    return empty


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
    slide = ag.call_claude(system, user)
    result = _render_single(slide, language, audience)
    if _empty_slides(result["pptx_base64"]):
        print("[/agent] one-slide came out BLANK — retrying with guidance", flush=True)
        retry = (user + "\n\nYour previous slide rendered COMPLETELY BLANK: the "
                 "content keys were wrong for the intent and got dropped. Use a "
                 "VALID structure for the intent (for multi-column-with-charts use "
                 "a compose grid of cells, never a 'columns' key) and return the "
                 "corrected slide.")
        slide = ag.call_claude(system, retry)
        result = _render_single(slide, language, audience)
    return slide, result


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
            return {"storyline": ag.call_claude(ag.SYSTEM_STORYLINE, user)}

        if req.mode == "generate":
            if req.storyline:
                user = ("Confirmed storyline (build the full deck from it):\n"
                        f"{req.storyline}\n")
            else:
                user = f"Brief: {req.prompt}\n"
            if req.language:
                user += f"language: {req.language}\n"
            if req.audience:
                user += f"audience: {req.audience}\n"
            spec = ag.call_claude(ag.SYSTEM_GENERATE, user)
            result = _render(spec)
            # Self-repair: rebuild while the gate blocks OR a slide came out
            # blank (the empty-canvas gap the gate misses), up to 2 attempts.
            attempts = 1
            while attempts < 2:
                gate = result["gate_result"]
                empty = _empty_slides(result["pptx_base64"])
                print(f"[gate attempt {attempts}] passed={gate['passed']} "
                      f"score={gate['overall_score']} empty_slides={empty}", flush=True)
                if gate["passed"] and not empty:
                    break
                fb = _gate_feedback(gate)
                if empty:
                    fb += ("\nBLANK SLIDES (rendered empty — wrong content keys "
                           "were dropped; rebuild these with a valid structure "
                           "for their intent): slides " + ", ".join(map(str, empty)))
                repair = (user + "\n\nYou previously produced this deck:\n"
                          + json.dumps(spec, ensure_ascii=False)
                          + "\n\nIt was REJECTED. Fix EVERY item below and return the "
                          "COMPLETE corrected deck_spec (keep what already passed):\n" + fb)
                spec = ag.call_claude(ag.SYSTEM_GENERATE, repair)
                result = _render(spec)
                attempts += 1
            return {"spec": spec, **result, "attempts": attempts,
                    "empty_slides": _empty_slides(result["pptx_base64"])}

        if req.mode == "generate_slide":
            user = f"Request: {req.prompt}\n"
            if req.language:
                user += f"language: {req.language}\n"
            if req.audience:
                user += f"audience: {req.audience}\n"
            slide, result = _one_slide(ag, ag.SYSTEM_GENERATE_SLIDE, user,
                                       req.language, req.audience)
            return {"spec": slide, **result,
                    "empty_slides": _empty_slides(result["pptx_base64"])}

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
            slide, result = _one_slide(ag, ag.SYSTEM_EDIT_SLIDE, user,
                                       req.language, req.audience)
            return {"spec": slide, **result,
                    "empty_slides": _empty_slides(result["pptx_base64"])}

        return JSONResponse(status_code=400, content={"error": f"unknown mode: {req.mode}"})
    except Exception as exc:                                  # noqa: BLE001
        import traceback
        traceback.print_exc()                                # surfaces in Render logs
        print(f"[/agent] {type(exc).__name__}: {exc}", flush=True)
        return JSONResponse(status_code=502,
                            content={"error": f"{type(exc).__name__}: {exc}"})
