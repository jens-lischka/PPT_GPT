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
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE / "src"))

from runtime import TemplateRuntime, __version__ as RUNTIME_VERSION  # noqa: E402
from compiler import SemanticCompiler                               # noqa: E402
from gate import run_gate                                           # noqa: E402

TEMPLATE = str(HERE / "ow_default.pptx")

app = FastAPI(title="OW deck renderer", version=RUNTIME_VERSION)

# Task pane origin(s). ALLOWED_ORIGINS is a comma-separated list; "*" (the
# test-phase default) allows any origin, which sidesteps origin-mismatch pain
# while the hosting story is in flux. Tighten to the exact pane origin for prod.
_allowed = os.environ.get("ALLOWED_ORIGINS", "*").strip()
_origins = ["*"] if _allowed == "*" else [o.strip() for o in _allowed.split(",") if o.strip()]
app.add_middleware(
    CORSMiddleware,
    allow_origins=_origins,
    allow_methods=["*"],
    allow_headers=["*"],
)


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
    storyline: dict[str, Any] | None = None
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


@app.get("/healthz")
def healthz():
    return {"ok": True, "runtime_version": RUNTIME_VERSION}


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


@app.post("/agent")
def agent(req: AgentRequest):
    """Claude-driven generation/edit collapsed into the renderer service (test
    deployment). Same request/response contract as the Supabase edge function."""
    import agent as ag  # lazy import: only /agent needs httpx + the API key
    try:
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
            return {"spec": spec, **_render(spec)}

        if req.mode == "generate_slide":
            user = f"Request: {req.prompt}\n"
            if req.language:
                user += f"language: {req.language}\n"
            if req.audience:
                user += f"audience: {req.audience}\n"
            slide = ag.call_claude(ag.SYSTEM_GENERATE_SLIDE, user)
            return {"spec": slide, **_render_single(slide, req.language, req.audience)}

        if req.mode == "edit_slide":
            if not req.slide_spec or not req.command:
                raise HTTPException(status_code=400,
                                    detail="edit_slide requires slide_spec and command")
            user = (f"Existing slide spec:\n{req.slide_spec}\n\n"
                    f"Edit command: {req.command}\n")
            if req.language:
                user += f"language: {req.language}\n"
            if req.audience:
                user += f"audience: {req.audience}\n"
            slide = ag.call_claude(ag.SYSTEM_EDIT_SLIDE, user)
            return {"spec": slide, **_render_single(slide, req.language, req.audience)}

        raise HTTPException(status_code=400, detail=f"unknown mode: {req.mode}")
    except HTTPException:
        raise
    except Exception as exc:                                  # noqa: BLE001
        raise HTTPException(status_code=502, detail=f"{type(exc).__name__}: {exc}")
