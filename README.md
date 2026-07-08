# OW Deck Builder — PowerPoint task-pane add-in (starter)

Continuation of the OW Presentation GPT as an Office.js add-in.
**Start by reading `CLAUDE.md`** — full context, phase plan, and the Phase-0
spike that decides whether the architecture holds.

Quick smoke test of the renderer (no Docker needed):

    cd renderer
    pip install -r requirements.txt -r runtime_requirements.txt
    uvicorn main:app --port 8123
    # then: curl localhost:8123/healthz

Repo map: `renderer/` (FastAPI + unchanged runtime v2.6.18) ·
`supabase/` (schema + edge function) · `addin/` (manifest + spike pane) ·
`testdata/` (golden briefs) · `ARCHITECTURE.md` (the original GPT system).
