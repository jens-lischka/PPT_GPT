# Phase 1 (generate) & Phase 2 (slide chat) — build notes

## Stability & quality update (2026-07-15)

- **Background jobs**: `POST /agent {background:true}` → `{job_id}`; pane polls
  `GET /agent/jobs/{id}` (live stage shown in the status line). No HTTP request
  outlives the host proxy window anymore — kills the 502/timeout class. Polls
  double as keep-warm. In-memory store (`renderer/jobs.py`), single instance.
- **Silent slide drops fixed**: near-miss intents (`quote`, `show_trend`, …)
  are aliased to registry names; still-unknown intents fail loudly pre-render;
  compiler "Slide N failed to compile" errors are captured and fed into the
  per-slide repair loop instead of vanishing. Zero-slide decks are rejected.
- **Edit reads the LIVE slide first**: the pane exports the selected slide
  (`exportAsBase64`, PowerPointApi 1.8; graceful fallback) and the backend
  parses it (texts, tables, charts incl. series values) — manual changes made
  after generation survive edits; untracked/foreign slides become editable.
- **Design quality**: DESIGN SELECTION rules added to the system prompts
  (visual-intent variety, KPI/stat cards, conclusions on data slides, sticker)
  — closes the gap vs. the custom GPT's card-rich output. A plain-string
  `insight` is now normalized so the colored callout card always renders.
- **Office ops hardened**: phased insert (snapshot → insert → idempotent
  tag/verify → delete-last), one retry on transient errors, tag write-back
  verification; deck/spec count mismatch is surfaced.
- **Transport**: shared httpx client (keep-alive) for Claude calls; pane fetch
  timeouts + one submit retry; keep-warm fires on settings save.

Phase 0 (spike) passed. This delivers the MVP generation loop and the
selection-aware slide-chat loop.

## What was built

**Renderer** (`renderer/`) — unchanged runtime v2.6.18. Both endpoints
re-verified against generate-style and edit-style payloads (see below).

**Edge function** (`supabase/functions/generate/index.ts`) — full rewrite,
four modes:
- `storyline` — Claude drafts an outline (no render) for the pane to confirm.
- `generate` — Claude drafts a full `deck_spec` → `/render` → optional persist.
- `generate_slide` — one slide → `/render-slide`.
- `edit_slide` — Claude edits a slide spec → `/render-slide` → optional revision.

Model `claude-opus-4-8` with adaptive thinking (per Anthropic guidance).
System prompt is distilled from the runtime itself in
`supabase/functions/_shared/system_prompt.ts` — all 28 intents, the
content-key traps, the waterfall explicit-end rule, and the voice/gate rules.
The gate result passes through untouched; the pane blocks insertion when
`gate.passed` is false (non-overridable invariant).

**Task pane** (`addin/src/app.html`, hosted at
`https://jens-lischka.github.io/PPT_GPT/app.html`, manifest
`addin/manifest.app.xml`):
- Phase 1: brief → **Propose storyline** → confirm → **Build deck** (shows gate;
  blocks insert on fail) → **Insert deck** → writes `OW_SLIDE_UUID` tags on every
  inserted slide.
- Phase 2: **Apply edit** on the selected slide = regenerate-and-replace (insert
  after old, tag new with the same UUID, delete old); **Create slide here** after
  the selection; **Which slide is selected?** reads the UUID tag.
- Spec/UUID map is cached in `localStorage`, so slide-chat works even before
  Supabase persistence is wired (persistence is best-effort until auth exists).

## Verified here (headless)

- Runtime `pytest -q`: **174 passed, 1 skipped**.
- Edge function + system prompt: `tsc --noEmit` clean. Pane JS: `node --check` clean.
- Render contract via the FastAPI `/render` and `/render-slide` code paths:
  - 6-slide de-DE generate-style deck (cover, contents, explain, table,
    waterfall, summary) → gate **PASS** (score 90).
  - single-slide table edit → gate **PASS** (score 100).

## Needs your infrastructure to run end-to-end (cannot run in the build env)

1. **Deploy the renderer** (Cloud Run):
   ```bash
   cd renderer
   gcloud run deploy ow-renderer --source . --region europe-west1 \
     --allow-unauthenticated --set-env-vars ALLOWED_ORIGINS=https://jens-lischka.github.io
   ```
   Note the service URL.
2. **Deploy the edge function**:
   ```bash
   supabase functions deploy generate
   supabase secrets set ANTHROPIC_API_KEY=sk-ant-… RENDERER_URL=<cloud-run-url> \
     ALLOWED_ORIGINS=https://jens-lischka.github.io
   ```
   Apply `supabase/migrations/0001_init.sql` for persistence. Persistence stays
   off until a user id is available (Azure AD via Supabase = Phase 3); until
   then the pane runs stateless with local spec caching.
3. **Load the pane**: sideload `addin/manifest.app.xml` (PowerPoint on the web →
   Insert → Add-ins → Upload My Add-in). Open **⚙︎ Settings**, paste the edge
   function URL (`https://YOURPROJECT.supabase.co/functions/v1/generate`) and the
   Supabase anon key, set language/audience, Save.

## Generation pipeline (renderer /agent)

Two ideas from the original GPT design, implemented in the wrapper (runtime
`src/` untouched):

1. **Convenience macros** (`_expand_macros` in `main.py`). The model emits an
   EASY shape and Python assembles the error-prone geometry. Today:
   `columns_layout` → a `compose` grid (N columns, each with heading + one
   visual (chart/table/kpi/image) + text). The model never writes grid
   mechanics, which was the cause of blank multi-column slides. Applied inside
   `_render`, so every path benefits.
2. **Slide-by-slide content pass** (`_generate_deck`). Instead of one giant
   call for the whole deck, generation is: storyline (fixes intent + action
   title per slide) → **one focused, parallel call per slide** → assemble →
   render + gate. Smaller per-call schema surface = far fewer mistakes.
   Failing slides (gate hard-fails, which name "slide N", plus blank-slide
   detection) are rebuilt **individually and in parallel**; a whole-deck repair
   is the fallback when overflow-splitting makes indices unmappable.

Belt-and-suspenders: `_empty_slides` flags any slide that rendered with no body
content (the empty-canvas gap the gate misses) and forces a rebuild;
`call_claude` retries 429/500/529 with backoff (parallel calls make these more
likely).

## Known limits (honest)

- The full chain (Claude → edge → renderer → PowerPoint) is untested until the
  two services are deployed and the pane is sideloaded — those need your keys
  and real M365.
- `edit_slide` changes intent/geometry via the renderer; a plain-text fast path
  (title typo without a server roundtrip) is noted in the plan but not yet in
  the pane — add if it proves worth the code.
- Adopting foreign (untracked) slides is Phase 3.
