# Phase 1 (generate) & Phase 2 (slide chat) — build notes

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

## Known limits (honest)

- The full chain (Claude → edge → renderer → PowerPoint) is untested until the
  two services are deployed and the pane is sideloaded — those need your keys
  and real M365.
- `edit_slide` changes intent/geometry via the renderer; a plain-text fast path
  (title typo without a server roundtrip) is noted in the plan but not yet in
  the pane — add if it proves worth the code.
- Adopting foreign (untracked) slides is Phase 3.
