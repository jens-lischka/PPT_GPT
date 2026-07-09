# Deploy — test phase (browser-only, no CLI / no admin)

Goal: get the pane working end-to-end with the least setup on a locked-down
machine. We collapse the backend into ONE service — the FastAPI renderer, which
also hosts the Claude `/agent` endpoint — and deploy it from this public repo to
Render.com. The pane is already hosted on GitHub Pages.

Prerequisite: an **Anthropic API key** (`sk-ant-…`) from https://console.anthropic.com.

## 1. Deploy the backend to Render.com (~3 min, all in the browser)

1. Sign in at https://render.com (GitHub login works; free plan is enough).
2. **New +** → **Blueprint**.
3. Connect / pick the repo **`jens-lischka/PPT_GPT`**. Render reads
   `render.yaml` and proposes a service **`ow-renderer`** (Docker, Frankfurt).
4. **Apply**. When prompted for the env var **`ANTHROPIC_API_KEY`**, paste your
   key (it is `sync: false`, so it is never stored in git).
5. Wait for the first build (a few minutes). When it's live, copy the service
   URL, e.g. `https://ow-renderer-xxxx.onrender.com`.
6. Sanity check in the browser: open `…onrender.com/healthz` → you should see
   `{"ok":true,"runtime_version":"2.6.18"}`.

> Free plan note: the service sleeps when idle, so the first request after a
> pause takes ~30–50 s to wake. Fine for testing.

## 2. Point the pane at it

1. In PowerPoint on the web: **Insert → Add-ins → Upload My Add-in** → upload
   `addin/manifest.app.xml` (download it raw from GitHub). The pane
   **OW Deck Builder** opens.
2. In the pane, open **⚙︎ Settings**:
   - **Backend URL**: `https://ow-renderer-xxxx.onrender.com/agent`
   - **Auth bearer token**: leave empty (only the Supabase path needs it).
   - **Language** / **Audience**: e.g. `de-DE` / `board`.
   - **Save**.

## 3. Use it

- **Generate**: type a brief → *Propose storyline* → review → *Build deck*
  (shows the gate; a failing gate blocks insertion) → *Insert deck*.
- **Edit a slide**: select a pane-generated slide → type a command → *Apply edit*
  (regenerate-and-replace). *Create slide here* adds one after the selection.

## CORS

`render.yaml` sets `ALLOWED_ORIGINS=https://jens-lischka.github.io` (the Pages
origin). If you host the pane elsewhere, update that env var in the Render
dashboard.

---

## Alternative: the documented two-service architecture (later)

For production the orchestration belongs in the Supabase edge function
(`supabase/functions/generate/index.ts`, Deno) with the renderer as a separate
Python host, plus Postgres persistence and Azure AD auth (Phase 3):

```bash
# renderer on Cloud Run
cd renderer && gcloud run deploy ow-renderer --source . --region europe-west1 \
  --allow-unauthenticated --set-env-vars ALLOWED_ORIGINS=https://jens-lischka.github.io
# edge function on Supabase
supabase functions deploy generate
supabase secrets set ANTHROPIC_API_KEY=sk-ant-… RENDERER_URL=<cloud-run-url> \
  ALLOWED_ORIGINS=https://jens-lischka.github.io
supabase db push        # applies migrations/0001_init.sql
```

Then set the pane's Backend URL to `https://YOURPROJECT.supabase.co/functions/v1/generate`
and paste the Supabase anon key as the bearer token.
