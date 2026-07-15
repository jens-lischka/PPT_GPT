# Corporate-readiness review — OW Deck Builder (branch claude/phase-0-spike-vapn73)

Reviewed: full stack — task pane (`addin/src/app.html` = `docs/app.html`), renderer service (`renderer/main.py`, `agent.py`), runtime v2.6.18 (`renderer/src/`, 11.7k LOC, untouched), Supabase path (`supabase/`), manifests, deploy config. Verified by execution where the sandbox allows: golden briefs render green (gate pass, scores 92–98), render path timed.

## Verdict

The architecture is sound and the prototype is further along than "test" suggests: server-side rendering with a deterministic gate, insert-via-base64, regenerate-and-replace editing, two-tier model use (Haiku happy path, Opus repair), per-slide parallel generation with surgical repair. **It is not corporate-ready.** The gaps are not in the rendering core (which is fast and tested) but in auth, privacy, transport reliability, and one real correctness bug found during this review.

## How and when the Python code runs (call flow + measured cost)

```
Pane (browser, GitHub Pages)
  └─ POST /agent (FastAPI on Render.com free tier, 1 instance, 1 worker)
       ├─ mode=storyline:      1 Claude call (Haiku, no thinking)         ~2–5 s
       ├─ mode=generate:       storyline (if needed)
       │                       → N parallel slide drafts (Haiku, ≤6 threads)
       │                       → _render()                                ~0.1 s
       │                       → failing slides → parallel repair (Opus + thinking)
       │                       → _render() again                          ~0.1 s
       ├─ mode=generate_slide / edit_slide: 1 Haiku call → _render_single
       │                       → on reject: 1 Opus call → render again
       └─ every _render(): load ow_default.pptx (10 ms) → SemanticCompiler.build
                           (20 ms) → save (50 ms) → gate (30 ms) → _analyze (1 parse)
```

Measured in this review: **a full deck render + gate costs ~0.1 s**. Claude calls are >95 % of wall time; an Opus repair pass with thinking can run 1–3 minutes. Free-tier cold start adds 30–50 s. **Conclusion: the Python renderer needs no optimization — all stability and latency work belongs in orchestration, transport, and hosting.**

## Critical — blockers for corporate use

**C1. `/agent` is unauthenticated and CORS is `*`.**
Anyone who finds the Render URL (it is discoverable: public repo, public Pages site) can make unlimited Opus calls on your Anthropic key. There is no rate limit, no quota, no origin lock. Fix: require a shared bearer token on `/agent` (the pane already has an auth-token field; the server just never checks one) and set `ALLOWED_ORIGINS=https://jens-lischka.github.io`.

**C2. Silently dropped slides corrupt edit tracking (found live in this review).**
When Claude emits a near-miss intent (`"quote"` instead of `"show_quote"`, `"show_trend"` instead of `"show_trend_with_key_message"` — both occurred with the shipped test briefs), the compiler logs an ERROR and skips the slide. The deck renders with fewer slides, the gate can still pass, and the pane maps `uuid[i] ↔ spec[i]` by index — so every slide after the dropped one is tracked against the wrong spec. Later edits then "randomly" change the wrong content. This is very likely one of your observed instabilities.
Status: fix drafted in the working copy (intent alias table + pre-render validation + a log-capture that turns drops into repairable feedback + zero-slide guard). Needs final wiring + verification.

**C3. Long builds die at the proxy as opaque 502s.**
A generate with an Opus repair pass can exceed Render's request window; the pane sees "backend 502" or a non-JSON error with all work lost. Fix drafted: background job pattern (`POST /agent {background:true}` → `job_id`; pane polls `GET /agent/jobs/{id}`, which also doubles as keep-warm). `jobs.py` written; endpoint + pane polling not yet wired.

**C4. Confidential deck content reaches third-party logs.**
The compiler logs slide titles and structure at INFO on every render; tracebacks can embed spec content; Render retains service logs. Deck content also transits Render (US company, Frankfurt region) and Anthropic. For corporate use you need: a `LOG_CONTENT=off` switch that silences content-bearing log lines, a DPA position on Render + Anthropic (API data is not used for training by default; retention terms should be confirmed), and sign-off that test-phase cloud processing covers real client content — currently only "approved for testing" per CLAUDE.md.

## High — fix before a pilot

- **H1. `/selftest` is public**: leaks key prefix, model, allowed origins, and performs a live Claude call unauthenticated. Put it behind the same token.
- **H2. Spec store is one browser's localStorage.** Clear the cache or switch machines (you use web AND Mac desktop — two separate stores!) and every existing deck becomes un-editable ("slide isn't tracked"). This is your third likely instability source. The Supabase persistence path exists but is unused and has no auth. Short term: accept + surface it clearly; medium term: wire Supabase with Azure AD.
- **H3. Pane fetch has no timeout/retry.** One network blip or cold start fails the click; keep-warm only runs while the pane is open and configured. Job polling (C3) largely solves this; add AbortController timeouts + one retry on submit.
- **H4. No retry on transient Office.js errors** during insert/replace; on failure mid-replace the old slide may already be gone from the intended position. The one-round-trip design minimizes but doesn't eliminate this. Add retry-once + post-insert tag verification.
- **H5. Free-tier hosting is itself an instability**: sleep/cold-start, 512 MB, single instance. A pilot needs Render paid (EU) or Cloud Run (eu-west) — the code is already stateless and fits both.
- **H6. No CI, unpinned deps** (`>=` ranges, no lockfile), and the 174-test suite runs nowhere automatically. The system prompt exists twice (`agent.py` and `system_prompt.ts`) with manual "keep in sync" comments — guaranteed drift. Pick the single-service path as canonical for now and mark the edge function as dormant.

## Medium — worthwhile optimizations (in priority order)

- **M1. Reuse one `httpx.Client`** in `agent.py` — currently every Claude call pays a fresh TLS handshake; with ~10 calls per deck that's free latency.
- **M2. Prompt-cache warm-up**: the 6 parallel Haiku slide calls race the cache write on first use (up to 6× cache-write cost). The sequential storyline call mostly primes it already; making `SYSTEM_GENERATE_SLIDE` share the exact cached prefix guarantees it.
- **M3. Insert-formatting default**: `useDestinationTheme` mangles slides when the user's deck was NOT started from `ow_default.pptx` (layout mapping fails). Detect mismatch (compare slide-master names via Office.js) and auto-fall back to `keepSourceFormatting` with a note, instead of relying on user discipline.
- **M4. Skip**: template-load caching (10 ms), binary transport instead of base64+gzip (~300 KB wire is fine), gate double-parse (30 ms). Not worth the complexity.

## Cost control (currently none)

No quota, no per-deck slide cap, no daily budget. Combined with C1 this is an open wallet. Minimum: auth (C1) + a slide-count cap (e.g. 15) + a simple daily request counter with a hard stop.

## What was already changed in this working copy (paused for your sign-off)

- `renderer/jobs.py` (new): in-memory background-job store with progress stages — not yet exposed via endpoints.
- `renderer/main.py`: intent alias table + pre-render validation (C2), compile-error capture feeding the repair loop (C2), per-slide repair now also fires when slides were dropped, zero-slide guard, progress hooks, `background` request field (not yet acted on).
- All changes are outside `src/` (runtime untouched, as required). Syntax-checked; golden briefs re-verified green.

## Recommended order

1. C1 + H1 — auth token + CORS lock (small, kills the biggest exposure)
2. C3 + H3 — finish job pattern + pane polling (kills the 502/timeout class)
3. C2 + H4 — finish drop-fix + Office retry/verify (kills the mis-tracking class)
4. C4 + H6 — log hygiene, pinned deps, CI note
5. H2/H5 — persistence + hosting decision (pilot gate)
6. M1–M3 — polish
Then: focus on presentation quality/content work with a stable base.
