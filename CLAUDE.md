# CLAUDE.md — OW Deck Builder: PowerPoint task-pane add-in

You are continuing a project that began as a Custom GPT ("OW Presentation GPT")
and is now being rebuilt as an Office.js task pane inside PowerPoint. This file
is the distilled context of that history. Read it fully before writing code.
`ARCHITECTURE.md` (repo root) describes the original GPT system in depth.

## What this project is

A task pane in PowerPoint where the user:
1. **Generates** an OW-branded deck from a prompt (storyline → confirm → build).
2. **Talks to the selected slide**: "change the data for this chart", "swap the
   two elements", "add an image", "move the chart to the left column".
3. **Creates new slides in place**: "create a four column slide with a heading,
   a paragraph, and a chart in each column" — with dummy or user-provided data.

## The one architectural fact that shapes everything

**Office.js cannot create or edit charts.** No chart API, no chartEx (our
treemaps), no freeform creation (our icons are native freeforms). Therefore:
**all rendering happens server-side in our existing Python runtime, and slides
enter PowerPoint via `presentation.insertSlidesFromBase64()`** (PowerPointApi
1.2+). Editing = regenerate-and-replace: re-render the one slide server-side,
insert at position, delete the old one. Never attempt to rebuild rendering in JS.

A small direct-manipulation fast path is allowed ONLY for plain text edits
(fix a typo in a title) — anything touching charts, icons, tables, or grid
geometry goes through the renderer.

## Repo layout

- `renderer/` — FastAPI wrapper around the UNCHANGED runtime.
  - `main.py` — `/render` (full deck), `/render-slide` (one slide), `/healthz`.
    Both endpoints tested green against the golden briefs.
  - `src/` — byte-identical copy of the GPT package's runtime
    **v2.6.18**. NEVER fork it here. To upgrade: replace wholesale from the
    package zip, update this pin, re-run `testdata/` briefs.
  - `ow_default.pptx` — the OW template (18 layouts). The renderer builds on
    this; the user should ALSO start their presentation from this template so
    `useDestinationTheme` maps layouts cleanly (spike question 2).
  - `Dockerfile` — targets Cloud Run (`$PORT`), also runs locally.
- `supabase/migrations/0001_init.sql` — decks / slides / revisions + RLS.
  Spec fragments live in Postgres, NOT in slide tags. Tags carry only a
  `slide_uuid` (tags have size limits; Postgres gives history + multi-user).
  `revisions` powers spec-level undo — PowerPoint-side undo can't be atomic
  for replace operations (insert + delete = 2 undo steps).
- `supabase/functions/generate/index.ts` — edge function skeleton
  (Deno). Orchestrates Claude API → renderer. The SYSTEM_PROMPT is a
  placeholder — port the distilled GPT instructions in Phase 1 (see below).
- `addin/` — task pane.
  - `manifest.xml` — sideloadable; requirement set kept at 1.1 minimum,
    feature-detect at runtime (`isSetSupported("PowerPointApi", "1.5")`).
  - `src/taskpane.html` — Phase-0 spike pane: insert test deck, render a
    golden brief via local renderer, log selection ids + tags.
  - `src/assets/spike_deck.b64` — pre-rendered 5-slide fidelity deck
    (chartEx treemap, icon freeforms, native table, waterfall, de-DE locale),
    gate score 100. This is the spike payload.
- `testdata/` — the three golden briefs from the GPT package's eval harness.

## PHASE 0 — the spike (do this first, it can kill the whole plan)

Two questions, both need real PowerPoint (Windows or Mac, current M365):

1. **Fidelity:** sideload the manifest, click "Insert test deck", visually
   inspect all 5 slides. Treemap renders? Icons crisp? Table styled? Waterfall
   bars + connectors? Numbers show `1.000`-style (de-DE LCID)?
2. **Master hygiene:** click insert 3×, then View → Slide Master. Count
   masters/layouts. If every insert duplicates the master set, we need a
   mitigation (start-from-template discipline, or post-insert cleanup, or
   `useDestinationTheme` tuning) BEFORE building more.

Sideload: `cd addin && npm install && npm run sideload` (or manual sideload of
`manifest.xml`; serve `src/` over https://localhost:3000 — Office requires
https even for localhost; `office-addin-dev-certs` handles the cert).

If fidelity fails on any element → STOP, report exactly what broke visually,
and we reassess (fallback options: image-based slide export for broken
elements, or dropping the affected intent from the pane's repertoire).

## Phase 1 — MVP generation

- Wire the edge function: port the system prompt. Sources, in the GPT package
  (`ow_presentation_gpt_v2_6_21_*.zip`): `00_gpt_instructions.md` (workflow +
  rules), `knowledge/01_readme_deck_system.md` (intent schema — canonical),
  `knowledge/13_intent_modes.md` (when to use which intent),
  `knowledge/60_voice_system.md` (voice rules). Distill; don't paste 15k
  tokens into a system prompt. The GPT's 7-phase flow compresses to:
  storyline proposal → user confirms in the pane → spec → render → insert.
- Persist deck + slides + revisions in Supabase on every successful render.
- Write `slide_uuid` into each inserted slide's tags
  (`slide.tags.add("OW_SLIDE_UUID", uuid)`) right after insert.
- Show the gate result in the pane. `passed=false` blocks the insert button
  (the gate is non-overridable — that invariant carries over from the GPT).

## Phase 2 — slide chat

- `getSelectedSlides()` → read `OW_SLIDE_UUID` tag → fetch spec from Supabase
  → send command + spec to edge function (`mode: "edit_slide"`) → render-slide
  → insert after the old slide → delete the old slide → update tag on the new
  slide → write a `revisions` row (`source: 'chat_edit'`, keep the prompt).
- Deck-level fields matter: pass `language` (and `audience`) from the deck row
  into `/render-slide`, or charts lose their locale.
- "Create a slide here": same flow, `mode: "generate"` with a 1-slide ask,
  insert after selection.
- The fast path: plain-text title/body edits via Office.js `TextRange` without
  a server roundtrip — offer it only when the command clearly touches text.

## Phase 3 — hardening

- Adopt foreign slides: `slide.exportAsBase64()` → renderer parses with
  python-pptx → propose a spec ("adopt this slide?").
- "Detach from spec" toggle per slide (`slides.detached` column exists).
- Manifest polish, hosting (Cloud Run for renderer; Supabase hosting for the
  pane or any static host), Azure AD auth via Supabase.

## Runtime knowledge you must not violate (from the GPT project)

- The deterministic gate (`gate.py`) runs on every render. Eight checks:
  overflow, action-titles, voice-lint, sources, density, charts, font tokens,
  content_keys. `passed` is binary and non-overridable.
- **content_keys check**: intents read EXACT slot names. Common traps we hit:
  `show_timeline` wants `milestones` (not `items`), `show_contents` wants
  `sections` with `{title, page}` (not `label`), `show_matrix` wants axis
  STRINGS + flat `items` with `row`/`col`. Unknown keys are silently dropped
  at render and flagged as warn_items by the gate.
- `show_waterfall` needs an EXPLICIT end value — `None` for a computed total
  is NOT supported; the canvas strategy fails with only a log warning and the
  slide renders with an empty canvas while the gate still passes (known gap;
  candidate fix in the runtime if it bites again).
- Language: `deck_spec["language"]` (BCP-47) drives chart number locale
  (`[$-407]#,##0` etc. on data labels AND value axis). If omitted, the
  compiler auto-detects from content via stopword heuristic (9 languages,
  en-US fallback). Explicit always wins.
- Diagram labels `"Head: body"` render bold-head/regular-body on pyramid and
  cycle (OW convention, implemented in `_render_diagram_label`).
- Voice: "transformation", "optimise/optimize" are banned terms in the
  consulting profile; action titles must be full sentences (≥4 words), not
  topic labels. The gate enforces both.
- Scatter/XY charts: python-pptx can't set data labels on them
  (`CT_ScatterChart` has no `dLbls`); the runtime skips with a warning.
  Don't offer scatter in the pane's suggestions unless asked.

## Deployment facts (test phase)

- Supabase project region: prefer `eu-central-1`. Cloud backend is approved
  for TESTING by the owner; production privacy posture is an open question —
  the renderer container is the only piece that hard-requires Python hosting,
  keep it swappable to on-prem.
- Renderer on Cloud Run: stateless, scale-to-zero is fine (2–4s cold start
  acceptable). `ALLOWED_ORIGINS` env var gates CORS.
- Secrets: `ANTHROPIC_API_KEY`, `RENDERER_URL` via `supabase secrets set`.

## Working agreements (carry over from the previous collaboration)

- Verify claims by running code, not by assertion. Every fix ships with a
  regression test where the codebase supports it.
- When a change touches the runtime: run its pytest suite (174 tests + 1 skip
  green at v2.6.18) AND the golden briefs before calling it done.
- Be explicit about what is NOT possible (Office.js chart API, atomic undo,
  offline) rather than working around it silently.
- The user speaks German and English; code and docs in English.
