# OW Presentation GPT — architecture and file map

*Runtime v2.6.18. Written for someone who wants to understand the whole system in one sitting, not a marketing overview.*

## 1. Mental model in one paragraph

The system takes a human's **request for a deck** (either as a natural-language brief in chat or as a hand-authored `deck_spec` JSON) and produces a corporate-branded PowerPoint file. The work is split between two very different environments: **the GPT** does everything language-shaped (understand the brief, pick a storyline, write action titles, choose intents per slide, produce a `deck_spec`), and **a Python runtime** does everything shape-shaped (choose the layout for each intent, place text and charts and icons in inches on the canvas, enforce the OW brand, run a deterministic quality gate). Neither half can do the other's job. The runtime never sees natural language — it consumes structured JSON. The GPT never sees pixels — it hands off to the runtime and reads back a machine-readable `gate_result.json`.

## 2. The two universes, and what "always loaded" means

- **GPT-side context.** Everything the model reads on every session lives here. The instructions field, the always-loaded knowledge markdown files, and any content the user pastes. These pay tokens per session forever, so we care what's in them.
- **Runtime sandbox.** A Python 3 sandbox with `python-pptx`, `lxml`, and `Pillow` available. The runtime zip is uploaded as GPT Knowledge, but *not* as text-in-context — the GPT extracts it to the sandbox and imports from it. The runtime never enters the context window. This is why `src/*.py` weighs 4,600 lines but costs zero session tokens.
- **Maintainer-side.** Docs, evaluation harness, tests, tooling. Not shipped to end users, not loaded by the GPT.

## 3. End-to-end flow of one session

The GPT is instructed to work in seven phases (`00_gpt_instructions.md`). What actually happens:

1. **Phase 1 — Bootstrap.** GPT ensures the runtime zip is in the sandbox and importable. On success, prints `RUNTIME OK v2.6.18` — the audit trail that runtime symbols really loaded, not stale cache. If this fails, everything stops here.
2. **Phase 2 — Ingest.** GPT reads always-loaded knowledge files. Understands the brief: audience, story, constraints.
3. **Phase 3 — Storyline.** GPT proposes a slide-by-slide storyline: intent + one-sentence action title per slide. The human confirms/edits before content is written. This is the point where cheap iteration happens.
4. **Phase 4 — Content pass.** GPT writes each slide's content (titles, bullets, chart data, quotes, KPIs, etc.) into a `deck_spec` JSON. Voice-lint rules apply (no banned terms like "transformation" or "optimise"; action titles must be full sentences).
5. **Phase 5 — Preflight.** GPT loads the on-demand knowledge relevant to what it's about to build (layout selection guide, spec authoring rules, icon index). Validates the spec locally by mentally checking against the schema.
6. **Phase 6 — Build.** GPT runs `python src/generate_deck.py --template ow_default.pptx --deck deck.json --output out.pptx` in the sandbox. Behind the scenes: compiler assigns layouts, strategies render, the gate runs, `gate_result.json` is written next to the `.pptx`.
7. **Phase 7 — Deliver.** GPT reads `gate_result.json`. If `passed: true`, hands the file to the user with a short summary. If `passed: false`, reads the `fail_items`, edits the spec, and loops to Phase 6. The gate is **non-overridable** — the GPT cannot deliver a failing deck by rationalizing.

Auto-language detection runs at the start of Phase 6 inside the compiler: it samples deck text (titles + bodies + insights) with a stopword-frequency heuristic across nine languages, and sets the chart number-format locale accordingly. Explicit `language: "de-DE"` in the spec always wins.

## 4. Package layout

The package that ships (`ow_presentation_gpt_v2_6_21_...zip`) contains three tiers of files:

```
package_root/
├── 00_gpt_instructions.md          ← GPT Instructions field (system prompt)
├── README.md                       ← changelog + maintainer overview (not loaded)
├── SETUP_NEW_GPT.md                ← how to instantiate a new GPT from this package
│
├── knowledge/                      ← files uploaded as GPT Knowledge
│   ├── (see §5 below)              ← 14 markdown files + runtime zip
│   └── deck_system_runtime_clean.zip  ← Python runtime, extracted to sandbox
│
└── _maintainer/                    ← never uploaded; maintainer-only
    ├── 51_portfolio_process.md     ← how to keep the deck portfolio consistent
    ├── 52_knowledge_manifest.md    ← authoritative list of which file loads when
    ├── requirements.txt            ← Python deps for running eval locally
    └── eval/                       ← evaluation harness (see §7)
```

## 5. The GPT-facing files — one line each

Numeric prefixes group files by role: `00` instructions, `01–02` runtime, `1x` standards, `2x` design, `50/60` reference. The manifest at `_maintainer/52_knowledge_manifest.md` is the canonical source of truth for load-timing.

### Instructions field (system prompt)

- **`00_gpt_instructions.md`** — the GPT's operating manual. Establishes the seven-phase workflow, the four hard rules (bootstrap first, never deliver a failing gate, respect the OW voice, own mistakes), and how the GPT should respond to briefs. **Loaded on every turn.**

### Always-loaded Knowledge (paid every session)

- **`01_readme_deck_system.md`** — the runtime's user-facing readme: what intents exist, what content each intent accepts, invariants, known limitations. The canonical schema owner.
- **`02_runtime_bootstrap.md`** — how the GPT locates, extracts, and preflight-checks the runtime zip in the sandbox. Recovery paths when the zip isn't where expected.
- **`11_presentation_standards.md`** — house rules that apply to every OW deck: action titles (full sentences, not labels), one message per slide, evidence discipline, source footnotes on data slides.
- **`12_brand_guidelines.md`** — OW visual identity in prose: navy/cream/gold palette, typography rules, "no bold on diagram labels" — the kind of style constraints the runtime can't enforce mechanically.
- **`13_intent_modes.md`** — descriptions of the 28 intents (what each is *for*, when to pick it, what data it needs). Complements `01_readme`'s schema view.
- **`14_content_structures.md`** — templates for the recurring content patterns (situation-complication-resolution, three-horizon, before/after).
- **`15_storyline_logic.md`** — how to sequence a deck: opening → context → findings → recommendation → next steps. When to insert section dividers.
- **`20_slide_design_capability_map.md`** — high-level map of "what the runtime can visually do" — used when a user asks whether X is possible before we start building.

### On-demand Knowledge (loaded only when relevant)

- **`16_deck_spec_authoring.md`** — the direct-authoring wrapper: how to write a `deck_spec` JSON/YAML by hand, syntax for charts, tables, adornments. Loaded only when the user hands us a spec or asks about specific spec fields.
- **`21_slide_design_selection_guide.md`** — 1,100-line decision tree for picking the right layout per intent. Loaded only in Phase 5 (preflight) when active layout selection is happening.
- **`22_experiences.md`** — hard-won troubleshooting notes: things that behave surprisingly (treemap renders blank in non-PowerPoint previews, chartEx quirks). Loaded when debugging.
- **`23_icon_index.md`** — searchable index of the 500 built-in icons. Loaded when an intent asks for icons and the GPT needs to pick names.
- **`50_design_system.md`** — machine-generated design tokens (font sizes, spacings, colors). Loaded when tuning brand-specific decisions. **Generated from code** by `generate_deck.py --list-tokens`, never hand-edited.
- **`60_voice_system.md`** — voice profiles (`consulting`, `longform`) with per-profile rules (max sentence length, banned words). Loaded when writing copy.

### The runtime zip

- **`deck_system_runtime_clean.zip`** — Python runtime, extracted to the sandbox at Phase 1 and imported from there. Never text-in-context.

## 6. Runtime internals — what runs when

When Phase 6 fires `generate_deck.py --template ow_default.pptx --deck deck.json --output out.pptx`, this is the pipeline:

```
deck.json  →  compiler.SemanticCompiler  →  runtime.TemplateRuntime  →  gate  →  gate_result.json
    │              │                              │                    │              │
    │              ├─ intent lookup               ├─ shape placement    ├─ overflow    ├─ passed
    │              ├─ layout scoring              ├─ style application  ├─ action-     ├─ fail_items
    │              ├─ strategy dispatch           ├─ chart / icon /     │  titles      └─ warn_items
    │              └─ language detection          │  table drawers      ├─ voice-lint
    │                                             └─ save .pptx         ├─ sources
    │                                                                   └─ density
    └─ your JSON (audience, language, slides[])
```

### `src/*.py` — role of each module

- **`generate_deck.py`** — CLI entry point. Argument parsing, load spec, instantiate `TemplateRuntime`, call `SemanticCompiler.build()`, save the `.pptx`, invoke the gate, exit with a status code. Also owns the `--list-tokens`, `--check-contrast`, `--lint-deck`, `--list-voice`, `--diff-tokens` maintenance sub-commands.
- **`compiler.py`** — semantic layer. Owns the 28 intents (`INTENTS` dict), the 32 strategies (`STRATEGIES` dict), layout scoring, capacity-based split logic (when a slide overflows, split it into two), audience-conditional composition routing, language detection wiring. **The place where "meaning → shapes" happens.**
- **`runtime.py`** — rendering primitives. `TemplateRuntime` holds the template Presentation and every low-level drawer: `draw_process_chevrons`, `draw_pyramid`, `draw_cycle`, `draw_org_chart`, `draw_treemap`, `draw_card_grid`, `add_conclusion`, `add_footnote`. Owns design tokens (`Tokens`), brand palette (`PALETTE`, `BRAND_COLORS`), and the OOXML mutations (chart normalization, `endParaRPr` repair, dLbls construction with locale-aware numFmt).
- **`gate.py`** — the deterministic quality gate. Runs eight checks after the deck is saved: overflow (text needs more room than the placeholder gives), action-titles (topic labels vs full sentences), voice-lint (banned words), sources (data slide without footnote), density (sparse content that should have visuals), chart validity, font tokens, and content-keys (author-supplied keys the intent's strategies don't read — silent-drop detector). Emits `gate_result.json` with `passed`, `overall_score` (0-100), `fail_items`, `warn_items`. **The GPT reads this; it does not decide "quality".**
- **`autolayout.py`** — solver used by capacity-based split logic. Given cells with different content, figures out if it fits in the available area or needs to overflow to another slide.
- **`cards.py`** — card primitives: `kpi_card`, `quote_card`, `stat_card`, `icon_card`. Registry + factory, sizes, styling.
- **`components.py`** — reusable shape components (dividers, callouts, chevron geometry helpers).
- **`design_components.py`** — higher-order visual primitives assembled from components (insight card, section number badge).
- **`design_lint.py`** — checks a deck's `design_choices` object against the brand rules. Called from `--lint-deck`.
- **`design_prose.py`** — generates the `50_design_system.md` file from code, so the design tokens documented for the GPT are always in sync with what `runtime.py` actually uses.
- **`icons.py`** — icon library (500 SVGs under `assets/icons/`). Resolves names (`"lightbulb"`, `"chart"`), draws SVG as native PowerPoint freeform shapes (not raster images), handles the OW navy-default coloring.
- **`images.py`** — image intake, placement, and the two image modes (`placeholder` for the "will be replaced" mode, `generate-and-embed` for the real-image mode).
- **`infographics.py`** — the `infographic` intent's non-standard drawers (label-in-shape compositions, badge grids).
- **`textmetrics.py`** — font-size-aware text height estimation. Given "this string in 18pt Noto Sans in a 4.2" wide box", predicts how many inches tall the wrapped text will be. Feeds the overflow check.
- **`voice_lint.py`** — pattern matching against the voice profile's banned-terms and sentence-length rules. Feeds the gate.
- **`voice_profiles.py`** — the profile registry (`consulting`, `longform`) with rules and stopwords.
- **`voice_prose.py`** — generates the `60_voice_system.md` file from the profiles, keeping the doc in sync with the code.
- **`bootstrap.py`** — the runtime's self-preflight: on import, verifies the template exists, the icon library resolves, no schema errors. Referenced from `02_runtime_bootstrap.md`.

### The template

- **`ow_default.pptx`** — the OW-branded PowerPoint template. 18 layouts: `Title Slide`, `Title and Content`, `Title Only`, `2 columns`, `3 columns`, `4 columns`, `2 columns 1/3 split`, `2 columns 2/3 split`, `Big portrait photo on right`, `CV - bio with photo`, `One pager credential`, `Section`, `Title Slide with Picture`, `Contents`, `Confidentiality`, `Qualification`, `Backcover`, `Blank`. The runtime writes into these layouts — colors, fonts, master styles come from the template, not from Python.

### The tests

- **`tests/*.py`** — 174 tests + 1 skip. Cover: gate checks, chart normalization, treemap validity, design system tokens, cover-subtitle exemption, KPI icon rendering, language auto-detection, LCID prefixing on axis and data labels, bold-head/regular-body diagram-label formatting on pyramid and cycle, and the content-keys silent-drop detector. Ship inside the runtime zip so anyone can re-verify after extraction; they cost zero context tokens.

### The examples

- **`examples/semantic_deck.json`** — 11 slides exercising overflow-split paths.
- **`examples/visuals_deck.json`** — 18 slides exercising every visual primitive.

## 7. The maintainer tooling

Never shipped to end users. Used to keep the package healthy.

- **`_maintainer/51_portfolio_process.md`** — how to keep the shipped example decks and the golden briefs consistent as the runtime evolves.
- **`_maintainer/52_knowledge_manifest.md`** — canonical list of which knowledge files are always-loaded, on-demand, or reference. Read this before deciding "should this new file be Core?"
- **`_maintainer/requirements.txt`** — Python dependencies for local development and eval.
- **`_maintainer/eval/`** — evaluation harness:
  - **`golden/*.json`** — three frozen briefs (board recommendation, data review, narrative briefing) that stand in for real deck types. Comparison across runs happens on these, so changes are apples-to-apples.
  - **`run_eval.py`** — builds each golden through the real `generate_deck.py` pipeline, reads back `gate_result.json`, writes `report.json` and `report.md`. Deterministic — runs anywhere. **The gate-based scorecard.**
  - **`render.py`** — drives real Microsoft PowerPoint (Windows COM / macOS AppleScript) to export slides as PNGs. Feeds the LLM judge. Only runs where Office is installed.
  - **`llm_judge.py`** — feeds rendered PNGs to Claude via the Anthropic API, scores Content/Design/Coherence 1–5. Complements the deterministic gate with semantic judgment. Needs `ANTHROPIC_API_KEY`.
  - **`report.json` / `report.md`** — the last eval run's scorecard. Committed per release so you can diff quality over time.

## 8. The four hard invariants

If any of these break, the system is broken:

1. **The gate is non-overridable.** `gate_result.json`'s `passed` boolean is binary and machine-generated. The GPT is instructed never to deliver a failing deck. If the gate is wrong (it sometimes is — see the cover-subtitle overflow fix), we fix the gate, not the delivery rule.
2. **Runtime never in context.** The zip lives in the sandbox. The moment any of its content lands in the GPT's context window, we've silently 20× the token cost of a session for no benefit.
3. **50/60 files are generated.** `50_design_system.md` and `60_voice_system.md` come from `generate_deck.py --list-tokens` / `--list-voice`. Never hand-edited. Enforced by `--diff-tokens` in CI.
4. **Explicit language wins over auto-detection.** Auto-detect is a convenience; the human writer's stated language is authoritative.

## 9. Tracing one build

Concrete example — the user says "build me a 20-slide retail-banking growth deck." What actually happens:

1. GPT reads `00_gpt_instructions.md` (context, always) → seven-phase workflow.
2. Phase 1: `python bootstrap.py` in sandbox → prints `RUNTIME OK v2.6.18`.
3. Phase 2: reads `11_presentation_standards.md`, `12_brand_guidelines.md`, `13_intent_modes.md`, `15_storyline_logic.md`.
4. Phase 3: proposes a 20-slide storyline (cover, exec summary, 4 dividers, 14 content slides). Human confirms.
5. Phase 4: writes titles as action sentences, drafts bullets, picks intents. Loads `21_slide_design_selection_guide.md` (Phase 5 on-demand) to check layout choices.
6. Phase 6: `generate_deck.py` fires. Compiler auto-detects language from titles + bodies (`en-US`, 120 stopword hits). For each slide: intent lookup → layout scoring → strategy dispatch. Runtime places shapes, applies theme, sets locale-aware `[$-409]#,##0` on chart axes and data labels. Gate runs. `out.pptx` and `gate_result.json` land in the sandbox.
7. Phase 7: GPT reads `gate_result.json`. Sees `passed: true, overall_score: 96`. Hands the file to the user with a short summary of what the gate did and did not catch (e.g. "quote slide has no source footnote — expected on a customer quote").

Total: 20 seconds of Python, one context-heavy phase (Phase 4), and a machine-readable pass/fail at the end. That's the whole system.

## 10. Where to look when something is wrong

Symptom → file to inspect first:

- Runtime import fails → `02_runtime_bootstrap.md`, then `bootstrap.py`.
- Gate says the deck is bad → `gate.py` for the check that failed, then the corresponding rule in `11_presentation_standards.md` or `12_brand_guidelines.md`.
- A slide renders with visible text missing where the author put content → `gate.py:_check_content_keys` and the resulting `[content_keys] slide N …` warn in `gate_result.json` — the author almost certainly passed a key the intent's strategies don't read (e.g. `items` on `show_timeline`, which wants `milestones`).
- Layout looks wrong → the intent in `compiler.py:INTENTS`, then the strategy it dispatched to.
- A shape isn't where you expect → `runtime.py`, specifically the drawer that owns that shape family.
- A chart renders with the wrong locale → `_apply_chart_axes` and `_build_ser_dLbls` in `runtime.py`; also confirm `deck_spec["language"]` or auto-detection log line.
- Pyramid or cycle diagram labels aren't formatted the way you expected → `_render_diagram_label` in `runtime.py` splits on `": "` for the bold-head/regular-body convention; label without the colon renders as a single regular-weight run.
- An icon doesn't appear → `icons.py:resolve` (name resolution), then the strategy that was supposed to place it (KPI / column / row).
- Design token drift → run `generate_deck.py --diff-tokens knowledge/50_design_system.md src/runtime.py`; if it flags anything, regenerate.

That's the whole system. Everything else is variations on these files.
