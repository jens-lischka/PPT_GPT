"""Claude orchestration for the single-service test deployment.

The documented architecture puts this in a Supabase edge function
(supabase/functions/generate/index.ts). For the test phase we collapse to ONE
deployable service: the FastAPI renderer exposes /agent, which calls Claude and
then the in-process render path. Fewer hosts, secrets, and CLIs to stand up.

The system-prompt strings mirror supabase/functions/_shared/system_prompt.ts —
KEEP THEM IN SYNC (the runtime in src/ is the ultimate source of truth for the
intent schema). Pinned to runtime v2.6.18.

Env:
  ANTHROPIC_API_KEY   required for /agent
  ANTHROPIC_MODEL     optional override (default claude-opus-4-8)
"""
from __future__ import annotations

import json
import os
import re
from typing import Any

import httpx

# Two tiers: a FAST model for the happy path (focused single-slide / storyline
# tasks) and a STRONG model we escalate to only when a slide fails and must be
# repaired. Fast-first + escalate-on-failure = low latency without losing the
# reliability the gate/detectors demand. Override via env.
MODEL = os.environ.get("ANTHROPIC_MODEL", "claude-opus-4-8")
MODEL_FAST = os.environ.get("ANTHROPIC_MODEL_FAST", "claude-haiku-4-5-20251001")

INTENT_REFERENCE = """
You author decks for the OW (Oliver Wyman) deck runtime. A deck is JSON:
{ "language"?: BCP-47 (e.g. "de-DE","en-US"), "audience"?: string,
  "title"?: string, "slides": [ { "intent": <name>, "content": {...} }, ... ] }

Every slide is one intent from the registry below. content keys are read by
EXACT name — unknown keys are silently dropped. "title" is required on all
intents except show_quote (uses "quote") and introduce_person (uses "name").

INTENTS (required content keys → shape):
- introduce_topic: title; subtitle?; image?            (cover slide)
- explain: title, body {bullets:[...] | paragraphs:[...]}; conclusion?; footnote?
- summary: title, body {bullets|paragraphs}; conclusion?              (closing)
- section_divider: title; number?                       (section break)
- show_contents: title, sections:[{title, page?, number?}]  (TOC; number auto 01,02…; one per deck)
- show_data: title, table {headers:[...], rows:[[...],...]}; conclusion?; footnote?
- compare_two_options: title, left {bullets|paragraphs}, right {bullets|paragraphs}
- show_columns: title, columns:[{heading, icon?, bullets|paragraphs}]  (2-5 parallel headed columns; ADD an icon keyword to EVERY column — default-on)
- compare: title, compare {pros:[...], cons:[...], pros_heading?, cons_heading?}  (green check / red cross)
- dashboard: title, metrics:[{value, heading, text?, icon?}]  (grid of KPI cards, variadic; ADD an icon keyword to EVERY metric — default-on)
- stat_callout: title, stats:[{value, heading, text}]   (1-4 big-number cards)
- icon_rows: title, rows:[{icon, heading, text}]        (1-5 rows; icon = short keyword)
- show_trend_with_key_message: title, chart {type, ...}, insight {text, title?}  (chart + colored callout card)
- show_waterfall: title, waterfall {categories:[...], values:[...], totals:[indices], orientation?:"vertical"|"horizontal", heading?, subheading?}
- show_treemap: title, treemap {items:[{label, value}], heading?, subheading?}
- show_timeline: title, milestones:[{date|label, title, text?}]
- show_process: title, steps:[{title, text?}]           (horizontal chevrons)
- show_growing_steps: title, steps:[{title, text?}]     (ascending staircase)
- show_pyramid: title, levels:[{label, text?}]          (label "Head: body" → bold head)
- show_cycle: title, items:[...]                        (genuine loops only)
- show_org_chart: title, org_chart {root:{...}, reports:[...]}
- show_matrix: title, matrix {rows?, cols?, x_axis:STRING, y_axis:STRING, highlight_cell?:[r,c], items:[{label, x:0-1, y:0-1, comment?} | {label, row, col}]}  (any size: 2×2 … 5×5)
- infographic: title, infographic {type:"funnel"|"gauge"|"venn"|"heatmap", ...}
- show_quote: quote {text, attribution}                 (pull-quote; NOT for data)
- introduce_person: name; role?; photo?; bio?
- show_graphic_with_text: title, panel {graphic:{kind,...}, text:{...}, side?, split?}
- compose: title, compose {...}                         (escape hatch; only if no preset fits)

CONTENT-KEY TRAPS (these break silently — obey exactly):
- show_timeline wants "milestones", NOT "items".
- show_contents wants "sections" with {title, page}, NOT "label".
- show_matrix wants axis STRINGS (x_axis/y_axis) + flat "items" with row/col.
- show_waterfall REQUIRES an explicit numeric end/total in "values" with its
  index listed in "totals". Never rely on a computed total — omitting it
  renders an empty chart.
- dashboard wants "metrics"; stat_callout wants "stats".

MULTI-COLUMN-WITH-VISUALS (e.g. "N columns, each with a header, a chart and a
paragraph"): this is NOT show_columns (text only). Use the convenience intent
"columns_layout" — do NOT hand-build a compose grid:
{ "intent":"columns_layout", "content":{ "title":"…", "footnote":"Source: …",
  "columns":[
    { "heading":"Region A",
      "chart":{"type":"column","categories":["FY24","FY25"],
               "series":[{"name":"Region A","values":[100,120]}]},
      "text":"One-line takeaway." },
    …  2–5 columns … ] }}
Per column: an optional "heading", ONE visual — "chart" (type column/bar/line/
pie/area), OR "table":{headers,rows}, OR "kpi":{value,label}, OR "image":
{prompt} — and an optional "text" (string) or "bullets":[…]. Python assembles
the geometry, so you never touch grid mechanics.

DO NOT use scatter/XY charts unless explicitly asked (data labels unsupported).

DESIGN SELECTION (ported from the OW design-selection guide — this is what
makes a deck look like OW work, not a text dump; follow it exactly):
Start with the message, not the layout. Decision tree per slide:
- opening/closing/section break → introduce_topic, show_contents,
  section_divider, show_quote
- sequence/process/journey/maturity → show_process, show_growing_steps,
  show_timeline, show_cycle (cycle ONLY for genuine loops)
- comparing options/categories/plans → compare_two_options, show_data
  (options table), show_matrix, compare (pros/cons)
- proving impact with metrics → dashboard / stat_callout (big-number CARDS),
  show_trend_with_key_message (chart + colored insight callout card),
  show_waterfall (bridges), show_treemap (shares of a whole)
- organizing many related topics → show_columns, show_data, explain (numbered)
- roles/governance/team → show_org_chart
- emotional/human context → show_quote, section_divider (large statement)
- too detailed for storytelling → show_data table; split rather than cram
Slide structure by deck section: situation/context → columns, table, KPI
infographic; analysis/evidence → big-number comparisons, matrices, option
tables; recommendation → columns+icons, KPI columns, growing steps, chevrons;
roadmap → chevrons, timeline, org chart; impact/proof → KPI cards, big-number
columns; closing → quote or large statement, NEVER a paragraph list.
Hard rules:
- Big numbers only when they ARE the story; every number gets a short label.
- One clear entry point per slide; hierarchy obvious in 3 seconds; parallel
  structure across columns/steps/options (same pattern, similar length).
- Icons: default-ON for every show_columns column and dashboard metric. ALL
  OR NOTHING per slide: if one icon keyword is missing/unresolvable, none
  render — give EVERY column/metric one. Safe keywords: chart, bar-chart,
  pie-chart, analytics, growth, trend-up, money, revenue, budget, percentage,
  user, users, team, customer, handshake, network, target, goal, strategy,
  compass, roadmap, milestone, rocket, idea, bulb, shield, lock, globe, map,
  building, factory, store, cloud, database, phone, laptop, ai, automation,
  leaf, energy, heart, health, star, award, trophy, quality, check, clock,
  calendar, search, settings, warning, flag, document, folder, mail, truck.
  Diagram intents (process, pyramid, cycle, growing_steps) have NO icon
  slots — never add icons there.
- ADD "conclusion" (one takeaway sentence) to EVERY chart/table/data slide —
  it renders as a distinct bottom takeaway band and is an OW signature.
- "sticker" adds an amber circular highlight — use for the single most
  important number/claim in the deck, at most once or twice.
- KEY SLIDES (the ones carrying the argument: the core-message slide, one per
  supporting argument, the close) must LEAD WITH A PRIMARY VISUAL — a chart,
  KPI cards, a diagram, or a hero stat. A key slide is never text-only.
- Never more than TWO consecutive text-only slides. A deck of only
  explain+show_data+summary is REJECTED QUALITY. A good 7-slide deck mixes
  ≥4 intent families and ≥2 card/chart slides.
- Density: simpler slides for executive storytelling, denser for analysis. If
  a slide would be ~3 thin bullets, add a stat/icon/graphic or merge it.

VOICE (enforced by a non-overridable gate — violations block the deck):
- Action titles: full sentences that state the "so what" (≥4 words, ≤15),
  active voice. NOT topic labels. e.g. "Two suppliers cover 80% of volume",
  NOT "Supplier overview".
- BANNED words/stems (consulting profile): transformation/transform, optimise/
  optimize, leverage, enable, enhance, empower, facilitate, utilise/utilize,
  synergy, seamless. BANNED phrases: "drive value", "best-in-class",
  "game-changing", "world-class", "cutting-edge", "move the needle",
  "low-hanging fruit", plus filler ("furthermore", "moreover", "delve",
  "it is worth noting", "in conclusion").
- Bullets: parallel, one idea each, ≤ ~3-4 per slide. Prefer specifics/numbers.
- Every data/claim slide should carry a "footnote" starting "Source: " (or
  "Quelle: " in German).

LANGUAGE: set deck "language" (BCP-47) — it drives chart number locale
(de-DE → 1.000 thousands). Write ALL slide text in that language.

Output STRICT JSON only. No prose, no markdown fences, no comments.
""".strip()

SYSTEM_STORYLINE = INTENT_REFERENCE + """

TASK: Propose a storyline (outline) for a deck from the user's brief. Do NOT
write full slide content yet. Return JSON:
{ "title": string, "language": BCP-47, "audience": string,
  "storyline": [ { "intent": <name>, "title": <action-title sentence>,
                   "note": <one line on what this slide will contain> } ] }

STORYLINE LOGIC (ported from the OW storyline system — apply in this order):
1. Infer the INTENT MODE from the brief: update/status → Inform · proposal/
   business case → Recommend · approval/options → Decide · stakeholders/
   rollout → Align · vision/launch → Inspire. Default Inform.
2. Write the CORE MESSAGE as one sentence (if it needs an "and", split it).
   LEAD WITH THE ANSWER: the first content slide reveals the recommendation or
   headline finding — a reader who stops after one minute knows the answer.
3. Build a PYRAMID: 2-4 mutually distinct supporting arguments, each carried
   by exactly one slide, evidence beneath. Opening pattern by mode:
   Recommend/Decide/Inform → answer first, then the case; Align → SCQA
   (situation-complication-question-answer); Inspire → from-state → to-state
   → the path, end on a motivating message. Decide MUST state the decision
   required; Align MUST state next steps and ownership.
4. The slide titles, read top to bottom on their own, must tell the whole
   story (headline sequence). Every title is a full action-title sentence.
5. KEY SLIDES (core message, one per argument, the close) get a VISUAL intent
   (chart/cards/diagram/hero stat) — never explain. The close is never a
   paragraph list. Apply the DESIGN SELECTION rules at the outline level.
6. Long decks with distinct parts: open each part with a section_divider
   (structural slides are free — they don't count against a requested count).
7. Match the DECK PROFILE to the brief: results/KPIs/budget → chart- and
   table-heavy; board/strategy/decision → lean, stat callouts, one idea per
   slide; vision/launch → hero statements and big stats, light text;
   otherwise balanced with a visual on most content slides.
5-9 content slides unless the brief implies otherwise."""

SYSTEM_GENERATE = INTENT_REFERENCE + """

TASK: Produce a COMPLETE deck_spec (the full JSON object with "slides") from
the user's brief and/or a confirmed storyline. Fill every required content key
with real, specific, on-brief content in the deck's language. Where the user
gave no data, invent plausible, clearly illustrative dummy values (round
numbers, obvious placeholders) — never leave a slot empty.

A deterministic gate will BLOCK the deck unless ALL of these hold — comply the
first time:
- Every content-slide title is a FULL SENTENCE with a subject and a verb that
  states the finding (the "so what"), 4-15 words. Never a topic label. Bad:
  "Overview", "Numbers", "Our market". Good: "Two suppliers cover 80% of
  volume at lower cost."
- No banned terms anywhere (transformation, optimise, leverage, enable,
  enhance, empower, facilitate, utilise, synergy, seamless, "drive value",
  "best-in-class", etc.).
- Every data/table/chart slide carries a "footnote" starting "Source: "
  (English) or "Quelle: " (German).
- Each slide is substantive: an explain/summary body has 2-4 parallel bullets,
  not one thin line. Prefer numbers and specifics."""

SYSTEM_GENERATE_SLIDE = INTENT_REFERENCE + """

TASK: Produce ONE slide. Return a single deck-spec slide object
{ "intent": <name>, "content": {...} } (NOT wrapped in "slides"). Pick the
intent that best fits the request; fill all required keys with specific
content or clearly illustrative dummy data in the requested language."""

SYSTEM_EDIT_SLIDE = INTENT_REFERENCE + """

TASK: You are given an existing slide and an edit command. Make the SMALLEST
change that satisfies the command and return the COMPLETE updated slide object
{ "intent": <name>, "content": {...} } (NOT wrapped in "slides").

You may receive TWO descriptions of the slide:
1. "Stored spec" — the spec it was originally generated from (may be absent
   for slides created outside the pane).
2. "Actual slide contents" — an inventory parsed from the LIVE slide right
   now: all text, tables (with cell values), charts (with type, categories,
   and series values), and picture counts.

THE ACTUAL SLIDE CONTENTS ARE GROUND TRUTH. The user may have manually edited
the slide after it was generated (changed numbers, added a chart, pasted
text). Your updated spec MUST preserve those manual changes — carry the
actual data, not the stored spec's stale data — unless the edit command
explicitly overrides them. If an element exists on the slide but not in the
stored spec (e.g. a second chart), include it in the updated spec (compose or
columns_layout can host multiple visuals). If there is no stored spec, first
derive the closest-fitting intent + content from the actual contents, then
apply the command.

Rules:
- KEEP the same intent and overall structure where possible. Preserve every
  field the command does not explicitly touch (titles, columns, footnotes, data).
- Change the intent ONLY if the command explicitly asks for a different element
  type (e.g. "turn this into a bar chart", "make it a table"). "Add icons",
  "reword", "add a column", "change the numbers" do NOT change the intent.
- For show_columns / dashboard, "add icons" means add an "icon" keyword to each
  column/metric — do not convert to a chart or another layout.
- Obey the content-key traps and voice rules. Keep all required keys valid."""


SYSTEM_SHAPE_OPS = """
You are a slide-layout engine for the Oliver Wyman PowerPoint template. You
receive a JSON inventory of ONE live slide — slide size, every shape
{id, name, type, left, top, width, height, text?, selected?} (inches, from the
top-left corner) — and a user command. You return operations the task pane
applies via Office.js. This is DIRECT manipulation of the user's existing
slide: never invent a redesign the command didn't ask for, and never touch
shapes the command doesn't concern.

OUTPUT — STRICT JSON only:
{ "note": <one short line: what you did or why nothing>, "ops": [ ... ] }
Coordinates in INCHES (floats). Colors as 6-digit hex WITHOUT '#'.

THE OW GRID (16:9 slide, 13.333 × 7.5 in):
- Title zone: y < 1.4 — the title placeholder lives here; leave it alone
  unless the command names it.
- CONTENT AREA: x 0.5, y 1.54, width 12.333, height 5.06 (bottom edge 6.6).
- Column grid: split the content width into N equal columns with 0.5 in
  gutters (0.25 in for tight card stacks). col_width = (12.333 - (N-1)*g)/N.
  Snap left edges to 0.5 + i*(col_width+g). Rows stack with 0.5 in gaps.
- Footer zone: y > 6.6 (source line, page number, "© Oliver Wyman") — never
  place or move content there; never touch shapes already there.
BRAND PALETTE: midnightblue 000F47 (primary text + strong fill), cream F7F3EE
(card fill), skyblue 82BAFF, lightblue CEECFF, gold FFBF00 (highlight, use
sparingly), grey 7B7974, lightgrey B9B6B1, pale EBE7E2, white FFFFFF.
Text on cream/white/pale: midnightblue. Text on midnightblue: white.

OPS (the ONLY vocabulary; "id" refers to an inventory shape id):
- {"op":"move","id",...any of left,top,width,height}
- {"op":"fill","id","color"}            {"op":"no_fill","id"}
- {"op":"line","id","color"?,"weight_pt"?}   {"op":"no_line","id"}
- {"op":"font","id","color"?,"size_pt"?,"bold"?}
- {"op":"text","id","text"}             (replaces the shape's text)
- {"op":"delete","id"}
- {"op":"add_textbox","text","left","top","width","height","size_pt"?,"bold"?,"color"?}
- {"op":"add_shape","shape":"rectangle"|"rounded_rectangle"|"oval"|"line",
   "left","top","width","height","fill"?,"line_color"?,"text"?,"font_color"?,
   "size_pt"?,"bold"?}
- {"op":"group","ids":[...]}            (2+ shapes)

TASK PATTERNS:
- "establish a grid" / "align to the grid" / "clean up": infer the intended
  layout from the shapes' ROUGH current positions (shapes at similar y = a
  row; similar x = a column). Pick the smallest fitting column grid, then
  emit move ops that snap every content shape to it — equal widths and equal
  heights for shapes playing the same role, aligned tops within a row,
  gutters exact. Do NOT move: the title placeholder (name contains 'Title'),
  anything in the footer zone, or shapes the user grouped deliberately.
- "add a component": compose it from add_shape/add_textbox ON grid positions
  in free space (or where the user says). Recipes:
  · KPI card = cream rounded_rectangle (≥2.0×1.4) + bold 28pt midnightblue
    value textbox + 12pt label textbox below the value.
  · Insight callout = pale rectangle full column width + bold 12pt heading
    'INSIGHT' + 12pt body text.
  · Section label = midnightblue rectangle, white bold 14pt text.
  · Divider line = "line" shape, grey, weight 1.
  Charts and tables CANNOT be built from shapes — if asked, return ops:[] and
  a note pointing to 'Apply edit' (the server-rendered path).
- Selected shapes ("these", "the selected …", or when shapes carry
  "selected":true and the command is generic like "align these left"):
  operate ONLY on the selected ones. Alignment = move ops you compute:
  align left → same left; center horizontally → same center x; distribute →
  equal spacing between edges; same size → copy the largest/median size.
- Fill/recolor: use the brand palette; map color words (blue → midnightblue,
  yellow/amber → gold, beige → cream).

RULES:
- Every shape stays inside the slide; content stays inside the content area.
- Do not overlap shapes unless the command asks for it (cards on panels are
  intentional overlaps — a textbox ON a rectangle is fine).
- Impossible/unclear command → {"note": short reason, "ops": []}.
""".strip()


def _relaxed_loads(s: str) -> Any:
    """json.loads, but tolerant of the usual LLM slips: // and /* */ comments
    and trailing commas before } or ]."""
    try:
        return json.loads(s)
    except json.JSONDecodeError:
        pass
    cleaned = re.sub(r"/\*.*?\*/", "", s, flags=re.S)
    cleaned = re.sub(r"(?m)//[^\n]*", "", cleaned)
    cleaned = re.sub(r",(\s*[}\]])", r"\1", cleaned)
    return json.loads(cleaned)


def _extract_json(raw: str) -> Any:
    t = raw.strip()
    t = re.sub(r"^```(?:json)?", "", t, flags=re.I).strip()
    t = re.sub(r"```$", "", t).strip()
    try:
        return _relaxed_loads(t)
    except json.JSONDecodeError:
        pass
    m = re.search(r"[\[{]", t)
    if not m:
        raise ValueError(f"no JSON in model output: {t[:200]}")
    start = m.start()
    open_c = t[start]
    close_c = "}" if open_c == "{" else "]"
    depth = 0
    in_str = False
    esc = False
    for i in range(start, len(t)):
        c = t[i]
        if in_str:
            if esc:
                esc = False
            elif c == "\\":
                esc = True
            elif c == '"':
                in_str = False
            continue
        if c == '"':
            in_str = True
        elif c == open_c:
            depth += 1
        elif c == close_c:
            depth -= 1
            if depth == 0:
                return _relaxed_loads(t[start:i + 1])
    raise ValueError(f"unbalanced JSON in model output: {t[:200]}")


# One shared client: connection keep-alive across the ~10 Claude calls a deck
# build makes (storyline + parallel fills + repairs) saves a TLS handshake per
# call. httpx.Client is thread-safe; parallel fills share it.
_client: httpx.Client | None = None


def _http() -> httpx.Client:
    global _client
    if _client is None:
        _client = httpx.Client(
            timeout=120.0,
            limits=httpx.Limits(max_connections=8, max_keepalive_connections=8))
    return _client


def _messages(system: str, user_text: str, max_tokens: int, thinking: bool,
              model: str) -> str:
    """One Anthropic Messages call; returns the concatenated text blocks.
    Retries transient overload/rate-limit (429/500/529) with fixed backoff."""
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise RuntimeError("ANTHROPIC_API_KEY is not set")
    payload: dict[str, Any] = {
        "model": model,
        "max_tokens": max_tokens,
        "system": [{"type": "text", "text": system,
                    "cache_control": {"type": "ephemeral"}}],
        "messages": [{"role": "user", "content": user_text}],
    }
    if thinking:
        payload["thinking"] = {"type": "adaptive"}
        payload["output_config"] = {"effort": "medium"}
    import time
    last = None
    for attempt in range(4):
        resp = _http().post(
            "https://api.anthropic.com/v1/messages",
            headers={"content-type": "application/json", "x-api-key": api_key,
                     "anthropic-version": "2023-06-01"},
            json=payload,
        )
        if resp.status_code == 200:
            data = resp.json()
            return "\n".join(b["text"] for b in data.get("content", [])
                             if b.get("type") == "text")
        last = f"Claude {resp.status_code}: {resp.text}"
        if resp.status_code in (429, 500, 529) and attempt < 3:
            time.sleep(2 * (attempt + 1))
            continue
        raise RuntimeError(last)
    raise RuntimeError(last or "Claude: exhausted retries")


def call_claude(system: str, user_text: str, *, max_tokens: int = 16000,
                thinking: bool = True, fast: bool = False) -> Any:
    """Call Claude and return parsed JSON. fast=True uses the fast model with
    thinking off (the happy path). If the model returns unparseable JSON, retry
    ONCE with an explicit strict-JSON instruction before failing."""
    model = MODEL_FAST if fast else MODEL
    if fast:
        thinking = False
    text = _messages(system, user_text, max_tokens, thinking, model)
    try:
        return _extract_json(text)
    except (json.JSONDecodeError, ValueError):
        strict = (user_text + "\n\nYour previous reply was NOT valid JSON. Reply "
                  "with STRICT JSON ONLY: double-quoted keys and string values, no "
                  "comments, no trailing commas, no ellipses (…), no prose, no "
                  "markdown fences.")
        text = _messages(system, strict, max_tokens, thinking, model)
        return _extract_json(text)
