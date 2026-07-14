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

MODEL = os.environ.get("ANTHROPIC_MODEL", "claude-opus-4-8")

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
- show_columns: title, columns:[{heading, bullets|paragraphs}]  (2-5 parallel headed columns)
- compare: title, compare {pros:[...], cons:[...], pros_heading?, cons_heading?}  (green check / red cross)
- dashboard: title, metrics:[{value, heading, text?}]   (grid of KPI cards, variadic)
- stat_callout: title, stats:[{value, heading, text}]   (1-4 big-number cards)
- icon_rows: title, rows:[{icon, heading, text}]        (1-5 rows; icon = short keyword)
- show_trend_with_key_message: title, chart {type, ...}, insight (text)   (chart + callout)
- show_waterfall: title, waterfall {categories:[...], values:[...], totals:[indices], orientation?:"vertical"|"horizontal", heading?, subheading?}
- show_treemap: title, treemap {items:[{label, value}], heading?, subheading?}
- show_timeline: title, milestones:[{date|label, title, text?}]
- show_process: title, steps:[{title, text?}]           (horizontal chevrons)
- show_growing_steps: title, steps:[{title, text?}]     (ascending staircase)
- show_pyramid: title, levels:[{label, text?}]          (label "Head: body" → bold head)
- show_cycle: title, items:[...]                        (genuine loops only)
- show_org_chart: title, org_chart {root:{...}, reports:[...]}
- show_matrix: title, matrix {x_axis:STRING, y_axis:STRING, items:[{row, col, label}]}
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
Aim for a tight narrative: cover → (contents if long) → situation →
complication → evidence (data/chart) → recommendation → summary. 5-9 slides
unless the brief implies otherwise. Titles must already be action titles."""

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

TASK: You are given one existing slide spec and an edit command. Make the
SMALLEST change that satisfies the command and return the COMPLETE updated
slide object { "intent": <name>, "content": {...} } (NOT wrapped in "slides").

Rules:
- KEEP the same intent and overall structure. Preserve every field the command
  does not explicitly touch (titles, other columns, footnotes, data).
- Change the intent ONLY if the command explicitly asks for a different element
  type (e.g. "turn this into a bar chart", "make it a table"). "Add icons",
  "reword", "add a column", "change the numbers" do NOT change the intent.
- For show_columns / dashboard, "add icons" means add an "icon" keyword to each
  column/metric — do not convert to a chart or another layout.
- Obey the content-key traps and voice rules. Keep all required keys valid."""


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


def _messages(system: str, user_text: str, max_tokens: int, thinking: bool) -> str:
    """One Anthropic Messages call; returns the concatenated text blocks.
    Retries transient overload/rate-limit (429/500/529) with fixed backoff."""
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise RuntimeError("ANTHROPIC_API_KEY is not set")
    payload: dict[str, Any] = {
        "model": MODEL,
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
        resp = httpx.post(
            "https://api.anthropic.com/v1/messages",
            headers={"content-type": "application/json", "x-api-key": api_key,
                     "anthropic-version": "2023-06-01"},
            json=payload, timeout=120.0,
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
                thinking: bool = True) -> Any:
    """Call Claude and return parsed JSON. If the model returns unparseable
    JSON, retry ONCE with an explicit strict-JSON instruction before failing."""
    text = _messages(system, user_text, max_tokens, thinking)
    try:
        return _extract_json(text)
    except (json.JSONDecodeError, ValueError):
        strict = (user_text + "\n\nYour previous reply was NOT valid JSON. Reply "
                  "with STRICT JSON ONLY: double-quoted keys and string values, no "
                  "comments, no trailing commas, no ellipses (…), no prose, no "
                  "markdown fences.")
        text = _messages(system, strict, max_tokens, thinking)
        return _extract_json(text)
