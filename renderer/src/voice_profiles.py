"""voice_profiles.py — the writing-voice system for the OW deck/report agent.

This is the voice analogue of the design system. Same architecture as the
design-system pipeline, for consistency:

  - voice_profiles.py (this file) — the NORMATIVE tokens: two selectable
    writing profiles (`consulting`, `longform`) as checkable rules, plus the
    generator that emits the `60_voice_system.md` reference.
  - voice_prose.py — the HUMAN-AUTHORED rationale: the OW house register
    (target voice) and the before/after exemplars.
  - voice_lint.py — deterministic lint (banned phrases, clutter, hedge budget,
    title length, bullet/prose discipline) + reviewer rubrics + draft metrics.

Design intent (what this system is and is NOT):
  - ONE profile per deliverable, chosen by AUDIENCE, declared in the brief
    (`voice_style: consulting | longform`). It is a *simplification* — a single
    surface register for the whole deck — not a per-slide setting and not a new
    set of intents.
  - It governs the SURFACE only (register, titles, bullets-vs-prose, sentence
    mechanics). The underlying ARGUMENT structure (pyramid, MECE, SCQA opening,
    headline ladder) is owned by `15_storyline_logic.md` and is referenced here,
    never duplicated. Clear roles: storyline = structure; voice = wording.
  - The house register in voice_prose applies to BOTH profiles and does not
    change; the profile only decides how that register is packaged.
"""
from __future__ import annotations

from typing import Dict, List


# ---------------------------------------------------------------------------
# The two profiles. Values are hard, checkable rules (the lint reads these).
# `audiences` drives selection; `select_when` is the human heuristic.
# ---------------------------------------------------------------------------
VOICE_PROFILES: Dict[str, dict] = {
    "consulting": {
        "register": "action-first",
        "one_line": "Answer-first slides built from action titles and short, "
                    "parallel bullets — one idea per slide, the title carries "
                    "the 'so what'.",
        "audiences": ["executives", "board", "steering committee",
                      "decision forum", "project team"],
        "select_when": "the deliverable is PRESENTED live or skimmed under time "
                       "pressure, the audience is a decision forum, and the goal "
                       "is a fast decision. Format: slides.",
        "title": {
            "style": "action-title",        # complete sentence, states the takeaway
            "max_words": 15,
            "voice": "active",
            "must_state_so_what": True,
        },
        "body": {
            "mode": "bullets",              # bullets-led
            "bullets_max_per_unit": 30,
            "bullets_are": "short parallel fragments, each a complete analytical "
                           "thought, never a label",
            "bullets_overflow": "engine continues onto the next slide automatically",
            "paragraphs": "discouraged (a short lead-in line is fine)",
            "lead": "answer-first (BLUF)",
            "source_every_number": True,
        },
        "sentence": {
            "register": "telegraphic but not cryptic; strong verbs, concrete nouns",
            "length_variation": "required",
            "hedge_budget_per_unit": 1,
            "point_of_view": "required",    # take a position
        },
        "structure_ref": "15_storyline_logic.md",   # the argument lives there
    },
    "longform": {
        "register": "editorial-prose",
        "one_line": "Flowing, topic-sentence-led paragraphs that develop the "
                    "argument across sentences, with very limited bullets — like "
                    "a printed strategy report.",
        "audiences": ["internal", "readers", "policy", "regulator",
                      "thought-leadership", "board (pre-read)"],
        "select_when": "the deliverable is READ off-platform (white paper, "
                       "thought leadership, policy/regulatory doc, board memo), "
                       "and the goal is to persuade through a developed argument. "
                       "Format: a document or text-heavy leave-behind.",
        "title": {
            "style": "evocative-or-declarative",   # metaphor+colon, claim, or framing Q
            "max_words": 14,
            "voice": "active",
            "must_state_so_what": False,   # the OPENING paragraph states it instead
        },
        "body": {
            "mode": "prose",               # paragraph-led
            "paragraph_topic_sentence": True,    # each para opens with its point
            "cohesion": "old-to-new (open with familiar, end with new/emphatic)",
            "nominalizations": "discouraged (characters as subjects, actions as verbs)",
            "bullets": "minimal — only for genuine enumeration",
            "paragraph_words": "≈40–90; vary length; never 4 same-length in a row",
            "lead": "answer-first opening paragraph",
        },
        "sentence": {
            "register": "active voice, strong verbs; cut clutter (Paramedic Method)",
            "length_variation": "required",
            "hedge_budget_per_paragraph": 1,
            "point_of_view": "required",
        },
        "structure_ref": "15_storyline_logic.md",
    },
}


def profile_names() -> List[str]:
    return sorted(VOICE_PROFILES.keys())


def select_profile_from_audience(audience: str) -> str:
    """Heuristic: map a free-text audience to a profile. Returns the profile name.
    Falls back to 'consulting' (the default for slide deliverables)."""
    a = (audience or "").lower()
    longform_signals = ("internal", "all-staff", "all staff", "employee",
                        "policy", "regulat", "white paper", "thought",
                        "readers", "memo", "pre-read", "preread", "report")
    if any(sig in a for sig in longform_signals):
        return "longform"
    return "consulting"


def _kv_table(d: dict, keymap=None) -> str:
    rows = []
    for k, v in d.items():
        label = (keymap or {}).get(k, k)
        if isinstance(v, bool):
            v = "yes" if v else "no"
        rows.append(f"| `{label}` | {v} |")
    return "\n".join(rows)


def voice_spec(profile: str) -> str:
    """Compact, prompt-injectable spec for ONE profile — the text the agent loads
    when that style is selected (progressive disclosure). Kept terse on purpose."""
    p = VOICE_PROFILES[profile]
    t, b = p["title"], p["body"]
    lines = [f"VOICE PROFILE: {profile} — {p['register']}",
             p["one_line"], "",
             "TITLES: " + (
                 f"action title (complete sentence, states the so-what), "
                 f"≤{t['max_words']} words, active voice"
                 if profile == "consulting" else
                 f"evocative or declarative, ≤{t['max_words']} words; the OPENING "
                 f"paragraph states the point"),
             "BODY: " + (
                 f"bullets-led, ≤{b['bullets_max_per_unit']} short parallel "
                 f"bullets/slide, each a complete thought; answer-first; source "
                 f"every number"
                 if profile == "consulting" else
                 "prose-led; each paragraph opens with its point (topic sentence); "
                 "old-to-new flow; bullets only for real enumeration"),
             "SENTENCES: strong verbs, concrete nouns; vary sentence length; "
             "take a position; one hedge max per " +
             ("slide" if profile == "consulting" else "paragraph") + ".",
             "STRUCTURE: the argument (pyramid/MECE/SCQA opening) is built per "
             "15_storyline_logic.md — this profile only sets the wording.",
             "HOUSE VOICE applies on top (see voice_prose / 60_voice_system.md)."]
    return "\n".join(lines)


def voice_markdown() -> str:
    """Render the full voice-system reference (`60_voice_system.md`) from the live
    profiles, merging the human-authored register + exemplars from voice_prose.
    Generated from code so it can't drift; printed by `--list-voice`."""
    try:
        import voice_prose as _vp
        HOUSE = _vp.HOUSE_VOICE.strip()
        SHARED = _vp.SHARED_FOUNDATION.strip()
        ANTIBLAND = _vp.ANTI_BLAND.strip()
        EXEMPLARS = _vp.EXEMPLARS
        SELECT = _vp.SELECTION_PROSE.strip()
        PROFILE_PROSE = _vp.PROFILE_PROSE
    except Exception:
        HOUSE = SHARED = ANTIBLAND = SELECT = ""
        EXEMPLARS, PROFILE_PROSE = {}, {}

    def profile_block(name: str) -> str:
        p = VOICE_PROFILES[name]
        rationale = PROFILE_PROSE.get(name, "").strip()
        title_tbl = _kv_table(p["title"])
        body_tbl = _kv_table(p["body"])
        sent_tbl = _kv_table(p["sentence"])
        ex = EXEMPLARS.get(name, [])
        ex_md = ""
        if ex:
            ex_md = "\n**Before → after**\n\n" + "\n\n".join(
                f"- *Generic:* {e['before']}\n- *{name.title()}:* {e['after']}"
                for e in ex) + "\n"
        return (
            f"### Profile: `{name}` — {p['register']}\n"
            f"{p['one_line']}\n\n"
            f"**Audiences:** {', '.join(p['audiences'])}  \n"
            f"**Select when:** {p['select_when']}\n\n"
            f"{rationale + chr(10) + chr(10) if rationale else ''}"
            f"**Titles**\n\n| Rule | Value |\n| --- | --- |\n{title_tbl}\n\n"
            f"**Body**\n\n| Rule | Value |\n| --- | --- |\n{body_tbl}\n\n"
            f"**Sentences**\n\n| Rule | Value |\n| --- | --- |\n{sent_tbl}\n"
            f"{ex_md}")

    profiles_md = "\n".join(profile_block(n) for n in profile_names())

    return f"""# OW Voice System — writing-style reference

One deliverable, one voice. The profile is chosen by **audience** and declared in
the brief (`voice_style: consulting | longform`). It sets the *surface* — register,
titles, bullets-vs-prose, sentence mechanics. The *argument structure* (pyramid,
MECE, SCQA opening, headline ladder) is owned by `15_storyline_logic.md` and is
not repeated here. The **house voice** below applies to both profiles and does
not change.

This file is generated from `voice_profiles.py` (+ `voice_prose.py`); edit those,
not this doc. Lint a draft with `generate_deck.py --lint-text`.

## House voice (applies to both profiles)
{HOUSE}

## Shared foundation (structure lives in 15_storyline_logic.md)
{SHARED}

## Choosing the profile
{SELECT}

## Profiles
{profiles_md}
## Avoiding bland / templated prose
{ANTIBLAND}

## Linting & review
A draft is checked by `generate_deck.py --lint-text <file> --profile <name>`:
deterministic rules flag banned filler, clutter words, over-budget hedging,
over-long titles (consulting), and bullets-where-prose-belongs (longform). The
judgement calls — does the title state a so-what, does each paragraph open with
its point, is there a point of view — are scored by the reviewer rubric (see
`voice_lint.REVIEWERS`). Run generate → lint → revise until the lint passes.
"""
