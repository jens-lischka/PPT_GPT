"""voice_prose.py — the HUMAN-AUTHORED layer of the voice system.

Owns the editorial judgement that tokens can't capture: the OW house register
(the target voice), the structural foundation pointer, the selection heuristic,
the anti-blandness guidance, per-profile rationale, and before/after exemplars.
Merged into `60_voice_system.md` by voice_profiles.voice_markdown().

Keep this prose; keep it referencing rules by name rather than restating numbers,
so it can't drift from voice_profiles.py.
"""
from __future__ import annotations

# The target voice for BOTH profiles. This is the OW register; the profile only
# decides how it is packaged (bullets vs prose).
HOUSE_VOICE = """
Precise, senior, understated. Sound like a trusted advisor who has done the work
and formed a view — confident through clarity, not adjectives. Let evidence and
structure persuade; state conclusions plainly and show judgement rather than
enthusiasm.

- **Take a position.** If the analysis supports a conclusion, state it. A bland,
  balanced, view-from-nowhere is the main failure mode to avoid.
- **Concrete verbs that say what happens:** decide, reduce, clarify, prioritise,
  fund, stop, scale, pilot, approve, defer, redesign. Avoid vague-active verbs
  (leverage, enable, enhance, transform, optimise, empower, facilitate) and
  generic claims ("drive value", "best-in-class", "game-changing") unless backed
  by specific evidence.
- **Interpret numbers; never just display them.** Say what a figure means for the
  business question.
- **Handle uncertainty honestly.** Distinguish fact, interpretation, and
  assumption. Don't overclaim ("clearly", "without question") or bury the view
  under hedging — calibrate ("far from certain") rather than hedge everything.
- **Mechanics:** sentence case for body; consistent terminology, capitalisation,
  and number formats; spell out acronyms on first use for mixed audiences; cut
  anything that doesn't change what the reader understands, decides, or asks next.
"""

# Pointer, not a copy — the argument structure is owned by 15_storyline_logic.md.
SHARED_FOUNDATION = """
Both profiles render the *same* argument; they differ only in wording. Before
either applies, the storyline is already built per `15_storyline_logic.md`: one
governing message on top, a small MECE set of supporting arguments beneath, the
evidence under each, and an opening pattern (pyramid or SCQA) chosen by Intent
Mode. Read the takeaway titles top-to-bottom and the argument must hold.

The voice profile does not re-decide structure. It decides whether that structure
surfaces as action titles + bullets (consulting) or as topic-sentence-led
paragraphs (longform).
"""

SELECTION_PROSE = """
One profile per deck, picked by how the audience will *consume* it:

- **`consulting`** when the deliverable is presented live or skimmed under time
  pressure by a decision forum (board, steering committee, executives) and the
  goal is a fast decision. The artefact is a deck.
- **`longform`** when the deliverable is read off-platform — thought leadership,
  a white paper, a policy or regulatory document, an internal briefing, a board
  pre-read — and the goal is to persuade through a developed argument. The
  artefact reads as a document.

Tie-breaker: *will they read it alone or watch me present it?* Read → longform;
present → consulting. The choice is declared in the brief as
`voice_style: consulting | longform`; the agent echoes the choice and a one-line
reason before generating, and infers-then-confirms it when the brief is silent.
"""

ANTI_BLAND = """
These apply to both profiles and are the highest-leverage quality rules:

- **State a view.** The single biggest lift. If the evidence points somewhere,
  say so.
- **Ban the AI-filler fingerprint.** No "furthermore / moreover / it's worth
  noting / in conclusion"; no "delve / unleash / pivotal / realm / tapestry /
  leverage"; no reflexive rule-of-three for rhythm.
- **Concrete over abstract.** Specific numbers, named mechanisms, vivid anchors —
  not "significant value" or "a number of factors".
- **Vary cadence.** Mix sentence lengths and openings; cap a fast passage with a
  short line; use parallelism deliberately, not by default.
- **Kill hedging stacks.** "may / might / could potentially" → one calibrated
  modal or a flat claim.
- **Read-aloud test.** If you wouldn't say it to a colleague, rewrite it.
"""

PROFILE_PROSE = {
    "consulting": (
        "The discipline is one idea per slide, proven by the title. Everything on "
        "the slide exists to support the title's claim; if an element doesn't, "
        "cut it (the 'so what?' test). Bullets carry judgement — each is a "
        "complete analytical thought (\"Speed is constrained by late "
        "clarification loops, not capacity\"), never a label (\"Process speed\"). "
        "Name the decision plainly (\"Decide whether to standardise intake before "
        "adding capacity\"), and interpret every figure rather than displaying it."
    ),
    "longform": (
        "The discipline is the paragraph, not the bullet. Each paragraph opens "
        "with its point and then earns it; sentences run old-to-new so the reader "
        "is carried forward; characters are subjects and their actions are verbs "
        "(\"we decided\", not \"a decision was taken\"). Titles may be evocative — "
        "a metaphor or claim with a plain subtitle, in the OW house manner — but "
        "the opening paragraph still states the answer. Reserve bullets for "
        "genuine enumeration; if a slide is drafting as fragments, it isn't yet "
        "longform."
    ),
}

# Before/after exemplars, merged into the generated doc per profile.
EXEMPLARS = {
    "consulting": [
        {"before": "Customer Churn Analysis",
         "after": "Churn rose to 14% in H1, driven entirely by SMB — fixing "
                  "onboarding recovers ~$8M ARR"},
        {"before": "We looked at the data and found several issues with "
                   "onboarding that may be contributing to customers leaving.",
         "after": "Three parallel bullets — \"Onboarding takes 3× longer than "
                  "competitors (18 vs 6 days)\"; \"60% of churned SMB accounts "
                  "never completed setup\"; \"Fixing setup recovers ~$8M ARR by "
                  "FY27\"."},
    ],
    "longform": [
        {"before": "There are a number of factors that could potentially be "
                   "considered as contributing to the increase in churn observed "
                   "during the period in question.",
         "after": "Churn climbed to 14% in the first half — and almost all of it "
                  "came from one place. Small-business customers, not enterprise "
                  "accounts, drove the rise. The cause is mundane but expensive: "
                  "onboarding takes three times longer than at competitors, and "
                  "three in five churned SMB accounts never finished setting up. "
                  "Fix the setup flow and roughly $8 million in recurring revenue "
                  "comes back."},
    ],
}
