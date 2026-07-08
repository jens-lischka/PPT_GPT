"""voice_lint.py — deterministic voice linting + reviewer rubrics + metrics.

The voice analogue of design_lint, modelled on Vale's rule types (existence,
substitution, occurrence) so prose gets code-like linting:

  lint_text(text, profile)  -> findings + passed (errors==0) + metrics
  REVIEWERS                  -> the judgement-call rubric (for an LLM reviewer
                                pass; regex can't score "is there a point of view")
  metrics(text)             -> instrumentation: hedge count, filler count,
                                sentence-length variation, avg sentence length

Severities: error (must fix), warning (should fix), info (FYI). The deck/agent
gate fails on errors only, like the design lint.
"""
from __future__ import annotations

import re
from typing import Dict, List

try:
    import voice_profiles as _vp
except Exception:                                    # pragma: no cover
    _vp = None

# --- Vale-style rule data ---------------------------------------------------
# existence: phrases that should never appear (AI filler / hedge stacks).
BANNED_PHRASES = [
    "furthermore", "moreover", "it's worth noting", "it is worth noting",
    "in conclusion", "needless to say", "delve", "unleash", "pivotal",
    "tapestry", "realm", "navigating the complexities", "it is important to note",
    "as we all know", "at the end of the day",
]
# substitution: clutter -> preferred (Zinsser / Paramedic Method).
CLUTTER_SWAP = {
    "utilize": "use", "utilise": "use", "leverage": "use",
    "in order to": "to", "at the present time": "now", "due to the fact that":
    "because", "a number of": "several", "facilitate": "help",
    "with regard to": "on", "in the event that": "if", "prior to": "before",
    "subsequent to": "after", "a majority of": "most", "in spite of the fact":
    "although",
}
# occurrence: hedges (counted against a per-profile budget).
HEDGES = ["may", "might", "could", "perhaps", "possibly", "potentially",
          "arguably", "seems to", "appears to", "relatively", "somewhat",
          "fairly", "rather", "quite", "to some extent"]

# Vague-active verb stems and generic-claim clichés — the OW house-voice
# blacklist. SINGLE SOURCE OF TRUTH: the build gate (gate.py) imports these,
# so the deck-build voice check and the draft voice lint never drift apart.
BANNED_VERB_STEMS = (
    "leverag", "enabl", "enhanc", "transform", "optimi", "empower",
    "facilitat", "utiliz", "utilis", "synerg", "seamless",
)
GENERIC_CLAIMS = (
    "drive value", "best-in-class", "best in class", "game-changing",
    "game changing", "unlock significant value", "world-class", "world class",
    "cutting-edge", "cutting edge", "move the needle", "low-hanging fruit",
)

_SENT_SPLIT = re.compile(r"[.!?]+(?:\s+|$)")
_WORD = re.compile(r"\b[\w'%$-]+\b")


def _finding(sev, rule, msg, ctx=""):
    return {"severity": sev, "rule": rule, "message": msg, "context": ctx}


def _sentences(text: str) -> List[str]:
    return [s.strip() for s in _SENT_SPLIT.split(text) if s.strip()]


def metrics(text: str) -> Dict:
    """Instrumentation (Stage 5): the numbers we track per draft to tune the
    profiles. Not a pass/fail — a read-out."""
    low = text.lower()
    sents = _sentences(text)
    lengths = [len(_WORD.findall(s)) for s in sents] or [0]
    avg = round(sum(lengths) / len(lengths), 1)
    # coefficient of variation of sentence length — low = monotone (an AI tell).
    if len(lengths) > 1 and avg:
        mean = sum(lengths) / len(lengths)
        var = sum((l - mean) ** 2 for l in lengths) / len(lengths)
        cv = round((var ** 0.5) / mean, 2)
    else:
        cv = 0.0
    hedges = sum(len(re.findall(rf"\b{re.escape(h)}\b", low)) for h in HEDGES)
    filler = sum(1 for p in BANNED_PHRASES if p in low)
    clutter = sum(1 for c in CLUTTER_SWAP if re.search(rf"\b{re.escape(c)}\b", low))
    return {
        "sentences": len(sents),
        "avg_sentence_words": avg,
        "length_variation_cv": cv,     # higher is better; < ~0.3 reads monotone
        "hedges": hedges,
        "filler_phrases": filler,
        "clutter_words": clutter,
    }


def lint_text(text: str, profile: str = "consulting") -> Dict:
    findings: List[dict] = []
    low = text.lower()
    prof = (_vp.VOICE_PROFILES.get(profile) if _vp else None) or {}

    # existence — banned filler
    for p in BANNED_PHRASES:
        if p in low:
            findings.append(_finding("error", "filler",
                            f"drop AI-filler phrase '{p}'", p))
    # substitution — clutter words
    for bad, good in CLUTTER_SWAP.items():
        if re.search(rf"\b{re.escape(bad)}\b", low):
            findings.append(_finding("error", "clutter",
                            f"replace '{bad}' with '{good}'", bad))

    sents = _sentences(text)
    # occurrence — hedge budget (per profile; longform budgets per-paragraph,
    # consulting per-unit; we approximate the whole text against the tighter rule)
    sent_block = prof.get("sentence", {})
    budget = sent_block.get("hedge_budget_per_unit") or \
        sent_block.get("hedge_budget_per_paragraph") or 1
    hedge_n = metrics(text)["hedges"]
    # scale the budget by rough unit count so long drafts aren't unfairly failed
    units = max(1, len(sents) // 4)
    if hedge_n > budget * units:
        findings.append(_finding("warning", "hedging",
                        f"{hedge_n} hedges exceed budget (~{budget}/unit) — "
                        f"state the claim or calibrate once", ""))

    # profile-specific surface rules
    if profile == "consulting":
        # title length: first non-empty line treated as the title
        first = next((l.strip() for l in text.splitlines() if l.strip()), "")
        tmax = prof.get("title", {}).get("max_words", 15)
        if first and len(_WORD.findall(first)) > tmax:
            findings.append(_finding("warning", "title-length",
                            f"title is >{tmax} words — tighten the action title",
                            first[:60]))
    elif profile == "longform":
        # prose discipline: flag bullet-heavy drafts
        bullets = len(re.findall(r"^\s*[-*•]\s+", text, flags=re.M))
        if bullets >= 4:
            findings.append(_finding("warning", "prose-discipline",
                            f"{bullets} bullets — longform should develop the "
                            f"argument in paragraphs, not fragments", ""))
        # paragraph monotony
        if metrics(text)["length_variation_cv"] < 0.25 and len(sents) >= 4:
            findings.append(_finding("info", "cadence",
                            "low sentence-length variation reads monotone — "
                            "vary sentence length", ""))

    errors = sum(1 for f in findings if f["severity"] == "error")
    warnings = sum(1 for f in findings if f["severity"] == "warning")
    info = sum(1 for f in findings if f["severity"] == "info")
    return {
        "profile": profile,
        "passed": errors == 0,
        "findings": findings,
        "summary": {"errors": errors, "warnings": warnings, "info": info},
        "metrics": metrics(text),
    }


# --- Reviewer rubrics (the judgement-call layer, for an LLM reviewer pass) ---
# Regex can't score these; the agent runs them as a self-review per profile.
REVIEWERS = {
    "consulting": [
        "Does the title state a so-what (a complete-sentence takeaway), not a label?",
        "Read the titles in sequence — do they tell the whole story (horizontal logic)?",
        "Is there one idea per slide, with every element supporting the title?",
        "Are bullets complete analytical thoughts, parallel, and answer-first?",
        "Is a clear position taken, with numbers interpreted not just shown?",
    ],
    "longform": [
        "Does each paragraph open with its point (topic sentence)?",
        "Do sentences flow old-to-new, carrying the reader forward?",
        "Is the answer stated in the opening paragraph?",
        "Are characters the subjects and their actions the verbs (no nominalisations)?",
        "Is sentence length varied, with a clear point of view?",
    ],
}
