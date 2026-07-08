"""Tests for the voice system: profiles, generator/prose merge, lint, metrics,
reviewer rubrics, and audience->profile selection."""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

import voice_profiles as vp
import voice_lint as vl


# --- profiles & generator --------------------------------------------------
def test_two_profiles_exist():
    assert set(vp.profile_names()) == {"consulting", "longform"}

def test_voice_doc_generates_with_both_profiles_and_prose():
    md = vp.voice_markdown()
    assert "## House voice" in md
    assert "### Profile: `consulting`" in md
    assert "### Profile: `longform`" in md
    # prose merged in (a distinctive phrase from voice_prose.HOUSE_VOICE)
    assert "Take a position" in md
    # references structure file rather than duplicating it
    assert "15_storyline_logic.md" in md

def test_voice_spec_is_compact_and_profile_specific():
    c = vp.voice_spec("consulting")
    l = vp.voice_spec("longform")
    assert "action title" in c and "bullets-led" in c
    assert "topic sentence" in l and "prose-led" in l

def test_exemplars_present_per_profile():
    md = vp.voice_markdown()
    assert "Before → after" in md


# --- selection -------------------------------------------------------------
def test_select_longform_for_reading_audiences():
    assert vp.select_profile_from_audience("internal all-staff briefing") == "longform"
    assert vp.select_profile_from_audience("a regulatory white paper") == "longform"
    assert vp.select_profile_from_audience("board pre-read memo") == "longform"

def test_select_consulting_for_decision_forums():
    assert vp.select_profile_from_audience("the board, for a decision") == "consulting"
    assert vp.select_profile_from_audience("executive steering committee") == "consulting"

def test_select_defaults_to_consulting():
    assert vp.select_profile_from_audience("") == "consulting"


# --- lint ------------------------------------------------------------------
def test_lint_flags_filler_as_error():
    rep = vl.lint_text("Furthermore, we should delve into this.", "consulting")
    assert not rep["passed"]
    assert any(f["rule"] == "filler" for f in rep["findings"])

def test_lint_flags_clutter_with_swap():
    rep = vl.lint_text("We will utilize the data in order to win.", "consulting")
    assert any(f["rule"] == "clutter" and "use" in f["message"] for f in rep["findings"])

def test_lint_clean_consulting_passes():
    rep = vl.lint_text("Churn rose 14% in H1, driven by SMB onboarding.", "consulting")
    assert rep["passed"]

def test_lint_longform_flags_bullet_heavy():
    text = "- one\n- two\n- three\n- four\n- five"
    rep = vl.lint_text(text, "longform")
    assert any(f["rule"] == "prose-discipline" for f in rep["findings"])

def test_lint_consulting_flags_long_title():
    long_title = ("This is an extremely long slide title that goes well beyond "
                  "the fifteen word action title limit for the consulting profile here")
    rep = vl.lint_text(long_title, "consulting")
    assert any(f["rule"] == "title-length" for f in rep["findings"])


# --- metrics (instrumentation) ---------------------------------------------
def test_metrics_counts_hedges_and_variation():
    m = vl.metrics("We may possibly do this. It might perhaps work. Short.")
    assert m["hedges"] >= 3
    assert m["sentences"] == 3
    assert "length_variation_cv" in m

def test_reviewers_cover_both_profiles():
    assert set(vl.REVIEWERS) == {"consulting", "longform"}
    assert all(len(qs) >= 3 for qs in vl.REVIEWERS.values())
