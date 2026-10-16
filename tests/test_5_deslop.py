"""Test 5: AI-default text is flagged; approved rewrites bring the score
under the limit and keep every number, name and run format."""
import yaml

import ai_check
import build
import deslop
from conftest import HERE, OUT

FIX = HERE / "fixtures" / "ai_default"


def _deck():
    out, _ = build.build(FIX / "deck_spec.yaml", OUT / "ai_default.pptx", quiet=True)
    return out


def test_ai_default_is_flagged():
    r = ai_check.check(_deck(), client="Tajín")
    bad, rate = ai_check.over_threshold(r)
    assert bad and r["score"] >= 40


def test_winning_deck_is_within_limit(winning):
    r = ai_check.check(winning / "Mexico 24H Case.pptx", client="Tajín")
    assert not ai_check.over_threshold(r)[0]


def test_approved_rewrites_fix_it():
    deck = _deck()
    r = deslop.apply(deck, FIX / "edits.yaml", OUT / "ai_default_fixed.pptx", client="Tajín")
    assert r["saved"], r.get("problems")
    after = ai_check.check(OUT / "ai_default_fixed.pptx", client="Tajín")
    assert not ai_check.over_threshold(after)[0]
    assert r["score_after"] < r["score_before"]


def test_number_change_is_refused(tmp_path):
    deck = _deck()
    edits = tmp_path / "bad.yaml"
    edits.write_text(yaml.safe_dump({"edits": [{"slide": 1, "box": "impact", "approved": True,
        "find": "USD ~130m", "replace": "USD ~150m"}]}))
    r = deslop.apply(deck, edits, tmp_path / "x.pptx")
    assert not r["saved"]


def test_unapproved_edits_are_ignored(tmp_path):
    deck = _deck()
    edits = tmp_path / "e.yaml"
    edits.write_text(yaml.safe_dump({"edits": [{"slide": 1, "box": "impact", "approved": False,
        "find": "a pivotal opportunity to", "replace": ""}]}))
    r = deslop.apply(deck, edits, tmp_path / "x.pptx")
    assert r["saved"] and r["applied"] == 0
