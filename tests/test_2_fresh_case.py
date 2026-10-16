"""Test 2: a fresh case end to end (IE Business Case Competition 2026,
Alma Carraovejas): rough team notes -> storyline -> spec -> build ->
navigation -> review -> AI check. The storyline approval is simulated."""
import time

import ai_check
import build
import deckkit as dk
import deslop
import navigate
import review
from conftest import HERE, OUT

FIX = HERE / "fixtures" / "fresh_case"


def test_fresh_case_end_to_end():
    timings = {}
    t = time.time()
    errors, warnings = build.validate(build.load_spec(FIX / "deck_spec.yaml"), build.load_catalog(), FIX)
    timings["check"] = time.time() - t
    assert not errors, errors
    t = time.time()
    out, rep = build.build(FIX / "deck_spec.yaml", OUT / "alma_deck.pptx", quiet=True)
    timings["build"] = time.time() - t
    nav = rep["navigation"]          # added by build.py
    assert nav["zooms"] >= 2 and not navigate.check_links(out)
    t = time.time()
    r = review.review(out, client="Alma Carraovejas")
    timings["review"] = time.time() - t
    highs = [f for f in r["findings"] if f["severity"] == "high"]
    assert not highs, highs
    assert not r["ai"]["above_limit"]
    t = time.time()
    work = deslop.propose(out, client="Alma Carraovejas")
    timings["deslop_propose"] = time.time() - t
    assert work["score"] <= ai_check.DEFAULT_THRESHOLD["score"]
    assert not dk.verify_package(out)
    print("timings (s):", {k: round(v, 2) for k, v in timings.items()})
    assert sum(timings.values()) < 60
