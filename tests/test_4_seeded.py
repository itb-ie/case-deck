"""Test 4: the review finds at least 9 of 10 planted defects, at most 2 false alarms."""
from pathlib import Path

import build
import review
import seeded
from conftest import HERE, OUT


def test_seeded_defects(golden_deck):
    spec = seeded.clean_spec(HERE / "fixtures" / "tajin" / "deck_spec.yaml", HERE / "fixtures" / "tajin" / "seeded_clean_spec.yaml")
    clean, _ = build.build(spec, OUT / "seeded_clean.pptx", quiet=True)
    defects = seeded.plant(clean, OUT / "seeded.pptx")
    base = review.review(clean, run_ai=False)
    r = review.review(OUT / "seeded.pptx", run_ai=False)
    found, false = seeded.score(r, defects, base)
    missed = [d for d in defects if d not in found]
    assert len(found) >= 9, f"missed {missed}"
    assert len(false) <= 2, false
