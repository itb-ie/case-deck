"""Test 1: rebuild the Tajín main body from its own content."""
import shutil

import pytest

import build
import deckkit as dk
import review
from conftest import OUT

EXPECTED = ["title", "exec_summary", "section_divider", "company_position", "trends", "hook_image",
            "matrix_takeaways", "three_cards", "three_cards", "section_divider", "prioritization",
            "strategy_tree", "framework", "framework", "pillar_overview", "initiative_steps", "initiative",
            "initiative", "framework", "pillar_overview", "initiative", "framework", "pillar_overview",
            "section_divider", "driver_tree", "financials", "risk_matrix", "mitigation", "roadmap", "roadmap",
            "exec_summary", "closing", "appendix_divider", "appendix_index", "text_content", "chart_takeaways",
            "harvey_table"]


def test_golden_structure(golden_deck):
    out, rep = golden_deck
    prs = dk.open_deck(out)
    got = [build.slide_meta(s).get("pattern") for s in prs.slides if not build.slide_meta(s).get("nav_generated")]
    assert got == EXPECTED
    assert rep.get("navigation"), "build must add navigation by default"
    assert rep["seconds"] < 120, "35+ slides must build in under 2 minutes"
    assert not dk.verify_package(out)


def test_golden_review_clean(golden_deck):
    out, _ = golden_deck
    r = review.review(out, run_ai=True)
    highs = [f for f in r["findings"] if f["severity"] == "high"]
    assert not highs, highs
    assert r["ai"] and not r["ai"].get("above_limit")


def test_golden_side_by_side(golden_deck, winning):
    if not shutil.which("qlmanage"):
        pytest.skip("no renderer")
    import golden_compare
    pairs = golden_compare.compare(winning, golden_deck[0], OUT / "golden_compare", engine="quicklook")
    assert len(pairs) >= len(EXPECTED) - 1
