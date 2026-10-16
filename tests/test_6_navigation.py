"""Test 6: zooms for every appendix slide, links resolve, re-running does
not duplicate, the navigator is hidden."""
import build
import deckkit as dk
import navigate
from conftest import OUT


def _count(path):
    prs = dk.open_deck(path)
    zooms = sum(1 for s in prs.slides for e in s.shapes._spTree.iter() if str(e.tag).endswith("}sldZmObj"))
    nav = sum(1 for s in prs.slides for e in s.shapes._spTree.iter("{%s}cNvPr" % dk.NS["p"]) if e.get("name", "").startswith("nav:"))
    return prs, zooms, nav


def test_navigation(golden_deck):
    out = OUT / "tajin_nav.pptx"
    r = navigate.navigate(golden_deck[0], out)
    prs, zooms, nav = _count(out)
    appendix = [s for s in prs.slides if build.slide_meta(s).get("appendix")
                and build.slide_meta(s).get("pattern") not in ("appendix_divider", "appendix_index")]
    navigator = [s for s in prs.slides if build.slide_meta(s).get("nav_generated") == "navigator"]
    assert len(navigator) == 1 and navigator[0]._element.get("show") == "0"
    index_zooms = [e for s in prs.slides if build.slide_meta(s).get("pattern") == "appendix_index"
                   for e in s.shapes._spTree.iter() if str(e.tag).endswith("}sldZmObj")]
    assert len(index_zooms) == len(appendix)
    assert all(z.find("{%s}zmPr" % dk.NS["pslz"]).get("returnToParent") == "1" for z in index_zooms)
    assert not navigate.check_links(out)
    assert not dk.verify_package(out)
    assert r["corner_links"] >= 25

    out2 = OUT / "tajin_nav2.pptx"
    navigate.navigate(out, out2)
    _, zooms2, nav2 = _count(out2)
    assert (zooms2, nav2) == (zooms, nav)
    assert len(dk.open_deck(out2).slides) == len(prs.slides)
    assert not dk.verify_package(out2)
