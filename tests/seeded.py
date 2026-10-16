"""Seeded-defect deck: a clean deck built from the golden spec (without the
risk slides), then 10 defects planted by editing the .pptx the way a team
might by hand. review.py must find at least 9 with at most 2 false alarms."""
from __future__ import annotations

import copy
import sys
from pathlib import Path

import yaml

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "scripts"))
import build  # noqa: E402
import deckkit as dk  # noqa: E402

KEEP = ["title", "exec_summary", "section_divider", "company_position", "trends", "three_cards",
        "framework", "pillar_overview", "initiative", "driver_tree", "financials", "roadmap",
        "exec_summary", "closing", "appendix_divider", "appendix_index", "text_content", "chart_takeaways"]


def clean_spec(golden_spec: Path, out_spec: Path):
    spec = yaml.safe_load(golden_spec.read_text(encoding="utf-8"))
    slides, used = [], {}
    for sl in spec["slides"]:
        pid = sl["pattern"]
        want = KEEP.count(pid)
        if used.get(pid, 0) < want and not sl.get("current"):
            s2 = dict(sl)
            s2.pop("takeaway", None)
            slides.append(s2)
            used[pid] = used.get(pid, 0) + 1
    spec["slides"] = slides
    spec["deck"]["output"] = "seeded_clean.pptx"
    out_spec.write_text(yaml.safe_dump(spec, allow_unicode=True, sort_keys=False), encoding="utf-8")
    return out_spec


def plant(clean_deck: Path, out_deck: Path):
    """Returns the list of planted defects: (slide number, review category)."""
    prs = dk.open_deck(clean_deck)
    slides = list(prs.slides)
    by_pattern = {}
    for i, s in enumerate(slides, 1):
        by_pattern.setdefault(build.slide_meta(s).get("pattern"), []).append((i, s))
    defects = []

    # 1. topic title on the trends slide
    n, s = by_pattern["trends"][0]
    dk.set_text(dk.placeholder(s, 0), "Market analysis")
    defects.append((n, "title"))
    # 2. headline figure in the exec summary no longer matches the financials
    n, s = by_pattern["exec_summary"][0]
    dk.set_text(dk.slot_shapes(s)["impact"], "Impact: contribute USD ~160m in revenue by 2030")
    defects.append((n, "numbers"))
    # 3. numbers but the Source line was deleted (company position)
    n, s = by_pattern["company_position"][0]
    dk.remove_shape(dk.slot_shapes(s)["source"])
    defects.append((n, "source"))
    # 4. "HOOK TBD" title on the three-cards slide
    n, s = by_pattern["three_cards"][0]
    dk.set_text(dk.placeholder(s, 0), "HOOK TBD")
    defects.append((n, "leftover"))
    # 5. leftover template text in the appendix text slide
    n, s = by_pattern["text_content"][0]
    body = dk.slot_shapes(s)["body"]
    dk.set_text(body, [p.text for p in body.text_frame.paragraphs] + ["Click to add text"])
    defects.append((n, "leftover"))
    # 6. a company from another case on the pillar overview
    n, s = by_pattern["pillar_overview"][0]
    dk.set_text(dk.slot_shapes(s)["i1_case"], "**MOL** grew loyalty through its app, with 700K users in two years")
    defects.append((n, "other-case"))
    # 7. text overflow: a long paragraph pasted into a small box (initiative)
    n, s = by_pattern["initiative"][0]
    dk.set_text(dk.slot_shapes(s)["b1_text"], "Develop co-branded snacks with global partners using Tajín’s chile-lime DNA, "
                "starting with two partners in the U.S. and Mexico, then extending to Spain and Japan with local flavours, "
                "shared shelf space, joint promotions in convenience stores, and a creator programme on TikTok that "
                "seeds the launch with sampling events in 12 cities and limited packs for each market")
    defects.append((n, "overflow"))
    # 8. 7 pt text on the roadmap
    n, s = by_pattern["roadmap"][0]
    for r in dk.slot_shapes(s)["ph2_text"].text_frame.paragraphs[0].runs:
        r.font.size = dk.Pt(7)
    defects.append((n, "font"))
    # 9. off-palette colour on the driver tree
    n, s = by_pattern["driver_tree"][0]
    from pptx.dml.color import RGBColor
    box = dk.slot_shapes(s)["root"]
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(0xFF, 0x00, 0xFF)
    defects.append((n, "colour"))
    # 10. missing risks section: the spec has no risk slides
    defects.append((None, "structure"))
    prs.save(str(out_deck))
    return defects


def score(review_result, defects, baseline=None):
    """(found, false_alarms, details). A defect counts as found when a finding
    with its category sits on its slide (or deck-level for structure)."""
    found, matched = [], set()
    for d_slide, cat in defects:
        hit = None
        for k, f in enumerate(review_result["findings"]):
            if f["category"] == cat and (f["slide"] == d_slide):
                if cat == "structure" and "risk" not in f["message"].lower():
                    continue
                hit = k
                break
        if hit is not None:
            matched.add(hit)
            found.append((d_slide, cat))
    base = {(f["slide"], f["category"], f["message"]) for f in (baseline or {}).get("findings", [])}
    false = [f for k, f in enumerate(review_result["findings"])
             if k not in matched and f["severity"] in ("high", "medium") and f["category"] != "ai-pattern"
             and (f["slide"], f["category"], f["message"]) not in base]
    return found, false
