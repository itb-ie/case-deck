#!/usr/bin/env python3
"""Golden test material: a deck_spec.yaml for the Tajín (Mexico 24H) main
body, written from the real deck, plus its images.

The spec text is generated with the same selectors curate.py used, so titles
and numbers match the winning deck exactly. Images are copied from the
winning deck into tests/fixtures/tajin/img/ (git-ignored). Team photos are
never copied.

  python tests/make_golden.py --winning "<path to winning_cases>"
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import yaml

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "scripts"))
import deckkit as dk  # noqa: E402
import curate  # noqa: E402

FIX = HERE / "fixtures" / "tajin"

# (slide, shape name, nth) -> file name, for slides that are not a pattern source
EXTRA_IMAGES = {
    "s10_image": (10, "Picture 2", 1),
    "s32_image": (32, "Picture 2", 1),
    "s06_image": (6, "Picture 5", 1),
}


def extract(winning: Path):
    src = dk.open_deck(winning / curate.TAJIN)
    img_dir = FIX / "img"
    img_dir.mkdir(parents=True, exist_ok=True)
    values = {}
    for spec in curate.PATTERNS:
        fname, n = spec["source"]
        if fname != curate.TAJIN or spec.get("internal"):
            continue
        s = src.slides[n - 1]
        vals, taken, keep = {}, set(), []
        for slot, sd in spec["slots"].items():
            sd = sd if isinstance(sd, dict) else {"sel": sd}
            sh = curate.find(s, sd["sel"], taken)
            if sh is None:
                continue
            keep.append(sh)
            taken.add(sh._element)
            t = sd.get("type", "text")
            if t == "text" and sh.has_text_frame:
                paras = [p.text.strip() for p in sh.text_frame.paragraphs]
                while paras and not paras[-1]:
                    paras.pop()
                while paras and not paras[0] and slot != "target_label":
                    paras.pop(0)
                vals[slot] = paras[0] if len(paras) == 1 else paras
            elif t == "image":
                blip = sh._element.find(".//{%s}blip" % dk.NS["a"])
                part = s.part.related_part(blip.get("{%s}embed" % dk.R_NS))
                fn = f"s{n:02d}_{slot}.{part.partname.ext}"
                (img_dir / fn).write_bytes(part.blob)
                vals[slot] = f"img/{fn}"
        values[spec["id"]] = vals
    for key, (n, name, nth) in EXTRA_IMAGES.items():
        hits = [sh for sh in src.slides[n - 1].shapes if sh.name == name]
        sh = hits[nth - 1]
        blip = sh._element.find(".//{%s}blip" % dk.NS["a"])
        part = src.slides[n - 1].part.related_part(blip.get("{%s}embed" % dk.R_NS))
        (img_dir / f"{key}.{part.partname.ext}").write_bytes(part.blob)
        values.setdefault("_extra", {})[key] = f"img/{key}.{part.partname.ext}"
    return values


NEUTRAL = {"m1_name": "Presenter 1", "m2_name": "Presenter 2", "m3_name": "Presenter 3", "m4_name": "Presenter 4"}


def make_spec(v):
    x = v["_extra"]
    team = dict(NEUTRAL)  # the repo is shared with students: no real names
    v["section_divider"]["presenter"] = "Presenter 4"
    fw = v["framework"]

    def tracker(k):
        return {"pattern": "framework", "current": k, "section": "Strategy",
                "slots": {s: fw[s] for s in fw if s not in ("p1_num", "p2_num", "p3_num")}}

    slides = [
        {"pattern": "title", "slots": {**v["title"], **team}},
        {"pattern": "exec_summary", "section": "Executive Summary", "slots": v["exec_summary"]},
        {"pattern": "section_divider", "slots": v["section_divider"]},
        {"pattern": "company_position", "section": "Analysis", "slots": v["company_position"]},
        {"pattern": "trends", "section": "Analysis", "slots": v["trends"]},
        {"pattern": "hook_image", "slots": {"image": x["s06_image"]},
         "notes": "Hook: tell the human story behind the photo before showing data."},
        {"pattern": "matrix_takeaways", "section": "Analysis", "slots": v["matrix_takeaways"]},
        {"pattern": "three_cards", "section": "Analysis", "slots": {
            **v["three_cards"],
            "c1_text": ["## " + v["three_cards"]["c1_text"][0], ""] + v["three_cards"]["c1_text"][2:],
            "c2_text": ["## " + v["three_cards"]["c2_text"][0], ""] + v["three_cards"]["c2_text"][2:],
            "c3_text": ["## " + v["three_cards"]["c3_text"][0], ""] + v["three_cards"]["c3_text"][2:]},
         "takeaway": "Tajín can unite its **cultural heritage** and **bold innovation** to transform everyday flavor into a "
                     "**shareable, purpose-driven experience** that resonates with Gen Z and Gen Alpha"},
        {"pattern": "section_divider", "slots": {"section": "Strategy", "presenter": "Presenters 2 & 3",
                                                 "image": x["s10_image"]}},
        {"pattern": "prioritization", "section": "Strategy", "slots": v["prioritization"]},
        {"pattern": "strategy_tree", "section": "Strategy", "slots": v["strategy_tree"]},
        {"pattern": "framework", "section": "Strategy", "slots": v["framework"]},
        tracker(1),
        {"pattern": "pillar_overview", "section": "Sabores", "slots": v["pillar_overview"]},
        {"pattern": "initiative_steps", "section": "Sabores", "slots": v["initiative_steps"]},
        {"pattern": "initiative", "section": "Sabores", "slots": {
            **v["initiative"],
            "kpis": ["## Velocity uplift", "Units per store per week in first 8 weeks vs. control group (target 25%)", "",
                     "## Repeat rate", "% of first-time buyers repurchasing within 60-90 days (target 35%)", "",
                     "## Household penetration", "Incremental points gained in target cohorts within 6 months (target +2-3 points)"]},
         "takeaway": "Become a global everyday favorite by connecting emotionally with new generations worldwide, blending "
                     "**cultural authenticity** with **modern versatility** and **creative collaboration**, and staying true to Mexican roots"},
        tracker(2),
        {"pattern": "pillar_overview", "section": "Alma", "slots": {
            "title": "Tajín can improve its brand perception through community and sustainability initiatives, approaching Gen Z and Alpha consumer habits",
            "i1_name": ["Flavor co-creation", "**”Hecho Juntos”**"],
            "i2_name": ["Zero waste program", "**“Tierra Mojada”**"],
            "i3_name": ["Stories in packaging", "**“Hecho con Alma”**"],
            "i1_desc": "Customers will be able to upload suggestions for Tajín flavors they would like to try, and a contest will be held annually to launch **limited editions of the winners**",
            "i2_desc": "Chili and lime peels will be sent **back to producers**, to be reused in the growing of the soil for sustainable farming",
            "i3_desc": "Each Tajín package tells the story of where its ingredients originate and their supply chain, building **emotional connection and transparency**",
            "i1_case": "**Lay’s** invites fans worldwide to invent new chip flavors and vote on favorites. Winning ideas are produced, turning consumers into co-creators",
            "i2_case": "**Tricorbraun** cut landfill waste to nearly zero by recycling, reusing materials, and optimizing processes, reducing disposal costs",
            "i3_case": "**Ben & Jerry’s** embeds activism into packaging with flavors like Justice Remix’d, making storytelling drive sales and emotional loyalty",
            "i1_stat1": "14.4M", "i1_label1": "Flavor submissions",
            "i1_stat2": "8%", "i1_label2": "Sales increase 3 months after launch",
            "i2_stat1": "27T", "i2_label1": "CO2 avoided annually",
            "i3_stat1": "1.7x", "i3_label1": "Sales uplift",
            "i3_stat2": "1.5x", "i3_label2": "ROI higher than industry benchmarks",
            "source": "Source: Team Analysis, Syte, Houzz, Furniture Today"}},
        {"pattern": "initiative", "section": "Alma", "slots": {
            "title": "Build a sustainability image, drift away from greenwashing and reduce costs by establishing a circular waste reduction chain",
            "breakdown_head": "Breakdown of “Tierra Mojada”",
            "b1_text": "**Collect and process:** gather lime and chili waste and compost it into organic fertilizer",
            "b2_text": "**Return:** redistribute compost to partner farms, enriching soil and boosting yields",
            "b3_text": "**Measure and share:** track soil health and farmer savings, and tell the stories",
            "kpis": ["## Tons of chili and lime waste collected", "Percentage of the production waste that is repurposed", "",
                     "## Farmers cost savings", "Total cost savings in chemical fertilizers or pesticides", "",
                     "## Brand perception uplift", "Increase in customers' perception of Tajín's sustainability image"],
            "source": "Source: Deloitte Global Gen Z Survey 2024"}},
        tracker(3),
        {"pattern": "pillar_overview", "section": "Llegada", "slots": {
            "title": "Activate three growth channels engineered around virality to scale Tajín internationally",
            "i1_name": ["Popup Stands", "**”Carrito de Tajín”**"],
            "i2_name": ["Golden Ticket", "**“Entrada Tapatía”**"],
            "i3_name": ["Targeted e-commerce", "**“El Chile digital”**"],
            "i1_desc": "Limited-time **street-market pop-ups** that sample Tajín on fresh fruit to drive first taste and local buzz",
            "i2_desc": "On-pack codes unlock a chance to win a **behind-the-scenes vacation** to Guadalajara that celebrates local culture and Tajín’s transparent manufacturing",
            "i3_desc": "Pair Amazon “add-to-cart” bundles with **TikTok Shop creator storefronts** and shoppable UGC to convert inspiration into checkout",
            "i1_case": "**Marsh** scaled staffed sampling in supermarkets after strong results, proving demos convert trial into purchase",
            "i2_case": "**Raw** ran a “Golden Ticket” sweepstakes that sent 10 fans to tour its Spain factory, turning brand transparency into a community event",
            "i3_case": "**e.l.f.** fused creator-led TikTok Shop drops with optimized Amazon PDPs and bundles, channeling viral moments into same-session checkout",
            "i1_stat1": "68%", "i1_label1": "of shoppers persuaded by sampling",
            "i1_stat2": "6x", "i1_label2": "Sales during sampling events",
            "i2_stat1": "4.8M", "i2_label1": "views on TikTok",
            "i3_stat1": "56%", "i3_label1": "Lower ad costs",
            "i3_stat2": "77%", "i3_label2": "Quarterly sales growth",
            "source": "Source: Supermarket News, TikTok, Google Trends, TikTok for Business, E.L.F"}},
        {"pattern": "section_divider", "slots": {"section": "Impact", "presenter": "Presenter 1", "image": x["s32_image"]}},
        {"pattern": "driver_tree", "section": "Impact", "slots": v["driver_tree"]},
        {"pattern": "financials", "section": "Impact", "slots": {
            **v["financials"],
            "chart": {"categories": ["2024", "2025E", "2026E", "2027E", "2028E", "2029E", "2030E"],
                      "series": {"Base growth": [300, 345, 397, 456, 525, 603, 694],
                                 "Penetration": [0, 10, 20, 30, 45, 60, 75],
                                 "Frequency": [0, 8, 16, 24, 30, 36, 45],
                                 "Avg. price": [0, 3, 6, 9, 12, 15, 18]}}},
         "notes": "Chart values approximate the 2025 chart for the golden test; the original model is in the appendix."},
        {"pattern": "risk_matrix", "section": "Impact", "top_risks": [1, 4, 6], "slots": v["risk_matrix"]},
        {"pattern": "mitigation", "section": "Impact", "slots": v["mitigation"]},
        {"pattern": "roadmap", "section": "Impact", "slots": v["roadmap"],
         "takeaway": "Create a **scaled roll-out** of the initiatives with consistent brand guidelines and **transparent sustainability practices** "
                     "to capitalize on the **~130m revenue addition** by 2030 opportunity"},
        {"pattern": "exec_summary", "section": "Executive Summary", "slots": v["exec_summary"]},
        {"pattern": "closing", "slots": {**team, "hero_image": v["title"]["hero_image"],
                                         "client_logo": v["title"]["client_logo"]}},
        {"pattern": "appendix_divider", "appendix": True, "slots": {"client_logo": v["appendix_divider"].get("client_logo", v["title"]["client_logo"])}},
        {"pattern": "appendix_index", "appendix": True, "id": "index", "section": "Appendix"},
        {"pattern": "text_content", "appendix": True, "group": "Strategy", "id": "case_questions", "section": "Appendix", "slots": {
            "title": "Appendix. Strategies - We tailored our initiatives in each strategy to concretely address the issues proposed by Tajín",
            "body": ["**Product:** improve quality and uses of Yaya; expand internationally by adapting flavors; create snacks to penetrate untapped markets",
                     "**Brand:** include traceability in packaging; develop a zero-waste program; co-create limited edition flavors with the community",
                     "**Channel:** approach untapped customers through e-commerce; create hybrid experiences through golden tickets; attract new customers through pop-up stands"],
            "source": "Source: Team Analysis"}},
        {"pattern": "chart_takeaways", "appendix": True, "group": "Financials", "id": "fin_model", "section": "Appendix", "slots": {
            "title": "Appendix. Financial model - revenue builds from penetration and frequency gains",
            "chart_head": "Revenue with strategies vs base case (USD m)",
            "unit": "Revenue, USD m",
            "chart": {"categories": ["2025E", "2026E", "2027E", "2028E", "2029E", "2030E"],
                      "series": {"Revenue with strategies": [366, 439, 519, 612, 714, 832]}},
            "takeaways": ["Base growth of 15% a year from historicals", "",
                          "Strategies add ~19% on top of the base case by 2030", "",
                          "Total cost of ~USD 47m over 6 years gives a 2.88x ROI"],
            "source": "Source: Team Analysis"}},
        {"pattern": "harvey_table", "appendix": True, "group": "Market", "id": "competitors", "section": "Appendix", "slots": {
            "title": "Appendix. Competitive analysis - Tajín leads on cultural authenticity but trails on digital reach",
            "table": [[" ", "Authenticity", "Gen Z reach", "Digital", "Range", "Price"],
                      ["Tajín", "", "", "", "", ""], ["Valentina", "", "", "", "", ""], ["Chamoy brands", "", "", "", "", ""],
                      ["Global spice brands", "", "", "", "", ""]],
            "h_r1_c1": 4, "h_r1_c2": 2, "h_r1_c3": 2, "h_r1_c4": 3, "h_r1_c5": 3,
            "h_r2_c1": 3, "h_r2_c2": 2, "h_r2_c3": 1, "h_r2_c4": 2, "h_r2_c5": 4,
            "h_r3_c1": 3, "h_r3_c2": 3, "h_r3_c3": 1, "h_r3_c4": 1, "h_r3_c5": 3,
            "h_r4_c1": 1, "h_r4_c2": 2, "h_r4_c3": 4, "h_r4_c4": 4, "h_r4_c5": 2,
            "takeaway": "Tajín's edge is authenticity; digital reach is the gap the Llegada pillar closes",
            "source": "Source: Team Analysis [ASSUMPTION: illustrative ratings for the golden test]"}},
    ]
    return {"deck": {"title": "Mejorando todo con Tajín", "client": "Tajín",
                     "client_logo": v["title"]["client_logo"], "output": "tajin_rebuild.pptx"},
            "slides": slides}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--winning", required=True)
    args = ap.parse_args()
    values = extract(Path(args.winning))
    spec = make_spec(values)
    FIX.mkdir(parents=True, exist_ok=True)
    with open(FIX / "deck_spec.yaml", "w", encoding="utf-8") as f:
        f.write("# Golden test: Tajín (Mexico 24H) main body rebuilt from its own content.\n"
                "# Generated by tests/make_golden.py; images come from the winning deck.\n")
        yaml.safe_dump(spec, f, sort_keys=False, allow_unicode=True, width=120)
    print("wrote", FIX / "deck_spec.yaml", "with", len(spec["slides"]), "slides")


if __name__ == "__main__":
    main()
