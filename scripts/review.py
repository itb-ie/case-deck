#!/usr/bin/env python3
"""Review a deck: mechanical checks, ranked slide by slide, with a fix for
each finding. Claude adds the judge view on top (see SKILL.md, mode Review).

Usage:
  python review.py deck.pptx                    # Markdown report to stdout
  python review.py deck.pptx --json             # findings as JSON
  python review.py deck.pptx --client "Tajín"   # name of the case company
  python review.py deck.pptx --no-ai            # skip the AI-pattern check
  python review.py deck.pptx --out review.md    # write the report to a file

Works on decks built with case-deck and on hand-made decks (it then infers
slide roles from titles and section tags).
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

import yaml

sys.path.insert(0, str(Path(__file__).resolve().parent))
import deckkit as dk  # noqa: E402
from build import slide_meta  # noqa: E402

SEV_ORDER = {"high": 0, "medium": 1, "low": 2}
PAST_CLIENTS = ["MOL", "Tajín", "Tajin", "Artex", "Atlas FC", "Atlas", "RIVUS", "IULIUS", "Muebles"]
EXEMPT_TITLE_PATTERNS = {"title", "exec_summary", "section_divider", "hook_image", "closing", "appendix_divider",
                         "appendix_index", "framework", "strategy_tree", "navigator", "takeaway_overlay"}
CONVENTION_TITLES = re.compile(r"overview of our approach|our perspective|^appendix|index|agenda|thank you|questions\??$|"
                               r"executive summary|^a \w+ strategy to guide", re.I)
VERBS = set("""
accelerate achieve activate add address adopt align allow attract avoid become boost build capture change close
combine connect convert create cut deliver deploy design develop differentiate double drive earn elevate embed enable
enhance enter establish estimate expand extend focus gain generate grow help improve increase integrate introduce
invest keep launch lead leads leverage link make maximize meet move need needs nurture offer open optimize outperform
partner phase pilot position prioritize produce protect prove reach reduce reinforce remain require requires retain
reuse scale secure select selected sell serve shift show shows stay strengthen support target test transform turn
unite unlock use win wins is are was were has have can could will would should must lead leads drives grows makes
brings gives takes keeps opens sets puts turns gets helps lets creates builds delivers generates produces reaches
""".split())
MONEY = re.compile(r"(?P<cur>USD|EUR|RON|MXN|GBP|CHF|\$|€|£)?\s?(?P<approx>~)?(?P<num>\d+(?:[.,]\d+)?)\s?(?P<unit>bn|b|m|k|million|billion)\b(?!\w)", re.I)
RATIO = re.compile(r"(?P<approx>~)?(?P<num>\d+(?:\.\d+)?)\s?x\b", re.I)
HAS_NUMBER = re.compile(r"(?<![\w.])(?:[$€£]\s?)?\d[\d.,]*\s?(?:%|x\b|m\b|bn\b|k\b|M\b|B\b)|\b\d{2,}(?:[.,]\d+)?\b")
YEARLIKE = re.compile(r"^(19|20)\d{2}E?$")
PLACEHOLDER = re.compile(r"\[(?!ASSUMPTION|SOURCE\?)[^\]\n]{1,80}\]")
OPEN_TAG = re.compile(r"\[(ASSUMPTION|SOURCE\?)[^\]]*\]")


def load_catalog():
    try:
        with open(dk.CATALOG, encoding="utf-8") as f:
            return yaml.safe_load(f)
    except OSError:
        return {}


def finding(slide, sev, cat, msg, fix):
    return {"slide": slide, "severity": sev, "category": cat, "message": msg, "fix": fix}


def client_of(prs, explicit=None):
    """The case company: given, stored by build.py, or guessed as the
    past-case name the deck mentions most (hand-made decks)."""
    if explicit:
        return explicit
    kw = prs.core_properties.keywords or ""
    m = re.search(r"client=([^;]+)", kw)
    if m:
        return m.group(1).strip()
    text = "\n".join(dk.slide_text(s) for s in prs.slides)
    counts = {n: len(re.findall(rf"(?<![\w-]){re.escape(n)}(?![\w-])", text)) for n in PAST_CLIENTS}
    best = max(counts, key=counts.get)
    return best if counts[best] >= 5 else None


def body_shapes(slide):
    for sh in dk.iter_shapes(slide.shapes):
        if sh.name.startswith(("nav:", "pattern:")):
            continue
        yield sh


def role_of(slide, meta, title):
    p = meta.get("pattern")
    if p:
        return p
    t = title.lower()
    if "overview of our approach" in t or "executive summary" in t:
        return "exec_summary"
    if t.startswith("appendix. index"):
        return "appendix_index"
    if t == "appendix":
        return "appendix_divider"
    if re.search(r"\brisk", t):
        return "risk_matrix"
    if re.search(r"phase|roadmap|implementation|timeline", t):
        return "roadmap"
    if re.search(r"revenue|roi\b|payback|financial|npv|margin|ebitda", t):
        return "financials"
    return ""


def title_of(slide):
    ph = dk.placeholder(slide, 0)
    if ph is not None and ph.has_text_frame:
        return ph.text_frame.text.strip()
    return ""


def is_topic_title(title):
    words = re.findall(r"[A-Za-zÀ-ÿ’']+", title)
    if not words:
        return False
    if re.search(r"\d", title):
        return False
    lower = [w.lower().strip("’'") for w in words]
    if any(w in VERBS for w in lower):
        return False
    if any(w.endswith(("ing", "ed")) and len(w) > 4 for w in lower[1:]):
        return False
    return len(words) <= 7


def money_values(text):
    out = []
    for m in MONEY.finditer(text):
        num = float(m.group("num").replace(",", "."))
        unit = m.group("unit").lower()
        mult = {"k": 0.001, "m": 1, "million": 1, "b": 1000, "bn": 1000, "billion": 1000}[unit]
        cur = (m.group("cur") or "").upper().replace("$", "USD").replace("€", "EUR").replace("£", "GBP")
        out.append({"value": num * mult, "cur": cur, "approx": bool(m.group("approx")), "raw": m.group(0).strip()})
    return out


def review(path, client=None, run_ai=True, catalog=None):
    path = Path(path)
    prs = dk.open_deck(path)
    catalog = catalog or load_catalog()
    client = client_of(prs, client)
    house_colors = set(dk.PALETTE) | set(catalog.get("house_colors", []))
    W, H = prs.slide_width, prs.slide_height
    findings, slides_info = [], []

    for i, s in enumerate(prs.slides, 1):
        meta = slide_meta(s)
        title = title_of(s)
        role = role_of(s, meta, title)
        text = dk.slide_text(s)
        hidden = s._element.get("show") == "0"
        slides_info.append({"n": i, "role": role, "title": title, "text": text, "meta": meta, "hidden": hidden,
                            "appendix": bool(meta.get("appendix")) or role in ("appendix_divider", "appendix_index"),
                            "takeaway": bool(meta.get("takeaway_of"))})
        if meta.get("takeaway_of") or meta.get("nav_generated"):
            continue

        # titles
        if title and role not in EXEMPT_TITLE_PATTERNS and not CONVENTION_TITLES.search(title) \
                and not dk.LEFTOVER.search(title):
            if is_topic_title(title):
                findings.append(finding(i, "medium", "title", f"Topic title, not an action title: “{title}”",
                                        "Write the takeaway as a sentence with a verb, and a number if you have one."))
            if len(title) > 150:
                findings.append(finding(i, "medium", "title", f"Title is {len(title)} characters (more than 2 lines).",
                                        "Cut it to one idea; move detail to the subtitle."))

        # placeholders, open tags, leftovers
        for m in PLACEHOLDER.finditer(text):
            findings.append(finding(i, "high", "unfilled", f"Placeholder text left: “{m.group(0)}”",
                                    "Fill it in or delete the box."))
        for m in OPEN_TAG.finditer(text):
            findings.append(finding(i, "medium", "open-tag", f"Open tag: “{m.group(0)[:70]}”",
                                    "Resolve it: find the source, or state the assumption in the appendix."))
        hook_tbd = bool(re.search(r"\bHOOK\s+TBD\b", text, re.I))
        if hook_tbd:
            findings.append(finding(i, "high", "leftover", "“HOOK TBD” is still on the slide.", "Write the hook or remove the title."))
        for m in dk.LEFTOVER.finditer(text):
            if hook_tbd and m.group(0).upper() == "TBD":
                continue
            findings.append(finding(i, "high", "leftover", f"Leftover text: “{m.group(0)}”",
                                    "Replace or delete it."))
        for name in PAST_CLIENTS:
            if client and (name.lower() in client.lower() or client.lower() in name.lower()):
                continue
            if re.search(rf"(?<![\w-]){re.escape(name)}(?![\w-])", text):
                findings.append(finding(i, "high", "other-case", f"Mentions “{name}”, a company from another case.",
                                        "Check whether this slide was copied from an old deck; replace the content."))
                break
        grey = [sh.name.replace("slot:", "") for sh in body_shapes(s) if sh.shape_type == 13 and dk.is_placeholder_image(sh)]
        if grey:
            findings.append(finding(i, "medium", "image", f"{len(grey)} image placeholder(s) still grey: {', '.join(grey)}.",
                                    "Add real photos (in the spec or in PowerPoint) or delete the boxes."))
        for sh in body_shapes(s):
            if sh.is_placeholder and sh.has_text_frame and not sh.text_frame.text.strip() and sh.placeholder_format.idx in (0, 1):
                findings.append(finding(i, "medium", "empty", "Empty title or subtitle placeholder.", "Fill it or delete it."))

        # numbers and sources
        body_text = "\n".join(dk.shape_text(sh) for sh in body_shapes(s)
                              if not sh.is_placeholder or sh.placeholder_format.idx not in (17,))
        nums = [m.group(0) for m in HAS_NUMBER.finditer(body_text)
                if not YEARLIKE.match(m.group(0).strip()) and not re.fullmatch(r"0?\d", m.group(0).strip())]
        has_source = bool(re.search(r"^\s*Sources?:", text, re.M | re.I))
        if nums and not has_source and role not in ("title", "closing", "section_divider", "appendix_divider",
                                                     "hook_image", "appendix_index", "navigator"):
            findings.append(finding(i, "medium", "source", f"Numbers ({', '.join(nums[:3])}) with no Source line.",
                                    "Add “Source: …” at the bottom (Team Analysis counts for your own estimates)."))

        # visual checks from shape geometry (top-level shapes: group children
        # use the group's own coordinate space)
        top_level = [x._element for x in s.shapes]
        for sh in body_shapes(s):
            if getattr(sh, "has_chart", False) and sh.has_chart:
                continue
            b = dk.bbox(sh)
            in_group = not any(sh._element is e for e in top_level)
            rotated = getattr(sh, "rotation", 0) not in (0, 0.0)
            if sh.has_text_frame and sh.text_frame.text.strip() and not in_group and not rotated:
                if b[2] > W + 0.1 * dk.IN or b[3] > H + 0.1 * dk.IN or b[0] < -0.1 * dk.IN or b[1] < -0.1 * dk.IN:
                    findings.append(finding(i, "medium", "layout", f"Text box “{dk.shape_text(sh)[:30]}” sits partly outside the slide.",
                                            "Move it back inside the slide."))
            if sh.has_text_frame and sh.text_frame.text.strip():
                for size in dk.font_sizes(sh):
                    if size < 9:
                        findings.append(finding(i, "medium", "font", f"{size:g} pt text: “{dk.shape_text(sh)[:30]}”",
                                                "Use 9 pt or more; judges read from a distance."))
                        break
                over = None if (in_group or rotated) else overflow(sh, s)
                if over:
                    findings.append(finding(i, "medium", "overflow", f"Text probably overflows its box: “{dk.shape_text(sh)[:40]}” ({over})",
                                            "Shorten the text or make the box bigger; check the render."))
            for el in sh._element.iter("{%s}srgbClr" % dk.NS["a"]):
                val = (el.get("val") or "").upper()
                if val and val not in house_colors:
                    findings.append(finding(i, "low", "colour", f"Colour #{val} is not in the house palette.",
                                            "Use navy 061F32, cyan 22A3DF, blue 3478AB, dark blue 1D4769 or gold F8BC3B."))
                    break
            fonts = {el.get("typeface") for el in sh._element.iter()
                     if str(el.tag).endswith("}latin") and el.get("typeface") and not el.get("typeface").startswith("+")}
            bad_fonts = {f for f in fonts if f not in ("Arial", "Arial Narrow") and f not in catalog.get("house_fonts", [])}
            if bad_fonts and sh.has_text_frame and sh.text_frame.text.strip():
                findings.append(finding(i, "low", "font", f"Font {', '.join(sorted(bad_fonts))} in “{dk.shape_text(sh)[:30]}”.",
                                        "Use Arial, like the template."))
        ph = dk.placeholder(s, 0)
        if ph is not None and abs((ph.left or 0) - int(0.61 * dk.IN)) > int(0.15 * dk.IN):
            findings.append(finding(i, "low", "layout", "Title is not in the standard position.",
                                    "Reset the slide (Home > Reset) so titles line up across slides."))

    findings += structure_checks(slides_info)
    findings += number_checks(slides_info)

    ai = None
    if run_ai:
        try:
            import ai_check
            ai = ai_check.check(path, client=client)
            bad, rate = ai_check.over_threshold(ai)
            per_slide = {}
            for f in ai["findings"]:
                m = re.match(r"slide (\d+)$", f["where"])
                if m:
                    per_slide.setdefault(int(m.group(1)), []).append(f)
            for n, fs in per_slide.items():
                words = ", ".join(sorted({f["text"] or f["type"] for f in fs})[:6])
                sev = "medium" if any(f["severity"] == "high" for f in fs) else "low"
                findings.append(finding(n, sev, "ai-pattern", f"{len(fs)} AI-writing pattern(s): {words}",
                                        "Run the de-slop mode or rewrite these phrases in your own words."))
            ai = {"score": ai["score"], "label": ai["label"], "rate_per_10_slides": round(rate, 1), "above_limit": bad}
        except Exception as e:  # detector missing etc.
            ai = {"error": str(e)}

    findings.sort(key=lambda f: (SEV_ORDER[f["severity"]], f["slide"] or 0))
    return {"path": str(path), "client": client, "slides": len(slides_info), "findings": findings,
            "titles": [(x["n"], x["title"]) for x in slides_info if x["title"] and not x["takeaway"] and not x["hidden"]],
            "ai": ai, "judge_pack": judge_pack(slides_info)}


def overflow(sh, slide=None):
    """Estimate whether text needs more lines than the box holds, or (for
    boxes that grow with their text) whether it grows into the next box."""
    tf = sh.text_frame
    body = tf._txBody.find("{%s}bodyPr" % dk.NS["a"])
    if body is not None and body.find("{%s}normAutofit" % dk.NS["a"]) is not None:
        return None  # PowerPoint shrinks the text to fit
    pt = dk.first_font_pt(sh, 14.0)
    if not sh.width or not sh.height or pt <= 0:
        return None
    if body is not None and body.get("wrap") == "none":
        return None  # text does not wrap
    if len(tf.text.split()) <= 3:
        return None  # short labels: estimates are unreliable
    lines = dk.estimate_lines(tf.text, sh.width - int(0.2 * dk.IN), pt, char_w=0.5)
    need = lines * pt * 1.2
    have = dk.Emu(sh.height).pt
    grows = body is not None and body.find("{%s}spAutoFit" % dk.NS["a"]) is not None
    if grows or need > have:
        bottom_emu = (sh.top or 0) + int(max(need, have) * 12700)
        if bottom_emu > 7.4 * dk.IN:
            return f"grows to {bottom_emu / dk.IN:.1f}\" from the top"
        if slide is not None and need > have * 1.3 and lines >= 3:
            left, right = sh.left or 0, (sh.left or 0) + (sh.width or 0)
            my_bottom = (sh.top or 0) + (sh.height or 0)
            for other in slide.shapes:
                if other._element is sh._element or not other.has_text_frame or not other.text_frame.text.strip():
                    continue
                if other.name.startswith(("nav:", "pattern:")):
                    continue
                o_top = other.top or 0
                o_left, o_right = other.left or 0, (other.left or 0) + (other.width or 0)
                if o_top >= my_bottom - int(0.05 * dk.IN) and o_top < bottom_emu - int(0.05 * dk.IN) \
                        and min(right, o_right) - max(left, o_left) > 0.3 * (sh.width or 1):
                    return f"about {lines} lines; it runs into “{other.text_frame.text.strip()[:25]}” below"
        if grows:
            return None
    if need > have * 1.6 and lines > 2:
        return f"about {lines} lines for a box that holds {max(1, int(have / (pt * 1.2)))}"
    return None


def structure_checks(info):
    out = []
    main = [x for x in info if not x["appendix"] and not x["takeaway"] and not x["hidden"]]
    roles = [x["role"] for x in main]
    titles = " ".join(x["title"].lower() for x in main)
    if "exec_summary" not in roles:
        out.append(finding(None, "high", "structure", "No executive summary (answer first).",
                           "Add an exec_summary slide as slide 2: the question, your pillars, one recommendation each, the headline impact."))
    elif roles.count("exec_summary") < 2:
        out.append(finding(None, "low", "structure", "The executive summary is not repeated at the end.",
                           "Repeat it just before the closing slide; judges score what they remember."))
    if not ({"financials", "driver_tree", "chart_takeaways"} & set(roles)) and not re.search(r"revenue|roi|payback|npv", titles):
        out.append(finding(None, "high", "structure", "No financial impact slide.",
                           "Add a driver tree and a financials slide (revenue, ROI, payback)."))
    if not ({"risk_matrix", "mitigation"} & set(roles)) and "risk" not in titles:
        out.append(finding(None, "high", "structure", "No risks slide.",
                           "Add a risk matrix and a mitigation slide: judges always ask what could go wrong."))
    if "roadmap" not in roles and not re.search(r"phase|roadmap|timeline|implementation", titles):
        out.append(finding(None, "high", "structure", "No implementation roadmap.",
                           "Add a roadmap: pilot first, then scale, with months and owners."))
    if not any(x["appendix"] for x in info):
        out.append(finding(None, "medium", "structure", "No appendix.",
                           "Add an appendix (divider + index + backup slides) for the Q&A."))
    if not any(x["role"] == "hook_image" for x in info) and not re.search(r"\bhook\b", titles):
        out.append(finding(None, "low", "structure", "No hook slide found.",
                           "Consider a full-bleed photo after the company context, with the story told out loud."))
    return out


def number_checks(info):
    """Headline figures in the exec summary and closing must match the financials."""
    out = []
    summary = [x for x in info if x["role"] == "exec_summary" and not x["takeaway"]]
    fin = [x for x in info if x["role"] in ("financials", "driver_tree", "chart_takeaways") and not x["takeaway"]]
    fin_money = [v for x in fin for v in money_values(x["text"])]
    fin_ratio = [float(m.group("num")) for x in fin for m in RATIO.finditer(x["text"])]
    if not fin_money:
        return out
    for x in summary:
        for v in money_values(x["text"]):
            same = [f for f in fin_money if (f["cur"] == v["cur"] or not f["cur"] or not v["cur"])]
            if not same:
                continue
            tol = 0.06 if (v["approx"] or any(f["approx"] for f in same)) else 0.02
            if not any(abs(f["value"] - v["value"]) <= tol * max(f["value"], v["value"]) for f in same):
                near = min(same, key=lambda f: abs(f["value"] - v["value"]))
                out.append(finding(x["n"], "high", "numbers",
                                   f"“{v['raw']}” here, but the financials say “{near['raw']}”.",
                                   "Use one figure everywhere; judges compare the summary with the model."))
        for m in RATIO.finditer(x["text"]):
            r = float(m.group("num"))
            if fin_ratio and not any(abs(r - f) <= 0.05 * max(r, f) for f in fin_ratio):
                out.append(finding(x["n"], "high", "numbers", f"ROI/ratio “{m.group(0)}” does not match the financials.",
                                   "Use the same ratio as the financials slide."))
    return out


def judge_pack(info):
    """What Claude needs for the judge view: storyline and key numbers."""
    lines = []
    for x in info:
        if x["takeaway"] or x["hidden"]:
            continue
        nums = sorted({m.group(0).strip() for m in HAS_NUMBER.finditer(x["text"]) if not YEARLIKE.match(m.group(0).strip())})[:8]
        lines.append({"slide": x["n"], "role": x["role"] or "?", "appendix": x["appendix"], "title": x["title"],
                      "numbers": nums})
    return lines


def to_markdown(r):
    out = [f"# Review: {Path(r['path']).name}", "",
           f"{r['slides']} slides" + (f", client: {r['client']}" if r["client"] else "") + ".", ""]
    if r["ai"]:
        if "error" in r["ai"]:
            out.append(f"AI-pattern check not run: {r['ai']['error']}")
        else:
            out.append(f"AI-pattern score {r['ai']['score']} ({r['ai']['label']}), {r['ai']['rate_per_10_slides']} findings per 10 slides"
                       + (" — above the limit." if r["ai"]["above_limit"] else ", within the limit.").replace(" —", ":"))
        out.append("")
    counts = {k: sum(1 for f in r["findings"] if f["severity"] == k) for k in SEV_ORDER}
    out.append(f"Findings: {counts['high']} high, {counts['medium']} medium, {counts['low']} low.")
    out.append("")
    out.append("## Findings (most severe first)")
    out.append("")
    out.append("| # | Slide | Severity | What | Fix |")
    out.append("|---|---|---|---|---|")
    for k, f in enumerate(r["findings"], 1):
        out.append(f"| {k} | {f['slide'] or 'deck'} | {f['severity']} | {f['message']} | {f['fix']} |")
    out.append("")
    out.append("## Read only the titles")
    out.append("")
    out.append("The titles alone should tell the whole story:")
    out.append("")
    for n, t in r["titles"]:
        out.append(f"{n}. {t}")
    return "\n".join(out)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("deck")
    ap.add_argument("--client")
    ap.add_argument("--no-ai", action="store_true")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--out")
    args = ap.parse_args()
    r = review(args.deck, args.client, not args.no_ai)
    text = json.dumps(r, indent=1, ensure_ascii=False, default=str) if args.json else to_markdown(r)
    if args.out:
        Path(args.out).write_text(text, encoding="utf-8")
        print(f"wrote {args.out}")
    else:
        print(text)


if __name__ == "__main__":
    main()
