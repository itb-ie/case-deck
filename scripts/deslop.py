#!/usr/bin/env python3
"""Remove AI writing patterns from a deck, with the team's approval.

Two steps, so nothing changes without a human yes:

  1. python deslop.py propose deck.pptx > worklist.json
     Lists every flagged paragraph (slide, box, text, what was flagged,
     the detector's suggestion). Claude drafts rewrites from it into an
     edits file and shows them to the team as a before/after table.

  2. python deslop.py apply deck.pptx edits.yaml -o deck_v2.pptx
     Applies only the approved edits, run by run (fonts, sizes, colours and
     bold stay as they were). Every number, currency, percentage, date,
     proper noun and “branded name” in an edited box must be identical
     before and after, or nothing is saved. Then re-runs the AI check and
     prints the score before and after.

edits.yaml:
  edits:
    - slide: 2
      find: "A robust, game-changing approach that delves into bold new flavors"
      replace: "Bold new flavors for Gen Z, tested first in Mexico City"
      approved: true          # only approved edits are applied
"""
from __future__ import annotations

import argparse
import copy
import json
import re
import sys
from collections import Counter
from pathlib import Path

import yaml

sys.path.insert(0, str(Path(__file__).resolve().parent))
import deckkit as dk  # noqa: E402
import ai_check  # noqa: E402

A = dk.NS["a"]
NUMBER = re.compile(r"[-+~]?(?:[$€£¥]\s?)?\d[\d.,]*\s?(?:%|x\b|bn\b|m\b|k\b|M\b|B\b|pp\b)?")
QUOTED = re.compile(r"[“\"«]([^”\"»]{2,60})[”\"»]")
CURRENCY = re.compile(r"\b(?:USD|EUR|RON|MXN|GBP|CHF)\b|[$€£¥]")
STOP = set("""A An And As At But By For From If In Into It Its Of On Or So The This That These Those To We Our Us
With Without Yet Each Every Moreover Furthermore Additionally Ultimately Overall Finally However Also Then
Their They Your You He She Not No Yes""".split())


# --------------------------------------------------------------------------
# Protected content
# --------------------------------------------------------------------------

def protected(text):
    """Numbers, currencies, dates, proper nouns and quoted names, as a multiset."""
    items = Counter()
    for m in NUMBER.finditer(text):
        items["num:" + re.sub(r"\s", "", m.group(0)).lstrip("+~")] += 1
    for m in CURRENCY.finditer(text):
        items["cur:" + m.group(0)] += 1
    for m in QUOTED.finditer(text):
        items["name:" + m.group(1).strip()] += 1
    for sentence in re.split(r"(?<=[.!?:])\s+|\n", text):
        all_words = re.findall(r"[\wÀ-ÿ’'&-]+", sentence)
        words = re.findall(r"[A-ZÀ-Þ][\wÀ-ÿ’'&.-]*", sentence)
        if all_words and len(words) / len(all_words) >= 0.6 and len(all_words) <= 8:
            continue  # Title Case heading: capitals carry no meaning here
        first = re.match(r"\W*([\wÀ-ÿ’']+)", sentence)
        first = first.group(1) if first else None
        for w in words:
            w = re.sub(r"[’']s$", "", w).rstrip(".’'")
            if w == first or w in STOP or len(w) < 2:
                continue
            items["noun:" + w] += 1
    return items


# --------------------------------------------------------------------------
# Worklist
# --------------------------------------------------------------------------

def propose(path, client=None):
    res = ai_check.check(path, client=client)
    prs = dk.open_deck(path)
    work = []
    flagged = {}
    for f in res["findings"]:
        m = re.match(r"slide (\d+)$", f["where"])
        if m:
            flagged.setdefault(int(m.group(1)), []).append(f)
    for n, fs in sorted(flagged.items()):
        s = prs.slides[n - 1]
        for sh in dk.iter_shapes(s.shapes):
            if sh.name.startswith(ai_check.SKIP_SHAPES) or not sh.has_text_frame:
                continue
            for p in sh.text_frame.paragraphs:
                t = "".join(r.text for r in p.runs)
                hits = [f for f in fs if f["text"] and f["text"].lower() in t.lower()]
                if hits:
                    work.append({"slide": n, "box": sh.name.replace("slot:", ""), "text": t,
                                 "flagged": sorted({h["text"] for h in hits}),
                                 "suggestions": sorted({h["suggestion"] for h in hits if h["suggestion"]}),
                                 "protected": sorted(protected(t).elements())})
    return {"deck": str(path), "score": res["score"], "label": res["label"], "items": work}


# --------------------------------------------------------------------------
# Applying edits run by run
# --------------------------------------------------------------------------

def _replace_in_paragraph(p_el, old, new):
    """Replace old with new inside one a:p, across runs. The new text takes
    the formatting of the run where the match starts; other runs keep theirs.
    Returns True when replaced."""
    runs = p_el.findall(f"{{{A}}}r")
    texts = [r.findtext(f"{{{A}}}t") or "" for r in runs]
    full = "".join(texts)
    idx = full.find(old)
    if idx < 0:
        low = full.lower().find(old.lower())
        if low < 0:
            return False
        idx = low
    end = idx + len(old)
    pos = 0
    first = True
    for r, t in zip(runs, texts):
        r_start, r_end = pos, pos + len(t)
        pos = r_end
        if r_end <= idx or r_start >= end:
            continue
        a_ = max(idx, r_start) - r_start
        b_ = min(end, r_end) - r_start
        t_el = r.find(f"{{{A}}}t")
        if first:
            t_el.text = t[:a_] + new + t[b_:]
            first = False
        else:
            remaining = t[:a_] + t[b_:]
            if remaining:
                t_el.text = remaining
            else:
                p_el.remove(r)
    return True


def _run_props(shape):
    out = []
    for r in shape._element.iter(f"{{{A}}}r"):
        rpr = r.find(f"{{{A}}}rPr")
        out.append(dk.etree.tostring(rpr, method="c14n") if rpr is not None else b"")
    return out


def apply(path, edits_path, out=None, client=None):
    path = Path(path)
    edits = yaml.safe_load(Path(edits_path).read_text(encoding="utf-8")) or {}
    edits = [e for e in edits.get("edits", []) if e.get("approved")]
    prs = dk.open_deck(path)
    before_score = ai_check.check(path, client=client)["score"]
    problems, applied, skipped = [], 0, []
    for e in edits:
        n = int(e["slide"])
        s = prs.slides[n - 1]
        done = False
        for sh in dk.iter_shapes(s.shapes):
            if not sh.has_text_frame or (e.get("box") and sh.name.replace("slot:", "") != e["box"]):
                continue
            text_before = sh.text_frame.text
            if e["find"].lower() not in text_before.lower():
                continue
            props_before = _run_props(sh)
            for p in sh.text_frame.paragraphs:
                if _replace_in_paragraph(p._p, e["find"], e["replace"]):
                    done = True
                    break
            text_after = sh.text_frame.text
            pb, pa = protected(text_before), protected(text_after)
            if pb != pa:
                lost = sorted((pb - pa).elements())
                added = sorted((pa - pb).elements())
                problems.append({"slide": n, "box": sh.name, "find": e["find"][:60],
                                 "lost": lost, "added": added})
            # formatting: every remaining run keeps a property set it had before
            props_after = _run_props(sh)
            if any(p not in props_before for p in props_after):
                problems.append({"slide": n, "box": sh.name, "find": e["find"][:60], "formatting": "changed"})
            if done:
                applied += 1
                break
        if not done:
            skipped.append(e["find"][:60])
    if problems:
        return {"saved": False, "problems": problems, "applied": applied, "skipped": skipped}
    out = Path(out) if out else path.with_name(path.stem + "_v2.pptx")
    prs.save(str(out))
    after_score = ai_check.check(out, client=client)["score"]
    return {"saved": True, "output": str(out), "applied": applied, "skipped": skipped,
            "score_before": before_score, "score_after": after_score}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    p1 = sub.add_parser("propose")
    p1.add_argument("deck")
    p1.add_argument("--client")
    p2 = sub.add_parser("apply")
    p2.add_argument("deck")
    p2.add_argument("edits")
    p2.add_argument("-o", "--out")
    p2.add_argument("--client")
    args = ap.parse_args()
    if args.cmd == "propose":
        print(json.dumps(propose(args.deck, args.client), indent=1, ensure_ascii=False))
        return
    r = apply(args.deck, args.edits, args.out, args.client)
    if not r["saved"]:
        print("Nothing saved: these edits would change protected content or formatting:")
        for p in r["problems"]:
            print(" ", json.dumps(p, ensure_ascii=False))
        sys.exit(2)
    print(f"Saved {r['output']}: {r['applied']} edit(s) applied. AI-pattern score {r['score_before']} -> {r['score_after']}.")
    for s in r["skipped"]:
        print("  not found, skipped:", s)


if __name__ == "__main__":
    main()
