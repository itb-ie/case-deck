#!/usr/bin/env python3
"""AI-pattern check for decks and text, built on the avoid-ai-writing
detector (github.com/conorbronsdon/avoid-ai-writing, MIT, Conor Bronsdon),
plus a small case-deck layer the general detector does not cover.

Usage:
  python ai_check.py deck.pptx                 # report, exit 1 above threshold
  python ai_check.py notes.md storyline.md     # Markdown / text files
  python ai_check.py deck.pptx --json          # machine-readable
  python ai_check.py deck.pptx --threshold 20  # custom limit (findings per deck)

Text between <!-- avoid-ai-writing:ignore-start --> and
<!-- avoid-ai-writing:ignore-end --> is skipped (the detector's CLI does not
honour these markers, so this script strips them first).
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import deckkit as dk  # noqa: E402

DETECTOR_VERSION = "3.36.0"
IGNORE = re.compile(r"<!--\s*avoid-ai-writing:ignore-start\s*-->.*?<!--\s*avoid-ai-writing:ignore-end\s*-->", re.S)

# Default limits, set from the five 2025 winning decks (see README "Baseline").
DEFAULT_THRESHOLD = {"findings_per_10_slides": 4.0, "score": 25}

# --- case-deck layer --------------------------------------------------------
LAYER = [
    ("em-dash", re.compile("—"), "high",
     "Replace the em-dash with a comma, colon, full stop or brackets."),
    ("not-x-but-y", re.compile(r"\b(?:it'?s|it is|this is|that'?s)\s+not\s+(?:just\s+|only\s+|merely\s+|about\s+)?[^.;:!?]{1,60}?[,;]\s*(?:it'?s|it is|but|rather)\b", re.I), "high",
     "Say what it is, directly. Drop the 'not X, but Y' frame."),
    ("not-just-but", re.compile(r"\bnot\s+(?:just|only|merely)\s+[^.;:!?]{1,60}?\bbut(?:\s+also)?\b", re.I), "high",
     "Say what it is, directly. Drop the 'not just X, but Y' frame."),
    ("imagine-opener", re.compile(r"\bimagine\s+(?:a\s+world|if|a\s+future)\b", re.I), "medium",
     "Open with a real person, place or number instead."),
    ("fast-paced-world", re.compile(r"\bin today'?s\s+(?:fast[- ]paced|ever[- ]changing|digital)\s+(?:world|landscape|age)\b", re.I), "high",
     "Cut the scene-setting; start with the fact."),
    ("chatbot-leftover", re.compile(r"\b(?:as an ai|i hope this helps|certainly!|here'?s (?:a|an|the) (?:revised|updated)|let me know if)\b", re.I), "high",
     "Remove the chatbot sentence."),
]
SKIP_SHAPES = ("nav:", "pattern:", "overlay:")


# --------------------------------------------------------------------------
# Detector
# --------------------------------------------------------------------------

def detector_cmd():
    here = Path(__file__).resolve().parent.parent
    for base in (here, here.parent / "course" / "ai" / "ai_pattern_check"):
        local = base / "node_modules" / ".bin" / "avoid-ai-writing"
        if local.exists():
            return [str(local)]
    if shutil.which("npx"):
        return ["npx", "--yes", "--package", f"avoid-ai-writing-detector@{DETECTOR_VERSION}", "avoid-ai-writing"]
    raise RuntimeError("avoid-ai-writing detector not found: run `npm install` in the skill folder (needs Node.js)")


def run_detector(text: str, context="general"):
    cmd = detector_cmd() + ["--source-mode", "rendered-markdown", "--context", context]
    res = subprocess.run(cmd, input=text, capture_output=True, text=True, timeout=120)
    if res.returncode != 0:
        raise RuntimeError(f"detector failed: {res.stderr.strip()[:300]}")
    return json.loads(res.stdout)


# --------------------------------------------------------------------------
# Text extraction
# --------------------------------------------------------------------------

def deck_units(path, include_notes=False):
    """[(label, text)] per slide; sources, footnotes, nav and specimens skipped."""
    prs = dk.open_deck(path)
    units = []
    for i, s in enumerate(prs.slides, 1):
        parts = []
        for sh in dk.iter_shapes(s.shapes):
            if sh.name.startswith(SKIP_SHAPES) or "Specimen" in sh.name:
                continue
            t = dk.shape_text(sh).replace("\v", " ").strip()
            if not t or re.match(r"^(Source|Footnote)s?:", t, re.I) or re.fullmatch(r"[\d\s.,%x+~\-–]*", t):
                continue
            parts.append(t)
        if include_notes and s.has_notes_slide:
            n = s.notes_slide.notes_text_frame.text.strip()
            if n:
                parts.append(n)
        units.append((f"slide {i}", "\n".join(parts)))
    return units


def file_units(path):
    text = Path(path).read_text(encoding="utf-8")
    text = IGNORE.sub("", text)
    return [(f"{Path(path).name} line {k}", line) for k, line in enumerate(text.splitlines(), 1)], text


# --------------------------------------------------------------------------
# Check
# --------------------------------------------------------------------------

def check(path, context="general", include_notes=False, client=None):
    path = Path(path)
    if path.suffix.lower() == ".pptx":
        units = deck_units(path, include_notes)
        whole = "\n\n".join(f"## {lab}\n{t}" for lab, t in units if t)
        n_slides = len(units)
    else:
        units, whole = file_units(path)
        n_slides = None
    whole = IGNORE.sub("", whole)
    det = run_detector(whole, context) if whole.strip() else {"score": 0, "label": "empty", "issues": []}

    findings = []
    for iss in det.get("issues", []):
        txt = str(iss.get("text") or "")
        where = "whole document"
        if txt and len(txt) < 80 and not txt.startswith(("Vocabulary diversity", "Sentence length")):
            hit = next((lab for lab, t in units if txt.lower() in t.lower()), None)
            where = hit or where
        findings.append({"where": where, "source": "detector", "type": iss.get("type"),
                         "severity": iss.get("severity", "medium"), "text": txt,
                         "suggestion": iss.get("suggestion") or ""})
    for lab, t in units:
        t = IGNORE.sub("", t)
        for name, rx, sev, fix in LAYER:
            for m in rx.finditer(t):
                findings.append({"where": lab, "source": "case-deck", "type": name, "severity": sev,
                                 "text": m.group(0)[:80], "suggestion": fix})
    if client and path.suffix.lower() == ".pptx":
        findings += generic_recommendations(path, client)
    detector_hits = sum(1 for f in findings if f["source"] == "detector" and f["where"] != "whole document")
    layer_hits = sum(1 for f in findings if f["source"] == "case-deck")
    return {"path": str(path), "score": det.get("score"), "label": det.get("label"),
            "slides": n_slides, "findings": findings,
            "counts": {"detector": detector_hits, "case_deck": layer_hits,
                       "document_level": sum(1 for f in findings if f["where"] == "whole document")}}


def generic_recommendations(path, client):
    """Strategy slides that never name the client read as generic."""
    from build import slide_meta
    prs = dk.open_deck(path)
    out = []
    names = [client.lower()] + [w.lower() for w in re.split(r"[\s/]+", client) if len(w) > 3]
    for i, s in enumerate(prs.slides, 1):
        if slide_meta(s).get("pattern") not in ("pillar_overview", "initiative", "initiative_steps", "exec_summary"):
            continue
        text = dk.slide_text(s).lower()
        if not any(n in text for n in names):
            out.append({"where": f"slide {i}", "source": "case-deck", "type": "generic-recommendation",
                        "severity": "low", "text": "",
                        "suggestion": f"Name {client} (its products, places, people) so the idea could not fit any company."})
    return out


def over_threshold(result, threshold=None):
    th = dict(DEFAULT_THRESHOLD)
    if threshold is not None:
        th["findings_per_10_slides"] = threshold
    hits = result["counts"]["detector"] + result["counts"]["case_deck"]
    if result["slides"]:
        rate = hits / max(1, result["slides"]) * 10
        return rate > th["findings_per_10_slides"] or (result["score"] or 0) > th["score"], rate
    return hits > 6 or (result["score"] or 0) > th["score"], hits


def print_report(r, threshold=None):
    bad, rate = over_threshold(r, threshold)
    unit = "findings per 10 slides" if r["slides"] else "findings"
    print(f"{r['path']}: score {r['score']} ({r['label']}), {rate:.1f} {unit} -> {'ABOVE' if bad else 'within'} limit")
    by_where = {}
    for f in r["findings"]:
        by_where.setdefault(f["where"], []).append(f)
    order = {"high": 0, "medium": 1, "low": 2}

    def key(w):
        m = re.search(r"(\d+)$", w)
        return (0 if w == "whole document" else 1, int(m.group(1)) if m else 0)

    for where in sorted(by_where, key=key):
        print(f"\n  {where}")
        for f in sorted(by_where[where], key=lambda f: order.get(f["severity"], 3)):
            txt = f" \"{f['text']}\"" if f["text"] else ""
            print(f"    [{f['severity']}] {f['type']}{txt} -> {f['suggestion']}")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("paths", nargs="+")
    ap.add_argument("--context", default="general", choices=["general", "technical", "marketing", "personal"])
    ap.add_argument("--notes", action="store_true", help="include speaker notes")
    ap.add_argument("--client", help="case company name, for the generic-recommendation check")
    ap.add_argument("--threshold", type=float, help="max findings per 10 slides (decks) or per file")
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args()
    results, failed = [], False
    files = []
    for p in args.paths:
        pp = Path(p)
        files += sorted(x for x in pp.rglob("*") if x.suffix.lower() in (".pptx", ".md", ".txt")) if pp.is_dir() else [pp]
    for f in files:
        r = check(f, args.context, args.notes, args.client)
        results.append(r)
        bad, _ = over_threshold(r, args.threshold)
        failed |= bad
        if not args.json:
            print_report(r, args.threshold)
            print()
    if args.json:
        print(json.dumps(results, indent=1, ensure_ascii=False))
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
