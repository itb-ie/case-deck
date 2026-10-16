#!/usr/bin/env python3
"""Write references/patterns.md from assets/catalog.yaml (run after curate.py)."""
from __future__ import annotations

import sys
from pathlib import Path

import yaml

sys.path.insert(0, str(Path(__file__).resolve().parent))
import deckkit as dk  # noqa: E402


def main():
    cat = yaml.safe_load(dk.CATALOG.read_text(encoding="utf-8"))
    out = ["# Pattern catalogue", "",
           "Generated from `assets/catalog.yaml` by `scripts/catalog_doc.py`. Every slide in a",
           "case-deck deck comes from one of these patterns. Slot names go under `slots:` in",
           "`deck_spec.yaml`; `max` is the number of characters the box holds.", "",
           "Contact sheet: `tests/output/patterns_contact_sheet.png` (run `python scripts/render.py assets/patterns.pptx --sheet ...`).", ""]
    for k, (pid, p) in enumerate(cat["patterns"].items(), 1):
        if p.get("internal"):
            continue
        out += [f"## {pid}", "", p["purpose"], "", f"When: {p['when_to_use']}  ", f"From: {p['source']}", ""]
        for rep in p.get("repeat", []):
            out.append(f"Repeats: `{rep['prefix']}1_...` to `{rep['prefix']}{rep['count']}_...` "
                       f"({rep['min']} to {rep['count']} groups; fill only the ones you use).")
            out.append("")
        if pid == "framework":
            out += ["Pillar tracker: add `current: 2` to brighten pillar 2.", ""]
        if pid == "risk_matrix":
            out += ["Stars: `top_risks: [1, 4]` keeps the star next to risks 1 and 4.", ""]
        out += ["| Slot | Type | Max | Notes |", "|---|---|---|---|"]
        for name, info in p["slots"].items():
            if info["type"] == "harvey" and not name.endswith("_c1") and not name.startswith("h_r1_"):
                continue
            t = info["type"]
            mx = str(info.get("max_chars", "")) if t == "text" else ""
            note = "optional" if info.get("optional") else ""
            if t == "image":
                note = (note + "; " if note else "") + f"image ({info.get('label', '')})"
            if t == "harvey":
                name = "h_r<row>_c<col>"
                note = "Harvey level 0-4"
            if t == "table":
                note = info.get("note", f"{info.get('rows')} x {info.get('cols')} table")
            if t == "chart":
                note = f"chart ({info.get('chart_type', '')}): categories + series"
            out.append(f"| `{name}` | {t} | {mx} | {note} |")
        out.append("")
    path = dk.SKILL_DIR / "references" / "patterns.md"
    path.write_text("\n".join(out), encoding="utf-8")
    print("wrote", path)


if __name__ == "__main__":
    main()
