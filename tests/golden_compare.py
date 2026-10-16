#!/usr/bin/env python3
"""Side-by-side renders: Tajín original (left) vs case-deck rebuild (right).

  python tests/golden_compare.py --winning "<winning_cases>" [--engine powerpoint|quicklook]
Writes tests/output/golden_compare/pair_XX.png and an overview sheet.
"""
from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "scripts"))
import render  # noqa: E402

logging.getLogger("pypdf").setLevel(logging.ERROR)

# rebuild slide -> original Tajín slide (None: no original, new appendix slide)
PAIRS = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 17, 18, 19, 20, 21, 23, 26, 27,
         32, 33, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, None, None]


def compare(winning: Path, rebuild: Path, out: Path, engine="powerpoint"):
    from PIL import Image, ImageDraw
    import build
    import deckkit as dk
    out.mkdir(parents=True, exist_ok=True)
    # compare without the hidden navigator slide that navigation inserts
    prs = dk.open_deck(rebuild)
    for s in list(prs.slides):
        if build.slide_meta(s).get("nav_generated"):
            dk.drop_slide(prs, s)
    rebuild = out / "rebuild_without_navigator.pptx"
    prs.save(str(rebuild))
    orig = render.render(winning / "Mexico 24H Case.pptx", sorted({p for p in PAIRS if p}),
                         out / "orig", 1200, engine)
    new = render.render(rebuild, list(range(1, len(PAIRS) + 1)), out / "rebuild", 1200, engine)
    orig_map = {int(p.stem.split("_")[1]): p for p in orig}
    new_map = {int(p.stem.split("_")[1]): p for p in new}
    pairs = []
    for k, o in enumerate(PAIRS, 1):
        right = new_map.get(k)
        left = orig_map.get(o) if o else None
        if right is None:
            continue
        r = Image.open(right).convert("RGB")
        l = Image.open(left).convert("RGB").resize(r.size) if left else Image.new("RGB", r.size, "white")
        canvas = Image.new("RGB", (r.width * 2 + 30, r.height + 40), "white")
        canvas.paste(l, (0, 40))
        canvas.paste(r, (r.width + 30, 40))
        d = ImageDraw.Draw(canvas)
        d.text((10, 10), f"Original slide {o}" if o else "No original (new appendix slide)", fill=(6, 31, 50))
        d.text((r.width + 40, 10), f"Rebuild slide {k}", fill=(6, 31, 50))
        f = out / f"pair_{k:02d}.png"
        canvas.save(f)
        pairs.append(f)
    render.contact_sheet(pairs, out / "overview.png", cols=2, thumb_w=900)
    return pairs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--winning", required=True)
    ap.add_argument("--rebuild", default=str(HERE / "output" / "tajin_rebuild.pptx"))
    ap.add_argument("--engine", default="powerpoint")
    args = ap.parse_args()
    pairs = compare(Path(args.winning), Path(args.rebuild), HERE / "output" / "golden_compare", args.engine)
    print(f"{len(pairs)} pairs in {HERE / 'output' / 'golden_compare'}")


if __name__ == "__main__":
    main()
