#!/usr/bin/env python3
"""Render slides of a .pptx to PNG.

macOS (default): makes a one-slide copy per slide and renders it with Quick
Look (qlmanage), which needs no LibreOffice and no open PowerPoint window.
--engine powerpoint uses Microsoft PowerPoint through AppleScript (macOS) or
COM through PowerShell (Windows).

Usage:
  python render.py deck.pptx --out renders/            # every slide
  python render.py deck.pptx --slides 1,5-8 --out r/   # some slides
  python render.py deck.pptx --sheet contact.png       # contact sheet too
"""
from __future__ import annotations

import argparse
import io
import logging
import platform
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import deckkit as dk  # noqa: E402

logging.getLogger("pypdf").setLevel(logging.ERROR)


def parse_range(spec: str | None, n: int):
    if not spec:
        return list(range(1, n + 1))
    out = []
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-")
            out.extend(range(int(a), int(b) + 1))
        elif part:
            out.append(int(part))
    return [k for k in out if 1 <= k <= n]


def _single_slide_copies(pptx: Path, slides, tmp: Path):
    data = pptx.read_bytes()
    files = []
    for k in slides:
        prs = dk.Presentation(io.BytesIO(data))
        dk.keep_only(prs, [k - 1])
        only = prs.slides[0]._element
        if only.get("show") == "0":
            del only.attrib["show"]  # Quick Look stalls on a deck whose only slide is hidden
        f = tmp / f"slide_{k:03d}.pptx"
        prs.save(str(f))
        files.append((k, f))
    return files


def render_quicklook(pptx: Path, slides, out: Path, size=1600):
    out.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td)
        files = _single_slide_copies(pptx, slides, tmp)
        for i in range(0, len(files), 40):
            batch = [str(f) for _, f in files[i:i + 40]]
            try:
                subprocess.run(["qlmanage", "-t", "-s", str(size), "-o", str(tmp), *batch],
                               check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                               timeout=30 + 5 * len(batch))
            except subprocess.TimeoutExpired:
                print("warning: Quick Look timed out (the file may be damaged)", file=sys.stderr)
        result = []
        for k, f in files:
            png = tmp / (f.name + ".png")
            if png.exists():
                dst = out / f"slide_{k:03d}.png"
                shutil.move(str(png), dst)
                result.append(dst)
            else:
                print(f"warning: Quick Look produced no image for slide {k}", file=sys.stderr)
        return result


PP_CONTAINER = Path.home() / "Library/Containers/com.microsoft.Powerpoint/Data/case-deck"


def _unhidden_copy(pptx, dst):
    """PDF export skips hidden slides; render a copy with every slide shown
    so page k is slide k."""
    prs = dk.open_deck(pptx)
    hidden = False
    for s in prs.slides:
        if s._element.get("show") == "0":
            del s._element.attrib["show"]
            hidden = True
    if hidden:
        prs.save(str(dst))
    else:
        shutil.copyfile(pptx, dst)


def powerpoint_pdf(pptx: Path, pdf: Path, timeout=180):
    """Export a deck to PDF with Microsoft PowerPoint (macOS or Windows).
    On macOS PowerPoint is sandboxed, so the work happens in its container
    folder and paths are passed in HFS form. Raises on failure or timeout
    (a timeout usually means PowerPoint is showing a repair dialog)."""
    pptx, pdf = Path(pptx).resolve(), Path(pdf).resolve()
    if platform.system() == "Darwin":
        PP_CONTAINER.mkdir(parents=True, exist_ok=True)
        src = PP_CONTAINER / pptx.name
        out = PP_CONTAINER / (pptx.stem + ".pdf")
        _unhidden_copy(pptx, src)
        if out.exists():
            out.unlink()
        script = f'''
        tell application "Microsoft PowerPoint"
            open POSIX file "{src}"
            delay 1
            set p to presentation 1
            save p in ((POSIX file "{out}") as text) as save as PDF
            close p saving no
        end tell'''
        subprocess.run(["osascript", "-e", script], check=True, timeout=timeout,
                       stdout=subprocess.DEVNULL)
        if not out.exists():
            raise RuntimeError("PowerPoint did not write a PDF")
        shutil.move(str(out), pdf)
        src.unlink(missing_ok=True)
    elif platform.system() == "Windows":
        src = pdf.with_suffix(".pptx")
        _unhidden_copy(pptx, src)
        ps = (f"$a=New-Object -ComObject PowerPoint.Application;"
              f"$p=$a.Presentations.Open('{src}',$true,$false,$false);"
              f"$p.SaveAs('{pdf}',32);$p.Close()")
        subprocess.run(["powershell", "-NoProfile", "-Command", ps], check=True, timeout=timeout)
    else:
        raise RuntimeError("PowerPoint rendering needs macOS or Windows")
    return pdf


def render_powerpoint(pptx: Path, slides, out: Path, size=1600):
    """PowerPoint draws every shape exactly (Quick Look gets pie shapes
    wrong). Slow-ish: one PowerPoint export plus one image per page."""
    from pypdf import PdfReader, PdfWriter
    out.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td)
        pdf = powerpoint_pdf(pptx, tmp / "deck.pdf")
        reader = PdfReader(str(pdf))
        result = []
        for k in slides:
            if k - 1 >= len(reader.pages):
                continue
            w = PdfWriter()
            w.add_page(reader.pages[k - 1])
            page = tmp / f"p{k}.pdf"
            with open(page, "wb") as f:
                w.write(f)
            dst = out / f"slide_{k:03d}.png"
            if platform.system() == "Darwin":
                subprocess.run(["sips", "-s", "format", "png", "-Z", str(size), str(page), "--out", str(dst)],
                               check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            result.append(dst)
        return result


def render(pptx, slides=None, out=None, size=1600, engine="auto"):
    """Render 1-based slide numbers; returns PNG paths (may be empty)."""
    pptx = Path(pptx)
    n = len(dk.open_deck(pptx).slides)
    slides = list(slides) if slides else list(range(1, n + 1))
    out = Path(out) if out else pptx.with_suffix("").parent / (pptx.stem + "_renders")
    if engine == "auto":
        engine = "quicklook" if shutil.which("qlmanage") else (
            "powerpoint" if platform.system() == "Windows" else "none")
    if engine == "quicklook":
        return render_quicklook(pptx, slides, out, size)
    if engine == "powerpoint":
        return render_powerpoint(pptx, slides, out, size)
    print("warning: no renderer available (needs macOS Quick Look or PowerPoint)", file=sys.stderr)
    return []


def contact_sheet(pngs, out, cols=4, thumb_w=480, labels=None):
    from PIL import Image, ImageDraw
    pngs = [Path(p) for p in pngs]
    if not pngs:
        return None
    thumbs = []
    for p in pngs:
        im = Image.open(p).convert("RGB")
        h = int(im.height * thumb_w / im.width)
        thumbs.append(im.resize((thumb_w, h)))
    th = max(t.height for t in thumbs)
    pad, lab = 12, 26
    rows = -(-len(thumbs) // cols)
    sheet = Image.new("RGB", (cols * (thumb_w + pad) + pad, rows * (th + lab + pad) + pad), "white")
    draw = ImageDraw.Draw(sheet)
    for i, t in enumerate(thumbs):
        x = pad + (i % cols) * (thumb_w + pad)
        y = pad + (i // cols) * (th + lab + pad)
        draw.text((x, y), labels[i] if labels else pngs[i].stem, fill=(6, 31, 50))
        sheet.paste(t, (x, y + lab))
        draw.rectangle([x - 1, y + lab - 1, x + thumb_w, y + lab + t.height], outline=(191, 191, 191))
    sheet.save(out)
    return Path(out)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pptx")
    ap.add_argument("--slides", help="e.g. 1,4-6 (1-based)")
    ap.add_argument("--out", help="output folder for PNGs")
    ap.add_argument("--size", type=int, default=1600)
    ap.add_argument("--engine", default="auto", choices=["auto", "quicklook", "powerpoint"])
    ap.add_argument("--sheet", help="also write a contact sheet PNG here")
    args = ap.parse_args()
    n = len(dk.open_deck(args.pptx).slides)
    pngs = render(args.pptx, parse_range(args.slides, n), args.out, args.size, args.engine)
    for p in pngs:
        print(p)
    if args.sheet:
        print(contact_sheet(pngs, args.sheet))


if __name__ == "__main__":
    main()
