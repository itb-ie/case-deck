#!/usr/bin/env python3
"""Build a deck from deck_spec.yaml by cloning patterns and filling slots.

Usage:
  python build.py deck_spec.yaml                 # writes the deck named in the spec
  python build.py deck_spec.yaml -o out.pptx
  python build.py deck_spec.yaml --check         # validate only, build nothing
  python build.py deck_spec.yaml --allow-samples # keep sample text in unfilled boxes

See references/storyline_format.md for the spec format.
"""
from __future__ import annotations

import argparse
import copy
import json
import re
import sys
import time
from pathlib import Path

import yaml

sys.path.insert(0, str(Path(__file__).resolve().parent))
import deckkit as dk  # noqa: E402

BRACKET = re.compile(r"\[[^\]]{1,80}\]")
OPEN_TAGS = re.compile(r"\[(ASSUMPTION|SOURCE\?)[^\]]*\]")


class SpecError(Exception):
    pass


# --------------------------------------------------------------------------
# Loading
# --------------------------------------------------------------------------

def load_catalog(path=dk.CATALOG):
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)


def load_spec(path):
    with open(path, encoding="utf-8") as f:
        spec = yaml.safe_load(f) or {}
    if "slides" not in spec:
        raise SpecError("the spec needs a `slides:` list")
    spec.setdefault("deck", {})
    return spec


def pattern_slides(prs):
    out = {}
    for s in prs.slides:
        for sh in s.shapes:
            if sh.name.startswith("pattern:"):
                out[sh.name[8:]] = s
    return out


def _flat(value):
    if isinstance(value, (list, tuple)):
        return " ".join(_flat(v) for v in value)
    return "" if value is None else str(value)


def has_placeholder_text(value) -> bool:
    text = BRACKET.sub(lambda m: "" if OPEN_TAGS.fullmatch(m.group(0)) else m.group(0), _flat(value))
    return bool(BRACKET.search(text))


def _visible_len(value):
    return len(re.sub(r"\*\*|## ", "", _flat(value)))


# --------------------------------------------------------------------------
# Validation
# --------------------------------------------------------------------------

def repeat_count(pattern, slots, explicit):
    """How many groups of each repeat prefix the slide uses."""
    counts = {}
    for rep in pattern.get("repeat", []):
        p = rep["prefix"]
        if explicit and p in explicit:
            counts[p] = int(explicit[p])
            continue
        ks = [int(m.group(1)) for s in slots for m in [re.match(rf"^{p}(\d+)_", s)] if m]
        counts[p] = max(ks) if ks else rep["count"]
    return counts


def validate(spec, catalog, base_dir, allow_samples=False):
    errors, warnings = [], []
    pats = catalog["patterns"]
    ids = set()
    for i, sl in enumerate(spec["slides"], 1):
        where = f"slide {i}"
        pid = sl.get("pattern")
        if pid not in pats or pats[pid].get("internal"):
            close = [p for p in pats if not pats[p].get("internal") and (pid or "")[:4] in p]
            errors.append(f"{where}: unknown pattern {pid!r}" + (f" (did you mean {close[0]}?)" if close else ""))
            continue
        pat = pats[pid]
        slots = sl.get("slots", {}) or {}
        for name, value in slots.items():
            if name not in pat["slots"]:
                similar = [s for s in pat["slots"] if s.split("_")[-1] == name.split("_")[-1]][:3]
                errors.append(f"{where} ({pid}): unknown slot {name!r}" + (f"; similar: {', '.join(similar)}" if similar else ""))
                continue
            info = pat["slots"][name]
            if info["type"] == "text":
                mx = info.get("max_chars")
                n = _visible_len(value)
                if mx and n > mx * 1.15:
                    errors.append(f"{where} ({pid}): {name} has {n} characters, the box fits about {mx}. Shorten it.")
                elif mx and n > mx:
                    warnings.append(f"{where} ({pid}): {name} has {n} characters for a box of about {mx}; check the render.")
            elif info["type"] in ("image", "photo"):
                if value and not (base_dir / str(value)).exists():
                    errors.append(f"{where} ({pid}): image for {name} not found: {value}")
            elif info["type"] == "chart":
                if not isinstance(value, dict) or "categories" not in value or "series" not in value:
                    errors.append(f"{where} ({pid}): chart needs categories: and series:")
                else:
                    for sname, vals in value["series"].items():
                        if len(vals) != len(value["categories"]):
                            errors.append(f"{where} ({pid}): series {sname!r} has {len(vals)} values for {len(value['categories'])} categories")
            elif info["type"] == "harvey":
                if value is not None and not (0 <= int(value) <= 4):
                    errors.append(f"{where} ({pid}): {name} must be 0-4")
        counts = repeat_count(pat, slots, sl.get("count"))
        for rep in pat.get("repeat", []):
            n = counts[rep["prefix"]]
            if not rep["min"] <= n <= rep["count"]:
                errors.append(f"{where} ({pid}): {n} groups of {rep['prefix']}* but this pattern takes {rep['min']}-{rep['count']}")
        # required text slots that keep only placeholder sample text
        for name, info in pat["slots"].items():
            if info.get("optional") or info["type"] != "text":
                continue
            m = None
            for rep in pat.get("repeat", []):
                m = re.match(rf"^{rep['prefix']}(\d+)_", name)
                if m and int(m.group(1)) > counts[rep["prefix"]]:
                    break
                m = None
            if m:
                continue  # group not used
            if name == "tracker" and (sl.get("section") or sl.get("appendix")):
                continue
            if name not in slots and has_placeholder_text(info.get("sample", "")):
                msg = f"{where} ({pid}): required box {name!r} is not filled"
                (warnings if allow_samples else errors).append(msg)
        for name, value in slots.items():
            if isinstance(value, list):
                for x, y in zip(value, value[1:]):
                    if isinstance(x, str) and isinstance(y, str) and re.search(r"\d$", x) and re.match(r"\d{3}\b", y):
                        warnings.append(f"{where} ({pid}): {name} has “{x}” followed by “{y}”: a number with a "
                                        "thousands comma was split into two items. Put the item in quotes.")
            if isinstance(value, (str, list)) and has_placeholder_text(value):
                warnings.append(f"{where} ({pid}): {name} still contains placeholder text: {_flat(value)[:60]}")
        if sl.get("id"):
            if sl["id"] in ids:
                errors.append(f"{where}: duplicate id {sl['id']!r}")
            ids.add(sl["id"])
        if pid == "framework" and sl.get("current"):
            n = counts.get("p", 3)
            if not 1 <= int(sl["current"]) <= n:
                errors.append(f"{where}: current must be between 1 and {n}")
    return errors, warnings


# --------------------------------------------------------------------------
# Building one slide
# --------------------------------------------------------------------------

def _shapes_of_group(slide, prefix, k):
    rx = re.compile(rf"^(slot:{prefix}{k}_|grp:{prefix}{k}:)")
    out = []
    for sh in slide.shapes:
        if rx.match(sh.name):
            out.append(sh)
            continue
        if sh.shape_type == 6:  # group containing slots of this group
            if any(rx.match(c.name) for c in dk.iter_shapes(sh.shapes)):
                out.append(sh)
    return out


def apply_repeat(slide, pat, counts):
    for rep in pat.get("repeat", []):
        p, n, full = rep["prefix"], counts[rep["prefix"]], rep["count"]
        if n >= full:
            continue
        groups = [_shapes_of_group(slide, p, k) for k in range(1, full + 1)]
        all_boxes = [dk.union(dk.bbox(s) for s in g) for g in groups if g]
        lo = 0 if rep["axis"] == "x" else 1
        span = (min(b[lo] for b in all_boxes), max(b[lo + 2] for b in all_boxes)) if all_boxes else None
        for g in groups[n:]:
            for sh in g:
                if sh._element.getparent() is None:
                    continue
                if sh.name.startswith("slot:"):
                    dk.delete_with_decor(slide, sh)
                else:
                    dk.remove_shape(sh)
        if rep["mode"] != "none":
            dk.redistribute(slide, groups[:n], n, rep["axis"], rep["mode"], span)


def tag_slide(slide, pid, meta):
    tag = None
    for sh in slide.shapes:
        if sh.name.startswith("pattern:"):
            tag = sh
    if tag is None:
        tag = slide.shapes.add_textbox(0, 0, 0, 0)
    tag.name = f"pattern:{pid}"
    cnv = tag._element.find(f".//{{{dk.NS['p']}}}cNvPr")
    cnv.set("descr", json.dumps(meta, ensure_ascii=False))


def slide_meta(slide):
    for sh in slide.shapes:
        if sh.name.startswith("pattern:"):
            cnv = sh._element.find(f".//{{{dk.NS['p']}}}cNvPr")
            try:
                meta = json.loads(cnv.get("descr") or "{}")
            except json.JSONDecodeError:
                meta = {}
            meta["pattern"] = sh.name[8:]
            return meta
    return {}


def fill_slide(slide, pat, sl, base_dir, report, allow_samples=False):
    slots = sl.get("slots", {}) or {}
    shapes = dk.slot_shapes(slide)
    counts = repeat_count(pat, slots, sl.get("count"))
    if "tracker" in shapes and "tracker" not in slots:
        if sl.get("appendix") and (not sl.get("section") or sl.get("section") == "Appendix"):
            slots = {**slots, "tracker": "Appendix"}
        elif sl.get("section"):
            slots = {**slots, "tracker": sl["section"]}

    # values
    for name, value in slots.items():
        sh = shapes.get(name)
        info = pat["slots"].get(name, {})
        if sh is None:
            continue
        t = info.get("type", "text")
        if t == "text":
            dk.set_text(sh, value)
        elif t == "image":
            if value:
                dk.replace_picture(slide, sh, base_dir / str(value))
        elif t == "photo":
            dk.set_photo_fill(slide, sh, base_dir / str(value) if value else None)
        elif t == "chart":
            dk.fill_chart(sh, value)
        elif t == "table":
            fill_table_slot(slide, sh, value, pat)
        elif t == "harvey":
            pass  # handled below
        elif t == "marker" and value is False:
            dk.delete_with_decor(slide, sh)

    # Harvey balls
    if any(i.get("type") == "harvey" for i in pat["slots"].values()):
        for name, info in pat["slots"].items():
            if info.get("type") != "harvey":
                continue
            pie = shapes.get(name)
            oval = next((s for s in slide.shapes if s.name == "hb:" + name[2:]), None)
            if pie is None or oval is None or pie._element.getparent() is None:
                continue
            dk.set_harvey(pie, oval, int(slots.get(name, 0) or 0))

    # framework as pillar tracker
    if sl.get("current"):
        sh = shapes.get(f"p{int(sl['current'])}_shade")
        if sh is not None:
            dk.remove_shape(sh)

    # risk stars
    if "top_risks" in sl:
        top = {int(k) for k in sl["top_risks"]}
        for name, sh in list(shapes.items()):
            m = re.match(r"^r(\d+)_star$", name)
            if m and int(m.group(1)) not in top and sh._element.getparent() is not None:
                dk.remove_shape(sh)

    apply_repeat(slide, pat, counts)

    # unfilled slots
    shapes = dk.slot_shapes(slide)
    for name, info in pat["slots"].items():
        sh = shapes.get(name)
        if sh is None or sh._element.getparent() is None or name in slots:
            continue
        t = info.get("type", "text")
        if t == "marker":
            m = re.match(r"^([a-z]+\d+)_", name)
            if m and name.endswith("_star"):
                if "top_risks" not in sl:
                    dk.remove_shape(sh)
                continue
            if m and not any(s.startswith(m.group(1) + "_") for s in slots) and info.get("optional"):
                dk.delete_with_decor(slide, sh)
            continue
        if t == "harvey":
            continue
        if t == "photo":
            report["placeholders"].append(f"{name} (team photo)")
            continue
        if name == "index_list":
            continue  # filled after the whole deck is built
        if info.get("optional"):
            dk.delete_with_decor(slide, sh)
        elif t in ("image", "photo"):
            report["placeholders"].append(f"{name} ({info.get('label', t)})")
        elif t == "text" and has_placeholder_text(info.get("sample", "")) and not allow_samples:
            report["unfilled"].append(name)

    for k, v in slots.items():
        if isinstance(v, (str, list)) and OPEN_TAGS.search(_flat(v)):
            report["open_tags"].append(f"{k}: {OPEN_TAGS.search(_flat(v)).group(0)}")

    if sl.get("notes"):
        slide.notes_slide.notes_text_frame.text = str(sl["notes"])


def fill_table_slot(slide, sh, rows, pat):
    """Native table: fill cells; drop unused rows/columns (and their Harvey balls)."""
    tbl = sh.table
    n_rows, n_cols = len(rows), max(len(r) for r in rows)
    dk.fill_table(sh, rows)
    for j in range(len(tbl.columns) - 1, n_cols - 1, -1):
        grid = tbl._tbl.tblGrid
        grid.remove(grid.findall(f"{{{dk.NS['a']}}}gridCol")[j])
        for tr in tbl._tbl.tr_lst:
            tcs = tr.findall(f"{{{dk.NS['a']}}}tc")
            tr.remove(tcs[j])
    for i in range(len(tbl.rows) - 1, n_rows - 1, -1):
        tbl._tbl.remove(tbl._tbl.tr_lst[i])
    # Harvey balls outside the used table area
    for s in list(slide.shapes):
        m = re.match(r"^(?:slot:h_|hb:)r(\d+)_c(\d+)$", s.name)
        if m and (int(m.group(1)) > n_rows - 1 or int(m.group(2)) > n_cols - 1):
            dk.remove_shape(s)


def add_takeaway(prs, built, overlay_slide, text, head=None):
    """Duplicate a built slide and lay the Key Takeaway band over it."""
    dup = dk.clone_slide(built, prs)
    for sh in overlay_slide.shapes:
        if sh.name.startswith("pattern:"):
            continue
        dup.shapes._spTree.append(copy.deepcopy(sh._element))
    shapes = dk.slot_shapes(dup)
    dk.set_text(shapes["takeaway"], text)
    if head:
        dk.set_text(shapes["takeaway_head"], head)
    for name in ("takeaway", "takeaway_head"):
        shapes[name].name = f"overlay:{name}"
    meta = slide_meta(built)
    meta["takeaway_of"] = meta.get("id") or True
    meta.pop("id", None)
    tag_slide(dup, meta.pop("pattern"), meta)
    return dup


# --------------------------------------------------------------------------
# Deck-level
# --------------------------------------------------------------------------

def set_client_logo(prs, logo_path=None):
    """Put the client logo where the master shows the school name
    (bottom right), as the winning decks did. Without a logo the school
    name stays. Returns the number of masters changed."""
    from PIL import Image
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    if not logo_path:
        return 0
    changed = 0
    for master in prs.slide_masters:
        logo_box = next((sh for sh in master.shapes if sh.name == "SlideLogoText"), None)
        if logo_box is None:
            continue
        right = logo_box.left + logo_box.width
        bottom = logo_box.top + logo_box.height
        dk.remove_shape(logo_box)
        img_part = prs.part.package.get_or_add_image_part(str(logo_path))
        rId = master.part.relate_to(img_part, RT.IMAGE)
        iw, ih = Image.open(logo_path).size
        box_w, box_h = int(0.62 * dk.IN), int(0.30 * dk.IN)
        scale = min(box_w / iw, box_h / ih)
        w, h = int(iw * scale), int(ih * scale)
        x, y = int(12.04 * dk.IN), int(bottom - h)
        ids = [int(e.get("id")) for e in master.shapes._spTree.iter() if e.tag.endswith("}cNvPr") and e.get("id", "").isdigit()]
        pic = dk.etree.fromstring(
            f'<p:pic xmlns:p="{dk.NS["p"]}" xmlns:a="{dk.NS["a"]}" xmlns:r="{dk.R_NS}">'
            f'<p:nvPicPr><p:cNvPr id="{max(ids + [1]) + 1}" name="ClientLogo"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr userDrawn="1"/></p:nvPicPr>'
            f'<p:blipFill><a:blip r:embed="{rId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>')
        master.shapes._spTree.append(pic)
        changed += 1
    return changed


def fill_appendix_index(prs):
    """Without an explicit index_list, list the appendix slides on the index."""
    slides = list(prs.slides)
    for i, s in enumerate(slides):
        meta = slide_meta(s)
        if meta.get("pattern") != "appendix_index":
            continue
        box = dk.slot_shapes(s).get("index_list")
        if box is None and meta.get("auto_index") is False:
            continue
        titles = []
        for later in slides[i + 1:]:
            m = slide_meta(later)
            if m.get("takeaway_of") or not m.get("appendix", True if not m else False):
                continue
            ph = dk.placeholder(later, 0)
            if ph is not None and ph.has_text_frame and ph.text_frame.text.strip():
                t = re.sub(r"^Appendix\.?\s*", "", ph.text_frame.text.strip())
                titles.append(t.split(" - ")[0].strip())
        if not titles:
            continue
        if box is None:
            continue
        dk.set_text(box, titles)


def build(spec_path, out=None, allow_samples=False, catalog=None, quiet=False, nav=True, nav_engine="auto"):
    t0 = time.time()
    spec_path = Path(spec_path)
    base_dir = spec_path.parent
    spec = load_spec(spec_path)
    catalog = catalog or load_catalog()
    errors, warnings = validate(spec, catalog, base_dir, allow_samples)
    if errors:
        raise SpecError("\n".join(errors))

    prs = dk.open_deck(dk.PATTERNS)
    pslides = pattern_slides(prs)
    originals = list(prs.slides)
    deck = spec["deck"]
    logo = deck.get("client_logo")
    set_client_logo(prs, base_dir / logo if logo else None)

    report = {"placeholders": [], "unfilled": [], "open_tags": [], "warnings": warnings, "slides": 0}
    section = None
    for i, sl in enumerate(spec["slides"], 1):
        pid = sl["pattern"]
        pat = catalog["patterns"][pid]
        built = dk.clone_slide(pslides[pid], prs)
        section = sl.get("section") or ("Appendix" if sl.get("appendix") else section)
        before = {k: len(v) for k, v in report.items() if isinstance(v, list)}
        fill_slide(built, pat, {**sl, "section": sl.get("section") or (section if pid not in ("title", "closing", "hook_image", "section_divider", "appendix_divider") else None)},
                   base_dir, report, allow_samples)
        for k in ("placeholders", "unfilled"):
            for j in range(before[k], len(report[k])):
                report[k][j] = f"slide {len(prs.slides) - len(originals)}: {report[k][j]}"
        meta = {"section": section, "appendix": bool(sl.get("appendix")), "group": sl.get("group"),
                "id": sl.get("id"), "hidden": bool(sl.get("hidden"))}
        tag_slide(built, pid, {k: v for k, v in meta.items() if v not in (None, False)})
        if sl.get("hidden"):
            built._element.set("show", "0")
        if sl.get("takeaway"):
            add_takeaway(prs, built, pslides["takeaway_overlay"], sl["takeaway"], sl.get("takeaway_head"))
    for s in originals:
        dk.drop_slide(prs, s)
    fill_appendix_index(prs)
    for s in prs.slides:
        dk.drop_unused_rels(s)
    report["slides"] = len(prs.slides)
    cp = prs.core_properties
    cp.author = str(deck.get("author", ""))
    cp.last_modified_by = ""
    if deck.get("title"):
        cp.title = str(deck["title"])
    cp.keywords = "case-deck" + (f"; client={deck['client']}" if deck.get("client") else "")
    out = Path(out) if out else base_dir / deck.get("output", spec_path.stem.replace("_spec", "") + ".pptx")
    lock = out.parent / f"~${out.name}"
    if lock.exists():
        print(f"warning: {lock.name} exists; close {out.name} in PowerPoint first.", file=sys.stderr)
    prs.save(str(out))
    report["build_seconds"] = round(time.time() - t0, 2)
    if nav and any(slide_meta(s).get("pattern") == "appendix_index" for s in prs.slides):
        import navigate
        engine = nav_engine if nav_engine != "auto" else "quicklook"
        report["navigation"] = navigate.navigate(out, out, engine)
        report["slides"] = len(dk.open_deck(out).slides)
    report["seconds"] = round(time.time() - t0, 2)
    report["output"] = str(out)
    if not quiet:
        print_report(report)
    return out, report


def print_report(r):
    print(f"Built {r['output']}: {r['slides']} slides in {r['seconds']}s")
    n = r.get("navigation")
    if n:
        print(f"Navigation: {n['zooms']} zooms, {n['corner_links']} 'Appendix ›' links, "
              f"{n['sections']} sections, hidden navigator {'added' if n['navigator'] else 'not needed'}")
    elif "navigation" not in r:
        print("No navigation (no appendix_index slide, or --no-nav).")
    if r["unfilled"]:
        print("\nBoxes still showing sample text:")
        for u in r["unfilled"]:
            print("  -", u)
    if r["placeholders"]:
        print("\nImage placeholders to replace by hand or in the spec:")
        for p in r["placeholders"]:
            print("  -", p)
    if r["open_tags"]:
        print("\nOpen [ASSUMPTION] / [SOURCE?] tags to resolve before presenting:")
        for t in r["open_tags"]:
            print("  -", t)
    if r["warnings"]:
        print("\nWarnings:")
        for w in r["warnings"]:
            print("  -", w)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("spec")
    ap.add_argument("-o", "--out")
    ap.add_argument("--check", action="store_true", help="validate only")
    ap.add_argument("--allow-samples", action="store_true", help="unfilled required boxes keep sample text")
    ap.add_argument("--no-nav", action="store_true", help="skip the appendix zooms and links")
    ap.add_argument("--engine", default="auto", choices=["auto", "quicklook", "powerpoint"],
                    help="renderer for the zoom thumbnails")
    args = ap.parse_args()
    try:
        if args.check:
            spec = load_spec(args.spec)
            errors, warnings = validate(spec, load_catalog(), Path(args.spec).parent, args.allow_samples)
            for w in warnings:
                print("warning:", w)
            for e in errors:
                print("error:", e)
            print("OK" if not errors else f"{len(errors)} error(s)")
            sys.exit(1 if errors else 0)
        build(args.spec, args.out, args.allow_samples, nav=not args.no_nav, nav_engine=args.engine)
    except SpecError as e:
        print("The spec has problems; nothing was built:\n" + str(e), file=sys.stderr)
        sys.exit(2)


if __name__ == "__main__":
    main()
