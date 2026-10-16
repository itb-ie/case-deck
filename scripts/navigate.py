#!/usr/bin/env python3
"""Add Q&A navigation to a deck, the way the Zoom demo deck does it.

  * PowerPoint sections: one per storyline section, plus one per appendix group
  * Appendix index: Slide Zooms to every backup slide (return to the index)
  * "Back to presentation" button on the index (jumps to the last slide shown)
  * "Appendix ›" link in the corner of every main slide
  * Hidden navigator slide after the title with zooms to each main section

Re-running replaces earlier navigation (shapes named nav:*), never duplicates.

Usage:
  python navigate.py deck.pptx                  # in place
  python navigate.py deck.pptx -o deck_nav.pptx
  python navigate.py deck.pptx --engine powerpoint   # sharper thumbnails
"""
from __future__ import annotations

import argparse
import io
import math
import re
import sys
import tempfile
import uuid
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import deckkit as dk  # noqa: E402
from build import slide_meta, tag_slide  # noqa: E402
from lxml import etree  # noqa: E402
from pptx.opc.constants import RELATIONSHIP_TYPE as RT  # noqa: E402
from pptx.util import Inches, Pt  # noqa: E402

A, P = dk.NS["a"], dk.NS["p"]
SECTION_EXT = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"
INDEX_AREA = (Inches(0.61), Inches(1.55), Inches(12.12), Inches(5.05))  # left, top, width, height
MAX_TILES = 12


# --------------------------------------------------------------------------
# Reading the deck
# --------------------------------------------------------------------------

def title_of(slide):
    ph = dk.placeholder(slide, 0)
    if ph is not None and ph.has_text_frame and ph.text_frame.text.strip():
        return ph.text_frame.text.strip()
    for sh in slide.shapes:
        if sh.name in ("slot:section", "slot:title") and sh.has_text_frame and sh.text_frame.text.strip():
            return sh.text_frame.text.strip()
    return ""


def short_title(t, n=60):
    t = re.sub(r"^Appendix\.?\s*", "", t).split(" - ")[0].strip()
    return t if len(t) <= n else t[: n - 1].rsplit(" ", 1)[0] + "…"


def classify(prs):
    """Per slide: role (title/main/appendix/index/navigator), section, group."""
    out = []
    in_appendix = False
    for s in prs.slides:
        meta = slide_meta(s)
        pattern = meta.get("pattern", "")
        tracker = dk.placeholder(s, 17)
        tracker_text = tracker.text_frame.text.strip() if tracker is not None and tracker.has_text_frame else ""
        if pattern == "appendix_divider" or (not meta and title_of(s).lower() == "appendix"):
            in_appendix = True
        role = "main"
        if meta.get("nav_generated") == "navigator":
            role = "navigator"
        elif meta.get("nav_generated") == "group_index":
            role = "group_index"
        elif pattern == "appendix_index" or (not meta and title_of(s).lower().startswith("appendix. index")):
            role = "index"
        elif meta.get("appendix") or in_appendix:
            role = "appendix"
        elif pattern == "title" or (not out and not meta):
            role = "title"
        if meta.get("takeaway_of"):
            role = role + "_takeaway" if role in ("main", "appendix") else role
        section = meta.get("section") or tracker_text or None
        group = meta.get("group") or (tracker_text if role.startswith("appendix") and tracker_text.lower() != "appendix" else None)
        out.append({"slide": s, "meta": meta, "role": role, "section": section, "group": group,
                    "hidden": s._element.get("show") == "0", "title": title_of(s)})
    return out


# --------------------------------------------------------------------------
# Cleaning earlier navigation
# --------------------------------------------------------------------------

def _nav_elements(slide):
    tree = slide.shapes._spTree
    found = []
    for child in tree:
        names = [c.get("name", "") for c in child.iter(f"{{{P}}}cNvPr")]
        if any(n.startswith("nav:") for n in names[:1]):
            found.append(child)
    return found


def clear_navigation(prs):
    for s in list(prs.slides):
        meta = slide_meta(s)
        if meta.get("nav_generated"):
            dk.drop_slide(prs, s)
            continue
        removed = _nav_elements(s)
        for el in removed:
            el.getparent().remove(el)
        if removed:
            for rId, rel in list(s.part.rels.items()):
                if rel.reltype in (RT.IMAGE, RT.SLIDE):
                    s.part.drop_rel(rId)  # only drops when no longer referenced
    _set_sections(prs, [])


# --------------------------------------------------------------------------
# Sections
# --------------------------------------------------------------------------

def _set_sections(prs, sections):
    """sections: list of (name, [slide_id, ...])"""
    pres = prs.part._element
    ext_lst = pres.find(f"{{{P}}}extLst")
    if ext_lst is not None:
        for ext in ext_lst.findall(f"{{{P}}}ext"):
            if ext.get("uri") == SECTION_EXT:
                ext_lst.remove(ext)
    if not sections:
        return
    if ext_lst is None:
        ext_lst = etree.SubElement(pres, f"{{{P}}}extLst")
    ext = etree.Element(f"{{{P}}}ext", uri=SECTION_EXT)
    p14 = dk.NS["p14"]
    lst = etree.SubElement(ext, f"{{{p14}}}sectionLst", nsmap={"p14": p14})
    for name, ids in sections:
        sec = etree.SubElement(lst, f"{{{p14}}}section", name=name, id="{" + str(uuid.uuid4()).upper() + "}")
        il = etree.SubElement(sec, f"{{{p14}}}sldIdLst")
        for i in ids:
            etree.SubElement(il, f"{{{p14}}}sldId", id=str(i))
    ext_lst.insert(0, ext)


def build_sections(info):
    sections = []
    current = None
    for row in info:
        role = row["role"]
        if role == "title":
            name = "Opening"
        elif role == "navigator":
            name = current or "Opening"
        elif role.startswith("appendix") or role in ("index", "group_index"):
            name = f"Appendix: {row['group']}" if row["group"] else "Appendix"
        else:
            name = row["section"] or current or "Main"
        if role == "main" and row["meta"].get("pattern") in ("section_divider",):
            # a divider opens the section named on it
            name = row["title"] or name
        if sections and sections[-1][0] == name:
            sections[-1][1].append(row["slide"].slide_id)
        else:
            sections.append((name, [row["slide"].slide_id]))
        current = name if not name.startswith("Appendix") else current
    return sections


# --------------------------------------------------------------------------
# Zooms and links
# --------------------------------------------------------------------------

def _next_id(slide):
    ids = [int(e.get("id")) for e in slide.shapes._spTree.iter(f"{{{P}}}cNvPr") if e.get("id", "").isdigit()]
    return max(ids + [1]) + 1


def add_zoom(slide, target, png, x, y, w, h, name, return_to_parent):
    """A Slide Zoom with a picture fallback, as in the demo deck."""
    _, img_rId = slide.part.get_or_add_image_part(io.BytesIO(png))
    sld_rId = slide.part.relate_to(target.part, RT.SLIDE)
    sid = _next_id(slide)
    guid = "{" + str(uuid.uuid4()).upper() + "}"
    xml = f'''<mc:AlternateContent xmlns:mc="{dk.NS['mc']}" xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{dk.R_NS}">
  <mc:Choice xmlns:pslz="{dk.NS['pslz']}" Requires="pslz">
    <p:graphicFrame>
      <p:nvGraphicFramePr><p:cNvPr id="{sid}" name="{name}"/>
        <p:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></p:xfrm>
      <a:graphic><a:graphicData uri="{dk.NS['pslz']}">
        <pslz:sldZm><pslz:sldZmObj sldId="{target.slide_id}">
          <pslz:zmPr id="{guid}" returnToParent="{1 if return_to_parent else 0}" transitionDur="1000">
            <p166:blipFill xmlns:p166="{dk.NS['p166']}"><a:blip r:embed="{img_rId}"/><a:stretch><a:fillRect/></a:stretch></p166:blipFill>
            <p166:spPr xmlns:p166="{dk.NS['p166']}"><a:xfrm><a:off x="0" y="0"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              <a:ln w="9525"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln></p166:spPr>
          </pslz:zmPr>
        </pslz:sldZmObj></pslz:sldZm>
      </a:graphicData></a:graphic>
    </p:graphicFrame>
  </mc:Choice>
  <mc:Fallback>
    <p:pic>
      <p:nvPicPr><p:cNvPr id="{sid}" name="{name}"><a:hlinkClick r:id="{sld_rId}" action="ppaction://hlinksldjump"/></p:cNvPr>
        <p:cNvPicPr><a:picLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>
      <p:blipFill><a:blip r:embed="{img_rId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
      <p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:ln w="9525"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln></p:spPr>
    </p:pic>
  </mc:Fallback>
</mc:AlternateContent>'''
    el = etree.fromstring(xml)
    slide.shapes._spTree.append(el)
    return el


def add_label(slide, text, x, y, w, h, name, size=10, color="061F32", bold=False, align="l"):
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    tb = slide.shapes.add_textbox(x, y, w, h)
    tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.02)
    tf.margin_top = tf.margin_bottom = Inches(0.01)
    tf.vertical_anchor = MSO_ANCHOR.TOP
    p = tf.paragraphs[0]
    p.alignment = {"l": PP_ALIGN.LEFT, "r": PP_ALIGN.RIGHT, "c": PP_ALIGN.CENTER}[align]
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.name = "Arial"
    r.font.color.rgb = RGBColor.from_string(color)
    return tb


def link_shape(slide, shape, target=None, action=None):
    """Make a shape clickable: jump to target slide, or run an action."""
    cnv = shape._element.find(f".//{{{P}}}cNvPr")
    for old in cnv.findall(f"{{{A}}}hlinkClick"):
        cnv.remove(old)
    h = etree.SubElement(cnv, f"{{{A}}}hlinkClick")
    if target is not None:
        h.set(f"{{{dk.R_NS}}}id", slide.part.relate_to(target.part, RT.SLIDE))
        h.set("action", "ppaction://hlinksldjump")
    else:
        h.set(f"{{{dk.R_NS}}}id", "")
        h.set("action", action)


def add_button(slide, text, x, y, w, h, name, target=None, action=None):
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import PP_ALIGN
    b = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    b.name = name
    b.fill.solid()
    b.fill.fore_color.rgb = RGBColor.from_string("061F32")
    b.line.fill.background()
    b.shadow.inherit = False
    tf = b.text_frame
    tf.margin_top = tf.margin_bottom = Inches(0.01)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = text
    r.font.size, r.font.bold, r.font.name = Pt(10), True, "Arial"
    r.font.color.rgb = RGBColor.from_string("FFFFFF")
    link_shape(slide, b, target, action)
    return b


def grid(n, area=INDEX_AREA, label_h=Inches(0.34), gap=Inches(0.22)):
    """Tile positions (x, y, w, h) for n 16:9 thumbnails with a label under each."""
    left, top, width, height = area
    best = None
    for cols in range(1, 8):
        rows = math.ceil(n / cols)
        w = (width - gap * (cols - 1)) / cols
        h = w * 9 / 16
        total = rows * (h + label_h) + gap * (rows - 1)
        if total > height:
            h = (height - gap * (rows - 1)) / rows - label_h
            w = h * 16 / 9
        if best is None or w > best[2]:
            best = (cols, rows, w, h)
    cols, rows, w, h = best
    row_w = cols * w + (cols - 1) * gap
    x0 = left + (width - row_w) / 2
    tiles = []
    for k in range(n):
        r, c = divmod(k, cols)
        tiles.append((int(x0 + c * (w + gap)), int(top + r * (h + label_h + gap)), int(w), int(h)))
    return tiles, label_h


# --------------------------------------------------------------------------
# Thumbnails
# --------------------------------------------------------------------------

def thumbnails(deck_path, slide_numbers, engine="quicklook"):
    """{slide_number: png bytes}; falls back to a labelled tile if rendering fails."""
    import render
    out = {}
    with tempfile.TemporaryDirectory() as td:
        try:
            pngs = render.render(deck_path, sorted(set(slide_numbers)), Path(td), 800, engine)
        except Exception as e:  # no renderer, PowerPoint missing, ...
            print(f"warning: thumbnails not rendered ({e}); using plain tiles", file=sys.stderr)
            pngs = []
        for p in pngs:
            out[int(p.stem.split("_")[1])] = _shrink(p.read_bytes())
    return out


def _shrink(png, width=640):
    from PIL import Image
    im = Image.open(io.BytesIO(png)).convert("RGB")
    im = im.resize((width, int(im.height * width / im.width)))
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=82)
    return buf.getvalue()


def tile_png(label):
    return _shrink(dk.placeholder_png(label, 1600, 900))


# --------------------------------------------------------------------------
# Main routine
# --------------------------------------------------------------------------

def navigate(deck_path, out_path=None, engine="quicklook", corner_links=True, navigator=True):
    deck_path = Path(deck_path)
    out_path = Path(out_path) if out_path else deck_path
    prs = dk.open_deck(deck_path)
    clear_navigation(prs)
    info = classify(prs)
    index_rows = [r for r in info if r["role"] == "index"]
    appendix = [r for r in info if r["role"] == "appendix"
                and r["meta"].get("pattern") != "appendix_divider" and r["title"].lower() != "appendix"]
    report = {"zooms": 0, "corner_links": 0, "sections": 0, "navigator": False, "group_indexes": 0}

    # group large appendices: one hidden-from-flow group index per group
    group_slides = []
    if index_rows and len(appendix) > MAX_TILES:
        groups = {}
        for r in appendix:
            groups.setdefault(r["group"] or "Other", []).append(r)
        index_slide = index_rows[0]["slide"]
        for gname, rows in groups.items():
            gs = dk.clone_slide(index_slide, prs)
            for sh in list(gs.shapes):
                if sh.name in ("slot:index_list",):
                    dk.remove_shape(sh)
            ph = dk.placeholder(gs, 0)
            if ph is not None:
                dk.set_text(ph, f"Appendix. {gname}")
            tag_slide(gs, "appendix_index", {"nav_generated": "group_index", "group": gname, "appendix": True})
            first = rows[0]["slide"]
            dk.move_slide(prs, gs, dk.slide_index(prs, first))
            group_slides.append((gname, gs, rows))
        report["group_indexes"] = len(group_slides)

    # navigator: hidden slide after the title with zooms to each main section
    nav_slide = None
    main_starts = []
    if navigator and index_rows:
        seen = set()
        for r in info:
            if r["role"] != "main":
                continue
            sec = r["section"] or "Main"
            if r["meta"].get("pattern") == "section_divider":
                sec = r["title"] or sec
            if sec not in seen:
                seen.add(sec)
                main_starts.append((sec, r["slide"]))
        if len(main_starts) >= 2:
            nav_slide = dk.clone_slide(index_rows[0]["slide"], prs)
            for sh in list(nav_slide.shapes):
                if sh.name == "slot:index_list":
                    dk.remove_shape(sh)
            ph = dk.placeholder(nav_slide, 0)
            if ph is not None:
                dk.set_text(ph, "Navigator. Main sections")
            tr = dk.placeholder(nav_slide, 17)
            if tr is not None:
                dk.set_text(tr, "Navigator")
            tag_slide(nav_slide, "navigator", {"nav_generated": "navigator"})
            nav_slide._element.set("show", "0")
            title_pos = next((i for i, r in enumerate(info) if r["role"] == "title"), 0)
            dk.move_slide(prs, nav_slide, title_pos + 1)
            report["navigator"] = True

    # save once so thumbnails show the slides as they are now
    with tempfile.TemporaryDirectory() as td:
        tmp_deck = Path(td) / "for_thumbs.pptx"
        prs.save(str(tmp_deck))
        numbers = {s.slide_id: i for i, s in enumerate(prs.slides, 1)}
        targets = [r["slide"] for r in appendix] + [s for _, s in main_starts] + [g for _, g, _ in group_slides]
        thumbs = thumbnails(tmp_deck, [numbers[t.slide_id] for t in targets], engine)

    def png_for(slide, label):
        return thumbs.get(numbers.get(slide.slide_id)) or tile_png(label)

    def fill_tiles(host, rows, return_to_parent):
        box = dk.slot_shapes(host).get("index_list")
        if box is not None:
            dk.remove_shape(box)
        tiles, label_h = grid(len(rows))
        for k, ((label, target), (x, y, w, h)) in enumerate(zip(rows, tiles), 1):
            add_zoom(host, target, png_for(target, label), x, y, w, h, f"nav:zoom:{k}", return_to_parent)
            add_label(host, label, x, y + h + Inches(0.03), w, label_h, f"nav:label:{k}", size=10)
            report["zooms"] += 1

    if index_rows:
        index_slide = index_rows[0]["slide"]
        if group_slides:
            fill_tiles(index_slide, [(f"{g} ({len(rows)})", gs) for g, gs, rows in group_slides], True)
            for g, gs, rows in group_slides:
                fill_tiles(gs, [(short_title(r["title"]) or f"Slide {k}", r["slide"]) for k, r in enumerate(rows, 1)], True)
                add_button(gs, "‹ Back to index", Inches(10.73), Inches(6.62), Inches(2.0), Inches(0.3), "nav:back_index", target=index_slide)
        else:
            fill_tiles(index_slide, [(short_title(r["title"]) or "Backup", r["slide"]) for r in appendix], True)
        add_button(index_slide, "‹ Back to presentation", Inches(10.73), Inches(6.62), Inches(2.0), Inches(0.3),
                   "nav:back", action="ppaction://hlinkshowjump?jump=lastslideviewed")
        if nav_slide is not None:
            add_button(index_slide, "Main sections", Inches(8.6), Inches(6.62), Inches(2.0), Inches(0.3),
                       "nav:to_navigator", target=nav_slide)
            fill_tiles(nav_slide, [(sec, s) for sec, s in main_starts], False)
            add_button(nav_slide, "Appendix index", Inches(10.73), Inches(6.62), Inches(2.0), Inches(0.3),
                       "nav:to_index", target=index_slide)

        # corner link on every main slide
        if corner_links:
            for r in classify(prs):
                if r["role"] not in ("main", "main_takeaway", "title"):
                    continue
                s = r["slide"]
                dark = s.slide_layout.name in ("Blank", "Blank dark", "Frontpage dark", "Custom Layout")
                lab = add_label(s, "Appendix ›", Inches(10.55), Inches(7.16), Inches(1.4), Inches(0.24),
                                "nav:appendix_link", size=9, color="FFFFFF" if dark else "7F7F7F", align="r")
                link_shape(s, lab, target=index_slide)
                report["corner_links"] += 1

    sections = build_sections(classify(prs))
    _set_sections(prs, sections)
    report["sections"] = len(sections)
    prs.save(str(out_path))
    report["output"] = str(out_path)
    return report


def check_links(path):
    """Every nav link and zoom points at a slide in the deck. Returns problems."""
    prs = dk.open_deck(path)
    ids = {s.slide_id for s in prs.slides}
    problems = []
    for i, s in enumerate(prs.slides, 1):
        for el in s.shapes._spTree.iter():
            tag = str(el.tag)
            if tag.endswith("}sldZmObj") and int(el.get("sldId")) not in ids:
                problems.append(f"slide {i}: zoom to missing slide id {el.get('sldId')}")
            if tag.endswith("}hlinkClick"):
                rid = el.get(f"{{{dk.R_NS}}}id")
                act = el.get("action", "")
                if rid:
                    rel = s.part.rels.get(rid)
                    if rel is None:
                        problems.append(f"slide {i}: link with unknown relationship {rid}")
                    elif rel.reltype == RT.SLIDE and rel.target_part not in [x.part for x in prs.slides]:
                        problems.append(f"slide {i}: link to a slide that is not in the deck")
                elif act and "jump=" not in act and act != "ppaction://noaction":
                    problems.append(f"slide {i}: link action without target: {act}")
    return problems


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("deck")
    ap.add_argument("-o", "--out")
    ap.add_argument("--engine", default="quicklook", choices=["quicklook", "powerpoint", "auto"])
    ap.add_argument("--no-corner-links", action="store_true")
    ap.add_argument("--no-navigator", action="store_true")
    args = ap.parse_args()
    r = navigate(args.deck, args.out, args.engine, not args.no_corner_links, not args.no_navigator)
    print(f"Navigation added to {r['output']}: {r['zooms']} zooms, {r['corner_links']} corner links, "
          f"{r['sections']} sections, navigator {'yes' if r['navigator'] else 'no'}"
          + (f", {r['group_indexes']} group indexes" if r["group_indexes"] else ""))
    problems = check_links(r["output"])
    for p in problems:
        print("problem:", p)
    sys.exit(1 if problems else 0)


if __name__ == "__main__":
    main()
