#!/usr/bin/env python3
"""Edit an existing deck: add, replace, move or delete slides. Works on decks
built by case-deck and on decks made by hand. Slides you do not touch stay
exactly as they were (checked byte by byte).

Usage:
  python edit.py edit_spec.yaml            # writes <deck>_v2.pptx unless output: is given
  python edit.py edit_spec.yaml --check    # validate only

edit_spec.yaml:
  deck: our_deck.pptx
  output: our_deck_v2.pptx          # optional; never overwrites unless overwrite: true
  operations:                       # slide numbers always refer to the ORIGINAL deck
    - add: {pattern: risk_matrix, after: 12, section: Impact, slots: {...}, top_risks: [1, 3]}
    - add: {pattern: chart_takeaways, at_end: true, appendix: true, group: Financials, slots: {...}}
    - replace: {slide: 7, pattern: initiative, slots: {...}}
    - move: {slide: 9, after: 4}
    - delete: {slide: 15}
"""
from __future__ import annotations

import argparse
import hashlib
import sys
from pathlib import Path

import yaml

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build  # noqa: E402
import deckkit as dk  # noqa: E402


class EditError(Exception):
    pass


def slide_fingerprint(slide):
    """Hash of the slide XML plus the content of everything it points to."""
    h = hashlib.sha1(dk.slide_xml(slide))
    for rId, rel in sorted(slide.part.rels.items()):
        h.update(rId.encode() + rel.reltype.encode())
        if rel.is_external:
            h.update(rel.target_ref.encode())
        elif rel.reltype.endswith("/slide"):
            h.update(b"slide-link")
        elif not rel.reltype.endswith(("/slideLayout", "/notesSlide")):
            h.update(hashlib.sha1(rel.target_part.blob).digest())
    return h.hexdigest()


def validate_ops(spec, catalog, n_slides, base_dir):
    errors = []
    ops = spec.get("operations") or []
    if not ops:
        errors.append("no operations")
    fake = {"slides": []}
    for k, op in enumerate(ops, 1):
        (kind, args), = op.items()
        if kind not in ("add", "replace", "move", "delete"):
            errors.append(f"operation {k}: unknown kind {kind!r}")
            continue
        for key in ("slide", "after", "before"):
            if key in args and not 1 <= int(args[key]) <= n_slides:
                errors.append(f"operation {k}: {key} {args[key]} is outside 1-{n_slides}")
        if kind in ("add", "replace"):
            fake["slides"].append({kk: vv for kk, vv in args.items() if kk not in ("after", "before", "slide", "at_end")})
    if fake["slides"]:
        e, _ = build.validate(fake, catalog, base_dir)
        errors += [x.replace("slide ", "new slide ", 1) for x in e]
    return errors


def edit(spec_path, check_only=False):
    spec_path = Path(spec_path)
    base_dir = spec_path.parent
    spec = yaml.safe_load(spec_path.read_text(encoding="utf-8"))
    deck_path = base_dir / spec["deck"]
    catalog = build.load_catalog()
    prs = dk.open_deck(deck_path)
    original = list(prs.slides)
    errors = validate_ops(spec, catalog, len(original), base_dir)
    if errors:
        raise EditError("\n".join(errors))
    if check_only:
        return None, {}
    before = {s.slide_id: slide_fingerprint(s) for s in original}

    patterns = dk.open_deck(dk.PATTERNS)
    pslides = build.pattern_slides(patterns)
    touched, deleted = set(), set()
    report = {"placeholders": [], "unfilled": [], "open_tags": [], "warnings": [], "added": 0}

    def new_slide(args):
        pid = args["pattern"]
        s = dk.clone_slide(pslides[pid], prs)
        build.fill_slide(s, catalog["patterns"][pid], args, base_dir, report)
        meta = {"section": args.get("section"), "appendix": bool(args.get("appendix")),
                "group": args.get("group"), "id": args.get("id")}
        build.tag_slide(s, pid, {k: v for k, v in meta.items() if v not in (None, False)})
        out = [s]
        if args.get("takeaway"):
            out.append(build.add_takeaway(prs, s, pslides["takeaway_overlay"], args["takeaway"], args.get("takeaway_head")))
        report["added"] += len(out)
        return out

    # positions are anchored to original slides, so order of operations does not shift numbers
    plan = []  # (anchor_slide_or_None, 'after'|'before'|'end', [new slides]) and moves
    for op in spec["operations"]:
        (kind, args), = op.items()
        if kind == "add":
            created = new_slide(args)
            if args.get("at_end"):
                plan.append((None, "end", created))
            elif "before" in args:
                plan.append((original[int(args["before"]) - 1], "before", created))
            else:
                plan.append((original[int(args.get("after", len(original))) - 1], "after", created))
        elif kind == "replace":
            old = original[int(args["slide"]) - 1]
            created = new_slide(args)
            plan.append((old, "after", created))
            deleted.add(old.slide_id)
        elif kind == "move":
            s = original[int(args["slide"]) - 1]
            if "before" in args:
                plan.append((original[int(args["before"]) - 1], "before", [s]))
            else:
                plan.append((original[int(args["after"]) - 1], "after", [s]))
        elif kind == "delete":
            deleted.add(original[int(args["slide"]) - 1].slide_id)

    for anchor, where, items in plan:
        for k, s in enumerate(items):
            cur = list(prs.slides)
            if where == "end":
                pos = len(cur) - 1
            else:
                ai = next(i for i, x in enumerate(cur) if x.slide_id == anchor.slide_id)
                if s.slide_id == anchor.slide_id:
                    continue
                si = next(i for i, x in enumerate(cur) if x.slide_id == s.slide_id)
                if where == "after":
                    pos = ai + 1 + k
                else:
                    pos = ai + k
                if si < pos:
                    pos -= 1
            dk.move_slide(prs, s, pos)
    for s in list(prs.slides):
        if s.slide_id in deleted:
            dk.drop_slide(prs, s)

    out = base_dir / spec.get("output", deck_path.stem + "_v2.pptx")
    if out.resolve() == deck_path.resolve() and not spec.get("overwrite"):
        raise EditError("output would overwrite the original deck; set overwrite: true to allow it")
    prs.save(str(out))

    had_nav = any(sh.name.startswith("nav:") for s in original for sh in s.shapes)
    if had_nav:
        import navigate
        report["navigation"] = navigate.navigate(out, out)

    # verify: every original slide that was not replaced/deleted is unchanged
    # (the appendix index is rebuilt by the navigation on purpose)
    after = dk.open_deck(out)
    changed = []
    for s in after.slides:
        if build.slide_meta(s).get("pattern") == "appendix_index" and had_nav:
            continue
        if s.slide_id in before and s.slide_id not in deleted:
            if slide_fingerprint(s) != before[s.slide_id]:
                changed.append(s.slide_id)
    report["output"] = str(out)
    report["slides"] = len(after.slides)
    report["unchanged_ok"] = not changed
    report["changed_ids"] = changed
    return out, report


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("spec")
    ap.add_argument("--check", action="store_true")
    args = ap.parse_args()
    try:
        out, r = edit(args.spec, args.check)
    except (EditError, build.SpecError) as e:
        print("The edit spec has problems; nothing was changed:\n" + str(e), file=sys.stderr)
        sys.exit(2)
    if args.check:
        print("OK")
        return
    print(f"Wrote {out}: {r['slides']} slides, {r['added']} new.")
    print("Untouched slides unchanged: " + ("yes" if r["unchanged_ok"] else f"NO ({r['changed_ids']})"))
    for k in ("unfilled", "placeholders", "open_tags"):
        for x in r[k]:
            print(f"  {k}: {x}")
    if r.get("navigation"):
        n = r["navigation"]
        print(f"Navigation refreshed: {n['zooms']} zooms on the index, {n['corner_links']} 'Appendix ›' links.")
    sys.exit(0 if r["unchanged_ok"] else 1)


if __name__ == "__main__":
    main()
