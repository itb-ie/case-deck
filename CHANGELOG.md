# Changelog

## v0.1.1 (2026-10-17)

- One file per deck: `build.py` adds the appendix navigation itself (`--no-nav` to skip), and `edit.py` refreshes it. Existing "Appendix ›" links on untouched slides stay as they are.
- Dropping a slide also removes links and zooms that pointed at it (fixes duplicate slide parts on save).
- Rendering a hidden slide no longer stalls Quick Look.

## v0.1.0 (2026-10-16)

First version for the case competition course.

- 26 slide patterns from the 2025 winning decks (Tajín 24H) and the Slideworks template, plus the Key Takeaway band.
- Build from `deck_spec.yaml`, with validation (unknown boxes, text too long for its box, missing images, required boxes left empty, numbers split by YAML commas).
- Edit existing decks, including decks made in other templates; untouched slides are verified byte for byte.
- Review: structure, action titles, matching numbers, sources, leftovers and other cases' names, overflow, fonts, colours, AI patterns; judge view through Claude with the course rubric.
- AI-pattern check on the avoid-ai-writing detector (pinned 3.36.0) plus a case-deck layer; de-slop applies approved edits only and protects numbers, names and formatting.
- Navigation: sections, appendix index with Slide Zooms (return to index), "Appendix ›" corner links, back button, hidden navigator.
- Tests: golden Tajín rebuild, fresh case (IE BCC 2026 Alma Carraovejas), edit of a real team deck, 10 seeded defects, de-slop, navigation, PowerPoint open check, originals untouched, self-test.

Manual slideshow test (README, "Navigation: the manual slideshow test"): **[to do by the instructor]**
