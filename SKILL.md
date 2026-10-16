---
name: case-deck
description: Build, edit, review and polish business case competition decks in the format of our 2025 winning decks (Slideworks template). Use when a team asks to make slides or a deck for a case, add or change slides in their deck, review or judge a deck, remove AI-sounding text, or add appendix navigation (Slide Zooms, links) for Q&A.
---

# case-deck

Turns a team's thinking into slides in the house style of the 2025 winning decks (Babeș-Bolyai and Guadalajara). The scripts do the layout; the team owns the storyline, every number and the final wording.

All paths below are relative to this skill's folder. Run scripts with `python3 <skill folder>/scripts/<name>.py`. Work files (storyline, spec, decks) go in the team's own project folder, never inside the skill folder.

## First message in a session

Say, in one or two lines:
- check your competition's rules on AI before using this during the case: some ban it (Queen's ICBC) or ban all electronics (Champions Trophy), others say nothing (CBS, BBU-ICC). When rules are silent, ask the organisers in writing, and disclose AI use where required.
- If `python3 scripts/selftest.py` has not been run on this laptop yet, run it now.

## Pick the mode from the request

| The team says... | Mode |
|---|---|
| "make our deck / slides from these notes / from the case" | 1 Build |
| "add a slide / replace slide 7 / move / delete" | 2 Edit |
| "review / check / judge our deck", "what will judges ask" | 3 Review |
| "make it sound less like AI", "de-slop", "rewrite the AI bits" | 4 AI-pattern pass |
| "add the appendix zooms / navigation / links" | 5 Navigation |

## Mode 1: Build

1. Read everything the team gives: notes, the case PDF, numbers, photos, an existing storyline. Read `references/house_style.md` and `references/patterns.md`.
2. If something essential is missing, ask at most 5 short questions: the case questions, the competition format (6H, 24H...), time left, who presents which section, 2 or 3 pillars.
3. Write `storyline.md` in the team's folder (format: `references/storyline_format.md`). Follow the house skeleton. Every number gets a source, or `[ASSUMPTION: ...]` or `[SOURCE?]`. Never invent a number, a statistic or a quote.
4. **Stop. Show the storyline and wait until the team approves or edits it.** No slides before that.
5. Write `deck_spec.yaml` from the approved storyline. Run `python3 scripts/build.py deck_spec.yaml --check` and fix every error (box too long: shorten the text, never shrink fonts).
6. Build: `python3 scripts/build.py deck_spec.yaml`. Then `python3 scripts/navigate.py <deck>.pptx`. Then render: `python3 scripts/render.py <deck>.pptx --sheet sheet.png` and look at the sheet.
7. Self-review: `python3 scripts/review.py <deck>.pptx`. Fix overflow, overlaps and leftovers in the spec and rebuild, **at most 2 rounds**.
8. Report: the deck path, slide count, open `[ASSUMPTION]`/`[SOURCE?]` items, image placeholders still grey, and what needs a human polish (photos, builds, final wording).

Images the team did not give stay as labelled grey placeholders. Never use stock or AI-generated images.

## Mode 2: Edit

1. Write `edit_spec.yaml` (format in `references/storyline_format.md`): `add`, `replace`, `move`, `delete`. Slide numbers refer to the deck as it is now.
2. `python3 scripts/edit.py edit_spec.yaml --check`, then without `--check`. It writes `<deck>_v2.pptx` and confirms that every untouched slide is unchanged.
3. If the deck had navigation, run `navigate.py` on the new file.
4. Works on decks made by hand in other templates too: new slides keep our style and colours.

## Mode 3: Review (slide quality plus the judge view)

1. Run `python3 scripts/review.py <deck>.pptx --client "<company>" --out review.md`. Render the deck and look at every slide.
2. Report, ranked by severity, slide by slide, with a concrete fix each:
   - what `review.py` found (structure, titles, numbers that do not match, missing sources, leftovers, other companies' names, overflow, fonts, colours, AI patterns);
   - what you see in the renders that the script cannot: crowded slides, overlaps, unclear charts, weak hooks;
   - the "read only the titles" test: do the titles alone tell the story?
3. **Judge view**, with `references/rubric.md`:
   - for each criterion, quote evidence from the deck (slide and words) **before** the score;
   - score twice, independently; report criteria where the two scores differ by more than 1;
   - list the 10-20 most likely judge questions, tagged (assumption, why-not-X, implementation, financials, risk, curveball);
   - list the **appendix gaps**: likely questions with no backup slide, each with a proposed backup slide (pattern + action title) the team can add with mode 2;
   - say that the score is feedback: AI judges favour long text and squeeze scores to the middle; human judges differ.

## Mode 4: AI-pattern pass

Read `references/deslop.md` first.
1. `python3 scripts/ai_check.py <deck>.pptx --client "<company>"` for the score, then `python3 scripts/deslop.py propose <deck>.pptx > worklist.json`.
2. For each flagged paragraph, draft a rewrite in the team's voice: specific numbers and names, the client's own words, short sentences. Do not add facts that are not already in the deck.
3. Show a before/after table per slide. **Apply only what the team approves:** write those lines to `edits.yaml` with `approved: true`.
4. `python3 scripts/deslop.py apply <deck>.pptx edits.yaml -o <deck>_v2.pptx`. If it refuses (a number or name would change), fix that rewrite and run again. Report the score before and after.

## Mode 5: Navigation

`python3 scripts/navigate.py <deck>.pptx` adds, as in the Zoom demo deck:
- PowerPoint sections (one per storyline section and appendix group);
- Slide Zooms to every backup slide on the appendix index; after a backup slide, PowerPoint returns to the index;
- a "‹ Back to presentation" button on the index (last slide viewed);
- a small "Appendix ›" link on every main slide;
- a hidden navigator slide after the title with zooms to each main section.

Re-running replaces earlier navigation. Use `--engine powerpoint` for sharper thumbnails on a Mac with PowerPoint.

## Things that are always true

- Pillars are 2 or 3, whichever fits the case. Never present 3 as the rule.
- Plain words for students. No internal codes.
- Text you write for the team (storylines, slide text, reviews) follows `references/deslop.md`: no em-dashes, no "not X but Y" frames, no stock phrases.
- Never modify the files in `assets/`.
- Mark anything you are unsure about as **[to confirm]**.
