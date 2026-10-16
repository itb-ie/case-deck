# case-deck

A Claude Code skill that helps case competition teams build slides fast, in the format of our 2025 winning decks (Babeș-Bolyai and Guadalajara), on the Slideworks template.

It can:
1. **Build** a deck from rough notes, a storyline or the case PDF (a storyline first, slides only after the team approves it);
2. **Edit** an existing deck: add, replace, move or delete slides, also in decks made by hand;
3. **Review** a deck for slide quality and as a judge would (rubric, likely questions, missing backup slides);
4. **Remove AI writing patterns**, only with the team's approval, never changing a number or a name;
5. Add **appendix navigation** for Q&A: Slide Zooms, an "Appendix ›" link on every slide, a back button.

The scripts do the layout. The team owns the storyline, every number and the final wording.

> **Before a competition:** check its rules on AI. Some ban it (Queen's ICBC) or ban all electronics (Champions Trophy); others say nothing (CBS, BBU-ICC). If the rules are silent, ask the organisers in writing, and disclose AI use where it is required.

## Install (once per laptop)

You need Claude Code, Python 3.9+ and Node.js 18+.

```bash
# put the skill where Claude Code finds it: in your project...
git clone <repo url> <your project>/.claude/skills/case-deck
# ...or for all your projects
git clone <repo url> ~/.claude/skills/case-deck

cd <that folder>
python3 -m pip install -r requirements.txt
npm install                      # the AI-pattern detector, so it also works offline
python3 scripts/selftest.py      # should end with "All good"
```

Update during the semester with `git pull` in the same folder.

## Use it

Open Claude Code in your team's project folder and talk normally:

- "Build our deck from notes.md and the case PDF. We are four, 6H format, presenters Ana, Ben, Carla, Diego."
- "Add a risk matrix after slide 12 with these risks..."
- "Review our deck like a judge. What will they ask?"
- "Some slides sound like AI. Fix them, but show me first."
- "Add the appendix zooms."

Claude writes a `storyline.md` first and waits for your OK. Then it writes `deck_spec.yaml` and builds the deck. You can also edit the spec yourself and rebuild: `python3 scripts/build.py deck_spec.yaml`.

What is in the skill:

| Path | What |
|---|---|
| `SKILL.md` | Instructions Claude follows for each mode |
| `assets/template.pptx` | The Slideworks template (do not edit) |
| `assets/patterns.pptx` | 26 slide patterns taken from the winning decks and the template |
| `assets/catalog.yaml`, `references/patterns.md` | What each pattern is for and its fillable boxes |
| `references/` | House style, judging rubric, spec formats, the AI-pattern pass |
| `scripts/` | `build`, `edit`, `navigate`, `render`, `review`, `ai_check`, `deslop`, `inspect_deck`, `selftest` |

## Speed (measured on a 2024 MacBook, PowerPoint installed)

| Step | Time |
|---|---|
| Build 37 slides from a spec | about 1 s |
| Add one slide to a 37-slide deck | under 1 s |
| Navigation (zooms, links, sections) | 2-3 s |
| Review (with AI check) | about 1 s |
| Render all slides (Quick Look / PowerPoint) | 0.5 s per slide / about 7 s for the deck |

## AI-pattern check: baseline

`scripts/ai_check.py` uses the [avoid-ai-writing](https://github.com/conorbronsdon/avoid-ai-writing) detector by Conor Bronsdon (MIT licence), npm package `avoid-ai-writing-detector` pinned at 3.36.0, plus a small case-deck layer (see `references/deslop.md`).

Scores of the 2025 winning decks (context `general`; `marketing` gives the same results):

| Deck | Slides | Score | Label | Findings per 10 slides |
|---|---|---|---|---|
| Romania 6H (RIVUS) | 41 | 5 | Minimal AI signals | 1.2 |
| Romania 24H (MOL) | 78 | 8 | Minimal AI signals | 2.4 |
| Mexico 6H (Artex) | 50 | 7 | Minimal AI signals | 1.4 |
| Mexico 24H (Tajín) | 71 | 7 | Minimal AI signals | 3.1 |
| Mexico 6H Final (Atlas FC) | 57 | 4 | Minimal AI signals | 1.2 |
| AI-default rewrite of the Tajín storyline (test fixture) | 4 | 51 | Moderate AI signals | 102.5 |

Default limit: 4 findings per 10 slides and a score of 25. All winning decks pass; the AI-default text fails. **[to confirm: how much AI the 2025 team used while writing these decks.]**

## Navigation: the manual slideshow test

Scripts check that every zoom and link points to the right slide, and that PowerPoint opens and exports the deck. The jump-back behaviour can only be checked by a person, once per release:

1. Open a navigated deck in PowerPoint and start the slideshow on a main slide (for example slide 5).
2. Click "Appendix ›" bottom right: the appendix index opens.
3. Click a backup slide's thumbnail: it zooms in.
4. Click to advance: PowerPoint returns to the index.
5. Click "‹ Back to presentation": you are back on slide 5.

Record the result in `CHANGELOG.md`.

## Known limits (v0.1)

- **Waterfall charts** from the template (Office 2016 chart type) are not in the pattern library; use `chart_takeaways` or build a waterfall by hand.
- **Rendering on Windows** needs PowerPoint (`--engine powerpoint`, untested on Windows so far). Without a renderer, builds and reviews still work; thumbnails become plain labelled tiles.
- **Quick Look draws pie shapes (Harvey balls) wrongly**; PowerPoint is correct. Use `--engine powerpoint` to check those slides.
- **Text fit** is estimated from box sizes and font sizes. Always look at the render.
- **Inline bold and italics** come only from `**bold**` in the spec; italics are set by hand.
- **Edits into another team's template**: new slides keep our look and colours; theme colours are frozen to our template.

## For the instructor

- Rebuild the pattern library after changing a source slide or a selector: `python3 scripts/curate.py --winning "<path to winning_cases>"`, then `python3 scripts/catalog_doc.py`.
- Run the tests (they look for `../winning_cases` and `../case comp cases`, or set `CASE_DECK_WINNING` and `CASE_DECK_CASES`):

  ```bash
  python3 -m pip install pytest pypdf
  PYTEST_DISABLE_PLUGIN_AUTOLOAD=1 python3 -m pytest -q
  ```

## Credits

- Patterns: the 2025 IE University case competition teams' winning decks, and the Slideworks template (licence for class use confirmed by the instructor).
- AI-pattern detection: avoid-ai-writing by Conor Bronsdon (MIT).
- Zoom navigation modelled on a PowerPoint Slide Zoom demo deck.
