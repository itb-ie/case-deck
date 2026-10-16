# The AI-pattern pass

Judges increasingly recognise AI-written slides, and generic wording costs credibility. This pass finds AI writing patterns and helps the team rewrite them in its own voice. It never changes facts.

## Engine

`scripts/ai_check.py` uses the **avoid-ai-writing** detector by Conor Bronsdon (github.com/conorbronsdon/avoid-ai-writing, MIT licence), npm package `avoid-ai-writing-detector` pinned at 3.36.0 in `package.json`. It covers 50+ pattern families: inflated significance, promotional words, vague attribution, a 3-tier word list with plain replacements, filler phrases, uniform sentence rhythm, title-case headings, chatbot leftovers.

On top, a small case-deck layer catches what the detector misses in decks. The examples below are specimens of what to avoid:

<!-- avoid-ai-writing:ignore-start -->
- em-dashes (—);
- "it's not X, it's Y" and "not just X, but Y";
- "In today's fast-paced world";
- "Imagine a world where...";
- chatbot sentences ("I hope this helps", "let me know if");
- strategy slides that never name the client (they could fit any company).
<!-- avoid-ai-writing:ignore-end -->

The detector's command-line tool ignores the `<!-- avoid-ai-writing:ignore-start -->` / `<!-- avoid-ai-writing:ignore-end -->` markers, so `ai_check.py` removes marked passages before checking. Use the markers in notes and storylines that quote bad examples on purpose.

## Baseline and limit

The five 2025 winning decks score 4-8 ("Minimal AI signals") with 1.2-3.1 findings per 10 slides. The default limit is 4 findings per 10 slides and a score of 25. An AI-default version of the Tajín storyline scores 51 with over 100 findings per 10 slides.

A single flagged word proves nothing: <!-- avoid-ai-writing:ignore-start -->"leverage" or "seamless"<!-- avoid-ai-writing:ignore-end --> can be the right word. Judges react to clusters of these words combined with vague content. The cure is specific: real numbers, real names and places, the client's products, branded initiative names, short sentences.

## The workflow (mode 4 in SKILL.md)

1. `uv run --project "$SKILL" python "$SKILL/scripts/deslop.py" propose deck.pptx > worklist.json` lists every flagged paragraph with the slide, the box, what was flagged and a suggested plain word.
2. Claude drafts a rewrite for each paragraph, in the team's voice, and shows a before/after table per slide.
3. The team approves, edits or rejects each line. Only lines marked `approved: true` go into `edits.yaml`.
4. `uv run --project "$SKILL" python "$SKILL/scripts/deslop.py" apply deck.pptx edits.yaml -o deck_v2.pptx` applies them run by run. Fonts, sizes, colours and bold stay.
5. The score before and after is printed. If it is still above the limit, repeat for the remaining slides.

## What never changes

`deslop.py` compares each edited box before and after and refuses to save if any of these changed: numbers, percentages, currencies, dates, proper nouns (companies, places, people, products) and quoted branded names. A rewrite that adds a new fact is also refused: new facts belong in the storyline, with a source, and the team decides about them.

A human team member makes the final call on wording. An AI rewrite of AI text still needs a human read.
