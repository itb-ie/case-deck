# Formats: storyline.md, deck_spec.yaml, edit_spec.yaml

## 1. storyline.md (written first, approved by the team)

One block per slide. Plain Markdown, easy to edit together.

```markdown
# Storyline: [deck title] ([client], [competition], [format, e.g. 6H])

## Slide 4 · Analysis · presenter: Ana
Pattern: company_position
Title: Tajín must focus its growth where cultural relevance and profitability align
Content:
- 3 numbers: 15+ products, 65+ countries, 35+ years
- strengths: portfolio, Mexican culture, heritage brand
- map of growth options by brand fit
Numbers: 15+ products (Source: Tajín website); 65+ countries (Provided case); 35+ years (Provided case)
Place: main
```

Rules:
- `Title:` is the action title, a full sentence with the takeaway.
- Every number in `Numbers:` has a source, or `[ASSUMPTION: why]` or `[SOURCE?]`.
- `Pattern:` comes from `references/patterns.md`.
- `Place:` is `main` or `appendix (group)`.

## 2. deck_spec.yaml (what build.py reads)

```yaml
deck:
  title: Mejorando todo con Tajín      # file title
  client: Tajín                        # used by the review and the AI check
  client_logo: img/tajin_logo.png      # optional; replaces the school name bottom right
  output: tajin_deck.pptx              # file name (next to the spec)

slides:
  - pattern: title
    slots:
      title: Mejorando todo con Tajín
      hero_image: img/hero.jpg          # paths are relative to the spec file
      m1_name: Ana López                # m1..m4: fill only the people you have
      m2_name: Ben Ortiz

  - pattern: exec_summary
    section: Executive Summary          # also fills the small tag top right
    slots:
      question: How can Tajín stay relevant for Gen Z and Gen Alpha?
      c1_head: Sabores de México        # c1..c3: use 2 or 3 columns
      c1_question: How can we innovate without diluting our essence?
      c1_reco: ["**Global Flavor Playgrounds**", "Expand Yaya and reinvent Tajín through global flavors"]
      c2_head: Llegar al mundo
      c2_question: Which channels build an engaged community?
      c2_reco: ["**Pop-ups and TikTok Shop**", "Sample on fresh fruit in street markets, then sell through creator storefronts"]
      impact: "Impact: contribute USD ~130m in revenue by 2030"
      source: "Source: Team Analysis"

  - pattern: section_divider
    slots: {section: Analysis, presenter: Ana López, image: img/divider1.jpg}

  - pattern: three_cards
    section: Analysis
    takeaway: "Tajín can unite its **cultural heritage** and **bold innovation**"   # adds the Key Takeaway slide after it
    slots:
      title: Tajín can position itself where culture, experience and convenience meet
      c1_text: ["## Innovating Flavor Portfolios", "", "Playful formats like Yaya turn flavor into participation"]
      # ... c2_text, c3_text, images, source

  - pattern: framework            # used again as the pillar tracker:
    current: 1                    # brightens pillar 1, shades the others
    section: Strategy
    slots: {...}

  - pattern: financials
    section: Impact
    slots:
      chart:
        categories: ["2024", "2025E", "2026E"]
        series:
          Base growth: [300, 345, 397]
          Penetration: [0, 10, 20]
      # ... kpi1_value, kpi1_label, etc.

  - pattern: risk_matrix
    section: Impact
    top_risks: [1, 4]             # stars next to risks 1 and 4
    slots: {r1_name: ..., r1_desc: ...}   # r1..r7; unused rows disappear

  - pattern: chart_takeaways
    appendix: true                # appendix slides get zooms on the index
    group: Financials             # groups them in the index and in sections
    id: fin_model                 # optional name for links
    slots: {...}
```

Text values:
- A string, or a list of strings (one paragraph each).
- `**bold**` marks bold words; the rest of that paragraph is not bold.
- `## ` at the start of a paragraph uses the box's heading style (first paragraph of two-style boxes like KPIs).
- `\v` inside a string is a line break without a new paragraph.

Other slide fields: `section`, `notes` (speaker notes), `takeaway` and `takeaway_head`, `appendix`, `group`, `id`, `hidden`, `count` (force the number of columns or rows, e.g. `count: {c: 2}`), `current` (framework), `top_risks` (risk matrix).

Tables (`harvey_table`): `table:` is a list of rows; the first row holds the criteria, the first column the options. Harvey levels go in `h_r<row>_c<col>: 0-4`.

Check before building: `uv run --project "$SKILL" python "$SKILL/scripts/build.py" deck_spec.yaml --check`. It lists unknown slots (with similar names), text longer than the box, missing images, and required boxes left empty.

## 3. edit_spec.yaml (what edit.py reads)

```yaml
deck: our_deck.pptx
output: our_deck_v2.pptx
operations:            # slide numbers refer to the ORIGINAL deck, before any change
  - add: {pattern: mitigation, after: 12, section: Impact, slots: {...}}
  - add: {pattern: chart_takeaways, at_end: true, appendix: true, group: Financials, slots: {...}}
  - replace: {slide: 7, pattern: initiative, slots: {...}}
  - move: {slide: 9, after: 4}
  - delete: {slide: 15}
```

## 4. edits.yaml (what deslop.py apply reads)

```yaml
edits:
  - slide: 2
    box: c1_reco          # optional: the slot name shown in the worklist
    find: "the exact words to replace"
    replace: "the team's version"
    approved: true        # only approved edits are applied
```
