# House style: how our winning decks are built

Source: the five 2025 winning decks (Babeș-Bolyai 6H and 24H, Guadalajara 6H, 24H and 6H Final). The Tajín 24H deck is the most complete example and most patterns come from it.

## The storyline (main body: about 35-40 slides for 15-20 minutes)

| # | Part | Pattern(s) | What it does |
|---|---|---|---|
| 1 | Title | `title` | Short, memorable deck title (often in the case's language), team photos and names |
| 2 | Executive summary | `exec_summary` | Answer first: the central question, your pillars as the case questions, one recommendation per pillar, the headline impact |
| 3 | Analysis divider | `section_divider` | Section name and who presents it. Every section names its presenter |
| 4 | Company position | `company_position` | 3 big numbers, what the company is good at, where it must now decide |
| 5 | Market | `trends` | 2-3 trends, each with 2 sourced numbers and a "so what" |
| 6 | Hook | `hook_image` | One full photo, no text. The presenter tells a real story out loud |
| 7 | Synthesis | `three_cards` or `matrix_takeaways` | Where the company's strengths meet the market's best opportunities |
| 8 | Strategy divider | `section_divider` | |
| 9 | Prioritisation | `prioritization` | The criteria and which options you chose and dropped |
| 10 | Our perspective | `strategy_tree` or `issue_tree` | The central question split into areas and initiatives |
| 11 | Framework | `framework` | 2 or 3 pillars, often an acronym in the case's language (SAL, SOL, GOL, AIM) |
| 12 | Pillar blocks | `framework` with `current:`, `pillar_overview`, `initiative`, `initiative_steps` | For each pillar: tracker, overview of its initiatives, one slide per initiative with breakdown, KPIs and budget |
| 13 | Impact divider | `section_divider` | |
| 14 | Financials | `driver_tree`, `financials` | Each initiative linked to a revenue driver; revenue created, ROI, payback, assumptions |
| 15 | Risks | `risk_matrix`, `mitigation` | Risks A-G by probability and impact, then how you reduce the top ones |
| 16 | Roadmap | `roadmap` | Pilot first, then scale, with months |
| 17 | Recap | `exec_summary` | The same summary again |
| 18 | Closing | `closing` | Team photos and names |
| 19+ | Appendix | `appendix_divider`, `appendix_index`, then backup slides | Financial model, sensitivity, assumptions, funding, competitor tables, personas, the case questions answered one by one |

6H and 24H decks have the same main body. A 24H deck adds depth in the appendix: the financial model, sensitivity, funding sources, personas, full risk tables.

## Rules for every slide

1. **Action title.** A full sentence with the takeaway: who does what, with which result. Use a number when you have one. "Phase implementation by deploying pilot programs on select regions and SKUs, then scale with an integrated narrative" is an action title; "Implementation plan" is a topic title. Test: read only the titles from start to end. They must tell the whole story.
2. **One message per slide.** If a slide needs two titles, it is two slides.
3. **Source line on every slide with data.** "Source: Team Analysis" is fine for your own estimates. Never invent a number. If you have to estimate, write `[ASSUMPTION: ...]` and explain the assumption in the appendix. If you still need to find a source, write `[SOURCE?]`; the review flags both.
4. **2 or 3 pillars.** Choose what fits the case. Two pillars are often short term and long term. Three is common but never a rule.
5. **Branded initiative names.** Give each initiative a short name in the case's language where it fits: "Manos Vivas", "Hecho Juntos", "Carrito de Tajín". Judges remember names.
6. **The hook.** A full photo placed after the company context, with a real person or moment told out loud. Start with a specific person, place or number <!-- avoid-ai-writing:ignore-start -->(not "Imagine a world where...")<!-- avoid-ai-writing:ignore-end -->.
7. **Key Takeaway band.** Close important sections with the dark band (add `takeaway:` to a slide in the spec). One sentence the judges should remember, with the key phrases in bold.
8. **Presenter per section.** Each divider names who presents. Hand-overs happen at dividers.
9. **Numbers match everywhere.** The impact on the executive summary, the financials and the recap must be the same figure.
10. **Look.** Arial; navy 061F32, cyan 22A3DF, blue 3478AB, dark blue 1D4769, gold F8BC3B, light grey F2F2F2; 12-14 pt body, 9 pt sources, 20-36 pt big numbers. The client's logo bottom right on every slide.

## What the template already does for you

- The section tag top right (the `tracker` slot) is filled from `section:` in the spec.
- Page numbers come from the layout.
- `navigate.py` adds the appendix index with zooms, a link to it on every main slide, and a hidden navigator of the main sections.

## Mistakes seen in the 2025 decks (the review checks for them)

- Slides copied from an older deck still showing another company (MOL slides inside the Artex deck).
- "HOOK TBD" left as a title.
- The executive summary says USD ~130m while the financials say USD 135m. Pick one figure.
- Empty text boxes and leftover template text.
