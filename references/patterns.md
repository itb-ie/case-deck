# Pattern catalogue

Generated from `assets/catalog.yaml` by `scripts/catalog_doc.py`. Every slide in a
case-deck deck comes from one of these patterns. Slot names go under `slots:` in
`deck_spec.yaml`; `max` is the number of characters the box holds.

Contact sheet: `tests/output/patterns_contact_sheet.png` (in the skill folder: `uv run scripts/render.py assets/patterns.pptx --sheet sheet.png`).

## title

Cover slide: full-width photo, deck title, team photos and names, client logo.

When: First slide of every deck.  
From: Mexico 24H Case.pptx slide 1

Repeats: `m1_...` to `m4_...` (2 to 4 groups; fill only the ones you use).

| Slot | Type | Max | Notes |
|---|---|---|---|
| `title` | text | 45 |  |
| `hero_image` | image |  | image (hero photo of the client) |
| `client_logo` | image |  | optional; image (client logo) |
| `m1_photo` | photo |  | optional |
| `m2_photo` | photo |  | optional |
| `m3_photo` | photo |  | optional |
| `m4_photo` | photo |  | optional |
| `m1_name` | text | 15 |  |
| `m2_name` | text | 15 |  |
| `m3_name` | text | 17 |  |
| `m4_name` | text | 16 |  |

## exec_summary

Answer-first overview: the central question, case questions as pillars, one recommendation per pillar, headline impact.

When: Slide 2, and again just before the closing slide.  
From: Mexico 24H Case.pptx slide 2

Repeats: `c1_...` to `c3_...` (2 to 3 groups; fill only the ones you use).

| Slot | Type | Max | Notes |
|---|---|---|---|
| `title` | text | 150 |  |
| `tracker` | text | 43 |  |
| `question` | text | 213 |  |
| `c1_num` | text | 8 |  |
| `c2_num` | text | 8 |  |
| `c3_num` | text | 8 |  |
| `c1_head` | text | 58 |  |
| `c2_head` | text | 58 |  |
| `c3_head` | text | 58 |  |
| `c1_question` | text | 144 |  |
| `c2_question` | text | 148 |  |
| `c3_question` | text | 148 |  |
| `c1_reco` | text | 141 |  |
| `c2_reco` | text | 166 |  |
| `c3_reco` | text | 138 |  |
| `impact` | text | 194 |  |
| `source` | text | 124 |  |

## section_divider

Section break with a full-bleed photo, the section name and who presents it.

When: Before Analysis, Strategy, Impact (each presenter hand-off).  
From: Mexico 24H Case.pptx slide 3

| Slot | Type | Max | Notes |
|---|---|---|---|
| `section` | text | 21 |  |
| `presenter` | text | 45 |  |
| `image` | image |  | image (section photo) |

## company_position

Where the company stands: 3 big numbers, what they mean, core competencies, and a growth-potential map.

When: First Analysis slide.  
From: Mexico 24H Case.pptx slide 4

| Slot | Type | Max | Notes |
|---|---|---|---|
| `title` | text | 150 |  |
| `tracker` | text | 43 |  |
| `left_head` | text | 47 |  |
| `right_head` | text | 62 |  |
| `stat1_value` | text | 8 |  |
| `stat2_value` | text | 8 |  |
| `stat3_value` | text | 8 |  |
| `stat1_label` | text | 31 |  |
| `stat2_label` | text | 33 |  |
| `stat3_label` | text | 31 |  |
| `strength1` | text | 71 |  |
| `strength2` | text | 71 |  |
| `strength3` | text | 65 |  |
| `comp_head` | text | 48 |  |
| `comp1` | text | 48 |  |
| `comp2` | text | 46 |  |
| `comp3` | text | 62 |  |
| `map_head` | text | 32 |  |
| `map_x` | text | 29 |  |
| `map1` | text | 66 | optional |
| `map2` | text | 54 | optional |
| `map3` | text | 42 | optional |
| `map4` | text | 42 | optional |
| `map5` | text | 28 | optional |
| `map6` | text | 60 | optional |
| `source` | text | 124 |  |

## trends

Market or customer trends: 2-3 columns, each with a photo, 2 sourced numbers and a one-line takeaway.

When: Analysis section, after the company slide.  
From: Mexico 24H Case.pptx slide 5

Repeats: `t1_...` to `t3_...` (2 to 3 groups; fill only the ones you use).

| Slot | Type | Max | Notes |
|---|---|---|---|
| `title` | text | 150 |  |
| `tracker` | text | 43 |  |
| `t1_image` | image |  | image (trend 1 photo) |
| `t1_head` | text | 21 |  |
| `t1_stat1` | text | 10 |  |
| `t1_label1` | text | 39 |  |
| `t1_stat2` | text | 9 |  |
| `t1_label2` | text | 47 |  |
| `t1_takeaway` | text | 41 |  |
| `t2_image` | image |  | image (trend 2 photo) |
| `t2_head` | text | 24 |  |
| `t2_stat1` | text | 9 |  |
| `t2_label1` | text | 47 |  |
| `t2_stat2` | text | 9 |  |
| `t2_label2` | text | 36 |  |
| `t2_takeaway` | text | 58 |  |
| `t3_image` | image |  | image (trend 3 photo) |
| `t3_head` | text | 25 |  |
| `t3_stat1` | text | 9 |  |
| `t3_label1` | text | 42 |  |
| `t3_stat2` | text | 9 |  |
| `t3_label2` | text | 43 |  |
| `t3_takeaway` | text | 65 |  |
| `source` | text | 149 |  |

## hook_image

The hook: one full-bleed photo, no text. The presenter tells the human story out loud.

When: After the company context, before the analysis gets deep.  
From: Mexico 24H Case.pptx slide 6

| Slot | Type | Max | Notes |
|---|---|---|---|
| `image` | image |  | image (hook photo (a real person or moment)) |

## matrix_takeaways

A 2x2 positioning map (e.g. products by profitability vs relevance) with key takeaways on the right.

When: Analysis: to show where to play.  
From: Mexico 24H Case.pptx slide 7

| Slot | Type | Max | Notes |
|---|---|---|---|
| `title` | text | 150 |  |
| `tracker` | text | 43 |  |
| `chart_head` | text | 72 |  |
| `y_axis` | text | 39 |  |
| `x_axis` | text | 33 |  |
| `takeaways_head` | text | 35 |  |
| `takeaways` | text | 544 |  |
| `item1` | text | 60 | optional |
| `item2` | text | 26 | optional |
| `item3` | text | 60 | optional |
| `item4` | text | 20 | optional |
| `item5` | text | 60 | optional |
| `item6` | text | 27 | optional |
| `item7` | text | 27 | optional |
| `item8` | text | 27 | optional |
| `source` | text | 140 |  |

## three_cards

2-3 photo cards, each with a bold heading and a short paragraph (e.g. strengths x opportunities).

When: Analysis synthesis, or any 'three things' message.  
From: Mexico 24H Case.pptx slide 8

Repeats: `c1_...` to `c3_...` (2 to 3 groups; fill only the ones you use).

| Slot | Type | Max | Notes |
|---|---|---|---|
| `title` | text | 150 |  |
| `subtitle` | text | 122 | optional |
| `tracker` | text | 43 |  |
| `c1_image` | image |  | image (card 1 photo) |
| `c2_image` | image |  | image (card 2 photo) |
| `c3_image` | image |  | image (card 3 photo) |
| `c1_text` | text | 187 |  |
| `c2_text` | text | 210 |  |
| `c3_text` | text | 174 |  |
| `source` | text | 124 |  |

## prioritization

Impact x feasibility matrix: the criteria used, the options as labelled dots in clusters, a legend and the conclusion.

When: Start of Strategy: show which options you chose and which you rejected.  
From: Mexico 24H Case.pptx slide 11

| Slot | Type | Max | Notes |
|---|---|---|---|
| `title` | text | 150 |  |
| `tracker` | text | 43 |  |
| `criterion1` | text | 44 |  |
| `criterion2` | text | 44 |  |
| `criterion3` | text | 46 |  |
| `criterion4` | text | 46 |  |
| `legend1` | text | 185 |  |
| `legend2` | text | 55 |  |
| `legend3` | text | 57 |  |
| `legend4` | text | 80 |  |
| `conclusion` | text | 106 |  |
| `opt1` | text | 30 | optional |
| `opt2` | text | 30 | optional |
| `opt3` | text | 30 | optional |
| `opt4` | text | 20 | optional |
| `opt5` | text | 30 | optional |
| `opt6` | text | 36 | optional |
| `opt7` | text | 30 | optional |
| `opt8` | text | 30 | optional |
| `opt9` | text | 14 | optional |
| `opt10` | text | 22 | optional |
| `opt11` | text | 34 | optional |
| `opt12` | text | 28 | optional |
| `opt13` | text | 24 | optional |
| `opt14` | text | 30 | optional |
| `opt15` | text | 50 | optional |
| `opt16` | text | 30 | optional |
| `opt17` | text | 20 | optional |
| `source` | text | 124 |  |

## strategy_tree

'Our perspective': the central question split into 3 areas, each with sub-areas and the initiatives under them.

When: Strategy section, after prioritisation. Delete boxes you do not need.  
From: Mexico 24H Case.pptx slide 12

| Slot | Type | Max | Notes |
|---|---|---|---|
| `title` | text | 150 |  |
| `tracker` | text | 43 |  |
| `question` | text | 216 |  |
| `b1` | text | 33 |  |
| `b2` | text | 33 |  |
| `b3` | text | 33 |  |
| `b1_l1` | text | 33 | optional |
| `b1_l2` | text | 35 | optional |
| `b1_l3` | text | 33 | optional |
| `b2_m1` | text | 16 | optional |
| `b2_m1_l1` | text | 36 | optional |
| `b2_m1_l2` | text | 33 | optional |
| `b2_m2` | text | 13 | optional |
| `b2_m2_l1` | text | 33 | optional |
| `b3_m1` | text | 13 | optional |
| `b3_m1_l1` | text | 33 | optional |
| `b3_m2` | text | 13 | optional |
| `b3_m2_l1` | text | 33 | optional |
| `b3_m2_l2` | text | 33 | optional |
| `source` | text | 124 |  |

## issue_tree

Regular issue tree: main question, 2 drivers, 2 secondary drivers each, 2 underlying drivers each.

When: Analysis or appendix, when the structure is symmetric.  
From: PowerPoint Template.pptx slide 174

| Slot | Type | Max | Notes |
|---|---|---|---|
| `title` | text | 150 |  |
| `tracker` | text | 43 |  |
| `root` | text | 104 |  |
| `d1` | text | 77 |  |
| `d2` | text | 77 |  |
| `d1_s1` | text | 32 | optional |
| `d1_s2` | text | 32 | optional |
| `d2_s1` | text | 32 | optional |
| `d2_s2` | text | 32 | optional |
| `d1_s1_u1` | text | 50 | optional |
| `d1_s1_u2` | text | 50 | optional |
| `d1_s2_u1` | text | 50 | optional |
| `d1_s2_u2` | text | 50 | optional |
| `d2_s1_u1` | text | 50 | optional |
| `d2_s1_u2` | text | 50 | optional |
| `d2_s2_u1` | text | 50 | optional |
| `d2_s2_u2` | text | 50 | optional |
| `source` | text | 140 |  |

## framework

The strategy framework (acronym): 2-3 photo cards with number, pillar name and the 'How do we...?' question.

When: Strategy section. Also used as the pillar tracker (set `current:`).  
From: Mexico 24H Case.pptx slide 13

Repeats: `p1_...` to `p3_...` (2 to 3 groups; fill only the ones you use).

Pillar tracker: add `current: 2` to brighten pillar 2.

| Slot | Type | Max | Notes |
|---|---|---|---|
| `title` | text | 150 |  |
| `subtitle` | text | 107 |  |
| `tracker` | text | 43 |  |
| `p1_image` | image |  | image (pillar 1 photo) |
| `p2_image` | image |  | image (pillar 2 photo) |
| `p3_image` | image |  | image (pillar 3 photo) |
| `p1_shade` | marker |  |  |
| `p2_shade` | marker |  |  |
| `p3_shade` | marker |  |  |
| `p1_num` | text | 8 |  |
| `p2_num` | text | 8 |  |
| `p3_num` | text | 8 |  |
| `p1_text` | text | 128 |  |
| `p2_text` | text | 128 |  |
| `p3_text` | text | 140 |  |
| `source` | text | 124 |  |

## pillar_overview

A pillar's key strategies: 2-3 rows, each a branded initiative, what it does, and a success case with numbers.

When: First slide of each pillar block.  
From: Mexico 24H Case.pptx slide 15

Repeats: `i1_...` to `i3_...` (2 to 3 groups; fill only the ones you use).

| Slot | Type | Max | Notes |
|---|---|---|---|
| `title` | text | 150 |  |
| `tracker` | text | 43 |  |
| `left_head` | text | 29 |  |
| `right_head` | text | 20 |  |
| `i1_num` | text | 8 |  |
| `i2_num` | text | 8 |  |
| `i3_num` | text | 8 |  |
| `i1_name` | text | 65 |  |
| `i2_name` | text | 65 |  |
| `i3_name` | text | 65 |  |
| `i1_desc` | text | 172 |  |
| `i2_desc` | text | 186 |  |
| `i3_desc` | text | 186 |  |
| `i1_case` | text | 204 |  |
| `i2_case` | text | 155 |  |
| `i3_case` | text | 155 |  |
| `i1_logo` | image |  | optional; image (case logo) |
| `i2_logo` | image |  | optional; image (case logo) |
| `i3_logo` | image |  | optional; image (case logo) |
| `i1_stat1` | text | 8 |  |
| `i1_label1` | text | 39 |  |
| `i1_stat2` | text | 8 | optional |
| `i1_label2` | text | 34 | optional |
| `i2_stat1` | text | 8 |  |
| `i2_label1` | text | 108 |  |
| `i3_stat1` | text | 8 |  |
| `i3_label1` | text | 16 |  |
| `i3_stat2` | text | 8 | optional |
| `i3_label2` | text | 38 | optional |
| `source` | text | 124 |  |

## initiative

One initiative in detail: breakdown in 3 points, partners, 3 KPIs with a target, and an example image.

When: Inside each pillar block, one per initiative.  
From: Mexico 24H Case.pptx slide 18

| Slot | Type | Max | Notes |
|---|---|---|---|
| `title` | text | 150 |  |
| `tracker` | text | 43 |  |
| `breakdown_head` | text | 37 |  |
| `b1_text` | text | 86 |  |
| `b2_text` | text | 100 |  |
| `b3_text` | text | 79 |  |
| `partners_head` | text | 42 | optional |
| `partner1` | image |  | optional; image (partner logo) |
| `partner2` | image |  | optional; image (partner logo) |
| `partner3` | image |  | optional; image (partner logo) |
| `partner4` | image |  | optional; image (partner logo) |
| `kpis_head` | text | 20 |  |
| `kpis` | text | 368 |  |
| `example_head` | text | 20 |  |
| `example_image` | image |  | image (example / mock-up) |
| `source` | text | 124 |  |

## initiative_steps

An initiative as 3 steps (e.g. Discover, Prototype, Launch), with example images, a target-audience number and 2 KPIs.

When: Inside a pillar block, when the initiative is a process.  
From: Mexico 24H Case.pptx slide 17

| Slot | Type | Max | Notes |
|---|---|---|---|
| `title` | text | 150 |  |
| `subtitle` | text | 131 |  |
| `tracker` | text | 43 |  |
| `breakdown_head` | text | 34 |  |
| `s1_name` | text | 44 |  |
| `s2_name` | text | 44 |  |
| `s3_name` | text | 44 |  |
| `s1_text` | text | 140 |  |
| `s2_text` | text | 138 |  |
| `s3_text` | text | 138 |  |
| `example_head` | text | 35 |  |
| `ex1_image` | image |  | image (example 1) |
| `ex2_image` | image |  | optional; image (example 2) |
| `ex3_image` | image |  | optional; image (example 3) |
| `target_head` | text | 17 |  |
| `target_stat` | text | 8 |  |
| `target_label` | text | 100 |  |
| `kpis_head` | text | 33 |  |
| `kpi1` | text | 83 |  |
| `kpi2` | text | 83 |  |
| `source` | text | 124 |  |

## driver_tree

Revenue driver tree (e.g. penetration x frequency x price) linked to each area's initiatives by number.

When: First Impact slide: how each initiative moves a revenue driver.  
From: Mexico 24H Case.pptx slide 33

Repeats: `a1_...` to `a3_...` (2 to 3 groups; fill only the ones you use).

| Slot | Type | Max | Notes |
|---|---|---|---|
| `title` | text | 150 |  |
| `tracker` | text | 43 |  |
| `left_head` | text | 47 |  |
| `right_head` | text | 47 |  |
| `root` | text | 39 |  |
| `d1` | text | 39 |  |
| `d2` | text | 39 |  |
| `d3` | text | 39 |  |
| `op1` | text | 8 |  |
| `op2` | text | 8 |  |
| `a1` | text | 44 |  |
| `a2` | text | 44 |  |
| `a3` | text | 44 |  |
| `a1_s1` | text | 36 |  |
| `a1_s2` | text | 36 | optional |
| `a1_s3` | text | 36 | optional |
| `a2_s1` | text | 36 |  |
| `a2_s2` | text | 36 | optional |
| `a2_s3` | text | 36 | optional |
| `a3_s1` | text | 36 |  |
| `a3_s2` | text | 36 | optional |
| `a3_s3` | text | 43 | optional |
| `a1_s1_driver` | text | 8 |  |
| `a1_s2_driver` | text | 8 | optional |
| `a1_s3_driver` | text | 8 | optional |
| `a2_s1_driver` | text | 8 |  |
| `a2_s2_driver` | text | 8 | optional |
| `a2_s3_driver` | text | 8 | optional |
| `a3_s1_driver` | text | 8 |  |
| `a3_s2_driver` | text | 8 | optional |
| `a3_s3_driver` | text | 8 | optional |
| `footnote` | text | 229 | optional |
| `source` | text | 124 |  |

## financials

Financial impact: headline KPI pills (revenue, ROI, payback), key assumptions, driver increases and costs, and a base vs strategy chart.

When: Impact section, right after the driver tree.  
From: Mexico 24H Case.pptx slide 35

| Slot | Type | Max | Notes |
|---|---|---|---|
| `title` | text | 150 |  |
| `subtitle` | text | 95 |  |
| `tracker` | text | 43 |  |
| `kpi1_value` | text | 16 |  |
| `kpi1_label` | text | 20 |  |
| `kpi2_value` | text | 9 |  |
| `kpi2_label` | text | 8 |  |
| `kpi3_value` | text | 8 |  |
| `kpi3_label` | text | 10 |  |
| `as1_value` | text | 8 |  |
| `as1_label` | text | 39 |  |
| `as2_value` | text | 8 | optional |
| `as2_label` | text | 35 | optional |
| `d1_name` | text | 12 |  |
| `d2_name` | text | 10 |  |
| `d3_name` | text | 14 |  |
| `inc_head` | text | 14 |  |
| `cost_head` | text | 16 |  |
| `d1_inc` | text | 8 |  |
| `d2_inc` | text | 8 |  |
| `d3_inc` | text | 8 |  |
| `d1_cost` | text | 8 |  |
| `d2_cost` | text | 8 |  |
| `d3_cost` | text | 8 |  |
| `total_inc` | text | 8 |  |
| `total_cost` | text | 8 |  |
| `chart_head` | text | 156 |  |
| `chart` | chart |  | chart (COLUMN_STACKED): categories + series |
| `contrib1` | text | 8 |  |
| `contrib2` | text | 8 |  |
| `contrib3` | text | 8 |  |
| `contrib4` | text | 8 |  |
| `contrib5` | text | 8 |  |
| `contrib6` | text | 8 |  |
| `growth_label` | text | 8 |  |
| `footnote` | text | 192 | optional |
| `source` | text | 124 |  |

## chart_takeaways

A native column chart with a description on top and key takeaways on the right.

When: Analysis or appendix, any time a number series makes the point.  
From: PowerPoint Template.pptx slide 11

| Slot | Type | Max | Notes |
|---|---|---|---|
| `title` | text | 150 |  |
| `subtitle` | text | 122 | optional |
| `tracker` | text | 43 |  |
| `chart_head` | text | 72 |  |
| `unit` | text | 21 |  |
| `chart` | chart |  | chart (COLUMN_CLUSTERED): categories + series |
| `takeaways_head` | text | 35 |  |
| `takeaways` | text | 378 |  |
| `footnote` | text | 161 | optional |
| `source` | text | 124 |  |

## harvey_table

Comparison table with Harvey balls (0-4): options or competitors against 5 criteria, plus a takeaway bar.

When: Analysis (competitors) or Strategy (option evaluation); also good in the appendix.  
From: PowerPoint Template.pptx slide 72

| Slot | Type | Max | Notes |
|---|---|---|---|
| `title` | text | 150 |  |
| `subtitle` | text | 122 | optional |
| `tracker` | text | 43 |  |
| `table` | table |  | First row: criteria; first column: options. Harvey levels go in h_r<row>_c<col> (0-4). |
| `takeaway` | text | 222 |  |
| `source` | text | 124 |  |
| `h_r<row>_c<col>` | harvey |  | Harvey level 0-4 |
| `h_r<row>_c<col>` | harvey |  | Harvey level 0-4 |
| `h_r<row>_c<col>` | harvey |  | Harvey level 0-4 |
| `h_r<row>_c<col>` | harvey |  | Harvey level 0-4 |
| `h_r<row>_c<col>` | harvey |  | Harvey level 0-4 |
| `h_r<row>_c<col>` | harvey |  | Harvey level 0-4 |
| `h_r<row>_c<col>` | harvey |  | Harvey level 0-4 |
| `h_r<row>_c<col>` | harvey |  | Harvey level 0-4 |
| `h_r<row>_c<col>` | harvey |  | Harvey level 0-4 |

## risk_matrix

Risk matrix: risks A-G placed by probability and impact, coloured by pillar, with a table of descriptions and top-risk stars.

When: Impact section, before mitigation.  
From: Mexico 24H Case.pptx slide 36

Stars: `top_risks: [1, 4]` keeps the star next to risks 1 and 4.

| Slot | Type | Max | Notes |
|---|---|---|---|
| `title` | text | 150 |  |
| `tracker` | text | 43 |  |
| `legend1` | text | 10 |  |
| `legend2` | text | 10 |  |
| `legend3` | text | 10 | optional |
| `r1_name` | text | 34 |  |
| `r2_name` | text | 34 |  |
| `r3_name` | text | 34 |  |
| `r4_name` | text | 34 | optional |
| `r5_name` | text | 34 | optional |
| `r6_name` | text | 34 | optional |
| `r7_name` | text | 34 | optional |
| `r1_desc` | text | 156 |  |
| `r2_desc` | text | 146 |  |
| `r3_desc` | text | 148 |  |
| `r4_desc` | text | 148 | optional |
| `r5_desc` | text | 119 | optional |
| `r6_desc` | text | 138 | optional |
| `r7_desc` | text | 139 | optional |
| `r1_letter` | marker |  |  |
| `r2_letter` | marker |  |  |
| `r3_letter` | marker |  |  |
| `r4_letter` | marker |  | optional |
| `r5_letter` | marker |  | optional |
| `r6_letter` | marker |  | optional |
| `r7_letter` | marker |  | optional |
| `r1_dot` | marker |  |  |
| `r2_dot` | marker |  |  |
| `r3_dot` | marker |  |  |
| `r4_dot` | marker |  | optional |
| `r5_dot` | marker |  | optional |
| `r6_dot` | marker |  | optional |
| `r7_dot` | marker |  | optional |
| `source` | text | 124 |  |
| `r1_star` | marker |  | optional |
| `r2_star` | marker |  | optional |
| `r3_star` | marker |  | optional |
| `r4_star` | marker |  | optional |
| `r5_star` | marker |  | optional |
| `r6_star` | marker |  | optional |
| `r7_star` | marker |  | optional |

## mitigation

Risk mitigation: for each top risk, the mitigation and the actionable steps.

When: Right after the risk matrix.  
From: Mexico 24H Case.pptx slide 37

Repeats: `m1_...` to `m3_...` (2 to 3 groups; fill only the ones you use).

| Slot | Type | Max | Notes |
|---|---|---|---|
| `title` | text | 150 |  |
| `tracker` | text | 43 |  |
| `steps_head` | text | 18 |  |
| `m1_letter` | text | 8 |  |
| `m2_letter` | text | 8 |  |
| `m3_letter` | text | 8 |  |
| `m1_name` | text | 19 |  |
| `m2_name` | text | 27 |  |
| `m3_name` | text | 19 |  |
| `m1_mitigation` | text | 170 |  |
| `m2_mitigation` | text | 158 |  |
| `m3_mitigation` | text | 155 |  |
| `m1_steps` | text | 335 |  |
| `m2_steps` | text | 335 |  |
| `m3_steps` | text | 335 |  |
| `footnote` | text | 161 | optional |
| `source` | text | 124 |  |

## roadmap

Implementation roadmap: 3-5 phases as chevrons with timing, what happens in each, and a rising phase line.

When: Impact section, last content slide before the recap.  
From: Mexico 24H Case.pptx slide 38

Repeats: `ph1_...` to `ph5_...` (3 to 5 groups; fill only the ones you use).

| Slot | Type | Max | Notes |
|---|---|---|---|
| `title` | text | 150 |  |
| `tracker` | text | 43 |  |
| `ph1_name` | text | 66 |  |
| `ph2_name` | text | 66 |  |
| `ph3_name` | text | 66 |  |
| `ph4_name` | text | 66 |  |
| `ph5_name` | text | 66 |  |
| `ph1_time` | text | 11 |  |
| `ph2_time` | text | 11 |  |
| `ph3_time` | text | 12 |  |
| `ph4_time` | text | 13 |  |
| `ph5_time` | text | 13 |  |
| `ph1_text` | text | 408 |  |
| `ph2_text` | text | 384 |  |
| `ph3_text` | text | 336 |  |
| `ph4_text` | text | 312 |  |
| `ph5_text` | text | 255 |  |
| `ph1_label` | text | 8 |  |
| `ph2_label` | text | 8 |  |
| `ph3_label` | text | 8 |  |
| `ph4_label` | text | 8 |  |
| `ph5_label` | text | 8 |  |
| `footnote` | text | 161 | optional |
| `source` | text | 124 |  |

## appendix_divider

Dark divider that opens the appendix.

When: After the closing slide, first slide of the appendix.  
From: Mexico 24H Case.pptx slide 42

| Slot | Type | Max | Notes |
|---|---|---|---|
| `section` | text | 36 |  |
| `client_logo` | image |  | optional; image (client logo) |

## appendix_index

Appendix index. navigate.py fills it with Slide Zooms to every backup slide; without navigation it lists them.

When: Second appendix slide.  
From: Mexico 24H Case.pptx slide 43

| Slot | Type | Max | Notes |
|---|---|---|---|
| `title` | text | 150 |  |
| `tracker` | text | 43 |  |
| `index_list` | text | 1926 | optional |

## text_content

Plain content slide: action title, optional subtitle, bullets. Use when no other pattern fits.

When: Anywhere; appendix text slides.  
From: Mexico 24H Case.pptx slide 43

| Slot | Type | Max | Notes |
|---|---|---|---|
| `title` | text | 150 |  |
| `tracker` | text | 43 |  |
| `body` | text | 1926 |  |
| `source` | text | 140 |  |

## closing

Closing slide: hero photo, team photos and names, client logo.

When: Last slide before the appendix.  
From: Mexico 24H Case.pptx slide 41

Repeats: `m1_...` to `m4_...` (2 to 4 groups; fill only the ones you use).

| Slot | Type | Max | Notes |
|---|---|---|---|
| `hero_image` | image |  | image (closing photo) |
| `client_logo` | image |  | optional; image (client logo) |
| `m1_photo` | photo |  | optional |
| `m2_photo` | photo |  | optional |
| `m3_photo` | photo |  | optional |
| `m4_photo` | photo |  | optional |
| `m1_name` | text | 15 |  |
| `m2_name` | text | 15 |  |
| `m3_name` | text | 17 |  |
| `m4_name` | text | 16 |  |
