# Course judging rubric (used by the review's judge view)

Adapted from published competition rubrics: CBS Case Competition (analysis 25, solution 25, feasibility 20, creativity 20, visuals 10) and NIBS (issue identification, analysis, alternatives and recommendation, implementation, delivery, creativity, Q&A). The same rubric is used by human judges in class, so AI and human scores can be compared.

| Criterion | Weight | 1 (weak) | 3 (solid) | 5 (winning) |
|---|---|---|---|---|
| Key issue | 10 | Restates the case | Names the central question | Reframes the question in a way the client had not seen, and the whole deck answers it |
| Analysis | 20 | Facts without a "so what" | Company and market facts with sources lead to conclusions | Few, sharp numbers that make the recommendation feel inevitable |
| Recommendation and alternatives | 20 | One idea, no options considered | Clear pillars; some options compared | Explicit criteria, options dropped with reasons, recommendation follows from them |
| Implementation and financials | 20 | No numbers, no plan | Roadmap and a revenue figure | Initiatives linked to revenue drivers, ROI and payback, phased roadmap, risks with mitigation |
| Creativity | 10 | Generic ideas that fit any company | Some ideas specific to the client | Branded, client-specific initiatives a judge will remember |
| Slides and storyline | 10 | Topic titles, crowded slides | Action titles, mostly one message per slide | Titles alone tell the story; clean, consistent, sourced |
| Q&A readiness | 10 | No appendix | Some backup slides | An appendix answers the likely questions and is one click away |

Score = sum of (score / 5 x weight), out of 100.

## How Claude uses it in the judge view

1. For each criterion, quote one piece of evidence from the deck (slide number and words) **before** giving the score.
2. Score twice, independently. Report any criterion where the two scores differ by more than 1 and say why.
3. List the 10-20 most likely judge questions, tagged: assumption, why-not-X, implementation, financials, risk, curveball.
4. List the appendix gaps: likely questions with no backup slide. For each, propose a backup slide (pattern and action title) the team can add with edit mode.
5. Say clearly that this is feedback. Research on AI judges shows they favour longer text and their own model's writing, and squeeze scores toward the middle. Human judges differ, and the team decides.
