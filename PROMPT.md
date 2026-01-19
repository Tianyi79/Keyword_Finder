You are helping me locate exact evidence passages in an academic PDF.
Do NOT paraphrase the paper. Your job is to produce search anchors I can paste into a local keyword finder.

IMPORTANT: Follow the exact output template below. Do not add extra sections.  
Formatting rules (VERY IMPORTANT):
- Separate keywords with commas.
- Each line must be the exact keyword/phrase itself, nothing else.

Generate search anchors I can paste into a local keyword finder (one per line).
STRICT RULES:
- NO bullets, NO numbering, NO prefixes (no "-" "•" "1."), keywords in comma-seperated lists.
- Prefer 2–6 word phrases and proper nouns over single common words.
- Do NOT output synonyms that would hit the same sentence/paragraph repeatedly.
- Do NOT output both a short term and its longer container phrase (pick the most specific one).
- Avoid generic evidence phrases that appear many times (e.g., “we propose”, “we show that”). Only include them if you also add a SPECIFIER (e.g., “we propose <method name>”).
- Cap the output to:
  [TECH ANCHORS] 18–28 lines total
  [EVIDENCE ANCHORS] 6–10 lines total
- Diversity constraint: anchors should point to DIFFERENT parts of the paper (method definition, key table/figure, main result claim, limitation paragraph, conclusion).

Correct:
we propose a
Table 2
contrastive loss

Wrong:
- we propose a
• Table 2
1) contrastive loss
keyword: contrastive loss

==================== OUTPUT TEMPLATE (MUST FOLLOW) ====================

[ONE-LINE TOPIC]
<one sentence: what this paper is about>

[SECTION SUMMARY]
Abstract: <1–2 sentences>
Introduction: <1–2 sentences>
Method: <1–2 sentences>
Experiments/Results: <1–2 sentences>
Limitations: <1–2 sentences or "Not stated clearly">
Conclusion/Future work: <1–2 sentences>

[KEYWORDS - TECHNICAL ANCHORS]
- Output 30–60 lines
- ONE keyword/phrase per line
- Prefer 2–5 word phrases over single words
- Prefer: method/model names, module/component names, unique coined terms, symbols/variables, loss names, dataset/benchmark names, metrics, table/figure anchors like "Table 2", "Figure 3", and distinctive nouns from captions
- Avoid generic words like: method, result, experiment, approach, paper, model, data, performance, analysis

<put keywords here, in a comma-seperated list>

[KEYWORDS - EVIDENCE ANCHORS]
- Output 15–25 lines
- ONE phrase per line
- These are “where the author makes a claim / compares / admits limits / concludes” anchors
- Prefer phrases that are likely to appear verbatim in the paper’s writing style
- Include some of these patterns if appropriate:
  "we propose", "we show that", "compared to", "in contrast", "our ablation", "sensitivity", "limitation", "trade-off", "failure case", "in conclusion", "future work", "we hypothesize"

<put keywords here, in a comma-seperated list>

[WARNINGS - HIGH FALSE POSITIVE TERMS]
- Output 5–15 items
- Terms I should avoid using alone because they are too generic OR cause substring false matches (e.g., "art" -> "departure")
- Format: term -> safer alternative phrase(s)

<put warnings here>

==================== END TEMPLATE ====================

Now process the paper text I provide and fill the template.


 