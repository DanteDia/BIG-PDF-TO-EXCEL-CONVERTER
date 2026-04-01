---
description: "Use when reviewing any client finding or requested change that could alter system behavior, structure, routing, parsing, calculations, output layout, naming, validation, or fiscal business logic. Requires checking CASE_FINDINGS_INDEX.md first and surfacing earlier contradictory or noisy case evidence before proceeding."
name: "Case Contradiction Check"
---
# Case Contradiction Check

Before proposing or implementing any meaningful change for this repo:

1. Classify the new observation by client, case/example, affected layer, affected artifact, and rule family.
2. Read [CASE_FINDINGS_INDEX.md](../../CASE_FINDINGS_INDEX.md) first.
3. Search for prior findings in the same rule family or with overlapping conditions.
4. If an older finding conflicts, or even only makes noise, surface it explicitly before acting.

This applies to all structural or behavioral areas, including:

1. Parsing, section detection, header interpretation, and workbook structure.
2. Price, quantity, currency, ARS/USD routing, and formula semantics.
3. Sheet placement, summary routing, column meaning, and naming decisions.
4. Merge behavior, carryover logic, stock logic, and guardrails.
5. Output behavior in PDF, client Excel, downloads, app behavior, and validation rules.

When surfacing a possible contradiction, always include:

1. Prior client or example name.
2. The exact earlier claim or assumption.
3. The evidence file or concrete data point.
4. Whether the new observation looks like a true contradiction, a narrower exception, or still needs manual review.

Decision protocol:

1. If the earlier and newer findings are compatible, label the newer one as a specialization.
2. If they appear incompatible, stop and ask for manual review instead of silently choosing one.
3. If the earlier record was only a working assumption and not an approved rule, mark it as superseded or pending review.
4. If the requested change affects structure or outputs beyond the immediate bug, check for prior user guidance on layout, naming, placement, or workflow before editing code.

Record-keeping protocol:

1. Add new validated findings to [CASE_FINDINGS_INDEX.md](../../CASE_FINDINGS_INDEX.md).
2. Distinguish `validated`, `user-assertion`, `pending-review`, and `superseded` entries.
3. For every new entry, capture client/comitente, layer, artifact, exact condition, and related rules or noisy antecedents.

Minimum contradiction check before acting:

1. What exact behavior or structure is being changed?
2. In which layer does it live: parser, postprocess, merge, render, app, export, or validation?
3. Did the user previously ask for the opposite or for a narrower rule in another case?
4. Is the new request a contradiction, a specialization, or a case-specific override?

Do not treat previous implementation work as permanent truth if a new user finding challenges its routing or interpretation. In that case, mark the prior assumption for review and present the conflict.