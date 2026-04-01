# Case Findings Index

## Purpose

This file is the repo-wide register of client findings, approved rules, user assertions, structural decisions, and pending contradictions.

Before changing business logic or any structural behavior, consult this file and check whether the new observation conflicts with an earlier case, narrows it, or supersedes a working assumption.

## Status Meanings

| Status | Meaning |
| --- | --- |
| `validated` | Approved by real case evidence and accepted as a rule or exception |
| `user-assertion` | Explicit user guidance not yet fully reconciled against prior cases |
| `pending-review` | Important open point that must be checked manually before hard-coding |
| `superseded` | Earlier assumption kept only for history; do not treat as current truth |

## Fields

| Field | Meaning |
| --- | --- |
| `Finding ID` | Stable identifier for referencing a case finding |
| `Source Type` | `validated-case`, `user-assertion`, or `working-assumption` |
| `Client / Example` | Comitente, client name, or example folder |
| `Layer` | Parser, postprocess, merge, render, app, export, validation, or workflow |
| `Artifact` | Specific sheet, PDF section, workbook output, app behavior, file, or layout decision |
| `Area` | Price, quantity, ARS/USD routing, rentas/dividendos routing, guardrail, etc. |
| `Condition` | Exact scope where the finding applies |
| `Claim` | What was observed or approved |
| `Evidence` | Concrete workbook, PDF source section, row, or file |
| `Related Rules / Tension` | Rule IDs or findings that may conflict or need manual comparison |

## Structural Scope

Contradiction checks are not limited to normalization rules. They also apply to:

1. Workbook shape, sheet creation, section mapping, and output placement.
2. Header semantics, column meaning, naming, and display layout.
3. Routing between ARS/USD buckets or between different sheets.
4. App/export behavior such as which workbook is downloaded or rendered.
5. Validation expectations, smoke baselines, and review workflow.

## Current Findings

| Finding ID | Status | Source Type | Client / Example | Layer | Artifact | Area | Condition | Claim | Evidence | Related Rules / Tension |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| CAN-ARS-RAW-PRICE-001 | `validated` | `validated-case` | Canullo / 10600 | `merge` | Resultado Ventas ARS price semantics | Price normalization | Visual-origin TP/ON/Letras in ARS result flow with low price already nominal | Keep raw nominal price; do not divide by 100 again in the ARS result flow | [NORMALIZATION_RULES.md](NORMALIZATION_RULES.md) rules `R-002` and `R-003`; Canullo approved workbooks under 10600 | Related to Parma/Sturman trade-price branches; not a quantity rule |
| CAN-RENTAS-SOURCE-001 | `validated` | `validated-case` | Canullo / 10600 | `merge` | Rentas Dividendos ARS/USD placement | Currency routing | Visual rentas/dividendos rows where row text may look USD-like but source sheet is ARS or USD | Source sheet ARS/USD wins for rentas/dividendos routing | [NORMALIZATION_RULES.md](NORMALIZATION_RULES.md) rules `R-007`, `R-008`, `R-015` | Important precedent for source-based routing, but not yet a proven rule for Resultado Ventas |
| APP-DOWNLOAD-VALUES-001 | `validated` | `validated-case` | Multi-case regression work | `app` | Downloaded workbook selection | Output behavior | App download/export path after merge | Default download must point to the final values workbook, not the wrong intermediate workbook | Prior fix validated during this conversation on app/export path and user verification | Structural/output precedent, not a normalization rule |
| PAR-USD-GUARDRAIL-001 | `validated` | `validated-case` | Parma / 11797 | `merge` | Resultado Ventas USD output | USD result guardrail | TP/ON/Letras rows where stock price scale becomes absurd versus trade nominal | Do not publish misleading numeric USD result; emit guardrail/audit instead | [NORMALIZATION_RULES.md](NORMALIZATION_RULES.md) rules `R-013`, `R-014`, `R-016`; Parma recheck workbooks under 11797 | Can coexist with later Sturman subcases as a narrower exception |
| PAR-VISUAL-TRADE-PRICE-001 | `validated` | `validated-case` | Parma / 11797 | `merge` | Visual Boletos/Resultados trade price | Trade price normalization | Visual TP/ON trade rows already in raw trade magnitude | Keep raw trade price in Visual Boletos/Resultados for the proven subcases | [NORMALIZATION_RULES.md](NORMALIZATION_RULES.md) rules `R-010`, `R-011`, `R-012` | Needs comparison against later micro-price and ARS-low-nominal exceptions |
| STU-USD-MICRO-PRICE-001 | `validated` | `validated-case` | Sturman / 11688 | `merge` | Resultado Ventas USD price regime | USD micro-price | Visual TP/Letras rows with effective USD micro-prices below `0.01` | Preserve micro-price regime; do not re-standardize it like normal 0.55/0.60 USD quotes | [NORMALIZATION_RULES.md](NORMALIZATION_RULES.md) rule `R-020`; Sturman `RECHECK11` targets around `M112 ~= 301355` and `M124 ~= 1565217` | Specializes Parma trade-price rules; not an ARS/USD routing rule |
| STU-VISUAL-QUANTITY-RESCUE-001 | `validated` | `validated-case` | Sturman / 11688 | `postprocess` | Visual Boletos quantity repair | Quantity rescue | Visual Boletos quantities truncated by column width with a unique counterpart in Resultado Ventas | Quantity may be rescued from Resultado Ventas by code/species + concertacion + liquidacion + currency bucket + buy/sell side | Sturman `RECHECK15` workbooks and current repo implementation in [pdf_converter/datalab/postprocess.py](pdf_converter/datalab/postprocess.py) | Assumes the matched Resultado Ventas row is in the correct ARS/USD bucket |
| STU-RESULTADO-ROUTING-ARS-001 | `pending-review` | `user-assertion` | Sturman / 11688 | `parser+merge` | Resultado Ventas ARS/USD placement | Resultado Ventas ARS/USD routing | Visual PDF source pages 5-11, especially subsection `ARS`, where operations appear under the central Resultado Ventas title with `ARS` shown at the left | If a movement is located inside the Visual source subsection `ARS`, it may not belong in final Resultado Ventas USD even if later heuristics or routing place it there | User review on 2026-04-01 referencing the Visual source layout and the example around Sturman rows 148-155 | Compare against `CAN-RENTAS-SOURCE-001` as a source-precedence precedent and against `STU-VISUAL-QUANTITY-RESCUE-001`, which fixed quantities assuming those rows still belonged in USD |
| STU-ROWS-148-155-USD-ASSUMPTION-001 | `pending-review` | `working-assumption` | Sturman / 11688 | `merge` | Resultado Ventas USD destination | Resultado Ventas USD assumption | BONO NACION rows 148-155 were corrected for quantity under the assumption that they should remain in Resultado Ventas USD | Quantity fix is technically validated, but routing to USD is not yet accepted as final truth after the user's new ARS-subsection observation | Sturman `RECHECK15` values workbook: row 148 `qty=1176470588`, row 149 `qty=-1176470588`, row 153 `qty=512974026`, row 154 `qty=-512974026` | Directly challenged by `STU-RESULTADO-ROUTING-ARS-001`; manual review required before treating the USD placement as approved |

## Review Protocol

When a new finding arrives:

1. Identify the closest existing findings in the same area, layer, artifact, or behavior class.
2. Decide whether the new claim is a contradiction, a specialization, or only a noisy neighbor.
3. If contradiction or noise exists, surface the earlier case details before changing code.
4. Only move a pending item to `validated` after manual confirmation with concrete evidence.

## Minimal Entry Rule

If the user asks to change anything structural or functional, try to register at least one of these dimensions in the index:

1. Parser or source interpretation.
2. Postprocess or OCR repair semantics.
3. Merge/routing/calculation behavior.
4. Output structure, naming, or download/export behavior.
5. Validation or regression expectation.