# Normalization Rules

## Purpose

This file is the starting table for how prices, quantities, and effective currency
should be interpreted before calculating stock, cost, and summary totals.

It must separate:

1. Rules validated by approved client cases.
2. Rules implemented as heuristics in code.
3. Rules still unknown and pending evidence.

The first approved seed comes from Canullo.

## Columns

| Rule ID | Field | Instrument match | Origin / source | Currency context | Current treatment | Confidence | Evidence |
| --- | --- | --- | --- | --- | --- | --- | --- |
| R-001 | Price | Obligaciones Negociables, Titulos Publicos, Letras | Generic | Any | Treat source price as quoted per 100 nominal units and divide by 100 to get nominal unit price. Quantity stays unchanged. | High | Current merge rule in `TIPOS_PRECIO_CADA_100` and `_normalize_nominal_price()` |
| R-002 | Price | Obligaciones Negociables, Titulos Publicos, Letras | Visual transaction rows | ARS result flow | If row is Visual-origin, ARS flow, and absolute price is low (`0 < abs(price) < 20`), treat it as already nominal and do not divide by 100 again. Quantity stays unchanged. | High | Validated by Canullo and encoded in `_uses_visual_ars_raw_nominal()` |
| R-003 | Initial cost price | Any instrument that would normally match `precio cada 100` | `PrecioTenenciasIniciales` | Initial stock / cost basis | Treat incoming value as already nominal. Do not divide by 100 again. Quantity stays unchanged. | High | Validated by Canullo and encoded in `_normalize_initial_cost_price()` |
| R-004 | Initial cost price | Obligaciones Negociables, Titulos Publicos, Letras | `PreciosInicialesEspecies` or generic initial-cost fallback | Initial stock / cost basis | Apply the same nominal-price normalization as normal transactions. Today this means divide by 100 unless a later validated exception says otherwise. Quantity stays unchanged. | Medium | Current code path through `_normalize_initial_cost_price()` -> `_normalize_nominal_price()` |
| R-005 | Interest and expenses | Obligaciones Negociables, Titulos Publicos, Letras | Resultado Ventas ARS materialization | ARS result flow | Divide `interes` and `gastos` by 100 when the instrument is quoted per 100. | Medium | Current implementation in `_materialize_resultado_ventas()` |
| R-006 | Standardized USD price | Obligaciones Negociables, Titulos Publicos, Letras | Resultado Ventas USD materialization | USD result flow | First convert standardized price to USD using exchange rate, then divide by 100 to get nominal USD price. Quantity stays unchanged. | Medium | Current implementation in `_materialize_resultado_ventas()` |
| R-007 | Effective currency for rentas/dividendos | Any | Visual `Rentas Dividendos ARS` sheet | Summary routing | Source sheet wins over row text. If a row comes from Visual ARS, route it as ARS and normalize effective currency to `Pesos` for downstream summary. | High | Validated by Canullo Tenaris and encoded in `_normalize_visual_rentas_currency()` |
| R-008 | Effective currency for rentas/dividendos | Any | Visual `Rentas Dividendos USD` sheet | Summary routing | Source sheet wins over row text. If a row comes from Visual USD, keep it in USD bucket; visible currency may still be `Dolar`, `Dolar MEP`, or `Dolar Cable`. | Medium | Current implementation mirrors ARS rule; approved negative evidence from Canullo |
| R-009 | Quantity | All current Canullo-confirmed paths | Any | Any | No quantity multiplier/divider validated from Canullo yet. Current safe rule is to preserve parsed quantity unless a separate OCR correction rule is explicitly proven. | High | Canullo fix was price/currency based, not quantity-based |
| R-010 | Trade price | Títulos Públicos, Obligaciones Negociables, Letras | Visual boletos / trade rows | ARS trade flow | If Visual-origin and price is already in trade magnitude for pesos (roughly `>= 100`), keep raw trade price; do not divide by 100 in Boletos or Resultados. Quantity stays unchanged. | High | Validated by Parma (`9237`, `9247`, `5923`, `80567`) and encoded in `_uses_visual_raw_trade_price()` |
| R-011 | Trade price | Títulos Públicos, Obligaciones Negociables, Letras | Visual boletos / trade rows | USD trade flow | If Visual-origin and dollar trade price is already in trade magnitude (roughly `0 < abs(price) < 2`), keep raw trade price; do not divide by 100 in Boletos or Resultados. Quantity stays unchanged. | High | Validated by Parma (`5921`, `9237`) and encoded in `_uses_visual_raw_trade_price()` |
| R-012 | Trade price | Cedears | Visual boletos / trade rows | Any | Cedears currently remain raw with no `/100` adjustment. Parma confirms this branch is already correct and should stay isolated from TP/ON rules. | High | Validated by Parma (`8526`) |
| R-013 | Initial stock price | Títulos Públicos, Obligaciones Negociables, Letras | USD results fallback from `PrecioTenenciasIniciales` | USD result flow | `PrecioTenenciasIniciales` fallback already comes in USD nominal for these rows and must not be divided again by `Valor USD Dia`. | High | Validated by Parma (`5923`, `80567`, `9247`) |
| R-014 | Guardrail | Títulos Públicos, Obligaciones Negociables, Letras | USD result rows | USD result flow | If stock price is absurd versus nominal trade price (ratio extremely small/large), do not emit a numeric result; place `|` and add audit alert. | High | Validated by Parma emergency protocol |
| R-015 | Effective currency for rentas/dividendos | Any | Visual rentas origin | Summary routing | Visual source sheet (`ARS` / `USD`) must win over row text when splitting final rentas/dividendos sheets. | High | Canullo + Parma |
| R-016 | Economic unit price | Títulos Públicos, Obligaciones Negociables, Letras | Resultado Ventas USD | USD result flow | If the normalized USD nominal is still around screen price each 100 (roughly `>= 10`), divide by 100 again to compare sale and stock on economic unit basis. | High | Validated by Parma (`80567`, `9247`) |
| R-017 | Effective currency for rentas/dividendos | Gallo-only rentas rows | Gallo origin with USD-signaled species name / emission | Summary routing | If a renta only comes from Gallo and the species/emission clearly signals USD/U$D, classify it as USD even if Gallo sheet is pesos. | High | Validated by Parma business rule for BONO PCIA BS AS |

## What Canullo proved

Canullo gave two approved facts that must stay frozen unless a later approved case disproves them:

1. `PrecioTenenciasIniciales` is not the same semantic source as `PreciosInicialesEspecies`.
2. A Visual ARS dividend row may contain USD-related text in `Moneda`, but still belongs to ARS if the source sheet is ARS.

## Open items

These items should not be converted into hard rules yet:

1. Whether all Visual ARS low-magnitude prices for `precio cada 100` instruments are always already nominal, or whether more sub-cases exist by instrument family.
2. Whether any quantity scaling rule exists by broker, section, or OCR layout beyond the separate boletos OCR correction heuristics.
3. Whether `PreciosInicialesEspecies` needs its own validated exception matrix by instrument family.
4. Some guarded TP/ON USD cases may still need finer stock-source repair beyond the emergency protocol, but they no longer publish misleading numeric results.

## Recommended next evolution

When a new approved case arrives, update this table with:

1. Exact input source.
2. Exact affected field.
3. Exact transformation applied.
4. Confidence level.
5. Approved case name.
6. Code path where the rule currently lives.

Once the table is stable enough, move the validated subset into a machine-readable rules file consumed by the merge layer.