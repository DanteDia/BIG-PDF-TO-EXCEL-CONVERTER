# Flores Gabriel 10946 Handoff

## Status

- Case registry status: `under-review`
- Finding: `FLO10946-USD-EXTERIOR-AUX-PRICE-001`
- Classification: `merge`, USD exterior action cost-basis fallback

## Evidence

- Broken output: [ejemplo 10946/10946_FLORES_GABRIEL_Resumen_Impositivo_20260522_1813.xlsx](../ejemplo%2010946/10946_FLORES_GABRIEL_Resumen_Impositivo_20260522_1813.xlsx) had six high `RES-USD-RATIO-001` triggers on `Resultado Ventas USD` rows 43, 45, 46, 47, 48, and 49.
- Source Gallo: [ejemplo 10946/IG_10946.PDF](../ejemplo%2010946/IG_10946.PDF) shows `07877 APPLE COMPUTER INC` foreign-action sales around `211` to `242` USD.
- Auxiliary initial prices: `PreciosInicialesEspecies` row for code `7877` has AAPL-US prices `250.44` / `243.78`, already USD unit prices.
- Broken generated rows used `0.208750426` as stock price, because the USD fallback divided `243.78` by `COTIZACION_INICIO_PERIODO`.

## Rule

For `Acciones` with `moneda_emision = Dolar Cable (exterior)`, a `PreciosInicialesEspecies` fallback is already a USD unit price. Do not divide it by the initial FX when resolving stock basis for `Resultado Ventas USD`. Keep the FX division for non-exterior USD fallbacks.

## Contradiction Check

This specializes `CIC-USD-POSITION-PRICE-001`. Cicero proved direct Gallo position prices for exterior actions are already USD; Flores Gabriel 10946 extends the same unit rule to `PreciosInicialesEspecies` fallback when no stronger initial/tenencias basis exists.

## Verification

- `python -m pytest test_usd_exterior_initial_price_fallback.py test_flores_intermediate_position.py test_md_to_excel_gallo_position_sections.py test_precio_tenencias_zero_cost.py test_economic_sanity_validation.py test_release_audit.py -q` -> 14 passed.
- Regenerated E2E output: [LOCAL_VERIFY_10946_20260522/FLORES_GABRIEL_10946_AFTERFIX_Resumen_Impositivo_FIXED_values.xlsx](../LOCAL_VERIFY_10946_20260522/FLORES_GABRIEL_10946_AFTERFIX_Resumen_Impositivo_FIXED_values.xlsx).
- Validation JSON: `issue_count=0`.
- `explain_audit.py` on regenerated values: `Issues: 0 severity={}`.
- AAPL rows now use stock price `243.78`, with normal losses and no audit alert.
