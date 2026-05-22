# Flores 13645 Handoff

## Status

- Case registry status: `under-review`
- Finding: `FLO-GALLO-INTERMEDIATE-POSITION-001`
- Classification: `parser+merge`, dated prior position fallback for Resultado Ventas ARS

## Evidence

- Gallo source: [ejemplo flores 13645/IG_13645.PDF](../ejemplo%20flores%2013645/IG_13645.PDF) page 8 has `POSICION AL 31/05/25` with `GLOB`, quantity `880`, price `6560`, amount `5,772,800`.
- Visual source: [ejemplo flores 13645/13645bolsa.pdf](../ejemplo%20flores%2013645/13645bolsa.pdf) page 1 has a later 30/06/2025 sale of code `8499` for quantity `880` at `6100.6932`.
- Existing generated PDF: [ejemplo flores 13645/13645_FLORES_OSCAR_ANIBAL_Resumen_Impositivo_20260521_1754.pdf](../ejemplo%20flores%2013645/13645_FLORES_OSCAR_ANIBAL_Resumen_Impositivo_20260521_1754.pdf) shows GLOBANT with a result around `-6.86M`, caused by generic fallback price `13900` instead of the dated 31/05 snapshot price.

## Rule

Use a later-than-01/01 Gallo position snapshot as prior stock/cost evidence only when all of these are true:

1. True `Posicion Inicial Gallo` and `PrecioTenenciasIniciales` do not provide a basis for the instrument.
2. The Gallo snapshot has an explicit `POSICION AL dd/mm/yy` date preserved by the parser.
3. The snapshot date is earlier than the sale/operation date.
4. The row is ARS equity-like (`Cedears` or `Acciones`).
5. The basis uses the snapshot's own unit price (`precio Tenencia Final Pesos`), not generic `PreciosInicialesEspecies`.

## Contradiction Check

This specializes `CIC-GALLO-POSITION-DATE-001`. Cicero proved that later position snapshots must not replace the true initial position. Flores adds the narrower case where a dated later snapshot can seed a later operation only after stronger initial/tenencias evidence is absent.

## Verification

Focused regression tests:

- `test_flores_intermediate_position.py`
- `test_md_to_excel_gallo_position_sections.py`

Generated local outputs are not promoted yet. Regenerate and review Flores values/PDF before moving the case out of `under-review`.
