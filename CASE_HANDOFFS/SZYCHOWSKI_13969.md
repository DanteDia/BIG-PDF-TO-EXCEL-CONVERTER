# Szychowski 13969 Handoff

## Status

- Case registry status: `under-review`
- Finding: `SZY-VISUAL-RENTAS-ONLY-001`
- Classification: `parser+merge`, Visual/Bolsa rentas-only report

## Evidence

- Source Bolsa: [13969 szychowski ines - faltan las rentas de bolsa/13969bolsa.datalab.md](../13969%20szychowski%20ines%20-%20faltan%20las%20rentas%20de%20bolsa/13969bolsa.datalab.md) contains only `Rentas y Dividendos`, `Resumen`, and `Posición de Títulos` sections.
- Source Bolsa USD rentas on 10/07/2025:
  - AL30 / code `5921`, renta `25.77`, gastos `0.04`.
  - GD30 / code `81086`, renta `3.46`.
  - Visual source total: `29.23`.
- Broken final PDF: [13969 szychowski ines - faltan las rentas de bolsa/13969_SZYCHOWSKI_MARIA_INES_Resumen_Impositivo_20260526_1531.pdf](../13969%20szychowski%20ines%20-%20faltan%20las%20rentas%20de%20bolsa/13969_SZYCHOWSKI_MARIA_INES_Resumen_Impositivo_20260526_1531.pdf) showed only Gallo January rentas (`27.65`, `3.32`) and omitted Bolsa July rentas.

## Rule

Visual reports must be detected as Visual even when they contain only rentas/resumen/position sections and no boletos/resultados. For Visual-origin rentas/dividendos, preserve the source `Importe` as the final importe; the source `Gastos` stays visible but must not be subtracted a second time.

## Contradiction Check

This specializes `CAN-RENTAS-SOURCE-001`. Source sheet routing still wins for rentas/dividendos; this case adds the missing parser entry point and clarifies Visual `Importe` semantics.

## Verification

- `python -m pytest test_md_to_excel_visual_rentas_only.py test_visual_rentas_importe_preserved.py test_md_to_excel_gallo_position_sections.py test_precio_tenencias_zero_cost.py test_economic_sanity_validation.py test_release_audit.py -q` -> 12 passed.
- Regenerated output: [LOCAL_VERIFY_13969_20260526/SZYCHOWSKI_13969_AFTERFIX_Resumen_Impositivo_FIXED_values.xlsx](../LOCAL_VERIFY_13969_20260526/SZYCHOWSKI_13969_AFTERFIX_Resumen_Impositivo_FIXED_values.xlsx).
- Validation JSON: `issue_count=0`.
- `explain_audit.py` on regenerated values: `Issues: 0 severity={}`.
- Final USD rentas total: `60.20` = Gallo `27.65 + 3.32` plus Bolsa `25.77 + 3.46`.
- Regenerated PDF has no `WARNING`, `ALERTA`, `VALIDACION`, `VALIDACIÓN`, or `REVISAR` tokens.
