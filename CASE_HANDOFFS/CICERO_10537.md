# Cicero 10537 Handoff

Last updated: 2026-05-21

## Status

- Current status: `under-review`
- Latest pushed commit with fixes: `6aed231` (`fix: stabilize Cicero recovered-cost review flow`)
- Latest local audit: `9 review`, `0 high`
- Latest rule family: Gallo position parsing, Precio Tenencias recovered cost, ARS economic sanity filtering.
- This case is not yet an approved smoke. It is safe to test through the app, but remaining review rows still need human review before treating it as release-clean.

## Inputs And Local Outputs

- Source OCR markdown folder: `Caso Cicero/WEBFLOW_E2E_20260519_221839/`
- Latest regenerated local folder: `LOCAL_VERIFY_CURRENT_CICERO_AFTERFIX/`
- Latest values workbook: `LOCAL_VERIFY_CURRENT_CICERO_AFTERFIX/CICERO_CURRENT_AFTERFIX_values.xlsx`
- Latest formulas workbook: `LOCAL_VERIFY_CURRENT_CICERO_AFTERFIX/CICERO_CURRENT_AFTERFIX_formulas.xlsx`
- Latest PDF: `LOCAL_VERIFY_CURRENT_CICERO_AFTERFIX/CICERO_CURRENT_AFTERFIX_final.pdf`
- Latest audit JSON: `LOCAL_RELEASE_AUDIT_CICERO_AFTERFIX.json`

These local outputs are investigation artifacts. Do not commit them unless the user explicitly asks to promote them.

## Rules Added Or Confirmed

- `CIC-GALLO-POSITION-DATE-001`: inline `POSICION AL 01/01/25` must be classified as `Posicion Inicial`; later dated `POSICION AL` blocks are `Posicion Final`.
- `CIC-GALLO-POSITION-CONTINUATION-001`: Gallo position continuation pages under category headings stay inside the active position section.
- `CIC-PRECIO-TENENCIAS-NEGATIVE-INVESTED-001`: positive quantity with negative `Importe invertido` in Precio Tenencias means recovered cost; remaining units have zero fiscal basis.
- `CIC-USD-POSITION-PRICE-001`: exterior actions with direct Gallo initial USD unit price must keep that price and must not be divided by FX again.
- `CIC-USD-GUARDRAIL-COST-001`: USD result sanity ratio uses cost denominator when cost exists.

See [CASE_FINDINGS_INDEX.md](../CASE_FINDINGS_INDEX.md) for evidence and compatibility notes.

## What Was Fixed

- Gallo parser now recognizes inline `POSICION AL` rows and routes by date.
- Gallo parser now keeps `TIT.PRIVADOS DEL EXTERIOR` continuation rows inside initial/final position blocks.
- Precio Tenencias postprocess preserves negative `Importe invertido` for positive quantities instead of flipping it with `abs()`.
- Merge marks those rows as `PrecioTenenciasCostoRecuperado`, assigns zero initial basis, and computes sale result as full proceeds when appropriate.
- Economic sanity validator no longer reports ARS ratio-one/recovered-cost rows as review triggers when the initial position is explicitly marked recovered-cost.
- The generated PDF was produced from the materialized values workbook and verified as a valid PDF with 59 pages.

## Verification

Focused tests:

```powershell
python -m pytest test_precio_tenencias_zero_cost.py test_economic_sanity_validation.py test_md_to_excel_gallo_position_sections.py test_release_audit.py
```

Latest result before push: `10 passed`.

Latest local audit command:

```powershell
python run_release_audit.py LOCAL_VERIFY_CURRENT_CICERO_AFTERFIX\CICERO_CURRENT_AFTERFIX_values.xlsx --json-output LOCAL_RELEASE_AUDIT_CICERO_AFTERFIX.json
```

Latest Cicero target result: `9 review`, `0 high`, all `RES-ARS-RATIO-001`.

## Remaining Review Rows

Latest remaining review rows in `Resultado Ventas ARS`:

| Rows | Code | Instrument | Reason |
| --- | --- | --- | --- |
| 33-34 | `30035` | SUPV / Grupo Supervielle | Result ratio around 81-82% of bruto; not recovered-cost noise. |
| 326-332 | `8481` | AMD CEDEAR | Result ratio around 81-87% of bruto; latest Precio Tenencias input is positive invested amount, not recovered cost. |

## Commands For Next Session

Check pushed state:

```powershell
git rev-parse HEAD
git rev-parse origin/main
git log -1 --oneline
```

Regenerate the PDF from the current values workbook if needed:

```powershell
$env:PYTHONIOENCODING='utf-8'; @'
from pathlib import Path
from pdf_converter.datalab.excel_to_pdf import ExcelToPdfExporter

values_path = Path(r'LOCAL_VERIFY_CURRENT_CICERO_AFTERFIX\CICERO_CURRENT_AFTERFIX_values.xlsx')
pdf_path = Path(r'LOCAL_VERIFY_CURRENT_CICERO_AFTERFIX\CICERO_CURRENT_AFTERFIX_final.pdf')
exporter = ExcelToPdfExporter(str(values_path), {'numero': '10537', 'nombre': 'CICERO MIGUEL'})
exporter.periodo_inicio = 'Enero 1'
exporter.periodo_fin = 'Diciembre 31'
exporter.anio = 2025
exporter.export_to_pdf(str(pdf_path))
print(pdf_path.resolve())
'@ | .\.venv\Scripts\python.exe -
```

## Cautions

- Do not infer that every high `Resultado/Bruto` ratio is harmless. The recovered-cost exception requires explicit `PrecioTenenciasCostoRecuperado` source in `Posicion Inicial Gallo`.
- Do not promote Cicero to approved smoke until SUPV/AMD remaining reviews are resolved or explicitly accepted.
- Do not rollback to the May 6 anchor for Cicero; prior comparison showed it was worse on this case.