# Ramsfelder Marcela / Comitente 13988

## Status

- Case status: `under-review`
- Source folder: `13988 falta tenencias`
- Inputs available: Visual/Bolsa PDF and Gallo PDF
- Missing input: Precio Tenencias PDF
- Latest local output folder: `LOCAL_VERIFY_13988_20260529`
- Latest values workbook: `LOCAL_VERIFY_13988_20260529/RAMSFELDER_13988_RECHECK_Resumen_Impositivo_FIXED_values.xlsx`
- Latest PDF: `LOCAL_VERIFY_13988_20260529/RAMSFELDER_13988_RECHECK_Resumen_Impositivo_FIXED.pdf`
- Latest validation: `issue_count=0`, `high=0`, `review=0`

## Original Problem

The case has no Precio Tenencias input. This does not affect the fiscal result because the available Visual and Gallo inputs have no sale rows, but the deterministic generation script still required `--precio-pdf` or `--precio-xlsx` and blocked generation.

## Fix

`generate_case_outputs.py` now treats Precio Tenencias as optional, matching `GalloVisualMerger` and the Streamlit merge path, which already accept `precio_tenencias_path=None`.

The script also creates nested output directories before saving converted intermediates and final outputs, so case prefixes like `LOCAL_VERIFY_13988_20260529/RAMSFELDER_13988_RECHECK` work from a fresh folder.

## Evidence

- Visual conversion found `Rentas Dividendos USD: 5 rows`, `Boletos: 0 rows`, `Resultado Ventas ARS: 0 rows`, and `Resultado Ventas USD: 0 rows`.
- Gallo conversion found `Renta Fija Dolares: 4 rows`; one income row flowed to final rentas.
- Final `Rentas Dividendos USD` has 6 rows total:
  - Gallo `58394` RENTA `58.32`
  - Visual `58394` RENTA `2595.55`
  - Visual `58394` AMORTIZACION `0`
  - Visual `58394` RENTA `2624.08`
  - Visual `58535` RENTA `1161.27`
  - Visual `58535` AMORTIZACION `0`
- Final `Resumen` USD `Rentas` and `Total`: `6439.22`.
- `Boletos`, `Resultado Ventas ARS`, and `Resultado Ventas USD` have zero data rows.

## Verification

```powershell
python -m pytest test_generate_case_outputs_optional_precio.py -q
python generate_case_outputs.py --visual-xlsx "LOCAL_VERIFY_13988_20260529/COMITENTE_13988_RECHECK_Visual_from_PDF.xlsx" --gallo-xlsx "LOCAL_VERIFY_13988_20260529/COMITENTE_13988_RECHECK_Gallo_from_PDF.xlsx" --case-prefix "LOCAL_VERIFY_13988_20260529/RAMSFELDER_13988_RECHECK" --client-number 13988 --client-name "RAMSFELDER, MARCELA" --period-start "Enero 1" --period-end "Diciembre 31"
python explain_audit.py "LOCAL_VERIFY_13988_20260529/RAMSFELDER_13988_RECHECK_Resumen_Impositivo_FIXED_values.xlsx"
```

Observed results:

- Optional Precio Tenencias regression: `1 passed`
- Generation: `DONE`
- Audit: `Issues: 0 severity={}`
- PDF: 11 pages, contains `13988`, `RAMSFELDER`, and `MARCELA`, with no `WARNING`/`ALERTA`/`VALIDACION`/`REVISAR` tokens in extracted text.

## Related Findings

- `RMS13988-MISSING-PRECIO-NO-SALES-001`
- Related precedent: `VIS-ONLY-MERGE-001`
