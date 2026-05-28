# Padova Javier / Comitente 10017

## Status

- Case status: `under-review`
- Source folder: `ejemplo comitente 10017 error de fecha formato`
- Latest local output folder: `LOCAL_VERIFY_10017_20260527`
- Latest values workbook: `LOCAL_VERIFY_10017_20260527/COMITENTE_10017_RECHECK_Resumen_Impositivo_FIXED_values.xlsx`
- Latest PDF: `LOCAL_VERIFY_10017_20260527/COMITENTE_10017_RECHECK_Resumen_Impositivo_FIXED.pdf`
- Latest validation: `issue_count=0`, `high=0`, `review=0`

## Original Problem

Middle Office reported that the Streamlit web process failed even though the source file dates looked correct. The reproduced error was:

```text
ValueError: could not convert string to float: 'Fecha: 27/05/26'
```

The failure happened during merge when `Posicion Inicial Gallo` was consumed and a metadata string reached a numeric quantity/price path.

## Root Cause

Datalab injected page-level Gallo metadata rows after `TOTAL POSICION AL 01/01/25` and before the next section. Because the parser still had `Posicion Inicial` active, these rows were appended as fake position rows, for example:

```text
Comitente: 10017 PADOVA JAVIER / Fecha: 27/05/26
Desde Fecha: 01/01/25 / Hasta Fecha: 31/05/25 / Hoja: 11
```

Those are report header/footer fields, not species rows.

## Fixes

1. Parser: `md_to_excel` now skips Gallo table rows made only of report metadata prefixes: `Comitente:`, `Fecha:`, `Desde Fecha:`, `Hasta Fecha:`, and `Hoja:`.
2. Merge: for USD fixed-income Gallo positions with a direct USD unit price, `Posicion Inicial Gallo` now uses that price as `PosicionInicialUSD` instead of overriding it with peso-based `PrecioTenenciasIniciales`.
3. Merge: `Resultado Ventas USD` no longer divides sub-2 nominal TP/ON stock prices by FX again; those are already unit USD prices.

## Evidence

- Before the parser fix, `Fecha: 27/05/26` appeared inside the converted `Posicion Inicial` data rows and the web-equivalent generation crashed.
- After the parser fix, the converted Gallo workbook no longer contains those metadata rows in `Posicion Inicial` and full generation completes.
- Before the USD fixed-income fix, GD30 code `81086` produced `RES-USD-RATIO-001` high findings due to stock basis around `0.0919` and then `0.00065` USD.
- After the fix, GD30 stock price is around `0.761` / `0.75985` USD and validation is clean.

## Verification

```powershell
python -m pytest test_usd_fixed_income_position_price.py test_md_to_excel_gallo_position_sections.py -q
python generate_case_outputs.py --visual-xlsx "LOCAL_VERIFY_10017_20260527/COMITENTE_10017_RECHECK_Visual_from_PDF.xlsx" --gallo-xlsx "LOCAL_VERIFY_10017_20260527/COMITENTE_10017_RECHECK_Gallo_from_PDF.xlsx" --precio-xlsx "LOCAL_VERIFY_10017_20260527/COMITENTE_10017_RECHECK_PrecioTenencias_from_PDF.xlsx" --case-prefix "LOCAL_VERIFY_10017_20260527/COMITENTE_10017_RECHECK" --client-number 10017 --client-name "PADOVA JAVIER" --period-start "Enero 1" --period-end "Diciembre 31"
python explain_audit.py "LOCAL_VERIFY_10017_20260527/COMITENTE_10017_RECHECK_Resumen_Impositivo_FIXED_values.xlsx"
```

Observed results:

- Focused tests: `5 passed`
- Generation: `DONE`
- Audit: `Issues: 0 severity={}`
- PDF: 18 pages, contains `PADOVA` and `10017`, no `WARNING`/`ALERTA`/`VALIDACION`/`REVISAR` tokens in extracted text.

## Related Findings

- `PAD10017-GALLO-PAGE-METADATA-001`
- `PAD10017-USD-FIXED-INCOME-POSITION-001`
- Related precedents: `CIC-GALLO-POSITION-DATE-001`, `CIC-GALLO-POSITION-CONTINUATION-001`, `CIC-USD-POSITION-PRICE-001`, `FLO10946-USD-EXTERIOR-AUX-PRICE-001`
