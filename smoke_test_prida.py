#!/usr/bin/env python
"""
Smoke test for Prida 10488 case.

Runs the full merge pipeline from frozen Gallo+Visual inputs and compares
EVERY cell in EVERY sheet against the approved baseline values workbook.

Usage:
    python smoke_test_prida.py --save       # save current output as new baseline
    python smoke_test_prida.py              # run smoke test against baseline
"""

import sys
from pathlib import Path

from smoke_test_common import SmokeTestConfig, run_cli


ROOT = Path(__file__).resolve().parent
PRIDA_DIR = ROOT / "Prida- Transferencia titulos tomar como compra"
BASELINE_DIR = ROOT / "SMOKE_BASELINE" / "PRIDA_20260416_APPROVED"

GALLO_PATH = PRIDA_DIR / "IG_10488.xlsx"
VISUAL_PATH = PRIDA_DIR / "10488bolsa.xlsx"

BASELINE_VALUES = BASELINE_DIR / "10488_PRIDA_baseline_values.xlsx"
BASELINE_JSON = BASELINE_DIR / "baseline_snapshot.json"


def run_pipeline():
    sys.path.insert(0, str(ROOT))
    from pdf_converter.datalab.merge_gallo_visual import GalloVisualMerger

    assert GALLO_PATH.exists(), f"Gallo not found: {GALLO_PATH}"
    assert VISUAL_PATH.exists(), f"Visual not found: {VISUAL_PATH}"

    merger = GalloVisualMerger(str(GALLO_PATH), str(VISUAL_PATH))
    _, workbook = merger.merge("both")
    return workbook


def build_summary_lines(workbook):
    lines = []

    resumen_sheet = workbook["Resumen"]
    lines.extend(["", "Resumen:"])
    for row in range(1, resumen_sheet.max_row + 1):
        values = [resumen_sheet.cell(row, col).value for col in range(1, resumen_sheet.max_column + 1)]
        lines.append(f"  Row {row}: {values}")

    boletos_sheet = workbook["Boletos"]
    headers = [boletos_sheet.cell(1, col).value for col in range(1, boletos_sheet.max_column + 1)]
    columns = {header: index + 1 for index, header in enumerate(headers)}
    lines.extend(["", f"Boletos: {boletos_sheet.max_row - 1} data rows", "Key boletos:"])
    for row in range(2, boletos_sheet.max_row + 1):
        boleto = str(boletos_sheet.cell(row, 4).value or "")
        if boleto in {"42583", "42598", "42580", "42582"}:
            precio_nominal = boletos_sheet.cell(row, columns.get("Precio Nominal", 20)).value
            moneda = boletos_sheet.cell(row, columns.get("Moneda", 5)).value
            bruto = boletos_sheet.cell(row, columns.get("Bruto", 13)).value
            lines.append(f"  B={boleto}: PN={precio_nominal}, Moneda={moneda}, Bruto={bruto}")

    lines.extend(
        [
            "",
            "Protected invariants:",
            "- B=42583  Moneda=Dolar MEP  PN=1.309     Bruto=3998.995  (VENTA USD override)",
            "- B=42598  Moneda=Dolar MEP  PN=1.13269   Bruto=-3460.368 (VENTA USD)",
            "- B=42580  Moneda=Pesos      PN=1.28608   Bruto=4998.993  (normal Pesos)",
            "- Synthetic initial positions for 81086/81090/81092/81274 from TRF history",
            "- 81090 Resultado Ventas USD: gains (not losses) from TRF cost basis",
            "- Full workbook values compared cell-by-cell by smoke test.",
        ]
    )
    return lines


CONFIG = SmokeTestConfig(
    title="PRIDA 10488",
    description="Smoke test - Prida 10488",
    baseline_dir=BASELINE_DIR,
    baseline_values=BASELINE_VALUES,
    baseline_json=BASELINE_JSON,
    input_labels=(("Gallo", GALLO_PATH), ("Visual", VISUAL_PATH)),
    run_pipeline=run_pipeline,
    summary_builder=build_summary_lines,
    baseline_load_data_only=True,
)


if __name__ == "__main__":
    raise SystemExit(run_cli(CONFIG))
