#!/usr/bin/env python
"""
Smoke test for Sigal 10374 case.

Runs the full merge pipeline from frozen Gallo+Visual inputs and compares
EVERY cell in EVERY sheet against the approved baseline values workbook.
Any difference is reported with sheet, row, column, header, old value, new value.

Usage:
    python smoke_test_sigal.py --save       # save current output as new baseline
    python smoke_test_sigal.py              # run smoke test against baseline
"""

import sys
from pathlib import Path

from smoke_test_common import SmokeTestConfig, run_cli


ROOT = Path(__file__).resolve().parent
SIGAL_DIR = ROOT / "Sigal numeros como texto mal interpretados"
BASELINE_DIR = ROOT / "SMOKE_BASELINE" / "SIGAL_20260414_APPROVED"

GALLO_PATH = SIGAL_DIR / "IG_10374_rebuilt.xlsx"
VISUAL_PATH = SIGAL_DIR / "_web_sim_visual.xlsx"

BASELINE_VALUES = BASELINE_DIR / "10374_SIGAL_ARIEL_baseline_values.xlsx"
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
        if boleto in {"79603", "81990", "82749", "84357"}:
            precio_nominal = boletos_sheet.cell(row, columns.get("Precio Nominal", 20)).value
            tipo_cambio = boletos_sheet.cell(row, columns.get("Tipo Cambio", 12)).value
            bruto = boletos_sheet.cell(row, columns.get("Bruto", 13)).value
            lines.append(f"  B={boleto}: PN={precio_nominal}, TC={tipo_cambio}, Bruto={bruto}")

    lines.extend(
        [
            "",
            "Protected invariants:",
            "- B=79603  Precio Nominal=1.4217   Bruto=5686800000",
            "- B=81990  Precio Nominal=1.4307   Bruto=-5722800000",
            "- B=82749  Precio Nominal=0.0011   TC=1280  Bruto=2600000",
            "- B=84357  Precio Nominal=0.0011   TC=1261  Bruto=2550000",
            "- Full workbook values compared cell-by-cell by smoke test.",
        ]
    )
    return lines


CONFIG = SmokeTestConfig(
    title="SIGAL 10374",
    description="Smoke test - Sigal 10374",
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
