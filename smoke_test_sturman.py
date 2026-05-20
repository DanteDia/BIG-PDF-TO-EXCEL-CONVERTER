#!/usr/bin/env python
"""
Smoke test for J_Sturman 2797 case (options + ejercicio compra titular).

Runs the full merge pipeline from frozen Visual input and compares
EVERY cell in EVERY sheet against the approved baseline values workbook.
Any difference is reported with sheet, row, column, header, old value, new value.

Key invariants protected:
- Ejercicio Compra Titular qty=-20 row is filtered out
- Ejercicio qty=2000 row is mapped to YPFD (code 710) as Acciones
- Running stock PPP for YPFD incorporates the ejercicio purchase

Usage:
    python smoke_test_sturman.py --save       # save current output as new baseline
    python smoke_test_sturman.py              # run smoke test against baseline
"""

import shutil
import sys
from pathlib import Path

from smoke_test_common import SmokeTestConfig, run_cli


ROOT = Path(__file__).resolve().parent
STURMAN_DIR = ROOT / "J_STURMAN_Opciones 4 ceros post coma mal interpretado"
BASELINE_DIR = ROOT / "SMOKE_BASELINE" / "STURMAN_20260417_APPROVED"

VISUAL_PATH = STURMAN_DIR / "2797bolsa_fixed4.xlsx"
GALLO_PATH = VISUAL_PATH
CLIENT_FINAL = STURMAN_DIR / "2797_J_STURMAN_SA_Resumen_Impositivo_20260417_1601_cliente_final_Smoke.xlsx"

BASELINE_VALUES = BASELINE_DIR / "2797_J_STURMAN_baseline_values.xlsx"
BASELINE_JSON = BASELINE_DIR / "baseline_snapshot.json"


def run_pipeline():
    sys.path.insert(0, str(ROOT))
    from pdf_converter.datalab.merge_gallo_visual import GalloVisualMerger

    assert VISUAL_PATH.exists(), f"Visual not found: {VISUAL_PATH}"

    merger = GalloVisualMerger(str(GALLO_PATH), str(VISUAL_PATH))
    _, workbook = merger.merge("both")
    return workbook


def copy_client_reference():
    if not CLIENT_FINAL.exists():
        return []
    destination = BASELINE_DIR / CLIENT_FINAL.name
    shutil.copy2(str(CLIENT_FINAL), str(destination))
    return [f"Copied client  : {destination.name}"]


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
        operation = str(boletos_sheet.cell(row, columns.get("Tipo Operación", 6)).value or "").lower()
        code = boletos_sheet.cell(row, columns.get("Cod.Instrum", 7)).value
        especie = boletos_sheet.cell(row, columns.get("Instrumento", 8)).value
        quantity = boletos_sheet.cell(row, columns.get("Cantidad", 10)).value
        if "ejercicio" in operation or str(code) == "710":
            lines.append(f"  Row {row}: oper={operation}, code={code}, especie={especie}, qty={quantity}")

    lines.extend(
        [
            "",
            "Protected invariants:",
            "- Ejercicio Compra Titular qty=-20 contract closure row is ABSENT from Boletos",
            "- Ejercicio qty=2000 row is mapped to YPFD (code 710) as Acciones",
            "- YPFC49000D / YPFC55000D resolve to YPFD via _resolve_option_underlying()",
            "- Full workbook values compared cell-by-cell by smoke test.",
        ]
    )
    return lines


CONFIG = SmokeTestConfig(
    title="J_STURMAN 2797",
    description="Smoke test - J_Sturman 2797",
    baseline_dir=BASELINE_DIR,
    baseline_values=BASELINE_VALUES,
    baseline_json=BASELINE_JSON,
    input_labels=(("Visual", VISUAL_PATH),),
    run_pipeline=run_pipeline,
    summary_builder=build_summary_lines,
    baseline_load_data_only=True,
    extra_save_actions=(copy_client_reference,),
)


if __name__ == "__main__":
    raise SystemExit(run_cli(CONFIG))
