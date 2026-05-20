#!/usr/bin/env python
"""
Smoke test for Koltan 13353 case.

Protects against regressions in:
- Full workbook cell-by-cell comparison
- 1000x quantity inflation artifacts

Frozen inputs from OCR run 2026-04-20 of Koltan 13353 PDFs.

Usage:
    python smoke_test_koltan_13353.py --save   # save current output as new baseline
    python smoke_test_koltan_13353.py          # run smoke test against baseline
"""

import sys
from pathlib import Path

from smoke_test_common import SmokeTestConfig, run_cli


ROOT = Path(__file__).resolve().parent
BASELINE_DIR = ROOT / "SMOKE_BASELINE" / "KOLTAN_13353_20260420_APPROVED"

GALLO_PATH = BASELINE_DIR / "13353_gallo_frozen.xlsx"
VISUAL_PATH = BASELINE_DIR / "13353_visual_frozen.xlsx"
PRECIO_PATH = BASELINE_DIR / "13353_precio_tenencias_frozen.xlsx"

BASELINE_VALUES = BASELINE_DIR / "13353_KOLTAN_baseline_values.xlsx"
BASELINE_JSON = BASELINE_DIR / "baseline_snapshot.json"

MAX_SANE_QTY = 10_000_000_000


def run_pipeline():
    sys.path.insert(0, str(ROOT))
    from pdf_converter.datalab.merge_gallo_visual import GalloVisualMerger

    assert GALLO_PATH.exists(), f"Gallo not found: {GALLO_PATH}"
    assert VISUAL_PATH.exists(), f"Visual not found: {VISUAL_PATH}"

    precio = str(PRECIO_PATH) if PRECIO_PATH.exists() else None
    merger = GalloVisualMerger(
        str(GALLO_PATH),
        str(VISUAL_PATH),
        precio_tenencias_path=precio,
    )
    _, workbook = merger.merge("both")
    return workbook


def check_inflation(workbook):
    bad = []
    for sheet_name in workbook.sheetnames:
        if sheet_name in CONFIG.aux_sheets:
            continue
        worksheet = workbook[sheet_name]
        for row in range(2, worksheet.max_row + 1):
            for col in range(1, worksheet.max_column + 1):
                value = worksheet.cell(row, col).value
                if isinstance(value, str):
                    continue
                try:
                    if abs(float(value)) > MAX_SANE_QTY:
                        header = worksheet.cell(1, col).value or f"Col{col}"
                        bad.append(f"  {sheet_name} Row {row} Col {col} ({header}): {value}")
                except (TypeError, ValueError):
                    pass
    return bad


def build_summary_lines(_workbook):
    return [
        "",
        "Protected invariants:",
        "- No quantity > 10,000,000,000 in non-aux sheets",
        "- Full workbook values compared cell-by-cell by smoke test",
    ]


CONFIG = SmokeTestConfig(
    title="KOLTAN 13353",
    description="Smoke test - Koltan 13353",
    baseline_dir=BASELINE_DIR,
    baseline_values=BASELINE_VALUES,
    baseline_json=BASELINE_JSON,
    input_labels=(("Gallo", GALLO_PATH), ("Visual", VISUAL_PATH), ("Precio", PRECIO_PATH)),
    run_pipeline=run_pipeline,
    summary_builder=build_summary_lines,
    baseline_load_data_only=False,
    save_guards=(("  Quantity inflation detected:", check_inflation),),
    run_guards=(("  Quantity inflation detected:", check_inflation),),
)


if __name__ == "__main__":
    raise SystemExit(run_cli(CONFIG))
