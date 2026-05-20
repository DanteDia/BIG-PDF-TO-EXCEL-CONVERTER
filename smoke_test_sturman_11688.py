#!/usr/bin/env python
"""
Smoke test for Sturman 11688 case.

Protects against regressions in:
- USD micro-price bruto/neto corruption (STU-MICRO-PRICE-RESCUE-BUG-001)
- Dolar Cable 1000x quantity inflation (STU-ANCHOR-1000X-DEFLATION-001)
- Futuros boletos bruto=0 rule (STU-FUTUROS-BOLETOS-001)
- Full workbook cell-by-cell comparison

Frozen inputs from OCR run 2026-04-20 of Sturman 11688 PDFs.
Reference: web output 20260420_0242 (pre-fix) and 20260420_1624 (post-fix).

Usage:
    python smoke_test_sturman_11688.py --save   # save current output as new baseline
    python smoke_test_sturman_11688.py           # run smoke test against baseline
"""

import sys
from pathlib import Path

from smoke_test_common import SmokeTestConfig, run_cli


ROOT = Path(__file__).resolve().parent
BASELINE_DIR = ROOT / "SMOKE_BASELINE" / "STURMAN_11688_20260420_APPROVED"

GALLO_PATH = BASELINE_DIR / "11688_gallo_frozen.xlsx"
VISUAL_PATH = BASELINE_DIR / "11688_visual_frozen.xlsx"
PRECIO_PATH = BASELINE_DIR / "11688_precio_tenencias_frozen.xlsx"

BASELINE_VALUES = BASELINE_DIR / "11688_STURMAN_JAVIER_baseline_values.xlsx"
BASELINE_JSON = BASELINE_DIR / "baseline_snapshot.json"

MAX_SANE_QTY = 10_000_000_000
MICROPRICE_EXPECTED = {
    "71174": 995000,
    "71873": 994975,
    "72719": 999000,
    "122342": 1000000,
}


def run_pipeline():
    sys.path.insert(0, str(ROOT))
    from pdf_converter.datalab.merge_gallo_visual import GalloVisualMerger

    assert GALLO_PATH.exists(), f"Gallo not found: {GALLO_PATH}"
    assert VISUAL_PATH.exists(), f"Visual not found: {VISUAL_PATH}"

    precio = str(PRECIO_PATH) if PRECIO_PATH.exists() else None
    merger = GalloVisualMerger(str(GALLO_PATH), str(VISUAL_PATH), precio_tenencias_path=precio)
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


def check_microprice(workbook):
    worksheet = workbook["Boletos"]
    bad = []
    found = {}
    for row in range(2, worksheet.max_row + 1):
        boleto = str(worksheet.cell(row, 4).value or "").strip()
        if boleto in MICROPRICE_EXPECTED:
            neto = worksheet.cell(row, 16).value
            expected = MICROPRICE_EXPECTED[boleto]
            try:
                if abs(float(neto or 0) - expected) > 1.0:
                    bad.append(f"  Bol {boleto}: neto={neto}, expected={expected}")
            except (TypeError, ValueError):
                bad.append(f"  Bol {boleto}: neto={neto!r} (non-numeric), expected={expected}")
            found[boleto] = neto

    missing = set(MICROPRICE_EXPECTED) - set(found)
    for boleto in sorted(missing):
        bad.append(f"  Bol {boleto}: NOT FOUND in Boletos sheet")
    return bad


def build_summary_lines(workbook):
    lines = ["", "Micro-price boletos (STU-MICRO-PRICE-RESCUE-BUG-001):"]
    boletos_sheet = workbook["Boletos"]
    for boleto, expected_neto in sorted(MICROPRICE_EXPECTED.items()):
        for row in range(2, boletos_sheet.max_row + 1):
            if str(boletos_sheet.cell(row, 4).value or "").strip() == boleto:
                actual = boletos_sheet.cell(row, 16).value
                lines.append(f"  Bol {boleto}: neto={actual} (expected={expected_neto})")
                break

    lines.extend(
        [
            "",
            "Protected invariants:",
            "- No quantity > 10,000,000,000 in any sheet (1000x inflation guard)",
            "- Bol 71174: neto=995,000 (not 1,053,529 - micro-price fix)",
            "- Bol 71873: neto=994,975 (not 1,053,503 - micro-price fix)",
            "- Bol 72719: neto=999,000 (not 1,057,765 - micro-price fix)",
            "- Bol 122342: neto=1,000,000 (not 1,058,824 - micro-price fix)",
            "- Futuros boletos: bruto=0, neto=gastos (no qty*precio)",
            "- Boletos Dolar Cable code 9249 qty=757575758 (not 757575758000)",
            "- Boletos Dolar Cable code 9323 qty=512974026 (not 512974026000)",
            "- Full workbook values compared cell-by-cell by smoke test",
            "",
            "Reference artifacts in this directory:",
            "- 11688_web_reference_20260420_0242.xlsx: web output BEFORE micro-price fix",
            "- 11688_web_reference_20260420_1624.pdf: web output AFTER micro-price fix",
        ]
    )
    return lines


CONFIG = SmokeTestConfig(
    title="STURMAN 11688",
    description="Smoke test - Sturman 11688",
    baseline_dir=BASELINE_DIR,
    baseline_values=BASELINE_VALUES,
    baseline_json=BASELINE_JSON,
    input_labels=(("Gallo", GALLO_PATH), ("Visual", VISUAL_PATH), ("Precio", PRECIO_PATH)),
    run_pipeline=run_pipeline,
    summary_builder=build_summary_lines,
    baseline_load_data_only=False,
    save_guards=(("  Quantity inflation detected:", check_inflation), ("  Micro-price boletos corrupted:", check_microprice)),
    run_guards=(("  Quantity inflation detected:", check_inflation), ("  Micro-price boletos corrupted:", check_microprice)),
)


if __name__ == "__main__":
    raise SystemExit(run_cli(CONFIG))
