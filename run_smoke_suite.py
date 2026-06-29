from __future__ import annotations

from pathlib import Path
from typing import Iterable
import sys

from compare_workbooks import compare_workbooks
from smoke_test_common import SmokeTestConfig, run_smoke
import verify_regression_cases as regression


ROOT = Path(__file__).resolve().parent
CANULLO_DIR = ROOT / "SMOKE_BASELINE" / "CANULLO_20260326_APPROVED"


def _assert(condition: bool, message: str, failures: list[str]) -> None:
    if not condition:
        failures.append(message)


def _check_glozman(failures: list[str]) -> None:
    glozman = regression._latest_existing(
        ROOT / "Ejemplo Glozman error moneda pesos en seccion USD" / "12766_GLOZMAN_DARIO_EDMUNDO_Resumen_Impositivo_FIXED_v2.xlsx",
        ROOT / "Ejemplo Glozman error moneda pesos en seccion USD" / "12766_GLOZMAN_DARIO_EDMUNDO_Resumen_Impositivo_FIXED.xlsx",
    )

    _assert(glozman.exists(), f"Falta workbook GLOZMAN: {glozman}", failures)
    if failures:
        return

    glozman_ars_rows = regression.count_non_empty_rows(glozman, "Rentas Dividendos ARS")
    glozman_usd_rows, glozman_bad = regression.check_currency_rows(glozman, "Rentas Dividendos USD", "USD")

    _assert(glozman_ars_rows == 26, f"GLOZMAN ARS rows esperado=26 actual={glozman_ars_rows}", failures)
    _assert(glozman_usd_rows == 6, f"GLOZMAN USD rows esperado=6 actual={glozman_usd_rows}", failures)
    _assert(not glozman_bad, f"GLOZMAN filas USD mal clasificadas: {glozman_bad}", failures)


def _check_canullo_approved(failures: list[str]) -> None:
    baseline = CANULLO_DIR / "10600_CANULLO_MARTHA_NOEMI_Resumen_Impositivo_FIXED_values.xlsx"
    current = ROOT / "10600_CANULLO_MARTHA_NOEMI_Resumen_Impositivo_FIXED_values.xlsx"

    _assert(baseline.exists(), f"Falta baseline CANULLO: {baseline}", failures)
    _assert(current.exists(), f"Falta workbook actual CANULLO: {current}", failures)
    if failures:
        return

    exact_diffs = compare_workbooks(
        baseline,
        current,
        tolerance=1e-9,
        max_diffs=25,
    )
    for diff in exact_diffs:
        failures.append(f"CANULLO exact diff: {diff}")


def _print_section(title: str, lines: Iterable[str]) -> None:
    print(f"\n=== {title} ===")
    for line in lines:
        print(line)


def _dedicated_smoke_configs() -> list[SmokeTestConfig]:
    from smoke_test_koltan_13353 import CONFIG as KOLTAN_CONFIG
    from smoke_test_prida import CONFIG as PRIDA_CONFIG
    from smoke_test_sigal import CONFIG as SIGAL_CONFIG
    from smoke_test_sturman import CONFIG as STURMAN_CONFIG
    from smoke_test_sturman_11688 import CONFIG as STURMAN_11688_CONFIG

    return [
        SIGAL_CONFIG,
        PRIDA_CONFIG,
        STURMAN_CONFIG,
        STURMAN_11688_CONFIG,
        KOLTAN_CONFIG,
    ]


def _run_dedicated_smokes(failures: list[str], passes: list[str]) -> None:
    smoke_labels = {
        "SIGAL 10374": "SIGAL 10374: cell-by-cell OK",
        "PRIDA 10488": "PRIDA 10488: cell-by-cell OK",
        "J_STURMAN 2797": "STURMAN 2797: cell-by-cell OK",
        "STURMAN 11688": "STURMAN 11688: cell-by-cell OK + micro-price guard + inflation guard",
        "KOLTAN 13353": "KOLTAN 13353: cell-by-cell OK + inflation guard",
    }

    for config in _dedicated_smoke_configs():
        print(f"\n>>> Running dedicated smoke: {config.title}")
        exit_code = run_smoke(config)
        if exit_code != 0:
            failures.append(f"{config.title} FAILED")
            continue
        passes.append(smoke_labels.get(config.title, f"{config.title}: OK"))


def main() -> int:
    failures: list[str] = []
    passes = [
        "GLOZMAN: ARS/USD split consistente",
        "CANULLO approved: workbook completo sin desvíos",
    ]

    _check_glozman(failures)
    _check_canullo_approved(failures)

    if not failures:
        _run_dedicated_smokes(failures, passes)

    if failures:
        _print_section("SMOKE FAIL", failures)
        return 1

    _print_section("SMOKE PASS", passes)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())