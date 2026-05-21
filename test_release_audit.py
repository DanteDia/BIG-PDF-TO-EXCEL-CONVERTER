from pathlib import Path

from openpyxl import Workbook

from run_release_audit import AuditTarget, audit_target, compare_to_baseline, summarize


def _save_workbook(path: Path, resultado: float) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Resultado Ventas ARS"
    ws.append(["Tipo Operación", "Bruto", "Resultado Calculado(final)"])
    ws.append(["Venta", 100, resultado])
    wb.save(path)


def test_audit_target_classifies_high_review(tmp_path):
    workbook_path = tmp_path / "case_values.xlsx"
    _save_workbook(workbook_path, 150)

    row = audit_target(AuditTarget("case", workbook_path, "approved-smoke", "test"), tmp_path)

    assert row["status"] == "high-review"
    assert row["counts_by_severity"]["high"] == 1
    assert row["counts_by_rule"]["RES-ARS-RATIO-001"] == 1


def test_summarize_counts_approved_high(tmp_path):
    rows = [
        {"bucket": "approved-smoke", "status": "high-review"},
        {"bucket": "approved-smoke", "status": "review"},
        {"bucket": "under-review", "status": "high-review"},
    ]

    summary = summarize(rows)

    assert summary["approved_high_count"] == 1
    assert summary["approved_review_count"] == 1
    assert summary["by_bucket"] == {"approved-smoke": 2, "under-review": 1}


def test_compare_to_baseline_allows_existing_high_but_flags_increase():
    baseline = {
        "targets": [
            {
                "label": "case",
                "path": "case.xlsx",
                "counts_by_severity": {"high": 2},
                "counts_by_rule": {"RES-ARS-RATIO-001": 2},
            }
        ]
    }

    same_rows = [
        {
            "label": "case",
            "path": "case.xlsx",
            "bucket": "approved-smoke",
            "counts_by_severity": {"high": 2},
            "counts_by_rule": {"RES-ARS-RATIO-001": 2},
        }
    ]
    worse_rows = [
        {
            "label": "case",
            "path": "case.xlsx",
            "bucket": "approved-smoke",
            "counts_by_severity": {"high": 3},
            "counts_by_rule": {"RES-ARS-RATIO-001": 3},
        }
    ]

    assert compare_to_baseline(same_rows, baseline) == []
    regressions = compare_to_baseline(worse_rows, baseline)
    assert any(regression.startswith("HIGH_INCREASE") for regression in regressions)
    assert any(regression.startswith("RESULT_RULE_INCREASE") for regression in regressions)
