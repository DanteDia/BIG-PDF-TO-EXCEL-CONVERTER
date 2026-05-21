from pathlib import Path

import yaml

import explain_audit
import repo_doctor


ROOT = Path(__file__).resolve().parent


def test_case_registry_has_required_cicero_context():
    registry = yaml.safe_load((ROOT / "CASE_REGISTRY.yml").read_text(encoding="utf-8"))

    cicero = registry["cases"]["cicero_10537"]

    assert cicero["status"] == "under-review"
    assert cicero["handoff"] == "CASE_HANDOFFS/CICERO_10537.md"
    assert cicero["latest_audit"]["high"] == 0
    assert cicero["latest_audit"]["review"] == 9
    assert "CIC-PRECIO-TENENCIAS-NEGATIVE-INVESTED-001" in cicero["findings"]


def test_repo_doctor_classifies_local_artifacts():
    assert repo_doctor.is_local_artifact("LOCAL_VERIFY_CURRENT_CICERO_AFTERFIX/file.xlsx")
    assert repo_doctor.is_local_artifact("LOCAL_RELEASE_AUDIT_CICERO_AFTERFIX.json")
    assert repo_doctor.is_local_artifact("SMOKE_BASELINE/CASE/output.pdf")
    assert not repo_doctor.is_local_artifact("pdf_converter/datalab/merge_gallo_visual.py")
    assert not repo_doctor.is_local_artifact("CASE_HANDOFFS/CICERO_10537.md")


def test_explain_audit_matches_registry_workbook_path():
    registry = yaml.safe_load((ROOT / "CASE_REGISTRY.yml").read_text(encoding="utf-8"))
    workbook = ROOT / "LOCAL_VERIFY_CURRENT_CICERO_AFTERFIX" / "CICERO_CURRENT_AFTERFIX_values.xlsx"

    case_id, case = explain_audit.find_case_for_workbook(workbook, registry)

    assert case_id == "cicero_10537"
    assert case["status"] == "under-review"