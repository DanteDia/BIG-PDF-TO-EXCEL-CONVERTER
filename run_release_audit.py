from __future__ import annotations

import argparse
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook

from pdf_converter.datalab.economic_sanity import validate_workbook


ROOT = Path(__file__).resolve().parent


@dataclass(frozen=True)
class AuditTarget:
    label: str
    path: Path
    bucket: str
    source: str
    expected_state: str = "review"


def _existing_latest(paths: Iterable[Path]) -> Path | None:
    existing = [path for path in paths if path.exists()]
    if not existing:
        return None
    return max(existing, key=lambda path: path.stat().st_mtime)


def smoke_baseline_targets(root: Path = ROOT) -> list[AuditTarget]:
    from smoke_test_koltan_13353 import CONFIG as KOLTAN_CONFIG
    from smoke_test_prida import CONFIG as PRIDA_CONFIG
    from smoke_test_sigal import CONFIG as SIGAL_CONFIG
    from smoke_test_sturman import CONFIG as STURMAN_CONFIG
    from smoke_test_sturman_11688 import CONFIG as STURMAN_11688_CONFIG

    configs = [
        SIGAL_CONFIG,
        PRIDA_CONFIG,
        STURMAN_CONFIG,
        STURMAN_11688_CONFIG,
        KOLTAN_CONFIG,
    ]
    targets = [
        AuditTarget(config.title, config.baseline_values, "approved-smoke", "smoke-baseline", "approved")
        for config in configs
    ]
    targets.extend(
        [
            AuditTarget(
                "AGUIAR pre Martin baseline",
                root / "SMOKE_BASELINE" / "AGUIAR_20260318_PRE_MARTIN" / "AGUIAR_SMOKE_values.xlsx",
                "approved-smoke",
                "run_smoke_suite",
                "approved",
            ),
            AuditTarget(
                "AGUIAR same input current",
                root / "SMOKE_BASELINE" / "AGUIAR_20260318_POST_MARTIN" / "AGUIAR_SAME_INPUT_values.xlsx",
                "approved-smoke",
                "run_smoke_suite",
                "approved",
            ),
            AuditTarget(
                "CANULLO approved baseline",
                root
                / "SMOKE_BASELINE"
                / "CANULLO_20260326_APPROVED"
                / "10600_CANULLO_MARTHA_NOEMI_Resumen_Impositivo_FIXED_values.xlsx",
                "approved-smoke",
                "run_smoke_suite",
                "approved",
            ),
            AuditTarget(
                "CANULLO current root",
                root / "10600_CANULLO_MARTHA_NOEMI_Resumen_Impositivo_FIXED_values.xlsx",
                "approved-smoke",
                "run_smoke_suite",
                "approved",
            ),
        ]
    )
    glozman = _existing_latest(
        [
            root
            / "Ejemplo Glozman error moneda pesos en seccion USD"
            / "12766_GLOZMAN_DARIO_EDMUNDO_Resumen_Impositivo_FIXED_v2.xlsx",
            root
            / "Ejemplo Glozman error moneda pesos en seccion USD"
            / "12766_GLOZMAN_DARIO_EDMUNDO_Resumen_Impositivo_FIXED.xlsx",
        ]
    )
    if glozman:
        targets.append(AuditTarget("GLOZMAN latest fixed", glozman, "approved-smoke", "run_smoke_suite", "approved"))
    return targets


def under_review_targets(root: Path = ROOT) -> list[AuditTarget]:
    patterns = [
        ("CICERO", root / "Caso Cicero", "**/*FIXED_values.xlsx"),
        ("DEL RIO", root / "Del RIO mal OCR tenencias por cantidad muy grande", "**/*values.xlsx"),
        ("HALAC", root / "Ejemplo HALAC LEON - error importe y precio desde gallo", "**/*values.xlsx"),
    ]
    targets: list[AuditTarget] = []
    for label, folder, pattern in patterns:
        if not folder.exists():
            continue
        matches = sorted(folder.glob(pattern), key=lambda path: path.stat().st_mtime, reverse=True)
        for path in matches[:3]:
            targets.append(AuditTarget(f"{label} {path.name}", path, "under-review", "case-output", "known-failing"))
    return targets


def issue_counts(issues: Iterable[dict]) -> tuple[dict[str, int], dict[str, int]]:
    by_severity: dict[str, int] = {}
    by_rule: dict[str, int] = {}
    for issue in issues:
        severity = str(issue.get("severity") or "unknown")
        rule_id = str(issue.get("rule_id") or "unknown")
        by_severity[severity] = by_severity.get(severity, 0) + 1
        by_rule[rule_id] = by_rule.get(rule_id, 0) + 1
    return by_severity, by_rule


def audit_target(target: AuditTarget, root: Path = ROOT) -> dict:
    relative_path = str(target.path.relative_to(root)) if target.path.is_relative_to(root) else str(target.path)
    row = {
        "label": target.label,
        "path": relative_path,
        "bucket": target.bucket,
        "source": target.source,
        "expected_state": target.expected_state,
    }
    if not target.path.exists():
        return {**row, "status": "missing", "issue_count": 0, "counts_by_severity": {}, "counts_by_rule": {}, "top_issues": []}

    try:
        workbook = load_workbook(target.path, data_only=True)
        report = validate_workbook(workbook).to_dict()
    except Exception as exc:
        return {
            **row,
            "status": "error",
            "error": f"{type(exc).__name__}: {exc}",
            "issue_count": 0,
            "counts_by_severity": {},
            "counts_by_rule": {},
            "top_issues": [],
        }

    issues = report["issues"]
    counts_by_severity, counts_by_rule = issue_counts(issues)
    status = "clean"
    if counts_by_severity.get("high", 0):
        status = "high-review"
    elif issues:
        status = "review"

    return {
        **row,
        "status": status,
        "issue_count": len(issues),
        "counts_by_severity": counts_by_severity,
        "counts_by_rule": counts_by_rule,
        "top_issues": issues[:10],
    }


def build_targets(root: Path, include_under_review: bool, extra_paths: Iterable[Path]) -> list[AuditTarget]:
    targets = smoke_baseline_targets(root)
    if include_under_review:
        targets.extend(under_review_targets(root))
    for path in extra_paths:
        resolved = path if path.is_absolute() else root / path
        targets.append(AuditTarget(path.name, resolved, "manual", "cli", "review"))
    seen: set[Path] = set()
    unique: list[AuditTarget] = []
    for target in targets:
        key = target.path.resolve()
        if key in seen:
            continue
        seen.add(key)
        unique.append(target)
    return unique


def summarize(rows: list[dict]) -> dict:
    summary = {
        "total": len(rows),
        "by_bucket": {},
        "by_status": {},
        "approved_high_count": 0,
        "approved_review_count": 0,
    }
    for row in rows:
        bucket = row["bucket"]
        status = row["status"]
        summary["by_bucket"][bucket] = summary["by_bucket"].get(bucket, 0) + 1
        summary["by_status"][status] = summary["by_status"].get(status, 0) + 1
        if bucket == "approved-smoke" and status == "high-review":
            summary["approved_high_count"] += 1
        if bucket == "approved-smoke" and status == "review":
            summary["approved_review_count"] += 1
    return summary


def _target_key(row: dict) -> str:
    return str(row.get("path") or row.get("label") or "")


def compare_to_baseline(rows: list[dict], baseline: dict) -> list[str]:
    baseline_rows = {_target_key(row): row for row in baseline.get("targets", [])}
    regressions: list[str] = []
    for row in rows:
        key = _target_key(row)
        previous = baseline_rows.get(key)
        if not previous:
            if row.get("bucket") == "approved-smoke" and row.get("counts_by_severity", {}).get("high", 0):
                regressions.append(f"NEW_TARGET_WITH_HIGH {row['label']} high={row['counts_by_severity']['high']}")
            continue
        current_high = int(row.get("counts_by_severity", {}).get("high", 0) or 0)
        previous_high = int(previous.get("counts_by_severity", {}).get("high", 0) or 0)
        if current_high > previous_high:
            regressions.append(f"HIGH_INCREASE {row['label']} {previous_high}->{current_high}")

        current_rules = row.get("counts_by_rule", {}) or {}
        previous_rules = previous.get("counts_by_rule", {}) or {}
        for rule_id, current_count in sorted(current_rules.items()):
            previous_count = int(previous_rules.get(rule_id, 0) or 0)
            if int(current_count or 0) > previous_count and str(rule_id).startswith("RES-"):
                regressions.append(f"RESULT_RULE_INCREASE {row['label']} {rule_id} {previous_count}->{current_count}")
    return regressions


def print_report(rows: list[dict], summary: dict) -> None:
    print("Release audit")
    print(f"Targets: {summary['total']}")
    print(f"By status: {summary['by_status']}")
    print(f"By bucket: {summary['by_bucket']}")
    print()
    for row in rows:
        counts = row["counts_by_severity"]
        rules = ", ".join(f"{rule}={count}" for rule, count in sorted(row["counts_by_rule"].items())[:5])
        print(
            f"[{row['bucket']}] {row['status']}: {row['label']} "
            f"issues={row['issue_count']} severity={counts} rules={rules}"
        )
        if row.get("error"):
            print(f"  error: {row['error']}")
        print(f"  path: {row['path']}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Run non-blocking economic sanity audit over release candidate workbooks.")
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--approved-only", action="store_true", help="Only audit approved smoke/reference targets.")
    parser.add_argument("--json-output", type=Path, help="Optional path to save full audit JSON.")
    parser.add_argument("--compare-baseline", type=Path, help="Compare against a prior audit JSON and fail on new high/result triggers.")
    parser.add_argument("--fail-on-approved-high", action="store_true", help="Exit non-zero if approved targets have high triggers.")
    parser.add_argument("paths", nargs="*", type=Path, help="Extra workbook paths to audit.")
    args = parser.parse_args()

    root = args.root.resolve()
    targets = build_targets(root, include_under_review=not args.approved_only, extra_paths=args.paths)
    rows = [audit_target(target, root) for target in targets]
    summary = summarize(rows)
    output = {"summary": summary, "targets": rows}

    print_report(rows, summary)
    regressions: list[str] = []
    if args.compare_baseline:
        baseline_path = args.compare_baseline if args.compare_baseline.is_absolute() else root / args.compare_baseline
        baseline = json.loads(baseline_path.read_text(encoding="utf-8"))
        regressions = compare_to_baseline(rows, baseline)
        if regressions:
            print("\nAudit regressions versus baseline:")
            for regression in regressions:
                print(f"  {regression}")
        else:
            print("\nAudit comparison: no new high/result triggers versus baseline.")
    if args.json_output:
        output_path = args.json_output if args.json_output.is_absolute() else root / args.json_output
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(json.dumps(output, indent=2, ensure_ascii=False), encoding="utf-8")
        print(f"\nSaved JSON: {output_path}")

    if args.fail_on_approved_high and summary["approved_high_count"]:
        return 1
    if regressions:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
