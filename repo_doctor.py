from __future__ import annotations

import argparse
import subprocess
from pathlib import Path
from typing import Any

import yaml


ROOT = Path(__file__).resolve().parent
LOCAL_OUTPUT_PREFIXES = ("LOCAL_VERIFY_", "LOCAL_RELEASE_AUDIT")


def run_git(args: list[str]) -> tuple[int, str]:
    completed = subprocess.run(
        ["git", *args],
        cwd=ROOT,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        check=False,
    )
    return completed.returncode, completed.stdout.strip()


def load_registry(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {"cases": {}}
    return yaml.safe_load(path.read_text(encoding="utf-8")) or {"cases": {}}


def status_lines() -> list[str]:
    code, output = run_git(["status", "--short"])
    if code != 0:
        return [f"ERROR reading git status: {output}"]
    return [line for line in output.splitlines() if line.strip()]


def is_local_artifact(path: str) -> bool:
    normalized = path.replace("\\", "/").strip('"')
    name = normalized.split("/", 1)[0]
    if name.startswith(LOCAL_OUTPUT_PREFIXES):
        return True
    if name == ".worktrees":
        return True
    if normalized.lower().endswith((".xlsx", ".xlsm", ".pdf")):
        return True
    return False


def summarize_dirty(lines: list[str]) -> tuple[list[str], list[str]]:
    local: list[str] = []
    source: list[str] = []
    for line in lines:
        path = line[3:].strip()
        if is_local_artifact(path):
            local.append(line)
        else:
            source.append(line)
    return source, local


def print_case_registry(registry: dict[str, Any]) -> None:
    cases = registry.get("cases") or {}
    print("\nCase registry")
    if not cases:
        print("  CASE_REGISTRY.yml missing or empty")
        return
    for case_id, case in cases.items():
        status = case.get("status", "unknown")
        client = case.get("client") or case_id
        audit = case.get("latest_audit") or {}
        high = audit.get("high")
        review = audit.get("review")
        signal = ""
        if high is not None or review is not None:
            signal = f" high={high or 0} review={review or 0}"
        print(f"  {case_id}: {status} - {client}{signal}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Print repo state, case registry, and local artifact warnings.")
    parser.add_argument("--strict", action="store_true", help="Exit non-zero when source/docs changes are pending.")
    parser.add_argument("--registry", type=Path, default=ROOT / "CASE_REGISTRY.yml")
    args = parser.parse_args()

    registry_path = args.registry if args.registry.is_absolute() else ROOT / args.registry
    registry = load_registry(registry_path)

    _, branch = run_git(["branch", "--show-current"])
    _, head = run_git(["rev-parse", "--short", "HEAD"])
    _, origin = run_git(["rev-parse", "--short", "origin/main"])
    _, last = run_git(["log", "-1", "--oneline"])
    dirty = status_lines()
    source_dirty, local_dirty = summarize_dirty(dirty)

    print("Repo doctor")
    print(f"  branch: {branch or 'unknown'}")
    print(f"  HEAD: {head or 'unknown'}")
    print(f"  origin/main: {origin or 'unknown'}")
    print(f"  aligned: {head == origin}")
    print(f"  last commit: {last or 'unknown'}")

    print_case_registry(registry)

    print("\nWorking tree")
    if not dirty:
        print("  clean")
    else:
        print(f"  source/docs pending: {len(source_dirty)}")
        for line in source_dirty[:20]:
            print(f"    {line}")
        if len(source_dirty) > 20:
            print(f"    ... {len(source_dirty) - 20} more")
        print(f"  local/generated pending: {len(local_dirty)}")
        for line in local_dirty[:20]:
            print(f"    {line}")
        if len(local_dirty) > 20:
            print(f"    ... {len(local_dirty) - 20} more")

    print("\nRead first")
    print("  CLAUDE.md")
    print("  CURRENT_STATE.md")
    print("  CASE_REGISTRY.yml")
    print("  CASE_FINDINGS_INDEX.md before behavior changes")

    if args.strict and source_dirty:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())