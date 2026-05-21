# Repository Operating Instructions

This file is the durable operating contract for AI agents and humans working in this repository. Follow it before creating new architecture or changing business behavior.

## First Files To Read

1. [CURRENT_STATE.md](CURRENT_STATE.md) for branch, case status, commands, and next priorities.
2. [CASE_REGISTRY.yml](CASE_REGISTRY.yml) for machine-readable case status.
3. The relevant handoff in [CASE_HANDOFFS](CASE_HANDOFFS) when working on a named case.
4. [CASE_FINDINGS_INDEX.md](CASE_FINDINGS_INDEX.md) before changing parser, postprocess, merge, routing, validation, output layout, or fiscal logic.

## Mandatory Workflow

1. Run or inspect `python repo_doctor.py` at the start of substantial work.
2. Classify the work: parser, postprocess, merge, validation, app/export, docs, or local artifact.
3. If behavior changes, check [CASE_FINDINGS_INDEX.md](CASE_FINDINGS_INDEX.md) and record whether the new rule is compatible, narrower, or superseding.
4. Add or update a focused test for any fiscal/parser/merge/validation rule.
5. Run the smallest useful verification first, then the broader audit when appropriate.
6. Commit source, tests, and docs only. Do not commit generated local outputs unless the user explicitly asks to promote them.

## Case State Vocabulary

- `approved-smoke`: blocks releases; frozen inputs and baseline are accepted.
- `approved-manual`: accepted by human review, not yet automated as smoke.
- `approved-smoke-candidate`: should be promoted into the blocking smoke suite.
- `under-review`: useful for testing and investigation, not release truth.
- `known-failing`: do not chase without a new request.
- `superseded`: kept for history only.

## Local Artifact Policy

Do not commit these by default:

- `LOCAL_VERIFY_*`
- `LOCAL_RELEASE_AUDIT*.json`
- disposable `.worktrees/`
- fresh case folders copied from user machines
- generated `.xlsx` / `.pdf` review outputs

Commit these by default when they change intentionally:

- source code
- focused tests
- [CASE_FINDINGS_INDEX.md](CASE_FINDINGS_INDEX.md)
- [CASE_REGISTRY.yml](CASE_REGISTRY.yml)
- [CURRENT_STATE.md](CURRENT_STATE.md)
- [CASE_HANDOFFS](CASE_HANDOFFS)

## Useful Commands

```powershell
python repo_doctor.py
python explain_audit.py <values-workbook.xlsx>
python run_release_audit.py <values-workbook.xlsx> --json-output LOCAL_RELEASE_AUDIT_CASE.json
python -m pytest test_precio_tenencias_zero_cost.py test_economic_sanity_validation.py test_md_to_excel_gallo_position_sections.py test_release_audit.py
```

## Design Rule

Prefer small tools that are actually invoked by the workflow over large documentation-only architecture. Any new process should either be read by agents at startup, consumed by a script, or used in a test/release command.