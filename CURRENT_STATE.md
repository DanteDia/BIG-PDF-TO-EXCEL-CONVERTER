# Current State

Last updated: 2026-05-21

## GitHub State

- Branch: `main`
- Latest pushed commit: `6aed231` (`fix: stabilize Cicero recovered-cost review flow`)
- Remote: `origin/main` on `DanteDia/BIG-PDF-TO-EXCEL-CONVERTER`
- Important note: the working tree may contain local case folders, generated Excel/PDF outputs, worktrees, and investigation scripts. Do not commit those unless they are explicitly promoted to repo evidence.

## Production Flow

1. Streamlit app: [export_validation/app_datalab.py](export_validation/app_datalab.py)
2. OCR/markdown parsing: [pdf_converter/datalab/md_to_excel.py](pdf_converter/datalab/md_to_excel.py)
3. Postprocess/OCR repair: [pdf_converter/datalab/postprocess.py](pdf_converter/datalab/postprocess.py)
4. Merge and fiscal calculations: [pdf_converter/datalab/merge_gallo_visual.py](pdf_converter/datalab/merge_gallo_visual.py)
5. Economic sanity review: [pdf_converter/datalab/economic_sanity.py](pdf_converter/datalab/economic_sanity.py)
6. Deterministic case replay: [generate_case_outputs.py](generate_case_outputs.py)
7. Release/audit matrix: [run_release_audit.py](run_release_audit.py)

## Context Sources To Read First

1. [CLAUDE.md](CLAUDE.md) for the durable agent/human operating contract.
2. [CASE_REGISTRY.yml](CASE_REGISTRY.yml) for machine-readable case status.
3. [CASE_FINDINGS_INDEX.md](CASE_FINDINGS_INDEX.md) for validated business and fiscal rules.
4. This file for current branch/case status.
5. The relevant handoff in [CASE_HANDOFFS](CASE_HANDOFFS).
6. The latest audit JSON for any case being discussed.

## Case Status

| Case | Status | Current Signal | Notes |
| --- | --- | --- | --- |
| Cicero 10537 | `under-review` | `9 review`, `0 high` on latest local audit | Fixes are pushed; remaining reviews are SUPV and AMD ratio rows. See [CASE_HANDOFFS/CICERO_10537.md](CASE_HANDOFFS/CICERO_10537.md). |
| Del Rio 10980 | `approved-smoke-candidate` | User confirmed it is an approved smoke | Should be promoted to the blocking smoke suite when doing release hardening. |
| Approved smoke baselines | `release-floor` | Use frozen inputs, not fresh OCR | Do not rebaseline without a documented case finding. |
| Fresh full-flow outputs | `investigation` | Useful for review, not automatic truth | Keep separate from approved smokes until manually accepted. |

## Validation Commands

Run focused tests for the latest Cicero/review-flow changes:

```powershell
python -m pytest test_precio_tenencias_zero_cost.py test_economic_sanity_validation.py test_md_to_excel_gallo_position_sections.py test_release_audit.py
```

Run the Cicero local audit after regenerating its workbook:

```powershell
python run_release_audit.py LOCAL_VERIFY_CURRENT_CICERO_AFTERFIX\CICERO_CURRENT_AFTERFIX_values.xlsx --json-output LOCAL_RELEASE_AUDIT_CICERO_AFTERFIX.json
```

Check pushed state:

```powershell
git rev-parse HEAD
git rev-parse origin/main
git log -1 --oneline
```

Run repo context checks:

```powershell
python repo_doctor.py
python explain_audit.py LOCAL_VERIFY_CURRENT_CICERO_AFTERFIX\CICERO_CURRENT_AFTERFIX_values.xlsx
```

## Commit Hygiene

- Commit source, tests, and docs only by default.
- Do not commit local generated outputs such as `LOCAL_VERIFY_*`, `LOCAL_RELEASE_AUDIT*.json`, case folders, or disposable worktrees unless explicitly asked.
- Before changing fiscal behavior, search [CASE_FINDINGS_INDEX.md](CASE_FINDINGS_INDEX.md) for prior rules and document whether the new rule is compatible, narrower, or superseding.

## Next Priorities

1. Test a new production case from the web/app now that `main` includes the Cicero recovered-cost and review-flow fixes.
2. Review Cicero's remaining `9` ARS ratio reviews: SUPV rows 33-34 and AMD rows 326-332 in the latest local output.
3. Promote Del Rio into the blocking smoke suite if release hardening continues.