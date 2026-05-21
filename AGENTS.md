# Agent Instructions

Read [CLAUDE.md](CLAUDE.md) first. It is the canonical operating contract for this repository.

In short:

1. Start from [CURRENT_STATE.md](CURRENT_STATE.md) and [CASE_REGISTRY.yml](CASE_REGISTRY.yml).
2. Use [CASE_HANDOFFS](CASE_HANDOFFS) for case-specific state.
3. Check [CASE_FINDINGS_INDEX.md](CASE_FINDINGS_INDEX.md) before changing behavior.
4. Use `python repo_doctor.py` and focused tests/audits before committing.
5. Do not commit local generated outputs unless explicitly promoted.