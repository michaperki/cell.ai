# MEMORY.md — Project State

Role: Live project state, architecture snapshot, and current focus. Update this file at the end of each work session.

Last updated: 2026-03-22

---

## Architecture Snapshot
- Platform: Windows app with headless CLI runner (see `SOUL.md` for operating constraints).
- Entry point: `Program.cs` (CLI flags for headless runs).
- Headless runner: `Core/HeadlessTestRunner.cs` (writes artifacts under `tests/output/`).
- Core spreadsheet engine and helpers: `Core/Spreadsheet.cs`, `Core/FormulaEngine.cs`, `Core/UndoManager.cs`.

## Current Focus
- In Progress
  - Stage 4 autonomy loop operational: regression gate + failure taxonomy + prompt versioning are in place.
  - Active prompt: v2 (system_planner_v2.md). Suite at 32/35 (gate PASS, 0 regressions).
  - test_39 still failing: cell_mismatch C9 ('active' vs 'actve') and under_generation (2<3). Next: try v3 with a few-shot categorical correction example or accept as model limit if no improvement.
- Next
  - Consider adding explicit examples for categorical normalization; evaluate impact on test_39 while maintaining gate PASS.
- Blocked
  - None noted.

## Recent Changes
- Stage 4 infrastructure landed:
  - Regression gate with baseline blessing and regression report (JSON + MD).
  - Failure taxonomy doc mapping real assertion keys to writable surfaces.
  - Prompt versioning: system prompt extracted to versioned files; run history records prompt_version.
- First Stage 4 experiment: v2 prompt adds categorical label typo correction guidance. test_39 unchanged; overall suite improved (31→32/35), gate PASS. v2 kept and blessed as baseline.

## Pointers & Artifacts
- Headless outputs: `tests/output/` (per-step artifacts: `*.before.workbook.json`, `*.workbook.json`, `*.plan.json`, `*.export.json`, `*.agent.txt`, `*.observations.json`).
- Aggregated results: `tests/output/results_summary.json`, `tests/output/scorecard.json`.
- Run history: `tests/output/run_history.jsonl` (append-only).
- Dashboard (agents): `tests/output/DASHBOARD.md`.
- Dashboard (humans): `tests/output/dashboard.workbook.json` (openable in app).
- Operating manual: `SOUL.md` (durable rules and session protocol).

## Maintenance Notes
- At session start: skim `SOUL.md` and this file.
- At session end: update the bullets under Current Focus (what’s in progress, what’s next, what’s blocked).
- Promote durable lessons from the journal into this file when they become stable operating truths.
