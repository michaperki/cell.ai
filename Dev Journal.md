# Dev Journal

> Chronological record of experiments, decisions, and fixes. For older entries, see `Dev Journal Archive.md`.

## Index (most recent first)

| Date | Entry | Key decision |
|------|-------|-------------|
| 03-22 | Stage 4 infra + first experiment | Kept v2 prompt; gate PASS |
| 03-22 | Restored live memory doc | Reintroduced `memory/MEMORY.md`; kept SOUL references |
| 03-22 | Provider-backed full suite + 8 new hard tests | SanitizePlanToBounds metadata bug fix; tests 37-44 covering all 15 capabilities; 2 real LLM failures found |
| 03-22 | Persistent memory restructure + dashboard workbook | SOUL.md, journal rotation, capability matrix, run history, dashboard.workbook.json |
| 03-22 | Provider-backed pass 30-36: strict gating OK | 7/7 tests, 9/9 steps, no stray commands |
| 03-22 | Autonomy Stage 2-3: Asserts, Scoring, Reflection | Step-level assertions, scorecard, improvement suggestions |
| 03-22 | Headless Test Runner and Autonomy | CLI headless runner, structured artifacts, Stage 1 complete |
| 03-21 | Unified Thread UX + Acceptance Tests | Single threaded AI session, inline action cards |
| 03-21 | UI Panel v2 | Mode chips, preview card, Ask thread |
| 03-21 | UI Polish Pass | Theme.cs, visual overhaul, AI write flash |
| 03-21 | AI Telemetry & Chat History | Usage parsing, status line, JSONL debug logging |
| 03-20 | Agent Loop + Two-Phase + Gating | Observe->reason->act, query intents, selection fencing |
| 03-20 | Planner Policy Refactor | Explicit AllowedCommands, WritePolicy, Schema |
| 03-20 | Planner Revise Loop | Auto-revision on policy violations |
| 03-20 | Major Feature Batch | Cross-sheet refs, insert/delete rows, batch fill, 11 functions |
| 03-20 | Hebrew Roots Test + Batch Vision | MVP use case, scale analysis, structured prompt template |
| 03-19 | Codebase Audit & E2E Test Suite | 15 initial tests, Test Runner, ENHANCEMENTS audit |
| 03-19 | UI Modernization | Flat styling, modern selection colors, double buffering |

> Entries before 03-21 are in `Dev Journal Archive.md`.

---

## Restore Live Memory Doc (2026-03-22)

What changed
- Created `memory/MEMORY.md` as the live project-state doc (architecture snapshot, Current Focus, pointers, maintenance notes).
- Seeded Current Focus from recent findings: Test 41 layout stacking issue; Test 44 unit scaling ("95k" as 95000); global policy for title writes in strict mode.

Rationale
- `SOUL.md` explicitly references `memory/MEMORY.md` and the session protocol expects reading/updating it each session.
- The previous memory document was intentionally archived during the restructure; `SOUL.md` now holds durable truths while `memory/MEMORY.md` tracks session-level state.

Files
- Added: `memory/MEMORY.md`

---

## Provider-backed full suite + 8 new hard tests (2026-03-22)

What changed
- Fixed `SanitizePlanToBounds` in HeadlessTestRunner — was creating new AIPlan without copying Provider/Model/Usage/LatencyMs. This caused 30/40 steps to appear as null-provider when they were actually hitting OpenAI.
- Added provider metadata preservation to ProviderChatPlanner's fallback JSON parser (the catch block that recovers malformed LLM JSON).
- Added error logging to both catch blocks in ProviderChatPlanner so silent fallbacks are visible.
- Created 8 new tests (37-44) with real assertions covering all 15 capability categories:
  - 37: Data cleaning (name casing, email lowering) — PASS
  - 38: Complex formulas (SUM, MAX, growth %) — PASS
  - 39: Ambiguity handling (observe messy data, fix status typos) — PASS
  - 40: Analyst reasoning (analyze trends, add profit column) — PASS
  - 41: Multi-turn deep (4-turn inventory build from scratch) — FAIL step 4: layout error, labels packed into wrong shape
  - 42: Safe write protection (insert rows without breaking formulas) — PASS
  - 43: Boundary stress (sparse data, orphan cells, summary from scattered values) — PASS
  - 44: Cross-sheet reasoning (aggregate employee data into department summary) — FAIL: unit scale error ("95k" read as 95)

Key findings
- Full suite now 35 tests. Provider-backed run: 33/35 pass, 2 fail.
- The 2 failures are genuine LLM reasoning errors, not infrastructure bugs.
- Test 41 failure: gpt-4o-mini packed `["Total Items:", "Total Value:"]` into one row instead of stacking vertically. Shape/layout planning degrades in later turns.
- Test 44 failure: prompt hinted "Alice 95k" → LLM computed 95+98+102=295 instead of 95000+98000+102000=295000.

---

## Persistent Memory Restructure + Dashboard Workbook (2026-03-22)

What changed
- Restructured the entire persistent memory system for multi-agent workflows.
- Created `SOUL.md` (agent operating manual) from root MEMORY.md + distilled THERAPY.md autonomy ladder + benchmark philosophy. Any agent reads this first.
- Slimmed Dev Journal from 1314 lines to ~120 with an index; archived older entries to `Dev Journal Archive.md`.
- Created `tests/CAPABILITY_MATRIX.md` (15 capability categories with test coverage) and `tests/capability_map.json` (machine-readable twin).
- Created `tests/FAILURE_LOG.md` for tracking recurring failure patterns across runs.
- Archived raw conversation docs (THERAPY.md, AGENTIC_TESTS.md, DASHBOARD.md, old MEMORY.md) to `archive/`.
- Added session protocol to SOUL.md (start/end of session checklist).

Dashboard & run history (HeadlessTestRunner.cs)
- After each headless suite run, appends a JSON line to `tests/output/run_history.jsonl` (timestamp, commit, provider, model, pass rate, failure keys).
- Generates `tests/output/dashboard.workbook.json` — a real workbook openable in the app with 5 sheets: Overview, Run History, Capabilities, Failures, Next Actions.
- Also generates `tests/output/DASHBOARD.md` (markdown twin for agent consumption).

Design decisions
- Workbook uses PascalCase properties to match SpreadsheetIO's WorkbookData loader.
- Git commit hash captured via `git rev-parse --short HEAD` at runtime.
- Run history is append-only; dashboard is regenerated each run from full history.
- Capability map kept as separate JSON (not embedded in runner) so it can be updated independently when tests are added.

---

## Provider-backed pass 30-36: strict gating OK (2026-03-22)

Run results
- Ran `tests/TEST_SPECS_30_36.json` via headless Release EXE with `.env` in place.
- 7/7 tests passed; 9/9 steps passed; pass rate 1.000.
- No assertion failures; no `unexpected_command:set_title`; no out-of-bounds writes.

Provider
- Test 31 (values-only gating) used OpenAI `gpt-4o-mini-2024-07-18` (latency ~1.15s; total_tokens ~1576).
- Other steps were heuristic/deterministic (provider fields null), as expected for observe-only and transform primitives.

Artifacts
- `tests/output/scorecard.json` (pass_rate = 1)
- `tests/output/results_summary.json` (per-step provider/tokens where applicable)
- `tests/output/improvement_suggestions.json` (empty)

Notes
- The previously observed `set_title` stray command did not reproduce under this run.
- Next: propagate default asserts to the full suite and decide global policy for title writes in strict mode.

## Headless Test Runner and Autonomy (2026-03-22)

What changed
- Added a headless test runner to the WinForms host (no separate project required). New CLI flags:
  - `--run-tests [tests/TEST_SPECS.json] [--output-dir tests/output] [--reflection]`
  - `--run-one <tests/test_XX_....workbook.json> [--output-dir tests/output] [--reflection]`
- The runner mirrors the UI Test Runner: selection setup, planning (chat or two-phase agent), optional apply, diffs, and consolidated exports.
- It writes per-step artifacts (`*.before.workbook.json`, `*.workbook.json`, `*.plan.json`, `*.export.json`, and for agent steps `*.agent.txt` + `*.observations.json`).
- It also writes `tests/output/results_summary.json` with per-test step summaries and totals.
- Optional `--reflection` emits `*.feedback.json` per step with heuristic "missing tools/commands, comms feedback, next capabilities".

How we run it
- From WSL, we invoke the Windows EXE via PowerShell; execution occurs on Windows, artifacts land under `tests/output/` and are visible to WSL.
- If the Debug UI is open, the Debug EXE can be file-locked. Use Release for headless: `bin/Release/net8.0-windows/SpreadsheetApp.exe ...`.

Provider vs Mock
- Provider selection remains env-first (Auto). `.env` must be placed next to the running EXE (Debug or Release folder) or env vars set globally.
- If `OPENAI_API_KEY`/`ANTHROPIC_API_KEY` is present, headless runs use the real provider; otherwise they fall back to the MockChatPlanner.

Validation on first pass
- Full suite reported "27 passed, 0 failed" with Mock. Provider-backed run on tests 30-36: 6/7 passed, 8/9 steps passed (0.889 pass rate).

Gotchas fixed
- Build error (`AIContext.SheetSummary` type) — corrected to `SheetSummary`.
- Results summary writer switched to safe writer.
- Debug EXE file lock avoided by using Release build.

---

## Autonomy Stage 2-3: Asserts, Scoring, Reflection (2026-03-22)

What changed
- Extended headless runner with step-level assertions and scoring.
  - New optional fields per step: `AllowedCommands`, `ExpectNoWrites`, `ExpectTranscript`, `ExpectedCells`, `MaxChangesOutside`, `MinChangesTotal`, `Strict`.
  - Runner evaluates assertions, records `assert_failures` per step in `*.export.json`, marks each step `passed`.
- Added outputs:
  - `tests/output/scorecard.json` — totals, passed steps, pass rate.
  - `tests/output/improvement_suggestions.json` — aggregated failure keys.

Run results (provider-backed, tests 30-36)
- 6/7 tests passed; 8/9 steps passed; pass rate 0.889.
- One failure: `test_35` step 1 — `unexpected_command:set_title` (strict allowlist correctly caught stray title write).

Next steps
- Propagate default asserts across full suite.
- Decide policy for `set_title` in strict mode.
- Enrich results_summary with provider/model and token/latency.

---

## Unified Thread UX + Acceptance Tests (2026-03-21)

What changed
- Replaced mode-first panes with one unified, scrollable thread.
- Inline Apply/Revise buttons on each proposal card.
- Default selection widened from 5x1 to 20x6.
- Intent heuristics for command gating (insert/delete/copy/move/format/validation/transform/sort/explain).
- Inline ghost acceptance logs to thread for traceability.

Tests added: 35 (unified thread: create then extend), 36 (answer-only, no writes).

---

## UI Panel v2 — Mode chips, Preview card, Ask thread (2026-03-21)

What changed
- Mode chips: Fill / Append / Transform / Ask with per-mode command gating.
- Preview card: clean summary + 5x6 sample table instead of terminal log.
- Ask mode: answer-only path with visible Q/A thread, no Apply button.
- Advanced drawer: collapsed by default, holds debug/policy/history.
- Context menu consolidated: AI > Explain Cell / Schema Fill / Ask About Sheet.

---

## AI Telemetry & Chat History (2026-03-21)

What changed
- Provider metadata: usage parsing + latency in OpenAI/Anthropic providers.
- Chat status line: provider/model, latency, tokens, remaining context.
- JSONL debug logging (opt-in via `AI_DEBUG_LOG=1`).
- Action Log enriched with Model/Tokens/Latency columns.
- Chat History viewer with JSON copy/save.

---

## UI Polish Pass (2026-03-21)

What changed
- New `UI/Theme.cs` centralizing all colors, fonts, button styles.
- Chat panel: ListBox -> RichTextBox with color-coded entries.
- MainForm: formula bar 32px, flat sheet tabs with blue accent, AI write flash (500ms).
- Dialogs: FindReplace, Settings, GenerateFill all themed.
- Plan preview diff overlay (green=add, yellow=modify, red=clear).
- Collapsible policy panel, focus indicators, animated "Thinking..." indicator.
## Stage 4 infra + first experiment (2026-03-22)

What changed
- Built the Stage 4 regression gate into the headless runner: generates `regression_report.json` + `REGRESSION.md` and contributes to exit code.
- Added `tests/FAILURE_TAXONOMY.md` aligned to real assertion keys, including distinct modes: `no_effect` vs `under_generation`, and `flaky_nondeterministic` (from test_44).
- Extracted system prompt to versioned files (`prompts/system_planner_v1.md`), added file-loading with fallback, and recorded `prompt_version` in run history.
- Created `prompts/system_planner_v2.md` with a minimal categorical label normalization rule (fix obvious single-character typos like "actve"→"active").

Experiment
- Target: test_39 (`cell_mismatch:C9 expected='active' got='actve'`; also `min_changes_total_not_met:2<3`).
- Hypothesis: Explicit categorical typo-correction guidance would fix C9.
- Result: test_39 remained failing. However overall suite improved from 31/35 to 32/35 with zero regressions (gate PASS). test_42 improved (no_effect addressed). test_44 remains numerics issue and flaky.

Decisions
- Kept v2 as active (strict net improvement); blessed baseline at 32/35.
- v1 was briefly restored then reverted back to v2; final state uses v2 by default.

Next
- Consider v3 with a few-shot example to enforce categorical correction, or accept as a model limit if no gain without regressions.
