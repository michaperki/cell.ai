# Failure Taxonomy — Writable Surface Mapping

Purpose: Provide a stable mapping from observed assertion failures to likely writable surfaces (system prompt, heuristics, test specs) and note distinct failure modes (e.g., flaky vs. consistent errors). This is a living document and should reflect real keys emitted by the headless runner.

Scope: Keys and examples align to the assertion system in `Core/HeadlessTestRunner.cs` today.

---

## Key Categories

- wrong_content
  - Keys: `cell_mismatch:<addr> expected='…' got='…'` (non-numeric or semantic mismatches)
  - Likely surface: system prompt guidance (layout/content rules) or test oracle if the expected value is wrong.
  - Notes: For formula outputs, also check `formula_error` below.

- wrong_numeric_scale
  - Keys: `cell_mismatch:<addr> expected='number' got='number'` where values differ systematically (e.g., 295 vs. 295000).
  - Likely surface: system prompt normalization/units rules; sometimes test spec if the oracle expects rounding.
  - Example: Test 44 variant (“95k” treated as 95) — expected 295000, different incorrect totals produced across runs.

- no_effect
  - Keys: `min_changes_total_not_met:0<Y` (zero writes when writes were expected)
  - Likely surface: planner gating/intent extraction or prompt clarity about destination and allowed commands.
  - Example: Test 42 — step expects some writes but total changes is zero.

- under_generation
  - Keys: `min_changes_total_not_met:X<Y` (X>0 but below minimum)
  - Likely surface: system prompt strength for fill/coverage; test spec threshold if too strict.
  - Example: Test 39 — expected 3 changes, only 2 were made.

- out_of_bounds
  - Keys: `changes_outside_exceeded:X>Y`
  - Likely surface: selection fencing (hard mode), planner range inference rules, or prompt clarity about range bounds.

- disallowed_command
  - Keys: `unexpected_command:<type>`, `unexpected_set_title`
  - Likely surface: AllowedCommands gating or planner’s tool routing; strengthen schema/constraints in the prompt.

- writes_forbidden
  - Keys: `expected_no_writes_but_found_changes`
  - Likely surface: command gating; ensure Answer/Q&A mode or empty allowlist is honored end-to-end.

- missing_transcript
  - Keys: `expected_transcript_missing`
  - Likely surface: agent observation phase not invoked; ensure two-phase loop and transcript capture are enabled for agent steps.

- formula_error
  - Keys: `cell_mismatch:<addr> … got='#ERR…'`
  - Likely surface: formula routing (use set_formula, not set_values) or prompt rules for relative references and header alignment.

- layout_violation
  - Keys: `cell_mismatch:<addr> …` but value appears in the wrong shape/position (qualitative judgment; not auto-detected yet)
  - Likely surface: prompt layout rules (e.g., vertical stacking vs single-row packing). Example: Test 41 packed labels into one row instead of stacking.

- flaky_nondeterministic
  - Signals: Failure payloads vary across runs at temperature=0 for the same test and expected value (e.g., different wrong numbers in `cell_mismatch` for the same address).
  - Detection: Compare `tests/output/run_history.jsonl` and/or `regression_report.json` across runs; if the `got` value changes while `expected` stays constant, mark as flaky.
  - Likely surface: provider/model nondeterminism; mitigation via stronger normalization/explicit constraints in the system prompt or by decomposing the task with observations.

---

## Mapping Guidance

- wrong_content → tighten system prompt content/layout rules; avoid over-specifying if the test oracle is uncertain.
- wrong_numeric_scale → add explicit unit/notation normalization rules in the system prompt (e.g., "interpret k-notation as thousands").
- no_effect → ensure destination range and commands are explicit; consider minimal fill scaffolding guidance.
- under_generation → consider lowering MinChanges thresholds in tests only if justified; otherwise strengthen planner coverage.
- out_of_bounds → reinforce selection fencing in the prompt and/or sanitize plan to selection hard bounds (already enabled in headless).
- disallowed_command → revisit AllowedCommands for the step or prune tools in the prompt via examples/negative guidance.
- writes_forbidden → ensure Ask/Q&A mode pathways never yield write commands.
- missing_transcript → agent loop must run observation phase; add explicit instruction or fallback queries.
- formula_error → route to set_formula with correct relative references; add examples for header offsets.
- layout_violation → add shape rules (vertical stacking, per-row widths) to prompt; supply explicit ranges in tests to reduce ambiguity.
- flaky_nondeterministic → consider intermediate calculations and explicit arithmetic; add normalization; if persistent, annotate test as flaky in FAILURE_LOG with notes.

---

## Examples (from recent runs)

- Test 39: `min_changes_total_not_met:2<3`, `cell_mismatch:C9 expected='active' got='actve'` → under_generation + wrong_content
- Test 42: `min_changes_total_not_met:0<4` → no_effect
- Test 44: `cell_mismatch:D2 expected='295000' got='318000'` vs later `got='240000'` → wrong_numeric_scale + flaky_nondeterministic
- Test 02: `cell_mismatch:B9 expected='=AVERAGE(B2:B7)' got=''` → formula_error (missing formula)

---

## Process Notes

- Start with document-only taxonomy (this file). No programmatic classifier yet.
- When patterns stabilize, optionally add lightweight categorization in regression reports.
- Keep failure keys consistent with the runner; if keys evolve, update this mapping.

