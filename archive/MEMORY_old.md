# Memory

## Operating constraints
- The app runs on Windows; the agent runs in WSL.
- The agent may run the Windows binary in headless mode (no UI) to execute tests and collect artifacts.
- Do not attempt to launch the desktop UI from WSL; use the headless CLI instead.
- Do not claim something was tested unless the headless outputs exist or Mike explicitly tested it and reported results.
- When debugging runtime behavior, rely on Mike’s screenshots, logs, and descriptions in addition to headless artifacts.

## Collaboration model
- Mike is product lead and tester.
- The agent should implement, inspect code, and reason from artifacts.
- The agent should not invent verification.
- The agent should not repeatedly ask to reproduce bugs if Mike already described the issue clearly.
- Prefer direct code inspection and targeted fixes over vague debugging advice.

## Product vision
- The spreadsheet AI should feel like one unified AI collaborator, not four separate tools.
- Planning/execution may exist internally, but the user-facing experience should feel like one continuous session.
- Answers, proposed changes, and applied changes should appear inline in a unified thread.
- Avoid fragmented mode-first UX unless absolutely necessary.
- Visible context matters: the user should understand what the AI is using and what it is about to change.

## UX principles
- Reduce clutter.
- Hide internals by default.
- Prefer user language over internal/dev language.
- Do not expose implementation details unless they help the user make a decision.
- The main pane should optimize for clarity, trust, and flow.
- Distinguish clearly between:
  - discussion/answer
  - proposed spreadsheet action
  - applied result

## Agent behavior expectations
- Solve the root UX/product issue, not only the nearest symptom.
- Before implementing, restate the higher-level interaction model being targeted.
- Evaluate proposed changes against the product vision, not just local correctness.
- Avoid getting stuck in narrow mode-specific patches when the real issue is architectural.
- If there is uncertainty, surface tradeoffs explicitly.

## Spreadsheet-specific assumptions
- Safe writes matter.
- Selection boundaries and write constraints should be visible and predictable.
- Follow-up requests should account for prior sheet state and prior conversation context.
- The agent should infer natural continuations when reasonable (for example, adding an adjacent column to an existing generated table).

## Communication rules
- Be explicit about assumptions.
- Do not say something “works” unless it has actually been verified by Mike.
- When reporting progress, separate:
  - what was changed
  - what is inferred
  - what still needs user validation


* We are **not using CI**.
* Do not suggest, set up, modify, debug, or rely on CI workflows.
* Do not propose GitHub Actions, pipeline fixes, or CI-based validation.
* Validation happens through Mike’s local Windows testing and the headless runner (local), not CI.

- Periodically update `Dev Journal.md` to track what we tried, what failed, what changed, and what we learned.
- Important implementation attempts and product lessons should not live only in chat.
- `Memory.md` stores durable truths; `Dev Journal.md` stores the chronological record of the work.

## Headless Test Runner (infra)
- We have a built-in headless runner in the WinForms host (Program CLI flags) that executes E2E tests without launching the UI.
- Commands (run from the Windows machine):
  - `--run-tests [tests/TEST_SPECS.json] [--output-dir tests/output] [--reflection]`
  - `--run-one <tests/test_XX_....workbook.json> [--output-dir tests/output] [--reflection]`
- Outputs (under the chosen `--output-dir`):
  - Per step: `*.before.workbook.json`, `*.workbook.json`, `*.plan.json`, `*.export.json`, and for agent steps `*.agent.txt` + `*.observations.json`.
  - Aggregated: `results_summary.json` with per-test step summaries and totals.
  - Optional dev reflection per step: `*.feedback.json` (heuristic suggestions for missing tools/commands and next capabilities).
- Provider selection remains env-first (Auto):
  - Place `.env` next to the running EXE (e.g., `bin/Debug/net8.0-windows/.env` or `bin/Release/net8.0-windows/.env`) or set environment variables.
  - If `OPENAI_API_KEY` or `ANTHROPIC_API_KEY` is present, the headless runner uses the real provider; otherwise it falls back to Mock.
- Gotcha: If the Debug UI is open, the Debug EXE may be file-locked. Use the Release build for headless runs (`bin/Release/net8.0-windows/SpreadsheetApp.exe`).


## Autonomy and self-improvement direction
- The long-term goal is not just task execution, but a bounded improvement loop.
- Build/run/test capability is a prerequisite primitive, not the final objective.
- The target progression is:
  1. execute tasks
  2. evaluate outcomes
  3. explain failures/successes
  4. propose improvements
  5. rerun and compare results
- The agent should think in terms of repeated experimental loops, not one-off task completion.
- More tools or more autonomy do not by themselves equal self-improvement; improvement requires evaluation, comparison, and learning across runs.

## Task loop vs meta-improvement loop
- Separate the task-execution loop from the meta-improvement loop.
- The task loop solves the immediate spreadsheet/user problem.
- The meta-improvement loop reviews artifacts from runs and proposes improvements to prompts, tool usage, tests, heuristics, or memory.
- The agent should explicitly identify which loop it is operating in.

## Structured artifacts expectation
- Meaningful runs should leave structured artifacts, not only chat output.
- Useful artifacts include:
  - task spec
  - context snapshot
  - plan
  - actions taken
  - result diff
  - evaluator judgment
  - reflection / lessons
  - candidate improvement
- Reflections should be grounded in artifacts, not vague intuition.

## Evaluation and benchmarks
- The system should be improved against stable benchmark tasks where possible.
- Do not treat isolated successes as proof of improvement.
- Prefer changes that improve repeatability, reduce human correction, and reduce regressions.
- When proposing a system change, explain how it would be evaluated.

## Writable surfaces for bounded self-improvement
- Self-improvement requires clearly defined writable surfaces.
- Writable surfaces may include:
  - prompts/instructions
  - planner heuristics
  - test cases
  - evaluation rubrics
  - memory summaries
  - dev journal entries
- Do not implicitly treat the whole codebase as freely self-modifiable.
- Be explicit about what may be changed automatically, what requires Mike review, and what is read-only.

## Reflection quality bar
- Reflection should identify:
  - what failed or succeeded
  - why
  - what change is proposed
  - how the change would be tested
- Avoid generic reflections that do not cash out into a concrete next experiment.

## Benchmark and regression mindset
- The agent should prefer changes that can be tested against existing headless tests or stable local workflows.
- Any claimed improvement should ideally be checked against prior artifacts or benchmark tasks.
- Avoid churn that changes behavior without improving measurably.

## Dev Journal and Memory roles
- Memory.md stores durable operating truths, product truths, and agent behavior constraints.
- Dev Journal.md stores chronological experiments, failures, fixes, and lessons.
- Durable lessons that repeatedly shape decisions should be promoted from Dev Journal.md into Memory.md.


# MISC

- Before proposing infrastructure work, explain which autonomy stage it unlocks: execution, evaluation, reflection, improvement, or promotion.

- Do not mistake local optimization of a tool or command for progress toward an improving system unless it clearly strengthens the experimental loop.
