# Stage 4 Prerequisites — Spec for Bounded Self-Improvement

**Purpose**: Before an agent can safely self-improve (modify prompts, heuristics, test specs), the system needs infrastructure that prevents regressions, classifies failures into actionable categories, tracks what changed, and records experiments. This document specs those pieces in build order.

**Context**: The autonomy ladder in SOUL.md defines Stage 4 as "can patch bounded surfaces (prompts, heuristics, tests, memory)." We are currently at Stage 2–3: the headless runner evaluates and reflects, but there is no automated comparison, no regression gate, and no structured experiment tracking. The agent that added token tracking is a good example of what Stage 4 looks like — but it happened because a human noticed the gap.

---

## 1. Regression Gate (Priority: Critical)

### Problem
The agent can run tests and see pass/fail, but cannot automatically answer: "did my change help or hurt compared to the last known-good state?" Without this, any prompt tweak is a gamble.

### Design

**Baseline snapshot** (`tests/output/baseline.json`)
- A copy of `results_summary.json` from the last blessed run.
- Created manually via CLI: `--bless-baseline` flag copies current `results_summary.json` → `baseline.json`.
- The agent should NEVER overwrite baseline without explicit human approval (Mike runs `--bless-baseline`).

**Comparison report** (`tests/output/regression_report.json`)
- Auto-generated after every headless run by comparing current `results_summary.json` against `baseline.json`.
- Schema:
```json
{
  "baseline_run": "2026-03-22T...",
  "current_run": "2026-03-22T...",
  "summary": {
    "baseline_pass_rate": 0.933,
    "current_pass_rate": 0.933,
    "delta": 0.0,
    "regressions": 0,
    "improvements": 0,
    "stable_pass": 7,
    "stable_fail": 1,
    "new_tests": 0
  },
  "tests": [
    {
      "file": "test_39_ambiguity_vague_prompt.workbook.json",
      "baseline_status": "fail",
      "current_status": "fail",
      "change": "stable_fail",
      "baseline_failures": ["min_changes_total_not_met:2<3", "cell_mismatch:C9"],
      "current_failures": ["min_changes_total_not_met:2<3", "cell_mismatch:C9"]
    },
    {
      "file": "test_41_multi_turn_deep.workbook.json",
      "baseline_status": "pass",
      "current_status": "pass",
      "change": "stable_pass"
    }
  ],
  "gate_result": "PASS"  // PASS if regressions == 0, FAIL otherwise
}
```

**Gate logic**
- `PASS`: zero regressions (no test went from pass → fail). Improvements and stable failures are informational.
- `FAIL`: any test regressed. The report lists which ones and their failure keys.
- A new test that fails on first run is NOT a regression (it has no baseline entry).

**Markdown companion** (`tests/output/REGRESSION.md`)
- Human-readable version of the report, auto-generated alongside the JSON.
- Format:
```
## Regression Report
Baseline: 2026-03-22 (7/8 pass, 93.3%)
Current:  2026-03-22 (7/8 pass, 93.3%)
Gate: PASS

### Regressions (0)
(none)

### Improvements (0)
(none)

### Stable Failures (1)
- test_39: min_changes_total_not_met, cell_mismatch:C9

### New Tests (0)
(none)
```

**Implementation location**: `Core/HeadlessTestRunner.cs`, in the post-run dashboard generation section. Add a `GenerateRegressionReport()` method called after `GenerateDashboard()`.

**CLI**:
- `--bless-baseline`: copies `results_summary.json` → `baseline.json` (human-only action).
- Regression report is always generated if `baseline.json` exists. If no baseline, skip with a note.

---

## 2. Failure Taxonomy & Surface Mapping (Priority: High)

### Problem
When the agent sees `cell_mismatch:C9 expected='active' got='actve'`, it doesn't know which writable surface to touch. Is this a system prompt issue? A test spec issue? A heuristic gap? Currently this mapping lives only in a human's head.

### Design

**Failure categories** (extend as patterns emerge):

| Category | Pattern in assert_failures | Likely surface | Example |
|----------|---------------------------|----------------|---------|
| `content_accuracy` | `cell_mismatch:*` where expected ≠ got but types match | System prompt (guidance rules) | "actve" vs "active" |
| `numeric_scale` | `cell_mismatch:*` where values differ by 10^n | System prompt (normalization rules) | 295 vs 295000 |
| `under_generation` | `min_changes_total_not_met:*` | System prompt or test spec (threshold too high) | 2 < 3 |
| `over_generation` | `max_changes_outside:*` | Sanitization heuristic or system prompt | Wrote outside selection |
| `wrong_command` | `disallowed_command:*` | Command gating heuristic | Used set_formula when only set_values allowed |
| `layout_violation` | `cell_mismatch:*` where value is in wrong position | System prompt (layout rules) | Label in wrong column |
| `formula_error` | `cell_mismatch:*` where got contains `#ERR` | Formula routing or system prompt | Broken formula reference |

**Surface mapping file** (`tests/FAILURE_TAXONOMY.md`)
- Living document. Each category maps to: description, detection heuristic, which writable surface to try first, and an example from real test history.
- The agent reads this before proposing a fix so it knows where to look.

**Programmatic classification** (optional, Phase 2)
- A method in `HeadlessTestRunner` that takes the `assert_failures` array and returns classified categories.
- Added to `regression_report.json` per-test: `"failure_categories": ["content_accuracy"]`.
- This is a nice-to-have; the taxonomy doc is the critical piece.

---

## 3. Prompt Versioning & Rollback (Priority: High)

### Problem
The system prompt is a single enormous string literal in `ProviderChatPlanner.cs:52`. There's no way to diff, snapshot, version, or rollback it. When the agent modifies it, the old version is gone.

### Design

**Extract to file**: `prompts/system_planner_v{N}.md`
- The system prompt text lives in a versioned file, loaded at runtime.
- `ProviderChatPlanner` reads from `prompts/system_planner.md` (a symlink or config pointing to the active version).
- Old versions are kept: `system_planner_v1.md`, `system_planner_v2.md`, etc.

**Version metadata** (`prompts/prompt_versions.json`):
```json
[
  {
    "version": 1,
    "file": "system_planner_v1.md",
    "created": "2026-03-22",
    "description": "Baseline: original system prompt from ProviderChatPlanner.cs",
    "author": "human"
  },
  {
    "version": 2,
    "file": "system_planner_v2.md",
    "created": "2026-03-22",
    "description": "Added k-notation normalization and two-column width rules",
    "author": "agent",
    "changes": "Added 2 rules to end of prompt",
    "baseline_pass_rate": 0.933,
    "result_pass_rate": 0.933
  }
]
```

**Runtime loading**:
- `Env.cs` or a new config reads `SYSTEM_PROMPT_PATH` (default: `prompts/system_planner.md`).
- `ProviderChatPlanner` loads the file at initialization instead of using a string literal.
- Fallback: if file not found, use the hardcoded string and log a warning.

**Rollback**: Change the active pointer back to a previous version file. The agent can do this; the regression gate validates the result.

**Run history enrichment**: `run_history.jsonl` gains a `prompt_version` field so you can correlate runs to prompt versions.

---

## 4. Structured Experiment Log (Priority: Medium)

### Problem
Dev Journal is narrative. When the agent tries something, there's no structured record of hypothesis → change → result → verdict. This makes it hard to avoid repeating failed experiments or to understand what's been tried.

### Design

**Experiment log** (`tests/experiments.jsonl`):
- Append-only, one JSON object per experiment.
- Schema:
```json
{
  "id": "exp_001",
  "timestamp": "2026-03-22T14:30:00Z",
  "author": "agent",
  "hypothesis": "Adding 'normalize k-notation to integers' to system prompt will fix test_44 numeric scale failures",
  "surface": "system_prompt",
  "target_tests": ["test_44_cross_sheet_reasoning"],
  "change_description": "Appended normalization rule to system prompt v2",
  "prompt_version_before": 1,
  "prompt_version_after": 2,
  "baseline_results": { "pass": 6, "fail": 2, "pass_rate": 0.75 },
  "experiment_results": { "pass": 7, "fail": 1, "pass_rate": 0.875 },
  "regressions": [],
  "improvements": ["test_44_cross_sheet_reasoning"],
  "verdict": "PROMOTED",
  "notes": "test_44 now passes. No regressions. Blessed as new baseline."
}
```

**Verdicts**: `PROMOTED` (change kept, baseline updated), `REVERTED` (regression or no improvement), `INCONCLUSIVE` (flaky or needs more data).

**Agent workflow** (the full Stage 4 loop):
1. Agent reads `FAILURE_TAXONOMY.md` + `regression_report.json` to pick a target failure
2. Agent reads `experiments.jsonl` to check what's been tried before for this failure
3. Agent formulates hypothesis, writes it to experiment log (status: `IN_PROGRESS`)
4. Agent makes the change (new prompt version, test spec edit, etc.)
5. Agent runs headless tests
6. Agent reads `regression_report.json` — if PASS and target improved, verdict = `PROMOTED`; if FAIL (regression), revert and verdict = `REVERTED`
7. Agent updates experiment log with results
8. If PROMOTED, agent requests human to `--bless-baseline`

**Key constraint**: Step 8 keeps the human in the loop. The agent can experiment freely but only a human blesses the new baseline. This is the safety valve.

---

## Implementation Order

```
Phase A (regression gate):
  1. Add --bless-baseline CLI flag
  2. Add GenerateRegressionReport() to HeadlessTestRunner
  3. Generate regression_report.json + REGRESSION.md after each run
  4. Bless current 37-44 results as first baseline

Phase B (failure taxonomy):
  5. Create tests/FAILURE_TAXONOMY.md with initial categories from known failures
  6. (Optional) Add programmatic classifier to regression report

Phase C (prompt versioning):
  7. Extract system prompt to prompts/system_planner_v1.md
  8. Add file-loading in ProviderChatPlanner with fallback
  9. Add prompt_version to run_history.jsonl
  10. Create prompts/prompt_versions.json

Phase D (experiment log):
  11. Create tests/experiments.jsonl schema
  12. Add "Stage 4 Prerequisites" section to SOUL.md documenting the workflow
```

---

## SOUL.md Addition (proposed)

Add after the "Writable surfaces" section:

```markdown
### Stage 4 Prerequisites
Before an agent operates in self-improvement mode, these must be in place:
1. **Regression gate**: Every run compares against a blessed baseline. Zero regressions required to keep changes.
2. **Failure taxonomy**: A living doc mapping failure patterns to writable surfaces. The agent reads this before proposing fixes.
3. **Prompt versioning**: System prompts live in versioned files, not code literals. Every run records which version was used.
4. **Experiment log**: Structured append-only log of hypothesis → change → result → verdict. The agent checks prior experiments before repeating attempts.
5. **Human-in-the-loop gate**: Only a human can bless a new baseline (`--bless-baseline`). The agent experiments freely but cannot promote its own changes as the new standard without approval.
```
