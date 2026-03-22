# Failure Log

Tracks recurring failure patterns across headless test runs. Updated after each provider-backed run. Source data: `tests/output/improvement_suggestions.json` and `tests/output/scorecard.json`.

## How to update

After a headless run:
1. Check `tests/output/scorecard.json` for pass rate.
2. Check `tests/output/improvement_suggestions.json` for new failure keys.
3. Add a row to the Run History table below.
4. If a failure key appears 3+ times across runs, promote it to the Recurring Patterns section.

---

## Run History

| Date | Spec file | Tests | Steps passed | Pass rate | New failure keys |
|------|-----------|-------|-------------|-----------|-----------------|
| 2026-03-22 | TEST_SPECS_30_36.json | 7 | 8/9 | 0.889 | `unexpected_command:set_title` |
| 2026-03-22 | TEST_SPECS_30_36.json | 7 | 9/9 | 1.000 | — |
| 2026-03-22 | TEST_SPECS.json | 27 | 40/40 | 1.000 | — (2 JSON parse recoveries; metadata bug fixed) |
| 2026-03-22 | TEST_SPECS_37_44.json | 8 | 15/17 | 0.882 | `min_changes_total_not_met`, `cell_mismatch:D2/D3/D4` |

## Recurring Patterns

*No patterns promoted yet. A failure key needs 3+ occurrences across runs to be listed here.*

## Failure Key Reference

These are the failure types tracked by the assertion system:

| Key pattern | Meaning |
|------------|---------|
| `unexpected_command:<type>` | Planner emitted a command not in AllowedCommands |
| `writes_outside_selection` | Changes detected outside the specified selection bounds |
| `missing_transcript` | Agent step produced no observation transcript |
| `unexpected_writes` | Writes occurred when ExpectNoWrites was set |
| `missing_expected_cell` | An ExpectedCells value was not found in the result |
| `insufficient_changes` | Fewer changes than MinChangesTotal |
| `excess_outside_changes` | More out-of-bounds changes than MaxChangesOutside |

## Notes

- The headless runner currently uses heuristic reflection. Provider-driven reflection (strict JSON) is a follow-up.
- Mock-planner runs are useful for structural coverage but don't test real model behavior. Provider-backed runs are the meaningful signal.
