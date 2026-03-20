# Backlog / Refinements

This file tracks follow-ups and refinements discovered while implementing the enhancement plan. It lives next to `ENHANCEMENT_PLAN.md`.

**Last updated: 2026-03-20**

---

## Active — AI / Chat UX

- **Chat closes on Apply (bug/UX) — DONE**
  - Chat now stays open after Apply; input clears and an “Applied …” summary is appended so users can continue multi‑turn flows.

- **Planner timeout too short (10s) — ADDRESSED**
  - Planning timeout increased to 30s. Consider making configurable if needed per‑provider.

- **Anthropic max_tokens config — DONE**
  - Added `ANTHROPIC_MAX_TOKENS` (default 2048) for planner/provider.

- **No progress indicator during planning — DONE**
  - Added a lightweight “Thinking…” status label in ChatAssistantForm during planning.

- **Plan revision UX — DONE**
  - Revise button appends feedback and replans; keeps current plan visible until replaced.

- **`set_values` auto-detect formulas — DONE**
  - Values starting with `=` are written as formulas and evaluated correctly.

- **AI error feedback loop — DONE**
  - After apply, we prompt if errors are detected and open Chat prefilled to attempt a repair (no auto-apply).

- **MockChatPlanner coverage — PARTIAL DONE**
  - Expanded to produce `set_formula`, `sort_range`, `clear_range`, `rename_sheet`, plus heuristics for expense tables, tax columns, and bonus columns. Further tuning still welcome.

---

## Active — Infrastructure

- Incremental recalc UI integration — DONE (gated)
  - Direct edits use `_sheet.RecalculateDirty` + thresholded `RefreshDirtyOrFull` with full fallback.

- Dependency extraction robustness — DONE
  - Uses `FormulaEngine.EnumerateReferences` (AST) for references and ranges.

- Performance / UX
  - Consider DataGridView VirtualMode for very large sheets.
  - Batch UI updates and avoid per-cell painting where possible.

- UI modernization (visual polish, zero perf cost) — DONE (first pass)
  - Grid: flat borders (`BorderStyle.None`), single-horizontal cell borders, subtle alternating row colors.
  - Headers: disabled `EnableHeadersVisualStyles`, flat custom-colored column/row headers.
  - Selection: modern accent color (`Color.FromArgb(200, 220, 240)`) instead of default royal blue, dark text on selection.
  - Double-buffered DataGridView via reflection to eliminate flicker.
  - Default cell font set to Segoe UI 9pt for consistency.
  - Future: menu icons, dark mode toggle, owner-drawn tabs.

- Workbook summary header detection — DONE
  - Heuristic picks the first non-empty text-dominant row.

---

## Active — I/O

- Multi-sheet workbook I/O
  - Add workbook Save/Open that serializes multiple sheets with names and formats.
  - For backward compatibility, still accept single-sheet files. Offer both "Save Sheet" and "Save Workbook".

- Async I/O adoption — DONE (CSV)
  - Import/Export CSV now async with busy guards; pattern matches Open/Save async.

---

## Active — Editing UX

- Clear contents for multi-cell selections — DONE
  - Delete/Backspace clears selection with a single bulk undo action; guarded prompt for formulas; incremental repaint in place.

---

## Active — Undo/Redo Polish

- Coalesce rapid edits to the same cell into a single undo action.
- Extend to multi-cell operations (copy/paste, fill, import) with grouped actions.
- Enable/disable Edit > Undo/Redo menu items based on stack state. — DONE

---

## Active — Error Messaging

- Normalize messages (e.g., use `#ERR:` prefix consistently via EvaluationResult).
- Review function argument errors for clarity and parity with spreadsheet norms.

---

## Active — Testing

- Add unit tests for: comparison operators, string literals and `&`, new functions (IF/AND/OR/NOT, string ops, math, VLOOKUP), absolute refs parsing, and dependency-graph incremental recalc.
- E2E test suite: 15 workbook test files in `tests/` with Test Runner UI (Test menu). See `tests/TEST_INDEX.md`.

---

## Active — AI UX (Generate Fill / Inline)

- Generate Fill: support range selection. Use selected rectangle as the target shape and seed the dialog with its rows/cols. Requires enabling DataGridView multi-select.
- Inline continuation: filter out suggestions that duplicate contiguous items above the cursor (case-insensitive, trimmed). Show only new items in the ghost panel and apply only those on accept.
- Chat vs Fill differentiation: keep Chat for multi-command plans; keep Generate Fill for explicit range fills. Ensure both respect write caps and record a single undo group per apply.
