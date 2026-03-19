# Backlog / Refinements

This file tracks follow-ups and refinements discovered while implementing the enhancement plan. It lives next to `ENHANCEMENT_PLAN.md`.

- Incremental recalc UI integration
  - After `_sheet.RecalculateDirty(...)`, update only affected cells and invalidate them individually (`grid.InvalidateCell(r, c)`) instead of refreshing the whole grid.
  - Wrap grid updates in `grid.SuspendLayout()` / `grid.ResumeLayout()` to reduce flicker.
  - Consider handling `CellValidated` instead of `CellEndEdit`, or call `grid.CommitEdit` explicitly before recalculation to ensure committed values.
  - Heuristic: if affected set > ~5% of cells, fall back to full-sheet refresh.

- Dependency extraction robustness
  - Reuse the existing parser (AST) to collect cell references and ranges for dependency tracking instead of ad‑hoc scanning.
  - Ensure references inside string literals are ignored (already handled); cover functions, nested expressions, and ranges uniformly.

- Performance / UX
  - Consider DataGridView VirtualMode for very large sheets.
  - Batch UI updates and avoid per-cell painting where possible.

- Multi‑sheet workbook I/O
  - Add workbook Save/Open that serializes multiple sheets with names and formats.
  - For backward compatibility, still accept single‑sheet files. Offer both “Save Sheet” and “Save Workbook”.

- Async I/O adoption
  - Open/Save now uses async IO with busy cursor and disables UI. Extend to other long-running operations (CSV import/export, future workbook IO) and add guard flags to avoid reentrancy.

- Editing UX
  - Clear contents for multi-cell selections with Delete/Backspace — implement as a single bulk undo action and refresh only affected cells when incremental repaint lands.
  - Consider a dedicated Edit→Clear Contents menu item mirroring Delete.

- Undo/Redo polish
  - Coalesce rapid edits to the same cell into a single undo action.
  - Extend to multi-cell operations (copy/paste, fill, import) with grouped actions.
  - Enable/disable Edit→Undo/Redo menu items based on stack state.

- Error messaging consistency
  - Normalize messages (e.g., use `#ERR:` prefix consistently via EvaluationResult).
  - Review function argument errors for clarity and parity with spreadsheet norms.

- Testing
  - Add unit tests for: comparison operators, string literals and `&`, new functions (IF/AND/OR/NOT, string ops, math, VLOOKUP), absolute refs parsing, and dependency‑graph incremental recalc.
 - AI UX fixes
   - Generate Fill: support range selection. Use selected rectangle as the target shape and seed the dialog with its rows/cols. Requires enabling DataGridView multi‑select.
   - Inline continuation: filter out suggestions that duplicate contiguous items above the cursor (case‑insensitive, trimmed). Show only new items in the ghost panel and apply only those on accept.
   - Chat vs Fill differentiation: keep Chat for multi‑command plans; keep Generate Fill for explicit range fills. Ensure both respect write caps and record a single undo group per apply.
