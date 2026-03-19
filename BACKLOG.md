# Backlog / Refinements

This file tracks follow-ups and refinements discovered while implementing the enhancement plan. It lives next to `ENHANCEMENT_PLAN.md`.

**Last updated: 2026-03-19**

---

## Active — AI / Chat UX

- **Chat closes on Apply (bug/UX)**
  - `ChatAssistantForm` calls `Close()` after apply (line 43), killing multi-turn iteration. The user must reopen chat for each follow-up, even though conversation history is maintained.
  - Fix: keep window open after apply, clear the input, show "Applied" confirmation in the list, let the user continue chatting.

- **Planner timeout too short (10s)**
  - `CancellationTokenSource(TimeSpan.FromSeconds(10))` in `ChatAssistantForm.DoPlanAsync`. Complex plans with large context (workbook summary + nearby + history) regularly exceed this, especially on Claude.
  - Raise to 20-30s, or make configurable via AppSettings.

- **Anthropic max_tokens=800 is low**
  - `AnthropicProvider` caps response at 800 tokens. A budget table with formulas can exceed this. OpenAI has no explicit limit, creating asymmetry.
  - Raise to at least 2048. Consider making configurable.

- **No progress indicator during planning**
  - The UI only disables the Plan button. No visual feedback that work is happening.
  - Add a "Thinking..." label or a simple progress spinner in the ChatAssistantForm.

- **No plan revision / rejection UX**
  - Only options are Apply or Close. No way to tell the AI "change the formulas but keep the values" or edit individual commands before applying.
  - Add a "Revise" button that appends user feedback to history and re-plans.

- **`set_values` doesn't auto-detect formulas**
  - If the AI puts `=SUM(...)` in a `set_values` command, it's written as literal text (not a formula). This is a common AI mistake.
  - In the apply logic, auto-route values starting with `=` to the formula write path as a safety net.

- **No AI error feedback loop**
  - If a formula the AI generates evaluates to `#ERR`, there's no path back. After apply, scan for new `#ERR` cells and auto-prompt: "These cells errored: [list]. Fix?"
  - Lighter lift than the full "error repair" enhancement — directly tied to the AI apply flow.

- **MockChatPlanner doesn't cover full command set**
  - Keyword heuristics only generate `set_values`, `create_sheet`, `set_title`. Can't produce `set_formula`, `sort_range`, `clear_range`, `rename_sheet` plans.
  - Expand for better offline development and testing.

---

## Active — Infrastructure

- Incremental recalc UI integration
  - Use incremental repaint for bulk ops (paste, clear, replace, AI apply, undo/redo). For direct edits, keep full refresh for now due to DataGridView repaint quirks after commit.
  - Explore `CellValidated` vs `CellEndEdit`, or explicit `CommitEdit` before recalculation to enable safe incremental edit repaint.
  - Wrap updates in `SuspendLayout`/`ResumeLayout`; keep 5% threshold fallback.

- Dependency extraction robustness
  - Reuse the existing parser (AST) to collect cell references and ranges for dependency tracking instead of ad-hoc scanning.
  - Ensure references inside string literals are ignored (already handled); cover functions, nested expressions, and ranges uniformly.

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

- Workbook summary header detection
  - Currently sends row 1 as headers to the AI. If the actual header row is row 2 (common when row 1 is a title), the AI gets wrong context.
  - Implement auto-detect header heuristic (first non-empty row with predominantly text values).

---

## Active — I/O

- Multi-sheet workbook I/O
  - Add workbook Save/Open that serializes multiple sheets with names and formats.
  - For backward compatibility, still accept single-sheet files. Offer both "Save Sheet" and "Save Workbook".

- Async I/O adoption
  - Open/Save now uses async IO with busy cursor and disables UI. Extend to other long-running operations (CSV import/export, future workbook IO) and add guard flags to avoid reentrancy.

---

## Active — Editing UX

- Clear contents for multi-cell selections with Delete/Backspace — implement as a single bulk undo action and refresh only affected cells when incremental repaint lands.
- Consider a dedicated Edit > Clear Contents menu item mirroring Delete.

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
