# Dev Journal

This journal captures decisions, hurdles, and fixes made while implementing the enhancement plan for Rustsheet.

## What We Shipped
- Core
  - Shared defaults for rows/cols (100 x 26)
  - Division by zero returns `#ERR: DIV/0!`
  - Dependency graph (dependencies + dependents) and `RecalculateDirty`
  - Undo/Redo with per-edit recording
- Formula Engine
  - Comparison operators: `=, <>, <, >, <=, >=` (return 1/0)
  - String literals and `&` concatenation with escaping (`""`)
  - New functions: IF, AND, OR, NOT; LEN, LEFT, RIGHT, MID, CONCATENATE; ABS, ROUND, CEILING, FLOOR, MOD; VLOOKUP (exact + approximate)
  - Argument validation with clear messages
  - Absolute refs parsing (`$A$1`) accepted and stripped
- UI
  - Copy/Cut/Paste (raw values) with menu + Ctrl shortcuts
  - Status bar (Cell, Raw, Value) with live updates
  - Keyboard navigation: Enter down, Tab right (when not editing)
  - Find & Replace dialog (Ctrl+F / Ctrl+H), Replace All, wrap-around
  - Cell formatting: bold, text/fill colors, alignment, simple number format (0.00)
  - Tabs for multi-sheet (Add/Rename/Remove, per-sheet undo)
  - Column resize enabled; default width set
  - Recent files (last 5) persisted in LocalAppData
- I/O
  - JSON save/load for single sheet, now includes formats
  - CSV import/export (display values)
  - Async Save/Load methods added (UI still sync)

## Hurdles and How We Solved Them
1) Dependent cells not updating after edit
   - Symptom: After editing `A1`, cells with `=A1+1` didn’t visually update until clicked.
   - Root cause: DataGridView didn’t repaint affected cells reliably with the initial incremental update; also a risk of missing edges in ad‑hoc reference scanning.
   - Fix: Perform full-sheet recalc + redraw on edit commit. This guarantees correctness now. We kept the dependency graph in place and added backlog items to re-enable efficient incremental repaint using the parser AST.

2) Designer and class scope breakages (CS8803 top-level statements)
   - Symptom: “Top-level statements must precede namespace” and invalid modifier errors.
   - Root cause: Patching placed blocks after the closing braces of `MainForm` and `Spreadsheet`, leaving methods/fields at top level.
   - Fix: Moved all injected blocks back inside their classes. Ensured the designer’s `InitializeComponent()` order is respected.

3) Null menu item crash on startup
   - Symptom: `ArgumentNullException` in `AddRange` during `InitializeComponent`.
   - Root cause: `formatToolStripMenuItem` was added to `menuStrip1.Items` before it was instantiated.
   - Fix: Instantiate `formatToolStripMenuItem` before `AddRange`. Verified all items are created before use.

4) Recent files warning and robustness
   - Symptom: Nullable warning and potential issues when reading recent list.
   - Fix: Guarded text when populating the list; wrapped initial load in try/catch; stored under `%LocalAppData%/SpreadsheetApp/recent.txt`.

5) Multi-sheet reentrancy
   - Symptom: Risk of `SelectedIndexChanged` firing while rebuilding tabs.
   - Fix: Added `_suppressTabChange` guard in `RefreshTabs()` to prevent reentrancy loops.

6) Parser extensions and precedence
   - Considerations: `ParseComparison()` above add/sub; `ParseConcat()` lowest; lexing multi-char ops (`<=`, `<>`, `>=`) and `=` distinct from formula prefix; string literal escapes (`""`).
   - Fix: Introduced new tokens and parser levels; ensured evaluation returns 1/0 for comparisons and concatenates display strings.

7) VLOOKUP specifics
   - Implemented exact match (text or numeric) and approximate numeric search (last `<=` target). Added argument checks and range parsing.

8) Environment constraints
   - The local agent could not run `dotnet`; relied on your builds to validate. Added global exception handlers to surface runtime issues via message boxes.

## Tradeoffs and Rationale
- Recalc strategy: Full-sheet recalc on edit is simpler and reliable for 100x26. The dependency graph enables future incremental path; we’ll use AST-driven ref collection to ensure correctness.
- Cell formatting: Focused on high-value basics (bold, colors, align, `0.00`). Left advanced number formats for later.
- Multi-sheet: Implemented UI first, deferred workbook-level Save/Open for a dedicated pass (tracked in BACKLOG).
- Async I/O: Added API first, UI adoption next (busy cursor/disable menus while loading).

## Next Steps (linked to BACKLOG)
- Incremental repaint after `_sheet.RecalculateDirty(...)` with AST-driven reference collection.
- Workbook Save/Open supporting multiple named sheets in one file; maintain backward compatibility with single-sheet JSON.
- Switch UI Open/Save to async (`await`), show wait cursor, and guard repeated actions.
- Add unit tests for parser, functions, VLOOKUP, formatting, and CSV.

## AI Integration Phase (v0.1–v0.3) — What We Built, What We Learned

### Goals
- Add small, reliable slices of AI with strict output contracts and snappy UX.
- Keep inference events cheap and cancellable (INFERENCE.md principles), and never auto‑commit changes.

### v0.1 Generate Fill (explicit, preview + accept)
- Abstractions: `IInferenceProvider` with `AIContext`/`AIResult` (2D plain‑text cells), plus a `MockInferenceProvider` for offline dev.
- UI: AI > Generate Fill… (Ctrl+I) opens `GenerateFillDialog` with prompt, shape pickers, and a preview grid. Accept commits all writes as one undo step.
- Undo/Redo: Upgraded `UndoManager` to support bulk edits and later composite operations.
- Key files: Core/AI/*, UI/AI/GenerateFillDialog.cs, UI/MainForm.cs.

### v0.2 Inline Ghost Fill (debounced list continuation)
- Triggers only when the current cell is empty and there are ≥2 non‑empty cells above in the column. Debounced ~200 ms, cancelled on edit/scroll/selection change.
- Overlay: A faint panel with “Generating…” spinner, suggestions list, and Apply/Dismiss buttons. Accept with Enter/Tab/double‑click/Ctrl+Shift+I.
- Caching: Tiny key over (sheet, column, recent‑items hash, optional title, N) to avoid recomputation when navigating.
- UX fixes:
  - Kept the overlay visible while composing modifiers (Ctrl/Shift) so Ctrl+Shift+I works.
  - Preselected the first suggestion for clarity; added Esc to dismiss.
- Guardrails: Never auto‑commit; writes are selection‑bounded and a single undo group.

### v0.3 Chat Assistant (plan → preview → apply)
- Planner: Started with `MockChatPlanner`, then added `ProviderChatPlanner` that uses the selected model (OpenAI/Anthropic) to produce a strict JSON plan:
  - `set_values(start:{row,col}, values[][])`
  - `set_title(start:{row,col}, rows, cols, text)`
  - `create_sheet(name)`
- Apply: `ApplyPlan` executes a safe subset only, then records a single composite undo containing (optional) sheet add + all cell edits.
- Append semantics: If the plan asks to “continue/append”, we allow a sentinel `StartRow=-1` and resolve it to the first empty row before apply.
- Limits: Hard write caps are easy to add; for now, we size by the returned payload.

### Provider Integration (env‑first)
- `.env` loader (Program start) + Auto provider selection:
  - OPENAI_API_KEY → OpenAI Chat Completions
  - ANTHROPIC_API_KEY → Anthropic Messages
  - Otherwise → Mock (offline)
- Settings moved to provider choice (Auto/OpenAI/Anthropic/Mock/External). Keys are sourced exclusively from env; we do not collect or store API secrets in the app.
- Test Connection: Quick probe runs a 1x1 fill, reports latency + sample. Reloads `.env` on click, and shows the expected `.env` path.
- External POST provider remains available in code for custom gateways (expects `{"cells":[...]}` or `{"text":"..."}`). UI for configuring it is deferred.

### Bugs / Gotchas We Hit (and fixed)
1) Menu crash on startup (null ToolStrip item).
   - Cause: `AI` menu added to `AddRange` before instantiation.
   - Fix: Instantiate the menu first, then call `menuStrip1.Items.Add(aiToolStripMenuItem)`.

2) Inline overlay vanished when pressing Ctrl/Shift (hotkey acceptance failed).
   - Cause: Key handler treated modifier keys as invalidation triggers.
   - Fix: Ignore pure modifiers; added a form‑level `ProcessCmdKey` for Ctrl+Shift+I.

3) Planner compile error (CS0206).
   - Cause: Using `out` to assign directly into properties.
   - Fix: Assign to locals, then write into properties.

4) Path hint typo for `.env` (missing path separator).
   - Fix: Always print `Path.Combine(baseDir, ".env")`.

5) Provider readiness in inline/Generate Fill.
   - Improvement: Gate requests on `ProviderReady()`; inline schedules only when provider is usable.

### Rationale / Design Tradeoffs
- Strict output contracts (JSON or plain lines) avoid UI surprises and simplify preview.
- Minimal context snapshots + caching keep latency and cost down for both inline and explicit fills.
- Composite undo makes multi‑command chat applies atomic and reversible.
- Env‑first secrets: simpler operationally (and safer) than storing keys in app settings; later, a gateway can meter usage for end users.

### How to Validate (manual)
- Explicit fill: Ctrl+I → prompt → Preview → Accept → single Undo restores.
- Inline: seed 2+ items, pause on the next empty cell → overlay → Accept/Dismiss; revisit to see cached suggestions.
- Chat: Plan → review commands → Apply; Undo once reverts sheet add + edits together; Redo reapplies.

### Known Limitations / Future Work
- Planner schema is minimal; add `rename_sheet`, `clear_range`, header awareness, and write caps.
- Provider planning returns free‑form text sometimes; we attempt robust JSON extraction but can strengthen prompts and parsing.
- External provider UI (endpoint/key) is deferred since we pivoted to env‑first; code remains if we revive a gateway.
- Observability: basic message boxes for errors; add in‑app diagnostics (latency, error counts) later.

## Post‑Testing Fixes (Generate Fill + Inline)
- Generate Fill now respects rectangular selection:
  - Enabled grid multi‑select and detect the selected rectangle; dialog seeds Rows/Cols from that shape.
  - Preview and Accept fill the exact range; records a single bulk undo.
- Inline list continuation deduplication:
  - Previously, provider echoes (e.g., “Hammer, Nails, Screwdriver”) could be shown again in the ghost panel and then applied, causing duplicates.
  - We now trim any leading repeats of items that already exist above the cursor, then show only new items.
  - If all returned items are repeats, the ghost shows “No suggestions” and disables Apply instead of flashing and disappearing.
  - Accept writes only the new items starting at the current empty cell; existing cells above are never overwritten.

## Additional Updates (Repaint, Workbook I/O, Chat Planner)

1) Incremental repaint strategy and reliability fix
   - Change: Introduced an incremental repaint helper that updates only affected cells using `_sheet.RecalculateDirty(...)` with a 5% threshold fallback to full refresh.
   - Applied to: paste, cut, Clear Contents, Replace/Replace All, AI apply, and Undo/Redo.
   - Edit reliability: After `CellEndEdit`, we restored a full-sheet `RefreshGridValues()` to guarantee dependent cells update immediately (fixes A2→A3 scenario). Incremental-on-edit remains in backlog for later when safe.

2) Multi-cell Clear Contents + safety
   - New: Edit → Clear Contents (Delete/Backspace) clears multi-selection in a single bulk undo group.
   - Added confirmation when clearing if any selected cells contain formulas.

3) Async Open/Save in UI
   - File → Open/Save now use async IO with a wait cursor and disabled UI during operations.
   - Pattern ready to be reused for CSV and workbook operations.

4) Workbook Save/Open (multi‑sheet) — first pass
   - Added `.workbook.json` format with `formatVersion = 1`, sheet names, cells, and formats.
   - Loader auto-detects workbook vs single-sheet JSON; backward compatible.
   - UI adds “Open Workbook…” / “Save Workbook As…” (async).

5) Chat planner commands extended
   - `clear_range`: clears a rectangular region; applied with composite undo and incremental repaint.
   - `rename_sheet`: renames a target sheet by index or name; updates tabs immediately.

6) Undo/Redo menu state
   - Edit → Undo/Redo enable/disable based on stack contents; stays in sync after operations.

7) Regression fix
   - Issue: Dependent cells not repainting after edit when we first tried incremental-on-edit.
   - Fix: Reverted edits to full refresh path; kept incremental for bulk ops.
