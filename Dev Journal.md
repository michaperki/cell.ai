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

## Post‑Update (Undo UX + Menu State)

1) Undo coalescing for rapid edits
   - Problem: Multiple quick changes to the same cell (e.g., paste/replace bursts) created many tiny undo steps.
   - Change: Added a 1s coalescing window in `UndoManager.RecordSet`. Rapid edits to the same cell merge into a single undo record, preserving the first old value and last new value. Redo stack clears as before. Non‑single operations reset the merge window.

2) Undo/Redo menu state sync
    - Problem: Edit → Undo/Redo enable state could lag until another UI event.
    
## Test Suite Automation + Multi‑turn Chat Learnings (2026‑03‑19)

1) Removed in‑sheet instructions; added automated specs
   - Change: Eliminated A1 instruction text from all `tests/test_*.workbook.json` files to avoid contaminating AI context.
   - Added `tests/TEST_SPECS.json` that defines per‑test AI chat steps (prompt, target sheet, and selection/range).
   - Test Runner now executes these steps via “Run Steps”, shows a live log, and can save snapshots to `tests/output/` after each step.

2) Chat dialog UX improvements (applied earlier in this pass)
   - Apply no longer closes the chat window; clears input and shows an “Applied …” summary. Added a lightweight “Thinking…” indicator and extended planning timeout to 30s.

3) Multi‑turn (Test 06) observations from snapshots
   - Symptom: Step 1 placed a repeated “Expense Table” across A3:C3, omitted the Amount header/data, and wrote totals prematurely. Step 2/3 produced formulas in the wrong columns (e.g., `=B4*0.1` in C4, then `=C4*0.1` in D4).
   - Likely causes:
     - Selection shape and start were not strongly conveyed; context had `Rows=5, Cols=1` regardless of selection, biasing single‑column plans.
     - System prompt allowed “extras” (titles/totals) not explicitly requested.
     - Header awareness is minimal; provider inferred columns loosely.
   - Fixes applied:
     - BuildPlannerContext now sets StartRow/StartCol to the selection’s top‑left and Rows/Cols to the selection shape. Nearby window also anchors at this start.
     - Provider system prompt tightened: “Only perform requested changes; do not add titles/totals unless asked; align writes to the indicated shape.”
     - Updated `TEST_SPECS.json` for Test 06 to anchor steps with explicit ranges: Step 1 `A3:C6`, Step 2 `D3:D6`, Step 3 `A6:E6`.

5) Plan, prompt, and system debug dumps
   - Added a Test Runner toggle “Dump plan JSON” that saves the raw provider plan (or a serialized fallback) to `tests/output/<test>_stepN.plan.json` alongside workbook snapshots. This aids post‑mortem analysis when model output doesn’t align with expectations.
   - Added “Dump user prompt” to persist the constructed user prompt string (selection + nearby + workbook summary + instruction) per step to `tests/output/<test>_stepN.user.txt` for reproducibility.
   - Added “Dump system prompt” to persist the system schema/rules sent to the planner to `tests/output/<test>_stepN.system.txt` for transparency.

6) Before/after cell diffs in Test Runner
   - The runner now logs a concise diff of cell changes on the active sheet after each applied step (added/changed/cleared), capped at 100 entries per step for readability.

4) Next steps
   - Add optional plan JSON dump to `tests/output/` for deeper debugging.
   - Enhance header detection in the workbook summary (identify actual header row when row 1 is a title).
   - Expand MockChatPlanner to cover `set_formula`, `sort_range`, `clear_range`, `rename_sheet` for offline parity.
   - Change: Centralized calls to update menu state after edits, bulk applies, clear contents, replace, AI applies, and after Undo/Redo itself. Also refreshes on sheet (re)initialization and on Edit menu opening.

Rationale: Smoother undo UX reduces noise and accidental multi‑undo clicks; keeping menu state in sync improves perceived reliability.

3) Multi‑cell clipboard (Copy/Paste/Cut)
   - Added rectangular selection support for Copy and Cut; clipboard format is TSV (rows by newlines, columns by tabs) with minimal quoting for tabs/newlines.
   - Paste detects TSV and writes a 2D block starting at the top‑left of the current selection; records a single bulk undo and uses incremental repaint.

## AI Enhancements (Context, Commands, Chat History)

1) New planner commands: set_formula and sort_range
   - set_formula: Distinct from set_values; writes raw formulas (kept as strings with a leading `=`). Schema: {"type":"set_formula","start":{row, col},"formulas":[["=..."],...]}. Applied as a bulk write with composite undo.
   - sort_range: Sorts a rectangular region by a specified column. Schema: {"type":"sort_range","start":{row, col},"rows":N,"cols":N,"sort_col":"B"|2,"order":"asc|desc","has_header":true|false}.
     - sort_col accepts a column letter (absolute) or a 1-based index relative to the start column.
     - Sorting compares numerically when both sides are numbers; otherwise case-insensitive text; blanks last. Header row (when has_header) is preserved.
     - Formats are not moved with rows in this pass; we only rewrite raw values to keep undo semantics simple.

2) Context enrichment for planning
   - AIContext now carries optional SelectionValues (up to 40x10), NearbyValues (20x10 window to the right/down of the cursor), and Workbook summaries (sheet names, used rows/cols, first-row headers).
   - ProviderChatPlanner includes a compact inline representation of this context in the user message, keeping the existing strict-JSON system prompt and write-cap guidance.

3) Conversation history in Chat Assistant
   - The chat dialog maintains an in-memory rolling history (last ~5 exchanges). History is sent with each plan request to enable multi-turn prompts.
   - Added a "Reset History" button to clear context quickly.
   - Preview shows command summaries and counts both value and formula writes.

4) Planner schema/prompt updates
   - Strict schema extended to include set_formula and sort_range. JSON parsing covers both, with sort_col supporting either column letters or relative indices.

Rationale / Notes
- We avoided extending MockChatPlanner further; the provider-backed planner is the main path. Mock remains for offline dev but does not emit new commands.
- Sorting does not adjust external references; formulas moved within the block keep their relative addresses. Cross-range formulas elsewhere may change results; we will consider reference-safe structural ops later.
- Context strings are bounded to keep token cost reasonable. If we need more fidelity later, we can switch to a structured JSON context pack.

4) Number formats extended
   - Added presets: 0, 0.00, #,##0, #,##0.00, 0%, 0.00%, $#,##0, $#,##0.00.
   - Implemented formatting in `GetDisplayWithFormat` for numeric cells; errors/text unchanged. Menu updated under Format → Number Format.

## Codebase Audit & E2E Test Suite (2026-03-19)

### ENHANCEMENTS.md Audit

Performed a comprehensive audit of ENHANCEMENTS.md against the actual codebase. Key findings:

- **AI Command Grammar:** 2 of 8 implemented (set_formula, sort_range). Remaining 6 (insert/delete rows/cols, set_format, move_range, copy_range, delete_sheet) are not started.
- **AI Context Enrichment:** 4 of 4 complete. SelectionValues, NearbyValues, Workbook summary, and conversation history all shipped and wired through ProviderChatPlanner.
- **New AI Interaction Modes:** ~0.5 of 5. Selection content is sent (enabling transforms) but no dedicated UX. Inline formula help, explain cell, error repair, and undo-aware re-planning are not started.
- **Formula Engine Gaps:** ~4.5 of 10 functions present. HLOOKUP, INDEX, MATCH, UPPER, LOWER, SUBSTITUTE are implemented. SUMIF/COUNTIF/AVERAGEIF, IFERROR, TRIM, PROPER, TEXT, TODAY/NOW, COUNTA, ISBLANK/ISNUMBER/ISTEXT, REPLACE are missing.
- **All other sections** (data validation, cross-sheet refs, observability, QoL) are not started.
- **Overall completion: ~25-30%**, with the highest-priority items (the original top-4 recommendations) fully shipped.

Updated ENHANCEMENTS.md with per-item status annotations, a new "Known Issues & Quick Wins" section, and revised priority recommendations.

### New Issues Discovered

Found 8 actionable issues during the codebase review, added to BACKLOG.md:

1. **Planner timeout too short (10s)** — complex plans with rich context time out regularly. Needs 20-30s.
2. **Anthropic max_tokens=800** — budget tables exceed this; OpenAI has no cap, creating asymmetry. Needs 2048+.
3. **Chat closes on Apply** — `ChatAssistantForm` calls `Close()` after apply, breaking multi-turn iteration. Should stay open.
4. **No plan revision UX** — only Apply or Close, no way to give feedback and re-plan.
5. **set_values doesn't auto-detect formulas** — values starting with `=` written as literal text. Should auto-route to formula path.
6. **No progress indicator during planning** — just a disabled button, no spinner or "Thinking..." label.
7. **Workbook summary sends row 1 as headers** — wrong when row 1 is a title and row 2 has actual headers.
8. **MockChatPlanner doesn't cover full command set** — can't generate set_formula, sort_range, clear_range, or rename_sheet plans offline.

### E2E Test Suite Created

Created 15 `.workbook.json` test files in `tests/`, each with cell A1 containing step-by-step testing instructions:

| Tests 01-07 | AI capabilities: set_values, set_formula, sort_range, create/rename sheet, clear_range, multi-turn chat, context awareness |
| Test 08 | Formula engine: 12 scenarios covering math, DIV/0, IF, strings, VLOOKUP, comparisons |
| Tests 09-13 | UX: undo/redo (single/bulk/composite), workbook I/O, find & replace, copy/paste/cut, cell formatting |
| Tests 14-15 | AI advanced: complex multi-command plans, cross-sheet workbook context |

Also created `tests/TEST_INDEX.md` as a reference index.

### Test Runner UI

Added a Test Runner form (`UI/TestRunnerForm.cs`) accessible via **Test > Test Runner** menu:

- Discovers all `test_*.workbook.json` files in the `tests/` directory automatically (walks up from bin to project root).
- Lists tests in sorted order with Prev/Next navigation (buttons + arrow keys).
- Shows the A1 instruction text as a preview before loading.
- "Load Test" button (or Enter) loads the selected workbook into the main spreadsheet.
- Added `LoadWorkbookFromPath(string path)` public method to MainForm for programmatic workbook loading.
- Added Test menu with Test Runner entry to `MainForm.Designer.cs`.

### Documents Updated

- **ENHANCEMENTS.md** — Full status annotations on every item, new sections for known issues and updated priorities, test suite reference.
- **BACKLOG.md** — Restructured with new "Active — AI / Chat UX" section containing the 8 discovered issues. Cleaned up existing items.
- **Roadmap.md** — Added E2E test suite reference in Quality & Verification section.
- **tests/TEST_INDEX.md** — New file documenting all 15 test workbooks with usage instructions.
