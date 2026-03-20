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

- UI — Docked Chat Pane
  - Added a right-side docked Chat assistant panel (Plan/Revise/Apply, rolling history). Toggle via AI > Toggle Chat or Ctrl+Shift+C. This is now the single chat surface; the former pop-out window was removed.

## AI Command Set Extensions (2026‑03‑20)

- Added new planner/execute commands (OpenAI path):
  - insert_cols, delete_cols — shift data and formats right/left with undo
  - delete_sheet — remove a named or indexed sheet safely (guard against last sheet)
  - copy_range, move_range — copy/move rectangular regions with formula reference rewrite and format preservation
  - set_format — apply bold/align/number format and colors to a range
  - set_validation (MVP) — list and number-between rules; enforced on CellEndEdit and AI applies
  - set_conditional_format (MVP) — apply formatting immediately when cell value meets a numeric threshold

- Planner updates
  - System JSON schema extended to include all commands above; strict parsing and AllowedCommands filtering maintained.
  - Hex color parsing for `#RRGGBB` and `#AARRGGBB` inputs.

- Tests added (E2E)
  - 23: insert/delete columns
  - 24: delete sheet
  - 25: copy/move range with formula rewrite
  - 26: set_format (bold/align/fill, number format)
  - 27: set_validation (list/number-between)
  - 28: set_conditional_format (threshold-based apply)

Notes
- Conditional formatting is applied immediately (not a persistent live rule engine yet).
- Validation prevents invalid user edits and AI writes; invalid cells are rejected with an explanatory message.

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

## Post‑Testing Fixes (AI Planner + Selection Sanitization)

1) Values-only enforcement (Test 16: Hebrew roots)
   - Symptom: Planner emitted a `set_title` despite the instruction “Use set_values only”; headers were duplicated into row 2.
   - Root cause: The system grammar advertised `set_title`, and we passed a “Title=” hint derived from the header cell above the selection, nudging the model toward title writes.
   - Fix: Add values-only gating. When the prompt contains “set_values only” or “do not add titles,” we (a) strengthen the planner instruction, (b) filter out disallowed command types (e.g., `set_title`, `set_formula`) from the returned plan before apply, and (c) de-bias the context by clearing the Title hint.

2) Ragged `set_values` shape within a selection
   - Symptom: A plan contained mixed-width rows (first rows only in column A, later rows spanning A..I). Sanitization used the first row’s width, so only A7:A9 were written and the multi-column data was dropped.
   - Fix: Sanitize by intersecting each row against the selection, compact rows with zero overlap (drop them), and build rows aligned to the selection’s left edge. Apply path now iterates per-row width instead of assuming a rectangular block.

3) Header-echo suppression
   - Symptom: Model repeated headers as the first row of data under the real header row.
   - Fix: When the first planned row matches the header cells above the selection (majority match), drop that row before apply.

Rationale
- Keeps AI writes inside the user’s intent and selection shape, prevents accidental title/format writes, and ensures multi-turn appends align correctly without losing the wide rows.

Validation
- Re-ran the E2E steps for Test 16 using the Test Runner: row 2 contains actual data (no header dup), and Step 2 fills B7:I9 for the three new roots while A7:A9 are populated once.

## Design Note — Observation‑First Agent Model (2026‑03‑20)

Summary
- We identified a structural gap in the current AI path: it is primarily a single‑shot planner that writes after a curated snapshot. For tasks like data cleanup, the agent needs to look around iteratively before deciding to act.

What’s missing
- Read/query tools as first‑class commands: the agent should be able to ask for uniques in a column, counts of blanks, a sample of rows, or a quick column profile. These are cheap, safe, and unblock understanding.
- A small observe→reason→act loop: run a few read steps autonomously, then propose a write plan for approval. The human sets the intent; the agent sequences the micro‑steps.

Decisions
- Add a read/query command set to the grammar: `get_range`, `sample_rows`, `unique_values`, `count_where`, `describe_column`, `selection_summary`.
- Implement a host‑controlled Agent Loop MVP (max 3 turns). Earlier steps may contain only read commands; the final step may contain writes subject to AllowedCommands/WritePolicy/selection bounds. All writes apply under a single composite undo.
- De‑prioritize single‑shot batch optimizations until the loop ships. The loop is the product; batching is a later optimization.

Planned acceptance (MVP)
- On the NJ High Schools dataset, the agent can: (1) compute unique city names and blank counts, (2) sample rows to infer normalization rules, (3) propose a values/formula plan to clean the column, (4) apply after user approval.

Implementation notes
- Queries return compact, bounded results (top‑K uniques, first N rows, aggregates) to keep tokens low. Full results are available via “copy all” in the transcript.
- Keep provider‑agnostic: initial loop can be host‑driven (planner emits query intents as JSON), with a later move to proper tool‑calling if desired.

Backlog/Roadmap updates
- New sections added for “AI Observation Tools” and “Agent Loop (MVP)”; Structured Batch Fill orchestration/costing marked Paused pending v0.3.

## Agent Loop MVP — City Normalization (2026‑03‑20)

What we shipped
- Host‑controlled agent loop that performs read‑only observations (selection summary, column profile, uniques, sample rows) and then proposes a plan. Writes occur only on Apply.
- Chat pane toggle: “Use Agent Loop (MVP)” augments prompts with observations and surfaces an observations transcript above the plan.
- Test Runner support: new `ai_agent` action to run the loop in automation, export the observations transcript, plan JSON, and workbook snapshots.
- Safety: values‑only/no‑titles filtering applied in agent automation when requested by the prompt; selection‑bounds sanitation before apply.

Validation
- Added test file `tests/test_29_ai_agent_city_cleanup.workbook.json` (NJ_HS sheet, messy City cases) and spec in `tests/TEST_SPECS.json`:
  - Step 1: `ai_agent`, apply=false — records observations and proposed plan (no writes).
  - Step 2: `ai_agent`, apply=true — applies the proposed `set_values` within B2:B12.
- Result: City names normalized (trim + Title Case) while other columns remain unchanged. Artifacts include plan JSON, agent transcript, and consolidated export.

Notes / next
- Current loop is host‑driven (no formal query tool‑calls in the provider schema yet). This proved the product loop; next iterations can expose queries as first‑class planner tools.
- Extend observation set (regex counts, histograms) and add simple gating for single‑text‑column normalizations to bias AllowedCommands to `set_values`.

## Planner Policy Refactor + Test 16 Learnings (2026‑03‑20)

What we saw
- Plans drifted from schema and policy:
  - Step 1 wrote 9 columns starting at B and used Hebrew in B (transliteration column), overrunning B..I. Cropping hid the mismatch.
  - Step 2’s user context forbade writing to A globally even though the instruction asked to add inputs in A7:A9.

Root cause
- Policies were implied in prose (heuristics scanning the prompt) instead of being explicit, typed constraints passed to the planner and enforced post‑parse.

Changes
- Added explicit AllowedCommands, WritePolicy, and Schema on AIContext.
  - AllowedCommands gates planner at both prompt and post‑parse filtering (e.g., set_values only).
  - WritePolicy carries writable columns and nuanced input‑column rules (e.g., “A read‑only for existing rows, allowed for empty rows in selection”).
  - Schema summarizes column headers and expected per‑row width bound to the selection.
- ProviderChatPlanner now renders a compact policy+schema section in the user message and filters disallowed commands by type.
- Test Runner gained structural assertions: after apply, verify that changed cells are inside the requested selection. This catches silent crops/overwrites and header edits.

Rationale
- Make constraints machine‑readable for the planner, visible for debugging (dumped to tests/output), and enforceable without domain‑specific tuning.

Validation
- For Test 16:
  - Step 1: policy expresses “Do not write to A; write only to B..I; exactly 8 values per row.”
  - Step 2: policy allows writes to A for empty rows 7..9 while keeping prior rows read‑only. Plan stays within A7:I9.

Follow-ups queued
- Planner revise loop: when the plan width/policy mismatches, request a repair instead of only cropping. Logged in BACKLOG. — Implemented below.
- Chat UI schema/policy preview: surface AllowedCommands, writable columns, and input-column rules pre-apply. Logged in ROADMAP/BACKLOG.
- Typed schema (optional): enable simple content-class hints (e.g., transliteration ASCII) to guide providers without overfitting to domains. Logged in BACKLOG.

## Post‑Update (Undo UX + Menu State)

1) Undo coalescing for rapid edits
   - Problem: Multiple quick changes to the same cell (e.g., paste/replace bursts) created many tiny undo steps.
   - Change: Added a 1s coalescing window in `UndoManager.RecordSet`. Rapid edits to the same cell merge into a single undo record, preserving the first old value and last new value. Redo stack clears as before. Non‑single operations reset the merge window.

2) Undo/Redo menu state sync
    - Problem: Edit → Undo/Redo enable state could lag until another UI event.

## Test Suite Automation + Multi‑turn Chat Learnings (2026‑03‑19)

## Planner Revise Loop (2026‑03‑20)

What we added
- ProviderChatPlanner now validates returned plans against the active selection bounds, AllowedCommands, WritePolicy (writable columns and input‑column rules), and expected per‑row width (selection width or schema length).
- On violations (e.g., writes outside selection, non‑writable columns, disallowed command types, ragged row widths), it issues a single automatic “revision” request to the provider. The revision message includes:
  - A compact constraints summary (bounds, allowed commands, writable columns, input‑column rules, expected width)
  - A short list of the first few detected problems
  - The prior plan JSON for grounding
- The corrected plan (if returned) is parsed and filtered again before being surfaced to the UI/Test Runner. If the provider cannot repair, we fall back to the original plan (the Test Runner still sanitizes to selection bounds during automation).

Why
- Reduces silent no‑ops and crop-induced surprises seen in tests (e.g., step‑2 of formula auto‑route writing to column A instead of C). Moves us from “sanitize and hope” to “enforce and repair”.

Notes
- One revision pass is performed to keep latency predictable. This integrates cleanly with the existing AllowedCommands/WritePolicy/Schema path from the earlier policy refactor.

## Validation Snapshot (2026‑03‑20, Tests 19/18/16)

Summary
- Test 19 (formula_autoroute):
  - Result: PASS. Step 2 now emits set_values with strings beginning with '=' (e.g., "=A2+B2") in C2:C4, which evaluate as formulas. No cycles or math errors.
  - Cause of prior replans: planner wrote formulas into B (self‑reference) and later returned plain numbers. Fixed by (a) simple‑fill command gating to set_values and (b) values‑only formula guidance.

- Test 18 (no_write_input_column):
  - Step 1: IMPROVED. Returns a single 2D set_values aligned to B:D (header echo appears as first row; apply path drops it when it matches headers).
  - Step 2: PARTIAL. Adds A5:A6 but did not reliably produce B5:D6 outputs on repair; out‑of‑selection writes were correctly rejected.
  - Next: strengthen repair hinting for append‑rows (explicitly: "fill only the selected rows for outputs; do not touch earlier rows"), and consider a two‑phase prompt (first inputs A, then outputs B:D for the same row span) when the model struggles to co‑emit both.

- Test 16 (hebrew_roots):
  - Step 1: IMPROVED. Plan width matches selection (8); respects write policy. Header echo detected and dropped on apply. Some content quality is still mixed (transliteration vs script), to be handled by typed schema hints.
  - Step 2: PARTIAL. Adds A7:A9; planner did not co‑emit B:I reliably on the first repair.

What we changed in this pass
- Recompute WritePolicy/Schema after Test Runner overrides location, avoiding mis‑stated writable columns in revision prompts.
- Values‑only formula guidance: in set_values‑only mode, explicitly request formulas as strings starting with '=' and suppress set_formula messaging.
- Heuristic command gating: for simple fills (no formulas/structural ops mentioned), restrict to set_values.

Next incremental fixes
- Add targeted repair phrasing for append scenarios: "Write outputs only in the selected rows (B..D for rows 5..6). Do not write to earlier rows." Include selection row span explicitly in the revision request.
- Typed schema hints (e.g., transliteration ASCII vs Hebrew script) to improve Test 16 content fidelity.

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

# Latest Updates (v0.3 Execution Plan)

1) Parser‑driven dependency extraction
- Spreadsheet now uses `FormulaEngine.EnumerateReferences` (AST walk) to extract references and ranges, ignoring string literals and nested constructs safely.
- Improves correctness of the dependency graph for incremental recalc.

2) Incremental‑on‑edit repaint (gated)
- On `CellEndEdit`, compute affected cells via `_sheet.RecalculateDirty` and refresh only those cells using the existing 5% threshold fallback to full refresh.
- Immediate dependent updates; full fallback remains for reliability.

3) Copy/Paste: absolute/relative reference rewriting
- Copy tags the clipboard with origin and a structured payload of raw values. Paste prefers that payload and rewrites cell references and ranges by delta, honoring `$` anchors.
- Bugfix: Resolved a case where pasting could yield malformed refs like `=B11` by anchoring at the active cell, ending edit before paste, and using a regex‑based rewriter outside string literals.

4) CSV async I/O + guards
- Added `ExportCsvAsync`/`ImportCsvAsync` and wired File → Import/Export CSV to async paths with `SetUiBusy` to disable menus and show wait cursor.
- Long operations do not block the UI.

5) Chat UX improvements
- Revise loop: Chat dialog adds “Revise” to append feedback and re‑plan without closing.
- Post‑apply error feedback: After apply, scan changed cells for `#ERR`; prompt to attempt a fix; if confirmed, open Chat with prefilled prompt and auto‑plan (review before applying).

6) Provider limits/config
- Anthropic `max_tokens` is controlled by `ANTHROPIC_MAX_TOKENS` (default 2048). Planning timeout configurable via `AI_PLAN_TIMEOUT_SEC` (default 30s).

7) Planner context header detection
- Heuristic picks the first non‑empty row with predominantly text values as the header row for the workbook summary, instead of always row 1.

## Quick Validate
- Edit A1 and observe immediate B1/C1 updates; large changes fallback to full refresh.
- Copy 2×2 formula block with `$` variants; paste elsewhere; verify `$` anchors preserved and relative refs rewritten.
- Import/Export CSV: menus disabled and wait cursor shown during operation.
- Chat: Plan → Revise with feedback; then Apply. Introduce an error and confirm the post‑apply prompt to repair.
   - Paste detects TSV and writes a 2D block starting at the top‑left of the current selection; records a single bulk undo and uses incremental repaint.

## AI Enhancements (Context, Commands, Chat History)

## Test 22–28 Findings (2026‑03‑20)

Summary of latest run based on `tests/output/` snapshots and plans.

- Test 22 — formula_ifs
  - Output: No snapshot found in `tests/output/` for this run.
  - Engine readiness: Core implements IFERROR, SUMIF, COUNTIF, AVERAGEIF in `Core/FormulaEngine.cs` and they’re exercised in the workbook.
  - Status: Likely PASS, but cannot confirm without a saved snapshot. Suggest re‑run with “Save snapshots” enabled to verify values in C2–D4.

- Test 23 — ai_insert_delete_cols
  - Step 1 expected: Insert column B; set B1='Inserted'; do not alter other data.
  - Observed plan: `insert_cols` then `set_values` for B1..B4. Snapshot shows duplicates in B2:B4 (Alpha/Beta/Gamma) and original Name still in C — data altered.
  - Step 2 expected: Delete column C.
  - Observed plan: Repeated `insert_cols` + `set_values` at C; snapshot shows further drift (headers/data misaligned).
  - Root causes:
    - FillMapping section is injected into the user message, biasing the model to generate `set_values` even for structural tasks.
    - AllowedCommands not restricted to structural ops for this prompt; WritePolicy allowed B writes across rows.
  - Status: FAIL.
  - Fixes:
    - Only include FillMapping when AllowedCommands includes `set_values` and the intent is a schema fill.
    - For “insert/delete column” prompts, set AllowedCommands to only the structural command(s) plus, when explicitly requested, a single‑cell `set_values` for the header (B1). Tighten WritePolicy to forbid non‑header writes for this step.
    - In revision message, explicitly state “Do not write any values except B1”.

- Test 24 — ai_delete_sheet
  - Observed plan: `delete_sheet` name=Temp.
  - Snapshot: Only Keep remains with its data.
  - Status: PASS.

- Test 25 — ai_copy_move_range
  - Step 1 expected: Copy A1:B3 → D1:E3 (preserve headers; rewrite formulas only where applicable).
  - Observed plan: Empty `commands` (no change in snapshot).
  - Step 2 expected: Move A1:B3 → F1:G3.
  - Observed plan: Both `copy_range` and `move_range` to F; snapshot shows F1/F2… partly right but header text “H1/H2” rewritten to formulas `=M1`/`=M2` (wrong), and originals not cleared.
  - Root causes:
    - SanitizePlanToBounds currently converts `copy_range` into a `clear_range` (bug in sanitation), so step 1 is effectively a no‑op.
    - Apply path rewrites non‑formula text that “looks like” a cell ref (e.g., “H1”), changing headers to refs (H1→M1). That rule is correct for autorouted formulas, but not for copy/move.
    - AllowedCommands should be exclusive per action: copy OR move, not both.
  - Status: FAIL.
  - Fixes:
    - SanitizePlanToBounds: preserve `copy_range` and `move_range` arguments instead of substituting `clear_range`.
    - In ApplyPlan for copy/move, only rewrite formulas (strings starting with “=”), never plain text even if it resembles a ref. Remove the `LooksLikeCellRefOrRange` branch for these commands.
    - Heuristics: if prompt contains “copy … to …” allow only `copy_range`; if it contains “move … to …” allow only `move_range`.

- Test 26 — ai_set_format
  - Step 1 expected: Format A1:D1 (bold, center, #EEEEEE) without changing values.
  - Observed plan: `set_values` overwriting header with first data row, then `set_format` — snapshot confirms header values changed.
  - Step 2 expected: Apply number format 0.00 to D2:D4 only.
  - Observed plan: `set_values` replaced formulas with literal numbers, then `set_format` with number_format applied.
  - Root causes:
    - FillMapping biases the planner toward `set_values` for non‑fill tasks.
    - General guidance text suggests “Use set_values for plain text …” even when the allowed command should be purely formatting.
  - Status: FAIL.
  - Fixes:
    - For “set format/format …” prompts, AllowedCommands = [`set_format`] only; remove FillMapping block and remove “Use set_values …” guidance.
    - Keep formulas intact; apply number_format only.

- Test 27 — ai_set_validation
  - Expected: `set_validation` list mode on B2:B10 with allowed [Low, Medium, High]; empty allowed.
  - Observed plan: `set_values` writing Low/Medium/High into cells; no `set_validation` command.
  - Root causes:
    - AllowedCommands not restricted; guidance biases toward `set_values`.
    - Workbook export doesn’t persist validations, so snapshots won’t capture rules even when applied (rules are enforced in‑memory/UI only).
  - Status: FAIL.
  - Fixes:
    - For prompts that contain “validation”, AllowedCommands = [`set_validation`]; remove FillMapping and “Use set_values …”.
    - Optional: extend workbook I/O to persist validations for verifiable snapshots.

- Test 28 — ai_set_conditional_format
  - Expected: `set_conditional_format` on A2:A5 with op “>” threshold 12; apply bold + #FFFF99 for matching cells; do not alter values.
  - Observed plan: `set_values` blanked A2:A3 and re‑wrote A4:A5; snapshot shows values changed and no formats applied.
  - Root causes:
    - AllowedCommands not set to `set_conditional_format`; FillMapping/guidance steer to `set_values`.
  - Status: FAIL.
  - Fixes:
    - For prompts that contain “conditional format/conditional formatting”, AllowedCommands = [`set_conditional_format`] only; remove FillMapping and value‑writing guidance.
    - Apply formatting immediately based on numeric value, leave cell contents untouched.

Process and prompt refinements to implement:
- Planner message shaping
  - Only append the FillMapping section when AllowedCommands includes `set_values` and the task intent is a schema/table fill. Omit for structural/formatting/validation tasks.
  - If AllowedCommands is set, remove generic guidance like “Use set_values …”; instead, explicitly state: “Use only: <list>.”
- AllowedCommands heuristics in `RunChatStepAsync`
  - Map intent keywords to explicit command sets:
    - “insert column/insert col” → [`insert_cols`] (+ allow a single `set_values` for a specified header cell if mentioned).
    - “delete column/delete col” → [`delete_cols`].
    - “copy … to …” → [`copy_range`].
    - “move … to …” → [`move_range`].
    - “set format/format …/bold/align/number format/color” → [`set_format`].
    - “validation/data validation/list/number between” → [`set_validation`].
    - “conditional format/conditional formatting/threshold” → [`set_conditional_format`].
  - Continue using values‑only gating for explicit “set_values only”.
- Sanitation and apply correctness
  - Fix SanitizePlanToBounds to preserve `copy_range` and `move_range` (do not convert to `clear_range`).
  - In ApplyPlan for copy/move, only rewrite references for formulas (raw starts with “=”); do not rewrite plain text that looks like refs.
  - Tighten header‑row protections: keep HeaderRowReadOnly by default; allow explicit header exceptions (e.g., “set B1 …”) via a narrow per‑step override rather than broad writable B column.
- Test runner ergonomics
  - Persist a minimal text log per step with diffs and any assertion failures to `tests/output/<test>_stepN.log.txt` to aid review.
  - For engine tests like 22, ensure “Save snapshots” is on so we can confirm outcomes from disk.

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

## UI Modernization (2026‑03‑19)

### Problem
WinForms defaults (3D sunken borders, royal-blue selection, system-drawn headers) made the app look dated. The question was whether C# inherently forces this — it doesn't; the defaults are just old.

### Changes (zero performance cost)

1) **Grid flat styling** (`MainForm.Designer.cs`)
   - Removed `Fixed3D` border; set `BorderStyle.None`.
   - Cell borders changed to `Single` with light gray gridlines (`228, 228, 228`).
   - Background set to white; grid background color matched.

2) **Modern selection colors**
   - Selection: soft blue (`200, 220, 240`) with dark text (`30, 30, 30`) instead of the default saturated royal blue with white text.

3) **Flat headers**
   - Disabled `EnableHeadersVisualStyles` to take control of header rendering.
   - Column and row headers: flat gray background (`240, 240, 240`), dark text, `Single` border style.
   - Column headers bumped to 28px height and centered.

4) **Font consistency**
   - Set Segoe UI 9pt as the default cell, column header, and row header font.

5) **Double buffering** (`MainForm.cs`)
   - Enabled via reflection (`DoubleBuffered = true`) on the DataGridView to eliminate scroll/resize flicker.

6) **Control dock order fix**
   - Reordered `Controls.Add` calls so the Fill-docked grid is added first (WinForms lays out in reverse add-order). Previously the grid could overlap tabs/headers, hiding column names on startup.

### Iteration
- Initially included alternating row colors (`245, 247, 250`) and `SingleHorizontal` cell borders, which made the grid look like a list rather than a spreadsheet. Removed alternating colors and switched to `Single` (full grid) based on feedback.

### What was NOT done (noted for future)
- Menu icons, dark mode toggle, owner-drawn flat tabs.
- These are tracked in BACKLOG.md under "UI modernization."

## Post‑Testing Fixes (Chat + Runner) — 2026‑03‑19

- Planner context alignment to selection
  - Change: In the Test Runner path, after `BuildPlannerContext`, we now parse the requested `location` (single cell or range) and overwrite `AIContext.StartRow/StartCol/Rows/Cols/Title/SelectionValues` to exactly match it.
  - Impact: Prevents over/under‑fills (e.g., stray `D7`) and makes provider outputs align to the visible selection.
  - File: UI/MainForm.cs (RunChatStepAsync).

- Immediate repaint after incremental updates
  - Change: After `RefreshDirtyOrFull`’s targeted cell updates, call `grid.Refresh()` to force a repaint and ensure the UI reflects each step instantly during automation.
  - Also scroll the grid to the selection when the Test Runner sets a range, so the changed region is visible.
  - File: UI/MainForm.cs.

- Mock totals labeling in multi‑turn (Test 06)
  - Change: When adding the total row, write "Total" into the Description column on the totals row, and keep totals in C/D.
  - File: Core/AI/MockChatPlanner.cs.

- Result verification
  - Step 1: A3:C6 headers + 3 rows.
  - Step 2: D3:D6 (Tax header + formulas down only data rows).
  - Step 3: Row 6 totals in C/D with B6 labeled "Total". Snapshots and on‑screen view now match.

## Planner Robustness + Workbook Semantics (2026‑03‑20)

1) Tolerant plan parsing for set_values/set_formula
   - Problem: Providers sometimes emit JSON numbers/booleans in `set_values.values`. Our parser previously treated every cell as string (`GetString()`), which caused parse exceptions and dropped whole commands.
   - Change: Coerce any JSON value to a string when parsing `set_values` (strings, numbers, booleans, objects via `GetRawText()`). `set_formula` parsing is also tolerant.
   - Impact: Plans that include numeric literals now apply reliably without losing the entire write block.

2) Richer workbook summary for planning
   - Added per‑sheet fields: `HeaderRowIndex`, `DataRowCountExcludingHeader`, and `UsedTopLeft`/`UsedBottomRight` addresses.
   - Header detection heuristic picks the first predominantly‑text non‑empty row; falls back to the first used row.
   - Impact: Providers can reason about “data rows (excluding header)” and the true used range without guessing from rough counts.

3) Selection bounding in automation
   - When executing AI steps via the Test Runner, we now sanitize the plan to the requested selection range: intersect write regions and drop out‑of‑bounds writes.
   - Impact: Prevents accidental spillover writes during automated runs while keeping normal chat behavior unchanged.

4) Provider stability
   - OpenAI/Anthropic planner calls now use `temperature=0.0`; Anthropic `max_tokens` increased to 2048.
   - Impact: Reduces variance and truncation for larger JSON plans without test‑specific prompts.

## Hebrew Roots Test + Batch Fill Vision (2026‑03‑20)

### MVP use case identified
The target workflow: user enters Hebrew roots in column A, headers across row 1 describe morphological forms (Verb Qal Perf/Imperf, Noun masc/fem, Adjective masc/fem, Semantic Domain), and AI fills the empty grid. This generalizes to any "input column + header schema → AI populates" pattern and drives the design of a new **Structured Batch Fill** feature.

### Scale analysis and batching strategy
- **≤40 rows:** Single API call. Current Generate Fill / Chat can handle this today with minor prompt tuning.
- **100 rows:** Needs 3–5 batches of 20–30 rows. System prompt + headers are identical across batches → prompt caching saves ~90% input cost on batches 2–N.
- **1,000 rows:** 20–50 batches. Needs parallel execution (3–5 concurrent), progress bar, per‑batch error recovery, incremental apply.
- **10,000+ rows:** Anthropic Batch API (async, 50% cost discount) or offline CSV workflow.

Key insight: the task is embarrassingly parallel (row N doesn't depend on row N−1), the schema (headers) is static, and input per row is tiny (a Hebrew root is 2–4 chars). This makes it ideal for aggressive caching and batching.

### Roadmap additions
- Added **AI v0.3.x — Structured Batch Fill** to Roadmap.md with three staged milestones (single‑shot → batch orchestration → async bulk).
- Added **Structured Batch Fill** section to ENHANCEMENTS.md (AI drag‑to‑fill gesture, batch orchestration, cost estimation, Anthropic Batch API).
- Added **Active — Structured Batch Fill** section to BACKLOG.md.

### UX gaps identified and added to roadmap
Three user‑reported missing spreadsheet primitives, added as **v0.3.x — Core Spreadsheet UX** in Roadmap.md:
1. **Formula bar / cell viewer** — no editable bar above the grid; only a read‑only status strip at the bottom. Fundamental UX gap.
2. **Clipboard selection outline** — no visual feedback on copy/cut (no marching ants or solid border).
3. **Fill Down (Ctrl+D) + drag‑fill handle** — not implemented. One of the most‑used spreadsheet interactions.

### Test 16: Hebrew roots morphology
- Created `tests/test_16_ai_hebrew_roots.workbook.json` with 9 headers and 5 pre‑filled roots (כתב, למד, שמר, מלך, קדש).
- Added 2‑step automation to `TEST_SPECS.json`: Step 1 fills B2:I6 for existing roots; Step 2 adds 3 more roots (פקד, שׁפט, ברך) and fills B7:I9.
- Updated `TEST_INDEX.md` (now 16 tests).

### Test 16 — first run observations
- **Bug fixed:** `ObjectDisposedException` in `TestRunnerForm.RunSelectedAsync` — the form's `_txtLog.AppendText` was called after the form was disposed (user closed the window while async steps were still running). Fixed by adding a `Log()` helper that guards on `IsDisposed`, replacing all direct `_txtLog.AppendText` calls, and adding an early‑exit check in the step loop.
- **AI accuracy issues (prompt tuning needed):**
  - The planner produced 8 separate `set_values` commands (one per column) instead of a single 2D block. Functionally correct but inefficient — more commands means more parsing and a larger plan JSON.
  - **Root‑to‑row misalignment:** Row 2 (כתב) received forms for שמר; row 3 (למד) received forms for כתב; rows shifted by one. The planner didn't anchor output rows to the input roots in column A.
  - **Transliteration column (B) received Hebrew verb forms** instead of Latin transliterations. The prompt asked for "transliteration of the root consonants" but the AI wrote vocalized Hebrew.
  - **Noun/Adjective columns had repetitions** from the verb forms (e.g., שָׁמֵר appeared in Noun, Adjective, and Fem Adjective for the same row).
  - Step 2 never ran (crashed on the disposed‑form bug before reaching it).

### Lessons for prompt engineering
- The planner prompt needs to **explicitly associate each output row with the root in column A** for that row. Current context sends NearbyValues which includes the roots, but the AI doesn't reliably pair them.
- Column B should specify "Latin/English transliteration" more clearly to avoid Hebrew script output.
- Consider a dedicated prompt template for "fill based on column A using row 1 headers" that is more structured than free‑form chat.
- For the batch fill feature, the prompt should send roots as an explicit numbered list mapped to output rows, not rely on spatial context inference.

## Formula Bar + Schema‑Fill Prompt Fix + Log() Crash (2026‑03‑20)

### What We Shipped

1) **Formula bar / cell viewer** (`MainForm.Designer.cs`, `MainForm.cs`)
   - Panel docked above the grid with cell‑name box (60px, left) + editable formula bar (fill).
   - Synced from `UpdateStatus()` via `_suppressFormulaBarSync` guard (suppressed while the bar has focus, like Excel cancel‑on‑blur).
   - Enter commits edit + records undo; Escape reverts; Tab commits + moves right.
   - Name box accepts typed addresses (e.g., "C5" + Enter navigates).

2) **NearbyValues context window shift** (`MainForm.cs` — `BuildPlannerContext`)
   - Shifted `startR` up by 1 row and `startC` left by 1 column so the planner can see headers above and input columns to the left of the selection.
   - Two‑line change: `Math.Max(0, sr - 1)` and `Math.Max(0, sc - 1)`.

3) **Schema‑fill detection** (`ProviderChatPlanner.cs`)
   - System prompt appended: "When filling a table from a list of inputs, combine all rows into a single set_values command with a 2D values array."
   - `TryBuildSchemaFillSection()` helper detects when NearbyValues row 0 has headers and rows 1+ have non‑empty col 0 values, then emits an explicit `FillMapping` section mapping each row to its input value and target columns.

4) **Crash logging** (`Program.cs`, `MainForm.cs`)
   - Global exception handlers now write to `crash.log` before showing MessageBox.
   - `RunChatStepAsync` wraps both `BuildPlannerContext()` and `PlanAsync()` in try/catch with file logging.

### The Log() Infinite Recursion Bug

**Symptom:** Every test runner execution caused an instant hard crash — no error dialog, no log file, process just died. Affected all tests, not just test 16.

**Debugging journey:**
- Added `crash.log` to `RunChatStepAsync` — no file created → crash was before that code.
- Added `crash.log` to global exception handlers in `Program.cs` — still no file → not a .NET exception.
- Disabled formula bar (Controls.Add commented out) — still crashed → not the formula bar.
- Disabled NearbyValues shift + schema‑fill helper — still crashed → not our Priority 1 changes.
- Concluded the crash predated all our changes entirely.

**Root cause:** `TestRunnerForm.Log()` (line 198‑202) had been refactored to add a disposed‑object guard but the body called itself instead of `_txtLog.AppendText(text)`:
```csharp
private void Log(string text)
{
    if (IsDisposed || _txtLog.IsDisposed) return;
    Log(text);  // ← INFINITE RECURSION → StackOverflowException
}
```
A `StackOverflowException` cannot be caught by any .NET exception handler — it terminates the process immediately. This is why no crash dialog appeared and no log file was written.

**Fix:** Changed `Log(text)` → `_txtLog.AppendText(text)`.

**Lesson:** When a .NET app dies with no exception dialog and no log, suspect `StackOverflowException` (uncatchable) or `AccessViolationException`. The disposed‑guard refactor in a previous session introduced this bug by accidentally creating a self‑recursive call.

### Test 16 Results (post‑fix, NearbyValues shift still reverted during this run)

**Step 1** (fill B2:I6 for 5 roots):
- Produced **8 separate set_values commands** (one per column) instead of a single 2D block — the system prompt instruction to combine wasn't enough without the schema‑fill helper active.
- **Root‑to‑row misalignment persists:** Row 2 (כתב) got forms for a different root; Transliteration column (B) received Hebrew text instead of Latin.
- **Semantic Domain (I) got Hebrew** instead of English keywords.
- Workbook summary headers saved partial alignment — the AI knew column names but couldn't see roots in column A (NearbyValues shift was reverted for crash isolation).

**Step 2** (add 3 more roots to A7:A9, fill B7:I9):
- Produced **1 set_values command** with a proper 3×9 2D array — much better structure.
- Transliteration (B7:B9) correctly in Latin: "paqad", "shapat", "barakh".
- Verb forms (C, D) in Hebrew with vowel pointing — correct.
- Noun/Adjective columns (F, G, H) mostly blank — the selection bounding sanitizer clipped them because the AI wrote to col 1 (A) which is outside the B‑I selection.
- Semantic Domain (I7:I9) correctly in English: "command", "judgment", "blessing".

**Step 2 is closer to the target behavior.** The key difference: step 2's prompt explicitly named the roots with transliterations, giving the AI a clear row↔input mapping. Step 1 relied on spatial context which was empty.

### Next Steps
- Re‑run test 16 with NearbyValues shift + schema‑fill helper **enabled** (now re‑enabled) to see if the FillMapping section fixes step 1's alignment.
- The schema‑fill helper should make step 1 behave more like step 2 by giving the planner explicit "Row N: input=X" mappings.
- Column B transliteration issue needs a prompt template fix (specify "Latin alphabet" more forcefully).

## Chat UI Unification + Docked Pane Fixes (2026‑03‑20 PM)

Findings
- Two chat surfaces (docked pane + pop‑out form) duplicated logic and confused users; policy/schema preview in the docked pane didn’t reflect current selection (stuck at 5×1).

Fixes
- Docked pane now refreshes its policy/schema preview on grid selection change and when the pane is shown. The preview matches the pop‑out.
- Revision guidance strengthened: prompts now include explicit row bounds and append‑only input rules (improves Test 18 behavior).
- Unique sheet name guard: `create_sheet` now auto‑suffixes duplicate names (e.g., “Name (2)”).
- Typed schema hint: when a column header contains “Transliteration”, the schema preview notes “latin alphabet” to steer providers away from Hebrew script in that column (improves Test 16).

Decision
- Unify chat surfaces. The docked pane is the canonical chat UI; the “Open Chat” command focuses the pane. The pop‑out window has been removed to avoid duplicate logic.

Validation (E2E 16–19 re‑run)
- Test 18: append‑only constraints surfaced; step 2 produced outputs in B5:D6 while preserving earlier rows (structural + intent pass).
- Test 19: values in A/B then formulas in C; if `create_sheet` appears, duplicate‑name protection prevents collisions.
- Test 16: step 1 still risks Hebrew in Transliteration; typed schema hint added; further typed schema support planned.

Next Work
- Refactor to a single ChatAssistantView + shared ChatSession used by both docked and (optional) pop‑out wrappers; deprecate the second codepath.
- Extend Test Runner structural assertions (per‑row width, header‑dup checks) and add content checks where safe.

## Agent Loop + Queries Expansion (2026‑03‑20, late PM)

What changed
- Query intents expanded and formalized: added `describe_column` and `count_where` to the existing set (`selection_summary`, `profile_column`, `unique_values`, `sample_rows`). Query‑only system schema updated; parser fills `AIPlan.Queries`; `AgentLoop` executes and appends compact transcript lines.
- Plan preview rationales: commands can include an optional `rationale` string (preview‑only); Chat panel shows a “reason:” line under each planned action.
- Test Runner observations export: `ai_agent` steps now emit `observations.json` with the transcript array for downstream tooling.
- Automation safety: default Selection Hard Mode enabled for Test Runner chat/agent runs to drop out‑of‑bounds writes before apply.

Follow‑ups
- Expand observation library (regex match counts, simple histograms).
- Allow Mock planner to emit sample rationales for demo purposes.

## Two‑Phase Agent Loop + Query Schema (2026‑03‑20 PM)

What we shipped
- Formal query intents in the planner schema: selection_summary, profile_column, unique_values, sample_rows. The planner can now return a top‑level `{"queries":[...]}` block when `AIContext.RequestQueriesOnly=true`.
- Host‑controlled two‑phase Agent Loop:
  1) Phase 1: request queries only, execute them deterministically via ObservationTools, and build an Observations transcript.
  2) Phase 2: send original prompt + transcript; planner returns write commands which we sanitize and apply.
- Fallback: when queries are unavailable (e.g., Mock), we use the previous single‑phase built‑in observations.

Safety and UX
- Selection Hard Mode: when enabled in context (and via Chat pane toggle), the planner drops any out‑of‑bounds set_values/set_formula writes after parsing.
- Chat pane adds policy toggles (input column policy + selection hard mode) and a "Copy Observations" button for convenient sharing of the transcript.

Tests
- Added `test_30_ai_agent_observe.workbook.json` and spec to run `ai_agent` with apply=false on a small 3‑column dataset; the Test Runner logs an Observations section and asserts it is present.
- Updated `test_29_ai_agent_city_cleanup` flow benefits from two‑phase loop: the transcript lists uniques/profile first, then the plan is proposed within selection bounds.

Persistence
- Freeze Panes state (top row / first column) is now persisted per sheet in workbook JSON and re‑applied on load.

Notes
- Query coverage can be expanded (e.g., `count_where`, `describe_column`) and per‑command rationales can be surfaced in the plan preview in a later pass.

## Docs Viewer + Markdown Rendering + JSON Export (2026-03-20)

What we shipped
- Added a lightweight Docs Viewer (Help > View Docs…) that lists root‑level Markdown files and renders them as HTML in‑app.
- Added Export Docs JSON (Help > Export Docs JSON) and a `--export-docs` CLI to write `docs/docs_index.json` containing files, sections, and raw content.
- Introduced a small `Core/DocsIndexer` to scan top‑level `*.md` and split sections by headings.
- Added a tiny offline Markdown renderer (`Core/Markdown.cs`) to cover headings, paragraphs, lists, blockquotes, horizontal rules, code blocks, inline code, bold/italic, links, and images. Styling is embedded CSS for simplicity.

Hurdles / fixes
- Initial regex strings used standard C# escapes which produced CS1009 “Unrecognized escape sequence” and CS1012 char literal issues on Windows. Rewrote patterns as verbatim strings (`@"..."`) and corrected char literals in the line splitter to use `'\r'` / `'\n'` directly.
- Ensured WebBrowser control is quiet (script errors suppressed) and safe (no drag‑drop).

Rationale
- Keep it minimal and offline so it’s easy to experiment with and doesn’t add package dependencies. JSON export provides a simple path to integrate docs elsewhere if desired.

Follow-ups
- Anchor navigation within the full document and auto-scroll to a selected section.
- Option to include nested docs (e.g., `tests/TEST_INDEX.md`) and a configurable include list.
- Basic link hygiene: open external links in the system browser and keep in-app view sandboxed.

## Core UX: Fill Down/Right + Clipboard Outline + Schema Fill v1 (2026-03-20)

What we shipped
- Fill Down (Ctrl+D): copies the top cell per selected column down to the rest of the selection. Relative references are rewritten via the existing paste-rewrite path; absolute anchors are preserved. Bulk undo + incremental repaint.
- Fill Right (Ctrl+R): copies the leftmost cell per selected row rightward with the same rewrite semantics. Bulk undo + incremental repaint.
- Drag‑fill handle (basic): small square at the bottom‑right of the selection; drag to extend the selection down/right/diagonally. For straight down/right extensions, detects simple numeric series from the last two values and increments accordingly; otherwise applies a repeating pattern with reference rewrite. Bulk undo + incremental repaint. More series types (dates, weekdays) are deferred.
- Clipboard selection outline (static v1): solid border for Copy; dashed border for Cut. Clears on Escape and Paste. Implemented via a lightweight overlay in the grid's Paint handler. Marching ants animation is deferred to the roadmap.
- Schema Fill v1 (menu action): AI > Fill Selected From Schema… triggers a values-only plan using headers as schema and the input column to the left, sanitized to the selection, applied as one grouped undo. Single-shot for small ranges; batching/drag gesture comes later.

Files
- Fill Down/Right and outline: `UI/MainForm.cs`, `UI/MainForm.Designer.cs`.
- Schema Fill v1 action: `UI/MainForm.cs`, `UI/MainForm.Designer.cs`.
- Tests: `tests/test_20_fill_down.workbook.json`, `tests/test_21_ai_schema_single_shot.workbook.json`, and `tests/TEST_INDEX.md` updated.

Notes
- Reference rewriting leverages the same code paths used for Copy/Paste to ensure consistency (handles single refs and ranges, respects string literals, preserves `$` anchors).
- Outline rendering is intentionally simple for now to avoid perf regressions. An animated border can be added with a small timer updating DashOffset.

Environment constraint / partner note
- Frustration acknowledged: asking to "run a smoke test" is not actionable here since the app builds and runs only in your Windows environment, and I cannot build or run it myself from this sandbox. Going forward, I’ll assume you will execute validation locally and I’ll focus on delivering ready-to-run changes with concise in-app verification steps.

Formula engine updates
- Implemented IFERROR, SUMIF, COUNTIF, and AVERAGEIF with range-size validation and comparison parsing (supports =, <>, <, <=, >, >= against numbers or strings).
## E2E Results — New Tests 17–19 (2026‑03‑20)

Test 17 — Append-only multi‑turn
- Step 1 (A2:D4):
  - Plan: one set_values for A..C and set_formula in D with =B(row)*C(row).
  - Applied: A2:D4 populated; D contains formulas; within bounds; no header edits.
- Step 2 (A5:D6 append):
  - Plan: set_values for A..C and set_formula in D for the two new rows; no changes to prior rows.
  - Applied: A5:D6 filled; A2:D4 unchanged; within bounds.
Observations: Pass (structural + intent). Minor: user prompt shows WritableColumns=A (should list A,B,C,D).

Test 18 — No‑write to input column + append
- Step 1 (B2:D4):
  - Plan: set_values at B2:D4 but returned empty strings; no writes to A (as intended).
  - Applied: sheet unchanged (content-wise); structurally within bounds.
- Step 2 (A5:D6):
  - Plan: set_values adds inputs in A5:A6 and fills B..D accordingly; existing rows untouched.
  - Applied: delta/epsilon rows appended with outputs; no modification to A2:A4; within bounds.
Observations: Structural pass; content for step 1 was empty (model choice). Policy respected (no A writes until append step).

Test 19 — Formula auto‑route from set_values
- Step 1 (A2:B4):
  - Plan: included create_sheet + set_title + set_values and set_formula. Selection bounds kept data in range but created a duplicate sheet named “FormulaAuto”.
  - Applied: two sheets now exist: original with headers; second with numbers at A2:B4.
- Step 2 (C2:C4, set_values only, write =A2+B2 etc.):
  - Plan: produced a set_values targeting A2:A4 (2/4/6), not C2:C4, and without '=' formulas.
  - Applied: sanitized to selection bounds → no overlapping cells → effectively no changes.
Observations: Structural guards prevented out‑of‑range writes; however, step 1 allowed create_sheet since not forbidden, causing a duplicate‑name sheet; step 2 did not demonstrate formula auto‑routing because the model did not emit '=' values into C under set_values.

Follow‑ups from tests (queued)
- Correct WritableColumns hint in the user prompt for multi‑column selections (currently shows only the first letter in some cases).
- Tighten default AllowedCommands for narrow tasks (or pass an explicit allowlist per step) to prevent unintended create_sheet/set_title when not asked.
- Optionally block duplicate‑name sheet creation unless explicitly requested, or auto‑rename/new sheet confirmation.
- Consider a brief “revise” on policy/shape mismatch (e.g., step 19.2 targeting A instead of C) to elicit repair rather than silently dropping.

## Major Feature Batch: Cross‑Sheet Refs, Grammar, Chat Unification, Batch Fill, Low‑Hanging Fruit (2026‑03‑20)

### What We Shipped

#### 1. Cross‑Sheet Formula References (`=Sheet2!A1`)
- **Parser:** Added `Exclamation` token type, `SheetCellRefNode(Sheet, Address)` and `SheetRangeNode(Sheet, A, B)` AST nodes. Handles both bare identifiers (`Sheet2!A1`) and single‑quoted names (`'My Sheet'!A1`).
- **Evaluator:** Added optional `_crossSheetResolver` callback (`Func<string, string, EvaluationResult>`) to `FormulaEngine`. `SheetCellRefNode` delegates to the resolver; `SheetRangeNode` works inside aggregate functions (SUM, COUNT, etc.) via the existing `AsNumbers` helper.
- **Spreadsheet.cs:** Added `CrossSheetResolver` property (`Func<string, int, int, EvaluationResult>`). `EvaluateCell` now passes a cross‑sheet adapter to the `FormulaEngine` constructor.
- **MainForm.cs:** `WireCrossSheetResolver()` maps sheet names from `_sheetNames` to `_sheets` and is called from `InitializeSheet()` whenever the active sheet changes.
- **Dependency tracking:** Cross‑sheet refs return early in `CollectRefs` — they don't create same‑sheet dependency edges. Cross‑sheet dependency propagation is deferred (editing Sheet2 won't auto‑recalc Sheet1 formulas referencing it unless Sheet1 is recalculated explicitly).
- Files: `Core/FormulaEngine.cs`, `Core/Spreadsheet.cs`, `UI/MainForm.cs`.

#### 2. AI Command Grammar: `insert_rows` / `delete_rows`
- Added `InsertRows` and `DeleteRows` to `AICommandType` enum and new `InsertRowsCommand` / `DeleteRowsCommand` classes in `Commands.cs`.
- `ProviderChatPlanner.cs`: Extended system prompt with JSON schemas (`{“type”:”insert_rows”,”at”:<1-based>,”count”:<int>}`), added `ParsePlan` cases, updated `IsCommandAllowedByList` and `DescribeType`.
- `MainForm.cs` `ApplyPlan`: `InsertRowsCommand` shifts data down (bottom‑up to avoid overwrites), clears inserted rows, copies formats. `DeleteRowsCommand` records old values, shifts up, clears vacated bottom rows. Both call `Recalculate()` and mark all affected rows for repaint.
- Files: `Core/AI/Commands.cs`, `Core/AI/ProviderChatPlanner.cs`, `UI/MainForm.cs`.

#### 3. Chat Surface Unification
- `ChatAssistantForm.cs` rewritten from 222 lines of duplicated `DoPlanAsync`/`BuildPolicyPreview`/history logic to a 30‑line thin wrapper that embeds a `ChatAssistantPanel` instance with a Close button. Constructor signature preserved — all callers (pop‑out, error repair dialog, test runner) work unchanged.
- File: `UI/AI/ChatAssistantForm.cs`.

#### 4. Batch Schema Fill Skeleton (Stage 2)
- New `Core/AI/BatchSchemaFiller.cs`: Splits large fills into configurable batches (default 30 rows), runs with `SemaphoreSlim`‑based concurrency (default 3 concurrent), per‑batch retry (up to 2), 60s per‑batch timeout, `Action<FillProgress>` callback for progress reporting.
- `CloneContextForBatch` builds a fresh `AIContext` per batch with headers + only that batch's input values (prompt‑cache friendly — system prompt + headers identical across batches).
- `MainForm.cs`: Added `BatchFillSelectedFromSchemaAsync()` — falls back to single‑shot for ≤40 rows, shows confirmation dialog with batch count, updates title bar with progress, applies each successful batch with bounds sanitization. Added menu item “Batch Schema Fill…” in the AI menu.
- Files: `Core/AI/BatchSchemaFiller.cs` (new), `UI/MainForm.cs`, `UI/MainForm.Designer.cs`.

#### 5. Structured Prompt Template for Schema Fill
- `TryBuildSchemaFillSection` in `ProviderChatPlanner.cs` now emits explicit ordering instructions: “Produce exactly one row of N values per input below, in the SAME order. Each output row MUST correspond to the input on that row. Do NOT reorder or skip inputs.” Plus total expected row/column counts.
- This directly addresses the root‑to‑row misalignment observed in Test 16 where the AI shuffled or omitted inputs.
- File: `Core/AI/ProviderChatPlanner.cs`.

#### 6. Formula Engine: 11 New Functions
- **COUNTA** — count non‑empty cells (vs COUNT which only counts numbers).
- **ISBLANK, ISNUMBER, ISTEXT** — type‑checking predicates (return 1/0).
- **TRIM** — strip leading/trailing whitespace.
- **PROPER** — title‑case via `CultureInfo.TextInfo.ToTitleCase`.
- **TODAY / NOW** — return date/datetime as `yyyy-MM-dd` / `yyyy-MM-dd HH:mm:ss` strings.
- **TEXT** — format a number with a format string; handles `%` patterns and maps `#` → `0` for .NET compat.
- **REPLACE** — replace characters by position (1‑based start, count, replacement text).
- File: `Core/FormulaEngine.cs`.

#### 7. Freeze Panes
- View menu with two checkable toggles: “Freeze Top Row” and “Freeze First Column”. Uses native DataGridView `Rows[0].Frozen` / `Columns[0].Frozen`.
- Files: `UI/MainForm.cs`, `UI/MainForm.Designer.cs`.

#### 8. AI Action Log
- New `Core/AI/AIActionLog.cs`: Session‑scoped list of `Entry` records (timestamp, prompt summary, command count, cell count, summary). `Record()` called at the top of `ApplyPlan`.
- `ShowAIActionLog()` in MainForm opens a `ListView`‑based dialog. Menu item “View Action Log…” in the AI menu.
- Files: `Core/AI/AIActionLog.cs` (new), `UI/MainForm.cs`, `UI/MainForm.Designer.cs`.

#### 9. Explain Cell
- AI menu item “Explain Cell…” builds a prompt from the active cell's raw content and evaluated value, then opens the docked chat pane (or dialog fallback) with `autoPlan: true`.
- Files: `UI/MainForm.cs`, `UI/MainForm.Designer.cs`.

#### 10. Schema Fill Hotkey + Selection Heuristic
- Bound to `Ctrl+Shift+F`. If only a single cell is selected, `SmartSchemaFillAsync` auto‑expands to the output rectangle: finds header row, rightmost header column, input column, last data row, then selects the empty output region before calling `FillSelectedFromSchemaAsync`.
- Files: `UI/MainForm.cs`, `UI/MainForm.Designer.cs`.

### Files Modified (9)
- `Core/FormulaEngine.cs` — cross‑sheet parser/evaluator + 11 new functions
- `Core/Spreadsheet.cs` — CrossSheetResolver property + wiring
- `Core/AI/Commands.cs` — InsertRows/DeleteRows types
- `Core/AI/ProviderChatPlanner.cs` — insert/delete grammar, structured prompt template
- `UI/AI/ChatAssistantForm.cs` — rewritten as thin wrapper
- `UI/MainForm.cs` — all feature integrations
- `UI/MainForm.Designer.cs` — menu items for all new features

### Files Created (2)
- `Core/AI/AIActionLog.cs` — session action log
- `Core/AI/BatchSchemaFiller.cs` — batch orchestration

### Known Limitations / Follow‑ups
- Cross‑sheet dependency propagation is not automatic: editing Sheet2 won't recalc Sheet1 formulas until Sheet1 is recalculated. Needs a workbook‑level recalc pass.
- `insert_cols` / `delete_cols` not yet implemented (only rows).
- Batch orchestration is skeleton — needs real‑world testing with >40 rows and cost estimation dialog.
- Explain Cell uses the planner (which returns JSON commands); a dedicated text‑only endpoint would be better for pure explanations.
- Freeze panes state doesn't persist across sheet switches or save/load.

## Agent Loop Gating Fix + New Agent Tests (2026‑03‑20 PM2)

What changed
- Agent loop now mirrors chat path gating: we de‑bias Title when the prompt forbids titles and pass an explicit AllowedCommands list derived from the prompt (e.g., ["set_values"] for “Use set_values only”). File: `UI/MainForm.cs:RunAgentStepAsync`.

Why
- Test 29 previously yielded no plan under OpenAI because the planner inferred `transform_range` from the prompt and our post‑filter removed it due to "Use set_values only", leaving an empty plan. By passing AllowedCommands upfront, the provider is steered to return `set_values` sized to the selection.

Results
- Test 29 now proposes a single `set_values` (11×1) and applies the expected five fixes within B2:B12 (trim + Title Case). Logs show the 5 cell changes.

New tests added
- Test 31 — agent values‑only gating: simple 1‑column city list; two steps (observe, apply). Expects `set_values` only and selection‑bounded normalization.
- Test 32 — agent transform_range: same dataset; prompt allows `transform_range normalize_city`; single apply step.
- Test 33 — agent observe‑only strict: observe transcript on a small 3‑column dataset with apply=false.
- Test 34 — selection fencing: selection is a subset of messy values; prompt enforces in‑selection normalization with `set_values` only; apply=true.

Next
- Make apply=false strictly Phase 1 (skip writes planning) so observe‑only runs leave `plan.commands` empty.
- Extend SelectionHardMode/SanitizePlan to fence `transform_range` writes to selection bounds.
- Add lightweight plan assertions in the Test Runner (e.g., error log if apply=false and commands were proposed; structural checks for OOB writes).
