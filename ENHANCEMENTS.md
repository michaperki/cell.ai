# Enhancement Ideas

Items not already covered in Roadmap.md, BACKLOG.md, or ENHANCEMENT_PLAN.md. AI expansions listed first.

**Last audited: 2026-03-19** — see status annotations on each item.

---

## AI Command Grammar Expansions (2 of 8 done)

The chat planner currently supports 7 command types (set_values, set_title, set_formula, sort_range, create_sheet, clear_range, rename_sheet). These are the remaining impactful additions:

### `set_formula` — IMPLEMENTED
Distinct from `set_values`, this writes formulas (not plain text) into cells. Currently `set_values` writes plain text only.
Schema: `{"type":"set_formula","start":{...},"formulas":[["=SUM(B2:B10)"]]}`.
Unlocks prompts like "add a total row" or "calculate the average salary in column D." The AI generating formulas that reference existing data is arguably the highest-value AI task in a spreadsheet.

### `sort_range` — IMPLEMENTED
Sort a rectangular region by a specified column, ascending or descending.
Schema: `{"type":"sort_range","start":{...},"rows":N,"cols":N,"sort_col":"B","order":"asc","has_header":true}`.
Enables "sort my expenses by amount" or "alphabetize the names column."

### `insert_rows` / `insert_cols` — NOT STARTED
Insert blank rows or columns at a position, shifting existing data down/right.
Schema: `{"type":"insert_rows","at":5,"count":2}`.
Currently there's no way to ask the AI to "add a header row above my data" without overwriting existing content.

### `delete_rows` / `delete_cols` — NOT STARTED
Remove rows or columns, shifting data up/left.
Enables "remove the empty rows" or "delete column C."

### `set_format` — NOT STARTED
Apply formatting (bold, color, number format) to a range.
Schema: `{"type":"set_format","start":{...},"rows":N,"cols":N,"bold":true,"number_format":"0.00"}`.
Deferred in the roadmap to v0.4 but enabling it only through the AI planner is a lighter lift than full format-command UI — the AI becomes the interface. Note: CellFormat infrastructure already exists in Core/CellFormat.cs.

### `move_range` — NOT STARTED
Cut a rectangle and paste it elsewhere.
Enables "move column A data to column C" type tasks.

### `copy_range` — NOT STARTED
Copy values (or formulas) from one range to another.
Enables "duplicate this table to Sheet2."

### `delete_sheet` — NOT STARTED
Currently missing from the command set. `create_sheet` and `rename_sheet` exist but not delete.
Enables "remove the scratch sheet."

---

## AI Context Enrichment (4 of 4 done)

All priority context items are shipped. The planner now receives rich context on every request.

### Send existing cell data around the selection — IMPLEMENTED
`NearbyValues` in AIContext sends a pipe-delimited snapshot of up to 20 rows x 10 columns around the current selection. The AI can now read headers, see existing data, and write formulas referencing real column names.

### Workbook-level summary — IMPLEMENTED
`Workbook` array of `SheetSummary` objects sent to planner, each containing sheet name, used-range dimensions, and first-row headers. Enables cross-sheet awareness.

### Selection content — IMPLEMENTED
`SelectionValues` in AIContext sends the current values of the selected cells as a 2D array, bounded to 20x10. Enables transform-style prompts.

### Conversation history — IMPLEMENTED
Rolling 10-message history (5 user/assistant exchanges) maintained in ChatAssistantForm's `_history` list. Sent as `[role:content]` pairs in the planner context. Reset button available.

---

## New AI Interaction Modes (0.5 of 5 done)

### Natural language formula help — NOT STARTED
The user types `=` in a cell, then presses a hotkey (or after a pause), and the AI suggests a formula based on the cell's position relative to headers and data. Differs from the v0.4 "function assistance" roadmap item by being inline and context-driven rather than a separate dialog.

### Explain cell — NOT STARTED
Right-click or hotkey on a cell with a formula, AI explains what it does in plain language. Low token cost since it's sending one formula string.

### Error repair — NOT STARTED
When a cell shows `#ERR:`, offer an AI action: "Fix this formula." Send the formula, the error message, and surrounding context. The AI returns a corrected formula. Tightly scoped and high-value.

### Selection-aware transforms — PARTIAL
Selection content is now sent to the AI planner (see context enrichment above), so the data flow exists. However, there's no dedicated UX flow for "transform selected data" — the user must manually describe the transformation in chat. A dedicated "Transform Selection" action would streamline this.

### Multi-turn chat with undo-aware re-planning — NOT STARTED
Multi-turn chat works (conversation history is shipped), but undo and chat are completely decoupled. If the user undoes an AI apply, the chat has no awareness of this. Detecting undo-of-AI and offering to revise would close the loop.

---

## Formula Engine Gaps (~7 of 10 done)

Functions not yet implemented that would both enrich the spreadsheet and expand what the AI can generate via `set_formula`:
Implemented: IFERROR, SUMIF, COUNTIF, AVERAGEIF
- **`SUMIF` / `COUNTIF` / `AVERAGEIF`** — NOT IMPLEMENTED. Conditional aggregates. Extremely common in real spreadsheet use.
- **`HLOOKUP`** — IMPLEMENTED. Horizontal lookup, complement to VLOOKUP.
- **`INDEX` / `MATCH`** — IMPLEMENTED. More flexible than VLOOKUP; the modern lookup pattern.
- **`IFERROR`** — NOT IMPLEMENTED. Wraps errors gracefully; `=IFERROR(A1/B1, 0)`.
- **`TRIM`** — NOT IMPLEMENTED. **`UPPER` / `LOWER`** — IMPLEMENTED. **`PROPER`** — NOT IMPLEMENTED. Text cleanup functions.
- **`TEXT`** — NOT IMPLEMENTED. Format a number as text with a format string: `=TEXT(A1, "0.00%")`.
- **`TODAY` / `NOW`** — NOT IMPLEMENTED. Date functions (even if just returning serial numbers or formatted strings).
- **`COUNTA`** — NOT IMPLEMENTED. Count non-empty cells (COUNT only counts numbers).
- **`ISBLANK` / `ISNUMBER` / `ISTEXT`** — NOT IMPLEMENTED. Type-checking functions.
- **`SUBSTITUTE`** — IMPLEMENTED. **`REPLACE`** — NOT IMPLEMENTED. String replacement.

---

## Data Validation & Structured Input — NOT STARTED

### Dropdown lists / data validation per cell
Define a list of allowed values for a cell or column. The AI could set these up: "make column B a dropdown with Yes/No/Maybe." Natural AI command extension (`set_validation`).

### Named ranges
Allow naming a range (e.g., `Expenses = A2:A50`). Named ranges improve both human formula authoring and AI context — the AI could reference named ranges instead of brittle cell addresses.

---

## Cross-Sheet References — NOT STARTED

The formula engine currently resolves cell references within a single sheet. Supporting `=Sheet2!A1` or `=SUM(Sheet2!B:B)` syntax would unlock real multi-sheet workbooks and let the AI write formulas that pull data across sheets.

---

## Observability & Trust — NOT STARTED

### AI action log
Keep a per-session log of every AI plan applied (timestamp, prompt, commands, cell count). Accessible from a menu item. Helps users understand what the AI changed, especially after stepping away.

### Token/cost estimation
Before sending a request, estimate token count from the context pack and show it in the chat UI. Helps users understand the cost of richer context.

### Dry-run diff view
The chat preview currently shows a plan summary list. A richer preview would show a before/after diff of affected cells — actual cell values changing, not just "Set 5x3 values at B2." Builds trust and catches errors before apply.

---

## Quality of Life — PARTIALLY STARTED

### Auto-detect header rows — IMPLEMENTED
Heuristic to identify the first non-empty row with text values as a header row. Used for AI context (send headers automatically) and workbook summary.

### Formula bar / cell viewer — IMPLEMENTED
A dedicated editable text area above the grid (like Excel's formula bar) showing the active cell's address (name box) and raw contents. The current status bar is read‑only and tucked at the bottom. This is a fundamental spreadsheet UX element — users expect to see and edit cell contents in a prominent bar, especially for long formulas that overflow the cell width.

### Clipboard selection outline — IMPLEMENTED (static v1)
Marching-ants animation deferred.
When the user copies or cuts a range, draw a visible border around the source cells — solid for copy, animated dashed ("marching ants") for cut. Clear on Escape or after paste. This is standard in Excel/Sheets and provides essential visual feedback about what's on the clipboard.

### Fill Down (Ctrl+D) and drag‑fill handle — IMPLEMENTED (basic)
Ctrl+D fills with reference rewriting. Drag handle extends selection; detects simple numeric series for straight down/right fills; otherwise repeats pattern.
Ctrl+D fills selected cells with the value/formula from the top cell of each column, rewriting relative references. A small drag handle at the bottom‑right corner of the selection extends a series or copies values downward/rightward. This is one of the most used spreadsheet interactions — critical for productivity.

### Drag-fill handle
Click and drag the corner of a selection to extend a series (1,2,3... or Mon,Tue,Wed...). Core spreadsheet interaction that's missing. (Subsumed by Fill Down / drag‑fill above.)

### Freeze panes — NOT STARTED
Freeze the top row(s) or left column(s) so headers stay visible during scrolling. DataGridView supports frozen columns/rows natively.

### Conditional formatting rules — NOT STARTED
Color cells based on value thresholds. The AI could set these: "highlight cells over 1000 in red." Another natural `set_conditional_format` command.

---

## Known Issues & Quick Wins (discovered 2026-03-19 audit)

These are actionable items found during codebase review that don't fit neatly into a feature category:

1. **Planner timeout too short (10s).** `ChatAssistantForm.DoPlanAsync` uses a 10s CancellationTokenSource. Complex plans with large context can exceed this. Raise to 20-30s or make configurable.
2. **Anthropic max_tokens=800 is low.** A budget table with formulas can exceed 800 tokens of JSON. OpenAI has no explicit limit, creating asymmetry. Raise to 2048+.
3. **Chat closes on Apply.** `ChatAssistantForm` calls `Close()` after apply, killing multi-turn iteration. Keep window open after apply.
4. **No plan revision UX.** Only options are Apply or Close. A "Revise" button that sends feedback to the planner would be valuable.
5. **`set_values` doesn't auto-detect formulas.** If the AI puts `=SUM(...)` in `set_values`, it's written as literal text. Auto-routing values starting with `=` to the formula path would be a safety net.
6. **No progress indicator during planning.** The UI only disables the Plan button. A "Thinking..." label or spinner would reassure users.
7. **Workbook summary only sends row 1 as headers.** If the actual header row is row 2 (common when row 1 is a title), the AI gets wrong context. Ties into the auto-detect-header QoL item.
8. **MockChatPlanner doesn't cover full command set.** It can't generate `set_formula` or `sort_range` plans, limiting offline development and testing.

---

## Structured Batch Fill — NOT STARTED

### MVP use case: Hebrew morphological forms
User enters Hebrew roots in column A, headers across row 1 describe desired forms (Verb, Noun, Adjective, etc.), AI fills the grid. Generalizes to any "input column + header schema → AI populates" pattern.

### AI drag‑to‑fill gesture — NOT STARTED
Select the empty output range; system detects header row as schema and leftmost non‑empty column as input; fires AI request(s) and fills. Distinct from Generate Fill (which requires a manual prompt) and Chat (which is conversational). This is the natural spreadsheet gesture for AI‑assisted data entry.

### Batch orchestration — NOT STARTED
For >40 rows, split into batches of 20–50 rows. Each batch sends identical system prompt + headers (prompt‑cache friendly) and only that batch's input values. Progress bar, per‑batch error recovery, incremental apply, cancel support.

### Async bulk via Anthropic Batch API — NOT STARTED
For 1,000+ rows, submit as a single async batch job (50% cost discount). Poll for completion, import results. Resumable and re‑runnable for failed rows.

### Cost estimation before fill — NOT STARTED
Before starting a batch fill, estimate token count and cost from the schema + row count. Show a confirmation dialog.

---

## Priority Recommendations (updated)

**Shipped (top 4):**
1. ~~Send existing cell data in AI context~~ — DONE
2. ~~`set_formula` command~~ — DONE
3. ~~Conversation history~~ — DONE
4. ~~`sort_range` command~~ — DONE

**Next priorities:**
5. **Fix chat-closes-on-apply** — low effort, high UX impact for multi-turn workflows
6. **Raise planner timeout and max_tokens** — low effort, fixes truncated/timed-out plans
7. **Auto-detect `=` in set_values** — safety net for AI mistakes, low effort
8. **`insert_rows` / `delete_rows`** — unlocks "add a row" without overwriting data
9. **SUMIF / COUNTIF / IFERROR** — most-wanted formula functions for AI-generated formulas
10. **Cross-sheet references** — unlocks real multi-sheet workbook value

---

## E2E Test Suite

A set of 15 test workbooks lives in `tests/`. Each file is a `.workbook.json` that opens via File > Open Workbook (or the Test Runner under the Test menu). Cell A1 in each file contains step-by-step testing instructions. See `tests/TEST_INDEX.md` for the full list.
