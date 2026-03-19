# Enhancement Ideas

Items not already covered in Roadmap.md, BACKLOG.md, or ENHANCEMENT_PLAN.md. AI expansions listed first.

---

## AI Command Grammar Expansions

The chat planner currently supports 5 command types (set_values, set_title, create_sheet, clear_range, rename_sheet). These are the most impactful additions:

### `set_formula` — IMPLEMENTED
Distinct from `set_values`, this writes formulas (not plain text) into cells. Currently `set_values` writes plain text only.
Schema: `{"type":"set_formula","start":{...},"formulas":[["=SUM(B2:B10)"]]}`.
Unlocks prompts like "add a total row" or "calculate the average salary in column D." The AI generating formulas that reference existing data is arguably the highest-value AI task in a spreadsheet.

### `sort_range` — IMPLEMENTED
Sort a rectangular region by a specified column, ascending or descending.
Schema: `{"type":"sort_range","start":{...},"rows":N,"cols":N,"sort_col":"B","order":"asc"}`.
Enables "sort my expenses by amount" or "alphabetize the names column." This is a common spreadsheet task with no current path to accomplish it through the AI.

### `insert_rows` / `insert_cols`
Insert blank rows or columns at a position, shifting existing data down/right.
Schema: `{"type":"insert_rows","at":5,"count":2}`.
Currently there's no way to ask the AI to "add a header row above my data" without overwriting existing content.

### `delete_rows` / `delete_cols`
Remove rows or columns, shifting data up/left.
Enables "remove the empty rows" or "delete column C."

### `set_format`
Apply formatting (bold, color, number format) to a range.
Schema: `{"type":"set_format","start":{...},"rows":N,"cols":N,"bold":true,"number_format":"0.00"}`.
Deferred in the roadmap to v0.4 but enabling it only through the AI planner is a lighter lift than full format-command UI — the AI becomes the interface.

### `move_range`
Cut a rectangle and paste it elsewhere.
Enables "move column A data to column C" type tasks.

### `copy_range`
Copy values (or formulas) from one range to another.
Enables "duplicate this table to Sheet2."

### `delete_sheet`
Currently missing from the command set.
Enables "remove the scratch sheet."

---

## AI Context Enrichment

The current context pack sent to the planner is minimal: sheet name, selection, rows/cols, title. Several improvements would dramatically improve plan quality.

### Send existing cell data around the selection
The AI doesn't know what's already in the sheet. Sending a snapshot of the used range (or at least a configurable window of nearby cells) would let the AI write formulas referencing real column headers, avoid overwriting data, and understand the structure it's working with. This is the single biggest gap — the AI is currently planning blind.

### Workbook-level summary
Send all sheet names, each sheet's used range dimensions, and the first row (headers) of each sheet. This lets the AI make cross-sheet plans ("create a summary sheet pulling totals from each monthly sheet").

### Selection content
When the user selects a range and asks the AI to transform it ("convert these dates to ISO format", "capitalize all names"), the AI needs the current values of the selected cells as input.

### Conversation history
The chat planner currently sends a single user message per plan. Maintaining a rolling conversation context (last 3-5 exchanges) would let the user iterate: "now add a column for tax" after "create an expense table." This turns the chat from single-shot into a real assistant.

---

## New AI Interaction Modes

### Natural language formula help
The user types `=` in a cell, then presses a hotkey (or after a pause), and the AI suggests a formula based on the cell's position relative to headers and data. For example, cursor in the cell below a column labeled "Total" next to a SUM-able range — the AI suggests `=SUM(B2:B10)`. Differs from the v0.4 "function assistance" roadmap item by being inline and context-driven rather than a separate dialog.

### Explain cell
Right-click or hotkey on a cell with a formula, AI explains what it does in plain language. Useful for inherited or complex workbooks. Low token cost since it's sending one formula string.

### Error repair
When a cell shows `#ERR:`, offer an AI action: "Fix this formula." Send the formula, the error message, and surrounding context. The AI returns a corrected formula. Tightly scoped and high-value.

### Selection-aware transforms
Select a range, open chat, and the AI can see the selected data and transform it: "parse these addresses into separate City/State/Zip columns", "convert currencies", "extract numbers from these strings." Requires sending selection content (see context enrichment above).

### Multi-turn chat with undo-aware re-planning
If the user undoes an AI apply, the chat could detect this and offer to revise the plan. Currently undo and chat are completely decoupled.

---

## Formula Engine Gaps

Functions not yet implemented that would both enrich the spreadsheet and expand what the AI can generate via `set_formula`:

- **`SUMIF` / `COUNTIF` / `AVERAGEIF`** — Conditional aggregates. Extremely common in real spreadsheet use.
- **`HLOOKUP`** — Horizontal lookup, complement to VLOOKUP.
- **`INDEX` / `MATCH`** — More flexible than VLOOKUP; the modern lookup pattern.
- **`IFERROR`** — Wraps errors gracefully; `=IFERROR(A1/B1, 0)`.
- **`TRIM` / `UPPER` / `LOWER` / `PROPER`** — Text cleanup functions.
- **`TEXT`** — Format a number as text with a format string: `=TEXT(A1, "0.00%")`.
- **`TODAY` / `NOW`** — Date functions (even if just returning serial numbers or formatted strings).
- **`COUNTA`** — Count non-empty cells (COUNT only counts numbers).
- **`ISBLANK` / `ISNUMBER` / `ISTEXT`** — Type-checking functions.
- **`SUBSTITUTE` / `REPLACE`** — String replacement.

---

## Data Validation & Structured Input

### Dropdown lists / data validation per cell
Define a list of allowed values for a cell or column. The AI could set these up: "make column B a dropdown with Yes/No/Maybe." Natural AI command extension (`set_validation`).

### Named ranges
Allow naming a range (e.g., `Expenses = A2:A50`). Named ranges improve both human formula authoring and AI context — the AI could reference named ranges instead of brittle cell addresses.

---

## Cross-Sheet References

The formula engine currently resolves cell references within a single sheet. Supporting `=Sheet2!A1` or `=SUM(Sheet2!B:B)` syntax would unlock real multi-sheet workbooks and let the AI write formulas that pull data across sheets.

---

## Observability & Trust

### AI action log
Keep a per-session log of every AI plan applied (timestamp, prompt, commands, cell count). Accessible from a menu item. Helps users understand what the AI changed, especially after stepping away.

### Token/cost estimation
Before sending a request, estimate token count from the context pack and show it in the chat UI. Helps users understand the cost of richer context.

### Dry-run diff view
The chat preview currently shows a plan summary list. A richer preview would show a before/after diff of affected cells — actual cell values changing, not just "Set 5x3 values at B2." Builds trust and catches errors before apply.

---

## Quality of Life

### Auto-detect header rows
Heuristic to identify the first non-empty row with text values as a header row. Use this for AI context (send headers automatically), for sort operations, and for formatting (auto-bold headers).

### Drag-fill handle
Click and drag the corner of a selection to extend a series (1,2,3... or Mon,Tue,Wed...). Core spreadsheet interaction that's missing.

### Freeze panes
Freeze the top row(s) or left column(s) so headers stay visible during scrolling. DataGridView supports frozen columns/rows natively.

### Conditional formatting rules
Color cells based on value thresholds. The AI could set these: "highlight cells over 1000 in red." Another natural `set_conditional_format` command.

---

## Priority Recommendations

The highest-impact items are:

1. **Send existing cell data in AI context** — so it stops planning blind
2. **`set_formula` command** — so the AI can write real formulas
3. **Conversation history** — for multi-turn chat
4. **`sort_range` command** — most-requested spreadsheet operation with no current AI path

Status update: `set_formula`, `sort_range`, basic context enrichment, and conversation history are now implemented. Remaining priorities: send richer selection/workbook content as structured JSON, and expand conversation memory.
