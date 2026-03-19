# Rustsheet Enhancement Plan

## Context
The spreadsheet app has a solid foundation (~838 lines of C#) with a formula engine, basic file I/O, and a WinForms UI. We're implementing 19 enhancements covering the formula engine, UI, architecture, and I/O.

## Phase 1: Core Infrastructure (no UI changes)

### 1A. Consolidate magic numbers — DONE
- **Files**: `Core/Spreadsheet.cs`, `IO/SpreadsheetIO.cs`
- Add `public const int DefaultRows = 100; public const int DefaultCols = 26;` to `Spreadsheet` class
- Replace hardcoded `100`/`26` in `SpreadsheetIO.LoadFromFile` with `Spreadsheet.DefaultRows`/`Spreadsheet.DefaultCols`
- `MainForm.cs` already has its own constants — point them at the shared ones

### 1B. Fix division by zero — return `#DIV/0!` error instead of NaN — DONE
- **File**: `Core/FormulaEngine.cs:61`
- Change `b == 0 ? double.NaN : a / b` → `b == 0 ? return FromError("DIV/0!") : FromNumber(a / b)`

### 1C. Dependency-graph recalculation — DONE (engine + maps; UI uses full refresh for now)
- **File**: `Core/Spreadsheet.cs`
- Add `_dependencies` dictionary: `Dictionary<(int,int), HashSet<(int,int)>>` mapping each cell to the cells it references
- Add `_dependents` reverse map: cells that depend on a given cell
- On `SetRaw`, parse formula to extract references, update both maps
- Add `RecalculateDirty(int row, int col)` that does a topological-sort BFS from the changed cell through dependents, only re-evaluating affected cells
- Keep `Recalculate()` as full recalc (used on file load)
- Change `MainForm.Grid_CellEndEdit` to call `RecalculateDirty` instead of full recalc

### 1D. Undo/Redo system — DONE
- **New file**: `Core/UndoManager.cs`
- Simple command stack: record `(int row, int col, string? oldRaw, string? newRaw)` on each `SetRaw`
- `Undo()`: restore oldRaw, recalculate
- `Redo()`: restore newRaw, recalculate
- Wire into `Spreadsheet.SetRaw` or wrap at the MainForm level

## Phase 2: Formula Engine Enhancements

### 2A. Add comparison operators to parser — DONE
- **File**: `Core/FormulaEngine.cs`
- Add token types: `Eq`, `NotEq`, `Lt`, `Gt`, `LtEq`, `GtEq`
- Add `ParseComparison()` precedence level between `ParseAddSub` and the top-level `ParseExpression`
- Comparisons return 1.0 (true) or 0.0 (false)
- Lexer: handle `=`, `<>`, `<`, `>`, `<=`, `>=` — note `=` is tricky since formulas start with `=` but that's stripped before parsing

### 2B. Add string literal support to parser — DONE
- **File**: `Core/FormulaEngine.cs`
- Add `StringLiteral` token type and `StringNode` AST node
- Lexer: when `"` encountered, read until matching `"`
- Add `&` as string concatenation operator (new token type `Ampersand`)
- Add `ParseConcat()` precedence level (lowest, below comparison)
- `EvalNode` for `&`: convert both sides to display strings and concatenate

### 2C. Add new formula functions — DONE
- **File**: `Core/FormulaEngine.cs` — `EvalFunc` method
- **Conditional**: `IF(cond, true_val, false_val)` — treats 0/empty as false, anything else as true
- **Logical**: `AND(...)`, `OR(...)`, `NOT(val)` — return 1.0/0.0
- **String**: `LEN(text)`, `LEFT(text, n)`, `RIGHT(text, n)`, `MID(text, start, len)`, `CONCATENATE(...)`
- **Math**: `ABS(n)`, `ROUND(n, digits)`, `CEILING(n)`, `FLOOR(n)`, `MOD(a, b)`
- **Lookup**: `VLOOKUP(search, range, col_index, [exact])` — search first column of range, return value from col_index column

### 2D. Function argument validation — DONE
- **File**: `Core/FormulaEngine.cs` — `EvalFunc` method
- For each function, check arg count and return clear error like `"SUM expects at least 1 argument"`
- `IF` requires exactly 2 or 3 args, `NOT` requires 1, `LEFT`/`RIGHT` require 2, `MID` requires 3, etc.

## Phase 3: UI Enhancements

### 3A. Copy/paste support — DONE
- **File**: `UI/MainForm.cs`
- Override `ProcessCmdKey` or handle `KeyDown` for Ctrl+C, Ctrl+V, Ctrl+X
- Copy: store selected cell's raw value to clipboard
- Paste: set cell raw from clipboard, recalculate
- Cut: copy then clear cell

### 3B. Status bar with cell info — DONE
- **Files**: `UI/MainForm.Designer.cs`, `UI/MainForm.cs`
- Add a `StatusStrip` with labels: cell address, raw formula, computed value
- Update on `SelectionChanged` event of the grid

### 3C. Keyboard navigation improvements — DONE
- **File**: `UI/MainForm.cs`
- Tab moves right, Enter moves down (after edit commit) — DataGridView may already do some of this; verify and enhance
- Handle arrow keys for navigation during non-edit mode

### 3D. Find and replace — DONE
- **New file**: `UI/FindReplaceForm.cs` (small dialog)
- **File**: `UI/MainForm.cs` — add Ctrl+F / Ctrl+H shortcuts
- Search through `_sheet.GetRaw()` for all cells, highlight/navigate to matches
- Replace: update raw value and recalculate

### 3E. Multi-sheet (tabs) support — PARTIAL (UI tabs done; workbook I/O pending)
- **Files**: `UI/MainForm.cs`, `UI/MainForm.Designer.cs`, `Core/Spreadsheet.cs`, `IO/SpreadsheetIO.cs`
- Add `TabControl` below menu, above grid
- `_sheets` list of `Spreadsheet` objects with names
- Add/remove/rename sheet tabs via right-click context menu
- Update IO format to save/load array of sheets
- Each tab switch swaps the active `Spreadsheet` and refreshes grid

### 3F. Cell formatting — DONE (basic number format support)
- **New file**: `Core/CellFormat.cs` — stores bold, foreground color, background color, number format, alignment
- **File**: `Core/Spreadsheet.cs` — add `_formats` dictionary parallel to `_raw`
- **File**: `UI/MainForm.cs` — apply formatting in `RefreshGridValues`, add Format menu or toolbar
- **File**: `IO/SpreadsheetIO.cs` — serialize/deserialize formats

### 3G. Column resize support — DONE
- **File**: `UI/MainForm.Designer.cs`
- Already has `AllowUserToResizeRows = true`; add `AllowUserToResizeColumns = true`
- Set a reasonable default column width instead of auto-size

## Phase 4: I/O Enhancements

### 4A. CSV import/export — DONE
- **File**: `IO/SpreadsheetIO.cs` — add `ExportCsv` and `ImportCsv` methods
- **File**: `UI/MainForm.cs` — add menu items for CSV import/export
- Export: iterate rows/columns, write comma-separated values (handle commas in values with quoting)
- Import: parse CSV, create new Spreadsheet, populate cells

### 4B. Async file I/O — PARTIAL (IO methods added; UI adoption pending)
- **File**: `IO/SpreadsheetIO.cs` — add `SaveToFileAsync` and `LoadFromFileAsync` using `File.WriteAllTextAsync`/`File.ReadAllTextAsync`
- **File**: `UI/MainForm.cs` — make save/open handlers async, show cursor wait

### 4C. Recent files list — DONE
- **File**: `UI/MainForm.cs`, `UI/MainForm.Designer.cs`
- Store last 5 file paths in a simple text file at `AppData/Local/SpreadsheetApp/recent.txt`
- Add "Recent Files" submenu under File menu
- Update on each Open/Save

### 4D. Absolute/relative cell references ($A$1) — DONE
- **File**: `Core/FormulaEngine.cs` — parser handles `$` prefix on column letters and row digits
- For now, parse and ignore `$` (strip it) so formulas like `=$A$1+B2` work
- Future copy/paste can use the distinction

## Implementation Order
1. Phase 1A (constants) — quick, unblocks nothing but good hygiene
2. Phase 1B (div/0) — one-line fix
3. Phase 2A (comparison operators) — needed by IF
4. Phase 2B (string literals + &) — needed by string functions
5. Phase 2C (new functions) — depends on 2A, 2B
6. Phase 2D (arg validation) — do alongside 2C
7. Phase 4D (absolute refs) — parser change, do with other parser work
8. Phase 1C (dependency graph) — architecture change, do before UI work
9. Phase 1D (undo/redo) — needs to hook into SetRaw
10. Phase 3B (status bar) — small UI addition
11. Phase 3A (copy/paste) — needs undo integration
12. Phase 3C (keyboard nav) — small
13. Phase 3G (column resize) — small
14. Phase 3D (find/replace) — new dialog
15. Phase 3F (cell formatting) — significant, new data model
16. Phase 3E (multi-sheet tabs) — significant, changes IO format
17. Phase 4A (CSV) — IO addition
18. Phase 4B (async IO) — refactor existing IO
19. Phase 4C (recent files) — UI + small file persistence

## Verification
- Build with `dotnet build` after each phase
- Manual testing: enter formulas with new functions, test IF/AND/OR, string ops, comparison operators
- Test undo/redo, copy/paste, find/replace
- Test save/load with new features (multi-sheet, formatting, CSV)
- Verify dependency-graph recalc by editing cells with chains of references
