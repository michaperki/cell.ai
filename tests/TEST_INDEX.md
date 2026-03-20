# E2E Test Suite Index

Each test is a `.workbook.json` file. Open via **Test > Test Runner** (navigate with Next/Prev) or **File > Open Workbook**. Automated steps and prompts are defined in `tests/TEST_SPECS.json` and can be executed from the Test Runner.

| # | File | Category | What it tests |
|---|------|----------|---------------|
| 01 | test_01_ai_set_values | AI | `set_values` command fills plain text, composite undo |
| 02 | test_02_ai_set_formula | AI | `set_formula` writes SUM/AVERAGE, formula bar verification |
| 03 | test_03_ai_sort_range | AI | `sort_range` descending, header preservation, undo |
| 04 | test_04_ai_create_rename_sheet | AI | `create_sheet` + `rename_sheet`, undo chain |
| 05 | test_05_ai_clear_range | AI | `clear_range` clears column without touching adjacent data |
| 06 | test_06_ai_multi_turn_chat | AI | 3-step multi-turn conversation building on prior context (auto-run) |
| 07 | test_07_ai_context_awareness | AI | AI reads existing headers/data to place formulas correctly |
| 08 | test_08_formula_engine | Engine | 12 formula scenarios: math, DIV/0, IF, strings, VLOOKUP, etc. |
| 09 | test_09_undo_redo | UX | Single undo, bulk undo (multi-cell delete), composite AI undo |
| 10 | test_10_workbook_io | I/O | Multi-sheet load, format persistence, CSV round-trip |
| 11 | test_11_find_replace | UX | Ctrl+F navigation, Ctrl+H replace all, undo replacements |
| 12 | test_12_copy_paste_cut | UX | Copy, cut, paste data+formulas across regions |
| 13 | test_13_cell_formatting | UX | Bold, number format, alignment, save/reopen persistence |
| 14 | test_14_ai_complex_plan | AI | Multi-command plan (set_values + set_formula in one shot) |
| 15 | test_15_ai_workbook_context | AI | AI reasoning across multiple sheets using workbook summary |
| 16 | test_16_ai_hebrew_roots | AI | Hebrew morphology fill: headers as schema + roots as input; multi-turn append |
| 20 | test_20_fill_down | UX | Keyboard Fill Down/Right; verify formula rewriting and bulk apply |
| 21 | test_21_ai_schema_single_shot | AI | Single-shot schema-driven values-only fill for a small selection |
| 22 | test_22_formula_ifs | Engine | IFERROR, SUMIF, COUNTIF, AVERAGEIF sample formulas |

## How to use

1. **Test Runner (recommended):** Test menu > Test Runner. Use Next/Prev to navigate. Click "Load Test" to load the workbook. Click "Run Steps" to execute prompts/selections from `TEST_SPECS.json`.
   - Optional: enable "Save snapshots" to write workbook snapshots.
   - Optional: enable "Dump plan JSON" to save the raw provider plan to `tests/output/` after each step.
   - Optional: enable "Dump user prompt" to save the constructed user prompt (context + instruction) that was sent to the provider.
   - Optional: enable "Dump system prompt" to save the system schema/rules used for planning.
2. **Manual:** File > Open Workbook, navigate to `tests/`, pick a file. You can still perform actions manually; A1 no longer contains instructions.
3. **AI tests** require an API key configured (OpenAI or Anthropic). Without one, the MockChatPlanner will respond with heuristic-based plans (limited coverage).

### TEST_SPECS.json format
- `tests/TEST_SPECS.json` defines automated steps per test file.
- Each test lists `steps` with:
  - `action`: currently only `ai_chat`
  - `prompt`: message sent to the chat planner
  - `location`: address or range to anchor context (e.g., `A3` or `A3:C6`). This sets the selection before planning.
  - `sheet` (optional): name of the sheet to activate for the step
  - `apply`: whether to apply the planned commands

Example:
{
  "file": "test_06_ai_multi_turn_chat.workbook.json",
  "steps": [
    { "action": "ai_chat", "prompt": "Create an expense table...", "location": "A3:C6", "apply": true },
    { "action": "ai_chat", "prompt": "Now add a Tax column...", "location": "D3:D6", "apply": true },
    { "action": "ai_chat", "prompt": "Add a total row...", "location": "A6:E6", "apply": true }
  ]
}
