# Backlog / Refinements

This file tracks follow-ups and refinements discovered while implementing the enhancement plan. It lives next to `ENHANCEMENT_PLAN.md`.

**Last updated: 2026-03-20**

---

## Active — AI / Chat UX

- **Docked Chat Pane — DONE**
  - Added a right-side docked Chat panel with Plan/Revise/Apply and rolling history. Toggle via AI > Toggle Chat or Ctrl+Shift+C. This is the only chat surface; the pop-out window was removed.

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

- **Planner revise loop — DONE**
  - ProviderChatPlanner validates plans against selection bounds, AllowedCommands, WritePolicy (writable columns + input-column rules), and expected per-row width. On violations, it issues a single automatic revision request with a compact summary of problems and constraints, then applies the corrected plan if returned. Falls back gracefully if the provider cannot repair.

- **Values-only formula guidance — DONE**
  - In values-only mode, planner instructs: "When formulas are needed, write them as strings beginning with '=' inside set_values cells so they evaluate as formulas." Removes "use set_formula" guidance in this mode.

- **Simple-fill command gating — DONE**
  - For prompts that clearly ask to fill/write simple values and do not mention formulas or structural ops, automatically gate AllowedCommands=["set_values"]. Reduces accidental set_formula/set_title in basic scenarios (e.g., Test 19 step 1).

- **Append-row repair phrasing — NOT STARTED**
  - Add revision messaging that explicitly instructs: "Write outputs only for the selected rows; do not modify earlier rows." Helps Test 18 step 2 produce B:D for new inputs.

- **Schema/policy preview in Chat UI — NOT STARTED**
  - Show the mapped columns (letters), AllowedCommands, and WritePolicy summary above the Plan/Apply actions to reduce surprises and catch over-constraints early.

- **Typed schema constraints (optional) — NOT STARTED**
  - Allow per-column hints like `allow_script=latin` (for transliteration columns) or `type=number` to improve content class fidelity without domain overfitting.

---

## Active — AI Observation Tools (new)

- Query command grammar — PARTIAL
  - Host‑side ObservationTools implemented for `selection_summary`, `unique_values`, `describe_column`, and safe `get_range` sampling (no side effects).
  - Planner query intents not yet formalized as schema/tool‑calls (keep host‑driven loop for MVP).

- Planner integration for queries — PARTIAL
  - Host executes observations and augments the user prompt with a concise transcript before planning.
  - Cap results with top‑K uniques and first N row samples to control tokens.

- UI transcript of observations — DONE (first pass)
  - Chat pane renders an observations section above the plan when agent loop is used. Transcript exported by Test Runner.

- Tests — PARTIAL
  - E2E added: `test_29_ai_agent_city_cleanup` (observe then apply). Unit coverage for ObservationTools still pending.

---

## Active — Agent Loop (MVP)

- Loop controller — DONE (first pass)
  - Host‑side loop (single observation phase → propose plan) wired in Chat pane and Test Runner.

- Safety gating — DONE (first pass)
  - Values‑only/no‑titles filtering honored in agent automation; plan is sanitized to selection before apply.

- UX — DONE (first pass)
  - Chat toggle “Use Agent Loop (MVP)” and transcript section; Test Runner `ai_agent` action with transcript export.

- Demo dataset — DONE (first pass)
  - `test_29_ai_agent_city_cleanup.workbook.json` added; uses messy City data for normalization.

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

## Active — Docs / Tools

- Internal Docs Viewer — EXPERIMENT (DONE)
  - Help > View Docs… renders root‑level Markdown via a lightweight offline converter.
  - Help > Export Docs JSON writes `docs/docs_index.json`; also available via `--export-docs` CLI.

- Follow‑ups
  - Anchor navigation inside a full-file render; auto‑scroll to the selected section.
  - Configurable include paths (e.g., include `tests/TEST_INDEX.md` and nested docs).
  - External link handling: open in system browser; keep viewer sandboxed.

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

- Formula bar / cell viewer — DONE
  - Implemented top formula bar with name box; synced to selection and editable with Enter/Tab commit.

- Clipboard selection outline — DONE (static v1)
  - Solid border for Copy, dashed for Cut; clears on Escape and Paste; drawn in grid Paint. Marching-ants animation is deferred (roadmap).

- Fill Down (Ctrl+D) — DONE; Fill Right (Ctrl+R) — DONE; Drag‑fill handle — DONE (basic)
  - Keyboard-first and drag-handle implementations with reference rewriting for formulas and ranges; grouped undo and incremental repaint. Series detection/increments remain in backlog.

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
- E2E test suite: 16 workbook test files in `tests/` with Test Runner UI (Test menu). See `tests/TEST_INDEX.md`.

- Add new E2E tests (next):
  - Append-only multi-turn table (no overwrites between steps; fixed width, selection bounded).
  - No-write to input column unless explicitly allowed (two-step variant adding new inputs in empty rows only).
  - Formula auto-routing from set_values (values starting with `=` evaluate as formulas).
  - Repair robustness for append-range outputs (explicit selection-row-only writes).

---

## Active — Structured Batch Fill (MVP: Hebrew roots) — PAUSED (reprioritized behind v0.3 Agent Loop)

- Schema Fill v1 (menu action) — DONE (single-shot)
  - AI > Fill Selected From Schema… generates a values-only set_values plan for the current selection using headers as schema and the input column to the left. Single-shot ≤40 rows; sanitized to selection; composite undo.

- AI drag‑to‑fill gesture — NOT STARTED
  - Select empty output range; system detects header row + input column; fires AI request(s); fills. Single-shot for ≤40 rows, batched beyond that.

- **Batch orchestration — NOT STARTED (Paused)**
  - Split large fills into 20–50 row batches. Same system prompt + headers per batch (prompt‑cache friendly). Progress bar, cancel, per‑batch retry, incremental apply.

- **Cost estimation dialog — NOT STARTED (Paused)**
  - Before batch fill, estimate tokens/cost from schema + row count. User confirms before proceeding.

- **Anthropic Batch API integration — NOT STARTED (stretch, Paused)**
  - For 1,000+ rows, submit async batch job at 50% cost. Poll completion, import results, resume on interruption.

- **Prompt template for schema‑driven fill — NOT STARTED (Paused)**
  - Current free‑form chat prompt doesn't reliably align output rows to input column values. Need a structured prompt template that explicitly maps each root (row) to its expected output columns. Discovered via Test 16 where the AI misaligned roots to rows and wrote Hebrew in the transliteration column.

- **TestRunnerForm disposed‑object crash — FIXED**
  - `_txtLog.AppendText` threw `ObjectDisposedException` when the form was closed mid‑run. Added `Log()` guard helper and early‑exit check in step loop.

---

## Active — AI UX (Generate Fill / Inline)

- Generate Fill: support range selection. Use selected rectangle as the target shape and seed the dialog with its rows/cols. Requires enabling DataGridView multi-select.
- Inline continuation: filter out suggestions that duplicate contiguous items above the cursor (case-insensitive, trimmed). Show only new items in the ghost panel and apply only those on accept.
- Chat vs Fill differentiation: keep Chat for multi-command plans; keep Generate Fill for explicit range fills. Ensure both respect write caps and record a single undo group per apply.
