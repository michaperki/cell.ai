# Rustsheet Roadmap

This roadmap organizes near-term work, medium-term polish, and stretch goals for the Rustsheet project. We’ll iterate this collaboratively.

## Vision & Scope
- A fast, reliable desktop spreadsheet with a solid formula engine, clean I/O, and pragmatic UI.
- Prioritize correctness and usability over breadth; grow features sustainably.

## Current State Snapshot
- Core, formula engine, and many UI features shipped; CSV I/O and async APIs added.
- Open/partial items: multi‑sheet workbook Save/Open; adopt async I/O in UI.
- Full-sheet recalc used after edits; dependency graph exists for future incremental repaint.

## Release Milestones

### v0.2 — Reliability & Multi‑sheet Parity (shipped, first pass)
- Workbook Save/Open for multiple named sheets — added `.workbook.json`, names + formats, backward compatible loader.
- Switch UI Open/Save to async with wait cursor and UI disable.
- Generate Fill respects selection shape (rows x cols), multi-select enabled, single bulk undo.
- Inline continuation deduplication and safe apply of only new items.
- Chat planner extensions: `clear_range`, `rename_sheet` (strict JSON), composite apply.

### v0.3 — Performance & UX Polish (near‑term)
- Incremental repaint: use RecalculateDirty for bulk ops (done); evaluate safe incremental-on-edit (CellValidated/CommitEdit), with 5% fallback.
- Copy/paste semantics honoring absolute/relative refs ($A$1 vs A1).
- Expanded number formatting presets and alignment tweaks.
- Undo polish: coalesce rapid edits; grouped actions state reflected in menu enable/disable.
- Async adoption: extend to CSV import/export and workbook flows; add guard flags and progress affordances.
- Unique sheet name guard: prevent duplicate names when creating sheets (auto‑suffix).
- **UI modernization (low-cost visual polish):**
  - Grid: flat borders, modern selection accent, subtle alternating row colors, single-horizontal cell borders.
  - Headers: flat-styled column/row headers with custom background instead of 3D sunken default.
  - Double-buffered DataGridView to eliminate scroll/resize flicker.
  - Consistent default cell font (Segoe UI 9pt).

### v0.3.x — Core Spreadsheet UX (near‑term, user‑reported gaps)
- **Formula bar / cell viewer:** Add a dedicated editable text area above the grid (like Excel's formula bar) that shows the active cell's address and raw contents. Clicking a cell populates it; editing there commits back to the cell. The current status bar at the bottom is read‑only and cramped — this is a fundamental spreadsheet affordance.
- **Clipboard selection outline:** When the user copies (Ctrl+C) or cuts (Ctrl+X) a range, draw a visible outline around the source cells — solid border for copy, animated dashed ("marching ants") for cut — matching Excel's behavior. Clear the outline on Escape or after paste.
- **Fill Down (Ctrl+D) and drag‑fill handle:** Ctrl+D fills the selection with the value/formula from the top cell of each column (rewriting relative refs). A small drag handle at the bottom‑right corner of the selection extends a series or copies values/formulas downward/rightward on drag. This is one of the most fundamental spreadsheet interactions.

### v0.4 — Interop & Distribution (stretch)
- Evaluate XLSX import/export or a stable interchange path.
- Packaging, versioning, and basic telemetry/crash reporting.

## AI Roadmap (Inference‑Driven)

### Principles (from INFERENCE.md)
- Trigger smartly, not per keystroke: explicit commands first; later, debounced inline suggestions (100–250 ms) with immediate cancel on edit.
- Send tiny, structured context snapshots: selection shape, nearby headers/title, sheet name; avoid full‑workbook uploads.
- Enforce structured outputs: JSON with exact cell payloads sized to the selection.
- Aggressive cancellation and timeouts: every request cancellable; default 10s timeout.
- Cache by prefix/shape/context to keep latency snappy; local gating drops low‑value requests.
- Provider‑agnostic: pluggable `IInferenceProvider` behind a small contract.

### AI v0.1 — Generate Fill (Preview) — DONE
- UX: AI > Generate Fill… (Ctrl+I) opens a modal prompt; preview fills the current selection; Accept applies in a single undo step.
- Context: selection address/shape; sheet name; nearest title/header cells; optional sample of existing items above/left.
- Contract: provider returns `{"cells": [[...],[...]]}` sized exactly to the selection; plain text only; no formatting/formulas.
- Implementation: introduce `IInferenceProvider`, `MockProvider` for offline dev; async call with spinner and Cancel; one-shot apply.
- Acceptance:
  - Supports rectangular selections; never writes outside selection.
  - Undo restores prior values; handles short/empty responses gracefully.
  - UI stays responsive; request cancellable; under 1s local mock, target <2–3s real provider.

### AI v0.2 — Inline Ghost Fill (Debounced) — SHIPPED (first pass)
- Triggers: when a user pauses after entering a header like “Grocery List” or after 2–3 items in a single column; debounce 150–250 ms; cancel on any keypress.
- UX: faint preview of the next N items below the cursor; Tab/Enter accepts first suggestion row; menu toggles enable/disable inline AI.
- Context: current column window (few rows above), header/title cell, sheet name; no cross‑sheet context yet.
- Caching: key by (sheet id, column key, recent items hash, header text) to avoid recomputation.
- Safety: suggestions never auto‑commit; no network calls when typing rapidly due to debounce + gating.
- Acceptance:
  - Only 1D list continuation; respects existing data boundaries; zero writes without explicit accept.
  - Cancelation immediate; UI never blocks; memory/CPU overhead minimal.

### AI v0.3 — Chat Assistant (Plan → Preview → Apply) — LIVE (expanded)
- UX: docked side panel with chat (toggle via AI > Toggle Chat or Ctrl+Shift+C); messages generate a dry‑run “plan” of commands; user reviews a diff/summary and applies all or step‑by‑step. The docked pane is the single chat surface; no pop‑out.
- Command grammar now includes:
  - `set_values(range, values2d)`
  - `set_formula(range, formulas2d)`
  - `sort_range(range, sort_col, order, has_header)`
  - `create_sheet(name)`
  - `set_title(range, text)`
  - `clear_range(range)`
  - `rename_sheet(name|index, new_name)`
- Context pack:
  - Selection content (bounded), nearby window values, and workbook summary (sheet names, used sizes, first-row headers).
- Limits & safety: hard cap on changed cells per apply (default 5,000); always preview; undo groups per apply; confirm cross‑sheet changes.
- Streaming & control: show plan as it streams; allow Cancel; retry uses cached context.
- Acceptance:
  - Plans are syntactically valid and scope‑limited; preview matches final apply; undo fully reverts.

### AI v0.3.x — Structured Batch Fill (MVP use case: Hebrew morphology)

The motivating use case: a user enters Hebrew roots in column A with headers describing desired morphological forms (Verb, Noun, Adjective, etc.) across the top row, then uses an "AI fill" gesture to populate the empty cells. This generalizes to any "input column + header schema → AI fills the grid" pattern.

#### Stage 1 — Single‑shot fill (10–40 rows)
- **UX — "AI drag to fill" gesture:** User selects the empty output range. The system detects:
  - Header row (row 1 or detected via heuristic) as the output schema.
  - Input column (leftmost non‑empty column adjacent to selection) as the seed data.
  - Fires a single AI request with headers + input values as context; applies result to the selection.
- **Context optimization:** Strip NearbyValues / full workbook summary. Send only: system prompt (task definition + JSON schema) + headers + input values. Minimal tokens.
- **Model selection:** Test smaller models (Haiku, GPT‑4o‑mini) for factual/linguistic lookup tasks where accuracy is high and creativity is unnecessary. Surface model choice in the fill dialog.
- Acceptance: Works end‑to‑end for ≤40 rows. Single undo group. Preview before apply.

#### Stage 2 — Batch orchestration (100–1,000 rows)
- **Batch planner:** Given the output region size and estimated per‑row token density, split work into N batches of 20–50 rows each. Each batch sends identical system prompt + headers (cache‑friendly) and only that batch's input values.
- **Prompt caching:** Leverage Anthropic prompt caching — system prompt + headers prefix is identical across batches, so batches 2–N pay ~90% less on input tokens.
- **Progress UI:** Progress bar (batch X of N, row Y of Z), Cancel button (finishes current batch, keeps completed work), incremental apply (write results to grid as each batch returns).
- **Error recovery:** Per‑batch retry (up to 2 retries on timeout/parse failure). Failed batches are logged and skippable; user can re‑run failures later.
- **Undo:** One composite undo per batch, plus a "Revert All Batches" option.
- **Cost estimation:** Before starting, show "~N batches, ~X tokens, ~$Y estimated — proceed?" confirmation.

#### Stage 3 — Async bulk processing (1,000–10,000+ rows)
- **Anthropic Batch API integration:** Submit all batches as a single async batch job (50% cost discount, 24h turnaround). Poll for completion; import results when ready.
- **Offline workflow:** Export roots → process externally → import filled CSV. The spreadsheet is the I/O surface, not the processing engine.
- **Resumability:** Persist batch progress to disk so interrupted jobs can resume.
- **Re‑run failures:** Identify rows that errored or returned empty; offer targeted re‑fill.

#### Design principles for batch fill
- **Schema is static:** Headers define what to produce. Send once, cache forever across batches.
- **Input is minimal:** A Hebrew root is 2–4 characters. Per‑row input cost is tiny.
- **Task is embarrassingly parallel:** Row N doesn't depend on row N−1. Batches can run concurrently (with rate‑limit cap of 3–5 concurrent).
- **Temperature 0.0:** Factual/linguistic lookup, not creative generation. Deterministic output.
- **Validation pass:** LLMs aren't perfect on irregular forms. The UX should support spot‑check → revise cycles (existing Chat Revise loop is reusable).

### AI v0.4+ — RAG, Functions, Interop (stretch)
- Retrieval across workbook for richer context; lightweight index of headers/keys.
- Function assistance: suggest formulas and explain results; opt‑in due to higher complexity.
- Interop: optional external knowledge packs (e.g., common grocery lists) cached locally.

### Executive Decisions (Opinionated Defaults)
- Start explicit, then inline: ship v0.1 as explicit command before any background suggestions.
- Structure over prose: models must return strict JSON; UI validates shape before preview.
- Minimal context by default: no full workbook uploads; opt‑in for richer context later.
- Hard limits: command caps, timeouts, and cancellation enforced uniformly.
- Provider‑agnostic: `IInferenceProvider` abstraction; `MockProvider` default in dev.
- Privacy switch: global “Allow AI” toggle with per‑workbook preference.

### Near-term Additions
- Schema/policy preview in Chat UI — IMPLEMENTED
  - Shows selection shape, AllowedCommands, Writable columns, input‑column rule, and a short schema list.
  - Available in both pop‑out assistant and docked chat pane; docked pane refreshes on selection change.
- Planner revise loop: IMPLEMENTED — on width/policy mismatch, the planner automatically requests a corrected plan instead of relying only on cropping.

- Chat surface unification — IMPLEMENTED
  - Docked chat pane is now the only chat surface. “Toggle Chat” shows/hides the pane.
  - The former pop‑out window has been removed to reduce duplication and maintenance.

### Technical Foundation
- Providers: Mock, OpenAI (`OpenAIProvider`), Anthropic (`AnthropicProvider`), optional External POST (`ExternalApiProvider`).
- Env‑first config: `.env`/system env — OPENAI_API_KEY / ANTHROPIC_API_KEY; default model via OPENAI_MODEL / ANTHROPIC_MODEL.
- Core: `Core/AI/IInferenceProvider.cs`, `Core/AI/AIContext.cs`, `Core/AI/AIResult.cs`.
- Inline: overlay + debounce + caching; Apply/Dismiss + hotkeys.
- Chat: `ProviderChatPlanner` (strict JSON plan) and `MockChatPlanner` fallback; composite undo for multi‑op apply.
- UI: `UI/AI/GenerateFillDialog.cs`, `UI/AI/ChatAssistantForm.cs`, AI menu entries + hotkeys; Settings + Test Connection.
- Observability: manual Test Connection; future in‑app counters.

### Test & Verification (AI)
- Unit: shape enforcement, JSON parsing, selection‑bounded writes, undo integrity.
- Scenario: grocery list single‑column fill; cancel mid‑request; preview accept; cache hit on repeat.
- Performance: debounce effectiveness; no main‑thread stalls; memory stable over repeated requests.

## Quality & Verification
- Build gate and quick manual checklist per release.
- Unit tests for parser, functions, I/O, and formatting.
- **E2E Test Suite**: 16 `.workbook.json` files in `tests/` covering AI commands, formula engine, undo/redo, I/O, and UX flows. Steps and prompts live in `tests/TEST_SPECS.json` and run via the Test Runner (Test menu). The runner can save after-step snapshots, dump provider plan JSON and the constructed user/system prompts, and logs a concise diff of changed cells per step. See `tests/TEST_INDEX.md`.
 - Structural assertions: after each applied step, assert that all modified cells are within the requested selection; optional schema-aware checks can be enabled for stricter validation.
 - Docked chat pane policy preview now stays in sync with the grid selection and on open/close.

### Planner Robustness (In Progress)
- Plan parser tolerates typed values in `set_values` and coerces to strings.
- Workbook summary includes `HeaderRowIndex`, `DataRowCountExcludingHeader`, and `UsedTopLeft/UsedBottomRight` to improve general reasoning.
- Test Runner sanitizes AI plans to selection bounds during automation to prevent out‑of‑range writes.
- Values-only enforcement for chat plans: block `set_title`/`set_formula` when the prompt explicitly requires `set_values` only; clear biasing Title hints in such cases.
- Selection sanitization handles ragged `set_values` arrays by intersecting per row and compacting rows with no overlap; prevents header duplication by dropping a header-echo first row.
- AllowedCommands + WritePolicy (new): pass explicit allowed command types and per‑selection write policies (writable columns, input‑column rules) through AIContext to the planner; filter disallowed commands post‑parse.
- Schema in context (new): include a compact, typed schema for the target selection (column headers, exact per‑row width) to reduce ambiguity and improve plan shape fidelity.
- Append‑only revise guidance — IMPLEMENTED: revision prompts include row bounds and explicit input‑column append‑only wording (helps Test 18).
- Typed schema hint for Transliteration — IMPLEMENTED: when a column name contains “Transliteration”, include “latin alphabet” in the schema preview to improve content fidelity (helps Test 16).

## File Format Versioning
- Introduced `formatVersion` (1) for workbook JSON; loader auto-detects workbook vs single-sheet for backward compatibility.
- Document forward compatibility and migration guidance as formats evolve.

## Risks & Mitigations
- UI reentrancy around tabs and async I/O → guards and thorough event handling.
- Performance regressions from repaint → measure vs. full‑sheet fallback.
- Format migrations → explicit versioning and migration helpers.

## Low‑Hanging Fruit (quick wins, high perceived value)

### Freeze Panes — IMPLEMENTED
View menu adds Freeze Top Row / Freeze First Column toggles, wiring DataGridView’s `Frozen` properties so headers remain visible during scroll.

### Formula Engine: COUNTA + IS* Functions — IMPLEMENTED
Added `COUNTA`, `ISBLANK`, `ISNUMBER`, and `ISTEXT` to the function dispatch, improving parity and enabling richer AI‑generated formulas.

### Explain Cell (AI right‑click action) — IMPLEMENTED
Right‑click or AI > Explain Cell… builds a prompt from the active cell’s raw contents and evaluated value, then opens the docked Chat pane (or dialog fallback) with auto‑plan. Planner is gated to no‑write commands for a plain‑language explanation; no grid writes occur.

### AI Action Log (session history) — IMPLEMENTED
Session‑scoped log of applied AI plans (timestamp, prompt, command count, cell count, summary). Accessible via AI > View Action Log… and rendered in a simple list view.

### Schema Fill Hotkey + Selection Heuristic — IMPLEMENTED
Added Smart Schema Fill (Ctrl+Shift+F). When a single cell is selected, it auto‑expands to the likely output rectangle using header/input detection before invoking schema fill.

## Backlog (Triage Pool)
- Advanced number formats (patterns), more functions, keyboard/selection refinements.
- Nice‑to‑have: XLSX interop, charts (out of scope unless prioritized).

## Notes
- Derived from enhancement plan and dev journal; we’ll refine scope and acceptance criteria together.
