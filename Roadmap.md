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
- **UI modernization (low-cost visual polish):**
  - Grid: flat borders, modern selection accent, subtle alternating row colors, single-horizontal cell borders.
  - Headers: flat-styled column/row headers with custom background instead of 3D sunken default.
  - Double-buffered DataGridView to eliminate scroll/resize flicker.
  - Consistent default cell font (Segoe UI 9pt).

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
- UX: docked side panel with chat; messages generate a dry‑run “plan” of commands; user reviews a diff/summary and applies all or step‑by‑step.
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
- **E2E Test Suite**: 15 `.workbook.json` files in `tests/` covering AI commands, formula engine, undo/redo, I/O, and UX flows. Each file's cell A1 contains step-by-step instructions. Accessible via Test > Test Runner menu or File > Open Workbook. See `tests/TEST_INDEX.md`.

## File Format Versioning
- Introduced `formatVersion` (1) for workbook JSON; loader auto-detects workbook vs single-sheet for backward compatibility.
- Document forward compatibility and migration guidance as formats evolve.

## Risks & Mitigations
- UI reentrancy around tabs and async I/O → guards and thorough event handling.
- Performance regressions from repaint → measure vs. full‑sheet fallback.
- Format migrations → explicit versioning and migration helpers.

## Backlog (Triage Pool)
- Advanced number formats (patterns), more functions, keyboard/selection refinements.
- Nice‑to‑have: XLSX interop, charts (out of scope unless prioritized).

## Notes
- Derived from enhancement plan and dev journal; we’ll refine scope and acceptance criteria together.
