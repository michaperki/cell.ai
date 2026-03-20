# Execution Plan (v0.3 focus + AI UX)

This plan prioritizes reliability, UX polish, and small-but-high‑leverage AI improvements. It is sequenced to land safe infrastructure first, then visible UX.

## Goals
- Ship safe incremental-on-edit repaint with robust dependency extraction.
- Improve copy/paste semantics for formulas (absolute vs relative refs).
- Extend async I/O to CSV; maintain responsive UI with progress/guards.
- Improve Chat UX: revise loop, formula auto-route, error feedback.
- Raise provider limits/config and enrich planner context (header detection).

## Phase 1 — Foundations & Performance (Status: DONE)

1) Incremental-on-edit repaint (safe path) — DONE
- Tasks
  - Reuse the existing dependency graph and incremental repaint helper for direct edits.
  - Gate with a reliability switch: attempt `RecalculateDirty` + `RefreshDirtyOrFull`, fall back to full refresh when affected >5% of grid.
  - Add fast `grid.CommitEdit`/`CellValidated` experimentation to avoid stale visuals.
- Acceptance
  - Edits to a cell update its dependents immediately without full-sheet repaint in typical cases.
  - No stale cells after editing chains (A2→A3 scenarios). Full fallback is triggered automatically when needed.
- Risks / Mitigations
  - DataGridView repaint quirks post-commit → keep full-refresh fallback and explicit `grid.Refresh()`.

2) Dependency extraction robustness (parser-driven) — DONE
- Tasks
  - Expose a public reference enumeration from the formula parser (AST), or reuse a shared tokenizer used by both `FormulaEngine` and `Spreadsheet.SetRaw`.
  - Replace ad‑hoc scanning in `Spreadsheet.ExtractReferences` with the shared parser walker.
- Acceptance
  - Correctly extracts references from functions, nested expressions, and ranges; ignores string literals.
  - Regression tests cover typical refs and ranges.
- Risks
  - Parser exposure scope creep → keep API minimal (refs only), no evaluation coupling.

## Phase 2 — UX & Correctness (Status: DONE)

3) Copy/Paste: absolute/relative refs — DONE
- Tasks
  - When pasting formulas, rewrite cell tokens by delta (row/col) honoring `$` anchors.
  - Keep plain text/values unchanged; preserve `$` in output.
- Acceptance
  - Copy/paste a block of formulas reproduces spreadsheet‑standard behavior for `$A$1`, `$A1`, `A$1`, and `A1`.
  - Unit tests: 2×2 formula block moved in four directions.

4) Async CSV I/O + guards — DONE
- Tasks
  - Add async wrappers for CSV export/import; use `SetUiBusy` and disable menus during operations.
  - Optionally show a simple status label for long imports.
- Acceptance
  - CSV open/save does not block the UI; cancel-safe and error dialogs on failure.

## Phase 3 — AI UX Improvements (Status: DONE)

5a) Docked Chat Pane — DONE
- Tasks
  - Extract ChatAssistantForm internals to a reusable UserControl and host it in a right-docked panel with a splitter.
  - Toggle via menu and Ctrl+Shift+C; persist width/visibility in settings; keep pop-out window as optional.
- Acceptance
  - Chat is available as a docked pane without blocking workflows; users can continue editing while planning and applying.

5) Chat “Revise” loop — DONE
- Tasks
  - Add a “Revise” button to `ChatAssistantForm` to append feedback into `_history` and re‑plan.
  - Keep current plan preview visible until replaced.
- Acceptance
  - Users can iteratively refine a plan without closing the dialog.

6) Auto-route formulas in set_values — DONE
- Tasks
  - In `ApplyPlan`, if a `set_values` cell begins with `=`, route it through the formula write path (or split into a `SetFormulaCommand` internally).
  - Preserve composite undo and repaint behavior.
- Acceptance
  - AI plans that mistakenly put formulas in `set_values` still produce working formulas in the grid.

7) Post-apply error feedback loop — DONE
- Tasks
  - After apply, scan changed cells for `#ERR` results; if any, surface a prompt “Some cells errored — try to fix?” that seeds a follow‑up chat prompt with context (locations + formulas + messages).
- Acceptance
  - One-click path to repair attempts; no auto‑writes without user approval.

8) Provider limits/config — DONE
- Tasks
  - Raise Anthropic `max_tokens` to 2048 (env/config driven) to reduce truncation.
  - Make planning timeout configurable (default 30s already applied).
- Acceptance
  - Long JSON plans (tables + formulas) no longer truncate under normal cases.

9) Header detection heuristic for planner context — DONE
- Tasks
  - Identify header row as the first non-empty row with predominantly text values; include that row’s values in workbook summary.
- Acceptance
  - Planner aligns better with real header rows when row 1 is a title.

## Phase 4 — Command Grammar Extensions (stretch)
- Insert/Delete rows/cols: shift data and adjust dependency graph; schema: `insert_rows/delete_rows`.
- Set format: apply bold/color/number format to ranges using existing `CellFormat`; schema: `set_format`.
- Move/Copy range: cut/copy rectangles; respect formulas and formats (initially values-only to keep undo simple).
- Delete sheet: complement `create_sheet`/`rename_sheet`.
- Acceptance
  - Parser/serializer updated in `ProviderChatPlanner`; `ApplyPlan` covers new commands with composite undo.

## Phase 5 — Testing & Verification (In progress)
- Unit tests
  - Formula rewrite on paste (absolute/relative), dependency extraction via parser, error‑free incremental-on-edit for small graphs.
- E2E tests
  - Add steps to `TEST_SPECS.json` for copy/paste refs and “Revise” flow.
  - Keep snapshots + plan/prompt dumps for debugging.

## Milestones & Order
1) Phase 1.1 (incremental-on-edit) → 1.2 (parser refs)
2) Phase 2.3 (paste semantics) → 2.4 (CSV async)
3) Phase 3.5–3.9 (AI UX tranche)
4) Phase 4 (grammar), timeboxed per command in order: insert/delete rows/cols → set_format → delete_sheet → move/copy range.
5) Phase 5 (tests) runs alongside each phase; add E2E after each tranche.

## Rollback & Observability
- Keep the full-refresh path behind a simple switch; log fallback rates to detect regressions.
- Preserve mock provider/planner paths to validate offline.

## Out-of-Scope (for this pass)
- XLSX interop and packaging/telemetry (tracked in Roadmap v0.4).
- Advanced number formats and conditional formatting rules.
