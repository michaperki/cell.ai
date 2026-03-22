# Capability Matrix

Tracks what the AI agent should be able to do, mapped to test coverage. Derived from the capability taxonomy in `SOUL.md`. See `TEST_INDEX.md` for full test details.

## Coverage Summary

| # | Capability | Tests | Coverage |
|---|-----------|-------|----------|
| 1 | **Table creation** — fresh structured table from vague prompt | 01, 06, 35 | Basic |
| 2 | **Table extension** — add columns/rows to existing table | 16, 17, 35 | Basic |
| 3 | **Transformation** — rewrite data in place or adjacent | 29, 31, 32 | Basic |
| 4 | **Q&A over sheet** — inspect and answer without writing | 30, 33, 36 | Good |
| 5 | **Multi-turn continuity** — maintain useful thread across turns | 06, 16, 17 | Basic |
| 6 | **Destination selection** — choose where to write based on context | 07, 14, 15, 35 | Basic |
| 7 | **Safe write judgment** — avoid destructive or out-of-bounds writes | 18, 34 | Basic |
| 8 | **Structural operations** — insert/delete/sort/copy/move | 03, 23, 24, 25 | Good |
| 9 | **Formula intelligence** — use formulas when appropriate | 02, 19, 22 | Basic |
| 10 | **Data cleaning** — normalize, deduplicate, standardize | 29, 32 | Minimal |
| 11 | **Classification / enrichment** — add inferred columns | 16, 21 | Minimal |
| 12 | **Ambiguity handling** — ask or choose well when prompt is unclear | — | None |
| 13 | **Preview / proposal quality** — clear, understandable action cards | 35, 36 | Minimal |
| 14 | **Boundary / stress** — empty sheets, single cells, sparse data | — | None |
| 15 | **Analyst reasoning** — insights, anomalies, summaries | — | None |

**Overall: 12/15 categories have at least one test. 3 gaps (ambiguity, boundary, analyst reasoning).**

## What each test exercises

Key: `C` = creation, `E` = extension, `T` = transform, `Q` = Q&A, `M` = multi-turn, `D` = destination, `S` = safe write, `St` = structural, `F` = formula, `Cl` = cleaning, `En` = enrichment

| Test | Capabilities | Notes |
|------|-------------|-------|
| 01 | C | set_values, composite undo |
| 02 | F | set_formula SUM/AVERAGE |
| 03 | St | sort_range descending |
| 06 | C, M | 3-step multi-turn build |
| 07 | D | Context-aware formula placement |
| 14 | D | Multi-command plan |
| 15 | D | Cross-sheet reasoning |
| 16 | E, En, M | Hebrew morphology fill, multi-turn append |
| 17 | E, M | Append-only multi-turn |
| 18 | S | No-write input column protection |
| 19 | F | Formula autoroute |
| 21 | En | Schema-driven single-shot fill |
| 22 | F | IFERROR, SUMIF, COUNTIF, AVERAGEIF |
| 23 | St | insert/delete cols |
| 24 | St | delete sheet |
| 25 | St | copy/move range with formula rewrite |
| 26 | — | set_format (visual only) |
| 27 | — | set_validation |
| 28 | — | set_conditional_format |
| 29 | T, Cl | Agent city cleanup |
| 30 | Q | Observe-only, transcript |
| 31 | T | Values-only gating |
| 32 | T, Cl | transform_range normalize |
| 33 | Q | Observe-only strict |
| 34 | S | Selection fencing |
| 35 | C, E, D | Unified thread: create then extend |
| 36 | Q | Answer-only, no writes |

## Priority gaps to fill

1. **Ambiguity handling** — test with vague prompts ("fix this", "make this better") and verify the agent asks for clarification or makes safe assumptions
2. **Boundary / stress** — empty sheet, single cell, giant selection, sparse data islands, formulas mixed with values
3. **Analyst reasoning** — "what are the top insights?", "find anomalies", "which rows need review?"
4. **Data cleaning depth** — dedup, phone normalization, date standardization, name splitting
5. **Multi-turn memory at depth** — 5+ turn conversations with interleaved writes and questions

## Failure types to track (see `FAILURE_LOG.md`)

- wrong destination
- wrong shape / ragged rows
- wrong headers / header echo
- failed follow-up memory
- destructive write / out-of-bounds
- bad proposal clarity
- answered instead of acting
- acted instead of answering
- overfit to selection
- ignored existing table
- unexpected command type
