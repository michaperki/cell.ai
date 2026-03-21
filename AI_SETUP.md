# AI Setup (Env‑First)

This app uses environment variables (or a local `.env` file) to configure AI providers. No API keys are stored by the app.

## Quick Start

1) Create a `.env` file next to the running EXE (during dev, `bin/Debug/net8.0-windows/.env`).

Example `.env`:

OPENAI_API_KEY=sk-...
OPENAI_MODEL=gpt-4o-mini

# Optional Anthropic
ANTHROPIC_API_KEY=sk-ant-...
ANTHROPIC_MODEL=claude-3-haiku-20240307

# Optional limits/timeouts
ANTHROPIC_MAX_TOKENS=2048
AI_PLAN_TIMEOUT_SEC=30

# Optional custom endpoints
OPENAI_BASE_URL=https://api.openai.com/v1/chat/completions
ANTHROPIC_BASE_URL=https://api.anthropic.com/v1/messages

2) Restart the app. In AI > Settings…, choose Provider = “Auto (env)” (recommended), or explicitly “OpenAI” / “Anthropic”.

3) Use “Test Connection” to confirm provider and latency.

## Providers

- Auto (env): picks OpenAI if `OPENAI_API_KEY` is set; otherwise Anthropic if `ANTHROPIC_API_KEY`; otherwise Mock (offline).
- OpenAI: uses Chat Completions; returns `{ "cells": [[...]] }` or falls back to line‑split.
- Anthropic: uses Messages; same strict output contract.
- Mock (local): deterministic suggestions; no network.
- External API: generic POST provider exists in code for future gateway; UI configuration is deferred.

## Features

- Generate Fill (Ctrl+I): prompt → preview → accept as a single undo; respects rectangular selection (rows × cols) and never writes outside the selection.
- Inline Suggestions: ghost overlay after a short pause when continuing a list; filters out items already present above; Apply/Dismiss buttons; Enter/Tab/Double‑click to accept.
- Chat Assistant (Ctrl+Shift+C): docked right-side panel with plan → preview → apply; composite undo for multi‑command changes. This is the single chat surface; the former pop‑out window has been removed.
  - Values-only enforcement: when your prompt says “Use set_values only” (or “do not add titles”), the planner will filter out disallowed commands like `set_title` and `set_formula`, and the context will omit Title hints to avoid nudging the model to write titles.
  - Agent Loop (two‑phase): enable “Use Agent Loop (MVP)” to let the assistant first run lightweight observations (uniques/profile/sample) and then propose a plan. Use “Copy Observations” to copy the transcript.
  - Policy toggles: set Input column policy (read‑only / append‑only empty / writable) and “Selection hard mode” to strictly prevent any out‑of‑bounds writes.
  - History viewer: click “History…” to browse the last messages and export/copy the conversation as JSON.
 - Explain Cell: opens the docked Chat panel prefilled with an explanation prompt for the active cell (no‑write mode).
 - Smart Schema Fill (Ctrl+Shift+F): when a single cell is selected, auto‑expands to the likely output rectangle (headers + input column) and invokes a schema‑guided values‑only plan in the Chat panel.
 - View Action Log: AI > View Action Log… shows a session‑scoped list of applied AI plans (timestamp, prompt summary, command count, cell count, summary).

## Troubleshooting

- If Test Connection uses Mock, restart the app after setting env vars or check that `.env` is in the same folder as the EXE.
- Inline suggestions only appear when the current cell is empty and there are ≥2 non‑empty cells above it in the same column.
- We enforce plain text outputs sized to the requested shape; no formatting or formulas are written by AI.

## Notes on Limits and Timeouts
- `ANTHROPIC_MAX_TOKENS` controls the Anthropic response budget for planning/fill (default 2048 if unset).
- `AI_PLAN_TIMEOUT_SEC` sets the Chat planning timeout (default 30s); planning is cancellable.

## Notes on Two‑Phase Planning
- In the first phase, the planner returns `{"queries": [...]}` only; the host executes them and builds an Observations transcript.
- In the second phase, the planner receives the original prompt plus Observations and returns `{"commands": [...]}`. All writes are still selection‑bounded and subject to the chosen policy and hard‑mode fencing.

## Telemetry & Debugging (proposed)

To help understand model cost and behavior, we plan to surface lightweight telemetry and optional debug logging across all AI entry points (Generate Fill, Inline, Chat/Planner, Agent Loop).

- Counters (UI): show per‑request latency, model name, token usage (prompt/input, completion/output, total), and estimated remaining context window if known for the model.
- Chat session: expose a “History…” view to browse the rolling conversation (last 10 messages kept by the app) and export as JSON.
- Action Log: augment entries with model and token counts when available; keep the current summary of command/cell counts.
- File logging (opt‑in): when enabled, write JSON Lines logs for each request under a `logs/ai/` folder with fields like timestamp, surface (`chat|generate_fill|inline|agent_phase1|agent_phase2`), provider/model, token usage, prompts (truncated), plan summary, selection bounds, and write counts.

Environment toggles (planned):
- `AI_DEBUG_LOG=1` — enable per‑request JSONL logging to `logs/ai/`.
- `AI_DEBUG_DIR=...` — override the default log directory.
- `AI_DEBUG_PROMPT=1` — include full prompts in logs (otherwise prompts are truncated to protect privacy).

Implementation notes:
- OpenAI Chat Completions and Anthropic Messages both return usage blocks (`usage.prompt_tokens`/`completion_tokens`, or `usage.input_tokens`/`output_tokens`). We will parse these in the provider and propagate usage to the Chat/Generate Fill UIs and Action Log.
- For models without explicit usage, approximate tokens by characters/4 with a note that it is an estimate.
- Context window remaining = `model_context_limit - prompt/input tokens` when the model’s limit is known; expose an env override (e.g., `OPENAI_CONTEXT_TOKENS`) for custom endpoints.
