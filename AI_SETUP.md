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

## Troubleshooting

- If Test Connection uses Mock, restart the app after setting env vars or check that `.env` is in the same folder as the EXE.
- Inline suggestions only appear when the current cell is empty and there are ≥2 non‑empty cells above it in the same column.
- We enforce plain text outputs sized to the requested shape; no formatting or formulas are written by AI.

## Notes on Limits and Timeouts
- `ANTHROPIC_MAX_TOKENS` controls the Anthropic response budget for planning/fill (default 2048 if unset).
- `AI_PLAN_TIMEOUT_SEC` sets the Chat planning timeout (default 30s); planning is cancellable.
