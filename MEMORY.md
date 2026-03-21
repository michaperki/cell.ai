# Memory

## Operating constraints
- Mike runs and tests the app; the agent does not.
- The app runs on Windows.
- The agent runs in WSL.
- Do not try to run the desktop app from the agent environment.
- Do not claim something was tested unless Mike explicitly tested it and reported results.
- When debugging runtime behavior, rely on Mike’s screenshots, logs, and descriptions.

## Collaboration model
- Mike is product lead and tester.
- The agent should implement, inspect code, and reason from artifacts.
- The agent should not invent verification.
- The agent should not repeatedly ask to reproduce bugs if Mike already described the issue clearly.
- Prefer direct code inspection and targeted fixes over vague debugging advice.

## Product vision
- The spreadsheet AI should feel like one unified AI collaborator, not four separate tools.
- Planning/execution may exist internally, but the user-facing experience should feel like one continuous session.
- Answers, proposed changes, and applied changes should appear inline in a unified thread.
- Avoid fragmented mode-first UX unless absolutely necessary.
- Visible context matters: the user should understand what the AI is using and what it is about to change.

## UX principles
- Reduce clutter.
- Hide internals by default.
- Prefer user language over internal/dev language.
- Do not expose implementation details unless they help the user make a decision.
- The main pane should optimize for clarity, trust, and flow.
- Distinguish clearly between:
  - discussion/answer
  - proposed spreadsheet action
  - applied result

## Agent behavior expectations
- Solve the root UX/product issue, not only the nearest symptom.
- Before implementing, restate the higher-level interaction model being targeted.
- Evaluate proposed changes against the product vision, not just local correctness.
- Avoid getting stuck in narrow mode-specific patches when the real issue is architectural.
- If there is uncertainty, surface tradeoffs explicitly.

## Spreadsheet-specific assumptions
- Safe writes matter.
- Selection boundaries and write constraints should be visible and predictable.
- Follow-up requests should account for prior sheet state and prior conversation context.
- The agent should infer natural continuations when reasonable (for example, adding an adjacent column to an existing generated table).

## Communication rules
- Be explicit about assumptions.
- Do not say something “works” unless it has actually been verified by Mike.
- When reporting progress, separate:
  - what was changed
  - what is inferred
  - what still needs user validation


* We are **not using CI**.
* Do not suggest, set up, modify, debug, or rely on CI workflows.
* Do not propose GitHub Actions, pipeline fixes, or CI-based validation.
* Validation happens through Mike’s local Windows testing, not CI.

- Periodically update `Dev Journal.md` to track what we tried, what failed, what changed, and what we learned.
- Important implementation attempts and product lessons should not live only in chat.
- `Memory.md` stores durable truths; `Dev Journal.md` stores the chronological record of the work.
