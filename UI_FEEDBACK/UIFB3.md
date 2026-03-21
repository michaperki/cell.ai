

Use the new session-model framing as the guiding principle for v2.

Please proceed on `ui-panel-v2`, with these decisions:

* Yes: hard-scope the right pane to exactly **Prompt → Preview → Apply**
* Yes: move everything else into an **Advanced / Inspector** drawer
* Yes: replace the terminal-style log with a clean **preview card**
* Yes: add explicit **mode selection**
* Yes: tighten copy throughout
* Yes: keep detailed history/debugging in a separate dialog, not in the main pane

A few important refinements:

**1. Optimize for a session-based AI model, not a generic chatbot**
The core UX problem is not just visual clutter — it is unclear session/context behavior.
The UI should make it obvious whether the user is in an ongoing AI session and what context is active.

For this pass, please add a lightweight session/status line near the prompt, for example:

* “Using current selection + headers”
* “Using prior AI context”
* “Fresh request”

Even if the underlying behavior is still simple, I want the UI moving toward explicit context visibility.

**2. Mode selector**
Prefer explicit chips / segmented control over pure auto-detect.

Use something like:

* Fill
* Append
* Transform
* Ask

Auto-detection can still happen under the hood, but the visible mode should be explicit and user-legible.

**3. Preview card**
Yes, replace the terminal pane with a card showing:

* what AI plans to do
* target range
* write count / shape
* small table preview
* warnings if relevant

Execution logs should not dominate the primary workflow.

**4. Copy changes**
Approved direction:

* “Selection hard mode” → something like **Strict to selection**
* “Let AI explore first” → **Analyze selection first**

In general, favor user language over internal language.

**5. Context menu**
Consolidate under a single **AI** submenu.
The top-level context menu should stay clean.

**6. Defaults**
Default to safer behavior:

* **Strict to selection: on**
* input columns remain protected unless the active mode explicitly allows writing there
* append behavior should be clearly scoped to append mode

**7. Theme**
Yes, use a single accent color for primary emphasis.
That sounds fine for this pass.

Most important product note:
Please design this as a **spreadsheet tool with an AI session**, not as a floating chat box. I want iterative context, but I do not want hidden context. The UI should increasingly answer:

* am I in an ongoing AI session?
* what is being used as context right now?
* what is the AI about to write?

Proceed with implementation.

---

Clarifying answers (2026-03-21)

1. Mode semantics and gating

* Yes: Fill = write into selected output cells only.
* Yes: Transform = modify/overwrite within the selection, including formulas, but still respect input-column protection unless explicitly overridden later.
* Yes on Append using `StartRow = -1` in the chat path for append-to-first-empty-row within the selected block.

2. Session status line

* Yes to:
  * “Fresh request”
  * “Using prior AI context (N)”
  * “Using current selection + headers”
* Do not add “Ignore history for this plan” yet. That adds complexity too early.
* Instead, make “Fresh request” / “Using prior AI context” clearly visible now; add per-plan history controls later if needed.

3. Preview card details

* Yes: cap table sample at 5 rows × 6 cols.
* Yes: for multi-command plans, show:
  * aggregated summary
  * first N write previews
  * human-readable summaries for structural ops like insert/delete/sort/format

4. Defaults per mode

Use these defaults:

* Fill: Strict to selection ON, Input read-only
* Append: Strict to selection ON, append-only into empty rows
* Transform: Strict to selection ON, Input read-only by default
* Ask: query-only, no Apply button

Tweak: in Ask, “Strict to selection” should not be emphasized since no writes occur. Keep it on internally, but the UI should focus on “context limited to current selection.”

5. Input column detection

Prefer inference for now, not exposed manual UI. Use the existing Smart Schema Fill-style heuristic for now if it is solid. If uncertain, surface that in the preview/inspector rather than adding more controls in the main pane.

6. Copy / labels

Approved:

* Strict to selection
* Analyze selection first
* Fill / Append / Transform / Ask

7. Persistence

* Persist last chosen mode across sessions
* Do not persist “Strict to selection” yet; keep it default-on each time for safety and predictability

8. Context menu

Confirmed structure:

* AI ▸
  * Explain Cell…
  * Smart Schema Fill…
  * Ask About Sheet…

Focus: optimize toward a spreadsheet tool with visible AI session context. The session/status line and preview card matter more than extra toggles.
