

I just tested the new UI and found an important issue in **Ask** mode.

The main problem:

* I entered **“What game is this from?”**
* The selected values were things like **Knight, Victory Point, Monopoly, Year of Plenty, Road Building**
* I expected a visible natural-language answer like **“These are development cards from Catan.”**
* Instead, there was **no visible AI answer**, and the preview area showed something misleading like:

  * **“AI will write 5 cell(s)”**
  * the existing selected values in a table
  * while also saying **“Ask mode: no writes (queries only)”**

So Ask mode currently feels broken or at least badly surfaced.

What I want changed:

1. **Ask mode needs its own response surface**

* In Ask mode, the main output should be a visible **answer card / response panel**
* Not a write preview card
* The answer should be readable in the main pane without opening history or advanced logs

2. **Do not show write-preview UI in Ask mode**

* No “AI will write N cells”
* No target-range write preview table
* No rendering of selected values as though they are proposed writes

3. **Ask mode should support plan-less/query-only responses**

* If the AI is answering a question, we should allow a response even when there are no commands to apply
* The primary artifact in Ask mode is the **textual answer**, not an AIPlan write preview

4. **Selection can still be shown as context, but clearly**

* It is fine to show:

  * “Using current selection B1:B5”
  * maybe a tiny context summary
* But that must be visually separate from the AI’s actual answer

My guess about the bug:

* Either Ask mode is still flowing through the normal write-preview path
* Or the preview renderer is reusing selected-cell contents / prior plan data
* Or the planner is returning something that the UI is misinterpreting as writes even though Ask is query-only

Please directly inspect and fix the code path for **Ask** mode in the panel:

* planner result handling
* preview rendering
* no-command/query-only behavior
* any stale preview state carried over from Fill/Append/Transform

Desired behavior for this exact case:

* Prompt: **“What game is this from?”**
* Selection: the five card names
* Result shown in-pane: something like
  **Answer:** These are development cards from *Catan*.
* No Apply button
* No write-count summary
* No fake table of writes

This is the highest-priority follow-up from my test, because it gets at the core distinction between **Ask** and the write-oriented modes.
