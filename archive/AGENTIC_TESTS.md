
Yes — instead of chasing the latest failure, build a **capability test suite**.

For your “Codex-like data analyst,” I’d test these abilities:

## 1. Table creation

Can it create clean structured tables from vague prompts?

Examples:

* “Create a weekly meal prep plan with calories, protein, cost, and justification.”
* “Make a 10-row CRM table of fake startup leads with company, contact, stage, and next step.”
* “Create a study plan for Hebrew for 4 weeks.”

What you’re testing:

* picks sensible headers
* chooses reasonable shape
* places table intelligently
* does not over-anchor to selection

## 2. Table extension

Can it understand “add onto this existing thing”?

Examples:

* “Add a column with roots.”
* “Add a justification column.”
* “Add confidence scores.”
* “Add a short summary column.”

What you’re testing:

* detects nearby table
* adds adjacent columns instead of writing somewhere random
* preserves row alignment
* understands follow-up turns

## 3. Transformation

Can it rewrite existing data in place or beside it?

Examples:

* “Uppercase these values.”
* “Translate this column to English.”
* “Normalize these dates.”
* “Extract first names.”
* “Turn these job titles into short categories.”

What you’re testing:

* in-place vs adjacent transform judgment
* schema preservation
* row-by-row consistency

## 4. Question answering over sheet context

Can it inspect the sheet and answer correctly?

Examples:

* “What game are these pieces from?”
* “What pattern do you notice in this list?”
* “Which rows look duplicated?”
* “What’s missing from this dataset?”
* “Summarize this table in plain English.”

What you’re testing:

* uses sheet as context
* gives visible answer
* references relevant data, not stale convo junk

## 5. Multi-turn memory / conversational continuity

Can it maintain a useful thread?

Examples:

* “Create a list of chess pieces.”
* “Add a column with point values.”
* “Now sort strongest to weakest.”
* “Which two are most valuable?”
* “What was my previous request?”

What you’re testing:

* remembers prior turns
* understands follow-up references
* doesn’t lose track of sheet state

## 6. Destination selection

This seems especially important for you.

Examples:

* Blank sheet selected: “Create a 7-day workout plan.”
* Cursor inside existing table: “Add a notes column.”
* Cursor far away from table: “Add Hebrew roots.”
* Small selected block: “Make a fresh budget table.”

What you’re testing:

* when to use selection
* when to extend existing structure
* when to start a new table elsewhere
* whether “strict to selection” really matters

## 7. Safe write judgment

Can it avoid dumb/destructive writes?

Examples:

* Existing filled table: “Add a summary.”
* Selection overlaps formulas: “Clean this up.”
* Crowded sheet: “Create a new table for meal prep.”
* User asks vague follow-up near existing data.

What you’re testing:

* doesn’t overwrite unrelated data
* proposes before destructive edits
* respects protected/input columns
* distinguishes safe small writes vs larger changes

## 8. Structural spreadsheet operations

Can it do more than just dump values?

Examples:

* “Insert a column for status after B.”
* “Sort this table by date descending.”
* “Add a totals row.”
* “Create a second sheet with a cleaned version.”
* “Move duplicates to another area.”

What you’re testing:

* insert/delete/sort behavior
* awareness of headers and table bounds
* action-card clarity for structural changes

## 9. Formula intelligence

Can it behave like a spreadsheet analyst, not just a text generator?

Examples:

* “Add a total cost column.”
* “Compute conversion rate.”
* “Add a running total.”
* “Highlight rows where revenue < cost.”
* “Create a summary block with averages.”

What you’re testing:

* chooses formulas when appropriate
* fills formulas consistently
* understands relative references
* distinguishes computed vs generated text

## 10. Data cleaning

Huge category for a data analyst agent.

Examples:

* “Normalize these country names.”
* “Fix inconsistent capitalization.”
* “Split full name into first/last.”
* “Standardize phone numbers.”
* “Identify likely duplicates.”

What you’re testing:

* pattern detection
* conservative cleanup
* whether it explains uncertain cases

## 11. Classification / enrichment

Can it add useful inferred columns?

Examples:

* “Classify these companies by industry.”
* “Mark these leads as hot/warm/cold.”
* “Estimate sentiment.”
* “Add part of speech.”
* “Group these expenses into categories.”

What you’re testing:

* contextual reasoning
* row-level consistency
* confidence / uncertainty handling

## 12. Ambiguity handling

Can it ask or choose well when prompt is unclear?

Examples:

* “Fix this.”
* “Make this better.”
* “Clean this up.”
* “Add the important stuff.”
* “Organize this.”

What you’re testing:

* asks follow-up when needed
* makes reasonable assumptions when safe
* doesn’t hallucinate a totally arbitrary operation

## 13. Preview / proposal quality

Even when the write is correct, is the proposal understandable?

Examples:

* large fresh table
* adjacent column add
* structural insert + populate
* destructive transform

What you’re testing:

* card explains intended result clearly
* header row visible
* destination legible
* user can tell whether to approve

## 14. Boundary / stress tests

Make sure it doesn’t fall apart at edges.

Examples:

* empty sheet
* single-cell selection
* giant selected block
* sparse sheet with islands of data
* merged-like visual patterns
* formulas mixed with values
* partially filled columns

What you’re testing:

* robustness
* graceful fallback
* no bizarre anchoring behavior

## 15. Analyst-style reasoning tasks

This is the “Codex-like data analyst” sweet spot.

Examples:

* “Given this sales table, what are the top three insights?”
* “Add a risk score per row and explain your logic.”
* “Find anomalies.”
* “Which rows need manual review?”
* “Make an executive summary block above this table.”

What you’re testing:

* reasoning over structured data
* mixing analysis + action
* ability to produce both answers and spreadsheet changes

---

## A very useful benchmark format

For each test, log:

* **Prompt**
* **Initial sheet state**
* **Selection**
* **Expected behavior**
* **Actual behavior**
* **Pass / fail / partial**
* **Failure type**

Failure types could be:

* wrong destination
* wrong shape
* wrong headers
* failed follow-up memory
* destructive write
* bad proposal clarity
* answered instead of acting
* acted instead of answering
* overfit to selection
* ignored existing table

---

## 10 high-value starter tests

If you want a lean initial suite, start here:

1. Create a fresh structured table on blank sheet
2. Add adjacent column to existing table
3. Transform one existing column
4. Ask a question about selected data
5. Multi-turn follow-up that extends prior work
6. Create new table without overusing current selection
7. Sort / totals / structural edit
8. Add formula column
9. Clean messy real-world data
10. Detect duplicates / anomalies and summarize findings

The deepest thing you want to test is probably this:

**Can the agent choose the right action type and the right destination from conversational context, without being micromanaged by modes or selection?**

That’s the heart of the product.

I can turn this into a proper `TESTS.md` with sections and numbered cases.
