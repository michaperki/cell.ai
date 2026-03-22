

Your current suite is messy, but it actually contains **different layers that should stay separate**:

* engine / product correctness
* AI command-planning behavior
* agent-loop behavior
* UI / unified-thread behavior

That separation is visible in the current index already: Engine, UX, I/O, AI, AI (Agent), UI/AI are all mixed in one place, which is the real problem. 

So the issue is less:

> “we have too many kinds of tests”

and more:

> “we don’t have a clean test taxonomy.”

Also, in `TEST_SPECS.json`, some tests are classic single-step AI command tests, some are multi-turn session tests, and some are true agent-loop tests with observe/apply phases. Those are not the same thing and should not be judged by the same standard. 

My recommendation:

## Keep 4 buckets

**1. Deterministic product tests**

* formula engine
* undo/redo
* workbook I/O
* copy/paste
* formatting
* future pivot-table engine behavior

These should be as non-AI as possible.

**2. AI command tests**

* can the planner choose the right command(s)?
* `set_values`, `set_formula`, `sort_range`, etc.
* selection fencing, no-write zones, schema fills

These test “can the assistant drive the spreadsheet primitives?”

**3. Agent loop tests**

* observe
* propose
* gated apply
* reflection
* artifact creation
* bounded iteration

These are the real “agentic loop” tests.

**4. Scenario / benchmark tests**

* analyst-style end-to-end jobs
* messy workbook + business request
* maybe uses pivot table, cleanup, summary, chart, explanation

These are the closest thing to “AI data analyst” tests.

## What I would change

I’d stop organizing around a single flat numbered list.

Instead make the suite feel like:

* `engine_*`
* `command_*`
* `agent_*`
* `scenario_*`

Right now `test_29` through `test_34` feel like the nucleus of the agent-loop suite already. 

So I would **promote those into a first-class suite**, not collapse everything into them.

## What I would probably remove

A few of the current AI tests look more like feature smoke tests than durable benchmarks. Those can stay, but they should not dominate your thinking.

The danger of “10 loop tests only” is:

* you lose coverage of core spreadsheet correctness
* you lose feature-specific regressions
* failures become harder to localize
* you stop knowing whether the loop failed or the product failed

That would make the system feel more agentic while actually making it harder to improve.

## Better mental model

Think of it like this:

* **product tests**: is the substrate correct?
* **command tests**: can AI use the substrate?
* **loop tests**: can the agent reason over tasks safely?
* **scenario tests**: does the whole thing resemble a useful analyst?

That is the ladder.

## Concrete restructure I’d do

Keep the existing content, but split into folders/suites like:

* `tests/engine/`
* `tests/command/`
* `tests/agent/`
* `tests/scenario/`

And maintain separate manifests:

* `ENGINE_SPECS.json`
* `COMMAND_SPECS.json`
* `AGENT_SPECS.json`
* `SCENARIO_SPECS.json`

Then your headline benchmark becomes maybe:

* ~15 deterministic core tests
* ~20 command tests
* ~10 agent-loop tests
* ~10 scenario tests

That is much healthier than replacing everything with 10 loop tests.

## One more important point

You should also define what “pass” means per bucket:

* engine: exact correctness
* command: correct operation/targeting/constraints
* agent: good observations, safe proposals, bounded writes, artifacts
* scenario: acceptable analyst outcome, maybe rubric-based not exact-match

That prevents the suite from turning into mush.

My blunt take: **don’t scrap; refactor the taxonomy.**
Your current suite is annoying because it is all in one pile, not because the categories are wrong.
