
Yes — that makes a lot of sense.

Once you have an auto-improvement loop, you need a **developer control tower** or else it just feels like random activity.

The dashboard should answer:

* Are we actually getting better?
* Which capabilities are improving?
* Which ones keep failing?
* Did a code change help or hurt?
* What should we work on next?

I’d think of it in 5 sections:

## 1. Overall score / trend

Top-level cards:

* total tests run
* pass rate
* partial rate
* regression count
* rolling score over time
* last run / last green run

This tells you whether the system is improving or thrashing.

## 2. Capability matrix view

Map results to the abilities you care about:

* table creation
* table extension
* transformation
* Q&A over sheet
* multi-turn continuity
* destination selection
* safe write judgment
* structural ops
* formula intelligence
* data cleaning
* enrichment/classification
* ambiguity handling
* preview quality
* boundary/stress
* analyst reasoning

For each:

* coverage count
* current pass rate
* trend
* most recent failures

This is probably the most important view.

## 3. Failure patterns

Aggregate failure types like:

* overfit to selection
* wrong destination
* wrong shape
* destructive write
* answered instead of acting
* acted instead of answering
* bad follow-up memory
* unclear proposal card
* wrong headers
* hidden gating / policy issue

You want to see:

* top recurring failures
* new failures introduced this week
* failures fixed recently

## 4. Run explorer

A per-run view:

* run id
* commit / branch
* provider / model
* suite(s) run
* pass/fail counts
* linked artifacts
* reflection summary

Click into a run and inspect:

* test case
* prompt
* sheet context
* expected vs actual
* plan json
* observations
* screenshots if available

## 5. Improvement loop / queue

This is the action layer:

* biggest current bottlenecks
* candidate fixes
* tests to add next
* flaky tests
* missing coverage areas
* recent code changes correlated with regressions

This is what turns the dashboard from reporting into guidance.

A really useful distinction:

* **capability progress**
* **product UX progress**
* **infrastructure progress**

Because otherwise “we improved” becomes muddy.

For example:

* capability progress = destination selection pass rate rose from 30% to 55%
* UX progress = unified thread now preserves history
* infrastructure progress = runner now emits step-level artifacts and scorecards

You probably want all three.

A strong MVP dashboard might just have:

* pass rate over time
* capability matrix heatmap
* recurring failure types
* recent runs table
* “what regressed / what improved” panel

That alone would already be valuable.

Most important design principle:
this should not be a vanity dashboard. It should help you decide:

* what to test next
* what to fix next
* whether a change actually helped

So yes — I think this is a very natural next piece for the system. It gives you visibility into whether the auto-improvement loop is producing real progress or just noise.

I can sketch the exact layout for a v1 dashboard.



Yes — that makes a lot of sense.

Once you have an auto-improvement loop, you need a **developer control tower** or else it just feels like random activity.

The dashboard should answer:

* Are we actually getting better?
* Which capabilities are improving?
* Which ones keep failing?
* Did a code change help or hurt?
* What should we work on next?

I’d think of it in 5 sections:

## 1. Overall score / trend

Top-level cards:

* total tests run
* pass rate
* partial rate
* regression count
* rolling score over time
* last run / last green run

This tells you whether the system is improving or thrashing.

## 2. Capability matrix view

Map results to the abilities you care about:

* table creation
* table extension
* transformation
* Q&A over sheet
* multi-turn continuity
* destination selection
* safe write judgment
* structural ops
* formula intelligence
* data cleaning
* enrichment/classification
* ambiguity handling
* preview quality
* boundary/stress
* analyst reasoning

For each:

* coverage count
* current pass rate
* trend
* most recent failures

This is probably the most important view.

## 3. Failure patterns

Aggregate failure types like:

* overfit to selection
* wrong destination
* wrong shape
* destructive write
* answered instead of acting
* acted instead of answering
* bad follow-up memory
* unclear proposal card
* wrong headers
* hidden gating / policy issue

You want to see:

* top recurring failures
* new failures introduced this week
* failures fixed recently

## 4. Run explorer

A per-run view:

* run id
* commit / branch
* provider / model
* suite(s) run
* pass/fail counts
* linked artifacts
* reflection summary

Click into a run and inspect:

* test case
* prompt
* sheet context
* expected vs actual
* plan json
* observations
* screenshots if available

## 5. Improvement loop / queue

This is the action layer:

* biggest current bottlenecks
* candidate fixes
* tests to add next
* flaky tests
* missing coverage areas
* recent code changes correlated with regressions

This is what turns the dashboard from reporting into guidance.

A really useful distinction:

* **capability progress**
* **product UX progress**
* **infrastructure progress**

Because otherwise “we improved” becomes muddy.

For example:

* capability progress = destination selection pass rate rose from 30% to 55%
* UX progress = unified thread now preserves history
* infrastructure progress = runner now emits step-level artifacts and scorecards

You probably want all three.

A strong MVP dashboard might just have:

* pass rate over time
* capability matrix heatmap
* recurring failure types
* recent runs table
* “what regressed / what improved” panel

That alone would already be valuable.

Most important design principle:
this should not be a vanity dashboard. It should help you decide:

* what to test next
* what to fix next
* whether a change actually helped

So yes — I think this is a very natural next piece for the system. It gives you visibility into whether the auto-improvement loop is producing real progress or just noise.

I can sketch the exact layout for a v1 dashboard.
