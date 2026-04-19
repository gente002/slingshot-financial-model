# The Three God Rules

**Portable edition — paste into the top of any project's CLAUDE.md or equivalent guidance doc.**

Three non-negotiable constraints on every design decision. Not guidelines — constraints. If a proposed design violates any one of them, it is wrong, regardless of how elegant or efficient it looks. Every feature, every module, every workflow must satisfy all three.

---

## Rule 1 — Never Get Stuck

**Principle.** The system always produces output of the expected shape. When an input is missing, ambiguous, or bad, the system substitutes the best available assumption and proceeds. Quality problems isolate individual records; they do not halt the run.

**Operational test.** Can this code path raise an unrecoverable error when a user runs it against imperfect real-world data? If yes, redesign. Every branch must have a defined fallback that yields a valid output with the same shape downstream consumers expect.

**Common misreading to avoid.** This is *not* "swallow all errors silently." Assumed values must be tagged as assumptions (see Rule 3) and logged. The rule is about *shape invariance under input variance*, not about hiding problems.

**Typical applications.** Data ingestion with missing fields, config loading with absent entries, computations on partial inputs, report generation against incomplete periods.

---

## Rule 2 — Never Lose Control

**Principle.** Users have a sanctioned path to override any value at any layer of the system — temporarily, with tracking, without breaking anything downstream. Automation sets defaults; humans retain authority.

**Operational test.** Pick any value the system computes or assumes. Can a user change it — at the record level, in bulk, or as a top-side aggregate adjustment that allocates down — without editing code, without corrupting state, and with the change visibly attributed to them? If no, the layer is missing an override surface.

**Common misreading to avoid.** This is *not* "expose every internal as a knob." The rule is about *sanctioned* paths — explicit override cells, adjustment tables, configurable assumptions — not about making everything mutable. Overrides are tracked, not silent. Recurring overrides should graduate into permanent configuration.

**Typical applications.** Input cells on calculation sheets, adjustment ledgers on pipelines, `actual vs assumption` flags on every value, rerun-safe edits to intermediate state.

---

## Rule 3 — Never Work Alone

**Principle.** Every workflow that produces something consequential is a multi-party collaboration: preview before commit, annotate values with comments or threads, converge on truth iteratively, confirm explicitly. Even a solo operator is collaborating with their future self and with downstream consumers.

**Operational test.** Before a change lands, can a second party (human reviewer, downstream system, future reader) see (a) *what* will change, (b) *why* it's changing, (c) the *downstream impact*, and (d) who approved it? Can they leave feedback tied to specific values without blocking the whole run?

**Common misreading to avoid.** This is *not* "require sign-off on everything." It is *not* "add chat to the app." The rule is about *making work visible and annotatable at the granularity of the value that matters*, so disagreement can be surfaced precisely rather than as vague complaints after the fact.

**Typical applications.** Dry-run / preview modes, per-cell or per-row comments, diff views on assumption changes, materiality thresholds that flag changes big enough to require review, convergence metrics (% of values that are actual vs assumed).

---

## How the rules compose

The rules interlock. Rule 1 guarantees output exists; Rule 2 guarantees that output is correctable; Rule 3 guarantees correction is visible and deliberate. A design that satisfies one by violating another — e.g., "never get stuck" by silently overwriting user overrides — is not a valid design. When rules appear to conflict, the usual resolution is to add a *layer* (a tracked assumption, a preview step, an annotation) rather than to pick one rule over another.

**When evaluating any proposed change, state explicitly how it satisfies each of the three rules. If you can't, the change isn't ready.**

---

## Origin

These rules originated in the Phronex Rapid Development Kit (RDK) — a framework for building Excel/VBA financial models with CSV persistence. The original canonical text lives in that project's `docs/RDK_Developer_Flywheel_v1.2.md`, written against a specific domain (insurance modeling, kernel/domain separation, multi-agent build process). This document is the domain-agnostic extract, suitable for transfer to any project where automation produces output consumed by humans who need to trust, override, and audit it.
