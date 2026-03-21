# OfficeIMO Markdown Correctness Backlog

This document turns the correctness roadmap into issue-sized work.

It is meant to answer:

- what we should do next
- what order we should do it in
- how we know a change is done

Use this together with `officeimo.markdown.correctness-roadmap.md`.

## Execution Rules

- Prefer correctness over breadth.
- Prefer typed AST and syntax-tree work over renderer-side recovery.
- Prefer isolated, reviewable changes over broad mixed refactors.
- Do not couple generic work to IntelligenceX-only behavior.
- Every backlog item should land with tests.

## Suggested Order

1. Tree invariants and source mapping
2. Canonical AST cleanup
3. Parser extension seams
4. Renderer and HTML-ingestion cleanup
5. Compatibility matrix and corpus coverage
6. Performance and benchmark evidence
7. Breadth expansion only after the above are stable

## Workstream A: Tree Invariants

### A1. Audit syntax-node builder association

Goal:
- ensure every semantic node that should map back to syntax does so consistently

Current audit:

- see `Docs/reviews/officeimo.markdown-syntax-association-audit-2026-03-21.md`
- see `Docs/reviews/officeimo.markdown-associated-object-hotspots-2026-03-21.md`

Done means:
- every relevant `BuildSyntaxNode` path sets the expected associated object
- missing coverage cases are fixed
- regression tests assert object-to-span lookup behavior

### A2. Add tree invariant test helpers

Goal:
- make it cheap to validate parent, sibling, root, and index consistency

Current coverage:

- shared helpers landed in `OfficeIMO.Tests/Markdown/MarkdownInvariantAssert.cs`
- representative invariant suite landed in `OfficeIMO.Tests/Markdown/Markdown_Tree_Invariant_Tests.cs`
- follow-up findings are captured in `Docs/reviews/officeimo.markdown-tree-invariant-findings-2026-03-21.md`

Done means:
- shared test helpers exist for syntax and semantic trees
- at least one focused invariant suite runs against representative documents

### A3. Add provenance-focused golden cases

Goal:
- prove that heading spans, block spans, and lookups survive realistic documents

Done means:
- golden inputs cover headings, nested lists, block quotes, tables, code fences, and footnotes
- tests assert spans and logical lookup targets, not only rendered text

## Workstream B: Canonical AST

### B1. Identify duplicated mutable node shapes

Goal:
- make the cleanup list explicit before refactoring node types one by one

Current inventory:

- see `Docs/reviews/officeimo.markdown-duplicated-node-shapes-inventory-2026-03-21.md`

Done means:
- a short inventory exists of nodes that keep both structural children and parallel text/state
- each node is tagged as low, medium, or high migration risk

### B2. Canonicalize one medium-complexity node first

Goal:
- establish the refactor pattern on one real node before applying it broadly

Suggested first target:

- `FootnoteDefinitionBlock`
- see `Docs/reviews/officeimo.markdown-footnote-canonicalization-sketch-2026-03-21.md`

Done means:
- one duplicated node shape is converted to a single primary representation
- convenience accessors remain where needed
- old behavior stays covered by tests

### B3. Define AST ownership rules

Goal:
- remove ambiguity about what is primary versus derived

Done means:
- a short design note or code comments define:
  - which collections own children
  - which text fields are derived
  - when cached computed values are acceptable

## Workstream C: Parser Extension Seams

### C1. Formalize block parser ordering rules

Goal:
- make extension behavior predictable when multiple parsers can claim similar input

Done means:
- ordering and conflict rules are documented
- tests cover precedence and fallback behavior

### C2. Design inline parser extension contracts

Goal:
- stop growing inline behavior as internal-only logic

Done means:
- a public or clearly intended contract exists for inline parser registration
- extension order and failure behavior are defined
- at least one non-trivial inline extension uses the seam

### C3. Separate parsing from normalization

Goal:
- avoid mixing repair logic into core parsing in ways that are hard to reason about

Done means:
- parser-owned behavior and normalization-owned behavior are explicitly separated
- normalization can be applied intentionally instead of implicitly everywhere

## Workstream D: Renderer And HTML Cleanup

### D1. Remove semantic rediscovery from renderer paths

Goal:
- stop recovering meaning from emitted HTML when the AST can carry it directly

Done means:
- at least one current HTML-string recovery path is replaced with typed AST or typed renderer contracts
- tests prove equivalent or better output without HTML rescanning

### D2. Promote fenced semantics into first-class typed contracts

Goal:
- keep diagrams, charts, dataviews, and future fenced extensions out of regex pipelines

Done means:
- fenced extensions claim languages before rendering
- typed semantic nodes or typed extension payload contracts exist
- HTML rendering consumes those contracts directly

### D3. Keep generic renderer neutral

Goal:
- prevent host-specific behavior from defining the base package

Done means:
- generic plugin set is clearly separated from host-specific aliases
- backwards compatibility uses adapters or feature-pack registration helpers

## Workstream E: HTML Ingestion Convergence

### E1. Audit HTML-to-AST fidelity gaps

Goal:
- find where HTML ingestion still degrades too early or creates parallel semantics

Done means:
- a focused inventory exists for unsupported or weakly represented HTML structures
- each gap is classified as:
  - representable in current AST
  - needs AST expansion
  - should remain raw HTML fallback

### E2. Add typed recovery for one high-value HTML gap

Goal:
- prove the preferred pattern in real code

Done means:
- one meaningful HTML structure currently handled weakly is upgraded to typed AST recovery
- portable and OfficeIMO markdown writer expectations are both tested

## Workstream F: Compatibility Evidence

### F1. Create a compatibility matrix

Goal:
- make support explicit instead of implied

Done means:
- the matrix distinguishes:
  - CommonMark behavior
  - GFM behavior
  - OfficeIMO extensions
  - host-only semantics
  - intentional deviations

### F2. Add CommonMark-focused corpus runs

Goal:
- move from curated cases toward formal compatibility evidence

Done means:
- corpus cases run in CI
- unsupported cases are tracked instead of silently ignored

### F3. Add GFM-focused corpus runs

Goal:
- prove expected behavior for tables, task lists, autolinks, and related extensions

Done means:
- GFM coverage exists in CI
- behavior is classified as pass, fail, or intentional deviation

### F4. Add cross-pipeline round-trip suites

Goal:
- make sure parser, AST, HTML, and Word projections remain aligned

Done means:
- suites cover markdown -> AST -> markdown
- suites cover markdown -> HTML -> AST -> markdown
- suites cover markdown -> Word -> markdown where deliberate degradation is expected and documented

## Workstream G: Performance Evidence

### G1. Expand benchmark corpora

Goal:
- move beyond synthetic micro-cases

Done means:
- benchmarks include README-style docs, long lists, nested block content, large tables, and mixed rich documents

### G2. Track allocations and transform costs

Goal:
- avoid accidental performance regressions while architecture improves

Done means:
- baseline metrics exist for parse, transform, and render paths
- major refactors are checked against those baselines

### G3. Compare against stable external baselines

Goal:
- make "competitive" measurable

Done means:
- benchmark inputs are fixed
- comparisons against Markdig or other stable baselines are documented carefully
- results are used for prioritization, not vanity reporting

## Workstream H: Public Product Story

### H1. Keep the stable docs architecture-first

Goal:
- avoid stale feature claims during heavy development

Done means:
- roadmap docs focus on contracts, guardrails, and compatibility status
- package READMEs are updated when behavior settles

### H2. Publish a short "what we mean by correct" note

Goal:
- make review standards obvious to contributors

Done means:
- contributors can quickly see that regex-heavy semantic recovery and duplicate mutable state are not acceptable defaults

## Pull Request Checklist

Before merging a markdown-architecture PR, ask:

1. Does this make the syntax or semantic model clearer?
2. Does this reduce, rather than increase, semantic rediscovery?
3. Does this preserve or improve source mapping?
4. Does this keep generic and host-specific behavior separated?
5. Does this come with tests that prove the behavior we care about?

If not, the change probably needs another pass.

## Recommended First Batch

If we want the best next moves with the lowest architectural regret, do these first:

1. A1: audit syntax-node builder association
2. A2: add invariant test helpers
3. B1: inventory duplicated mutable node shapes
4. C1: formalize block parser ordering rules
5. D1: replace one HTML-string semantic recovery path with a typed path
6. F1: create the compatibility matrix

That batch improves trust in the architecture without forcing a single giant refactor.
