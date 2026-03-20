# OfficeIMO.Markdown Markdig Competitor Roadmap

This document turns the recent review into a practical roadmap for making `OfficeIMO.Markdown` feel like a real general-purpose markdown engine instead of a host-specific parser with good extras.

## Current Assessment

`OfficeIMO.Markdown` already has several things many markdown libraries never reach:

- a typed public object model
- a syntax tree with source spans and lookup helpers
- block parser extensibility
- post-parse document transforms
- HTML import/export and Word-oriented integration points
- a large markdown test surface

That is a strong base. It is not yet at true Markdig-competitor level for one reason more than any other: the parser, AST, extension API, and renderer model are not yet centered around one canonical, lossless tree with broad plugin seams.

## Where We Are Strong

- The library already models markdown as structured data instead of treating it as strings all the way through.
- The syntax tree is detailed enough to support diagnostics, editor integrations, and source mapping.
- The reader has better host-specific semantic modeling than many generic markdown libraries.
- The test suite is already large enough to support incremental refactoring if we keep tightening parity coverage.

## Where We Are Behind Markdig

Markdig’s differentiators are not only feature count. The bigger advantages are:

- a mature extension model for parsing and rendering
- broad CommonMark and GFM compatibility expectations
- clear seams for custom parsers, renderers, and pipeline composition
- confidence that extensions can cooperate without fighting the core model

Today `OfficeIMO.Markdown` is still behind on those points.

## Immediate Bugs To Fix

These are not roadmap items. They should stay fixed and covered by tests.

- `SemanticFencedBlock` must attach itself as the syntax node `AssociatedObject`, otherwise object-level `SourceSpan` mapping is incomplete.
- `FootnoteDefinitionBlock` must do the same for the same reason.

## Main Architectural Gaps

### 1. The AST is not fully canonical

Several public nodes keep multiple parallel representations of the same content. Examples include:

- `CalloutBlock`
- `FootnoteDefinitionBlock`
- `ListItem`

That creates drift risk for transforms, renderers, and future extension authors.

Target direction:

- each public node should have one primary structural representation
- convenience text/views should be derived, not separately stored
- rewrites should only need to update one representation

### 2. Inline extensibility is still too closed

The block pipeline is becoming extensible, but inline parsing still lives mainly as internal reader logic.

Target direction:

- add first-class inline parser extensions
- support ordered inline parser registration
- support post-inline normalization passes separately from parsing

### 3. Syntax tree and semantic tree are still too coupled

Right now the system behaves like a hybrid between a parse tree and a semantic object graph.

Target direction:

- keep `MarkdownSyntaxNode` as the lossless parse tree
- keep `MarkdownDoc` and child objects as the semantic tree
- make mapping between them explicit and dependable
- ensure transforms that replace semantic nodes also have a clear final-syntax rebuild path

### 4. Rendering is not yet a broad plugin surface

To compete seriously, rendering should not feel like a fixed HTML path with a few escape hatches.

Target direction:

- formalize renderer abstractions over the semantic tree
- let extensions register node renderers cleanly
- make non-HTML writers a first-class long-term goal

### 5. Standards and parity coverage are still too curated

The current tests are strong, but they are still closer to curated product coverage than to formal compatibility coverage.

Target direction:

- add CommonMark spec corpus coverage
- add GFM-focused corpus coverage
- track intentional deviations explicitly
- add performance and allocation baselines against representative public corpora

## Recommended Phases

### Phase 0: Stabilize source mapping and tree invariants

- keep object-level `SourceSpan` coverage growing for all AST nodes
- audit every `ISyntaxMarkdownBlock.BuildSyntaxNode` implementation for `AssociatedObject` consistency
- add invariants for parent/root/index/sibling binding across both syntax and semantic trees

### Phase 1: Make the public semantic AST canonical

- remove duplicate storage where blocks keep both raw text and block/inline children as peers
- define clear ownership rules for child nodes
- add lightweight computed helpers for legacy ergonomics instead of duplicated state

### Phase 2: Open the parser properly

- formalize block parser extension ordering and conflict rules
- add inline parser extension contracts
- add post-parse normalization hooks distinct from semantic transforms
- document which layer owns normalization vs parsing vs semantic upgrades

### Phase 3: Separate lossless syntax from semantic meaning

- make parse-tree fidelity a first-class contract
- ensure final syntax trees can be rebuilt reliably after transforms
- support syntax-to-semantic and semantic-to-syntax lookup without hidden mismatch risk

### Phase 4: Expand renderer and writer architecture

- add extensible renderer contracts over semantic nodes
- keep HTML as the first renderer, not the only design center
- prepare for alternate outputs such as markdown rewrite, plain text, and host-specific projections

### Phase 5: Prove compatibility and performance

- run CommonMark and GFM corpora in CI
- classify unsupported cases into backlog buckets
- expand benchmarks beyond synthetic samples into real README/docs corpora
- measure throughput, allocations, and transform costs against Markdig on stable benchmark inputs

## What “Competitive With Markdig” Should Mean

The realistic target is not “identical to Markdig.”

A better definition is:

- reliable CommonMark and major GFM behavior
- stable plugin model
- canonical AST and dependable source mapping
- strong renderer/story for custom semantics
- enough performance and compatibility data that users can choose it confidently

If `OfficeIMO.Markdown` reaches that bar, it can be a credible alternative even if it keeps a different object model and different host-oriented strengths.

## Practical Near-Term Backlog

1. Audit all syntax-node builders for `AssociatedObject` coverage and add regression tests where missing.
2. Pick one duplicated public node shape and make it canonical as the pattern for the rest.
3. Design an inline parser extension API before adding more inline features.
4. Add a documented compatibility matrix for CommonMark, GFM, OfficeIMO extensions, and host-only semantics.
5. Add formal benchmark inputs from public docs/readme-style corpora.
6. Create a parity backlog that separates parser gaps, AST gaps, renderer gaps, and performance gaps.

## Recommendation

The best path is:

1. fix correctness and tree invariants first
2. make the AST canonical
3. open the parser and renderer seams
4. then push hard on standards parity and benchmark evidence

That ordering keeps the architecture healthy while feature coverage grows. It gives `OfficeIMO.Markdown` a real chance to become a serious markdown platform instead of only an effective internal subsystem.
