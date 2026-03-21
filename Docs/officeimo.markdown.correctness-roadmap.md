# OfficeIMO Markdown Correctness Roadmap

This document is the quality bar for `OfficeIMO.Markdown`, `OfficeIMO.MarkdownRenderer`, and `OfficeIMO.Markdown.Html`.

It is intentionally not a feature inventory.

For issue-sized follow-up work, use `officeimo.markdown.correctness-backlog.md` alongside this roadmap.

Why:

- the implementation is moving quickly
- package READMEs may temporarily lag behind the code
- we still need one stable source of truth for what "good work" means

The goal is simple:

- do not ship "just enough"
- do not solve structural problems with regex-heavy postprocessing
- do not trade away AST correctness for short-term convenience
- make the stack credible as a serious markdown platform and ingestion engine

## What "Correct" Means

For this stack, correct does not mean "identical to Markdig" or "identical to MarkItDown".

It means:

- parsing is driven by syntax-aware readers, not string hacks
- the lossless syntax tree and semantic model have clear responsibilities
- transforms operate on typed structures instead of reparsing emitted text or HTML
- renderers and writers are extension seams, not the place where semantics are rediscovered
- source mapping, spans, and provenance are dependable enough for diagnostics and citations
- compatibility gaps are explicit, tested, and intentional

## Hard Rules

The following should be treated as architectural guardrails, not suggestions.

1. Do not use regex as the primary semantic engine for markdown or HTML structure.
   Regex may still be acceptable for narrow validation, token cleanup, or targeted diagnostics.
   It is not acceptable as the main way to discover fenced blocks, list structure, tables, headings, or HTML block semantics after a document has already been rendered.

2. Do not keep multiple mutable representations of the same meaning unless there is a very strong reason.
   If a node has child blocks or child inlines, that structure should be primary.
   Convenience text should be derived where possible.

3. Do not let host-specific behavior define the generic core.
   IntelligenceX-specific aliases, payload conventions, and transcript repair rules should layer on top of the generic packages, not shape their core APIs.

4. Do not flatten rich structure too early.
   If HTML or markdown can still be represented faithfully in the AST, keep it there.
   Text-only degradation should happen at the final projection boundary, not in the middle of the pipeline.

5. Do not add new syntax features without deciding where they belong.
   Every addition should answer:
   - parser responsibility
   - syntax-tree representation
   - semantic-tree representation
   - writer/render path
   - transform story
   - provenance/source-span story
   - compatibility-test story

## Preferred Architecture

The preferred stack shape is:

1. Source text or HTML enters a syntax-aware reader.
2. The reader builds a lossless parse representation with dependable source spans.
3. A semantic document model is built from that representation.
4. Typed transforms operate on the semantic model.
5. Writers and renderers project the semantic model into HTML, markdown, plain text, Word, or host-specific outputs.

That means:

- markdown semantics should not be rediscovered from emitted HTML
- HTML ingestion should map into the same semantic world instead of inventing a parallel one
- plugin behavior should register into parser, transform, and renderer seams explicitly

## Package Responsibilities

### OfficeIMO.Markdown

Owns:

- markdown parsing
- syntax tree contracts
- semantic AST contracts
- source spans and provenance
- generic normalization primitives
- generic markdown writing

Should not own:

- product-specific transcript policy
- host-specific UI/runtime behavior
- renderer-only postprocessing that compensates for missing AST semantics

### OfficeIMO.MarkdownRenderer

Owns:

- rendering pipeline
- renderer plugins and feature packs
- host-facing presentation contracts
- renderer-time transforms that are truly presentation concerns

Should not own:

- primary markdown parsing semantics
- chat transcript repair that is specific to one host
- regex recovery of semantics that should already exist in the AST

### OfficeIMO.Markdown.Html

Owns:

- HTML ingestion into the markdown semantic world
- HTML-to-AST mapping
- configurable preservation and fallback rules

Should not own:

- a disconnected second markdown object model
- silent flattening of unsupported but preservable structure

## What Good Extensions Look Like

A good extension:

- declares what syntax it claims
- maps to typed nodes or typed transforms
- participates in source mapping cleanly
- composes with other extensions through explicit ordering and contracts
- can be tested at parser, AST, and renderer levels independently

A weak extension:

- waits until HTML output exists
- scans strings to recover structure
- stores duplicate text and structure that can drift
- only works because one specific host happens to emit one specific shape

## Proof Requirements

We should trust improvements only when they are backed by at least one of these:

- corpus tests
- round-trip tests
- source-span/provenance tests
- compatibility matrices
- benchmark data

For mature features, we should want several of them.

## Minimum Evidence For New Work

### New parser behavior

Requires:

- positive tests
- negative tests
- ambiguous-edge-case tests
- source-span assertions where applicable

### New AST node shape

Requires:

- invariants for parent/child ownership
- syntax-to-semantic mapping coverage
- writer/render coverage
- transform safety checks

### New renderer/plugin feature

Requires:

- proof that semantics are available before HTML emission when they should be
- plugin composition coverage
- fallback behavior coverage

### New import path from HTML

Requires:

- typed AST mapping where structure is representable
- explicit raw HTML fallback when not representable
- portable-writer and OfficeIMO-writer expectations

## Priority Order

When there is pressure to move quickly, prefer work in this order:

1. correctness and invariants
2. canonical AST shape
3. extension seams
4. compatibility coverage
5. performance tuning
6. breadth expansion

That order matters.

Adding more syntax or more formats on top of shaky internals will make the system harder to fix later.

## Near-Term Roadmap

### Phase 1: Lock the invariants

- audit syntax-node builders for dependable object association and source spans
- keep semantic-tree ownership rules explicit
- remove duplicated mutable state where practical

### Phase 2: Finish the parser and AST story

- make inline extensibility first-class
- keep block and inline extension ordering explicit
- define when normalization is parser-owned versus transform-owned

### Phase 3: Remove semantic rediscovery

- stop relying on post-HTML regex discovery for semantics that can be represented in AST form
- promote fenced semantics and other extension points into typed contracts before rendering

### Phase 4: Separate generic and host-specific behavior

- keep generic packages neutral
- move host-specific aliases and transcript policy behind adapters or host packages
- keep backwards compatibility through registration helpers, not through core-package leakage

### Phase 5: Prove the claims

- run CommonMark and GFM-focused corpora
- add cross-pipeline round-trip suites
- add benchmark corpora from real README/docs-style inputs
- maintain a compatibility matrix that distinguishes generic behavior from OfficeIMO-only extensions

## How To Evaluate A Proposed Shortcut

Ask these questions:

1. Does this bypass the AST or syntax tree?
2. Does this rediscover semantics from emitted HTML or markdown text?
3. Does this introduce duplicate mutable state?
4. Does this blur generic and host-specific responsibilities?
5. Will this be harder to reason about once more extensions are added?

If the answer to any of those is "yes", the default decision should be "do not merge yet".

## Coordination Note

Because the markdown packages are under active development, this document should stay more stable than package READMEs.

That means:

- use this file as the decision-making guide while implementation is moving
- update package READMEs when behavior settles
- prefer new focused design docs over broad README churn during active refactors

## Recommendation

The path to being a real Markdig or MarkItDown challenger is not "support more inputs at any cost".

It is:

1. keep the architecture principled
2. keep the AST canonical
3. keep extension seams explicit
4. prove compatibility and performance with data
5. only then broaden the surface aggressively

If we hold that line, OfficeIMO can be both ambitious and trustworthy.
