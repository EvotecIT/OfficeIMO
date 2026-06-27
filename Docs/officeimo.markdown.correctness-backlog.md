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

## Lossless Roundtrip Design

The target design for editor-grade source preservation lives in `officeimo.markdown.lossless-roundtrip-design.md`.

Treat `ToMarkdown()` as semantic markdown generation until that design is implemented. Lossless work should preserve original source slices for unchanged syntax nodes and report diagnostics whenever it falls back to generated markdown.
`MarkdownRoundtripWriter.WriteUnchanged` now covers the narrow unchanged-document case for trivia-preserved parse results, and `WriteWithSourceEdit` / `WriteWithSourceEdits` apply explicit native source edits back to original input when source spans remap safely. The writer reports fallback diagnostics when it cannot claim byte preservation.

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

Current coverage:

- callout kind syntax nodes now map their token span back to `CalloutBlock.KindSourceSpan`
- native callout projections expose the same kind token span for editor/read-model consumers
- callout title syntax nodes associate with `CalloutBlock.TitleInlines`, and native callout projections expose the explicit title span for snapshots and source edits
- footnote definitions now expose label token spans through `FootnoteDefinitionBlock.LabelSourceSpan` and a specialized native footnote projection
- public callout construction can now preserve rich bodies as child blocks through structured constructors and the fluent body builder instead of flattening lists/code/tables into a raw body string
- definition-list group/value syntax nodes now map source spans back to `DefinitionListGroup` and `DefinitionListDefinition` semantic objects
- definition lists now expose definition body blocks through public `ChildBlocks`, matching the same underlying objects owned by `DefinitionListDefinition`
- native definition-list projection now exposes grouped terms, definition body children, term inline runs, and snapshot DTOs instead of reporting definition lists as unsupported native blocks
- native thematic-break projection now exposes CommonMark horizontal rules as first-class native blocks with source spans and snapshots
- reference-style link definitions now have effective parse-result/native metadata with definition-level spans, label/destination/title token spans, snapshots, and source-edit coverage instead of living only in internal parser state and syntax nodes
- inline footnote references now expose label metadata spans in syntax and native snapshots, with native source-edit coverage for the label token
- inline links, images, and linked images now expose target/title and alt/source/title-style metadata spans in native snapshots, with native source-edit coverage for replacing those tokens without replacing the whole paragraph
- parsed formatting sequence inlines now expose native `openingMarker` and `closingMarker` metadata with source spans, snapshot coverage, and metadata source-edit coverage for nested emphasis/strong runs
- nested emphasis inlines now have native source-edit coverage through `MarkdownRoundtripWriter`, preserving original surrounding markdown and CRLF input while replacing the span-backed inline content
- fenced code and semantic fenced blocks now expose info-string/content source spans in native projections and snapshots, with source-edit coverage for replacing those tokens without replacing the full block
- CommonMark fenced-code smoke coverage now includes escaped/entity-decoded language tokens, tilde fences, longer closing fences, unclosed fences, empty fences, indented fences, blockquoted fences, and invalid backtick info strings that must remain paragraph/code-span text
- raw HTML and HTML comment blocks now project as source-addressable native HTML blocks with snapshots and block-level source edits instead of falling through unsupported native projection paths
- CommonMark HTML block smoke coverage now includes official raw table, type 1 `pre`/`script`, comment, processing-instruction, CDATA, paragraph-interrupt, and Markdown-between-raw-tags behavior, with AST assertions for source-addressable raw/comment blocks and split paragraphs
- raw HTML renderer security-profile coverage now proves `Strip`, `Escape`, and `Sanitize` behavior across CommonMark HTML block types 1-7
- CommonMark paragraph and line-break smoke coverage now includes official blank-line paragraph splitting, paragraph indentation normalization, indented-code paragraph boundaries, trailing-space hard breaks, backslash hard breaks, and soft line breaks
- CommonMark emphasis smoke coverage now includes underscore delimiter edge cases for leading whitespace, punctuation adjacency, intraword ASCII text, digits, and CJK text, with AST source-span coverage for a valid punctuation-adjacent opener
- nested list-item quote parsing now stops before valid non-one ordered child-list markers instead of swallowing them as lazy blockquote paragraph text
- nested list-item quote/ordered-list source remapping now keeps quote markers, ordered-list markers, syntax associations, native snapshots, and targeted source edits on original columns
- headings now expose level/text token spans in syntax nodes, native projections, and snapshots, with source-edit coverage for replacing heading markers or heading text without replacing the full block
- GFM pipe-table alignment rows now have a dedicated syntax node and native snapshot field span, with source-edit coverage for replacing the separator row without replacing the full table
- GFM pipe-table smoke coverage now includes no-leading-pipe tables, one-column delimiter rows, paragraph-to-table boundaries, short delimiter cells, aligned columns, header/delimiter mismatch rejection, body-row padding/truncation to the header width, reference links inside table cells, adjacent empty cells, compact inline emphasis, inline formatting in table headers/body cells, non-table pipe-row rejection, minimal header-only tables, raw inline HTML/break tags, escaped pipes inside table-cell code spans, and broader table backslash escaping
- GFM raw HTML smoke coverage now includes the cmark-gfm HTML tag filter, preserving dangerous tags as source-addressable raw HTML nodes while rendering the filtered leading `<` as `&lt;`; focused renderer coverage now asserts the same filtering for raw HTML blocks and inline tags when raw HTML is otherwise allowed
- GFM strikethrough smoke coverage now includes official delimiter-run edge cases for single, double, long, and mismatched tilde runs
- GFM autolink smoke coverage now includes plus-tag email local parts, invalid email-like tokens, bare `mailto:`/`xmpp:` URLs, Unicode URL destinations, `www` host underscore rules, quoted autolinks, trailing punctuation trimming, and the upstream ignored malformed-email case as a parser stability/source-mapping regression; Markdig parity tests now keep plus-tag plain email comparison in the portable profile where bare emails are intentionally literal
- GFM footnote smoke coverage now includes first-reference ordering and repeated references to the same definition, including GitHub-style repeated backrefs and independent source spans for each repeated reference label
- list-item paragraph syntax nodes now associate to `ParagraphBlock` objects instead of their nested `InlineSequence`
- sequence inline syntax nodes now associate back to their wrapper objects across strong, emphasis, strong-emphasis, strikethrough, and highlight grouped inline content

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

Current coverage:

- representative semantic-object provenance coverage now walks a mixed document with heading, blockquote/list, definition list, table-cell nested blocks, code fence, semantic fenced block, paragraph, and footnote objects and asserts every expected semantic object has a mapped source span

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

Current coverage:

- `MarkdownBlockParserExtension` registrations at the same placement run in registration order
- the first successful block parser consumes the block and later matching parsers are not invoked
- disabled block parser extensions are skipped and core parsers continue as fallback
- placement anchors are covered across earlier/later core parser conflicts, including extension claims before reference definitions and late extensions that cannot preempt GFM tables
- built-in block extensions and custom block extensions follow the same registration-order conflict rule when registered at the same placement

### C2. Design inline parser extension contracts

Goal:
- stop growing inline behavior as internal-only logic

Done means:
- a public or clearly intended contract exists for inline parser registration
- extension order and failure behavior are defined
- at least one non-trivial inline extension uses the seam

Current coverage:

- `MarkdownInlineParserExtension` is public and covered by custom inline parser tests
- inline parser extensions now have registration-order, first-success, false-fallback, disabled-extension, and core-parser fallback coverage
- syntax-aware custom inline parser extensions now project inline-backed `Unknown` syntax nodes into native `Other` inline snapshots, preserving custom syntax kind, source span, markdown/text, and nested inline children for host read models
- `MarkdownInlineTransformExtension` separates post-parse inline AST normalization from token recognition
- `HtmlOptions.InlineRenderExtensions` lets hosts/packages override HTML for specific inline semantic types without editing those node types
- inline HTML render extensions now have explicit ordering and null-fallback tests: later matching registrations win, and `null` falls back to the inline's contextual/default renderer

### C3. Separate parsing from normalization

Goal:
- avoid mixing repair logic into core parsing in ways that are hard to reason about

Done means:
- parser-owned behavior and normalization-owned behavior are explicitly separated
- normalization can be applied intentionally instead of implicitly everywhere

Current coverage:

- post-parse inline AST transforms run after core inline parsing and built-in input normalization
- tests cover transform ordering, replacement sequences, `null` no-op behavior, disabled-extension skipping, nested inline containers, and source-span preservation for reused nodes

## Workstream D: Renderer And HTML Cleanup

### D1. Remove semantic rediscovery from renderer paths

Goal:
- stop recovering meaning from emitted HTML when the AST can carry it directly

Done means:
- at least one current HTML-string recovery path is replaced with typed AST or typed renderer contracts
- tests prove equivalent or better output without HTML rescanning

Current coverage:

- renderer-owned fenced code block registrations are added to the reader as semantic fenced block extensions before parse
- `Markdown_Renderer_FenceConversion_Tests` proves a matching custom fence renders through the semantic AST renderer and does not invoke the `CodeBlock` HTML fallback

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

### F5. Implement lossless trivia/source-slice mode

Goal:
- make markdown -> parse -> targeted edit -> markdown preserve untouched source text where supported

Current coverage:

- native source edit helpers can replace a span-backed fenced code block while preserving surrounding normalized source
- native source edit helpers can also replace a span-backed inline token while preserving surrounding normalized source
- native source edit helpers can address parsed formatting opening/closing marker metadata without replacing the full inline content
- source-edit roundtrip coverage now includes replacing nested emphasis inline content in preserved original markdown while retaining surrounding formatting markers and CRLF trivia
- syntax-backed parse results can materialize normalized source slices for span-backed nodes
- `MarkdownReaderOptions.PreserveTrivia` retains raw reader input as parse-result metadata while keeping existing source spans tied to normalized markdown, and line-ending-equivalent original input, including CRLF and standalone CR, can now materialize original source slices through line/column coordinates
- `MarkdownRoundtripWriter.WriteUnchanged` returns the captured original markdown byte-for-byte for unchanged parse results and reports diagnostics when it falls back to generated markdown
- `MarkdownRoundtripWriter.WriteWithSourceEdit` and `WriteWithSourceEdits` apply explicit native source edits to preserved original markdown when each edit can be remapped safely, and fall back to normalized markdown with diagnostics when they cannot
- these are useful groundwork, not a first-class lossless trivia mode

Done means:
- parser captures enough trivia and source slices for supported block and inline cases
- unchanged supported fixtures roundtrip byte-for-byte
- targeted edits preserve surrounding source byte-for-byte
- diagnostics report every fallback to generated markdown

## Workstream G: Performance Evidence

### G1. Expand benchmark corpora

Goal:
- move beyond synthetic micro-cases

Current coverage:

- `OfficeIMO.Markdown.Benchmarks` uses stable in-source corpora for README-style docs, chat/transcript documents, technical docs, mixed rich AST content, long nested lists, large pipe tables, and normalization-heavy transcript artifacts

Done means:
- benchmarks include README-style docs, long lists, nested block content, large tables, and mixed rich documents

### G2. Track allocations and transform costs

Goal:
- avoid accidental performance regressions while architecture improves

Current coverage:

- BenchmarkDotNet `MemoryDiagnoser` is enabled for parse, syntax-tree parse, HTML render, and normalization/document-transform benchmarks
- `MarkdownTransformBenchmarks` measures OfficeIMO baseline parse, parse with normalization transforms, syntax-tree parse with transform diagnostics, and markdown generation after transforms

Done means:
- baseline metrics exist for parse, transform, and render paths
- major refactors are checked against those baselines

### G3. Compare against stable external baselines

Goal:
- make "competitive" measurable

Current coverage:

- `OfficeIMO.Tests` and `OfficeIMO.Markdown.Benchmarks` both reference Markdig `1.3.2`; a guardrail test keeps the versions aligned
- `dotnet list ... package --outdated` reported no newer Markdig package for the parity test or benchmark projects on 2026-06-27

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
