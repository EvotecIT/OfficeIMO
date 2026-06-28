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

Current scoreboard:

- compatibility matrix: `Docs/officeimo.markdown.compatibility-matrix.md`
- parity gap plan: `Docs/officeimo.markdown.markdig-parity-gap-plan.md`
- lossless roundtrip design: `Docs/officeimo.markdown.lossless-roundtrip-design.md`
- external parity baseline: Markdig `1.3.2`
- standards smoke baseline: 277 CommonMark `0.31.2` fixtures, 36 cmark-gfm extension fixtures, and a focused upstream ignored-autolink crash regression
- package guardrail baseline: tests and benchmarks must keep the same Markdig package version

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

## Source-Mapping Fixes To Keep Covered

These are not speculative roadmap items anymore. They are core invariants that should stay covered by tests.

- `SemanticFencedBlock` attaches itself as the syntax node `AssociatedObject`, so object-level `SourceSpan` mapping remains complete for semantic fenced extensions.
- `FootnoteDefinitionBlock` does the same, including structured paragraph child ownership in the final syntax tree.
- `ListItem` paragraph syntax associates to `ParagraphBlock` objects, not nested inline sequences.
- Sequence inline wrappers associate to the wrapper object across strong, emphasis, strong-emphasis, strikethrough, and highlight nodes.
- `DefinitionListBlock` group/value syntax associates to `DefinitionListGroup` and `DefinitionListDefinition` objects.

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

Recent progress:

- parsed raw HTML declaration blocks now expose source-backed opening marker, body, and closing marker syntax tokens plus native source fields, caret lookup, source edits, snapshot projection, and a CommonMark declaration smoke fixture while malformed raw HTML stays on the whole-block fallback
- parsed raw HTML processing-instruction and CDATA blocks now expose source-backed opening marker, body, and closing marker syntax tokens plus native source fields, caret lookup, source edits, and snapshot projection while malformed raw HTML stays on the whole-block fallback
- parsed simple raw HTML tag frames now expose source-backed opening tag, body, and closing tag syntax tokens plus native source fields, caret lookup, source edits, and snapshot projection while malformed raw HTML stays on the whole-block fallback
- the Markdig comparison baseline is now guarded so `OfficeIMO.Tests`, `OfficeIMO.Markdown.Benchmarks`, and the public compatibility docs cannot silently drift to different Markdig versions
- parsed HTML comment blocks now expose source-backed opening marker, body, and closing marker syntax tokens plus native source fields, caret lookup, source edits, and snapshot projection without claiming broader raw-HTML tag tokenization
- parsed footnote definitions now expose source-backed opening and separator marker syntax tokens plus native source fields, caret lookup, source edits, and snapshot projection while keeping label/body semantics stable
- parsed reference link definitions now expose source-backed opening and separator marker syntax tokens plus native metadata, source-field enumeration, caret lookup, source edits, and snapshot fields around single-line and multiline labels
- blockquote-contained reference definitions now participate in the document-level pre-scan before earlier paragraphs resolve shortcut links, and native reference-definition spans still point at the original quoted columns
- parsed inline footnote references now expose source-backed opening and closing delimiter syntax tokens while native label/marker metadata, source edits, and semantic traversal stay clean
- parsed inline/reference links and images now expose source-backed delimiter marker syntax tokens while native metadata/source edits and semantic inline traversal remain stable
- parsed hard breaks now expose source-backed marker syntax tokens for two-space, backslash, and inline HTML break spellings, while native hard-break metadata/source edits and semantic traversal stay clean
- parsed backslash escapes and decoded HTML entities now expose source-backed syntax token children under `InlineText`, while native text metadata/source edits and semantic inline traversal remain stable
- parsed code spans now expose opening delimiter, raw content, and closing delimiter syntax tokens, while native code inlines keep their existing source-addressable marker/content metadata and semantic traversal shape
- the GFM smoke lane now covers escaped table pipes, pipes inside code spans, escaped pipes inside table-cell code spans, broader table backslash escaping, no-leading-pipe tables, one-column delimiter rows, paragraph-to-table boundaries, aligned delimiter rows, header/delimiter mismatch rejection in both shorter and longer delimiter-row directions, body-row padding/truncation, reference links inside table cells, adjacent empty cells, compact inline emphasis in table cells, inline formatting in table headers/body cells, non-table pipe-row rejection, minimal header-only tables, raw inline HTML and break tags inside table cells, list-nested and blockquote-nested table syntax ownership, the cmark-gfm HTML tag filter, nested task lists, uppercase checked task markers, task-marker whitespace boundaries, non-task bracket-marker list items, plus-tag email local parts, invalid email-like tokens, bare `mailto:`/`xmpp:` autolinks, Unicode URL destinations, `www` host underscore rules, quoted/trailing-punctuation autolinks, an upstream ignored autolink crash regression, nested emphasis and delimiter-run edge cases inside strikethrough, and footnote ordering by first reference
- the CommonMark smoke lane now covers more code-span delimiter cases, underscore emphasis, digit-adjacent emphasis, broader absolute-URI autolinks, expanded official HTML block behavior for raw tables, type 1 blocks, comments, processing instructions, CDATA, and paragraph interruption, backslash-escaped punctuation, named entities, paragraph blank-line/indentation handling, hard and soft line breaks, the full official thematic-break, ATX-heading, and Setext-heading example sets including list-boundary, nested-list-item, indentation, closing-marker, escaped-marker, paragraph-interruption, empty-heading, indented-underline rejection, trailing-backslash, code-span, and HTML-looking heading text cases, compact nested blockquotes, blockquote lazy-continuation cases, nested blockquote list-continuations, list/code boundaries, shallow list indentation, and HTML-comment list boundaries
- CommonMark backslash escapes now use the full ASCII punctuation set, and inline entity references are decoded through the reader instead of remaining escaped text
- double-tilde strikethrough parsing now prefers the full `~~` delimiter frame instead of accidentally nesting two single-tilde spans when the GFM profile enables single-tilde strikethrough
- `FootnoteDefinitionBlock` now stores structured content through one canonical block list; paragraph-block and inline paragraph views are derived from that list instead of being stored as parallel state
- `ListItem` rewrite projection is now owned by `ListItem.ReplaceBlockChildren`, so recursive rewriters no longer duplicate the mapping from block children back into lead content, additional paragraphs, and nested blocks
- parsed list items now expose the list marker token, and task-list items expose the `[ ]`, `[x]`, or `[X]` marker token, as semantic/native source spans with snapshot and source-edit coverage
- parsed blockquotes now expose each explicit `>` marker token as semantic/native source spans, including nested quote remapping, snapshot projection, and marker-only source-edit coverage
- lossless/trivia roundtrip now has a concrete design target covering source slices, trivia capture, change tracking, roundtrip diagnostics, and phased implementation gates
- inline HTML rendering now has a type-targeted `HtmlOptions.InlineRenderExtensions` override seam, matching the existing direction for block HTML and markdown writer overrides
- inline HTML render extensions now have contract coverage for last-registration-wins ordering, `null` fallback to contextual/default inline rendering, and contextual renderers that can read body context, final syntax nodes, normalized source slices, and parsed inline source spans
- inline Markdown serialization now has a type-targeted `MarkdownWriteOptions.InlineRenderExtensions` override seam with contract coverage for last-registration-wins ordering, `null` fallback to default inline Markdown rendering, and contextual renderers that can read write context, final syntax nodes, normalized source slices, and parsed inline source spans
- block HTML render extensions can now read final syntax nodes plus normalized/original source slices from parsed documents while builder-only documents keep returning no source context
- block Markdown writer extensions can now read final syntax nodes plus normalized/original source slices from parsed documents, including CRLF-preserving original slices when `PreserveTrivia` is enabled
- block parser extensions now have contract coverage for same-placement registration order, first-success consumption, disabled-extension fallback to core parsers, multi-placement conflicts with later core parsers, and built-in/custom callout precedence
- custom block syntax builders can now ask the public builder context for owned child-container syntax, so extension blocks can reuse the same child syntax ownership rules as core containers
- custom block syntax builders can now set explicit inline-container associated objects through the public builder context, so wrapper syntax nodes can point at extension-owned semantic owners instead of being forced to associate with the inline sequence itself
- inline parser extensions now have contract coverage for registration order, first-success consumption at the current marker, false fallback to later extensions, disabled-extension skipping, and core parser fallback
- post-parse inline AST transforms now have a dedicated `MarkdownReaderOptions.InlineTransformExtensions` seam with contract coverage for ordering, replacement sequences, disabled-extension skipping, nested inline containers, and source-span preservation for reused nodes
- renderer fenced-code registrations now have semantic AST evidence: matching custom fences are parsed as semantic fenced blocks and render without falling back through default code-block HTML
- parsed callouts now carry the `[!KIND]` token span on the semantic block and native projection, giving editor tooling a direct source handle for callout markers
- parsed callout titles now carry through from `CalloutTitle` syntax association to native title field spans, with source-edit coverage for replacing just the title token
- public callout construction can now use structured body blocks through `CalloutBlock` constructors and the `MarkdownDoc.Callout(..., Action<MarkdownDoc>)` builder, so rich callout bodies no longer have to be flattened to raw markdown strings
- details summaries now carry through to native summary field spans, with source-edit coverage for replacing the summary element without replacing the whole disclosure block
- `DefinitionListBlock.ChildBlocks` now exposes the same structured definition body blocks already owned by `DefinitionListDefinition`, aligning definition lists with the public child-container shape used by other nested blocks
- `TableCell.ChildBlocks` now exposes each cell body through the same child-container shape as tables, lists, definitions, footnotes, and other nested semantic blocks
- parsed `TableCell` objects now expose their owned syntax children through the same generic syntax-child path used by other structured containers, so table syntax generation no longer needs a cell-specific syntax branch
- parsed `ListItem` objects now expose their owned block children through the same generic syntax-child owner path, keeping list-item paragraph and nested-block syntax aligned with the canonical `BlockChildren` projection
- parsed `DefinitionListBlock` objects now expose grouped definition syntax through the generic syntax-child owner path while keeping public `ChildBlocks` as the flattened definition-body projection
- parsed `TableBlock` objects now expose row/header/alignment syntax through the generic syntax-child owner path while keeping public `ChildBlocks` as the flattened cell-body projection
- definition lists now have a specialized native projection with grouped terms, definition bodies, nested child blocks, source spans, and snapshot DTOs instead of falling through the generic native `Other` block path
- footnote definitions now expose `LabelSourceSpan` and have a specialized native projection with label metadata and nested body blocks instead of falling through the generic native `Other` block path
- CommonMark thematic breaks now have a specialized native projection and snapshot shape instead of reporting horizontal rules as unsupported native blocks
- CommonMark fenced-code coverage now includes escaped/entity-decoded language info strings, tilde fences, longer closing fences, unclosed fences, empty fences, indented fences, blockquoted fences, non-opening backtick info strings, and language-plus-metadata examples
- reference-style link definitions are now exposed as effective parse-result and native-document metadata with definition-level spans, label-frame/destination/title token spans, snapshots, and native source-edit coverage
- inline footnote references now expose their label as syntax/native metadata with token source spans, snapshot spans, and native source-edit coverage
- fenced code and semantic fenced blocks now expose source-addressable opening fence, info-string, content, and closing fence spans through native projections and snapshots, with source-edit coverage for replacing only those tokens and nested blockquote/list/footnote remapping coverage
- raw HTML and HTML comment blocks now have native projection/source-edit evidence, so editor hosts can treat them as source-addressable blocks rather than unsupported fallbacks
- headings now expose source-addressable level and text spans through syntax nodes, native projections, and snapshots, with source-edit coverage for replacing only those tokens
- GFM pipe-table alignment rows now have a dedicated syntax node and native snapshot field span, with source-edit coverage for replacing the separator row without replacing the full table
- source-mapping coverage now includes a representative mixed semantic object graph, including heading, blockquote/list, definition-list group/value, table-cell nested blocks, code fence, semantic fenced block, paragraph, and footnote semantic objects
- parsed results now expose final syntax-tree associated-object lookup helpers for line, position, containing-span, and overlapping-span navigation, including typed overloads that resolve the nearest live semantic object after transforms
- sequence inline wrapper coverage now proves grouped strong, emphasis, strong-emphasis, strikethrough, and highlight syntax nodes map to wrapper objects rather than nested inline sequences
- native source-edit coverage now proves span-backed fenced code blocks and inline tokens can be replaced while preserving surrounding normalized source
- native source-edit helpers now cover source-backed list-item content spans, including nested list edits that preserve marker tokens and surrounding original trivia
- native document navigation now resolves source-backed list items and table cells directly by caret position, including list markers, task markers, nested list items, header/body table cells, and list-nested table cells with snapshot/source-edit coverage
- native document navigation now also resolves source-backed definition-list groups, terms, and definition bodies by caret position, including document-order enumeration for editor hosts
- parser-core source-slice coverage now proves syntax-backed parse results can materialize normalized source slices for span-backed nodes, `PreserveTrivia` can retain raw reader input, and line-ending-equivalent original input, including CRLF and standalone CR, can materialize original source slices without repointing normalized source spans
- roundtrip-writer coverage now proves unchanged trivia-backed parse results can return captured original markdown byte-for-byte, explicit native source edits can preserve original source around edited spans, and non-trivia, transformed, unsafe-map, or overlapping-edit cases report fallback diagnostics
- native inline metadata coverage now proves link target/title, image alt/source/title, and linked-image alt/source/image-title/link-target/link-title tokens carry source spans into snapshots and can be source-edited without replacing the surrounding paragraph
- native document navigation now enumerates source-backed inline metadata leaves and resolves metadata by caret position, so editor hosts can directly target link, image, footnote-reference, autolink, and inline-token spans
- native block source fields now provide first-class enumeration, name filtering, caret-position lookup, and source-edit targets over heading, fence, quote-marker, callout, details, footnote, table-alignment, and thematic-break token spans
- native reference-definition source fields now provide first-class enumeration, name filtering, caret-position lookup, source-edit targets, and snapshot projection over label-frame, destination, and title token spans
- native block snapshots now expose the same source-backed block fields through a repeat-aware `SourceFields` list, so serialized host projections can target the same tokens as live navigation APIs

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
4. Keep the compatibility matrix current for CommonMark, GFM, OfficeIMO extensions, and host-only semantics.
5. Add formal benchmark inputs from public docs/readme-style corpora.
6. Create a parity backlog that separates parser gaps, AST gaps, renderer gaps, and performance gaps.

## Recommendation

The best path is:

1. fix correctness and tree invariants first
2. make the AST canonical
3. open the parser and renderer seams
4. then push hard on standards parity and benchmark evidence

That ordering keeps the architecture healthy while feature coverage grows. It gives `OfficeIMO.Markdown` a real chance to become a serious markdown platform instead of only an effective internal subsystem.
