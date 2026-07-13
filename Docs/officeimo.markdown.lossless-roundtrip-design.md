# OfficeIMO.Markdown Lossless Roundtrip Design

This design defines the target shape for a lossless markdown mode. It is not a claim that the current parser is lossless today.

The goal is to let editor and rewrite hosts preserve source text that they do not intentionally change, while still exposing the existing semantic `MarkdownDoc` model for rendering, transforms, and document integrations.

## Design Principles

- Keep `MarkdownSyntaxNode` as the parse-tree entry point, not the semantic model.
- Keep `MarkdownDoc` as the semantic tree; do not add trivia-only concerns to semantic blocks and inlines.
- Preserve original source slices for unchanged syntax nodes.
- Re-emit generated markdown only for changed nodes, inserted nodes, or nodes that never had source.
- Keep normalization opt-in. Lossless mode must not silently apply cleanup transforms.
- Make unsupported lossless edits explicit through diagnostics rather than silently formatting the whole document.

## Proposed Public Shape

The first implementation should introduce a small set of opt-in contracts:

- `MarkdownReaderOptions.PreserveTrivia`
  enables token/trivia capture during parsing.
- `MarkdownParseResult.SourceText`
  remains the normalized source used for spans; lossless mode also records the original source when normalization changed it.
- `MarkdownSyntaxTrivia`
  represents whitespace, blank lines, delimiters, markers, indentation, and trailing text that belongs to syntax fidelity rather than semantic meaning.
- `MarkdownSourceSlice`
  points at a stable source range and can materialize the original text for unchanged nodes.
- `MarkdownRoundtripWriter`
  writes a document from syntax plus semantic changes, using source slices when possible and generated markdown when required.
- `MarkdownRoundtripDiagnostic`
  reports when a node cannot be preserved losslessly and why.

These names are design targets. Final implementation can adjust names if the resulting API is clearer.

## Syntax Tree Responsibilities

Lossless syntax nodes should be able to answer:

- what exact source range they own
- which source range is semantic content
- which leading and trailing trivia belongs to the node
- which delimiters or markers formed the node
- whether the node was parsed from original source or generated after a transform
- whether child spans fully cover the node or leave source gaps

Examples:

- a heading node owns the opening `#` markers, spacing after the marker, heading inline content, optional trailing hashes, and line ending
- a list item owns marker indentation, marker text, post-marker padding, child blocks, blank lines between children, and continuation indentation
- a fenced code block owns opening fence, info string trivia, raw content, closing fence, and unmatched trailing text if the block is incomplete
- an inline emphasis node owns opening delimiters, inner inline nodes, closing delimiters, and delimiter-run metadata

## Semantic Tree Responsibilities

Semantic blocks and inlines should stay focused on meaning:

- `FootnoteDefinitionBlock.ChildBlocks` owns structured footnote content; paragraph views are derived
- list items should converge on block children as the primary content model
- definition lists should converge on one structural model for terms and definitions
- table cells should expose one canonical block/inline model rather than parallel mutable views

Semantic objects may hold source spans for mapping and diagnostics, but they should not become trivia containers.

## Write Algorithm

The roundtrip writer should use this decision order:

1. If a syntax node and all descendants are unchanged, write the original source slice.
2. If a node is unchanged but a child changed, write preserved trivia around generated child output.
3. If a node was inserted or has no usable source slice, write generated markdown from the semantic node.
4. If a transform replaced a semantic node while preserving a source span, write generated markdown for that node and preserve surrounding sibling trivia.
5. If preservation is ambiguous, emit generated markdown and add a diagnostic.

This keeps the common edit path small: update one node, preserve everything else byte-for-byte where possible.

## Change Tracking

Implementation should avoid diffing rendered markdown. Instead:

- bind each semantic object to its source syntax node through `AssociatedObject`
- record generated/replaced nodes during transform pipelines
- let transforms optionally report affected nodes and affected source spans
- derive unchanged status from syntax identity plus semantic object identity

The existing original/final syntax tree split is a useful starting point, but it needs explicit preservation metadata before it can support editor-grade roundtrip.

## Implementation Phases

### Phase 1: Trivia Model

- add trivia/slice types
- capture leading/trailing line trivia for top-level blocks
- expose diagnostics for source gaps
- add fixtures for headings, paragraphs, fenced code, blockquotes, lists, tables, and footnotes

### Phase 2: Block Roundtrip

- preserve untouched top-level blocks from source slices
- regenerate only changed top-level blocks
- preserve blank lines between blocks
- prove markdown -> parse -> roundtrip byte stability for supported block cases

### Phase 3: Nested Block Roundtrip

- add trivia preservation for list items, blockquotes, footnotes, callouts, tables, and definition lists
- preserve continuation indentation and blank lines inside containers
- add source-span and source-slice tests for nested edits

### Phase 4: Inline Roundtrip

- capture delimiter runs, escaped characters, raw autolink spelling, link destination/title spelling, image label spelling, and code-span fence lengths
- preserve unchanged inline spelling even when semantic rendering would normalize it
- regenerate only edited inline ranges

### Phase 5: Public Editor API

- expose non-mutating source edits over syntax and semantic objects
- expose diagnostics for unsupported lossless edits
- document fallback behavior for transforms that intentionally normalize content

## Completion Criteria

Lossless mode should not be called editor-grade until:

1. unchanged supported fixtures roundtrip byte-for-byte
2. one-block edits preserve surrounding source byte-for-byte
3. nested container edits preserve unrelated indentation and blank lines
4. inline edits preserve unrelated delimiters and escaping
5. diagnostics identify every fallback to generated markdown
6. compatibility docs distinguish semantic markdown writing from lossless source rewriting

## Current Status

Current `OfficeIMO.Markdown` has strong source spans, syntax-to-semantic associations, native projections, source edit helpers, and an initial source-slice primitive. Syntax-backed parse results can now materialize normalized-source slices for span-backed syntax nodes, and native documents now expose document-level source trivia for empty lines, whitespace-only lines, leading horizontal whitespace, trailing horizontal whitespace, tabs, and line endings through live projections, source-order enumeration, position lookup, snapshots, and normalized/original source-slice APIs. Document-level trivia columns, source maps, offset-less line/column source slices, offset-less line/column native source edits, and visual-column start-offset lookups for prefix preservation now share one tab-expanded source-column model instead of parallel local implementations. Callout/alert headers and structured details opening/closing tags plus summary opening/text/closing fields now expose source-backed fields in native projections, so editor hosts can address those container tokens without rescanning raw strings. When `MarkdownReaderOptions.PreserveTrivia` is enabled, parse results also retain the raw reader input as `OriginalMarkdown`; original-source slices are exposed when the raw input is byte-identical to the normalized span backing text or differs only by line-ending normalization (`\r\n`, `\n`, or standalone `\r`), in which case offset-backed spans are mapped across equivalent line-ending tokens before falling back to line/column coordinates. `MarkdownRoundtripWriter.WriteUnchanged` can now return the captured original markdown byte-for-byte for unchanged parse results and reports diagnostics when it must fall back to generated markdown because trivia was not preserved or document transforms changed the result. `MarkdownRoundtripWriter.WriteWithSourceEdit` and `WriteWithSourceEdits` can also apply native span-backed source edits to the preserved original input when each edit can be remapped safely, while falling back to normalized markdown with diagnostics when trivia is missing, transforms changed the document, original-source mapping is unsafe, or edits overlap. Source-edit and transform fallback diagnostics carry the source span that could not map back to original input, made ordering ambiguous, lacked preserved trivia, or was changed by a transform when one is available, giving editor hosts a precise source target for recovery UI.

Span-backed native edits can already replace a fenced code block, an inline token, parsed formatting opening/closing marker metadata, and selected document-level source trivia, including line-ending trivia, while preserving surrounding normalized source, and the roundtrip writer now has first-step original-source preservation for those edits when `PreserveTrivia` is enabled. The parser still does not capture enough delimiter trivia, full original/normalized offset mapping for transformed or generated nodes, complete delimiter metadata across all inline spellings, or changed-node identity to promise general lossless markdown rewriting.

Until the broader design is implemented, `ToMarkdown()` should be treated as semantic markdown generation, and `MarkdownRoundtripWriter` should be treated as a conservative unchanged-document and explicit-source-edit primitive rather than a general byte-preserving edit writer.
