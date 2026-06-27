# OfficeIMO.Markdown Compatibility Matrix

This matrix tracks the work needed to make `OfficeIMO.Markdown` a credible Markdig-class Markdown engine while keeping the OfficeIMO-specific document strengths explicit.

Status values:

- `Covered`: behavior is implemented and protected by focused tests.
- `Partial`: behavior exists, but coverage or edge-case conformance is incomplete.
- `Intentional`: OfficeIMO behavior deliberately differs from CommonMark, GFM, or Markdig.
- `Gap`: behavior is missing or too weak to claim.
- `Planned`: design direction is known, but implementation has not landed.

## Baseline

| Area | Current baseline |
| --- | --- |
| External comparison package | Markdig `1.3.2` in `OfficeIMO.Tests` and `OfficeIMO.Markdown.Benchmarks`; guarded by `PackageDependencyGuardrailTests.MarkdownParityProjects_UseTheSameCurrentMarkdigBaseline` |
| CommonMark reference | 165 CommonMark `0.31.2` smoke fixtures |
| GFM reference | 32 cmark-gfm extension smoke fixtures plus a focused upstream ignored-autolink crash regression |
| OfficeIMO core package | `OfficeIMO.Markdown` owns parsing, semantic AST, syntax tree, writing, and HTML projection |
| Host renderer package | `OfficeIMO.MarkdownRenderer` owns WebView/browser shell rendering and incremental updates |

## Standards Coverage

| Capability | CommonMark profile | GFM profile | OfficeIMO profile | Evidence | Status | Next action |
| --- | --- | --- | --- | --- | --- | --- |
| ATX headings | Yes | Yes | Yes | `Markdown_CommonMark_Examples_Tests`, `Markdown_Reader_Profile_Tests` | Covered | Expand from smoke cases to full spec import |
| Setext headings | Yes | Yes | Yes | Markdig parity cases, profile tests, official CommonMark setext/container interaction fixtures | Partial | Add remaining full-spec list/blockquote interaction cases |
| Paragraphs, entities, and breaks | Yes | Yes | Yes | 165 CommonMark smoke fixtures, including official paragraph blank-line/indentation, hard-break, and soft-break examples; renderer tests; hard-break/entity fixtures; fence info-string entity fixtures | Partial | Classify remaining entity and inline-break edge cases in full matrix |
| Thematic breaks | Yes | Yes | Yes | CommonMark smoke fixtures, reader tests, native thematic-break projection tests | Covered | Keep source-span/native projection coverage aligned |
| Fenced code blocks | Yes | Yes | Yes, plus semantic fenced blocks | Markdig parity cases, semantic fenced block tests, native opening/info/content/closing source-span and source-edit tests including blockquote/list container remapping, 29 CommonMark fence delimiter/indentation/container/info-string smoke fixtures | Partial | Add remaining full-spec inventory |
| Indented code blocks | Yes | Yes | Yes | Markdig parity cases, profile tests, official CommonMark list/code-boundary fixtures | Partial | Expand deeper list-item and blockquote indentation fixtures |
| Blockquotes | Yes | Yes | Yes, plus callout recognition by default | Markdig parity cases, compatibility notes, official CommonMark compact nested blockquote, lazy-continuation, container-boundary fixtures, and nested list-item quote-to-ordered-list boundary regression coverage | Partial | Expand remaining lazy-continuation corpus and document intentional callout behavior |
| Ordered/unordered lists | Yes | Yes | Yes | Markdig parity cases, profile tests, invariant tests, official CommonMark shallow-indentation, HTML-comment boundary, nested blockquote list-continuation fixtures, and non-one ordered-marker child-list regression coverage after nested quotes | Partial | Continue canonical `ListItem` cleanup and full CommonMark list corpus |
| Inline emphasis and escapes | Yes | Yes | Yes | Markdig parity cases plus CommonMark emphasis/backslash-escape smoke fixtures | Partial | Add delimiter-run spec breadth and CJK/intraword classifications |
| Links and reference links | Yes | Yes | Yes | Markdig parity cases, profile tests, parse-result/native reference-definition metadata and source-edit tests | Partial | Expand URI normalization and reference-definition corpus |
| Images | Paragraph inline images in standards profiles | Paragraph inline images in standards profiles | Standalone image block promotion enabled by default | Profile tests, syntax tests | Intentional | Keep standards behavior separate from OfficeIMO default promotion |
| Raw HTML | Optional | Optional with cmark-gfm tag filter render option | Optional | HTML block/inline tests, expanded CommonMark HTML block fixtures for raw tables, type 1 blocks, comments, processing instructions, CDATA, and paragraph interruption, native raw/comment HTML projection and source-edit tests, renderer security-profile tests proving `Strip`, `Escape`, and `Sanitize` behavior across CommonMark HTML block types 1-7, and GFM HTML tag filter fixtures/tests for raw blocks and inlines | Partial | Expand from smoke fixtures toward full CommonMark/GFM HTML corpus and classify sanitizer allowlist differences |

## GFM And Extension Coverage

| Capability | Current behavior | Evidence | Status | Next action |
| --- | --- | --- | --- | --- |
| Pipe tables | Enabled in GFM and OfficeIMO profiles; GFM requires a delimiter row and keeps cells inline-only, while OfficeIMO can still opt into headerless/structured cells | Table parser tests, GFM smoke fixtures for escaped pipes, code-span pipes, escaped pipes inside table-cell code spans, broader table backslash escaping, no-leading-pipe tables, one-column delimiter rows, paragraph-to-table boundaries, alignment, header/delimiter mismatch rejection, body-row padding/truncation, reference links inside cells, adjacent empty cells, compact inline emphasis, inline formatting in headers/body cells, non-table pipe-row rejection, minimal header-only tables, and raw inline HTML/break tags; alignment-row syntax/native span and source-edit tests | Partial | Add broader GFM table corpus for remaining delimiter/escaping interactions |
| Task lists | Enabled in GFM and OfficeIMO profiles | GFM fixtures covering nested items, uppercase checked markers, whitespace after markers, and non-task bracket-marker list text; profile tests | Covered | Keep GitHub HTML shape and marker-boundary tests current |
| Strikethrough | `~~text~~`; single-tilde in GFM profile | Profile tests, GFM nested-emphasis and delimiter-run fixtures | Covered | Expand full cmark-gfm delimiter edge corpus |
| Footnotes | Enabled in GFM and OfficeIMO profiles | GFM fixtures for GitHub HTML shape and first-reference ordering, footnote definition/reference metadata tests | Partial | Compare rendered footnote shape against cmark-gfm across more cases |
| Alerts/callouts | OfficeIMO callout extension; disabled in CommonMark/GFM profiles unless explicitly registered | Built-in extension tests and profile tests | Intentional | Keep registration explicit and matrixed |
| Autolinks | Configurable bare URL, selected GFM bare schemes, `www`, email, and angle autolinks; OfficeIMO/GFM profiles autolink plus-tag plain emails while CommonMark/portable profiles keep them literal | Markdig parity cases, profile tests, 7 CommonMark autolink fixtures, GFM fixtures for bare URLs, bare `mailto:`/`xmpp:` URLs, Unicode destinations, plus-tag emails, invalid email-like tokens, quoted/trailing punctuation trimming, `www` host underscore rules, and a focused upstream ignored-autolink crash regression | Partial | Expand remaining extended-autolink coverage |
| Front matter | OfficeIMO default only | reader/profile tests | Intentional | Keep front matter disabled in standards profiles |
| Semantic fenced blocks | OfficeIMO extension contract | semantic fenced block tests, native visual tests, renderer semantic-fence conversion tests | Intentional | Keep parser, AST, renderer, and native projection coverage aligned |
| Parser/render extensions | Block, fenced-block, inline parser, post-inline transform, block HTML, inline HTML, markdown block, and markdown inline override seams | block extension tests, block parser ordering/disabled-fallback tests, multi-placement parser conflict tests against reference definitions, GFM tables, and built-in callouts, inline parser ordering/fallback/disabled-extension tests, inline transform ordering/replacement/nested/span-preservation tests, custom renderer override tests including inline ordering and null-fallback behavior, semantic fenced renderer tests proving code-block fallback is bypassed | Partial | Add syntax-node-shape renderer design if needed |

## AST And Editor-Grade Coverage

| Contract | Current behavior | Evidence | Status | Next action |
| --- | --- | --- | --- | --- |
| Semantic AST | `MarkdownDoc` plus typed block/inline model | public API, tree invariant tests, footnote block-primary cleanup, list-item paragraph-block ownership tests, callout structured body construction tests, definition-list public child-block projection tests | Partial | Continue canonicalizing duplicated node shapes one type at a time |
| Syntax tree | `MarkdownSyntaxNode` with kind, literal, source span, children, custom kind, associated object | syntax tests, invariant tests | Covered | Keep expanding associated-object coverage |
| Original/final tree split | `MarkdownParseResult.SyntaxTree` and `FinalSyntaxTree` | transform and invariant tests | Covered | Add more transform replacement fixtures |
| Source spans | Line/column/offset-aware span model, with token-level spans for selected structured nodes | syntax tests, native document tests, heading level/text source-span tests, table alignment-row source-span tests, code-fence opening/info/content/closing source-span tests with blockquote/list/footnote remapping, callout kind/title source-span tests, list/task marker source-span tests, quote marker source-span tests, footnote definition/reference label source-span tests, inline link target/title, image alt/source/title, and linked-image alt/source/image-title/link-target/link-title metadata source-edit tests, representative semantic-object provenance test covering code fences and semantic fenced blocks | Partial | Add remaining high-risk nested/container provenance golden cases |
| Syntax-to-semantic mapping | `AssociatedObject` plus object-level and selected token-level `SourceSpan` binding | `MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent`, mixed semantic-object provenance test, table-cell/list-item/definition-list/code-fence/semantic-fence/callout/footnote/sequence-inline-wrapper source-mapping tests | Partial | Continue hotspot audit for any syntax wrappers or token fields without semantic owners |
| Native snapshot | `MarkdownNativeDocument` and snapshot DTOs for hosts | native document tests, native heading level/text field spans, table alignment-row field span, list/task marker span/source-edit coverage, quote marker span/source-edit coverage, native callout kind/title and details summary field-span coverage, native definition-list, footnote definition/reference, thematic-break projection, reference-definition metadata/span tests, raw/comment HTML block source spans and source edits, fenced-code opening/info/content/closing field spans including nested blockquote/list/footnote snapshots, reference-definition, inline footnote label, inline link target/title, inline image alt/source/title, and linked-image metadata source-edit preservation tests | Partial | Add more edit/roundtrip fixtures as lossless trivia mode lands |
| Lossless trivia/roundtrip | Early groundwork exists: syntax-backed parse results expose normalized source slices for span-backed nodes, `PreserveTrivia` retains raw reader input, line-ending-equivalent original input, including CRLF and standalone CR, can materialize original source slices through line/column coordinates, and `MarkdownRoundtripWriter` can preserve unchanged trivia-backed parse results byte-for-byte or apply explicit native source edits to original input when every edit remaps safely. It reports fallback diagnostics for generated or normalized output when trivia is missing, transforms changed the document, original-source mapping is unsafe, or edits overlap. Full trivia capture, implicit changed-node detection, and general byte-preserving edit writing are not implemented. | `Docs/officeimo.markdown.lossless-roundtrip-design.md`, source-slice tests, roundtrip-writer tests, fenced-code and inline-token source-edit preservation tests | Partial | Implement parser-owned trivia capture, full original/normalized offset mapping, changed-node tracking, and generated-node roundtrip diagnostics before broad editor app claims |

## Known AST Cleanup Targets

| Node | Risk | Direction | Status |
| --- | --- | --- | --- |
| `FootnoteDefinitionBlock` | Medium | Block children are now the primary structured content; paragraph and inline views are derived, with legacy text fallback for direct string construction; parsed labels expose token-level source spans and native projection metadata | Partial |
| `ListItem` | High | Paragraph blocks now own list-item paragraph syntax; `BlockChildren` provides the mixed AST-style block projection while legacy inline views remain compatibility helpers; rewrite projection is centralized in `ListItem.ReplaceBlockChildren`; parsed items expose list-marker source spans, and parsed task items expose task-marker source spans through semantic and native projections | Partial |
| `DefinitionListBlock` | High | Groups/definitions are structured and legacy tuple items are adapters; public `ChildBlocks` exposes definition body blocks as the derived child-container projection; native projection now exposes groups, terms, definition bodies, nested blocks, and snapshots; group/value syntax maps source spans back to semantic group/definition objects | Partial |
| `TableBlock` | High | Typed `TableCell` model carries row/column metadata, spans, blocks, and syntax ownership; continue treating raw/inline views as caches or adapters | Partial |
| `CalloutBlock` | Medium | Title inlines and child blocks are primary for parsed callouts; public structured constructors and the fluent body builder can now create the same child-block shape; body/title text helpers are derived; parsed kind and explicit title tokens now carry semantic/native source spans | Partial |

## Parity Gates

Before claiming Markdig-class parity, require:

1. Current Markdig package baseline in tests and benchmarks.
2. Full CommonMark `0.31.2` corpus imported or an explicit skipped-case inventory.
3. Broader cmark-gfm corpus coverage for enabled GFM extensions.
4. Compatibility matrix updated for every intentional deviation.
5. Tree invariant and associated-object tests for all canonicalized AST nodes.
6. Benchmarks against stable README/docs/chat corpora with parse, syntax-tree parse, HTML render, allocation, and transform costs.
7. Implement the documented lossless/trivia mode design before editor-grade roundtrip claims.
