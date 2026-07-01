# OfficeIMO.Markdown Markdig Extension Compatibility Matrix

This matrix turns the Markdig `1.3.2` extension inventory into execution lanes. It is generated from `MarkdigExtensionInventory`, so the open work stays tied to reflected Markdig pipeline entry points instead of drifting into ad hoc fixture lists.

Use it as the control board for parity slices:

- If the `Engine parser` lane is open, improve `OfficeIMO.Markdown` behavior before adding more proof.
- If only `Proof` is open, the behavior already exists and needs focused Markdig/source/writer evidence.
- If `Decision` says optional, deferred, renderer policy, or intentional, make the scope call before implementing parser behavior.

## Summary

Current inventory: 33 Markdig extension-family rows; 13 covered, 7 partial, 3 intentional, 10 gap.

| Metric | Count |
| --- | ---: |
| Markdig extension-family rows | 33 |
| Covered | 13 |
| Partial | 7 |
| Intentional | 3 |
| Gap | 10 |

## Execution Matrix

| Markdig entry point | Status | Decision | Engine parser | AST/source | Writer/render | Proof | Next non-looping action |
| --- | --- | --- | --- | --- | --- | --- | --- |
| `UseAbbreviations` | `Covered` | Core engine | Covered | Covered | Covered | Covered | Keep abbreviation comparison, list-contained source-token, native source-edit, and writer fixtures aligned as lossless trivia work expands. |
| `UseAdvancedExtensions` | `Intentional` | Intentional difference | Not planned | Not planned | Intentional difference | Documented | Keep this row as a roll-up guard; do not implement as a broad on switch. |
| `UseAlertBlocks` | `Covered` | Renderer/host policy | Covered | Covered | Covered | Covered | Keep alert comparison, source/native, and writer fixtures aligned as broader GFM and lossless trivia work expands. |
| `UseAutoIdentifiers` | `Covered` | Renderer/host policy | Covered | Covered | Covered | Covered | Keep slug-style and heading-source fixtures aligned as broader renderer profiles evolve. |
| `UseAutoLinks` | `Covered` | Core engine | Covered | Covered | Covered | Covered | Keep broader GFM fixture breadth separate from the Markdig UseAutoLinks row. |
| `UseBootstrap` | `Intentional` | Renderer/host policy | Not planned | Not planned | Intentional difference | Documented | Keep theme/rendering presets separate from parser parity. |
| `UseCjkFriendlyEmphasis` | `Covered` | Core engine | Covered | Covered | Covered | Covered | Keep the CJK-friendly option aligned with future emphasis delimiter rewrites and do not enable it by default in CommonMark/GFM profiles. |
| `UseCitations` | `Gap` | Deferred | Deferred | Deferred | Deferred | Needs real consumer requirement | Decide whether citations are in scope after core CommonMark/GFM closure. |
| `UseCustomContainers` | `Gap` | Optional/core block extension | Missing colon-fence parser | Missing child-block source mapping | Missing renderer/writer seams | Missing Markdig comparison fixtures | Route to block parser extensions plus renderer/writer source-slice contracts. |
| `UseDefinitionLists` | `Covered` | Core engine | Covered | Covered | Covered | Covered | Keep definition-list comparison, native source, no generated-child diagnostic, and writer reparse fixtures aligned as lossless trivia work expands. |
| `UseDiagrams` | `Partial` | Renderer/host policy | Semantic fences exist | Needs language/source mapping decision | Needs renderer package ownership | Needs renderer comparison fixtures | Compare Mermaid/Nomnom-style cases and decide renderer-package ownership. |
| `UseEmojiAndSmiley` | `Gap` | Optional inline transform | Missing shortcode/smiley transform | Needs source metadata policy | Needs writer literal/normalized policy | Needs opt-in transform fixtures | Keep normalization separate from an optional inline replacement extension. |
| `UseEmphasisExtras` | `Covered` | Core engine | Covered | Covered | Covered | Covered | Keep emphasis-extra delimiter cases aligned as broader GFM and lossless trivia coverage expands. |
| `UseFigures` | `Partial` | Core image plus optional syntax | Partial image/figure behavior | Needs Markdown figure syntax source model | Needs renderer/writer contract | Needs syntax-vs-import proof | Separate HTML-import figure recovery from Markdown parser extension support. |
| `UseFooters` | `Gap` | Deferred | Deferred | Deferred | Deferred | Needs real consumer requirement | Leave out of scope unless document footer semantics become a Markdown requirement. |
| `UseFootnotes` | `Covered` | Core engine | Covered | Covered | Covered | Covered | Keep the GFM footnote fixture corpus and structured-body writer coverage current. |
| `UseGenericAttributes` | `Partial` | Core engine | Partial target coverage; strong-emphasis, single-character id grammar boundaries, list/blockquote-contained fenced-code, blockquote-contained lists, pipe-table trailing attribute boundaries, footnote/definition continuation boundaries, soft-line-break continuation/trailing text, and typed inline HTML wrapper breadth covered | Needs arbitrary-shape source propagation | Needs broader writer/render propagation | Needs remaining block/inline target proof | Extend the shared attribute parser/writer to more block and inline families, then promote once writer/source propagation and token-level coverage are proven across arbitrary shapes. |
| `UseGlobalization` | `Gap` | Deferred | Deferred | Deferred | Deferred | Needs real consumer requirement | Revisit only if a real consumer needs culture-sensitive Markdown behavior. |
| `UseGridTables` | `Gap` | Optional block parser | Missing grid-table parser | Missing table source model | Missing HTML/Markdown writer behavior | Missing malformed fallback fixtures | Decide if grid tables belong in core or an optional extension package. |
| `UseJiraLinks` | `Gap` | Optional link extension | Missing issue-key parser | Needs source metadata | Needs resolver/render policy | Needs opt-in fixtures | Treat as optional link extension after core link/source mapping is stable. |
| `UseListExtras` | `Partial` | Core opt-in parser | Alpha and roman ordered markers are parsed for the scoped Markdig syntax, including nested lower-roman lists after parent text and inside blockquotes | Uses canonical OrderedListBlock/ListItem marker style, delimiter, marker text, syntax marker spans, and nested/container listMarker source edits; needs broader edge breadth | HTML type/start and parsed-marker Markdown writing are covered for scoped cases | Has nested lower-roman Markdig comparison, blockquote nested-list comparison, and blockquote/nested-container source-edit reparse proof; needs remaining breadth before promotion | Expand remaining source-edit and reparse breadth before promoting to Covered. |
| `UseMathematics` | `Partial` | Optional parser plus renderer policy | Missing delimiter parity | Needs math node/source metadata | Needs renderer handoff and writer policy | Needs inline/block math fixtures | Define math parser ownership and compare inline/block math fixtures. |
| `UseMediaLinks` | `Partial` | Renderer policy plus optional parser | Missing shortcut parser | Needs source metadata for providers | Needs safe renderer output policy | Needs provider comparison fixtures | Route shortcut media providers through renderer/host extension seams if in scope. |
| `UseNonAsciiNoEscape` | `Covered` | Renderer/host policy | Covered | Covered | Covered | Covered | Keep direct encoder audits and focused non-ASCII render-policy tests current when adding new HTML output paths. |
| `UsePipeTables` | `Covered` | Core engine | Covered | Covered | Covered | Covered | Keep the GFM table fixture corpus and table-cell source-edit coverage current. |
| `UsePragmaLines` | `Gap` | Deferred | Deferred | Deferred | Deferred | Needs real consumer requirement | Leave out of core unless a concrete document workflow needs it. |
| `UsePreciseSourceLocation` | `Partial` | Cross-cutting source architecture | Partial parser spans | Has native block/snapshot field accessors, semantic HeadingBlock level/text source spans, semantic LinkInline/ImageInline/ImageLinkInline source spans for link URL/title, image alt/source/title, and linked-image target/title fields, semantic TextRun escape source spans, semantic decoded entity source-text spans, semantic HardBreakInline marker source spans, semantic CodeSpanInline content source spans, semantic AbbreviationInline text/title source spans, semantic ImageBlock source spans for standalone and linked image alt/path/title/link target/link title tokens, semantic CodeBlock and SemanticFencedBlock info/content source spans, structured details opening/closing tag and summary opening/text/closing semantic/syntax/native fields, native list-item paragraph projections/source slices/source-backed canonical reconciliation/original-preserving source edits, document-level abbreviation-definition source fields/snapshots/source slices/source edits, tab-aware line/column source slices, document-level blank-line, horizontal-whitespace with tab-expanded columns, and line-ending trivia source slices/edits, native inline/metadata/source-edit-target source slices, generated-node source-slice and source-edit failure metadata, custom block parser context source slices, custom inline parser context source slices, inline transform context source slices, document-transform context source slices, and reason-aware original mapping failures carried by native source edits; still needs full lossless trivia and mapping | Has explicit source edits, fallback diagnostics, and machine-readable original-source fallback reasons; needs broader roundtrip behavior | Needs broader source-edit and original-mapping proof | Continue Phase 3 source-map and trivia work before claiming parity. |
| `UseReferralLinks` | `Gap` | Renderer policy | Not parser-owned | Needs link metadata decision | Missing opt-in rel policy | Needs renderer-policy tests | Treat as renderer policy work if requested. |
| `UseSelfPipeline` | `Intentional` | Intentional difference | Not planned | Not planned | Intentional difference | Documented | Keep extension composition in OfficeIMO reader/render/write options. |
| `UseSmartyPants` | `Gap` | Optional inline transform | Missing smart punctuation transform | Needs source/edit behavior | Needs writer/escaping policy | Needs opt-in transform fixtures | Consider as an optional inline transform after delimiter parsing stabilizes. |
| `UseSoftlineBreakAsHardlineBreak` | `Covered` | Core engine | Covered | Covered | Covered | Covered | Keep the option covered alongside paragraph/list source-map and writer fixtures. |
| `UseTaskLists` | `Covered` | Core engine | Covered | Covered | Covered | Covered | Keep the GFM fixture corpus and marker source-edit coverage current. |
| `UseYamlFrontMatter` | `Covered` | Core engine | Covered | Covered | Covered | Covered | Keep raw YAML, parsed-entry helpers, and front-matter source-edit fixtures aligned as lossless trivia work expands. |

## Work Checklist

- [ ] Pick one row and one open lane before implementation starts.
- [ ] If the row needs engine work, change parser/AST/source/writer/render behavior in the owning layer first.
- [ ] If the row only needs proof, add focused Markdig comparisons, source/native snapshots, writer checks, renderer checks, or generated inventory assertions.
- [ ] If the row is optional, deferred, renderer-owned, or intentional, record that scope decision before adding syntax.
- [ ] Promote a row to `Covered` only when the matrix lanes and the inventory promotion bar agree.

## Immediate Queue

- [ ] Continue `UseGenericAttributes` only after probing remaining Markdig-supported block and inline targets. Avoid another standalone-attribute sweep unless Markdig evidence requires it.
- [x] Keep `UseDefinitionLists` covered while broader source/trivia work evolves; do not reopen it without new Markdig evidence.
- [x] Keep the `UseAlertBlocks` titled-callout boundary explicit: OfficeIMO mode keeps rich titles; Markdig-compatible mode treats titled markers as ordinary blockquotes.
- [ ] Continue `UsePreciseSourceLocation` through the broader lossless AST/source model; the current native block, semantic HeadingBlock level/text source spans, semantic LinkInline/ImageInline/ImageLinkInline source spans, semantic TextRun escape source spans, semantic decoded entity source-text spans, semantic HardBreakInline marker source spans, semantic CodeSpanInline content source spans, semantic AbbreviationInline text/title source spans, semantic ImageBlock source spans, semantic CodeBlock and SemanticFencedBlock info/content source spans, structured details and summary semantic/syntax/native tag and text fields, editable list-item paragraph, editable native table row, editable document source trivia including tab-expanded columns and line endings, tab-aware line/column source slices, inline, metadata, source-edit target, generated-node source-slice and source-edit failure metadata, machine-readable roundtrip fallback reasons, custom block parser context, custom inline parser context, inline transform context, and document-transform context source-slice APIs improve editor-grade source addressing but do not close full trivia parity.
