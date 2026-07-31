# OfficeIMO.Markdown compatibility matrix

This matrix describes the current standards profiles, OfficeIMO extensions, source-model behavior, and documented limits. It is based on the generated [CommonMark inventory](officeimo.markdown.commonmark-inventory.md), generated [GFM inventory](officeimo.markdown.gfm-inventory.md), focused parser/renderer tests, and round-trip/source-edit contracts.

Open Markdown work is tracked in the repository [roadmap](ROADMAP.md).

## Status meanings

- **Covered:** implemented and protected by focused evidence.
- **Partial:** useful behavior exists with a named incomplete edge or source/writer limit.
- **Intentional:** the OfficeIMO profile deliberately adds behavior outside a standards profile.
- **Gap:** unavailable; no parser, transform, or renderer option implements the named behavior. The route names only the candidate owner for a possible implementation.
- **Unsupported:** input remains literal, source-preserved, diagnosed, or rejected according to the named profile.

## Evidence baseline

| Evidence | Current result |
| --- | --- |
| CommonMark smoke fixtures | 316 of 652 official CommonMark `0.31.2` examples pinned as smoke fixtures |
| CommonMark full inventory | 651 of 652 official CommonMark `0.31.2` examples match; 1 failures |
| GFM smoke fixtures | 52 cmark-gfm extension smoke fixtures plus focused OfficeIMO supplements for upstream ignored-autolink crash and query/fragment autolink regressions |
| GFM generated inventory | 52 tracked GFM fixtures; 52 passing, 0 failing |
| Engine ownership | `OfficeIMO.Markdown` owns parsing, semantic AST, syntax tree, writing, source edits, and HTML projection |
| Host rendering | `OfficeIMO.MarkdownRenderer` owns the WebView/browser shell and incremental host updates |

## Extension-family inventory

This is the published routing inventory for optional and profile-specific extension families represented in the compatibility test corpus. It is not a competitive scorecard: each row states whether the behavior belongs in the core parser, an opt-in extension, renderer/host policy, or an intentional/deferred boundary. The inventory test verifies every reflected row against this table so per-family coverage cannot disappear during documentation consolidation.

| Metric | Count |
| --- | ---: |
| Extension-family rows | 33 |
| Covered | 13 |
| Partial | 8 |
| Intentional | 3 |
| Gap | 9 |

| Extension family | Status | Current route |
| --- | --- | --- |
| Abbreviations | `Covered` | Core parser, opt-in |
| Advanced extension bundle | `Intentional` | Intentional bundle guard |
| Alert blocks | `Covered` | Core parser plus renderer policy |
| Auto identifiers | `Covered` | Core renderer option |
| Extended autolinks | `Covered` | Core parser, profile-gated |
| Bootstrap renderer helpers | `Intentional` | Renderer theme policy |
| CJK-friendly emphasis | `Covered` | Core delimiter parser option |
| Citations | `Gap` | Unavailable; candidate owner: Optional parser extension, deferred |
| Custom containers | `Partial` | Core extension seam plus optional built-in parser |
| Definition lists | `Covered` | Core parser, opt-in/profile-gated |
| Diagrams | `Partial` | Renderer/host policy over semantic fences |
| Emoji and smiley | `Gap` | Unavailable; candidate owner: Optional inline transform |
| Emphasis extras | `Covered` | Core inline parser, profile-gated |
| Figures | `Partial` | Core image AST plus optional parser syntax |
| Footers | `Gap` | Unavailable; candidate owner: Deferred document semantics |
| Footnotes | `Covered` | Core parser, GFM profile |
| Generic attributes | `Partial` | Core AST/source architecture |
| Globalization | `Gap` | Unavailable; candidate owner: Deferred compatibility option |
| Grid tables | `Gap` | Unavailable; candidate owner: Optional block parser extension |
| Jira links | `Gap` | Unavailable; candidate owner: Optional link inline extension |
| List extras | `Partial` | Core parser, opt-in |
| Mathematics | `Partial` | Optional parser plus renderer/host policy |
| Media links | `Partial` | Renderer/host policy with optional link parser |
| Non-ASCII no-escape rendering | `Covered` | Renderer escaping policy |
| Pipe tables | `Covered` | Core parser, GFM profile |
| Pragma lines | `Gap` | Unavailable; candidate owner: Deferred metadata parser |
| Precise source location | `Partial` | Cross-cutting core source architecture |
| Referral links | `Gap` | Unavailable; candidate owner: Renderer policy |
| Self pipeline | `Intentional` | Intentional composition difference |
| SmartyPants | `Gap` | Unavailable; candidate owner: Optional inline transform |
| Soft line break as hard line break | `Covered` | Core parser option |
| Task lists | `Covered` | Core parser, GFM profile |
| YAML front matter | `Covered` | Core parser, OfficeIMO profile |

<!-- extension-partial-boundaries:start -->
### Partial-family boundaries

These are the exact current implementation boundaries and promotion requirements for every `Partial` family in the structured extension inventory.

#### Custom containers

- **OfficeIMO state:** OfficeIMO now exposes opt-in MarkdownReaderOptions.CustomContainers for Markdig-style colon-fenced custom containers, with a semantic CustomContainerBlock, root and nested container parsing, child block parsing, Markdig-compatible div/class HTML for scoped cases, Markdown writing from the parsed block, and source-backed syntax tokens for opening fence, info, nested children, and closing fence. Native projection exposes a dedicated block kind with fence, info, body, and closing-fence source fields, snapshots, source slices, caret lookup, lossless source edits, and reparse proof. Focused Markdig comparison covers simple containers, first-token class behavior, list children, blockquote-contained containers, list-item-contained tight-list containers, nested longer outer fences, no-space info starts, unclosed containers, shorter-fence nested-container fallback, and trailing-text fence lines that start nested containers. List-contained custom containers now preserve remapped opening fence, info, child paragraph, and closing fence source spans, and tight-list custom-container HTML suppresses child paragraph wrappers to match Markdig. Tight-list custom-container HTML rendering now routes syntax/type renderer overrides through the shared dispatcher, including nested child overrides and source-slice access for the custom-container syntax node. Generated Markdown writing now lengthens outer colon fences around nested generated custom containers, including list-item-contained containers, so writer output reparses without collapsing nested ownership. Broader container ownership plus remaining writer breadth remain partial.
- **Promotion bar:** Complete remaining blockquote/container breadth and broader writer behavior before promotion.

#### Diagrams

- **OfficeIMO state:** OfficeIMO has semantic fenced blocks and visual renderer hooks, but not Markdig diagram extension parity.
- **Promotion bar:** Named diagram language mapping, renderer package ownership, source/writer behavior, and comparison fixtures.

#### Figures

- **OfficeIMO state:** OfficeIMO has image/figure import and publisher figure rendering paths, but not Markdig figure syntax parity.
- **Promotion bar:** Separate HTML-import figure recovery from Markdown figure syntax, then prove renderer/writer/source behavior.

#### Generic attributes

- **OfficeIMO state:** OfficeIMO now has generic attribute storage on semantic MarkdownObject nodes and MarkdownSyntaxNode nodes, with fenced-code id/classes/attributes projected from MarkdownCodeFenceInfo through ordinary CodeBlock and SemanticFencedBlock parser paths. Ordinary CodeBlock HTML rendering now projects explicit fence-info attribute blocks onto the code element, including attribute-only info strings such as ```{#code .wide}```, language-plus-attribute info strings such as ```cs {#code .wide}```, and opaque-info-prefix forms such as ```cs linenums {#code .wide}``` without leaking `linenums` as an HTML attribute; source-backed standalone generic attributes before fenced code also render on the code element to match Markdig's UseGenericAttributes behavior. Single-character shorthand ids such as `{#h .wide}` stay literal like Markdig across headings, paragraphs, setext content, and fenced-code info strings, without syntax/native `attributes` metadata. Semantic fenced-block default fallback keeps its host semantic renderer boundary while CodeBlock fallback renderers receive the attributed CodeBlock. Fence-info attributes expose native `attributes` source fields and UI-safe snapshot fields for code and semantic fenced blocks so editor hosts can locate and replace the explicit `{...}` attribute segment without rewriting the language token, opaque fence options, or whole info string. Opt-in MarkdownReaderOptions.GenericAttributes now parses Markdig-style trailing attribute blocks for ATX headings, including closing-marker-before-attribute and attribute-before-closing-marker forms, Setext headings, paragraphs, standalone attribute blocks before fenced code, headings, setext headings including blank-separated standalone attribute forms, paragraphs, inline-image paragraphs in portable/Markdig profiles, root ordered lists, root unordered lists, pipe tables, OfficeIMO-default typed image blocks, dash setext/thematic forms, and indented code, root ordered/unordered list items, nested list items, blockquote-contained list items, definition-list-looking text without UseDefinitionLists, definition-list terms, definition-list definition first paragraphs, and Markdig-style pipe-table cells that promote attributes to the owning table, while standalone attributes before HTML blocks are consumed without metadata and blockquote paragraph, heading, and standalone-before-blockquote attribute blocks remain literal to match Markdig's container behavior. Standalone attributes before reference-link-definition-looking lines now produce attributed literal paragraphs without registering reference definitions, with source-backed native edit proof. Standalone fenced-code, standalone setext-heading, list-contained fenced-code, blockquote-contained fenced-code, blockquote-contained list, list, pipe-table, typed image-block, dash setext/thematic, and indented-code-derived paragraph attributes are source-backed in syntax/native/source-edit APIs; standalone pipe-table attributes inside blockquotes and list items now attach to the nested table like Markdig and remain source-editable on the original attribute line; pipe-table-looking paragraph runs followed by standalone generic attributes now stay paragraphs without table syntax or `attributes` metadata, matching Markdig's table/attribute boundary. List- and blockquote-contained standalone attributes before fenced code attach to the nested CodeBlock, match Markdig HTML, and preserve the attribute line through native source edits. Blockquote-contained standalone attributes before unordered, ordered, and task lists attach to the nested list, match Markdig HTML, and preserve the attribute line through native source edits across unordered, ordered, and task-list paths. List-contained standalone attributes before nested task lists now attach to the nested list, match Markdig HTML, and preserve the attribute line through native source edits. List attributes project to the top-level `&lt;ol&gt;`/`&lt;ul&gt;` element, pipe-table attributes project to the `&lt;table&gt;` element, typed image-block attributes project to the `&lt;img&gt;` element, dash setext/thematic attributes produce an empty h2, and indented-code attributes produce an attributed paragraph. Paragraph attribute blocks preserve Markdig's consumed separator whitespace in HTML and Markdown writing, including thematic-break-like paragraph lines such as `--- {#id}`, `*** {#id}`, and `___ {#id}`, no-space bare-URL paragraph attribute blocks such as `https://example.com{#id}` keep literal URL text and no-space Markdown writing, no-space abbreviation-ending paragraph attribute blocks target the owning paragraph like Markdig when UseAbbreviations is combined with UseGenericAttributes, ordinary no-space plain-text paragraph attribute blocks such as `word{#id}` and `C++{#id}` target the owning paragraph without stealing paired inline delimiter targets, unmatched trailing backtick runs such as `text`{#id}` and ``{#id}` target the owning paragraph like Markdig while valid code spans still target the code span, escaped final punctuation such as `\*{#id}` and `\`{#id}` targets the owning paragraph like Markdig, decoded character references such as `&copy;{#copy .wide}` consume following valid attribute blocks without rendering literal attribute text or emitting native `attributes` metadata like Markdig, escaped character-reference-looking text such as `\&copy;{#copy .wide}` targets the owning paragraph like Markdig, and standalone attribute continuation lines at the end of paragraphs are consumed without metadata or rendered output like Markdig, including soft and hard line-break forms; soft line-break continuation attributes preserve trailing text and source columns after the consumed attribute segment. Paragraph-contained attributes embedded at the end of nested link labels, image alt text, linked-image alt text, emphasis content, and strong content now promote to the paragraph owner like Markdig, strip the literal attribute text from nested content, and remain source-backed in syntax/native projections. List-item attribute blocks are consumed for Markdig-compatible HTML without projecting attributes onto &lt;li&gt;, preserve Markdig's consumed separator whitespace, write trailing attribute blocks while preserving the captured separator whitespace, and expose semantic attributes plus syntax/native/source-edit proof; focused Markdig interaction proof covers the same consumption behavior when UseTaskLists is also enabled. List-contained ATX and loose nested headings now keep trailing generic attribute text literal like Markdig and suppress automatic ids derived from that literal marker, while fenced-code attributes inside list items still attach to the code block with native source-field proof. Definition-list term attributes project to `&lt;dt&gt;` and remain source-backed on the semantic term, while first definition-value paragraph attributes are consumed without projecting onto `&lt;dd&gt;` and later continuation paragraph attribute blocks after a blank line remain literal without native `attributes` metadata, matching Markdig's rendered behavior. Footnote definition first body paragraphs consume and project generic attributes when UseFootnotes and UseGenericAttributes are combined, later continuation paragraph attribute blocks after a blank line remain literal without native `attributes` metadata, standalone generic attributes before footnote definitions are consumed without metadata to match Markdig's boundary, and footnote references consume following attribute blocks without rendering literal text or native `attributes` metadata to match Markdig's reference behavior. No-space inline attribute blocks attach to links, reference links, images, reference images, linked images, emphasis, strong, code spans, angle autolinks, superscript, and subscript nodes. Triple-delimiter strong-emphasis attributes render with Markdig-compatible attributes on both nested emphasis tags while Markdown writing preserves the single source attribute block. Markdig leaves strikethrough, highlight, and inserted emphasis-extra attribute blocks literal, and OfficeIMO follows that boundary. Raw inline HTML, typed inline HTML wrappers such as `&lt;u&gt;...&lt;/u&gt;`, `&lt;sup&gt;...&lt;/sup&gt;`, `&lt;sub&gt;...&lt;/sub&gt;`, `&lt;ins&gt;...&lt;/ins&gt;`, and `&lt;q&gt;...&lt;/q&gt;`, and inline HTML break markers such as `&lt;br&gt;` and `&lt;br /&gt;` consume a following generic attribute block without projecting it into rendered HTML, matching Markdig's rendered output for those shapes while preserving source-backed trailing text after consumed attributes. Those attributes flow through semantic/syntax storage, default HTML rendering, Markdown writing, and reparse proof for the covered shapes. Generic attribute blocks on covered block and inline shapes are source-backed as dedicated GenericAttributeBlock syntax tokens and in native projections as `attributes` source fields/metadata, with syntax navigation, preserved-trivia source slicing, snapshot, and source-edit proof. Inline attribute consumption now preserves source-backed trailing text after attributed targets and consumed-without-metadata targets such as footnote references, soft line-break continuations, and typed inline HTML wrappers, so editor hosts can still create exact original slices for the remaining text after the `{...}` block. It still does not parse generic attributes for arbitrary block families or every inline family.
- **Promotion bar:** Remaining arbitrary block-family parsing, complete inline-family breadth, and broader Markdown writer/source preservation across arbitrary shapes.

#### List extras

- **OfficeIMO state:** OfficeIMO exposes opt-in MarkdownReaderOptions.ListExtras for Markdig-style single-letter alphabetic ordered markers, roman ordered markers up to xxxix, dot and parenthesis delimiters, lower/upper marker families, nested list-extra markers, marker-style HTML type attributes, parsed marker text preservation, syntax list-marker spans, and Markdown writer preservation for parsed marker spelling. Focused Markdig comparison covers lower/upper alpha, lower/upper roman, non-1 starts, parenthesis delimiters, double-letter alpha fallback, mixed marker-family list splitting, nested list-extra markers including lower-roman lists after blank lines, after parent item text, and inside blockquote-contained list containers, default opt-in boundaries, and marker source metadata. Native source-field/source-edit proof covers list-extra markers inside blockquotes and nested unordered-list containers, with reparsed AST checks for edited alpha and roman markers.
- **Promotion bar:** Broaden remaining list-extra edge coverage, native source-edit APIs, and writer reparse proof before promotion.

#### Mathematics

- **OfficeIMO state:** OfficeIMO has math-oriented semantic/rendering paths through host options, but not Markdig math delimiter parity.
- **Promotion bar:** Inline/block math delimiters, AST/source/native metadata, writer preservation, and renderer handoff contract.

#### Media links

- **OfficeIMO state:** OfficeIMO has image/media document semantics, but not Markdig media-link provider parity.
- **Promotion bar:** Provider model, safe renderer output, writer preservation, and source metadata for shortcut media links.

#### Precise source location

- **OfficeIMO state:** OfficeIMO has syntax/source/native spans, source slices, original-source slices when trivia is preserved, reason-aware original-source slice failure reporting, inspectable parse-result/native source mappings that pair normalized slices with original slices and exact/line-ending-equivalent/unavailable mapping kinds, source edits, roundtrip diagnostics with machine-readable original-source failure reasons, addressable native block/snapshot source fields including repeated fields by occurrence index, native block source-field, list item/list-item paragraph, table row/table cell, definition-list group/term/definition, inline metadata, reference-definition field, abbreviation-definition field, and source-trivia snapshot raw normalized/original text plus original-source failure reasons, semantic HeadingBlock level/text source spans, semantic LinkInline/ImageInline/ImageLinkInline source spans for link URL/title, image alt/source/title, and linked-image target/title fields, semantic TextRun escape source spans, semantic decoded entity source-text spans, semantic HardBreakInline marker source spans, semantic CodeSpanInline content source spans, semantic AbbreviationInline text/title source spans, semantic ImageBlock source spans for standalone and linked image alt/path/title/link target/link title tokens, semantic CodeBlock and SemanticFencedBlock info/content source spans, semantic CustomContainerBlock name source spans, source-order inline metadata snapshots, native inline/inline-metadata source-slice APIs for source-backed link targets, titles, link/image delimiters, formatting and emphasis-extra delimiters, footnote definition markers, front-matter fences, code and semantic fenced block markers, formatting content, escaped-character markers, decoded entity source text, hard-break markers, raw inline HTML fragments, structured details opening/closing tag and summary opening/text/closing semantic/syntax/native fields, raw HTML block comment/tag/CDATA/declaration/processing-instruction markers, per-column table alignment markers, pipe-table row delimiters, native table row projections/source slices/source edits/snapshots, and similar metadata, source-slice APIs aligned with native source-edit targets for blocks, list item content, list-item paragraphs, table cells, table rows, definition-list objects, reference definitions, reference-definition fields, abbreviation definitions, abbreviation-definition fields, details opening/closing tags, summary opening/text/closing fields, custom-container names, and document-level source trivia, paragraph-level native projections, source slices, source-backed canonical syntax reconciliation, and original-preserving source edits for parsed list-item paragraphs, custom block parser context normalized source-slice APIs for parser-created spans and relative line ranges, custom inline parser context normalized source-slice APIs for claimed inline ranges, inline transform context normalized source-slice APIs for source-backed inline nodes and spans, emphasis sequence semantic node source spans, document-transform context normalized/original source-slice APIs for parsed model objects and syntax spans, block/inline HTML renderer contexts and Markdown writer contexts with normalized/original source-slice APIs plus original-source failure reasons, parse-result generated syntax diagnostics with syntax paths, index paths, source fallback anchors, and associated semantic object details, generated syntax-node original-source slice rejection with a dedicated generated-node failure reason including native inline and inline-metadata slices, native source edits that carry known original-source failure metadata at creation time for generated nodes, missing preserved trivia, non-equivalent original input, unmappable original spans, and similar mapping failures without duplicating the primary preserve-trivia roundtrip diagnostic, plus document-level source trivia, snapshots, source-order enumeration, position lookup, and normalized/original source slices for blank lines, whitespace-only lines, leading/trailing horizontal whitespace, tabs, and line endings. Document-level trivia columns and generic line/column source-slice fallback now expand tabs with the same tab-stop model as source maps. Original-source slices now map line-ending-equivalent normalized spans back to the original CRLF, LF, or standalone CR spelling through a shared mapper used by parse results and transform contexts, and line-ending trivia source edits preserve original source bytes around the changed trivia when possible. Full lossless trivia/original mapping is still partial.
- **Promotion bar:** Complete lossless trivia/original mapping, generated-node roundtrip semantics, and source-edit coverage before claiming parity.
<!-- extension-partial-boundaries:end -->

## Standards profiles

| Capability | CommonMark profile | GFM profile | OfficeIMO profile | Status and boundary |
| --- | --- | --- | --- | --- |
| ATX and Setext headings | Yes | Yes | Yes | Covered, including heading/source marker spans and Setext source edits |
| Paragraphs, entities, escapes, hard/soft breaks | Yes | Yes | Yes | Partial; broad grammar is covered while complete lossless trivia remains open |
| Thematic breaks | Yes | Yes | Yes | Covered with marker syntax and native projection |
| Fenced and indented code | Yes | Yes | Yes, plus semantic fenced blocks | Partial; core grammar is broad and semantic fenced blocks are an OfficeIMO extension |
| Blockquotes | Yes | Yes | Yes, plus opt-in callout recognition | Partial source/lossless breadth; callouts are intentional profile behavior |
| Ordered and unordered lists | Yes | Yes | Yes | Partial; parsing is broad while canonical subobject/source ownership continues to improve |
| Emphasis, strong, escapes, and code spans | Yes | Yes | Yes | Partial; standards grammar is broad and delimiter/trivia completeness remains open |
| Links, references, images, and autolinks | Yes | Yes | Yes | Partial; profile-specific bare-link and standalone-image behavior is explicit |
| Raw HTML | Optional | Optional with GFM tag-filter policy | Optional | Grammar and output security are separate options; one CommonMark raw-HTML inventory case remains open |
| Pipe tables | No | Yes | Yes | Covered for the tracked GFM table corpus, source spans, edits, and writer round trips |
| Task lists | No | Yes | Yes | Covered with marker/source metadata and GitHub HTML output |
| Strikethrough | No | Yes | Yes | Covered for the tracked profile grammar |
| Footnotes | Extension | Yes in the OfficeIMO GFM profile | Yes | Covered for structured bodies, ordering, backreferences, source metadata, and writer round trips |
| Front matter | No | No | Optional | Intentional OfficeIMO extension with source-backed metadata |
| Alerts and callouts | No | Optional | Optional | Intentional extension; standards profiles do not enable it implicitly |

## Source, editing, and extension contracts

| Area | Current contract | Status |
| --- | --- | --- |
| Semantic model | Typed public block and inline model used by transforms, converters, and writers | Covered |
| Syntax model | Final syntax tree with source spans, token/field nodes, semantic association, and navigation helpers | Partial; some subobject associations and full trivia coverage remain open |
| Original source | Normalized slices are available for span-backed nodes; line-ending-equivalent original slices are available when trivia is preserved | Partial; general original-to-normalized mapping remains open |
| Round-trip writer | Unchanged trivia-backed input can be emitted byte-for-byte; safe native edits apply to original input; fallback produces diagnostics | Partial; arbitrary generated/multi-node edits are not fully byte preserving |
| Parser extensions | Ordered block, fenced-block, and inline parser seams with explicit fallback and source-aware contexts | Covered for the public extension contract |
| Transforms | Ordered post-parse transforms with source-impact diagnostics | Covered for the current transform contract |
| Renderer/writer overrides | Type-targeted and syntax-kind-targeted HTML/Markdown overrides with nested rendering contexts | Covered for the public extension contract |
| Security profiles | Raw HTML, URL handling, images, sanitizer behavior, and host renderer policy are explicit | Covered for the documented profiles |

## Current limits

- Trivia and delimiter-token coverage is incomplete for some built-in and optional syntax nodes.
- Original-to-normalized mapping remains incomplete for every combination of CRLF/LF/CR, tabs, nested containers, transforms, and generated nodes.
- General byte-preserving writing after arbitrary tree mutation is not claimed.
- Optional grid tables, mathematics, media providers, figures, and diagram languages require an explicit product contract before they become parser defaults.
- HTML sanitization is a rendering policy, not part of Markdown grammar.
- Performance results are evidence for measured corpora and options, not a semantic compatibility claim.
