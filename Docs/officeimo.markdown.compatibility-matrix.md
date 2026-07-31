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

- **OfficeIMO state:** Opt-in colon-fenced containers support root and nested blocks, child parsing, HTML rendering, Markdown writing, syntax/native fields, source slices, and source edits. Remaining container interactions and writer breadth keep this family partial.
- **Promotion bar:** Complete remaining blockquote and container interactions plus broader writer behavior.

#### Diagrams

- **OfficeIMO state:** Semantic fenced blocks and visual renderer hooks exist; named diagram-language mapping and a complete renderer handoff contract remain open.
- **Promotion bar:** Define named diagram-language mapping, renderer-package ownership, source/writer behavior, and focused fixtures.

#### Figures

- **OfficeIMO state:** Image and figure import plus publisher rendering paths exist; a dedicated Markdown figure syntax and its source/writer contract remain open.
- **Promotion bar:** Separate HTML-import figure recovery from authored Markdown figure syntax, then prove renderer, writer, and source behavior.

#### Generic attributes

- **OfficeIMO state:** Generic attributes are stored on semantic and syntax nodes and are source-backed for the covered heading, paragraph, code, list, table, image, definition-list, footnote, link, image-link, emphasis, and inline-code shapes. Arbitrary block and inline families remain incomplete.
- **Promotion bar:** Complete arbitrary block-family parsing, inline-family breadth, and writer/source preservation across supported shapes.

#### List extras

- **OfficeIMO state:** Opt-in alphabetic and Roman ordered markers support nested parsing, marker-style HTML, source metadata and edits, and Markdown writer preservation. Remaining edge, source-edit, and reparse coverage keeps this family partial.
- **Promotion bar:** Broaden remaining list-marker edges, native source edits, and writer reparse proof.

#### Mathematics

- **OfficeIMO state:** Math-oriented semantic and rendering hooks exist, but inline and block delimiter parsing does not yet have a complete AST, source, writer, and renderer contract.
- **Promotion bar:** Define inline and block delimiters, AST/source/native metadata, writer preservation, and renderer handoff.

#### Media links

- **OfficeIMO state:** Image and media semantics exist, but shortcut media providers do not yet have a complete parser, safe-renderer, source, and writer contract.
- **Promotion bar:** Define the provider model, safe renderer output, writer preservation, and source metadata for shortcut media links.

#### Precise source location

- **OfficeIMO state:** Syntax, semantic, native, transform, renderer, writer, and source-edit APIs expose broad normalized and original source evidence. Complete lossless trivia, original mapping, generated-node semantics, and arbitrary source edits remain partial.
- **Promotion bar:** Complete lossless trivia and original mapping, generated-node round-trip semantics, and source-edit coverage.
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
