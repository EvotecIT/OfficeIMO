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

These entries explain what works today and what remains outside each `Partial` compatibility family.

#### Custom containers

- **Current behavior:** Opt-in colon-fenced containers support root, nested, blockquote-contained, and list-contained shapes with child parsing, HTML rendering, Markdown writing, syntax fields, source slices, source edits, and stable reparse.
- **Limit:** Other optional container shapes are not recognized. Unsupported syntax remains literal rather than receiving partial parse, render, or source-edit behavior.

#### Diagrams

- **Current behavior:** Semantic fenced blocks and visual renderer hooks exist; named diagram-language mapping and a complete renderer handoff contract remain open.
- **Limit:** Named diagram-language parsing and a complete renderer handoff are not available.

#### Figures

- **Current behavior:** Image and figure import plus publisher rendering paths exist; a dedicated Markdown figure syntax and its source/writer contract remain open.
- **Limit:** HTML figure recovery does not provide a dedicated authored Markdown figure syntax or a source-preserving writer contract.

#### Generic attributes

- **Current behavior:** Generic attributes have end-to-end ownership for every supported target family, including callouts, details blocks, and custom containers: semantic and syntax storage, exact source fields, HTML projection, Markdown writing, source edits, and stable reparse. Targets outside that declared set remain literal or deliberately consumed according to the documented profile boundary.
- **Limit:** Only the documented target families accept attributes. Other targets remain literal or follow the selected profile's documented consumption rule.

#### List extras

- **Current behavior:** Opt-in alphabetic and Roman ordered markers support nested parsing, marker-style HTML, source metadata and edits, and Markdown writer preservation.
- **Limit:** Some list-marker edge cases and source-edit round trips remain outside the supported subset.

#### Mathematics

- **Current behavior:** Math-oriented semantic and rendering hooks exist, but inline and block delimiter parsing does not yet have a complete AST, source, writer, and renderer contract.
- **Limit:** Built-in inline and block math delimiter parsing, source metadata, Markdown writing, and renderer handoff are not available as one complete contract.

#### Media links

- **Current behavior:** Image and media semantics exist, but shortcut media providers do not yet have a complete parser, safe-renderer, source, and writer contract.
- **Limit:** Shortcut media-provider syntax does not yet have a complete parser, safe HTML output, source mapping, and Markdown writer contract.

#### Precise source location

- **Current behavior:** The source contract is field-bounded: documented source-backed fields expose normalized spans, exact or line-ending-equivalent original mappings, stable semantic associations, and source edits. Generated or transformed nodes are spanless with machine-readable unavailable reasons, and arbitrary semantic edits use normalized writing.
- **Limit:** Exact locations are not returned for arbitrary fields, transformed nodes, or generated nodes. Arbitrary semantic edits use normalized Markdown writing rather than a lossless source patch.
<!-- extension-partial-boundaries:end -->

## Standards profiles

| Capability | CommonMark profile | GFM profile | OfficeIMO profile | Status and boundary |
| --- | --- | --- | --- | --- |
| ATX and Setext headings | Yes | Yes | Yes | Covered, including heading/source marker spans and Setext source edits |
| Paragraphs, entities, escapes, hard/soft breaks | Yes | Yes | Yes | Covered grammar; exact source claims are limited to the documented source-backed fields |
| Thematic breaks | Yes | Yes | Yes | Covered with marker syntax and native projection |
| Fenced and indented code | Yes | Yes | Yes, plus semantic fenced blocks | Partial; core grammar is broad and semantic fenced blocks are an OfficeIMO extension |
| Blockquotes | Yes | Yes | Yes, plus opt-in callout recognition | Covered grammar; plain blockquotes keep literal attribute boundaries while callouts own their documented source fields |
| Ordered and unordered lists | Yes | Yes | Yes | Partial; parsing is broad while canonical subobject/source ownership continues to improve |
| Emphasis, strong, escapes, and code spans | Yes | Yes | Yes | Covered grammar; delimiter locations are exact only for documented source-backed fields |
| Links, references, images, and autolinks | Yes | Yes | Yes | Partial; profile-specific bare-link and standalone-image behavior is explicit |
| Raw HTML | Optional | Optional with GFM tag-filter policy | Optional | Grammar and output security are separate options; the CommonMark 0.31.2 inventory is complete |
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
| Syntax model | Final syntax tree with source spans, token/field nodes, semantic association, and navigation helpers | Covered for documented source-backed fields; unlisted and generated fields are explicitly unavailable |
| Original source | Normalized slices are available for span-backed nodes; exact or line-ending-equivalent original slices are available when trivia is preserved | Covered for the bounded field contract; unsafe or unavailable mappings return machine-readable reasons |
| Round-trip writer | Unchanged trivia-backed input can be emitted byte-for-byte; safe native edits apply to original input; fallback produces diagnostics | Covered for unchanged input and validated source-field edits; arbitrary semantic mutation intentionally uses normalized writing |
| Parser extensions | Ordered block, fenced-block, and inline parser seams with explicit fallback and source-aware contexts | Covered for the public extension contract |
| Transforms | Ordered post-parse transforms with source-impact diagnostics | Covered for the current transform contract |
| Renderer/writer overrides | Type-targeted and syntax-kind-targeted HTML/Markdown overrides with nested rendering contexts | Covered for the public extension contract |
| Security profiles | Raw HTML, URL handling, images, sanitizer behavior, and host renderer policy are explicit | Covered for the documented profiles |

## Deliberate boundaries

- Trivia and delimiter locations are claimed only for documented source-backed fields; unsupported optional syntax and unlisted fields remain literal or unavailable.
- Original-to-normalized mapping succeeds only when exact or line-ending-equivalent. CRLF/LF/CR and tab-aware mapped fields are covered; transformed and generated nodes return an unavailable reason.
- General byte-preserving writing after arbitrary tree mutation is not claimed.
- Optional grid tables, mathematics, media providers, figures, and diagram languages require an explicit product contract before they become parser defaults.
- HTML sanitization is a rendering policy, not part of Markdown grammar.
- Performance results are evidence for measured corpora and options, not a semantic compatibility claim.
