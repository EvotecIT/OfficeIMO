# OfficeIMO.Markdown Markdig Extension Inventory

This report compares the Markdig `1.3.2` extension-family entry points reflected from the local comparison package with the current `OfficeIMO.Markdown` support story.

Status values:

- `Covered`: implemented and protected by focused evidence.
- `Partial`: real OfficeIMO support exists, but Markdig breadth, options, source mapping, writer behavior, or renderer behavior is incomplete.
- `Intentional`: the Markdig entry point is a bundle, helper, or renderer policy that OfficeIMO should model differently.
- `Gap`: no meaningful OfficeIMO equivalent exists yet.

Route values name the owning layer for future work. Scope decisions collapse those routes into execution buckets, so missing behavior is fixed in the reusable engine, optional extension, renderer/host policy, deferred backlog, or intentionally documented difference instead of drifting into ad hoc tests.

Refresh command:

```powershell
$env:OFFICEIMO_UPDATE_MARKDIG_INVENTORY = '1'
dotnet test OfficeIMO.Tests\OfficeIMO.Tests.csproj --framework net8.0 --filter "FullyQualifiedName~Markdown_Markdig_Extension_Inventory_Tests"
Remove-Item Env:\OFFICEIMO_UPDATE_MARKDIG_INVENTORY
```

## Summary

| Metric | Count |
| --- | ---: |
| Markdig extension-family rows | 33 |
| Covered | 10 |
| Partial | 9 |
| Intentional | 3 |
| Gap | 11 |

## Extension Families

| Markdig entry point | Family | Status | Scope decision | Route | Promotion bar | OfficeIMO state | Next action |
| --- | --- | --- | --- | --- | --- | --- | --- |
| `UseAbbreviations` | Abbreviations | `Covered` | Core engine | Core parser, opt-in | Keep Markdig comparison, syntax/native/source-edit, and writer fixtures current. | OfficeIMO has opt-in abbreviation definitions through MarkdownReaderOptions.Abbreviations, case-sensitive/later-wins document-wide definition collection, consumed definition syntax nodes including empty-title and list-item-contained definitions, AbbreviationInline semantic nodes, HTML <abbr> rendering, Markdig comparison cases for top-level text, Unicode text, unresolved bracket text, emphasis, link labels, blockquotes, lists, list-item definitions, dash/opening-punctuation boundaries, and pipe-table cells when UsePipeTables is also enabled, syntax/native metadata for visible text plus definition title source edits, nested container/table-cell AST propagation, definition-preserving Markdown writing for parse-owned definitions with front matter, empty-title, list-contained definitions, and reparse-stability coverage, list-contained definition source-token navigation and native title edits, and Markdig-style literal non-ASCII text rendering through HtmlOptions.EscapeNonAsciiText = false. | Keep abbreviation comparison, list-contained source-token, native source-edit, and writer fixtures aligned as lossless trivia work expands. |
| `UseAdvancedExtensions` | Advanced extension bundle | `Intentional` | Intentional difference | Intentional bundle guard | Keep individual feature rows authoritative; do not add a broad bundle switch. | OfficeIMO should track individual feature families instead of claiming bundle parity. | Keep this row as a roll-up guard; do not implement as a broad on switch. |
| `UseAlertBlocks` | Alert blocks | `Partial` | Renderer/host policy | Core parser plus renderer policy | Alert/callout AST fields, source spans, renderer callbacks, writer output, and Markdig/GFM comparison fixtures. | OfficeIMO has callout blocks and GitHub-style callout parsing, but not Markdig's alert rendering callback shape. | Align callout/alert syntax, AST fields, source spans, and renderer customization explicitly. |
| `UseAutoIdentifiers` | Auto identifiers | `Covered` | Renderer/host policy | Core renderer option | Keep slug-style and source metadata fixtures current. | OfficeIMO has automatic heading ids with duplicate-slug tracking, an opt-out HTML switch, Markdig default and GitHub-compatible slug styles, GFM HTML profile wiring, heading traversal APIs, and source-backed heading syntax/native metadata. | Keep slug-style and heading-source fixtures aligned as broader renderer profiles evolve. |
| `UseAutoLinks` | Extended autolinks | `Covered` | Core engine | Core parser, profile-gated | Keep Markdig/GFM autolink fixtures, source metadata, writer preservation, and URL/text-rendering profile evidence current. | OfficeIMO has profile-sensitive bare URL/email autolinks with Markdig-style previous-character, domain-without-period, query/fragment special-character, balanced-parenthesis trailing-punctuation, punctuation-before-closing-parenthesis preservation, single trailing punctuation/underscore trimming, optional trailing semicolon retention, optional trailing quote retention with paired single-quote literal fallback, lowercase www. prefix matching, optional www-host and url-host underscore rejection, optional user-info authority rejection for Markdig-compatible http/www/ftp literals, optional closing-bracket URL consumption, lowercase bare scheme matching, profile-selectable bare scheme prefixes for Markdig-compatible mailto:, ftp://, and tel: behavior while OfficeIMO/GFM can keep xmpp:, apostrophe-started bare scheme literal fallback, bare mailto path/query/fragment targets with address-only display, optional Markdig-compatible mailto semicolon and address-only colon/dash handling, Markdig-compatible href tilde and quote percent-encoding, default/Markdig-style IDNA host rendering, cmark-gfm-style percent-encoded Unicode hosts through the GFM HTML profile, literal non-ASCII display text through the explicit HTML text policy, GFM/table-cell coverage, source-backed target and angle-marker metadata, and Markdown writer preservation for parsed bare and angle autolink spelling. The focused Markdig AutoLinks and AutoLinks+PipeTables comparison lanes pass across the covered option matrix. | Keep broader GFM fixture breadth separate from the Markdig UseAutoLinks row. |
| `UseBootstrap` | Bootstrap renderer helpers | `Intentional` | Renderer/host policy | Renderer theme policy | Keep parser parity separate from optional theme presets. | This is renderer-theme behavior rather than a core Markdown syntax family for OfficeIMO. | Keep theme/rendering presets separate from parser parity. |
| `UseCjkFriendlyEmphasis` | CJK-friendly emphasis | `Partial` | Core engine | Core delimiter parser option | Delimiter-run option with CJK comparison fixtures and source-token stability. | OfficeIMO has selected CJK-adjacent emphasis regression coverage, but not a Markdig-compatible CJK emphasis option. | Fold into the CommonMark emphasis delimiter rewrite and keep CJK-specific fixtures explicit. |
| `UseCitations` | Citations | `Gap` | Deferred | Optional parser extension, deferred | Citation AST, renderer/writer contract, and real consumer need after core/GFM closure. | No citation AST or renderer contract exists. | Decide whether citations are in scope after core CommonMark/GFM closure. |
| `UseCustomContainers` | Custom containers | `Gap` | Core engine | Core extension seam plus optional built-in parser | Container parser contract, child-block source mapping, renderer/writer source-slice APIs, and Markdig fixtures. | OfficeIMO has semantic block extension seams, but not Markdig custom container syntax parity. | Route to block parser extensions plus renderer/writer source-slice contracts. |
| `UseDefinitionLists` | Definition lists | `Partial` | Core engine | Core parser, opt-in/profile-gated | Remaining source-map and writer edge breadth for marker groups, lazy continuation, loose definitions, and reparsing. | OfficeIMO has structured definition-list AST, Markdig-style colon-marker term grouping, multiple-definition parsing, source/native projection, profile-correct HTML comparison coverage, grouped Markdown writer preservation for reparsing, Markdig lazy paragraph, nested block, loose-definition, edge-continuation, setext-continuation, and empty-marker first-continuation coverage, parsed and generated definition marker syntax tokens, native source-backed marker fields/source edits, loose-definition writer preservation, blank-separated marker-group writer preservation, tight nested-list writer preservation, setext-continuation writer reparse proof, and typed plus source-field multiline definition-body edits that keep continuation indentation valid for simple and marker forms, but full source-map and writer edge breadth is not closed. | Broaden remaining Markdig definition-list source-map and writer edge cases before promotion. |
| `UseDiagrams` | Diagrams | `Partial` | Renderer/host policy | Renderer/host policy over semantic fences | Named diagram language mapping, renderer package ownership, source/writer behavior, and comparison fixtures. | OfficeIMO has semantic fenced blocks and visual renderer hooks, but not Markdig diagram extension parity. | Compare Mermaid/Nomnom-style cases and decide renderer-package ownership. |
| `UseEmojiAndSmiley` | Emoji and smiley | `Gap` | Optional extension | Optional inline transform | Shortcode/smiley tables, opt-in profile behavior, source metadata, writer rules, and no conflict with Unicode normalization. | OfficeIMO has emoji word-join normalization only, not shortcode/smiley expansion. | Keep normalization separate from an optional inline replacement extension. |
| `UseEmphasisExtras` | Emphasis extras | `Covered` | Core engine | Core inline parser, profile-gated | Keep delimiter fixtures aligned with GFM and lossless source work. | OfficeIMO has strikethrough, inserted-text, highlight/mark, superscript, and subscript inline nodes with Markdig comparison cases, parser-owned source marker metadata, native projection, HTML rendering, Markdown writing, and explicit GFM single-tilde strikethrough profile coverage. | Keep emphasis-extra delimiter cases aligned as broader GFM and lossless trivia coverage expands. |
| `UseFigures` | Figures | `Partial` | Core engine | Core image AST plus optional parser syntax | Separate HTML-import figure recovery from Markdown figure syntax, then prove renderer/writer/source behavior. | OfficeIMO has image/figure import and publisher figure rendering paths, but not Markdig figure syntax parity. | Separate HTML-import figure recovery from Markdown parser extension support. |
| `UseFooters` | Footers | `Gap` | Deferred | Deferred document semantics | Only implement if Markdown-authored footer semantics become a real document requirement. | No footer block parser or semantic node exists. | Leave out of scope unless document footer semantics become a Markdown requirement. |
| `UseFootnotes` | Footnotes | `Covered` | Core engine | Core parser, GFM profile | Keep GFM footnote fixture corpus and structured writer proof current. | OfficeIMO has GFM footnote parsing and GitHub HTML rendering for first-reference ordering, repeated-reference backrefs, missing/unused definitions, nested block bodies, source/native label and marker spans, and structured Markdown writer roundtrip proof. | Keep the GFM footnote fixture corpus and structured-body writer coverage current. |
| `UseGenericAttributes` | Generic attributes | `Partial` | Core engine | Core AST/source architecture | Remaining arbitrary block-family parsing, complete inline-family breadth, broader Markdown writer preservation, and full token-level syntax/source coverage across blocks/inlines. | OfficeIMO now has generic attribute storage on semantic MarkdownObject nodes and MarkdownSyntaxNode nodes, with fenced-code id/classes/attributes projected from MarkdownCodeFenceInfo through ordinary CodeBlock and SemanticFencedBlock parser paths. Default fenced-code HTML rendering projects those attributes onto the HTML pre wrapper, and semantic fenced-block code fallback renderers receive the attributed CodeBlock. Opt-in MarkdownReaderOptions.GenericAttributes now parses Markdig-style trailing attribute blocks for ATX headings, Setext headings, and paragraphs, plus no-space inline attribute blocks on links, images, emphasis, strong, code spans, strikethrough, highlight, inserted, superscript, and subscript nodes. Those attributes flow through semantic/syntax storage, default HTML rendering, Markdown writing, and reparse proof for the covered shapes. Generic attribute blocks on covered block and inline shapes are also source-backed in native projections as `attributes` source fields/metadata, with preserved-trivia source-edit proof. It still does not parse generic attributes for arbitrary block families, and full syntax-token/source coverage across arbitrary blocks and inlines is not complete. | Extend the shared attribute parser/writer to more block families, then promote once writer/source propagation and token-level coverage are proven across arbitrary blocks and inlines. |
| `UseGlobalization` | Globalization | `Gap` | Deferred | Deferred compatibility option | Only implement with a concrete culture-sensitive behavior contract and fixtures. | No Markdig globalization extension equivalent is documented for OfficeIMO. | Revisit only if a real consumer needs culture-sensitive Markdown behavior. |
| `UseGridTables` | Grid tables | `Gap` | Optional extension | Optional block parser extension | Grid table AST/source model, HTML/Markdown writer behavior, malformed-table fallback, and Markdig/Pandoc-style fixtures. | OfficeIMO has pipe tables only; grid table parsing is absent. | Decide if grid tables belong in core or an optional extension package. |
| `UseJiraLinks` | Jira links | `Gap` | Optional extension | Optional link inline extension | Configurable issue-key resolver, renderer policy, writer preservation, and source metadata without affecting ordinary text. | No Jira-link shortcut parser exists. | Treat as optional link extension after core link/source mapping is stable. |
| `UseListExtras` | List extras | `Gap` | Optional extension | Optional parser work after list cleanup | Inventory Markdig list-extra syntax, choose supported forms, and prove canonical ListItem/source behavior. | OfficeIMO list work is focused on CommonMark/GFM task behavior, not Markdig list extras. | Inventory Markdig list-extra syntax before choosing scope. |
| `UseMathematics` | Mathematics | `Partial` | Renderer/host policy | Optional parser plus renderer/host policy | Inline/block math delimiters, AST/source/native metadata, writer preservation, and renderer handoff contract. | OfficeIMO has math-oriented semantic/rendering paths through host options, but not Markdig math delimiter parity. | Define math parser ownership and compare inline/block math fixtures. |
| `UseMediaLinks` | Media links | `Partial` | Renderer/host policy | Renderer/host policy with optional link parser | Provider model, safe renderer output, writer preservation, and source metadata for shortcut media links. | OfficeIMO has image/media document semantics, but not Markdig media-link provider parity. | Route shortcut media providers through renderer/host extension seams if in scope. |
| `UseNonAsciiNoEscape` | Non-ASCII no-escape rendering | `Covered` | Renderer/host policy | Renderer escaping policy | Keep renderer escape-policy coverage aligned as new HTML output paths are introduced. | OfficeIMO exposes HtmlOptions.EscapeNonAsciiText so Markdig/GFM-style HTML output can keep non-ASCII visible text literal while preserving the historical .NET encoder behavior by default. The GitHub-flavored HTML profile enables literal non-ASCII text rendering, and inline text, link display text, code block text, captions, simple quote text, abbreviation output, TOC labels/titles/anchors, heading helper text and generated attributes, page titles, body/head/asset metadata, footnote ids/backrefs, code and callout classes, link/image policy attributes, raw-HTML escape output, image-blocked placeholder text, portable fallback helper text, image/link title and alt attributes, picture-source descriptor attributes, sanitizer escape output, and custom HTML render-extension helper APIs use the explicit policy. URL-bearing attributes remain routed through the URL attribute encoder. | Keep direct encoder audits and focused non-ASCII render-policy tests current when adding new HTML output paths. |
| `UsePipeTables` | Pipe tables | `Covered` | Core engine | Core parser, GFM profile | Keep GFM table corpus and table-cell source-edit coverage current. | OfficeIMO has GFM pipe-table parsing with delimiter-row validation, escaped/code-span pipe handling, body-row padding/truncation, container ownership, semantic table/cell AST, syntax/native source spans, GitHub HTML rendering, and aligned Markdown writer roundtrip proof. | Keep the GFM table fixture corpus and table-cell source-edit coverage current. |
| `UsePragmaLines` | Pragma lines | `Gap` | Deferred | Deferred metadata parser | Only implement if a concrete workflow needs pragma metadata with source-preserving writer behavior. | No pragma-line parser or semantic contract exists. | Leave out of core unless a concrete document workflow needs it. |
| `UsePreciseSourceLocation` | Precise source location | `Partial` | Core engine | Cross-cutting core source architecture | Complete lossless trivia/original mapping, generated-node diagnostics, and source-edit coverage before claiming parity. | OfficeIMO has syntax/source/native spans and source slices, but full lossless trivia/original mapping is still partial. | Continue Phase 3 source-map and trivia work before claiming parity. |
| `UseReferralLinks` | Referral links | `Gap` | Renderer/host policy | Renderer policy | Only implement as an opt-in link-rendering policy with safe defaults and tests. | No Markdig-compatible referral-link renderer policy exists. | Treat as renderer policy work if requested. |
| `UseSelfPipeline` | Self pipeline | `Intentional` | Intentional difference | Intentional composition difference | Keep extension composition in OfficeIMO options rather than mirroring Markdig pipeline helpers. | This is a Markdig pipeline composition helper, not a Markdown feature OfficeIMO should mirror directly. | Keep extension composition in OfficeIMO reader/render/write options. |
| `UseSmartyPants` | SmartyPants | `Gap` | Optional extension | Optional inline transform | Smart punctuation transform with opt-in profile, source/edit behavior, writer policy, and escaping rules. | No SmartyPants inline transform exists. | Consider as an optional inline transform after delimiter parsing stabilizes. |
| `UseSoftlineBreakAsHardlineBreak` | Soft line break as hard line break | `Covered` | Core engine | Core parser option | Keep option covered alongside paragraph/list source-map and writer fixtures. | OfficeIMO exposes an explicit reader option that parses ordinary paragraph soft breaks as hard breaks while keeping CommonMark/GFM defaults unchanged, rendering HTML breaks, writing normalized hard-break markdown, and avoiding fake source marker metadata. | Keep the option covered alongside paragraph/list source-map and writer fixtures. |
| `UseTaskLists` | Task lists | `Covered` | Core engine | Core parser, GFM profile | Keep GFM task marker source-edit coverage current. | OfficeIMO has GFM task-list parsing for checked, unchecked, uppercase, nested, and invalid tight-marker cases; semantic AST flags; exact marker source spans; native snapshots/source edits; GitHub HTML rendering; and Markdown writer roundtrip proof. | Keep the GFM fixture corpus and marker source-edit coverage current. |
| `UseYamlFrontMatter` | YAML front matter | `Covered` | Core engine | Core parser, OfficeIMO profile | Keep raw YAML helpers and front-matter source-edit fixtures aligned with lossless work. | OfficeIMO preserves YAML front matter as a top-of-document raw YAML AST payload with body and fence source spans, structured key/value helpers for simple entries, native source fields and snapshots, HTML omission, and Markdown writer roundtrip behavior. | Keep raw YAML, parsed-entry helpers, and front-matter source-edit fixtures aligned as lossless trivia work expands. |

## Reflected Pipeline Entry Points

These public Markdig pipeline-builder methods are reflected from the local package so package upgrades cannot silently add a new `Use*` extension family without updating this report.

| Method | Tracked as extension family |
| --- | --- |
| `Configure` | No |
| `ConfigureNewLine` | No |
| `DisableHeadings` | No |
| `DisableHtml` | No |
| `EnableTrackTrivia` | No |
| `Use` | No |
| `UseAbbreviations` | Yes |
| `UseAdvancedExtensions` | Yes |
| `UseAlertBlocks` | Yes |
| `UseAutoIdentifiers` | Yes |
| `UseAutoLinks` | Yes |
| `UseBootstrap` | Yes |
| `UseCitations` | Yes |
| `UseCjkFriendlyEmphasis` | Yes |
| `UseCustomContainers` | Yes |
| `UseDefinitionLists` | Yes |
| `UseDiagrams` | Yes |
| `UseEmojiAndSmiley` | Yes |
| `UseEmphasisExtras` | Yes |
| `UseFigures` | Yes |
| `UseFooters` | Yes |
| `UseFootnotes` | Yes |
| `UseGenericAttributes` | Yes |
| `UseGlobalization` | Yes |
| `UseGridTables` | Yes |
| `UseJiraLinks` | Yes |
| `UseListExtras` | Yes |
| `UseMathematics` | Yes |
| `UseMediaLinks` | Yes |
| `UseNonAsciiNoEscape` | Yes |
| `UsePipeTables` | Yes |
| `UsePragmaLines` | Yes |
| `UsePreciseSourceLocation` | Yes |
| `UseReferralLinks` | Yes |
| `UseSelfPipeline` | Yes |
| `UseSmartyPants` | Yes |
| `UseSoftlineBreakAsHardlineBreak` | Yes |
| `UseTaskLists` | Yes |
| `UseYamlFrontMatter` | Yes |

## Next Use

- Use this inventory to decide whether an upcoming slice is parser grammar, AST/source mapping, renderer/writer behavior, extension seam work, or an intentional profile difference.
- Keep `Partial` rows honest: promote them to `Covered` only when parser, AST/source, renderer, writer, and fixture evidence all match the claimed scope.
- Use the `Scope decision`, `Route`, and `Promotion bar` columns before implementation so every slice moves the right owner instead of creating another local workaround.
- Add fixtures or engine work by row, not by nearby test names.
