# OfficeIMO.Markdown Markdig Extension Inventory

This report compares the Markdig `1.3.2` extension-family entry points reflected from the local comparison package with the current `OfficeIMO.Markdown` support story.

Status values:

- `Covered`: implemented and protected by focused evidence.
- `Partial`: real OfficeIMO support exists, but Markdig breadth, options, source mapping, writer behavior, or renderer behavior is incomplete.
- `Intentional`: the Markdig entry point is a bundle, helper, or renderer policy that OfficeIMO should model differently.
- `Gap`: no meaningful OfficeIMO equivalent exists yet.

Route values name the owning layer for future work, so missing behavior is fixed in the reusable engine, optional extension, renderer/host policy, or intentionally documented difference instead of drifting into ad hoc tests.

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
| Covered | 7 |
| Partial | 11 |
| Intentional | 4 |
| Gap | 11 |

## Extension Families

| Markdig entry point | Family | Status | Route | Promotion bar | OfficeIMO state | Next action |
| --- | --- | --- | --- | --- | --- | --- |
| `UseAbbreviations` | Abbreviations | `Partial` | Core parser, opt-in | Remaining Markdig abbreviation edge breadth plus promotion-level source/native coverage. | OfficeIMO has opt-in abbreviation definitions through MarkdownReaderOptions.Abbreviations, case-sensitive/later-wins document-wide definition collection, consumed definition syntax nodes, AbbreviationInline semantic nodes, HTML <abbr> rendering, Markdig comparison cases for top-level text, emphasis, link labels, blockquotes, lists, and pipe-table cells when UsePipeTables is also enabled, syntax/native metadata for visible text plus definition title source edits, nested container/table-cell AST propagation, and definition-preserving Markdown writing for parse-owned definitions with front matter and reparse-stability coverage, but broader Markdig edge breadth is not complete. | Broaden the remaining Markdig abbreviation edge cases and promote only after any remaining parser/source/writer deltas are explicit. |
| `UseAdvancedExtensions` | Advanced extension bundle | `Intentional` | Intentional bundle guard | Keep individual feature rows authoritative; do not add a broad bundle switch. | OfficeIMO should track individual feature families instead of claiming bundle parity. | Keep this row as a roll-up guard; do not implement as a broad on switch. |
| `UseAlertBlocks` | Alert blocks | `Partial` | Core parser plus renderer policy | Alert/callout AST fields, source spans, renderer callbacks, writer output, and Markdig/GFM comparison fixtures. | OfficeIMO has callout blocks and GitHub-style callout parsing, but not Markdig's alert rendering callback shape. | Align callout/alert syntax, AST fields, source spans, and renderer customization explicitly. |
| `UseAutoIdentifiers` | Auto identifiers | `Covered` | Core renderer option | Keep slug-style and source metadata fixtures current. | OfficeIMO has automatic heading ids with duplicate-slug tracking, an opt-out HTML switch, Markdig default and GitHub-compatible slug styles, GFM HTML profile wiring, heading traversal APIs, and source-backed heading syntax/native metadata. | Keep slug-style and heading-source fixtures aligned as broader renderer profiles evolve. |
| `UseAutoLinks` | Extended autolinks | `Partial` | Core parser, profile-gated | Full extended URL/email/scheme/boundary corpus, writer preservation, and native/source metadata. | OfficeIMO has profile-sensitive bare URL/email autolinks with Markdig-style previous-character, domain-without-period, query/fragment special-character, balanced-parenthesis trailing-punctuation, punctuation-before-closing-parenthesis preservation, lowercase www. prefix matching, lowercase bare scheme matching, profile-selectable bare scheme prefixes for Markdig-compatible mailto:, ftp://, and tel: behavior while OfficeIMO/GFM can keep xmpp:, default/Markdig-style HTML URL rendering can normalize Unicode hostnames to IDNA, the GFM HTML profile keeps cmark-gfm-style percent-encoded Unicode hosts, GFM/table-cell coverage exists, source-backed target and angle-marker metadata exists, and Markdown writer preservation for parsed bare and angle autolink spelling exists, but broader Markdig/GFM edge breadth is not complete. | Broaden remaining GFM/Markdig autolink edge cases before promotion. |
| `UseBootstrap` | Bootstrap renderer helpers | `Intentional` | Renderer theme policy | Keep parser parity separate from optional theme presets. | This is renderer-theme behavior rather than a core Markdown syntax family for OfficeIMO. | Keep theme/rendering presets separate from parser parity. |
| `UseCjkFriendlyEmphasis` | CJK-friendly emphasis | `Partial` | Core delimiter parser option | Delimiter-run option with CJK comparison fixtures and source-token stability. | OfficeIMO has selected CJK-adjacent emphasis regression coverage, but not a Markdig-compatible CJK emphasis option. | Fold into the CommonMark emphasis delimiter rewrite and keep CJK-specific fixtures explicit. |
| `UseCitations` | Citations | `Gap` | Optional parser extension, deferred | Citation AST, renderer/writer contract, and real consumer need after core/GFM closure. | No citation AST or renderer contract exists. | Decide whether citations are in scope after core CommonMark/GFM closure. |
| `UseCustomContainers` | Custom containers | `Gap` | Core extension seam plus optional built-in parser | Container parser contract, child-block source mapping, renderer/writer source-slice APIs, and Markdig fixtures. | OfficeIMO has semantic block extension seams, but not Markdig custom container syntax parity. | Route to block parser extensions plus renderer/writer source-slice contracts. |
| `UseDefinitionLists` | Definition lists | `Partial` | Core parser, opt-in/profile-gated | Remaining source-map and writer edge breadth for marker groups, lazy continuation, nested blocks, and reparsing. | OfficeIMO has structured definition-list AST, Markdig-style colon-marker term grouping, multiple-definition parsing, source/native projection, profile-correct HTML comparison coverage, grouped Markdown writer preservation for reparsing, Markdig lazy paragraph, nested block, loose-definition, edge-continuation, and empty-marker first-continuation coverage, loose-definition writer preservation, and blank-separated marker-group writer preservation, but full source-map and writer edge breadth is not closed. | Broaden remaining Markdig definition-list source-map and writer edge cases before promotion. |
| `UseDiagrams` | Diagrams | `Partial` | Renderer/host policy over semantic fences | Named diagram language mapping, renderer package ownership, source/writer behavior, and comparison fixtures. | OfficeIMO has semantic fenced blocks and visual renderer hooks, but not Markdig diagram extension parity. | Compare Mermaid/Nomnom-style cases and decide renderer-package ownership. |
| `UseEmojiAndSmiley` | Emoji and smiley | `Gap` | Optional inline transform | Shortcode/smiley tables, opt-in profile behavior, source metadata, writer rules, and no conflict with Unicode normalization. | OfficeIMO has emoji word-join normalization only, not shortcode/smiley expansion. | Keep normalization separate from an optional inline replacement extension. |
| `UseEmphasisExtras` | Emphasis extras | `Covered` | Core inline parser, profile-gated | Keep delimiter fixtures aligned with GFM and lossless source work. | OfficeIMO has strikethrough, inserted-text, highlight/mark, superscript, and subscript inline nodes with Markdig comparison cases, parser-owned source marker metadata, native projection, HTML rendering, Markdown writing, and explicit GFM single-tilde strikethrough profile coverage. | Keep emphasis-extra delimiter cases aligned as broader GFM and lossless trivia coverage expands. |
| `UseFigures` | Figures | `Partial` | Core image AST plus optional parser syntax | Separate HTML-import figure recovery from Markdown figure syntax, then prove renderer/writer/source behavior. | OfficeIMO has image/figure import and publisher figure rendering paths, but not Markdig figure syntax parity. | Separate HTML-import figure recovery from Markdown parser extension support. |
| `UseFooters` | Footers | `Gap` | Deferred document semantics | Only implement if Markdown-authored footer semantics become a real document requirement. | No footer block parser or semantic node exists. | Leave out of scope unless document footer semantics become a Markdown requirement. |
| `UseFootnotes` | Footnotes | `Covered` | Core parser, GFM profile | Keep GFM footnote fixture corpus and structured writer proof current. | OfficeIMO has GFM footnote parsing and GitHub HTML rendering for first-reference ordering, repeated-reference backrefs, missing/unused definitions, nested block bodies, source/native label and marker spans, and structured Markdown writer roundtrip proof. | Keep the GFM footnote fixture corpus and structured-body writer coverage current. |
| `UseGenericAttributes` | Generic attributes | `Partial` | Core AST/source architecture | Attribute storage on semantic and syntax nodes, renderer/writer propagation, and source-edit coverage across blocks/inlines. | OfficeIMO captures fenced-code brace metadata, but not generic attributes on arbitrary blocks/inlines. | Design attribute storage on semantic and syntax nodes before broad support. |
| `UseGlobalization` | Globalization | `Gap` | Deferred compatibility option | Only implement with a concrete culture-sensitive behavior contract and fixtures. | No Markdig globalization extension equivalent is documented for OfficeIMO. | Revisit only if a real consumer needs culture-sensitive Markdown behavior. |
| `UseGridTables` | Grid tables | `Gap` | Optional block parser extension | Grid table AST/source model, HTML/Markdown writer behavior, malformed-table fallback, and Markdig/Pandoc-style fixtures. | OfficeIMO has pipe tables only; grid table parsing is absent. | Decide if grid tables belong in core or an optional extension package. |
| `UseJiraLinks` | Jira links | `Gap` | Optional link inline extension | Configurable issue-key resolver, renderer policy, writer preservation, and source metadata without affecting ordinary text. | No Jira-link shortcut parser exists. | Treat as optional link extension after core link/source mapping is stable. |
| `UseListExtras` | List extras | `Gap` | Optional parser work after list cleanup | Inventory Markdig list-extra syntax, choose supported forms, and prove canonical ListItem/source behavior. | OfficeIMO list work is focused on CommonMark/GFM task behavior, not Markdig list extras. | Inventory Markdig list-extra syntax before choosing scope. |
| `UseMathematics` | Mathematics | `Partial` | Optional parser plus renderer/host policy | Inline/block math delimiters, AST/source/native metadata, writer preservation, and renderer handoff contract. | OfficeIMO has math-oriented semantic/rendering paths through host options, but not Markdig math delimiter parity. | Define math parser ownership and compare inline/block math fixtures. |
| `UseMediaLinks` | Media links | `Partial` | Renderer/host policy with optional link parser | Provider model, safe renderer output, writer preservation, and source metadata for shortcut media links. | OfficeIMO has image/media document semantics, but not Markdig media-link provider parity. | Route shortcut media providers through renderer/host extension seams if in scope. |
| `UseNonAsciiNoEscape` | Non-ASCII no-escape rendering | `Intentional` | Renderer escaping policy | Document profile differences when output claims broaden. | OfficeIMO keeps escaping behavior profile/renderer-owned instead of mirroring this Markdig switch. | Document any renderer escaping profile differences when output claims broaden. |
| `UsePipeTables` | Pipe tables | `Covered` | Core parser, GFM profile | Keep GFM table corpus and table-cell source-edit coverage current. | OfficeIMO has GFM pipe-table parsing with delimiter-row validation, escaped/code-span pipe handling, body-row padding/truncation, container ownership, semantic table/cell AST, syntax/native source spans, GitHub HTML rendering, and aligned Markdown writer roundtrip proof. | Keep the GFM table fixture corpus and table-cell source-edit coverage current. |
| `UsePragmaLines` | Pragma lines | `Gap` | Deferred metadata parser | Only implement if a concrete workflow needs pragma metadata with source-preserving writer behavior. | No pragma-line parser or semantic contract exists. | Leave out of core unless a concrete document workflow needs it. |
| `UsePreciseSourceLocation` | Precise source location | `Partial` | Cross-cutting core source architecture | Complete lossless trivia/original mapping, generated-node diagnostics, and source-edit coverage before claiming parity. | OfficeIMO has syntax/source/native spans and source slices, but full lossless trivia/original mapping is still partial. | Continue Phase 3 source-map and trivia work before claiming parity. |
| `UseReferralLinks` | Referral links | `Gap` | Renderer policy | Only implement as an opt-in link-rendering policy with safe defaults and tests. | No Markdig-compatible referral-link renderer policy exists. | Treat as renderer policy work if requested. |
| `UseSelfPipeline` | Self pipeline | `Intentional` | Intentional composition difference | Keep extension composition in OfficeIMO options rather than mirroring Markdig pipeline helpers. | This is a Markdig pipeline composition helper, not a Markdown feature OfficeIMO should mirror directly. | Keep extension composition in OfficeIMO reader/render/write options. |
| `UseSmartyPants` | SmartyPants | `Gap` | Optional inline transform | Smart punctuation transform with opt-in profile, source/edit behavior, writer policy, and escaping rules. | No SmartyPants inline transform exists. | Consider as an optional inline transform after delimiter parsing stabilizes. |
| `UseSoftlineBreakAsHardlineBreak` | Soft line break as hard line break | `Covered` | Core parser option | Keep option covered alongside paragraph/list source-map and writer fixtures. | OfficeIMO exposes an explicit reader option that parses ordinary paragraph soft breaks as hard breaks while keeping CommonMark/GFM defaults unchanged, rendering HTML breaks, writing normalized hard-break markdown, and avoiding fake source marker metadata. | Keep the option covered alongside paragraph/list source-map and writer fixtures. |
| `UseTaskLists` | Task lists | `Covered` | Core parser, GFM profile | Keep GFM task marker source-edit coverage current. | OfficeIMO has GFM task-list parsing for checked, unchecked, uppercase, nested, and invalid tight-marker cases; semantic AST flags; exact marker source spans; native snapshots/source edits; GitHub HTML rendering; and Markdown writer roundtrip proof. | Keep the GFM fixture corpus and marker source-edit coverage current. |
| `UseYamlFrontMatter` | YAML front matter | `Covered` | Core parser, OfficeIMO profile | Keep raw YAML helpers and front-matter source-edit fixtures aligned with lossless work. | OfficeIMO preserves YAML front matter as a top-of-document raw YAML AST payload with body and fence source spans, structured key/value helpers for simple entries, native source fields and snapshots, HTML omission, and Markdown writer roundtrip behavior. | Keep raw YAML, parsed-entry helpers, and front-matter source-edit fixtures aligned as lossless trivia work expands. |

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
- Use the `Route` and `Promotion bar` columns before implementation so every slice moves the right owner instead of creating another local workaround.
- Add fixtures or engine work by row, not by nearby test names.
