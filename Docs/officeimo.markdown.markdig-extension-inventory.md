# OfficeIMO.Markdown Markdig Extension Inventory

This report compares the Markdig `1.3.2` extension-family entry points reflected from the local comparison package with the current `OfficeIMO.Markdown` support story.

Status values:

- `Covered`: implemented and protected by focused evidence.
- `Partial`: real OfficeIMO support exists, but Markdig breadth, options, source mapping, writer behavior, or renderer behavior is incomplete.
- `Intentional`: the Markdig entry point is a bundle, helper, or renderer policy that OfficeIMO should model differently.
- `Gap`: no meaningful OfficeIMO equivalent exists yet.

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
| Partial | 10 |
| Intentional | 4 |
| Gap | 12 |

## Extension Families

| Markdig entry point | Family | Status | OfficeIMO state | Next action |
| --- | --- | --- | --- | --- |
| `UseAbbreviations` | Abbreviations | `Gap` | No OfficeIMO abbreviation parser or renderer contract exists yet. | Decide whether abbreviation expansion belongs in core or an optional inline extension. |
| `UseAdvancedExtensions` | Advanced extension bundle | `Intentional` | OfficeIMO should track individual feature families instead of claiming bundle parity. | Keep this row as a roll-up guard; do not implement as a broad on switch. |
| `UseAlertBlocks` | Alert blocks | `Partial` | OfficeIMO has callout blocks and GitHub-style callout parsing, but not Markdig's alert rendering callback shape. | Align callout/alert syntax, AST fields, source spans, and renderer customization explicitly. |
| `UseAutoIdentifiers` | Auto identifiers | `Covered` | OfficeIMO has automatic heading ids with duplicate-slug tracking, an opt-out HTML switch, Markdig default and GitHub-compatible slug styles, GFM HTML profile wiring, heading traversal APIs, and source-backed heading syntax/native metadata. | Keep slug-style and heading-source fixtures aligned as broader renderer profiles evolve. |
| `UseAutoLinks` | Extended autolinks | `Partial` | OfficeIMO has profile-sensitive bare URL/email autolinks with Markdig-style previous-character, domain-without-period, query/fragment special-character, ftp://, and tel: behavior, GFM coverage, source-backed target and angle-marker metadata, and Markdown writer preservation for parsed bare and angle autolink spelling, but broader Markdig/GFM edge breadth is not complete. | Broaden remaining GFM/Markdig autolink edge cases before promotion. |
| `UseBootstrap` | Bootstrap renderer helpers | `Intentional` | This is renderer-theme behavior rather than a core Markdown syntax family for OfficeIMO. | Keep theme/rendering presets separate from parser parity. |
| `UseCjkFriendlyEmphasis` | CJK-friendly emphasis | `Partial` | OfficeIMO has selected CJK-adjacent emphasis regression coverage, but not a Markdig-compatible CJK emphasis option. | Fold into the CommonMark emphasis delimiter rewrite and keep CJK-specific fixtures explicit. |
| `UseCitations` | Citations | `Gap` | No citation AST or renderer contract exists. | Decide whether citations are in scope after core CommonMark/GFM closure. |
| `UseCustomContainers` | Custom containers | `Gap` | OfficeIMO has semantic block extension seams, but not Markdig custom container syntax parity. | Route to block parser extensions plus renderer/writer source-slice contracts. |
| `UseDefinitionLists` | Definition lists | `Partial` | OfficeIMO has structured definition-list AST, Markdig-style colon-marker term grouping, multiple-definition parsing, source/native projection, profile-correct HTML comparison coverage, grouped Markdown writer preservation for reparsing, Markdig lazy paragraph, nested block, loose-definition, edge-continuation, and empty-marker first-continuation coverage, loose-definition writer preservation, and blank-separated marker-group writer preservation, but full source-map and writer edge breadth is not closed. | Broaden remaining Markdig definition-list source-map and writer edge cases before promotion. |
| `UseDiagrams` | Diagrams | `Partial` | OfficeIMO has semantic fenced blocks and visual renderer hooks, but not Markdig diagram extension parity. | Compare Mermaid/Nomnom-style cases and decide renderer-package ownership. |
| `UseEmojiAndSmiley` | Emoji and smiley | `Gap` | OfficeIMO has emoji word-join normalization only, not shortcode/smiley expansion. | Keep normalization separate from an optional inline replacement extension. |
| `UseEmphasisExtras` | Emphasis extras | `Covered` | OfficeIMO has strikethrough, inserted-text, highlight/mark, superscript, and subscript inline nodes with Markdig comparison cases, parser-owned source marker metadata, native projection, HTML rendering, Markdown writing, and explicit GFM single-tilde strikethrough profile coverage. | Keep emphasis-extra delimiter cases aligned as broader GFM and lossless trivia coverage expands. |
| `UseFigures` | Figures | `Partial` | OfficeIMO has image/figure import and publisher figure rendering paths, but not Markdig figure syntax parity. | Separate HTML-import figure recovery from Markdown parser extension support. |
| `UseFooters` | Footers | `Gap` | No footer block parser or semantic node exists. | Leave out of scope unless document footer semantics become a Markdown requirement. |
| `UseFootnotes` | Footnotes | `Covered` | OfficeIMO has GFM footnote parsing and GitHub HTML rendering for first-reference ordering, repeated-reference backrefs, missing/unused definitions, nested block bodies, source/native label and marker spans, and structured Markdown writer roundtrip proof. | Keep the GFM footnote fixture corpus and structured-body writer coverage current. |
| `UseGenericAttributes` | Generic attributes | `Partial` | OfficeIMO captures fenced-code brace metadata, but not generic attributes on arbitrary blocks/inlines. | Design attribute storage on semantic and syntax nodes before broad support. |
| `UseGlobalization` | Globalization | `Gap` | No Markdig globalization extension equivalent is documented for OfficeIMO. | Revisit only if a real consumer needs culture-sensitive Markdown behavior. |
| `UseGridTables` | Grid tables | `Gap` | OfficeIMO has pipe tables only; grid table parsing is absent. | Decide if grid tables belong in core or an optional extension package. |
| `UseJiraLinks` | Jira links | `Gap` | No Jira-link shortcut parser exists. | Treat as optional link extension after core link/source mapping is stable. |
| `UseListExtras` | List extras | `Gap` | OfficeIMO list work is focused on CommonMark/GFM task behavior, not Markdig list extras. | Inventory Markdig list-extra syntax before choosing scope. |
| `UseMathematics` | Mathematics | `Partial` | OfficeIMO has math-oriented semantic/rendering paths through host options, but not Markdig math delimiter parity. | Define math parser ownership and compare inline/block math fixtures. |
| `UseMediaLinks` | Media links | `Partial` | OfficeIMO has image/media document semantics, but not Markdig media-link provider parity. | Route shortcut media providers through renderer/host extension seams if in scope. |
| `UseNonAsciiNoEscape` | Non-ASCII no-escape rendering | `Intentional` | OfficeIMO keeps escaping behavior profile/renderer-owned instead of mirroring this Markdig switch. | Document any renderer escaping profile differences when output claims broaden. |
| `UsePipeTables` | Pipe tables | `Covered` | OfficeIMO has GFM pipe-table parsing with delimiter-row validation, escaped/code-span pipe handling, body-row padding/truncation, container ownership, semantic table/cell AST, syntax/native source spans, GitHub HTML rendering, and aligned Markdown writer roundtrip proof. | Keep the GFM table fixture corpus and table-cell source-edit coverage current. |
| `UsePragmaLines` | Pragma lines | `Gap` | No pragma-line parser or semantic contract exists. | Leave out of core unless a concrete document workflow needs it. |
| `UsePreciseSourceLocation` | Precise source location | `Partial` | OfficeIMO has syntax/source/native spans and source slices, but full lossless trivia/original mapping is still partial. | Continue Phase 3 source-map and trivia work before claiming parity. |
| `UseReferralLinks` | Referral links | `Gap` | No Markdig-compatible referral-link renderer policy exists. | Treat as renderer policy work if requested. |
| `UseSelfPipeline` | Self pipeline | `Intentional` | This is a Markdig pipeline composition helper, not a Markdown feature OfficeIMO should mirror directly. | Keep extension composition in OfficeIMO reader/render/write options. |
| `UseSmartyPants` | SmartyPants | `Gap` | No SmartyPants inline transform exists. | Consider as an optional inline transform after delimiter parsing stabilizes. |
| `UseSoftlineBreakAsHardlineBreak` | Soft line break as hard line break | `Covered` | OfficeIMO exposes an explicit reader option that parses ordinary paragraph soft breaks as hard breaks while keeping CommonMark/GFM defaults unchanged, rendering HTML breaks, writing normalized hard-break markdown, and avoiding fake source marker metadata. | Keep the option covered alongside paragraph/list source-map and writer fixtures. |
| `UseTaskLists` | Task lists | `Covered` | OfficeIMO has GFM task-list parsing for checked, unchecked, uppercase, nested, and invalid tight-marker cases; semantic AST flags; exact marker source spans; native snapshots/source edits; GitHub HTML rendering; and Markdown writer roundtrip proof. | Keep the GFM fixture corpus and marker source-edit coverage current. |
| `UseYamlFrontMatter` | YAML front matter | `Covered` | OfficeIMO preserves YAML front matter as a top-of-document raw YAML AST payload with body and fence source spans, structured key/value helpers for simple entries, native source fields and snapshots, HTML omission, and Markdown writer roundtrip behavior. | Keep raw YAML, parsed-entry helpers, and front-matter source-edit fixtures aligned as lossless trivia work expands. |

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
- Add fixtures or engine work by row, not by nearby test names.
