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
| Covered | 1 |
| Partial | 14 |
| Intentional | 4 |
| Gap | 14 |

## Extension Families

| Markdig entry point | Family | Status | OfficeIMO state | Next action |
| --- | --- | --- | --- | --- |
| `UseAbbreviations` | Abbreviations | `Gap` | No OfficeIMO abbreviation parser or renderer contract exists yet. | Decide whether abbreviation expansion belongs in core or an optional inline extension. |
| `UseAdvancedExtensions` | Advanced extension bundle | `Intentional` | OfficeIMO should track individual feature families instead of claiming bundle parity. | Keep this row as a roll-up guard; do not implement as a broad on switch. |
| `UseAlertBlocks` | Alert blocks | `Partial` | OfficeIMO has callout blocks and GitHub-style callout parsing, but not Markdig's alert rendering callback shape. | Align callout/alert syntax, AST fields, source spans, and renderer customization explicitly. |
| `UseAutoIdentifiers` | Auto identifiers | `Gap` | Heading ids are not tracked as a Markdig-compatible extension family. | Design slug generation, duplicate handling, and source/native metadata before enabling. |
| `UseAutoLinks` | Extended autolinks | `Partial` | OfficeIMO has profile-sensitive bare URL/email autolinks with GFM coverage, but Markdig option parity is not complete. | Broaden GFM/Markdig autolink cases and document profile differences. |
| `UseBootstrap` | Bootstrap renderer helpers | `Intentional` | This is renderer-theme behavior rather than a core Markdown syntax family for OfficeIMO. | Keep theme/rendering presets separate from parser parity. |
| `UseCjkFriendlyEmphasis` | CJK-friendly emphasis | `Partial` | OfficeIMO has selected CJK-adjacent emphasis regression coverage, but not a Markdig-compatible CJK emphasis option. | Fold into the CommonMark emphasis delimiter rewrite and keep CJK-specific fixtures explicit. |
| `UseCitations` | Citations | `Gap` | No citation AST or renderer contract exists. | Decide whether citations are in scope after core CommonMark/GFM closure. |
| `UseCustomContainers` | Custom containers | `Gap` | OfficeIMO has semantic block extension seams, but not Markdig custom container syntax parity. | Route to block parser extensions plus renderer/writer source-slice contracts. |
| `UseDefinitionLists` | Definition lists | `Partial` | OfficeIMO has structured definition-list AST, syntax, native projection, and HTML coverage, but full Markdig syntax breadth is not inventoried. | Add Markdig spec cases and keep canonical AST cleanup moving. |
| `UseDiagrams` | Diagrams | `Partial` | OfficeIMO has semantic fenced blocks and visual renderer hooks, but not Markdig diagram extension parity. | Compare Mermaid/Nomnom-style cases and decide renderer-package ownership. |
| `UseEmojiAndSmiley` | Emoji and smiley | `Gap` | OfficeIMO has emoji word-join normalization only, not shortcode/smiley expansion. | Keep normalization separate from an optional inline replacement extension. |
| `UseEmphasisExtras` | Emphasis extras | `Partial` | OfficeIMO has strikethrough and highlight/mark-style inlines, but not the full Markdig emphasis-extra set. | Inventory exact delimiter options before changing inline parsing. |
| `UseFigures` | Figures | `Partial` | OfficeIMO has image/figure import and publisher figure rendering paths, but not Markdig figure syntax parity. | Separate HTML-import figure recovery from Markdown parser extension support. |
| `UseFooters` | Footers | `Gap` | No footer block parser or semantic node exists. | Leave out of scope unless document footer semantics become a Markdown requirement. |
| `UseFootnotes` | Footnotes | `Partial` | OfficeIMO has footnote definitions/references, source spans, native metadata, and GFM smoke coverage, but not broad Markdig spec coverage. | Expand footnote corpus and preserve label/body source mapping. |
| `UseGenericAttributes` | Generic attributes | `Partial` | OfficeIMO captures fenced-code brace metadata, but not generic attributes on arbitrary blocks/inlines. | Design attribute storage on semantic and syntax nodes before broad support. |
| `UseGlobalization` | Globalization | `Gap` | No Markdig globalization extension equivalent is documented for OfficeIMO. | Revisit only if a real consumer needs culture-sensitive Markdown behavior. |
| `UseGridTables` | Grid tables | `Gap` | OfficeIMO has pipe tables only; grid table parsing is absent. | Decide if grid tables belong in core or an optional extension package. |
| `UseJiraLinks` | Jira links | `Gap` | No Jira-link shortcut parser exists. | Treat as optional link extension after core link/source mapping is stable. |
| `UseListExtras` | List extras | `Gap` | OfficeIMO list work is focused on CommonMark/GFM task behavior, not Markdig list extras. | Inventory Markdig list-extra syntax before choosing scope. |
| `UseMathematics` | Mathematics | `Partial` | OfficeIMO has math-oriented semantic/rendering paths through host options, but not Markdig math delimiter parity. | Define math parser ownership and compare inline/block math fixtures. |
| `UseMediaLinks` | Media links | `Partial` | OfficeIMO has image/media document semantics, but not Markdig media-link provider parity. | Route shortcut media providers through renderer/host extension seams if in scope. |
| `UseNonAsciiNoEscape` | Non-ASCII no-escape rendering | `Intentional` | OfficeIMO keeps escaping behavior profile/renderer-owned instead of mirroring this Markdig switch. | Document any renderer escaping profile differences when output claims broaden. |
| `UsePipeTables` | Pipe tables | `Partial` | OfficeIMO has GFM pipe-table parsing, AST/source mapping, and tracked GFM fixtures, but broader table corpus coverage is still open. | Expand GFM/Markdig table fixtures for malformed delimiters and containers. |
| `UsePragmaLines` | Pragma lines | `Gap` | No pragma-line parser or semantic contract exists. | Leave out of core unless a concrete document workflow needs it. |
| `UsePreciseSourceLocation` | Precise source location | `Partial` | OfficeIMO has syntax/source/native spans and source slices, but full lossless trivia/original mapping is still partial. | Continue Phase 3 source-map and trivia work before claiming parity. |
| `UseReferralLinks` | Referral links | `Gap` | No Markdig-compatible referral-link renderer policy exists. | Treat as renderer policy work if requested. |
| `UseSelfPipeline` | Self pipeline | `Intentional` | This is a Markdig pipeline composition helper, not a Markdown feature OfficeIMO should mirror directly. | Keep extension composition in OfficeIMO reader/render/write options. |
| `UseSmartyPants` | SmartyPants | `Gap` | No SmartyPants inline transform exists. | Consider as an optional inline transform after delimiter parsing stabilizes. |
| `UseSoftlineBreakAsHardlineBreak` | Soft line break as hard line break | `Gap` | OfficeIMO has hard/soft break nodes, but no Markdig-compatible softbreak-as-hardbreak switch. | Add only as an explicit profile/render option with tests. |
| `UseTaskLists` | Task lists | `Covered` | OfficeIMO has GFM task-list parsing for checked, unchecked, uppercase, nested, and invalid tight-marker cases; semantic AST flags; exact marker source spans; native snapshots/source edits; GitHub HTML rendering; and Markdown writer roundtrip proof. | Keep the GFM fixture corpus and marker source-edit coverage current. |
| `UseYamlFrontMatter` | YAML front matter | `Partial` | OfficeIMO has front matter blocks with key/value and fence source spans, but not a Markdig YAML object-model parity claim. | Separate raw YAML preservation from parsed metadata helpers. |

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
