# OfficeIMO.Markdown CommonMark Inventory

This report is generated from the checked-in official CommonMark `0.31.2` spec JSON and the current `OfficeIMO.Markdown` CommonMark profile.

Refresh command:

```powershell
$env:OFFICEIMO_UPDATE_COMMONMARK_INVENTORY = '1'
dotnet test OfficeIMO.Tests\OfficeIMO.Tests.csproj --framework net8.0 --filter "FullyQualifiedName~Markdown_CommonMark_Inventory_Tests"
Remove-Item Env:\OFFICEIMO_UPDATE_COMMONMARK_INVENTORY
```

## Summary

| Metric | Count |
| --- | ---: |
| Official examples | 652 |
| Pinned smoke fixtures | 296 |
| Passing pinned fixtures | 296 |
| Passing unpinned examples | 336 |
| Failing examples | 20 |
| Intentional deviations | 0 |

## Section Inventory

| Section | Official | Pinned | Passing pinned | Passing unpinned | Failing | Intentional |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Tabs | 11 | 0 | 0 | 10 | 1 | 0 |
| Backslash escapes | 13 | 8 | 8 | 5 | 0 | 0 |
| Entity and numeric character references | 17 | 7 | 7 | 8 | 2 | 0 |
| Precedence | 1 | 0 | 0 | 1 | 0 | 0 |
| Thematic breaks | 19 | 19 | 19 | 0 | 0 | 0 |
| ATX headings | 18 | 18 | 18 | 0 | 0 | 0 |
| Setext headings | 27 | 27 | 27 | 0 | 0 | 0 |
| Indented code blocks | 12 | 1 | 1 | 10 | 1 | 0 |
| Fenced code blocks | 29 | 29 | 29 | 0 | 0 | 0 |
| HTML blocks | 44 | 19 | 19 | 24 | 1 | 0 |
| Link reference definitions | 27 | 10 | 10 | 17 | 0 | 0 |
| Paragraphs | 8 | 8 | 8 | 0 | 0 | 0 |
| Blank lines | 1 | 0 | 0 | 1 | 0 | 0 |
| Block quotes | 25 | 10 | 10 | 12 | 3 | 0 |
| List items | 48 | 38 | 38 | 9 | 1 | 0 |
| Lists | 26 | 26 | 26 | 0 | 0 | 0 |
| Inlines | 1 | 0 | 0 | 1 | 0 | 0 |
| Code spans | 22 | 8 | 8 | 14 | 0 | 0 |
| Emphasis and strong emphasis | 132 | 11 | 11 | 113 | 8 | 0 |
| Links | 90 | 27 | 27 | 63 | 0 | 0 |
| Images | 22 | 2 | 2 | 20 | 0 | 0 |
| Autolinks | 19 | 12 | 12 | 7 | 0 | 0 |
| Raw HTML | 20 | 8 | 8 | 12 | 0 | 0 |
| Hard line breaks | 15 | 6 | 6 | 6 | 3 | 0 |
| Soft line breaks | 2 | 2 | 2 | 0 | 0 | 0 |
| Textual content | 3 | 0 | 0 | 3 | 0 | 0 |

## Failure Clusters

| Cluster | Failing | Sections | First examples |
| --- | ---: | --- | --- |
| Emphasis delimiter algorithm | 8 | Emphasis and strong emphasis | #408, #418, #432, #438, #441, #450, #453, #470 |
| Container indentation and continuation | 6 | Block quotes, Indented code blocks, List items, Tabs | #9, #111, #231, #242, #252, #264 |
| Inline precedence and line-break grammar | 3 | Hard line breaks | #642, #643, #644 |
| CommonMark entity decoder | 2 | Entity and numeric character references | #25, #26 |
| HTML block/raw HTML grammar | 1 | HTML blocks | #174 |

## Next Use

- Use the failure clusters to pick parser work by root cause, not by nearby example number.
- When a parser slice lands, refresh this report and promote newly passing examples into `commonmark-0.31.2-smoke.json` only when the engine contract is understood.
- Keep intentional deviations at zero unless the compatibility matrix explains the profile difference.
