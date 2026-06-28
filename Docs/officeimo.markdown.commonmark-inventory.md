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
| Pinned smoke fixtures | 271 |
| Passing pinned fixtures | 271 |
| Passing unpinned examples | 322 |
| Failing examples | 59 |
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
| HTML blocks | 44 | 19 | 19 | 22 | 3 | 0 |
| Link reference definitions | 27 | 10 | 10 | 17 | 0 | 0 |
| Paragraphs | 8 | 8 | 8 | 0 | 0 | 0 |
| Blank lines | 1 | 0 | 0 | 1 | 0 | 0 |
| Block quotes | 25 | 10 | 10 | 12 | 3 | 0 |
| List items | 48 | 38 | 38 | 9 | 1 | 0 |
| Lists | 26 | 26 | 26 | 0 | 0 | 0 |
| Inlines | 1 | 0 | 0 | 1 | 0 | 0 |
| Code spans | 22 | 5 | 5 | 13 | 4 | 0 |
| Emphasis and strong emphasis | 132 | 11 | 11 | 112 | 9 | 0 |
| Links | 90 | 12 | 12 | 62 | 16 | 0 |
| Images | 22 | 2 | 2 | 20 | 0 | 0 |
| Autolinks | 19 | 8 | 8 | 7 | 4 | 0 |
| Raw HTML | 20 | 6 | 6 | 3 | 11 | 0 |
| Hard line breaks | 15 | 5 | 5 | 6 | 4 | 0 |
| Soft line breaks | 2 | 2 | 2 | 0 | 0 | 0 |
| Textual content | 3 | 0 | 0 | 3 | 0 | 0 |

## Failure Clusters

| Cluster | Failing | Sections | First examples |
| --- | ---: | --- | --- |
| Link/image/reference grammar | 16 | Links | #491, #518, #519, #520, #523, #524, #525, #526, #531, #532, #533, #536 |
| HTML block/raw HTML grammar | 14 | HTML blocks, Raw HTML | #148, #174, #191, #615, #619, #621, #622, #624, #625, #626, #627, #628 |
| Emphasis delimiter algorithm | 9 | Emphasis and strong emphasis | #353, #408, #418, #432, #438, #441, #450, #453, #470 |
| Container indentation and continuation | 6 | Block quotes, Indented code blocks, List items, Tabs | #9, #111, #231, #242, #252, #264 |
| Autolink grammar | 4 | Autolinks | #602, #606, #609, #610 |
| Code span normalization and precedence | 4 | Code spans | #333, #334, #336, #342 |
| Inline precedence and line-break grammar | 4 | Hard line breaks | #641, #642, #643, #644 |
| CommonMark entity decoder | 2 | Entity and numeric character references | #25, #26 |

## Next Use

- Use the failure clusters to pick parser work by root cause, not by nearby example number.
- When a parser slice lands, refresh this report and promote newly passing examples into `commonmark-0.31.2-smoke.json` only when the engine contract is understood.
- Keep intentional deviations at zero unless the compatibility matrix explains the profile difference.
