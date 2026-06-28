# OfficeIMO.Markdown GFM Inventory

This report is generated from the checked-in cmark-gfm extension smoke fixtures and the current `OfficeIMO.Markdown` GitHub Flavored Markdown profile.

Refresh command:

```powershell
$env:OFFICEIMO_UPDATE_GFM_INVENTORY = '1'
dotnet test OfficeIMO.Tests\OfficeIMO.Tests.csproj --framework net8.0 --filter "FullyQualifiedName~Markdown_GitHubFlavoredMarkdown_Inventory_Tests"
Remove-Item Env:\OFFICEIMO_UPDATE_GFM_INVENTORY
```

## Summary

| Metric | Count |
| --- | ---: |
| Tracked fixtures | 36 |
| Upstream cmark-gfm fixtures | 33 |
| OfficeIMO supplement fixtures | 3 |
| Passing fixtures | 36 |
| Failing fixtures | 0 |
| Intentional deviations | 0 |

## Section Inventory

| Section | Tracked | Upstream | Supplements | Passing | Failing | Intentional |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Tables | 19 | 16 | 3 | 19 | 0 | 0 |
| Strikethroughs | 3 | 3 | 0 | 3 | 0 | 0 |
| Autolinks | 4 | 4 | 0 | 4 | 0 | 0 |
| HTML tag filter | 1 | 1 | 0 | 1 | 0 | 0 |
| Task lists | 5 | 5 | 0 | 5 | 0 | 0 |
| Footnotes | 3 | 3 | 0 | 3 | 0 | 0 |
| Interop | 1 | 1 | 0 | 1 | 0 | 0 |

## Source Inventory

| Source | Tracked | Passing | Failing |
| --- | ---: | ---: | ---: |
| github/cmark-gfm test/extensions.txt | 33 | 33 | 0 |
| officeimo/gfm-container-table-smoke | 2 | 2 | 0 |
| officeimo/gfm-table-smoke | 1 | 1 | 0 |

## Failure Clusters

| Cluster | Failing | Sections | First fixture indexes |
| --- | ---: | --- | --- |

## Next Use

- Use the section inventory to pick GFM expansion work by enabled extension family.
- Keep upstream cmark-gfm fixtures and OfficeIMO supplement fixtures separated when adding new cases.
- When a GFM parser or renderer slice lands, refresh this report and promote new upstream examples only after the behavior contract is understood.
