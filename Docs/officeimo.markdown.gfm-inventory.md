# OfficeIMO.Markdown GFM Inventory

This report is generated from the checked-in cmark-gfm extension smoke fixtures and the current `OfficeIMO.Markdown` GitHub Flavored Markdown profile.
Upstream evidence is pinned to cmark-gfm `0.29.0.gfm.13` at commit `587a12bb54d95ac37241377e6ddc93ea0e45439b`; recorded hashes cover the extension examples, full specification, and pathological-test inventory.

Refresh command:

```powershell
$env:OFFICEIMO_UPDATE_GFM_INVENTORY = '1'
dotnet test OfficeIMO.Markdown.Tests\OfficeIMO.Markdown.Tests.csproj --framework net8.0 --filter "FullyQualifiedName~Markdown_GitHubFlavoredMarkdown_Inventory_Tests"
Remove-Item Env:\OFFICEIMO_UPDATE_GFM_INVENTORY
```

## Summary

| Metric | Count |
| --- | ---: |
| Tracked fixtures | 52 |
| Upstream cmark-gfm fixtures | 48 |
| OfficeIMO supplement fixtures | 4 |
| Passing fixtures | 52 |
| Failing fixtures | 0 |
| Intentional deviations | 0 |
| Pinned upstream source files | 3 |

## Upstream Provenance

| Source | Bytes | SHA-256 | Use |
| --- | ---: | --- | --- |
| `test/extensions.txt` | 21274 | `a2a45e98be9fca95f564f927265a0f63beea6cae5369d1cf4bde44caa51b2a3a` | Extension-family examples |
| `test/spec.txt` | 216680 | `7d8e5814befec287ac116786d81ff14e0adc9b13295b4494649e995408fd871c` | Full GFM specification examples |
| `test/pathological_tests.py` | 5778 | `b200aa0fd6c3199cc0fdaff59c759f862f8f18b5824dc4b33afd8892376aaf69` | Adversarial performance-case inventory |

Checked-in examples stay bounded and auditable; upstream files are identified by immutable commit and hash rather than copied wholesale.

## Section Inventory

| Section | Tracked | Upstream | Supplements | Passing | Failing | Intentional |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Tables | 26 | 23 | 3 | 26 | 0 | 0 |
| Strikethroughs | 3 | 3 | 0 | 3 | 0 | 0 |
| Autolinks | 12 | 11 | 1 | 12 | 0 | 0 |
| HTML tag filter | 2 | 2 | 0 | 2 | 0 | 0 |
| Task lists | 5 | 5 | 0 | 5 | 0 | 0 |
| Footnotes | 3 | 3 | 0 | 3 | 0 | 0 |
| Interop | 1 | 1 | 0 | 1 | 0 | 0 |

## Source Inventory

| Source | Tracked | Passing | Failing |
| --- | ---: | ---: | ---: |
| github/cmark-gfm spec.txt autolinks extension | 7 | 7 | 0 |
| github/cmark-gfm spec.txt tagfilter extension | 1 | 1 | 0 |
| github/cmark-gfm test/extensions.txt | 33 | 33 | 0 |
| github/cmark-gfm test/spec.txt tables extension | 7 | 7 | 0 |
| officeimo/gfm-autolink-smoke | 1 | 1 | 0 |
| officeimo/gfm-container-table-smoke | 2 | 2 | 0 |
| officeimo/gfm-table-smoke | 1 | 1 | 0 |

## Failure Clusters

| Cluster | Failing | Sections | First fixture indexes |
| --- | ---: | --- | --- |

## Next Use

- Use the section inventory to pick GFM expansion work by enabled extension family.
- Keep upstream cmark-gfm fixtures and OfficeIMO supplement fixtures separated when adding new cases.
- When a GFM parser or renderer slice lands, refresh this report and promote new upstream examples only after the behavior contract is understood.
