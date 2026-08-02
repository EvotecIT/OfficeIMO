# PowerPoint workflow baseline

This is the first regression-budget input, not a budget. It records one isolated cold Release run refreshed on 2026-08-02 using .NET 8.0.29, Windows 10.0.26200, x64, and a 32-logical-processor machine. Each workflow runs in a new process, captures peak working set before semantic validation, and validates the measured artifact afterward. Repeat the run on representative Windows and non-Windows agents before selecting thresholds.

## OfficeIMO workflows

Elapsed time is milliseconds. Allocation and peak working set are MiB.

| Scale | Workflow | Elapsed | Allocated | Peak |
| --- | --- | ---: | ---: | ---: |
| Small, 3 slides | Create/save | 291.4 | 9.7 | 69.4 |
| Small, 3 slides | Open/edit/save | 18.0 | 4.0 | 71.5 |
| Small, 3 slides | Image export | 297.9 | 51.8 | 102.1 |
| Small, 3 slides | PDF export | 453.2 | 223.5 | 263.4 |
| Normal, 30 slides | Create/save | 351.0 | 20.8 | 80.6 |
| Normal, 30 slides | Open/edit/save | 82.9 | 15.1 | 83.1 |
| Normal, 30 slides | Image export | 1615.0 | 492.5 | 161.5 |
| Normal, 30 slides | PDF export | 640.8 | 271.2 | 276.4 |
| Large, 120 slides | Create/save | 531.0 | 63.6 | 114.9 |
| Large, 120 slides | Open/edit/save | 284.8 | 55.9 | 118.6 |
| Large, 120 slides | Image export | 3838.2 | 1972.1 | 165.3 |
| Large, 120 slides | PDF export | 948.3 | 454.9 | 276.8 |

Create/save and open/edit/save remain comfortably bounded through 120 slides. Image export is the allocation-heavy lane because every raster payload is materialized and validated; it is linear enough to establish a useful baseline, but it should be the first lane to receive an allocation budget after cross-machine variance is known. PDF export has a larger fixed working set from font discovery and embedding, while incremental cost stays controlled.

The baseline run also motivated two immediate fixes: shared path filling no longer uses a per-pixel point-in-path scan, and system font metadata discovery no longer reads full font payloads merely to identify candidates. The current figures are after those corrections.

## Optional ShapeCrawler comparison

ShapeCrawler 0.79.4 was run cold on the same machine and runtime with the same slide dimensions, background/style pattern, editable text, vector panels, tables, two-series clustered bar charts, and every-tenth-slide edit cadence. Both lanes compile and use the same semantic validator after timing and peak-working-set capture to verify expected text, styling, table contents, chart data, and edit markers before running Open XML validation. Shape counts still differ because each library exposes compound table and chart content differently, so compare the complete workflow rather than raw shape totals.

| Scale | Workflow | OfficeIMO elapsed | ShapeCrawler elapsed | OfficeIMO allocated | ShapeCrawler allocated |
| --- | --- | ---: | ---: | ---: | ---: |
| Small | Create/save | 291.4 | 283.8 | 9.7 | 10.4 |
| Small | Open/edit/save | 18.0 | 9.6 | 4.0 | 2.9 |
| Normal | Create/save | 351.0 | 355.0 | 20.8 | 48.7 |
| Normal | Open/edit/save | 82.9 | 25.3 | 15.1 | 7.5 |
| Large | Create/save | 531.0 | 689.6 | 63.6 | 189.4 |
| Large | Open/edit/save | 284.8 | 96.3 | 55.9 | 25.6 |

OfficeIMO creates the normal corpus at parity and the large corpus faster, with substantially fewer managed allocations at both scales. ShapeCrawler is faster and allocates less in the edit lane. That edit gap is visible but not currently pathological: OfficeIMO stays below 0.3 seconds and 60 MiB allocated for the 120-slide workload. Keep it under measurement and investigate changes that materially worsen the curve rather than optimizing against a single workstation result.
