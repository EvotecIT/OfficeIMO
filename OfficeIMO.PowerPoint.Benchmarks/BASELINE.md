# PowerPoint workflow baseline

This is the first regression-budget input, not a budget. It records one isolated cold Release run refreshed on 2026-08-02 using .NET 8.0.29, Windows 10.0.26200, x64, and a 32-logical-processor machine. Each workflow runs in a new process, captures peak working set before semantic validation, and validates the measured artifact afterward. Open-workflow source fixtures are authored by the parent before the probe starts; reading the fixture from disk remains part of the measured operation. Repeat the run on representative Windows and non-Windows agents before selecting thresholds.

## OfficeIMO workflows

Elapsed time is milliseconds. Allocation and peak working set are MiB.

| Scale | Workflow | Elapsed | Allocated | Peak |
| --- | --- | ---: | ---: | ---: |
| Small, 3 slides | Create/save | 344.9 | 9.7 | 69.3 |
| Small, 3 slides | Open/edit/save | 196.1 | 6.8 | 61.5 |
| Small, 3 slides | Image export | 459.9 | 55.2 | 95.3 |
| Small, 3 slides | PDF export | 709.8 | 226.8 | 240.3 |
| Normal, 30 slides | Create/save | 333.5 | 20.8 | 80.7 |
| Normal, 30 slides | Open/edit/save | 250.0 | 17.9 | 72.7 |
| Normal, 30 slides | Image export | 1717.9 | 495.9 | 155.6 |
| Normal, 30 slides | PDF export | 715.1 | 274.6 | 252.2 |
| Large, 120 slides | Create/save | 522.9 | 63.7 | 115.0 |
| Large, 120 slides | Open/edit/save | 449.2 | 59.0 | 110.5 |
| Large, 120 slides | Image export | 3727.7 | 1975.9 | 160.7 |
| Large, 120 slides | PDF export | 1005.5 | 458.6 | 270.7 |

Create/save and open/edit/save remain comfortably bounded through 120 slides. Image export is the allocation-heavy lane because every raster payload is materialized and validated; it is linear enough to establish a useful baseline, but it should be the first lane to receive an allocation budget after cross-machine variance is known. PDF export has a larger fixed working set from font discovery and embedding, while incremental cost stays controlled.

The baseline run also motivated two immediate fixes: shared path filling no longer uses a per-pixel point-in-path scan, and system font metadata discovery no longer reads full font payloads merely to identify candidates. The current figures are after those corrections.

## Optional ShapeCrawler comparison

ShapeCrawler 0.79.4 was run cold on the same machine and runtime with the same slide dimensions, background/style pattern, editable text, vector panels, tables, two-series clustered bar charts, and every-tenth-slide edit cadence. Both lanes compile and use the same semantic validator after timing and peak-working-set capture to verify expected text, styling, table contents, chart data, and edit markers before running Open XML validation. Shape counts still differ because each library exposes compound table and chart content differently, so compare the complete workflow rather than raw shape totals.

| Scale | Workflow | OfficeIMO elapsed | ShapeCrawler elapsed | OfficeIMO allocated | ShapeCrawler allocated |
| --- | --- | ---: | ---: | ---: | ---: |
| Small | Create/save | 344.9 | 324.4 | 9.7 | 10.4 |
| Small | Open/edit/save | 196.1 | 193.9 | 6.8 | 6.3 |
| Normal | Create/save | 333.5 | 340.5 | 20.8 | 48.7 |
| Normal | Open/edit/save | 250.0 | 201.1 | 17.9 | 11.0 |
| Large | Create/save | 522.9 | 717.2 | 63.7 | 189.4 |
| Large | Open/edit/save | 449.2 | 260.7 | 59.0 | 29.3 |

OfficeIMO creates the normal corpus at parity and the large corpus faster, with substantially fewer managed allocations at both scales. ShapeCrawler is faster and allocates less in the edit lane. That edit gap is visible but not currently pathological: OfficeIMO stays below 0.5 seconds and 60 MiB allocated for the 120-slide workload. Keep it under measurement and investigate changes that materially worsen the curve rather than optimizing against a single workstation result.
