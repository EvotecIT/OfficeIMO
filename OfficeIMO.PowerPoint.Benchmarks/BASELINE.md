# PowerPoint workflow baseline

This is the first regression-budget input, not a budget. It records one isolated Release run on 2026-08-01 using .NET 8.0.29, Windows 10.0.26200, x64, and a 32-logical-processor machine. Repeat the run on representative Windows and non-Windows agents before selecting thresholds.

## OfficeIMO workflows

Elapsed time is milliseconds. Allocation and peak working set are MiB.

| Scale | Workflow | Elapsed | Allocated | Peak |
| --- | --- | ---: | ---: | ---: |
| Small, 3 slides | Create/save | 17.5 | 4.9 | 81.1 |
| Small, 3 slides | Open/edit/save | 11.3 | 4.0 | 85.5 |
| Small, 3 slides | Image export | 187.9 | 49.0 | 118.0 |
| Small, 3 slides | PDF export | 55.1 | 128.8 | 274.9 |
| Normal, 30 slides | Create/save | 64.4 | 16.4 | 109.7 |
| Normal, 30 slides | Open/edit/save | 68.8 | 15.1 | 119.0 |
| Normal, 30 slides | Image export | 704.1 | 490.3 | 161.1 |
| Normal, 30 slides | PDF export | 165.4 | 176.7 | 301.4 |
| Large, 120 slides | Create/save | 279.7 | 59.5 | 125.1 |
| Large, 120 slides | Open/edit/save | 283.9 | 56.0 | 141.3 |
| Large, 120 slides | Image export | 3044.0 | 1972.1 | 177.9 |
| Large, 120 slides | PDF export | 661.7 | 361.2 | 304.6 |

Create/save and open/edit/save remain comfortably bounded through 120 slides. Image export is the allocation-heavy lane because every raster payload is materialized and validated; it is linear enough to establish a useful baseline, but it should be the first lane to receive an allocation budget after cross-machine variance is known. PDF export has a larger fixed working set from font discovery and embedding, while incremental cost stays controlled.

The baseline run also motivated two immediate fixes: shared path filling no longer uses a per-pixel point-in-path scan, and system font metadata discovery no longer reads full font payloads merely to identify candidates. The current figures are after those corrections.

## Optional ShapeCrawler comparison

ShapeCrawler 0.79.4 was run on the same machine and runtime with the same slide dimensions, background/style pattern, editable text, vector panels, tables, two-series clustered bar charts, and every-tenth-slide edit cadence. Both lanes reopen every output and run Open XML validation outside the measured interval. Shape counts still differ because each library exposes compound table and chart content differently, so compare the complete workflow rather than raw shape totals.

| Scale | Workflow | OfficeIMO elapsed | ShapeCrawler elapsed | OfficeIMO allocated | ShapeCrawler allocated |
| --- | --- | ---: | ---: | ---: | ---: |
| Small | Create/save | 17.5 | 18.5 | 4.9 | 6.7 |
| Small | Open/edit/save | 11.3 | 15.3 | 4.0 | 2.9 |
| Normal | Create/save | 64.4 | 95.0 | 16.4 | 45.7 |
| Normal | Open/edit/save | 68.8 | 35.7 | 15.1 | 7.6 |
| Large | Create/save | 279.7 | 430.4 | 59.5 | 188.6 |
| Large | Open/edit/save | 283.9 | 80.9 | 56.0 | 25.8 |

OfficeIMO creates the normal and large corpora faster with substantially fewer managed allocations. ShapeCrawler is faster and allocates less in the edit lane. That edit gap is visible but not currently pathological: OfficeIMO stays below 0.3 seconds and 60 MiB allocated for the 120-slide workload. Keep it under measurement and investigate changes that materially worsen the curve rather than optimizing against a single workstation result.
