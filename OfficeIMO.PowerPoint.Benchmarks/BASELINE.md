# PowerPoint workflow baseline

This is the first regression-budget input, not a budget. It records one isolated cold Release run refreshed on 2026-08-02 using .NET 8.0.29, Windows 10.0.26200, x64, and a 32-logical-processor machine. Each workflow runs in a new process, captures peak working set before semantic validation, and validates the measured artifact afterward. Open-workflow source fixtures are authored once by the OfficeIMO parent before either probe starts; reading the exact same Small (33,366 bytes), Normal (102,327 bytes), or Large (338,801 bytes) package from disk remains part of both measured operations. Repeat the run on representative Windows and non-Windows agents before selecting thresholds.

## OfficeIMO workflows

Elapsed time is milliseconds. Allocation and peak working set are MiB.

| Scale | Workflow | Elapsed | Allocated | Peak |
| --- | --- | ---: | ---: | ---: |
| Small, 3 slides | Create/save | 291.7 | 9.6 | 70.6 |
| Small, 3 slides | Open/edit/save | 190.7 | 6.8 | 62.6 |
| Small, 3 slides | Image export | 471.5 | 55.2 | 96.7 |
| Small, 3 slides | PDF export | 609.9 | 226.9 | 241.6 |
| Normal, 30 slides | Create/save | 365.5 | 20.8 | 82.0 |
| Normal, 30 slides | Open/edit/save | 284.1 | 17.9 | 74.1 |
| Normal, 30 slides | Image export | 1847.9 | 495.9 | 157.2 |
| Normal, 30 slides | PDF export | 794.8 | 274.7 | 253.9 |
| Large, 120 slides | Create/save | 558.7 | 63.6 | 116.1 |
| Large, 120 slides | Open/edit/save | 451.5 | 59.0 | 111.9 |
| Large, 120 slides | Image export | 3534.2 | 1975.9 | 159.5 |
| Large, 120 slides | PDF export | 1200.6 | 458.7 | 273.7 |

Create/save and open/edit/save remain comfortably bounded through 120 slides. Image export is the allocation-heavy lane because every raster payload is materialized and validated; it is linear enough to establish a useful baseline, but it should be the first lane to receive an allocation budget after cross-machine variance is known. PDF export has a larger fixed working set from font discovery and embedding, while incremental cost stays controlled.

The baseline run also motivated two immediate fixes: shared path filling no longer uses a per-pixel point-in-path scan, and system font metadata discovery no longer reads full font payloads merely to identify candidates. The current figures are after those corrections.

## Optional ShapeCrawler comparison

ShapeCrawler 0.79.4 was run cold on the same machine and runtime. Both open/edit/save lanes consumed the exact same prebuilt PPTX bytes and used the same every-tenth-slide edit cadence. Create/save remains a producer comparison with equivalent slide dimensions, background/style pattern, editable text, vector panels, tables, and two-series clustered bar charts. Both lanes compile and use the same semantic validator after timing and peak-working-set capture to verify expected text, styling, table contents, chart data, and edit markers before running Open XML validation. Shape counts still differ because each library exposes compound table and chart content differently, so compare the complete workflow rather than raw shape totals.

| Scale | Workflow | OfficeIMO elapsed | ShapeCrawler elapsed | OfficeIMO allocated | ShapeCrawler allocated |
| --- | --- | ---: | ---: | ---: | ---: |
| Small | Create/save | 291.7 | 263.1 | 9.6 | 10.4 |
| Small | Open/edit/save | 190.7 | 199.3 | 6.8 | 6.4 |
| Normal | Create/save | 365.5 | 350.2 | 20.8 | 48.7 |
| Normal | Open/edit/save | 284.1 | 226.3 | 17.9 | 11.5 |
| Large | Create/save | 558.7 | 669.2 | 63.6 | 189.3 |
| Large | Open/edit/save | 451.5 | 275.1 | 59.0 | 31.3 |

OfficeIMO creates the normal corpus at parity and the large corpus faster, with substantially fewer managed allocations at both scales. On the exact shared input, OfficeIMO is slightly faster for the Small edit lane; ShapeCrawler is faster and allocates less at Normal and Large. That larger-scale edit gap is visible but not currently pathological: OfficeIMO stays below 0.5 seconds and 60 MiB allocated for the 120-slide workload. Keep it under measurement and investigate changes that materially worsen the curve rather than optimizing against a single workstation result.
