# OfficeIMO.Excel Benchmark Report

Generated: 2026-05-31T06:37:06.3174278Z
Run mode: quick
Publish: False
Machine: EVOMAGIC (32 processors)

## How to Read
- Mean: average elapsed time for the measured operation. Lower is better.
- Allocated: managed memory allocated by the measured operation. Lower is better.
- Package: generated XLSX file size when package profiling is available.
- Ratio to OfficeIMO compares each library against OfficeIMO for the same scenario and row count.
- Quick runs are useful for engineering direction; full runs should be used for public claims.
- Benchmarks are machine-specific and should be treated as reproducible evidence, not universal guarantees.

## At a Glance

| Rows | Artifact | Workload | Category | OfficeIMO wins | OfficeIMO losses | Biggest loss |
| ---: | --- | --- | --- | ---: | ---: | --- |
| 2500 | dense-helloworld-comparison | read | Other | 0 | 2 | dense-helloworld-read-range vs Sylvan.Data.Excel (1.20x) |
| 2500 | package-profile | package | Package size | 41 | 13 | write-insertobjects-legacy-dictionaries-direct vs LargeXlsx (1.53x) |
| 2500 | speed-comparison | mutate | AutoFit and mutation | 2 | 0 |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | 4 | 0 |  |
| 2500 | speed-comparison | other | Object projection | 2 | 0 |  |
| 2500 | speed-comparison | other | Range and table read | 4 | 0 |  |
| 2500 | speed-comparison | other | Real-world feature mix | 6 | 0 |  |
| 2500 | speed-comparison | other | Real-world report | 2 | 0 |  |
| 2500 | speed-comparison | other | Report workbook | 4 | 0 |  |
| 2500 | speed-comparison | read | Formula write/read | 1 | 0 |  |
| 2500 | speed-comparison | read | Other | 1 | 2 | shared-string-read vs Sylvan.Data.Excel (1.61x) |
| 2500 | speed-comparison | read | Range and table read | 3 | 4 | read-datatable vs Sylvan.Data.Excel (1.96x) |
| 2500 | speed-comparison | read | Streaming read | 1 | 3 | read-range-stream vs Sylvan.Data.Excel (2.43x) |
| 2500 | speed-comparison | read | Typed object read | 0 | 2 | read-objects vs Sylvan.Data.Excel (1.15x) |
| 2500 | speed-comparison | write | AutoFit and mutation | 5 | 0 |  |
| 2500 | speed-comparison | write | Cell writer | 7 | 0 |  |
| 2500 | speed-comparison | write | DataSet table export | 3 | 0 |  |
| 2500 | speed-comparison | write | DataTable table export | 4 | 0 |  |
| 2500 | speed-comparison | write | Formatted report write | 1 | 0 |  |
| 2500 | speed-comparison | write | Formula write/read | 1 | 0 |  |
| 2500 | speed-comparison | write | Other | 2 | 2 | write-powershell-mixed-objects-direct vs LargeXlsx (1.30x) |
| 2500 | speed-comparison | write | Plain cell export | 1 | 3 | append-plain-rows vs LargeXlsx (1.43x) |
| 2500 | speed-comparison | write | Plain streaming export | 1 | 1 | write-datareader-plain vs Sylvan.Data.Excel (1.14x) |
| 2500 | speed-comparison | write | Plain string export | 0 | 1 | write-blog-2023-20-string-columns vs LargeXlsx (1.07x) |
| 2500 | speed-comparison | write | Shared string write | 4 | 0 |  |
| 2500 | speed-comparison | write | Typed object export | 0 | 3 | write-insertobjects-direct vs LargeXlsx (1.38x) |
| 25000 | dense-helloworld-comparison | read | Other | 0 | 2 | dense-helloworld-read-stream vs Sylvan.Data.Excel (1.29x) |
| 25000 | package-profile | package | Package size | 42 | 12 | write-insertobjects-legacy-dictionaries-direct vs LargeXlsx (1.52x) |
| 25000 | speed-comparison | mutate | AutoFit and mutation | 2 | 0 |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | 4 | 0 |  |
| 25000 | speed-comparison | other | Object projection | 2 | 0 |  |
| 25000 | speed-comparison | other | Range and table read | 4 | 0 |  |
| 25000 | speed-comparison | other | Real-world feature mix | 6 | 0 |  |
| 25000 | speed-comparison | other | Real-world report | 2 | 0 |  |
| 25000 | speed-comparison | other | Report workbook | 4 | 0 |  |
| 25000 | speed-comparison | read | Formula write/read | 1 | 0 |  |
| 25000 | speed-comparison | read | Other | 2 | 1 | shared-string-read vs Sylvan.Data.Excel (1.14x) |
| 25000 | speed-comparison | read | Range and table read | 2 | 5 | read-top-range vs Sylvan.Data.Excel (1.32x) |
| 25000 | speed-comparison | read | Streaming read | 1 | 3 | read-top-range-stream-small-chunks vs Sylvan.Data.Excel (1.33x) |
| 25000 | speed-comparison | read | Typed object read | 1 | 1 | read-objects vs Sylvan.Data.Excel (1.29x) |
| 25000 | speed-comparison | write | AutoFit and mutation | 5 | 0 |  |
| 25000 | speed-comparison | write | Cell writer | 7 | 0 |  |
| 25000 | speed-comparison | write | DataSet table export | 3 | 0 |  |
| 25000 | speed-comparison | write | DataTable table export | 3 | 1 | write-datatable-direct vs LargeXlsx (1.03x) |
| 25000 | speed-comparison | write | Formatted report write | 1 | 0 |  |
| 25000 | speed-comparison | write | Formula write/read | 1 | 0 |  |
| 25000 | speed-comparison | write | Other | 2 | 2 | write-powershell-psobject-mixed-direct vs LargeXlsx (1.23x) |
| 25000 | speed-comparison | write | Plain cell export | 1 | 3 | append-plain-rows vs LargeXlsx (1.29x) |
| 25000 | speed-comparison | write | Plain streaming export | 1 | 1 | write-datareader-plain vs Sylvan.Data.Excel (1.24x) |
| 25000 | speed-comparison | write | Plain string export | 0 | 1 | write-blog-2023-20-string-columns vs LargeXlsx (1.12x) |
| 25000 | speed-comparison | write | Shared string write | 4 | 0 |  |
| 25000 | speed-comparison | write | Typed object export | 0 | 3 | write-insertobjects-flat-dictionaries-direct vs LargeXlsx (1.43x) |

## Rows

| Rows | Artifact | Scenario | Library | Mean | Allocated | Package | Best | Outcome |
| ---: | --- | --- | --- | ---: | ---: | ---: | --- | --- |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | Sylvan.Data.Excel | 6.74 ms | 362.3 KB |  | Sylvan.Data.Excel | 17.0% faster than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | OfficeIMO.Excel | 8.12 ms | 2.4 MB |  | Sylvan.Data.Excel | Loss +20.4% |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | ExcelDataReader | 16.65 ms | 6.7 MB |  | Sylvan.Data.Excel | 105.1% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | MiniExcel | 20.71 ms | 21.0 MB |  | Sylvan.Data.Excel | 155.1% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | Sylvan.Data.Excel | 6.24 ms | 362.3 KB |  | Sylvan.Data.Excel | 15.4% faster than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | OfficeIMO.Excel | 7.38 ms | 2.4 MB |  | Sylvan.Data.Excel | Loss +18.2% |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | ExcelDataReader | 15.69 ms | 6.7 MB |  | Sylvan.Data.Excel | 112.7% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | MiniExcel | 20.40 ms | 21.0 MB |  | Sylvan.Data.Excel | 176.6% slower than OfficeIMO |
| 2500 | package-profile | append-plain-rows | LargeXlsx | 1.79 ms | 296.4 KB | 63.1 KB | LargeXlsx | 20.2% faster than OfficeIMO |
| 2500 | package-profile | append-plain-rows | OfficeIMO.Excel | 2.25 ms | 1.5 MB | 63.0 KB | LargeXlsx | Loss +25.3% |
| 2500 | package-profile | append-plain-rows | MiniExcel | 4.52 ms | 19.2 MB | 68.1 KB | LargeXlsx | 101.5% slower than OfficeIMO |
| 2500 | package-profile | append-plain-rows | ClosedXML | 17.74 ms | 10.9 MB | 59.8 KB | LargeXlsx | 690.0% slower than OfficeIMO |
| 2500 | package-profile | append-plain-rows | EPPlus | 29.62 ms | 14.0 MB | 56.9 KB | LargeXlsx | 1219.2% slower than OfficeIMO |
| 2500 | package-profile | autofit-existing | OfficeIMO.Excel | 8.84 ms | 1.9 MB | 143.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | autofit-existing | EPPlus | 88.48 ms | 49.5 MB | 115.0 KB | OfficeIMO.Excel | 900.8% slower than OfficeIMO |
| 2500 | package-profile | autofit-existing | ClosedXML | 154.22 ms | 82.6 MB | 121.0 KB | OfficeIMO.Excel | 1644.3% slower than OfficeIMO |
| 2500 | package-profile | large-shared-strings | OfficeIMO.Excel | 2.36 ms | 2.4 MB | 55.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | large-shared-strings | MiniExcel | 4.81 ms | 20.6 MB | 60.7 KB | OfficeIMO.Excel | 104.0% slower than OfficeIMO |
| 2500 | package-profile | large-shared-strings | ClosedXML | 13.22 ms | 11.0 MB | 50.3 KB | OfficeIMO.Excel | 460.8% slower than OfficeIMO |
| 2500 | package-profile | large-shared-strings | EPPlus | 24.93 ms | 12.5 MB | 48.1 KB | OfficeIMO.Excel | 957.8% slower than OfficeIMO |
| 2500 | package-profile | realworld-autofilter | OfficeIMO.Excel | 4.61 ms | 1.3 MB | 142.4 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-autofilter | ClosedXML | 38.08 ms | 21.7 MB | 120.2 KB | OfficeIMO.Excel | 726.1% slower than OfficeIMO |
| 2500 | package-profile | realworld-autofilter | EPPlus | 48.85 ms | 24.1 MB | 114.2 KB | OfficeIMO.Excel | 959.7% slower than OfficeIMO |
| 2500 | package-profile | realworld-charts | OfficeIMO.Excel | 6.37 ms | 1.8 MB | 147.6 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-charts | EPPlus | 50.42 ms | 26.5 MB | 117.0 KB | OfficeIMO.Excel | 692.0% slower than OfficeIMO |
| 2500 | package-profile | realworld-conditional-formatting | OfficeIMO.Excel | 4.19 ms | 1.4 MB | 142.7 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-conditional-formatting | ClosedXML | 34.23 ms | 21.8 MB | 120.3 KB | OfficeIMO.Excel | 716.2% slower than OfficeIMO |
| 2500 | package-profile | realworld-conditional-formatting | EPPlus | 44.59 ms | 24.2 MB | 114.3 KB | OfficeIMO.Excel | 963.3% slower than OfficeIMO |
| 2500 | package-profile | realworld-data-validation | OfficeIMO.Excel | 4.94 ms | 1.3 MB | 142.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-data-validation | ClosedXML | 37.64 ms | 21.7 MB | 120.3 KB | OfficeIMO.Excel | 661.7% slower than OfficeIMO |
| 2500 | package-profile | realworld-data-validation | EPPlus | 49.86 ms | 24.1 MB | 114.2 KB | OfficeIMO.Excel | 909.0% slower than OfficeIMO |
| 2500 | package-profile | realworld-freeze-panes | OfficeIMO.Excel | 4.03 ms | 1.3 MB | 142.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-freeze-panes | ClosedXML | 38.72 ms | 21.7 MB | 120.2 KB | OfficeIMO.Excel | 860.4% slower than OfficeIMO |
| 2500 | package-profile | realworld-freeze-panes | EPPlus | 49.75 ms | 24.2 MB | 114.3 KB | OfficeIMO.Excel | 1133.9% slower than OfficeIMO |
| 2500 | package-profile | realworld-pivot-table | OfficeIMO.Excel | 10.88 ms | 5.4 MB | 200.4 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-pivot-table | EPPlus | 53.85 ms | 28.8 MB | 117.4 KB | OfficeIMO.Excel | 394.8% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-all-in-one | OfficeIMO.Excel | 14.06 ms | 6.1 MB | 206.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-all-in-one | EPPlus | 81.69 ms | 53.3 MB | 121.8 KB | OfficeIMO.Excel | 481.1% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-chart-first | OfficeIMO.Excel | 13.49 ms | 6.1 MB | 206.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-chart-first | EPPlus | 84.74 ms | 53.3 MB | 121.8 KB | OfficeIMO.Excel | 528.0% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-core | OfficeIMO.Excel | 4.93 ms | 1.5 MB | 143.9 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-core | EPPlus | 77.21 ms | 46.2 MB | 115.6 KB | OfficeIMO.Excel | 1465.4% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-core | ClosedXML | 96.60 ms | 68.2 MB | 121.5 KB | OfficeIMO.Excel | 1858.6% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-extra-column | OfficeIMO.Excel | 14.28 ms | 6.2 MB | 219.1 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-extra-column | EPPlus | 92.46 ms | 57.8 MB | 128.4 KB | OfficeIMO.Excel | 547.5% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-no-autofit | OfficeIMO.Excel | 11.95 ms | 6.0 MB | 206.4 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-no-autofit | EPPlus | 53.86 ms | 32.1 MB | 121.8 KB | OfficeIMO.Excel | 350.9% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-post-mutation | OfficeIMO.Excel | 12.69 ms | 6.1 MB | 206.6 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-post-mutation | EPPlus | 90.40 ms | 53.3 MB | 121.9 KB | OfficeIMO.Excel | 612.4% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-shuffled-columns | OfficeIMO.Excel | 15.33 ms | 6.1 MB | 211.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-shuffled-columns | EPPlus | 84.62 ms | 53.3 MB | 124.3 KB | OfficeIMO.Excel | 451.8% slower than OfficeIMO |
| 2500 | package-profile | report-workbook | OfficeIMO.Excel | 17.65 ms | 7.1 MB | 275.6 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | report-workbook | EPPlus | 108.51 ms | 75.7 MB | 161.8 KB | OfficeIMO.Excel | 514.8% slower than OfficeIMO |
| 2500 | package-profile | report-workbook-core | OfficeIMO.Excel | 7.96 ms | 2.6 MB | 187.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | report-workbook-core | ClosedXML | 119.39 ms | 94.9 MB | 165.1 KB | OfficeIMO.Excel | 1400.3% slower than OfficeIMO |
| 2500 | package-profile | report-workbook-core | EPPlus | 119.81 ms | 70.3 MB | 157.2 KB | OfficeIMO.Excel | 1405.6% slower than OfficeIMO |
| 2500 | package-profile | report-workbook-datatable | OfficeIMO.Excel | 17.85 ms | 7.4 MB | 275.6 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | report-workbook-datatable | EPPlus | 122.78 ms | 64.4 MB | 161.8 KB | OfficeIMO.Excel | 587.6% slower than OfficeIMO |
| 2500 | package-profile | report-workbook-datatable-core | OfficeIMO.Excel | 7.45 ms | 2.9 MB | 187.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | report-workbook-datatable-core | EPPlus | 109.00 ms | 59.1 MB | 157.2 KB | OfficeIMO.Excel | 1363.6% slower than OfficeIMO |
| 2500 | package-profile | report-workbook-datatable-core | ClosedXML | 121.77 ms | 80.9 MB | 165.1 KB | OfficeIMO.Excel | 1535.1% slower than OfficeIMO |
| 2500 | package-profile | write-blog-2023-20-string-columns | LargeXlsx | 5.07 ms | 857.6 KB | 237.7 KB | LargeXlsx | 5.5% faster than OfficeIMO |
| 2500 | package-profile | write-blog-2023-20-string-columns | OfficeIMO.Excel | 5.36 ms | 1.6 MB | 216.7 KB | LargeXlsx | Loss +5.8% |
| 2500 | package-profile | write-blog-2023-20-string-columns | MiniExcel | 22.32 ms | 35.1 MB | 235.3 KB | LargeXlsx | 316.1% slower than OfficeIMO |
| 2500 | package-profile | write-blog-2023-20-string-columns | ClosedXML | 104.61 ms | 69.8 MB | 257.2 KB | LargeXlsx | 1849.9% slower than OfficeIMO |
| 2500 | package-profile | write-bulk-report | OfficeIMO.Excel | 4.33 ms | 1.4 MB | 143.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-bulk-report | MiniExcel | 9.96 ms | 26.2 MB | 153.8 KB | OfficeIMO.Excel | 129.8% slower than OfficeIMO |
| 2500 | package-profile | write-bulk-report | EPPlus | 79.07 ms | 46.1 MB | 115.0 KB | OfficeIMO.Excel | 1724.6% slower than OfficeIMO |
| 2500 | package-profile | write-bulk-report | ClosedXML | 91.62 ms | 57.0 MB | 121.0 KB | OfficeIMO.Excel | 2014.0% slower than OfficeIMO |
| 2500 | package-profile | write-cellformula | OfficeIMO.Excel | 3.39 ms | 1.4 MB | 66.6 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellformula | ClosedXML | 24.79 ms | 11.8 MB | 70.6 KB | OfficeIMO.Excel | 631.0% slower than OfficeIMO |
| 2500 | package-profile | write-cellformula | EPPlus | 44.04 ms | 17.7 MB | 62.1 KB | OfficeIMO.Excel | 1198.7% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-empty-strings | OfficeIMO.Excel | 2.37 ms | 1.7 MB | 44.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-empty-strings | ClosedXML | 13.31 ms | 9.7 MB | 44.9 KB | OfficeIMO.Excel | 461.8% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-empty-strings | EPPlus | 24.48 ms | 11.5 MB | 42.0 KB | OfficeIMO.Excel | 933.1% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-numbers | OfficeIMO.Excel | 2.49 ms | 1.1 MB | 47.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-numbers | ClosedXML | 12.57 ms | 9.0 MB | 45.9 KB | OfficeIMO.Excel | 405.7% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-numbers | EPPlus | 25.76 ms | 12.6 MB | 43.7 KB | OfficeIMO.Excel | 936.5% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-mixed | OfficeIMO.Excel | 3.22 ms | 1.7 MB | 61.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-object-mixed | ClosedXML | 20.79 ms | 11.6 MB | 59.5 KB | OfficeIMO.Excel | 546.3% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-mixed | EPPlus | 32.49 ms | 15.3 MB | 58.9 KB | OfficeIMO.Excel | 910.2% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-sparse | OfficeIMO.Excel | 3.17 ms | 1.5 MB | 62.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-object-sparse | ClosedXML | 17.85 ms | 11.0 MB | 52.5 KB | OfficeIMO.Excel | 462.8% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-sparse | EPPlus | 31.58 ms | 14.6 MB | 54.2 KB | OfficeIMO.Excel | 895.7% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 3.11 ms | 1.5 MB | 62.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-object-sparse-batch | ClosedXML | 24.55 ms | 11.0 MB | 52.5 KB | OfficeIMO.Excel | 688.9% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-sparse-batch | EPPlus | 32.53 ms | 14.6 MB | 54.2 KB | OfficeIMO.Excel | 945.5% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-scalars | OfficeIMO.Excel | 2.13 ms | 1.1 MB | 46.3 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-scalars | ClosedXML | 14.10 ms | 8.8 MB | 45.4 KB | OfficeIMO.Excel | 563.0% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-scalars | EPPlus | 27.22 ms | 12.5 MB | 42.4 KB | OfficeIMO.Excel | 1179.7% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings | OfficeIMO.Excel | 3.87 ms | 2.6 MB | 55.1 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-strings | ClosedXML | 14.43 ms | 11.0 MB | 50.3 KB | OfficeIMO.Excel | 272.7% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings | EPPlus | 25.28 ms | 12.5 MB | 48.1 KB | OfficeIMO.Excel | 553.1% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings-distinct | OfficeIMO.Excel | 2.43 ms | 2.3 MB | 51.8 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-strings-distinct | ClosedXML | 16.10 ms | 12.8 MB | 61.9 KB | OfficeIMO.Excel | 562.6% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings-distinct | EPPlus | 28.89 ms | 13.6 MB | 61.5 KB | OfficeIMO.Excel | 1089.0% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings-repeated | OfficeIMO.Excel | 2.14 ms | 1.5 MB | 40.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-strings-repeated | ClosedXML | 11.76 ms | 9.0 MB | 38.8 KB | OfficeIMO.Excel | 450.4% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings-repeated | EPPlus | 23.25 ms | 11.1 MB | 34.8 KB | OfficeIMO.Excel | 988.3% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-temporal | OfficeIMO.Excel | 2.80 ms | 1.4 MB | 63.3 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-temporal | ClosedXML | 18.76 ms | 9.5 MB | 54.5 KB | OfficeIMO.Excel | 571.0% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-temporal | EPPlus | 30.46 ms | 14.4 MB | 53.1 KB | OfficeIMO.Excel | 989.6% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 1.39 ms | 447.0 KB | 47.3 KB | LargeXlsx | 18.9% faster than OfficeIMO |
| 2500 | package-profile | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 1.71 ms | 1.1 MB | 48.2 KB | LargeXlsx | Loss +23.2% |
| 2500 | package-profile | write-cellvalues-headerless-rectangle-direct | ClosedXML | 13.32 ms | 10.0 MB | 53.0 KB | LargeXlsx | 678.7% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-headerless-rectangle-direct | EPPlus | 25.99 ms | 12.7 MB | 52.5 KB | LargeXlsx | 1420.0% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | LargeXlsx | 3.87 ms | 758.3 KB | 138.4 KB | LargeXlsx | 16.5% faster than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 4.64 ms | 2.0 MB | 138.0 KB | LargeXlsx | Loss +19.8% |
| 2500 | package-profile | write-cellvalues-rectangle-direct | MiniExcel | 9.47 ms | 22.7 MB | 153.7 KB | LargeXlsx | 104.0% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | ClosedXML | 40.20 ms | 21.7 MB | 120.1 KB | LargeXlsx | 766.4% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | EPPlus | 45.62 ms | 24.1 MB | 114.1 KB | LargeXlsx | 883.2% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | Sylvan.Data.Excel | 3.69 ms | 758.7 KB | 78.5 KB | Sylvan.Data.Excel | 17.3% faster than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | OfficeIMO.Excel | 4.46 ms | 1.7 MB | 138.0 KB | Sylvan.Data.Excel | Loss +21.0% |
| 2500 | package-profile | write-datareader-plain | LargeXlsx | 4.51 ms | 1.0 MB | 138.4 KB | Sylvan.Data.Excel | Tie vs OfficeIMO |
| 2500 | package-profile | write-datareader-plain | MiniExcel | 8.62 ms | 22.5 MB | 153.6 KB | Sylvan.Data.Excel | 93.2% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | ClosedXML | 31.74 ms | 11.3 MB | 120.1 KB | Sylvan.Data.Excel | 611.5% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | EPPlus | 43.07 ms | 16.3 MB | 114.9 KB | Sylvan.Data.Excel | 865.4% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table | OfficeIMO.Excel | 4.83 ms | 1.7 MB | 138.8 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-datareader-table | MiniExcel | 8.15 ms | 22.5 MB | 153.6 KB | OfficeIMO.Excel | 68.9% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table | ClosedXML | 42.68 ms | 18.6 MB | 120.9 KB | OfficeIMO.Excel | 784.2% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table | EPPlus | 42.98 ms | 16.3 MB | 114.9 KB | OfficeIMO.Excel | 790.4% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table-autofit | OfficeIMO.Excel | 5.27 ms | 1.7 MB | 138.9 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-datareader-table-autofit | MiniExcel | 10.06 ms | 26.0 MB | 153.8 KB | OfficeIMO.Excel | 90.9% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table-autofit | EPPlus | 66.73 ms | 37.4 MB | 115.1 KB | OfficeIMO.Excel | 1166.8% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table-autofit | ClosedXML | 87.41 ms | 57.0 MB | 121.0 KB | OfficeIMO.Excel | 1559.2% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-direct-export | OfficeIMO.Excel | 4.53 ms | 2.1 MB | 131.1 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-dataset-sparse-direct-export | LargeXlsx | 5.15 ms | 1.1 MB | 164.2 KB | OfficeIMO.Excel | 13.7% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-direct-export | MiniExcel | 11.67 ms | 29.0 MB | 180.5 KB | OfficeIMO.Excel | 157.5% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-direct-export | ClosedXML | 67.27 ms | 26.8 MB | 159.4 KB | OfficeIMO.Excel | 1384.5% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-direct-export | EPPlus | 67.79 ms | 21.4 MB | 144.5 KB | OfficeIMO.Excel | 1395.9% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-tables | OfficeIMO.Excel | 5.66 ms | 2.8 MB | 176.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-dataset-sparse-tables | MiniExcel | 11.50 ms | 29.0 MB | 180.5 KB | OfficeIMO.Excel | 103.2% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-tables | ClosedXML | 67.63 ms | 26.8 MB | 159.4 KB | OfficeIMO.Excel | 1094.5% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-tables | EPPlus | 71.10 ms | 21.4 MB | 144.5 KB | OfficeIMO.Excel | 1155.8% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables | OfficeIMO.Excel | 5.01 ms | 2.0 MB | 138.9 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-dataset-tables | MiniExcel | 9.63 ms | 28.0 MB | 156.4 KB | OfficeIMO.Excel | 92.3% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables | EPPlus | 41.45 ms | 18.3 MB | 116.6 KB | OfficeIMO.Excel | 727.8% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables | ClosedXML | 43.41 ms | 18.4 MB | 123.4 KB | OfficeIMO.Excel | 766.7% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables-autofit | OfficeIMO.Excel | 5.08 ms | 2.0 MB | 139.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-dataset-tables-autofit | MiniExcel | 10.95 ms | 31.1 MB | 156.6 KB | OfficeIMO.Excel | 115.2% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables-autofit | EPPlus | 74.32 ms | 40.5 MB | 116.9 KB | OfficeIMO.Excel | 1361.5% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables-autofit | ClosedXML | 95.69 ms | 55.4 MB | 123.7 KB | OfficeIMO.Excel | 1781.8% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | LargeXlsx | 3.99 ms | 1.1 MB | 138.4 KB | LargeXlsx | 18.3% faster than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | OfficeIMO.Excel | 4.89 ms | 1.7 MB | 138.0 KB | LargeXlsx | Loss +22.4% |
| 2500 | package-profile | write-datatable-direct | MiniExcel | 9.58 ms | 22.5 MB | 153.7 KB | LargeXlsx | 96.0% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | ClosedXML | 43.65 ms | 11.3 MB | 120.1 KB | LargeXlsx | 792.8% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | EPPlus | 44.21 ms | 16.3 MB | 114.9 KB | LargeXlsx | 804.1% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-table-direct | OfficeIMO.Excel | 4.74 ms | 1.7 MB | 138.8 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-datatable-table-direct | MiniExcel | 9.45 ms | 22.5 MB | 153.7 KB | OfficeIMO.Excel | 99.5% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-table-direct | ClosedXML | 40.97 ms | 18.6 MB | 120.9 KB | OfficeIMO.Excel | 764.9% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-table-direct | EPPlus | 42.90 ms | 16.3 MB | 114.9 KB | OfficeIMO.Excel | 805.5% slower than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | LargeXlsx | 3.34 ms | 758.3 KB | 138.4 KB | LargeXlsx | 26.7% faster than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 4.56 ms | 1.7 MB | 138.0 KB | LargeXlsx | Loss +36.4% |
| 2500 | package-profile | write-fluent-rowsfrom-direct | MiniExcel | 10.34 ms | 22.7 MB | 153.7 KB | LargeXlsx | 126.7% slower than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | ClosedXML | 32.63 ms | 11.3 MB | 120.1 KB | LargeXlsx | 615.4% slower than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | EPPlus | 43.58 ms | 16.3 MB | 114.9 KB | LargeXlsx | 855.4% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 5.36 ms | 1.3 MB | 142.4 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 70.55 ms | 37.4 MB | 115.1 KB | OfficeIMO.Excel | 1215.1% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 91.83 ms | 49.7 MB | 120.2 KB | OfficeIMO.Excel | 1611.7% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | LargeXlsx | 4.02 ms | 758.3 KB | 138.4 KB | LargeXlsx | 14.2% faster than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | OfficeIMO.Excel | 4.68 ms | 1.3 MB | 142.3 KB | LargeXlsx | Loss +16.6% |
| 2500 | package-profile | write-insertobjects-direct | MiniExcel | 10.79 ms | 22.7 MB | 153.7 KB | LargeXlsx | 130.4% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | ClosedXML | 40.84 ms | 11.3 MB | 120.1 KB | LargeXlsx | 771.8% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | EPPlus | 48.52 ms | 16.3 MB | 114.9 KB | LargeXlsx | 935.9% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 6.49 ms | 1.5 MB | 138.1 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 67.74 ms | 37.4 MB | 115.1 KB | OfficeIMO.Excel | 943.6% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 83.81 ms | 49.7 MB | 120.2 KB | OfficeIMO.Excel | 1191.1% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 3.94 ms | 758.3 KB | 138.4 KB | LargeXlsx | 31.0% faster than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 5.71 ms | 1.5 MB | 138.0 KB | LargeXlsx | Loss +45.0% |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | MiniExcel | 9.88 ms | 22.7 MB | 153.7 KB | LargeXlsx | 73.0% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | ClosedXML | 35.22 ms | 11.3 MB | 120.1 KB | LargeXlsx | 516.6% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | EPPlus | 44.96 ms | 16.3 MB | 114.9 KB | LargeXlsx | 687.2% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | LargeXlsx | 3.95 ms | 758.3 KB | 138.4 KB | LargeXlsx | 34.8% faster than OfficeIMO |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | OfficeIMO.Excel | 6.06 ms | 1.7 MB | 142.3 KB | LargeXlsx | Loss +53.3% |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | MiniExcel | 10.60 ms | 22.7 MB | 153.7 KB | LargeXlsx | 74.9% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | ClosedXML | 36.87 ms | 11.3 MB | 120.1 KB | LargeXlsx | 508.5% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | EPPlus | 43.50 ms | 16.3 MB | 114.9 KB | LargeXlsx | 617.8% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.93 ms | 1.3 MB | 142.4 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 58.21 ms | 27.9 MB | 120.2 KB | OfficeIMO.Excel | 1080.2% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 61.51 ms | 26.7 MB | 115.0 KB | OfficeIMO.Excel | 1147.1% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-mixed-objects-direct | LargeXlsx | 5.31 ms | 802.5 KB | 182.6 KB | LargeXlsx | 31.4% faster than OfficeIMO |
| 2500 | package-profile | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 7.74 ms | 2.3 MB | 183.1 KB | LargeXlsx | Loss +45.8% |
| 2500 | package-profile | write-powershell-mixed-objects-direct | MiniExcel | 10.69 ms | 24.6 MB | 194.0 KB | LargeXlsx | 38.1% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-mixed-objects-direct | ClosedXML | 45.32 ms | 16.6 MB | 161.0 KB | LargeXlsx | 485.5% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-mixed-objects-direct | EPPlus | 59.97 ms | 19.6 MB | 152.1 KB | LargeXlsx | 674.8% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-mixed-direct | LargeXlsx | 5.00 ms | 802.5 KB | 182.6 KB | LargeXlsx | 5.0% faster than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 5.26 ms | 1.5 MB | 182.4 KB | LargeXlsx | Loss +5.2% |
| 2500 | package-profile | write-powershell-psobject-mixed-direct | MiniExcel | 10.38 ms | 24.6 MB | 194.0 KB | LargeXlsx | 97.2% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-mixed-direct | ClosedXML | 45.53 ms | 16.6 MB | 161.0 KB | LargeXlsx | 764.8% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-mixed-direct | EPPlus | 56.18 ms | 19.6 MB | 152.1 KB | LargeXlsx | 967.1% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-wide-direct | LargeXlsx | 23.92 ms | 2.7 MB | 644.6 KB | LargeXlsx | 7.0% faster than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 25.72 ms | 4.4 MB | 651.0 KB | LargeXlsx | Loss +7.5% |
| 2500 | package-profile | write-powershell-psobject-wide-direct | MiniExcel | 41.68 ms | 47.3 MB | 674.4 KB | LargeXlsx | 62.0% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-wide-direct | ClosedXML | 141.20 ms | 50.4 MB | 615.5 KB | LargeXlsx | 448.9% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-wide-direct | EPPlus | 182.54 ms | 67.5 MB | 548.9 KB | LargeXlsx | 609.6% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | LargeXlsx | 1.95 ms | 296.4 KB |  | LargeXlsx | 30.3% faster than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | OfficeIMO.Excel | 2.79 ms | 1.5 MB |  | LargeXlsx | Loss +43.5% |
| 2500 | speed-comparison | append-plain-rows | MiniExcel | 5.35 ms | 19.2 MB |  | LargeXlsx | 91.6% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | EPPlus 4.5.3.3 | 16.83 ms | 0 B |  | LargeXlsx | 502.5% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | ClosedXML | 18.25 ms | 10.9 MB |  | LargeXlsx | 553.1% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | EPPlus | 30.11 ms | 14.0 MB |  | LargeXlsx | 977.4% slower than OfficeIMO |
| 2500 | speed-comparison | autofit-existing | OfficeIMO.Excel | 9.68 ms | 1.9 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | autofit-existing | EPPlus | 104.81 ms | 49.5 MB |  | OfficeIMO.Excel | 983.2% slower than OfficeIMO |
| 2500 | speed-comparison | autofit-existing | EPPlus 4.5.3.3 | 112.86 ms | 0 B |  | OfficeIMO.Excel | 1066.4% slower than OfficeIMO |
| 2500 | speed-comparison | autofit-existing | ClosedXML | 181.55 ms | 82.6 MB |  | OfficeIMO.Excel | 1776.3% slower than OfficeIMO |
| 2500 | speed-comparison | build-object-datatable-dictionaries | OfficeIMO.Excel | 1.67 ms | 564.2 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | build-object-datatable-typed | OfficeIMO.Excel | 1.60 ms | 856.9 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-cells | OfficeIMO.Excel | 7.95 ms | 2.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-cells | ClosedXML | 40.45 ms | 16.6 MB |  | OfficeIMO.Excel | 409.0% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-cells | EPPlus | 40.75 ms | 19.7 MB |  | OfficeIMO.Excel | 412.8% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-first-column-from-wide-sheet | OfficeIMO.Excel | 5.17 ms | 523.5 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-first-column-from-wide-sheet | EPPlus | 32.06 ms | 12.8 MB |  | OfficeIMO.Excel | 520.1% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-first-column-from-wide-sheet | ClosedXML | 44.98 ms | 15.1 MB |  | OfficeIMO.Excel | 770.1% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-range | OfficeIMO.Excel | 7.62 ms | 2.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-range | EPPlus | 37.58 ms | 19.7 MB |  | OfficeIMO.Excel | 393.4% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-range | ClosedXML | 38.81 ms | 16.6 MB |  | OfficeIMO.Excel | 409.5% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-top-range | OfficeIMO.Excel | 0.77 ms | 285.5 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-top-range | EPPlus | 28.35 ms | 12.1 MB |  | OfficeIMO.Excel | 3592.3% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-top-range | ClosedXML | 37.77 ms | 15.0 MB |  | OfficeIMO.Excel | 4819.5% slower than OfficeIMO |
| 2500 | speed-comparison | formula-heavy-read | OfficeIMO.Excel | 3.23 ms | 706.7 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | formula-heavy-read | EPPlus 4.5.3.3 | 17.95 ms | 0 B |  | OfficeIMO.Excel | 456.1% slower than OfficeIMO |
| 2500 | speed-comparison | formula-heavy-read | EPPlus | 23.17 ms | 7.5 MB |  | OfficeIMO.Excel | 618.1% slower than OfficeIMO |
| 2500 | speed-comparison | formula-heavy-read | ClosedXML | 25.61 ms | 8.1 MB |  | OfficeIMO.Excel | 693.7% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | OfficeIMO.Excel | 2.99 ms | 2.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | large-shared-strings | MiniExcel | 5.00 ms | 20.6 MB |  | OfficeIMO.Excel | 67.2% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | EPPlus 4.5.3.3 | 12.65 ms | 0 B |  | OfficeIMO.Excel | 323.1% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | ClosedXML | 16.12 ms | 11.0 MB |  | OfficeIMO.Excel | 439.2% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | EPPlus | 26.57 ms | 12.5 MB |  | OfficeIMO.Excel | 788.4% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | OfficeIMO.Excel | 1.33 ms | 177.3 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | large-sparse-column-read | Sylvan.Data.Excel | 1.48 ms | 316.6 KB |  | OfficeIMO.Excel | 10.9% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | ExcelDataReader | 2.49 ms | 4.0 MB |  | OfficeIMO.Excel | 87.1% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | ClosedXML | 4.79 ms | 4.3 MB |  | OfficeIMO.Excel | 259.2% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | EPPlus 4.5.3.3 | 10.56 ms | 0 B |  | OfficeIMO.Excel | 692.3% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | MiniExcel | 17.43 ms | 45.1 MB |  | OfficeIMO.Excel | 1207.5% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | EPPlus | 48.44 ms | 42.1 MB |  | OfficeIMO.Excel | 3534.1% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | Sylvan.Data.Excel | 1.61 ms | 316.6 KB |  | Sylvan.Data.Excel | 30.8% faster than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | OfficeIMO.Excel | 2.33 ms | 177.3 KB |  | Sylvan.Data.Excel | Loss +44.4% |
| 2500 | speed-comparison | large-sparse-row-read | ExcelDataReader | 2.73 ms | 4.0 MB |  | Sylvan.Data.Excel | 17.3% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | ClosedXML | 4.32 ms | 4.3 MB |  | Sylvan.Data.Excel | 85.5% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | EPPlus 4.5.3.3 | 13.49 ms | 0 B |  | Sylvan.Data.Excel | 478.8% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | MiniExcel | 14.22 ms | 45.1 MB |  | Sylvan.Data.Excel | 510.0% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | EPPlus | 47.28 ms | 42.1 MB |  | Sylvan.Data.Excel | 1928.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | Sylvan.Data.Excel | 5.13 ms | 655.2 KB |  | Sylvan.Data.Excel | 18.8% faster than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | OfficeIMO.Excel | 6.31 ms | 374.5 KB |  | Sylvan.Data.Excel | Loss +23.2% |
| 2500 | speed-comparison | read-bottom-range | ExcelDataReader | 14.84 ms | 5.9 MB |  | Sylvan.Data.Excel | 135.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | MiniExcel | 17.45 ms | 18.2 MB |  | Sylvan.Data.Excel | 176.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | EPPlus | 31.90 ms | 12.1 MB |  | Sylvan.Data.Excel | 405.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | ClosedXML | 40.72 ms | 15.0 MB |  | Sylvan.Data.Excel | 544.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | OfficeIMO.Excel | 5.85 ms | 377.9 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | read-bottom-range-stream | Sylvan.Data.Excel | 6.08 ms | 655.2 KB |  | OfficeIMO.Excel | 4.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | ExcelDataReader | 17.06 ms | 5.9 MB |  | OfficeIMO.Excel | 191.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | MiniExcel | 20.60 ms | 18.2 MB |  | OfficeIMO.Excel | 252.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | EPPlus | 35.11 ms | 12.1 MB |  | OfficeIMO.Excel | 500.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | ClosedXML | 44.83 ms | 15.0 MB |  | OfficeIMO.Excel | 666.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | Sylvan.Data.Excel | 10.99 ms | 2.2 MB |  | Sylvan.Data.Excel | 48.9% faster than OfficeIMO |
| 2500 | speed-comparison | read-datatable | ExcelDataReader | 16.89 ms | 7.5 MB |  | Sylvan.Data.Excel | 21.5% faster than OfficeIMO |
| 2500 | speed-comparison | read-datatable | MiniExcel | 19.08 ms | 17.8 MB |  | Sylvan.Data.Excel | 11.3% faster than OfficeIMO |
| 2500 | speed-comparison | read-datatable | OfficeIMO.Excel | 21.50 ms | 3.5 MB |  | Sylvan.Data.Excel | Loss +95.7% |
| 2500 | speed-comparison | read-datatable | EPPlus | 44.57 ms | 21.2 MB |  | Sylvan.Data.Excel | 107.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | EPPlus 4.5.3.3 | 45.30 ms | 0 B |  | Sylvan.Data.Excel | 110.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | ClosedXML | 46.27 ms | 17.9 MB |  | Sylvan.Data.Excel | 115.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | OfficeIMO.Excel | 5.97 ms | 543.0 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | Sylvan.Data.Excel | 6.47 ms | 733.5 KB |  | OfficeIMO.Excel | 8.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | MiniExcel | 14.59 ms | 15.5 MB |  | OfficeIMO.Excel | 144.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | ExcelDataReader | 15.25 ms | 5.9 MB |  | OfficeIMO.Excel | 155.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | EPPlus | 32.44 ms | 12.8 MB |  | OfficeIMO.Excel | 443.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | ClosedXML | 43.22 ms | 15.1 MB |  | OfficeIMO.Excel | 623.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | Sylvan.Data.Excel | 8.37 ms | 895.3 KB |  | Sylvan.Data.Excel | 13.1% faster than OfficeIMO |
| 2500 | speed-comparison | read-objects | OfficeIMO.Excel | 9.62 ms | 2.4 MB |  | Sylvan.Data.Excel | Loss +15.0% |
| 2500 | speed-comparison | read-objects | ExcelDataReader | 16.84 ms | 6.2 MB |  | Sylvan.Data.Excel | 75.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | MiniExcel | 19.73 ms | 18.0 MB |  | Sylvan.Data.Excel | 105.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | EPPlus 4.5.3.3 | 39.04 ms | 0 B |  | Sylvan.Data.Excel | 305.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | ClosedXML | 41.15 ms | 16.5 MB |  | Sylvan.Data.Excel | 327.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | EPPlus | 41.83 ms | 20.9 MB |  | Sylvan.Data.Excel | 334.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | Sylvan.Data.Excel | 6.56 ms | 831.0 KB |  | Sylvan.Data.Excel | 10.3% faster than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | OfficeIMO.Excel | 7.32 ms | 2.4 MB |  | Sylvan.Data.Excel | Loss +11.4% |
| 2500 | speed-comparison | read-objects-stream | ExcelDataReader | 16.47 ms | 6.1 MB |  | Sylvan.Data.Excel | 125.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | MiniExcel | 19.62 ms | 18.0 MB |  | Sylvan.Data.Excel | 168.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | EPPlus 4.5.3.3 | 37.46 ms | 0 B |  | Sylvan.Data.Excel | 412.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | EPPlus | 45.75 ms | 20.8 MB |  | Sylvan.Data.Excel | 525.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | ClosedXML | 47.14 ms | 16.5 MB |  | Sylvan.Data.Excel | 544.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | OfficeIMO.Excel | 11.52 ms | 2.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | read-range | Sylvan.Data.Excel | 15.83 ms | 655.0 KB |  | OfficeIMO.Excel | 37.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | MiniExcel | 27.33 ms | 18.2 MB |  | OfficeIMO.Excel | 137.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | ExcelDataReader | 29.17 ms | 5.9 MB |  | OfficeIMO.Excel | 153.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | EPPlus 4.5.3.3 | 41.98 ms | 0 B |  | OfficeIMO.Excel | 264.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | EPPlus | 42.16 ms | 19.7 MB |  | OfficeIMO.Excel | 266.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | ClosedXML | 92.11 ms | 16.5 MB |  | OfficeIMO.Excel | 699.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | OfficeIMO.Excel | 6.46 ms | 2.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | read-range-decimal | Sylvan.Data.Excel | 7.72 ms | 750.3 KB |  | OfficeIMO.Excel | 19.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | ExcelDataReader | 14.77 ms | 5.9 MB |  | OfficeIMO.Excel | 128.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | MiniExcel | 16.22 ms | 18.2 MB |  | OfficeIMO.Excel | 151.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | EPPlus | 33.72 ms | 19.7 MB |  | OfficeIMO.Excel | 421.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | ClosedXML | 38.28 ms | 16.3 MB |  | OfficeIMO.Excel | 492.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | Sylvan.Data.Excel | 6.84 ms | 655.2 KB |  | Sylvan.Data.Excel | 58.8% faster than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | OfficeIMO.Excel | 16.61 ms | 2.7 MB |  | Sylvan.Data.Excel | Loss +142.7% |
| 2500 | speed-comparison | read-range-stream | ExcelDataReader | 17.35 ms | 5.9 MB |  | Sylvan.Data.Excel | 4.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | MiniExcel | 19.48 ms | 18.2 MB |  | Sylvan.Data.Excel | 17.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | EPPlus 4.5.3.3 | 38.28 ms | 0 B |  | Sylvan.Data.Excel | 130.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | EPPlus | 45.78 ms | 19.7 MB |  | Sylvan.Data.Excel | 175.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | ClosedXML | 46.72 ms | 16.3 MB |  | Sylvan.Data.Excel | 181.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | Sylvan.Data.Excel | 0.60 ms | 348.5 KB |  | Sylvan.Data.Excel | 21.8% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range | OfficeIMO.Excel | 0.77 ms | 296.0 KB |  | Sylvan.Data.Excel | Loss +27.8% |
| 2500 | speed-comparison | read-top-range | MiniExcel | 0.99 ms | 869.0 KB |  | Sylvan.Data.Excel | 29.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | ExcelDataReader | 6.38 ms | 1.9 MB |  | Sylvan.Data.Excel | 727.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | EPPlus | 33.20 ms | 12.1 MB |  | Sylvan.Data.Excel | 4209.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | EPPlus 4.5.3.3 | 35.67 ms | 0 B |  | Sylvan.Data.Excel | 4530.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | ClosedXML | 44.00 ms | 15.0 MB |  | Sylvan.Data.Excel | 5610.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | Sylvan.Data.Excel | 0.66 ms | 348.5 KB |  | Sylvan.Data.Excel | 5.4% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | OfficeIMO.Excel | 0.69 ms | 299.4 KB |  | Sylvan.Data.Excel | Loss +5.7% |
| 2500 | speed-comparison | read-top-range-stream | MiniExcel | 1.21 ms | 869.0 KB |  | Sylvan.Data.Excel | 74.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | ExcelDataReader | 6.59 ms | 1.9 MB |  | Sylvan.Data.Excel | 850.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | EPPlus 4.5.3.3 | 33.54 ms | 0 B |  | Sylvan.Data.Excel | 4733.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | EPPlus | 35.07 ms | 12.1 MB |  | Sylvan.Data.Excel | 4954.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | ClosedXML | 47.34 ms | 15.0 MB |  | Sylvan.Data.Excel | 6722.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | Sylvan.Data.Excel | 0.56 ms | 348.5 KB |  | Sylvan.Data.Excel | 21.2% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | OfficeIMO.Excel | 0.72 ms | 300.2 KB |  | Sylvan.Data.Excel | Loss +26.9% |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | MiniExcel | 1.01 ms | 869.0 KB |  | Sylvan.Data.Excel | 41.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | ExcelDataReader | 6.03 ms | 1.9 MB |  | Sylvan.Data.Excel | 742.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | EPPlus | 31.18 ms | 12.1 MB |  | Sylvan.Data.Excel | 4258.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | ClosedXML | 45.49 ms | 15.0 MB |  | Sylvan.Data.Excel | 6259.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-used-range | Sylvan.Data.Excel | 6.77 ms | 655.2 KB |  | Sylvan.Data.Excel | 20.5% faster than OfficeIMO |
| 2500 | speed-comparison | read-used-range | OfficeIMO.Excel | 8.52 ms | 2.7 MB |  | Sylvan.Data.Excel | Loss +25.8% |
| 2500 | speed-comparison | read-used-range | ExcelDataReader | 14.77 ms | 5.9 MB |  | Sylvan.Data.Excel | 73.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-used-range | MiniExcel | 19.20 ms | 18.2 MB |  | Sylvan.Data.Excel | 125.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-used-range | EPPlus | 36.49 ms | 19.7 MB |  | Sylvan.Data.Excel | 328.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-used-range | ClosedXML | 68.92 ms | 16.4 MB |  | Sylvan.Data.Excel | 709.2% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-autofilter | OfficeIMO.Excel | 4.26 ms | 1.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-autofilter | EPPlus 4.5.3.3 | 28.49 ms | 0 B |  | OfficeIMO.Excel | 568.1% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-autofilter | ClosedXML | 42.14 ms | 21.7 MB |  | OfficeIMO.Excel | 888.2% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-autofilter | EPPlus | 50.65 ms | 24.1 MB |  | OfficeIMO.Excel | 1087.7% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-charts | OfficeIMO.Excel | 6.57 ms | 1.8 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-charts | EPPlus 4.5.3.3 | 34.29 ms | 0 B |  | OfficeIMO.Excel | 422.1% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-charts | EPPlus | 54.79 ms | 26.5 MB |  | OfficeIMO.Excel | 734.2% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-conditional-formatting | OfficeIMO.Excel | 4.53 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-conditional-formatting | EPPlus 4.5.3.3 | 30.75 ms | 0 B |  | OfficeIMO.Excel | 578.7% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-conditional-formatting | ClosedXML | 37.66 ms | 21.8 MB |  | OfficeIMO.Excel | 731.1% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-conditional-formatting | EPPlus | 44.53 ms | 24.2 MB |  | OfficeIMO.Excel | 882.8% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-data-validation | OfficeIMO.Excel | 4.76 ms | 1.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-data-validation | EPPlus 4.5.3.3 | 28.70 ms | 0 B |  | OfficeIMO.Excel | 502.6% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-data-validation | ClosedXML | 40.88 ms | 21.7 MB |  | OfficeIMO.Excel | 758.2% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-data-validation | EPPlus | 47.64 ms | 24.1 MB |  | OfficeIMO.Excel | 900.3% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-freeze-panes | OfficeIMO.Excel | 4.25 ms | 1.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-freeze-panes | EPPlus 4.5.3.3 | 28.87 ms | 0 B |  | OfficeIMO.Excel | 579.3% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-freeze-panes | ClosedXML | 38.50 ms | 21.7 MB |  | OfficeIMO.Excel | 806.0% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-freeze-panes | EPPlus | 46.20 ms | 24.2 MB |  | OfficeIMO.Excel | 987.3% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-pivot-table | OfficeIMO.Excel | 10.98 ms | 5.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-pivot-table | EPPlus 4.5.3.3 | 33.03 ms | 0 B |  | OfficeIMO.Excel | 200.9% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-pivot-table | EPPlus | 55.47 ms | 28.8 MB |  | OfficeIMO.Excel | 405.4% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-all-in-one | OfficeIMO.Excel | 13.87 ms | 6.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-all-in-one | EPPlus | 85.20 ms | 53.3 MB |  | OfficeIMO.Excel | 514.2% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-all-in-one | EPPlus 4.5.3.3 | 85.69 ms | 0 B |  | OfficeIMO.Excel | 517.8% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-chart-first | OfficeIMO.Excel | 13.43 ms | 6.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-chart-first | EPPlus 4.5.3.3 | 81.33 ms | 0 B |  | OfficeIMO.Excel | 505.7% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-chart-first | EPPlus | 88.79 ms | 53.3 MB |  | OfficeIMO.Excel | 561.3% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-core | OfficeIMO.Excel | 5.27 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-core | EPPlus | 78.53 ms | 46.2 MB |  | OfficeIMO.Excel | 1390.6% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-core | EPPlus 4.5.3.3 | 85.01 ms | 0 B |  | OfficeIMO.Excel | 1513.7% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-core | ClosedXML | 92.13 ms | 68.2 MB |  | OfficeIMO.Excel | 1648.9% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-extra-column | OfficeIMO.Excel | 14.14 ms | 6.2 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-extra-column | EPPlus 4.5.3.3 | 88.80 ms | 0 B |  | OfficeIMO.Excel | 527.8% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-extra-column | EPPlus | 89.14 ms | 57.8 MB |  | OfficeIMO.Excel | 530.2% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-no-autofit | OfficeIMO.Excel | 13.55 ms | 6.0 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-no-autofit | EPPlus 4.5.3.3 | 39.42 ms | 0 B |  | OfficeIMO.Excel | 190.9% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-no-autofit | EPPlus | 54.57 ms | 32.1 MB |  | OfficeIMO.Excel | 302.6% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-post-mutation | OfficeIMO.Excel | 14.68 ms | 6.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-post-mutation | EPPlus 4.5.3.3 | 78.95 ms | 0 B |  | OfficeIMO.Excel | 437.8% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-post-mutation | EPPlus | 85.11 ms | 53.3 MB |  | OfficeIMO.Excel | 479.7% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-shuffled-columns | OfficeIMO.Excel | 15.82 ms | 6.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-shuffled-columns | EPPlus 4.5.3.3 | 81.45 ms | 0 B |  | OfficeIMO.Excel | 414.9% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-shuffled-columns | EPPlus | 84.62 ms | 53.3 MB |  | OfficeIMO.Excel | 434.9% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook | OfficeIMO.Excel | 26.10 ms | 7.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | report-workbook | EPPlus 4.5.3.3 | 108.50 ms | 0 B |  | OfficeIMO.Excel | 315.6% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook | EPPlus | 124.26 ms | 75.7 MB |  | OfficeIMO.Excel | 376.0% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-core | OfficeIMO.Excel | 7.48 ms | 2.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | report-workbook-core | EPPlus 4.5.3.3 | 103.99 ms | 0 B |  | OfficeIMO.Excel | 1290.3% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-core | EPPlus | 119.77 ms | 70.3 MB |  | OfficeIMO.Excel | 1501.3% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-core | ClosedXML | 123.93 ms | 94.9 MB |  | OfficeIMO.Excel | 1556.8% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-datatable | OfficeIMO.Excel | 18.81 ms | 7.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | report-workbook-datatable | EPPlus 4.5.3.3 | 102.31 ms | 0 B |  | OfficeIMO.Excel | 443.9% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-datatable | EPPlus | 125.61 ms | 64.4 MB |  | OfficeIMO.Excel | 567.8% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-datatable-core | OfficeIMO.Excel | 7.58 ms | 2.9 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | report-workbook-datatable-core | EPPlus 4.5.3.3 | 100.61 ms | 0 B |  | OfficeIMO.Excel | 1227.8% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-datatable-core | EPPlus | 114.61 ms | 59.1 MB |  | OfficeIMO.Excel | 1412.6% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-datatable-core | ClosedXML | 119.46 ms | 80.9 MB |  | OfficeIMO.Excel | 1476.7% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | Sylvan.Data.Excel | 2.41 ms | 518.6 KB |  | Sylvan.Data.Excel | 38.0% faster than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | OfficeIMO.Excel | 3.89 ms | 1.0 MB |  | Sylvan.Data.Excel | Loss +61.4% |
| 2500 | speed-comparison | shared-string-read | ExcelDataReader | 6.73 ms | 2.6 MB |  | Sylvan.Data.Excel | 73.0% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | MiniExcel | 7.48 ms | 7.4 MB |  | Sylvan.Data.Excel | 92.2% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | ClosedXML | 20.53 ms | 9.3 MB |  | Sylvan.Data.Excel | 427.1% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | EPPlus 4.5.3.3 | 20.53 ms | 0 B |  | Sylvan.Data.Excel | 427.3% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | EPPlus | 23.82 ms | 10.1 MB |  | Sylvan.Data.Excel | 511.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-blog-2023-20-string-columns | LargeXlsx | 4.22 ms | 857.6 KB |  | LargeXlsx | 6.4% faster than OfficeIMO |
| 2500 | speed-comparison | write-blog-2023-20-string-columns | OfficeIMO.Excel | 4.51 ms | 1.6 MB |  | LargeXlsx | Loss +6.9% |
| 2500 | speed-comparison | write-blog-2023-20-string-columns | MiniExcel | 16.08 ms | 35.1 MB |  | LargeXlsx | 256.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-blog-2023-20-string-columns | ClosedXML | 82.94 ms | 69.8 MB |  | LargeXlsx | 1740.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | OfficeIMO.Excel | 7.20 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-bulk-report | MiniExcel | 14.97 ms | 26.2 MB |  | OfficeIMO.Excel | 107.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | EPPlus 4.5.3.3 | 106.75 ms | 0 B |  | OfficeIMO.Excel | 1382.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | EPPlus | 132.23 ms | 48.0 MB |  | OfficeIMO.Excel | 1735.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | ClosedXML | 239.90 ms | 57.0 MB |  | OfficeIMO.Excel | 3230.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellformula | OfficeIMO.Excel | 4.33 ms | 1.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellformula | EPPlus 4.5.3.3 | 17.00 ms | 0 B |  | OfficeIMO.Excel | 292.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellformula | ClosedXML | 23.08 ms | 11.8 MB |  | OfficeIMO.Excel | 433.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellformula | EPPlus | 52.62 ms | 17.7 MB |  | OfficeIMO.Excel | 1115.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-empty-strings | OfficeIMO.Excel | 2.17 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-empty-strings | ClosedXML | 11.88 ms | 9.7 MB |  | OfficeIMO.Excel | 446.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-empty-strings | EPPlus | 20.60 ms | 11.5 MB |  | OfficeIMO.Excel | 847.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-numbers | OfficeIMO.Excel | 2.01 ms | 1.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-numbers | ClosedXML | 10.90 ms | 9.0 MB |  | OfficeIMO.Excel | 443.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-numbers | EPPlus 4.5.3.3 | 13.09 ms | 0 B |  | OfficeIMO.Excel | 552.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-numbers | EPPlus | 21.77 ms | 12.6 MB |  | OfficeIMO.Excel | 986.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-mixed | OfficeIMO.Excel | 3.34 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-object-mixed | EPPlus 4.5.3.3 | 15.61 ms | 0 B |  | OfficeIMO.Excel | 367.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-mixed | ClosedXML | 15.63 ms | 11.6 MB |  | OfficeIMO.Excel | 368.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-mixed | EPPlus | 28.99 ms | 15.3 MB |  | OfficeIMO.Excel | 767.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-sparse | OfficeIMO.Excel | 3.47 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-object-sparse | ClosedXML | 14.80 ms | 11.0 MB |  | OfficeIMO.Excel | 326.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-sparse | EPPlus | 27.67 ms | 14.6 MB |  | OfficeIMO.Excel | 697.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 3.65 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-object-sparse-batch | ClosedXML | 16.35 ms | 11.0 MB |  | OfficeIMO.Excel | 348.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-sparse-batch | EPPlus | 28.41 ms | 14.6 MB |  | OfficeIMO.Excel | 679.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-scalars | OfficeIMO.Excel | 1.85 ms | 1.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-scalars | ClosedXML | 10.26 ms | 8.8 MB |  | OfficeIMO.Excel | 453.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-scalars | EPPlus 4.5.3.3 | 12.40 ms | 0 B |  | OfficeIMO.Excel | 568.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-scalars | EPPlus | 22.34 ms | 12.5 MB |  | OfficeIMO.Excel | 1104.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings | OfficeIMO.Excel | 2.78 ms | 2.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-strings | ClosedXML | 10.96 ms | 11.0 MB |  | OfficeIMO.Excel | 294.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings | EPPlus 4.5.3.3 | 14.80 ms | 0 B |  | OfficeIMO.Excel | 432.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings | EPPlus | 19.78 ms | 12.5 MB |  | OfficeIMO.Excel | 611.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings-distinct | OfficeIMO.Excel | 2.66 ms | 2.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-strings-distinct | ClosedXML | 14.50 ms | 12.8 MB |  | OfficeIMO.Excel | 444.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings-distinct | EPPlus | 26.15 ms | 13.6 MB |  | OfficeIMO.Excel | 882.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings-repeated | OfficeIMO.Excel | 2.13 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-strings-repeated | ClosedXML | 10.64 ms | 9.0 MB |  | OfficeIMO.Excel | 399.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings-repeated | EPPlus | 19.63 ms | 11.1 MB |  | OfficeIMO.Excel | 821.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-temporal | OfficeIMO.Excel | 3.11 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-temporal | EPPlus 4.5.3.3 | 15.51 ms | 0 B |  | OfficeIMO.Excel | 399.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-temporal | ClosedXML | 15.59 ms | 9.5 MB |  | OfficeIMO.Excel | 402.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-temporal | EPPlus | 24.48 ms | 14.4 MB |  | OfficeIMO.Excel | 688.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 1.24 ms | 447.0 KB |  | LargeXlsx | 21.3% faster than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 1.57 ms | 1.1 MB |  | LargeXlsx | Loss +27.0% |
| 2500 | speed-comparison | write-cellvalues-headerless-rectangle-direct | ClosedXML | 12.07 ms | 10.0 MB |  | LargeXlsx | 667.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-headerless-rectangle-direct | EPPlus | 23.40 ms | 12.7 MB |  | LargeXlsx | 1386.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | LargeXlsx | 3.20 ms | 758.3 KB |  | LargeXlsx | 22.9% faster than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 4.14 ms | 2.0 MB |  | LargeXlsx | Loss +29.7% |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | MiniExcel | 9.01 ms | 22.7 MB |  | LargeXlsx | 117.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 31.77 ms | 0 B |  | LargeXlsx | 667.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | ClosedXML | 32.28 ms | 21.7 MB |  | LargeXlsx | 679.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus | 43.19 ms | 24.1 MB |  | LargeXlsx | 942.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-sparse-rectangle-direct | OfficeIMO.Excel | 2.27 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalues-sparse-rectangle-direct | ClosedXML | 14.29 ms | 11.0 MB |  | OfficeIMO.Excel | 530.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-sparse-rectangle-direct | EPPlus | 24.41 ms | 14.6 MB |  | OfficeIMO.Excel | 976.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | Sylvan.Data.Excel | 3.92 ms | 758.6 KB |  | Sylvan.Data.Excel | 12.3% faster than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | OfficeIMO.Excel | 4.47 ms | 1.7 MB |  | Sylvan.Data.Excel | Loss +14.0% |
| 2500 | speed-comparison | write-datareader-plain | LargeXlsx | 8.05 ms | 1.0 MB |  | Sylvan.Data.Excel | 80.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | MiniExcel | 8.37 ms | 22.5 MB |  | Sylvan.Data.Excel | 87.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | ClosedXML | 26.10 ms | 11.3 MB |  | Sylvan.Data.Excel | 483.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | EPPlus 4.5.3.3 | 27.49 ms | 0 B |  | Sylvan.Data.Excel | 514.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | EPPlus | 35.56 ms | 16.3 MB |  | Sylvan.Data.Excel | 695.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | OfficeIMO.Excel | 5.22 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datareader-table | MiniExcel | 8.19 ms | 22.5 MB |  | OfficeIMO.Excel | 56.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | ClosedXML | 38.75 ms | 18.6 MB |  | OfficeIMO.Excel | 641.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | EPPlus 4.5.3.3 | 40.00 ms | 0 B |  | OfficeIMO.Excel | 665.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | EPPlus | 41.39 ms | 16.3 MB |  | OfficeIMO.Excel | 692.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table-autofit | OfficeIMO.Excel | 6.87 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datareader-table-autofit | MiniExcel | 8.40 ms | 26.0 MB |  | OfficeIMO.Excel | 22.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table-autofit | EPPlus | 54.87 ms | 37.4 MB |  | OfficeIMO.Excel | 698.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table-autofit | EPPlus 4.5.3.3 | 63.29 ms | 0 B |  | OfficeIMO.Excel | 820.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table-autofit | ClosedXML | 74.64 ms | 57.0 MB |  | OfficeIMO.Excel | 985.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-headerless-tables | OfficeIMO.Excel | 5.96 ms | 2.0 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-headerless-tables | MiniExcel | 15.52 ms | 28.5 MB |  | OfficeIMO.Excel | 160.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-headerless-tables | ClosedXML | 53.71 ms | 18.5 MB |  | OfficeIMO.Excel | 800.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-headerless-tables | EPPlus | 95.66 ms | 18.0 MB |  | OfficeIMO.Excel | 1503.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | OfficeIMO.Excel | 6.99 ms | 2.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | LargeXlsx | 11.06 ms | 1.1 MB |  | OfficeIMO.Excel | 58.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | MiniExcel | 16.74 ms | 29.3 MB |  | OfficeIMO.Excel | 139.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | EPPlus | 62.15 ms | 21.4 MB |  | OfficeIMO.Excel | 788.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | ClosedXML | 74.89 ms | 26.8 MB |  | OfficeIMO.Excel | 970.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-tables | OfficeIMO.Excel | 7.23 ms | 2.8 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-sparse-tables | MiniExcel | 17.99 ms | 29.8 MB |  | OfficeIMO.Excel | 148.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-tables | ClosedXML | 62.65 ms | 26.8 MB |  | OfficeIMO.Excel | 766.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-tables | EPPlus | 75.41 ms | 22.1 MB |  | OfficeIMO.Excel | 942.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | OfficeIMO.Excel | 6.40 ms | 2.0 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-tables | MiniExcel | 17.91 ms | 28.0 MB |  | OfficeIMO.Excel | 179.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | EPPlus 4.5.3.3 | 50.19 ms | 0 B |  | OfficeIMO.Excel | 684.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | EPPlus | 93.79 ms | 19.0 MB |  | OfficeIMO.Excel | 1365.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | ClosedXML | 108.60 ms | 18.4 MB |  | OfficeIMO.Excel | 1597.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables-autofit | OfficeIMO.Excel | 6.06 ms | 2.0 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-tables-autofit | MiniExcel | 16.98 ms | 31.6 MB |  | OfficeIMO.Excel | 179.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables-autofit | EPPlus | 144.10 ms | 42.4 MB |  | OfficeIMO.Excel | 2276.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables-autofit | ClosedXML | 146.69 ms | 55.4 MB |  | OfficeIMO.Excel | 2319.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | OfficeIMO.Excel | 5.01 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datatable-direct | LargeXlsx | 7.23 ms | 1.1 MB |  | OfficeIMO.Excel | 44.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | MiniExcel | 10.41 ms | 22.5 MB |  | OfficeIMO.Excel | 107.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | ClosedXML | 31.52 ms | 11.3 MB |  | OfficeIMO.Excel | 528.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | EPPlus | 42.67 ms | 16.3 MB |  | OfficeIMO.Excel | 750.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | EPPlus 4.5.3.3 | 54.65 ms | 0 B |  | OfficeIMO.Excel | 990.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-object-table-direct | OfficeIMO.Excel | 4.57 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datatable-object-table-direct | MiniExcel | 9.61 ms | 22.3 MB |  | OfficeIMO.Excel | 110.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-object-table-direct | ClosedXML | 40.64 ms | 18.3 MB |  | OfficeIMO.Excel | 789.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-object-table-direct | EPPlus | 44.57 ms | 16.0 MB |  | OfficeIMO.Excel | 876.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | OfficeIMO.Excel | 4.70 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datatable-table-direct | MiniExcel | 10.20 ms | 22.5 MB |  | OfficeIMO.Excel | 117.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | EPPlus | 42.18 ms | 16.3 MB |  | OfficeIMO.Excel | 797.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | ClosedXML | 48.36 ms | 18.6 MB |  | OfficeIMO.Excel | 929.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | EPPlus 4.5.3.3 | 49.12 ms | 0 B |  | OfficeIMO.Excel | 945.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-dictionary-objects-table-direct | OfficeIMO.Excel | 6.69 ms | 2.0 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | LargeXlsx | 3.63 ms | 758.3 KB |  | LargeXlsx | 21.0% faster than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 4.60 ms | 1.7 MB |  | LargeXlsx | Loss +26.6% |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | MiniExcel | 9.76 ms | 22.7 MB |  | LargeXlsx | 112.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 30.21 ms | 0 B |  | LargeXlsx | 556.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | ClosedXML | 33.96 ms | 11.3 MB |  | LargeXlsx | 638.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus | 46.83 ms | 16.3 MB |  | LargeXlsx | 918.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 5.95 ms | 1.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 64.25 ms | 37.4 MB |  | OfficeIMO.Excel | 979.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 82.94 ms | 49.7 MB |  | OfficeIMO.Excel | 1292.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | LargeXlsx | 4.11 ms | 758.3 KB |  | LargeXlsx | 27.3% faster than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | OfficeIMO.Excel | 5.65 ms | 1.3 MB |  | LargeXlsx | Loss +37.5% |
| 2500 | speed-comparison | write-insertobjects-direct | MiniExcel | 11.13 ms | 22.7 MB |  | LargeXlsx | 96.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | EPPlus 4.5.3.3 | 26.76 ms | 0 B |  | LargeXlsx | 373.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | ClosedXML | 41.70 ms | 11.3 MB |  | LargeXlsx | 637.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | EPPlus | 54.21 ms | 16.3 MB |  | LargeXlsx | 858.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 5.93 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 74.10 ms | 37.4 MB |  | OfficeIMO.Excel | 1150.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 86.37 ms | 49.7 MB |  | OfficeIMO.Excel | 1357.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 4.09 ms | 758.3 KB |  | LargeXlsx | 24.1% faster than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 5.39 ms | 1.5 MB |  | LargeXlsx | Loss +31.8% |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | MiniExcel | 10.07 ms | 22.7 MB |  | LargeXlsx | 86.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | ClosedXML | 40.54 ms | 11.3 MB |  | LargeXlsx | 651.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | EPPlus | 49.90 ms | 16.3 MB |  | LargeXlsx | 825.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.96 ms | 1.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 55.27 ms | 27.9 MB |  | OfficeIMO.Excel | 1013.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 60.14 ms | 26.7 MB |  | OfficeIMO.Excel | 1111.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | LargeXlsx | 5.36 ms | 802.5 KB |  | LargeXlsx | 23.2% faster than OfficeIMO |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 6.99 ms | 2.3 MB |  | LargeXlsx | Loss +30.3% |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | MiniExcel | 11.77 ms | 24.6 MB |  | LargeXlsx | 68.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | ClosedXML | 49.70 ms | 16.6 MB |  | LargeXlsx | 611.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | EPPlus | 58.11 ms | 19.6 MB |  | LargeXlsx | 731.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-mixed-direct | LargeXlsx | 6.13 ms | 802.5 KB |  | LargeXlsx | 20.1% faster than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 7.68 ms | 1.5 MB |  | LargeXlsx | Loss +25.2% |
| 2500 | speed-comparison | write-powershell-psobject-mixed-direct | MiniExcel | 11.71 ms | 24.6 MB |  | LargeXlsx | 52.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-mixed-direct | ClosedXML | 49.52 ms | 16.6 MB |  | LargeXlsx | 545.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-mixed-direct | EPPlus | 57.50 ms | 19.6 MB |  | LargeXlsx | 648.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 23.95 ms | 4.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-powershell-psobject-wide-direct | LargeXlsx | 25.83 ms | 2.7 MB |  | OfficeIMO.Excel | 7.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-wide-direct | MiniExcel | 43.78 ms | 47.3 MB |  | OfficeIMO.Excel | 82.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-wide-direct | ClosedXML | 151.06 ms | 50.4 MB |  | OfficeIMO.Excel | 530.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-wide-direct | EPPlus | 182.97 ms | 67.5 MB |  | OfficeIMO.Excel | 663.9% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | Sylvan.Data.Excel | 37.49 ms | 394.1 KB |  | Sylvan.Data.Excel | 22.1% faster than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | OfficeIMO.Excel | 48.16 ms | 23.1 MB |  | Sylvan.Data.Excel | Loss +28.4% |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | ExcelDataReader | 106.96 ms | 67.9 MB |  | Sylvan.Data.Excel | 122.1% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | MiniExcel | 142.41 ms | 210.3 MB |  | Sylvan.Data.Excel | 195.7% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | Sylvan.Data.Excel | 37.07 ms | 394.1 KB |  | Sylvan.Data.Excel | 22.3% faster than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | OfficeIMO.Excel | 47.73 ms | 23.8 MB |  | Sylvan.Data.Excel | Loss +28.7% |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | ExcelDataReader | 107.61 ms | 67.9 MB |  | Sylvan.Data.Excel | 125.5% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | MiniExcel | 141.46 ms | 210.3 MB |  | Sylvan.Data.Excel | 196.4% slower than OfficeIMO |
| 25000 | package-profile | append-plain-rows | LargeXlsx | 11.10 ms | 2.7 MB | 605.0 KB | LargeXlsx | 22.1% faster than OfficeIMO |
| 25000 | package-profile | append-plain-rows | OfficeIMO.Excel | 14.25 ms | 10.6 MB | 610.4 KB | LargeXlsx | Loss +28.4% |
| 25000 | package-profile | append-plain-rows | MiniExcel | 29.97 ms | 56.9 MB | 642.3 KB | LargeXlsx | 110.3% slower than OfficeIMO |
| 25000 | package-profile | append-plain-rows | ClosedXML | 128.65 ms | 101.8 MB | 540.6 KB | LargeXlsx | 802.7% slower than OfficeIMO |
| 25000 | package-profile | append-plain-rows | EPPlus | 198.66 ms | 98.0 MB | 525.6 KB | LargeXlsx | 1293.8% slower than OfficeIMO |
| 25000 | package-profile | autofit-existing | OfficeIMO.Excel | 85.90 ms | 15.3 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | autofit-existing | EPPlus | 452.28 ms | 245.1 MB | 1.1 MB | OfficeIMO.Excel | 426.6% slower than OfficeIMO |
| 25000 | package-profile | autofit-existing | ClosedXML | 1.32 s | 810.3 MB | 1.1 MB | OfficeIMO.Excel | 1431.4% slower than OfficeIMO |
| 25000 | package-profile | large-shared-strings | OfficeIMO.Excel | 14.85 ms | 15.4 MB | 529.7 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | large-shared-strings | MiniExcel | 27.87 ms | 72.0 MB | 581.0 KB | OfficeIMO.Excel | 87.6% slower than OfficeIMO |
| 25000 | package-profile | large-shared-strings | ClosedXML | 108.66 ms | 101.8 MB | 460.1 KB | OfficeIMO.Excel | 631.5% slower than OfficeIMO |
| 25000 | package-profile | large-shared-strings | EPPlus | 184.64 ms | 82.4 MB | 444.7 KB | OfficeIMO.Excel | 1143.0% slower than OfficeIMO |
| 25000 | package-profile | realworld-autofilter | OfficeIMO.Excel | 30.76 ms | 11.2 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-autofilter | ClosedXML | 294.92 ms | 205.7 MB | 1.1 MB | OfficeIMO.Excel | 858.7% slower than OfficeIMO |
| 25000 | package-profile | realworld-autofilter | EPPlus | 367.81 ms | 206.9 MB | 1.1 MB | OfficeIMO.Excel | 1095.6% slower than OfficeIMO |
| 25000 | package-profile | realworld-charts | OfficeIMO.Excel | 36.44 ms | 12.3 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-charts | EPPlus | 374.31 ms | 209.9 MB | 1.1 MB | OfficeIMO.Excel | 927.3% slower than OfficeIMO |
| 25000 | package-profile | realworld-conditional-formatting | OfficeIMO.Excel | 30.36 ms | 11.3 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-conditional-formatting | ClosedXML | 286.34 ms | 205.8 MB | 1.1 MB | OfficeIMO.Excel | 843.1% slower than OfficeIMO |
| 25000 | package-profile | realworld-conditional-formatting | EPPlus | 361.78 ms | 206.9 MB | 1.1 MB | OfficeIMO.Excel | 1091.6% slower than OfficeIMO |
| 25000 | package-profile | realworld-data-validation | OfficeIMO.Excel | 32.26 ms | 11.2 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-data-validation | ClosedXML | 308.34 ms | 205.7 MB | 1.1 MB | OfficeIMO.Excel | 855.9% slower than OfficeIMO |
| 25000 | package-profile | realworld-data-validation | EPPlus | 377.00 ms | 206.9 MB | 1.1 MB | OfficeIMO.Excel | 1068.7% slower than OfficeIMO |
| 25000 | package-profile | realworld-freeze-panes | OfficeIMO.Excel | 31.38 ms | 11.2 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-freeze-panes | ClosedXML | 290.90 ms | 205.7 MB | 1.1 MB | OfficeIMO.Excel | 826.9% slower than OfficeIMO |
| 25000 | package-profile | realworld-freeze-panes | EPPlus | 372.05 ms | 206.9 MB | 1.1 MB | OfficeIMO.Excel | 1085.5% slower than OfficeIMO |
| 25000 | package-profile | realworld-pivot-table | OfficeIMO.Excel | 81.89 ms | 41.2 MB | 1.9 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-pivot-table | EPPlus | 395.47 ms | 225.4 MB | 1.1 MB | OfficeIMO.Excel | 382.9% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-all-in-one | OfficeIMO.Excel | 88.60 ms | 42.7 MB | 1.9 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-all-in-one | EPPlus | 422.81 ms | 270.6 MB | 1.1 MB | OfficeIMO.Excel | 377.2% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-chart-first | OfficeIMO.Excel | 90.24 ms | 42.5 MB | 1.9 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-chart-first | EPPlus | 424.94 ms | 270.6 MB | 1.1 MB | OfficeIMO.Excel | 370.9% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-core | OfficeIMO.Excel | 34.89 ms | 11.4 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-core | EPPlus | 413.51 ms | 249.1 MB | 1.1 MB | OfficeIMO.Excel | 1085.3% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-core | ClosedXML | 784.41 ms | 664.2 MB | 1.1 MB | OfficeIMO.Excel | 2148.5% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-extra-column | OfficeIMO.Excel | 99.55 ms | 44.5 MB | 2.1 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-extra-column | EPPlus | 469.40 ms | 295.7 MB | 1.1 MB | OfficeIMO.Excel | 371.5% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-no-autofit | OfficeIMO.Excel | 82.57 ms | 42.6 MB | 1.9 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-no-autofit | EPPlus | 389.25 ms | 229.3 MB | 1.1 MB | OfficeIMO.Excel | 371.4% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-post-mutation | OfficeIMO.Excel | 91.54 ms | 42.7 MB | 1.9 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-post-mutation | EPPlus | 433.54 ms | 270.6 MB | 1.1 MB | OfficeIMO.Excel | 373.6% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-shuffled-columns | OfficeIMO.Excel | 100.92 ms | 42.7 MB | 2.0 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-shuffled-columns | EPPlus | 435.26 ms | 270.6 MB | 1.1 MB | OfficeIMO.Excel | 331.3% slower than OfficeIMO |
| 25000 | package-profile | report-workbook | OfficeIMO.Excel | 114.88 ms | 57.8 MB | 2.6 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | report-workbook | EPPlus | 569.92 ms | 356.2 MB | 1.5 MB | OfficeIMO.Excel | 396.1% slower than OfficeIMO |
| 25000 | package-profile | report-workbook-core | OfficeIMO.Excel | 48.41 ms | 10.7 MB | 1.8 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | report-workbook-core | EPPlus | 578.88 ms | 334.8 MB | 1.5 MB | OfficeIMO.Excel | 1095.7% slower than OfficeIMO |
| 25000 | package-profile | report-workbook-core | ClosedXML | 1.19 s | 952.9 MB | 1.5 MB | OfficeIMO.Excel | 2348.8% slower than OfficeIMO |
| 25000 | package-profile | report-workbook-datatable | OfficeIMO.Excel | 125.75 ms | 60.5 MB | 2.6 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | report-workbook-datatable | EPPlus | 589.39 ms | 242.0 MB | 1.5 MB | OfficeIMO.Excel | 368.7% slower than OfficeIMO |
| 25000 | package-profile | report-workbook-datatable-core | OfficeIMO.Excel | 48.98 ms | 13.4 MB | 1.8 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | report-workbook-datatable-core | EPPlus | 519.39 ms | 220.7 MB | 1.5 MB | OfficeIMO.Excel | 960.4% slower than OfficeIMO |
| 25000 | package-profile | report-workbook-datatable-core | ClosedXML | 1.02 s | 812.7 MB | 1.5 MB | OfficeIMO.Excel | 1979.1% slower than OfficeIMO |
| 25000 | package-profile | write-blog-2023-20-string-columns | LargeXlsx | 52.98 ms | 10.5 MB | 2.4 MB | LargeXlsx | 15.1% faster than OfficeIMO |
| 25000 | package-profile | write-blog-2023-20-string-columns | OfficeIMO.Excel | 62.44 ms | 11.4 MB | 2.2 MB | LargeXlsx | Loss +17.8% |
| 25000 | package-profile | write-blog-2023-20-string-columns | MiniExcel | 189.99 ms | 221.6 MB | 2.4 MB | LargeXlsx | 204.3% slower than OfficeIMO |
| 25000 | package-profile | write-blog-2023-20-string-columns | ClosedXML | 1.11 s | 742.0 MB | 2.5 MB | LargeXlsx | 1684.0% slower than OfficeIMO |
| 25000 | package-profile | write-bulk-report | OfficeIMO.Excel | 35.39 ms | 11.3 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-bulk-report | MiniExcel | 66.86 ms | 122.6 MB | 1.5 MB | OfficeIMO.Excel | 88.9% slower than OfficeIMO |
| 25000 | package-profile | write-bulk-report | EPPlus | 419.29 ms | 249.0 MB | 1.1 MB | OfficeIMO.Excel | 1084.7% slower than OfficeIMO |
| 25000 | package-profile | write-bulk-report | ClosedXML | 718.65 ms | 552.7 MB | 1.1 MB | OfficeIMO.Excel | 1930.5% slower than OfficeIMO |
| 25000 | package-profile | write-cellformula | OfficeIMO.Excel | 24.74 ms | 9.9 MB | 670.3 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellformula | ClosedXML | 231.30 ms | 111.2 MB | 643.2 KB | OfficeIMO.Excel | 834.7% slower than OfficeIMO |
| 25000 | package-profile | write-cellformula | EPPlus | 407.68 ms | 137.4 MB | 593.9 KB | OfficeIMO.Excel | 1547.6% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-empty-strings | OfficeIMO.Excel | 14.96 ms | 6.7 MB | 451.4 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-empty-strings | ClosedXML | 152.22 ms | 90.7 MB | 398.1 KB | OfficeIMO.Excel | 917.8% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-empty-strings | EPPlus | 196.67 ms | 72.7 MB | 390.6 KB | OfficeIMO.Excel | 1215.0% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-numbers | OfficeIMO.Excel | 22.16 ms | 5.8 MB | 462.6 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-numbers | ClosedXML | 141.31 ms | 82.2 MB | 411.4 KB | OfficeIMO.Excel | 537.7% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-numbers | EPPlus | 237.33 ms | 84.4 MB | 406.5 KB | OfficeIMO.Excel | 971.0% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-mixed | OfficeIMO.Excel | 26.90 ms | 8.1 MB | 585.1 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-object-mixed | ClosedXML | 204.31 ms | 108.5 MB | 532.9 KB | OfficeIMO.Excel | 659.6% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-mixed | EPPlus | 256.47 ms | 110.6 MB | 544.3 KB | OfficeIMO.Excel | 853.5% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-sparse | OfficeIMO.Excel | 24.34 ms | 7.2 MB | 607.1 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-object-sparse | ClosedXML | 199.93 ms | 102.8 MB | 468.0 KB | OfficeIMO.Excel | 721.5% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-sparse | EPPlus | 306.84 ms | 103.8 MB | 494.4 KB | OfficeIMO.Excel | 1160.8% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 20.17 ms | 7.2 MB | 607.1 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-object-sparse-batch | ClosedXML | 171.99 ms | 102.8 MB | 468.0 KB | OfficeIMO.Excel | 752.6% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-sparse-batch | EPPlus | 245.03 ms | 103.8 MB | 494.4 KB | OfficeIMO.Excel | 1114.7% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-scalars | OfficeIMO.Excel | 14.54 ms | 6.0 MB | 441.9 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-scalars | ClosedXML | 125.06 ms | 80.7 MB | 394.9 KB | OfficeIMO.Excel | 760.3% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-scalars | EPPlus | 227.11 ms | 83.1 MB | 379.3 KB | OfficeIMO.Excel | 1462.3% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings | OfficeIMO.Excel | 21.69 ms | 15.0 MB | 527.8 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-strings | ClosedXML | 155.26 ms | 101.8 MB | 460.1 KB | OfficeIMO.Excel | 616.0% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings | EPPlus | 217.36 ms | 82.4 MB | 444.7 KB | OfficeIMO.Excel | 902.3% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings-distinct | OfficeIMO.Excel | 18.79 ms | 13.5 MB | 499.0 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-strings-distinct | ClosedXML | 189.65 ms | 128.4 MB | 555.3 KB | OfficeIMO.Excel | 909.4% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings-distinct | EPPlus | 262.96 ms | 95.4 MB | 565.1 KB | OfficeIMO.Excel | 1299.6% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings-repeated | OfficeIMO.Excel | 15.44 ms | 7.3 MB | 376.0 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-strings-repeated | ClosedXML | 133.40 ms | 82.5 MB | 331.8 KB | OfficeIMO.Excel | 764.1% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings-repeated | EPPlus | 214.21 ms | 68.4 MB | 300.8 KB | OfficeIMO.Excel | 1287.5% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-temporal | OfficeIMO.Excel | 27.28 ms | 7.3 MB | 620.5 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-temporal | ClosedXML | 196.97 ms | 87.2 MB | 483.0 KB | OfficeIMO.Excel | 622.1% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-temporal | EPPlus | 273.95 ms | 101.4 MB | 495.1 KB | OfficeIMO.Excel | 904.3% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 13.10 ms | 3.4 MB | 443.4 KB | LargeXlsx | 11.4% faster than OfficeIMO |
| 25000 | package-profile | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 14.79 ms | 6.8 MB | 455.5 KB | LargeXlsx | Loss +12.9% |
| 25000 | package-profile | write-cellvalues-headerless-rectangle-direct | ClosedXML | 161.16 ms | 93.8 MB | 467.5 KB | LargeXlsx | 989.8% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-headerless-rectangle-direct | EPPlus | 227.82 ms | 85.4 MB | 484.1 KB | LargeXlsx | 1440.6% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | LargeXlsx | 33.84 ms | 5.5 MB | 1.4 MB | LargeXlsx | 24.7% faster than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 44.94 ms | 15.7 MB | 1.4 MB | LargeXlsx | Loss +32.8% |
| 25000 | package-profile | write-cellvalues-rectangle-direct | MiniExcel | 89.91 ms | 91.1 MB | 1.5 MB | LargeXlsx | 100.1% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | ClosedXML | 356.62 ms | 205.7 MB | 1.1 MB | LargeXlsx | 693.6% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | EPPlus | 431.05 ms | 206.9 MB | 1.1 MB | LargeXlsx | 859.2% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | Sylvan.Data.Excel | 38.53 ms | 5.6 MB | 755.4 KB | Sylvan.Data.Excel | 18.8% faster than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | LargeXlsx | 45.15 ms | 8.2 MB | 1.4 MB | Sylvan.Data.Excel | 4.9% faster than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | OfficeIMO.Excel | 47.47 ms | 12.7 MB | 1.4 MB | Sylvan.Data.Excel | Loss +23.2% |
| 25000 | package-profile | write-datareader-plain | MiniExcel | 85.99 ms | 90.0 MB | 1.5 MB | Sylvan.Data.Excel | 81.1% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | ClosedXML | 340.53 ms | 101.8 MB | 1.1 MB | Sylvan.Data.Excel | 617.3% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | EPPlus | 407.27 ms | 114.7 MB | 1.1 MB | Sylvan.Data.Excel | 757.9% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table | OfficeIMO.Excel | 51.03 ms | 12.7 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-datareader-table | MiniExcel | 101.73 ms | 90.0 MB | 1.5 MB | OfficeIMO.Excel | 99.4% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table | EPPlus | 452.81 ms | 114.7 MB | 1.1 MB | OfficeIMO.Excel | 787.4% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table | ClosedXML | 512.33 ms | 169.3 MB | 1.1 MB | OfficeIMO.Excel | 904.0% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table-autofit | OfficeIMO.Excel | 45.06 ms | 12.7 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-datareader-table-autofit | MiniExcel | 94.73 ms | 121.6 MB | 1.5 MB | OfficeIMO.Excel | 110.2% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table-autofit | EPPlus | 433.23 ms | 156.0 MB | 1.1 MB | OfficeIMO.Excel | 861.4% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table-autofit | ClosedXML | 972.91 ms | 552.9 MB | 1.1 MB | OfficeIMO.Excel | 2059.0% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-direct-export | OfficeIMO.Excel | 31.35 ms | 9.6 MB | 1.3 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-dataset-sparse-direct-export | LargeXlsx | 34.77 ms | 9.0 MB | 1.6 MB | OfficeIMO.Excel | 10.9% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-direct-export | MiniExcel | 94.69 ms | 105.6 MB | 1.8 MB | OfficeIMO.Excel | 202.1% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-direct-export | EPPlus | 475.12 ms | 132.5 MB | 1.4 MB | OfficeIMO.Excel | 1415.6% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-direct-export | ClosedXML | 516.93 ms | 273.8 MB | 1.5 MB | OfficeIMO.Excel | 1549.0% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-tables | OfficeIMO.Excel | 37.94 ms | 13.1 MB | 1.8 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-dataset-sparse-tables | MiniExcel | 95.12 ms | 105.6 MB | 1.8 MB | OfficeIMO.Excel | 150.7% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-tables | EPPlus | 501.01 ms | 132.5 MB | 1.4 MB | OfficeIMO.Excel | 1220.4% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-tables | ClosedXML | 528.34 ms | 273.8 MB | 1.5 MB | OfficeIMO.Excel | 1292.4% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables | OfficeIMO.Excel | 32.75 ms | 10.0 MB | 1.3 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-dataset-tables | MiniExcel | 74.48 ms | 94.8 MB | 1.5 MB | OfficeIMO.Excel | 127.4% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables | EPPlus | 315.41 ms | 108.2 MB | 1.1 MB | OfficeIMO.Excel | 863.2% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables | ClosedXML | 334.31 ms | 168.0 MB | 1.1 MB | OfficeIMO.Excel | 920.9% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables-autofit | OfficeIMO.Excel | 37.93 ms | 10.1 MB | 1.3 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-dataset-tables-autofit | MiniExcel | 80.85 ms | 125.9 MB | 1.5 MB | OfficeIMO.Excel | 113.1% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables-autofit | EPPlus | 413.80 ms | 190.8 MB | 1.1 MB | OfficeIMO.Excel | 990.9% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables-autofit | ClosedXML | 724.66 ms | 537.2 MB | 1.1 MB | OfficeIMO.Excel | 1810.4% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | LargeXlsx | 46.80 ms | 9.3 MB | 1.4 MB | LargeXlsx | 3.9% faster than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | OfficeIMO.Excel | 48.70 ms | 12.4 MB | 1.4 MB | LargeXlsx | Loss +4.1% |
| 25000 | package-profile | write-datatable-direct | MiniExcel | 110.39 ms | 90.2 MB | 1.5 MB | LargeXlsx | 126.7% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | ClosedXML | 367.17 ms | 101.8 MB | 1.1 MB | LargeXlsx | 653.9% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | EPPlus | 443.45 ms | 114.7 MB | 1.1 MB | LargeXlsx | 810.6% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-table-direct | OfficeIMO.Excel | 43.59 ms | 12.4 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-datatable-table-direct | MiniExcel | 108.29 ms | 90.2 MB | 1.5 MB | OfficeIMO.Excel | 148.4% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-table-direct | EPPlus | 437.71 ms | 114.7 MB | 1.1 MB | OfficeIMO.Excel | 904.2% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-table-direct | ClosedXML | 501.41 ms | 169.3 MB | 1.1 MB | OfficeIMO.Excel | 1050.3% slower than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | LargeXlsx | 27.15 ms | 5.5 MB | 1.4 MB | LargeXlsx | 13.5% faster than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 31.39 ms | 12.6 MB | 1.4 MB | LargeXlsx | Loss +15.6% |
| 25000 | package-profile | write-fluent-rowsfrom-direct | MiniExcel | 64.33 ms | 91.1 MB | 1.5 MB | LargeXlsx | 105.0% slower than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | ClosedXML | 263.69 ms | 101.8 MB | 1.1 MB | LargeXlsx | 740.1% slower than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | EPPlus | 323.15 ms | 114.7 MB | 1.1 MB | LargeXlsx | 929.6% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 41.71 ms | 11.2 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 477.62 ms | 156.0 MB | 1.1 MB | OfficeIMO.Excel | 1045.1% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 819.12 ms | 485.3 MB | 1.1 MB | OfficeIMO.Excel | 1863.9% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | LargeXlsx | 37.23 ms | 5.5 MB | 1.4 MB | LargeXlsx | 11.1% faster than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | OfficeIMO.Excel | 41.89 ms | 11.2 MB | 1.4 MB | LargeXlsx | Loss +12.5% |
| 25000 | package-profile | write-insertobjects-direct | MiniExcel | 83.62 ms | 91.1 MB | 1.5 MB | LargeXlsx | 99.6% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | ClosedXML | 342.11 ms | 101.8 MB | 1.1 MB | LargeXlsx | 716.6% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | EPPlus | 462.26 ms | 114.7 MB | 1.1 MB | LargeXlsx | 1003.4% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 41.91 ms | 9.9 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 376.33 ms | 156.0 MB | 1.1 MB | OfficeIMO.Excel | 798.0% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 660.30 ms | 485.3 MB | 1.1 MB | OfficeIMO.Excel | 1475.6% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 28.36 ms | 5.5 MB | 1.4 MB | LargeXlsx | 23.7% faster than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 37.17 ms | 9.9 MB | 1.4 MB | LargeXlsx | Loss +31.1% |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | MiniExcel | 83.02 ms | 91.1 MB | 1.5 MB | LargeXlsx | 123.3% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | ClosedXML | 276.14 ms | 101.8 MB | 1.1 MB | LargeXlsx | 642.8% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | EPPlus | 335.68 ms | 114.7 MB | 1.1 MB | LargeXlsx | 803.0% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | LargeXlsx | 25.79 ms | 5.5 MB | 1.4 MB | LargeXlsx | 34.1% faster than OfficeIMO |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | OfficeIMO.Excel | 39.14 ms | 15.4 MB | 1.4 MB | LargeXlsx | Loss +51.8% |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | MiniExcel | 59.78 ms | 91.1 MB | 1.5 MB | LargeXlsx | 52.7% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | ClosedXML | 257.84 ms | 101.8 MB | 1.1 MB | LargeXlsx | 558.7% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | EPPlus | 322.95 ms | 114.7 MB | 1.1 MB | LargeXlsx | 725.1% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 33.39 ms | 11.2 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 373.66 ms | 135.1 MB | 1.1 MB | OfficeIMO.Excel | 1019.0% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 441.73 ms | 269.0 MB | 1.1 MB | OfficeIMO.Excel | 1222.9% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-mixed-objects-direct | LargeXlsx | 37.33 ms | 5.9 MB | 1.8 MB | LargeXlsx | 17.1% faster than OfficeIMO |
| 25000 | package-profile | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 45.00 ms | 10.3 MB | 1.8 MB | LargeXlsx | Loss +20.6% |
| 25000 | package-profile | write-powershell-mixed-objects-direct | MiniExcel | 79.84 ms | 111.3 MB | 1.9 MB | LargeXlsx | 77.4% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-mixed-objects-direct | ClosedXML | 359.92 ms | 175.3 MB | 1.5 MB | LargeXlsx | 699.8% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-mixed-objects-direct | EPPlus | 438.18 ms | 141.5 MB | 1.4 MB | LargeXlsx | 873.7% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-mixed-direct | LargeXlsx | 37.01 ms | 5.9 MB | 1.8 MB | LargeXlsx | 12.7% faster than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 42.38 ms | 9.7 MB | 1.8 MB | LargeXlsx | Loss +14.5% |
| 25000 | package-profile | write-powershell-psobject-mixed-direct | MiniExcel | 79.29 ms | 111.3 MB | 1.9 MB | LargeXlsx | 87.1% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-mixed-direct | ClosedXML | 371.02 ms | 175.3 MB | 1.5 MB | LargeXlsx | 775.4% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-mixed-direct | EPPlus | 448.05 ms | 141.5 MB | 1.4 MB | LargeXlsx | 957.1% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 190.90 ms | 35.3 MB | 6.6 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-powershell-psobject-wide-direct | LargeXlsx | 198.09 ms | 22.7 MB | 6.5 MB | OfficeIMO.Excel | 3.8% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-wide-direct | MiniExcel | 336.82 ms | 339.8 MB | 6.8 MB | OfficeIMO.Excel | 76.4% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-wide-direct | ClosedXML | 1.17 s | 476.0 MB | 6.0 MB | OfficeIMO.Excel | 511.1% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-wide-direct | EPPlus | 1.51 s | 549.7 MB | 5.3 MB | OfficeIMO.Excel | 693.1% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | LargeXlsx | 15.01 ms | 2.7 MB |  | LargeXlsx | 22.7% faster than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | OfficeIMO.Excel | 19.42 ms | 10.6 MB |  | LargeXlsx | Loss +29.4% |
| 25000 | speed-comparison | append-plain-rows | MiniExcel | 40.00 ms | 56.9 MB |  | LargeXlsx | 106.0% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | EPPlus 4.5.3.3 | 117.61 ms | 0 B |  | LargeXlsx | 505.6% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | ClosedXML | 150.45 ms | 101.8 MB |  | LargeXlsx | 674.7% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | EPPlus | 226.42 ms | 98.0 MB |  | LargeXlsx | 1065.9% slower than OfficeIMO |
| 25000 | speed-comparison | autofit-existing | OfficeIMO.Excel | 84.25 ms | 15.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | autofit-existing | EPPlus | 504.64 ms | 245.1 MB |  | OfficeIMO.Excel | 499.0% slower than OfficeIMO |
| 25000 | speed-comparison | autofit-existing | EPPlus 4.5.3.3 | 652.49 ms | 0 B |  | OfficeIMO.Excel | 674.5% slower than OfficeIMO |
| 25000 | speed-comparison | autofit-existing | ClosedXML | 1.56 s | 810.4 MB |  | OfficeIMO.Excel | 1755.3% slower than OfficeIMO |
| 25000 | speed-comparison | build-object-datatable-dictionaries | OfficeIMO.Excel | 12.67 ms | 5.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | build-object-datatable-typed | OfficeIMO.Excel | 10.47 ms | 7.9 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-cells | OfficeIMO.Excel | 55.93 ms | 24.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-cells | EPPlus | 318.84 ms | 183.0 MB |  | OfficeIMO.Excel | 470.0% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-cells | ClosedXML | 440.20 ms | 162.6 MB |  | OfficeIMO.Excel | 687.0% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-first-column-from-wide-sheet | OfficeIMO.Excel | 41.78 ms | 3.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-first-column-from-wide-sheet | EPPlus | 273.17 ms | 112.8 MB |  | OfficeIMO.Excel | 553.8% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-first-column-from-wide-sheet | ClosedXML | 394.13 ms | 147.4 MB |  | OfficeIMO.Excel | 843.3% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-range | OfficeIMO.Excel | 67.22 ms | 24.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-range | EPPlus | 302.40 ms | 183.0 MB |  | OfficeIMO.Excel | 349.8% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-range | ClosedXML | 390.28 ms | 162.6 MB |  | OfficeIMO.Excel | 480.6% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-top-range | OfficeIMO.Excel | 0.74 ms | 285.3 KB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-top-range | EPPlus | 257.85 ms | 103.1 MB |  | OfficeIMO.Excel | 34874.7% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-top-range | ClosedXML | 396.23 ms | 145.9 MB |  | OfficeIMO.Excel | 53645.7% slower than OfficeIMO |
| 25000 | speed-comparison | formula-heavy-read | OfficeIMO.Excel | 19.08 ms | 6.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | formula-heavy-read | EPPlus 4.5.3.3 | 73.40 ms | 0 B |  | OfficeIMO.Excel | 284.7% slower than OfficeIMO |
| 25000 | speed-comparison | formula-heavy-read | EPPlus | 151.27 ms | 69.2 MB |  | OfficeIMO.Excel | 692.9% slower than OfficeIMO |
| 25000 | speed-comparison | formula-heavy-read | ClosedXML | 170.22 ms | 77.7 MB |  | OfficeIMO.Excel | 792.2% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | OfficeIMO.Excel | 14.54 ms | 15.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | large-shared-strings | MiniExcel | 28.29 ms | 72.0 MB |  | OfficeIMO.Excel | 94.5% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | EPPlus 4.5.3.3 | 89.78 ms | 0 B |  | OfficeIMO.Excel | 517.3% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | ClosedXML | 110.05 ms | 101.8 MB |  | OfficeIMO.Excel | 656.7% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | EPPlus | 183.41 ms | 82.4 MB |  | OfficeIMO.Excel | 1161.1% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | OfficeIMO.Excel | 0.84 ms | 177.3 KB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | large-sparse-column-read | Sylvan.Data.Excel | 1.04 ms | 316.6 KB |  | OfficeIMO.Excel | 23.9% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | ExcelDataReader | 1.80 ms | 4.0 MB |  | OfficeIMO.Excel | 113.6% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | ClosedXML | 3.70 ms | 4.3 MB |  | OfficeIMO.Excel | 339.4% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | MiniExcel | 12.33 ms | 45.1 MB |  | OfficeIMO.Excel | 1366.2% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | EPPlus 4.5.3.3 | 32.60 ms | 0 B |  | OfficeIMO.Excel | 3775.5% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | EPPlus | 93.94 ms | 42.1 MB |  | OfficeIMO.Excel | 11067.9% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | OfficeIMO.Excel | 0.83 ms | 177.4 KB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | large-sparse-row-read | Sylvan.Data.Excel | 1.02 ms | 316.6 KB |  | OfficeIMO.Excel | 22.8% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | ExcelDataReader | 1.97 ms | 4.0 MB |  | OfficeIMO.Excel | 136.9% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | ClosedXML | 3.36 ms | 4.3 MB |  | OfficeIMO.Excel | 304.4% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | MiniExcel | 12.87 ms | 45.1 MB |  | OfficeIMO.Excel | 1448.6% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | EPPlus 4.5.3.3 | 36.10 ms | 0 B |  | OfficeIMO.Excel | 4243.3% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | EPPlus | 94.92 ms | 42.1 MB |  | OfficeIMO.Excel | 11321.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | OfficeIMO.Excel | 37.82 ms | 1.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-bottom-range | Sylvan.Data.Excel | 58.11 ms | 3.5 MB |  | OfficeIMO.Excel | 53.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | ExcelDataReader | 148.27 ms | 59.8 MB |  | OfficeIMO.Excel | 292.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | MiniExcel | 165.33 ms | 182.1 MB |  | OfficeIMO.Excel | 337.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | EPPlus | 276.62 ms | 103.1 MB |  | OfficeIMO.Excel | 631.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | ClosedXML | 407.13 ms | 145.9 MB |  | OfficeIMO.Excel | 976.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | OfficeIMO.Excel | 31.74 ms | 1.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-bottom-range-stream | Sylvan.Data.Excel | 38.23 ms | 3.5 MB |  | OfficeIMO.Excel | 20.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | ExcelDataReader | 104.82 ms | 59.8 MB |  | OfficeIMO.Excel | 230.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | MiniExcel | 116.31 ms | 182.1 MB |  | OfficeIMO.Excel | 266.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | EPPlus | 230.07 ms | 103.1 MB |  | OfficeIMO.Excel | 624.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | ClosedXML | 310.27 ms | 145.9 MB |  | OfficeIMO.Excel | 877.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | Sylvan.Data.Excel | 80.47 ms | 18.0 MB |  | Sylvan.Data.Excel | 2.4% faster than OfficeIMO |
| 25000 | speed-comparison | read-datatable | OfficeIMO.Excel | 82.41 ms | 33.8 MB |  | Sylvan.Data.Excel | Loss +2.4% |
| 25000 | speed-comparison | read-datatable | ExcelDataReader | 177.66 ms | 74.3 MB |  | Sylvan.Data.Excel | 115.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | MiniExcel | 203.27 ms | 177.0 MB |  | Sylvan.Data.Excel | 146.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | EPPlus 4.5.3.3 | 210.09 ms | 0 B |  | Sylvan.Data.Excel | 154.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | EPPlus | 382.50 ms | 197.5 MB |  | Sylvan.Data.Excel | 364.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | ClosedXML | 450.16 ms | 174.3 MB |  | Sylvan.Data.Excel | 446.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | OfficeIMO.Excel | 41.76 ms | 3.9 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | Sylvan.Data.Excel | 52.65 ms | 4.2 MB |  | OfficeIMO.Excel | 26.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | MiniExcel | 131.17 ms | 154.9 MB |  | OfficeIMO.Excel | 214.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | ExcelDataReader | 153.14 ms | 59.8 MB |  | OfficeIMO.Excel | 266.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | EPPlus | 279.75 ms | 112.8 MB |  | OfficeIMO.Excel | 569.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | ClosedXML | 405.28 ms | 147.4 MB |  | OfficeIMO.Excel | 870.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | Sylvan.Data.Excel | 43.41 ms | 5.7 MB |  | Sylvan.Data.Excel | 22.6% faster than OfficeIMO |
| 25000 | speed-comparison | read-objects | OfficeIMO.Excel | 56.11 ms | 23.0 MB |  | Sylvan.Data.Excel | Loss +29.3% |
| 25000 | speed-comparison | read-objects | ExcelDataReader | 116.41 ms | 62.0 MB |  | Sylvan.Data.Excel | 107.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | MiniExcel | 143.73 ms | 179.4 MB |  | Sylvan.Data.Excel | 156.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | EPPlus 4.5.3.3 | 242.79 ms | 0 B |  | Sylvan.Data.Excel | 332.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | EPPlus | 336.83 ms | 194.9 MB |  | Sylvan.Data.Excel | 500.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | ClosedXML | 361.17 ms | 161.7 MB |  | Sylvan.Data.Excel | 543.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | OfficeIMO.Excel | 57.68 ms | 22.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-objects-stream | Sylvan.Data.Excel | 86.11 ms | 5.2 MB |  | OfficeIMO.Excel | 49.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | ExcelDataReader | 152.38 ms | 61.5 MB |  | OfficeIMO.Excel | 164.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | MiniExcel | 224.89 ms | 178.9 MB |  | OfficeIMO.Excel | 289.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | EPPlus 4.5.3.3 | 244.91 ms | 0 B |  | OfficeIMO.Excel | 324.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | EPPlus | 403.29 ms | 194.7 MB |  | OfficeIMO.Excel | 599.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | ClosedXML | 470.11 ms | 161.5 MB |  | OfficeIMO.Excel | 715.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | Sylvan.Data.Excel | 62.62 ms | 3.5 MB |  | Sylvan.Data.Excel | 11.8% faster than OfficeIMO |
| 25000 | speed-comparison | read-range | OfficeIMO.Excel | 70.99 ms | 25.5 MB |  | Sylvan.Data.Excel | Loss +13.4% |
| 25000 | speed-comparison | read-range | ExcelDataReader | 155.87 ms | 59.8 MB |  | Sylvan.Data.Excel | 119.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | MiniExcel | 159.61 ms | 182.1 MB |  | Sylvan.Data.Excel | 124.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | EPPlus 4.5.3.3 | 235.32 ms | 0 B |  | Sylvan.Data.Excel | 231.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | EPPlus | 306.92 ms | 183.0 MB |  | Sylvan.Data.Excel | 332.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | ClosedXML | 402.78 ms | 159.8 MB |  | Sylvan.Data.Excel | 467.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | Sylvan.Data.Excel | 61.40 ms | 4.4 MB |  | Sylvan.Data.Excel | 17.4% faster than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | OfficeIMO.Excel | 74.37 ms | 26.1 MB |  | Sylvan.Data.Excel | Loss +21.1% |
| 25000 | speed-comparison | read-range-decimal | ExcelDataReader | 136.18 ms | 59.8 MB |  | Sylvan.Data.Excel | 83.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | MiniExcel | 162.06 ms | 182.1 MB |  | Sylvan.Data.Excel | 117.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | EPPlus | 393.26 ms | 183.0 MB |  | Sylvan.Data.Excel | 428.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | ClosedXML | 427.27 ms | 159.8 MB |  | Sylvan.Data.Excel | 474.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | Sylvan.Data.Excel | 45.29 ms | 3.5 MB |  | Sylvan.Data.Excel | 22.7% faster than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | OfficeIMO.Excel | 58.57 ms | 26.3 MB |  | Sylvan.Data.Excel | Loss +29.3% |
| 25000 | speed-comparison | read-range-stream | ExcelDataReader | 116.60 ms | 59.8 MB |  | Sylvan.Data.Excel | 99.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | MiniExcel | 121.85 ms | 182.1 MB |  | Sylvan.Data.Excel | 108.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | EPPlus 4.5.3.3 | 167.71 ms | 0 B |  | Sylvan.Data.Excel | 186.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | EPPlus | 340.04 ms | 183.0 MB |  | Sylvan.Data.Excel | 480.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | ClosedXML | 368.83 ms | 159.8 MB |  | Sylvan.Data.Excel | 529.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | Sylvan.Data.Excel | 0.51 ms | 348.5 KB |  | Sylvan.Data.Excel | 24.0% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range | OfficeIMO.Excel | 0.67 ms | 296.0 KB |  | Sylvan.Data.Excel | Loss +31.6% |
| 25000 | speed-comparison | read-top-range | MiniExcel | 0.96 ms | 869.0 KB |  | Sylvan.Data.Excel | 42.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | ExcelDataReader | 50.63 ms | 16.7 MB |  | Sylvan.Data.Excel | 7436.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | EPPlus 4.5.3.3 | 259.78 ms | 0 B |  | Sylvan.Data.Excel | 38568.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | EPPlus | 273.27 ms | 103.1 MB |  | Sylvan.Data.Excel | 40578.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | ClosedXML | 400.35 ms | 145.9 MB |  | Sylvan.Data.Excel | 59492.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | Sylvan.Data.Excel | 0.42 ms | 348.5 KB |  | Sylvan.Data.Excel | 19.8% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | OfficeIMO.Excel | 0.53 ms | 299.3 KB |  | Sylvan.Data.Excel | Loss +24.7% |
| 25000 | speed-comparison | read-top-range-stream | MiniExcel | 0.82 ms | 869.0 KB |  | Sylvan.Data.Excel | 56.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | ExcelDataReader | 36.07 ms | 16.7 MB |  | Sylvan.Data.Excel | 6741.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | EPPlus 4.5.3.3 | 169.07 ms | 0 B |  | Sylvan.Data.Excel | 31968.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | EPPlus | 228.12 ms | 103.1 MB |  | Sylvan.Data.Excel | 43169.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | ClosedXML | 315.57 ms | 145.9 MB |  | Sylvan.Data.Excel | 59757.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | Sylvan.Data.Excel | 0.41 ms | 348.5 KB |  | Sylvan.Data.Excel | 25.0% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | OfficeIMO.Excel | 0.54 ms | 300.0 KB |  | Sylvan.Data.Excel | Loss +33.3% |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | MiniExcel | 0.83 ms | 869.0 KB |  | Sylvan.Data.Excel | 52.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | ExcelDataReader | 39.64 ms | 16.7 MB |  | Sylvan.Data.Excel | 7210.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | EPPlus | 224.36 ms | 103.1 MB |  | Sylvan.Data.Excel | 41271.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | ClosedXML | 317.56 ms | 145.9 MB |  | Sylvan.Data.Excel | 58458.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-used-range | Sylvan.Data.Excel | 66.41 ms | 3.5 MB |  | Sylvan.Data.Excel | 8.1% faster than OfficeIMO |
| 25000 | speed-comparison | read-used-range | OfficeIMO.Excel | 72.30 ms | 25.5 MB |  | Sylvan.Data.Excel | Loss +8.9% |
| 25000 | speed-comparison | read-used-range | ExcelDataReader | 148.97 ms | 59.8 MB |  | Sylvan.Data.Excel | 106.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-used-range | MiniExcel | 158.05 ms | 182.1 MB |  | Sylvan.Data.Excel | 118.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-used-range | EPPlus | 297.28 ms | 183.0 MB |  | Sylvan.Data.Excel | 311.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-used-range | ClosedXML | 450.53 ms | 159.8 MB |  | Sylvan.Data.Excel | 523.2% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-autofilter | OfficeIMO.Excel | 32.65 ms | 11.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-autofilter | EPPlus 4.5.3.3 | 233.13 ms | 0 B |  | OfficeIMO.Excel | 613.9% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-autofilter | ClosedXML | 309.25 ms | 205.7 MB |  | OfficeIMO.Excel | 847.0% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-autofilter | EPPlus | 367.66 ms | 206.9 MB |  | OfficeIMO.Excel | 1025.9% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-charts | OfficeIMO.Excel | 32.68 ms | 12.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-charts | EPPlus 4.5.3.3 | 240.13 ms | 0 B |  | OfficeIMO.Excel | 634.8% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-charts | EPPlus | 370.85 ms | 209.9 MB |  | OfficeIMO.Excel | 1034.7% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-conditional-formatting | OfficeIMO.Excel | 31.89 ms | 11.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-conditional-formatting | EPPlus 4.5.3.3 | 239.57 ms | 0 B |  | OfficeIMO.Excel | 651.3% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-conditional-formatting | ClosedXML | 300.39 ms | 205.8 MB |  | OfficeIMO.Excel | 842.1% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-conditional-formatting | EPPlus | 362.73 ms | 206.9 MB |  | OfficeIMO.Excel | 1037.6% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-data-validation | OfficeIMO.Excel | 32.47 ms | 11.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-data-validation | EPPlus 4.5.3.3 | 238.04 ms | 0 B |  | OfficeIMO.Excel | 633.2% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-data-validation | ClosedXML | 294.79 ms | 205.7 MB |  | OfficeIMO.Excel | 808.0% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-data-validation | EPPlus | 369.52 ms | 206.9 MB |  | OfficeIMO.Excel | 1038.2% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-freeze-panes | OfficeIMO.Excel | 33.62 ms | 11.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-freeze-panes | EPPlus 4.5.3.3 | 231.67 ms | 0 B |  | OfficeIMO.Excel | 589.1% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-freeze-panes | ClosedXML | 298.44 ms | 205.7 MB |  | OfficeIMO.Excel | 787.7% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-freeze-panes | EPPlus | 372.05 ms | 206.9 MB |  | OfficeIMO.Excel | 1006.6% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-pivot-table | OfficeIMO.Excel | 83.49 ms | 41.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-pivot-table | EPPlus 4.5.3.3 | 238.99 ms | 0 B |  | OfficeIMO.Excel | 186.2% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-pivot-table | EPPlus | 394.56 ms | 225.4 MB |  | OfficeIMO.Excel | 372.6% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-all-in-one | OfficeIMO.Excel | 91.44 ms | 42.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-all-in-one | EPPlus | 425.47 ms | 270.6 MB |  | OfficeIMO.Excel | 365.3% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-all-in-one | EPPlus 4.5.3.3 | 616.24 ms | 0 B |  | OfficeIMO.Excel | 573.9% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-chart-first | OfficeIMO.Excel | 89.50 ms | 42.5 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-chart-first | EPPlus | 434.47 ms | 270.6 MB |  | OfficeIMO.Excel | 385.4% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-chart-first | EPPlus 4.5.3.3 | 514.29 ms | 0 B |  | OfficeIMO.Excel | 474.6% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-core | OfficeIMO.Excel | 34.81 ms | 11.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-core | EPPlus | 400.27 ms | 249.1 MB |  | OfficeIMO.Excel | 1049.9% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-core | EPPlus 4.5.3.3 | 470.96 ms | 0 B |  | OfficeIMO.Excel | 1253.0% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-core | ClosedXML | 801.68 ms | 664.2 MB |  | OfficeIMO.Excel | 2203.2% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-extra-column | OfficeIMO.Excel | 98.59 ms | 44.5 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-extra-column | EPPlus | 473.03 ms | 295.7 MB |  | OfficeIMO.Excel | 379.8% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-extra-column | EPPlus 4.5.3.3 | 595.57 ms | 0 B |  | OfficeIMO.Excel | 504.1% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-no-autofit | OfficeIMO.Excel | 89.46 ms | 42.6 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-no-autofit | EPPlus 4.5.3.3 | 271.20 ms | 0 B |  | OfficeIMO.Excel | 203.2% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-no-autofit | EPPlus | 403.76 ms | 229.3 MB |  | OfficeIMO.Excel | 351.4% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-post-mutation | OfficeIMO.Excel | 93.32 ms | 42.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-post-mutation | EPPlus | 416.94 ms | 270.6 MB |  | OfficeIMO.Excel | 346.8% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-post-mutation | EPPlus 4.5.3.3 | 477.41 ms | 0 B |  | OfficeIMO.Excel | 411.6% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-shuffled-columns | OfficeIMO.Excel | 105.35 ms | 42.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-shuffled-columns | EPPlus | 434.19 ms | 270.6 MB |  | OfficeIMO.Excel | 312.2% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-shuffled-columns | EPPlus 4.5.3.3 | 488.51 ms | 0 B |  | OfficeIMO.Excel | 363.7% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook | OfficeIMO.Excel | 116.92 ms | 57.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | report-workbook | EPPlus | 603.69 ms | 356.2 MB |  | OfficeIMO.Excel | 416.3% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook | EPPlus 4.5.3.3 | 618.55 ms | 0 B |  | OfficeIMO.Excel | 429.1% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-core | OfficeIMO.Excel | 46.03 ms | 10.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | report-workbook-core | EPPlus | 527.81 ms | 334.8 MB |  | OfficeIMO.Excel | 1046.8% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-core | EPPlus 4.5.3.3 | 616.09 ms | 0 B |  | OfficeIMO.Excel | 1238.6% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-core | ClosedXML | 1.14 s | 952.9 MB |  | OfficeIMO.Excel | 2367.8% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-datatable | OfficeIMO.Excel | 126.00 ms | 60.5 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | report-workbook-datatable | EPPlus | 565.03 ms | 242.0 MB |  | OfficeIMO.Excel | 348.4% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-datatable | EPPlus 4.5.3.3 | 697.91 ms | 0 B |  | OfficeIMO.Excel | 453.9% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-datatable-core | OfficeIMO.Excel | 49.99 ms | 13.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | report-workbook-datatable-core | EPPlus | 542.99 ms | 220.7 MB |  | OfficeIMO.Excel | 986.1% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-datatable-core | EPPlus 4.5.3.3 | 617.24 ms | 0 B |  | OfficeIMO.Excel | 1134.6% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-datatable-core | ClosedXML | 1.09 s | 812.7 MB |  | OfficeIMO.Excel | 2083.2% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | Sylvan.Data.Excel | 15.68 ms | 1.9 MB |  | Sylvan.Data.Excel | 12.6% faster than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | OfficeIMO.Excel | 17.94 ms | 9.0 MB |  | Sylvan.Data.Excel | Loss +14.4% |
| 25000 | speed-comparison | shared-string-read | ExcelDataReader | 41.52 ms | 24.4 MB |  | Sylvan.Data.Excel | 131.5% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | MiniExcel | 46.74 ms | 72.7 MB |  | Sylvan.Data.Excel | 160.6% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | EPPlus 4.5.3.3 | 83.96 ms | 0 B |  | Sylvan.Data.Excel | 368.1% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | EPPlus | 146.12 ms | 87.3 MB |  | Sylvan.Data.Excel | 714.6% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | ClosedXML | 151.07 ms | 88.3 MB |  | Sylvan.Data.Excel | 742.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-blog-2023-20-string-columns | LargeXlsx | 51.12 ms | 10.5 MB |  | LargeXlsx | 10.7% faster than OfficeIMO |
| 25000 | speed-comparison | write-blog-2023-20-string-columns | OfficeIMO.Excel | 57.24 ms | 11.4 MB |  | LargeXlsx | Loss +12.0% |
| 25000 | speed-comparison | write-blog-2023-20-string-columns | MiniExcel | 191.91 ms | 221.6 MB |  | LargeXlsx | 235.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-blog-2023-20-string-columns | ClosedXML | 1.09 s | 742.0 MB |  | LargeXlsx | 1806.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | OfficeIMO.Excel | 43.87 ms | 11.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-bulk-report | MiniExcel | 86.28 ms | 122.6 MB |  | OfficeIMO.Excel | 96.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | EPPlus | 461.26 ms | 249.0 MB |  | OfficeIMO.Excel | 951.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | EPPlus 4.5.3.3 | 515.73 ms | 0 B |  | OfficeIMO.Excel | 1075.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | ClosedXML | 939.50 ms | 552.7 MB |  | OfficeIMO.Excel | 2041.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellformula | OfficeIMO.Excel | 28.29 ms | 9.9 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellformula | EPPlus 4.5.3.3 | 134.02 ms | 0 B |  | OfficeIMO.Excel | 373.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellformula | ClosedXML | 199.57 ms | 111.2 MB |  | OfficeIMO.Excel | 605.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellformula | EPPlus | 366.58 ms | 137.4 MB |  | OfficeIMO.Excel | 1195.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-empty-strings | OfficeIMO.Excel | 18.17 ms | 6.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-empty-strings | ClosedXML | 150.23 ms | 90.7 MB |  | OfficeIMO.Excel | 726.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-empty-strings | EPPlus | 190.35 ms | 72.7 MB |  | OfficeIMO.Excel | 947.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-numbers | OfficeIMO.Excel | 21.13 ms | 5.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-numbers | EPPlus 4.5.3.3 | 101.78 ms | 0 B |  | OfficeIMO.Excel | 381.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-numbers | ClosedXML | 126.71 ms | 82.2 MB |  | OfficeIMO.Excel | 499.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-numbers | EPPlus | 222.06 ms | 84.4 MB |  | OfficeIMO.Excel | 951.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-mixed | OfficeIMO.Excel | 26.87 ms | 8.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-object-mixed | EPPlus 4.5.3.3 | 120.69 ms | 0 B |  | OfficeIMO.Excel | 349.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-mixed | ClosedXML | 209.39 ms | 108.5 MB |  | OfficeIMO.Excel | 679.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-mixed | EPPlus | 253.86 ms | 110.6 MB |  | OfficeIMO.Excel | 844.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-sparse | OfficeIMO.Excel | 25.88 ms | 7.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-object-sparse | ClosedXML | 175.49 ms | 102.8 MB |  | OfficeIMO.Excel | 578.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-sparse | EPPlus | 243.66 ms | 103.8 MB |  | OfficeIMO.Excel | 841.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 20.35 ms | 7.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-object-sparse-batch | ClosedXML | 183.44 ms | 102.8 MB |  | OfficeIMO.Excel | 801.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-sparse-batch | EPPlus | 249.45 ms | 103.8 MB |  | OfficeIMO.Excel | 1125.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-scalars | OfficeIMO.Excel | 16.44 ms | 6.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-scalars | EPPlus 4.5.3.3 | 103.46 ms | 0 B |  | OfficeIMO.Excel | 529.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-scalars | ClosedXML | 114.72 ms | 80.7 MB |  | OfficeIMO.Excel | 597.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-scalars | EPPlus | 219.56 ms | 83.1 MB |  | OfficeIMO.Excel | 1235.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings | OfficeIMO.Excel | 23.58 ms | 15.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-strings | EPPlus 4.5.3.3 | 100.17 ms | 0 B |  | OfficeIMO.Excel | 324.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings | ClosedXML | 133.86 ms | 101.8 MB |  | OfficeIMO.Excel | 467.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings | EPPlus | 214.83 ms | 82.4 MB |  | OfficeIMO.Excel | 811.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings-distinct | OfficeIMO.Excel | 17.30 ms | 13.5 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-strings-distinct | ClosedXML | 186.98 ms | 128.4 MB |  | OfficeIMO.Excel | 980.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings-distinct | EPPlus | 249.13 ms | 95.4 MB |  | OfficeIMO.Excel | 1339.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings-repeated | OfficeIMO.Excel | 18.09 ms | 7.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-strings-repeated | ClosedXML | 123.25 ms | 82.5 MB |  | OfficeIMO.Excel | 581.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings-repeated | EPPlus | 188.32 ms | 68.4 MB |  | OfficeIMO.Excel | 941.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-temporal | OfficeIMO.Excel | 28.12 ms | 7.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-temporal | EPPlus 4.5.3.3 | 121.82 ms | 0 B |  | OfficeIMO.Excel | 333.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-temporal | ClosedXML | 189.26 ms | 87.2 MB |  | OfficeIMO.Excel | 573.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-temporal | EPPlus | 223.14 ms | 101.4 MB |  | OfficeIMO.Excel | 693.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 12.85 ms | 3.4 MB |  | LargeXlsx | 10.0% faster than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 14.28 ms | 6.8 MB |  | LargeXlsx | Loss +11.2% |
| 25000 | speed-comparison | write-cellvalues-headerless-rectangle-direct | ClosedXML | 146.01 ms | 93.8 MB |  | LargeXlsx | 922.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-headerless-rectangle-direct | EPPlus | 220.48 ms | 85.4 MB |  | LargeXlsx | 1443.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | LargeXlsx | 35.54 ms | 5.5 MB |  | LargeXlsx | 10.2% faster than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 39.59 ms | 15.7 MB |  | LargeXlsx | Loss +11.4% |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | MiniExcel | 85.24 ms | 91.1 MB |  | LargeXlsx | 115.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 247.08 ms | 0 B |  | LargeXlsx | 524.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | ClosedXML | 369.50 ms | 205.7 MB |  | LargeXlsx | 833.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus | 443.19 ms | 206.9 MB |  | LargeXlsx | 1019.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-sparse-rectangle-direct | OfficeIMO.Excel | 20.83 ms | 7.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalues-sparse-rectangle-direct | ClosedXML | 178.41 ms | 102.8 MB |  | OfficeIMO.Excel | 756.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-sparse-rectangle-direct | EPPlus | 250.41 ms | 103.8 MB |  | OfficeIMO.Excel | 1102.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | Sylvan.Data.Excel | 33.41 ms | 5.6 MB |  | Sylvan.Data.Excel | 19.4% faster than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | OfficeIMO.Excel | 41.47 ms | 12.7 MB |  | Sylvan.Data.Excel | Loss +24.1% |
| 25000 | speed-comparison | write-datareader-plain | LargeXlsx | 43.55 ms | 8.2 MB |  | Sylvan.Data.Excel | 5.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | MiniExcel | 87.81 ms | 90.0 MB |  | Sylvan.Data.Excel | 111.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | EPPlus 4.5.3.3 | 235.54 ms | 0 B |  | Sylvan.Data.Excel | 468.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | ClosedXML | 347.17 ms | 101.8 MB |  | Sylvan.Data.Excel | 737.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | EPPlus | 394.28 ms | 114.7 MB |  | Sylvan.Data.Excel | 850.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | OfficeIMO.Excel | 46.10 ms | 12.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datareader-table | MiniExcel | 91.11 ms | 90.0 MB |  | OfficeIMO.Excel | 97.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | EPPlus 4.5.3.3 | 227.93 ms | 0 B |  | OfficeIMO.Excel | 394.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | EPPlus | 394.89 ms | 114.7 MB |  | OfficeIMO.Excel | 756.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | ClosedXML | 480.31 ms | 169.3 MB |  | OfficeIMO.Excel | 941.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table-autofit | OfficeIMO.Excel | 48.83 ms | 12.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datareader-table-autofit | MiniExcel | 99.23 ms | 121.6 MB |  | OfficeIMO.Excel | 103.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table-autofit | EPPlus | 432.69 ms | 156.0 MB |  | OfficeIMO.Excel | 786.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table-autofit | EPPlus 4.5.3.3 | 479.10 ms | 0 B |  | OfficeIMO.Excel | 881.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table-autofit | ClosedXML | 965.39 ms | 552.9 MB |  | OfficeIMO.Excel | 1877.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-headerless-tables | OfficeIMO.Excel | 56.02 ms | 12.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-headerless-tables | MiniExcel | 132.28 ms | 94.8 MB |  | OfficeIMO.Excel | 136.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-headerless-tables | ClosedXML | 595.67 ms | 168.0 MB |  | OfficeIMO.Excel | 963.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-headerless-tables | EPPlus | 679.97 ms | 108.6 MB |  | OfficeIMO.Excel | 1113.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | OfficeIMO.Excel | 45.09 ms | 9.6 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | LargeXlsx | 51.49 ms | 9.0 MB |  | OfficeIMO.Excel | 14.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | MiniExcel | 120.11 ms | 105.6 MB |  | OfficeIMO.Excel | 166.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | EPPlus | 659.20 ms | 132.5 MB |  | OfficeIMO.Excel | 1362.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | ClosedXML | 686.28 ms | 273.8 MB |  | OfficeIMO.Excel | 1422.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-tables | OfficeIMO.Excel | 54.08 ms | 13.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-sparse-tables | MiniExcel | 129.31 ms | 105.6 MB |  | OfficeIMO.Excel | 139.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-tables | EPPlus | 703.23 ms | 132.5 MB |  | OfficeIMO.Excel | 1200.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-tables | ClosedXML | 761.84 ms | 273.8 MB |  | OfficeIMO.Excel | 1308.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | OfficeIMO.Excel | 45.61 ms | 10.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-tables | MiniExcel | 94.87 ms | 94.8 MB |  | OfficeIMO.Excel | 108.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | EPPlus 4.5.3.3 | 217.44 ms | 0 B |  | OfficeIMO.Excel | 376.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | EPPlus | 424.85 ms | 108.2 MB |  | OfficeIMO.Excel | 831.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | ClosedXML | 444.96 ms | 168.0 MB |  | OfficeIMO.Excel | 875.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables-autofit | OfficeIMO.Excel | 47.84 ms | 10.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-tables-autofit | MiniExcel | 105.19 ms | 125.9 MB |  | OfficeIMO.Excel | 119.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables-autofit | EPPlus | 512.22 ms | 190.8 MB |  | OfficeIMO.Excel | 970.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables-autofit | ClosedXML | 953.09 ms | 537.2 MB |  | OfficeIMO.Excel | 1892.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | LargeXlsx | 38.39 ms | 9.3 MB |  | LargeXlsx | 3.0% faster than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | OfficeIMO.Excel | 39.58 ms | 12.4 MB |  | LargeXlsx | Loss +3.1% |
| 25000 | speed-comparison | write-datatable-direct | MiniExcel | 99.26 ms | 90.2 MB |  | LargeXlsx | 150.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | EPPlus 4.5.3.3 | 227.07 ms | 0 B |  | LargeXlsx | 473.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | ClosedXML | 348.73 ms | 101.8 MB |  | LargeXlsx | 781.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | EPPlus | 400.16 ms | 114.7 MB |  | LargeXlsx | 911.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-object-table-direct | OfficeIMO.Excel | 44.19 ms | 9.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datatable-object-table-direct | MiniExcel | 103.35 ms | 87.6 MB |  | OfficeIMO.Excel | 133.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-object-table-direct | EPPlus | 377.68 ms | 112.0 MB |  | OfficeIMO.Excel | 754.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-object-table-direct | ClosedXML | 473.81 ms | 166.7 MB |  | OfficeIMO.Excel | 972.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | OfficeIMO.Excel | 45.13 ms | 12.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datatable-table-direct | MiniExcel | 100.46 ms | 90.2 MB |  | OfficeIMO.Excel | 122.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | EPPlus 4.5.3.3 | 227.11 ms | 0 B |  | OfficeIMO.Excel | 403.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | EPPlus | 411.34 ms | 114.7 MB |  | OfficeIMO.Excel | 811.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | ClosedXML | 429.33 ms | 169.3 MB |  | OfficeIMO.Excel | 851.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-dictionary-objects-table-direct | OfficeIMO.Excel | 57.67 ms | 14.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | LargeXlsx | 34.79 ms | 5.5 MB |  | LargeXlsx | 12.3% faster than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 39.68 ms | 12.6 MB |  | LargeXlsx | Loss +14.1% |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | MiniExcel | 70.71 ms | 91.1 MB |  | LargeXlsx | 78.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 253.69 ms | 0 B |  | LargeXlsx | 539.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | ClosedXML | 323.67 ms | 101.8 MB |  | LargeXlsx | 715.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus | 388.83 ms | 114.7 MB |  | LargeXlsx | 879.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 44.97 ms | 11.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 448.91 ms | 156.0 MB |  | OfficeIMO.Excel | 898.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 851.90 ms | 485.3 MB |  | OfficeIMO.Excel | 1794.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | LargeXlsx | 33.17 ms | 5.5 MB |  | LargeXlsx | 15.4% faster than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | OfficeIMO.Excel | 39.19 ms | 11.2 MB |  | LargeXlsx | Loss +18.2% |
| 25000 | speed-comparison | write-insertobjects-direct | MiniExcel | 74.70 ms | 91.1 MB |  | LargeXlsx | 90.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | EPPlus 4.5.3.3 | 235.18 ms | 0 B |  | LargeXlsx | 500.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | ClosedXML | 347.88 ms | 101.8 MB |  | LargeXlsx | 787.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | EPPlus | 396.57 ms | 114.7 MB |  | LargeXlsx | 911.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 51.37 ms | 9.9 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 442.64 ms | 156.0 MB |  | OfficeIMO.Excel | 761.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 815.38 ms | 485.3 MB |  | OfficeIMO.Excel | 1487.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 33.47 ms | 5.5 MB |  | LargeXlsx | 29.9% faster than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 47.75 ms | 9.9 MB |  | LargeXlsx | Loss +42.6% |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | MiniExcel | 79.99 ms | 91.1 MB |  | LargeXlsx | 67.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | ClosedXML | 337.67 ms | 101.8 MB |  | LargeXlsx | 607.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | EPPlus | 378.78 ms | 114.7 MB |  | LargeXlsx | 693.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 39.87 ms | 11.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 424.75 ms | 135.1 MB |  | OfficeIMO.Excel | 965.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 507.98 ms | 269.0 MB |  | OfficeIMO.Excel | 1174.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | LargeXlsx | 50.66 ms | 5.9 MB |  | LargeXlsx | 3.6% faster than OfficeIMO |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 52.53 ms | 10.3 MB |  | LargeXlsx | Loss +3.7% |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | MiniExcel | 107.39 ms | 111.3 MB |  | LargeXlsx | 104.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | ClosedXML | 440.11 ms | 175.3 MB |  | LargeXlsx | 737.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | EPPlus | 517.02 ms | 141.5 MB |  | LargeXlsx | 884.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-mixed-direct | LargeXlsx | 45.19 ms | 5.9 MB |  | LargeXlsx | 18.6% faster than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 55.52 ms | 9.7 MB |  | LargeXlsx | Loss +22.9% |
| 25000 | speed-comparison | write-powershell-psobject-mixed-direct | MiniExcel | 104.10 ms | 111.3 MB |  | LargeXlsx | 87.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-mixed-direct | ClosedXML | 438.68 ms | 175.3 MB |  | LargeXlsx | 690.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-mixed-direct | EPPlus | 509.73 ms | 141.5 MB |  | LargeXlsx | 818.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 226.59 ms | 35.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-powershell-psobject-wide-direct | LargeXlsx | 247.61 ms | 22.7 MB |  | OfficeIMO.Excel | 9.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-wide-direct | MiniExcel | 437.78 ms | 339.8 MB |  | OfficeIMO.Excel | 93.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-wide-direct | ClosedXML | 1.46 s | 476.0 MB |  | OfficeIMO.Excel | 544.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-wide-direct | EPPlus | 1.82 s | 549.7 MB |  | OfficeIMO.Excel | 702.0% slower than OfficeIMO |
