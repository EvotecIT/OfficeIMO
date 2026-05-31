# OfficeIMO.Excel Benchmark Report

Generated: 2026-05-31T10:12:18.6622834Z
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
| 2500 | dense-helloworld-comparison | read | Other | 0 | 2 | dense-helloworld-read-stream vs Sylvan.Data.Excel (1.48x) |
| 2500 | package-profile | package | Package size | 43 | 11 | write-insertobjects-flat-dictionaries-direct vs LargeXlsx (1.62x) |
| 2500 | speed-comparison | mutate | AutoFit and mutation | 2 | 0 |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | 4 | 0 |  |
| 2500 | speed-comparison | other | Object projection | 2 | 0 |  |
| 2500 | speed-comparison | other | Range and table read | 4 | 0 |  |
| 2500 | speed-comparison | other | Real-world feature mix | 6 | 0 |  |
| 2500 | speed-comparison | other | Real-world report | 2 | 0 |  |
| 2500 | speed-comparison | other | Report workbook | 4 | 0 |  |
| 2500 | speed-comparison | read | Formula write/read | 1 | 0 |  |
| 2500 | speed-comparison | read | Other | 1 | 2 | shared-string-read vs Sylvan.Data.Excel (1.86x) |
| 2500 | speed-comparison | read | Range and table read | 4 | 3 | read-used-range vs Sylvan.Data.Excel (2.89x) |
| 2500 | speed-comparison | read | Streaming read | 1 | 3 | read-range-stream vs Sylvan.Data.Excel (1.30x) |
| 2500 | speed-comparison | read | Typed object read | 0 | 2 | read-objects vs Sylvan.Data.Excel (1.09x) |
| 2500 | speed-comparison | write | AutoFit and mutation | 5 | 0 |  |
| 2500 | speed-comparison | write | Cell writer | 7 | 0 |  |
| 2500 | speed-comparison | write | DataSet table export | 3 | 0 |  |
| 2500 | speed-comparison | write | DataTable table export | 4 | 0 |  |
| 2500 | speed-comparison | write | Formatted report write | 1 | 0 |  |
| 2500 | speed-comparison | write | Formula write/read | 1 | 0 |  |
| 2500 | speed-comparison | write | Other | 2 | 2 | write-powershell-psobject-mixed-direct vs LargeXlsx (1.31x) |
| 2500 | speed-comparison | write | Plain cell export | 1 | 3 | append-plain-rows vs LargeXlsx (1.64x) |
| 2500 | speed-comparison | write | Plain streaming export | 1 | 1 | write-datareader-plain vs Sylvan.Data.Excel (1.22x) |
| 2500 | speed-comparison | write | Plain string export | 1 | 0 |  |
| 2500 | speed-comparison | write | Shared string write | 4 | 0 |  |
| 2500 | speed-comparison | write | Typed object export | 0 | 3 | write-insertobjects-flat-dictionaries-direct vs LargeXlsx (1.44x) |
| 10000 | focused-package-profile | package | Package size | 1 | 0 |  |
| 25000 | dense-helloworld-comparison | read | Other | 0 | 2 | dense-helloworld-read-stream vs Sylvan.Data.Excel (1.22x) |
| 25000 | package-profile | package | Package size | 42 | 12 | write-insertobjects-legacy-dictionaries-direct vs LargeXlsx (1.52x) |
| 25000 | speed-comparison | mutate | AutoFit and mutation | 1 | 1 | realworld-report-no-autofit vs EPPlus 4.5.3.3 (1.15x) |
| 25000 | speed-comparison | other | Anti-cheat report variants | 4 | 0 |  |
| 25000 | speed-comparison | other | Object projection | 2 | 0 |  |
| 25000 | speed-comparison | other | Range and table read | 4 | 0 |  |
| 25000 | speed-comparison | other | Real-world feature mix | 6 | 0 |  |
| 25000 | speed-comparison | other | Real-world report | 2 | 0 |  |
| 25000 | speed-comparison | other | Report workbook | 4 | 0 |  |
| 25000 | speed-comparison | read | Formula write/read | 1 | 0 |  |
| 25000 | speed-comparison | read | Other | 2 | 1 | shared-string-read vs Sylvan.Data.Excel (1.29x) |
| 25000 | speed-comparison | read | Range and table read | 3 | 4 | read-used-range vs Sylvan.Data.Excel (1.93x) |
| 25000 | speed-comparison | read | Streaming read | 1 | 3 | read-range-stream vs Sylvan.Data.Excel (1.58x) |
| 25000 | speed-comparison | read | Typed object read | 1 | 1 | read-objects vs Sylvan.Data.Excel (1.12x) |
| 25000 | speed-comparison | write | AutoFit and mutation | 5 | 0 |  |
| 25000 | speed-comparison | write | Cell writer | 7 | 0 |  |
| 25000 | speed-comparison | write | DataSet table export | 3 | 0 |  |
| 25000 | speed-comparison | write | DataTable table export | 3 | 1 | write-datatable-direct vs LargeXlsx (1.14x) |
| 25000 | speed-comparison | write | Formatted report write | 1 | 0 |  |
| 25000 | speed-comparison | write | Formula write/read | 1 | 0 |  |
| 25000 | speed-comparison | write | Other | 2 | 2 | write-powershell-psobject-mixed-direct vs LargeXlsx (1.13x) |
| 25000 | speed-comparison | write | Plain cell export | 2 | 2 | append-plain-rows vs LargeXlsx (1.29x) |
| 25000 | speed-comparison | write | Plain streaming export | 1 | 1 | write-datareader-plain vs Sylvan.Data.Excel (1.26x) |
| 25000 | speed-comparison | write | Plain string export | 0 | 1 | write-blog-2023-20-string-columns vs LargeXlsx (1.11x) |
| 25000 | speed-comparison | write | Shared string write | 4 | 0 |  |
| 25000 | speed-comparison | write | Typed object export | 0 | 3 | write-insertobjects-flat-dictionaries-direct vs LargeXlsx (1.36x) |
| 300000 | focused-package-profile | package | Package size | 0 | 1 | write-blog-2023-20-string-columns vs LargeXlsx (1.10x) |

## Rows

| Rows | Artifact | Scenario | Library | Mean | Allocated | Package | Best | Outcome |
| ---: | --- | --- | --- | ---: | ---: | ---: | --- | --- |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | Sylvan.Data.Excel | 3.90 ms | 362.3 KB |  | Sylvan.Data.Excel | 29.4% faster than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | OfficeIMO.Excel | 5.52 ms | 2.4 MB |  | Sylvan.Data.Excel | Loss +41.6% |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | ExcelDataReader | 11.23 ms | 6.7 MB |  | Sylvan.Data.Excel | 103.6% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | MiniExcel | 15.89 ms | 21.0 MB |  | Sylvan.Data.Excel | 188.0% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | Sylvan.Data.Excel | 3.78 ms | 362.3 KB |  | Sylvan.Data.Excel | 32.3% faster than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | OfficeIMO.Excel | 5.58 ms | 2.4 MB |  | Sylvan.Data.Excel | Loss +47.7% |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | ExcelDataReader | 11.02 ms | 6.7 MB |  | Sylvan.Data.Excel | 97.5% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | MiniExcel | 15.79 ms | 21.0 MB |  | Sylvan.Data.Excel | 183.1% slower than OfficeIMO |
| 2500 | package-profile | append-plain-rows | LargeXlsx | 1.32 ms | 296.4 KB | 63.1 KB | LargeXlsx | 31.7% faster than OfficeIMO |
| 2500 | package-profile | append-plain-rows | OfficeIMO.Excel | 1.94 ms | 1.5 MB | 63.0 KB | LargeXlsx | Loss +46.5% |
| 2500 | package-profile | append-plain-rows | MiniExcel | 4.13 ms | 19.2 MB | 68.1 KB | LargeXlsx | 112.7% slower than OfficeIMO |
| 2500 | package-profile | append-plain-rows | ClosedXML | 14.60 ms | 10.9 MB | 59.8 KB | LargeXlsx | 652.3% slower than OfficeIMO |
| 2500 | package-profile | append-plain-rows | EPPlus | 25.32 ms | 14.0 MB | 56.9 KB | LargeXlsx | 1205.0% slower than OfficeIMO |
| 2500 | package-profile | autofit-existing | OfficeIMO.Excel | 8.13 ms | 1.9 MB | 143.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | autofit-existing | EPPlus | 74.38 ms | 49.5 MB | 115.0 KB | OfficeIMO.Excel | 815.0% slower than OfficeIMO |
| 2500 | package-profile | autofit-existing | ClosedXML | 130.45 ms | 82.6 MB | 121.0 KB | OfficeIMO.Excel | 1504.9% slower than OfficeIMO |
| 2500 | package-profile | large-shared-strings | OfficeIMO.Excel | 1.99 ms | 2.4 MB | 55.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | large-shared-strings | MiniExcel | 3.97 ms | 20.6 MB | 60.7 KB | OfficeIMO.Excel | 100.0% slower than OfficeIMO |
| 2500 | package-profile | large-shared-strings | ClosedXML | 12.15 ms | 11.0 MB | 50.3 KB | OfficeIMO.Excel | 511.4% slower than OfficeIMO |
| 2500 | package-profile | large-shared-strings | EPPlus | 22.24 ms | 12.5 MB | 48.1 KB | OfficeIMO.Excel | 1019.6% slower than OfficeIMO |
| 2500 | package-profile | realworld-autofilter | OfficeIMO.Excel | 3.79 ms | 1.3 MB | 142.4 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-autofilter | ClosedXML | 29.97 ms | 21.7 MB | 120.2 KB | OfficeIMO.Excel | 690.3% slower than OfficeIMO |
| 2500 | package-profile | realworld-autofilter | EPPlus | 41.02 ms | 24.1 MB | 114.2 KB | OfficeIMO.Excel | 981.8% slower than OfficeIMO |
| 2500 | package-profile | realworld-charts | OfficeIMO.Excel | 4.88 ms | 1.8 MB | 147.6 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-charts | EPPlus | 42.93 ms | 26.5 MB | 117.0 KB | OfficeIMO.Excel | 779.3% slower than OfficeIMO |
| 2500 | package-profile | realworld-conditional-formatting | OfficeIMO.Excel | 3.85 ms | 1.4 MB | 142.7 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-conditional-formatting | ClosedXML | 32.25 ms | 21.8 MB | 120.3 KB | OfficeIMO.Excel | 737.2% slower than OfficeIMO |
| 2500 | package-profile | realworld-conditional-formatting | EPPlus | 43.29 ms | 24.2 MB | 114.3 KB | OfficeIMO.Excel | 1023.9% slower than OfficeIMO |
| 2500 | package-profile | realworld-data-validation | OfficeIMO.Excel | 3.83 ms | 1.3 MB | 142.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-data-validation | ClosedXML | 30.94 ms | 21.7 MB | 120.3 KB | OfficeIMO.Excel | 708.1% slower than OfficeIMO |
| 2500 | package-profile | realworld-data-validation | EPPlus | 41.18 ms | 24.1 MB | 114.2 KB | OfficeIMO.Excel | 975.7% slower than OfficeIMO |
| 2500 | package-profile | realworld-freeze-panes | OfficeIMO.Excel | 3.61 ms | 1.3 MB | 142.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-freeze-panes | ClosedXML | 33.80 ms | 21.7 MB | 120.2 KB | OfficeIMO.Excel | 835.1% slower than OfficeIMO |
| 2500 | package-profile | realworld-freeze-panes | EPPlus | 43.01 ms | 24.2 MB | 114.3 KB | OfficeIMO.Excel | 1089.8% slower than OfficeIMO |
| 2500 | package-profile | realworld-pivot-table | OfficeIMO.Excel | 15.31 ms | 14.1 MB | 200.4 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-pivot-table | EPPlus | 46.94 ms | 28.8 MB | 117.4 KB | OfficeIMO.Excel | 206.6% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-all-in-one | OfficeIMO.Excel | 16.26 ms | 14.9 MB | 206.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-all-in-one | EPPlus | 72.08 ms | 53.3 MB | 121.8 KB | OfficeIMO.Excel | 343.2% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-chart-first | OfficeIMO.Excel | 10.81 ms | 6.1 MB | 206.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-chart-first | EPPlus | 71.07 ms | 53.3 MB | 121.8 KB | OfficeIMO.Excel | 557.4% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-core | OfficeIMO.Excel | 4.23 ms | 1.5 MB | 143.9 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-core | EPPlus | 66.83 ms | 46.2 MB | 115.6 KB | OfficeIMO.Excel | 1478.9% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-core | ClosedXML | 81.88 ms | 68.2 MB | 121.5 KB | OfficeIMO.Excel | 1834.5% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-extra-column | OfficeIMO.Excel | 17.23 ms | 16.0 MB | 219.1 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-extra-column | EPPlus | 78.31 ms | 57.8 MB | 128.4 KB | OfficeIMO.Excel | 354.6% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-no-autofit | OfficeIMO.Excel | 16.02 ms | 14.9 MB | 206.4 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-no-autofit | EPPlus | 50.30 ms | 32.1 MB | 121.8 KB | OfficeIMO.Excel | 214.0% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-post-mutation | OfficeIMO.Excel | 18.81 ms | 14.9 MB | 206.6 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-post-mutation | EPPlus | 79.51 ms | 53.3 MB | 121.9 KB | OfficeIMO.Excel | 322.7% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-shuffled-columns | OfficeIMO.Excel | 17.37 ms | 14.9 MB | 211.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-shuffled-columns | EPPlus | 69.76 ms | 53.3 MB | 124.3 KB | OfficeIMO.Excel | 301.7% slower than OfficeIMO |
| 2500 | package-profile | report-workbook | OfficeIMO.Excel | 22.48 ms | 18.7 MB | 275.6 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | report-workbook | EPPlus | 90.39 ms | 75.7 MB | 161.8 KB | OfficeIMO.Excel | 302.1% slower than OfficeIMO |
| 2500 | package-profile | report-workbook-core | OfficeIMO.Excel | 6.19 ms | 2.6 MB | 187.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | report-workbook-core | EPPlus | 96.47 ms | 70.3 MB | 157.2 KB | OfficeIMO.Excel | 1459.6% slower than OfficeIMO |
| 2500 | package-profile | report-workbook-core | ClosedXML | 102.22 ms | 94.9 MB | 165.1 KB | OfficeIMO.Excel | 1552.6% slower than OfficeIMO |
| 2500 | package-profile | report-workbook-datatable | OfficeIMO.Excel | 22.24 ms | 18.9 MB | 275.6 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | report-workbook-datatable | EPPlus | 95.47 ms | 64.4 MB | 161.8 KB | OfficeIMO.Excel | 329.3% slower than OfficeIMO |
| 2500 | package-profile | report-workbook-datatable-core | OfficeIMO.Excel | 5.77 ms | 2.9 MB | 187.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | report-workbook-datatable-core | EPPlus | 89.29 ms | 59.1 MB | 157.2 KB | OfficeIMO.Excel | 1446.7% slower than OfficeIMO |
| 2500 | package-profile | report-workbook-datatable-core | ClosedXML | 96.65 ms | 80.9 MB | 165.1 KB | OfficeIMO.Excel | 1574.3% slower than OfficeIMO |
| 2500 | package-profile | write-blog-2023-20-string-columns | OfficeIMO.Excel | 4.77 ms | 1.6 MB | 216.7 KB | OfficeIMO.Excel, LargeXlsx | Win |
| 2500 | package-profile | write-blog-2023-20-string-columns | LargeXlsx | 4.83 ms | 857.6 KB | 237.7 KB | OfficeIMO.Excel, LargeXlsx | Tie vs OfficeIMO |
| 2500 | package-profile | write-blog-2023-20-string-columns | MiniExcel | 18.27 ms | 35.1 MB | 235.3 KB | OfficeIMO.Excel, LargeXlsx | 283.4% slower than OfficeIMO |
| 2500 | package-profile | write-blog-2023-20-string-columns | ClosedXML | 91.84 ms | 69.8 MB | 257.2 KB | OfficeIMO.Excel, LargeXlsx | 1827.0% slower than OfficeIMO |
| 2500 | package-profile | write-bulk-report | OfficeIMO.Excel | 3.92 ms | 1.4 MB | 143.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-bulk-report | MiniExcel | 8.71 ms | 26.2 MB | 153.8 KB | OfficeIMO.Excel | 122.1% slower than OfficeIMO |
| 2500 | package-profile | write-bulk-report | EPPlus | 66.39 ms | 46.1 MB | 115.0 KB | OfficeIMO.Excel | 1591.8% slower than OfficeIMO |
| 2500 | package-profile | write-bulk-report | ClosedXML | 72.27 ms | 57.0 MB | 121.0 KB | OfficeIMO.Excel | 1741.8% slower than OfficeIMO |
| 2500 | package-profile | write-cellformula | OfficeIMO.Excel | 2.28 ms | 1.4 MB | 66.6 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellformula | ClosedXML | 16.98 ms | 11.8 MB | 70.6 KB | OfficeIMO.Excel | 644.7% slower than OfficeIMO |
| 2500 | package-profile | write-cellformula | EPPlus | 35.88 ms | 17.7 MB | 62.1 KB | OfficeIMO.Excel | 1473.0% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-empty-strings | OfficeIMO.Excel | 2.08 ms | 1.7 MB | 44.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-empty-strings | ClosedXML | 11.51 ms | 9.7 MB | 44.9 KB | OfficeIMO.Excel | 454.3% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-empty-strings | EPPlus | 21.40 ms | 11.5 MB | 42.0 KB | OfficeIMO.Excel | 931.0% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-numbers | OfficeIMO.Excel | 1.85 ms | 1.1 MB | 47.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-numbers | ClosedXML | 11.08 ms | 9.0 MB | 45.9 KB | OfficeIMO.Excel | 499.7% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-numbers | EPPlus | 21.82 ms | 12.6 MB | 43.7 KB | OfficeIMO.Excel | 1080.5% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-mixed | OfficeIMO.Excel | 2.34 ms | 1.7 MB | 61.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-object-mixed | ClosedXML | 18.02 ms | 11.6 MB | 59.5 KB | OfficeIMO.Excel | 671.7% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-mixed | EPPlus | 29.09 ms | 15.3 MB | 58.9 KB | OfficeIMO.Excel | 1145.5% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-sparse | OfficeIMO.Excel | 2.56 ms | 1.5 MB | 62.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-object-sparse | ClosedXML | 14.66 ms | 11.0 MB | 52.5 KB | OfficeIMO.Excel | 472.6% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-sparse | EPPlus | 26.12 ms | 14.6 MB | 54.2 KB | OfficeIMO.Excel | 920.5% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 2.41 ms | 1.5 MB | 62.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-object-sparse-batch | ClosedXML | 14.07 ms | 11.0 MB | 52.5 KB | OfficeIMO.Excel | 484.0% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-sparse-batch | EPPlus | 24.61 ms | 14.6 MB | 54.2 KB | OfficeIMO.Excel | 921.6% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-scalars | OfficeIMO.Excel | 1.77 ms | 1.1 MB | 46.3 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-scalars | ClosedXML | 12.12 ms | 8.8 MB | 45.4 KB | OfficeIMO.Excel | 584.4% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-scalars | EPPlus | 21.94 ms | 12.5 MB | 42.4 KB | OfficeIMO.Excel | 1138.7% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings | OfficeIMO.Excel | 2.45 ms | 2.6 MB | 55.1 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-strings | ClosedXML | 11.36 ms | 11.0 MB | 50.3 KB | OfficeIMO.Excel | 364.6% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings | EPPlus | 20.94 ms | 12.5 MB | 48.1 KB | OfficeIMO.Excel | 756.0% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings-distinct | OfficeIMO.Excel | 2.03 ms | 2.3 MB | 51.8 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-strings-distinct | ClosedXML | 16.87 ms | 12.8 MB | 61.9 KB | OfficeIMO.Excel | 730.2% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings-distinct | EPPlus | 25.70 ms | 13.6 MB | 61.5 KB | OfficeIMO.Excel | 1165.0% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings-repeated | OfficeIMO.Excel | 1.75 ms | 1.5 MB | 40.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-strings-repeated | ClosedXML | 10.45 ms | 9.0 MB | 38.8 KB | OfficeIMO.Excel | 497.4% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings-repeated | EPPlus | 20.73 ms | 11.1 MB | 34.8 KB | OfficeIMO.Excel | 1085.5% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-temporal | OfficeIMO.Excel | 2.71 ms | 1.4 MB | 63.3 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-temporal | ClosedXML | 15.56 ms | 9.5 MB | 54.5 KB | OfficeIMO.Excel | 474.2% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-temporal | EPPlus | 25.15 ms | 14.4 MB | 53.1 KB | OfficeIMO.Excel | 828.3% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 1.49 ms | 447.0 KB | 47.3 KB | LargeXlsx | 21.0% faster than OfficeIMO |
| 2500 | package-profile | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 1.89 ms | 1.1 MB | 48.2 KB | LargeXlsx | Loss +26.5% |
| 2500 | package-profile | write-cellvalues-headerless-rectangle-direct | ClosedXML | 14.29 ms | 10.0 MB | 53.0 KB | LargeXlsx | 656.7% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-headerless-rectangle-direct | EPPlus | 25.36 ms | 12.7 MB | 52.5 KB | LargeXlsx | 1243.5% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | LargeXlsx | 3.37 ms | 758.3 KB | 138.4 KB | LargeXlsx | 15.8% faster than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 4.00 ms | 2.0 MB | 138.0 KB | LargeXlsx | Loss +18.7% |
| 2500 | package-profile | write-cellvalues-rectangle-direct | MiniExcel | 7.87 ms | 22.7 MB | 153.7 KB | LargeXlsx | 96.8% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | ClosedXML | 33.40 ms | 21.7 MB | 120.1 KB | LargeXlsx | 734.9% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | EPPlus | 43.01 ms | 24.1 MB | 114.1 KB | LargeXlsx | 975.3% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | Sylvan.Data.Excel | 3.16 ms | 758.7 KB | 78.5 KB | Sylvan.Data.Excel | 22.8% faster than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | OfficeIMO.Excel | 4.09 ms | 1.7 MB | 138.0 KB | Sylvan.Data.Excel | Loss +29.5% |
| 2500 | package-profile | write-datareader-plain | LargeXlsx | 4.39 ms | 1.0 MB | 138.4 KB | Sylvan.Data.Excel | 7.3% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | MiniExcel | 7.73 ms | 22.5 MB | 153.6 KB | Sylvan.Data.Excel | 89.0% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | ClosedXML | 28.51 ms | 11.3 MB | 120.1 KB | Sylvan.Data.Excel | 597.0% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | EPPlus | 40.25 ms | 16.3 MB | 114.9 KB | Sylvan.Data.Excel | 884.0% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table | OfficeIMO.Excel | 4.41 ms | 1.7 MB | 138.8 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-datareader-table | MiniExcel | 7.87 ms | 22.5 MB | 153.6 KB | OfficeIMO.Excel | 78.4% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table | ClosedXML | 38.55 ms | 18.6 MB | 120.9 KB | OfficeIMO.Excel | 774.1% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table | EPPlus | 40.92 ms | 16.3 MB | 114.9 KB | OfficeIMO.Excel | 827.7% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table-autofit | OfficeIMO.Excel | 4.50 ms | 1.7 MB | 138.9 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-datareader-table-autofit | MiniExcel | 7.94 ms | 26.0 MB | 153.8 KB | OfficeIMO.Excel | 76.5% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table-autofit | EPPlus | 59.30 ms | 37.4 MB | 115.1 KB | OfficeIMO.Excel | 1217.8% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table-autofit | ClosedXML | 75.17 ms | 57.0 MB | 121.0 KB | OfficeIMO.Excel | 1570.5% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-direct-export | OfficeIMO.Excel | 3.96 ms | 2.1 MB | 131.1 KB | OfficeIMO.Excel, LargeXlsx | Win |
| 2500 | package-profile | write-dataset-sparse-direct-export | LargeXlsx | 4.00 ms | 1.1 MB | 164.2 KB | OfficeIMO.Excel, LargeXlsx | Tie vs OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-direct-export | MiniExcel | 10.08 ms | 29.0 MB | 180.5 KB | OfficeIMO.Excel, LargeXlsx | 154.5% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-direct-export | ClosedXML | 57.73 ms | 26.8 MB | 159.4 KB | OfficeIMO.Excel, LargeXlsx | 1357.0% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-direct-export | EPPlus | 57.77 ms | 21.4 MB | 144.5 KB | OfficeIMO.Excel, LargeXlsx | 1358.0% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-tables | OfficeIMO.Excel | 4.64 ms | 2.8 MB | 176.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-dataset-sparse-tables | MiniExcel | 9.45 ms | 29.0 MB | 180.5 KB | OfficeIMO.Excel | 103.9% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-tables | ClosedXML | 52.85 ms | 26.8 MB | 159.4 KB | OfficeIMO.Excel | 1040.1% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-tables | EPPlus | 60.78 ms | 21.4 MB | 144.5 KB | OfficeIMO.Excel | 1211.3% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables | OfficeIMO.Excel | 4.05 ms | 2.0 MB | 138.9 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-dataset-tables | MiniExcel | 8.28 ms | 28.0 MB | 156.4 KB | OfficeIMO.Excel | 104.5% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables | ClosedXML | 34.47 ms | 18.4 MB | 123.4 KB | OfficeIMO.Excel | 751.4% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables | EPPlus | 38.50 ms | 18.3 MB | 116.6 KB | OfficeIMO.Excel | 850.8% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables-autofit | OfficeIMO.Excel | 4.15 ms | 2.0 MB | 139.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-dataset-tables-autofit | MiniExcel | 8.49 ms | 31.1 MB | 156.6 KB | OfficeIMO.Excel | 104.4% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables-autofit | EPPlus | 58.03 ms | 40.5 MB | 116.9 KB | OfficeIMO.Excel | 1297.9% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables-autofit | ClosedXML | 71.73 ms | 55.4 MB | 123.7 KB | OfficeIMO.Excel | 1628.1% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | LargeXlsx | 3.59 ms | 1.1 MB | 138.4 KB | LargeXlsx | 13.6% faster than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | OfficeIMO.Excel | 4.15 ms | 1.7 MB | 138.0 KB | LargeXlsx | Loss +15.8% |
| 2500 | package-profile | write-datatable-direct | MiniExcel | 7.92 ms | 22.5 MB | 153.7 KB | LargeXlsx | 90.7% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | ClosedXML | 28.11 ms | 11.3 MB | 120.1 KB | LargeXlsx | 576.6% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | EPPlus | 37.53 ms | 16.3 MB | 114.9 KB | LargeXlsx | 803.3% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-table-direct | OfficeIMO.Excel | 4.21 ms | 1.7 MB | 138.8 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-datatable-table-direct | MiniExcel | 8.70 ms | 22.5 MB | 153.7 KB | OfficeIMO.Excel | 106.7% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-table-direct | EPPlus | 40.58 ms | 16.3 MB | 114.9 KB | OfficeIMO.Excel | 863.9% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-table-direct | ClosedXML | 43.49 ms | 18.6 MB | 120.9 KB | OfficeIMO.Excel | 933.0% slower than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | LargeXlsx | 3.18 ms | 758.3 KB | 138.4 KB | LargeXlsx | 19.3% faster than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 3.94 ms | 1.7 MB | 138.0 KB | LargeXlsx | Loss +23.9% |
| 2500 | package-profile | write-fluent-rowsfrom-direct | MiniExcel | 8.50 ms | 22.7 MB | 153.7 KB | LargeXlsx | 115.9% slower than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | ClosedXML | 28.49 ms | 11.3 MB | 120.1 KB | LargeXlsx | 623.5% slower than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | EPPlus | 38.50 ms | 16.3 MB | 114.9 KB | LargeXlsx | 877.5% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.11 ms | 1.3 MB | 142.4 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 55.35 ms | 37.4 MB | 115.1 KB | OfficeIMO.Excel | 1246.4% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 63.84 ms | 49.7 MB | 120.2 KB | OfficeIMO.Excel | 1453.1% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | LargeXlsx | 3.07 ms | 758.3 KB | 138.4 KB | LargeXlsx | 19.9% faster than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | OfficeIMO.Excel | 3.83 ms | 1.3 MB | 142.3 KB | LargeXlsx | Loss +24.9% |
| 2500 | package-profile | write-insertobjects-direct | MiniExcel | 8.43 ms | 22.7 MB | 153.7 KB | LargeXlsx | 119.8% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | ClosedXML | 29.38 ms | 11.3 MB | 120.1 KB | LargeXlsx | 666.1% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | EPPlus | 39.15 ms | 16.3 MB | 114.9 KB | LargeXlsx | 920.9% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 5.23 ms | 1.5 MB | 138.1 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 55.32 ms | 37.4 MB | 115.1 KB | OfficeIMO.Excel | 957.6% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 66.50 ms | 49.7 MB | 120.2 KB | OfficeIMO.Excel | 1171.1% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 3.05 ms | 758.3 KB | 138.4 KB | LargeXlsx | 38.4% faster than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 4.95 ms | 1.5 MB | 138.0 KB | LargeXlsx | Loss +62.5% |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | MiniExcel | 7.83 ms | 22.7 MB | 153.7 KB | LargeXlsx | 58.3% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | ClosedXML | 27.99 ms | 11.3 MB | 120.1 KB | LargeXlsx | 465.6% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | EPPlus | 37.85 ms | 16.3 MB | 114.9 KB | LargeXlsx | 665.0% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | LargeXlsx | 3.11 ms | 758.3 KB | 138.4 KB | LargeXlsx | 32.4% faster than OfficeIMO |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | OfficeIMO.Excel | 4.59 ms | 1.7 MB | 142.3 KB | LargeXlsx | Loss +47.9% |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | MiniExcel | 7.83 ms | 22.7 MB | 153.7 KB | LargeXlsx | 70.4% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | ClosedXML | 27.65 ms | 11.3 MB | 120.1 KB | LargeXlsx | 501.7% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | EPPlus | 35.97 ms | 16.3 MB | 114.9 KB | LargeXlsx | 682.8% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 3.61 ms | 1.3 MB | 142.4 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 42.02 ms | 27.9 MB | 120.2 KB | OfficeIMO.Excel | 1063.8% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 44.83 ms | 26.7 MB | 115.0 KB | OfficeIMO.Excel | 1141.7% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-mixed-objects-direct | LargeXlsx | 4.27 ms | 802.5 KB | 182.6 KB | LargeXlsx | 23.4% faster than OfficeIMO |
| 2500 | package-profile | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 5.57 ms | 2.3 MB | 183.1 KB | LargeXlsx | Loss +30.5% |
| 2500 | package-profile | write-powershell-mixed-objects-direct | MiniExcel | 8.33 ms | 24.6 MB | 194.0 KB | LargeXlsx | 49.4% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-mixed-objects-direct | ClosedXML | 35.10 ms | 16.6 MB | 161.0 KB | LargeXlsx | 529.8% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-mixed-objects-direct | EPPlus | 47.57 ms | 19.6 MB | 152.1 KB | LargeXlsx | 753.7% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-mixed-direct | LargeXlsx | 4.10 ms | 802.5 KB | 182.6 KB | LargeXlsx | 8.2% faster than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 4.47 ms | 1.5 MB | 182.4 KB | LargeXlsx | Loss +9.0% |
| 2500 | package-profile | write-powershell-psobject-mixed-direct | MiniExcel | 8.80 ms | 24.6 MB | 194.0 KB | LargeXlsx | 96.8% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-mixed-direct | ClosedXML | 36.57 ms | 16.6 MB | 161.0 KB | LargeXlsx | 717.7% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-mixed-direct | EPPlus | 49.63 ms | 19.6 MB | 152.1 KB | LargeXlsx | 1009.9% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 19.19 ms | 4.4 MB | 651.0 KB | OfficeIMO.Excel, LargeXlsx | Win |
| 2500 | package-profile | write-powershell-psobject-wide-direct | LargeXlsx | 19.48 ms | 2.7 MB | 644.6 KB | OfficeIMO.Excel, LargeXlsx | Tie vs OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-wide-direct | MiniExcel | 32.81 ms | 47.3 MB | 674.4 KB | OfficeIMO.Excel, LargeXlsx | 71.0% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-wide-direct | ClosedXML | 117.33 ms | 50.4 MB | 615.5 KB | OfficeIMO.Excel, LargeXlsx | 511.5% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-wide-direct | EPPlus | 157.85 ms | 67.5 MB | 548.9 KB | OfficeIMO.Excel, LargeXlsx | 722.6% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | LargeXlsx | 1.51 ms | 296.4 KB |  | LargeXlsx | 38.9% faster than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | OfficeIMO.Excel | 2.47 ms | 1.5 MB |  | LargeXlsx | Loss +63.6% |
| 2500 | speed-comparison | append-plain-rows | MiniExcel | 4.64 ms | 19.2 MB |  | LargeXlsx | 88.1% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | ClosedXML | 16.10 ms | 10.9 MB |  | LargeXlsx | 552.6% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | EPPlus 4.5.3.3 | 19.29 ms | 0 B |  | LargeXlsx | 681.6% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | EPPlus | 25.44 ms | 14.0 MB |  | LargeXlsx | 930.7% slower than OfficeIMO |
| 2500 | speed-comparison | autofit-existing | OfficeIMO.Excel | 8.38 ms | 1.9 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | autofit-existing | EPPlus | 76.78 ms | 49.5 MB |  | OfficeIMO.Excel | 816.8% slower than OfficeIMO |
| 2500 | speed-comparison | autofit-existing | EPPlus 4.5.3.3 | 103.66 ms | 0 B |  | OfficeIMO.Excel | 1137.7% slower than OfficeIMO |
| 2500 | speed-comparison | autofit-existing | ClosedXML | 136.52 ms | 82.7 MB |  | OfficeIMO.Excel | 1530.0% slower than OfficeIMO |
| 2500 | speed-comparison | build-object-datatable-dictionaries | OfficeIMO.Excel | 1.54 ms | 564.2 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | build-object-datatable-typed | OfficeIMO.Excel | 1.31 ms | 856.9 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-cells | OfficeIMO.Excel | 6.35 ms | 2.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-cells | EPPlus | 29.73 ms | 19.7 MB |  | OfficeIMO.Excel | 368.3% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-cells | ClosedXML | 31.81 ms | 16.6 MB |  | OfficeIMO.Excel | 400.9% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-first-column-from-wide-sheet | OfficeIMO.Excel | 3.99 ms | 523.4 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-first-column-from-wide-sheet | EPPlus | 24.83 ms | 12.8 MB |  | OfficeIMO.Excel | 521.7% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-first-column-from-wide-sheet | ClosedXML | 32.17 ms | 15.1 MB |  | OfficeIMO.Excel | 705.5% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-range | OfficeIMO.Excel | 6.68 ms | 2.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-range | EPPlus | 30.43 ms | 19.7 MB |  | OfficeIMO.Excel | 355.4% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-range | ClosedXML | 31.78 ms | 16.6 MB |  | OfficeIMO.Excel | 375.7% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-top-range | OfficeIMO.Excel | 0.64 ms | 285.3 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-top-range | EPPlus | 23.36 ms | 12.1 MB |  | OfficeIMO.Excel | 3524.8% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-top-range | ClosedXML | 29.03 ms | 15.0 MB |  | OfficeIMO.Excel | 4406.0% slower than OfficeIMO |
| 2500 | speed-comparison | formula-heavy-read | OfficeIMO.Excel | 2.11 ms | 706.6 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | formula-heavy-read | EPPlus 4.5.3.3 | 13.28 ms | 0 B |  | OfficeIMO.Excel | 530.8% slower than OfficeIMO |
| 2500 | speed-comparison | formula-heavy-read | ClosedXML | 18.24 ms | 8.1 MB |  | OfficeIMO.Excel | 766.0% slower than OfficeIMO |
| 2500 | speed-comparison | formula-heavy-read | EPPlus | 18.33 ms | 7.5 MB |  | OfficeIMO.Excel | 770.6% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | OfficeIMO.Excel | 2.00 ms | 2.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | large-shared-strings | MiniExcel | 4.69 ms | 20.6 MB |  | OfficeIMO.Excel | 134.1% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | EPPlus 4.5.3.3 | 10.18 ms | 0 B |  | OfficeIMO.Excel | 407.9% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | ClosedXML | 11.30 ms | 11.0 MB |  | OfficeIMO.Excel | 463.9% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | EPPlus | 25.44 ms | 12.5 MB |  | OfficeIMO.Excel | 1169.7% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | OfficeIMO.Excel | 0.83 ms | 177.2 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | large-sparse-column-read | Sylvan.Data.Excel | 1.02 ms | 316.6 KB |  | OfficeIMO.Excel | 22.4% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | ExcelDataReader | 1.46 ms | 4.0 MB |  | OfficeIMO.Excel | 74.9% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | ClosedXML | 3.87 ms | 4.3 MB |  | OfficeIMO.Excel | 364.6% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | EPPlus 4.5.3.3 | 11.29 ms | 0 B |  | OfficeIMO.Excel | 1255.4% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | MiniExcel | 11.96 ms | 45.1 MB |  | OfficeIMO.Excel | 1335.9% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | EPPlus | 37.89 ms | 42.1 MB |  | OfficeIMO.Excel | 4450.0% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | Sylvan.Data.Excel | 0.96 ms | 316.6 KB |  | Sylvan.Data.Excel | 43.5% faster than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | ExcelDataReader | 1.47 ms | 4.0 MB |  | Sylvan.Data.Excel | 12.9% faster than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | OfficeIMO.Excel | 1.69 ms | 177.2 KB |  | Sylvan.Data.Excel | Loss +77.0% |
| 2500 | speed-comparison | large-sparse-row-read | ClosedXML | 3.73 ms | 4.3 MB |  | Sylvan.Data.Excel | 120.9% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | EPPlus 4.5.3.3 | 10.98 ms | 0 B |  | Sylvan.Data.Excel | 549.6% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | MiniExcel | 12.06 ms | 45.1 MB |  | Sylvan.Data.Excel | 613.2% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | EPPlus | 38.97 ms | 42.1 MB |  | Sylvan.Data.Excel | 2204.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | OfficeIMO.Excel | 4.49 ms | 374.5 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | read-bottom-range | Sylvan.Data.Excel | 4.65 ms | 655.2 KB |  | OfficeIMO.Excel | 3.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | ExcelDataReader | 10.62 ms | 5.9 MB |  | OfficeIMO.Excel | 136.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | MiniExcel | 13.63 ms | 18.2 MB |  | OfficeIMO.Excel | 203.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | EPPlus | 24.14 ms | 12.1 MB |  | OfficeIMO.Excel | 437.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | ClosedXML | 30.06 ms | 15.0 MB |  | OfficeIMO.Excel | 569.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | OfficeIMO.Excel | 3.84 ms | 377.7 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | read-bottom-range-stream | Sylvan.Data.Excel | 4.62 ms | 655.2 KB |  | OfficeIMO.Excel | 20.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | ExcelDataReader | 10.39 ms | 5.9 MB |  | OfficeIMO.Excel | 170.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | MiniExcel | 12.30 ms | 18.2 MB |  | OfficeIMO.Excel | 220.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | EPPlus | 24.51 ms | 12.1 MB |  | OfficeIMO.Excel | 538.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | ClosedXML | 31.13 ms | 15.0 MB |  | OfficeIMO.Excel | 711.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | Sylvan.Data.Excel | 6.02 ms | 2.2 MB |  | Sylvan.Data.Excel | 14.4% faster than OfficeIMO |
| 2500 | speed-comparison | read-datatable | OfficeIMO.Excel | 7.03 ms | 3.5 MB |  | Sylvan.Data.Excel | Loss +16.8% |
| 2500 | speed-comparison | read-datatable | MiniExcel | 13.69 ms | 17.8 MB |  | Sylvan.Data.Excel | 94.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | ExcelDataReader | 14.67 ms | 7.5 MB |  | Sylvan.Data.Excel | 108.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | ClosedXML | 32.63 ms | 17.9 MB |  | Sylvan.Data.Excel | 364.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | EPPlus | 35.24 ms | 21.2 MB |  | Sylvan.Data.Excel | 401.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | EPPlus 4.5.3.3 | 42.02 ms | 0 B |  | Sylvan.Data.Excel | 497.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | OfficeIMO.Excel | 3.98 ms | 542.9 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | Sylvan.Data.Excel | 4.44 ms | 733.5 KB |  | OfficeIMO.Excel | 11.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | MiniExcel | 10.24 ms | 15.5 MB |  | OfficeIMO.Excel | 157.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | ExcelDataReader | 10.64 ms | 5.9 MB |  | OfficeIMO.Excel | 167.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | EPPlus | 23.96 ms | 12.8 MB |  | OfficeIMO.Excel | 501.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | ClosedXML | 30.30 ms | 15.1 MB |  | OfficeIMO.Excel | 660.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | Sylvan.Data.Excel | 6.90 ms | 895.3 KB |  | Sylvan.Data.Excel | 8.6% faster than OfficeIMO |
| 2500 | speed-comparison | read-objects | OfficeIMO.Excel | 7.55 ms | 2.4 MB |  | Sylvan.Data.Excel | Loss +9.4% |
| 2500 | speed-comparison | read-objects | ExcelDataReader | 13.57 ms | 6.2 MB |  | Sylvan.Data.Excel | 79.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | MiniExcel | 13.80 ms | 18.0 MB |  | Sylvan.Data.Excel | 82.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | ClosedXML | 32.04 ms | 16.5 MB |  | Sylvan.Data.Excel | 324.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | EPPlus | 33.46 ms | 20.9 MB |  | Sylvan.Data.Excel | 343.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | EPPlus 4.5.3.3 | 35.96 ms | 0 B |  | Sylvan.Data.Excel | 376.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | Sylvan.Data.Excel | 4.55 ms | 831.0 KB |  | Sylvan.Data.Excel | 3.7% faster than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | OfficeIMO.Excel | 4.72 ms | 2.4 MB |  | Sylvan.Data.Excel | Loss +3.8% |
| 2500 | speed-comparison | read-objects-stream | ExcelDataReader | 10.37 ms | 6.1 MB |  | Sylvan.Data.Excel | 119.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | MiniExcel | 12.33 ms | 18.0 MB |  | Sylvan.Data.Excel | 161.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | ClosedXML | 29.87 ms | 16.5 MB |  | Sylvan.Data.Excel | 532.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | EPPlus | 33.92 ms | 20.8 MB |  | Sylvan.Data.Excel | 617.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | EPPlus 4.5.3.3 | 34.05 ms | 0 B |  | Sylvan.Data.Excel | 620.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | OfficeIMO.Excel | 11.45 ms | 2.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | read-range | Sylvan.Data.Excel | 11.70 ms | 655.0 KB |  | OfficeIMO.Excel | 2.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | ExcelDataReader | 21.82 ms | 5.9 MB |  | OfficeIMO.Excel | 90.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | MiniExcel | 22.84 ms | 18.2 MB |  | OfficeIMO.Excel | 99.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | EPPlus 4.5.3.3 | 38.35 ms | 0 B |  | OfficeIMO.Excel | 235.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | EPPlus | 38.88 ms | 19.7 MB |  | OfficeIMO.Excel | 239.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | ClosedXML | 80.79 ms | 16.5 MB |  | OfficeIMO.Excel | 605.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | OfficeIMO.Excel | 5.71 ms | 2.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | read-range-decimal | Sylvan.Data.Excel | 6.22 ms | 750.3 KB |  | OfficeIMO.Excel | 8.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | ExcelDataReader | 11.46 ms | 5.9 MB |  | OfficeIMO.Excel | 100.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | MiniExcel | 14.88 ms | 18.2 MB |  | OfficeIMO.Excel | 160.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | ClosedXML | 32.82 ms | 16.3 MB |  | OfficeIMO.Excel | 475.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | EPPlus | 34.62 ms | 19.7 MB |  | OfficeIMO.Excel | 506.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | Sylvan.Data.Excel | 4.96 ms | 655.2 KB |  | Sylvan.Data.Excel | 23.3% faster than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | OfficeIMO.Excel | 6.47 ms | 2.7 MB |  | Sylvan.Data.Excel | Loss +30.4% |
| 2500 | speed-comparison | read-range-stream | ExcelDataReader | 11.35 ms | 5.9 MB |  | Sylvan.Data.Excel | 75.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | MiniExcel | 12.23 ms | 18.2 MB |  | Sylvan.Data.Excel | 89.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | EPPlus | 32.63 ms | 19.7 MB |  | Sylvan.Data.Excel | 404.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | ClosedXML | 32.73 ms | 16.3 MB |  | Sylvan.Data.Excel | 405.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | EPPlus 4.5.3.3 | 36.89 ms | 0 B |  | Sylvan.Data.Excel | 470.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | Sylvan.Data.Excel | 0.55 ms | 348.5 KB |  | Sylvan.Data.Excel | 14.7% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range | OfficeIMO.Excel | 0.64 ms | 296.0 KB |  | Sylvan.Data.Excel | Loss +17.2% |
| 2500 | speed-comparison | read-top-range | MiniExcel | 0.84 ms | 869.0 KB |  | Sylvan.Data.Excel | 30.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | ExcelDataReader | 4.53 ms | 1.9 MB |  | Sylvan.Data.Excel | 607.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | EPPlus | 26.31 ms | 12.1 MB |  | Sylvan.Data.Excel | 4008.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | EPPlus 4.5.3.3 | 29.93 ms | 0 B |  | Sylvan.Data.Excel | 4574.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | ClosedXML | 36.41 ms | 15.0 MB |  | Sylvan.Data.Excel | 5586.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | Sylvan.Data.Excel | 0.46 ms | 348.5 KB |  | Sylvan.Data.Excel | 22.7% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | OfficeIMO.Excel | 0.60 ms | 299.4 KB |  | Sylvan.Data.Excel | Loss +29.4% |
| 2500 | speed-comparison | read-top-range-stream | MiniExcel | 0.72 ms | 869.0 KB |  | Sylvan.Data.Excel | 20.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | ExcelDataReader | 4.57 ms | 1.9 MB |  | Sylvan.Data.Excel | 662.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | EPPlus | 24.16 ms | 12.1 MB |  | Sylvan.Data.Excel | 3924.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | EPPlus 4.5.3.3 | 31.25 ms | 0 B |  | Sylvan.Data.Excel | 5106.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | ClosedXML | 31.69 ms | 15.0 MB |  | Sylvan.Data.Excel | 5178.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | Sylvan.Data.Excel | 0.44 ms | 348.5 KB |  | Sylvan.Data.Excel | 21.6% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | OfficeIMO.Excel | 0.56 ms | 300.2 KB |  | Sylvan.Data.Excel | Loss +27.6% |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | MiniExcel | 0.68 ms | 869.0 KB |  | Sylvan.Data.Excel | 22.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | ExcelDataReader | 4.35 ms | 1.9 MB |  | Sylvan.Data.Excel | 680.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | EPPlus | 24.60 ms | 12.1 MB |  | Sylvan.Data.Excel | 4317.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | ClosedXML | 31.07 ms | 15.0 MB |  | Sylvan.Data.Excel | 5479.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-used-range | Sylvan.Data.Excel | 4.89 ms | 655.2 KB |  | Sylvan.Data.Excel | 65.4% faster than OfficeIMO |
| 2500 | speed-comparison | read-used-range | ExcelDataReader | 11.34 ms | 5.9 MB |  | Sylvan.Data.Excel | 19.7% faster than OfficeIMO |
| 2500 | speed-comparison | read-used-range | OfficeIMO.Excel | 14.12 ms | 3.4 MB |  | Sylvan.Data.Excel | Loss +188.6% |
| 2500 | speed-comparison | read-used-range | MiniExcel | 15.15 ms | 18.2 MB |  | Sylvan.Data.Excel | 7.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-used-range | EPPlus | 33.18 ms | 19.7 MB |  | Sylvan.Data.Excel | 134.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-used-range | ClosedXML | 63.52 ms | 16.4 MB |  | Sylvan.Data.Excel | 349.7% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-autofilter | OfficeIMO.Excel | 4.46 ms | 1.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-autofilter | ClosedXML | 36.07 ms | 21.7 MB |  | OfficeIMO.Excel | 709.0% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-autofilter | EPPlus 4.5.3.3 | 36.75 ms | 0 B |  | OfficeIMO.Excel | 724.1% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-autofilter | EPPlus | 49.80 ms | 24.1 MB |  | OfficeIMO.Excel | 1016.7% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-charts | OfficeIMO.Excel | 5.07 ms | 1.8 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-charts | EPPlus 4.5.3.3 | 36.97 ms | 0 B |  | OfficeIMO.Excel | 628.5% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-charts | EPPlus | 44.00 ms | 26.5 MB |  | OfficeIMO.Excel | 767.1% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-conditional-formatting | OfficeIMO.Excel | 4.24 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-conditional-formatting | EPPlus 4.5.3.3 | 35.01 ms | 0 B |  | OfficeIMO.Excel | 725.9% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-conditional-formatting | ClosedXML | 36.02 ms | 21.8 MB |  | OfficeIMO.Excel | 749.9% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-conditional-formatting | EPPlus | 44.88 ms | 24.2 MB |  | OfficeIMO.Excel | 958.8% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-data-validation | OfficeIMO.Excel | 3.66 ms | 1.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-data-validation | ClosedXML | 32.62 ms | 21.7 MB |  | OfficeIMO.Excel | 790.2% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-data-validation | EPPlus 4.5.3.3 | 36.82 ms | 0 B |  | OfficeIMO.Excel | 904.9% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-data-validation | EPPlus | 43.28 ms | 24.1 MB |  | OfficeIMO.Excel | 1081.1% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-freeze-panes | OfficeIMO.Excel | 4.62 ms | 1.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-freeze-panes | EPPlus 4.5.3.3 | 36.27 ms | 0 B |  | OfficeIMO.Excel | 685.0% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-freeze-panes | ClosedXML | 50.73 ms | 21.7 MB |  | OfficeIMO.Excel | 997.9% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-freeze-panes | EPPlus | 53.41 ms | 24.2 MB |  | OfficeIMO.Excel | 1055.9% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-pivot-table | OfficeIMO.Excel | 14.65 ms | 14.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-pivot-table | EPPlus 4.5.3.3 | 37.98 ms | 0 B |  | OfficeIMO.Excel | 159.2% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-pivot-table | EPPlus | 46.65 ms | 28.8 MB |  | OfficeIMO.Excel | 218.4% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-all-in-one | OfficeIMO.Excel | 22.00 ms | 14.9 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-all-in-one | EPPlus 4.5.3.3 | 80.50 ms | 0 B |  | OfficeIMO.Excel | 265.9% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-all-in-one | EPPlus | 86.01 ms | 53.3 MB |  | OfficeIMO.Excel | 291.0% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-chart-first | OfficeIMO.Excel | 11.62 ms | 6.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-chart-first | EPPlus | 77.67 ms | 53.3 MB |  | OfficeIMO.Excel | 568.7% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-chart-first | EPPlus 4.5.3.3 | 83.69 ms | 0 B |  | OfficeIMO.Excel | 620.4% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-core | OfficeIMO.Excel | 5.69 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-core | EPPlus 4.5.3.3 | 73.91 ms | 0 B |  | OfficeIMO.Excel | 1199.3% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-core | EPPlus | 86.27 ms | 46.2 MB |  | OfficeIMO.Excel | 1416.5% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-core | ClosedXML | 100.27 ms | 68.2 MB |  | OfficeIMO.Excel | 1662.6% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-extra-column | OfficeIMO.Excel | 40.21 ms | 16.0 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-extra-column | EPPlus 4.5.3.3 | 88.63 ms | 0 B |  | OfficeIMO.Excel | 120.4% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-extra-column | EPPlus | 170.20 ms | 57.8 MB |  | OfficeIMO.Excel | 323.3% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-no-autofit | OfficeIMO.Excel | 19.21 ms | 14.9 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-no-autofit | EPPlus 4.5.3.3 | 43.60 ms | 0 B |  | OfficeIMO.Excel | 126.9% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-no-autofit | EPPlus | 54.19 ms | 32.1 MB |  | OfficeIMO.Excel | 182.1% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-post-mutation | OfficeIMO.Excel | 25.82 ms | 14.9 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-post-mutation | EPPlus 4.5.3.3 | 76.72 ms | 0 B |  | OfficeIMO.Excel | 197.2% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-post-mutation | EPPlus | 111.68 ms | 53.3 MB |  | OfficeIMO.Excel | 332.6% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-shuffled-columns | OfficeIMO.Excel | 18.97 ms | 14.9 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-shuffled-columns | EPPlus 4.5.3.3 | 78.88 ms | 0 B |  | OfficeIMO.Excel | 315.9% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-shuffled-columns | EPPlus | 89.02 ms | 53.3 MB |  | OfficeIMO.Excel | 369.3% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook | OfficeIMO.Excel | 42.69 ms | 18.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | report-workbook | EPPlus 4.5.3.3 | 85.84 ms | 0 B |  | OfficeIMO.Excel | 101.1% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook | EPPlus | 119.63 ms | 75.7 MB |  | OfficeIMO.Excel | 180.2% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-core | OfficeIMO.Excel | 7.48 ms | 2.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | report-workbook-core | EPPlus 4.5.3.3 | 75.29 ms | 0 B |  | OfficeIMO.Excel | 907.0% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-core | ClosedXML | 124.29 ms | 94.9 MB |  | OfficeIMO.Excel | 1562.4% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-core | EPPlus | 130.20 ms | 70.3 MB |  | OfficeIMO.Excel | 1641.3% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-datatable | OfficeIMO.Excel | 30.54 ms | 18.9 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | report-workbook-datatable | EPPlus 4.5.3.3 | 75.79 ms | 0 B |  | OfficeIMO.Excel | 148.2% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-datatable | EPPlus | 132.53 ms | 64.4 MB |  | OfficeIMO.Excel | 334.0% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-datatable-core | OfficeIMO.Excel | 7.50 ms | 2.9 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | report-workbook-datatable-core | EPPlus 4.5.3.3 | 70.54 ms | 0 B |  | OfficeIMO.Excel | 840.2% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-datatable-core | EPPlus | 125.14 ms | 59.1 MB |  | OfficeIMO.Excel | 1568.0% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-datatable-core | ClosedXML | 135.90 ms | 80.9 MB |  | OfficeIMO.Excel | 1711.3% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | Sylvan.Data.Excel | 1.95 ms | 518.6 KB |  | Sylvan.Data.Excel | 46.2% faster than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | OfficeIMO.Excel | 3.62 ms | 1.0 MB |  | Sylvan.Data.Excel | Loss +85.9% |
| 2500 | speed-comparison | shared-string-read | ExcelDataReader | 4.76 ms | 2.6 MB |  | Sylvan.Data.Excel | 31.4% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | MiniExcel | 6.07 ms | 7.4 MB |  | Sylvan.Data.Excel | 67.6% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | EPPlus 4.5.3.3 | 12.69 ms | 0 B |  | Sylvan.Data.Excel | 250.3% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | ClosedXML | 16.29 ms | 9.3 MB |  | Sylvan.Data.Excel | 349.9% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | EPPlus | 22.01 ms | 10.1 MB |  | Sylvan.Data.Excel | 507.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-blog-2023-20-string-columns | OfficeIMO.Excel | 4.76 ms | 1.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-blog-2023-20-string-columns | LargeXlsx | 5.04 ms | 857.6 KB |  | OfficeIMO.Excel | 6.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-blog-2023-20-string-columns | MiniExcel | 18.38 ms | 35.1 MB |  | OfficeIMO.Excel | 286.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-blog-2023-20-string-columns | ClosedXML | 89.58 ms | 69.8 MB |  | OfficeIMO.Excel | 1783.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | OfficeIMO.Excel | 6.61 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-bulk-report | MiniExcel | 26.19 ms | 26.2 MB |  | OfficeIMO.Excel | 296.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | EPPlus 4.5.3.3 | 114.98 ms | 0 B |  | OfficeIMO.Excel | 1639.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | EPPlus | 228.66 ms | 48.0 MB |  | OfficeIMO.Excel | 3359.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | ClosedXML | 293.37 ms | 57.0 MB |  | OfficeIMO.Excel | 4338.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellformula | OfficeIMO.Excel | 2.79 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellformula | ClosedXML | 18.59 ms | 11.8 MB |  | OfficeIMO.Excel | 566.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellformula | EPPlus 4.5.3.3 | 23.82 ms | 0 B |  | OfficeIMO.Excel | 753.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellformula | EPPlus | 40.43 ms | 17.7 MB |  | OfficeIMO.Excel | 1348.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-empty-strings | OfficeIMO.Excel | 2.31 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-empty-strings | ClosedXML | 11.96 ms | 9.7 MB |  | OfficeIMO.Excel | 418.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-empty-strings | EPPlus | 20.68 ms | 11.5 MB |  | OfficeIMO.Excel | 796.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-numbers | OfficeIMO.Excel | 2.41 ms | 1.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-numbers | ClosedXML | 11.72 ms | 9.0 MB |  | OfficeIMO.Excel | 387.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-numbers | EPPlus 4.5.3.3 | 16.63 ms | 0 B |  | OfficeIMO.Excel | 591.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-numbers | EPPlus | 22.77 ms | 12.6 MB |  | OfficeIMO.Excel | 846.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-mixed | OfficeIMO.Excel | 3.00 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-object-mixed | ClosedXML | 17.88 ms | 11.6 MB |  | OfficeIMO.Excel | 495.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-mixed | EPPlus 4.5.3.3 | 21.94 ms | 0 B |  | OfficeIMO.Excel | 630.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-mixed | EPPlus | 25.36 ms | 15.3 MB |  | OfficeIMO.Excel | 744.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-sparse | OfficeIMO.Excel | 2.91 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-object-sparse | ClosedXML | 15.85 ms | 11.0 MB |  | OfficeIMO.Excel | 445.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-sparse | EPPlus | 27.04 ms | 14.6 MB |  | OfficeIMO.Excel | 830.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 2.85 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-object-sparse-batch | ClosedXML | 14.63 ms | 11.0 MB |  | OfficeIMO.Excel | 413.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-sparse-batch | EPPlus | 26.72 ms | 14.6 MB |  | OfficeIMO.Excel | 837.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-scalars | OfficeIMO.Excel | 2.48 ms | 1.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-scalars | ClosedXML | 13.07 ms | 8.8 MB |  | OfficeIMO.Excel | 426.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-scalars | EPPlus 4.5.3.3 | 17.73 ms | 0 B |  | OfficeIMO.Excel | 614.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-scalars | EPPlus | 24.64 ms | 12.5 MB |  | OfficeIMO.Excel | 893.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings | OfficeIMO.Excel | 3.09 ms | 2.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-strings | ClosedXML | 12.66 ms | 11.0 MB |  | OfficeIMO.Excel | 309.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings | EPPlus 4.5.3.3 | 20.51 ms | 0 B |  | OfficeIMO.Excel | 563.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings | EPPlus | 22.29 ms | 12.5 MB |  | OfficeIMO.Excel | 621.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings-distinct | OfficeIMO.Excel | 2.56 ms | 2.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-strings-distinct | ClosedXML | 15.67 ms | 12.8 MB |  | OfficeIMO.Excel | 512.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings-distinct | EPPlus | 25.22 ms | 13.6 MB |  | OfficeIMO.Excel | 886.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings-repeated | OfficeIMO.Excel | 2.31 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-strings-repeated | ClosedXML | 11.31 ms | 9.0 MB |  | OfficeIMO.Excel | 388.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings-repeated | EPPlus | 21.01 ms | 11.1 MB |  | OfficeIMO.Excel | 807.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-temporal | OfficeIMO.Excel | 3.13 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-temporal | ClosedXML | 16.16 ms | 9.5 MB |  | OfficeIMO.Excel | 415.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-temporal | EPPlus 4.5.3.3 | 18.25 ms | 0 B |  | OfficeIMO.Excel | 482.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-temporal | EPPlus | 26.01 ms | 14.4 MB |  | OfficeIMO.Excel | 729.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 1.30 ms | 447.0 KB |  | LargeXlsx | 23.7% faster than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 1.70 ms | 1.1 MB |  | LargeXlsx | Loss +31.0% |
| 2500 | speed-comparison | write-cellvalues-headerless-rectangle-direct | ClosedXML | 14.13 ms | 10.0 MB |  | LargeXlsx | 731.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-headerless-rectangle-direct | EPPlus | 23.28 ms | 12.7 MB |  | LargeXlsx | 1268.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | LargeXlsx | 3.54 ms | 758.3 KB |  | LargeXlsx | 17.0% faster than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 4.27 ms | 2.0 MB |  | LargeXlsx | Loss +20.5% |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | MiniExcel | 9.58 ms | 22.7 MB |  | LargeXlsx | 124.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | ClosedXML | 35.11 ms | 21.7 MB |  | LargeXlsx | 722.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 39.94 ms | 0 B |  | LargeXlsx | 835.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus | 45.56 ms | 24.1 MB |  | LargeXlsx | 967.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-sparse-rectangle-direct | OfficeIMO.Excel | 2.71 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalues-sparse-rectangle-direct | ClosedXML | 16.69 ms | 11.0 MB |  | OfficeIMO.Excel | 516.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-sparse-rectangle-direct | EPPlus | 25.96 ms | 14.6 MB |  | OfficeIMO.Excel | 858.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | Sylvan.Data.Excel | 3.66 ms | 758.6 KB |  | Sylvan.Data.Excel | 18.1% faster than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | LargeXlsx | 3.82 ms | 1.0 MB |  | Sylvan.Data.Excel | 14.5% faster than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | OfficeIMO.Excel | 4.47 ms | 1.7 MB |  | Sylvan.Data.Excel | Loss +22.1% |
| 2500 | speed-comparison | write-datareader-plain | MiniExcel | 9.18 ms | 22.5 MB |  | Sylvan.Data.Excel | 105.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | ClosedXML | 30.40 ms | 11.3 MB |  | Sylvan.Data.Excel | 580.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | EPPlus 4.5.3.3 | 32.21 ms | 0 B |  | Sylvan.Data.Excel | 620.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | EPPlus | 41.21 ms | 16.3 MB |  | Sylvan.Data.Excel | 821.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | OfficeIMO.Excel | 4.73 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datareader-table | MiniExcel | 8.96 ms | 22.5 MB |  | OfficeIMO.Excel | 89.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | EPPlus 4.5.3.3 | 35.33 ms | 0 B |  | OfficeIMO.Excel | 647.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | EPPlus | 39.86 ms | 16.3 MB |  | OfficeIMO.Excel | 743.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | ClosedXML | 41.15 ms | 18.6 MB |  | OfficeIMO.Excel | 770.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table-autofit | OfficeIMO.Excel | 5.13 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datareader-table-autofit | MiniExcel | 9.47 ms | 26.0 MB |  | OfficeIMO.Excel | 84.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table-autofit | EPPlus | 56.55 ms | 37.4 MB |  | OfficeIMO.Excel | 1003.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table-autofit | EPPlus 4.5.3.3 | 74.15 ms | 0 B |  | OfficeIMO.Excel | 1346.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table-autofit | ClosedXML | 83.68 ms | 57.0 MB |  | OfficeIMO.Excel | 1532.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-headerless-tables | OfficeIMO.Excel | 7.40 ms | 2.0 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-headerless-tables | MiniExcel | 13.88 ms | 28.0 MB |  | OfficeIMO.Excel | 87.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-headerless-tables | ClosedXML | 55.46 ms | 18.5 MB |  | OfficeIMO.Excel | 649.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-headerless-tables | EPPlus | 68.09 ms | 17.3 MB |  | OfficeIMO.Excel | 820.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | OfficeIMO.Excel | 7.82 ms | 2.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | LargeXlsx | 11.68 ms | 1.1 MB |  | OfficeIMO.Excel | 49.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | MiniExcel | 12.17 ms | 29.0 MB |  | OfficeIMO.Excel | 55.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | ClosedXML | 69.90 ms | 26.8 MB |  | OfficeIMO.Excel | 794.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | EPPlus | 74.05 ms | 21.4 MB |  | OfficeIMO.Excel | 847.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-tables | OfficeIMO.Excel | 8.42 ms | 2.8 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-sparse-tables | MiniExcel | 14.68 ms | 29.0 MB |  | OfficeIMO.Excel | 74.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-tables | EPPlus | 80.34 ms | 21.4 MB |  | OfficeIMO.Excel | 854.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-tables | ClosedXML | 86.61 ms | 26.8 MB |  | OfficeIMO.Excel | 928.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | OfficeIMO.Excel | 6.43 ms | 2.0 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-tables | MiniExcel | 23.18 ms | 28.5 MB |  | OfficeIMO.Excel | 260.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | EPPlus 4.5.3.3 | 43.50 ms | 0 B |  | OfficeIMO.Excel | 577.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | ClosedXML | 102.12 ms | 18.4 MB |  | OfficeIMO.Excel | 1489.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | EPPlus | 108.94 ms | 19.0 MB |  | OfficeIMO.Excel | 1595.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables-autofit | OfficeIMO.Excel | 8.68 ms | 2.0 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-tables-autofit | MiniExcel | 25.40 ms | 31.4 MB |  | OfficeIMO.Excel | 192.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables-autofit | EPPlus | 123.19 ms | 41.1 MB |  | OfficeIMO.Excel | 1318.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables-autofit | ClosedXML | 341.18 ms | 55.4 MB |  | OfficeIMO.Excel | 3829.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | OfficeIMO.Excel | 5.30 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datatable-direct | LargeXlsx | 9.57 ms | 1.1 MB |  | OfficeIMO.Excel | 80.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | MiniExcel | 10.75 ms | 22.5 MB |  | OfficeIMO.Excel | 102.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | ClosedXML | 38.37 ms | 11.3 MB |  | OfficeIMO.Excel | 624.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | EPPlus 4.5.3.3 | 40.92 ms | 0 B |  | OfficeIMO.Excel | 672.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | EPPlus | 51.25 ms | 16.3 MB |  | OfficeIMO.Excel | 867.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-object-table-direct | OfficeIMO.Excel | 4.72 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datatable-object-table-direct | MiniExcel | 8.84 ms | 22.3 MB |  | OfficeIMO.Excel | 87.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-object-table-direct | ClosedXML | 39.20 ms | 18.3 MB |  | OfficeIMO.Excel | 731.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-object-table-direct | EPPlus | 39.57 ms | 16.0 MB |  | OfficeIMO.Excel | 739.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | OfficeIMO.Excel | 4.99 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datatable-table-direct | MiniExcel | 8.24 ms | 22.5 MB |  | OfficeIMO.Excel | 65.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | ClosedXML | 41.21 ms | 18.6 MB |  | OfficeIMO.Excel | 725.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | EPPlus 4.5.3.3 | 41.96 ms | 0 B |  | OfficeIMO.Excel | 740.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | EPPlus | 42.63 ms | 16.3 MB |  | OfficeIMO.Excel | 753.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-dictionary-objects-table-direct | OfficeIMO.Excel | 6.84 ms | 2.0 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | LargeXlsx | 3.22 ms | 758.3 KB |  | LargeXlsx | 19.7% faster than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 4.00 ms | 1.7 MB |  | LargeXlsx | Loss +24.5% |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | MiniExcel | 9.83 ms | 22.7 MB |  | LargeXlsx | 145.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | ClosedXML | 30.57 ms | 11.3 MB |  | LargeXlsx | 663.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 35.34 ms | 0 B |  | LargeXlsx | 782.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus | 43.05 ms | 16.3 MB |  | LargeXlsx | 975.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.03 ms | 1.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 54.80 ms | 37.4 MB |  | OfficeIMO.Excel | 1261.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 63.73 ms | 49.7 MB |  | OfficeIMO.Excel | 1483.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | LargeXlsx | 3.14 ms | 758.3 KB |  | LargeXlsx | 11.9% faster than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | OfficeIMO.Excel | 3.56 ms | 1.3 MB |  | LargeXlsx | Loss +13.5% |
| 2500 | speed-comparison | write-insertobjects-direct | MiniExcel | 8.48 ms | 22.7 MB |  | LargeXlsx | 137.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | ClosedXML | 27.41 ms | 11.3 MB |  | LargeXlsx | 669.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | EPPlus 4.5.3.3 | 36.04 ms | 0 B |  | LargeXlsx | 911.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | EPPlus | 41.99 ms | 16.3 MB |  | LargeXlsx | 1078.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 5.32 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 60.04 ms | 37.4 MB |  | OfficeIMO.Excel | 1027.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 66.04 ms | 49.7 MB |  | OfficeIMO.Excel | 1140.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 3.06 ms | 758.3 KB |  | LargeXlsx | 30.3% faster than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 4.39 ms | 1.5 MB |  | LargeXlsx | Loss +43.6% |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | MiniExcel | 9.37 ms | 22.7 MB |  | LargeXlsx | 113.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | ClosedXML | 27.12 ms | 11.3 MB |  | LargeXlsx | 517.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | EPPlus | 42.16 ms | 16.3 MB |  | LargeXlsx | 860.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 3.98 ms | 1.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 45.30 ms | 27.9 MB |  | OfficeIMO.Excel | 1039.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 53.16 ms | 26.7 MB |  | OfficeIMO.Excel | 1236.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | LargeXlsx | 4.37 ms | 802.5 KB |  | LargeXlsx | 22.4% faster than OfficeIMO |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 5.62 ms | 2.3 MB |  | LargeXlsx | Loss +28.8% |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | MiniExcel | 8.92 ms | 24.6 MB |  | LargeXlsx | 58.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | ClosedXML | 39.17 ms | 16.6 MB |  | LargeXlsx | 596.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | EPPlus | 51.84 ms | 19.6 MB |  | LargeXlsx | 821.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-mixed-direct | LargeXlsx | 4.27 ms | 802.5 KB |  | LargeXlsx | 23.7% faster than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 5.59 ms | 1.5 MB |  | LargeXlsx | Loss +31.1% |
| 2500 | speed-comparison | write-powershell-psobject-mixed-direct | MiniExcel | 8.44 ms | 24.6 MB |  | LargeXlsx | 51.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-mixed-direct | ClosedXML | 39.06 ms | 16.6 MB |  | LargeXlsx | 598.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-mixed-direct | EPPlus | 52.27 ms | 19.6 MB |  | LargeXlsx | 834.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-wide-direct | LargeXlsx | 20.83 ms | 2.7 MB |  | LargeXlsx, OfficeIMO.Excel | Tie vs OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 20.86 ms | 4.4 MB |  | LargeXlsx, OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-powershell-psobject-wide-direct | MiniExcel | 35.09 ms | 47.3 MB |  | LargeXlsx, OfficeIMO.Excel | 68.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-wide-direct | ClosedXML | 120.21 ms | 50.4 MB |  | LargeXlsx, OfficeIMO.Excel | 476.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-wide-direct | EPPlus | 159.19 ms | 67.5 MB |  | LargeXlsx, OfficeIMO.Excel | 663.3% slower than OfficeIMO |
| 10000 | focused-package-profile | write-blog-2023-20-string-columns | OfficeIMO.Excel | 33.01 ms | 7.6 MB | 880.4 KB | OfficeIMO.Excel | Win |
| 10000 | focused-package-profile | write-blog-2023-20-string-columns | LargeXlsx | 88.91 ms | 3.1 MB | 970.2 KB | OfficeIMO.Excel | 2.69x vs best |
| 10000 | focused-package-profile | write-blog-2023-20-string-columns | MiniExcel | 142.47 ms | 96.2 MB | 957.6 KB | OfficeIMO.Excel | 4.32x vs best |
| 10000 | focused-package-profile | write-blog-2023-20-string-columns | ClosedXML | 725.30 ms | 280.2 MB | 1,015.4 KB | OfficeIMO.Excel | 21.97x vs best |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | Sylvan.Data.Excel | 38.38 ms | 394.1 KB |  | Sylvan.Data.Excel | 15.6% faster than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | OfficeIMO.Excel | 45.46 ms | 23.1 MB |  | Sylvan.Data.Excel | Loss +18.5% |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | ExcelDataReader | 114.89 ms | 67.9 MB |  | Sylvan.Data.Excel | 152.7% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | MiniExcel | 145.24 ms | 210.3 MB |  | Sylvan.Data.Excel | 219.5% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | Sylvan.Data.Excel | 36.76 ms | 394.1 KB |  | Sylvan.Data.Excel | 18.1% faster than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | OfficeIMO.Excel | 44.90 ms | 23.8 MB |  | Sylvan.Data.Excel | Loss +22.2% |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | ExcelDataReader | 115.71 ms | 67.9 MB |  | Sylvan.Data.Excel | 157.7% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | MiniExcel | 147.87 ms | 210.3 MB |  | Sylvan.Data.Excel | 229.3% slower than OfficeIMO |
| 25000 | package-profile | append-plain-rows | LargeXlsx | 10.99 ms | 2.7 MB | 605.0 KB | LargeXlsx | 24.6% faster than OfficeIMO |
| 25000 | package-profile | append-plain-rows | OfficeIMO.Excel | 14.59 ms | 10.6 MB | 610.4 KB | LargeXlsx | Loss +32.7% |
| 25000 | package-profile | append-plain-rows | MiniExcel | 33.66 ms | 56.9 MB | 642.3 KB | LargeXlsx | 130.8% slower than OfficeIMO |
| 25000 | package-profile | append-plain-rows | ClosedXML | 133.64 ms | 101.8 MB | 540.6 KB | LargeXlsx | 816.2% slower than OfficeIMO |
| 25000 | package-profile | append-plain-rows | EPPlus | 209.65 ms | 98.0 MB | 525.6 KB | LargeXlsx | 1337.4% slower than OfficeIMO |
| 25000 | package-profile | autofit-existing | OfficeIMO.Excel | 79.41 ms | 15.3 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | autofit-existing | EPPlus | 491.40 ms | 245.1 MB | 1.1 MB | OfficeIMO.Excel | 518.8% slower than OfficeIMO |
| 25000 | package-profile | autofit-existing | ClosedXML | 1.39 s | 810.3 MB | 1.1 MB | OfficeIMO.Excel | 1645.7% slower than OfficeIMO |
| 25000 | package-profile | large-shared-strings | OfficeIMO.Excel | 14.86 ms | 15.4 MB | 529.7 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | large-shared-strings | MiniExcel | 29.02 ms | 72.0 MB | 581.0 KB | OfficeIMO.Excel | 95.3% slower than OfficeIMO |
| 25000 | package-profile | large-shared-strings | ClosedXML | 110.59 ms | 101.8 MB | 460.1 KB | OfficeIMO.Excel | 644.4% slower than OfficeIMO |
| 25000 | package-profile | large-shared-strings | EPPlus | 180.83 ms | 82.4 MB | 444.7 KB | OfficeIMO.Excel | 1117.2% slower than OfficeIMO |
| 25000 | package-profile | realworld-autofilter | OfficeIMO.Excel | 37.72 ms | 11.2 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-autofilter | ClosedXML | 362.90 ms | 205.7 MB | 1.1 MB | OfficeIMO.Excel | 862.1% slower than OfficeIMO |
| 25000 | package-profile | realworld-autofilter | EPPlus | 435.34 ms | 206.9 MB | 1.1 MB | OfficeIMO.Excel | 1054.1% slower than OfficeIMO |
| 25000 | package-profile | realworld-charts | OfficeIMO.Excel | 33.59 ms | 12.3 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-charts | EPPlus | 350.87 ms | 209.9 MB | 1.1 MB | OfficeIMO.Excel | 944.5% slower than OfficeIMO |
| 25000 | package-profile | realworld-conditional-formatting | OfficeIMO.Excel | 34.63 ms | 11.3 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-conditional-formatting | ClosedXML | 313.89 ms | 205.8 MB | 1.1 MB | OfficeIMO.Excel | 806.3% slower than OfficeIMO |
| 25000 | package-profile | realworld-conditional-formatting | EPPlus | 384.78 ms | 206.9 MB | 1.1 MB | OfficeIMO.Excel | 1011.0% slower than OfficeIMO |
| 25000 | package-profile | realworld-data-validation | OfficeIMO.Excel | 31.39 ms | 11.2 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-data-validation | ClosedXML | 304.58 ms | 205.7 MB | 1.1 MB | OfficeIMO.Excel | 870.4% slower than OfficeIMO |
| 25000 | package-profile | realworld-data-validation | EPPlus | 369.82 ms | 206.9 MB | 1.1 MB | OfficeIMO.Excel | 1078.3% slower than OfficeIMO |
| 25000 | package-profile | realworld-freeze-panes | OfficeIMO.Excel | 32.83 ms | 11.2 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-freeze-panes | ClosedXML | 309.57 ms | 205.7 MB | 1.1 MB | OfficeIMO.Excel | 843.1% slower than OfficeIMO |
| 25000 | package-profile | realworld-freeze-panes | EPPlus | 387.10 ms | 206.9 MB | 1.1 MB | OfficeIMO.Excel | 1079.3% slower than OfficeIMO |
| 25000 | package-profile | realworld-pivot-table | OfficeIMO.Excel | 226.20 ms | 128.8 MB | 1.9 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-pivot-table | EPPlus | 389.18 ms | 225.4 MB | 1.1 MB | OfficeIMO.Excel | 72.0% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-all-in-one | OfficeIMO.Excel | 301.25 ms | 130.3 MB | 1.9 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-all-in-one | EPPlus | 601.05 ms | 270.6 MB | 1.1 MB | OfficeIMO.Excel | 99.5% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-chart-first | OfficeIMO.Excel | 161.07 ms | 42.5 MB | 1.9 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-chart-first | EPPlus | 788.97 ms | 270.6 MB | 1.1 MB | OfficeIMO.Excel | 389.8% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-core | OfficeIMO.Excel | 35.75 ms | 11.4 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-core | EPPlus | 389.91 ms | 249.1 MB | 1.1 MB | OfficeIMO.Excel | 990.7% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-core | ClosedXML | 771.95 ms | 664.2 MB | 1.1 MB | OfficeIMO.Excel | 2059.5% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-extra-column | OfficeIMO.Excel | 352.98 ms | 141.4 MB | 2.1 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-extra-column | EPPlus | 527.93 ms | 295.7 MB | 1.1 MB | OfficeIMO.Excel | 49.6% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-no-autofit | OfficeIMO.Excel | 417.93 ms | 130.3 MB | 1.9 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-no-autofit | EPPlus | 559.86 ms | 229.3 MB | 1.1 MB | OfficeIMO.Excel | 34.0% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-post-mutation | OfficeIMO.Excel | 243.05 ms | 130.3 MB | 1.9 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-post-mutation | EPPlus | 429.98 ms | 270.6 MB | 1.1 MB | OfficeIMO.Excel | 76.9% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-shuffled-columns | OfficeIMO.Excel | 347.07 ms | 130.4 MB | 2.0 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-shuffled-columns | EPPlus | 507.12 ms | 270.6 MB | 1.1 MB | OfficeIMO.Excel | 46.1% slower than OfficeIMO |
| 25000 | package-profile | report-workbook | OfficeIMO.Excel | 327.22 ms | 171.1 MB | 2.6 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | report-workbook | EPPlus | 560.40 ms | 356.2 MB | 1.5 MB | OfficeIMO.Excel | 71.3% slower than OfficeIMO |
| 25000 | package-profile | report-workbook-core | OfficeIMO.Excel | 60.71 ms | 10.7 MB | 1.8 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | report-workbook-core | EPPlus | 669.00 ms | 334.8 MB | 1.5 MB | OfficeIMO.Excel | 1002.0% slower than OfficeIMO |
| 25000 | package-profile | report-workbook-core | ClosedXML | 1.34 s | 952.9 MB | 1.5 MB | OfficeIMO.Excel | 2099.4% slower than OfficeIMO |
| 25000 | package-profile | report-workbook-datatable | OfficeIMO.Excel | 373.27 ms | 173.8 MB | 2.6 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | report-workbook-datatable | EPPlus | 558.15 ms | 242.0 MB | 1.5 MB | OfficeIMO.Excel | 49.5% slower than OfficeIMO |
| 25000 | package-profile | report-workbook-datatable-core | OfficeIMO.Excel | 49.96 ms | 13.4 MB | 1.8 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | report-workbook-datatable-core | EPPlus | 508.54 ms | 220.7 MB | 1.5 MB | OfficeIMO.Excel | 917.8% slower than OfficeIMO |
| 25000 | package-profile | report-workbook-datatable-core | ClosedXML | 1.02 s | 812.7 MB | 1.5 MB | OfficeIMO.Excel | 1940.2% slower than OfficeIMO |
| 25000 | package-profile | write-blog-2023-20-string-columns | LargeXlsx | 39.19 ms | 10.5 MB | 2.4 MB | LargeXlsx | 10.3% faster than OfficeIMO |
| 25000 | package-profile | write-blog-2023-20-string-columns | OfficeIMO.Excel | 43.72 ms | 11.4 MB | 2.2 MB | LargeXlsx | Loss +11.5% |
| 25000 | package-profile | write-blog-2023-20-string-columns | MiniExcel | 147.10 ms | 221.6 MB | 2.4 MB | LargeXlsx | 236.5% slower than OfficeIMO |
| 25000 | package-profile | write-blog-2023-20-string-columns | ClosedXML | 908.18 ms | 742.0 MB | 2.5 MB | LargeXlsx | 1977.4% slower than OfficeIMO |
| 25000 | package-profile | write-bulk-report | OfficeIMO.Excel | 36.78 ms | 11.3 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-bulk-report | MiniExcel | 69.37 ms | 122.6 MB | 1.5 MB | OfficeIMO.Excel | 88.6% slower than OfficeIMO |
| 25000 | package-profile | write-bulk-report | EPPlus | 412.29 ms | 249.0 MB | 1.1 MB | OfficeIMO.Excel | 1021.0% slower than OfficeIMO |
| 25000 | package-profile | write-bulk-report | ClosedXML | 830.33 ms | 552.7 MB | 1.1 MB | OfficeIMO.Excel | 2157.7% slower than OfficeIMO |
| 25000 | package-profile | write-cellformula | OfficeIMO.Excel | 18.30 ms | 9.9 MB | 670.3 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellformula | ClosedXML | 171.37 ms | 111.2 MB | 643.2 KB | OfficeIMO.Excel | 836.3% slower than OfficeIMO |
| 25000 | package-profile | write-cellformula | EPPlus | 299.14 ms | 137.4 MB | 593.9 KB | OfficeIMO.Excel | 1534.4% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-empty-strings | OfficeIMO.Excel | 12.16 ms | 6.7 MB | 451.4 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-empty-strings | ClosedXML | 113.45 ms | 90.7 MB | 398.1 KB | OfficeIMO.Excel | 833.2% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-empty-strings | EPPlus | 176.51 ms | 72.7 MB | 390.6 KB | OfficeIMO.Excel | 1351.9% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-numbers | OfficeIMO.Excel | 14.84 ms | 5.8 MB | 462.6 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-numbers | ClosedXML | 110.39 ms | 82.2 MB | 411.4 KB | OfficeIMO.Excel | 643.9% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-numbers | EPPlus | 193.38 ms | 84.4 MB | 406.5 KB | OfficeIMO.Excel | 1203.2% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-mixed | OfficeIMO.Excel | 16.32 ms | 8.1 MB | 585.1 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-object-mixed | ClosedXML | 147.22 ms | 108.5 MB | 532.9 KB | OfficeIMO.Excel | 801.9% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-mixed | EPPlus | 206.18 ms | 110.6 MB | 544.3 KB | OfficeIMO.Excel | 1163.2% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-sparse | OfficeIMO.Excel | 18.01 ms | 7.2 MB | 607.1 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-object-sparse | ClosedXML | 134.04 ms | 102.8 MB | 468.0 KB | OfficeIMO.Excel | 644.4% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-sparse | EPPlus | 204.02 ms | 103.8 MB | 494.4 KB | OfficeIMO.Excel | 1033.1% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 18.94 ms | 7.2 MB | 607.1 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-object-sparse-batch | ClosedXML | 147.75 ms | 102.8 MB | 468.0 KB | OfficeIMO.Excel | 680.1% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-sparse-batch | EPPlus | 212.85 ms | 103.8 MB | 494.4 KB | OfficeIMO.Excel | 1023.8% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-scalars | OfficeIMO.Excel | 10.96 ms | 6.0 MB | 441.9 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-scalars | ClosedXML | 98.87 ms | 80.7 MB | 394.9 KB | OfficeIMO.Excel | 801.9% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-scalars | EPPlus | 184.68 ms | 83.1 MB | 379.3 KB | OfficeIMO.Excel | 1584.6% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings | OfficeIMO.Excel | 16.52 ms | 15.0 MB | 527.8 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-strings | ClosedXML | 115.64 ms | 101.8 MB | 460.1 KB | OfficeIMO.Excel | 599.9% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings | EPPlus | 233.76 ms | 82.4 MB | 444.7 KB | OfficeIMO.Excel | 1314.7% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings-distinct | OfficeIMO.Excel | 12.73 ms | 13.5 MB | 499.0 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-strings-distinct | ClosedXML | 151.01 ms | 128.4 MB | 555.3 KB | OfficeIMO.Excel | 1086.2% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings-distinct | EPPlus | 209.45 ms | 95.4 MB | 565.1 KB | OfficeIMO.Excel | 1545.2% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings-repeated | OfficeIMO.Excel | 11.60 ms | 7.3 MB | 376.0 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-strings-repeated | ClosedXML | 97.75 ms | 82.5 MB | 331.8 KB | OfficeIMO.Excel | 742.8% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings-repeated | EPPlus | 157.09 ms | 68.4 MB | 300.8 KB | OfficeIMO.Excel | 1254.5% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-temporal | OfficeIMO.Excel | 18.80 ms | 7.3 MB | 620.5 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-temporal | ClosedXML | 140.91 ms | 87.2 MB | 483.0 KB | OfficeIMO.Excel | 649.4% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-temporal | EPPlus | 193.51 ms | 101.4 MB | 495.1 KB | OfficeIMO.Excel | 929.1% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 10.16 ms | 3.4 MB | 443.4 KB | LargeXlsx | 3.2% faster than OfficeIMO |
| 25000 | package-profile | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 10.50 ms | 6.8 MB | 455.5 KB | LargeXlsx | Loss +3.3% |
| 25000 | package-profile | write-cellvalues-headerless-rectangle-direct | ClosedXML | 126.65 ms | 93.8 MB | 467.5 KB | LargeXlsx | 1106.4% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-headerless-rectangle-direct | EPPlus | 197.12 ms | 85.4 MB | 484.1 KB | LargeXlsx | 1777.6% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | LargeXlsx | 31.14 ms | 5.5 MB | 1.4 MB | LargeXlsx | 13.1% faster than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 35.85 ms | 15.7 MB | 1.4 MB | LargeXlsx | Loss +15.1% |
| 25000 | package-profile | write-cellvalues-rectangle-direct | MiniExcel | 68.64 ms | 91.1 MB | 1.5 MB | LargeXlsx | 91.5% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | ClosedXML | 322.56 ms | 205.7 MB | 1.1 MB | LargeXlsx | 799.8% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | EPPlus | 447.57 ms | 206.9 MB | 1.1 MB | LargeXlsx | 1148.5% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | Sylvan.Data.Excel | 29.75 ms | 5.6 MB | 755.4 KB | Sylvan.Data.Excel | 25.9% faster than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | LargeXlsx | 34.57 ms | 8.2 MB | 1.4 MB | Sylvan.Data.Excel | 13.9% faster than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | OfficeIMO.Excel | 40.13 ms | 12.7 MB | 1.4 MB | Sylvan.Data.Excel | Loss +34.9% |
| 25000 | package-profile | write-datareader-plain | MiniExcel | 76.72 ms | 90.0 MB | 1.5 MB | Sylvan.Data.Excel | 91.2% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | ClosedXML | 293.14 ms | 101.8 MB | 1.1 MB | Sylvan.Data.Excel | 630.5% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | EPPlus | 360.44 ms | 114.7 MB | 1.1 MB | Sylvan.Data.Excel | 798.2% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table | OfficeIMO.Excel | 38.33 ms | 12.7 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-datareader-table | MiniExcel | 71.65 ms | 90.0 MB | 1.5 MB | OfficeIMO.Excel | 86.9% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table | EPPlus | 348.87 ms | 114.7 MB | 1.1 MB | OfficeIMO.Excel | 810.1% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table | ClosedXML | 397.09 ms | 169.3 MB | 1.1 MB | OfficeIMO.Excel | 935.9% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table-autofit | OfficeIMO.Excel | 40.84 ms | 12.7 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-datareader-table-autofit | MiniExcel | 74.04 ms | 121.6 MB | 1.5 MB | OfficeIMO.Excel | 81.3% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table-autofit | EPPlus | 373.41 ms | 156.0 MB | 1.1 MB | OfficeIMO.Excel | 814.3% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table-autofit | ClosedXML | 790.87 ms | 552.9 MB | 1.1 MB | OfficeIMO.Excel | 1836.4% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-direct-export | OfficeIMO.Excel | 34.06 ms | 9.6 MB | 1.3 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-dataset-sparse-direct-export | LargeXlsx | 40.18 ms | 9.0 MB | 1.6 MB | OfficeIMO.Excel | 18.0% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-direct-export | MiniExcel | 103.24 ms | 105.6 MB | 1.8 MB | OfficeIMO.Excel | 203.1% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-direct-export | EPPlus | 533.62 ms | 132.5 MB | 1.4 MB | OfficeIMO.Excel | 1466.7% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-direct-export | ClosedXML | 599.66 ms | 273.8 MB | 1.5 MB | OfficeIMO.Excel | 1660.6% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-tables | OfficeIMO.Excel | 63.10 ms | 13.1 MB | 1.8 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-dataset-sparse-tables | MiniExcel | 141.93 ms | 105.6 MB | 1.8 MB | OfficeIMO.Excel | 124.9% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-tables | EPPlus | 658.67 ms | 132.5 MB | 1.4 MB | OfficeIMO.Excel | 943.8% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-tables | ClosedXML | 707.99 ms | 273.8 MB | 1.5 MB | OfficeIMO.Excel | 1021.9% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables | OfficeIMO.Excel | 36.76 ms | 10.0 MB | 1.3 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-dataset-tables | MiniExcel | 82.30 ms | 94.8 MB | 1.5 MB | OfficeIMO.Excel | 123.9% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables | EPPlus | 343.40 ms | 108.2 MB | 1.1 MB | OfficeIMO.Excel | 834.2% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables | ClosedXML | 381.65 ms | 168.0 MB | 1.1 MB | OfficeIMO.Excel | 938.3% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables-autofit | OfficeIMO.Excel | 41.77 ms | 10.1 MB | 1.3 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-dataset-tables-autofit | MiniExcel | 94.28 ms | 125.9 MB | 1.5 MB | OfficeIMO.Excel | 125.7% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables-autofit | EPPlus | 439.32 ms | 190.8 MB | 1.1 MB | OfficeIMO.Excel | 951.7% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables-autofit | ClosedXML | 835.89 ms | 537.2 MB | 1.1 MB | OfficeIMO.Excel | 1901.0% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | LargeXlsx | 32.82 ms | 9.3 MB | 1.4 MB | LargeXlsx | 7.5% faster than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | OfficeIMO.Excel | 35.46 ms | 12.4 MB | 1.4 MB | LargeXlsx | Loss +8.1% |
| 25000 | package-profile | write-datatable-direct | MiniExcel | 84.58 ms | 90.2 MB | 1.5 MB | LargeXlsx | 138.5% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | ClosedXML | 307.30 ms | 101.8 MB | 1.1 MB | LargeXlsx | 766.5% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | EPPlus | 377.26 ms | 114.7 MB | 1.1 MB | LargeXlsx | 963.8% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-table-direct | OfficeIMO.Excel | 39.72 ms | 12.4 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-datatable-table-direct | MiniExcel | 83.95 ms | 90.2 MB | 1.5 MB | OfficeIMO.Excel | 111.3% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-table-direct | EPPlus | 357.03 ms | 114.7 MB | 1.1 MB | OfficeIMO.Excel | 798.8% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-table-direct | ClosedXML | 423.24 ms | 169.3 MB | 1.1 MB | OfficeIMO.Excel | 965.4% slower than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | LargeXlsx | 30.91 ms | 5.5 MB | 1.4 MB | LargeXlsx | 6.4% faster than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 33.04 ms | 12.6 MB | 1.4 MB | LargeXlsx | Loss +6.9% |
| 25000 | package-profile | write-fluent-rowsfrom-direct | MiniExcel | 70.99 ms | 91.1 MB | 1.5 MB | LargeXlsx | 114.9% slower than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | ClosedXML | 281.13 ms | 101.8 MB | 1.1 MB | LargeXlsx | 751.0% slower than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | EPPlus | 372.84 ms | 114.7 MB | 1.1 MB | LargeXlsx | 1028.6% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 35.17 ms | 11.2 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 375.88 ms | 156.0 MB | 1.1 MB | OfficeIMO.Excel | 968.8% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 677.49 ms | 485.3 MB | 1.1 MB | OfficeIMO.Excel | 1826.4% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | LargeXlsx | 26.24 ms | 5.5 MB | 1.4 MB | LargeXlsx | 15.6% faster than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | OfficeIMO.Excel | 31.08 ms | 11.2 MB | 1.4 MB | LargeXlsx | Loss +18.5% |
| 25000 | package-profile | write-insertobjects-direct | MiniExcel | 62.38 ms | 91.1 MB | 1.5 MB | LargeXlsx | 100.7% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | ClosedXML | 259.10 ms | 101.8 MB | 1.1 MB | LargeXlsx | 733.6% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | EPPlus | 315.46 ms | 114.7 MB | 1.1 MB | LargeXlsx | 914.9% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 40.26 ms | 9.9 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 361.93 ms | 156.0 MB | 1.1 MB | OfficeIMO.Excel | 799.1% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 657.21 ms | 485.3 MB | 1.1 MB | OfficeIMO.Excel | 1532.5% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 27.38 ms | 5.5 MB | 1.4 MB | LargeXlsx | 27.0% faster than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 37.52 ms | 9.9 MB | 1.4 MB | LargeXlsx | Loss +37.0% |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | MiniExcel | 63.99 ms | 91.1 MB | 1.5 MB | LargeXlsx | 70.6% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | ClosedXML | 283.20 ms | 101.8 MB | 1.1 MB | LargeXlsx | 654.8% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | EPPlus | 330.53 ms | 114.7 MB | 1.1 MB | LargeXlsx | 780.9% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | LargeXlsx | 25.93 ms | 5.5 MB | 1.4 MB | LargeXlsx | 34.3% faster than OfficeIMO |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | OfficeIMO.Excel | 39.44 ms | 15.4 MB | 1.4 MB | LargeXlsx | Loss +52.1% |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | MiniExcel | 65.10 ms | 91.1 MB | 1.5 MB | LargeXlsx | 65.1% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | ClosedXML | 267.03 ms | 101.8 MB | 1.1 MB | LargeXlsx | 577.0% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | EPPlus | 315.02 ms | 114.7 MB | 1.1 MB | LargeXlsx | 698.7% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 32.79 ms | 11.2 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 360.78 ms | 135.1 MB | 1.1 MB | OfficeIMO.Excel | 1000.2% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 431.01 ms | 269.0 MB | 1.1 MB | OfficeIMO.Excel | 1214.4% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-mixed-objects-direct | LargeXlsx | 36.26 ms | 5.9 MB | 1.8 MB | LargeXlsx | 13.8% faster than OfficeIMO |
| 25000 | package-profile | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 42.04 ms | 10.3 MB | 1.8 MB | LargeXlsx | Loss +15.9% |
| 25000 | package-profile | write-powershell-mixed-objects-direct | MiniExcel | 78.74 ms | 111.3 MB | 1.9 MB | LargeXlsx | 87.3% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-mixed-objects-direct | ClosedXML | 359.28 ms | 175.3 MB | 1.5 MB | LargeXlsx | 754.6% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-mixed-objects-direct | EPPlus | 436.99 ms | 141.5 MB | 1.4 MB | LargeXlsx | 939.5% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-mixed-direct | LargeXlsx | 37.54 ms | 5.9 MB | 1.8 MB | LargeXlsx | 13.1% faster than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 43.20 ms | 9.7 MB | 1.8 MB | LargeXlsx | Loss +15.1% |
| 25000 | package-profile | write-powershell-psobject-mixed-direct | MiniExcel | 81.33 ms | 111.3 MB | 1.9 MB | LargeXlsx | 88.3% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-mixed-direct | ClosedXML | 370.89 ms | 175.3 MB | 1.5 MB | LargeXlsx | 758.6% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-mixed-direct | EPPlus | 453.83 ms | 141.5 MB | 1.4 MB | LargeXlsx | 950.6% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 221.70 ms | 35.3 MB | 6.6 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-powershell-psobject-wide-direct | LargeXlsx | 252.35 ms | 22.7 MB | 6.5 MB | OfficeIMO.Excel | 13.8% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-wide-direct | MiniExcel | 401.93 ms | 339.8 MB | 6.8 MB | OfficeIMO.Excel | 81.3% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-wide-direct | ClosedXML | 1.46 s | 476.0 MB | 6.0 MB | OfficeIMO.Excel | 557.9% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-wide-direct | EPPlus | 2.10 s | 549.7 MB | 5.3 MB | OfficeIMO.Excel | 846.6% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | LargeXlsx | 11.00 ms | 2.7 MB |  | LargeXlsx | 22.3% faster than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | OfficeIMO.Excel | 14.17 ms | 10.6 MB |  | LargeXlsx | Loss +28.8% |
| 25000 | speed-comparison | append-plain-rows | MiniExcel | 29.69 ms | 56.9 MB |  | LargeXlsx | 109.5% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | EPPlus 4.5.3.3 | 107.14 ms | 0 B |  | LargeXlsx | 656.2% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | ClosedXML | 127.56 ms | 101.8 MB |  | LargeXlsx | 800.4% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | EPPlus | 191.33 ms | 98.0 MB |  | LargeXlsx | 1250.5% slower than OfficeIMO |
| 25000 | speed-comparison | autofit-existing | OfficeIMO.Excel | 84.43 ms | 15.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | autofit-existing | EPPlus | 465.82 ms | 245.1 MB |  | OfficeIMO.Excel | 451.7% slower than OfficeIMO |
| 25000 | speed-comparison | autofit-existing | EPPlus 4.5.3.3 | 594.23 ms | 0 B |  | OfficeIMO.Excel | 603.8% slower than OfficeIMO |
| 25000 | speed-comparison | autofit-existing | ClosedXML | 1.39 s | 810.3 MB |  | OfficeIMO.Excel | 1546.1% slower than OfficeIMO |
| 25000 | speed-comparison | build-object-datatable-dictionaries | OfficeIMO.Excel | 9.37 ms | 5.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | build-object-datatable-typed | OfficeIMO.Excel | 7.64 ms | 7.9 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-cells | OfficeIMO.Excel | 56.17 ms | 24.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-cells | EPPlus | 339.28 ms | 183.0 MB |  | OfficeIMO.Excel | 504.0% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-cells | ClosedXML | 398.90 ms | 162.6 MB |  | OfficeIMO.Excel | 610.1% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-first-column-from-wide-sheet | OfficeIMO.Excel | 42.80 ms | 3.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-first-column-from-wide-sheet | EPPlus | 310.59 ms | 112.8 MB |  | OfficeIMO.Excel | 625.6% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-first-column-from-wide-sheet | ClosedXML | 366.22 ms | 147.4 MB |  | OfficeIMO.Excel | 755.6% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-range | OfficeIMO.Excel | 56.08 ms | 24.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-range | EPPlus | 324.06 ms | 183.0 MB |  | OfficeIMO.Excel | 477.9% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-range | ClosedXML | 484.17 ms | 162.6 MB |  | OfficeIMO.Excel | 763.4% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-top-range | OfficeIMO.Excel | 0.81 ms | 285.2 KB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-top-range | EPPlus | 354.23 ms | 103.1 MB |  | OfficeIMO.Excel | 43854.4% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-top-range | ClosedXML | 420.23 ms | 145.9 MB |  | OfficeIMO.Excel | 52044.7% slower than OfficeIMO |
| 25000 | speed-comparison | formula-heavy-read | OfficeIMO.Excel | 26.43 ms | 6.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | formula-heavy-read | EPPlus 4.5.3.3 | 77.38 ms | 0 B |  | OfficeIMO.Excel | 192.8% slower than OfficeIMO |
| 25000 | speed-comparison | formula-heavy-read | EPPlus | 218.21 ms | 69.2 MB |  | OfficeIMO.Excel | 725.7% slower than OfficeIMO |
| 25000 | speed-comparison | formula-heavy-read | ClosedXML | 284.10 ms | 77.7 MB |  | OfficeIMO.Excel | 975.0% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | OfficeIMO.Excel | 17.77 ms | 15.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | large-shared-strings | MiniExcel | 31.49 ms | 72.0 MB |  | OfficeIMO.Excel | 77.2% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | EPPlus 4.5.3.3 | 100.12 ms | 0 B |  | OfficeIMO.Excel | 463.3% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | ClosedXML | 127.26 ms | 101.8 MB |  | OfficeIMO.Excel | 616.0% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | EPPlus | 172.71 ms | 82.4 MB |  | OfficeIMO.Excel | 871.8% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | OfficeIMO.Excel | 1.04 ms | 179.8 KB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | large-sparse-column-read | Sylvan.Data.Excel | 1.44 ms | 316.6 KB |  | OfficeIMO.Excel | 38.7% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | ExcelDataReader | 2.05 ms | 4.0 MB |  | OfficeIMO.Excel | 96.8% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | ClosedXML | 3.55 ms | 4.3 MB |  | OfficeIMO.Excel | 240.4% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | MiniExcel | 14.46 ms | 45.1 MB |  | OfficeIMO.Excel | 1288.7% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | EPPlus 4.5.3.3 | 17.64 ms | 0 B |  | OfficeIMO.Excel | 1593.2% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | EPPlus | 95.10 ms | 42.1 MB |  | OfficeIMO.Excel | 9030.8% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | OfficeIMO.Excel | 0.84 ms | 177.2 KB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | large-sparse-row-read | Sylvan.Data.Excel | 1.02 ms | 316.6 KB |  | OfficeIMO.Excel | 22.3% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | ExcelDataReader | 1.72 ms | 4.0 MB |  | OfficeIMO.Excel | 104.8% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | ClosedXML | 3.61 ms | 4.3 MB |  | OfficeIMO.Excel | 330.9% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | MiniExcel | 15.64 ms | 45.1 MB |  | OfficeIMO.Excel | 1767.2% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | EPPlus 4.5.3.3 | 16.88 ms | 0 B |  | OfficeIMO.Excel | 1915.0% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | EPPlus | 100.23 ms | 42.1 MB |  | OfficeIMO.Excel | 11866.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | OfficeIMO.Excel | 35.65 ms | 1.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-bottom-range | Sylvan.Data.Excel | 40.48 ms | 3.5 MB |  | OfficeIMO.Excel | 13.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | ExcelDataReader | 114.77 ms | 59.8 MB |  | OfficeIMO.Excel | 221.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | MiniExcel | 124.93 ms | 182.1 MB |  | OfficeIMO.Excel | 250.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | EPPlus | 233.28 ms | 103.1 MB |  | OfficeIMO.Excel | 554.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | ClosedXML | 358.20 ms | 145.9 MB |  | OfficeIMO.Excel | 904.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | OfficeIMO.Excel | 32.60 ms | 1.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-bottom-range-stream | Sylvan.Data.Excel | 37.16 ms | 3.5 MB |  | OfficeIMO.Excel | 14.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | ExcelDataReader | 105.42 ms | 59.8 MB |  | OfficeIMO.Excel | 223.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | MiniExcel | 114.12 ms | 182.1 MB |  | OfficeIMO.Excel | 250.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | EPPlus | 210.05 ms | 103.1 MB |  | OfficeIMO.Excel | 544.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | ClosedXML | 306.06 ms | 145.9 MB |  | OfficeIMO.Excel | 838.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | Sylvan.Data.Excel | 64.86 ms | 18.0 MB |  | Sylvan.Data.Excel | 11.3% faster than OfficeIMO |
| 25000 | speed-comparison | read-datatable | OfficeIMO.Excel | 73.16 ms | 33.8 MB |  | Sylvan.Data.Excel | Loss +12.8% |
| 25000 | speed-comparison | read-datatable | ExcelDataReader | 164.24 ms | 74.3 MB |  | Sylvan.Data.Excel | 124.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | MiniExcel | 211.57 ms | 177.0 MB |  | Sylvan.Data.Excel | 189.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | EPPlus 4.5.3.3 | 231.49 ms | 0 B |  | Sylvan.Data.Excel | 216.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | EPPlus | 329.36 ms | 197.5 MB |  | Sylvan.Data.Excel | 350.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | ClosedXML | 418.26 ms | 174.3 MB |  | Sylvan.Data.Excel | 471.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | OfficeIMO.Excel | 41.21 ms | 3.9 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | Sylvan.Data.Excel | 55.52 ms | 4.2 MB |  | OfficeIMO.Excel | 34.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | MiniExcel | 118.53 ms | 154.9 MB |  | OfficeIMO.Excel | 187.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | ExcelDataReader | 126.28 ms | 59.8 MB |  | OfficeIMO.Excel | 206.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | EPPlus | 243.80 ms | 112.8 MB |  | OfficeIMO.Excel | 491.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | ClosedXML | 351.77 ms | 147.4 MB |  | OfficeIMO.Excel | 753.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | Sylvan.Data.Excel | 47.62 ms | 5.7 MB |  | Sylvan.Data.Excel | 10.7% faster than OfficeIMO |
| 25000 | speed-comparison | read-objects | OfficeIMO.Excel | 53.33 ms | 23.0 MB |  | Sylvan.Data.Excel | Loss +12.0% |
| 25000 | speed-comparison | read-objects | ExcelDataReader | 118.13 ms | 62.0 MB |  | Sylvan.Data.Excel | 121.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | MiniExcel | 139.60 ms | 179.4 MB |  | Sylvan.Data.Excel | 161.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | EPPlus 4.5.3.3 | 169.30 ms | 0 B |  | Sylvan.Data.Excel | 217.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | EPPlus | 273.78 ms | 194.9 MB |  | Sylvan.Data.Excel | 413.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | ClosedXML | 356.08 ms | 161.7 MB |  | Sylvan.Data.Excel | 567.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | Sylvan.Data.Excel | 45.12 ms | 5.2 MB |  | Sylvan.Data.Excel, OfficeIMO.Excel | Tie vs OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | OfficeIMO.Excel | 45.81 ms | 22.8 MB |  | Sylvan.Data.Excel, OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-objects-stream | ExcelDataReader | 111.86 ms | 61.5 MB |  | Sylvan.Data.Excel, OfficeIMO.Excel | 144.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | MiniExcel | 122.92 ms | 178.9 MB |  | Sylvan.Data.Excel, OfficeIMO.Excel | 168.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | EPPlus 4.5.3.3 | 163.40 ms | 0 B |  | Sylvan.Data.Excel, OfficeIMO.Excel | 256.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | EPPlus | 245.97 ms | 194.7 MB |  | Sylvan.Data.Excel, OfficeIMO.Excel | 437.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | ClosedXML | 327.12 ms | 161.5 MB |  | Sylvan.Data.Excel, OfficeIMO.Excel | 614.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | Sylvan.Data.Excel | 73.75 ms | 3.5 MB |  | Sylvan.Data.Excel | 5.1% faster than OfficeIMO |
| 25000 | speed-comparison | read-range | OfficeIMO.Excel | 77.69 ms | 25.5 MB |  | Sylvan.Data.Excel | Loss +5.4% |
| 25000 | speed-comparison | read-range | MiniExcel | 170.99 ms | 182.1 MB |  | Sylvan.Data.Excel | 120.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | EPPlus 4.5.3.3 | 180.13 ms | 0 B |  | Sylvan.Data.Excel | 131.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | ExcelDataReader | 185.62 ms | 59.8 MB |  | Sylvan.Data.Excel | 138.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | EPPlus | 481.14 ms | 183.0 MB |  | Sylvan.Data.Excel | 519.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | ClosedXML | 515.27 ms | 159.8 MB |  | Sylvan.Data.Excel | 563.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | OfficeIMO.Excel | 88.39 ms | 26.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-range-decimal | Sylvan.Data.Excel | 105.24 ms | 4.4 MB |  | OfficeIMO.Excel | 19.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | ExcelDataReader | 221.67 ms | 59.8 MB |  | OfficeIMO.Excel | 150.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | MiniExcel | 333.49 ms | 182.1 MB |  | OfficeIMO.Excel | 277.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | EPPlus | 522.21 ms | 183.0 MB |  | OfficeIMO.Excel | 490.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | ClosedXML | 698.43 ms | 159.8 MB |  | OfficeIMO.Excel | 690.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | Sylvan.Data.Excel | 45.85 ms | 3.5 MB |  | Sylvan.Data.Excel | 36.7% faster than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | OfficeIMO.Excel | 72.39 ms | 26.3 MB |  | Sylvan.Data.Excel | Loss +57.9% |
| 25000 | speed-comparison | read-range-stream | ExcelDataReader | 132.48 ms | 59.8 MB |  | Sylvan.Data.Excel | 83.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | MiniExcel | 136.69 ms | 182.1 MB |  | Sylvan.Data.Excel | 88.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | EPPlus 4.5.3.3 | 166.05 ms | 0 B |  | Sylvan.Data.Excel | 129.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | EPPlus | 299.44 ms | 183.0 MB |  | Sylvan.Data.Excel | 313.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | ClosedXML | 377.28 ms | 159.8 MB |  | Sylvan.Data.Excel | 421.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | Sylvan.Data.Excel | 0.49 ms | 348.5 KB |  | Sylvan.Data.Excel | 27.9% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range | OfficeIMO.Excel | 0.68 ms | 296.0 KB |  | Sylvan.Data.Excel | Loss +38.7% |
| 25000 | speed-comparison | read-top-range | MiniExcel | 1.10 ms | 869.0 KB |  | Sylvan.Data.Excel | 61.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | ExcelDataReader | 47.34 ms | 16.7 MB |  | Sylvan.Data.Excel | 6841.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | EPPlus 4.5.3.3 | 147.60 ms | 0 B |  | Sylvan.Data.Excel | 21544.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | EPPlus | 303.47 ms | 103.1 MB |  | Sylvan.Data.Excel | 44398.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | ClosedXML | 443.69 ms | 145.9 MB |  | Sylvan.Data.Excel | 64959.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | Sylvan.Data.Excel | 0.42 ms | 348.5 KB |  | Sylvan.Data.Excel | 23.0% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | OfficeIMO.Excel | 0.54 ms | 299.3 KB |  | Sylvan.Data.Excel | Loss +29.9% |
| 25000 | speed-comparison | read-top-range-stream | MiniExcel | 0.79 ms | 869.0 KB |  | Sylvan.Data.Excel | 46.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | ExcelDataReader | 37.83 ms | 16.7 MB |  | Sylvan.Data.Excel | 6890.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | EPPlus 4.5.3.3 | 135.07 ms | 0 B |  | Sylvan.Data.Excel | 24862.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | EPPlus | 217.49 ms | 103.1 MB |  | Sylvan.Data.Excel | 40094.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | ClosedXML | 310.88 ms | 145.9 MB |  | Sylvan.Data.Excel | 57352.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | Sylvan.Data.Excel | 0.41 ms | 348.5 KB |  | Sylvan.Data.Excel | 26.9% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | OfficeIMO.Excel | 0.56 ms | 300.0 KB |  | Sylvan.Data.Excel | Loss +36.8% |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | MiniExcel | 0.68 ms | 869.0 KB |  | Sylvan.Data.Excel | 21.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | ExcelDataReader | 39.13 ms | 16.7 MB |  | Sylvan.Data.Excel | 6878.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | EPPlus | 220.47 ms | 103.1 MB |  | Sylvan.Data.Excel | 39215.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | ClosedXML | 311.29 ms | 145.9 MB |  | Sylvan.Data.Excel | 55411.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-used-range | Sylvan.Data.Excel | 145.45 ms | 3.5 MB |  | Sylvan.Data.Excel | 48.1% faster than OfficeIMO |
| 25000 | speed-comparison | read-used-range | ExcelDataReader | 217.22 ms | 59.8 MB |  | Sylvan.Data.Excel | 22.5% faster than OfficeIMO |
| 25000 | speed-comparison | read-used-range | MiniExcel | 230.58 ms | 182.1 MB |  | Sylvan.Data.Excel | 17.7% faster than OfficeIMO |
| 25000 | speed-comparison | read-used-range | OfficeIMO.Excel | 280.18 ms | 33.4 MB |  | Sylvan.Data.Excel | Loss +92.6% |
| 25000 | speed-comparison | read-used-range | EPPlus | 801.83 ms | 183.0 MB |  | Sylvan.Data.Excel | 186.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-used-range | ClosedXML | 860.09 ms | 159.8 MB |  | Sylvan.Data.Excel | 207.0% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-autofilter | OfficeIMO.Excel | 32.00 ms | 11.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-autofilter | EPPlus 4.5.3.3 | 266.67 ms | 0 B |  | OfficeIMO.Excel | 733.5% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-autofilter | ClosedXML | 294.02 ms | 205.7 MB |  | OfficeIMO.Excel | 818.9% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-autofilter | EPPlus | 348.45 ms | 206.9 MB |  | OfficeIMO.Excel | 989.1% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-charts | OfficeIMO.Excel | 34.42 ms | 12.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-charts | EPPlus 4.5.3.3 | 329.28 ms | 0 B |  | OfficeIMO.Excel | 856.8% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-charts | EPPlus | 356.03 ms | 209.9 MB |  | OfficeIMO.Excel | 934.5% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-conditional-formatting | OfficeIMO.Excel | 36.56 ms | 11.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-conditional-formatting | EPPlus 4.5.3.3 | 289.19 ms | 0 B |  | OfficeIMO.Excel | 691.1% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-conditional-formatting | ClosedXML | 324.45 ms | 205.8 MB |  | OfficeIMO.Excel | 787.6% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-conditional-formatting | EPPlus | 396.52 ms | 206.9 MB |  | OfficeIMO.Excel | 984.7% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-data-validation | OfficeIMO.Excel | 34.22 ms | 11.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-data-validation | ClosedXML | 301.17 ms | 205.7 MB |  | OfficeIMO.Excel | 780.0% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-data-validation | EPPlus 4.5.3.3 | 331.44 ms | 0 B |  | OfficeIMO.Excel | 868.4% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-data-validation | EPPlus | 362.22 ms | 206.9 MB |  | OfficeIMO.Excel | 958.4% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-freeze-panes | OfficeIMO.Excel | 31.21 ms | 11.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-freeze-panes | EPPlus 4.5.3.3 | 241.19 ms | 0 B |  | OfficeIMO.Excel | 672.9% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-freeze-panes | ClosedXML | 301.69 ms | 205.7 MB |  | OfficeIMO.Excel | 866.8% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-freeze-panes | EPPlus | 386.96 ms | 206.9 MB |  | OfficeIMO.Excel | 1140.0% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-pivot-table | OfficeIMO.Excel | 185.03 ms | 128.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-pivot-table | EPPlus 4.5.3.3 | 282.74 ms | 0 B |  | OfficeIMO.Excel | 52.8% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-pivot-table | EPPlus | 374.84 ms | 225.4 MB |  | OfficeIMO.Excel | 102.6% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-all-in-one | OfficeIMO.Excel | 358.93 ms | 130.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-all-in-one | EPPlus 4.5.3.3 | 459.78 ms | 0 B |  | OfficeIMO.Excel | 28.1% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-all-in-one | EPPlus | 628.05 ms | 270.6 MB |  | OfficeIMO.Excel | 75.0% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-chart-first | OfficeIMO.Excel | 99.49 ms | 42.5 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-chart-first | EPPlus 4.5.3.3 | 461.26 ms | 0 B |  | OfficeIMO.Excel | 363.6% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-chart-first | EPPlus | 510.86 ms | 270.6 MB |  | OfficeIMO.Excel | 413.5% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-core | OfficeIMO.Excel | 35.30 ms | 11.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-core | EPPlus | 441.33 ms | 249.1 MB |  | OfficeIMO.Excel | 1150.1% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-core | EPPlus 4.5.3.3 | 478.68 ms | 0 B |  | OfficeIMO.Excel | 1255.8% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-core | ClosedXML | 886.05 ms | 664.2 MB |  | OfficeIMO.Excel | 2409.7% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-extra-column | OfficeIMO.Excel | 370.85 ms | 141.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-extra-column | EPPlus 4.5.3.3 | 505.82 ms | 0 B |  | OfficeIMO.Excel | 36.4% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-extra-column | EPPlus | 519.47 ms | 295.7 MB |  | OfficeIMO.Excel | 40.1% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-no-autofit | EPPlus 4.5.3.3 | 264.87 ms | 0 B |  | EPPlus 4.5.3.3 | 13.4% faster than OfficeIMO |
| 25000 | speed-comparison | realworld-report-no-autofit | OfficeIMO.Excel | 305.85 ms | 130.3 MB |  | EPPlus 4.5.3.3 | Loss +15.5% |
| 25000 | speed-comparison | realworld-report-no-autofit | EPPlus | 589.21 ms | 229.3 MB |  | EPPlus 4.5.3.3 | 92.6% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-post-mutation | OfficeIMO.Excel | 321.76 ms | 130.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-post-mutation | EPPlus 4.5.3.3 | 474.57 ms | 0 B |  | OfficeIMO.Excel | 47.5% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-post-mutation | EPPlus | 494.52 ms | 270.6 MB |  | OfficeIMO.Excel | 53.7% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-shuffled-columns | OfficeIMO.Excel | 304.37 ms | 130.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-shuffled-columns | EPPlus | 474.39 ms | 270.6 MB |  | OfficeIMO.Excel | 55.9% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-shuffled-columns | EPPlus 4.5.3.3 | 517.28 ms | 0 B |  | OfficeIMO.Excel | 70.0% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook | OfficeIMO.Excel | 339.08 ms | 171.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | report-workbook | EPPlus | 578.04 ms | 356.2 MB |  | OfficeIMO.Excel | 70.5% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook | EPPlus 4.5.3.3 | 1.10 s | 0 B |  | OfficeIMO.Excel | 223.7% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-core | OfficeIMO.Excel | 49.50 ms | 10.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | report-workbook-core | EPPlus | 535.90 ms | 334.8 MB |  | OfficeIMO.Excel | 982.7% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-core | EPPlus 4.5.3.3 | 788.24 ms | 0 B |  | OfficeIMO.Excel | 1492.5% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-core | ClosedXML | 1.11 s | 952.9 MB |  | OfficeIMO.Excel | 2137.3% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-datatable | OfficeIMO.Excel | 343.99 ms | 173.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | report-workbook-datatable | EPPlus | 547.81 ms | 242.0 MB |  | OfficeIMO.Excel | 59.2% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-datatable | EPPlus 4.5.3.3 | 694.45 ms | 0 B |  | OfficeIMO.Excel | 101.9% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-datatable-core | OfficeIMO.Excel | 59.82 ms | 13.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | report-workbook-datatable-core | EPPlus | 569.30 ms | 220.7 MB |  | OfficeIMO.Excel | 851.7% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-datatable-core | EPPlus 4.5.3.3 | 668.83 ms | 0 B |  | OfficeIMO.Excel | 1018.1% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-datatable-core | ClosedXML | 1.22 s | 812.7 MB |  | OfficeIMO.Excel | 1946.7% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | Sylvan.Data.Excel | 19.81 ms | 1.9 MB |  | Sylvan.Data.Excel | 22.6% faster than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | OfficeIMO.Excel | 25.58 ms | 9.0 MB |  | Sylvan.Data.Excel | Loss +29.2% |
| 25000 | speed-comparison | shared-string-read | ExcelDataReader | 67.07 ms | 24.4 MB |  | Sylvan.Data.Excel | 162.2% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | MiniExcel | 79.38 ms | 72.7 MB |  | Sylvan.Data.Excel | 210.3% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | EPPlus 4.5.3.3 | 93.82 ms | 0 B |  | Sylvan.Data.Excel | 266.7% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | EPPlus | 186.17 ms | 87.3 MB |  | Sylvan.Data.Excel | 627.7% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | ClosedXML | 225.16 ms | 88.3 MB |  | Sylvan.Data.Excel | 780.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-blog-2023-20-string-columns | LargeXlsx | 38.87 ms | 10.5 MB |  | LargeXlsx | 9.8% faster than OfficeIMO |
| 25000 | speed-comparison | write-blog-2023-20-string-columns | OfficeIMO.Excel | 43.08 ms | 11.4 MB |  | LargeXlsx | Loss +10.8% |
| 25000 | speed-comparison | write-blog-2023-20-string-columns | MiniExcel | 143.92 ms | 221.6 MB |  | LargeXlsx | 234.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-blog-2023-20-string-columns | ClosedXML | 945.66 ms | 742.0 MB |  | LargeXlsx | 2095.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | OfficeIMO.Excel | 35.64 ms | 11.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-bulk-report | MiniExcel | 68.60 ms | 122.6 MB |  | OfficeIMO.Excel | 92.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | EPPlus | 392.93 ms | 249.0 MB |  | OfficeIMO.Excel | 1002.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | EPPlus 4.5.3.3 | 536.36 ms | 0 B |  | OfficeIMO.Excel | 1405.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | ClosedXML | 753.29 ms | 552.7 MB |  | OfficeIMO.Excel | 2013.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellformula | OfficeIMO.Excel | 19.84 ms | 9.9 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellformula | EPPlus 4.5.3.3 | 159.35 ms | 0 B |  | OfficeIMO.Excel | 703.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellformula | ClosedXML | 178.59 ms | 111.2 MB |  | OfficeIMO.Excel | 800.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellformula | EPPlus | 305.37 ms | 137.4 MB |  | OfficeIMO.Excel | 1439.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-empty-strings | OfficeIMO.Excel | 11.80 ms | 6.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-empty-strings | ClosedXML | 117.84 ms | 90.7 MB |  | OfficeIMO.Excel | 898.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-empty-strings | EPPlus | 169.41 ms | 72.7 MB |  | OfficeIMO.Excel | 1335.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-numbers | OfficeIMO.Excel | 14.08 ms | 5.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-numbers | EPPlus 4.5.3.3 | 90.82 ms | 0 B |  | OfficeIMO.Excel | 545.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-numbers | ClosedXML | 105.01 ms | 82.2 MB |  | OfficeIMO.Excel | 646.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-numbers | EPPlus | 181.09 ms | 84.4 MB |  | OfficeIMO.Excel | 1186.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-mixed | OfficeIMO.Excel | 17.23 ms | 8.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-object-mixed | EPPlus 4.5.3.3 | 123.47 ms | 0 B |  | OfficeIMO.Excel | 616.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-mixed | ClosedXML | 165.09 ms | 108.5 MB |  | OfficeIMO.Excel | 858.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-mixed | EPPlus | 211.89 ms | 110.6 MB |  | OfficeIMO.Excel | 1130.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-sparse | OfficeIMO.Excel | 19.01 ms | 7.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-object-sparse | ClosedXML | 146.53 ms | 102.8 MB |  | OfficeIMO.Excel | 671.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-sparse | EPPlus | 210.63 ms | 103.8 MB |  | OfficeIMO.Excel | 1008.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 17.31 ms | 7.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-object-sparse-batch | ClosedXML | 148.91 ms | 102.8 MB |  | OfficeIMO.Excel | 760.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-sparse-batch | EPPlus | 223.17 ms | 103.8 MB |  | OfficeIMO.Excel | 1189.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-scalars | OfficeIMO.Excel | 11.20 ms | 6.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-scalars | EPPlus 4.5.3.3 | 92.70 ms | 0 B |  | OfficeIMO.Excel | 728.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-scalars | ClosedXML | 99.54 ms | 80.7 MB |  | OfficeIMO.Excel | 789.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-scalars | EPPlus | 181.22 ms | 83.1 MB |  | OfficeIMO.Excel | 1518.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings | OfficeIMO.Excel | 17.93 ms | 15.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-strings | EPPlus 4.5.3.3 | 89.32 ms | 0 B |  | OfficeIMO.Excel | 398.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings | ClosedXML | 116.92 ms | 101.8 MB |  | OfficeIMO.Excel | 551.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings | EPPlus | 180.55 ms | 82.4 MB |  | OfficeIMO.Excel | 906.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings-distinct | OfficeIMO.Excel | 13.01 ms | 13.5 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-strings-distinct | ClosedXML | 150.14 ms | 128.4 MB |  | OfficeIMO.Excel | 1053.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings-distinct | EPPlus | 206.81 ms | 95.4 MB |  | OfficeIMO.Excel | 1489.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings-repeated | OfficeIMO.Excel | 11.70 ms | 7.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-strings-repeated | ClosedXML | 100.29 ms | 82.5 MB |  | OfficeIMO.Excel | 757.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings-repeated | EPPlus | 152.86 ms | 68.4 MB |  | OfficeIMO.Excel | 1206.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-temporal | OfficeIMO.Excel | 19.09 ms | 7.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-temporal | EPPlus 4.5.3.3 | 103.26 ms | 0 B |  | OfficeIMO.Excel | 440.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-temporal | ClosedXML | 152.16 ms | 87.2 MB |  | OfficeIMO.Excel | 696.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-temporal | EPPlus | 198.66 ms | 101.4 MB |  | OfficeIMO.Excel | 940.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 14.30 ms | 6.8 MB |  | OfficeIMO.Excel, LargeXlsx | Win |
| 25000 | speed-comparison | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 14.37 ms | 3.4 MB |  | OfficeIMO.Excel, LargeXlsx | Tie vs OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-headerless-rectangle-direct | ClosedXML | 215.52 ms | 93.8 MB |  | OfficeIMO.Excel, LargeXlsx | 1407.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-headerless-rectangle-direct | EPPlus | 238.62 ms | 85.4 MB |  | OfficeIMO.Excel, LargeXlsx | 1568.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | LargeXlsx | 30.17 ms | 5.5 MB |  | LargeXlsx | 13.3% faster than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 34.79 ms | 15.7 MB |  | LargeXlsx | Loss +15.3% |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | MiniExcel | 64.70 ms | 91.1 MB |  | LargeXlsx | 85.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 239.21 ms | 0 B |  | LargeXlsx | 587.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | ClosedXML | 303.26 ms | 205.7 MB |  | LargeXlsx | 771.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus | 369.82 ms | 206.9 MB |  | LargeXlsx | 962.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-sparse-rectangle-direct | OfficeIMO.Excel | 21.76 ms | 7.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalues-sparse-rectangle-direct | ClosedXML | 192.13 ms | 102.8 MB |  | OfficeIMO.Excel | 783.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-sparse-rectangle-direct | EPPlus | 268.93 ms | 103.8 MB |  | OfficeIMO.Excel | 1136.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | Sylvan.Data.Excel | 34.88 ms | 5.6 MB |  | Sylvan.Data.Excel | 20.9% faster than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | LargeXlsx | 43.45 ms | 8.2 MB |  | Sylvan.Data.Excel | Tie vs OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | OfficeIMO.Excel | 44.10 ms | 12.7 MB |  | Sylvan.Data.Excel | Loss +26.5% |
| 25000 | speed-comparison | write-datareader-plain | MiniExcel | 85.60 ms | 90.0 MB |  | Sylvan.Data.Excel | 94.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | EPPlus 4.5.3.3 | 223.51 ms | 0 B |  | Sylvan.Data.Excel | 406.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | ClosedXML | 334.79 ms | 101.8 MB |  | Sylvan.Data.Excel | 659.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | EPPlus | 389.85 ms | 114.7 MB |  | Sylvan.Data.Excel | 784.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | OfficeIMO.Excel | 35.38 ms | 12.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datareader-table | MiniExcel | 69.40 ms | 90.0 MB |  | OfficeIMO.Excel | 96.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | EPPlus 4.5.3.3 | 224.92 ms | 0 B |  | OfficeIMO.Excel | 535.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | ClosedXML | 365.06 ms | 169.3 MB |  | OfficeIMO.Excel | 931.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | EPPlus | 371.70 ms | 114.7 MB |  | OfficeIMO.Excel | 950.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table-autofit | OfficeIMO.Excel | 38.73 ms | 12.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datareader-table-autofit | MiniExcel | 73.10 ms | 121.6 MB |  | OfficeIMO.Excel | 88.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table-autofit | EPPlus | 408.69 ms | 156.0 MB |  | OfficeIMO.Excel | 955.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table-autofit | EPPlus 4.5.3.3 | 437.25 ms | 0 B |  | OfficeIMO.Excel | 1029.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table-autofit | ClosedXML | 839.51 ms | 552.9 MB |  | OfficeIMO.Excel | 2067.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-headerless-tables | OfficeIMO.Excel | 33.84 ms | 12.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-headerless-tables | MiniExcel | 75.43 ms | 94.8 MB |  | OfficeIMO.Excel | 122.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-headerless-tables | ClosedXML | 358.59 ms | 168.0 MB |  | OfficeIMO.Excel | 959.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-headerless-tables | EPPlus | 421.15 ms | 108.6 MB |  | OfficeIMO.Excel | 1144.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | OfficeIMO.Excel | 41.75 ms | 9.6 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | LargeXlsx | 57.57 ms | 9.0 MB |  | OfficeIMO.Excel | 37.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | MiniExcel | 182.53 ms | 105.6 MB |  | OfficeIMO.Excel | 337.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | ClosedXML | 715.69 ms | 273.8 MB |  | OfficeIMO.Excel | 1614.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | EPPlus | 730.53 ms | 132.5 MB |  | OfficeIMO.Excel | 1649.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-tables | OfficeIMO.Excel | 38.39 ms | 13.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-sparse-tables | MiniExcel | 97.36 ms | 105.6 MB |  | OfficeIMO.Excel | 153.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-tables | EPPlus | 473.93 ms | 132.5 MB |  | OfficeIMO.Excel | 1134.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-tables | ClosedXML | 511.60 ms | 273.8 MB |  | OfficeIMO.Excel | 1232.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | OfficeIMO.Excel | 34.56 ms | 10.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-tables | MiniExcel | 79.96 ms | 94.8 MB |  | OfficeIMO.Excel | 131.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | EPPlus 4.5.3.3 | 213.50 ms | 0 B |  | OfficeIMO.Excel | 517.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | EPPlus | 330.24 ms | 108.2 MB |  | OfficeIMO.Excel | 855.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | ClosedXML | 364.23 ms | 168.0 MB |  | OfficeIMO.Excel | 953.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables-autofit | OfficeIMO.Excel | 36.71 ms | 10.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-tables-autofit | MiniExcel | 78.03 ms | 125.9 MB |  | OfficeIMO.Excel | 112.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables-autofit | EPPlus | 386.42 ms | 190.8 MB |  | OfficeIMO.Excel | 952.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables-autofit | ClosedXML | 724.06 ms | 537.2 MB |  | OfficeIMO.Excel | 1872.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | LargeXlsx | 30.70 ms | 9.3 MB |  | LargeXlsx | 11.9% faster than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | OfficeIMO.Excel | 34.85 ms | 12.4 MB |  | LargeXlsx | Loss +13.5% |
| 25000 | speed-comparison | write-datatable-direct | MiniExcel | 83.37 ms | 90.2 MB |  | LargeXlsx | 139.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | EPPlus 4.5.3.3 | 216.60 ms | 0 B |  | LargeXlsx | 521.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | ClosedXML | 275.12 ms | 101.8 MB |  | LargeXlsx | 689.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | EPPlus | 372.68 ms | 114.7 MB |  | LargeXlsx | 969.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-object-table-direct | OfficeIMO.Excel | 34.13 ms | 9.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datatable-object-table-direct | MiniExcel | 75.33 ms | 87.6 MB |  | OfficeIMO.Excel | 120.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-object-table-direct | EPPlus | 315.89 ms | 112.0 MB |  | OfficeIMO.Excel | 825.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-object-table-direct | ClosedXML | 361.05 ms | 166.7 MB |  | OfficeIMO.Excel | 957.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | OfficeIMO.Excel | 32.76 ms | 12.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datatable-table-direct | MiniExcel | 74.70 ms | 90.2 MB |  | OfficeIMO.Excel | 128.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | EPPlus 4.5.3.3 | 239.78 ms | 0 B |  | OfficeIMO.Excel | 632.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | ClosedXML | 360.22 ms | 169.3 MB |  | OfficeIMO.Excel | 999.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | EPPlus | 360.86 ms | 114.7 MB |  | OfficeIMO.Excel | 1001.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-dictionary-objects-table-direct | OfficeIMO.Excel | 45.43 ms | 14.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | LargeXlsx | 30.29 ms | 5.5 MB |  | LargeXlsx | 13.7% faster than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 35.11 ms | 12.6 MB |  | LargeXlsx | Loss +15.9% |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | MiniExcel | 69.65 ms | 91.1 MB |  | LargeXlsx | 98.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 248.03 ms | 0 B |  | LargeXlsx | 606.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | ClosedXML | 281.67 ms | 101.8 MB |  | LargeXlsx | 702.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus | 367.65 ms | 114.7 MB |  | LargeXlsx | 947.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 36.06 ms | 11.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 365.29 ms | 156.0 MB |  | OfficeIMO.Excel | 913.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 683.69 ms | 485.3 MB |  | OfficeIMO.Excel | 1796.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | LargeXlsx | 29.04 ms | 5.5 MB |  | LargeXlsx | 11.0% faster than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | OfficeIMO.Excel | 32.62 ms | 11.2 MB |  | LargeXlsx | Loss +12.3% |
| 25000 | speed-comparison | write-insertobjects-direct | MiniExcel | 61.94 ms | 91.1 MB |  | LargeXlsx | 89.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | ClosedXML | 274.41 ms | 101.8 MB |  | LargeXlsx | 741.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | EPPlus 4.5.3.3 | 286.11 ms | 0 B |  | LargeXlsx | 777.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | EPPlus | 332.68 ms | 114.7 MB |  | LargeXlsx | 920.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 40.10 ms | 9.9 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 362.06 ms | 156.0 MB |  | OfficeIMO.Excel | 803.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 680.80 ms | 485.3 MB |  | OfficeIMO.Excel | 1597.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 28.64 ms | 5.5 MB |  | LargeXlsx | 26.3% faster than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 38.86 ms | 9.9 MB |  | LargeXlsx | Loss +35.7% |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | MiniExcel | 65.42 ms | 91.1 MB |  | LargeXlsx | 68.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | ClosedXML | 284.76 ms | 101.8 MB |  | LargeXlsx | 632.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | EPPlus | 339.38 ms | 114.7 MB |  | LargeXlsx | 773.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 33.77 ms | 11.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 362.44 ms | 135.1 MB |  | OfficeIMO.Excel | 973.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 437.83 ms | 269.0 MB |  | OfficeIMO.Excel | 1196.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | LargeXlsx | 41.10 ms | 5.9 MB |  | LargeXlsx | 10.8% faster than OfficeIMO |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 46.09 ms | 10.3 MB |  | LargeXlsx | Loss +12.1% |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | MiniExcel | 83.52 ms | 111.3 MB |  | LargeXlsx | 81.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | ClosedXML | 390.53 ms | 175.3 MB |  | LargeXlsx | 747.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | EPPlus | 470.46 ms | 141.5 MB |  | LargeXlsx | 920.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-mixed-direct | LargeXlsx | 40.56 ms | 5.9 MB |  | LargeXlsx | 11.6% faster than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 45.90 ms | 9.7 MB |  | LargeXlsx | Loss +13.2% |
| 25000 | speed-comparison | write-powershell-psobject-mixed-direct | MiniExcel | 88.65 ms | 111.3 MB |  | LargeXlsx | 93.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-mixed-direct | ClosedXML | 383.93 ms | 175.3 MB |  | LargeXlsx | 736.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-mixed-direct | EPPlus | 472.44 ms | 141.5 MB |  | LargeXlsx | 929.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 253.66 ms | 35.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-powershell-psobject-wide-direct | LargeXlsx | 263.61 ms | 22.7 MB |  | OfficeIMO.Excel | 3.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-wide-direct | MiniExcel | 446.74 ms | 339.8 MB |  | OfficeIMO.Excel | 76.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-wide-direct | ClosedXML | 1.34 s | 476.0 MB |  | OfficeIMO.Excel | 428.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-wide-direct | EPPlus | 2.03 s | 549.7 MB |  | OfficeIMO.Excel | 700.8% slower than OfficeIMO |
| 300000 | focused-package-profile | write-blog-2023-20-string-columns | LargeXlsx | 636.47 ms | 93.1 MB | 28.6 MB | LargeXlsx | Win |
| 300000 | focused-package-profile | write-blog-2023-20-string-columns | OfficeIMO.Excel | 701.65 ms | 173.4 MB | 26.6 MB | LargeXlsx | 1.10x vs best |
| 300000 | focused-package-profile | write-blog-2023-20-string-columns | MiniExcel | 2.28 s | 2.46 GB | 28.5 MB | LargeXlsx | 3.58x vs best |
| 300000 | focused-package-profile | write-blog-2023-20-string-columns | ClosedXML | 15.89 s | 8.51 GB | 31.0 MB | LargeXlsx | 24.97x vs best |
