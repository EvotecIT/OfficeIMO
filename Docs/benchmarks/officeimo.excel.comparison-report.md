# OfficeIMO.Excel Benchmark Report

Generated: 2026-05-31T10:45:56.6849738Z
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
| 2500 | dense-helloworld-comparison | read | Other | 0 | 2 | dense-helloworld-read-stream vs Sylvan.Data.Excel (1.52x) |
| 2500 | package-profile | package | Package size | 41 | 13 | write-insertobjects-legacy-dictionaries-direct vs LargeXlsx (1.50x) |
| 2500 | speed-comparison | mutate | AutoFit and mutation | 2 | 0 |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | 4 | 0 |  |
| 2500 | speed-comparison | other | Object projection | 2 | 0 |  |
| 2500 | speed-comparison | other | Range and table read | 4 | 0 |  |
| 2500 | speed-comparison | other | Real-world feature mix | 6 | 0 |  |
| 2500 | speed-comparison | other | Real-world report | 2 | 0 |  |
| 2500 | speed-comparison | other | Report workbook | 4 | 0 |  |
| 2500 | speed-comparison | read | Formula write/read | 1 | 0 |  |
| 2500 | speed-comparison | read | Other | 1 | 2 | shared-string-read vs Sylvan.Data.Excel (1.74x) |
| 2500 | speed-comparison | read | Range and table read | 4 | 3 | read-used-range vs Sylvan.Data.Excel (3.17x) |
| 2500 | speed-comparison | read | Streaming read | 1 | 3 | read-top-range-stream-small-chunks vs Sylvan.Data.Excel (1.29x) |
| 2500 | speed-comparison | read | Typed object read | 2 | 0 |  |
| 2500 | speed-comparison | write | AutoFit and mutation | 5 | 0 |  |
| 2500 | speed-comparison | write | Cell writer | 7 | 0 |  |
| 2500 | speed-comparison | write | DataSet table export | 3 | 0 |  |
| 2500 | speed-comparison | write | DataTable table export | 4 | 0 |  |
| 2500 | speed-comparison | write | Formatted report write | 1 | 0 |  |
| 2500 | speed-comparison | write | Formula write/read | 1 | 0 |  |
| 2500 | speed-comparison | write | Other | 1 | 3 | write-powershell-mixed-objects-direct vs LargeXlsx (1.31x) |
| 2500 | speed-comparison | write | Plain cell export | 3 | 1 | append-plain-rows vs LargeXlsx (1.57x) |
| 2500 | speed-comparison | write | Plain streaming export | 1 | 1 | write-datareader-plain vs Sylvan.Data.Excel (1.30x) |
| 2500 | speed-comparison | write | Plain string export | 1 | 0 |  |
| 2500 | speed-comparison | write | Shared string write | 4 | 0 |  |
| 2500 | speed-comparison | write | Typed object export | 0 | 3 | write-insertobjects-flat-dictionaries-direct vs LargeXlsx (1.35x) |
| 10000 | focused-package-profile | package | Package size | 1 | 0 |  |
| 25000 | dense-helloworld-comparison | read | Other | 0 | 2 | dense-helloworld-read-stream vs Sylvan.Data.Excel (1.31x) |
| 25000 | package-profile | package | Package size | 43 | 11 | write-insertobjects-legacy-dictionaries-direct vs LargeXlsx (1.52x) |
| 25000 | speed-comparison | mutate | AutoFit and mutation | 1 | 1 | realworld-report-no-autofit vs EPPlus 4.5.3.3 (1.19x) |
| 25000 | speed-comparison | other | Anti-cheat report variants | 4 | 0 |  |
| 25000 | speed-comparison | other | Object projection | 2 | 0 |  |
| 25000 | speed-comparison | other | Range and table read | 4 | 0 |  |
| 25000 | speed-comparison | other | Real-world feature mix | 6 | 0 |  |
| 25000 | speed-comparison | other | Real-world report | 2 | 0 |  |
| 25000 | speed-comparison | other | Report workbook | 4 | 0 |  |
| 25000 | speed-comparison | read | Formula write/read | 1 | 0 |  |
| 25000 | speed-comparison | read | Other | 2 | 1 | shared-string-read vs Sylvan.Data.Excel (1.14x) |
| 25000 | speed-comparison | read | Range and table read | 3 | 4 | read-used-range vs Sylvan.Data.Excel (2.12x) |
| 25000 | speed-comparison | read | Streaming read | 2 | 2 | read-top-range-stream-small-chunks vs Sylvan.Data.Excel (1.25x) |
| 25000 | speed-comparison | read | Typed object read | 1 | 1 | read-objects vs Sylvan.Data.Excel (1.16x) |
| 25000 | speed-comparison | write | AutoFit and mutation | 5 | 0 |  |
| 25000 | speed-comparison | write | Cell writer | 7 | 0 |  |
| 25000 | speed-comparison | write | DataSet table export | 3 | 0 |  |
| 25000 | speed-comparison | write | DataTable table export | 3 | 1 | write-datatable-direct vs LargeXlsx (1.11x) |
| 25000 | speed-comparison | write | Formatted report write | 1 | 0 |  |
| 25000 | speed-comparison | write | Formula write/read | 1 | 0 |  |
| 25000 | speed-comparison | write | Other | 2 | 2 | write-powershell-psobject-mixed-direct vs LargeXlsx (1.15x) |
| 25000 | speed-comparison | write | Plain cell export | 1 | 3 | append-plain-rows vs LargeXlsx (1.42x) |
| 25000 | speed-comparison | write | Plain streaming export | 0 | 2 | write-datareader-plain vs Sylvan.Data.Excel (1.34x) |
| 25000 | speed-comparison | write | Plain string export | 0 | 1 | write-blog-2023-20-string-columns vs LargeXlsx (1.20x) |
| 25000 | speed-comparison | write | Shared string write | 4 | 0 |  |
| 25000 | speed-comparison | write | Typed object export | 0 | 3 | write-insertobjects-flat-dictionaries-direct vs LargeXlsx (1.32x) |
| 300000 | focused-package-profile | package | Package size | 0 | 1 | write-blog-2023-20-string-columns vs LargeXlsx (1.10x) |

## Rows

| Rows | Artifact | Scenario | Library | Mean | Allocated | Package | Best | Outcome |
| ---: | --- | --- | --- | ---: | ---: | ---: | --- | --- |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | Sylvan.Data.Excel | 4.50 ms | 362.3 KB |  | Sylvan.Data.Excel | 24.4% faster than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | OfficeIMO.Excel | 5.95 ms | 2.4 MB |  | Sylvan.Data.Excel | Loss +32.3% |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | ExcelDataReader | 11.52 ms | 6.7 MB |  | Sylvan.Data.Excel | 93.6% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | MiniExcel | 16.15 ms | 21.0 MB |  | Sylvan.Data.Excel | 171.3% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | Sylvan.Data.Excel | 4.14 ms | 362.3 KB |  | Sylvan.Data.Excel | 34.3% faster than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | OfficeIMO.Excel | 6.31 ms | 2.4 MB |  | Sylvan.Data.Excel | Loss +52.2% |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | ExcelDataReader | 12.91 ms | 6.7 MB |  | Sylvan.Data.Excel | 104.7% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | MiniExcel | 18.16 ms | 21.0 MB |  | Sylvan.Data.Excel | 188.0% slower than OfficeIMO |
| 2500 | package-profile | append-plain-rows | LargeXlsx | 1.45 ms | 296.4 KB | 63.1 KB | LargeXlsx | 23.0% faster than OfficeIMO |
| 2500 | package-profile | append-plain-rows | OfficeIMO.Excel | 1.89 ms | 1.5 MB | 63.0 KB | LargeXlsx | Loss +29.9% |
| 2500 | package-profile | append-plain-rows | MiniExcel | 3.84 ms | 19.2 MB | 68.1 KB | LargeXlsx | 103.8% slower than OfficeIMO |
| 2500 | package-profile | append-plain-rows | ClosedXML | 14.83 ms | 10.9 MB | 59.8 KB | LargeXlsx | 686.5% slower than OfficeIMO |
| 2500 | package-profile | append-plain-rows | EPPlus | 25.58 ms | 14.0 MB | 56.9 KB | LargeXlsx | 1256.5% slower than OfficeIMO |
| 2500 | package-profile | autofit-existing | OfficeIMO.Excel | 8.03 ms | 1.9 MB | 143.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | autofit-existing | EPPlus | 77.65 ms | 49.5 MB | 115.0 KB | OfficeIMO.Excel | 866.8% slower than OfficeIMO |
| 2500 | package-profile | autofit-existing | ClosedXML | 129.85 ms | 82.6 MB | 121.0 KB | OfficeIMO.Excel | 1516.7% slower than OfficeIMO |
| 2500 | package-profile | large-shared-strings | OfficeIMO.Excel | 2.07 ms | 2.4 MB | 55.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | large-shared-strings | MiniExcel | 4.65 ms | 20.6 MB | 60.7 KB | OfficeIMO.Excel | 124.2% slower than OfficeIMO |
| 2500 | package-profile | large-shared-strings | ClosedXML | 13.20 ms | 11.0 MB | 50.3 KB | OfficeIMO.Excel | 536.8% slower than OfficeIMO |
| 2500 | package-profile | large-shared-strings | EPPlus | 24.38 ms | 12.5 MB | 48.1 KB | OfficeIMO.Excel | 1076.1% slower than OfficeIMO |
| 2500 | package-profile | realworld-autofilter | OfficeIMO.Excel | 3.85 ms | 1.3 MB | 142.4 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-autofilter | ClosedXML | 30.44 ms | 21.7 MB | 120.2 KB | OfficeIMO.Excel | 691.5% slower than OfficeIMO |
| 2500 | package-profile | realworld-autofilter | EPPlus | 42.56 ms | 24.1 MB | 114.2 KB | OfficeIMO.Excel | 1006.7% slower than OfficeIMO |
| 2500 | package-profile | realworld-charts | OfficeIMO.Excel | 4.87 ms | 1.8 MB | 147.6 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-charts | EPPlus | 41.15 ms | 26.5 MB | 117.0 KB | OfficeIMO.Excel | 744.5% slower than OfficeIMO |
| 2500 | package-profile | realworld-conditional-formatting | OfficeIMO.Excel | 3.88 ms | 1.4 MB | 142.7 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-conditional-formatting | ClosedXML | 31.02 ms | 21.8 MB | 120.3 KB | OfficeIMO.Excel | 700.2% slower than OfficeIMO |
| 2500 | package-profile | realworld-conditional-formatting | EPPlus | 42.23 ms | 24.2 MB | 114.3 KB | OfficeIMO.Excel | 989.3% slower than OfficeIMO |
| 2500 | package-profile | realworld-data-validation | OfficeIMO.Excel | 3.77 ms | 1.3 MB | 142.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-data-validation | ClosedXML | 34.03 ms | 21.7 MB | 120.3 KB | OfficeIMO.Excel | 802.4% slower than OfficeIMO |
| 2500 | package-profile | realworld-data-validation | EPPlus | 41.69 ms | 24.1 MB | 114.2 KB | OfficeIMO.Excel | 1005.4% slower than OfficeIMO |
| 2500 | package-profile | realworld-freeze-panes | OfficeIMO.Excel | 3.83 ms | 1.3 MB | 142.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-freeze-panes | ClosedXML | 31.93 ms | 21.7 MB | 120.2 KB | OfficeIMO.Excel | 734.3% slower than OfficeIMO |
| 2500 | package-profile | realworld-freeze-panes | EPPlus | 46.38 ms | 24.2 MB | 114.3 KB | OfficeIMO.Excel | 1111.9% slower than OfficeIMO |
| 2500 | package-profile | realworld-pivot-table | OfficeIMO.Excel | 13.50 ms | 14.1 MB | 200.4 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-pivot-table | EPPlus | 45.40 ms | 28.8 MB | 117.4 KB | OfficeIMO.Excel | 236.2% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-all-in-one | OfficeIMO.Excel | 16.27 ms | 14.9 MB | 206.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-all-in-one | EPPlus | 73.85 ms | 53.3 MB | 121.8 KB | OfficeIMO.Excel | 354.0% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-chart-first | OfficeIMO.Excel | 10.29 ms | 6.1 MB | 206.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-chart-first | EPPlus | 68.14 ms | 53.3 MB | 121.8 KB | OfficeIMO.Excel | 562.5% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-core | OfficeIMO.Excel | 4.45 ms | 1.5 MB | 143.9 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-core | EPPlus | 71.66 ms | 46.2 MB | 115.6 KB | OfficeIMO.Excel | 1509.7% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-core | ClosedXML | 80.71 ms | 68.2 MB | 121.5 KB | OfficeIMO.Excel | 1713.0% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-extra-column | OfficeIMO.Excel | 17.40 ms | 16.0 MB | 219.1 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-extra-column | EPPlus | 75.60 ms | 57.8 MB | 128.4 KB | OfficeIMO.Excel | 334.6% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-no-autofit | OfficeIMO.Excel | 14.90 ms | 14.9 MB | 206.4 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-no-autofit | EPPlus | 44.06 ms | 32.1 MB | 121.8 KB | OfficeIMO.Excel | 195.7% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-post-mutation | OfficeIMO.Excel | 15.91 ms | 14.9 MB | 206.6 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-post-mutation | EPPlus | 73.35 ms | 53.3 MB | 121.9 KB | OfficeIMO.Excel | 360.9% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-shuffled-columns | OfficeIMO.Excel | 17.47 ms | 14.9 MB | 211.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-shuffled-columns | EPPlus | 74.73 ms | 53.3 MB | 124.3 KB | OfficeIMO.Excel | 327.8% slower than OfficeIMO |
| 2500 | package-profile | report-workbook | OfficeIMO.Excel | 23.84 ms | 18.7 MB | 275.6 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | report-workbook | EPPlus | 99.75 ms | 75.7 MB | 161.8 KB | OfficeIMO.Excel | 318.4% slower than OfficeIMO |
| 2500 | package-profile | report-workbook-core | OfficeIMO.Excel | 5.79 ms | 2.6 MB | 187.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | report-workbook-core | EPPlus | 93.82 ms | 70.3 MB | 157.2 KB | OfficeIMO.Excel | 1519.6% slower than OfficeIMO |
| 2500 | package-profile | report-workbook-core | ClosedXML | 99.15 ms | 94.9 MB | 165.1 KB | OfficeIMO.Excel | 1611.7% slower than OfficeIMO |
| 2500 | package-profile | report-workbook-datatable | OfficeIMO.Excel | 22.56 ms | 18.9 MB | 275.6 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | report-workbook-datatable | EPPlus | 99.40 ms | 64.4 MB | 161.8 KB | OfficeIMO.Excel | 340.6% slower than OfficeIMO |
| 2500 | package-profile | report-workbook-datatable-core | OfficeIMO.Excel | 5.75 ms | 2.9 MB | 187.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | report-workbook-datatable-core | EPPlus | 92.65 ms | 59.1 MB | 157.2 KB | OfficeIMO.Excel | 1510.6% slower than OfficeIMO |
| 2500 | package-profile | report-workbook-datatable-core | ClosedXML | 96.01 ms | 80.9 MB | 165.1 KB | OfficeIMO.Excel | 1569.1% slower than OfficeIMO |
| 2500 | package-profile | write-blog-2023-20-string-columns | LargeXlsx | 3.94 ms | 857.6 KB | 237.7 KB | LargeXlsx | 18.6% faster than OfficeIMO |
| 2500 | package-profile | write-blog-2023-20-string-columns | OfficeIMO.Excel | 4.84 ms | 1.6 MB | 216.7 KB | LargeXlsx | Loss +22.8% |
| 2500 | package-profile | write-blog-2023-20-string-columns | MiniExcel | 16.17 ms | 35.1 MB | 235.3 KB | LargeXlsx | 234.2% slower than OfficeIMO |
| 2500 | package-profile | write-blog-2023-20-string-columns | ClosedXML | 92.68 ms | 69.8 MB | 257.2 KB | LargeXlsx | 1815.3% slower than OfficeIMO |
| 2500 | package-profile | write-bulk-report | OfficeIMO.Excel | 4.52 ms | 1.4 MB | 143.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-bulk-report | MiniExcel | 8.06 ms | 26.2 MB | 153.8 KB | OfficeIMO.Excel | 78.2% slower than OfficeIMO |
| 2500 | package-profile | write-bulk-report | EPPlus | 64.65 ms | 46.1 MB | 115.0 KB | OfficeIMO.Excel | 1329.2% slower than OfficeIMO |
| 2500 | package-profile | write-bulk-report | ClosedXML | 72.32 ms | 57.0 MB | 121.0 KB | OfficeIMO.Excel | 1498.8% slower than OfficeIMO |
| 2500 | package-profile | write-cellformula | OfficeIMO.Excel | 2.31 ms | 1.4 MB | 66.6 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellformula | ClosedXML | 17.18 ms | 11.8 MB | 70.6 KB | OfficeIMO.Excel | 644.6% slower than OfficeIMO |
| 2500 | package-profile | write-cellformula | EPPlus | 37.17 ms | 17.7 MB | 62.1 KB | OfficeIMO.Excel | 1510.9% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-empty-strings | OfficeIMO.Excel | 1.92 ms | 1.7 MB | 44.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-empty-strings | ClosedXML | 13.09 ms | 9.7 MB | 44.9 KB | OfficeIMO.Excel | 580.4% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-empty-strings | EPPlus | 25.15 ms | 11.5 MB | 42.0 KB | OfficeIMO.Excel | 1207.1% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-numbers | OfficeIMO.Excel | 1.85 ms | 1.1 MB | 47.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-numbers | ClosedXML | 10.86 ms | 9.0 MB | 45.9 KB | OfficeIMO.Excel | 485.7% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-numbers | EPPlus | 22.47 ms | 12.6 MB | 43.7 KB | OfficeIMO.Excel | 1112.2% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-mixed | OfficeIMO.Excel | 2.54 ms | 1.7 MB | 61.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-object-mixed | ClosedXML | 16.14 ms | 11.6 MB | 59.5 KB | OfficeIMO.Excel | 534.6% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-mixed | EPPlus | 26.30 ms | 15.3 MB | 58.9 KB | OfficeIMO.Excel | 934.2% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-sparse | OfficeIMO.Excel | 2.61 ms | 1.5 MB | 62.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-object-sparse | ClosedXML | 15.77 ms | 11.0 MB | 52.5 KB | OfficeIMO.Excel | 504.9% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-sparse | EPPlus | 25.38 ms | 14.6 MB | 54.2 KB | OfficeIMO.Excel | 873.2% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 2.32 ms | 1.5 MB | 62.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-object-sparse-batch | ClosedXML | 15.00 ms | 11.0 MB | 52.5 KB | OfficeIMO.Excel | 547.2% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-sparse-batch | EPPlus | 26.06 ms | 14.6 MB | 54.2 KB | OfficeIMO.Excel | 1024.5% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-scalars | OfficeIMO.Excel | 1.73 ms | 1.1 MB | 46.3 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-scalars | ClosedXML | 11.16 ms | 8.8 MB | 45.4 KB | OfficeIMO.Excel | 543.9% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-scalars | EPPlus | 23.92 ms | 12.5 MB | 42.4 KB | OfficeIMO.Excel | 1280.2% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings | OfficeIMO.Excel | 2.54 ms | 2.6 MB | 55.1 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-strings | ClosedXML | 12.19 ms | 11.0 MB | 50.3 KB | OfficeIMO.Excel | 380.7% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings | EPPlus | 24.36 ms | 12.5 MB | 48.1 KB | OfficeIMO.Excel | 861.1% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings-distinct | OfficeIMO.Excel | 2.41 ms | 2.3 MB | 51.8 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-strings-distinct | ClosedXML | 15.38 ms | 12.8 MB | 61.9 KB | OfficeIMO.Excel | 539.2% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings-distinct | EPPlus | 26.81 ms | 13.6 MB | 61.5 KB | OfficeIMO.Excel | 1014.3% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings-repeated | OfficeIMO.Excel | 1.74 ms | 1.5 MB | 40.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-strings-repeated | ClosedXML | 10.61 ms | 9.0 MB | 38.8 KB | OfficeIMO.Excel | 509.4% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings-repeated | EPPlus | 20.33 ms | 11.1 MB | 34.8 KB | OfficeIMO.Excel | 1067.6% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-temporal | OfficeIMO.Excel | 3.07 ms | 1.4 MB | 63.3 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-temporal | ClosedXML | 16.52 ms | 9.5 MB | 54.5 KB | OfficeIMO.Excel | 437.8% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-temporal | EPPlus | 27.36 ms | 14.4 MB | 53.1 KB | OfficeIMO.Excel | 790.9% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 1.19 ms | 447.0 KB | 47.3 KB | LargeXlsx | 18.6% faster than OfficeIMO |
| 2500 | package-profile | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 1.47 ms | 1.1 MB | 48.2 KB | LargeXlsx | Loss +22.9% |
| 2500 | package-profile | write-cellvalues-headerless-rectangle-direct | ClosedXML | 11.73 ms | 10.0 MB | 53.0 KB | LargeXlsx | 700.6% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-headerless-rectangle-direct | EPPlus | 21.13 ms | 12.7 MB | 52.5 KB | LargeXlsx | 1341.8% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | LargeXlsx | 3.16 ms | 758.3 KB | 138.4 KB | LargeXlsx | 17.8% faster than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 3.85 ms | 2.0 MB | 138.0 KB | LargeXlsx | Loss +21.6% |
| 2500 | package-profile | write-cellvalues-rectangle-direct | MiniExcel | 7.35 ms | 22.7 MB | 153.7 KB | LargeXlsx | 91.0% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | ClosedXML | 32.88 ms | 21.7 MB | 120.1 KB | LargeXlsx | 754.6% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | EPPlus | 40.39 ms | 24.1 MB | 114.1 KB | LargeXlsx | 949.9% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | Sylvan.Data.Excel | 3.27 ms | 758.7 KB | 78.5 KB | Sylvan.Data.Excel | 14.9% faster than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | LargeXlsx | 3.79 ms | 1.0 MB | 138.4 KB | Sylvan.Data.Excel | Tie vs OfficeIMO |
| 2500 | package-profile | write-datareader-plain | OfficeIMO.Excel | 3.84 ms | 1.7 MB | 138.0 KB | Sylvan.Data.Excel | Loss +17.5% |
| 2500 | package-profile | write-datareader-plain | MiniExcel | 7.28 ms | 22.5 MB | 153.6 KB | Sylvan.Data.Excel | 89.5% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | ClosedXML | 26.63 ms | 11.3 MB | 120.1 KB | Sylvan.Data.Excel | 593.5% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | EPPlus | 37.52 ms | 16.3 MB | 114.9 KB | Sylvan.Data.Excel | 877.0% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table | OfficeIMO.Excel | 4.09 ms | 1.7 MB | 138.8 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-datareader-table | MiniExcel | 6.86 ms | 22.5 MB | 153.6 KB | OfficeIMO.Excel | 67.7% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table | EPPlus | 35.40 ms | 16.3 MB | 114.9 KB | OfficeIMO.Excel | 765.1% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table | ClosedXML | 37.70 ms | 18.6 MB | 120.9 KB | OfficeIMO.Excel | 821.2% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table-autofit | OfficeIMO.Excel | 4.11 ms | 1.7 MB | 138.9 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-datareader-table-autofit | MiniExcel | 7.22 ms | 26.0 MB | 153.8 KB | OfficeIMO.Excel | 75.8% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table-autofit | EPPlus | 54.54 ms | 37.4 MB | 115.1 KB | OfficeIMO.Excel | 1227.5% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table-autofit | ClosedXML | 72.64 ms | 57.0 MB | 121.0 KB | OfficeIMO.Excel | 1668.1% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-direct-export | LargeXlsx | 3.99 ms | 1.1 MB | 164.2 KB | LargeXlsx | 3.9% faster than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-direct-export | OfficeIMO.Excel | 4.16 ms | 2.1 MB | 131.1 KB | LargeXlsx | Loss +4.1% |
| 2500 | package-profile | write-dataset-sparse-direct-export | MiniExcel | 9.88 ms | 29.0 MB | 180.5 KB | LargeXlsx | 137.7% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-direct-export | ClosedXML | 52.16 ms | 26.8 MB | 159.4 KB | LargeXlsx | 1154.8% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-direct-export | EPPlus | 58.52 ms | 21.4 MB | 144.5 KB | LargeXlsx | 1307.8% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-tables | OfficeIMO.Excel | 4.48 ms | 2.8 MB | 176.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-dataset-sparse-tables | MiniExcel | 9.47 ms | 29.0 MB | 180.5 KB | OfficeIMO.Excel | 111.3% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-tables | ClosedXML | 53.16 ms | 26.8 MB | 159.4 KB | OfficeIMO.Excel | 1086.1% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-tables | EPPlus | 58.99 ms | 21.4 MB | 144.5 KB | OfficeIMO.Excel | 1216.1% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables | OfficeIMO.Excel | 4.01 ms | 2.0 MB | 138.9 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-dataset-tables | MiniExcel | 8.23 ms | 28.0 MB | 156.4 KB | OfficeIMO.Excel | 104.9% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables | ClosedXML | 34.94 ms | 18.4 MB | 123.4 KB | OfficeIMO.Excel | 770.3% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables | EPPlus | 39.68 ms | 18.3 MB | 116.6 KB | OfficeIMO.Excel | 888.2% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables-autofit | OfficeIMO.Excel | 4.69 ms | 2.0 MB | 139.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-dataset-tables-autofit | MiniExcel | 7.99 ms | 31.1 MB | 156.6 KB | OfficeIMO.Excel | 70.6% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables-autofit | EPPlus | 61.04 ms | 40.5 MB | 116.9 KB | OfficeIMO.Excel | 1202.6% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables-autofit | ClosedXML | 73.00 ms | 55.4 MB | 123.7 KB | OfficeIMO.Excel | 1457.8% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | LargeXlsx | 3.44 ms | 1.1 MB | 138.4 KB | LargeXlsx | 7.8% faster than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | OfficeIMO.Excel | 3.73 ms | 1.7 MB | 138.0 KB | LargeXlsx | Loss +8.5% |
| 2500 | package-profile | write-datatable-direct | MiniExcel | 7.56 ms | 22.5 MB | 153.7 KB | LargeXlsx | 102.6% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | ClosedXML | 27.52 ms | 11.3 MB | 120.1 KB | LargeXlsx | 637.3% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | EPPlus | 38.92 ms | 16.3 MB | 114.9 KB | LargeXlsx | 942.9% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-table-direct | OfficeIMO.Excel | 3.76 ms | 1.7 MB | 138.8 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-datatable-table-direct | MiniExcel | 7.53 ms | 22.5 MB | 153.7 KB | OfficeIMO.Excel | 100.5% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-table-direct | EPPlus | 37.17 ms | 16.3 MB | 114.9 KB | OfficeIMO.Excel | 889.5% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-table-direct | ClosedXML | 38.16 ms | 18.6 MB | 120.9 KB | OfficeIMO.Excel | 915.9% slower than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | LargeXlsx | 3.03 ms | 758.3 KB | 138.4 KB | LargeXlsx | 26.2% faster than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 4.10 ms | 1.7 MB | 138.0 KB | LargeXlsx | Loss +35.4% |
| 2500 | package-profile | write-fluent-rowsfrom-direct | MiniExcel | 7.48 ms | 22.7 MB | 153.7 KB | LargeXlsx | 82.5% slower than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | ClosedXML | 30.08 ms | 11.3 MB | 120.1 KB | LargeXlsx | 633.6% slower than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | EPPlus | 37.28 ms | 16.3 MB | 114.9 KB | LargeXlsx | 809.3% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.14 ms | 1.3 MB | 142.4 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 54.28 ms | 37.4 MB | 115.1 KB | OfficeIMO.Excel | 1209.9% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 66.16 ms | 49.7 MB | 120.2 KB | OfficeIMO.Excel | 1496.7% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | LargeXlsx | 3.06 ms | 758.3 KB | 138.4 KB | LargeXlsx | 13.5% faster than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | OfficeIMO.Excel | 3.53 ms | 1.3 MB | 142.3 KB | LargeXlsx | Loss +15.6% |
| 2500 | package-profile | write-insertobjects-direct | MiniExcel | 7.53 ms | 22.7 MB | 153.7 KB | LargeXlsx | 113.1% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | ClosedXML | 29.11 ms | 11.3 MB | 120.1 KB | LargeXlsx | 723.6% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | EPPlus | 38.49 ms | 16.3 MB | 114.9 KB | LargeXlsx | 989.0% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 5.25 ms | 1.5 MB | 138.1 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 56.38 ms | 37.4 MB | 115.1 KB | OfficeIMO.Excel | 974.3% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 65.83 ms | 49.7 MB | 120.2 KB | OfficeIMO.Excel | 1154.5% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 3.23 ms | 758.3 KB | 138.4 KB | LargeXlsx | 20.9% faster than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 4.09 ms | 1.5 MB | 138.0 KB | LargeXlsx | Loss +26.5% |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | MiniExcel | 7.89 ms | 22.7 MB | 153.7 KB | LargeXlsx | 93.0% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | ClosedXML | 28.56 ms | 11.3 MB | 120.1 KB | LargeXlsx | 598.3% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | EPPlus | 38.07 ms | 16.3 MB | 114.9 KB | LargeXlsx | 830.7% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | LargeXlsx | 3.15 ms | 758.3 KB | 138.4 KB | LargeXlsx | 33.2% faster than OfficeIMO |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | OfficeIMO.Excel | 4.72 ms | 1.7 MB | 142.3 KB | LargeXlsx | Loss +49.7% |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | MiniExcel | 7.45 ms | 22.7 MB | 153.7 KB | LargeXlsx | 57.7% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | ClosedXML | 30.94 ms | 11.3 MB | 120.1 KB | LargeXlsx | 555.5% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | EPPlus | 36.48 ms | 16.3 MB | 114.9 KB | LargeXlsx | 672.9% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 3.71 ms | 1.3 MB | 142.4 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 45.10 ms | 27.9 MB | 120.2 KB | OfficeIMO.Excel | 1116.6% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 47.16 ms | 26.7 MB | 115.0 KB | OfficeIMO.Excel | 1172.4% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-mixed-objects-direct | LargeXlsx | 4.03 ms | 802.5 KB | 182.6 KB | LargeXlsx | 27.2% faster than OfficeIMO |
| 2500 | package-profile | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 5.53 ms | 2.3 MB | 183.1 KB | LargeXlsx | Loss +37.4% |
| 2500 | package-profile | write-powershell-mixed-objects-direct | MiniExcel | 8.49 ms | 24.6 MB | 194.0 KB | LargeXlsx | 53.4% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-mixed-objects-direct | ClosedXML | 39.62 ms | 16.6 MB | 161.0 KB | LargeXlsx | 616.0% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-mixed-objects-direct | EPPlus | 51.41 ms | 19.6 MB | 152.1 KB | LargeXlsx | 829.0% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-mixed-direct | LargeXlsx | 4.03 ms | 802.5 KB | 182.6 KB | LargeXlsx | 9.0% faster than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 4.42 ms | 1.5 MB | 182.4 KB | LargeXlsx | Loss +9.9% |
| 2500 | package-profile | write-powershell-psobject-mixed-direct | MiniExcel | 8.14 ms | 24.6 MB | 194.0 KB | LargeXlsx | 84.0% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-mixed-direct | ClosedXML | 38.42 ms | 16.6 MB | 161.0 KB | LargeXlsx | 768.4% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-mixed-direct | EPPlus | 47.14 ms | 19.6 MB | 152.1 KB | LargeXlsx | 965.4% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 18.33 ms | 4.4 MB | 651.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-powershell-psobject-wide-direct | LargeXlsx | 20.07 ms | 2.7 MB | 644.6 KB | OfficeIMO.Excel | 9.5% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-wide-direct | MiniExcel | 35.39 ms | 47.3 MB | 674.4 KB | OfficeIMO.Excel | 93.1% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-wide-direct | ClosedXML | 119.38 ms | 50.4 MB | 615.5 KB | OfficeIMO.Excel | 551.5% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-wide-direct | EPPlus | 154.02 ms | 67.5 MB | 548.9 KB | OfficeIMO.Excel | 740.5% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | LargeXlsx | 1.38 ms | 296.4 KB |  | LargeXlsx | 36.2% faster than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | OfficeIMO.Excel | 2.16 ms | 1.5 MB |  | LargeXlsx | Loss +56.7% |
| 2500 | speed-comparison | append-plain-rows | MiniExcel | 3.97 ms | 19.2 MB |  | LargeXlsx | 83.4% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | ClosedXML | 16.46 ms | 10.9 MB |  | LargeXlsx | 661.5% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | EPPlus 4.5.3.3 | 20.58 ms | 0 B |  | LargeXlsx | 851.8% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | EPPlus | 24.70 ms | 14.0 MB |  | LargeXlsx | 1042.9% slower than OfficeIMO |
| 2500 | speed-comparison | autofit-existing | OfficeIMO.Excel | 8.19 ms | 1.9 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | autofit-existing | EPPlus | 73.54 ms | 49.5 MB |  | OfficeIMO.Excel | 797.9% slower than OfficeIMO |
| 2500 | speed-comparison | autofit-existing | EPPlus 4.5.3.3 | 95.04 ms | 0 B |  | OfficeIMO.Excel | 1060.4% slower than OfficeIMO |
| 2500 | speed-comparison | autofit-existing | ClosedXML | 126.85 ms | 82.9 MB |  | OfficeIMO.Excel | 1448.8% slower than OfficeIMO |
| 2500 | speed-comparison | build-object-datatable-dictionaries | OfficeIMO.Excel | 1.36 ms | 564.2 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | build-object-datatable-typed | OfficeIMO.Excel | 1.09 ms | 856.9 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-cells | OfficeIMO.Excel | 5.42 ms | 2.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-cells | EPPlus | 31.13 ms | 19.7 MB |  | OfficeIMO.Excel | 474.8% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-cells | ClosedXML | 31.24 ms | 16.6 MB |  | OfficeIMO.Excel | 476.8% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-first-column-from-wide-sheet | OfficeIMO.Excel | 3.62 ms | 523.4 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-first-column-from-wide-sheet | EPPlus | 23.93 ms | 12.8 MB |  | OfficeIMO.Excel | 561.5% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-first-column-from-wide-sheet | ClosedXML | 29.69 ms | 15.1 MB |  | OfficeIMO.Excel | 720.6% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-range | OfficeIMO.Excel | 5.88 ms | 2.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-range | EPPlus | 29.07 ms | 19.7 MB |  | OfficeIMO.Excel | 394.5% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-range | ClosedXML | 31.89 ms | 16.6 MB |  | OfficeIMO.Excel | 442.5% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-top-range | OfficeIMO.Excel | 0.78 ms | 285.2 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-top-range | EPPlus | 24.20 ms | 12.1 MB |  | OfficeIMO.Excel | 3017.3% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-top-range | ClosedXML | 31.31 ms | 15.0 MB |  | OfficeIMO.Excel | 3933.4% slower than OfficeIMO |
| 2500 | speed-comparison | formula-heavy-read | OfficeIMO.Excel | 2.14 ms | 706.6 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | formula-heavy-read | EPPlus 4.5.3.3 | 13.21 ms | 0 B |  | OfficeIMO.Excel | 516.3% slower than OfficeIMO |
| 2500 | speed-comparison | formula-heavy-read | ClosedXML | 17.34 ms | 8.1 MB |  | OfficeIMO.Excel | 708.8% slower than OfficeIMO |
| 2500 | speed-comparison | formula-heavy-read | EPPlus | 19.35 ms | 7.5 MB |  | OfficeIMO.Excel | 802.7% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | OfficeIMO.Excel | 2.04 ms | 2.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | large-shared-strings | MiniExcel | 3.80 ms | 20.6 MB |  | OfficeIMO.Excel | 86.2% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | EPPlus 4.5.3.3 | 9.86 ms | 0 B |  | OfficeIMO.Excel | 382.9% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | ClosedXML | 12.09 ms | 11.0 MB |  | OfficeIMO.Excel | 492.1% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | EPPlus | 21.62 ms | 12.5 MB |  | OfficeIMO.Excel | 958.4% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | OfficeIMO.Excel | 0.81 ms | 177.2 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | large-sparse-column-read | Sylvan.Data.Excel | 0.95 ms | 316.6 KB |  | OfficeIMO.Excel | 17.3% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | ExcelDataReader | 1.45 ms | 4.0 MB |  | OfficeIMO.Excel | 78.5% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | ClosedXML | 3.98 ms | 4.3 MB |  | OfficeIMO.Excel | 391.8% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | EPPlus 4.5.3.3 | 11.00 ms | 0 B |  | OfficeIMO.Excel | 1258.6% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | MiniExcel | 12.08 ms | 45.1 MB |  | OfficeIMO.Excel | 1391.7% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | EPPlus | 38.48 ms | 42.1 MB |  | OfficeIMO.Excel | 4650.2% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | Sylvan.Data.Excel | 0.96 ms | 316.6 KB |  | Sylvan.Data.Excel | 38.2% faster than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | ExcelDataReader | 1.44 ms | 4.0 MB |  | Sylvan.Data.Excel | 7.7% faster than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | OfficeIMO.Excel | 1.56 ms | 177.2 KB |  | Sylvan.Data.Excel | Loss +61.8% |
| 2500 | speed-comparison | large-sparse-row-read | ClosedXML | 4.10 ms | 4.3 MB |  | Sylvan.Data.Excel | 162.8% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | EPPlus 4.5.3.3 | 10.65 ms | 0 B |  | Sylvan.Data.Excel | 582.6% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | MiniExcel | 12.98 ms | 45.1 MB |  | Sylvan.Data.Excel | 732.4% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | EPPlus | 38.29 ms | 42.1 MB |  | Sylvan.Data.Excel | 2354.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | OfficeIMO.Excel | 4.52 ms | 374.5 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | read-bottom-range | Sylvan.Data.Excel | 4.70 ms | 655.2 KB |  | OfficeIMO.Excel | 4.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | ExcelDataReader | 10.34 ms | 5.9 MB |  | OfficeIMO.Excel | 128.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | MiniExcel | 15.51 ms | 18.2 MB |  | OfficeIMO.Excel | 243.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | EPPlus | 27.26 ms | 12.1 MB |  | OfficeIMO.Excel | 503.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | ClosedXML | 36.83 ms | 15.0 MB |  | OfficeIMO.Excel | 715.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | OfficeIMO.Excel | 3.99 ms | 377.7 KB |  | OfficeIMO.Excel, Sylvan.Data.Excel | Win |
| 2500 | speed-comparison | read-bottom-range-stream | Sylvan.Data.Excel | 4.04 ms | 655.2 KB |  | OfficeIMO.Excel, Sylvan.Data.Excel | Tie vs OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | ExcelDataReader | 10.09 ms | 5.9 MB |  | OfficeIMO.Excel, Sylvan.Data.Excel | 152.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | MiniExcel | 13.68 ms | 18.2 MB |  | OfficeIMO.Excel, Sylvan.Data.Excel | 242.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | EPPlus | 23.69 ms | 12.1 MB |  | OfficeIMO.Excel, Sylvan.Data.Excel | 493.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | ClosedXML | 30.60 ms | 15.0 MB |  | OfficeIMO.Excel, Sylvan.Data.Excel | 666.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | Sylvan.Data.Excel | 6.17 ms | 2.2 MB |  | Sylvan.Data.Excel | 23.8% faster than OfficeIMO |
| 2500 | speed-comparison | read-datatable | OfficeIMO.Excel | 8.09 ms | 3.5 MB |  | Sylvan.Data.Excel | Loss +31.2% |
| 2500 | speed-comparison | read-datatable | ExcelDataReader | 13.36 ms | 7.5 MB |  | Sylvan.Data.Excel | 65.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | MiniExcel | 15.85 ms | 17.8 MB |  | Sylvan.Data.Excel | 95.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | EPPlus | 35.13 ms | 21.2 MB |  | Sylvan.Data.Excel | 334.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | EPPlus 4.5.3.3 | 35.47 ms | 0 B |  | Sylvan.Data.Excel | 338.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | ClosedXML | 38.43 ms | 17.9 MB |  | Sylvan.Data.Excel | 375.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | OfficeIMO.Excel | 3.86 ms | 542.8 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | Sylvan.Data.Excel | 4.36 ms | 733.5 KB |  | OfficeIMO.Excel | 13.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | ExcelDataReader | 10.50 ms | 5.9 MB |  | OfficeIMO.Excel | 172.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | MiniExcel | 10.76 ms | 15.5 MB |  | OfficeIMO.Excel | 178.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | EPPlus | 23.88 ms | 12.8 MB |  | OfficeIMO.Excel | 518.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | ClosedXML | 30.10 ms | 15.1 MB |  | OfficeIMO.Excel | 680.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | OfficeIMO.Excel | 7.01 ms | 2.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | read-objects | Sylvan.Data.Excel | 7.16 ms | 895.3 KB |  | OfficeIMO.Excel | 2.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | ExcelDataReader | 13.68 ms | 6.2 MB |  | OfficeIMO.Excel | 95.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | MiniExcel | 14.31 ms | 18.0 MB |  | OfficeIMO.Excel | 104.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | EPPlus 4.5.3.3 | 32.53 ms | 0 B |  | OfficeIMO.Excel | 364.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | ClosedXML | 32.62 ms | 16.5 MB |  | OfficeIMO.Excel | 365.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | EPPlus | 34.50 ms | 20.9 MB |  | OfficeIMO.Excel | 392.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | OfficeIMO.Excel | 4.72 ms | 2.4 MB |  | OfficeIMO.Excel, Sylvan.Data.Excel | Win |
| 2500 | speed-comparison | read-objects-stream | Sylvan.Data.Excel | 4.73 ms | 831.0 KB |  | OfficeIMO.Excel, Sylvan.Data.Excel | Tie vs OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | ExcelDataReader | 10.48 ms | 6.1 MB |  | OfficeIMO.Excel, Sylvan.Data.Excel | 121.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | MiniExcel | 13.56 ms | 18.0 MB |  | OfficeIMO.Excel, Sylvan.Data.Excel | 187.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | EPPlus 4.5.3.3 | 29.22 ms | 0 B |  | OfficeIMO.Excel, Sylvan.Data.Excel | 518.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | ClosedXML | 30.23 ms | 16.5 MB |  | OfficeIMO.Excel, Sylvan.Data.Excel | 540.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | EPPlus | 34.75 ms | 20.8 MB |  | OfficeIMO.Excel, Sylvan.Data.Excel | 635.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | OfficeIMO.Excel | 9.03 ms | 2.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | read-range | Sylvan.Data.Excel | 15.52 ms | 655.0 KB |  | OfficeIMO.Excel | 71.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | ExcelDataReader | 23.31 ms | 5.9 MB |  | OfficeIMO.Excel | 158.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | MiniExcel | 23.87 ms | 18.2 MB |  | OfficeIMO.Excel | 164.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | EPPlus 4.5.3.3 | 35.79 ms | 0 B |  | OfficeIMO.Excel | 296.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | EPPlus | 37.27 ms | 19.7 MB |  | OfficeIMO.Excel | 312.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | ClosedXML | 78.00 ms | 16.5 MB |  | OfficeIMO.Excel | 763.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | OfficeIMO.Excel | 4.93 ms | 2.7 MB |  | OfficeIMO.Excel, Sylvan.Data.Excel | Win |
| 2500 | speed-comparison | read-range-decimal | Sylvan.Data.Excel | 5.02 ms | 750.3 KB |  | OfficeIMO.Excel, Sylvan.Data.Excel | Tie vs OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | ExcelDataReader | 10.30 ms | 5.9 MB |  | OfficeIMO.Excel, Sylvan.Data.Excel | 108.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | MiniExcel | 12.48 ms | 18.2 MB |  | OfficeIMO.Excel, Sylvan.Data.Excel | 153.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | EPPlus | 29.09 ms | 19.7 MB |  | OfficeIMO.Excel, Sylvan.Data.Excel | 490.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | ClosedXML | 31.46 ms | 16.3 MB |  | OfficeIMO.Excel, Sylvan.Data.Excel | 538.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | Sylvan.Data.Excel | 4.40 ms | 655.2 KB |  | Sylvan.Data.Excel | 20.3% faster than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | OfficeIMO.Excel | 5.52 ms | 2.7 MB |  | Sylvan.Data.Excel | Loss +25.4% |
| 2500 | speed-comparison | read-range-stream | ExcelDataReader | 11.03 ms | 5.9 MB |  | Sylvan.Data.Excel | 99.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | MiniExcel | 12.18 ms | 18.2 MB |  | Sylvan.Data.Excel | 120.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | ClosedXML | 29.61 ms | 16.3 MB |  | Sylvan.Data.Excel | 435.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | EPPlus 4.5.3.3 | 29.77 ms | 0 B |  | Sylvan.Data.Excel | 438.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | EPPlus | 31.45 ms | 19.7 MB |  | Sylvan.Data.Excel | 469.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | Sylvan.Data.Excel | 0.44 ms | 348.4 KB |  | Sylvan.Data.Excel | 17.7% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range | OfficeIMO.Excel | 0.53 ms | 296.0 KB |  | Sylvan.Data.Excel | Loss +21.6% |
| 2500 | speed-comparison | read-top-range | MiniExcel | 0.76 ms | 869.0 KB |  | Sylvan.Data.Excel | 42.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | ExcelDataReader | 4.28 ms | 1.9 MB |  | Sylvan.Data.Excel | 707.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | EPPlus | 23.97 ms | 12.1 MB |  | Sylvan.Data.Excel | 4426.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | EPPlus 4.5.3.3 | 28.96 ms | 0 B |  | Sylvan.Data.Excel | 5370.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | ClosedXML | 31.47 ms | 15.0 MB |  | Sylvan.Data.Excel | 5844.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | Sylvan.Data.Excel | 0.43 ms | 348.5 KB |  | Sylvan.Data.Excel | 21.2% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | OfficeIMO.Excel | 0.54 ms | 299.4 KB |  | Sylvan.Data.Excel | Loss +26.9% |
| 2500 | speed-comparison | read-top-range-stream | MiniExcel | 0.79 ms | 869.0 KB |  | Sylvan.Data.Excel | 44.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | ExcelDataReader | 4.36 ms | 1.9 MB |  | Sylvan.Data.Excel | 702.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | EPPlus | 23.88 ms | 12.1 MB |  | Sylvan.Data.Excel | 4298.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | EPPlus 4.5.3.3 | 25.78 ms | 0 B |  | Sylvan.Data.Excel | 4647.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | ClosedXML | 31.10 ms | 15.0 MB |  | Sylvan.Data.Excel | 5628.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | Sylvan.Data.Excel | 0.40 ms | 348.5 KB |  | Sylvan.Data.Excel | 22.6% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | OfficeIMO.Excel | 0.52 ms | 300.0 KB |  | Sylvan.Data.Excel | Loss +29.2% |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | MiniExcel | 0.97 ms | 869.0 KB |  | Sylvan.Data.Excel | 86.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | ExcelDataReader | 4.34 ms | 1.9 MB |  | Sylvan.Data.Excel | 732.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | EPPlus | 23.53 ms | 12.1 MB |  | Sylvan.Data.Excel | 4408.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | ClosedXML | 29.79 ms | 15.0 MB |  | Sylvan.Data.Excel | 5608.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-used-range | Sylvan.Data.Excel | 4.52 ms | 655.2 KB |  | Sylvan.Data.Excel | 68.4% faster than OfficeIMO |
| 2500 | speed-comparison | read-used-range | ExcelDataReader | 10.88 ms | 5.9 MB |  | Sylvan.Data.Excel | 24.0% faster than OfficeIMO |
| 2500 | speed-comparison | read-used-range | MiniExcel | 11.98 ms | 18.2 MB |  | Sylvan.Data.Excel | 16.3% faster than OfficeIMO |
| 2500 | speed-comparison | read-used-range | OfficeIMO.Excel | 14.31 ms | 3.4 MB |  | Sylvan.Data.Excel | Loss +216.7% |
| 2500 | speed-comparison | read-used-range | EPPlus | 29.62 ms | 19.7 MB |  | Sylvan.Data.Excel | 107.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-used-range | ClosedXML | 67.87 ms | 16.4 MB |  | Sylvan.Data.Excel | 374.3% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-autofilter | OfficeIMO.Excel | 3.62 ms | 1.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-autofilter | ClosedXML | 30.89 ms | 21.7 MB |  | OfficeIMO.Excel | 752.6% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-autofilter | EPPlus 4.5.3.3 | 35.14 ms | 0 B |  | OfficeIMO.Excel | 869.9% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-autofilter | EPPlus | 41.08 ms | 24.1 MB |  | OfficeIMO.Excel | 1034.0% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-charts | OfficeIMO.Excel | 5.03 ms | 1.8 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-charts | EPPlus 4.5.3.3 | 37.08 ms | 0 B |  | OfficeIMO.Excel | 637.6% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-charts | EPPlus | 41.05 ms | 26.5 MB |  | OfficeIMO.Excel | 716.6% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-conditional-formatting | OfficeIMO.Excel | 3.59 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-conditional-formatting | ClosedXML | 31.85 ms | 21.8 MB |  | OfficeIMO.Excel | 786.7% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-conditional-formatting | EPPlus 4.5.3.3 | 36.01 ms | 0 B |  | OfficeIMO.Excel | 902.2% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-conditional-formatting | EPPlus | 40.15 ms | 24.2 MB |  | OfficeIMO.Excel | 1017.7% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-data-validation | OfficeIMO.Excel | 3.66 ms | 1.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-data-validation | ClosedXML | 30.28 ms | 21.7 MB |  | OfficeIMO.Excel | 727.9% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-data-validation | EPPlus 4.5.3.3 | 34.68 ms | 0 B |  | OfficeIMO.Excel | 848.2% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-data-validation | EPPlus | 38.32 ms | 24.1 MB |  | OfficeIMO.Excel | 947.7% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-freeze-panes | OfficeIMO.Excel | 3.65 ms | 1.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-freeze-panes | ClosedXML | 32.77 ms | 21.7 MB |  | OfficeIMO.Excel | 796.9% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-freeze-panes | EPPlus 4.5.3.3 | 34.24 ms | 0 B |  | OfficeIMO.Excel | 837.1% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-freeze-panes | EPPlus | 42.07 ms | 24.2 MB |  | OfficeIMO.Excel | 1051.4% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-pivot-table | OfficeIMO.Excel | 14.05 ms | 14.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-pivot-table | EPPlus 4.5.3.3 | 35.46 ms | 0 B |  | OfficeIMO.Excel | 152.5% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-pivot-table | EPPlus | 44.12 ms | 28.8 MB |  | OfficeIMO.Excel | 214.1% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-all-in-one | OfficeIMO.Excel | 16.82 ms | 14.9 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-all-in-one | EPPlus | 72.15 ms | 53.3 MB |  | OfficeIMO.Excel | 329.0% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-all-in-one | EPPlus 4.5.3.3 | 76.93 ms | 0 B |  | OfficeIMO.Excel | 357.5% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-chart-first | OfficeIMO.Excel | 10.76 ms | 6.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-chart-first | EPPlus | 73.65 ms | 53.3 MB |  | OfficeIMO.Excel | 584.1% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-chart-first | EPPlus 4.5.3.3 | 74.45 ms | 0 B |  | OfficeIMO.Excel | 591.6% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-core | OfficeIMO.Excel | 4.27 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-core | EPPlus 4.5.3.3 | 60.12 ms | 0 B |  | OfficeIMO.Excel | 1308.2% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-core | EPPlus | 65.98 ms | 46.2 MB |  | OfficeIMO.Excel | 1445.5% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-core | ClosedXML | 78.52 ms | 68.2 MB |  | OfficeIMO.Excel | 1739.3% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-extra-column | OfficeIMO.Excel | 17.92 ms | 16.0 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-extra-column | EPPlus | 73.98 ms | 57.8 MB |  | OfficeIMO.Excel | 312.7% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-extra-column | EPPlus 4.5.3.3 | 77.86 ms | 0 B |  | OfficeIMO.Excel | 334.4% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-no-autofit | OfficeIMO.Excel | 17.25 ms | 14.9 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-no-autofit | EPPlus 4.5.3.3 | 42.81 ms | 0 B |  | OfficeIMO.Excel | 148.2% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-no-autofit | EPPlus | 49.27 ms | 32.1 MB |  | OfficeIMO.Excel | 185.6% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-post-mutation | OfficeIMO.Excel | 16.05 ms | 14.9 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-post-mutation | EPPlus 4.5.3.3 | 61.57 ms | 0 B |  | OfficeIMO.Excel | 283.5% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-post-mutation | EPPlus | 68.85 ms | 53.3 MB |  | OfficeIMO.Excel | 328.9% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-shuffled-columns | OfficeIMO.Excel | 18.01 ms | 14.9 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-shuffled-columns | EPPlus 4.5.3.3 | 72.43 ms | 0 B |  | OfficeIMO.Excel | 302.1% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-shuffled-columns | EPPlus | 72.89 ms | 53.3 MB |  | OfficeIMO.Excel | 304.7% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook | OfficeIMO.Excel | 29.93 ms | 18.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | report-workbook | EPPlus 4.5.3.3 | 77.70 ms | 0 B |  | OfficeIMO.Excel | 159.6% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook | EPPlus | 94.07 ms | 75.7 MB |  | OfficeIMO.Excel | 214.3% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-core | OfficeIMO.Excel | 6.41 ms | 2.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | report-workbook-core | EPPlus 4.5.3.3 | 70.56 ms | 0 B |  | OfficeIMO.Excel | 1000.7% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-core | EPPlus | 97.70 ms | 70.3 MB |  | OfficeIMO.Excel | 1423.9% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-core | ClosedXML | 103.64 ms | 94.9 MB |  | OfficeIMO.Excel | 1516.6% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-datatable | OfficeIMO.Excel | 23.96 ms | 18.9 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | report-workbook-datatable | EPPlus 4.5.3.3 | 67.32 ms | 0 B |  | OfficeIMO.Excel | 180.9% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-datatable | EPPlus | 102.61 ms | 64.4 MB |  | OfficeIMO.Excel | 328.2% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-datatable-core | OfficeIMO.Excel | 6.28 ms | 2.9 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | report-workbook-datatable-core | EPPlus 4.5.3.3 | 63.60 ms | 0 B |  | OfficeIMO.Excel | 913.4% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-datatable-core | EPPlus | 97.86 ms | 59.1 MB |  | OfficeIMO.Excel | 1459.3% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-datatable-core | ClosedXML | 102.13 ms | 80.9 MB |  | OfficeIMO.Excel | 1527.4% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | Sylvan.Data.Excel | 2.04 ms | 518.6 KB |  | Sylvan.Data.Excel | 42.5% faster than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | OfficeIMO.Excel | 3.56 ms | 1.0 MB |  | Sylvan.Data.Excel | Loss +73.9% |
| 2500 | speed-comparison | shared-string-read | ExcelDataReader | 4.69 ms | 2.6 MB |  | Sylvan.Data.Excel | 31.9% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | MiniExcel | 5.45 ms | 7.4 MB |  | Sylvan.Data.Excel | 53.2% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | EPPlus 4.5.3.3 | 11.88 ms | 0 B |  | Sylvan.Data.Excel | 234.2% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | ClosedXML | 15.84 ms | 9.3 MB |  | Sylvan.Data.Excel | 345.4% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | EPPlus | 20.07 ms | 10.1 MB |  | Sylvan.Data.Excel | 464.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-blog-2023-20-string-columns | OfficeIMO.Excel | 4.76 ms | 1.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-blog-2023-20-string-columns | LargeXlsx | 7.44 ms | 857.6 KB |  | OfficeIMO.Excel | 56.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-blog-2023-20-string-columns | MiniExcel | 18.50 ms | 35.1 MB |  | OfficeIMO.Excel | 288.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-blog-2023-20-string-columns | ClosedXML | 92.91 ms | 69.8 MB |  | OfficeIMO.Excel | 1850.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | OfficeIMO.Excel | 6.22 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-bulk-report | MiniExcel | 15.69 ms | 26.2 MB |  | OfficeIMO.Excel | 152.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | EPPlus 4.5.3.3 | 110.55 ms | 0 B |  | OfficeIMO.Excel | 1676.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | EPPlus | 126.60 ms | 48.0 MB |  | OfficeIMO.Excel | 1934.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | ClosedXML | 225.25 ms | 57.0 MB |  | OfficeIMO.Excel | 3519.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellformula | OfficeIMO.Excel | 3.13 ms | 1.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellformula | ClosedXML | 18.98 ms | 11.8 MB |  | OfficeIMO.Excel | 505.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellformula | EPPlus 4.5.3.3 | 22.73 ms | 0 B |  | OfficeIMO.Excel | 625.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellformula | EPPlus | 36.93 ms | 17.7 MB |  | OfficeIMO.Excel | 1078.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-empty-strings | OfficeIMO.Excel | 2.50 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-empty-strings | ClosedXML | 12.34 ms | 9.7 MB |  | OfficeIMO.Excel | 393.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-empty-strings | EPPlus | 23.29 ms | 11.5 MB |  | OfficeIMO.Excel | 831.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-numbers | OfficeIMO.Excel | 2.60 ms | 1.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-numbers | ClosedXML | 11.62 ms | 9.0 MB |  | OfficeIMO.Excel | 346.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-numbers | EPPlus 4.5.3.3 | 18.22 ms | 0 B |  | OfficeIMO.Excel | 599.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-numbers | EPPlus | 22.11 ms | 12.6 MB |  | OfficeIMO.Excel | 749.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-mixed | OfficeIMO.Excel | 3.11 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-object-mixed | ClosedXML | 17.41 ms | 11.6 MB |  | OfficeIMO.Excel | 459.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-mixed | EPPlus 4.5.3.3 | 22.40 ms | 0 B |  | OfficeIMO.Excel | 619.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-mixed | EPPlus | 24.72 ms | 15.3 MB |  | OfficeIMO.Excel | 693.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-sparse | OfficeIMO.Excel | 3.15 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-object-sparse | ClosedXML | 14.82 ms | 11.0 MB |  | OfficeIMO.Excel | 370.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-sparse | EPPlus | 27.15 ms | 14.6 MB |  | OfficeIMO.Excel | 761.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 3.00 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-object-sparse-batch | ClosedXML | 15.83 ms | 11.0 MB |  | OfficeIMO.Excel | 427.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-sparse-batch | EPPlus | 27.97 ms | 14.6 MB |  | OfficeIMO.Excel | 832.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-scalars | OfficeIMO.Excel | 2.68 ms | 1.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-scalars | ClosedXML | 10.76 ms | 8.8 MB |  | OfficeIMO.Excel | 302.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-scalars | EPPlus 4.5.3.3 | 18.36 ms | 0 B |  | OfficeIMO.Excel | 586.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-scalars | EPPlus | 22.72 ms | 12.5 MB |  | OfficeIMO.Excel | 749.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings | OfficeIMO.Excel | 3.17 ms | 2.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-strings | ClosedXML | 13.86 ms | 11.0 MB |  | OfficeIMO.Excel | 337.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings | EPPlus 4.5.3.3 | 20.63 ms | 0 B |  | OfficeIMO.Excel | 550.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings | EPPlus | 20.89 ms | 12.5 MB |  | OfficeIMO.Excel | 558.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings-distinct | OfficeIMO.Excel | 2.48 ms | 2.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-strings-distinct | ClosedXML | 14.71 ms | 12.8 MB |  | OfficeIMO.Excel | 492.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings-distinct | EPPlus | 23.99 ms | 13.6 MB |  | OfficeIMO.Excel | 866.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings-repeated | OfficeIMO.Excel | 2.27 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-strings-repeated | ClosedXML | 10.98 ms | 9.0 MB |  | OfficeIMO.Excel | 384.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings-repeated | EPPlus | 19.15 ms | 11.1 MB |  | OfficeIMO.Excel | 745.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-temporal | OfficeIMO.Excel | 3.26 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-temporal | ClosedXML | 15.70 ms | 9.5 MB |  | OfficeIMO.Excel | 381.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-temporal | EPPlus 4.5.3.3 | 20.54 ms | 0 B |  | OfficeIMO.Excel | 529.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-temporal | EPPlus | 24.11 ms | 14.4 MB |  | OfficeIMO.Excel | 638.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 1.58 ms | 1.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 2.93 ms | 447.0 KB |  | OfficeIMO.Excel | 85.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-headerless-rectangle-direct | ClosedXML | 15.37 ms | 10.0 MB |  | OfficeIMO.Excel | 875.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-headerless-rectangle-direct | EPPlus | 23.31 ms | 12.7 MB |  | OfficeIMO.Excel | 1379.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 4.33 ms | 2.0 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | LargeXlsx | 7.05 ms | 758.3 KB |  | OfficeIMO.Excel | 62.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | MiniExcel | 8.87 ms | 22.7 MB |  | OfficeIMO.Excel | 104.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | ClosedXML | 33.24 ms | 21.7 MB |  | OfficeIMO.Excel | 667.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 36.64 ms | 0 B |  | OfficeIMO.Excel | 745.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus | 43.54 ms | 24.1 MB |  | OfficeIMO.Excel | 904.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-sparse-rectangle-direct | OfficeIMO.Excel | 3.38 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalues-sparse-rectangle-direct | ClosedXML | 15.37 ms | 11.0 MB |  | OfficeIMO.Excel | 354.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-sparse-rectangle-direct | EPPlus | 25.35 ms | 14.6 MB |  | OfficeIMO.Excel | 650.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | Sylvan.Data.Excel | 3.61 ms | 758.6 KB |  | Sylvan.Data.Excel | 22.9% faster than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | OfficeIMO.Excel | 4.68 ms | 1.7 MB |  | Sylvan.Data.Excel | Loss +29.7% |
| 2500 | speed-comparison | write-datareader-plain | MiniExcel | 7.46 ms | 22.5 MB |  | Sylvan.Data.Excel | 59.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | LargeXlsx | 8.21 ms | 1.0 MB |  | Sylvan.Data.Excel | 75.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | ClosedXML | 26.90 ms | 11.3 MB |  | Sylvan.Data.Excel | 475.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | EPPlus 4.5.3.3 | 33.60 ms | 0 B |  | Sylvan.Data.Excel | 618.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | EPPlus | 34.50 ms | 16.3 MB |  | Sylvan.Data.Excel | 637.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | OfficeIMO.Excel | 4.39 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datareader-table | MiniExcel | 8.93 ms | 22.5 MB |  | OfficeIMO.Excel | 103.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | EPPlus | 35.16 ms | 16.3 MB |  | OfficeIMO.Excel | 701.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | ClosedXML | 35.72 ms | 18.6 MB |  | OfficeIMO.Excel | 714.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | EPPlus 4.5.3.3 | 38.37 ms | 0 B |  | OfficeIMO.Excel | 774.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table-autofit | OfficeIMO.Excel | 5.10 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datareader-table-autofit | MiniExcel | 8.19 ms | 26.0 MB |  | OfficeIMO.Excel | 60.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table-autofit | EPPlus | 55.09 ms | 37.4 MB |  | OfficeIMO.Excel | 980.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table-autofit | EPPlus 4.5.3.3 | 72.00 ms | 0 B |  | OfficeIMO.Excel | 1311.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table-autofit | ClosedXML | 75.69 ms | 57.0 MB |  | OfficeIMO.Excel | 1383.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-headerless-tables | OfficeIMO.Excel | 5.26 ms | 2.0 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-headerless-tables | MiniExcel | 14.35 ms | 28.5 MB |  | OfficeIMO.Excel | 173.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-headerless-tables | ClosedXML | 42.13 ms | 18.5 MB |  | OfficeIMO.Excel | 701.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-headerless-tables | EPPlus | 71.86 ms | 18.0 MB |  | OfficeIMO.Excel | 1267.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | OfficeIMO.Excel | 5.69 ms | 2.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | LargeXlsx | 8.87 ms | 1.1 MB |  | OfficeIMO.Excel | 55.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | MiniExcel | 15.63 ms | 29.8 MB |  | OfficeIMO.Excel | 174.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | EPPlus | 56.81 ms | 21.6 MB |  | OfficeIMO.Excel | 898.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | ClosedXML | 59.53 ms | 26.8 MB |  | OfficeIMO.Excel | 946.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-tables | OfficeIMO.Excel | 6.70 ms | 2.8 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-sparse-tables | MiniExcel | 18.22 ms | 29.8 MB |  | OfficeIMO.Excel | 172.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-tables | ClosedXML | 60.91 ms | 26.8 MB |  | OfficeIMO.Excel | 809.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-tables | EPPlus | 67.57 ms | 22.0 MB |  | OfficeIMO.Excel | 909.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | OfficeIMO.Excel | 4.93 ms | 2.0 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-tables | MiniExcel | 14.69 ms | 28.0 MB |  | OfficeIMO.Excel | 198.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | EPPlus 4.5.3.3 | 40.60 ms | 0 B |  | OfficeIMO.Excel | 723.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | ClosedXML | 76.68 ms | 18.4 MB |  | OfficeIMO.Excel | 1455.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | EPPlus | 78.63 ms | 19.0 MB |  | OfficeIMO.Excel | 1495.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables-autofit | OfficeIMO.Excel | 5.49 ms | 2.0 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-tables-autofit | MiniExcel | 15.64 ms | 31.4 MB |  | OfficeIMO.Excel | 185.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables-autofit | EPPlus | 137.37 ms | 42.4 MB |  | OfficeIMO.Excel | 2403.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables-autofit | ClosedXML | 190.94 ms | 55.4 MB |  | OfficeIMO.Excel | 3380.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | OfficeIMO.Excel | 5.04 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datatable-direct | LargeXlsx | 6.80 ms | 1.1 MB |  | OfficeIMO.Excel | 35.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | MiniExcel | 11.17 ms | 22.9 MB |  | OfficeIMO.Excel | 121.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | ClosedXML | 27.12 ms | 11.3 MB |  | OfficeIMO.Excel | 438.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | EPPlus | 36.65 ms | 16.3 MB |  | OfficeIMO.Excel | 627.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | EPPlus 4.5.3.3 | 39.39 ms | 0 B |  | OfficeIMO.Excel | 681.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-object-table-direct | OfficeIMO.Excel | 4.53 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datatable-object-table-direct | MiniExcel | 8.76 ms | 22.3 MB |  | OfficeIMO.Excel | 93.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-object-table-direct | EPPlus | 34.01 ms | 16.0 MB |  | OfficeIMO.Excel | 651.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-object-table-direct | ClosedXML | 34.51 ms | 18.3 MB |  | OfficeIMO.Excel | 661.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | OfficeIMO.Excel | 4.40 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datatable-table-direct | MiniExcel | 9.00 ms | 22.5 MB |  | OfficeIMO.Excel | 104.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | EPPlus | 35.55 ms | 16.3 MB |  | OfficeIMO.Excel | 708.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | ClosedXML | 35.62 ms | 18.6 MB |  | OfficeIMO.Excel | 710.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | EPPlus 4.5.3.3 | 38.88 ms | 0 B |  | OfficeIMO.Excel | 784.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-dictionary-objects-table-direct | OfficeIMO.Excel | 5.46 ms | 2.0 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | LargeXlsx | 3.27 ms | 758.3 KB |  | LargeXlsx | 15.3% faster than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 3.86 ms | 1.7 MB |  | LargeXlsx | Loss +18.1% |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | MiniExcel | 7.93 ms | 22.7 MB |  | LargeXlsx | 105.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | ClosedXML | 27.35 ms | 11.3 MB |  | LargeXlsx | 608.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 41.48 ms | 0 B |  | LargeXlsx | 974.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus | 41.80 ms | 16.3 MB |  | LargeXlsx | 982.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 3.96 ms | 1.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 52.62 ms | 37.4 MB |  | OfficeIMO.Excel | 1227.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 63.25 ms | 49.7 MB |  | OfficeIMO.Excel | 1495.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | LargeXlsx | 3.16 ms | 758.3 KB |  | LargeXlsx | 15.0% faster than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | OfficeIMO.Excel | 3.71 ms | 1.3 MB |  | LargeXlsx | Loss +17.7% |
| 2500 | speed-comparison | write-insertobjects-direct | MiniExcel | 8.79 ms | 22.7 MB |  | LargeXlsx | 136.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | ClosedXML | 26.91 ms | 11.3 MB |  | LargeXlsx | 624.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | EPPlus 4.5.3.3 | 39.61 ms | 0 B |  | LargeXlsx | 966.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | EPPlus | 39.75 ms | 16.3 MB |  | LargeXlsx | 970.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.52 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 52.76 ms | 37.4 MB |  | OfficeIMO.Excel | 1067.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 65.95 ms | 49.7 MB |  | OfficeIMO.Excel | 1359.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 3.21 ms | 758.3 KB |  | LargeXlsx | 25.8% faster than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 4.33 ms | 1.5 MB |  | LargeXlsx | Loss +34.7% |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | MiniExcel | 7.98 ms | 22.7 MB |  | LargeXlsx | 84.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | ClosedXML | 28.38 ms | 11.3 MB |  | LargeXlsx | 555.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | EPPlus | 42.13 ms | 16.3 MB |  | LargeXlsx | 873.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.06 ms | 1.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 44.44 ms | 27.9 MB |  | OfficeIMO.Excel | 993.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 45.60 ms | 26.7 MB |  | OfficeIMO.Excel | 1021.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | LargeXlsx | 4.15 ms | 802.5 KB |  | LargeXlsx | 23.5% faster than OfficeIMO |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 5.42 ms | 2.3 MB |  | LargeXlsx | Loss +30.7% |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | MiniExcel | 8.67 ms | 24.6 MB |  | LargeXlsx | 60.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | ClosedXML | 36.99 ms | 16.6 MB |  | LargeXlsx | 582.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | EPPlus | 47.39 ms | 19.6 MB |  | LargeXlsx | 774.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-mixed-direct | LargeXlsx | 4.13 ms | 802.5 KB |  | LargeXlsx | 21.1% faster than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 5.24 ms | 1.5 MB |  | LargeXlsx | Loss +26.7% |
| 2500 | speed-comparison | write-powershell-psobject-mixed-direct | MiniExcel | 8.41 ms | 24.6 MB |  | LargeXlsx | 60.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-mixed-direct | ClosedXML | 38.16 ms | 16.6 MB |  | LargeXlsx | 628.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-mixed-direct | EPPlus | 47.12 ms | 19.6 MB |  | LargeXlsx | 799.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-wide-direct | LargeXlsx | 20.03 ms | 2.7 MB |  | LargeXlsx | 9.8% faster than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 22.20 ms | 4.4 MB |  | LargeXlsx | Loss +10.8% |
| 2500 | speed-comparison | write-powershell-psobject-wide-direct | MiniExcel | 34.75 ms | 47.3 MB |  | LargeXlsx | 56.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-wide-direct | ClosedXML | 119.19 ms | 50.4 MB |  | LargeXlsx | 436.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-wide-direct | EPPlus | 156.73 ms | 67.5 MB |  | LargeXlsx | 606.0% slower than OfficeIMO |
| 10000 | focused-package-profile | write-blog-2023-20-string-columns | OfficeIMO.Excel | 33.01 ms | 7.6 MB | 880.4 KB | OfficeIMO.Excel | Win |
| 10000 | focused-package-profile | write-blog-2023-20-string-columns | LargeXlsx | 88.91 ms | 3.1 MB | 970.2 KB | OfficeIMO.Excel | 2.69x vs best |
| 10000 | focused-package-profile | write-blog-2023-20-string-columns | MiniExcel | 142.47 ms | 96.2 MB | 957.6 KB | OfficeIMO.Excel | 4.32x vs best |
| 10000 | focused-package-profile | write-blog-2023-20-string-columns | ClosedXML | 725.30 ms | 280.2 MB | 1,015.4 KB | OfficeIMO.Excel | 21.97x vs best |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | Sylvan.Data.Excel | 73.68 ms | 394.1 KB |  | Sylvan.Data.Excel | 9.2% faster than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | OfficeIMO.Excel | 81.10 ms | 23.1 MB |  | Sylvan.Data.Excel | Loss +10.1% |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | MiniExcel | 194.48 ms | 210.3 MB |  | Sylvan.Data.Excel | 139.8% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | ExcelDataReader | 206.19 ms | 67.9 MB |  | Sylvan.Data.Excel | 154.2% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | Sylvan.Data.Excel | 40.87 ms | 394.1 KB |  | Sylvan.Data.Excel | 23.4% faster than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | OfficeIMO.Excel | 53.33 ms | 23.8 MB |  | Sylvan.Data.Excel | Loss +30.5% |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | ExcelDataReader | 124.69 ms | 67.9 MB |  | Sylvan.Data.Excel | 133.8% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | MiniExcel | 165.54 ms | 210.3 MB |  | Sylvan.Data.Excel | 210.4% slower than OfficeIMO |
| 25000 | package-profile | append-plain-rows | LargeXlsx | 11.83 ms | 2.7 MB | 605.0 KB | LargeXlsx | 17.2% faster than OfficeIMO |
| 25000 | package-profile | append-plain-rows | OfficeIMO.Excel | 14.28 ms | 10.6 MB | 610.4 KB | LargeXlsx | Loss +20.8% |
| 25000 | package-profile | append-plain-rows | MiniExcel | 30.72 ms | 56.9 MB | 642.3 KB | LargeXlsx | 115.1% slower than OfficeIMO |
| 25000 | package-profile | append-plain-rows | ClosedXML | 136.51 ms | 101.8 MB | 540.6 KB | LargeXlsx | 855.8% slower than OfficeIMO |
| 25000 | package-profile | append-plain-rows | EPPlus | 206.42 ms | 98.0 MB | 525.6 KB | LargeXlsx | 1345.3% slower than OfficeIMO |
| 25000 | package-profile | autofit-existing | OfficeIMO.Excel | 81.97 ms | 15.3 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | autofit-existing | EPPlus | 451.56 ms | 245.1 MB | 1.1 MB | OfficeIMO.Excel | 450.9% slower than OfficeIMO |
| 25000 | package-profile | autofit-existing | ClosedXML | 1.33 s | 810.3 MB | 1.1 MB | OfficeIMO.Excel | 1524.7% slower than OfficeIMO |
| 25000 | package-profile | large-shared-strings | OfficeIMO.Excel | 23.48 ms | 15.4 MB | 529.7 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | large-shared-strings | MiniExcel | 50.51 ms | 72.0 MB | 581.0 KB | OfficeIMO.Excel | 115.1% slower than OfficeIMO |
| 25000 | package-profile | large-shared-strings | ClosedXML | 169.37 ms | 101.8 MB | 460.1 KB | OfficeIMO.Excel | 621.2% slower than OfficeIMO |
| 25000 | package-profile | large-shared-strings | EPPlus | 308.49 ms | 82.4 MB | 444.7 KB | OfficeIMO.Excel | 1213.5% slower than OfficeIMO |
| 25000 | package-profile | realworld-autofilter | OfficeIMO.Excel | 32.88 ms | 11.2 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-autofilter | ClosedXML | 306.11 ms | 205.7 MB | 1.1 MB | OfficeIMO.Excel | 830.9% slower than OfficeIMO |
| 25000 | package-profile | realworld-autofilter | EPPlus | 420.88 ms | 206.9 MB | 1.1 MB | OfficeIMO.Excel | 1179.9% slower than OfficeIMO |
| 25000 | package-profile | realworld-charts | OfficeIMO.Excel | 34.79 ms | 12.3 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-charts | EPPlus | 447.66 ms | 209.9 MB | 1.1 MB | OfficeIMO.Excel | 1186.8% slower than OfficeIMO |
| 25000 | package-profile | realworld-conditional-formatting | OfficeIMO.Excel | 32.38 ms | 11.3 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-conditional-formatting | ClosedXML | 311.49 ms | 205.8 MB | 1.1 MB | OfficeIMO.Excel | 861.9% slower than OfficeIMO |
| 25000 | package-profile | realworld-conditional-formatting | EPPlus | 421.72 ms | 206.9 MB | 1.1 MB | OfficeIMO.Excel | 1202.3% slower than OfficeIMO |
| 25000 | package-profile | realworld-data-validation | OfficeIMO.Excel | 35.11 ms | 11.2 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-data-validation | ClosedXML | 334.51 ms | 205.7 MB | 1.1 MB | OfficeIMO.Excel | 852.6% slower than OfficeIMO |
| 25000 | package-profile | realworld-data-validation | EPPlus | 428.30 ms | 206.9 MB | 1.1 MB | OfficeIMO.Excel | 1119.7% slower than OfficeIMO |
| 25000 | package-profile | realworld-freeze-panes | OfficeIMO.Excel | 33.44 ms | 11.2 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-freeze-panes | ClosedXML | 309.66 ms | 205.7 MB | 1.1 MB | OfficeIMO.Excel | 826.1% slower than OfficeIMO |
| 25000 | package-profile | realworld-freeze-panes | EPPlus | 430.69 ms | 206.9 MB | 1.1 MB | OfficeIMO.Excel | 1188.1% slower than OfficeIMO |
| 25000 | package-profile | realworld-pivot-table | OfficeIMO.Excel | 282.43 ms | 128.8 MB | 1.9 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-pivot-table | EPPlus | 459.68 ms | 225.4 MB | 1.1 MB | OfficeIMO.Excel | 62.8% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-all-in-one | OfficeIMO.Excel | 259.80 ms | 130.3 MB | 1.9 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-all-in-one | EPPlus | 462.30 ms | 270.6 MB | 1.1 MB | OfficeIMO.Excel | 77.9% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-chart-first | OfficeIMO.Excel | 91.62 ms | 42.5 MB | 1.9 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-chart-first | EPPlus | 460.14 ms | 270.6 MB | 1.1 MB | OfficeIMO.Excel | 402.2% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-core | OfficeIMO.Excel | 59.32 ms | 11.4 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-core | EPPlus | 683.01 ms | 249.1 MB | 1.1 MB | OfficeIMO.Excel | 1051.4% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-core | ClosedXML | 1.30 s | 664.2 MB | 1.1 MB | OfficeIMO.Excel | 2093.5% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-extra-column | OfficeIMO.Excel | 358.46 ms | 141.4 MB | 2.1 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-extra-column | EPPlus | 518.72 ms | 295.7 MB | 1.1 MB | OfficeIMO.Excel | 44.7% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-no-autofit | OfficeIMO.Excel | 265.40 ms | 130.3 MB | 1.9 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-no-autofit | EPPlus | 441.51 ms | 229.3 MB | 1.1 MB | OfficeIMO.Excel | 66.4% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-post-mutation | OfficeIMO.Excel | 274.82 ms | 130.3 MB | 1.9 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-post-mutation | EPPlus | 461.37 ms | 270.6 MB | 1.1 MB | OfficeIMO.Excel | 67.9% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-shuffled-columns | OfficeIMO.Excel | 296.24 ms | 130.4 MB | 2.0 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-shuffled-columns | EPPlus | 482.21 ms | 270.6 MB | 1.1 MB | OfficeIMO.Excel | 62.8% slower than OfficeIMO |
| 25000 | package-profile | report-workbook | OfficeIMO.Excel | 332.42 ms | 171.1 MB | 2.6 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | report-workbook | EPPlus | 576.90 ms | 356.2 MB | 1.5 MB | OfficeIMO.Excel | 73.5% slower than OfficeIMO |
| 25000 | package-profile | report-workbook-core | OfficeIMO.Excel | 49.38 ms | 10.7 MB | 1.8 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | report-workbook-core | EPPlus | 571.21 ms | 334.8 MB | 1.5 MB | OfficeIMO.Excel | 1056.7% slower than OfficeIMO |
| 25000 | package-profile | report-workbook-core | ClosedXML | 1.11 s | 952.9 MB | 1.5 MB | OfficeIMO.Excel | 2149.1% slower than OfficeIMO |
| 25000 | package-profile | report-workbook-datatable | OfficeIMO.Excel | 358.15 ms | 173.8 MB | 2.6 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | report-workbook-datatable | EPPlus | 583.72 ms | 242.0 MB | 1.5 MB | OfficeIMO.Excel | 63.0% slower than OfficeIMO |
| 25000 | package-profile | report-workbook-datatable-core | OfficeIMO.Excel | 48.24 ms | 13.4 MB | 1.8 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | report-workbook-datatable-core | EPPlus | 517.17 ms | 220.7 MB | 1.5 MB | OfficeIMO.Excel | 972.1% slower than OfficeIMO |
| 25000 | package-profile | report-workbook-datatable-core | ClosedXML | 1.03 s | 812.7 MB | 1.5 MB | OfficeIMO.Excel | 2028.8% slower than OfficeIMO |
| 25000 | package-profile | write-blog-2023-20-string-columns | LargeXlsx | 38.47 ms | 10.5 MB | 2.4 MB | LargeXlsx | 10.7% faster than OfficeIMO |
| 25000 | package-profile | write-blog-2023-20-string-columns | OfficeIMO.Excel | 43.08 ms | 11.4 MB | 2.2 MB | LargeXlsx | Loss +12.0% |
| 25000 | package-profile | write-blog-2023-20-string-columns | MiniExcel | 145.45 ms | 221.6 MB | 2.4 MB | LargeXlsx | 237.6% slower than OfficeIMO |
| 25000 | package-profile | write-blog-2023-20-string-columns | ClosedXML | 889.75 ms | 742.0 MB | 2.5 MB | LargeXlsx | 1965.5% slower than OfficeIMO |
| 25000 | package-profile | write-bulk-report | OfficeIMO.Excel | 35.34 ms | 11.3 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-bulk-report | MiniExcel | 68.93 ms | 122.6 MB | 1.5 MB | OfficeIMO.Excel | 95.1% slower than OfficeIMO |
| 25000 | package-profile | write-bulk-report | EPPlus | 401.74 ms | 249.0 MB | 1.1 MB | OfficeIMO.Excel | 1036.9% slower than OfficeIMO |
| 25000 | package-profile | write-bulk-report | ClosedXML | 789.94 ms | 552.7 MB | 1.1 MB | OfficeIMO.Excel | 2135.4% slower than OfficeIMO |
| 25000 | package-profile | write-cellformula | OfficeIMO.Excel | 19.37 ms | 9.9 MB | 670.3 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellformula | ClosedXML | 167.75 ms | 111.2 MB | 643.2 KB | OfficeIMO.Excel | 766.2% slower than OfficeIMO |
| 25000 | package-profile | write-cellformula | EPPlus | 301.99 ms | 137.4 MB | 593.9 KB | OfficeIMO.Excel | 1459.4% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-empty-strings | OfficeIMO.Excel | 11.89 ms | 6.7 MB | 451.4 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-empty-strings | ClosedXML | 119.57 ms | 90.7 MB | 398.1 KB | OfficeIMO.Excel | 906.0% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-empty-strings | EPPlus | 175.67 ms | 72.7 MB | 390.6 KB | OfficeIMO.Excel | 1378.0% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-numbers | OfficeIMO.Excel | 17.33 ms | 5.8 MB | 462.6 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-numbers | ClosedXML | 112.19 ms | 82.2 MB | 411.4 KB | OfficeIMO.Excel | 547.5% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-numbers | EPPlus | 193.00 ms | 84.4 MB | 406.5 KB | OfficeIMO.Excel | 1013.9% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-mixed | OfficeIMO.Excel | 17.11 ms | 8.1 MB | 585.1 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-object-mixed | ClosedXML | 155.97 ms | 108.5 MB | 532.9 KB | OfficeIMO.Excel | 811.6% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-mixed | EPPlus | 202.66 ms | 110.6 MB | 544.3 KB | OfficeIMO.Excel | 1084.5% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-sparse | OfficeIMO.Excel | 20.02 ms | 7.2 MB | 607.1 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-object-sparse | ClosedXML | 137.51 ms | 102.8 MB | 468.0 KB | OfficeIMO.Excel | 586.8% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-sparse | EPPlus | 214.53 ms | 103.8 MB | 494.4 KB | OfficeIMO.Excel | 971.5% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 17.81 ms | 7.2 MB | 607.1 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-object-sparse-batch | ClosedXML | 137.98 ms | 102.8 MB | 468.0 KB | OfficeIMO.Excel | 674.9% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-sparse-batch | EPPlus | 206.62 ms | 103.8 MB | 494.4 KB | OfficeIMO.Excel | 1060.4% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-scalars | OfficeIMO.Excel | 11.00 ms | 6.0 MB | 441.9 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-scalars | ClosedXML | 100.15 ms | 80.7 MB | 394.9 KB | OfficeIMO.Excel | 810.6% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-scalars | EPPlus | 193.67 ms | 83.1 MB | 379.3 KB | OfficeIMO.Excel | 1661.0% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings | OfficeIMO.Excel | 16.75 ms | 15.0 MB | 527.8 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-strings | ClosedXML | 113.57 ms | 101.8 MB | 460.1 KB | OfficeIMO.Excel | 578.2% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings | EPPlus | 185.44 ms | 82.4 MB | 444.7 KB | OfficeIMO.Excel | 1007.4% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings-distinct | OfficeIMO.Excel | 12.87 ms | 13.5 MB | 499.0 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-strings-distinct | ClosedXML | 151.68 ms | 128.4 MB | 555.3 KB | OfficeIMO.Excel | 1078.2% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings-distinct | EPPlus | 215.34 ms | 95.4 MB | 565.1 KB | OfficeIMO.Excel | 1572.7% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings-repeated | OfficeIMO.Excel | 12.00 ms | 7.3 MB | 376.0 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-strings-repeated | ClosedXML | 95.79 ms | 82.5 MB | 331.8 KB | OfficeIMO.Excel | 697.9% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings-repeated | EPPlus | 151.89 ms | 68.4 MB | 300.8 KB | OfficeIMO.Excel | 1165.2% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-temporal | OfficeIMO.Excel | 22.92 ms | 7.3 MB | 620.5 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-temporal | ClosedXML | 150.35 ms | 87.2 MB | 483.0 KB | OfficeIMO.Excel | 556.1% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-temporal | EPPlus | 195.57 ms | 101.4 MB | 495.1 KB | OfficeIMO.Excel | 753.4% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 9.74 ms | 3.4 MB | 443.4 KB | LargeXlsx | 7.0% faster than OfficeIMO |
| 25000 | package-profile | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 10.47 ms | 6.8 MB | 455.5 KB | LargeXlsx | Loss +7.5% |
| 25000 | package-profile | write-cellvalues-headerless-rectangle-direct | ClosedXML | 124.73 ms | 93.8 MB | 467.5 KB | LargeXlsx | 1090.9% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-headerless-rectangle-direct | EPPlus | 202.06 ms | 85.4 MB | 484.1 KB | LargeXlsx | 1829.3% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | LargeXlsx | 28.54 ms | 5.5 MB | 1.4 MB | LargeXlsx | 17.5% faster than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 34.60 ms | 15.7 MB | 1.4 MB | LargeXlsx | Loss +21.2% |
| 25000 | package-profile | write-cellvalues-rectangle-direct | MiniExcel | 68.77 ms | 91.1 MB | 1.5 MB | LargeXlsx | 98.8% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | ClosedXML | 311.95 ms | 205.7 MB | 1.1 MB | LargeXlsx | 801.6% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | EPPlus | 401.78 ms | 206.9 MB | 1.1 MB | LargeXlsx | 1061.2% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | Sylvan.Data.Excel | 28.26 ms | 5.6 MB | 755.4 KB | Sylvan.Data.Excel | 25.8% faster than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | LargeXlsx | 36.80 ms | 8.2 MB | 1.4 MB | Sylvan.Data.Excel | 3.5% faster than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | OfficeIMO.Excel | 38.11 ms | 12.7 MB | 1.4 MB | Sylvan.Data.Excel | Loss +34.9% |
| 25000 | package-profile | write-datareader-plain | MiniExcel | 75.73 ms | 90.0 MB | 1.5 MB | Sylvan.Data.Excel | 98.7% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | ClosedXML | 284.04 ms | 101.8 MB | 1.1 MB | Sylvan.Data.Excel | 645.3% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | EPPlus | 346.01 ms | 114.7 MB | 1.1 MB | Sylvan.Data.Excel | 807.8% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table | OfficeIMO.Excel | 38.27 ms | 12.7 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-datareader-table | MiniExcel | 71.35 ms | 90.0 MB | 1.5 MB | OfficeIMO.Excel | 86.5% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table | EPPlus | 349.62 ms | 114.7 MB | 1.1 MB | OfficeIMO.Excel | 813.6% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table | ClosedXML | 408.08 ms | 169.3 MB | 1.1 MB | OfficeIMO.Excel | 966.4% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table-autofit | OfficeIMO.Excel | 40.72 ms | 12.7 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-datareader-table-autofit | MiniExcel | 73.24 ms | 121.6 MB | 1.5 MB | OfficeIMO.Excel | 79.9% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table-autofit | EPPlus | 380.69 ms | 156.0 MB | 1.1 MB | OfficeIMO.Excel | 834.9% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table-autofit | ClosedXML | 791.18 ms | 552.9 MB | 1.1 MB | OfficeIMO.Excel | 1842.9% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-direct-export | OfficeIMO.Excel | 33.90 ms | 9.6 MB | 1.3 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-dataset-sparse-direct-export | LargeXlsx | 37.40 ms | 9.0 MB | 1.6 MB | OfficeIMO.Excel | 10.3% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-direct-export | MiniExcel | 100.47 ms | 105.6 MB | 1.8 MB | OfficeIMO.Excel | 196.3% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-direct-export | EPPlus | 497.14 ms | 132.5 MB | 1.4 MB | OfficeIMO.Excel | 1366.3% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-direct-export | ClosedXML | 535.38 ms | 273.8 MB | 1.5 MB | OfficeIMO.Excel | 1479.1% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-tables | OfficeIMO.Excel | 41.08 ms | 13.1 MB | 1.8 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-dataset-sparse-tables | MiniExcel | 99.02 ms | 105.6 MB | 1.8 MB | OfficeIMO.Excel | 141.0% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-tables | EPPlus | 518.34 ms | 132.5 MB | 1.4 MB | OfficeIMO.Excel | 1161.8% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-tables | ClosedXML | 543.21 ms | 273.8 MB | 1.5 MB | OfficeIMO.Excel | 1222.3% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables | OfficeIMO.Excel | 36.50 ms | 10.0 MB | 1.3 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-dataset-tables | MiniExcel | 84.67 ms | 94.8 MB | 1.5 MB | OfficeIMO.Excel | 132.0% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables | EPPlus | 349.91 ms | 108.2 MB | 1.1 MB | OfficeIMO.Excel | 858.7% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables | ClosedXML | 372.66 ms | 168.0 MB | 1.1 MB | OfficeIMO.Excel | 921.1% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables-autofit | OfficeIMO.Excel | 38.99 ms | 10.1 MB | 1.3 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-dataset-tables-autofit | MiniExcel | 83.17 ms | 125.9 MB | 1.5 MB | OfficeIMO.Excel | 113.3% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables-autofit | EPPlus | 405.08 ms | 190.8 MB | 1.1 MB | OfficeIMO.Excel | 938.9% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables-autofit | ClosedXML | 742.48 ms | 537.2 MB | 1.1 MB | OfficeIMO.Excel | 1804.2% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | LargeXlsx | 30.82 ms | 9.3 MB | 1.4 MB | LargeXlsx | 12.3% faster than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | OfficeIMO.Excel | 35.16 ms | 12.4 MB | 1.4 MB | LargeXlsx | Loss +14.1% |
| 25000 | package-profile | write-datatable-direct | MiniExcel | 76.69 ms | 90.2 MB | 1.5 MB | LargeXlsx | 118.1% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | ClosedXML | 282.73 ms | 101.8 MB | 1.1 MB | LargeXlsx | 704.1% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | EPPlus | 336.46 ms | 114.7 MB | 1.1 MB | LargeXlsx | 857.0% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-table-direct | OfficeIMO.Excel | 35.23 ms | 12.4 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-datatable-table-direct | MiniExcel | 79.22 ms | 90.2 MB | 1.5 MB | OfficeIMO.Excel | 124.9% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-table-direct | EPPlus | 343.84 ms | 114.7 MB | 1.1 MB | OfficeIMO.Excel | 876.0% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-table-direct | ClosedXML | 384.98 ms | 169.3 MB | 1.1 MB | OfficeIMO.Excel | 992.8% slower than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | LargeXlsx | 28.33 ms | 5.5 MB | 1.4 MB | LargeXlsx | 12.6% faster than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 32.40 ms | 12.6 MB | 1.4 MB | LargeXlsx | Loss +14.4% |
| 25000 | package-profile | write-fluent-rowsfrom-direct | MiniExcel | 64.27 ms | 91.1 MB | 1.5 MB | LargeXlsx | 98.4% slower than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | ClosedXML | 275.05 ms | 101.8 MB | 1.1 MB | LargeXlsx | 748.9% slower than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | EPPlus | 339.85 ms | 114.7 MB | 1.1 MB | LargeXlsx | 948.9% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 36.13 ms | 11.2 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 375.91 ms | 156.0 MB | 1.1 MB | OfficeIMO.Excel | 940.4% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 689.41 ms | 485.3 MB | 1.1 MB | OfficeIMO.Excel | 1808.0% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | OfficeIMO.Excel | 32.30 ms | 11.2 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-insertobjects-direct | LargeXlsx | 41.68 ms | 5.5 MB | 1.4 MB | OfficeIMO.Excel | 29.0% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | MiniExcel | 65.64 ms | 91.1 MB | 1.5 MB | OfficeIMO.Excel | 103.2% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | ClosedXML | 275.34 ms | 101.8 MB | 1.1 MB | OfficeIMO.Excel | 752.5% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | EPPlus | 348.44 ms | 114.7 MB | 1.1 MB | OfficeIMO.Excel | 978.8% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 40.38 ms | 9.9 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 353.09 ms | 156.0 MB | 1.1 MB | OfficeIMO.Excel | 774.3% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 646.00 ms | 485.3 MB | 1.1 MB | OfficeIMO.Excel | 1499.6% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 29.20 ms | 5.5 MB | 1.4 MB | LargeXlsx | 22.5% faster than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 37.68 ms | 9.9 MB | 1.4 MB | LargeXlsx | Loss +29.0% |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | MiniExcel | 62.43 ms | 91.1 MB | 1.5 MB | LargeXlsx | 65.7% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | ClosedXML | 264.44 ms | 101.8 MB | 1.1 MB | LargeXlsx | 601.9% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | EPPlus | 324.52 ms | 114.7 MB | 1.1 MB | LargeXlsx | 761.3% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | LargeXlsx | 27.64 ms | 5.5 MB | 1.4 MB | LargeXlsx | 34.2% faster than OfficeIMO |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | OfficeIMO.Excel | 42.03 ms | 15.4 MB | 1.4 MB | LargeXlsx | Loss +52.0% |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | MiniExcel | 64.31 ms | 91.1 MB | 1.5 MB | LargeXlsx | 53.0% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | ClosedXML | 284.90 ms | 101.8 MB | 1.1 MB | LargeXlsx | 577.9% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | EPPlus | 340.14 ms | 114.7 MB | 1.1 MB | LargeXlsx | 709.3% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 33.96 ms | 11.2 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 342.32 ms | 135.1 MB | 1.1 MB | OfficeIMO.Excel | 908.1% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 415.95 ms | 269.0 MB | 1.1 MB | OfficeIMO.Excel | 1124.9% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-mixed-objects-direct | LargeXlsx | 41.35 ms | 5.9 MB | 1.8 MB | LargeXlsx | 8.2% faster than OfficeIMO |
| 25000 | package-profile | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 45.07 ms | 10.3 MB | 1.8 MB | LargeXlsx | Loss +9.0% |
| 25000 | package-profile | write-powershell-mixed-objects-direct | MiniExcel | 84.17 ms | 111.3 MB | 1.9 MB | LargeXlsx | 86.7% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-mixed-objects-direct | ClosedXML | 373.70 ms | 175.3 MB | 1.5 MB | LargeXlsx | 729.1% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-mixed-objects-direct | EPPlus | 456.24 ms | 141.5 MB | 1.4 MB | LargeXlsx | 912.2% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-mixed-direct | LargeXlsx | 36.93 ms | 5.9 MB | 1.8 MB | LargeXlsx | 12.8% faster than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 42.36 ms | 9.7 MB | 1.8 MB | LargeXlsx | Loss +14.7% |
| 25000 | package-profile | write-powershell-psobject-mixed-direct | MiniExcel | 79.53 ms | 111.3 MB | 1.9 MB | LargeXlsx | 87.8% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-mixed-direct | ClosedXML | 363.78 ms | 175.3 MB | 1.5 MB | LargeXlsx | 758.8% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-mixed-direct | EPPlus | 442.11 ms | 141.5 MB | 1.4 MB | LargeXlsx | 943.8% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 188.19 ms | 35.3 MB | 6.6 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-powershell-psobject-wide-direct | LargeXlsx | 207.43 ms | 22.7 MB | 6.5 MB | OfficeIMO.Excel | 10.2% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-wide-direct | MiniExcel | 326.91 ms | 339.8 MB | 6.8 MB | OfficeIMO.Excel | 73.7% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-wide-direct | ClosedXML | 1.17 s | 476.0 MB | 6.0 MB | OfficeIMO.Excel | 522.5% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-wide-direct | EPPlus | 1.48 s | 549.7 MB | 5.3 MB | OfficeIMO.Excel | 688.4% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | LargeXlsx | 11.27 ms | 2.7 MB |  | LargeXlsx | 29.6% faster than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | OfficeIMO.Excel | 16.00 ms | 10.6 MB |  | LargeXlsx | Loss +42.0% |
| 25000 | speed-comparison | append-plain-rows | MiniExcel | 33.04 ms | 56.9 MB |  | LargeXlsx | 106.4% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | ClosedXML | 137.30 ms | 101.8 MB |  | LargeXlsx | 757.9% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | EPPlus | 206.71 ms | 98.0 MB |  | LargeXlsx | 1191.5% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | EPPlus 4.5.3.3 | 209.43 ms | 0 B |  | LargeXlsx | 1208.5% slower than OfficeIMO |
| 25000 | speed-comparison | autofit-existing | OfficeIMO.Excel | 109.90 ms | 15.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | autofit-existing | EPPlus 4.5.3.3 | 630.47 ms | 0 B |  | OfficeIMO.Excel | 473.7% slower than OfficeIMO |
| 25000 | speed-comparison | autofit-existing | EPPlus | 673.41 ms | 245.1 MB |  | OfficeIMO.Excel | 512.7% slower than OfficeIMO |
| 25000 | speed-comparison | autofit-existing | ClosedXML | 2.60 s | 810.4 MB |  | OfficeIMO.Excel | 2262.8% slower than OfficeIMO |
| 25000 | speed-comparison | build-object-datatable-dictionaries | OfficeIMO.Excel | 9.68 ms | 5.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | build-object-datatable-typed | OfficeIMO.Excel | 9.18 ms | 7.9 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-cells | OfficeIMO.Excel | 49.67 ms | 24.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-cells | EPPlus | 271.74 ms | 183.0 MB |  | OfficeIMO.Excel | 447.1% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-cells | ClosedXML | 323.82 ms | 162.6 MB |  | OfficeIMO.Excel | 551.9% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-first-column-from-wide-sheet | OfficeIMO.Excel | 34.87 ms | 3.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-first-column-from-wide-sheet | EPPlus | 240.57 ms | 112.8 MB |  | OfficeIMO.Excel | 589.9% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-first-column-from-wide-sheet | ClosedXML | 325.60 ms | 147.4 MB |  | OfficeIMO.Excel | 833.7% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-range | OfficeIMO.Excel | 47.32 ms | 24.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-range | EPPlus | 274.79 ms | 183.0 MB |  | OfficeIMO.Excel | 480.8% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-range | ClosedXML | 346.83 ms | 162.6 MB |  | OfficeIMO.Excel | 633.0% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-top-range | OfficeIMO.Excel | 0.66 ms | 285.4 KB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-top-range | EPPlus | 233.46 ms | 103.1 MB |  | OfficeIMO.Excel | 35206.8% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-top-range | ClosedXML | 316.82 ms | 145.9 MB |  | OfficeIMO.Excel | 47812.8% slower than OfficeIMO |
| 25000 | speed-comparison | formula-heavy-read | OfficeIMO.Excel | 18.36 ms | 6.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | formula-heavy-read | EPPlus 4.5.3.3 | 88.07 ms | 0 B |  | OfficeIMO.Excel | 379.6% slower than OfficeIMO |
| 25000 | speed-comparison | formula-heavy-read | EPPlus | 147.84 ms | 69.2 MB |  | OfficeIMO.Excel | 705.2% slower than OfficeIMO |
| 25000 | speed-comparison | formula-heavy-read | ClosedXML | 171.08 ms | 77.7 MB |  | OfficeIMO.Excel | 831.7% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | OfficeIMO.Excel | 14.40 ms | 15.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | large-shared-strings | MiniExcel | 28.27 ms | 72.0 MB |  | OfficeIMO.Excel | 96.4% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | EPPlus 4.5.3.3 | 101.45 ms | 0 B |  | OfficeIMO.Excel | 604.7% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | ClosedXML | 106.16 ms | 101.8 MB |  | OfficeIMO.Excel | 637.4% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | EPPlus | 174.77 ms | 82.4 MB |  | OfficeIMO.Excel | 1114.0% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | OfficeIMO.Excel | 0.93 ms | 177.2 KB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | large-sparse-column-read | Sylvan.Data.Excel | 1.01 ms | 316.6 KB |  | OfficeIMO.Excel | 8.1% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | ExcelDataReader | 1.50 ms | 4.0 MB |  | OfficeIMO.Excel | 60.6% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | ClosedXML | 4.18 ms | 4.3 MB |  | OfficeIMO.Excel | 347.1% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | MiniExcel | 13.50 ms | 45.1 MB |  | OfficeIMO.Excel | 1345.6% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | EPPlus 4.5.3.3 | 21.73 ms | 0 B |  | OfficeIMO.Excel | 2226.0% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | EPPlus | 98.85 ms | 42.1 MB |  | OfficeIMO.Excel | 10483.1% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | OfficeIMO.Excel | 0.84 ms | 177.3 KB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | large-sparse-row-read | Sylvan.Data.Excel | 0.93 ms | 316.6 KB |  | OfficeIMO.Excel | 10.6% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | ExcelDataReader | 1.43 ms | 4.0 MB |  | OfficeIMO.Excel | 69.7% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | ClosedXML | 5.25 ms | 4.3 MB |  | OfficeIMO.Excel | 524.6% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | MiniExcel | 13.88 ms | 45.1 MB |  | OfficeIMO.Excel | 1550.6% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | EPPlus 4.5.3.3 | 23.31 ms | 0 B |  | OfficeIMO.Excel | 2671.0% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | EPPlus | 98.70 ms | 42.1 MB |  | OfficeIMO.Excel | 11634.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | OfficeIMO.Excel | 32.61 ms | 1.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-bottom-range | Sylvan.Data.Excel | 39.34 ms | 3.5 MB |  | OfficeIMO.Excel | 20.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | ExcelDataReader | 116.04 ms | 59.8 MB |  | OfficeIMO.Excel | 255.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | MiniExcel | 135.54 ms | 182.1 MB |  | OfficeIMO.Excel | 315.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | EPPlus | 244.19 ms | 103.1 MB |  | OfficeIMO.Excel | 648.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | ClosedXML | 363.10 ms | 145.9 MB |  | OfficeIMO.Excel | 1013.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | OfficeIMO.Excel | 34.14 ms | 1.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-bottom-range-stream | Sylvan.Data.Excel | 47.18 ms | 3.5 MB |  | OfficeIMO.Excel | 38.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | ExcelDataReader | 114.55 ms | 59.8 MB |  | OfficeIMO.Excel | 235.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | MiniExcel | 126.08 ms | 182.1 MB |  | OfficeIMO.Excel | 269.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | EPPlus | 239.79 ms | 103.1 MB |  | OfficeIMO.Excel | 602.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | ClosedXML | 356.22 ms | 145.9 MB |  | OfficeIMO.Excel | 943.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | Sylvan.Data.Excel | 64.53 ms | 18.0 MB |  | Sylvan.Data.Excel | 4.3% faster than OfficeIMO |
| 25000 | speed-comparison | read-datatable | OfficeIMO.Excel | 67.43 ms | 33.8 MB |  | Sylvan.Data.Excel | Loss +4.5% |
| 25000 | speed-comparison | read-datatable | ExcelDataReader | 164.53 ms | 74.3 MB |  | Sylvan.Data.Excel | 144.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | MiniExcel | 178.36 ms | 177.0 MB |  | Sylvan.Data.Excel | 164.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | EPPlus | 306.76 ms | 197.5 MB |  | Sylvan.Data.Excel | 355.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | ClosedXML | 401.63 ms | 174.3 MB |  | Sylvan.Data.Excel | 495.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | EPPlus 4.5.3.3 | 410.01 ms | 0 B |  | Sylvan.Data.Excel | 508.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | OfficeIMO.Excel | 34.81 ms | 3.9 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | Sylvan.Data.Excel | 45.36 ms | 4.2 MB |  | OfficeIMO.Excel | 30.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | MiniExcel | 100.80 ms | 154.9 MB |  | OfficeIMO.Excel | 189.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | ExcelDataReader | 108.24 ms | 59.8 MB |  | OfficeIMO.Excel | 210.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | EPPlus | 230.17 ms | 112.8 MB |  | OfficeIMO.Excel | 561.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | ClosedXML | 324.69 ms | 147.4 MB |  | OfficeIMO.Excel | 832.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | Sylvan.Data.Excel | 45.24 ms | 5.7 MB |  | Sylvan.Data.Excel | 13.7% faster than OfficeIMO |
| 25000 | speed-comparison | read-objects | OfficeIMO.Excel | 52.42 ms | 23.0 MB |  | Sylvan.Data.Excel | Loss +15.9% |
| 25000 | speed-comparison | read-objects | ExcelDataReader | 117.71 ms | 62.0 MB |  | Sylvan.Data.Excel | 124.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | MiniExcel | 143.83 ms | 179.4 MB |  | Sylvan.Data.Excel | 174.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | EPPlus 4.5.3.3 | 210.03 ms | 0 B |  | Sylvan.Data.Excel | 300.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | EPPlus | 258.15 ms | 194.9 MB |  | Sylvan.Data.Excel | 392.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | ClosedXML | 316.28 ms | 161.7 MB |  | Sylvan.Data.Excel | 503.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | OfficeIMO.Excel | 113.12 ms | 22.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-objects-stream | Sylvan.Data.Excel | 155.32 ms | 5.2 MB |  | OfficeIMO.Excel | 37.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | EPPlus 4.5.3.3 | 184.71 ms | 0 B |  | OfficeIMO.Excel | 63.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | ExcelDataReader | 192.75 ms | 61.5 MB |  | OfficeIMO.Excel | 70.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | MiniExcel | 339.62 ms | 178.9 MB |  | OfficeIMO.Excel | 200.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | ClosedXML | 538.91 ms | 161.5 MB |  | OfficeIMO.Excel | 376.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | EPPlus | 584.59 ms | 194.7 MB |  | OfficeIMO.Excel | 416.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | Sylvan.Data.Excel | 45.83 ms | 3.5 MB |  | Sylvan.Data.Excel | 11.4% faster than OfficeIMO |
| 25000 | speed-comparison | read-range | OfficeIMO.Excel | 51.75 ms | 25.5 MB |  | Sylvan.Data.Excel | Loss +12.9% |
| 25000 | speed-comparison | read-range | ExcelDataReader | 114.17 ms | 59.8 MB |  | Sylvan.Data.Excel | 120.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | MiniExcel | 127.55 ms | 182.1 MB |  | Sylvan.Data.Excel | 146.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | EPPlus | 253.79 ms | 183.0 MB |  | Sylvan.Data.Excel | 390.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | ClosedXML | 328.36 ms | 159.8 MB |  | Sylvan.Data.Excel | 534.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | EPPlus 4.5.3.3 | 384.90 ms | 0 B |  | Sylvan.Data.Excel | 643.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | Sylvan.Data.Excel | 53.67 ms | 4.4 MB |  | Sylvan.Data.Excel, OfficeIMO.Excel | Tie vs OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | OfficeIMO.Excel | 54.18 ms | 26.1 MB |  | Sylvan.Data.Excel, OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-range-decimal | ExcelDataReader | 116.42 ms | 59.8 MB |  | Sylvan.Data.Excel, OfficeIMO.Excel | 114.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | MiniExcel | 124.69 ms | 182.1 MB |  | Sylvan.Data.Excel, OfficeIMO.Excel | 130.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | EPPlus | 261.30 ms | 183.0 MB |  | Sylvan.Data.Excel, OfficeIMO.Excel | 382.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | ClosedXML | 327.20 ms | 159.8 MB |  | Sylvan.Data.Excel, OfficeIMO.Excel | 503.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | Sylvan.Data.Excel | 76.02 ms | 3.5 MB |  | Sylvan.Data.Excel | 14.3% faster than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | OfficeIMO.Excel | 88.69 ms | 26.3 MB |  | Sylvan.Data.Excel | Loss +16.7% |
| 25000 | speed-comparison | read-range-stream | MiniExcel | 230.22 ms | 182.1 MB |  | Sylvan.Data.Excel | 159.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | ExcelDataReader | 270.79 ms | 59.8 MB |  | Sylvan.Data.Excel | 205.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | EPPlus 4.5.3.3 | 304.24 ms | 0 B |  | Sylvan.Data.Excel | 243.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | ClosedXML | 538.78 ms | 159.8 MB |  | Sylvan.Data.Excel | 507.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | EPPlus | 553.94 ms | 183.0 MB |  | Sylvan.Data.Excel | 524.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | Sylvan.Data.Excel | 0.52 ms | 348.5 KB |  | Sylvan.Data.Excel | 21.0% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range | OfficeIMO.Excel | 0.66 ms | 296.1 KB |  | Sylvan.Data.Excel | Loss +26.6% |
| 25000 | speed-comparison | read-top-range | MiniExcel | 0.93 ms | 869.0 KB |  | Sylvan.Data.Excel | 39.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | ExcelDataReader | 52.14 ms | 16.7 MB |  | Sylvan.Data.Excel | 7758.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | EPPlus | 328.85 ms | 103.1 MB |  | Sylvan.Data.Excel | 49468.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | ClosedXML | 356.60 ms | 145.9 MB |  | Sylvan.Data.Excel | 53651.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | EPPlus 4.5.3.3 | 570.15 ms | 0 B |  | Sylvan.Data.Excel | 85838.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | OfficeIMO.Excel | 0.61 ms | 299.5 KB |  | OfficeIMO.Excel, Sylvan.Data.Excel | Win |
| 25000 | speed-comparison | read-top-range-stream | Sylvan.Data.Excel | 0.62 ms | 348.5 KB |  | OfficeIMO.Excel, Sylvan.Data.Excel | Tie vs OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | MiniExcel | 0.77 ms | 869.0 KB |  | OfficeIMO.Excel, Sylvan.Data.Excel | 24.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | ExcelDataReader | 42.85 ms | 16.7 MB |  | OfficeIMO.Excel, Sylvan.Data.Excel | 6867.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | EPPlus 4.5.3.3 | 228.05 ms | 0 B |  | OfficeIMO.Excel, Sylvan.Data.Excel | 36981.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | EPPlus | 248.42 ms | 103.1 MB |  | OfficeIMO.Excel, Sylvan.Data.Excel | 40293.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | ClosedXML | 342.46 ms | 145.9 MB |  | OfficeIMO.Excel, Sylvan.Data.Excel | 55584.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | Sylvan.Data.Excel | 0.58 ms | 348.5 KB |  | Sylvan.Data.Excel | 19.8% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | OfficeIMO.Excel | 0.72 ms | 300.2 KB |  | Sylvan.Data.Excel | Loss +24.6% |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | MiniExcel | 0.94 ms | 869.0 KB |  | Sylvan.Data.Excel | 30.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | ExcelDataReader | 46.85 ms | 16.7 MB |  | Sylvan.Data.Excel | 6426.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | EPPlus | 298.74 ms | 103.1 MB |  | Sylvan.Data.Excel | 41515.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | ClosedXML | 370.70 ms | 145.9 MB |  | Sylvan.Data.Excel | 51539.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-used-range | Sylvan.Data.Excel | 46.15 ms | 3.5 MB |  | Sylvan.Data.Excel | 52.8% faster than OfficeIMO |
| 25000 | speed-comparison | read-used-range | OfficeIMO.Excel | 97.82 ms | 33.4 MB |  | Sylvan.Data.Excel | Loss +112.0% |
| 25000 | speed-comparison | read-used-range | ExcelDataReader | 119.94 ms | 59.8 MB |  | Sylvan.Data.Excel | 22.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-used-range | MiniExcel | 126.05 ms | 182.1 MB |  | Sylvan.Data.Excel | 28.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-used-range | EPPlus | 262.09 ms | 183.0 MB |  | Sylvan.Data.Excel | 167.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-used-range | ClosedXML | 341.54 ms | 159.8 MB |  | Sylvan.Data.Excel | 249.1% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-autofilter | OfficeIMO.Excel | 31.43 ms | 11.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-autofilter | EPPlus 4.5.3.3 | 233.13 ms | 0 B |  | OfficeIMO.Excel | 641.7% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-autofilter | ClosedXML | 287.67 ms | 205.7 MB |  | OfficeIMO.Excel | 815.2% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-autofilter | EPPlus | 340.51 ms | 206.9 MB |  | OfficeIMO.Excel | 983.4% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-charts | OfficeIMO.Excel | 32.55 ms | 12.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-charts | EPPlus 4.5.3.3 | 247.47 ms | 0 B |  | OfficeIMO.Excel | 660.2% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-charts | EPPlus | 359.42 ms | 209.9 MB |  | OfficeIMO.Excel | 1004.1% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-conditional-formatting | OfficeIMO.Excel | 33.10 ms | 11.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-conditional-formatting | EPPlus 4.5.3.3 | 244.45 ms | 0 B |  | OfficeIMO.Excel | 638.6% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-conditional-formatting | ClosedXML | 312.23 ms | 205.8 MB |  | OfficeIMO.Excel | 843.4% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-conditional-formatting | EPPlus | 367.28 ms | 206.9 MB |  | OfficeIMO.Excel | 1009.7% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-data-validation | OfficeIMO.Excel | 32.92 ms | 11.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-data-validation | EPPlus 4.5.3.3 | 258.20 ms | 0 B |  | OfficeIMO.Excel | 684.3% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-data-validation | ClosedXML | 310.28 ms | 205.7 MB |  | OfficeIMO.Excel | 842.4% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-data-validation | EPPlus | 364.82 ms | 206.9 MB |  | OfficeIMO.Excel | 1008.1% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-freeze-panes | OfficeIMO.Excel | 31.83 ms | 11.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-freeze-panes | EPPlus 4.5.3.3 | 274.84 ms | 0 B |  | OfficeIMO.Excel | 763.6% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-freeze-panes | ClosedXML | 289.98 ms | 205.7 MB |  | OfficeIMO.Excel | 811.1% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-freeze-panes | EPPlus | 342.74 ms | 206.9 MB |  | OfficeIMO.Excel | 976.9% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-pivot-table | OfficeIMO.Excel | 216.45 ms | 128.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-pivot-table | EPPlus 4.5.3.3 | 249.58 ms | 0 B |  | OfficeIMO.Excel | 15.3% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-pivot-table | EPPlus | 381.95 ms | 225.4 MB |  | OfficeIMO.Excel | 76.5% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-all-in-one | OfficeIMO.Excel | 288.90 ms | 130.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-all-in-one | EPPlus | 472.62 ms | 270.6 MB |  | OfficeIMO.Excel | 63.6% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-all-in-one | EPPlus 4.5.3.3 | 474.61 ms | 0 B |  | OfficeIMO.Excel | 64.3% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-chart-first | OfficeIMO.Excel | 89.72 ms | 42.5 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-chart-first | EPPlus | 438.47 ms | 270.6 MB |  | OfficeIMO.Excel | 388.7% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-chart-first | EPPlus 4.5.3.3 | 467.21 ms | 0 B |  | OfficeIMO.Excel | 420.7% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-core | OfficeIMO.Excel | 35.23 ms | 11.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-core | EPPlus | 395.54 ms | 249.1 MB |  | OfficeIMO.Excel | 1022.8% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-core | EPPlus 4.5.3.3 | 467.60 ms | 0 B |  | OfficeIMO.Excel | 1227.4% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-core | ClosedXML | 796.25 ms | 664.2 MB |  | OfficeIMO.Excel | 2160.3% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-extra-column | OfficeIMO.Excel | 306.11 ms | 141.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-extra-column | EPPlus | 461.40 ms | 295.7 MB |  | OfficeIMO.Excel | 50.7% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-extra-column | EPPlus 4.5.3.3 | 517.68 ms | 0 B |  | OfficeIMO.Excel | 69.1% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-no-autofit | EPPlus 4.5.3.3 | 243.01 ms | 0 B |  | EPPlus 4.5.3.3 | 16.1% faster than OfficeIMO |
| 25000 | speed-comparison | realworld-report-no-autofit | OfficeIMO.Excel | 289.71 ms | 130.3 MB |  | EPPlus 4.5.3.3 | Loss +19.2% |
| 25000 | speed-comparison | realworld-report-no-autofit | EPPlus | 439.16 ms | 229.3 MB |  | EPPlus 4.5.3.3 | 51.6% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-post-mutation | OfficeIMO.Excel | 258.71 ms | 130.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-post-mutation | EPPlus | 447.48 ms | 270.6 MB |  | OfficeIMO.Excel | 73.0% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-post-mutation | EPPlus 4.5.3.3 | 476.17 ms | 0 B |  | OfficeIMO.Excel | 84.1% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-shuffled-columns | OfficeIMO.Excel | 286.40 ms | 130.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-shuffled-columns | EPPlus | 460.49 ms | 270.6 MB |  | OfficeIMO.Excel | 60.8% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-shuffled-columns | EPPlus 4.5.3.3 | 483.91 ms | 0 B |  | OfficeIMO.Excel | 69.0% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook | OfficeIMO.Excel | 483.21 ms | 171.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | report-workbook | EPPlus | 612.18 ms | 356.2 MB |  | OfficeIMO.Excel | 26.7% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook | EPPlus 4.5.3.3 | 691.69 ms | 0 B |  | OfficeIMO.Excel | 43.1% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-core | OfficeIMO.Excel | 51.79 ms | 10.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | report-workbook-core | EPPlus | 588.39 ms | 334.8 MB |  | OfficeIMO.Excel | 1036.0% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-core | EPPlus 4.5.3.3 | 674.43 ms | 0 B |  | OfficeIMO.Excel | 1202.2% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-core | ClosedXML | 1.24 s | 952.9 MB |  | OfficeIMO.Excel | 2303.2% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-datatable | OfficeIMO.Excel | 434.69 ms | 173.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | report-workbook-datatable | EPPlus | 556.91 ms | 242.0 MB |  | OfficeIMO.Excel | 28.1% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-datatable | EPPlus 4.5.3.3 | 686.85 ms | 0 B |  | OfficeIMO.Excel | 58.0% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-datatable-core | OfficeIMO.Excel | 51.48 ms | 13.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | report-workbook-datatable-core | EPPlus | 587.14 ms | 220.7 MB |  | OfficeIMO.Excel | 1040.4% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-datatable-core | EPPlus 4.5.3.3 | 653.00 ms | 0 B |  | OfficeIMO.Excel | 1168.3% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-datatable-core | ClosedXML | 1.14 s | 812.7 MB |  | OfficeIMO.Excel | 2123.1% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | Sylvan.Data.Excel | 17.00 ms | 1.9 MB |  | Sylvan.Data.Excel | 12.0% faster than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | OfficeIMO.Excel | 19.32 ms | 9.0 MB |  | Sylvan.Data.Excel | Loss +13.6% |
| 25000 | speed-comparison | shared-string-read | ExcelDataReader | 45.78 ms | 24.4 MB |  | Sylvan.Data.Excel | 137.0% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | MiniExcel | 52.21 ms | 72.7 MB |  | Sylvan.Data.Excel | 170.3% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | EPPlus 4.5.3.3 | 95.96 ms | 0 B |  | Sylvan.Data.Excel | 396.8% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | EPPlus | 146.58 ms | 87.3 MB |  | Sylvan.Data.Excel | 658.9% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | ClosedXML | 153.00 ms | 88.3 MB |  | Sylvan.Data.Excel | 692.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-blog-2023-20-string-columns | LargeXlsx | 39.59 ms | 10.5 MB |  | LargeXlsx | 16.8% faster than OfficeIMO |
| 25000 | speed-comparison | write-blog-2023-20-string-columns | OfficeIMO.Excel | 47.57 ms | 11.4 MB |  | LargeXlsx | Loss +20.1% |
| 25000 | speed-comparison | write-blog-2023-20-string-columns | MiniExcel | 149.96 ms | 221.6 MB |  | LargeXlsx | 215.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-blog-2023-20-string-columns | ClosedXML | 917.30 ms | 742.0 MB |  | LargeXlsx | 1828.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | OfficeIMO.Excel | 38.05 ms | 11.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-bulk-report | MiniExcel | 66.51 ms | 122.6 MB |  | OfficeIMO.Excel | 74.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | EPPlus | 401.29 ms | 249.0 MB |  | OfficeIMO.Excel | 954.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | EPPlus 4.5.3.3 | 501.28 ms | 0 B |  | OfficeIMO.Excel | 1217.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | ClosedXML | 775.84 ms | 552.7 MB |  | OfficeIMO.Excel | 1938.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellformula | OfficeIMO.Excel | 20.47 ms | 9.9 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellformula | ClosedXML | 184.73 ms | 111.2 MB |  | OfficeIMO.Excel | 802.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellformula | EPPlus 4.5.3.3 | 197.28 ms | 0 B |  | OfficeIMO.Excel | 863.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellformula | EPPlus | 341.28 ms | 137.4 MB |  | OfficeIMO.Excel | 1567.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-empty-strings | OfficeIMO.Excel | 12.51 ms | 6.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-empty-strings | ClosedXML | 117.53 ms | 90.7 MB |  | OfficeIMO.Excel | 839.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-empty-strings | EPPlus | 171.50 ms | 72.7 MB |  | OfficeIMO.Excel | 1270.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-numbers | OfficeIMO.Excel | 15.63 ms | 5.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-numbers | ClosedXML | 112.58 ms | 82.2 MB |  | OfficeIMO.Excel | 620.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-numbers | EPPlus 4.5.3.3 | 118.17 ms | 0 B |  | OfficeIMO.Excel | 655.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-numbers | EPPlus | 200.85 ms | 84.4 MB |  | OfficeIMO.Excel | 1184.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-mixed | OfficeIMO.Excel | 26.55 ms | 8.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-object-mixed | EPPlus 4.5.3.3 | 155.26 ms | 0 B |  | OfficeIMO.Excel | 484.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-mixed | ClosedXML | 232.31 ms | 108.5 MB |  | OfficeIMO.Excel | 775.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-mixed | EPPlus | 286.79 ms | 110.6 MB |  | OfficeIMO.Excel | 980.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-sparse | OfficeIMO.Excel | 29.96 ms | 7.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-object-sparse | ClosedXML | 221.95 ms | 102.8 MB |  | OfficeIMO.Excel | 640.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-sparse | EPPlus | 363.91 ms | 103.8 MB |  | OfficeIMO.Excel | 1114.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 22.18 ms | 7.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-object-sparse-batch | ClosedXML | 168.72 ms | 102.8 MB |  | OfficeIMO.Excel | 660.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-sparse-batch | EPPlus | 242.52 ms | 103.8 MB |  | OfficeIMO.Excel | 993.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-scalars | OfficeIMO.Excel | 11.28 ms | 6.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-scalars | ClosedXML | 106.95 ms | 80.7 MB |  | OfficeIMO.Excel | 848.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-scalars | EPPlus 4.5.3.3 | 120.06 ms | 0 B |  | OfficeIMO.Excel | 964.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-scalars | EPPlus | 203.42 ms | 83.1 MB |  | OfficeIMO.Excel | 1703.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings | OfficeIMO.Excel | 16.70 ms | 15.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-strings | EPPlus 4.5.3.3 | 93.86 ms | 0 B |  | OfficeIMO.Excel | 462.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings | ClosedXML | 123.65 ms | 101.8 MB |  | OfficeIMO.Excel | 640.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings | EPPlus | 193.53 ms | 82.4 MB |  | OfficeIMO.Excel | 1058.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings-distinct | OfficeIMO.Excel | 13.10 ms | 13.5 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-strings-distinct | ClosedXML | 149.32 ms | 128.4 MB |  | OfficeIMO.Excel | 1039.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings-distinct | EPPlus | 210.80 ms | 95.4 MB |  | OfficeIMO.Excel | 1508.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings-repeated | OfficeIMO.Excel | 11.53 ms | 7.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-strings-repeated | ClosedXML | 96.29 ms | 82.5 MB |  | OfficeIMO.Excel | 734.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings-repeated | EPPlus | 154.33 ms | 68.4 MB |  | OfficeIMO.Excel | 1238.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-temporal | OfficeIMO.Excel | 30.82 ms | 7.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-temporal | EPPlus 4.5.3.3 | 140.73 ms | 0 B |  | OfficeIMO.Excel | 356.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-temporal | ClosedXML | 243.04 ms | 87.2 MB |  | OfficeIMO.Excel | 688.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-temporal | EPPlus | 301.75 ms | 101.4 MB |  | OfficeIMO.Excel | 879.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 10.09 ms | 3.4 MB |  | LargeXlsx | 9.7% faster than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 11.18 ms | 6.8 MB |  | LargeXlsx | Loss +10.8% |
| 25000 | speed-comparison | write-cellvalues-headerless-rectangle-direct | ClosedXML | 122.21 ms | 93.8 MB |  | LargeXlsx | 993.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-headerless-rectangle-direct | EPPlus | 183.59 ms | 85.4 MB |  | LargeXlsx | 1542.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | LargeXlsx | 31.20 ms | 5.5 MB |  | LargeXlsx | 13.5% faster than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 36.05 ms | 15.7 MB |  | LargeXlsx | Loss +15.6% |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | MiniExcel | 68.11 ms | 91.1 MB |  | LargeXlsx | 88.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 230.16 ms | 0 B |  | LargeXlsx | 538.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | ClosedXML | 303.69 ms | 205.7 MB |  | LargeXlsx | 742.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus | 359.14 ms | 206.9 MB |  | LargeXlsx | 896.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-sparse-rectangle-direct | OfficeIMO.Excel | 15.77 ms | 7.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalues-sparse-rectangle-direct | ClosedXML | 147.00 ms | 102.8 MB |  | OfficeIMO.Excel | 832.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-sparse-rectangle-direct | EPPlus | 207.24 ms | 103.8 MB |  | OfficeIMO.Excel | 1214.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | Sylvan.Data.Excel | 27.72 ms | 5.6 MB |  | Sylvan.Data.Excel | 25.3% faster than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | LargeXlsx | 35.56 ms | 8.2 MB |  | Sylvan.Data.Excel | 4.2% faster than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | OfficeIMO.Excel | 37.13 ms | 12.7 MB |  | Sylvan.Data.Excel | Loss +33.9% |
| 25000 | speed-comparison | write-datareader-plain | MiniExcel | 72.48 ms | 90.0 MB |  | Sylvan.Data.Excel | 95.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | EPPlus 4.5.3.3 | 222.26 ms | 0 B |  | Sylvan.Data.Excel | 498.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | ClosedXML | 278.61 ms | 101.8 MB |  | Sylvan.Data.Excel | 650.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | EPPlus | 340.77 ms | 114.7 MB |  | Sylvan.Data.Excel | 817.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | OfficeIMO.Excel | 39.84 ms | 12.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datareader-table | MiniExcel | 70.82 ms | 90.0 MB |  | OfficeIMO.Excel | 77.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | EPPlus 4.5.3.3 | 215.71 ms | 0 B |  | OfficeIMO.Excel | 441.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | EPPlus | 341.89 ms | 114.7 MB |  | OfficeIMO.Excel | 758.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | ClosedXML | 374.21 ms | 169.3 MB |  | OfficeIMO.Excel | 839.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table-autofit | OfficeIMO.Excel | 39.05 ms | 12.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datareader-table-autofit | MiniExcel | 78.25 ms | 121.6 MB |  | OfficeIMO.Excel | 100.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table-autofit | EPPlus | 374.16 ms | 156.0 MB |  | OfficeIMO.Excel | 858.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table-autofit | EPPlus 4.5.3.3 | 433.61 ms | 0 B |  | OfficeIMO.Excel | 1010.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table-autofit | ClosedXML | 787.58 ms | 552.9 MB |  | OfficeIMO.Excel | 1917.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-headerless-tables | OfficeIMO.Excel | 75.48 ms | 12.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-headerless-tables | MiniExcel | 148.96 ms | 94.8 MB |  | OfficeIMO.Excel | 97.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-headerless-tables | ClosedXML | 636.53 ms | 168.0 MB |  | OfficeIMO.Excel | 743.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-headerless-tables | EPPlus | 827.09 ms | 108.6 MB |  | OfficeIMO.Excel | 995.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | LargeXlsx | 41.23 ms | 9.0 MB |  | LargeXlsx | 9.5% faster than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | OfficeIMO.Excel | 45.54 ms | 9.6 MB |  | LargeXlsx | Loss +10.5% |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | MiniExcel | 99.68 ms | 105.6 MB |  | LargeXlsx | 118.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | EPPlus | 591.72 ms | 132.5 MB |  | LargeXlsx | 1199.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | ClosedXML | 691.56 ms | 273.8 MB |  | LargeXlsx | 1418.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-tables | OfficeIMO.Excel | 68.65 ms | 13.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-sparse-tables | MiniExcel | 157.28 ms | 105.6 MB |  | OfficeIMO.Excel | 129.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-tables | EPPlus | 909.51 ms | 132.5 MB |  | OfficeIMO.Excel | 1224.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-tables | ClosedXML | 945.16 ms | 273.8 MB |  | OfficeIMO.Excel | 1276.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | OfficeIMO.Excel | 33.25 ms | 10.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-tables | MiniExcel | 78.32 ms | 94.8 MB |  | OfficeIMO.Excel | 135.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | EPPlus 4.5.3.3 | 210.73 ms | 0 B |  | OfficeIMO.Excel | 533.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | EPPlus | 329.90 ms | 108.2 MB |  | OfficeIMO.Excel | 892.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | ClosedXML | 340.81 ms | 168.0 MB |  | OfficeIMO.Excel | 925.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables-autofit | OfficeIMO.Excel | 38.25 ms | 10.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-tables-autofit | MiniExcel | 79.15 ms | 125.9 MB |  | OfficeIMO.Excel | 106.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables-autofit | EPPlus | 380.13 ms | 190.8 MB |  | OfficeIMO.Excel | 893.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables-autofit | ClosedXML | 746.95 ms | 537.2 MB |  | OfficeIMO.Excel | 1853.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | LargeXlsx | 31.41 ms | 9.3 MB |  | LargeXlsx | 9.7% faster than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | OfficeIMO.Excel | 34.78 ms | 12.4 MB |  | LargeXlsx | Loss +10.7% |
| 25000 | speed-comparison | write-datatable-direct | MiniExcel | 75.79 ms | 90.2 MB |  | LargeXlsx | 117.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | EPPlus 4.5.3.3 | 214.19 ms | 0 B |  | LargeXlsx | 515.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | ClosedXML | 267.94 ms | 101.8 MB |  | LargeXlsx | 670.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | EPPlus | 320.36 ms | 114.7 MB |  | LargeXlsx | 821.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-object-table-direct | OfficeIMO.Excel | 37.25 ms | 9.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datatable-object-table-direct | MiniExcel | 81.26 ms | 87.6 MB |  | OfficeIMO.Excel | 118.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-object-table-direct | EPPlus | 317.89 ms | 112.0 MB |  | OfficeIMO.Excel | 753.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-object-table-direct | ClosedXML | 374.24 ms | 166.7 MB |  | OfficeIMO.Excel | 904.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | OfficeIMO.Excel | 34.52 ms | 12.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datatable-table-direct | MiniExcel | 81.29 ms | 90.2 MB |  | OfficeIMO.Excel | 135.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | EPPlus 4.5.3.3 | 216.54 ms | 0 B |  | OfficeIMO.Excel | 527.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | EPPlus | 338.12 ms | 114.7 MB |  | OfficeIMO.Excel | 879.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | ClosedXML | 368.22 ms | 169.3 MB |  | OfficeIMO.Excel | 966.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-dictionary-objects-table-direct | OfficeIMO.Excel | 50.66 ms | 14.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | LargeXlsx | 26.99 ms | 5.5 MB |  | LargeXlsx | 18.1% faster than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 32.94 ms | 12.6 MB |  | LargeXlsx | Loss +22.0% |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | MiniExcel | 65.78 ms | 91.1 MB |  | LargeXlsx | 99.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | ClosedXML | 273.55 ms | 101.8 MB |  | LargeXlsx | 730.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus | 364.75 ms | 114.7 MB |  | LargeXlsx | 1007.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 405.64 ms | 0 B |  | LargeXlsx | 1131.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 41.09 ms | 11.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 392.24 ms | 156.0 MB |  | OfficeIMO.Excel | 854.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 777.21 ms | 485.3 MB |  | OfficeIMO.Excel | 1791.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | LargeXlsx | 30.38 ms | 5.5 MB |  | LargeXlsx | 7.7% faster than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | OfficeIMO.Excel | 32.90 ms | 11.2 MB |  | LargeXlsx | Loss +8.3% |
| 25000 | speed-comparison | write-insertobjects-direct | MiniExcel | 72.21 ms | 91.1 MB |  | LargeXlsx | 119.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | ClosedXML | 282.68 ms | 101.8 MB |  | LargeXlsx | 759.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | EPPlus | 333.73 ms | 114.7 MB |  | LargeXlsx | 914.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | EPPlus 4.5.3.3 | 340.19 ms | 0 B |  | LargeXlsx | 934.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 43.12 ms | 9.9 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 375.58 ms | 156.0 MB |  | OfficeIMO.Excel | 770.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 717.46 ms | 485.3 MB |  | OfficeIMO.Excel | 1563.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 28.20 ms | 5.5 MB |  | LargeXlsx | 24.1% faster than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 37.15 ms | 9.9 MB |  | LargeXlsx | Loss +31.7% |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | MiniExcel | 66.86 ms | 91.1 MB |  | LargeXlsx | 80.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | ClosedXML | 284.87 ms | 101.8 MB |  | LargeXlsx | 666.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | EPPlus | 333.21 ms | 114.7 MB |  | LargeXlsx | 796.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 34.88 ms | 11.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 421.61 ms | 135.1 MB |  | OfficeIMO.Excel | 1108.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 434.63 ms | 269.0 MB |  | OfficeIMO.Excel | 1146.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | LargeXlsx | 41.56 ms | 5.9 MB |  | LargeXlsx | 6.3% faster than OfficeIMO |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 44.33 ms | 10.3 MB |  | LargeXlsx | Loss +6.7% |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | MiniExcel | 86.31 ms | 111.3 MB |  | LargeXlsx | 94.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | ClosedXML | 391.17 ms | 175.3 MB |  | LargeXlsx | 782.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | EPPlus | 459.50 ms | 141.5 MB |  | LargeXlsx | 936.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-mixed-direct | LargeXlsx | 38.84 ms | 5.9 MB |  | LargeXlsx | 12.9% faster than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 44.61 ms | 9.7 MB |  | LargeXlsx | Loss +14.8% |
| 25000 | speed-comparison | write-powershell-psobject-mixed-direct | MiniExcel | 83.41 ms | 111.3 MB |  | LargeXlsx | 87.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-mixed-direct | ClosedXML | 382.51 ms | 175.3 MB |  | LargeXlsx | 757.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-mixed-direct | EPPlus | 455.91 ms | 141.5 MB |  | LargeXlsx | 922.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 198.92 ms | 35.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-powershell-psobject-wide-direct | LargeXlsx | 217.02 ms | 22.7 MB |  | OfficeIMO.Excel | 9.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-wide-direct | MiniExcel | 337.07 ms | 339.8 MB |  | OfficeIMO.Excel | 69.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-wide-direct | ClosedXML | 1.19 s | 476.0 MB |  | OfficeIMO.Excel | 499.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-wide-direct | EPPlus | 1.52 s | 549.7 MB |  | OfficeIMO.Excel | 665.8% slower than OfficeIMO |
| 300000 | focused-package-profile | write-blog-2023-20-string-columns | LargeXlsx | 636.47 ms | 93.1 MB | 28.6 MB | LargeXlsx | Win |
| 300000 | focused-package-profile | write-blog-2023-20-string-columns | OfficeIMO.Excel | 701.65 ms | 173.4 MB | 26.6 MB | LargeXlsx | 1.10x vs best |
| 300000 | focused-package-profile | write-blog-2023-20-string-columns | MiniExcel | 2.28 s | 2.46 GB | 28.5 MB | LargeXlsx | 3.58x vs best |
| 300000 | focused-package-profile | write-blog-2023-20-string-columns | ClosedXML | 15.89 s | 8.51 GB | 31.0 MB | LargeXlsx | 24.97x vs best |
