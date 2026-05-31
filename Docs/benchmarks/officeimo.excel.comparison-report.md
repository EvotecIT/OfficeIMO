# OfficeIMO.Excel Benchmark Report

Generated: 2026-05-31T11:23:34.1407225Z
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
| 2500 | dense-helloworld-comparison | read | Other | 0 | 2 | dense-helloworld-read-range vs Sylvan.Data.Excel (1.49x) |
| 2500 | package-profile | package | Package size | 43 | 11 | write-insertobjects-legacy-dictionaries-direct vs LargeXlsx (1.51x) |
| 2500 | speed-comparison | mutate | AutoFit and mutation | 2 | 0 |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | 4 | 0 |  |
| 2500 | speed-comparison | other | Object projection | 2 | 0 |  |
| 2500 | speed-comparison | other | Range and table read | 4 | 0 |  |
| 2500 | speed-comparison | other | Real-world feature mix | 6 | 0 |  |
| 2500 | speed-comparison | other | Real-world report | 2 | 0 |  |
| 2500 | speed-comparison | other | Report workbook | 4 | 0 |  |
| 2500 | speed-comparison | read | Formula write/read | 1 | 0 |  |
| 2500 | speed-comparison | read | Other | 1 | 2 | large-sparse-row-read vs Sylvan.Data.Excel (1.70x) |
| 2500 | speed-comparison | read | Range and table read | 1 | 6 | read-used-range vs Sylvan.Data.Excel (2.13x) |
| 2500 | speed-comparison | read | Streaming read | 1 | 3 | read-top-range-stream vs Sylvan.Data.Excel (1.49x) |
| 2500 | speed-comparison | read | Typed object read | 1 | 1 | read-objects-stream vs Sylvan.Data.Excel (1.07x) |
| 2500 | speed-comparison | write | AutoFit and mutation | 5 | 0 |  |
| 2500 | speed-comparison | write | Cell writer | 7 | 0 |  |
| 2500 | speed-comparison | write | DataSet table export | 3 | 0 |  |
| 2500 | speed-comparison | write | DataTable table export | 4 | 0 |  |
| 2500 | speed-comparison | write | Formatted report write | 1 | 0 |  |
| 2500 | speed-comparison | write | Formula write/read | 1 | 0 |  |
| 2500 | speed-comparison | write | Other | 2 | 2 | write-powershell-mixed-objects-direct vs LargeXlsx (1.51x) |
| 2500 | speed-comparison | write | Plain cell export | 1 | 3 | append-plain-rows vs LargeXlsx (1.38x) |
| 2500 | speed-comparison | write | Plain streaming export | 1 | 1 | write-datareader-plain vs Sylvan.Data.Excel (1.22x) |
| 2500 | speed-comparison | write | Plain string export | 0 | 1 | write-blog-2023-20-string-columns vs LargeXlsx (1.02x) |
| 2500 | speed-comparison | write | Shared string write | 4 | 0 |  |
| 2500 | speed-comparison | write | Typed object export | 0 | 3 | write-fluent-rowsfrom-direct vs LargeXlsx (1.49x) |
| 10000 | focused-package-profile | package | Package size | 1 | 0 |  |
| 25000 | dense-helloworld-comparison | read | Other | 1 | 1 | dense-helloworld-read-range vs Sylvan.Data.Excel (1.17x) |
| 25000 | package-profile | package | Package size | 43 | 11 | write-insertobjects-legacy-dictionaries-direct vs LargeXlsx (1.55x) |
| 25000 | speed-comparison | mutate | AutoFit and mutation | 1 | 1 | realworld-report-no-autofit vs EPPlus 4.5.3.3 (1.07x) |
| 25000 | speed-comparison | other | Anti-cheat report variants | 4 | 0 |  |
| 25000 | speed-comparison | other | Object projection | 2 | 0 |  |
| 25000 | speed-comparison | other | Range and table read | 4 | 0 |  |
| 25000 | speed-comparison | other | Real-world feature mix | 6 | 0 |  |
| 25000 | speed-comparison | other | Real-world report | 2 | 0 |  |
| 25000 | speed-comparison | other | Report workbook | 4 | 0 |  |
| 25000 | speed-comparison | read | Formula write/read | 1 | 0 |  |
| 25000 | speed-comparison | read | Other | 1 | 2 | shared-string-read vs Sylvan.Data.Excel (1.09x) |
| 25000 | speed-comparison | read | Range and table read | 2 | 5 | read-used-range vs Sylvan.Data.Excel (2.11x) |
| 25000 | speed-comparison | read | Streaming read | 1 | 3 | read-top-range-stream-small-chunks vs Sylvan.Data.Excel (1.30x) |
| 25000 | speed-comparison | read | Typed object read | 2 | 0 |  |
| 25000 | speed-comparison | write | AutoFit and mutation | 5 | 0 |  |
| 25000 | speed-comparison | write | Cell writer | 7 | 0 |  |
| 25000 | speed-comparison | write | DataSet table export | 3 | 0 |  |
| 25000 | speed-comparison | write | DataTable table export | 3 | 1 | write-datatable-direct vs LargeXlsx (1.09x) |
| 25000 | speed-comparison | write | Formatted report write | 1 | 0 |  |
| 25000 | speed-comparison | write | Formula write/read | 1 | 0 |  |
| 25000 | speed-comparison | write | Other | 2 | 2 | write-powershell-mixed-objects-direct vs LargeXlsx (1.25x) |
| 25000 | speed-comparison | write | Plain cell export | 1 | 3 | append-plain-rows vs LargeXlsx (1.31x) |
| 25000 | speed-comparison | write | Plain streaming export | 1 | 1 | write-datareader-plain vs Sylvan.Data.Excel (1.23x) |
| 25000 | speed-comparison | write | Plain string export | 0 | 1 | write-blog-2023-20-string-columns vs LargeXlsx (1.08x) |
| 25000 | speed-comparison | write | Shared string write | 4 | 0 |  |
| 25000 | speed-comparison | write | Typed object export | 0 | 3 | write-insertobjects-flat-dictionaries-direct vs LargeXlsx (1.29x) |
| 300000 | focused-package-profile | package | Package size | 0 | 1 | write-blog-2023-20-string-columns vs LargeXlsx (1.10x) |

## Rows

| Rows | Artifact | Scenario | Library | Mean | Allocated | Package | Best | Outcome |
| ---: | --- | --- | --- | ---: | ---: | ---: | --- | --- |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | Sylvan.Data.Excel | 4.11 ms | 362.3 KB |  | Sylvan.Data.Excel | 33.0% faster than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | OfficeIMO.Excel | 6.14 ms | 2.4 MB |  | Sylvan.Data.Excel | Loss +49.2% |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | ExcelDataReader | 11.15 ms | 6.7 MB |  | Sylvan.Data.Excel | 81.7% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | MiniExcel | 18.30 ms | 21.0 MB |  | Sylvan.Data.Excel | 198.2% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | Sylvan.Data.Excel | 4.12 ms | 362.3 KB |  | Sylvan.Data.Excel | 31.6% faster than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | OfficeIMO.Excel | 6.02 ms | 2.4 MB |  | Sylvan.Data.Excel | Loss +46.1% |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | ExcelDataReader | 10.48 ms | 6.7 MB |  | Sylvan.Data.Excel | 74.0% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | MiniExcel | 16.70 ms | 21.0 MB |  | Sylvan.Data.Excel | 177.4% slower than OfficeIMO |
| 2500 | package-profile | append-plain-rows | LargeXlsx | 1.43 ms | 296.4 KB | 63.1 KB | LargeXlsx | 27.4% faster than OfficeIMO |
| 2500 | package-profile | append-plain-rows | OfficeIMO.Excel | 1.97 ms | 1.5 MB | 63.0 KB | LargeXlsx | Loss +37.8% |
| 2500 | package-profile | append-plain-rows | MiniExcel | 4.47 ms | 19.2 MB | 68.1 KB | LargeXlsx | 126.9% slower than OfficeIMO |
| 2500 | package-profile | append-plain-rows | ClosedXML | 14.80 ms | 10.9 MB | 59.8 KB | LargeXlsx | 650.6% slower than OfficeIMO |
| 2500 | package-profile | append-plain-rows | EPPlus | 28.58 ms | 14.0 MB | 56.9 KB | LargeXlsx | 1349.6% slower than OfficeIMO |
| 2500 | package-profile | autofit-existing | OfficeIMO.Excel | 8.84 ms | 1.9 MB | 143.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | autofit-existing | EPPlus | 74.58 ms | 49.5 MB | 115.0 KB | OfficeIMO.Excel | 743.3% slower than OfficeIMO |
| 2500 | package-profile | autofit-existing | ClosedXML | 134.49 ms | 82.6 MB | 121.0 KB | OfficeIMO.Excel | 1420.7% slower than OfficeIMO |
| 2500 | package-profile | large-shared-strings | OfficeIMO.Excel | 2.17 ms | 2.4 MB | 55.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | large-shared-strings | MiniExcel | 4.07 ms | 20.6 MB | 60.7 KB | OfficeIMO.Excel | 87.5% slower than OfficeIMO |
| 2500 | package-profile | large-shared-strings | ClosedXML | 13.57 ms | 11.0 MB | 50.3 KB | OfficeIMO.Excel | 524.6% slower than OfficeIMO |
| 2500 | package-profile | large-shared-strings | EPPlus | 22.59 ms | 12.5 MB | 48.1 KB | OfficeIMO.Excel | 939.7% slower than OfficeIMO |
| 2500 | package-profile | realworld-autofilter | OfficeIMO.Excel | 3.67 ms | 1.3 MB | 142.4 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-autofilter | ClosedXML | 33.17 ms | 21.7 MB | 120.2 KB | OfficeIMO.Excel | 804.8% slower than OfficeIMO |
| 2500 | package-profile | realworld-autofilter | EPPlus | 40.85 ms | 24.1 MB | 114.2 KB | OfficeIMO.Excel | 1014.1% slower than OfficeIMO |
| 2500 | package-profile | realworld-charts | OfficeIMO.Excel | 5.91 ms | 1.8 MB | 147.6 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-charts | EPPlus | 44.08 ms | 26.5 MB | 117.0 KB | OfficeIMO.Excel | 646.2% slower than OfficeIMO |
| 2500 | package-profile | realworld-conditional-formatting | OfficeIMO.Excel | 3.74 ms | 1.4 MB | 142.7 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-conditional-formatting | ClosedXML | 30.31 ms | 21.8 MB | 120.3 KB | OfficeIMO.Excel | 709.4% slower than OfficeIMO |
| 2500 | package-profile | realworld-conditional-formatting | EPPlus | 43.51 ms | 24.2 MB | 114.3 KB | OfficeIMO.Excel | 1062.0% slower than OfficeIMO |
| 2500 | package-profile | realworld-data-validation | OfficeIMO.Excel | 3.55 ms | 1.3 MB | 142.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-data-validation | ClosedXML | 30.75 ms | 21.7 MB | 120.3 KB | OfficeIMO.Excel | 766.5% slower than OfficeIMO |
| 2500 | package-profile | realworld-data-validation | EPPlus | 40.78 ms | 24.1 MB | 114.2 KB | OfficeIMO.Excel | 1049.0% slower than OfficeIMO |
| 2500 | package-profile | realworld-freeze-panes | OfficeIMO.Excel | 3.71 ms | 1.3 MB | 142.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-freeze-panes | ClosedXML | 32.11 ms | 21.7 MB | 120.2 KB | OfficeIMO.Excel | 765.7% slower than OfficeIMO |
| 2500 | package-profile | realworld-freeze-panes | EPPlus | 43.72 ms | 24.2 MB | 114.3 KB | OfficeIMO.Excel | 1078.7% slower than OfficeIMO |
| 2500 | package-profile | realworld-pivot-table | OfficeIMO.Excel | 14.18 ms | 14.1 MB | 200.4 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-pivot-table | EPPlus | 44.79 ms | 28.8 MB | 117.4 KB | OfficeIMO.Excel | 215.9% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-all-in-one | OfficeIMO.Excel | 16.62 ms | 14.9 MB | 206.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-all-in-one | EPPlus | 72.49 ms | 53.3 MB | 121.8 KB | OfficeIMO.Excel | 336.2% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-chart-first | OfficeIMO.Excel | 11.04 ms | 6.0 MB | 206.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-chart-first | EPPlus | 72.66 ms | 53.3 MB | 121.8 KB | OfficeIMO.Excel | 558.4% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-core | OfficeIMO.Excel | 4.35 ms | 1.5 MB | 143.9 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-core | EPPlus | 64.46 ms | 46.2 MB | 115.6 KB | OfficeIMO.Excel | 1382.9% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-core | ClosedXML | 81.73 ms | 68.2 MB | 121.5 KB | OfficeIMO.Excel | 1780.2% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-extra-column | OfficeIMO.Excel | 17.46 ms | 16.0 MB | 219.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-extra-column | EPPlus | 81.43 ms | 57.8 MB | 128.4 KB | OfficeIMO.Excel | 366.3% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-no-autofit | OfficeIMO.Excel | 17.08 ms | 14.9 MB | 206.4 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-no-autofit | EPPlus | 48.34 ms | 32.1 MB | 121.8 KB | OfficeIMO.Excel | 183.1% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-post-mutation | OfficeIMO.Excel | 17.79 ms | 14.9 MB | 206.6 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-post-mutation | EPPlus | 72.15 ms | 53.3 MB | 121.9 KB | OfficeIMO.Excel | 305.6% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-shuffled-columns | OfficeIMO.Excel | 21.14 ms | 14.9 MB | 211.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-shuffled-columns | EPPlus | 79.17 ms | 53.3 MB | 124.3 KB | OfficeIMO.Excel | 274.4% slower than OfficeIMO |
| 2500 | package-profile | report-workbook | OfficeIMO.Excel | 22.58 ms | 18.7 MB | 275.6 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | report-workbook | EPPlus | 89.88 ms | 75.7 MB | 161.8 KB | OfficeIMO.Excel | 298.0% slower than OfficeIMO |
| 2500 | package-profile | report-workbook-core | OfficeIMO.Excel | 6.57 ms | 2.6 MB | 187.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | report-workbook-core | EPPlus | 99.68 ms | 70.3 MB | 157.2 KB | OfficeIMO.Excel | 1417.9% slower than OfficeIMO |
| 2500 | package-profile | report-workbook-core | ClosedXML | 105.93 ms | 94.9 MB | 165.1 KB | OfficeIMO.Excel | 1513.0% slower than OfficeIMO |
| 2500 | package-profile | report-workbook-datatable | OfficeIMO.Excel | 23.43 ms | 18.9 MB | 275.6 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | report-workbook-datatable | EPPlus | 105.46 ms | 64.4 MB | 161.8 KB | OfficeIMO.Excel | 350.1% slower than OfficeIMO |
| 2500 | package-profile | report-workbook-datatable-core | OfficeIMO.Excel | 6.04 ms | 2.9 MB | 187.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | report-workbook-datatable-core | EPPlus | 92.36 ms | 59.1 MB | 157.2 KB | OfficeIMO.Excel | 1430.2% slower than OfficeIMO |
| 2500 | package-profile | report-workbook-datatable-core | ClosedXML | 105.59 ms | 80.9 MB | 165.1 KB | OfficeIMO.Excel | 1649.2% slower than OfficeIMO |
| 2500 | package-profile | write-blog-2023-20-string-columns | LargeXlsx | 6.23 ms | 857.6 KB | 237.7 KB | LargeXlsx | 10.4% faster than OfficeIMO |
| 2500 | package-profile | write-blog-2023-20-string-columns | OfficeIMO.Excel | 6.95 ms | 1.6 MB | 216.7 KB | LargeXlsx | Loss +11.6% |
| 2500 | package-profile | write-blog-2023-20-string-columns | MiniExcel | 30.59 ms | 35.1 MB | 235.3 KB | LargeXlsx | 340.4% slower than OfficeIMO |
| 2500 | package-profile | write-blog-2023-20-string-columns | ClosedXML | 118.88 ms | 69.8 MB | 257.2 KB | LargeXlsx | 1611.2% slower than OfficeIMO |
| 2500 | package-profile | write-bulk-report | OfficeIMO.Excel | 4.24 ms | 1.4 MB | 143.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-bulk-report | MiniExcel | 8.01 ms | 26.2 MB | 153.8 KB | OfficeIMO.Excel | 88.8% slower than OfficeIMO |
| 2500 | package-profile | write-bulk-report | EPPlus | 68.09 ms | 46.1 MB | 115.0 KB | OfficeIMO.Excel | 1505.6% slower than OfficeIMO |
| 2500 | package-profile | write-bulk-report | ClosedXML | 74.11 ms | 57.0 MB | 121.0 KB | OfficeIMO.Excel | 1647.3% slower than OfficeIMO |
| 2500 | package-profile | write-cellformula | OfficeIMO.Excel | 2.38 ms | 1.4 MB | 66.6 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellformula | ClosedXML | 19.85 ms | 11.8 MB | 70.6 KB | OfficeIMO.Excel | 733.6% slower than OfficeIMO |
| 2500 | package-profile | write-cellformula | EPPlus | 37.48 ms | 17.7 MB | 62.1 KB | OfficeIMO.Excel | 1474.4% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-empty-strings | OfficeIMO.Excel | 3.06 ms | 1.7 MB | 44.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-empty-strings | ClosedXML | 20.92 ms | 9.7 MB | 44.9 KB | OfficeIMO.Excel | 584.8% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-empty-strings | EPPlus | 34.60 ms | 11.5 MB | 42.0 KB | OfficeIMO.Excel | 1032.5% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-numbers | OfficeIMO.Excel | 2.39 ms | 1.1 MB | 47.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-numbers | ClosedXML | 15.61 ms | 9.0 MB | 45.9 KB | OfficeIMO.Excel | 552.7% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-numbers | EPPlus | 29.60 ms | 12.6 MB | 43.7 KB | OfficeIMO.Excel | 1137.9% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-mixed | OfficeIMO.Excel | 2.58 ms | 1.7 MB | 61.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-object-mixed | ClosedXML | 21.69 ms | 11.6 MB | 59.5 KB | OfficeIMO.Excel | 742.2% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-mixed | EPPlus | 31.93 ms | 15.3 MB | 58.9 KB | OfficeIMO.Excel | 1139.5% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-sparse | OfficeIMO.Excel | 2.94 ms | 1.5 MB | 62.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-object-sparse | ClosedXML | 17.41 ms | 11.0 MB | 52.5 KB | OfficeIMO.Excel | 492.7% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-sparse | EPPlus | 26.56 ms | 14.6 MB | 54.2 KB | OfficeIMO.Excel | 804.2% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 2.46 ms | 1.5 MB | 62.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-object-sparse-batch | ClosedXML | 16.73 ms | 11.0 MB | 52.5 KB | OfficeIMO.Excel | 579.5% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-sparse-batch | EPPlus | 27.12 ms | 14.6 MB | 54.2 KB | OfficeIMO.Excel | 1001.7% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-scalars | OfficeIMO.Excel | 2.54 ms | 1.1 MB | 46.3 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-scalars | ClosedXML | 16.62 ms | 8.8 MB | 45.4 KB | OfficeIMO.Excel | 553.2% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-scalars | EPPlus | 37.06 ms | 12.5 MB | 42.4 KB | OfficeIMO.Excel | 1356.7% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings | OfficeIMO.Excel | 3.73 ms | 2.6 MB | 55.1 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-strings | EPPlus | 28.71 ms | 12.5 MB | 48.1 KB | OfficeIMO.Excel | 669.6% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings | ClosedXML | 28.87 ms | 11.0 MB | 50.3 KB | OfficeIMO.Excel | 673.9% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings-distinct | OfficeIMO.Excel | 3.19 ms | 2.3 MB | 51.8 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-strings-distinct | ClosedXML | 25.17 ms | 12.8 MB | 61.9 KB | OfficeIMO.Excel | 690.2% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings-distinct | EPPlus | 42.93 ms | 13.6 MB | 61.5 KB | OfficeIMO.Excel | 1247.8% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings-repeated | OfficeIMO.Excel | 2.75 ms | 1.5 MB | 40.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-strings-repeated | ClosedXML | 17.78 ms | 9.0 MB | 38.8 KB | OfficeIMO.Excel | 547.8% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings-repeated | EPPlus | 31.75 ms | 11.1 MB | 34.8 KB | OfficeIMO.Excel | 1056.4% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-temporal | OfficeIMO.Excel | 3.64 ms | 1.4 MB | 63.3 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-temporal | ClosedXML | 18.73 ms | 9.5 MB | 54.5 KB | OfficeIMO.Excel | 413.9% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-temporal | EPPlus | 35.39 ms | 14.4 MB | 53.1 KB | OfficeIMO.Excel | 871.0% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 1.51 ms | 447.0 KB | 47.3 KB | LargeXlsx, OfficeIMO.Excel | Tie vs OfficeIMO |
| 2500 | package-profile | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 1.51 ms | 1.1 MB | 48.2 KB | LargeXlsx, OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalues-headerless-rectangle-direct | ClosedXML | 13.63 ms | 10.0 MB | 53.0 KB | LargeXlsx, OfficeIMO.Excel | 805.6% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-headerless-rectangle-direct | EPPlus | 23.86 ms | 12.7 MB | 52.5 KB | LargeXlsx, OfficeIMO.Excel | 1485.5% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | LargeXlsx | 3.18 ms | 758.3 KB | 138.4 KB | LargeXlsx | 28.3% faster than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 4.44 ms | 2.0 MB | 138.0 KB | LargeXlsx | Loss +39.4% |
| 2500 | package-profile | write-cellvalues-rectangle-direct | MiniExcel | 8.55 ms | 22.7 MB | 153.7 KB | LargeXlsx | 92.7% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | ClosedXML | 35.02 ms | 21.7 MB | 120.1 KB | LargeXlsx | 688.9% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | EPPlus | 41.90 ms | 24.1 MB | 114.1 KB | LargeXlsx | 843.8% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | Sylvan.Data.Excel | 3.13 ms | 758.7 KB | 78.5 KB | Sylvan.Data.Excel | 24.6% faster than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | LargeXlsx | 3.73 ms | 1.0 MB | 138.4 KB | Sylvan.Data.Excel | 10.2% faster than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | OfficeIMO.Excel | 4.15 ms | 1.7 MB | 138.0 KB | Sylvan.Data.Excel | Loss +32.7% |
| 2500 | package-profile | write-datareader-plain | MiniExcel | 7.41 ms | 22.5 MB | 153.6 KB | Sylvan.Data.Excel | 78.6% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | ClosedXML | 27.71 ms | 11.3 MB | 120.1 KB | Sylvan.Data.Excel | 567.9% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | EPPlus | 38.77 ms | 16.3 MB | 114.9 KB | Sylvan.Data.Excel | 834.6% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table | OfficeIMO.Excel | 4.19 ms | 1.7 MB | 138.8 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-datareader-table | MiniExcel | 7.35 ms | 22.5 MB | 153.6 KB | OfficeIMO.Excel | 75.5% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table | ClosedXML | 35.81 ms | 18.6 MB | 120.9 KB | OfficeIMO.Excel | 755.1% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table | EPPlus | 37.12 ms | 16.3 MB | 114.9 KB | OfficeIMO.Excel | 786.4% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table-autofit | OfficeIMO.Excel | 4.54 ms | 1.7 MB | 138.9 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-datareader-table-autofit | MiniExcel | 7.44 ms | 26.0 MB | 153.8 KB | OfficeIMO.Excel | 64.0% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table-autofit | EPPlus | 57.57 ms | 37.4 MB | 115.1 KB | OfficeIMO.Excel | 1169.3% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table-autofit | ClosedXML | 73.07 ms | 57.0 MB | 121.0 KB | OfficeIMO.Excel | 1510.9% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-direct-export | OfficeIMO.Excel | 3.95 ms | 2.1 MB | 131.1 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-dataset-sparse-direct-export | LargeXlsx | 4.67 ms | 1.1 MB | 164.2 KB | OfficeIMO.Excel | 18.3% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-direct-export | MiniExcel | 10.26 ms | 29.0 MB | 180.5 KB | OfficeIMO.Excel | 160.0% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-direct-export | ClosedXML | 53.85 ms | 26.8 MB | 159.4 KB | OfficeIMO.Excel | 1264.1% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-direct-export | EPPlus | 58.83 ms | 21.4 MB | 144.5 KB | OfficeIMO.Excel | 1390.2% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-tables | OfficeIMO.Excel | 4.77 ms | 2.8 MB | 176.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-dataset-sparse-tables | MiniExcel | 10.26 ms | 29.0 MB | 180.5 KB | OfficeIMO.Excel | 115.2% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-tables | ClosedXML | 53.79 ms | 26.8 MB | 159.4 KB | OfficeIMO.Excel | 1027.8% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-tables | EPPlus | 63.17 ms | 21.4 MB | 144.5 KB | OfficeIMO.Excel | 1224.4% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables | OfficeIMO.Excel | 3.98 ms | 2.0 MB | 138.9 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-dataset-tables | MiniExcel | 7.91 ms | 28.0 MB | 156.4 KB | OfficeIMO.Excel | 98.8% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables | ClosedXML | 35.57 ms | 18.4 MB | 123.4 KB | OfficeIMO.Excel | 794.0% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables | EPPlus | 38.03 ms | 18.3 MB | 116.6 KB | OfficeIMO.Excel | 855.9% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables-autofit | OfficeIMO.Excel | 4.33 ms | 2.0 MB | 139.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-dataset-tables-autofit | MiniExcel | 8.49 ms | 31.1 MB | 156.6 KB | OfficeIMO.Excel | 95.9% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables-autofit | EPPlus | 60.29 ms | 40.5 MB | 116.9 KB | OfficeIMO.Excel | 1291.3% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables-autofit | ClosedXML | 76.08 ms | 55.4 MB | 123.7 KB | OfficeIMO.Excel | 1655.7% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | LargeXlsx | 3.55 ms | 1.1 MB | 138.4 KB | LargeXlsx | 11.1% faster than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | OfficeIMO.Excel | 3.99 ms | 1.7 MB | 138.0 KB | LargeXlsx | Loss +12.5% |
| 2500 | package-profile | write-datatable-direct | MiniExcel | 8.08 ms | 22.5 MB | 153.7 KB | LargeXlsx | 102.7% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | ClosedXML | 27.66 ms | 11.3 MB | 120.1 KB | LargeXlsx | 593.7% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | EPPlus | 37.18 ms | 16.3 MB | 114.9 KB | LargeXlsx | 832.4% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-table-direct | OfficeIMO.Excel | 3.92 ms | 1.7 MB | 138.8 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-datatable-table-direct | MiniExcel | 7.99 ms | 22.5 MB | 153.7 KB | OfficeIMO.Excel | 103.7% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-table-direct | ClosedXML | 36.52 ms | 18.6 MB | 120.9 KB | OfficeIMO.Excel | 831.2% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-table-direct | EPPlus | 37.11 ms | 16.3 MB | 114.9 KB | OfficeIMO.Excel | 846.1% slower than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | LargeXlsx | 3.12 ms | 758.3 KB | 138.4 KB | LargeXlsx | 20.0% faster than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 3.90 ms | 1.7 MB | 138.0 KB | LargeXlsx | Loss +25.0% |
| 2500 | package-profile | write-fluent-rowsfrom-direct | MiniExcel | 8.20 ms | 22.7 MB | 153.7 KB | LargeXlsx | 110.1% slower than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | ClosedXML | 30.76 ms | 11.3 MB | 120.1 KB | LargeXlsx | 688.4% slower than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | EPPlus | 40.22 ms | 16.3 MB | 114.9 KB | LargeXlsx | 930.9% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.05 ms | 1.3 MB | 142.4 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 56.19 ms | 37.4 MB | 115.1 KB | OfficeIMO.Excel | 1289.0% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 67.02 ms | 49.7 MB | 120.2 KB | OfficeIMO.Excel | 1556.5% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | LargeXlsx | 3.54 ms | 758.3 KB | 138.4 KB | LargeXlsx | 13.3% faster than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | OfficeIMO.Excel | 4.08 ms | 1.3 MB | 142.3 KB | LargeXlsx | Loss +15.4% |
| 2500 | package-profile | write-insertobjects-direct | MiniExcel | 8.02 ms | 22.7 MB | 153.7 KB | LargeXlsx | 96.4% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | ClosedXML | 31.89 ms | 11.3 MB | 120.1 KB | LargeXlsx | 680.7% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | EPPlus | 38.15 ms | 16.3 MB | 114.9 KB | LargeXlsx | 834.0% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 5.42 ms | 1.5 MB | 138.1 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 56.70 ms | 37.4 MB | 115.1 KB | OfficeIMO.Excel | 946.2% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 68.16 ms | 49.7 MB | 120.2 KB | OfficeIMO.Excel | 1157.7% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 3.23 ms | 758.3 KB | 138.4 KB | LargeXlsx | 24.0% faster than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 4.24 ms | 1.5 MB | 138.0 KB | LargeXlsx | Loss +31.5% |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | MiniExcel | 8.31 ms | 22.7 MB | 153.7 KB | LargeXlsx | 95.9% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | ClosedXML | 28.40 ms | 11.3 MB | 120.1 KB | LargeXlsx | 569.3% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | EPPlus | 41.11 ms | 16.3 MB | 114.9 KB | LargeXlsx | 869.1% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | LargeXlsx | 3.12 ms | 758.3 KB | 138.4 KB | LargeXlsx | 34.0% faster than OfficeIMO |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | OfficeIMO.Excel | 4.73 ms | 1.7 MB | 142.3 KB | LargeXlsx | Loss +51.5% |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | MiniExcel | 8.14 ms | 22.7 MB | 153.7 KB | LargeXlsx | 72.1% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | ClosedXML | 27.60 ms | 11.3 MB | 120.1 KB | LargeXlsx | 483.2% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | EPPlus | 37.43 ms | 16.3 MB | 114.9 KB | LargeXlsx | 690.9% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.24 ms | 1.3 MB | 142.4 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 46.87 ms | 27.9 MB | 120.2 KB | OfficeIMO.Excel | 1004.9% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 49.91 ms | 26.7 MB | 115.0 KB | OfficeIMO.Excel | 1076.7% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-mixed-objects-direct | LargeXlsx | 4.30 ms | 802.5 KB | 182.6 KB | LargeXlsx | 25.6% faster than OfficeIMO |
| 2500 | package-profile | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 5.79 ms | 2.3 MB | 183.1 KB | LargeXlsx | Loss +34.4% |
| 2500 | package-profile | write-powershell-mixed-objects-direct | MiniExcel | 8.84 ms | 24.6 MB | 194.0 KB | LargeXlsx | 52.8% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-mixed-objects-direct | ClosedXML | 36.32 ms | 16.6 MB | 161.0 KB | LargeXlsx | 527.7% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-mixed-objects-direct | EPPlus | 48.77 ms | 19.6 MB | 152.1 KB | LargeXlsx | 743.1% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-mixed-direct | LargeXlsx | 4.24 ms | 802.5 KB | 182.6 KB | LargeXlsx | 5.1% faster than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 4.47 ms | 1.5 MB | 182.4 KB | LargeXlsx | Loss +5.3% |
| 2500 | package-profile | write-powershell-psobject-mixed-direct | MiniExcel | 8.56 ms | 24.6 MB | 194.0 KB | LargeXlsx | 91.8% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-mixed-direct | ClosedXML | 35.02 ms | 16.6 MB | 161.0 KB | LargeXlsx | 684.1% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-mixed-direct | EPPlus | 48.37 ms | 19.6 MB | 152.1 KB | LargeXlsx | 983.2% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 20.44 ms | 4.4 MB | 651.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-powershell-psobject-wide-direct | LargeXlsx | 21.67 ms | 2.7 MB | 644.6 KB | OfficeIMO.Excel | 6.0% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-wide-direct | MiniExcel | 36.87 ms | 47.3 MB | 674.4 KB | OfficeIMO.Excel | 80.4% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-wide-direct | ClosedXML | 122.71 ms | 50.4 MB | 615.5 KB | OfficeIMO.Excel | 500.4% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-wide-direct | EPPlus | 161.19 ms | 67.5 MB | 548.9 KB | OfficeIMO.Excel | 688.7% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | LargeXlsx | 2.36 ms | 296.4 KB |  | LargeXlsx | 27.7% faster than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | OfficeIMO.Excel | 3.26 ms | 1.5 MB |  | LargeXlsx | Loss +38.3% |
| 2500 | speed-comparison | append-plain-rows | MiniExcel | 7.03 ms | 19.2 MB |  | LargeXlsx | 115.5% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | EPPlus 4.5.3.3 | 20.05 ms | 0 B |  | LargeXlsx | 515.1% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | ClosedXML | 23.02 ms | 10.9 MB |  | LargeXlsx | 606.2% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | EPPlus | 46.35 ms | 14.0 MB |  | LargeXlsx | 1321.4% slower than OfficeIMO |
| 2500 | speed-comparison | autofit-existing | OfficeIMO.Excel | 9.89 ms | 1.9 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | autofit-existing | EPPlus | 76.06 ms | 49.5 MB |  | OfficeIMO.Excel | 668.7% slower than OfficeIMO |
| 2500 | speed-comparison | autofit-existing | EPPlus 4.5.3.3 | 93.36 ms | 0 B |  | OfficeIMO.Excel | 843.5% slower than OfficeIMO |
| 2500 | speed-comparison | autofit-existing | ClosedXML | 129.48 ms | 82.6 MB |  | OfficeIMO.Excel | 1208.6% slower than OfficeIMO |
| 2500 | speed-comparison | build-object-datatable-dictionaries | OfficeIMO.Excel | 1.35 ms | 564.2 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | build-object-datatable-typed | OfficeIMO.Excel | 1.18 ms | 856.9 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-cells | OfficeIMO.Excel | 6.06 ms | 2.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-cells | EPPlus | 28.00 ms | 19.7 MB |  | OfficeIMO.Excel | 362.4% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-cells | ClosedXML | 32.33 ms | 16.6 MB |  | OfficeIMO.Excel | 433.8% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-first-column-from-wide-sheet | OfficeIMO.Excel | 4.33 ms | 523.5 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-first-column-from-wide-sheet | EPPlus | 28.29 ms | 12.8 MB |  | OfficeIMO.Excel | 553.9% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-first-column-from-wide-sheet | ClosedXML | 32.57 ms | 15.1 MB |  | OfficeIMO.Excel | 652.9% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-range | OfficeIMO.Excel | 6.16 ms | 2.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-range | EPPlus | 26.99 ms | 19.7 MB |  | OfficeIMO.Excel | 338.0% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-range | ClosedXML | 31.77 ms | 16.6 MB |  | OfficeIMO.Excel | 415.6% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-top-range | OfficeIMO.Excel | 0.69 ms | 285.4 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-top-range | EPPlus | 24.05 ms | 12.1 MB |  | OfficeIMO.Excel | 3410.7% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-top-range | ClosedXML | 31.39 ms | 15.0 MB |  | OfficeIMO.Excel | 4482.3% slower than OfficeIMO |
| 2500 | speed-comparison | formula-heavy-read | OfficeIMO.Excel | 2.65 ms | 706.8 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | formula-heavy-read | EPPlus 4.5.3.3 | 13.21 ms | 0 B |  | OfficeIMO.Excel | 398.0% slower than OfficeIMO |
| 2500 | speed-comparison | formula-heavy-read | ClosedXML | 18.76 ms | 8.1 MB |  | OfficeIMO.Excel | 607.5% slower than OfficeIMO |
| 2500 | speed-comparison | formula-heavy-read | EPPlus | 24.84 ms | 7.5 MB |  | OfficeIMO.Excel | 836.6% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | OfficeIMO.Excel | 2.82 ms | 2.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | large-shared-strings | MiniExcel | 5.09 ms | 20.6 MB |  | OfficeIMO.Excel | 80.5% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | EPPlus 4.5.3.3 | 10.20 ms | 0 B |  | OfficeIMO.Excel | 261.8% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | ClosedXML | 16.18 ms | 11.0 MB |  | OfficeIMO.Excel | 474.1% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | EPPlus | 25.89 ms | 12.5 MB |  | OfficeIMO.Excel | 818.4% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | OfficeIMO.Excel | 0.96 ms | 177.3 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | large-sparse-column-read | Sylvan.Data.Excel | 1.01 ms | 316.6 KB |  | OfficeIMO.Excel | 5.4% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | ExcelDataReader | 1.69 ms | 4.0 MB |  | OfficeIMO.Excel | 76.1% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | ClosedXML | 3.55 ms | 4.3 MB |  | OfficeIMO.Excel | 270.0% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | EPPlus 4.5.3.3 | 11.16 ms | 0 B |  | OfficeIMO.Excel | 1062.0% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | MiniExcel | 11.67 ms | 45.1 MB |  | OfficeIMO.Excel | 1115.3% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | EPPlus | 39.05 ms | 42.1 MB |  | OfficeIMO.Excel | 3965.1% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | Sylvan.Data.Excel | 1.04 ms | 316.6 KB |  | Sylvan.Data.Excel | 41.3% faster than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | ExcelDataReader | 1.56 ms | 4.0 MB |  | Sylvan.Data.Excel | 11.7% faster than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | OfficeIMO.Excel | 1.77 ms | 177.5 KB |  | Sylvan.Data.Excel | Loss +70.3% |
| 2500 | speed-comparison | large-sparse-row-read | ClosedXML | 3.96 ms | 4.3 MB |  | Sylvan.Data.Excel | 124.2% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | EPPlus 4.5.3.3 | 11.49 ms | 0 B |  | Sylvan.Data.Excel | 550.5% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | MiniExcel | 12.03 ms | 45.1 MB |  | Sylvan.Data.Excel | 581.4% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | EPPlus | 39.90 ms | 42.1 MB |  | Sylvan.Data.Excel | 2159.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | Sylvan.Data.Excel | 4.28 ms | 655.2 KB |  | Sylvan.Data.Excel | 5.6% faster than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | OfficeIMO.Excel | 4.54 ms | 374.7 KB |  | Sylvan.Data.Excel | Loss +6.0% |
| 2500 | speed-comparison | read-bottom-range | ExcelDataReader | 10.65 ms | 5.9 MB |  | Sylvan.Data.Excel | 134.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | MiniExcel | 13.34 ms | 18.2 MB |  | Sylvan.Data.Excel | 193.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | EPPlus | 24.49 ms | 12.1 MB |  | Sylvan.Data.Excel | 439.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | ClosedXML | 31.99 ms | 15.0 MB |  | Sylvan.Data.Excel | 604.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | OfficeIMO.Excel | 3.76 ms | 378.0 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | read-bottom-range-stream | Sylvan.Data.Excel | 4.16 ms | 655.2 KB |  | OfficeIMO.Excel | 10.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | ExcelDataReader | 10.40 ms | 5.9 MB |  | OfficeIMO.Excel | 176.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | MiniExcel | 12.94 ms | 18.2 MB |  | OfficeIMO.Excel | 244.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | EPPlus | 24.12 ms | 12.1 MB |  | OfficeIMO.Excel | 541.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | ClosedXML | 30.29 ms | 15.0 MB |  | OfficeIMO.Excel | 705.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | Sylvan.Data.Excel | 6.11 ms | 2.2 MB |  | Sylvan.Data.Excel | 20.3% faster than OfficeIMO |
| 2500 | speed-comparison | read-datatable | OfficeIMO.Excel | 7.66 ms | 3.5 MB |  | Sylvan.Data.Excel | Loss +25.4% |
| 2500 | speed-comparison | read-datatable | ExcelDataReader | 12.83 ms | 7.5 MB |  | Sylvan.Data.Excel | 67.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | MiniExcel | 14.51 ms | 17.8 MB |  | Sylvan.Data.Excel | 89.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | EPPlus | 30.29 ms | 21.2 MB |  | Sylvan.Data.Excel | 295.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | ClosedXML | 34.96 ms | 17.9 MB |  | Sylvan.Data.Excel | 356.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | EPPlus 4.5.3.3 | 36.78 ms | 0 B |  | Sylvan.Data.Excel | 380.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | Sylvan.Data.Excel | 4.77 ms | 733.5 KB |  | Sylvan.Data.Excel | 8.5% faster than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | OfficeIMO.Excel | 5.21 ms | 543.1 KB |  | Sylvan.Data.Excel | Loss +9.3% |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | ExcelDataReader | 11.71 ms | 5.9 MB |  | Sylvan.Data.Excel | 124.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | MiniExcel | 13.90 ms | 15.5 MB |  | Sylvan.Data.Excel | 166.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | EPPlus | 26.03 ms | 12.8 MB |  | Sylvan.Data.Excel | 399.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | ClosedXML | 33.42 ms | 15.1 MB |  | Sylvan.Data.Excel | 541.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | OfficeIMO.Excel | 7.55 ms | 2.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | read-objects | Sylvan.Data.Excel | 7.81 ms | 895.3 KB |  | OfficeIMO.Excel | 3.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | ExcelDataReader | 13.98 ms | 6.2 MB |  | OfficeIMO.Excel | 85.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | MiniExcel | 16.71 ms | 18.0 MB |  | OfficeIMO.Excel | 121.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | EPPlus 4.5.3.3 | 30.13 ms | 0 B |  | OfficeIMO.Excel | 299.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | EPPlus | 32.08 ms | 20.9 MB |  | OfficeIMO.Excel | 325.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | ClosedXML | 33.94 ms | 16.5 MB |  | OfficeIMO.Excel | 349.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | Sylvan.Data.Excel | 5.00 ms | 831.0 KB |  | Sylvan.Data.Excel | 6.5% faster than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | OfficeIMO.Excel | 5.35 ms | 2.4 MB |  | Sylvan.Data.Excel | Loss +6.9% |
| 2500 | speed-comparison | read-objects-stream | ExcelDataReader | 11.00 ms | 6.1 MB |  | Sylvan.Data.Excel | 105.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | MiniExcel | 14.30 ms | 18.0 MB |  | Sylvan.Data.Excel | 167.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | EPPlus | 31.15 ms | 20.8 MB |  | Sylvan.Data.Excel | 482.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | EPPlus 4.5.3.3 | 31.48 ms | 0 B |  | Sylvan.Data.Excel | 488.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | ClosedXML | 31.76 ms | 16.5 MB |  | Sylvan.Data.Excel | 493.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | Sylvan.Data.Excel | 18.04 ms | 655.0 KB |  | Sylvan.Data.Excel | 4.4% faster than OfficeIMO |
| 2500 | speed-comparison | read-range | OfficeIMO.Excel | 18.87 ms | 2.6 MB |  | Sylvan.Data.Excel | Loss +4.6% |
| 2500 | speed-comparison | read-range | ExcelDataReader | 30.78 ms | 5.9 MB |  | Sylvan.Data.Excel | 63.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | MiniExcel | 33.01 ms | 18.2 MB |  | Sylvan.Data.Excel | 75.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | EPPlus 4.5.3.3 | 35.98 ms | 0 B |  | Sylvan.Data.Excel | 90.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | EPPlus | 68.13 ms | 19.7 MB |  | Sylvan.Data.Excel | 261.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | ClosedXML | 125.88 ms | 16.5 MB |  | Sylvan.Data.Excel | 567.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | OfficeIMO.Excel | 5.26 ms | 2.7 MB |  | OfficeIMO.Excel, Sylvan.Data.Excel | Win |
| 2500 | speed-comparison | read-range-decimal | Sylvan.Data.Excel | 5.36 ms | 750.3 KB |  | OfficeIMO.Excel, Sylvan.Data.Excel | Tie vs OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | ExcelDataReader | 10.59 ms | 5.9 MB |  | OfficeIMO.Excel, Sylvan.Data.Excel | 101.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | MiniExcel | 12.90 ms | 18.2 MB |  | OfficeIMO.Excel, Sylvan.Data.Excel | 145.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | EPPlus | 28.19 ms | 19.7 MB |  | OfficeIMO.Excel, Sylvan.Data.Excel | 435.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | ClosedXML | 30.94 ms | 16.3 MB |  | OfficeIMO.Excel, Sylvan.Data.Excel | 487.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | Sylvan.Data.Excel | 5.45 ms | 655.2 KB |  | Sylvan.Data.Excel | 12.0% faster than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | OfficeIMO.Excel | 6.19 ms | 2.7 MB |  | Sylvan.Data.Excel | Loss +13.6% |
| 2500 | speed-comparison | read-range-stream | ExcelDataReader | 11.37 ms | 5.9 MB |  | Sylvan.Data.Excel | 83.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | MiniExcel | 12.77 ms | 18.2 MB |  | Sylvan.Data.Excel | 106.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | EPPlus | 28.07 ms | 19.7 MB |  | Sylvan.Data.Excel | 353.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | EPPlus 4.5.3.3 | 29.80 ms | 0 B |  | Sylvan.Data.Excel | 381.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | ClosedXML | 31.41 ms | 16.3 MB |  | Sylvan.Data.Excel | 407.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | Sylvan.Data.Excel | 0.51 ms | 348.4 KB |  | Sylvan.Data.Excel | 16.5% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range | OfficeIMO.Excel | 0.61 ms | 296.2 KB |  | Sylvan.Data.Excel | Loss +19.7% |
| 2500 | speed-comparison | read-top-range | MiniExcel | 0.83 ms | 869.0 KB |  | Sylvan.Data.Excel | 36.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | ExcelDataReader | 4.61 ms | 1.9 MB |  | Sylvan.Data.Excel | 655.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | EPPlus | 25.45 ms | 12.1 MB |  | Sylvan.Data.Excel | 4071.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | EPPlus 4.5.3.3 | 27.90 ms | 0 B |  | Sylvan.Data.Excel | 4472.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | ClosedXML | 32.49 ms | 15.0 MB |  | Sylvan.Data.Excel | 5225.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | Sylvan.Data.Excel | 0.45 ms | 348.5 KB |  | Sylvan.Data.Excel | 32.9% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | OfficeIMO.Excel | 0.67 ms | 299.5 KB |  | Sylvan.Data.Excel | Loss +49.0% |
| 2500 | speed-comparison | read-top-range-stream | MiniExcel | 0.81 ms | 869.0 KB |  | Sylvan.Data.Excel | 20.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | ExcelDataReader | 4.51 ms | 1.9 MB |  | Sylvan.Data.Excel | 573.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | EPPlus | 24.47 ms | 12.1 MB |  | Sylvan.Data.Excel | 3558.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | EPPlus 4.5.3.3 | 27.19 ms | 0 B |  | Sylvan.Data.Excel | 3965.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | ClosedXML | 31.88 ms | 15.0 MB |  | Sylvan.Data.Excel | 4666.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | Sylvan.Data.Excel | 0.43 ms | 348.5 KB |  | Sylvan.Data.Excel | 21.9% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | OfficeIMO.Excel | 0.55 ms | 300.3 KB |  | Sylvan.Data.Excel | Loss +28.1% |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | MiniExcel | 0.79 ms | 869.0 KB |  | Sylvan.Data.Excel | 44.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | ExcelDataReader | 4.44 ms | 1.9 MB |  | Sylvan.Data.Excel | 706.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | EPPlus | 24.52 ms | 12.1 MB |  | Sylvan.Data.Excel | 4354.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | ClosedXML | 32.07 ms | 15.0 MB |  | Sylvan.Data.Excel | 5725.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-used-range | Sylvan.Data.Excel | 5.03 ms | 655.2 KB |  | Sylvan.Data.Excel | 53.1% faster than OfficeIMO |
| 2500 | speed-comparison | read-used-range | ExcelDataReader | 10.67 ms | 5.9 MB |  | Sylvan.Data.Excel | Tie vs OfficeIMO |
| 2500 | speed-comparison | read-used-range | OfficeIMO.Excel | 10.71 ms | 3.4 MB |  | Sylvan.Data.Excel | Loss +113.1% |
| 2500 | speed-comparison | read-used-range | MiniExcel | 13.16 ms | 18.2 MB |  | Sylvan.Data.Excel | 22.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-used-range | EPPlus | 30.58 ms | 19.7 MB |  | Sylvan.Data.Excel | 185.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-used-range | ClosedXML | 56.90 ms | 16.4 MB |  | Sylvan.Data.Excel | 431.3% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-autofilter | OfficeIMO.Excel | 3.68 ms | 1.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-autofilter | EPPlus 4.5.3.3 | 26.80 ms | 0 B |  | OfficeIMO.Excel | 627.7% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-autofilter | ClosedXML | 32.55 ms | 21.7 MB |  | OfficeIMO.Excel | 783.8% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-autofilter | EPPlus | 42.40 ms | 24.1 MB |  | OfficeIMO.Excel | 1051.1% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-charts | OfficeIMO.Excel | 5.59 ms | 1.8 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-charts | EPPlus 4.5.3.3 | 28.24 ms | 0 B |  | OfficeIMO.Excel | 405.6% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-charts | EPPlus | 45.97 ms | 26.5 MB |  | OfficeIMO.Excel | 723.1% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-conditional-formatting | OfficeIMO.Excel | 3.77 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-conditional-formatting | EPPlus 4.5.3.3 | 27.43 ms | 0 B |  | OfficeIMO.Excel | 626.8% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-conditional-formatting | ClosedXML | 31.31 ms | 21.8 MB |  | OfficeIMO.Excel | 729.8% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-conditional-formatting | EPPlus | 41.11 ms | 24.2 MB |  | OfficeIMO.Excel | 989.5% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-data-validation | OfficeIMO.Excel | 3.64 ms | 1.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-data-validation | EPPlus 4.5.3.3 | 29.17 ms | 0 B |  | OfficeIMO.Excel | 702.0% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-data-validation | ClosedXML | 30.93 ms | 21.7 MB |  | OfficeIMO.Excel | 750.6% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-data-validation | EPPlus | 39.76 ms | 24.1 MB |  | OfficeIMO.Excel | 993.4% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-freeze-panes | OfficeIMO.Excel | 3.89 ms | 1.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-freeze-panes | EPPlus 4.5.3.3 | 27.35 ms | 0 B |  | OfficeIMO.Excel | 602.2% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-freeze-panes | ClosedXML | 32.34 ms | 21.7 MB |  | OfficeIMO.Excel | 730.4% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-freeze-panes | EPPlus | 42.87 ms | 24.2 MB |  | OfficeIMO.Excel | 1000.8% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-pivot-table | OfficeIMO.Excel | 17.61 ms | 14.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-pivot-table | EPPlus 4.5.3.3 | 25.77 ms | 0 B |  | OfficeIMO.Excel | 46.3% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-pivot-table | EPPlus | 51.21 ms | 28.8 MB |  | OfficeIMO.Excel | 190.8% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-all-in-one | OfficeIMO.Excel | 17.09 ms | 14.9 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-all-in-one | EPPlus 4.5.3.3 | 72.25 ms | 0 B |  | OfficeIMO.Excel | 322.9% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-all-in-one | EPPlus | 72.35 ms | 53.3 MB |  | OfficeIMO.Excel | 323.5% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-chart-first | OfficeIMO.Excel | 11.72 ms | 6.0 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-chart-first | EPPlus | 74.98 ms | 53.3 MB |  | OfficeIMO.Excel | 539.6% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-chart-first | EPPlus 4.5.3.3 | 81.24 ms | 0 B |  | OfficeIMO.Excel | 593.0% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-core | OfficeIMO.Excel | 4.44 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-core | EPPlus 4.5.3.3 | 64.30 ms | 0 B |  | OfficeIMO.Excel | 1349.1% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-core | EPPlus | 67.85 ms | 46.2 MB |  | OfficeIMO.Excel | 1429.1% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-core | ClosedXML | 77.24 ms | 68.2 MB |  | OfficeIMO.Excel | 1640.7% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-extra-column | OfficeIMO.Excel | 18.72 ms | 16.0 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-extra-column | EPPlus | 81.06 ms | 57.8 MB |  | OfficeIMO.Excel | 333.0% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-extra-column | EPPlus 4.5.3.3 | 86.64 ms | 0 B |  | OfficeIMO.Excel | 362.9% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-no-autofit | OfficeIMO.Excel | 15.66 ms | 14.9 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-no-autofit | EPPlus 4.5.3.3 | 35.21 ms | 0 B |  | OfficeIMO.Excel | 124.8% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-no-autofit | EPPlus | 45.36 ms | 32.1 MB |  | OfficeIMO.Excel | 189.6% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-post-mutation | OfficeIMO.Excel | 16.75 ms | 14.9 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-post-mutation | EPPlus | 69.78 ms | 53.3 MB |  | OfficeIMO.Excel | 316.6% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-post-mutation | EPPlus 4.5.3.3 | 70.97 ms | 0 B |  | OfficeIMO.Excel | 323.7% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-shuffled-columns | OfficeIMO.Excel | 17.87 ms | 14.9 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-shuffled-columns | EPPlus | 72.01 ms | 53.3 MB |  | OfficeIMO.Excel | 302.9% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-shuffled-columns | EPPlus 4.5.3.3 | 75.53 ms | 0 B |  | OfficeIMO.Excel | 322.6% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook | OfficeIMO.Excel | 30.29 ms | 18.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | report-workbook | EPPlus 4.5.3.3 | 76.31 ms | 0 B |  | OfficeIMO.Excel | 151.9% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook | EPPlus | 93.24 ms | 75.7 MB |  | OfficeIMO.Excel | 207.8% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-core | OfficeIMO.Excel | 6.31 ms | 2.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | report-workbook-core | EPPlus 4.5.3.3 | 66.16 ms | 0 B |  | OfficeIMO.Excel | 948.9% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-core | EPPlus | 97.68 ms | 70.3 MB |  | OfficeIMO.Excel | 1448.5% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-core | ClosedXML | 105.25 ms | 94.9 MB |  | OfficeIMO.Excel | 1568.6% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-datatable | OfficeIMO.Excel | 22.91 ms | 18.9 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | report-workbook-datatable | EPPlus 4.5.3.3 | 67.79 ms | 0 B |  | OfficeIMO.Excel | 195.9% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-datatable | EPPlus | 95.98 ms | 64.4 MB |  | OfficeIMO.Excel | 318.9% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-datatable-core | OfficeIMO.Excel | 5.94 ms | 2.9 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | report-workbook-datatable-core | EPPlus 4.5.3.3 | 71.02 ms | 0 B |  | OfficeIMO.Excel | 1095.1% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-datatable-core | EPPlus | 97.63 ms | 59.1 MB |  | OfficeIMO.Excel | 1542.9% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-datatable-core | ClosedXML | 104.58 ms | 80.9 MB |  | OfficeIMO.Excel | 1660.0% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | Sylvan.Data.Excel | 3.31 ms | 518.6 KB |  | Sylvan.Data.Excel | 23.2% faster than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | OfficeIMO.Excel | 4.31 ms | 1.0 MB |  | Sylvan.Data.Excel | Loss +30.1% |
| 2500 | speed-comparison | shared-string-read | ExcelDataReader | 8.32 ms | 2.6 MB |  | Sylvan.Data.Excel | 93.1% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | MiniExcel | 8.43 ms | 7.4 MB |  | Sylvan.Data.Excel | 95.6% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | EPPlus 4.5.3.3 | 13.09 ms | 0 B |  | Sylvan.Data.Excel | 203.7% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | ClosedXML | 21.93 ms | 9.3 MB |  | Sylvan.Data.Excel | 409.0% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | EPPlus | 24.50 ms | 10.1 MB |  | Sylvan.Data.Excel | 468.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-blog-2023-20-string-columns | LargeXlsx | 5.58 ms | 857.6 KB |  | LargeXlsx | Tie vs OfficeIMO |
| 2500 | speed-comparison | write-blog-2023-20-string-columns | OfficeIMO.Excel | 5.69 ms | 1.6 MB |  | LargeXlsx | Loss +2.0% |
| 2500 | speed-comparison | write-blog-2023-20-string-columns | MiniExcel | 21.43 ms | 35.1 MB |  | LargeXlsx | 276.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-blog-2023-20-string-columns | ClosedXML | 111.34 ms | 69.8 MB |  | LargeXlsx | 1856.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | OfficeIMO.Excel | 6.56 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-bulk-report | MiniExcel | 15.10 ms | 26.2 MB |  | OfficeIMO.Excel | 130.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | EPPlus | 129.84 ms | 48.0 MB |  | OfficeIMO.Excel | 1880.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | EPPlus 4.5.3.3 | 133.11 ms | 0 B |  | OfficeIMO.Excel | 1930.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | ClosedXML | 244.18 ms | 57.0 MB |  | OfficeIMO.Excel | 3624.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellformula | OfficeIMO.Excel | 3.05 ms | 1.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellformula | ClosedXML | 18.24 ms | 11.8 MB |  | OfficeIMO.Excel | 497.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellformula | EPPlus | 36.88 ms | 17.7 MB |  | OfficeIMO.Excel | 1107.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellformula | EPPlus 4.5.3.3 | 47.67 ms | 0 B |  | OfficeIMO.Excel | 1461.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-empty-strings | OfficeIMO.Excel | 2.30 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-empty-strings | ClosedXML | 12.05 ms | 9.7 MB |  | OfficeIMO.Excel | 423.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-empty-strings | EPPlus | 23.75 ms | 11.5 MB |  | OfficeIMO.Excel | 932.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-numbers | OfficeIMO.Excel | 2.13 ms | 1.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-numbers | ClosedXML | 12.57 ms | 9.0 MB |  | OfficeIMO.Excel | 489.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-numbers | EPPlus | 21.55 ms | 12.6 MB |  | OfficeIMO.Excel | 909.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-numbers | EPPlus 4.5.3.3 | 31.66 ms | 0 B |  | OfficeIMO.Excel | 1383.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-mixed | OfficeIMO.Excel | 3.21 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-object-mixed | ClosedXML | 16.73 ms | 11.6 MB |  | OfficeIMO.Excel | 420.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-mixed | EPPlus | 25.04 ms | 15.3 MB |  | OfficeIMO.Excel | 679.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-mixed | EPPlus 4.5.3.3 | 46.58 ms | 0 B |  | OfficeIMO.Excel | 1349.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-sparse | OfficeIMO.Excel | 3.24 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-object-sparse | ClosedXML | 15.17 ms | 11.0 MB |  | OfficeIMO.Excel | 368.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-sparse | EPPlus | 24.37 ms | 14.6 MB |  | OfficeIMO.Excel | 652.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 3.13 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-object-sparse-batch | ClosedXML | 15.18 ms | 11.0 MB |  | OfficeIMO.Excel | 384.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-sparse-batch | EPPlus | 25.76 ms | 14.6 MB |  | OfficeIMO.Excel | 721.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-scalars | OfficeIMO.Excel | 2.25 ms | 1.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-scalars | ClosedXML | 10.94 ms | 8.8 MB |  | OfficeIMO.Excel | 386.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-scalars | EPPlus | 23.77 ms | 12.5 MB |  | OfficeIMO.Excel | 956.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-scalars | EPPlus 4.5.3.3 | 24.86 ms | 0 B |  | OfficeIMO.Excel | 1005.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings | OfficeIMO.Excel | 3.58 ms | 2.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-strings | ClosedXML | 17.65 ms | 11.0 MB |  | OfficeIMO.Excel | 393.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings | EPPlus | 22.66 ms | 12.5 MB |  | OfficeIMO.Excel | 533.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings | EPPlus 4.5.3.3 | 43.64 ms | 0 B |  | OfficeIMO.Excel | 1119.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings-distinct | OfficeIMO.Excel | 2.99 ms | 2.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-strings-distinct | ClosedXML | 18.34 ms | 12.8 MB |  | OfficeIMO.Excel | 513.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings-distinct | EPPlus | 29.41 ms | 13.6 MB |  | OfficeIMO.Excel | 884.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings-repeated | OfficeIMO.Excel | 2.59 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-strings-repeated | ClosedXML | 11.66 ms | 9.0 MB |  | OfficeIMO.Excel | 349.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings-repeated | EPPlus | 25.15 ms | 11.1 MB |  | OfficeIMO.Excel | 870.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-temporal | OfficeIMO.Excel | 3.93 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-temporal | ClosedXML | 15.73 ms | 9.5 MB |  | OfficeIMO.Excel | 299.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-temporal | EPPlus | 25.20 ms | 14.4 MB |  | OfficeIMO.Excel | 540.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-temporal | EPPlus 4.5.3.3 | 34.88 ms | 0 B |  | OfficeIMO.Excel | 786.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 1.29 ms | 447.0 KB |  | LargeXlsx | 22.1% faster than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 1.65 ms | 1.1 MB |  | LargeXlsx | Loss +28.4% |
| 2500 | speed-comparison | write-cellvalues-headerless-rectangle-direct | ClosedXML | 13.40 ms | 10.0 MB |  | LargeXlsx | 711.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-headerless-rectangle-direct | EPPlus | 22.91 ms | 12.7 MB |  | LargeXlsx | 1288.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | LargeXlsx | 3.35 ms | 758.3 KB |  | LargeXlsx | 25.5% faster than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 4.49 ms | 2.0 MB |  | LargeXlsx | Loss +34.3% |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | MiniExcel | 9.13 ms | 22.7 MB |  | LargeXlsx | 103.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | ClosedXML | 32.78 ms | 21.7 MB |  | LargeXlsx | 629.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus | 43.32 ms | 24.1 MB |  | LargeXlsx | 863.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 89.15 ms | 0 B |  | LargeXlsx | 1883.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-sparse-rectangle-direct | OfficeIMO.Excel | 2.39 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalues-sparse-rectangle-direct | ClosedXML | 15.85 ms | 11.0 MB |  | OfficeIMO.Excel | 564.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-sparse-rectangle-direct | EPPlus | 25.38 ms | 14.6 MB |  | OfficeIMO.Excel | 963.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | Sylvan.Data.Excel | 3.69 ms | 758.6 KB |  | Sylvan.Data.Excel | 18.0% faster than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | OfficeIMO.Excel | 4.51 ms | 1.7 MB |  | Sylvan.Data.Excel | Loss +22.0% |
| 2500 | speed-comparison | write-datareader-plain | LargeXlsx | 5.41 ms | 1.0 MB |  | Sylvan.Data.Excel | 19.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | MiniExcel | 7.68 ms | 22.5 MB |  | Sylvan.Data.Excel | 70.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | ClosedXML | 28.25 ms | 11.3 MB |  | Sylvan.Data.Excel | 526.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | EPPlus | 37.61 ms | 16.3 MB |  | Sylvan.Data.Excel | 734.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | EPPlus 4.5.3.3 | 89.67 ms | 0 B |  | Sylvan.Data.Excel | 1889.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | OfficeIMO.Excel | 5.60 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datareader-table | MiniExcel | 8.96 ms | 22.5 MB |  | OfficeIMO.Excel | 60.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | ClosedXML | 37.88 ms | 18.6 MB |  | OfficeIMO.Excel | 577.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | EPPlus | 44.08 ms | 16.3 MB |  | OfficeIMO.Excel | 687.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | EPPlus 4.5.3.3 | 52.50 ms | 0 B |  | OfficeIMO.Excel | 838.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table-autofit | OfficeIMO.Excel | 5.07 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datareader-table-autofit | MiniExcel | 10.12 ms | 26.0 MB |  | OfficeIMO.Excel | 99.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table-autofit | EPPlus | 57.27 ms | 37.4 MB |  | OfficeIMO.Excel | 1029.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table-autofit | ClosedXML | 76.39 ms | 57.0 MB |  | OfficeIMO.Excel | 1407.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table-autofit | EPPlus 4.5.3.3 | 206.69 ms | 0 B |  | OfficeIMO.Excel | 3977.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-headerless-tables | OfficeIMO.Excel | 5.45 ms | 2.0 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-headerless-tables | MiniExcel | 15.49 ms | 28.5 MB |  | OfficeIMO.Excel | 184.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-headerless-tables | ClosedXML | 46.14 ms | 18.5 MB |  | OfficeIMO.Excel | 745.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-headerless-tables | EPPlus | 58.12 ms | 17.3 MB |  | OfficeIMO.Excel | 965.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | OfficeIMO.Excel | 4.71 ms | 2.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | LargeXlsx | 9.72 ms | 1.1 MB |  | OfficeIMO.Excel | 106.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | MiniExcel | 12.64 ms | 29.0 MB |  | OfficeIMO.Excel | 168.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | ClosedXML | 58.09 ms | 26.8 MB |  | OfficeIMO.Excel | 1132.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | EPPlus | 62.40 ms | 21.4 MB |  | OfficeIMO.Excel | 1224.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-tables | OfficeIMO.Excel | 8.68 ms | 2.8 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-sparse-tables | MiniExcel | 16.11 ms | 29.3 MB |  | OfficeIMO.Excel | 85.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-tables | EPPlus | 63.07 ms | 21.4 MB |  | OfficeIMO.Excel | 626.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-tables | ClosedXML | 77.00 ms | 26.8 MB |  | OfficeIMO.Excel | 787.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | OfficeIMO.Excel | 5.50 ms | 2.0 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-tables | MiniExcel | 15.88 ms | 28.0 MB |  | OfficeIMO.Excel | 188.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | EPPlus 4.5.3.3 | 47.36 ms | 0 B |  | OfficeIMO.Excel | 761.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | ClosedXML | 83.23 ms | 18.4 MB |  | OfficeIMO.Excel | 1413.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | EPPlus | 85.71 ms | 19.0 MB |  | OfficeIMO.Excel | 1458.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables-autofit | OfficeIMO.Excel | 7.29 ms | 2.0 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-tables-autofit | MiniExcel | 18.66 ms | 31.6 MB |  | OfficeIMO.Excel | 155.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables-autofit | EPPlus | 147.03 ms | 42.4 MB |  | OfficeIMO.Excel | 1915.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables-autofit | ClosedXML | 201.58 ms | 55.4 MB |  | OfficeIMO.Excel | 2663.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | OfficeIMO.Excel | 4.22 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datatable-direct | LargeXlsx | 8.70 ms | 1.1 MB |  | OfficeIMO.Excel | 106.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | MiniExcel | 9.25 ms | 22.5 MB |  | OfficeIMO.Excel | 119.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | ClosedXML | 29.38 ms | 11.3 MB |  | OfficeIMO.Excel | 596.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | EPPlus | 41.04 ms | 16.3 MB |  | OfficeIMO.Excel | 872.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | EPPlus 4.5.3.3 | 42.05 ms | 0 B |  | OfficeIMO.Excel | 896.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-object-table-direct | OfficeIMO.Excel | 4.88 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datatable-object-table-direct | MiniExcel | 8.37 ms | 22.3 MB |  | OfficeIMO.Excel | 71.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-object-table-direct | ClosedXML | 36.16 ms | 18.3 MB |  | OfficeIMO.Excel | 640.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-object-table-direct | EPPlus | 37.03 ms | 16.0 MB |  | OfficeIMO.Excel | 658.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | OfficeIMO.Excel | 4.94 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datatable-table-direct | MiniExcel | 9.11 ms | 22.5 MB |  | OfficeIMO.Excel | 84.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | ClosedXML | 38.67 ms | 18.6 MB |  | OfficeIMO.Excel | 683.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | EPPlus | 42.78 ms | 16.3 MB |  | OfficeIMO.Excel | 766.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | EPPlus 4.5.3.3 | 44.89 ms | 0 B |  | OfficeIMO.Excel | 809.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-dictionary-objects-table-direct | OfficeIMO.Excel | 5.62 ms | 2.0 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | LargeXlsx | 3.27 ms | 758.3 KB |  | LargeXlsx | 33.0% faster than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 4.88 ms | 1.7 MB |  | LargeXlsx | Loss +49.2% |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | MiniExcel | 8.66 ms | 22.7 MB |  | LargeXlsx | 77.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | ClosedXML | 32.30 ms | 11.3 MB |  | LargeXlsx | 562.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus | 41.39 ms | 16.3 MB |  | LargeXlsx | 749.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 67.00 ms | 0 B |  | LargeXlsx | 1274.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.02 ms | 1.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 53.84 ms | 37.4 MB |  | OfficeIMO.Excel | 1239.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 64.62 ms | 49.7 MB |  | OfficeIMO.Excel | 1507.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | LargeXlsx | 3.16 ms | 758.3 KB |  | LargeXlsx | 27.7% faster than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | OfficeIMO.Excel | 4.36 ms | 1.3 MB |  | LargeXlsx | Loss +38.2% |
| 2500 | speed-comparison | write-insertobjects-direct | MiniExcel | 9.87 ms | 22.7 MB |  | LargeXlsx | 126.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | ClosedXML | 28.65 ms | 11.3 MB |  | LargeXlsx | 556.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | EPPlus | 41.34 ms | 16.3 MB |  | LargeXlsx | 847.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | EPPlus 4.5.3.3 | 72.83 ms | 0 B |  | LargeXlsx | 1569.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.62 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 54.67 ms | 37.4 MB |  | OfficeIMO.Excel | 1084.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 67.78 ms | 49.7 MB |  | OfficeIMO.Excel | 1368.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 3.20 ms | 758.3 KB |  | LargeXlsx | 26.4% faster than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 4.35 ms | 1.5 MB |  | LargeXlsx | Loss +35.9% |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | MiniExcel | 8.26 ms | 22.7 MB |  | LargeXlsx | 89.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | ClosedXML | 28.31 ms | 11.3 MB |  | LargeXlsx | 550.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | EPPlus | 41.45 ms | 16.3 MB |  | LargeXlsx | 852.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 3.90 ms | 1.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 44.06 ms | 27.9 MB |  | OfficeIMO.Excel | 1030.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 49.31 ms | 26.7 MB |  | OfficeIMO.Excel | 1165.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | LargeXlsx | 4.74 ms | 802.5 KB |  | LargeXlsx | 33.8% faster than OfficeIMO |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 7.16 ms | 2.3 MB |  | LargeXlsx | Loss +51.0% |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | MiniExcel | 10.11 ms | 24.6 MB |  | LargeXlsx | 41.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | ClosedXML | 45.73 ms | 16.6 MB |  | LargeXlsx | 538.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | EPPlus | 60.05 ms | 19.6 MB |  | LargeXlsx | 738.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-mixed-direct | LargeXlsx | 5.21 ms | 802.5 KB |  | LargeXlsx | 14.9% faster than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 6.13 ms | 1.5 MB |  | LargeXlsx | Loss +17.6% |
| 2500 | speed-comparison | write-powershell-psobject-mixed-direct | MiniExcel | 10.48 ms | 24.6 MB |  | LargeXlsx | 71.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-mixed-direct | ClosedXML | 44.07 ms | 16.6 MB |  | LargeXlsx | 619.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-mixed-direct | EPPlus | 56.45 ms | 19.6 MB |  | LargeXlsx | 821.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 21.68 ms | 4.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-powershell-psobject-wide-direct | LargeXlsx | 23.90 ms | 2.7 MB |  | OfficeIMO.Excel | 10.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-wide-direct | MiniExcel | 37.56 ms | 47.3 MB |  | OfficeIMO.Excel | 73.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-wide-direct | ClosedXML | 137.17 ms | 50.4 MB |  | OfficeIMO.Excel | 532.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-wide-direct | EPPlus | 166.39 ms | 67.5 MB |  | OfficeIMO.Excel | 667.6% slower than OfficeIMO |
| 10000 | focused-package-profile | write-blog-2023-20-string-columns | OfficeIMO.Excel | 33.01 ms | 7.6 MB | 880.4 KB | OfficeIMO.Excel | Win |
| 10000 | focused-package-profile | write-blog-2023-20-string-columns | LargeXlsx | 88.91 ms | 3.1 MB | 970.2 KB | OfficeIMO.Excel | 2.69x vs best |
| 10000 | focused-package-profile | write-blog-2023-20-string-columns | MiniExcel | 142.47 ms | 96.2 MB | 957.6 KB | OfficeIMO.Excel | 4.32x vs best |
| 10000 | focused-package-profile | write-blog-2023-20-string-columns | ClosedXML | 725.30 ms | 280.2 MB | 1,015.4 KB | OfficeIMO.Excel | 21.97x vs best |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | Sylvan.Data.Excel | 68.52 ms | 394.1 KB |  | Sylvan.Data.Excel | 14.3% faster than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | OfficeIMO.Excel | 79.91 ms | 23.1 MB |  | Sylvan.Data.Excel | Loss +16.6% |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | ExcelDataReader | 179.67 ms | 67.9 MB |  | Sylvan.Data.Excel | 124.8% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | MiniExcel | 247.23 ms | 210.3 MB |  | Sylvan.Data.Excel | 209.4% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | OfficeIMO.Excel | 62.41 ms | 23.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | Sylvan.Data.Excel | 63.74 ms | 394.1 KB |  | OfficeIMO.Excel | 2.1% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | ExcelDataReader | 196.90 ms | 67.9 MB |  | OfficeIMO.Excel | 215.5% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | MiniExcel | 216.80 ms | 210.3 MB |  | OfficeIMO.Excel | 247.4% slower than OfficeIMO |
| 25000 | package-profile | append-plain-rows | LargeXlsx | 15.21 ms | 2.7 MB | 605.0 KB | LargeXlsx | 25.8% faster than OfficeIMO |
| 25000 | package-profile | append-plain-rows | OfficeIMO.Excel | 20.48 ms | 10.6 MB | 610.4 KB | LargeXlsx | Loss +34.7% |
| 25000 | package-profile | append-plain-rows | MiniExcel | 41.72 ms | 56.9 MB | 642.3 KB | LargeXlsx | 103.7% slower than OfficeIMO |
| 25000 | package-profile | append-plain-rows | ClosedXML | 173.50 ms | 101.8 MB | 540.6 KB | LargeXlsx | 747.1% slower than OfficeIMO |
| 25000 | package-profile | append-plain-rows | EPPlus | 279.73 ms | 98.0 MB | 525.6 KB | LargeXlsx | 1265.8% slower than OfficeIMO |
| 25000 | package-profile | autofit-existing | OfficeIMO.Excel | 99.77 ms | 15.3 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | autofit-existing | EPPlus | 679.38 ms | 245.1 MB | 1.1 MB | OfficeIMO.Excel | 581.0% slower than OfficeIMO |
| 25000 | package-profile | autofit-existing | ClosedXML | 2.15 s | 810.4 MB | 1.1 MB | OfficeIMO.Excel | 2050.0% slower than OfficeIMO |
| 25000 | package-profile | large-shared-strings | OfficeIMO.Excel | 22.57 ms | 15.4 MB | 529.7 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | large-shared-strings | MiniExcel | 57.58 ms | 72.0 MB | 581.0 KB | OfficeIMO.Excel | 155.1% slower than OfficeIMO |
| 25000 | package-profile | large-shared-strings | ClosedXML | 171.92 ms | 101.8 MB | 460.1 KB | OfficeIMO.Excel | 661.8% slower than OfficeIMO |
| 25000 | package-profile | large-shared-strings | EPPlus | 419.40 ms | 82.4 MB | 444.7 KB | OfficeIMO.Excel | 1758.4% slower than OfficeIMO |
| 25000 | package-profile | realworld-autofilter | OfficeIMO.Excel | 49.51 ms | 11.2 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-autofilter | ClosedXML | 428.62 ms | 205.7 MB | 1.1 MB | OfficeIMO.Excel | 765.8% slower than OfficeIMO |
| 25000 | package-profile | realworld-autofilter | EPPlus | 547.71 ms | 206.9 MB | 1.1 MB | OfficeIMO.Excel | 1006.4% slower than OfficeIMO |
| 25000 | package-profile | realworld-charts | OfficeIMO.Excel | 34.86 ms | 12.3 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-charts | EPPlus | 423.05 ms | 209.9 MB | 1.1 MB | OfficeIMO.Excel | 1113.7% slower than OfficeIMO |
| 25000 | package-profile | realworld-conditional-formatting | OfficeIMO.Excel | 46.99 ms | 11.3 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-conditional-formatting | ClosedXML | 393.79 ms | 205.8 MB | 1.1 MB | OfficeIMO.Excel | 738.0% slower than OfficeIMO |
| 25000 | package-profile | realworld-conditional-formatting | EPPlus | 546.43 ms | 206.9 MB | 1.1 MB | OfficeIMO.Excel | 1062.9% slower than OfficeIMO |
| 25000 | package-profile | realworld-data-validation | OfficeIMO.Excel | 43.64 ms | 11.2 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-data-validation | ClosedXML | 379.88 ms | 205.7 MB | 1.1 MB | OfficeIMO.Excel | 770.5% slower than OfficeIMO |
| 25000 | package-profile | realworld-data-validation | EPPlus | 510.66 ms | 206.9 MB | 1.1 MB | OfficeIMO.Excel | 1070.1% slower than OfficeIMO |
| 25000 | package-profile | realworld-freeze-panes | OfficeIMO.Excel | 49.71 ms | 11.2 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-freeze-panes | ClosedXML | 426.98 ms | 205.7 MB | 1.1 MB | OfficeIMO.Excel | 758.9% slower than OfficeIMO |
| 25000 | package-profile | realworld-freeze-panes | EPPlus | 642.83 ms | 206.9 MB | 1.1 MB | OfficeIMO.Excel | 1193.1% slower than OfficeIMO |
| 25000 | package-profile | realworld-pivot-table | OfficeIMO.Excel | 385.23 ms | 128.8 MB | 1.9 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-pivot-table | EPPlus | 574.75 ms | 225.4 MB | 1.1 MB | OfficeIMO.Excel | 49.2% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-all-in-one | OfficeIMO.Excel | 279.05 ms | 130.3 MB | 1.9 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-all-in-one | EPPlus | 509.88 ms | 270.6 MB | 1.1 MB | OfficeIMO.Excel | 82.7% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-chart-first | OfficeIMO.Excel | 96.51 ms | 42.5 MB | 1.9 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-chart-first | EPPlus | 484.82 ms | 270.6 MB | 1.1 MB | OfficeIMO.Excel | 402.4% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-core | OfficeIMO.Excel | 54.69 ms | 11.4 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-core | EPPlus | 616.79 ms | 249.1 MB | 1.1 MB | OfficeIMO.Excel | 1027.8% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-core | ClosedXML | 1.29 s | 664.2 MB | 1.1 MB | OfficeIMO.Excel | 2251.8% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-extra-column | OfficeIMO.Excel | 477.94 ms | 141.4 MB | 2.1 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-extra-column | EPPlus | 684.69 ms | 295.7 MB | 1.1 MB | OfficeIMO.Excel | 43.3% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-no-autofit | OfficeIMO.Excel | 269.74 ms | 130.3 MB | 1.9 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-no-autofit | EPPlus | 455.50 ms | 229.3 MB | 1.1 MB | OfficeIMO.Excel | 68.9% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-post-mutation | OfficeIMO.Excel | 355.35 ms | 130.3 MB | 1.9 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-post-mutation | EPPlus | 597.79 ms | 270.6 MB | 1.1 MB | OfficeIMO.Excel | 68.2% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-shuffled-columns | OfficeIMO.Excel | 301.21 ms | 130.4 MB | 2.0 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-shuffled-columns | EPPlus | 523.34 ms | 270.6 MB | 1.1 MB | OfficeIMO.Excel | 73.7% slower than OfficeIMO |
| 25000 | package-profile | report-workbook | OfficeIMO.Excel | 618.76 ms | 171.1 MB | 2.6 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | report-workbook | EPPlus | 746.43 ms | 356.2 MB | 1.5 MB | OfficeIMO.Excel | 20.6% slower than OfficeIMO |
| 25000 | package-profile | report-workbook-core | OfficeIMO.Excel | 72.67 ms | 10.7 MB | 1.8 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | report-workbook-core | EPPlus | 755.83 ms | 334.8 MB | 1.5 MB | OfficeIMO.Excel | 940.1% slower than OfficeIMO |
| 25000 | package-profile | report-workbook-core | ClosedXML | 1.70 s | 952.9 MB | 1.5 MB | OfficeIMO.Excel | 2235.6% slower than OfficeIMO |
| 25000 | package-profile | report-workbook-datatable | OfficeIMO.Excel | 597.30 ms | 173.8 MB | 2.6 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | report-workbook-datatable | EPPlus | 696.60 ms | 242.0 MB | 1.5 MB | OfficeIMO.Excel | 16.6% slower than OfficeIMO |
| 25000 | package-profile | report-workbook-datatable-core | OfficeIMO.Excel | 66.63 ms | 13.4 MB | 1.8 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | report-workbook-datatable-core | EPPlus | 721.72 ms | 220.7 MB | 1.5 MB | OfficeIMO.Excel | 983.1% slower than OfficeIMO |
| 25000 | package-profile | report-workbook-datatable-core | ClosedXML | 1.51 s | 812.7 MB | 1.5 MB | OfficeIMO.Excel | 2167.7% slower than OfficeIMO |
| 25000 | package-profile | write-blog-2023-20-string-columns | LargeXlsx | 42.28 ms | 10.5 MB | 2.4 MB | LargeXlsx | 13.5% faster than OfficeIMO |
| 25000 | package-profile | write-blog-2023-20-string-columns | OfficeIMO.Excel | 48.91 ms | 11.4 MB | 2.2 MB | LargeXlsx | Loss +15.7% |
| 25000 | package-profile | write-blog-2023-20-string-columns | MiniExcel | 157.40 ms | 221.6 MB | 2.4 MB | LargeXlsx | 221.9% slower than OfficeIMO |
| 25000 | package-profile | write-blog-2023-20-string-columns | ClosedXML | 923.70 ms | 742.0 MB | 2.5 MB | LargeXlsx | 1788.7% slower than OfficeIMO |
| 25000 | package-profile | write-bulk-report | OfficeIMO.Excel | 47.67 ms | 11.3 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-bulk-report | MiniExcel | 96.96 ms | 122.6 MB | 1.5 MB | OfficeIMO.Excel | 103.4% slower than OfficeIMO |
| 25000 | package-profile | write-bulk-report | EPPlus | 594.96 ms | 249.0 MB | 1.1 MB | OfficeIMO.Excel | 1148.0% slower than OfficeIMO |
| 25000 | package-profile | write-bulk-report | ClosedXML | 1.17 s | 552.7 MB | 1.1 MB | OfficeIMO.Excel | 2352.5% slower than OfficeIMO |
| 25000 | package-profile | write-cellformula | OfficeIMO.Excel | 29.76 ms | 9.9 MB | 670.3 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellformula | ClosedXML | 265.63 ms | 111.2 MB | 643.2 KB | OfficeIMO.Excel | 792.4% slower than OfficeIMO |
| 25000 | package-profile | write-cellformula | EPPlus | 526.43 ms | 137.4 MB | 593.9 KB | OfficeIMO.Excel | 1668.6% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-empty-strings | OfficeIMO.Excel | 16.61 ms | 6.7 MB | 451.4 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-empty-strings | ClosedXML | 157.36 ms | 90.7 MB | 398.1 KB | OfficeIMO.Excel | 847.4% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-empty-strings | EPPlus | 211.34 ms | 72.7 MB | 390.6 KB | OfficeIMO.Excel | 1172.4% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-numbers | OfficeIMO.Excel | 19.93 ms | 5.8 MB | 462.6 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-numbers | ClosedXML | 138.15 ms | 82.2 MB | 411.4 KB | OfficeIMO.Excel | 593.3% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-numbers | EPPlus | 280.33 ms | 84.4 MB | 406.5 KB | OfficeIMO.Excel | 1306.8% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-mixed | OfficeIMO.Excel | 25.30 ms | 8.1 MB | 585.1 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-object-mixed | ClosedXML | 207.93 ms | 108.5 MB | 532.9 KB | OfficeIMO.Excel | 721.9% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-mixed | EPPlus | 316.62 ms | 110.6 MB | 544.3 KB | OfficeIMO.Excel | 1151.5% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-sparse | OfficeIMO.Excel | 26.63 ms | 7.2 MB | 607.1 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-object-sparse | ClosedXML | 203.87 ms | 102.8 MB | 468.0 KB | OfficeIMO.Excel | 665.5% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-sparse | EPPlus | 295.96 ms | 103.8 MB | 494.4 KB | OfficeIMO.Excel | 1011.3% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 23.82 ms | 7.2 MB | 607.1 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-object-sparse-batch | ClosedXML | 209.02 ms | 102.8 MB | 468.0 KB | OfficeIMO.Excel | 777.4% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-sparse-batch | EPPlus | 319.47 ms | 103.8 MB | 494.4 KB | OfficeIMO.Excel | 1241.0% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-scalars | OfficeIMO.Excel | 15.47 ms | 6.0 MB | 441.9 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-scalars | ClosedXML | 132.52 ms | 80.7 MB | 394.9 KB | OfficeIMO.Excel | 756.7% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-scalars | EPPlus | 309.48 ms | 83.1 MB | 379.3 KB | OfficeIMO.Excel | 1900.7% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings | OfficeIMO.Excel | 26.95 ms | 15.0 MB | 527.8 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-strings | ClosedXML | 241.21 ms | 101.8 MB | 460.1 KB | OfficeIMO.Excel | 795.2% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings | EPPlus | 385.67 ms | 82.4 MB | 444.7 KB | OfficeIMO.Excel | 1331.3% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings-distinct | OfficeIMO.Excel | 16.89 ms | 13.5 MB | 499.0 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-strings-distinct | ClosedXML | 201.59 ms | 128.4 MB | 555.3 KB | OfficeIMO.Excel | 1093.9% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings-distinct | EPPlus | 286.84 ms | 95.4 MB | 565.1 KB | OfficeIMO.Excel | 1598.8% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings-repeated | OfficeIMO.Excel | 17.02 ms | 7.3 MB | 376.0 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-strings-repeated | ClosedXML | 139.32 ms | 82.5 MB | 331.8 KB | OfficeIMO.Excel | 718.6% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings-repeated | EPPlus | 241.25 ms | 68.4 MB | 300.8 KB | OfficeIMO.Excel | 1317.6% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-temporal | OfficeIMO.Excel | 28.45 ms | 7.3 MB | 620.5 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-temporal | ClosedXML | 207.47 ms | 87.2 MB | 483.0 KB | OfficeIMO.Excel | 629.2% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-temporal | EPPlus | 289.33 ms | 101.4 MB | 495.1 KB | OfficeIMO.Excel | 917.0% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 10.67 ms | 3.4 MB | 443.4 KB | LargeXlsx | 19.6% faster than OfficeIMO |
| 25000 | package-profile | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 13.28 ms | 6.8 MB | 455.5 KB | LargeXlsx | Loss +24.4% |
| 25000 | package-profile | write-cellvalues-headerless-rectangle-direct | ClosedXML | 193.91 ms | 93.8 MB | 467.5 KB | LargeXlsx | 1359.7% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-headerless-rectangle-direct | EPPlus | 284.85 ms | 85.4 MB | 484.1 KB | LargeXlsx | 2044.3% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 44.56 ms | 15.7 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalues-rectangle-direct | LargeXlsx | 55.26 ms | 5.5 MB | 1.4 MB | OfficeIMO.Excel | 24.0% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | MiniExcel | 120.41 ms | 91.1 MB | 1.5 MB | OfficeIMO.Excel | 170.2% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | ClosedXML | 451.48 ms | 205.7 MB | 1.1 MB | OfficeIMO.Excel | 913.2% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | EPPlus | 655.73 ms | 206.9 MB | 1.1 MB | OfficeIMO.Excel | 1371.6% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | Sylvan.Data.Excel | 42.77 ms | 5.6 MB | 755.4 KB | Sylvan.Data.Excel | 15.7% faster than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | LargeXlsx | 46.73 ms | 8.2 MB | 1.4 MB | Sylvan.Data.Excel | 7.9% faster than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | OfficeIMO.Excel | 50.74 ms | 12.7 MB | 1.4 MB | Sylvan.Data.Excel | Loss +18.6% |
| 25000 | package-profile | write-datareader-plain | MiniExcel | 100.30 ms | 90.0 MB | 1.5 MB | Sylvan.Data.Excel | 97.7% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | ClosedXML | 365.56 ms | 101.8 MB | 1.1 MB | Sylvan.Data.Excel | 620.5% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | EPPlus | 441.71 ms | 114.7 MB | 1.1 MB | Sylvan.Data.Excel | 770.5% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table | OfficeIMO.Excel | 48.21 ms | 12.7 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-datareader-table | MiniExcel | 106.36 ms | 90.0 MB | 1.5 MB | OfficeIMO.Excel | 120.6% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table | EPPlus | 439.80 ms | 114.7 MB | 1.1 MB | OfficeIMO.Excel | 812.3% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table | ClosedXML | 485.38 ms | 169.3 MB | 1.1 MB | OfficeIMO.Excel | 906.8% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table-autofit | OfficeIMO.Excel | 55.01 ms | 12.7 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-datareader-table-autofit | MiniExcel | 101.47 ms | 121.6 MB | 1.5 MB | OfficeIMO.Excel | 84.5% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table-autofit | EPPlus | 471.82 ms | 156.0 MB | 1.1 MB | OfficeIMO.Excel | 757.7% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table-autofit | ClosedXML | 1.04 s | 552.9 MB | 1.1 MB | OfficeIMO.Excel | 1786.8% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-direct-export | OfficeIMO.Excel | 41.29 ms | 9.6 MB | 1.3 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-dataset-sparse-direct-export | LargeXlsx | 46.59 ms | 9.0 MB | 1.6 MB | OfficeIMO.Excel | 12.8% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-direct-export | MiniExcel | 122.25 ms | 105.6 MB | 1.8 MB | OfficeIMO.Excel | 196.1% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-direct-export | EPPlus | 585.11 ms | 132.5 MB | 1.4 MB | OfficeIMO.Excel | 1317.0% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-direct-export | ClosedXML | 675.09 ms | 273.8 MB | 1.5 MB | OfficeIMO.Excel | 1534.9% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-tables | OfficeIMO.Excel | 47.16 ms | 13.1 MB | 1.8 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-dataset-sparse-tables | MiniExcel | 123.44 ms | 105.6 MB | 1.8 MB | OfficeIMO.Excel | 161.7% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-tables | EPPlus | 618.63 ms | 132.5 MB | 1.4 MB | OfficeIMO.Excel | 1211.7% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-tables | ClosedXML | 632.51 ms | 273.8 MB | 1.5 MB | OfficeIMO.Excel | 1241.1% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables | OfficeIMO.Excel | 48.11 ms | 10.0 MB | 1.3 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-dataset-tables | MiniExcel | 105.90 ms | 94.8 MB | 1.5 MB | OfficeIMO.Excel | 120.1% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables | EPPlus | 469.53 ms | 108.2 MB | 1.1 MB | OfficeIMO.Excel | 875.9% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables | ClosedXML | 492.47 ms | 168.0 MB | 1.1 MB | OfficeIMO.Excel | 923.5% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables-autofit | OfficeIMO.Excel | 48.13 ms | 10.1 MB | 1.3 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-dataset-tables-autofit | MiniExcel | 107.20 ms | 125.9 MB | 1.5 MB | OfficeIMO.Excel | 122.7% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables-autofit | EPPlus | 498.11 ms | 190.8 MB | 1.1 MB | OfficeIMO.Excel | 935.0% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables-autofit | ClosedXML | 962.62 ms | 537.2 MB | 1.1 MB | OfficeIMO.Excel | 1900.1% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | LargeXlsx | 47.80 ms | 9.3 MB | 1.4 MB | LargeXlsx | 19.7% faster than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | OfficeIMO.Excel | 59.54 ms | 12.4 MB | 1.4 MB | LargeXlsx | Loss +24.6% |
| 25000 | package-profile | write-datatable-direct | MiniExcel | 115.01 ms | 90.2 MB | 1.5 MB | LargeXlsx | 93.2% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | ClosedXML | 405.60 ms | 101.8 MB | 1.1 MB | LargeXlsx | 581.2% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | EPPlus | 537.54 ms | 114.7 MB | 1.1 MB | LargeXlsx | 802.8% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-table-direct | OfficeIMO.Excel | 42.27 ms | 12.4 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-datatable-table-direct | MiniExcel | 99.61 ms | 90.2 MB | 1.5 MB | OfficeIMO.Excel | 135.6% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-table-direct | EPPlus | 435.10 ms | 114.7 MB | 1.1 MB | OfficeIMO.Excel | 929.3% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-table-direct | ClosedXML | 465.71 ms | 169.3 MB | 1.1 MB | OfficeIMO.Excel | 1001.7% slower than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | LargeXlsx | 37.82 ms | 5.5 MB | 1.4 MB | LargeXlsx | 19.2% faster than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 46.79 ms | 12.6 MB | 1.4 MB | LargeXlsx | Loss +23.7% |
| 25000 | package-profile | write-fluent-rowsfrom-direct | MiniExcel | 93.77 ms | 91.1 MB | 1.5 MB | LargeXlsx | 100.4% slower than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | ClosedXML | 366.98 ms | 101.8 MB | 1.1 MB | LargeXlsx | 684.3% slower than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | EPPlus | 492.49 ms | 114.7 MB | 1.1 MB | LargeXlsx | 952.6% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 66.78 ms | 11.2 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 745.50 ms | 156.0 MB | 1.1 MB | OfficeIMO.Excel | 1016.3% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 1.10 s | 485.3 MB | 1.1 MB | OfficeIMO.Excel | 1552.8% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | LargeXlsx | 38.19 ms | 5.5 MB | 1.4 MB | LargeXlsx | 10.2% faster than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | OfficeIMO.Excel | 42.53 ms | 11.2 MB | 1.4 MB | LargeXlsx | Loss +11.4% |
| 25000 | package-profile | write-insertobjects-direct | MiniExcel | 91.58 ms | 91.1 MB | 1.5 MB | LargeXlsx | 115.3% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | ClosedXML | 352.64 ms | 101.8 MB | 1.1 MB | LargeXlsx | 729.1% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | EPPlus | 511.31 ms | 114.7 MB | 1.1 MB | LargeXlsx | 1102.2% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 60.51 ms | 9.9 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 533.66 ms | 156.0 MB | 1.1 MB | OfficeIMO.Excel | 781.9% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 942.44 ms | 485.3 MB | 1.1 MB | OfficeIMO.Excel | 1457.4% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 47.87 ms | 5.5 MB | 1.4 MB | LargeXlsx | 26.3% faster than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 65.00 ms | 9.9 MB | 1.4 MB | LargeXlsx | Loss +35.8% |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | MiniExcel | 112.24 ms | 91.1 MB | 1.5 MB | LargeXlsx | 72.7% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | ClosedXML | 417.42 ms | 101.8 MB | 1.1 MB | LargeXlsx | 542.2% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | EPPlus | 623.67 ms | 114.7 MB | 1.1 MB | LargeXlsx | 859.5% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | LargeXlsx | 39.99 ms | 5.5 MB | 1.4 MB | LargeXlsx | 35.5% faster than OfficeIMO |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | OfficeIMO.Excel | 62.00 ms | 15.4 MB | 1.4 MB | LargeXlsx | Loss +55.0% |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | MiniExcel | 92.30 ms | 91.1 MB | 1.5 MB | LargeXlsx | 48.9% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | ClosedXML | 369.70 ms | 101.8 MB | 1.1 MB | LargeXlsx | 496.3% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | EPPlus | 507.55 ms | 114.7 MB | 1.1 MB | LargeXlsx | 718.7% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 48.05 ms | 11.2 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 525.77 ms | 135.1 MB | 1.1 MB | OfficeIMO.Excel | 994.1% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 599.75 ms | 269.0 MB | 1.1 MB | OfficeIMO.Excel | 1148.1% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-mixed-objects-direct | LargeXlsx | 62.30 ms | 5.9 MB | 1.8 MB | LargeXlsx | 10.8% faster than OfficeIMO |
| 25000 | package-profile | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 69.88 ms | 10.3 MB | 1.8 MB | LargeXlsx | Loss +12.2% |
| 25000 | package-profile | write-powershell-mixed-objects-direct | MiniExcel | 153.24 ms | 111.3 MB | 1.9 MB | LargeXlsx | 119.3% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-mixed-objects-direct | ClosedXML | 591.88 ms | 175.3 MB | 1.5 MB | LargeXlsx | 747.0% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-mixed-objects-direct | EPPlus | 775.13 ms | 141.5 MB | 1.4 MB | LargeXlsx | 1009.3% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-mixed-direct | LargeXlsx | 59.05 ms | 5.9 MB | 1.8 MB | LargeXlsx | 27.1% faster than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 81.03 ms | 9.7 MB | 1.8 MB | LargeXlsx | Loss +37.2% |
| 25000 | package-profile | write-powershell-psobject-mixed-direct | MiniExcel | 150.22 ms | 111.3 MB | 1.9 MB | LargeXlsx | 85.4% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-mixed-direct | ClosedXML | 608.27 ms | 175.3 MB | 1.5 MB | LargeXlsx | 650.6% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-mixed-direct | EPPlus | 826.02 ms | 141.5 MB | 1.4 MB | LargeXlsx | 919.4% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 275.51 ms | 35.3 MB | 6.6 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-powershell-psobject-wide-direct | LargeXlsx | 297.83 ms | 22.7 MB | 6.5 MB | OfficeIMO.Excel | 8.1% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-wide-direct | MiniExcel | 495.23 ms | 339.8 MB | 6.8 MB | OfficeIMO.Excel | 79.8% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-wide-direct | ClosedXML | 1.73 s | 476.0 MB | 6.0 MB | OfficeIMO.Excel | 526.4% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-wide-direct | EPPlus | 2.22 s | 549.8 MB | 5.3 MB | OfficeIMO.Excel | 705.6% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | LargeXlsx | 11.63 ms | 2.7 MB |  | LargeXlsx | 23.4% faster than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | OfficeIMO.Excel | 15.18 ms | 10.6 MB |  | LargeXlsx | Loss +30.5% |
| 25000 | speed-comparison | append-plain-rows | MiniExcel | 37.69 ms | 56.9 MB |  | LargeXlsx | 148.2% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | EPPlus 4.5.3.3 | 115.51 ms | 0 B |  | LargeXlsx | 660.7% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | ClosedXML | 142.31 ms | 101.8 MB |  | LargeXlsx | 837.2% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | EPPlus | 224.80 ms | 98.0 MB |  | LargeXlsx | 1380.4% slower than OfficeIMO |
| 25000 | speed-comparison | autofit-existing | OfficeIMO.Excel | 101.95 ms | 15.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | autofit-existing | EPPlus 4.5.3.3 | 613.54 ms | 0 B |  | OfficeIMO.Excel | 501.8% slower than OfficeIMO |
| 25000 | speed-comparison | autofit-existing | EPPlus | 700.21 ms | 245.1 MB |  | OfficeIMO.Excel | 586.9% slower than OfficeIMO |
| 25000 | speed-comparison | autofit-existing | ClosedXML | 1.96 s | 810.4 MB |  | OfficeIMO.Excel | 1824.8% slower than OfficeIMO |
| 25000 | speed-comparison | build-object-datatable-dictionaries | OfficeIMO.Excel | 9.68 ms | 5.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | build-object-datatable-typed | OfficeIMO.Excel | 8.13 ms | 7.9 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-cells | OfficeIMO.Excel | 56.77 ms | 24.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-cells | EPPlus | 323.34 ms | 183.0 MB |  | OfficeIMO.Excel | 469.6% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-cells | ClosedXML | 386.18 ms | 162.6 MB |  | OfficeIMO.Excel | 580.3% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-first-column-from-wide-sheet | OfficeIMO.Excel | 35.46 ms | 3.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-first-column-from-wide-sheet | EPPlus | 330.23 ms | 112.8 MB |  | OfficeIMO.Excel | 831.3% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-first-column-from-wide-sheet | ClosedXML | 345.81 ms | 147.4 MB |  | OfficeIMO.Excel | 875.2% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-range | OfficeIMO.Excel | 51.77 ms | 24.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-range | ClosedXML | 348.20 ms | 162.6 MB |  | OfficeIMO.Excel | 572.5% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-range | EPPlus | 360.84 ms | 183.0 MB |  | OfficeIMO.Excel | 596.9% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-top-range | OfficeIMO.Excel | 0.64 ms | 285.5 KB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-top-range | EPPlus | 274.16 ms | 103.1 MB |  | OfficeIMO.Excel | 43066.3% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-top-range | ClosedXML | 323.07 ms | 145.9 MB |  | OfficeIMO.Excel | 50766.6% slower than OfficeIMO |
| 25000 | speed-comparison | formula-heavy-read | OfficeIMO.Excel | 19.74 ms | 6.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | formula-heavy-read | EPPlus 4.5.3.3 | 76.43 ms | 0 B |  | OfficeIMO.Excel | 287.1% slower than OfficeIMO |
| 25000 | speed-comparison | formula-heavy-read | EPPlus | 169.52 ms | 69.2 MB |  | OfficeIMO.Excel | 758.6% slower than OfficeIMO |
| 25000 | speed-comparison | formula-heavy-read | ClosedXML | 171.06 ms | 77.7 MB |  | OfficeIMO.Excel | 766.4% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | OfficeIMO.Excel | 14.81 ms | 15.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | large-shared-strings | MiniExcel | 31.06 ms | 72.0 MB |  | OfficeIMO.Excel | 109.6% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | EPPlus 4.5.3.3 | 97.28 ms | 0 B |  | OfficeIMO.Excel | 556.6% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | ClosedXML | 115.98 ms | 101.8 MB |  | OfficeIMO.Excel | 682.9% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | EPPlus | 233.73 ms | 82.4 MB |  | OfficeIMO.Excel | 1477.7% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | Sylvan.Data.Excel | 1.03 ms | 316.6 KB |  | Sylvan.Data.Excel | 7.4% faster than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | OfficeIMO.Excel | 1.11 ms | 177.4 KB |  | Sylvan.Data.Excel | Loss +8.0% |
| 25000 | speed-comparison | large-sparse-column-read | ExcelDataReader | 1.66 ms | 4.0 MB |  | Sylvan.Data.Excel | 49.3% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | ClosedXML | 3.40 ms | 4.3 MB |  | Sylvan.Data.Excel | 205.5% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | MiniExcel | 13.28 ms | 45.1 MB |  | Sylvan.Data.Excel | 1091.9% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | EPPlus 4.5.3.3 | 18.84 ms | 0 B |  | Sylvan.Data.Excel | 1590.8% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | EPPlus | 95.85 ms | 42.1 MB |  | Sylvan.Data.Excel | 8502.1% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | OfficeIMO.Excel | 1.17 ms | 177.5 KB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | large-sparse-row-read | Sylvan.Data.Excel | 1.22 ms | 316.6 KB |  | OfficeIMO.Excel | 4.4% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | ExcelDataReader | 1.90 ms | 4.0 MB |  | OfficeIMO.Excel | 62.8% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | ClosedXML | 7.32 ms | 4.3 MB |  | OfficeIMO.Excel | 526.3% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | MiniExcel | 14.30 ms | 45.1 MB |  | OfficeIMO.Excel | 1123.2% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | EPPlus 4.5.3.3 | 17.41 ms | 0 B |  | OfficeIMO.Excel | 1389.0% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | EPPlus | 102.74 ms | 42.1 MB |  | OfficeIMO.Excel | 8689.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | OfficeIMO.Excel | 33.96 ms | 1.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-bottom-range | Sylvan.Data.Excel | 39.51 ms | 3.5 MB |  | OfficeIMO.Excel | 16.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | ExcelDataReader | 107.22 ms | 59.8 MB |  | OfficeIMO.Excel | 215.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | MiniExcel | 123.57 ms | 182.1 MB |  | OfficeIMO.Excel | 263.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | EPPlus | 232.30 ms | 103.1 MB |  | OfficeIMO.Excel | 584.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | ClosedXML | 314.30 ms | 145.9 MB |  | OfficeIMO.Excel | 825.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | Sylvan.Data.Excel | 40.08 ms | 3.5 MB |  | Sylvan.Data.Excel | 3.8% faster than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | OfficeIMO.Excel | 41.65 ms | 1.1 MB |  | Sylvan.Data.Excel | Loss +3.9% |
| 25000 | speed-comparison | read-bottom-range-stream | ExcelDataReader | 110.65 ms | 59.8 MB |  | Sylvan.Data.Excel | 165.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | MiniExcel | 130.41 ms | 182.1 MB |  | Sylvan.Data.Excel | 213.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | EPPlus | 293.56 ms | 103.1 MB |  | Sylvan.Data.Excel | 604.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | ClosedXML | 361.26 ms | 145.9 MB |  | Sylvan.Data.Excel | 767.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | Sylvan.Data.Excel | 64.73 ms | 18.0 MB |  | Sylvan.Data.Excel | 6.6% faster than OfficeIMO |
| 25000 | speed-comparison | read-datatable | OfficeIMO.Excel | 69.30 ms | 33.8 MB |  | Sylvan.Data.Excel | Loss +7.1% |
| 25000 | speed-comparison | read-datatable | ExcelDataReader | 141.88 ms | 74.3 MB |  | Sylvan.Data.Excel | 104.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | MiniExcel | 159.44 ms | 177.0 MB |  | Sylvan.Data.Excel | 130.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | EPPlus 4.5.3.3 | 216.10 ms | 0 B |  | Sylvan.Data.Excel | 211.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | EPPlus | 301.85 ms | 197.5 MB |  | Sylvan.Data.Excel | 335.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | ClosedXML | 369.47 ms | 174.3 MB |  | Sylvan.Data.Excel | 433.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | OfficeIMO.Excel | 72.82 ms | 3.9 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | Sylvan.Data.Excel | 122.76 ms | 4.2 MB |  | OfficeIMO.Excel | 68.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | MiniExcel | 190.82 ms | 154.9 MB |  | OfficeIMO.Excel | 162.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | ExcelDataReader | 216.25 ms | 59.8 MB |  | OfficeIMO.Excel | 197.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | ClosedXML | 481.16 ms | 147.4 MB |  | OfficeIMO.Excel | 560.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | EPPlus | 527.98 ms | 112.8 MB |  | OfficeIMO.Excel | 625.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | OfficeIMO.Excel | 98.10 ms | 23.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-objects | Sylvan.Data.Excel | 108.46 ms | 5.7 MB |  | OfficeIMO.Excel | 10.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | EPPlus 4.5.3.3 | 172.90 ms | 0 B |  | OfficeIMO.Excel | 76.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | ExcelDataReader | 275.48 ms | 62.0 MB |  | OfficeIMO.Excel | 180.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | MiniExcel | 398.09 ms | 179.4 MB |  | OfficeIMO.Excel | 305.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | EPPlus | 555.37 ms | 194.9 MB |  | OfficeIMO.Excel | 466.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | ClosedXML | 680.52 ms | 161.7 MB |  | OfficeIMO.Excel | 593.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | OfficeIMO.Excel | 50.40 ms | 22.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-objects-stream | Sylvan.Data.Excel | 61.18 ms | 5.2 MB |  | OfficeIMO.Excel | 21.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | ExcelDataReader | 145.73 ms | 61.5 MB |  | OfficeIMO.Excel | 189.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | MiniExcel | 148.76 ms | 178.9 MB |  | OfficeIMO.Excel | 195.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | EPPlus 4.5.3.3 | 185.17 ms | 0 B |  | OfficeIMO.Excel | 267.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | EPPlus | 330.94 ms | 194.7 MB |  | OfficeIMO.Excel | 556.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | ClosedXML | 461.81 ms | 161.5 MB |  | OfficeIMO.Excel | 816.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | Sylvan.Data.Excel | 84.14 ms | 3.5 MB |  | Sylvan.Data.Excel | 4.3% faster than OfficeIMO |
| 25000 | speed-comparison | read-range | OfficeIMO.Excel | 87.94 ms | 25.5 MB |  | Sylvan.Data.Excel | Loss +4.5% |
| 25000 | speed-comparison | read-range | MiniExcel | 179.25 ms | 182.1 MB |  | Sylvan.Data.Excel | 103.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | ExcelDataReader | 183.81 ms | 59.8 MB |  | Sylvan.Data.Excel | 109.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | EPPlus 4.5.3.3 | 187.35 ms | 0 B |  | Sylvan.Data.Excel | 113.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | ClosedXML | 475.77 ms | 159.8 MB |  | Sylvan.Data.Excel | 441.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | EPPlus | 536.57 ms | 183.0 MB |  | Sylvan.Data.Excel | 510.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | Sylvan.Data.Excel | 60.33 ms | 4.4 MB |  | Sylvan.Data.Excel | 4.4% faster than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | OfficeIMO.Excel | 63.12 ms | 26.1 MB |  | Sylvan.Data.Excel | Loss +4.6% |
| 25000 | speed-comparison | read-range-decimal | ExcelDataReader | 121.99 ms | 59.8 MB |  | Sylvan.Data.Excel | 93.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | MiniExcel | 136.11 ms | 182.1 MB |  | Sylvan.Data.Excel | 115.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | EPPlus | 280.17 ms | 183.0 MB |  | Sylvan.Data.Excel | 343.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | ClosedXML | 362.61 ms | 159.8 MB |  | Sylvan.Data.Excel | 474.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | Sylvan.Data.Excel | 47.95 ms | 3.5 MB |  | Sylvan.Data.Excel | 11.1% faster than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | OfficeIMO.Excel | 53.97 ms | 26.3 MB |  | Sylvan.Data.Excel | Loss +12.5% |
| 25000 | speed-comparison | read-range-stream | ExcelDataReader | 125.44 ms | 59.8 MB |  | Sylvan.Data.Excel | 132.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | MiniExcel | 131.07 ms | 182.1 MB |  | Sylvan.Data.Excel | 142.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | EPPlus 4.5.3.3 | 219.92 ms | 0 B |  | Sylvan.Data.Excel | 307.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | EPPlus | 276.84 ms | 183.0 MB |  | Sylvan.Data.Excel | 412.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | ClosedXML | 361.79 ms | 159.8 MB |  | Sylvan.Data.Excel | 570.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | Sylvan.Data.Excel | 0.52 ms | 348.5 KB |  | Sylvan.Data.Excel | 15.6% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range | OfficeIMO.Excel | 0.61 ms | 296.3 KB |  | Sylvan.Data.Excel | Loss +18.5% |
| 25000 | speed-comparison | read-top-range | MiniExcel | 0.89 ms | 869.0 KB |  | Sylvan.Data.Excel | 44.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | ExcelDataReader | 46.69 ms | 16.7 MB |  | Sylvan.Data.Excel | 7493.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | EPPlus 4.5.3.3 | 163.29 ms | 0 B |  | Sylvan.Data.Excel | 26456.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | EPPlus | 271.67 ms | 103.1 MB |  | Sylvan.Data.Excel | 44084.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | ClosedXML | 371.92 ms | 145.9 MB |  | Sylvan.Data.Excel | 60387.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | OfficeIMO.Excel | 0.75 ms | 302.3 KB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-top-range-stream | MiniExcel | 1.06 ms | 869.0 KB |  | OfficeIMO.Excel | 40.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | Sylvan.Data.Excel | 1.15 ms | 348.5 KB |  | OfficeIMO.Excel | 52.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | ExcelDataReader | 56.83 ms | 16.7 MB |  | OfficeIMO.Excel | 7453.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | EPPlus 4.5.3.3 | 161.00 ms | 0 B |  | OfficeIMO.Excel | 21297.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | EPPlus | 301.87 ms | 103.1 MB |  | OfficeIMO.Excel | 40019.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | ClosedXML | 465.83 ms | 145.9 MB |  | OfficeIMO.Excel | 61810.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | Sylvan.Data.Excel | 0.44 ms | 348.5 KB |  | Sylvan.Data.Excel | 23.4% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | OfficeIMO.Excel | 0.57 ms | 300.3 KB |  | Sylvan.Data.Excel | Loss +30.5% |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | MiniExcel | 0.87 ms | 869.0 KB |  | Sylvan.Data.Excel | 51.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | ExcelDataReader | 45.96 ms | 16.7 MB |  | Sylvan.Data.Excel | 7905.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | EPPlus | 242.10 ms | 103.1 MB |  | Sylvan.Data.Excel | 42070.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | ClosedXML | 326.86 ms | 145.9 MB |  | Sylvan.Data.Excel | 56833.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-used-range | Sylvan.Data.Excel | 60.27 ms | 3.5 MB |  | Sylvan.Data.Excel | 52.6% faster than OfficeIMO |
| 25000 | speed-comparison | read-used-range | OfficeIMO.Excel | 127.12 ms | 33.4 MB |  | Sylvan.Data.Excel | Loss +110.9% |
| 25000 | speed-comparison | read-used-range | ExcelDataReader | 143.90 ms | 59.8 MB |  | Sylvan.Data.Excel | 13.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-used-range | MiniExcel | 152.95 ms | 182.1 MB |  | Sylvan.Data.Excel | 20.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-used-range | EPPlus | 359.84 ms | 183.0 MB |  | Sylvan.Data.Excel | 183.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-used-range | ClosedXML | 374.17 ms | 159.8 MB |  | Sylvan.Data.Excel | 194.3% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-autofilter | OfficeIMO.Excel | 32.26 ms | 11.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-autofilter | ClosedXML | 294.16 ms | 205.7 MB |  | OfficeIMO.Excel | 811.8% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-autofilter | EPPlus 4.5.3.3 | 367.46 ms | 0 B |  | OfficeIMO.Excel | 1039.0% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-autofilter | EPPlus | 396.06 ms | 206.9 MB |  | OfficeIMO.Excel | 1127.6% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-charts | OfficeIMO.Excel | 33.47 ms | 12.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-charts | EPPlus | 414.56 ms | 209.9 MB |  | OfficeIMO.Excel | 1138.7% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-charts | EPPlus 4.5.3.3 | 456.99 ms | 0 B |  | OfficeIMO.Excel | 1265.4% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-conditional-formatting | OfficeIMO.Excel | 33.86 ms | 11.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-conditional-formatting | ClosedXML | 311.94 ms | 205.8 MB |  | OfficeIMO.Excel | 821.2% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-conditional-formatting | EPPlus | 383.82 ms | 206.9 MB |  | OfficeIMO.Excel | 1033.5% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-conditional-formatting | EPPlus 4.5.3.3 | 399.47 ms | 0 B |  | OfficeIMO.Excel | 1079.7% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-data-validation | OfficeIMO.Excel | 31.38 ms | 11.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-data-validation | ClosedXML | 296.52 ms | 205.7 MB |  | OfficeIMO.Excel | 845.0% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-data-validation | EPPlus | 395.18 ms | 206.9 MB |  | OfficeIMO.Excel | 1159.5% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-data-validation | EPPlus 4.5.3.3 | 434.49 ms | 0 B |  | OfficeIMO.Excel | 1284.8% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-freeze-panes | OfficeIMO.Excel | 31.51 ms | 11.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-freeze-panes | ClosedXML | 293.64 ms | 205.7 MB |  | OfficeIMO.Excel | 832.0% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-freeze-panes | EPPlus 4.5.3.3 | 370.53 ms | 0 B |  | OfficeIMO.Excel | 1076.0% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-freeze-panes | EPPlus | 394.58 ms | 206.9 MB |  | OfficeIMO.Excel | 1152.4% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-pivot-table | OfficeIMO.Excel | 258.35 ms | 128.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-pivot-table | EPPlus 4.5.3.3 | 351.05 ms | 0 B |  | OfficeIMO.Excel | 35.9% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-pivot-table | EPPlus | 446.27 ms | 225.4 MB |  | OfficeIMO.Excel | 72.7% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-all-in-one | OfficeIMO.Excel | 317.70 ms | 130.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-all-in-one | EPPlus | 474.92 ms | 270.6 MB |  | OfficeIMO.Excel | 49.5% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-all-in-one | EPPlus 4.5.3.3 | 511.23 ms | 0 B |  | OfficeIMO.Excel | 60.9% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-chart-first | OfficeIMO.Excel | 94.45 ms | 42.5 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-chart-first | EPPlus | 465.71 ms | 270.6 MB |  | OfficeIMO.Excel | 393.1% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-chart-first | EPPlus 4.5.3.3 | 512.61 ms | 0 B |  | OfficeIMO.Excel | 442.7% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-core | OfficeIMO.Excel | 40.57 ms | 11.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-core | EPPlus | 435.70 ms | 249.1 MB |  | OfficeIMO.Excel | 974.0% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-core | EPPlus 4.5.3.3 | 799.72 ms | 0 B |  | OfficeIMO.Excel | 1871.4% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-core | ClosedXML | 914.32 ms | 664.2 MB |  | OfficeIMO.Excel | 2153.9% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-extra-column | OfficeIMO.Excel | 408.57 ms | 141.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-extra-column | EPPlus | 552.72 ms | 295.7 MB |  | OfficeIMO.Excel | 35.3% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-extra-column | EPPlus 4.5.3.3 | 862.35 ms | 0 B |  | OfficeIMO.Excel | 111.1% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-no-autofit | EPPlus 4.5.3.3 | 252.39 ms | 0 B |  | EPPlus 4.5.3.3 | 6.9% faster than OfficeIMO |
| 25000 | speed-comparison | realworld-report-no-autofit | OfficeIMO.Excel | 271.08 ms | 130.3 MB |  | EPPlus 4.5.3.3 | Loss +7.4% |
| 25000 | speed-comparison | realworld-report-no-autofit | EPPlus | 435.15 ms | 229.3 MB |  | EPPlus 4.5.3.3 | 60.5% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-post-mutation | OfficeIMO.Excel | 289.75 ms | 130.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-post-mutation | EPPlus | 478.18 ms | 270.6 MB |  | OfficeIMO.Excel | 65.0% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-post-mutation | EPPlus 4.5.3.3 | 766.56 ms | 0 B |  | OfficeIMO.Excel | 164.6% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-shuffled-columns | OfficeIMO.Excel | 344.54 ms | 130.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-shuffled-columns | EPPlus | 559.22 ms | 270.6 MB |  | OfficeIMO.Excel | 62.3% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-shuffled-columns | EPPlus 4.5.3.3 | 797.85 ms | 0 B |  | OfficeIMO.Excel | 131.6% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook | OfficeIMO.Excel | 493.30 ms | 171.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | report-workbook | EPPlus | 722.33 ms | 356.2 MB |  | OfficeIMO.Excel | 46.4% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook | EPPlus 4.5.3.3 | 1.01 s | 0 B |  | OfficeIMO.Excel | 104.9% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-core | OfficeIMO.Excel | 52.78 ms | 10.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | report-workbook-core | EPPlus | 571.33 ms | 334.8 MB |  | OfficeIMO.Excel | 982.4% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-core | EPPlus 4.5.3.3 | 944.71 ms | 0 B |  | OfficeIMO.Excel | 1689.8% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-core | ClosedXML | 1.23 s | 952.9 MB |  | OfficeIMO.Excel | 2235.4% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-datatable | OfficeIMO.Excel | 450.29 ms | 173.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | report-workbook-datatable | EPPlus | 581.27 ms | 242.0 MB |  | OfficeIMO.Excel | 29.1% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-datatable | EPPlus 4.5.3.3 | 856.43 ms | 0 B |  | OfficeIMO.Excel | 90.2% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-datatable-core | OfficeIMO.Excel | 53.43 ms | 13.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | report-workbook-datatable-core | EPPlus | 620.91 ms | 220.7 MB |  | OfficeIMO.Excel | 1062.1% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-datatable-core | EPPlus 4.5.3.3 | 677.08 ms | 0 B |  | OfficeIMO.Excel | 1167.3% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-datatable-core | ClosedXML | 1.22 s | 812.7 MB |  | OfficeIMO.Excel | 2182.6% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | Sylvan.Data.Excel | 17.14 ms | 1.9 MB |  | Sylvan.Data.Excel | 8.0% faster than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | OfficeIMO.Excel | 18.62 ms | 9.0 MB |  | Sylvan.Data.Excel | Loss +8.6% |
| 25000 | speed-comparison | shared-string-read | ExcelDataReader | 41.87 ms | 24.4 MB |  | Sylvan.Data.Excel | 124.8% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | MiniExcel | 50.58 ms | 72.7 MB |  | Sylvan.Data.Excel | 171.6% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | EPPlus 4.5.3.3 | 95.07 ms | 0 B |  | Sylvan.Data.Excel | 410.5% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | ClosedXML | 148.53 ms | 88.3 MB |  | Sylvan.Data.Excel | 697.6% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | EPPlus | 151.94 ms | 87.3 MB |  | Sylvan.Data.Excel | 716.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-blog-2023-20-string-columns | LargeXlsx | 53.45 ms | 10.5 MB |  | LargeXlsx | 7.5% faster than OfficeIMO |
| 25000 | speed-comparison | write-blog-2023-20-string-columns | OfficeIMO.Excel | 57.79 ms | 11.4 MB |  | LargeXlsx | Loss +8.1% |
| 25000 | speed-comparison | write-blog-2023-20-string-columns | MiniExcel | 207.19 ms | 221.6 MB |  | LargeXlsx | 258.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-blog-2023-20-string-columns | ClosedXML | 1.15 s | 742.0 MB |  | LargeXlsx | 1883.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | OfficeIMO.Excel | 38.72 ms | 11.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-bulk-report | MiniExcel | 70.28 ms | 122.6 MB |  | OfficeIMO.Excel | 81.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | EPPlus | 399.97 ms | 249.0 MB |  | OfficeIMO.Excel | 933.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | EPPlus 4.5.3.3 | 531.79 ms | 0 B |  | OfficeIMO.Excel | 1273.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | ClosedXML | 789.63 ms | 552.7 MB |  | OfficeIMO.Excel | 1939.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellformula | OfficeIMO.Excel | 27.62 ms | 9.9 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellformula | EPPlus 4.5.3.3 | 141.51 ms | 0 B |  | OfficeIMO.Excel | 412.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellformula | ClosedXML | 232.97 ms | 111.2 MB |  | OfficeIMO.Excel | 743.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellformula | EPPlus | 385.41 ms | 137.4 MB |  | OfficeIMO.Excel | 1295.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-empty-strings | OfficeIMO.Excel | 17.09 ms | 6.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-empty-strings | ClosedXML | 152.66 ms | 90.7 MB |  | OfficeIMO.Excel | 793.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-empty-strings | EPPlus | 215.18 ms | 72.7 MB |  | OfficeIMO.Excel | 1159.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-numbers | OfficeIMO.Excel | 20.11 ms | 5.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-numbers | EPPlus 4.5.3.3 | 106.95 ms | 0 B |  | OfficeIMO.Excel | 431.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-numbers | ClosedXML | 133.69 ms | 82.2 MB |  | OfficeIMO.Excel | 564.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-numbers | EPPlus | 256.54 ms | 84.4 MB |  | OfficeIMO.Excel | 1175.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-mixed | OfficeIMO.Excel | 21.80 ms | 8.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-object-mixed | EPPlus 4.5.3.3 | 118.04 ms | 0 B |  | OfficeIMO.Excel | 441.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-mixed | ClosedXML | 190.87 ms | 108.5 MB |  | OfficeIMO.Excel | 775.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-mixed | EPPlus | 240.28 ms | 110.6 MB |  | OfficeIMO.Excel | 1002.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-sparse | OfficeIMO.Excel | 23.70 ms | 7.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-object-sparse | ClosedXML | 175.28 ms | 102.8 MB |  | OfficeIMO.Excel | 639.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-sparse | EPPlus | 237.84 ms | 103.8 MB |  | OfficeIMO.Excel | 903.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 21.95 ms | 7.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-object-sparse-batch | ClosedXML | 180.85 ms | 102.8 MB |  | OfficeIMO.Excel | 724.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-sparse-batch | EPPlus | 257.43 ms | 103.8 MB |  | OfficeIMO.Excel | 1072.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-scalars | OfficeIMO.Excel | 15.30 ms | 6.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-scalars | EPPlus 4.5.3.3 | 104.52 ms | 0 B |  | OfficeIMO.Excel | 583.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-scalars | ClosedXML | 130.43 ms | 80.7 MB |  | OfficeIMO.Excel | 752.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-scalars | EPPlus | 284.44 ms | 83.1 MB |  | OfficeIMO.Excel | 1759.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings | OfficeIMO.Excel | 22.75 ms | 15.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-strings | EPPlus 4.5.3.3 | 96.25 ms | 0 B |  | OfficeIMO.Excel | 323.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings | ClosedXML | 144.35 ms | 101.8 MB |  | OfficeIMO.Excel | 534.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings | EPPlus | 221.35 ms | 82.4 MB |  | OfficeIMO.Excel | 873.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings-distinct | OfficeIMO.Excel | 18.28 ms | 13.5 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-strings-distinct | ClosedXML | 202.58 ms | 128.4 MB |  | OfficeIMO.Excel | 1008.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings-distinct | EPPlus | 222.71 ms | 95.4 MB |  | OfficeIMO.Excel | 1118.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings-repeated | OfficeIMO.Excel | 15.99 ms | 7.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-strings-repeated | ClosedXML | 133.66 ms | 82.5 MB |  | OfficeIMO.Excel | 736.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings-repeated | EPPlus | 187.59 ms | 68.4 MB |  | OfficeIMO.Excel | 1073.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-temporal | OfficeIMO.Excel | 23.99 ms | 7.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-temporal | EPPlus 4.5.3.3 | 106.13 ms | 0 B |  | OfficeIMO.Excel | 342.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-temporal | ClosedXML | 184.93 ms | 87.2 MB |  | OfficeIMO.Excel | 670.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-temporal | EPPlus | 233.42 ms | 101.4 MB |  | OfficeIMO.Excel | 873.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 12.82 ms | 3.4 MB |  | LargeXlsx | 12.1% faster than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 14.58 ms | 6.8 MB |  | LargeXlsx | Loss +13.8% |
| 25000 | speed-comparison | write-cellvalues-headerless-rectangle-direct | ClosedXML | 151.27 ms | 93.8 MB |  | LargeXlsx | 937.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-headerless-rectangle-direct | EPPlus | 228.81 ms | 85.4 MB |  | LargeXlsx | 1469.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | LargeXlsx | 44.35 ms | 5.5 MB |  | LargeXlsx | 9.9% faster than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 49.20 ms | 15.7 MB |  | LargeXlsx | Loss +10.9% |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | MiniExcel | 91.57 ms | 91.1 MB |  | LargeXlsx | 86.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 272.30 ms | 0 B |  | LargeXlsx | 453.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | ClosedXML | 419.05 ms | 205.7 MB |  | LargeXlsx | 751.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus | 490.75 ms | 206.9 MB |  | LargeXlsx | 897.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-sparse-rectangle-direct | OfficeIMO.Excel | 21.40 ms | 7.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalues-sparse-rectangle-direct | ClosedXML | 183.67 ms | 102.8 MB |  | OfficeIMO.Excel | 758.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-sparse-rectangle-direct | EPPlus | 251.47 ms | 103.8 MB |  | OfficeIMO.Excel | 1075.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | Sylvan.Data.Excel | 40.92 ms | 5.6 MB |  | Sylvan.Data.Excel | 18.6% faster than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | LargeXlsx | 46.96 ms | 8.2 MB |  | Sylvan.Data.Excel | 6.5% faster than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | OfficeIMO.Excel | 50.25 ms | 12.7 MB |  | Sylvan.Data.Excel | Loss +22.8% |
| 25000 | speed-comparison | write-datareader-plain | MiniExcel | 112.84 ms | 90.0 MB |  | Sylvan.Data.Excel | 124.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | EPPlus 4.5.3.3 | 229.11 ms | 0 B |  | Sylvan.Data.Excel | 356.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | ClosedXML | 412.77 ms | 101.8 MB |  | Sylvan.Data.Excel | 721.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | EPPlus | 528.04 ms | 114.7 MB |  | Sylvan.Data.Excel | 950.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | OfficeIMO.Excel | 53.21 ms | 12.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datareader-table | MiniExcel | 119.97 ms | 90.0 MB |  | OfficeIMO.Excel | 125.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | EPPlus 4.5.3.3 | 229.34 ms | 0 B |  | OfficeIMO.Excel | 331.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | ClosedXML | 464.92 ms | 169.3 MB |  | OfficeIMO.Excel | 773.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | EPPlus | 494.11 ms | 114.7 MB |  | OfficeIMO.Excel | 828.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table-autofit | OfficeIMO.Excel | 56.22 ms | 12.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datareader-table-autofit | MiniExcel | 118.37 ms | 121.6 MB |  | OfficeIMO.Excel | 110.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table-autofit | EPPlus 4.5.3.3 | 489.87 ms | 0 B |  | OfficeIMO.Excel | 771.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table-autofit | EPPlus | 651.63 ms | 156.0 MB |  | OfficeIMO.Excel | 1059.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table-autofit | ClosedXML | 1.51 s | 552.9 MB |  | OfficeIMO.Excel | 2583.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-headerless-tables | OfficeIMO.Excel | 38.51 ms | 12.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-headerless-tables | MiniExcel | 97.64 ms | 94.8 MB |  | OfficeIMO.Excel | 153.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-headerless-tables | ClosedXML | 425.52 ms | 168.0 MB |  | OfficeIMO.Excel | 1004.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-headerless-tables | EPPlus | 513.80 ms | 108.6 MB |  | OfficeIMO.Excel | 1234.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | OfficeIMO.Excel | 32.65 ms | 9.6 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | LargeXlsx | 39.32 ms | 9.0 MB |  | OfficeIMO.Excel | 20.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | MiniExcel | 98.65 ms | 105.6 MB |  | OfficeIMO.Excel | 202.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | EPPlus | 474.14 ms | 132.5 MB |  | OfficeIMO.Excel | 1352.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | ClosedXML | 515.34 ms | 273.8 MB |  | OfficeIMO.Excel | 1478.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-tables | OfficeIMO.Excel | 48.69 ms | 13.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-sparse-tables | MiniExcel | 104.23 ms | 105.6 MB |  | OfficeIMO.Excel | 114.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-tables | EPPlus | 509.93 ms | 132.5 MB |  | OfficeIMO.Excel | 947.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-tables | ClosedXML | 633.19 ms | 273.8 MB |  | OfficeIMO.Excel | 1200.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | OfficeIMO.Excel | 34.40 ms | 10.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-tables | MiniExcel | 77.99 ms | 94.8 MB |  | OfficeIMO.Excel | 126.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | EPPlus 4.5.3.3 | 227.84 ms | 0 B |  | OfficeIMO.Excel | 562.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | EPPlus | 319.12 ms | 108.2 MB |  | OfficeIMO.Excel | 827.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | ClosedXML | 357.54 ms | 168.0 MB |  | OfficeIMO.Excel | 939.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables-autofit | OfficeIMO.Excel | 37.89 ms | 10.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-tables-autofit | MiniExcel | 77.22 ms | 125.9 MB |  | OfficeIMO.Excel | 103.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables-autofit | EPPlus | 374.24 ms | 190.8 MB |  | OfficeIMO.Excel | 887.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables-autofit | ClosedXML | 728.51 ms | 537.2 MB |  | OfficeIMO.Excel | 1822.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | LargeXlsx | 30.42 ms | 9.3 MB |  | LargeXlsx | 8.2% faster than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | OfficeIMO.Excel | 33.13 ms | 12.4 MB |  | LargeXlsx | Loss +8.9% |
| 25000 | speed-comparison | write-datatable-direct | MiniExcel | 78.08 ms | 90.2 MB |  | LargeXlsx | 135.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | EPPlus 4.5.3.3 | 218.81 ms | 0 B |  | LargeXlsx | 560.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | ClosedXML | 269.30 ms | 101.8 MB |  | LargeXlsx | 712.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | EPPlus | 317.45 ms | 114.7 MB |  | LargeXlsx | 858.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-object-table-direct | OfficeIMO.Excel | 37.43 ms | 9.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datatable-object-table-direct | MiniExcel | 78.90 ms | 87.6 MB |  | OfficeIMO.Excel | 110.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-object-table-direct | EPPlus | 319.82 ms | 112.0 MB |  | OfficeIMO.Excel | 754.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-object-table-direct | ClosedXML | 363.95 ms | 166.7 MB |  | OfficeIMO.Excel | 872.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | OfficeIMO.Excel | 34.71 ms | 12.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datatable-table-direct | MiniExcel | 79.05 ms | 90.2 MB |  | OfficeIMO.Excel | 127.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | EPPlus 4.5.3.3 | 226.71 ms | 0 B |  | OfficeIMO.Excel | 553.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | EPPlus | 325.33 ms | 114.7 MB |  | OfficeIMO.Excel | 837.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | ClosedXML | 363.47 ms | 169.3 MB |  | OfficeIMO.Excel | 947.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-dictionary-objects-table-direct | OfficeIMO.Excel | 50.71 ms | 14.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | LargeXlsx | 42.05 ms | 5.5 MB |  | LargeXlsx | 7.4% faster than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 45.41 ms | 12.6 MB |  | LargeXlsx | Loss +8.0% |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | MiniExcel | 107.83 ms | 91.1 MB |  | LargeXlsx | 137.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 247.48 ms | 0 B |  | LargeXlsx | 444.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | ClosedXML | 397.45 ms | 101.8 MB |  | LargeXlsx | 775.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus | 490.36 ms | 114.7 MB |  | LargeXlsx | 979.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 49.88 ms | 11.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 517.50 ms | 156.0 MB |  | OfficeIMO.Excel | 937.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 972.48 ms | 485.3 MB |  | OfficeIMO.Excel | 1849.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | LargeXlsx | 41.97 ms | 5.5 MB |  | LargeXlsx | 5.2% faster than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | OfficeIMO.Excel | 44.27 ms | 11.2 MB |  | LargeXlsx | Loss +5.5% |
| 25000 | speed-comparison | write-insertobjects-direct | MiniExcel | 99.77 ms | 91.1 MB |  | LargeXlsx | 125.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | EPPlus 4.5.3.3 | 236.34 ms | 0 B |  | LargeXlsx | 433.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | ClosedXML | 356.55 ms | 101.8 MB |  | LargeXlsx | 705.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | EPPlus | 455.43 ms | 114.7 MB |  | LargeXlsx | 928.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 58.50 ms | 9.9 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 491.66 ms | 156.0 MB |  | OfficeIMO.Excel | 740.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 1.02 s | 485.3 MB |  | OfficeIMO.Excel | 1642.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 38.19 ms | 5.5 MB |  | LargeXlsx | 22.4% faster than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 49.23 ms | 9.9 MB |  | LargeXlsx | Loss +28.9% |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | MiniExcel | 81.82 ms | 91.1 MB |  | LargeXlsx | 66.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | ClosedXML | 352.25 ms | 101.8 MB |  | LargeXlsx | 615.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | EPPlus | 399.54 ms | 114.7 MB |  | LargeXlsx | 711.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 45.87 ms | 11.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 513.26 ms | 135.1 MB |  | OfficeIMO.Excel | 1019.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 585.29 ms | 269.0 MB |  | OfficeIMO.Excel | 1176.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | LargeXlsx | 55.05 ms | 5.9 MB |  | LargeXlsx | 19.8% faster than OfficeIMO |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 68.65 ms | 10.3 MB |  | LargeXlsx | Loss +24.7% |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | MiniExcel | 120.76 ms | 111.3 MB |  | LargeXlsx | 75.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | ClosedXML | 531.28 ms | 175.3 MB |  | LargeXlsx | 673.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | EPPlus | 602.15 ms | 141.5 MB |  | LargeXlsx | 777.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-mixed-direct | LargeXlsx | 53.86 ms | 5.9 MB |  | LargeXlsx | 11.2% faster than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 60.67 ms | 9.7 MB |  | LargeXlsx | Loss +12.7% |
| 25000 | speed-comparison | write-powershell-psobject-mixed-direct | MiniExcel | 115.79 ms | 111.3 MB |  | LargeXlsx | 90.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-mixed-direct | ClosedXML | 509.78 ms | 175.3 MB |  | LargeXlsx | 740.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-mixed-direct | EPPlus | 566.85 ms | 141.5 MB |  | LargeXlsx | 834.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 284.19 ms | 35.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-powershell-psobject-wide-direct | LargeXlsx | 292.23 ms | 22.7 MB |  | OfficeIMO.Excel | 2.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-wide-direct | MiniExcel | 447.67 ms | 339.8 MB |  | OfficeIMO.Excel | 57.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-wide-direct | ClosedXML | 1.53 s | 476.0 MB |  | OfficeIMO.Excel | 436.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-wide-direct | EPPlus | 1.99 s | 549.8 MB |  | OfficeIMO.Excel | 601.5% slower than OfficeIMO |
| 300000 | focused-package-profile | write-blog-2023-20-string-columns | LargeXlsx | 636.47 ms | 93.1 MB | 28.6 MB | LargeXlsx | Win |
| 300000 | focused-package-profile | write-blog-2023-20-string-columns | OfficeIMO.Excel | 701.65 ms | 173.4 MB | 26.6 MB | LargeXlsx | 1.10x vs best |
| 300000 | focused-package-profile | write-blog-2023-20-string-columns | MiniExcel | 2.28 s | 2.46 GB | 28.5 MB | LargeXlsx | 3.58x vs best |
| 300000 | focused-package-profile | write-blog-2023-20-string-columns | ClosedXML | 15.89 s | 8.51 GB | 31.0 MB | LargeXlsx | 24.97x vs best |
