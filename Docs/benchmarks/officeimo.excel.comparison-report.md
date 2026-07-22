# OfficeIMO.Excel Benchmark Report

Generated: 2026-05-31T18:44:43.8878299+00:00
Run mode: quick
Publish: False
Machine: EVOMAGIC ( processors)

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
| 2500 | dense-helloworld-comparison | read | Other | 1 | 1 | dense-helloworld-read-range vs Sylvan.Data.Excel (1.35x) |
| 2500 | package-profile | package | Package size | 44 | 10 | write-insertobjects-legacy-dictionaries-direct vs LargeXlsx (1.52x) |
| 2500 | speed-comparison | mutate | AutoFit and mutation | 2 | 0 |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | 4 | 0 |  |
| 2500 | speed-comparison | other | Object projection | 2 | 0 |  |
| 2500 | speed-comparison | other | Range and table read | 4 | 0 |  |
| 2500 | speed-comparison | other | Real-world feature mix | 6 | 0 |  |
| 2500 | speed-comparison | other | Real-world report | 2 | 0 |  |
| 2500 | speed-comparison | other | Report workbook | 4 | 0 |  |
| 2500 | speed-comparison | read | Formula write/read | 1 | 0 |  |
| 2500 | speed-comparison | read | Other | 1 | 2 | large-sparse-row-read vs Sylvan.Data.Excel (1.52x) |
| 2500 | speed-comparison | read | Range and table read | 4 | 3 | read-used-range vs Sylvan.Data.Excel (2.88x) |
| 2500 | speed-comparison | read | Streaming read | 1 | 3 | read-top-range-stream-small-chunks vs Sylvan.Data.Excel (1.62x) |
| 2500 | speed-comparison | read | Typed object read | 2 | 0 |  |
| 2500 | speed-comparison | write | AutoFit and mutation | 5 | 0 |  |
| 2500 | speed-comparison | write | Cell writer | 7 | 0 |  |
| 2500 | speed-comparison | write | DataSet table export | 3 | 0 |  |
| 2500 | speed-comparison | write | DataTable table export | 4 | 0 |  |
| 2500 | speed-comparison | write | Formatted report write | 1 | 0 |  |
| 2500 | speed-comparison | write | Formula write/read | 1 | 0 |  |
| 2500 | speed-comparison | write | Other | 2 | 2 | write-powershell-psobject-mixed-direct vs LargeXlsx (1.38x) |
| 2500 | speed-comparison | write | Plain cell export | 2 | 2 | append-plain-rows vs LargeXlsx (1.63x) |
| 2500 | speed-comparison | write | Plain streaming export | 1 | 1 | write-datareader-plain vs Sylvan.Data.Excel (1.03x) |
| 2500 | speed-comparison | write | Plain string export | 0 | 1 | write-blog-2023-20-string-columns vs LargeXlsx (1.03x) |
| 2500 | speed-comparison | write | Shared string write | 4 | 0 |  |
| 2500 | speed-comparison | write | Typed object export | 0 | 3 | write-insertobjects-flat-dictionaries-direct vs LargeXlsx (1.46x) |
| 10000 | focused-package-profile | package | Package size | 1 | 0 |  |
| 25000 | dense-helloworld-comparison | read | Other | 0 | 2 | dense-helloworld-read-range vs Sylvan.Data.Excel (1.21x) |
| 25000 | package-profile | package | Package size | 42 | 12 | write-insertobjects-legacy-dictionaries-direct vs LargeXlsx (1.61x) |
| 25000 | speed-comparison | mutate | AutoFit and mutation | 2 | 0 |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | 4 | 0 |  |
| 25000 | speed-comparison | other | Object projection | 2 | 0 |  |
| 25000 | speed-comparison | other | Range and table read | 4 | 0 |  |
| 25000 | speed-comparison | other | Real-world feature mix | 6 | 0 |  |
| 25000 | speed-comparison | other | Real-world report | 2 | 0 |  |
| 25000 | speed-comparison | other | Report workbook | 4 | 0 |  |
| 25000 | speed-comparison | read | Formula write/read | 1 | 0 |  |
| 25000 | speed-comparison | read | Other | 2 | 1 | shared-string-read vs Sylvan.Data.Excel (1.20x) |
| 25000 | speed-comparison | read | Range and table read | 2 | 5 | read-used-range vs Sylvan.Data.Excel (1.93x) |
| 25000 | speed-comparison | read | Streaming read | 1 | 3 | read-top-range-stream-small-chunks vs Sylvan.Data.Excel (1.24x) |
| 25000 | speed-comparison | read | Typed object read | 1 | 1 | read-objects vs Sylvan.Data.Excel (1.19x) |
| 25000 | speed-comparison | write | AutoFit and mutation | 5 | 0 |  |
| 25000 | speed-comparison | write | Cell writer | 7 | 0 |  |
| 25000 | speed-comparison | write | DataSet table export | 3 | 0 |  |
| 25000 | speed-comparison | write | DataTable table export | 4 | 0 |  |
| 25000 | speed-comparison | write | Formatted report write | 1 | 0 |  |
| 25000 | speed-comparison | write | Formula write/read | 1 | 0 |  |
| 25000 | speed-comparison | write | Other | 2 | 2 | write-powershell-mixed-objects-direct vs LargeXlsx (1.12x) |
| 25000 | speed-comparison | write | Plain cell export | 2 | 2 | append-plain-rows vs LargeXlsx (1.45x) |
| 25000 | speed-comparison | write | Plain streaming export | 1 | 1 | write-datareader-plain vs Sylvan.Data.Excel (1.49x) |
| 25000 | speed-comparison | write | Plain string export | 0 | 1 | write-blog-2023-20-string-columns vs LargeXlsx (1.26x) |
| 25000 | speed-comparison | write | Shared string write | 4 | 0 |  |
| 25000 | speed-comparison | write | Typed object export | 0 | 3 | write-insertobjects-flat-dictionaries-direct vs LargeXlsx (1.34x) |
| 300000 | focused-package-profile | package | Package size | 0 | 1 | write-blog-2023-20-string-columns vs LargeXlsx (1.10x) |

## Rows

| Rows | Artifact | Scenario | Library | Mean | Allocated | Package | Best | Outcome |
| ---: | --- | --- | --- | ---: | ---: | ---: | --- | --- |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | Sylvan.Data.Excel | 4.44 ms | 362.3 KB |  | Sylvan.Data.Excel | 25.7% faster than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | OfficeIMO.Excel | 5.98 ms | 2.4 MB |  | Sylvan.Data.Excel | Loss +34.6% |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | ExcelDataReader | 11.13 ms | 6.7 MB |  | Sylvan.Data.Excel | 86.1% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | MiniExcel | 16.89 ms | 21.0 MB |  | Sylvan.Data.Excel | 182.4% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | OfficeIMO.Excel | 5.67 ms | 2.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | Sylvan.Data.Excel | 6.72 ms | 362.3 KB |  | OfficeIMO.Excel | 18.5% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | ExcelDataReader | 10.82 ms | 6.7 MB |  | OfficeIMO.Excel | 91.0% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | MiniExcel | 21.82 ms | 21.0 MB |  | OfficeIMO.Excel | 285.1% slower than OfficeIMO |
| 2500 | package-profile | append-plain-rows | LargeXlsx | 1.55 ms | 296.4 KB | 63.1 KB | LargeXlsx | 23.8% faster than OfficeIMO |
| 2500 | package-profile | append-plain-rows | OfficeIMO.Excel | 2.04 ms | 1.5 MB | 63.0 KB | LargeXlsx | Loss +31.2% |
| 2500 | package-profile | append-plain-rows | MiniExcel | 5.08 ms | 19.2 MB | 68.1 KB | LargeXlsx | 149.3% slower than OfficeIMO |
| 2500 | package-profile | append-plain-rows | ClosedXML | 17.06 ms | 10.9 MB | 59.8 KB | LargeXlsx | 737.4% slower than OfficeIMO |
| 2500 | package-profile | append-plain-rows | EPPlus | 30.39 ms | 14.0 MB | 56.9 KB | LargeXlsx | 1391.4% slower than OfficeIMO |
| 2500 | package-profile | autofit-existing | OfficeIMO.Excel | 8.95 ms | 1.9 MB | 143.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | autofit-existing | EPPlus | 85.46 ms | 49.5 MB | 115.0 KB | OfficeIMO.Excel | 854.6% slower than OfficeIMO |
| 2500 | package-profile | autofit-existing | ClosedXML | 142.40 ms | 82.6 MB | 121.0 KB | OfficeIMO.Excel | 1490.6% slower than OfficeIMO |
| 2500 | package-profile | large-shared-strings | OfficeIMO.Excel | 2.13 ms | 2.4 MB | 55.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | large-shared-strings | MiniExcel | 4.75 ms | 20.6 MB | 60.7 KB | OfficeIMO.Excel | 122.7% slower than OfficeIMO |
| 2500 | package-profile | large-shared-strings | ClosedXML | 13.59 ms | 11.0 MB | 50.3 KB | OfficeIMO.Excel | 536.6% slower than OfficeIMO |
| 2500 | package-profile | large-shared-strings | EPPlus | 33.29 ms | 12.5 MB | 48.1 KB | OfficeIMO.Excel | 1459.4% slower than OfficeIMO |
| 2500 | package-profile | realworld-autofilter | OfficeIMO.Excel | 3.93 ms | 1.3 MB | 142.4 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-autofilter | ClosedXML | 31.75 ms | 21.7 MB | 120.2 KB | OfficeIMO.Excel | 708.1% slower than OfficeIMO |
| 2500 | package-profile | realworld-autofilter | EPPlus | 45.17 ms | 24.1 MB | 114.2 KB | OfficeIMO.Excel | 1049.9% slower than OfficeIMO |
| 2500 | package-profile | realworld-charts | OfficeIMO.Excel | 5.01 ms | 1.8 MB | 147.6 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-charts | EPPlus | 45.03 ms | 26.5 MB | 117.0 KB | OfficeIMO.Excel | 798.3% slower than OfficeIMO |
| 2500 | package-profile | realworld-conditional-formatting | OfficeIMO.Excel | 3.98 ms | 1.4 MB | 142.7 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-conditional-formatting | ClosedXML | 32.06 ms | 21.8 MB | 120.3 KB | OfficeIMO.Excel | 705.7% slower than OfficeIMO |
| 2500 | package-profile | realworld-conditional-formatting | EPPlus | 41.71 ms | 24.2 MB | 114.3 KB | OfficeIMO.Excel | 948.1% slower than OfficeIMO |
| 2500 | package-profile | realworld-data-validation | OfficeIMO.Excel | 4.66 ms | 1.3 MB | 142.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-data-validation | ClosedXML | 39.65 ms | 21.7 MB | 120.3 KB | OfficeIMO.Excel | 751.6% slower than OfficeIMO |
| 2500 | package-profile | realworld-data-validation | EPPlus | 43.46 ms | 24.1 MB | 114.2 KB | OfficeIMO.Excel | 833.4% slower than OfficeIMO |
| 2500 | package-profile | realworld-freeze-panes | OfficeIMO.Excel | 4.21 ms | 1.3 MB | 142.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-freeze-panes | ClosedXML | 51.06 ms | 21.7 MB | 120.2 KB | OfficeIMO.Excel | 1113.0% slower than OfficeIMO |
| 2500 | package-profile | realworld-freeze-panes | EPPlus | 51.43 ms | 24.2 MB | 114.3 KB | OfficeIMO.Excel | 1121.9% slower than OfficeIMO |
| 2500 | package-profile | realworld-pivot-table | OfficeIMO.Excel | 40.43 ms | 15.3 MB | 200.4 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-pivot-table | EPPlus | 75.85 ms | 28.8 MB | 117.4 KB | OfficeIMO.Excel | 87.6% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-all-in-one | OfficeIMO.Excel | 18.79 ms | 16.1 MB | 206.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-all-in-one | EPPlus | 109.93 ms | 53.3 MB | 121.8 KB | OfficeIMO.Excel | 484.9% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-chart-first | OfficeIMO.Excel | 17.90 ms | 7.3 MB | 206.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-chart-first | EPPlus | 92.37 ms | 53.3 MB | 121.8 KB | OfficeIMO.Excel | 416.1% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-core | OfficeIMO.Excel | 6.40 ms | 1.5 MB | 143.9 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-core | EPPlus | 95.81 ms | 46.2 MB | 115.6 KB | OfficeIMO.Excel | 1398.0% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-core | ClosedXML | 125.19 ms | 68.2 MB | 121.5 KB | OfficeIMO.Excel | 1857.4% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-extra-column | OfficeIMO.Excel | 24.24 ms | 17.2 MB | 219.1 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-extra-column | EPPlus | 110.91 ms | 57.8 MB | 128.4 KB | OfficeIMO.Excel | 357.5% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-no-autofit | OfficeIMO.Excel | 18.99 ms | 16.1 MB | 206.4 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-no-autofit | EPPlus | 55.14 ms | 32.1 MB | 121.8 KB | OfficeIMO.Excel | 190.3% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-post-mutation | OfficeIMO.Excel | 19.99 ms | 16.1 MB | 206.6 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-post-mutation | EPPlus | 89.44 ms | 53.3 MB | 121.9 KB | OfficeIMO.Excel | 347.5% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-shuffled-columns | OfficeIMO.Excel | 18.80 ms | 16.1 MB | 211.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-shuffled-columns | EPPlus | 124.37 ms | 53.3 MB | 124.3 KB | OfficeIMO.Excel | 561.6% slower than OfficeIMO |
| 2500 | package-profile | report-workbook | OfficeIMO.Excel | 24.52 ms | 20.0 MB | 275.6 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | report-workbook | EPPlus | 97.75 ms | 75.7 MB | 161.8 KB | OfficeIMO.Excel | 298.6% slower than OfficeIMO |
| 2500 | package-profile | report-workbook-core | OfficeIMO.Excel | 6.63 ms | 2.6 MB | 187.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | report-workbook-core | EPPlus | 102.38 ms | 70.3 MB | 157.2 KB | OfficeIMO.Excel | 1444.5% slower than OfficeIMO |
| 2500 | package-profile | report-workbook-core | ClosedXML | 108.22 ms | 94.9 MB | 165.1 KB | OfficeIMO.Excel | 1532.6% slower than OfficeIMO |
| 2500 | package-profile | report-workbook-datatable | OfficeIMO.Excel | 22.31 ms | 20.3 MB | 275.6 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | report-workbook-datatable | EPPlus | 108.48 ms | 64.4 MB | 161.8 KB | OfficeIMO.Excel | 386.3% slower than OfficeIMO |
| 2500 | package-profile | report-workbook-datatable-core | OfficeIMO.Excel | 6.43 ms | 2.9 MB | 187.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | report-workbook-datatable-core | EPPlus | 108.72 ms | 59.1 MB | 157.2 KB | OfficeIMO.Excel | 1591.4% slower than OfficeIMO |
| 2500 | package-profile | report-workbook-datatable-core | ClosedXML | 136.50 ms | 80.9 MB | 165.1 KB | OfficeIMO.Excel | 2023.5% slower than OfficeIMO |
| 2500 | package-profile | write-blog-2023-20-string-columns | LargeXlsx | 4.01 ms | 857.6 KB | 237.7 KB | LargeXlsx | 14.9% faster than OfficeIMO |
| 2500 | package-profile | write-blog-2023-20-string-columns | OfficeIMO.Excel | 4.71 ms | 1.6 MB | 216.7 KB | LargeXlsx | Loss +17.5% |
| 2500 | package-profile | write-blog-2023-20-string-columns | MiniExcel | 17.78 ms | 35.1 MB | 235.3 KB | LargeXlsx | 277.6% slower than OfficeIMO |
| 2500 | package-profile | write-blog-2023-20-string-columns | ClosedXML | 93.21 ms | 69.8 MB | 257.2 KB | LargeXlsx | 1879.9% slower than OfficeIMO |
| 2500 | package-profile | write-bulk-report | OfficeIMO.Excel | 4.71 ms | 1.4 MB | 143.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-bulk-report | MiniExcel | 9.16 ms | 26.2 MB | 153.8 KB | OfficeIMO.Excel | 94.4% slower than OfficeIMO |
| 2500 | package-profile | write-bulk-report | EPPlus | 74.44 ms | 46.1 MB | 115.0 KB | OfficeIMO.Excel | 1480.2% slower than OfficeIMO |
| 2500 | package-profile | write-bulk-report | ClosedXML | 74.63 ms | 57.0 MB | 121.0 KB | OfficeIMO.Excel | 1484.2% slower than OfficeIMO |
| 2500 | package-profile | write-cellformula | OfficeIMO.Excel | 2.60 ms | 1.4 MB | 66.6 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellformula | ClosedXML | 22.39 ms | 11.8 MB | 70.6 KB | OfficeIMO.Excel | 761.6% slower than OfficeIMO |
| 2500 | package-profile | write-cellformula | EPPlus | 46.10 ms | 17.7 MB | 62.1 KB | OfficeIMO.Excel | 1674.0% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-empty-strings | OfficeIMO.Excel | 1.98 ms | 1.7 MB | 44.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-empty-strings | ClosedXML | 12.24 ms | 9.7 MB | 44.9 KB | OfficeIMO.Excel | 517.4% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-empty-strings | EPPlus | 31.05 ms | 11.5 MB | 42.0 KB | OfficeIMO.Excel | 1466.6% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-numbers | OfficeIMO.Excel | 2.19 ms | 1.1 MB | 47.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-numbers | ClosedXML | 12.28 ms | 9.0 MB | 45.9 KB | OfficeIMO.Excel | 461.1% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-numbers | EPPlus | 28.87 ms | 12.6 MB | 43.7 KB | OfficeIMO.Excel | 1219.4% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-mixed | OfficeIMO.Excel | 2.58 ms | 1.7 MB | 61.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-object-mixed | ClosedXML | 18.52 ms | 11.6 MB | 59.5 KB | OfficeIMO.Excel | 618.7% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-mixed | EPPlus | 31.22 ms | 15.3 MB | 58.9 KB | OfficeIMO.Excel | 1111.2% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-sparse | OfficeIMO.Excel | 3.03 ms | 1.5 MB | 62.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-object-sparse | ClosedXML | 16.95 ms | 11.0 MB | 52.5 KB | OfficeIMO.Excel | 458.8% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-sparse | EPPlus | 31.93 ms | 14.6 MB | 54.2 KB | OfficeIMO.Excel | 952.6% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 3.09 ms | 1.5 MB | 62.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-object-sparse-batch | ClosedXML | 17.17 ms | 11.0 MB | 52.5 KB | OfficeIMO.Excel | 455.4% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-sparse-batch | EPPlus | 32.60 ms | 14.6 MB | 54.2 KB | OfficeIMO.Excel | 954.4% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-scalars | OfficeIMO.Excel | 1.92 ms | 1.1 MB | 46.3 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-scalars | ClosedXML | 12.18 ms | 8.8 MB | 45.4 KB | OfficeIMO.Excel | 532.8% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-scalars | EPPlus | 28.89 ms | 12.5 MB | 42.4 KB | OfficeIMO.Excel | 1401.0% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings | OfficeIMO.Excel | 3.43 ms | 2.6 MB | 55.1 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-strings | ClosedXML | 17.24 ms | 11.0 MB | 50.3 KB | OfficeIMO.Excel | 402.0% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings | EPPlus | 39.97 ms | 12.5 MB | 48.1 KB | OfficeIMO.Excel | 1064.3% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings-distinct | OfficeIMO.Excel | 2.19 ms | 2.3 MB | 51.8 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-strings-distinct | ClosedXML | 17.77 ms | 12.8 MB | 61.9 KB | OfficeIMO.Excel | 709.9% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings-distinct | EPPlus | 32.79 ms | 13.6 MB | 61.5 KB | OfficeIMO.Excel | 1394.5% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings-repeated | OfficeIMO.Excel | 1.87 ms | 1.5 MB | 40.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-strings-repeated | ClosedXML | 12.02 ms | 9.0 MB | 38.8 KB | OfficeIMO.Excel | 541.9% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings-repeated | EPPlus | 29.12 ms | 11.1 MB | 34.8 KB | OfficeIMO.Excel | 1454.8% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-temporal | OfficeIMO.Excel | 3.23 ms | 1.4 MB | 63.3 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-temporal | ClosedXML | 17.84 ms | 9.5 MB | 54.5 KB | OfficeIMO.Excel | 451.7% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-temporal | EPPlus | 33.32 ms | 14.4 MB | 53.1 KB | OfficeIMO.Excel | 930.4% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 1.42 ms | 447.0 KB | 47.3 KB | LargeXlsx | 10.6% faster than OfficeIMO |
| 2500 | package-profile | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 1.59 ms | 1.1 MB | 48.2 KB | LargeXlsx | Loss +11.8% |
| 2500 | package-profile | write-cellvalues-headerless-rectangle-direct | ClosedXML | 17.41 ms | 10.0 MB | 53.0 KB | LargeXlsx | 992.8% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-headerless-rectangle-direct | EPPlus | 28.68 ms | 12.7 MB | 52.5 KB | LargeXlsx | 1700.4% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | LargeXlsx | 3.15 ms | 758.3 KB | 138.4 KB | LargeXlsx | 27.2% faster than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 4.32 ms | 2.0 MB | 138.0 KB | LargeXlsx | Loss +37.4% |
| 2500 | package-profile | write-cellvalues-rectangle-direct | MiniExcel | 9.36 ms | 22.7 MB | 153.7 KB | LargeXlsx | 116.7% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | ClosedXML | 35.29 ms | 21.7 MB | 120.1 KB | LargeXlsx | 716.6% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | EPPlus | 50.31 ms | 24.1 MB | 114.1 KB | LargeXlsx | 1063.9% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | Sylvan.Data.Excel | 3.04 ms | 758.7 KB | 78.5 KB | Sylvan.Data.Excel | 32.6% faster than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | OfficeIMO.Excel | 4.51 ms | 1.7 MB | 138.0 KB | Sylvan.Data.Excel | Loss +48.4% |
| 2500 | package-profile | write-datareader-plain | LargeXlsx | 4.54 ms | 1.0 MB | 138.4 KB | Sylvan.Data.Excel | Tie vs OfficeIMO |
| 2500 | package-profile | write-datareader-plain | MiniExcel | 10.24 ms | 22.5 MB | 153.6 KB | Sylvan.Data.Excel | 127.0% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | ClosedXML | 29.12 ms | 11.3 MB | 120.1 KB | Sylvan.Data.Excel | 545.7% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | EPPlus | 47.50 ms | 16.3 MB | 114.9 KB | Sylvan.Data.Excel | 953.1% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table | OfficeIMO.Excel | 4.96 ms | 1.7 MB | 138.8 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-datareader-table | MiniExcel | 9.06 ms | 22.5 MB | 153.6 KB | OfficeIMO.Excel | 82.7% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table | ClosedXML | 37.73 ms | 18.6 MB | 120.9 KB | OfficeIMO.Excel | 660.5% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table | EPPlus | 45.48 ms | 16.3 MB | 114.9 KB | OfficeIMO.Excel | 816.8% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table-autofit | OfficeIMO.Excel | 4.81 ms | 1.7 MB | 138.9 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-datareader-table-autofit | MiniExcel | 8.70 ms | 26.0 MB | 153.8 KB | OfficeIMO.Excel | 80.9% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table-autofit | EPPlus | 58.63 ms | 37.4 MB | 115.1 KB | OfficeIMO.Excel | 1119.2% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table-autofit | ClosedXML | 80.64 ms | 57.0 MB | 121.0 KB | OfficeIMO.Excel | 1576.7% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-direct-export | LargeXlsx | 4.35 ms | 1.1 MB | 164.2 KB | LargeXlsx | 3.9% faster than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-direct-export | OfficeIMO.Excel | 4.53 ms | 2.1 MB | 131.1 KB | LargeXlsx | Loss +4.0% |
| 2500 | package-profile | write-dataset-sparse-direct-export | MiniExcel | 12.93 ms | 29.0 MB | 180.5 KB | LargeXlsx | 185.7% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-direct-export | ClosedXML | 53.63 ms | 26.8 MB | 159.4 KB | LargeXlsx | 1084.4% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-direct-export | EPPlus | 57.77 ms | 21.4 MB | 144.5 KB | LargeXlsx | 1175.9% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-tables | OfficeIMO.Excel | 4.57 ms | 2.8 MB | 176.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-dataset-sparse-tables | MiniExcel | 9.87 ms | 29.0 MB | 180.5 KB | OfficeIMO.Excel | 116.1% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-tables | ClosedXML | 53.58 ms | 26.8 MB | 159.4 KB | OfficeIMO.Excel | 1073.5% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-tables | EPPlus | 56.04 ms | 21.4 MB | 144.5 KB | OfficeIMO.Excel | 1127.4% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables | OfficeIMO.Excel | 3.97 ms | 2.0 MB | 138.9 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-dataset-tables | MiniExcel | 9.22 ms | 28.0 MB | 156.4 KB | OfficeIMO.Excel | 132.3% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables | ClosedXML | 35.17 ms | 18.4 MB | 123.4 KB | OfficeIMO.Excel | 786.5% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables | EPPlus | 40.60 ms | 18.3 MB | 116.6 KB | OfficeIMO.Excel | 923.4% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables-autofit | OfficeIMO.Excel | 4.42 ms | 2.0 MB | 139.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-dataset-tables-autofit | MiniExcel | 8.60 ms | 31.1 MB | 156.6 KB | OfficeIMO.Excel | 94.5% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables-autofit | EPPlus | 60.89 ms | 40.5 MB | 116.9 KB | OfficeIMO.Excel | 1277.7% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables-autofit | ClosedXML | 77.37 ms | 55.4 MB | 123.7 KB | OfficeIMO.Excel | 1650.5% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | OfficeIMO.Excel | 4.44 ms | 1.7 MB | 138.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-datatable-direct | LargeXlsx | 4.70 ms | 1.1 MB | 138.4 KB | OfficeIMO.Excel | 6.0% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | MiniExcel | 9.60 ms | 22.5 MB | 153.7 KB | OfficeIMO.Excel | 116.3% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | ClosedXML | 30.59 ms | 11.3 MB | 120.1 KB | OfficeIMO.Excel | 589.5% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | EPPlus | 52.17 ms | 16.3 MB | 114.9 KB | OfficeIMO.Excel | 1076.0% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-table-direct | OfficeIMO.Excel | 3.95 ms | 1.7 MB | 138.8 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-datatable-table-direct | MiniExcel | 12.17 ms | 22.5 MB | 153.7 KB | OfficeIMO.Excel | 208.4% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-table-direct | ClosedXML | 38.19 ms | 18.6 MB | 120.9 KB | OfficeIMO.Excel | 867.5% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-table-direct | EPPlus | 49.41 ms | 16.3 MB | 114.9 KB | OfficeIMO.Excel | 1151.8% slower than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | LargeXlsx | 3.27 ms | 758.3 KB | 138.4 KB | LargeXlsx | 19.3% faster than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 4.06 ms | 1.7 MB | 138.0 KB | LargeXlsx | Loss +23.9% |
| 2500 | package-profile | write-fluent-rowsfrom-direct | MiniExcel | 8.66 ms | 22.7 MB | 153.7 KB | LargeXlsx | 113.4% slower than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | ClosedXML | 29.25 ms | 11.3 MB | 120.1 KB | LargeXlsx | 621.4% slower than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | EPPlus | 49.20 ms | 16.3 MB | 114.9 KB | LargeXlsx | 1113.4% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 5.08 ms | 1.3 MB | 142.4 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 74.15 ms | 37.4 MB | 115.1 KB | OfficeIMO.Excel | 1358.5% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 78.42 ms | 49.7 MB | 120.2 KB | OfficeIMO.Excel | 1442.6% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | OfficeIMO.Excel | 4.15 ms | 1.3 MB | 142.3 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-insertobjects-direct | LargeXlsx | 4.39 ms | 758.3 KB | 138.4 KB | OfficeIMO.Excel | 5.7% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | MiniExcel | 9.07 ms | 22.7 MB | 153.7 KB | OfficeIMO.Excel | 118.4% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | ClosedXML | 37.29 ms | 11.3 MB | 120.1 KB | OfficeIMO.Excel | 797.7% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | EPPlus | 50.65 ms | 16.3 MB | 114.9 KB | OfficeIMO.Excel | 1119.4% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 5.56 ms | 1.5 MB | 138.1 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 62.64 ms | 37.4 MB | 115.1 KB | OfficeIMO.Excel | 1025.9% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 68.76 ms | 49.7 MB | 120.2 KB | OfficeIMO.Excel | 1136.0% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 4.03 ms | 758.3 KB | 138.4 KB | LargeXlsx | 20.5% faster than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 5.07 ms | 1.5 MB | 138.0 KB | LargeXlsx | Loss +25.8% |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | MiniExcel | 10.09 ms | 22.7 MB | 153.7 KB | LargeXlsx | 98.8% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | ClosedXML | 33.31 ms | 11.3 MB | 120.1 KB | LargeXlsx | 556.5% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | EPPlus | 48.27 ms | 16.3 MB | 114.9 KB | LargeXlsx | 851.4% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | LargeXlsx | 3.36 ms | 758.3 KB | 138.4 KB | LargeXlsx | 34.2% faster than OfficeIMO |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | OfficeIMO.Excel | 5.12 ms | 1.7 MB | 142.3 KB | LargeXlsx | Loss +52.0% |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | MiniExcel | 10.49 ms | 22.7 MB | 153.7 KB | LargeXlsx | 105.0% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | ClosedXML | 31.38 ms | 11.3 MB | 120.1 KB | LargeXlsx | 513.4% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | EPPlus | 44.74 ms | 16.3 MB | 114.9 KB | LargeXlsx | 774.5% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.66 ms | 1.3 MB | 142.4 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 56.10 ms | 27.9 MB | 120.2 KB | OfficeIMO.Excel | 1103.9% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 62.00 ms | 26.7 MB | 115.0 KB | OfficeIMO.Excel | 1230.5% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 5.92 ms | 2.3 MB | 183.1 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-powershell-mixed-objects-direct | LargeXlsx | 6.24 ms | 802.5 KB | 182.6 KB | OfficeIMO.Excel | 5.2% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-mixed-objects-direct | MiniExcel | 9.50 ms | 24.6 MB | 194.0 KB | OfficeIMO.Excel | 60.3% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-mixed-objects-direct | ClosedXML | 39.28 ms | 16.6 MB | 161.0 KB | OfficeIMO.Excel | 563.1% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-mixed-objects-direct | EPPlus | 56.36 ms | 19.6 MB | 152.1 KB | OfficeIMO.Excel | 851.2% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-mixed-direct | LargeXlsx | 4.41 ms | 802.5 KB | 182.6 KB | LargeXlsx | 11.6% faster than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 4.99 ms | 1.5 MB | 182.4 KB | LargeXlsx | Loss +13.1% |
| 2500 | package-profile | write-powershell-psobject-mixed-direct | MiniExcel | 9.44 ms | 24.6 MB | 194.0 KB | LargeXlsx | 89.0% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-mixed-direct | ClosedXML | 37.90 ms | 16.6 MB | 161.0 KB | LargeXlsx | 659.1% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-mixed-direct | EPPlus | 54.63 ms | 19.6 MB | 152.1 KB | LargeXlsx | 994.2% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 22.03 ms | 4.4 MB | 651.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-powershell-psobject-wide-direct | LargeXlsx | 25.70 ms | 2.7 MB | 644.6 KB | OfficeIMO.Excel | 16.7% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-wide-direct | MiniExcel | 46.09 ms | 47.3 MB | 674.4 KB | OfficeIMO.Excel | 109.3% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-wide-direct | ClosedXML | 141.65 ms | 50.4 MB | 615.5 KB | OfficeIMO.Excel | 543.1% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-wide-direct | EPPlus | 197.30 ms | 67.5 MB | 548.9 KB | OfficeIMO.Excel | 795.8% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | LargeXlsx | 1.66 ms | 296.4 KB |  | LargeXlsx | 38.5% faster than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | OfficeIMO.Excel | 2.71 ms | 1.5 MB |  | LargeXlsx | Loss +62.7% |
| 2500 | speed-comparison | append-plain-rows | MiniExcel | 4.91 ms | 19.2 MB |  | LargeXlsx | 81.3% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | ClosedXML | 15.74 ms | 10.9 MB |  | LargeXlsx | 481.3% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | EPPlus 4.5.3.3 | 16.50 ms | 0 B |  | LargeXlsx | 509.7% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | EPPlus | 28.16 ms | 14.0 MB |  | LargeXlsx | 940.2% slower than OfficeIMO |
| 2500 | speed-comparison | autofit-existing | OfficeIMO.Excel | 9.58 ms | 1.9 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | autofit-existing | EPPlus 4.5.3.3 | 96.76 ms | 0 B |  | OfficeIMO.Excel | 909.8% slower than OfficeIMO |
| 2500 | speed-comparison | autofit-existing | EPPlus | 153.26 ms | 49.5 MB |  | OfficeIMO.Excel | 1499.4% slower than OfficeIMO |
| 2500 | speed-comparison | autofit-existing | ClosedXML | 246.28 ms | 82.6 MB |  | OfficeIMO.Excel | 2470.1% slower than OfficeIMO |
| 2500 | speed-comparison | build-object-datatable-dictionaries | OfficeIMO.Excel | 1.58 ms | 564.2 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | build-object-datatable-typed | OfficeIMO.Excel | 1.24 ms | 856.9 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-cells | OfficeIMO.Excel | 6.05 ms | 2.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-cells | ClosedXML | 32.07 ms | 16.6 MB |  | OfficeIMO.Excel | 430.3% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-cells | EPPlus | 34.27 ms | 19.7 MB |  | OfficeIMO.Excel | 466.7% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-first-column-from-wide-sheet | OfficeIMO.Excel | 3.93 ms | 526.1 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-first-column-from-wide-sheet | EPPlus | 23.08 ms | 12.8 MB |  | OfficeIMO.Excel | 487.8% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-first-column-from-wide-sheet | ClosedXML | 31.65 ms | 15.1 MB |  | OfficeIMO.Excel | 705.8% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-range | OfficeIMO.Excel | 5.77 ms | 2.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-range | EPPlus | 34.80 ms | 19.7 MB |  | OfficeIMO.Excel | 502.8% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-range | ClosedXML | 36.64 ms | 16.6 MB |  | OfficeIMO.Excel | 534.7% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-top-range | OfficeIMO.Excel | 0.74 ms | 285.5 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-top-range | EPPlus | 26.30 ms | 12.1 MB |  | OfficeIMO.Excel | 3474.9% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-top-range | ClosedXML | 38.74 ms | 15.0 MB |  | OfficeIMO.Excel | 5165.6% slower than OfficeIMO |
| 2500 | speed-comparison | formula-heavy-read | OfficeIMO.Excel | 2.66 ms | 709.4 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | formula-heavy-read | EPPlus 4.5.3.3 | 13.45 ms | 0 B |  | OfficeIMO.Excel | 405.8% slower than OfficeIMO |
| 2500 | speed-comparison | formula-heavy-read | ClosedXML | 16.37 ms | 8.1 MB |  | OfficeIMO.Excel | 515.7% slower than OfficeIMO |
| 2500 | speed-comparison | formula-heavy-read | EPPlus | 22.42 ms | 7.5 MB |  | OfficeIMO.Excel | 743.5% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | OfficeIMO.Excel | 2.05 ms | 2.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | large-shared-strings | MiniExcel | 3.97 ms | 20.6 MB |  | OfficeIMO.Excel | 94.3% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | ClosedXML | 12.83 ms | 11.0 MB |  | OfficeIMO.Excel | 527.2% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | EPPlus 4.5.3.3 | 14.14 ms | 0 B |  | OfficeIMO.Excel | 591.3% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | EPPlus | 25.40 ms | 12.5 MB |  | OfficeIMO.Excel | 1142.1% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | OfficeIMO.Excel | 1.20 ms | 177.4 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | large-sparse-column-read | Sylvan.Data.Excel | 1.47 ms | 316.6 KB |  | OfficeIMO.Excel | 22.0% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | ExcelDataReader | 2.40 ms | 4.0 MB |  | OfficeIMO.Excel | 99.9% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | ClosedXML | 5.24 ms | 4.3 MB |  | OfficeIMO.Excel | 335.4% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | EPPlus 4.5.3.3 | 10.55 ms | 0 B |  | OfficeIMO.Excel | 776.6% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | MiniExcel | 16.38 ms | 45.1 MB |  | OfficeIMO.Excel | 1261.4% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | EPPlus | 43.15 ms | 42.1 MB |  | OfficeIMO.Excel | 3487.1% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | Sylvan.Data.Excel | 1.73 ms | 316.6 KB |  | Sylvan.Data.Excel | 34.4% faster than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | OfficeIMO.Excel | 2.64 ms | 177.5 KB |  | Sylvan.Data.Excel | Loss +52.4% |
| 2500 | speed-comparison | large-sparse-row-read | ExcelDataReader | 2.88 ms | 4.0 MB |  | Sylvan.Data.Excel | 9.1% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | EPPlus 4.5.3.3 | 14.09 ms | 0 B |  | Sylvan.Data.Excel | 434.4% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | MiniExcel | 18.39 ms | 45.1 MB |  | Sylvan.Data.Excel | 597.6% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | EPPlus | 78.22 ms | 42.1 MB |  | Sylvan.Data.Excel | 2867.0% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | ClosedXML | 130.35 ms | 4.3 MB |  | Sylvan.Data.Excel | 4844.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | OfficeIMO.Excel | 4.63 ms | 374.5 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | read-bottom-range | Sylvan.Data.Excel | 5.01 ms | 655.2 KB |  | OfficeIMO.Excel | 8.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | ExcelDataReader | 10.92 ms | 5.9 MB |  | OfficeIMO.Excel | 136.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | MiniExcel | 16.73 ms | 18.2 MB |  | OfficeIMO.Excel | 261.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | EPPlus | 25.85 ms | 12.1 MB |  | OfficeIMO.Excel | 458.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | ClosedXML | 32.55 ms | 15.0 MB |  | OfficeIMO.Excel | 603.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | OfficeIMO.Excel | 3.53 ms | 377.8 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | read-bottom-range-stream | Sylvan.Data.Excel | 4.00 ms | 655.2 KB |  | OfficeIMO.Excel | 13.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | ExcelDataReader | 10.07 ms | 5.9 MB |  | OfficeIMO.Excel | 184.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | MiniExcel | 12.40 ms | 18.2 MB |  | OfficeIMO.Excel | 251.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | EPPlus | 24.50 ms | 12.1 MB |  | OfficeIMO.Excel | 593.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | ClosedXML | 28.47 ms | 15.0 MB |  | OfficeIMO.Excel | 705.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | Sylvan.Data.Excel | 6.39 ms | 2.2 MB |  | Sylvan.Data.Excel | 19.8% faster than OfficeIMO |
| 2500 | speed-comparison | read-datatable | OfficeIMO.Excel | 7.97 ms | 3.5 MB |  | Sylvan.Data.Excel | Loss +24.7% |
| 2500 | speed-comparison | read-datatable | MiniExcel | 13.58 ms | 17.8 MB |  | Sylvan.Data.Excel | 70.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | ExcelDataReader | 13.74 ms | 7.5 MB |  | Sylvan.Data.Excel | 72.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | EPPlus | 34.54 ms | 21.2 MB |  | Sylvan.Data.Excel | 333.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | ClosedXML | 34.64 ms | 17.9 MB |  | Sylvan.Data.Excel | 334.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | EPPlus 4.5.3.3 | 37.35 ms | 0 B |  | Sylvan.Data.Excel | 368.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | Sylvan.Data.Excel | 4.62 ms | 733.5 KB |  | Sylvan.Data.Excel, OfficeIMO.Excel | Tie vs OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | OfficeIMO.Excel | 4.69 ms | 551.0 KB |  | Sylvan.Data.Excel, OfficeIMO.Excel | Win |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | MiniExcel | 10.02 ms | 15.5 MB |  | Sylvan.Data.Excel, OfficeIMO.Excel | 113.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | ExcelDataReader | 10.20 ms | 5.9 MB |  | Sylvan.Data.Excel, OfficeIMO.Excel | 117.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | EPPlus | 25.44 ms | 12.8 MB |  | Sylvan.Data.Excel, OfficeIMO.Excel | 442.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | ClosedXML | 40.08 ms | 15.1 MB |  | Sylvan.Data.Excel, OfficeIMO.Excel | 754.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | OfficeIMO.Excel | 13.24 ms | 2.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | read-objects | Sylvan.Data.Excel | 18.49 ms | 895.3 KB |  | OfficeIMO.Excel | 39.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | EPPlus 4.5.3.3 | 31.66 ms | 0 B |  | OfficeIMO.Excel | 139.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | ExcelDataReader | 40.24 ms | 6.2 MB |  | OfficeIMO.Excel | 204.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | MiniExcel | 44.25 ms | 18.0 MB |  | OfficeIMO.Excel | 234.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | EPPlus | 49.69 ms | 20.9 MB |  | OfficeIMO.Excel | 275.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | ClosedXML | 67.44 ms | 16.5 MB |  | OfficeIMO.Excel | 409.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | OfficeIMO.Excel | 7.59 ms | 2.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | read-objects-stream | Sylvan.Data.Excel | 8.39 ms | 831.0 KB |  | OfficeIMO.Excel | 10.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | ExcelDataReader | 15.25 ms | 6.1 MB |  | OfficeIMO.Excel | 100.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | MiniExcel | 19.76 ms | 18.0 MB |  | OfficeIMO.Excel | 160.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | EPPlus 4.5.3.3 | 28.45 ms | 0 B |  | OfficeIMO.Excel | 274.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | EPPlus | 43.60 ms | 20.8 MB |  | OfficeIMO.Excel | 474.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | ClosedXML | 50.37 ms | 16.5 MB |  | OfficeIMO.Excel | 563.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | OfficeIMO.Excel | 9.55 ms | 2.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | read-range | Sylvan.Data.Excel | 12.29 ms | 655.0 KB |  | OfficeIMO.Excel | 28.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | MiniExcel | 20.15 ms | 18.2 MB |  | OfficeIMO.Excel | 111.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | ExcelDataReader | 21.67 ms | 5.9 MB |  | OfficeIMO.Excel | 126.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | EPPlus 4.5.3.3 | 34.03 ms | 0 B |  | OfficeIMO.Excel | 256.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | EPPlus | 40.39 ms | 19.7 MB |  | OfficeIMO.Excel | 323.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | ClosedXML | 79.44 ms | 16.5 MB |  | OfficeIMO.Excel | 732.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | OfficeIMO.Excel | 6.14 ms | 2.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | read-range-decimal | Sylvan.Data.Excel | 6.26 ms | 750.3 KB |  | OfficeIMO.Excel | 2.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | ExcelDataReader | 13.25 ms | 5.9 MB |  | OfficeIMO.Excel | 115.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | MiniExcel | 14.05 ms | 18.2 MB |  | OfficeIMO.Excel | 128.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | ClosedXML | 30.94 ms | 16.3 MB |  | OfficeIMO.Excel | 404.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | EPPlus | 33.17 ms | 19.7 MB |  | OfficeIMO.Excel | 440.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | Sylvan.Data.Excel | 4.51 ms | 655.2 KB |  | Sylvan.Data.Excel | 24.1% faster than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | OfficeIMO.Excel | 5.94 ms | 2.7 MB |  | Sylvan.Data.Excel | Loss +31.8% |
| 2500 | speed-comparison | read-range-stream | ExcelDataReader | 12.22 ms | 5.9 MB |  | Sylvan.Data.Excel | 105.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | MiniExcel | 13.20 ms | 18.2 MB |  | Sylvan.Data.Excel | 122.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | EPPlus | 30.90 ms | 19.7 MB |  | Sylvan.Data.Excel | 420.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | EPPlus 4.5.3.3 | 31.47 ms | 0 B |  | Sylvan.Data.Excel | 429.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | ClosedXML | 32.20 ms | 16.3 MB |  | Sylvan.Data.Excel | 442.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | Sylvan.Data.Excel | 0.53 ms | 348.5 KB |  | Sylvan.Data.Excel | 12.6% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range | OfficeIMO.Excel | 0.61 ms | 296.0 KB |  | Sylvan.Data.Excel | Loss +14.5% |
| 2500 | speed-comparison | read-top-range | MiniExcel | 0.84 ms | 869.0 KB |  | Sylvan.Data.Excel | 38.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | ExcelDataReader | 4.70 ms | 1.9 MB |  | Sylvan.Data.Excel | 672.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | EPPlus | 27.47 ms | 12.1 MB |  | Sylvan.Data.Excel | 4415.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | EPPlus 4.5.3.3 | 30.54 ms | 0 B |  | Sylvan.Data.Excel | 4918.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | ClosedXML | 34.56 ms | 15.0 MB |  | Sylvan.Data.Excel | 5579.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | Sylvan.Data.Excel | 0.42 ms | 348.5 KB |  | Sylvan.Data.Excel | 21.2% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | OfficeIMO.Excel | 0.53 ms | 299.4 KB |  | Sylvan.Data.Excel | Loss +27.0% |
| 2500 | speed-comparison | read-top-range-stream | MiniExcel | 0.74 ms | 869.0 KB |  | Sylvan.Data.Excel | 39.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | ExcelDataReader | 4.63 ms | 1.9 MB |  | Sylvan.Data.Excel | 770.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | EPPlus | 24.44 ms | 12.1 MB |  | Sylvan.Data.Excel | 4498.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | EPPlus 4.5.3.3 | 28.84 ms | 0 B |  | Sylvan.Data.Excel | 5326.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | ClosedXML | 29.86 ms | 15.0 MB |  | Sylvan.Data.Excel | 5518.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | Sylvan.Data.Excel | 0.42 ms | 348.5 KB |  | Sylvan.Data.Excel | 38.1% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | OfficeIMO.Excel | 0.67 ms | 300.1 KB |  | Sylvan.Data.Excel | Loss +61.5% |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | MiniExcel | 0.87 ms | 869.0 KB |  | Sylvan.Data.Excel | 29.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | ExcelDataReader | 5.22 ms | 1.9 MB |  | Sylvan.Data.Excel | 674.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | EPPlus | 26.20 ms | 12.1 MB |  | Sylvan.Data.Excel | 3782.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | ClosedXML | 30.94 ms | 15.0 MB |  | Sylvan.Data.Excel | 4484.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-used-range | Sylvan.Data.Excel | 4.67 ms | 655.2 KB |  | Sylvan.Data.Excel | 65.3% faster than OfficeIMO |
| 2500 | speed-comparison | read-used-range | ExcelDataReader | 10.48 ms | 5.9 MB |  | Sylvan.Data.Excel | 22.1% faster than OfficeIMO |
| 2500 | speed-comparison | read-used-range | MiniExcel | 12.41 ms | 18.2 MB |  | Sylvan.Data.Excel | 7.8% faster than OfficeIMO |
| 2500 | speed-comparison | read-used-range | OfficeIMO.Excel | 13.46 ms | 3.4 MB |  | Sylvan.Data.Excel | Loss +188.4% |
| 2500 | speed-comparison | read-used-range | EPPlus | 37.20 ms | 19.7 MB |  | Sylvan.Data.Excel | 176.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-used-range | ClosedXML | 52.00 ms | 16.4 MB |  | Sylvan.Data.Excel | 286.2% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-autofilter | OfficeIMO.Excel | 3.76 ms | 1.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-autofilter | ClosedXML | 30.06 ms | 21.7 MB |  | OfficeIMO.Excel | 698.7% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-autofilter | EPPlus 4.5.3.3 | 31.90 ms | 0 B |  | OfficeIMO.Excel | 747.5% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-autofilter | EPPlus | 39.23 ms | 24.1 MB |  | OfficeIMO.Excel | 942.2% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-charts | OfficeIMO.Excel | 5.08 ms | 1.8 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-charts | EPPlus 4.5.3.3 | 34.43 ms | 0 B |  | OfficeIMO.Excel | 578.4% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-charts | EPPlus | 44.67 ms | 26.5 MB |  | OfficeIMO.Excel | 780.2% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-conditional-formatting | OfficeIMO.Excel | 3.99 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-conditional-formatting | ClosedXML | 31.01 ms | 21.8 MB |  | OfficeIMO.Excel | 676.8% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-conditional-formatting | EPPlus 4.5.3.3 | 35.13 ms | 0 B |  | OfficeIMO.Excel | 779.9% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-conditional-formatting | EPPlus | 40.54 ms | 24.2 MB |  | OfficeIMO.Excel | 915.3% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-data-validation | OfficeIMO.Excel | 3.79 ms | 1.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-data-validation | EPPlus 4.5.3.3 | 31.43 ms | 0 B |  | OfficeIMO.Excel | 729.9% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-data-validation | ClosedXML | 31.48 ms | 21.7 MB |  | OfficeIMO.Excel | 731.2% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-data-validation | EPPlus | 47.86 ms | 24.1 MB |  | OfficeIMO.Excel | 1163.7% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-freeze-panes | OfficeIMO.Excel | 3.54 ms | 1.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-freeze-panes | ClosedXML | 30.17 ms | 21.7 MB |  | OfficeIMO.Excel | 752.0% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-freeze-panes | EPPlus 4.5.3.3 | 32.39 ms | 0 B |  | OfficeIMO.Excel | 814.5% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-freeze-panes | EPPlus | 46.01 ms | 24.2 MB |  | OfficeIMO.Excel | 1199.3% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-pivot-table | OfficeIMO.Excel | 15.54 ms | 15.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-pivot-table | EPPlus 4.5.3.3 | 35.76 ms | 0 B |  | OfficeIMO.Excel | 130.2% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-pivot-table | EPPlus | 50.67 ms | 28.8 MB |  | OfficeIMO.Excel | 226.1% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-all-in-one | OfficeIMO.Excel | 18.80 ms | 16.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-all-in-one | EPPlus | 78.40 ms | 53.3 MB |  | OfficeIMO.Excel | 317.0% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-all-in-one | EPPlus 4.5.3.3 | 81.58 ms | 0 B |  | OfficeIMO.Excel | 333.9% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-chart-first | OfficeIMO.Excel | 12.31 ms | 7.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-chart-first | EPPlus 4.5.3.3 | 71.50 ms | 0 B |  | OfficeIMO.Excel | 480.8% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-chart-first | EPPlus | 72.95 ms | 53.3 MB |  | OfficeIMO.Excel | 492.6% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-core | OfficeIMO.Excel | 4.45 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-core | EPPlus | 66.04 ms | 46.2 MB |  | OfficeIMO.Excel | 1383.1% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-core | EPPlus 4.5.3.3 | 67.08 ms | 0 B |  | OfficeIMO.Excel | 1406.4% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-core | ClosedXML | 87.76 ms | 68.2 MB |  | OfficeIMO.Excel | 1870.9% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-extra-column | OfficeIMO.Excel | 18.17 ms | 17.2 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-extra-column | EPPlus | 75.44 ms | 57.8 MB |  | OfficeIMO.Excel | 315.2% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-extra-column | EPPlus 4.5.3.3 | 83.91 ms | 0 B |  | OfficeIMO.Excel | 361.7% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-no-autofit | OfficeIMO.Excel | 15.80 ms | 16.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-no-autofit | EPPlus 4.5.3.3 | 41.12 ms | 0 B |  | OfficeIMO.Excel | 160.2% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-no-autofit | EPPlus | 52.21 ms | 32.1 MB |  | OfficeIMO.Excel | 230.3% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-post-mutation | OfficeIMO.Excel | 16.10 ms | 16.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-post-mutation | EPPlus 4.5.3.3 | 69.99 ms | 0 B |  | OfficeIMO.Excel | 334.8% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-post-mutation | EPPlus | 77.29 ms | 53.3 MB |  | OfficeIMO.Excel | 380.1% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-shuffled-columns | OfficeIMO.Excel | 22.37 ms | 16.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-shuffled-columns | EPPlus 4.5.3.3 | 70.10 ms | 0 B |  | OfficeIMO.Excel | 213.4% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-shuffled-columns | EPPlus | 70.13 ms | 53.3 MB |  | OfficeIMO.Excel | 213.5% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook | OfficeIMO.Excel | 31.04 ms | 20.0 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | report-workbook | EPPlus 4.5.3.3 | 95.11 ms | 0 B |  | OfficeIMO.Excel | 206.5% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook | EPPlus | 118.61 ms | 75.7 MB |  | OfficeIMO.Excel | 282.2% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-core | OfficeIMO.Excel | 7.25 ms | 2.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | report-workbook-core | EPPlus 4.5.3.3 | 82.84 ms | 0 B |  | OfficeIMO.Excel | 1041.8% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-core | EPPlus | 117.98 ms | 70.3 MB |  | OfficeIMO.Excel | 1526.1% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-core | ClosedXML | 119.32 ms | 94.9 MB |  | OfficeIMO.Excel | 1544.7% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-datatable | OfficeIMO.Excel | 27.85 ms | 20.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | report-workbook-datatable | EPPlus 4.5.3.3 | 84.67 ms | 0 B |  | OfficeIMO.Excel | 204.0% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-datatable | EPPlus | 109.46 ms | 64.4 MB |  | OfficeIMO.Excel | 293.1% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-datatable-core | OfficeIMO.Excel | 6.52 ms | 2.9 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | report-workbook-datatable-core | EPPlus 4.5.3.3 | 84.15 ms | 0 B |  | OfficeIMO.Excel | 1191.0% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-datatable-core | EPPlus | 100.44 ms | 59.1 MB |  | OfficeIMO.Excel | 1440.9% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-datatable-core | ClosedXML | 103.43 ms | 80.9 MB |  | OfficeIMO.Excel | 1486.9% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | Sylvan.Data.Excel | 2.35 ms | 518.6 KB |  | Sylvan.Data.Excel | 22.2% faster than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | OfficeIMO.Excel | 3.02 ms | 1.0 MB |  | Sylvan.Data.Excel | Loss +28.5% |
| 2500 | speed-comparison | shared-string-read | ExcelDataReader | 5.47 ms | 2.6 MB |  | Sylvan.Data.Excel | 81.1% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | MiniExcel | 8.74 ms | 7.4 MB |  | Sylvan.Data.Excel | 189.2% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | EPPlus 4.5.3.3 | 15.48 ms | 0 B |  | Sylvan.Data.Excel | 412.5% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | ClosedXML | 15.49 ms | 9.3 MB |  | Sylvan.Data.Excel | 412.8% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | EPPlus | 20.58 ms | 10.1 MB |  | Sylvan.Data.Excel | 581.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-blog-2023-20-string-columns | LargeXlsx | 4.31 ms | 857.6 KB |  | LargeXlsx | 3.1% faster than OfficeIMO |
| 2500 | speed-comparison | write-blog-2023-20-string-columns | OfficeIMO.Excel | 4.45 ms | 1.6 MB |  | LargeXlsx | Loss +3.2% |
| 2500 | speed-comparison | write-blog-2023-20-string-columns | MiniExcel | 16.45 ms | 35.1 MB |  | LargeXlsx | 269.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-blog-2023-20-string-columns | ClosedXML | 88.92 ms | 69.8 MB |  | LargeXlsx | 1897.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | OfficeIMO.Excel | 5.80 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-bulk-report | MiniExcel | 13.37 ms | 26.2 MB |  | OfficeIMO.Excel | 130.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | EPPlus 4.5.3.3 | 104.37 ms | 0 B |  | OfficeIMO.Excel | 1698.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | EPPlus | 127.91 ms | 48.0 MB |  | OfficeIMO.Excel | 2103.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | ClosedXML | 219.04 ms | 57.0 MB |  | OfficeIMO.Excel | 3674.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellformula | OfficeIMO.Excel | 4.36 ms | 1.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellformula | EPPlus 4.5.3.3 | 19.69 ms | 0 B |  | OfficeIMO.Excel | 352.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellformula | ClosedXML | 26.18 ms | 11.8 MB |  | OfficeIMO.Excel | 501.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellformula | EPPlus | 51.94 ms | 17.7 MB |  | OfficeIMO.Excel | 1092.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-empty-strings | OfficeIMO.Excel | 2.63 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-empty-strings | ClosedXML | 15.57 ms | 9.7 MB |  | OfficeIMO.Excel | 492.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-empty-strings | EPPlus | 22.53 ms | 11.5 MB |  | OfficeIMO.Excel | 756.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-numbers | OfficeIMO.Excel | 3.20 ms | 1.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-numbers | ClosedXML | 12.05 ms | 9.0 MB |  | OfficeIMO.Excel | 276.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-numbers | EPPlus 4.5.3.3 | 17.22 ms | 0 B |  | OfficeIMO.Excel | 437.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-numbers | EPPlus | 25.62 ms | 12.6 MB |  | OfficeIMO.Excel | 700.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-mixed | OfficeIMO.Excel | 3.41 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-object-mixed | ClosedXML | 17.03 ms | 11.6 MB |  | OfficeIMO.Excel | 399.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-mixed | EPPlus 4.5.3.3 | 17.30 ms | 0 B |  | OfficeIMO.Excel | 407.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-mixed | EPPlus | 27.64 ms | 15.3 MB |  | OfficeIMO.Excel | 711.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-sparse | OfficeIMO.Excel | 3.72 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-object-sparse | ClosedXML | 15.37 ms | 11.0 MB |  | OfficeIMO.Excel | 312.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-sparse | EPPlus | 28.52 ms | 14.6 MB |  | OfficeIMO.Excel | 665.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 3.08 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-object-sparse-batch | ClosedXML | 15.05 ms | 11.0 MB |  | OfficeIMO.Excel | 389.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-sparse-batch | EPPlus | 26.93 ms | 14.6 MB |  | OfficeIMO.Excel | 775.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-scalars | OfficeIMO.Excel | 2.57 ms | 1.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-scalars | ClosedXML | 12.12 ms | 8.8 MB |  | OfficeIMO.Excel | 371.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-scalars | EPPlus 4.5.3.3 | 15.86 ms | 0 B |  | OfficeIMO.Excel | 517.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-scalars | EPPlus | 25.37 ms | 12.5 MB |  | OfficeIMO.Excel | 886.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings | OfficeIMO.Excel | 2.85 ms | 2.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-strings | ClosedXML | 11.80 ms | 11.0 MB |  | OfficeIMO.Excel | 313.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings | EPPlus 4.5.3.3 | 17.38 ms | 0 B |  | OfficeIMO.Excel | 509.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings | EPPlus | 23.73 ms | 12.5 MB |  | OfficeIMO.Excel | 732.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings-distinct | OfficeIMO.Excel | 2.50 ms | 2.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-strings-distinct | ClosedXML | 14.40 ms | 12.8 MB |  | OfficeIMO.Excel | 475.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings-distinct | EPPlus | 23.43 ms | 13.6 MB |  | OfficeIMO.Excel | 836.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings-repeated | OfficeIMO.Excel | 2.06 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-strings-repeated | ClosedXML | 12.47 ms | 9.0 MB |  | OfficeIMO.Excel | 504.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings-repeated | EPPlus | 23.44 ms | 11.1 MB |  | OfficeIMO.Excel | 1035.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-temporal | OfficeIMO.Excel | 3.64 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-temporal | EPPlus 4.5.3.3 | 17.59 ms | 0 B |  | OfficeIMO.Excel | 383.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-temporal | ClosedXML | 17.89 ms | 9.5 MB |  | OfficeIMO.Excel | 391.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-temporal | EPPlus | 29.04 ms | 14.4 MB |  | OfficeIMO.Excel | 697.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 1.43 ms | 447.0 KB |  | LargeXlsx | 16.7% faster than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 1.72 ms | 1.1 MB |  | LargeXlsx | Loss +20.0% |
| 2500 | speed-comparison | write-cellvalues-headerless-rectangle-direct | ClosedXML | 17.27 ms | 10.0 MB |  | LargeXlsx | 905.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-headerless-rectangle-direct | EPPlus | 24.65 ms | 12.7 MB |  | LargeXlsx | 1335.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 4.25 ms | 2.0 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | LargeXlsx | 4.67 ms | 758.3 KB |  | OfficeIMO.Excel | 9.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | MiniExcel | 9.25 ms | 22.7 MB |  | OfficeIMO.Excel | 117.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | ClosedXML | 33.28 ms | 21.7 MB |  | OfficeIMO.Excel | 683.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 35.40 ms | 0 B |  | OfficeIMO.Excel | 733.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus | 43.02 ms | 24.1 MB |  | OfficeIMO.Excel | 912.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-sparse-rectangle-direct | OfficeIMO.Excel | 2.46 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalues-sparse-rectangle-direct | ClosedXML | 15.15 ms | 11.0 MB |  | OfficeIMO.Excel | 517.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-sparse-rectangle-direct | EPPlus | 26.15 ms | 14.6 MB |  | OfficeIMO.Excel | 964.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | Sylvan.Data.Excel | 4.70 ms | 758.6 KB |  | Sylvan.Data.Excel | 2.8% faster than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | OfficeIMO.Excel | 4.84 ms | 1.7 MB |  | Sylvan.Data.Excel | Loss +2.9% |
| 2500 | speed-comparison | write-datareader-plain | LargeXlsx | 8.72 ms | 1.0 MB |  | Sylvan.Data.Excel | 80.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | MiniExcel | 9.95 ms | 22.5 MB |  | Sylvan.Data.Excel | 105.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | ClosedXML | 28.01 ms | 11.3 MB |  | Sylvan.Data.Excel | 478.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | EPPlus 4.5.3.3 | 32.76 ms | 0 B |  | Sylvan.Data.Excel | 576.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | EPPlus | 37.06 ms | 16.3 MB |  | Sylvan.Data.Excel | 665.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | OfficeIMO.Excel | 4.67 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datareader-table | MiniExcel | 8.66 ms | 22.5 MB |  | OfficeIMO.Excel | 85.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | ClosedXML | 38.22 ms | 18.6 MB |  | OfficeIMO.Excel | 718.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | EPPlus 4.5.3.3 | 39.58 ms | 0 B |  | OfficeIMO.Excel | 747.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | EPPlus | 44.60 ms | 16.3 MB |  | OfficeIMO.Excel | 854.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table-autofit | OfficeIMO.Excel | 7.58 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datareader-table-autofit | MiniExcel | 9.51 ms | 26.0 MB |  | OfficeIMO.Excel | 25.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table-autofit | EPPlus | 59.00 ms | 37.4 MB |  | OfficeIMO.Excel | 678.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table-autofit | EPPlus 4.5.3.3 | 78.02 ms | 0 B |  | OfficeIMO.Excel | 929.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table-autofit | ClosedXML | 165.11 ms | 57.0 MB |  | OfficeIMO.Excel | 2079.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-headerless-tables | OfficeIMO.Excel | 4.95 ms | 2.0 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-headerless-tables | MiniExcel | 14.88 ms | 28.5 MB |  | OfficeIMO.Excel | 200.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-headerless-tables | ClosedXML | 66.79 ms | 18.5 MB |  | OfficeIMO.Excel | 1248.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-headerless-tables | EPPlus | 101.58 ms | 18.0 MB |  | OfficeIMO.Excel | 1950.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | OfficeIMO.Excel | 7.33 ms | 2.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | LargeXlsx | 13.53 ms | 1.1 MB |  | OfficeIMO.Excel | 84.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | MiniExcel | 18.08 ms | 29.7 MB |  | OfficeIMO.Excel | 146.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | EPPlus | 70.09 ms | 21.8 MB |  | OfficeIMO.Excel | 856.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | ClosedXML | 89.91 ms | 26.8 MB |  | OfficeIMO.Excel | 1127.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-tables | OfficeIMO.Excel | 7.34 ms | 2.8 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-sparse-tables | MiniExcel | 19.25 ms | 29.8 MB |  | OfficeIMO.Excel | 162.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-tables | ClosedXML | 93.18 ms | 26.8 MB |  | OfficeIMO.Excel | 1170.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-tables | EPPlus | 105.75 ms | 22.1 MB |  | OfficeIMO.Excel | 1341.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | OfficeIMO.Excel | 5.55 ms | 2.0 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-tables | MiniExcel | 13.89 ms | 28.0 MB |  | OfficeIMO.Excel | 150.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | EPPlus 4.5.3.3 | 54.07 ms | 0 B |  | OfficeIMO.Excel | 875.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | ClosedXML | 63.06 ms | 18.4 MB |  | OfficeIMO.Excel | 1037.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | EPPlus | 70.75 ms | 19.0 MB |  | OfficeIMO.Excel | 1175.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables-autofit | OfficeIMO.Excel | 5.97 ms | 2.0 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-tables-autofit | MiniExcel | 15.95 ms | 31.1 MB |  | OfficeIMO.Excel | 167.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables-autofit | EPPlus | 151.95 ms | 42.4 MB |  | OfficeIMO.Excel | 2446.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables-autofit | ClosedXML | 182.55 ms | 55.4 MB |  | OfficeIMO.Excel | 2958.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | OfficeIMO.Excel | 7.70 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datatable-direct | LargeXlsx | 7.97 ms | 1.1 MB |  | OfficeIMO.Excel | 3.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | MiniExcel | 10.21 ms | 22.5 MB |  | OfficeIMO.Excel | 32.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | ClosedXML | 37.21 ms | 11.3 MB |  | OfficeIMO.Excel | 383.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | EPPlus | 43.91 ms | 16.3 MB |  | OfficeIMO.Excel | 470.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | EPPlus 4.5.3.3 | 45.86 ms | 0 B |  | OfficeIMO.Excel | 495.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-object-table-direct | OfficeIMO.Excel | 5.28 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datatable-object-table-direct | MiniExcel | 11.52 ms | 22.3 MB |  | OfficeIMO.Excel | 118.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-object-table-direct | ClosedXML | 41.38 ms | 18.3 MB |  | OfficeIMO.Excel | 683.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-object-table-direct | EPPlus | 46.89 ms | 16.0 MB |  | OfficeIMO.Excel | 787.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | OfficeIMO.Excel | 4.80 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datatable-table-direct | MiniExcel | 10.10 ms | 22.5 MB |  | OfficeIMO.Excel | 110.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | EPPlus 4.5.3.3 | 42.21 ms | 0 B |  | OfficeIMO.Excel | 780.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | EPPlus | 48.37 ms | 16.3 MB |  | OfficeIMO.Excel | 908.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | ClosedXML | 49.33 ms | 18.6 MB |  | OfficeIMO.Excel | 928.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-dictionary-objects-table-direct | OfficeIMO.Excel | 7.13 ms | 2.0 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | LargeXlsx | 3.38 ms | 758.3 KB |  | LargeXlsx | 18.4% faster than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 4.14 ms | 1.7 MB |  | LargeXlsx | Loss +22.5% |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | MiniExcel | 8.36 ms | 22.7 MB |  | LargeXlsx | 101.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | ClosedXML | 30.22 ms | 11.3 MB |  | LargeXlsx | 629.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 37.16 ms | 0 B |  | LargeXlsx | 797.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus | 39.29 ms | 16.3 MB |  | LargeXlsx | 848.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 15.06 ms | 1.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 85.75 ms | 37.4 MB |  | OfficeIMO.Excel | 469.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 118.60 ms | 49.7 MB |  | OfficeIMO.Excel | 687.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | LargeXlsx | 4.09 ms | 758.3 KB |  | LargeXlsx | 30.7% faster than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | OfficeIMO.Excel | 5.91 ms | 1.3 MB |  | LargeXlsx | Loss +44.3% |
| 2500 | speed-comparison | write-insertobjects-direct | MiniExcel | 11.66 ms | 22.7 MB |  | LargeXlsx | 97.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | EPPlus 4.5.3.3 | 35.44 ms | 0 B |  | LargeXlsx | 499.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | ClosedXML | 59.53 ms | 11.3 MB |  | LargeXlsx | 907.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | EPPlus | 64.43 ms | 16.3 MB |  | LargeXlsx | 990.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 7.52 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 70.60 ms | 37.4 MB |  | OfficeIMO.Excel | 838.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 95.09 ms | 49.7 MB |  | OfficeIMO.Excel | 1164.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 4.94 ms | 758.3 KB |  | LargeXlsx | 31.7% faster than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 7.23 ms | 1.5 MB |  | LargeXlsx | Loss +46.3% |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | MiniExcel | 16.89 ms | 22.7 MB |  | LargeXlsx | 133.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | ClosedXML | 40.65 ms | 11.3 MB |  | LargeXlsx | 462.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | EPPlus | 58.71 ms | 16.3 MB |  | LargeXlsx | 712.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 7.62 ms | 1.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 71.20 ms | 27.9 MB |  | OfficeIMO.Excel | 834.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 78.97 ms | 26.7 MB |  | OfficeIMO.Excel | 936.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | LargeXlsx | 4.54 ms | 802.5 KB |  | LargeXlsx | 24.9% faster than OfficeIMO |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 6.04 ms | 2.3 MB |  | LargeXlsx | Loss +33.1% |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | MiniExcel | 9.72 ms | 24.6 MB |  | LargeXlsx | 61.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | ClosedXML | 49.04 ms | 16.6 MB |  | LargeXlsx | 712.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | EPPlus | 59.34 ms | 19.6 MB |  | LargeXlsx | 882.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-mixed-direct | LargeXlsx | 4.85 ms | 802.5 KB |  | LargeXlsx | 27.6% faster than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 6.71 ms | 1.5 MB |  | LargeXlsx | Loss +38.2% |
| 2500 | speed-comparison | write-powershell-psobject-mixed-direct | MiniExcel | 14.01 ms | 24.6 MB |  | LargeXlsx | 108.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-mixed-direct | ClosedXML | 66.92 ms | 16.6 MB |  | LargeXlsx | 897.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-mixed-direct | EPPlus | 71.05 ms | 19.6 MB |  | LargeXlsx | 958.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 22.25 ms | 4.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-powershell-psobject-wide-direct | LargeXlsx | 41.84 ms | 2.7 MB |  | OfficeIMO.Excel | 88.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-wide-direct | MiniExcel | 55.69 ms | 47.3 MB |  | OfficeIMO.Excel | 150.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-wide-direct | ClosedXML | 128.12 ms | 50.4 MB |  | OfficeIMO.Excel | 475.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-wide-direct | EPPlus | 176.60 ms | 67.5 MB |  | OfficeIMO.Excel | 693.6% slower than OfficeIMO |
| 10000 | focused-package-profile | write-blog-2023-20-string-columns | OfficeIMO.Excel | 33.01 ms | 7.6 MB | 880.4 KB | OfficeIMO.Excel | Win |
| 10000 | focused-package-profile | write-blog-2023-20-string-columns | LargeXlsx | 88.91 ms | 3.1 MB | 970.2 KB | OfficeIMO.Excel | 2.69x vs best |
| 10000 | focused-package-profile | write-blog-2023-20-string-columns | MiniExcel | 142.47 ms | 96.2 MB | 957.6 KB | OfficeIMO.Excel | 4.32x vs best |
| 10000 | focused-package-profile | write-blog-2023-20-string-columns | ClosedXML | 725.30 ms | 280.2 MB | 1,015.4 KB | OfficeIMO.Excel | 21.97x vs best |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | Sylvan.Data.Excel | 61.06 ms | 394.1 KB |  | Sylvan.Data.Excel | 17.5% faster than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | OfficeIMO.Excel | 73.99 ms | 23.1 MB |  | Sylvan.Data.Excel | Loss +21.2% |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | ExcelDataReader | 162.93 ms | 67.9 MB |  | Sylvan.Data.Excel | 120.2% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | MiniExcel | 217.43 ms | 210.3 MB |  | Sylvan.Data.Excel | 193.9% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | Sylvan.Data.Excel | 66.62 ms | 394.1 KB |  | Sylvan.Data.Excel | 3.4% faster than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | OfficeIMO.Excel | 68.98 ms | 23.8 MB |  | Sylvan.Data.Excel | Loss +3.5% |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | ExcelDataReader | 194.35 ms | 67.9 MB |  | Sylvan.Data.Excel | 181.7% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | MiniExcel | 244.62 ms | 210.3 MB |  | Sylvan.Data.Excel | 254.6% slower than OfficeIMO |
| 25000 | package-profile | append-plain-rows | LargeXlsx | 14.33 ms | 2.7 MB | 605.0 KB | LargeXlsx | 29.9% faster than OfficeIMO |
| 25000 | package-profile | append-plain-rows | OfficeIMO.Excel | 20.45 ms | 10.6 MB | 610.4 KB | LargeXlsx | Loss +42.7% |
| 25000 | package-profile | append-plain-rows | MiniExcel | 36.40 ms | 56.9 MB | 642.3 KB | LargeXlsx | 78.0% slower than OfficeIMO |
| 25000 | package-profile | append-plain-rows | ClosedXML | 163.87 ms | 101.8 MB | 540.6 KB | LargeXlsx | 701.3% slower than OfficeIMO |
| 25000 | package-profile | append-plain-rows | EPPlus | 244.50 ms | 98.0 MB | 525.6 KB | LargeXlsx | 1095.6% slower than OfficeIMO |
| 25000 | package-profile | autofit-existing | OfficeIMO.Excel | 92.81 ms | 15.3 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | autofit-existing | EPPlus | 594.57 ms | 245.1 MB | 1.1 MB | OfficeIMO.Excel | 540.6% slower than OfficeIMO |
| 25000 | package-profile | autofit-existing | ClosedXML | 1.78 s | 810.3 MB | 1.1 MB | OfficeIMO.Excel | 1816.9% slower than OfficeIMO |
| 25000 | package-profile | large-shared-strings | OfficeIMO.Excel | 20.06 ms | 15.4 MB | 529.7 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | large-shared-strings | MiniExcel | 39.29 ms | 72.0 MB | 581.0 KB | OfficeIMO.Excel | 95.9% slower than OfficeIMO |
| 25000 | package-profile | large-shared-strings | ClosedXML | 160.95 ms | 101.8 MB | 460.1 KB | OfficeIMO.Excel | 702.3% slower than OfficeIMO |
| 25000 | package-profile | large-shared-strings | EPPlus | 283.73 ms | 82.4 MB | 444.7 KB | OfficeIMO.Excel | 1314.4% slower than OfficeIMO |
| 25000 | package-profile | realworld-autofilter | OfficeIMO.Excel | 43.84 ms | 11.2 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-autofilter | ClosedXML | 398.52 ms | 205.7 MB | 1.1 MB | OfficeIMO.Excel | 809.0% slower than OfficeIMO |
| 25000 | package-profile | realworld-autofilter | EPPlus | 477.44 ms | 206.9 MB | 1.1 MB | OfficeIMO.Excel | 989.0% slower than OfficeIMO |
| 25000 | package-profile | realworld-charts | OfficeIMO.Excel | 45.87 ms | 12.3 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-charts | EPPlus | 513.82 ms | 209.9 MB | 1.1 MB | OfficeIMO.Excel | 1020.1% slower than OfficeIMO |
| 25000 | package-profile | realworld-conditional-formatting | OfficeIMO.Excel | 42.72 ms | 11.3 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-conditional-formatting | ClosedXML | 401.24 ms | 205.8 MB | 1.1 MB | OfficeIMO.Excel | 839.3% slower than OfficeIMO |
| 25000 | package-profile | realworld-conditional-formatting | EPPlus | 522.13 ms | 206.9 MB | 1.1 MB | OfficeIMO.Excel | 1122.3% slower than OfficeIMO |
| 25000 | package-profile | realworld-data-validation | OfficeIMO.Excel | 42.86 ms | 11.2 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-data-validation | ClosedXML | 383.78 ms | 205.7 MB | 1.1 MB | OfficeIMO.Excel | 795.4% slower than OfficeIMO |
| 25000 | package-profile | realworld-data-validation | EPPlus | 527.65 ms | 206.9 MB | 1.1 MB | OfficeIMO.Excel | 1131.1% slower than OfficeIMO |
| 25000 | package-profile | realworld-freeze-panes | OfficeIMO.Excel | 42.53 ms | 11.2 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-freeze-panes | ClosedXML | 395.05 ms | 205.7 MB | 1.1 MB | OfficeIMO.Excel | 828.9% slower than OfficeIMO |
| 25000 | package-profile | realworld-freeze-panes | EPPlus | 463.69 ms | 206.9 MB | 1.1 MB | OfficeIMO.Excel | 990.3% slower than OfficeIMO |
| 25000 | package-profile | realworld-pivot-table | OfficeIMO.Excel | 411.32 ms | 140.3 MB | 1.9 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-pivot-table | EPPlus | 602.49 ms | 225.4 MB | 1.1 MB | OfficeIMO.Excel | 46.5% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-all-in-one | OfficeIMO.Excel | 307.01 ms | 141.7 MB | 1.9 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-all-in-one | EPPlus | 527.21 ms | 270.6 MB | 1.1 MB | OfficeIMO.Excel | 71.7% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-chart-first | OfficeIMO.Excel | 119.04 ms | 54.0 MB | 1.9 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-chart-first | EPPlus | 535.98 ms | 270.6 MB | 1.1 MB | OfficeIMO.Excel | 350.3% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-core | OfficeIMO.Excel | 46.76 ms | 11.4 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-core | EPPlus | 516.28 ms | 249.1 MB | 1.1 MB | OfficeIMO.Excel | 1004.1% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-core | ClosedXML | 1.09 s | 664.2 MB | 1.1 MB | OfficeIMO.Excel | 2227.1% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-extra-column | OfficeIMO.Excel | 325.78 ms | 153.2 MB | 2.1 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-extra-column | EPPlus | 572.39 ms | 295.7 MB | 1.1 MB | OfficeIMO.Excel | 75.7% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-no-autofit | OfficeIMO.Excel | 320.07 ms | 141.7 MB | 1.9 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-no-autofit | EPPlus | 498.50 ms | 229.3 MB | 1.1 MB | OfficeIMO.Excel | 55.7% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-post-mutation | OfficeIMO.Excel | 336.55 ms | 141.8 MB | 1.9 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-post-mutation | EPPlus | 564.93 ms | 270.6 MB | 1.1 MB | OfficeIMO.Excel | 67.9% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-shuffled-columns | OfficeIMO.Excel | 325.90 ms | 141.8 MB | 2.0 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-shuffled-columns | EPPlus | 537.19 ms | 270.6 MB | 1.1 MB | OfficeIMO.Excel | 64.8% slower than OfficeIMO |
| 25000 | package-profile | report-workbook | OfficeIMO.Excel | 443.58 ms | 191.8 MB | 2.6 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | report-workbook | EPPlus | 711.76 ms | 356.2 MB | 1.5 MB | OfficeIMO.Excel | 60.5% slower than OfficeIMO |
| 25000 | package-profile | report-workbook-core | OfficeIMO.Excel | 63.24 ms | 10.7 MB | 1.8 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | report-workbook-core | EPPlus | 673.99 ms | 334.8 MB | 1.5 MB | OfficeIMO.Excel | 965.7% slower than OfficeIMO |
| 25000 | package-profile | report-workbook-core | ClosedXML | 1.45 s | 952.9 MB | 1.5 MB | OfficeIMO.Excel | 2186.8% slower than OfficeIMO |
| 25000 | package-profile | report-workbook-datatable | OfficeIMO.Excel | 467.76 ms | 194.4 MB | 2.6 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | report-workbook-datatable | EPPlus | 710.89 ms | 242.0 MB | 1.5 MB | OfficeIMO.Excel | 52.0% slower than OfficeIMO |
| 25000 | package-profile | report-workbook-datatable-core | OfficeIMO.Excel | 76.78 ms | 13.4 MB | 1.8 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | report-workbook-datatable-core | EPPlus | 678.16 ms | 220.7 MB | 1.5 MB | OfficeIMO.Excel | 783.3% slower than OfficeIMO |
| 25000 | package-profile | report-workbook-datatable-core | ClosedXML | 1.42 s | 812.7 MB | 1.5 MB | OfficeIMO.Excel | 1752.4% slower than OfficeIMO |
| 25000 | package-profile | write-blog-2023-20-string-columns | LargeXlsx | 55.09 ms | 10.5 MB | 2.4 MB | LargeXlsx | 11.6% faster than OfficeIMO |
| 25000 | package-profile | write-blog-2023-20-string-columns | OfficeIMO.Excel | 62.30 ms | 11.4 MB | 2.2 MB | LargeXlsx | Loss +13.1% |
| 25000 | package-profile | write-blog-2023-20-string-columns | MiniExcel | 199.36 ms | 221.6 MB | 2.4 MB | LargeXlsx | 220.0% slower than OfficeIMO |
| 25000 | package-profile | write-blog-2023-20-string-columns | ClosedXML | 1.19 s | 742.0 MB | 2.5 MB | LargeXlsx | 1805.3% slower than OfficeIMO |
| 25000 | package-profile | write-bulk-report | OfficeIMO.Excel | 41.76 ms | 11.3 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-bulk-report | MiniExcel | 100.58 ms | 122.6 MB | 1.5 MB | OfficeIMO.Excel | 140.9% slower than OfficeIMO |
| 25000 | package-profile | write-bulk-report | EPPlus | 507.40 ms | 249.0 MB | 1.1 MB | OfficeIMO.Excel | 1115.1% slower than OfficeIMO |
| 25000 | package-profile | write-bulk-report | ClosedXML | 1.02 s | 552.7 MB | 1.1 MB | OfficeIMO.Excel | 2346.5% slower than OfficeIMO |
| 25000 | package-profile | write-cellformula | OfficeIMO.Excel | 31.55 ms | 9.9 MB | 670.3 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellformula | ClosedXML | 266.92 ms | 111.2 MB | 643.2 KB | OfficeIMO.Excel | 746.1% slower than OfficeIMO |
| 25000 | package-profile | write-cellformula | EPPlus | 443.25 ms | 137.4 MB | 593.9 KB | OfficeIMO.Excel | 1305.1% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-empty-strings | OfficeIMO.Excel | 16.17 ms | 6.7 MB | 451.4 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-empty-strings | ClosedXML | 154.30 ms | 90.7 MB | 398.1 KB | OfficeIMO.Excel | 854.2% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-empty-strings | EPPlus | 205.95 ms | 72.7 MB | 390.6 KB | OfficeIMO.Excel | 1173.6% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-numbers | OfficeIMO.Excel | 17.61 ms | 5.8 MB | 462.6 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-numbers | ClosedXML | 143.04 ms | 82.2 MB | 411.4 KB | OfficeIMO.Excel | 712.3% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-numbers | EPPlus | 231.65 ms | 84.4 MB | 406.5 KB | OfficeIMO.Excel | 1215.5% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-mixed | OfficeIMO.Excel | 22.69 ms | 8.1 MB | 585.1 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-object-mixed | ClosedXML | 206.68 ms | 108.5 MB | 532.9 KB | OfficeIMO.Excel | 811.0% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-mixed | EPPlus | 267.33 ms | 110.6 MB | 544.3 KB | OfficeIMO.Excel | 1078.3% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-sparse | OfficeIMO.Excel | 25.91 ms | 7.2 MB | 607.1 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-object-sparse | ClosedXML | 181.20 ms | 102.8 MB | 468.0 KB | OfficeIMO.Excel | 599.3% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-sparse | EPPlus | 258.58 ms | 103.8 MB | 494.4 KB | OfficeIMO.Excel | 898.0% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 24.04 ms | 7.2 MB | 607.1 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-object-sparse-batch | ClosedXML | 204.65 ms | 102.8 MB | 468.0 KB | OfficeIMO.Excel | 751.2% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-sparse-batch | EPPlus | 291.04 ms | 103.8 MB | 494.4 KB | OfficeIMO.Excel | 1110.6% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-scalars | OfficeIMO.Excel | 15.41 ms | 6.0 MB | 441.9 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-scalars | ClosedXML | 126.06 ms | 80.7 MB | 394.9 KB | OfficeIMO.Excel | 718.1% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-scalars | EPPlus | 226.89 ms | 83.1 MB | 379.3 KB | OfficeIMO.Excel | 1372.4% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings | OfficeIMO.Excel | 21.93 ms | 15.0 MB | 527.8 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-strings | ClosedXML | 149.89 ms | 101.8 MB | 460.1 KB | OfficeIMO.Excel | 583.6% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings | EPPlus | 227.21 ms | 82.4 MB | 444.7 KB | OfficeIMO.Excel | 936.3% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings-distinct | OfficeIMO.Excel | 17.75 ms | 13.5 MB | 499.0 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-strings-distinct | ClosedXML | 202.70 ms | 128.4 MB | 555.3 KB | OfficeIMO.Excel | 1042.0% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings-distinct | EPPlus | 264.92 ms | 95.4 MB | 565.1 KB | OfficeIMO.Excel | 1392.6% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings-repeated | OfficeIMO.Excel | 16.46 ms | 7.3 MB | 376.0 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-strings-repeated | ClosedXML | 132.97 ms | 82.5 MB | 331.8 KB | OfficeIMO.Excel | 707.7% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings-repeated | EPPlus | 192.12 ms | 68.4 MB | 300.8 KB | OfficeIMO.Excel | 1067.1% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-temporal | OfficeIMO.Excel | 28.28 ms | 7.3 MB | 620.5 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-temporal | ClosedXML | 193.06 ms | 87.2 MB | 483.0 KB | OfficeIMO.Excel | 582.7% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-temporal | EPPlus | 248.15 ms | 101.4 MB | 495.1 KB | OfficeIMO.Excel | 777.6% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 12.59 ms | 3.4 MB | 443.4 KB | LargeXlsx | 14.2% faster than OfficeIMO |
| 25000 | package-profile | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 14.67 ms | 6.8 MB | 455.5 KB | LargeXlsx | Loss +16.5% |
| 25000 | package-profile | write-cellvalues-headerless-rectangle-direct | ClosedXML | 163.44 ms | 93.8 MB | 467.5 KB | LargeXlsx | 1013.8% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-headerless-rectangle-direct | EPPlus | 236.59 ms | 85.4 MB | 484.1 KB | LargeXlsx | 1512.3% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | LargeXlsx | 38.15 ms | 5.5 MB | 1.4 MB | LargeXlsx | 18.1% faster than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 46.60 ms | 15.7 MB | 1.4 MB | LargeXlsx | Loss +22.1% |
| 25000 | package-profile | write-cellvalues-rectangle-direct | MiniExcel | 94.53 ms | 91.1 MB | 1.5 MB | LargeXlsx | 102.9% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | ClosedXML | 417.65 ms | 205.7 MB | 1.1 MB | LargeXlsx | 796.3% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | EPPlus | 502.29 ms | 206.9 MB | 1.1 MB | LargeXlsx | 977.9% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | Sylvan.Data.Excel | 40.07 ms | 5.6 MB | 755.4 KB | Sylvan.Data.Excel | 32.0% faster than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | OfficeIMO.Excel | 58.95 ms | 12.7 MB | 1.4 MB | Sylvan.Data.Excel | Loss +47.1% |
| 25000 | package-profile | write-datareader-plain | LargeXlsx | 71.81 ms | 8.2 MB | 1.4 MB | Sylvan.Data.Excel | 21.8% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | MiniExcel | 111.30 ms | 90.0 MB | 1.5 MB | Sylvan.Data.Excel | 88.8% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | ClosedXML | 466.33 ms | 101.8 MB | 1.1 MB | Sylvan.Data.Excel | 691.0% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | EPPlus | 512.75 ms | 114.7 MB | 1.1 MB | Sylvan.Data.Excel | 769.8% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table | OfficeIMO.Excel | 46.01 ms | 12.7 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-datareader-table | MiniExcel | 93.22 ms | 90.0 MB | 1.5 MB | OfficeIMO.Excel | 102.6% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table | EPPlus | 468.99 ms | 114.7 MB | 1.1 MB | OfficeIMO.Excel | 919.3% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table | ClosedXML | 492.09 ms | 169.3 MB | 1.1 MB | OfficeIMO.Excel | 969.6% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table-autofit | OfficeIMO.Excel | 55.69 ms | 12.7 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-datareader-table-autofit | MiniExcel | 103.68 ms | 121.6 MB | 1.5 MB | OfficeIMO.Excel | 86.2% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table-autofit | EPPlus | 466.75 ms | 156.0 MB | 1.1 MB | OfficeIMO.Excel | 738.1% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table-autofit | ClosedXML | 1.09 s | 552.9 MB | 1.1 MB | OfficeIMO.Excel | 1858.6% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-direct-export | OfficeIMO.Excel | 42.31 ms | 9.6 MB | 1.3 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-dataset-sparse-direct-export | LargeXlsx | 48.29 ms | 9.0 MB | 1.6 MB | OfficeIMO.Excel | 14.1% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-direct-export | MiniExcel | 120.23 ms | 105.6 MB | 1.8 MB | OfficeIMO.Excel | 184.2% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-direct-export | EPPlus | 590.56 ms | 132.5 MB | 1.4 MB | OfficeIMO.Excel | 1295.9% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-direct-export | ClosedXML | 679.04 ms | 273.8 MB | 1.5 MB | OfficeIMO.Excel | 1505.1% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-tables | OfficeIMO.Excel | 49.85 ms | 13.1 MB | 1.8 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-dataset-sparse-tables | MiniExcel | 123.98 ms | 105.6 MB | 1.8 MB | OfficeIMO.Excel | 148.7% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-tables | EPPlus | 592.86 ms | 132.5 MB | 1.4 MB | OfficeIMO.Excel | 1089.3% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-tables | ClosedXML | 712.67 ms | 273.8 MB | 1.5 MB | OfficeIMO.Excel | 1329.6% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables | OfficeIMO.Excel | 49.25 ms | 10.0 MB | 1.3 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-dataset-tables | MiniExcel | 108.39 ms | 94.8 MB | 1.5 MB | OfficeIMO.Excel | 120.1% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables | EPPlus | 416.33 ms | 108.2 MB | 1.1 MB | OfficeIMO.Excel | 745.3% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables | ClosedXML | 478.62 ms | 168.0 MB | 1.1 MB | OfficeIMO.Excel | 871.8% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables-autofit | OfficeIMO.Excel | 69.82 ms | 10.1 MB | 1.3 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-dataset-tables-autofit | MiniExcel | 131.66 ms | 125.9 MB | 1.5 MB | OfficeIMO.Excel | 88.6% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables-autofit | EPPlus | 643.58 ms | 190.8 MB | 1.1 MB | OfficeIMO.Excel | 821.7% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables-autofit | ClosedXML | 1.23 s | 537.2 MB | 1.1 MB | OfficeIMO.Excel | 1665.8% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | LargeXlsx | 39.84 ms | 9.3 MB | 1.4 MB | LargeXlsx | 8.7% faster than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | OfficeIMO.Excel | 43.62 ms | 12.4 MB | 1.4 MB | LargeXlsx | Loss +9.5% |
| 25000 | package-profile | write-datatable-direct | MiniExcel | 97.14 ms | 90.2 MB | 1.5 MB | LargeXlsx | 122.7% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | ClosedXML | 340.09 ms | 101.8 MB | 1.1 MB | LargeXlsx | 679.6% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | EPPlus | 387.77 ms | 114.7 MB | 1.1 MB | LargeXlsx | 788.9% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-table-direct | OfficeIMO.Excel | 46.22 ms | 12.4 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-datatable-table-direct | MiniExcel | 113.39 ms | 90.2 MB | 1.5 MB | OfficeIMO.Excel | 145.4% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-table-direct | EPPlus | 420.29 ms | 114.7 MB | 1.1 MB | OfficeIMO.Excel | 809.4% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-table-direct | ClosedXML | 499.93 ms | 169.3 MB | 1.1 MB | OfficeIMO.Excel | 981.7% slower than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | LargeXlsx | 32.53 ms | 5.5 MB | 1.4 MB | LargeXlsx | 22.5% faster than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 41.97 ms | 12.6 MB | 1.4 MB | LargeXlsx | Loss +29.0% |
| 25000 | package-profile | write-fluent-rowsfrom-direct | MiniExcel | 84.91 ms | 91.1 MB | 1.5 MB | LargeXlsx | 102.3% slower than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | ClosedXML | 339.85 ms | 101.8 MB | 1.1 MB | LargeXlsx | 709.8% slower than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | EPPlus | 390.85 ms | 114.7 MB | 1.1 MB | LargeXlsx | 831.3% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 54.71 ms | 11.2 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 629.14 ms | 156.0 MB | 1.1 MB | OfficeIMO.Excel | 1050.0% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 994.61 ms | 485.3 MB | 1.1 MB | OfficeIMO.Excel | 1718.1% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | LargeXlsx | 35.13 ms | 5.5 MB | 1.4 MB | LargeXlsx | 18.0% faster than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | OfficeIMO.Excel | 42.85 ms | 11.2 MB | 1.4 MB | LargeXlsx | Loss +22.0% |
| 25000 | package-profile | write-insertobjects-direct | MiniExcel | 89.61 ms | 91.1 MB | 1.5 MB | LargeXlsx | 109.1% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | ClosedXML | 364.78 ms | 101.8 MB | 1.1 MB | LargeXlsx | 751.2% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | EPPlus | 435.73 ms | 114.7 MB | 1.1 MB | LargeXlsx | 916.7% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 59.48 ms | 9.9 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 477.41 ms | 156.0 MB | 1.1 MB | OfficeIMO.Excel | 702.6% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 926.61 ms | 485.3 MB | 1.1 MB | OfficeIMO.Excel | 1457.7% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 36.21 ms | 5.5 MB | 1.4 MB | LargeXlsx | 30.3% faster than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 51.95 ms | 9.9 MB | 1.4 MB | LargeXlsx | Loss +43.5% |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | MiniExcel | 84.91 ms | 91.1 MB | 1.5 MB | LargeXlsx | 63.5% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | ClosedXML | 345.72 ms | 101.8 MB | 1.1 MB | LargeXlsx | 565.5% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | EPPlus | 424.55 ms | 114.7 MB | 1.1 MB | LargeXlsx | 717.3% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | LargeXlsx | 38.66 ms | 5.5 MB | 1.4 MB | LargeXlsx | 38.1% faster than OfficeIMO |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | OfficeIMO.Excel | 62.41 ms | 15.4 MB | 1.4 MB | LargeXlsx | Loss +61.4% |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | MiniExcel | 90.62 ms | 91.1 MB | 1.5 MB | LargeXlsx | 45.2% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | ClosedXML | 348.79 ms | 101.8 MB | 1.1 MB | LargeXlsx | 458.9% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | EPPlus | 422.81 ms | 114.7 MB | 1.1 MB | LargeXlsx | 577.5% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 46.77 ms | 11.2 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 424.52 ms | 135.1 MB | 1.1 MB | OfficeIMO.Excel | 807.6% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 545.26 ms | 269.0 MB | 1.1 MB | OfficeIMO.Excel | 1065.7% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-mixed-objects-direct | LargeXlsx | 51.04 ms | 5.9 MB | 1.8 MB | LargeXlsx | 11.9% faster than OfficeIMO |
| 25000 | package-profile | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 57.93 ms | 10.3 MB | 1.8 MB | LargeXlsx | Loss +13.5% |
| 25000 | package-profile | write-powershell-mixed-objects-direct | MiniExcel | 109.78 ms | 111.3 MB | 1.9 MB | LargeXlsx | 89.5% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-mixed-objects-direct | ClosedXML | 488.96 ms | 175.3 MB | 1.5 MB | LargeXlsx | 744.0% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-mixed-objects-direct | EPPlus | 556.52 ms | 141.5 MB | 1.4 MB | LargeXlsx | 860.6% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-mixed-direct | LargeXlsx | 49.15 ms | 5.9 MB | 1.8 MB | LargeXlsx | 9.7% faster than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 54.45 ms | 9.7 MB | 1.8 MB | LargeXlsx | Loss +10.8% |
| 25000 | package-profile | write-powershell-psobject-mixed-direct | MiniExcel | 108.65 ms | 111.3 MB | 1.9 MB | LargeXlsx | 99.5% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-mixed-direct | ClosedXML | 473.38 ms | 175.3 MB | 1.5 MB | LargeXlsx | 769.4% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-mixed-direct | EPPlus | 538.07 ms | 141.5 MB | 1.4 MB | LargeXlsx | 888.2% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 271.73 ms | 35.3 MB | 6.6 MB | OfficeIMO.Excel, LargeXlsx | Win |
| 25000 | package-profile | write-powershell-psobject-wide-direct | LargeXlsx | 275.39 ms | 22.7 MB | 6.5 MB | OfficeIMO.Excel, LargeXlsx | Tie vs OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-wide-direct | MiniExcel | 454.22 ms | 339.8 MB | 6.8 MB | OfficeIMO.Excel, LargeXlsx | 67.2% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-wide-direct | ClosedXML | 1.45 s | 476.0 MB | 6.0 MB | OfficeIMO.Excel, LargeXlsx | 435.3% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-wide-direct | EPPlus | 1.87 s | 549.7 MB | 5.3 MB | OfficeIMO.Excel, LargeXlsx | 588.1% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | LargeXlsx | 11.57 ms | 2.7 MB |  | LargeXlsx | 31.1% faster than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | OfficeIMO.Excel | 16.80 ms | 10.6 MB |  | LargeXlsx | Loss +45.2% |
| 25000 | speed-comparison | append-plain-rows | MiniExcel | 31.96 ms | 56.9 MB |  | LargeXlsx | 90.3% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | EPPlus 4.5.3.3 | 144.17 ms | 0 B |  | LargeXlsx | 758.4% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | ClosedXML | 144.22 ms | 101.8 MB |  | LargeXlsx | 758.7% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | EPPlus | 212.19 ms | 98.0 MB |  | LargeXlsx | 1163.4% slower than OfficeIMO |
| 25000 | speed-comparison | autofit-existing | OfficeIMO.Excel | 85.60 ms | 15.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | autofit-existing | EPPlus | 460.13 ms | 245.1 MB |  | OfficeIMO.Excel | 437.5% slower than OfficeIMO |
| 25000 | speed-comparison | autofit-existing | EPPlus 4.5.3.3 | 779.69 ms | 0 B |  | OfficeIMO.Excel | 810.9% slower than OfficeIMO |
| 25000 | speed-comparison | autofit-existing | ClosedXML | 1.46 s | 810.3 MB |  | OfficeIMO.Excel | 1607.4% slower than OfficeIMO |
| 25000 | speed-comparison | build-object-datatable-dictionaries | OfficeIMO.Excel | 8.95 ms | 5.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | build-object-datatable-typed | OfficeIMO.Excel | 8.15 ms | 7.9 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-cells | OfficeIMO.Excel | 52.17 ms | 24.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-cells | EPPlus | 267.12 ms | 183.0 MB |  | OfficeIMO.Excel | 412.0% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-cells | ClosedXML | 348.65 ms | 162.6 MB |  | OfficeIMO.Excel | 568.2% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-first-column-from-wide-sheet | OfficeIMO.Excel | 35.19 ms | 3.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-first-column-from-wide-sheet | EPPlus | 227.30 ms | 112.8 MB |  | OfficeIMO.Excel | 545.9% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-first-column-from-wide-sheet | ClosedXML | 309.41 ms | 147.4 MB |  | OfficeIMO.Excel | 779.3% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-range | OfficeIMO.Excel | 49.51 ms | 24.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-range | EPPlus | 279.89 ms | 183.0 MB |  | OfficeIMO.Excel | 465.3% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-range | ClosedXML | 359.43 ms | 162.6 MB |  | OfficeIMO.Excel | 626.0% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-top-range | OfficeIMO.Excel | 0.61 ms | 285.3 KB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-top-range | EPPlus | 226.94 ms | 103.1 MB |  | OfficeIMO.Excel | 37213.7% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-top-range | ClosedXML | 331.90 ms | 145.9 MB |  | OfficeIMO.Excel | 54470.9% slower than OfficeIMO |
| 25000 | speed-comparison | formula-heavy-read | OfficeIMO.Excel | 28.98 ms | 6.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | formula-heavy-read | EPPlus 4.5.3.3 | 97.60 ms | 0 B |  | OfficeIMO.Excel | 236.8% slower than OfficeIMO |
| 25000 | speed-comparison | formula-heavy-read | EPPlus | 193.07 ms | 69.2 MB |  | OfficeIMO.Excel | 566.3% slower than OfficeIMO |
| 25000 | speed-comparison | formula-heavy-read | ClosedXML | 235.57 ms | 77.6 MB |  | OfficeIMO.Excel | 713.0% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | OfficeIMO.Excel | 20.03 ms | 15.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | large-shared-strings | MiniExcel | 37.70 ms | 72.0 MB |  | OfficeIMO.Excel | 88.2% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | EPPlus 4.5.3.3 | 106.29 ms | 0 B |  | OfficeIMO.Excel | 430.7% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | ClosedXML | 146.81 ms | 101.8 MB |  | OfficeIMO.Excel | 633.0% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | EPPlus | 220.49 ms | 82.4 MB |  | OfficeIMO.Excel | 1000.8% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | OfficeIMO.Excel | 0.86 ms | 177.5 KB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | large-sparse-column-read | Sylvan.Data.Excel | 1.22 ms | 316.6 KB |  | OfficeIMO.Excel | 41.8% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | ExcelDataReader | 1.45 ms | 4.0 MB |  | OfficeIMO.Excel | 68.5% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | ClosedXML | 3.71 ms | 4.3 MB |  | OfficeIMO.Excel | 331.2% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | MiniExcel | 13.56 ms | 45.1 MB |  | OfficeIMO.Excel | 1474.5% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | EPPlus 4.5.3.3 | 28.10 ms | 0 B |  | OfficeIMO.Excel | 3162.4% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | EPPlus | 101.17 ms | 42.1 MB |  | OfficeIMO.Excel | 11645.3% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | OfficeIMO.Excel | 0.86 ms | 177.6 KB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | large-sparse-row-read | Sylvan.Data.Excel | 1.04 ms | 316.6 KB |  | OfficeIMO.Excel | 20.3% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | ExcelDataReader | 1.76 ms | 4.0 MB |  | OfficeIMO.Excel | 103.7% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | ClosedXML | 4.16 ms | 4.3 MB |  | OfficeIMO.Excel | 380.8% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | MiniExcel | 13.58 ms | 45.1 MB |  | OfficeIMO.Excel | 1470.0% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | EPPlus 4.5.3.3 | 27.99 ms | 0 B |  | OfficeIMO.Excel | 3136.9% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | EPPlus | 95.81 ms | 42.1 MB |  | OfficeIMO.Excel | 10980.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | OfficeIMO.Excel | 32.46 ms | 1.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-bottom-range | Sylvan.Data.Excel | 38.21 ms | 3.5 MB |  | OfficeIMO.Excel | 17.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | ExcelDataReader | 104.17 ms | 59.8 MB |  | OfficeIMO.Excel | 220.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | MiniExcel | 118.85 ms | 182.1 MB |  | OfficeIMO.Excel | 266.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | EPPlus | 227.00 ms | 103.1 MB |  | OfficeIMO.Excel | 599.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | ClosedXML | 317.74 ms | 145.9 MB |  | OfficeIMO.Excel | 878.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | OfficeIMO.Excel | 36.09 ms | 1.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-bottom-range-stream | Sylvan.Data.Excel | 48.42 ms | 3.5 MB |  | OfficeIMO.Excel | 34.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | ExcelDataReader | 132.34 ms | 59.8 MB |  | OfficeIMO.Excel | 266.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | MiniExcel | 136.78 ms | 182.1 MB |  | OfficeIMO.Excel | 279.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | EPPlus | 333.00 ms | 103.1 MB |  | OfficeIMO.Excel | 822.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | ClosedXML | 381.75 ms | 145.9 MB |  | OfficeIMO.Excel | 957.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | Sylvan.Data.Excel | 58.59 ms | 18.0 MB |  | Sylvan.Data.Excel | 6.2% faster than OfficeIMO |
| 25000 | speed-comparison | read-datatable | OfficeIMO.Excel | 62.47 ms | 33.8 MB |  | Sylvan.Data.Excel | Loss +6.6% |
| 25000 | speed-comparison | read-datatable | ExcelDataReader | 140.43 ms | 74.3 MB |  | Sylvan.Data.Excel | 124.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | MiniExcel | 156.96 ms | 177.0 MB |  | Sylvan.Data.Excel | 151.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | EPPlus 4.5.3.3 | 256.28 ms | 0 B |  | Sylvan.Data.Excel | 310.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | EPPlus | 278.70 ms | 197.5 MB |  | Sylvan.Data.Excel | 346.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | ClosedXML | 357.61 ms | 174.3 MB |  | Sylvan.Data.Excel | 472.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | OfficeIMO.Excel | 36.67 ms | 3.9 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | Sylvan.Data.Excel | 49.37 ms | 4.2 MB |  | OfficeIMO.Excel | 34.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | MiniExcel | 97.14 ms | 154.9 MB |  | OfficeIMO.Excel | 164.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | ExcelDataReader | 115.78 ms | 59.8 MB |  | OfficeIMO.Excel | 215.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | EPPlus | 237.72 ms | 112.8 MB |  | OfficeIMO.Excel | 548.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | ClosedXML | 328.79 ms | 147.4 MB |  | OfficeIMO.Excel | 796.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | Sylvan.Data.Excel | 47.14 ms | 5.7 MB |  | Sylvan.Data.Excel | 16.0% faster than OfficeIMO |
| 25000 | speed-comparison | read-objects | OfficeIMO.Excel | 56.11 ms | 23.0 MB |  | Sylvan.Data.Excel | Loss +19.0% |
| 25000 | speed-comparison | read-objects | ExcelDataReader | 114.10 ms | 62.0 MB |  | Sylvan.Data.Excel | 103.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | MiniExcel | 138.95 ms | 179.4 MB |  | Sylvan.Data.Excel | 147.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | EPPlus 4.5.3.3 | 239.53 ms | 0 B |  | Sylvan.Data.Excel | 326.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | EPPlus | 258.29 ms | 194.9 MB |  | Sylvan.Data.Excel | 360.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | ClosedXML | 330.96 ms | 161.7 MB |  | Sylvan.Data.Excel | 489.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | OfficeIMO.Excel | 45.62 ms | 22.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-objects-stream | Sylvan.Data.Excel | 46.80 ms | 5.2 MB |  | OfficeIMO.Excel | 2.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | ExcelDataReader | 107.50 ms | 61.5 MB |  | OfficeIMO.Excel | 135.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | MiniExcel | 119.80 ms | 178.9 MB |  | OfficeIMO.Excel | 162.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | EPPlus 4.5.3.3 | 233.78 ms | 0 B |  | OfficeIMO.Excel | 412.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | EPPlus | 248.56 ms | 194.7 MB |  | OfficeIMO.Excel | 444.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | ClosedXML | 322.58 ms | 161.5 MB |  | OfficeIMO.Excel | 607.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | Sylvan.Data.Excel | 43.73 ms | 3.5 MB |  | Sylvan.Data.Excel | 17.0% faster than OfficeIMO |
| 25000 | speed-comparison | read-range | OfficeIMO.Excel | 52.68 ms | 25.5 MB |  | Sylvan.Data.Excel | Loss +20.5% |
| 25000 | speed-comparison | read-range | ExcelDataReader | 112.35 ms | 59.8 MB |  | Sylvan.Data.Excel | 113.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | MiniExcel | 123.08 ms | 182.1 MB |  | Sylvan.Data.Excel | 133.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | EPPlus | 257.22 ms | 183.0 MB |  | Sylvan.Data.Excel | 388.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | EPPlus 4.5.3.3 | 260.85 ms | 0 B |  | Sylvan.Data.Excel | 395.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | ClosedXML | 336.19 ms | 159.8 MB |  | Sylvan.Data.Excel | 538.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | Sylvan.Data.Excel | 51.66 ms | 4.4 MB |  | Sylvan.Data.Excel | 4.8% faster than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | OfficeIMO.Excel | 54.25 ms | 26.1 MB |  | Sylvan.Data.Excel | Loss +5.0% |
| 25000 | speed-comparison | read-range-decimal | ExcelDataReader | 118.38 ms | 59.8 MB |  | Sylvan.Data.Excel | 118.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | MiniExcel | 135.67 ms | 182.1 MB |  | Sylvan.Data.Excel | 150.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | EPPlus | 270.08 ms | 183.0 MB |  | Sylvan.Data.Excel | 397.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | ClosedXML | 366.96 ms | 159.8 MB |  | Sylvan.Data.Excel | 576.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | Sylvan.Data.Excel | 45.35 ms | 3.5 MB |  | Sylvan.Data.Excel | 12.5% faster than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | OfficeIMO.Excel | 51.84 ms | 26.3 MB |  | Sylvan.Data.Excel | Loss +14.3% |
| 25000 | speed-comparison | read-range-stream | ExcelDataReader | 115.16 ms | 59.8 MB |  | Sylvan.Data.Excel | 122.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | MiniExcel | 123.80 ms | 182.1 MB |  | Sylvan.Data.Excel | 138.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | EPPlus 4.5.3.3 | 211.23 ms | 0 B |  | Sylvan.Data.Excel | 307.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | EPPlus | 257.52 ms | 183.0 MB |  | Sylvan.Data.Excel | 396.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | ClosedXML | 350.82 ms | 159.8 MB |  | Sylvan.Data.Excel | 576.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | Sylvan.Data.Excel | 0.51 ms | 348.5 KB |  | Sylvan.Data.Excel | 11.8% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range | OfficeIMO.Excel | 0.57 ms | 296.0 KB |  | Sylvan.Data.Excel | Loss +13.4% |
| 25000 | speed-comparison | read-top-range | MiniExcel | 0.79 ms | 869.0 KB |  | Sylvan.Data.Excel | 38.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | ExcelDataReader | 40.72 ms | 16.7 MB |  | Sylvan.Data.Excel | 6991.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | EPPlus 4.5.3.3 | 193.50 ms | 0 B |  | Sylvan.Data.Excel | 33599.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | EPPlus | 244.03 ms | 103.1 MB |  | Sylvan.Data.Excel | 42399.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | ClosedXML | 331.00 ms | 145.9 MB |  | Sylvan.Data.Excel | 57545.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | Sylvan.Data.Excel | 0.45 ms | 348.5 KB |  | Sylvan.Data.Excel | 16.4% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | OfficeIMO.Excel | 0.54 ms | 299.3 KB |  | Sylvan.Data.Excel | Loss +19.6% |
| 25000 | speed-comparison | read-top-range-stream | MiniExcel | 1.18 ms | 869.0 KB |  | Sylvan.Data.Excel | 117.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | ExcelDataReader | 42.43 ms | 16.7 MB |  | Sylvan.Data.Excel | 7726.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | EPPlus 4.5.3.3 | 192.82 ms | 0 B |  | Sylvan.Data.Excel | 35468.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | EPPlus | 226.60 ms | 103.1 MB |  | Sylvan.Data.Excel | 41700.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | ClosedXML | 345.53 ms | 145.9 MB |  | Sylvan.Data.Excel | 63639.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | Sylvan.Data.Excel | 0.43 ms | 348.5 KB |  | Sylvan.Data.Excel | 19.3% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | OfficeIMO.Excel | 0.54 ms | 300.1 KB |  | Sylvan.Data.Excel | Loss +23.9% |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | MiniExcel | 0.72 ms | 869.0 KB |  | Sylvan.Data.Excel | 32.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | ExcelDataReader | 40.17 ms | 16.7 MB |  | Sylvan.Data.Excel | 7353.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | EPPlus | 346.90 ms | 103.1 MB |  | Sylvan.Data.Excel | 64272.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | ClosedXML | 358.96 ms | 145.9 MB |  | Sylvan.Data.Excel | 66510.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-used-range | Sylvan.Data.Excel | 45.12 ms | 3.5 MB |  | Sylvan.Data.Excel | 48.1% faster than OfficeIMO |
| 25000 | speed-comparison | read-used-range | OfficeIMO.Excel | 86.94 ms | 33.4 MB |  | Sylvan.Data.Excel | Loss +92.7% |
| 25000 | speed-comparison | read-used-range | ExcelDataReader | 114.29 ms | 59.8 MB |  | Sylvan.Data.Excel | 31.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-used-range | MiniExcel | 127.39 ms | 182.1 MB |  | Sylvan.Data.Excel | 46.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-used-range | EPPlus | 273.40 ms | 183.0 MB |  | Sylvan.Data.Excel | 214.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-used-range | ClosedXML | 346.12 ms | 159.8 MB |  | Sylvan.Data.Excel | 298.1% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-autofilter | OfficeIMO.Excel | 43.12 ms | 11.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-autofilter | EPPlus 4.5.3.3 | 336.57 ms | 0 B |  | OfficeIMO.Excel | 680.6% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-autofilter | ClosedXML | 394.06 ms | 205.7 MB |  | OfficeIMO.Excel | 814.0% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-autofilter | EPPlus | 433.24 ms | 206.9 MB |  | OfficeIMO.Excel | 904.9% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-charts | OfficeIMO.Excel | 44.45 ms | 12.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-charts | EPPlus 4.5.3.3 | 288.53 ms | 0 B |  | OfficeIMO.Excel | 549.1% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-charts | EPPlus | 452.20 ms | 209.9 MB |  | OfficeIMO.Excel | 917.3% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-conditional-formatting | OfficeIMO.Excel | 42.96 ms | 11.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-conditional-formatting | EPPlus 4.5.3.3 | 305.31 ms | 0 B |  | OfficeIMO.Excel | 610.7% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-conditional-formatting | ClosedXML | 373.78 ms | 205.8 MB |  | OfficeIMO.Excel | 770.1% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-conditional-formatting | EPPlus | 430.02 ms | 206.9 MB |  | OfficeIMO.Excel | 901.1% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-data-validation | OfficeIMO.Excel | 43.52 ms | 11.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-data-validation | EPPlus 4.5.3.3 | 291.80 ms | 0 B |  | OfficeIMO.Excel | 570.5% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-data-validation | ClosedXML | 367.24 ms | 205.7 MB |  | OfficeIMO.Excel | 743.9% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-data-validation | EPPlus | 428.36 ms | 206.9 MB |  | OfficeIMO.Excel | 884.3% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-freeze-panes | OfficeIMO.Excel | 41.39 ms | 11.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-freeze-panes | EPPlus 4.5.3.3 | 317.32 ms | 0 B |  | OfficeIMO.Excel | 666.7% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-freeze-panes | ClosedXML | 382.25 ms | 205.7 MB |  | OfficeIMO.Excel | 823.6% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-freeze-panes | EPPlus | 458.18 ms | 206.9 MB |  | OfficeIMO.Excel | 1007.1% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-pivot-table | EPPlus 4.5.3.3 | 295.79 ms | 0 B |  | EPPlus 4.5.3.3, OfficeIMO.Excel | Tie vs OfficeIMO |
| 25000 | speed-comparison | realworld-pivot-table | OfficeIMO.Excel | 299.58 ms | 140.3 MB |  | EPPlus 4.5.3.3, OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-pivot-table | EPPlus | 464.34 ms | 225.4 MB |  | EPPlus 4.5.3.3, OfficeIMO.Excel | 55.0% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-all-in-one | OfficeIMO.Excel | 226.59 ms | 141.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-all-in-one | EPPlus | 435.10 ms | 270.6 MB |  | OfficeIMO.Excel | 92.0% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-all-in-one | EPPlus 4.5.3.3 | 644.05 ms | 0 B |  | OfficeIMO.Excel | 184.2% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-chart-first | OfficeIMO.Excel | 93.06 ms | 54.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-chart-first | EPPlus | 434.66 ms | 270.6 MB |  | OfficeIMO.Excel | 367.1% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-chart-first | EPPlus 4.5.3.3 | 595.63 ms | 0 B |  | OfficeIMO.Excel | 540.1% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-core | OfficeIMO.Excel | 46.34 ms | 11.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-core | EPPlus | 471.30 ms | 249.1 MB |  | OfficeIMO.Excel | 917.1% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-core | EPPlus 4.5.3.3 | 736.59 ms | 0 B |  | OfficeIMO.Excel | 1489.6% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-core | ClosedXML | 1.04 s | 664.2 MB |  | OfficeIMO.Excel | 2148.1% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-extra-column | OfficeIMO.Excel | 356.27 ms | 153.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-extra-column | EPPlus | 528.93 ms | 295.7 MB |  | OfficeIMO.Excel | 48.5% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-extra-column | EPPlus 4.5.3.3 | 1.02 s | 0 B |  | OfficeIMO.Excel | 185.5% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-no-autofit | OfficeIMO.Excel | 234.38 ms | 141.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-no-autofit | EPPlus 4.5.3.3 | 307.49 ms | 0 B |  | OfficeIMO.Excel | 31.2% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-no-autofit | EPPlus | 405.18 ms | 229.3 MB |  | OfficeIMO.Excel | 72.9% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-post-mutation | OfficeIMO.Excel | 320.89 ms | 141.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-post-mutation | EPPlus | 515.92 ms | 270.6 MB |  | OfficeIMO.Excel | 60.8% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-post-mutation | EPPlus 4.5.3.3 | 804.51 ms | 0 B |  | OfficeIMO.Excel | 150.7% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-shuffled-columns | OfficeIMO.Excel | 253.35 ms | 141.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-shuffled-columns | EPPlus | 463.37 ms | 270.6 MB |  | OfficeIMO.Excel | 82.9% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-shuffled-columns | EPPlus 4.5.3.3 | 622.17 ms | 0 B |  | OfficeIMO.Excel | 145.6% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook | OfficeIMO.Excel | 342.77 ms | 191.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | report-workbook | EPPlus | 571.24 ms | 356.2 MB |  | OfficeIMO.Excel | 66.7% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook | EPPlus 4.5.3.3 | 841.42 ms | 0 B |  | OfficeIMO.Excel | 145.5% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-core | OfficeIMO.Excel | 51.45 ms | 10.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | report-workbook-core | EPPlus | 563.86 ms | 334.8 MB |  | OfficeIMO.Excel | 995.9% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-core | EPPlus 4.5.3.3 | 857.15 ms | 0 B |  | OfficeIMO.Excel | 1565.9% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-core | ClosedXML | 1.19 s | 952.9 MB |  | OfficeIMO.Excel | 2204.5% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-datatable | OfficeIMO.Excel | 334.91 ms | 194.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | report-workbook-datatable | EPPlus | 560.46 ms | 242.0 MB |  | OfficeIMO.Excel | 67.3% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-datatable | EPPlus 4.5.3.3 | 854.84 ms | 0 B |  | OfficeIMO.Excel | 155.2% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-datatable-core | OfficeIMO.Excel | 50.73 ms | 13.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | report-workbook-datatable-core | EPPlus | 527.51 ms | 220.7 MB |  | OfficeIMO.Excel | 939.9% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-datatable-core | EPPlus 4.5.3.3 | 872.50 ms | 0 B |  | OfficeIMO.Excel | 1620.0% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-datatable-core | ClosedXML | 1.06 s | 812.7 MB |  | OfficeIMO.Excel | 1979.9% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | Sylvan.Data.Excel | 26.16 ms | 1.9 MB |  | Sylvan.Data.Excel | 16.4% faster than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | OfficeIMO.Excel | 31.30 ms | 9.0 MB |  | Sylvan.Data.Excel | Loss +19.7% |
| 25000 | speed-comparison | shared-string-read | ExcelDataReader | 62.56 ms | 24.4 MB |  | Sylvan.Data.Excel | 99.9% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | MiniExcel | 64.67 ms | 72.7 MB |  | Sylvan.Data.Excel | 106.6% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | EPPlus 4.5.3.3 | 111.48 ms | 0 B |  | Sylvan.Data.Excel | 256.2% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | EPPlus | 183.69 ms | 87.3 MB |  | Sylvan.Data.Excel | 486.9% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | ClosedXML | 219.18 ms | 88.3 MB |  | Sylvan.Data.Excel | 600.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-blog-2023-20-string-columns | LargeXlsx | 38.84 ms | 10.5 MB |  | LargeXlsx | 20.6% faster than OfficeIMO |
| 25000 | speed-comparison | write-blog-2023-20-string-columns | OfficeIMO.Excel | 48.91 ms | 11.4 MB |  | LargeXlsx | Loss +26.0% |
| 25000 | speed-comparison | write-blog-2023-20-string-columns | MiniExcel | 159.38 ms | 221.6 MB |  | LargeXlsx | 225.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-blog-2023-20-string-columns | ClosedXML | 937.90 ms | 742.0 MB |  | LargeXlsx | 1817.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | OfficeIMO.Excel | 41.74 ms | 11.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-bulk-report | MiniExcel | 69.58 ms | 122.6 MB |  | OfficeIMO.Excel | 66.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | EPPlus | 458.95 ms | 249.0 MB |  | OfficeIMO.Excel | 999.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | EPPlus 4.5.3.3 | 602.19 ms | 0 B |  | OfficeIMO.Excel | 1342.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | ClosedXML | 868.51 ms | 552.7 MB |  | OfficeIMO.Excel | 1980.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellformula | OfficeIMO.Excel | 21.59 ms | 9.9 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellformula | EPPlus 4.5.3.3 | 190.19 ms | 0 B |  | OfficeIMO.Excel | 781.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellformula | ClosedXML | 191.72 ms | 111.2 MB |  | OfficeIMO.Excel | 788.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellformula | EPPlus | 343.57 ms | 137.4 MB |  | OfficeIMO.Excel | 1491.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-empty-strings | OfficeIMO.Excel | 12.34 ms | 6.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-empty-strings | ClosedXML | 115.57 ms | 90.7 MB |  | OfficeIMO.Excel | 836.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-empty-strings | EPPlus | 168.93 ms | 72.7 MB |  | OfficeIMO.Excel | 1268.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-numbers | OfficeIMO.Excel | 14.87 ms | 5.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-numbers | ClosedXML | 106.97 ms | 82.2 MB |  | OfficeIMO.Excel | 619.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-numbers | EPPlus 4.5.3.3 | 123.82 ms | 0 B |  | OfficeIMO.Excel | 732.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-numbers | EPPlus | 185.49 ms | 84.4 MB |  | OfficeIMO.Excel | 1147.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-mixed | OfficeIMO.Excel | 17.26 ms | 8.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-object-mixed | EPPlus 4.5.3.3 | 152.59 ms | 0 B |  | OfficeIMO.Excel | 783.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-mixed | ClosedXML | 169.43 ms | 108.5 MB |  | OfficeIMO.Excel | 881.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-mixed | EPPlus | 216.54 ms | 110.6 MB |  | OfficeIMO.Excel | 1154.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-sparse | OfficeIMO.Excel | 19.43 ms | 7.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-object-sparse | ClosedXML | 146.59 ms | 102.8 MB |  | OfficeIMO.Excel | 654.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-sparse | EPPlus | 221.18 ms | 103.8 MB |  | OfficeIMO.Excel | 1038.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 17.85 ms | 7.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-object-sparse-batch | ClosedXML | 146.67 ms | 102.8 MB |  | OfficeIMO.Excel | 721.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-sparse-batch | EPPlus | 222.85 ms | 103.8 MB |  | OfficeIMO.Excel | 1148.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-scalars | OfficeIMO.Excel | 10.97 ms | 6.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-scalars | ClosedXML | 103.62 ms | 80.7 MB |  | OfficeIMO.Excel | 844.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-scalars | EPPlus 4.5.3.3 | 126.15 ms | 0 B |  | OfficeIMO.Excel | 1050.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-scalars | EPPlus | 180.24 ms | 83.1 MB |  | OfficeIMO.Excel | 1543.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings | OfficeIMO.Excel | 17.16 ms | 15.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-strings | ClosedXML | 114.32 ms | 101.8 MB |  | OfficeIMO.Excel | 566.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings | EPPlus 4.5.3.3 | 114.75 ms | 0 B |  | OfficeIMO.Excel | 568.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings | EPPlus | 191.80 ms | 82.4 MB |  | OfficeIMO.Excel | 1017.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings-distinct | OfficeIMO.Excel | 12.68 ms | 13.5 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-strings-distinct | ClosedXML | 154.91 ms | 128.4 MB |  | OfficeIMO.Excel | 1121.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings-distinct | EPPlus | 209.41 ms | 95.4 MB |  | OfficeIMO.Excel | 1551.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings-repeated | OfficeIMO.Excel | 12.81 ms | 7.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-strings-repeated | ClosedXML | 98.22 ms | 82.5 MB |  | OfficeIMO.Excel | 667.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings-repeated | EPPlus | 145.99 ms | 68.4 MB |  | OfficeIMO.Excel | 1040.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-temporal | OfficeIMO.Excel | 26.79 ms | 7.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-temporal | EPPlus 4.5.3.3 | 137.39 ms | 0 B |  | OfficeIMO.Excel | 412.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-temporal | ClosedXML | 146.06 ms | 87.2 MB |  | OfficeIMO.Excel | 445.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-temporal | EPPlus | 198.90 ms | 101.4 MB |  | OfficeIMO.Excel | 642.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 11.45 ms | 6.8 MB |  | OfficeIMO.Excel, LargeXlsx | Win |
| 25000 | speed-comparison | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 11.49 ms | 3.4 MB |  | OfficeIMO.Excel, LargeXlsx | Tie vs OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-headerless-rectangle-direct | ClosedXML | 122.84 ms | 93.8 MB |  | OfficeIMO.Excel, LargeXlsx | 972.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-headerless-rectangle-direct | EPPlus | 184.56 ms | 85.4 MB |  | OfficeIMO.Excel, LargeXlsx | 1512.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | LargeXlsx | 29.47 ms | 5.5 MB |  | LargeXlsx | 14.9% faster than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 34.62 ms | 15.7 MB |  | LargeXlsx | Loss +17.5% |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | MiniExcel | 63.13 ms | 91.1 MB |  | LargeXlsx | 82.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 292.85 ms | 0 B |  | LargeXlsx | 745.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | ClosedXML | 293.21 ms | 205.7 MB |  | LargeXlsx | 746.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus | 353.97 ms | 206.9 MB |  | LargeXlsx | 922.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-sparse-rectangle-direct | OfficeIMO.Excel | 21.86 ms | 7.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalues-sparse-rectangle-direct | ClosedXML | 141.69 ms | 102.8 MB |  | OfficeIMO.Excel | 548.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-sparse-rectangle-direct | EPPlus | 206.85 ms | 103.8 MB |  | OfficeIMO.Excel | 846.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | Sylvan.Data.Excel | 27.98 ms | 5.6 MB |  | Sylvan.Data.Excel | 32.8% faster than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | LargeXlsx | 35.55 ms | 8.2 MB |  | Sylvan.Data.Excel | 14.7% faster than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | OfficeIMO.Excel | 41.67 ms | 12.7 MB |  | Sylvan.Data.Excel | Loss +48.9% |
| 25000 | speed-comparison | write-datareader-plain | MiniExcel | 71.40 ms | 90.0 MB |  | Sylvan.Data.Excel | 71.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | EPPlus 4.5.3.3 | 269.37 ms | 0 B |  | Sylvan.Data.Excel | 546.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | ClosedXML | 275.70 ms | 101.8 MB |  | Sylvan.Data.Excel | 561.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | EPPlus | 317.33 ms | 114.7 MB |  | Sylvan.Data.Excel | 661.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | OfficeIMO.Excel | 37.14 ms | 12.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datareader-table | MiniExcel | 80.41 ms | 90.0 MB |  | OfficeIMO.Excel | 116.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | EPPlus 4.5.3.3 | 315.77 ms | 0 B |  | OfficeIMO.Excel | 750.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | EPPlus | 326.33 ms | 114.7 MB |  | OfficeIMO.Excel | 778.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | ClosedXML | 377.19 ms | 169.3 MB |  | OfficeIMO.Excel | 915.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table-autofit | OfficeIMO.Excel | 42.59 ms | 12.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datareader-table-autofit | MiniExcel | 71.88 ms | 121.6 MB |  | OfficeIMO.Excel | 68.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table-autofit | EPPlus | 370.44 ms | 156.0 MB |  | OfficeIMO.Excel | 769.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table-autofit | EPPlus 4.5.3.3 | 610.52 ms | 0 B |  | OfficeIMO.Excel | 1333.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table-autofit | ClosedXML | 775.61 ms | 552.9 MB |  | OfficeIMO.Excel | 1721.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-headerless-tables | OfficeIMO.Excel | 39.94 ms | 12.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-headerless-tables | MiniExcel | 86.38 ms | 94.8 MB |  | OfficeIMO.Excel | 116.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-headerless-tables | ClosedXML | 367.90 ms | 168.0 MB |  | OfficeIMO.Excel | 821.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-headerless-tables | EPPlus | 487.20 ms | 108.6 MB |  | OfficeIMO.Excel | 1119.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | OfficeIMO.Excel | 34.90 ms | 9.6 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | LargeXlsx | 43.39 ms | 9.0 MB |  | OfficeIMO.Excel | 24.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | MiniExcel | 106.07 ms | 105.6 MB |  | OfficeIMO.Excel | 203.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | EPPlus | 599.00 ms | 132.5 MB |  | OfficeIMO.Excel | 1616.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | ClosedXML | 605.72 ms | 273.8 MB |  | OfficeIMO.Excel | 1635.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-tables | OfficeIMO.Excel | 45.31 ms | 13.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-sparse-tables | MiniExcel | 122.85 ms | 105.6 MB |  | OfficeIMO.Excel | 171.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-tables | ClosedXML | 561.97 ms | 273.8 MB |  | OfficeIMO.Excel | 1140.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-tables | EPPlus | 681.54 ms | 132.5 MB |  | OfficeIMO.Excel | 1404.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | OfficeIMO.Excel | 38.33 ms | 10.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-tables | MiniExcel | 85.48 ms | 94.8 MB |  | OfficeIMO.Excel | 123.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | EPPlus 4.5.3.3 | 337.04 ms | 0 B |  | OfficeIMO.Excel | 779.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | ClosedXML | 365.94 ms | 168.0 MB |  | OfficeIMO.Excel | 854.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | EPPlus | 395.10 ms | 108.2 MB |  | OfficeIMO.Excel | 930.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables-autofit | OfficeIMO.Excel | 43.41 ms | 10.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-tables-autofit | MiniExcel | 89.68 ms | 125.9 MB |  | OfficeIMO.Excel | 106.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables-autofit | EPPlus | 522.87 ms | 190.8 MB |  | OfficeIMO.Excel | 1104.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables-autofit | ClosedXML | 809.32 ms | 537.2 MB |  | OfficeIMO.Excel | 1764.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | LargeXlsx | 36.05 ms | 9.3 MB |  | LargeXlsx, OfficeIMO.Excel | Tie vs OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | OfficeIMO.Excel | 36.47 ms | 12.4 MB |  | LargeXlsx, OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datatable-direct | MiniExcel | 85.97 ms | 90.2 MB |  | LargeXlsx, OfficeIMO.Excel | 135.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | ClosedXML | 295.14 ms | 101.8 MB |  | LargeXlsx, OfficeIMO.Excel | 709.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | EPPlus 4.5.3.3 | 404.09 ms | 0 B |  | LargeXlsx, OfficeIMO.Excel | 1008.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | EPPlus | 419.23 ms | 114.7 MB |  | LargeXlsx, OfficeIMO.Excel | 1049.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-object-table-direct | OfficeIMO.Excel | 37.27 ms | 9.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datatable-object-table-direct | MiniExcel | 83.85 ms | 87.6 MB |  | OfficeIMO.Excel | 125.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-object-table-direct | EPPlus | 321.94 ms | 112.0 MB |  | OfficeIMO.Excel | 763.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-object-table-direct | ClosedXML | 378.33 ms | 166.7 MB |  | OfficeIMO.Excel | 915.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | OfficeIMO.Excel | 35.86 ms | 12.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datatable-table-direct | MiniExcel | 91.28 ms | 90.2 MB |  | OfficeIMO.Excel | 154.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | EPPlus 4.5.3.3 | 368.88 ms | 0 B |  | OfficeIMO.Excel | 928.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | EPPlus | 396.58 ms | 114.7 MB |  | OfficeIMO.Excel | 1006.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | ClosedXML | 428.25 ms | 169.3 MB |  | OfficeIMO.Excel | 1094.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-dictionary-objects-table-direct | OfficeIMO.Excel | 47.26 ms | 14.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | LargeXlsx | 28.88 ms | 5.5 MB |  | LargeXlsx | 18.1% faster than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 35.25 ms | 12.6 MB |  | LargeXlsx | Loss +22.1% |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | MiniExcel | 67.04 ms | 91.1 MB |  | LargeXlsx | 90.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | ClosedXML | 293.48 ms | 101.8 MB |  | LargeXlsx | 732.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 308.31 ms | 0 B |  | LargeXlsx | 774.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus | 342.29 ms | 114.7 MB |  | LargeXlsx | 871.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 39.51 ms | 11.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 369.72 ms | 156.0 MB |  | OfficeIMO.Excel | 835.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 680.91 ms | 485.3 MB |  | OfficeIMO.Excel | 1623.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | LargeXlsx | 28.36 ms | 5.5 MB |  | LargeXlsx | 19.5% faster than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | OfficeIMO.Excel | 35.22 ms | 11.2 MB |  | LargeXlsx | Loss +24.2% |
| 25000 | speed-comparison | write-insertobjects-direct | MiniExcel | 69.38 ms | 91.1 MB |  | LargeXlsx | 97.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | ClosedXML | 286.17 ms | 101.8 MB |  | LargeXlsx | 712.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | EPPlus 4.5.3.3 | 310.27 ms | 0 B |  | LargeXlsx | 781.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | EPPlus | 351.40 ms | 114.7 MB |  | LargeXlsx | 897.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 42.96 ms | 9.9 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 399.61 ms | 156.0 MB |  | OfficeIMO.Excel | 830.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 736.55 ms | 485.3 MB |  | OfficeIMO.Excel | 1614.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 29.60 ms | 5.5 MB |  | LargeXlsx | 25.3% faster than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 39.65 ms | 9.9 MB |  | LargeXlsx | Loss +34.0% |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | MiniExcel | 68.24 ms | 91.1 MB |  | LargeXlsx | 72.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | ClosedXML | 289.15 ms | 101.8 MB |  | LargeXlsx | 629.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | EPPlus | 349.78 ms | 114.7 MB |  | LargeXlsx | 782.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 37.23 ms | 11.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 359.10 ms | 135.1 MB |  | OfficeIMO.Excel | 864.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 427.09 ms | 269.0 MB |  | OfficeIMO.Excel | 1047.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | LargeXlsx | 41.13 ms | 5.9 MB |  | LargeXlsx | 11.0% faster than OfficeIMO |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 46.21 ms | 10.3 MB |  | LargeXlsx | Loss +12.3% |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | MiniExcel | 89.58 ms | 111.3 MB |  | LargeXlsx | 93.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | ClosedXML | 403.61 ms | 175.3 MB |  | LargeXlsx | 773.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | EPPlus | 493.62 ms | 141.5 MB |  | LargeXlsx | 968.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-mixed-direct | LargeXlsx | 46.42 ms | 5.9 MB |  | LargeXlsx | 4.7% faster than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 48.71 ms | 9.7 MB |  | LargeXlsx | Loss +4.9% |
| 25000 | speed-comparison | write-powershell-psobject-mixed-direct | MiniExcel | 89.95 ms | 111.3 MB |  | LargeXlsx | 84.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-mixed-direct | ClosedXML | 410.67 ms | 175.3 MB |  | LargeXlsx | 743.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-mixed-direct | EPPlus | 487.10 ms | 141.5 MB |  | LargeXlsx | 900.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 216.40 ms | 35.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-powershell-psobject-wide-direct | LargeXlsx | 238.08 ms | 22.7 MB |  | OfficeIMO.Excel | 10.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-wide-direct | MiniExcel | 368.92 ms | 339.8 MB |  | OfficeIMO.Excel | 70.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-wide-direct | ClosedXML | 1.33 s | 476.0 MB |  | OfficeIMO.Excel | 512.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-wide-direct | EPPlus | 1.71 s | 549.7 MB |  | OfficeIMO.Excel | 690.5% slower than OfficeIMO |
| 300000 | focused-package-profile | write-blog-2023-20-string-columns | LargeXlsx | 636.47 ms | 93.1 MB | 28.6 MB | LargeXlsx | Win |
| 300000 | focused-package-profile | write-blog-2023-20-string-columns | OfficeIMO.Excel | 701.65 ms | 173.4 MB | 26.6 MB | LargeXlsx | 1.10x vs best |
| 300000 | focused-package-profile | write-blog-2023-20-string-columns | MiniExcel | 2.28 s | 2.46 GB | 28.5 MB | LargeXlsx | 3.58x vs best |
| 300000 | focused-package-profile | write-blog-2023-20-string-columns | ClosedXML | 15.89 s | 8.51 GB | 31.0 MB | LargeXlsx | 24.97x vs best |
