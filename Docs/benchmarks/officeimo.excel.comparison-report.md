# OfficeIMO.Excel Benchmark Report

Generated: 2026-05-24T11:44:30.9671195Z
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
| 2500 | dense-helloworld-comparison | read | Other | 0 | 2 | dense-helloworld-read-range vs Sylvan.Data.Excel (1.51x) |
| 2500 | package-profile | package | Package size | 27 | 8 | write-cellvalues-rectangle-direct vs LargeXlsx (1.57x) |
| 2500 | speed-comparison | mutate | AutoFit and mutation | 1 | 0 |  |
| 2500 | speed-comparison | other | Object projection | 2 | 0 |  |
| 2500 | speed-comparison | other | Range and table read | 4 | 0 |  |
| 2500 | speed-comparison | read | Formula write/read | 1 | 0 |  |
| 2500 | speed-comparison | read | Other | 0 | 3 | large-sparse-row-read vs Sylvan.Data.Excel (2.65x) |
| 2500 | speed-comparison | read | Range and table read | 1 | 6 | read-top-range vs Sylvan.Data.Excel (4.23x) |
| 2500 | speed-comparison | read | Streaming read | 0 | 4 | read-top-range-stream-small-chunks vs Sylvan.Data.Excel (3.80x) |
| 2500 | speed-comparison | read | Typed object read | 0 | 2 | read-objects-stream vs Sylvan.Data.Excel (1.29x) |
| 2500 | speed-comparison | write | AutoFit and mutation | 5 | 0 |  |
| 2500 | speed-comparison | write | Cell writer | 7 | 0 |  |
| 2500 | speed-comparison | write | DataSet table export | 3 | 0 |  |
| 2500 | speed-comparison | write | DataTable table export | 4 | 0 |  |
| 2500 | speed-comparison | write | Formatted report write | 1 | 0 |  |
| 2500 | speed-comparison | write | Formula write/read | 1 | 0 |  |
| 2500 | speed-comparison | write | Other | 1 | 1 | write-powershell-mixed-objects-direct vs LargeXlsx (1.08x) |
| 2500 | speed-comparison | write | Plain cell export | 1 | 3 | write-cellvalues-headerless-rectangle-direct vs LargeXlsx (2.54x) |
| 2500 | speed-comparison | write | Plain streaming export | 1 | 1 | write-datareader-plain vs Sylvan.Data.Excel (1.37x) |
| 2500 | speed-comparison | write | Plain string export | 0 | 1 | write-blog-2023-20-string-columns vs LargeXlsx (1.13x) |
| 2500 | speed-comparison | write | Shared string write | 4 | 0 |  |
| 2500 | speed-comparison | write | Typed object export | 0 | 3 | write-insertobjects-flat-dictionaries-direct vs LargeXlsx (1.35x) |
| 10000 | focused-package-profile | package | Package size | 1 | 0 |  |
| 25000 | dense-helloworld-comparison | read | Other | 0 | 2 | dense-helloworld-read-range vs Sylvan.Data.Excel (1.24x) |
| 25000 | package-profile | package | Package size | 24 | 11 | append-plain-rows vs LargeXlsx (1.52x) |
| 25000 | speed-comparison | mutate | AutoFit and mutation | 1 | 0 |  |
| 25000 | speed-comparison | other | Object projection | 2 | 0 |  |
| 25000 | speed-comparison | other | Range and table read | 4 | 0 |  |
| 25000 | speed-comparison | read | Formula write/read | 1 | 0 |  |
| 25000 | speed-comparison | read | Other | 0 | 3 | large-sparse-column-read vs Sylvan.Data.Excel (1.72x) |
| 25000 | speed-comparison | read | Range and table read | 2 | 5 | read-top-range vs Sylvan.Data.Excel (4.42x) |
| 25000 | speed-comparison | read | Streaming read | 1 | 3 | read-top-range-stream vs Sylvan.Data.Excel (4.53x) |
| 25000 | speed-comparison | read | Typed object read | 0 | 2 | read-objects-stream vs Sylvan.Data.Excel (1.13x) |
| 25000 | speed-comparison | write | AutoFit and mutation | 5 | 0 |  |
| 25000 | speed-comparison | write | Cell writer | 7 | 0 |  |
| 25000 | speed-comparison | write | DataSet table export | 3 | 0 |  |
| 25000 | speed-comparison | write | DataTable table export | 3 | 1 | write-datatable-direct vs LargeXlsx (1.07x) |
| 25000 | speed-comparison | write | Formatted report write | 1 | 0 |  |
| 25000 | speed-comparison | write | Formula write/read | 1 | 0 |  |
| 25000 | speed-comparison | write | Other | 1 | 1 | write-powershell-mixed-objects-direct vs LargeXlsx (1.18x) |
| 25000 | speed-comparison | write | Plain cell export | 1 | 3 | append-plain-rows vs LargeXlsx (1.56x) |
| 25000 | speed-comparison | write | Plain streaming export | 1 | 1 | write-datareader-plain vs Sylvan.Data.Excel (1.29x) |
| 25000 | speed-comparison | write | Plain string export | 0 | 1 | write-blog-2023-20-string-columns vs LargeXlsx (1.10x) |
| 25000 | speed-comparison | write | Shared string write | 4 | 0 |  |
| 25000 | speed-comparison | write | Typed object export | 0 | 3 | write-insertobjects-flat-dictionaries-direct vs LargeXlsx (1.32x) |
| 300000 | focused-package-profile | package | Package size | 0 | 1 | write-blog-2023-20-string-columns vs LargeXlsx (1.10x) |

## Rows

| Rows | Artifact | Scenario | Library | Mean | Allocated | Package | Best | Outcome |
| ---: | --- | --- | --- | ---: | ---: | ---: | --- | --- |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | Sylvan.Data.Excel | 6.41 ms | 362.3 KB |  | Sylvan.Data.Excel | 33.9% faster than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | OfficeIMO.Excel | 9.70 ms | 2.4 MB |  | Sylvan.Data.Excel | Loss +51.3% |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | ExcelDataReader | 17.83 ms | 6.7 MB |  | Sylvan.Data.Excel | 83.9% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | MiniExcel | 25.26 ms | 21.0 MB |  | Sylvan.Data.Excel | 160.5% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | Sylvan.Data.Excel | 6.24 ms | 362.3 KB |  | Sylvan.Data.Excel | 31.5% faster than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | OfficeIMO.Excel | 9.12 ms | 2.5 MB |  | Sylvan.Data.Excel | Loss +46.0% |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | ExcelDataReader | 19.66 ms | 6.7 MB |  | Sylvan.Data.Excel | 115.6% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | MiniExcel | 25.47 ms | 21.0 MB |  | Sylvan.Data.Excel | 179.4% slower than OfficeIMO |
| 2500 | package-profile | append-plain-rows | LargeXlsx | 2.19 ms | 296.4 KB | 63.1 KB | LargeXlsx | 30.9% faster than OfficeIMO |
| 2500 | package-profile | append-plain-rows | OfficeIMO.Excel | 3.17 ms | 1.6 MB | 64.5 KB | LargeXlsx | Loss +44.7% |
| 2500 | package-profile | append-plain-rows | MiniExcel | 6.53 ms | 19.2 MB | 68.1 KB | LargeXlsx | 106.4% slower than OfficeIMO |
| 2500 | package-profile | append-plain-rows | ClosedXML | 23.06 ms | 10.9 MB | 59.8 KB | LargeXlsx | 628.5% slower than OfficeIMO |
| 2500 | package-profile | append-plain-rows | EPPlus | 38.87 ms | 14.0 MB | 56.9 KB | LargeXlsx | 1127.9% slower than OfficeIMO |
| 2500 | package-profile | autofit-existing | OfficeIMO.Excel | 43.26 ms | 13.6 MB | 138.9 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | autofit-existing | EPPlus | 96.11 ms | 49.5 MB | 115.0 KB | OfficeIMO.Excel | 122.2% slower than OfficeIMO |
| 2500 | package-profile | autofit-existing | ClosedXML | 186.71 ms | 82.6 MB | 121.0 KB | OfficeIMO.Excel | 331.6% slower than OfficeIMO |
| 2500 | package-profile | large-shared-strings | OfficeIMO.Excel | 3.14 ms | 2.1 MB | 55.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | large-shared-strings | MiniExcel | 6.06 ms | 20.6 MB | 60.7 KB | OfficeIMO.Excel | 92.8% slower than OfficeIMO |
| 2500 | package-profile | large-shared-strings | ClosedXML | 19.81 ms | 11.0 MB | 50.3 KB | OfficeIMO.Excel | 530.5% slower than OfficeIMO |
| 2500 | package-profile | large-shared-strings | EPPlus | 27.62 ms | 12.5 MB | 48.1 KB | OfficeIMO.Excel | 779.0% slower than OfficeIMO |
| 2500 | package-profile | write-blog-2023-20-string-columns | LargeXlsx | 5.41 ms | 849.6 KB | 237.7 KB | LargeXlsx | 17.2% faster than OfficeIMO |
| 2500 | package-profile | write-blog-2023-20-string-columns | OfficeIMO.Excel | 6.54 ms | 2.0 MB | 216.7 KB | LargeXlsx | Loss +20.8% |
| 2500 | package-profile | write-blog-2023-20-string-columns | MiniExcel | 25.19 ms | 35.1 MB | 235.3 KB | LargeXlsx | 285.5% slower than OfficeIMO |
| 2500 | package-profile | write-blog-2023-20-string-columns | ClosedXML | 124.41 ms | 69.8 MB | 257.2 KB | LargeXlsx | 1803.7% slower than OfficeIMO |
| 2500 | package-profile | write-bulk-report | OfficeIMO.Excel | 6.85 ms | 1.5 MB | 138.9 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-bulk-report | MiniExcel | 10.85 ms | 26.2 MB | 153.8 KB | OfficeIMO.Excel | 58.4% slower than OfficeIMO |
| 2500 | package-profile | write-bulk-report | EPPlus | 70.33 ms | 46.0 MB | 115.0 KB | OfficeIMO.Excel | 926.5% slower than OfficeIMO |
| 2500 | package-profile | write-bulk-report | ClosedXML | 98.93 ms | 57.0 MB | 121.0 KB | OfficeIMO.Excel | 1344.1% slower than OfficeIMO |
| 2500 | package-profile | write-cellformula | OfficeIMO.Excel | 4.43 ms | 1.1 MB | 66.6 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellformula | ClosedXML | 30.50 ms | 11.8 MB | 70.6 KB | OfficeIMO.Excel | 588.4% slower than OfficeIMO |
| 2500 | package-profile | write-cellformula | EPPlus | 64.98 ms | 17.7 MB | 62.1 KB | OfficeIMO.Excel | 1366.6% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-empty-strings | OfficeIMO.Excel | 3.02 ms | 1.4 MB | 44.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-empty-strings | ClosedXML | 16.30 ms | 9.7 MB | 44.9 KB | OfficeIMO.Excel | 440.2% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-empty-strings | EPPlus | 22.71 ms | 11.4 MB | 42.0 KB | OfficeIMO.Excel | 652.5% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-numbers | OfficeIMO.Excel | 2.37 ms | 939.5 KB | 47.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-numbers | ClosedXML | 15.40 ms | 9.0 MB | 45.9 KB | OfficeIMO.Excel | 549.5% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-numbers | EPPlus | 27.83 ms | 12.5 MB | 43.7 KB | OfficeIMO.Excel | 1074.2% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-mixed | OfficeIMO.Excel | 4.60 ms | 1.4 MB | 61.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-object-mixed | ClosedXML | 29.75 ms | 11.6 MB | 59.5 KB | OfficeIMO.Excel | 546.2% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-mixed | EPPlus | 39.47 ms | 15.2 MB | 58.9 KB | OfficeIMO.Excel | 757.3% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-sparse | OfficeIMO.Excel | 5.74 ms | 1.2 MB | 62.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-object-sparse | ClosedXML | 26.35 ms | 11.0 MB | 52.5 KB | OfficeIMO.Excel | 358.9% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-sparse | EPPlus | 49.07 ms | 14.5 MB | 54.2 KB | OfficeIMO.Excel | 754.5% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 4.42 ms | 1.2 MB | 62.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-object-sparse-batch | ClosedXML | 37.10 ms | 11.0 MB | 52.5 KB | OfficeIMO.Excel | 739.7% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-sparse-batch | EPPlus | 67.81 ms | 14.6 MB | 54.2 KB | OfficeIMO.Excel | 1434.9% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-scalars | OfficeIMO.Excel | 2.83 ms | 957.6 KB | 46.3 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-scalars | ClosedXML | 23.96 ms | 8.8 MB | 45.4 KB | OfficeIMO.Excel | 747.4% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-scalars | EPPlus | 29.94 ms | 12.5 MB | 42.4 KB | OfficeIMO.Excel | 959.2% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings | OfficeIMO.Excel | 4.70 ms | 2.2 MB | 55.1 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-strings | ClosedXML | 15.79 ms | 11.0 MB | 50.3 KB | OfficeIMO.Excel | 235.8% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings | EPPlus | 25.02 ms | 12.4 MB | 48.1 KB | OfficeIMO.Excel | 432.3% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings-distinct | OfficeIMO.Excel | 3.01 ms | 2.1 MB | 51.8 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-strings-distinct | ClosedXML | 23.31 ms | 12.8 MB | 61.9 KB | OfficeIMO.Excel | 675.2% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings-distinct | EPPlus | 30.89 ms | 13.5 MB | 61.5 KB | OfficeIMO.Excel | 927.3% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings-repeated | OfficeIMO.Excel | 2.40 ms | 1.2 MB | 40.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-strings-repeated | ClosedXML | 14.94 ms | 9.0 MB | 38.8 KB | OfficeIMO.Excel | 522.0% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings-repeated | EPPlus | 18.98 ms | 11.0 MB | 34.8 KB | OfficeIMO.Excel | 690.0% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-temporal | OfficeIMO.Excel | 5.14 ms | 1.2 MB | 63.3 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-temporal | ClosedXML | 29.19 ms | 9.5 MB | 54.5 KB | OfficeIMO.Excel | 467.4% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-temporal | EPPlus | 38.17 ms | 14.3 MB | 53.1 KB | OfficeIMO.Excel | 641.9% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 1.81 ms | 439.0 KB | 47.3 KB | LargeXlsx | 11.6% faster than OfficeIMO |
| 2500 | package-profile | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 2.04 ms | 1.1 MB | 48.2 KB | LargeXlsx | Loss +13.1% |
| 2500 | package-profile | write-cellvalues-headerless-rectangle-direct | ClosedXML | 17.38 ms | 10.0 MB | 53.0 KB | LargeXlsx | 750.3% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-headerless-rectangle-direct | EPPlus | 27.36 ms | 12.7 MB | 52.5 KB | LargeXlsx | 1238.4% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | LargeXlsx | 3.12 ms | 750.2 KB | 138.4 KB | LargeXlsx | 36.3% faster than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 4.90 ms | 1.7 MB | 138.0 KB | LargeXlsx | Loss +56.9% |
| 2500 | package-profile | write-cellvalues-rectangle-direct | MiniExcel | 9.49 ms | 22.7 MB | 153.7 KB | LargeXlsx | 93.7% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | ClosedXML | 31.22 ms | 21.7 MB | 120.1 KB | LargeXlsx | 536.9% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | EPPlus | 39.70 ms | 24.0 MB | 114.1 KB | LargeXlsx | 709.9% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | Sylvan.Data.Excel | 3.50 ms | 750.7 KB | 78.5 KB | Sylvan.Data.Excel | 12.5% faster than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | OfficeIMO.Excel | 4.00 ms | 1.4 MB | 138.0 KB | Sylvan.Data.Excel | Loss +14.3% |
| 2500 | package-profile | write-datareader-plain | LargeXlsx | 4.00 ms | 1.0 MB | 138.4 KB | Sylvan.Data.Excel | Tie vs OfficeIMO |
| 2500 | package-profile | write-datareader-plain | MiniExcel | 7.69 ms | 22.5 MB | 153.6 KB | Sylvan.Data.Excel | 92.2% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | ClosedXML | 27.28 ms | 11.3 MB | 120.1 KB | Sylvan.Data.Excel | 582.3% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | EPPlus | 41.36 ms | 16.2 MB | 114.9 KB | Sylvan.Data.Excel | 934.6% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table | OfficeIMO.Excel | 4.24 ms | 1.4 MB | 138.8 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-datareader-table | MiniExcel | 7.65 ms | 22.5 MB | 153.6 KB | OfficeIMO.Excel | 80.3% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table | ClosedXML | 36.52 ms | 18.6 MB | 120.9 KB | OfficeIMO.Excel | 761.2% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table | EPPlus | 38.20 ms | 16.2 MB | 114.9 KB | OfficeIMO.Excel | 800.9% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table-autofit | OfficeIMO.Excel | 4.37 ms | 1.4 MB | 138.9 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-datareader-table-autofit | MiniExcel | 7.87 ms | 26.0 MB | 153.8 KB | OfficeIMO.Excel | 80.3% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table-autofit | EPPlus | 54.69 ms | 37.4 MB | 115.1 KB | OfficeIMO.Excel | 1152.8% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table-autofit | ClosedXML | 77.71 ms | 57.0 MB | 121.0 KB | OfficeIMO.Excel | 1680.2% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-direct-export | OfficeIMO.Excel | 5.18 ms | 1.6 MB | 131.1 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-dataset-sparse-direct-export | LargeXlsx | 5.40 ms | 1.1 MB | 164.2 KB | OfficeIMO.Excel | 4.2% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-direct-export | MiniExcel | 12.76 ms | 29.0 MB | 180.5 KB | OfficeIMO.Excel | 146.1% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-direct-export | EPPlus | 63.95 ms | 21.3 MB | 144.5 KB | OfficeIMO.Excel | 1133.4% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-direct-export | ClosedXML | 71.82 ms | 26.8 MB | 159.4 KB | OfficeIMO.Excel | 1285.2% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-tables | OfficeIMO.Excel | 5.80 ms | 2.3 MB | 176.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-dataset-sparse-tables | MiniExcel | 13.36 ms | 29.0 MB | 180.5 KB | OfficeIMO.Excel | 130.3% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-tables | EPPlus | 67.15 ms | 21.3 MB | 144.5 KB | OfficeIMO.Excel | 1057.5% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-tables | ClosedXML | 71.30 ms | 26.8 MB | 159.4 KB | OfficeIMO.Excel | 1129.0% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables | OfficeIMO.Excel | 5.77 ms | 1.5 MB | 138.9 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-dataset-tables | MiniExcel | 10.01 ms | 28.0 MB | 156.4 KB | OfficeIMO.Excel | 73.6% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables | ClosedXML | 43.43 ms | 18.4 MB | 123.4 KB | OfficeIMO.Excel | 653.2% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables | EPPlus | 44.07 ms | 18.2 MB | 116.6 KB | OfficeIMO.Excel | 664.3% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables-autofit | OfficeIMO.Excel | 5.70 ms | 1.5 MB | 139.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-dataset-tables-autofit | MiniExcel | 11.97 ms | 31.0 MB | 156.6 KB | OfficeIMO.Excel | 109.9% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables-autofit | EPPlus | 76.12 ms | 40.4 MB | 116.9 KB | OfficeIMO.Excel | 1235.1% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables-autofit | ClosedXML | 95.13 ms | 55.4 MB | 123.7 KB | OfficeIMO.Excel | 1568.6% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | OfficeIMO.Excel | 4.01 ms | 1.4 MB | 138.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-datatable-direct | LargeXlsx | 4.14 ms | 1.1 MB | 138.4 KB | OfficeIMO.Excel | 3.5% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | MiniExcel | 8.63 ms | 22.5 MB | 153.7 KB | OfficeIMO.Excel | 115.4% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | ClosedXML | 30.31 ms | 11.3 MB | 120.1 KB | OfficeIMO.Excel | 656.8% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | EPPlus | 40.48 ms | 16.2 MB | 114.9 KB | OfficeIMO.Excel | 910.6% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-table-direct | OfficeIMO.Excel | 3.88 ms | 1.4 MB | 138.8 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-datatable-table-direct | MiniExcel | 7.80 ms | 22.5 MB | 153.7 KB | OfficeIMO.Excel | 101.1% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-table-direct | ClosedXML | 36.21 ms | 18.6 MB | 120.9 KB | OfficeIMO.Excel | 833.6% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-table-direct | EPPlus | 36.41 ms | 16.2 MB | 114.9 KB | OfficeIMO.Excel | 838.7% slower than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 5.36 ms | 1.4 MB | 138.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-fluent-rowsfrom-direct | LargeXlsx | 12.29 ms | 758.3 KB | 138.4 KB | OfficeIMO.Excel | 129.4% slower than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | MiniExcel | 12.79 ms | 22.7 MB | 153.7 KB | OfficeIMO.Excel | 138.7% slower than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | ClosedXML | 53.29 ms | 11.3 MB | 120.1 KB | OfficeIMO.Excel | 894.9% slower than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | EPPlus | 82.54 ms | 16.3 MB | 114.9 KB | OfficeIMO.Excel | 1441.0% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 8.61 ms | 1.4 MB | 138.1 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 128.74 ms | 37.4 MB | 115.1 KB | OfficeIMO.Excel | 1395.2% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 165.08 ms | 49.7 MB | 120.2 KB | OfficeIMO.Excel | 1817.2% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | LargeXlsx | 7.27 ms | 758.3 KB | 138.4 KB | LargeXlsx | 4.3% faster than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | OfficeIMO.Excel | 7.60 ms | 1.4 MB | 138.0 KB | LargeXlsx | Loss +4.5% |
| 2500 | package-profile | write-insertobjects-direct | MiniExcel | 22.07 ms | 22.7 MB | 153.7 KB | LargeXlsx | 190.4% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | ClosedXML | 46.94 ms | 11.3 MB | 120.1 KB | LargeXlsx | 517.7% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | EPPlus | 68.39 ms | 16.3 MB | 114.9 KB | LargeXlsx | 799.9% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 7.17 ms | 749.9 KB | 142.4 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 122.54 ms | 37.4 MB | 115.1 KB | OfficeIMO.Excel | 1608.7% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 143.73 ms | 49.7 MB | 120.2 KB | OfficeIMO.Excel | 1904.2% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 5.01 ms | 758.3 KB | 138.4 KB | LargeXlsx | 6.4% faster than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 5.36 ms | 742.1 KB | 142.3 KB | LargeXlsx | Loss +6.9% |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | MiniExcel | 11.45 ms | 22.7 MB | 153.7 KB | LargeXlsx | 113.7% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | ClosedXML | 40.06 ms | 11.3 MB | 120.1 KB | LargeXlsx | 647.9% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | EPPlus | 49.66 ms | 16.3 MB | 114.9 KB | LargeXlsx | 827.1% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | OfficeIMO.Excel | 7.21 ms | 1.6 MB | 142.3 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | LargeXlsx | 7.68 ms | 758.3 KB | 138.4 KB | OfficeIMO.Excel | 6.5% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | MiniExcel | 12.53 ms | 22.7 MB | 153.7 KB | OfficeIMO.Excel | 73.6% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | ClosedXML | 42.51 ms | 11.3 MB | 120.1 KB | OfficeIMO.Excel | 489.3% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | EPPlus | 53.99 ms | 16.3 MB | 114.9 KB | OfficeIMO.Excel | 648.5% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 6.90 ms | 1.4 MB | 138.1 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 66.54 ms | 27.9 MB | 120.2 KB | OfficeIMO.Excel | 864.0% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 73.83 ms | 26.7 MB | 115.0 KB | OfficeIMO.Excel | 969.7% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-mixed-objects-direct | LargeXlsx | 6.13 ms | 802.5 KB | 182.6 KB | LargeXlsx | 5.1% faster than OfficeIMO |
| 2500 | package-profile | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 6.46 ms | 1.1 MB | 182.4 KB | LargeXlsx | Loss +5.4% |
| 2500 | package-profile | write-powershell-mixed-objects-direct | MiniExcel | 12.70 ms | 24.6 MB | 194.0 KB | LargeXlsx | 96.7% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-mixed-objects-direct | ClosedXML | 53.74 ms | 16.6 MB | 161.0 KB | LargeXlsx | 732.0% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-mixed-objects-direct | EPPlus | 63.75 ms | 19.6 MB | 152.1 KB | LargeXlsx | 886.9% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | LargeXlsx | 1.62 ms | 288.4 KB |  | LargeXlsx | 35.8% faster than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | OfficeIMO.Excel | 2.53 ms | 1.6 MB |  | LargeXlsx | Loss +55.7% |
| 2500 | speed-comparison | append-plain-rows | MiniExcel | 4.18 ms | 19.2 MB |  | LargeXlsx | 65.6% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | ClosedXML | 13.91 ms | 10.9 MB |  | LargeXlsx | 450.7% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | EPPlus | 20.81 ms | 13.9 MB |  | LargeXlsx | 723.5% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | EPPlus 4.5.3.3 | 21.10 ms | 0 B |  | LargeXlsx | 735.0% slower than OfficeIMO |
| 2500 | speed-comparison | autofit-existing | OfficeIMO.Excel | 30.04 ms | 13.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | autofit-existing | EPPlus | 71.87 ms | 49.5 MB |  | OfficeIMO.Excel | 139.2% slower than OfficeIMO |
| 2500 | speed-comparison | autofit-existing | ClosedXML | 127.67 ms | 82.8 MB |  | OfficeIMO.Excel | 325.0% slower than OfficeIMO |
| 2500 | speed-comparison | autofit-existing | EPPlus 4.5.3.3 | 136.44 ms | 0 B |  | OfficeIMO.Excel | 354.2% slower than OfficeIMO |
| 2500 | speed-comparison | build-object-datatable-dictionaries | OfficeIMO.Excel | 1.85 ms | 564.2 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | build-object-datatable-typed | OfficeIMO.Excel | 1.25 ms | 856.9 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-cells | OfficeIMO.Excel | 7.24 ms | 2.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-cells | EPPlus | 23.22 ms | 18.4 MB |  | OfficeIMO.Excel | 220.5% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-cells | ClosedXML | 31.92 ms | 19.9 MB |  | OfficeIMO.Excel | 340.7% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-first-column-from-wide-sheet | OfficeIMO.Excel | 4.89 ms | 679.5 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-first-column-from-wide-sheet | EPPlus | 17.16 ms | 11.5 MB |  | OfficeIMO.Excel | 250.7% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-first-column-from-wide-sheet | ClosedXML | 34.43 ms | 18.3 MB |  | OfficeIMO.Excel | 603.4% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-range | OfficeIMO.Excel | 7.12 ms | 2.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-range | EPPlus | 24.22 ms | 18.4 MB |  | OfficeIMO.Excel | 240.2% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-range | ClosedXML | 33.06 ms | 19.9 MB |  | OfficeIMO.Excel | 364.4% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-top-range | OfficeIMO.Excel | 1.68 ms | 433.9 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-top-range | EPPlus | 18.03 ms | 10.9 MB |  | OfficeIMO.Excel | 974.8% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-top-range | ClosedXML | 32.23 ms | 18.3 MB |  | OfficeIMO.Excel | 1821.7% slower than OfficeIMO |
| 2500 | speed-comparison | formula-heavy-read | OfficeIMO.Excel | 3.23 ms | 777.5 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | formula-heavy-read | EPPlus | 12.88 ms | 7.5 MB |  | OfficeIMO.Excel | 299.2% slower than OfficeIMO |
| 2500 | speed-comparison | formula-heavy-read | EPPlus 4.5.3.3 | 16.76 ms | 0 B |  | OfficeIMO.Excel | 419.6% slower than OfficeIMO |
| 2500 | speed-comparison | formula-heavy-read | ClosedXML | 18.37 ms | 8.1 MB |  | OfficeIMO.Excel | 469.4% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | OfficeIMO.Excel | 1.94 ms | 2.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | large-shared-strings | MiniExcel | 3.84 ms | 20.6 MB |  | OfficeIMO.Excel | 98.3% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | ClosedXML | 11.64 ms | 11.0 MB |  | OfficeIMO.Excel | 500.6% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | EPPlus | 17.12 ms | 12.4 MB |  | OfficeIMO.Excel | 783.1% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | EPPlus 4.5.3.3 | 17.82 ms | 0 B |  | OfficeIMO.Excel | 819.5% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | Sylvan.Data.Excel | 1.11 ms | 316.6 KB |  | Sylvan.Data.Excel | 37.8% faster than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | ExcelDataReader | 1.45 ms | 4.0 MB |  | Sylvan.Data.Excel | 18.7% faster than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | OfficeIMO.Excel | 1.78 ms | 248.9 KB |  | Sylvan.Data.Excel | Loss +60.7% |
| 2500 | speed-comparison | large-sparse-column-read | ClosedXML | 3.51 ms | 4.3 MB |  | Sylvan.Data.Excel | 97.0% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | MiniExcel | 10.88 ms | 45.1 MB |  | Sylvan.Data.Excel | 511.3% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | EPPlus 4.5.3.3 | 16.91 ms | 0 B |  | Sylvan.Data.Excel | 849.7% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | EPPlus | 30.45 ms | 42.1 MB |  | Sylvan.Data.Excel | 1610.1% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | Sylvan.Data.Excel | 1.00 ms | 316.6 KB |  | Sylvan.Data.Excel | 62.3% faster than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | ExcelDataReader | 1.44 ms | 4.0 MB |  | Sylvan.Data.Excel | 46.0% faster than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | OfficeIMO.Excel | 2.66 ms | 249.0 KB |  | Sylvan.Data.Excel | Loss +165.5% |
| 2500 | speed-comparison | large-sparse-row-read | ClosedXML | 3.18 ms | 4.3 MB |  | Sylvan.Data.Excel | 19.3% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | MiniExcel | 10.94 ms | 45.1 MB |  | Sylvan.Data.Excel | 310.9% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | EPPlus 4.5.3.3 | 16.00 ms | 0 B |  | Sylvan.Data.Excel | 500.7% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | EPPlus | 30.27 ms | 42.1 MB |  | Sylvan.Data.Excel | 1036.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | Sylvan.Data.Excel | 4.17 ms | 370.6 KB |  | Sylvan.Data.Excel | 9.8% faster than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | OfficeIMO.Excel | 4.62 ms | 516.7 KB |  | Sylvan.Data.Excel | Loss +10.8% |
| 2500 | speed-comparison | read-bottom-range | ExcelDataReader | 10.65 ms | 6.0 MB |  | Sylvan.Data.Excel | 130.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | MiniExcel | 13.08 ms | 18.2 MB |  | Sylvan.Data.Excel | 182.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | EPPlus | 17.59 ms | 10.9 MB |  | Sylvan.Data.Excel | 280.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | ClosedXML | 37.40 ms | 18.3 MB |  | Sylvan.Data.Excel | 708.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | Sylvan.Data.Excel | 4.32 ms | 370.6 KB |  | Sylvan.Data.Excel | 7.5% faster than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | OfficeIMO.Excel | 4.67 ms | 520.1 KB |  | Sylvan.Data.Excel | Loss +8.1% |
| 2500 | speed-comparison | read-bottom-range-stream | ExcelDataReader | 11.11 ms | 6.0 MB |  | Sylvan.Data.Excel | 137.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | MiniExcel | 13.18 ms | 18.2 MB |  | Sylvan.Data.Excel | 182.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | EPPlus | 17.51 ms | 10.9 MB |  | Sylvan.Data.Excel | 275.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | ClosedXML | 37.30 ms | 18.3 MB |  | Sylvan.Data.Excel | 698.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | Sylvan.Data.Excel | 6.79 ms | 1.9 MB |  | Sylvan.Data.Excel | 18.4% faster than OfficeIMO |
| 2500 | speed-comparison | read-datatable | OfficeIMO.Excel | 8.32 ms | 3.5 MB |  | Sylvan.Data.Excel | Loss +22.5% |
| 2500 | speed-comparison | read-datatable | ExcelDataReader | 13.16 ms | 7.6 MB |  | Sylvan.Data.Excel | 58.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | MiniExcel | 16.75 ms | 17.8 MB |  | Sylvan.Data.Excel | 101.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | EPPlus | 28.47 ms | 20.0 MB |  | Sylvan.Data.Excel | 242.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | ClosedXML | 37.78 ms | 21.1 MB |  | Sylvan.Data.Excel | 354.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | EPPlus 4.5.3.3 | 40.68 ms | 0 B |  | Sylvan.Data.Excel | 389.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | Sylvan.Data.Excel | 4.47 ms | 442.8 KB |  | Sylvan.Data.Excel | 16.0% faster than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | OfficeIMO.Excel | 5.32 ms | 699.1 KB |  | Sylvan.Data.Excel | Loss +19.0% |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | MiniExcel | 11.04 ms | 15.4 MB |  | Sylvan.Data.Excel | 107.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | ExcelDataReader | 12.00 ms | 6.0 MB |  | Sylvan.Data.Excel | 125.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | EPPlus | 18.18 ms | 11.5 MB |  | Sylvan.Data.Excel | 241.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | ClosedXML | 33.55 ms | 18.3 MB |  | Sylvan.Data.Excel | 530.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | Sylvan.Data.Excel | 7.42 ms | 610.6 KB |  | Sylvan.Data.Excel | 11.8% faster than OfficeIMO |
| 2500 | speed-comparison | read-objects | OfficeIMO.Excel | 8.41 ms | 2.4 MB |  | Sylvan.Data.Excel | Loss +13.4% |
| 2500 | speed-comparison | read-objects | MiniExcel | 14.39 ms | 18.0 MB |  | Sylvan.Data.Excel | 71.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | ExcelDataReader | 14.50 ms | 6.3 MB |  | Sylvan.Data.Excel | 72.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | EPPlus | 29.62 ms | 19.6 MB |  | Sylvan.Data.Excel | 252.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | ClosedXML | 35.40 ms | 19.8 MB |  | Sylvan.Data.Excel | 320.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | EPPlus 4.5.3.3 | 37.12 ms | 0 B |  | Sylvan.Data.Excel | 341.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | Sylvan.Data.Excel | 5.04 ms | 546.4 KB |  | Sylvan.Data.Excel | 22.2% faster than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | OfficeIMO.Excel | 6.48 ms | 2.4 MB |  | Sylvan.Data.Excel | Loss +28.6% |
| 2500 | speed-comparison | read-objects-stream | ExcelDataReader | 11.96 ms | 6.2 MB |  | Sylvan.Data.Excel | 84.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | MiniExcel | 13.91 ms | 17.9 MB |  | Sylvan.Data.Excel | 114.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | EPPlus | 25.91 ms | 19.6 MB |  | Sylvan.Data.Excel | 299.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | ClosedXML | 34.03 ms | 19.8 MB |  | Sylvan.Data.Excel | 425.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | EPPlus 4.5.3.3 | 35.63 ms | 0 B |  | Sylvan.Data.Excel | 449.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | OfficeIMO.Excel | 12.34 ms | 2.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | read-range | Sylvan.Data.Excel | 14.60 ms | 370.3 KB |  | OfficeIMO.Excel | 18.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | MiniExcel | 20.39 ms | 18.2 MB |  | OfficeIMO.Excel | 65.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | ExcelDataReader | 25.61 ms | 6.0 MB |  | OfficeIMO.Excel | 107.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | EPPlus | 25.87 ms | 18.4 MB |  | OfficeIMO.Excel | 109.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | EPPlus 4.5.3.3 | 39.46 ms | 0 B |  | OfficeIMO.Excel | 219.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | ClosedXML | 77.04 ms | 19.7 MB |  | OfficeIMO.Excel | 524.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | Sylvan.Data.Excel | 5.11 ms | 465.6 KB |  | Sylvan.Data.Excel | 51.8% faster than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | OfficeIMO.Excel | 10.59 ms | 2.7 MB |  | Sylvan.Data.Excel | Loss +107.5% |
| 2500 | speed-comparison | read-range-decimal | ExcelDataReader | 11.65 ms | 6.0 MB |  | Sylvan.Data.Excel | 10.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | MiniExcel | 16.78 ms | 18.2 MB |  | Sylvan.Data.Excel | 58.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | EPPlus | 25.55 ms | 18.4 MB |  | Sylvan.Data.Excel | 141.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | ClosedXML | 34.81 ms | 19.6 MB |  | Sylvan.Data.Excel | 228.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | Sylvan.Data.Excel | 4.63 ms | 370.6 KB |  | Sylvan.Data.Excel | 52.9% faster than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | OfficeIMO.Excel | 9.81 ms | 2.7 MB |  | Sylvan.Data.Excel | Loss +112.1% |
| 2500 | speed-comparison | read-range-stream | ExcelDataReader | 11.27 ms | 6.0 MB |  | Sylvan.Data.Excel | 14.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | MiniExcel | 12.66 ms | 18.2 MB |  | Sylvan.Data.Excel | 29.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | EPPlus | 24.76 ms | 18.4 MB |  | Sylvan.Data.Excel | 152.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | EPPlus 4.5.3.3 | 33.86 ms | 0 B |  | Sylvan.Data.Excel | 245.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | ClosedXML | 34.78 ms | 19.6 MB |  | Sylvan.Data.Excel | 254.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | Sylvan.Data.Excel | 0.43 ms | 367.3 KB |  | Sylvan.Data.Excel | 76.3% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range | MiniExcel | 0.77 ms | 959.8 KB |  | Sylvan.Data.Excel | 57.9% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range | OfficeIMO.Excel | 1.84 ms | 439.5 KB |  | Sylvan.Data.Excel | Loss +322.8% |
| 2500 | speed-comparison | read-top-range | ExcelDataReader | 4.38 ms | 1.9 MB |  | Sylvan.Data.Excel | 138.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | EPPlus | 17.86 ms | 10.9 MB |  | Sylvan.Data.Excel | 872.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | EPPlus 4.5.3.3 | 32.52 ms | 0 B |  | Sylvan.Data.Excel | 1670.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | ClosedXML | 35.24 ms | 18.3 MB |  | Sylvan.Data.Excel | 1817.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | Sylvan.Data.Excel | 0.43 ms | 367.4 KB |  | Sylvan.Data.Excel | 73.1% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | MiniExcel | 0.79 ms | 959.8 KB |  | Sylvan.Data.Excel | 51.2% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | OfficeIMO.Excel | 1.61 ms | 442.9 KB |  | Sylvan.Data.Excel | Loss +271.8% |
| 2500 | speed-comparison | read-top-range-stream | ExcelDataReader | 4.58 ms | 1.9 MB |  | Sylvan.Data.Excel | 183.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | EPPlus | 17.20 ms | 10.9 MB |  | Sylvan.Data.Excel | 965.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | EPPlus 4.5.3.3 | 28.47 ms | 0 B |  | Sylvan.Data.Excel | 1663.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | ClosedXML | 34.78 ms | 18.3 MB |  | Sylvan.Data.Excel | 2054.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | Sylvan.Data.Excel | 0.41 ms | 367.4 KB |  | Sylvan.Data.Excel | 73.7% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | MiniExcel | 0.80 ms | 959.8 KB |  | Sylvan.Data.Excel | 48.9% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | OfficeIMO.Excel | 1.57 ms | 443.7 KB |  | Sylvan.Data.Excel | Loss +279.7% |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | ExcelDataReader | 4.28 ms | 1.9 MB |  | Sylvan.Data.Excel | 172.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | EPPlus | 18.36 ms | 10.9 MB |  | Sylvan.Data.Excel | 1068.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | ClosedXML | 32.50 ms | 18.3 MB |  | Sylvan.Data.Excel | 1967.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-used-range | Sylvan.Data.Excel | 7.73 ms | 370.6 KB |  | Sylvan.Data.Excel | 49.9% faster than OfficeIMO |
| 2500 | speed-comparison | read-used-range | ExcelDataReader | 11.68 ms | 6.0 MB |  | Sylvan.Data.Excel | 24.3% faster than OfficeIMO |
| 2500 | speed-comparison | read-used-range | OfficeIMO.Excel | 15.44 ms | 3.3 MB |  | Sylvan.Data.Excel | Loss +99.7% |
| 2500 | speed-comparison | read-used-range | MiniExcel | 20.96 ms | 18.2 MB |  | Sylvan.Data.Excel | 35.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-used-range | EPPlus | 27.86 ms | 18.4 MB |  | Sylvan.Data.Excel | 80.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-used-range | ClosedXML | 83.58 ms | 19.7 MB |  | Sylvan.Data.Excel | 441.4% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | Sylvan.Data.Excel | 2.04 ms | 518.6 KB |  | Sylvan.Data.Excel | 38.2% faster than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | OfficeIMO.Excel | 3.31 ms | 1.1 MB |  | Sylvan.Data.Excel | Loss +61.8% |
| 2500 | speed-comparison | shared-string-read | ExcelDataReader | 4.90 ms | 2.5 MB |  | Sylvan.Data.Excel | 48.0% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | MiniExcel | 5.27 ms | 7.3 MB |  | Sylvan.Data.Excel | 59.2% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | EPPlus | 14.32 ms | 10.1 MB |  | Sylvan.Data.Excel | 332.8% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | ClosedXML | 16.75 ms | 9.3 MB |  | Sylvan.Data.Excel | 406.3% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | EPPlus 4.5.3.3 | 21.81 ms | 0 B |  | Sylvan.Data.Excel | 559.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-blog-2023-20-string-columns | LargeXlsx | 6.72 ms | 849.6 KB |  | LargeXlsx | 11.7% faster than OfficeIMO |
| 2500 | speed-comparison | write-blog-2023-20-string-columns | OfficeIMO.Excel | 7.61 ms | 2.0 MB |  | LargeXlsx | Loss +13.2% |
| 2500 | speed-comparison | write-blog-2023-20-string-columns | MiniExcel | 23.39 ms | 35.1 MB |  | LargeXlsx | 207.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-blog-2023-20-string-columns | ClosedXML | 117.38 ms | 69.8 MB |  | LargeXlsx | 1441.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | OfficeIMO.Excel | 5.87 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-bulk-report | MiniExcel | 16.04 ms | 26.2 MB |  | OfficeIMO.Excel | 173.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | EPPlus 4.5.3.3 | 103.42 ms | 0 B |  | OfficeIMO.Excel | 1661.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | EPPlus | 126.49 ms | 47.9 MB |  | OfficeIMO.Excel | 2054.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | ClosedXML | 218.79 ms | 57.0 MB |  | OfficeIMO.Excel | 3627.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellformula | OfficeIMO.Excel | 3.98 ms | 1.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellformula | EPPlus 4.5.3.3 | 23.31 ms | 0 B |  | OfficeIMO.Excel | 485.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellformula | ClosedXML | 27.80 ms | 11.7 MB |  | OfficeIMO.Excel | 598.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellformula | EPPlus | 49.06 ms | 17.6 MB |  | OfficeIMO.Excel | 1132.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-empty-strings | OfficeIMO.Excel | 3.70 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-empty-strings | ClosedXML | 15.47 ms | 9.7 MB |  | OfficeIMO.Excel | 318.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-empty-strings | EPPlus | 19.34 ms | 11.4 MB |  | OfficeIMO.Excel | 423.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-numbers | OfficeIMO.Excel | 3.43 ms | 939.5 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-numbers | ClosedXML | 16.32 ms | 9.0 MB |  | OfficeIMO.Excel | 375.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-numbers | EPPlus 4.5.3.3 | 17.93 ms | 0 B |  | OfficeIMO.Excel | 422.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-numbers | EPPlus | 25.89 ms | 12.5 MB |  | OfficeIMO.Excel | 654.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-mixed | OfficeIMO.Excel | 3.98 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-object-mixed | EPPlus 4.5.3.3 | 21.09 ms | 0 B |  | OfficeIMO.Excel | 429.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-mixed | ClosedXML | 22.23 ms | 11.6 MB |  | OfficeIMO.Excel | 458.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-mixed | EPPlus | 30.59 ms | 15.2 MB |  | OfficeIMO.Excel | 668.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-sparse | OfficeIMO.Excel | 3.47 ms | 1.2 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-object-sparse | ClosedXML | 19.16 ms | 11.0 MB |  | OfficeIMO.Excel | 452.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-sparse | EPPlus | 31.69 ms | 14.5 MB |  | OfficeIMO.Excel | 813.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 3.60 ms | 1.2 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-object-sparse-batch | ClosedXML | 21.15 ms | 11.0 MB |  | OfficeIMO.Excel | 488.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-sparse-batch | EPPlus | 31.36 ms | 14.5 MB |  | OfficeIMO.Excel | 772.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-scalars | OfficeIMO.Excel | 2.89 ms | 957.6 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-scalars | ClosedXML | 16.03 ms | 8.8 MB |  | OfficeIMO.Excel | 455.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-scalars | EPPlus 4.5.3.3 | 16.64 ms | 0 B |  | OfficeIMO.Excel | 476.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-scalars | EPPlus | 26.96 ms | 12.5 MB |  | OfficeIMO.Excel | 834.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings | OfficeIMO.Excel | 3.88 ms | 2.2 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-strings | ClosedXML | 17.35 ms | 11.0 MB |  | OfficeIMO.Excel | 346.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings | EPPlus 4.5.3.3 | 17.63 ms | 0 B |  | OfficeIMO.Excel | 354.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings | EPPlus | 22.25 ms | 12.4 MB |  | OfficeIMO.Excel | 473.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings-distinct | OfficeIMO.Excel | 2.83 ms | 2.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-strings-distinct | ClosedXML | 18.93 ms | 12.8 MB |  | OfficeIMO.Excel | 569.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings-distinct | EPPlus | 23.95 ms | 13.5 MB |  | OfficeIMO.Excel | 747.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings-repeated | OfficeIMO.Excel | 2.89 ms | 1.2 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-strings-repeated | ClosedXML | 14.53 ms | 9.0 MB |  | OfficeIMO.Excel | 402.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings-repeated | EPPlus | 20.54 ms | 11.0 MB |  | OfficeIMO.Excel | 610.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-temporal | OfficeIMO.Excel | 3.74 ms | 1.2 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-temporal | EPPlus 4.5.3.3 | 19.78 ms | 0 B |  | OfficeIMO.Excel | 428.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-temporal | ClosedXML | 22.72 ms | 9.5 MB |  | OfficeIMO.Excel | 506.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-temporal | EPPlus | 28.06 ms | 14.3 MB |  | OfficeIMO.Excel | 649.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 1.83 ms | 439.0 KB |  | LargeXlsx | 60.6% faster than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 4.64 ms | 1.1 MB |  | LargeXlsx | Loss +153.8% |
| 2500 | speed-comparison | write-cellvalues-headerless-rectangle-direct | ClosedXML | 17.16 ms | 10.0 MB |  | LargeXlsx | 270.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-headerless-rectangle-direct | EPPlus | 27.88 ms | 12.7 MB |  | LargeXlsx | 501.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | LargeXlsx | 6.11 ms | 750.2 KB |  | LargeXlsx | 3.5% faster than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 6.33 ms | 1.7 MB |  | LargeXlsx | Loss +3.6% |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | MiniExcel | 12.78 ms | 22.7 MB |  | LargeXlsx | 102.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 41.29 ms | 0 B |  | LargeXlsx | 552.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | ClosedXML | 45.37 ms | 21.7 MB |  | LargeXlsx | 617.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus | 54.31 ms | 24.0 MB |  | LargeXlsx | 758.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-sparse-rectangle-direct | OfficeIMO.Excel | 3.58 ms | 1.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalues-sparse-rectangle-direct | ClosedXML | 21.00 ms | 11.0 MB |  | OfficeIMO.Excel | 487.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-sparse-rectangle-direct | EPPlus | 29.21 ms | 14.5 MB |  | OfficeIMO.Excel | 716.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | Sylvan.Data.Excel | 4.70 ms | 750.5 KB |  | Sylvan.Data.Excel | 27.0% faster than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | OfficeIMO.Excel | 6.44 ms | 1.4 MB |  | Sylvan.Data.Excel | Loss +37.0% |
| 2500 | speed-comparison | write-datareader-plain | MiniExcel | 11.30 ms | 22.5 MB |  | Sylvan.Data.Excel | 75.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | LargeXlsx | 11.46 ms | 1.0 MB |  | Sylvan.Data.Excel | 78.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | EPPlus 4.5.3.3 | 37.95 ms | 0 B |  | Sylvan.Data.Excel | 489.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | ClosedXML | 41.06 ms | 11.3 MB |  | Sylvan.Data.Excel | 538.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | EPPlus | 49.16 ms | 16.2 MB |  | Sylvan.Data.Excel | 663.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | OfficeIMO.Excel | 5.97 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datareader-table | MiniExcel | 12.52 ms | 22.5 MB |  | OfficeIMO.Excel | 109.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | EPPlus 4.5.3.3 | 34.35 ms | 0 B |  | OfficeIMO.Excel | 475.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | ClosedXML | 47.05 ms | 18.6 MB |  | OfficeIMO.Excel | 688.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | EPPlus | 49.34 ms | 16.2 MB |  | OfficeIMO.Excel | 726.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table-autofit | OfficeIMO.Excel | 6.85 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datareader-table-autofit | MiniExcel | 11.95 ms | 26.0 MB |  | OfficeIMO.Excel | 74.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table-autofit | EPPlus | 71.54 ms | 37.4 MB |  | OfficeIMO.Excel | 944.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table-autofit | ClosedXML | 103.45 ms | 57.0 MB |  | OfficeIMO.Excel | 1409.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table-autofit | EPPlus 4.5.3.3 | 117.67 ms | 0 B |  | OfficeIMO.Excel | 1617.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-headerless-tables | OfficeIMO.Excel | 6.68 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-headerless-tables | MiniExcel | 22.39 ms | 28.5 MB |  | OfficeIMO.Excel | 235.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-headerless-tables | ClosedXML | 66.77 ms | 18.5 MB |  | OfficeIMO.Excel | 898.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-headerless-tables | EPPlus | 127.33 ms | 17.9 MB |  | OfficeIMO.Excel | 1804.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | OfficeIMO.Excel | 7.75 ms | 1.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | LargeXlsx | 12.00 ms | 1.1 MB |  | OfficeIMO.Excel | 54.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | MiniExcel | 17.73 ms | 29.3 MB |  | OfficeIMO.Excel | 128.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | EPPlus | 69.02 ms | 21.3 MB |  | OfficeIMO.Excel | 790.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | ClosedXML | 81.29 ms | 26.8 MB |  | OfficeIMO.Excel | 949.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-tables | OfficeIMO.Excel | 8.30 ms | 2.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-sparse-tables | MiniExcel | 23.44 ms | 29.8 MB |  | OfficeIMO.Excel | 182.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-tables | ClosedXML | 79.00 ms | 26.8 MB |  | OfficeIMO.Excel | 852.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-tables | EPPlus | 84.27 ms | 21.8 MB |  | OfficeIMO.Excel | 915.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | OfficeIMO.Excel | 6.48 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-tables | MiniExcel | 16.96 ms | 28.0 MB |  | OfficeIMO.Excel | 162.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | EPPlus 4.5.3.3 | 33.43 ms | 0 B |  | OfficeIMO.Excel | 416.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | ClosedXML | 77.04 ms | 18.4 MB |  | OfficeIMO.Excel | 1089.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | EPPlus | 79.97 ms | 18.9 MB |  | OfficeIMO.Excel | 1134.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables-autofit | OfficeIMO.Excel | 7.21 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-tables-autofit | MiniExcel | 20.69 ms | 31.4 MB |  | OfficeIMO.Excel | 186.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables-autofit | EPPlus | 180.90 ms | 42.4 MB |  | OfficeIMO.Excel | 2408.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables-autofit | ClosedXML | 253.05 ms | 55.4 MB |  | OfficeIMO.Excel | 3408.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | OfficeIMO.Excel | 5.01 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datatable-direct | LargeXlsx | 8.88 ms | 1.1 MB |  | OfficeIMO.Excel | 77.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | MiniExcel | 12.64 ms | 22.5 MB |  | OfficeIMO.Excel | 152.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | EPPlus 4.5.3.3 | 32.64 ms | 0 B |  | OfficeIMO.Excel | 551.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | ClosedXML | 33.45 ms | 11.3 MB |  | OfficeIMO.Excel | 567.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | EPPlus | 41.43 ms | 16.2 MB |  | OfficeIMO.Excel | 726.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-object-table-direct | OfficeIMO.Excel | 5.74 ms | 1.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datatable-object-table-direct | MiniExcel | 10.49 ms | 22.2 MB |  | OfficeIMO.Excel | 82.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-object-table-direct | EPPlus | 42.38 ms | 15.9 MB |  | OfficeIMO.Excel | 638.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-object-table-direct | ClosedXML | 45.32 ms | 18.3 MB |  | OfficeIMO.Excel | 689.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | OfficeIMO.Excel | 5.71 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datatable-table-direct | MiniExcel | 12.63 ms | 22.5 MB |  | OfficeIMO.Excel | 121.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | EPPlus 4.5.3.3 | 34.34 ms | 0 B |  | OfficeIMO.Excel | 501.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | EPPlus | 41.67 ms | 16.2 MB |  | OfficeIMO.Excel | 629.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | ClosedXML | 48.05 ms | 18.6 MB |  | OfficeIMO.Excel | 741.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-dictionary-objects-table-direct | OfficeIMO.Excel | 7.62 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | LargeXlsx | 2.99 ms | 750.2 KB |  | LargeXlsx | 16.4% faster than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 3.58 ms | 1.4 MB |  | LargeXlsx | Loss +19.6% |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | MiniExcel | 7.49 ms | 22.7 MB |  | LargeXlsx | 109.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | ClosedXML | 27.39 ms | 11.3 MB |  | LargeXlsx | 664.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus | 35.04 ms | 16.2 MB |  | LargeXlsx | 878.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 38.25 ms | 0 B |  | LargeXlsx | 968.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 5.75 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 69.22 ms | 37.4 MB |  | OfficeIMO.Excel | 1102.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 90.57 ms | 49.7 MB |  | OfficeIMO.Excel | 1473.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | LargeXlsx | 4.52 ms | 750.2 KB |  | LargeXlsx | 14.7% faster than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | OfficeIMO.Excel | 5.30 ms | 1.4 MB |  | LargeXlsx | Loss +17.2% |
| 2500 | speed-comparison | write-insertobjects-direct | MiniExcel | 13.19 ms | 22.7 MB |  | LargeXlsx | 149.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | EPPlus 4.5.3.3 | 39.78 ms | 0 B |  | LargeXlsx | 651.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | ClosedXML | 40.29 ms | 11.3 MB |  | LargeXlsx | 660.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | EPPlus | 50.40 ms | 16.2 MB |  | LargeXlsx | 851.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.56 ms | 741.8 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 53.46 ms | 37.4 MB |  | OfficeIMO.Excel | 1072.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 63.56 ms | 49.7 MB |  | OfficeIMO.Excel | 1294.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 4.00 ms | 750.2 KB |  | LargeXlsx | 25.8% faster than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 5.39 ms | 734.1 KB |  | LargeXlsx | Loss +34.7% |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | MiniExcel | 9.00 ms | 22.7 MB |  | LargeXlsx | 67.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | ClosedXML | 33.13 ms | 11.3 MB |  | LargeXlsx | 515.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | EPPlus | 40.21 ms | 16.2 MB |  | LargeXlsx | 646.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 5.55 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 57.93 ms | 27.9 MB |  | OfficeIMO.Excel | 943.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 58.10 ms | 26.6 MB |  | OfficeIMO.Excel | 947.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | LargeXlsx | 4.17 ms | 794.5 KB |  | LargeXlsx | 7.4% faster than OfficeIMO |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 4.50 ms | 1.1 MB |  | LargeXlsx | Loss +8.0% |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | MiniExcel | 8.69 ms | 24.6 MB |  | LargeXlsx | 92.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | ClosedXML | 34.65 ms | 16.6 MB |  | LargeXlsx | 669.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | EPPlus | 44.33 ms | 19.6 MB |  | LargeXlsx | 884.5% slower than OfficeIMO |
| 10000 | focused-package-profile | write-blog-2023-20-string-columns | OfficeIMO.Excel | 33.01 ms | 7.6 MB | 880.4 KB | OfficeIMO.Excel | Win |
| 10000 | focused-package-profile | write-blog-2023-20-string-columns | LargeXlsx | 88.91 ms | 3.1 MB | 970.2 KB | OfficeIMO.Excel | 2.69x vs best |
| 10000 | focused-package-profile | write-blog-2023-20-string-columns | MiniExcel | 142.47 ms | 96.2 MB | 957.6 KB | OfficeIMO.Excel | 4.32x vs best |
| 10000 | focused-package-profile | write-blog-2023-20-string-columns | ClosedXML | 725.30 ms | 280.2 MB | 1,015.4 KB | OfficeIMO.Excel | 21.97x vs best |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | Sylvan.Data.Excel | 55.25 ms | 394.1 KB |  | Sylvan.Data.Excel | 19.6% faster than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | OfficeIMO.Excel | 68.71 ms | 23.1 MB |  | Sylvan.Data.Excel | Loss +24.4% |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | ExcelDataReader | 163.40 ms | 67.9 MB |  | Sylvan.Data.Excel | 137.8% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | MiniExcel | 190.66 ms | 210.3 MB |  | Sylvan.Data.Excel | 177.5% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | Sylvan.Data.Excel | 57.91 ms | 394.1 KB |  | Sylvan.Data.Excel | 18.0% faster than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | OfficeIMO.Excel | 70.63 ms | 23.9 MB |  | Sylvan.Data.Excel | Loss +22.0% |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | ExcelDataReader | 158.86 ms | 67.9 MB |  | Sylvan.Data.Excel | 124.9% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | MiniExcel | 214.64 ms | 210.3 MB |  | Sylvan.Data.Excel | 203.9% slower than OfficeIMO |
| 25000 | package-profile | append-plain-rows | LargeXlsx | 16.04 ms | 2.7 MB | 605.0 KB | LargeXlsx | 34.4% faster than OfficeIMO |
| 25000 | package-profile | append-plain-rows | OfficeIMO.Excel | 24.44 ms | 11.4 MB | 622.5 KB | LargeXlsx | Loss +52.4% |
| 25000 | package-profile | append-plain-rows | MiniExcel | 44.05 ms | 56.9 MB | 642.3 KB | LargeXlsx | 80.3% slower than OfficeIMO |
| 25000 | package-profile | append-plain-rows | ClosedXML | 184.98 ms | 101.8 MB | 540.6 KB | LargeXlsx | 656.9% slower than OfficeIMO |
| 25000 | package-profile | append-plain-rows | EPPlus | 201.04 ms | 97.9 MB | 525.6 KB | LargeXlsx | 722.6% slower than OfficeIMO |
| 25000 | package-profile | autofit-existing | OfficeIMO.Excel | 474.61 ms | 132.8 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | autofit-existing | EPPlus | 579.77 ms | 245.0 MB | 1.1 MB | OfficeIMO.Excel | 22.2% slower than OfficeIMO |
| 25000 | package-profile | autofit-existing | ClosedXML | 1.86 s | 810.1 MB | 1.1 MB | OfficeIMO.Excel | 290.9% slower than OfficeIMO |
| 25000 | package-profile | large-shared-strings | OfficeIMO.Excel | 17.93 ms | 15.0 MB | 529.7 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | large-shared-strings | MiniExcel | 37.33 ms | 72.0 MB | 581.0 KB | OfficeIMO.Excel | 108.2% slower than OfficeIMO |
| 25000 | package-profile | large-shared-strings | ClosedXML | 136.10 ms | 101.8 MB | 460.1 KB | OfficeIMO.Excel | 659.0% slower than OfficeIMO |
| 25000 | package-profile | large-shared-strings | EPPlus | 160.97 ms | 82.4 MB | 444.7 KB | OfficeIMO.Excel | 797.6% slower than OfficeIMO |
| 25000 | package-profile | write-blog-2023-20-string-columns | LargeXlsx | 58.30 ms | 10.5 MB | 2.4 MB | LargeXlsx | 6.0% faster than OfficeIMO |
| 25000 | package-profile | write-blog-2023-20-string-columns | OfficeIMO.Excel | 61.99 ms | 12.5 MB | 2.2 MB | LargeXlsx | Loss +6.3% |
| 25000 | package-profile | write-blog-2023-20-string-columns | MiniExcel | 206.70 ms | 221.6 MB | 2.4 MB | LargeXlsx | 233.4% slower than OfficeIMO |
| 25000 | package-profile | write-blog-2023-20-string-columns | ClosedXML | 1.27 s | 742.0 MB | 2.5 MB | LargeXlsx | 1950.2% slower than OfficeIMO |
| 25000 | package-profile | write-bulk-report | OfficeIMO.Excel | 56.75 ms | 12.4 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-bulk-report | MiniExcel | 102.61 ms | 122.6 MB | 1.5 MB | OfficeIMO.Excel | 80.8% slower than OfficeIMO |
| 25000 | package-profile | write-bulk-report | EPPlus | 493.29 ms | 248.9 MB | 1.1 MB | OfficeIMO.Excel | 769.3% slower than OfficeIMO |
| 25000 | package-profile | write-bulk-report | ClosedXML | 1.21 s | 552.7 MB | 1.1 MB | OfficeIMO.Excel | 2039.7% slower than OfficeIMO |
| 25000 | package-profile | write-cellformula | OfficeIMO.Excel | 28.86 ms | 9.3 MB | 670.3 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellformula | ClosedXML | 244.63 ms | 111.2 MB | 643.2 KB | OfficeIMO.Excel | 747.6% slower than OfficeIMO |
| 25000 | package-profile | write-cellformula | EPPlus | 353.87 ms | 137.4 MB | 593.9 KB | OfficeIMO.Excel | 1126.2% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-empty-strings | OfficeIMO.Excel | 17.57 ms | 6.6 MB | 451.4 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-empty-strings | ClosedXML | 153.52 ms | 90.7 MB | 398.1 KB | OfficeIMO.Excel | 774.0% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-empty-strings | EPPlus | 160.59 ms | 72.7 MB | 390.6 KB | OfficeIMO.Excel | 814.2% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-numbers | OfficeIMO.Excel | 22.31 ms | 5.7 MB | 462.6 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-numbers | ClosedXML | 148.09 ms | 82.2 MB | 411.4 KB | OfficeIMO.Excel | 563.9% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-numbers | EPPlus | 198.06 ms | 84.3 MB | 406.5 KB | OfficeIMO.Excel | 787.9% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-mixed | OfficeIMO.Excel | 27.22 ms | 7.8 MB | 585.1 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-object-mixed | EPPlus | 213.77 ms | 110.5 MB | 544.3 KB | OfficeIMO.Excel | 685.5% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-mixed | ClosedXML | 219.55 ms | 108.5 MB | 532.9 KB | OfficeIMO.Excel | 706.7% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-sparse | OfficeIMO.Excel | 29.93 ms | 7.0 MB | 607.1 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-object-sparse | ClosedXML | 193.29 ms | 102.7 MB | 468.0 KB | OfficeIMO.Excel | 545.8% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-sparse | EPPlus | 233.63 ms | 103.8 MB | 494.4 KB | OfficeIMO.Excel | 680.5% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 23.41 ms | 7.0 MB | 607.1 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-object-sparse-batch | ClosedXML | 200.12 ms | 102.7 MB | 468.0 KB | OfficeIMO.Excel | 754.8% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-sparse-batch | EPPlus | 240.93 ms | 103.8 MB | 494.4 KB | OfficeIMO.Excel | 929.2% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-scalars | OfficeIMO.Excel | 16.95 ms | 5.8 MB | 441.9 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-scalars | ClosedXML | 144.48 ms | 80.6 MB | 394.9 KB | OfficeIMO.Excel | 752.6% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-scalars | EPPlus | 186.69 ms | 83.1 MB | 379.3 KB | OfficeIMO.Excel | 1001.6% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings | OfficeIMO.Excel | 23.06 ms | 14.7 MB | 527.8 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-strings | ClosedXML | 152.25 ms | 101.8 MB | 460.1 KB | OfficeIMO.Excel | 560.3% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings | EPPlus | 167.28 ms | 82.4 MB | 444.7 KB | OfficeIMO.Excel | 625.5% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings-distinct | OfficeIMO.Excel | 20.39 ms | 13.3 MB | 499.0 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-strings-distinct | ClosedXML | 204.76 ms | 128.4 MB | 555.3 KB | OfficeIMO.Excel | 904.0% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings-distinct | EPPlus | 206.26 ms | 95.4 MB | 565.1 KB | OfficeIMO.Excel | 911.4% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings-repeated | OfficeIMO.Excel | 18.04 ms | 7.0 MB | 376.0 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-strings-repeated | ClosedXML | 136.24 ms | 82.5 MB | 331.8 KB | OfficeIMO.Excel | 655.0% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings-repeated | EPPlus | 140.99 ms | 68.3 MB | 300.8 KB | OfficeIMO.Excel | 681.4% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-temporal | OfficeIMO.Excel | 30.65 ms | 7.1 MB | 620.5 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-temporal | ClosedXML | 201.16 ms | 87.2 MB | 483.0 KB | OfficeIMO.Excel | 556.4% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-temporal | EPPlus | 220.50 ms | 101.3 MB | 495.1 KB | OfficeIMO.Excel | 619.5% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 13.52 ms | 3.4 MB | 443.4 KB | LargeXlsx | 9.4% faster than OfficeIMO |
| 25000 | package-profile | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 14.93 ms | 7.1 MB | 455.5 KB | LargeXlsx | Loss +10.4% |
| 25000 | package-profile | write-cellvalues-headerless-rectangle-direct | ClosedXML | 156.14 ms | 93.8 MB | 467.5 KB | LargeXlsx | 946.1% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-headerless-rectangle-direct | EPPlus | 176.07 ms | 85.3 MB | 484.1 KB | LargeXlsx | 1079.6% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | LargeXlsx | 43.22 ms | 5.5 MB | 1.4 MB | LargeXlsx | 16.0% faster than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 51.42 ms | 15.3 MB | 1.4 MB | LargeXlsx | Loss +19.0% |
| 25000 | package-profile | write-cellvalues-rectangle-direct | MiniExcel | 94.29 ms | 91.1 MB | 1.5 MB | LargeXlsx | 83.4% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | EPPlus | 421.55 ms | 206.8 MB | 1.1 MB | LargeXlsx | 719.8% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | ClosedXML | 426.81 ms | 205.7 MB | 1.1 MB | LargeXlsx | 730.0% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | Sylvan.Data.Excel | 41.31 ms | 5.6 MB | 755.4 KB | Sylvan.Data.Excel | 21.6% faster than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | LargeXlsx | 52.29 ms | 8.1 MB | 1.4 MB | Sylvan.Data.Excel | Tie vs OfficeIMO |
| 25000 | package-profile | write-datareader-plain | OfficeIMO.Excel | 52.68 ms | 12.4 MB | 1.4 MB | Sylvan.Data.Excel | Loss +27.5% |
| 25000 | package-profile | write-datareader-plain | MiniExcel | 105.90 ms | 90.0 MB | 1.5 MB | Sylvan.Data.Excel | 101.0% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | ClosedXML | 394.06 ms | 101.8 MB | 1.1 MB | Sylvan.Data.Excel | 648.1% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | EPPlus | 409.65 ms | 114.6 MB | 1.1 MB | Sylvan.Data.Excel | 677.6% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table | OfficeIMO.Excel | 54.87 ms | 12.4 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-datareader-table | MiniExcel | 105.42 ms | 90.0 MB | 1.5 MB | OfficeIMO.Excel | 92.1% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table | EPPlus | 410.15 ms | 114.6 MB | 1.1 MB | OfficeIMO.Excel | 647.5% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table | ClosedXML | 551.06 ms | 169.3 MB | 1.1 MB | OfficeIMO.Excel | 904.3% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table-autofit | OfficeIMO.Excel | 58.90 ms | 12.4 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-datareader-table-autofit | MiniExcel | 119.78 ms | 121.6 MB | 1.5 MB | OfficeIMO.Excel | 103.4% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table-autofit | EPPlus | 466.69 ms | 155.9 MB | 1.1 MB | OfficeIMO.Excel | 692.4% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table-autofit | ClosedXML | 1.20 s | 552.9 MB | 1.1 MB | OfficeIMO.Excel | 1944.4% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-direct-export | OfficeIMO.Excel | 51.16 ms | 9.3 MB | 1.3 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-dataset-sparse-direct-export | LargeXlsx | 61.41 ms | 9.0 MB | 1.6 MB | OfficeIMO.Excel | 20.0% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-direct-export | MiniExcel | 145.19 ms | 105.6 MB | 1.8 MB | OfficeIMO.Excel | 183.8% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-direct-export | EPPlus | 672.92 ms | 132.5 MB | 1.4 MB | OfficeIMO.Excel | 1215.2% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-direct-export | ClosedXML | 807.52 ms | 273.8 MB | 1.5 MB | OfficeIMO.Excel | 1478.3% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-tables | OfficeIMO.Excel | 59.60 ms | 12.8 MB | 1.8 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-dataset-sparse-tables | MiniExcel | 143.73 ms | 105.6 MB | 1.8 MB | OfficeIMO.Excel | 141.1% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-tables | EPPlus | 650.81 ms | 132.5 MB | 1.4 MB | OfficeIMO.Excel | 991.9% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-tables | ClosedXML | 811.59 ms | 273.8 MB | 1.5 MB | OfficeIMO.Excel | 1261.6% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables | OfficeIMO.Excel | 55.79 ms | 9.6 MB | 1.3 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-dataset-tables | MiniExcel | 117.54 ms | 94.8 MB | 1.5 MB | OfficeIMO.Excel | 110.7% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables | EPPlus | 382.34 ms | 108.1 MB | 1.1 MB | OfficeIMO.Excel | 585.4% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables | ClosedXML | 523.64 ms | 168.0 MB | 1.1 MB | OfficeIMO.Excel | 838.6% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables-autofit | OfficeIMO.Excel | 56.87 ms | 9.6 MB | 1.3 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-dataset-tables-autofit | MiniExcel | 119.78 ms | 125.8 MB | 1.5 MB | OfficeIMO.Excel | 110.6% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables-autofit | EPPlus | 490.82 ms | 190.7 MB | 1.1 MB | OfficeIMO.Excel | 763.1% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables-autofit | ClosedXML | 1.11 s | 537.2 MB | 1.1 MB | OfficeIMO.Excel | 1845.1% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | LargeXlsx | 46.50 ms | 9.3 MB | 1.4 MB | LargeXlsx | 9.9% faster than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | OfficeIMO.Excel | 51.63 ms | 12.1 MB | 1.4 MB | LargeXlsx | Loss +11.0% |
| 25000 | package-profile | write-datatable-direct | MiniExcel | 113.69 ms | 90.2 MB | 1.5 MB | LargeXlsx | 120.2% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | EPPlus | 393.57 ms | 114.6 MB | 1.1 MB | LargeXlsx | 662.3% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | ClosedXML | 393.82 ms | 101.8 MB | 1.1 MB | LargeXlsx | 662.8% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-table-direct | OfficeIMO.Excel | 54.26 ms | 12.1 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-datatable-table-direct | MiniExcel | 117.75 ms | 90.2 MB | 1.5 MB | OfficeIMO.Excel | 117.0% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-table-direct | EPPlus | 417.49 ms | 114.6 MB | 1.1 MB | OfficeIMO.Excel | 669.5% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-table-direct | ClosedXML | 567.64 ms | 169.3 MB | 1.1 MB | OfficeIMO.Excel | 946.2% slower than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | LargeXlsx | 38.05 ms | 5.5 MB | 1.4 MB | LargeXlsx | 19.8% faster than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 47.42 ms | 12.3 MB | 1.4 MB | LargeXlsx | Loss +24.6% |
| 25000 | package-profile | write-fluent-rowsfrom-direct | MiniExcel | 88.33 ms | 91.1 MB | 1.5 MB | LargeXlsx | 86.3% slower than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | EPPlus | 370.59 ms | 114.6 MB | 1.1 MB | LargeXlsx | 681.5% slower than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | ClosedXML | 372.14 ms | 101.8 MB | 1.1 MB | LargeXlsx | 684.8% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 52.64 ms | 12.3 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 428.17 ms | 155.9 MB | 1.1 MB | OfficeIMO.Excel | 713.4% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 947.32 ms | 485.3 MB | 1.1 MB | OfficeIMO.Excel | 1699.7% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | LargeXlsx | 38.55 ms | 5.5 MB | 1.4 MB | LargeXlsx | 20.0% faster than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | OfficeIMO.Excel | 48.22 ms | 12.3 MB | 1.4 MB | LargeXlsx | Loss +25.1% |
| 25000 | package-profile | write-insertobjects-direct | MiniExcel | 91.32 ms | 91.1 MB | 1.5 MB | LargeXlsx | 89.4% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | EPPlus | 369.82 ms | 114.6 MB | 1.1 MB | LargeXlsx | 667.0% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | ClosedXML | 373.40 ms | 101.8 MB | 1.1 MB | LargeXlsx | 674.4% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 58.29 ms | 6.9 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 418.95 ms | 155.9 MB | 1.1 MB | OfficeIMO.Excel | 618.8% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 954.95 ms | 485.3 MB | 1.1 MB | OfficeIMO.Excel | 1538.4% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 38.55 ms | 5.5 MB | 1.4 MB | LargeXlsx | 25.4% faster than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 51.67 ms | 6.9 MB | 1.4 MB | LargeXlsx | Loss +34.0% |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | MiniExcel | 89.23 ms | 91.1 MB | 1.5 MB | LargeXlsx | 72.7% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | ClosedXML | 366.84 ms | 101.8 MB | 1.1 MB | LargeXlsx | 610.0% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | EPPlus | 370.36 ms | 114.6 MB | 1.1 MB | LargeXlsx | 616.8% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | LargeXlsx | 40.68 ms | 5.5 MB | 1.4 MB | LargeXlsx | 31.5% faster than OfficeIMO |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | OfficeIMO.Excel | 59.39 ms | 15.2 MB | 1.4 MB | LargeXlsx | Loss +46.0% |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | MiniExcel | 94.16 ms | 91.1 MB | 1.5 MB | LargeXlsx | 58.6% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | EPPlus | 379.73 ms | 114.6 MB | 1.1 MB | LargeXlsx | 539.4% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | ClosedXML | 380.01 ms | 101.8 MB | 1.1 MB | LargeXlsx | 539.9% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 49.22 ms | 12.3 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 396.09 ms | 135.0 MB | 1.1 MB | OfficeIMO.Excel | 704.8% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 593.53 ms | 269.0 MB | 1.1 MB | OfficeIMO.Excel | 1105.9% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-mixed-objects-direct | LargeXlsx | 53.76 ms | 5.9 MB | 1.8 MB | LargeXlsx | 11.8% faster than OfficeIMO |
| 25000 | package-profile | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 60.97 ms | 7.1 MB | 1.8 MB | LargeXlsx | Loss +13.4% |
| 25000 | package-profile | write-powershell-mixed-objects-direct | MiniExcel | 115.84 ms | 111.3 MB | 1.9 MB | LargeXlsx | 90.0% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-mixed-objects-direct | ClosedXML | 509.25 ms | 175.3 MB | 1.5 MB | LargeXlsx | 735.3% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-mixed-objects-direct | EPPlus | 524.91 ms | 141.5 MB | 1.4 MB | LargeXlsx | 760.9% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | LargeXlsx | 17.83 ms | 2.7 MB |  | LargeXlsx | 35.8% faster than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | OfficeIMO.Excel | 27.80 ms | 11.4 MB |  | LargeXlsx | Loss +55.9% |
| 25000 | speed-comparison | append-plain-rows | MiniExcel | 53.39 ms | 56.9 MB |  | LargeXlsx | 92.1% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | EPPlus 4.5.3.3 | 131.78 ms | 0 B |  | LargeXlsx | 374.1% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | ClosedXML | 204.47 ms | 101.8 MB |  | LargeXlsx | 635.6% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | EPPlus | 217.72 ms | 98.0 MB |  | LargeXlsx | 683.3% slower than OfficeIMO |
| 25000 | speed-comparison | autofit-existing | OfficeIMO.Excel | 425.06 ms | 132.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | autofit-existing | EPPlus | 498.57 ms | 245.0 MB |  | OfficeIMO.Excel | 17.3% slower than OfficeIMO |
| 25000 | speed-comparison | autofit-existing | EPPlus 4.5.3.3 | 909.70 ms | 0 B |  | OfficeIMO.Excel | 114.0% slower than OfficeIMO |
| 25000 | speed-comparison | autofit-existing | ClosedXML | 1.72 s | 810.1 MB |  | OfficeIMO.Excel | 305.0% slower than OfficeIMO |
| 25000 | speed-comparison | build-object-datatable-dictionaries | OfficeIMO.Excel | 14.84 ms | 5.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | build-object-datatable-typed | OfficeIMO.Excel | 10.41 ms | 7.9 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-cells | OfficeIMO.Excel | 56.30 ms | 22.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-cells | EPPlus | 190.25 ms | 169.8 MB |  | OfficeIMO.Excel | 237.9% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-cells | ClosedXML | 373.06 ms | 194.5 MB |  | OfficeIMO.Excel | 562.7% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-first-column-from-wide-sheet | OfficeIMO.Excel | 144.98 ms | 3.9 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-first-column-from-wide-sheet | EPPlus | 178.72 ms | 99.6 MB |  | OfficeIMO.Excel | 23.3% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-first-column-from-wide-sheet | ClosedXML | 358.72 ms | 179.2 MB |  | OfficeIMO.Excel | 147.4% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-range | OfficeIMO.Excel | 78.13 ms | 22.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-range | EPPlus | 263.45 ms | 169.8 MB |  | OfficeIMO.Excel | 237.2% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-range | ClosedXML | 552.71 ms | 194.5 MB |  | OfficeIMO.Excel | 607.4% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-top-range | OfficeIMO.Excel | 2.65 ms | 445.0 KB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-top-range | EPPlus | 231.31 ms | 89.9 MB |  | OfficeIMO.Excel | 8627.4% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-top-range | ClosedXML | 610.72 ms | 177.7 MB |  | OfficeIMO.Excel | 22942.4% slower than OfficeIMO |
| 25000 | speed-comparison | formula-heavy-read | OfficeIMO.Excel | 21.12 ms | 6.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | formula-heavy-read | EPPlus 4.5.3.3 | 111.90 ms | 0 B |  | OfficeIMO.Excel | 429.9% slower than OfficeIMO |
| 25000 | speed-comparison | formula-heavy-read | EPPlus | 114.41 ms | 69.2 MB |  | OfficeIMO.Excel | 441.8% slower than OfficeIMO |
| 25000 | speed-comparison | formula-heavy-read | ClosedXML | 175.75 ms | 77.6 MB |  | OfficeIMO.Excel | 732.3% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | OfficeIMO.Excel | 17.94 ms | 15.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | large-shared-strings | MiniExcel | 29.31 ms | 72.0 MB |  | OfficeIMO.Excel | 63.4% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | ClosedXML | 112.12 ms | 101.8 MB |  | OfficeIMO.Excel | 525.0% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | EPPlus | 132.55 ms | 82.4 MB |  | OfficeIMO.Excel | 638.9% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | EPPlus 4.5.3.3 | 272.45 ms | 0 B |  | OfficeIMO.Excel | 1418.7% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | Sylvan.Data.Excel | 1.09 ms | 316.6 KB |  | Sylvan.Data.Excel | 42.0% faster than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | ExcelDataReader | 1.72 ms | 4.0 MB |  | Sylvan.Data.Excel | 8.3% faster than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | OfficeIMO.Excel | 1.87 ms | 249.0 KB |  | Sylvan.Data.Excel | Loss +72.4% |
| 25000 | speed-comparison | large-sparse-column-read | ClosedXML | 3.52 ms | 4.3 MB |  | Sylvan.Data.Excel | 87.9% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | MiniExcel | 11.18 ms | 45.1 MB |  | Sylvan.Data.Excel | 496.7% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | EPPlus 4.5.3.3 | 13.45 ms | 0 B |  | Sylvan.Data.Excel | 617.9% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | EPPlus | 58.52 ms | 42.1 MB |  | Sylvan.Data.Excel | 3023.3% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | Sylvan.Data.Excel | 1.03 ms | 316.6 KB |  | Sylvan.Data.Excel | 39.8% faster than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | ExcelDataReader | 1.42 ms | 4.0 MB |  | Sylvan.Data.Excel | 16.9% faster than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | OfficeIMO.Excel | 1.71 ms | 249.1 KB |  | Sylvan.Data.Excel | Loss +66.0% |
| 25000 | speed-comparison | large-sparse-row-read | ClosedXML | 3.39 ms | 4.3 MB |  | Sylvan.Data.Excel | 98.3% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | MiniExcel | 11.12 ms | 45.1 MB |  | Sylvan.Data.Excel | 551.5% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | EPPlus 4.5.3.3 | 15.64 ms | 0 B |  | Sylvan.Data.Excel | 816.4% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | EPPlus | 57.65 ms | 42.1 MB |  | Sylvan.Data.Excel | 3276.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | OfficeIMO.Excel | 35.43 ms | 1.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-bottom-range | Sylvan.Data.Excel | 40.18 ms | 400.4 KB |  | OfficeIMO.Excel | 13.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | ExcelDataReader | 117.72 ms | 60.1 MB |  | OfficeIMO.Excel | 232.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | MiniExcel | 154.33 ms | 180.7 MB |  | OfficeIMO.Excel | 335.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | EPPlus | 174.65 ms | 89.9 MB |  | OfficeIMO.Excel | 393.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | ClosedXML | 349.20 ms | 177.7 MB |  | OfficeIMO.Excel | 885.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | OfficeIMO.Excel | 36.18 ms | 1.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-bottom-range-stream | Sylvan.Data.Excel | 38.91 ms | 400.4 KB |  | OfficeIMO.Excel | 7.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | ExcelDataReader | 115.37 ms | 60.1 MB |  | OfficeIMO.Excel | 218.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | MiniExcel | 128.03 ms | 180.7 MB |  | OfficeIMO.Excel | 253.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | EPPlus | 158.12 ms | 89.9 MB |  | OfficeIMO.Excel | 337.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | ClosedXML | 353.33 ms | 177.7 MB |  | OfficeIMO.Excel | 876.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | Sylvan.Data.Excel | 55.46 ms | 14.9 MB |  | Sylvan.Data.Excel | 8.4% faster than OfficeIMO |
| 25000 | speed-comparison | read-datatable | OfficeIMO.Excel | 60.55 ms | 32.5 MB |  | Sylvan.Data.Excel | Loss +9.2% |
| 25000 | speed-comparison | read-datatable | ExcelDataReader | 134.07 ms | 74.6 MB |  | Sylvan.Data.Excel | 121.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | MiniExcel | 149.64 ms | 175.7 MB |  | Sylvan.Data.Excel | 147.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | EPPlus | 206.42 ms | 184.3 MB |  | Sylvan.Data.Excel | 240.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | EPPlus 4.5.3.3 | 260.12 ms | 0 B |  | Sylvan.Data.Excel | 329.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | ClosedXML | 364.70 ms | 206.1 MB |  | Sylvan.Data.Excel | 502.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | Sylvan.Data.Excel | 44.30 ms | 1.1 MB |  | Sylvan.Data.Excel, OfficeIMO.Excel | Tie vs OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | OfficeIMO.Excel | 45.05 ms | 4.1 MB |  | Sylvan.Data.Excel, OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | MiniExcel | 107.30 ms | 153.6 MB |  | Sylvan.Data.Excel, OfficeIMO.Excel | 138.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | ExcelDataReader | 117.17 ms | 60.1 MB |  | Sylvan.Data.Excel, OfficeIMO.Excel | 160.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | EPPlus | 205.84 ms | 99.6 MB |  | Sylvan.Data.Excel, OfficeIMO.Excel | 356.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | ClosedXML | 411.75 ms | 179.2 MB |  | Sylvan.Data.Excel, OfficeIMO.Excel | 814.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | Sylvan.Data.Excel | 44.32 ms | 2.6 MB |  | Sylvan.Data.Excel | 11.7% faster than OfficeIMO |
| 25000 | speed-comparison | read-objects | OfficeIMO.Excel | 50.18 ms | 21.7 MB |  | Sylvan.Data.Excel | Loss +13.2% |
| 25000 | speed-comparison | read-objects | ExcelDataReader | 113.13 ms | 62.3 MB |  | Sylvan.Data.Excel | 125.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | MiniExcel | 131.88 ms | 178.0 MB |  | Sylvan.Data.Excel | 162.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | EPPlus | 194.35 ms | 181.7 MB |  | Sylvan.Data.Excel | 287.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | EPPlus 4.5.3.3 | 226.02 ms | 0 B |  | Sylvan.Data.Excel | 350.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | ClosedXML | 346.02 ms | 193.5 MB |  | Sylvan.Data.Excel | 589.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | Sylvan.Data.Excel | 63.79 ms | 2.1 MB |  | Sylvan.Data.Excel | 11.7% faster than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | OfficeIMO.Excel | 72.25 ms | 21.7 MB |  | Sylvan.Data.Excel | Loss +13.3% |
| 25000 | speed-comparison | read-objects-stream | ExcelDataReader | 148.07 ms | 61.8 MB |  | Sylvan.Data.Excel | 104.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | MiniExcel | 166.45 ms | 177.5 MB |  | Sylvan.Data.Excel | 130.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | EPPlus 4.5.3.3 | 216.02 ms | 0 B |  | Sylvan.Data.Excel | 199.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | EPPlus | 230.93 ms | 181.5 MB |  | Sylvan.Data.Excel | 219.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | ClosedXML | 447.80 ms | 193.3 MB |  | Sylvan.Data.Excel | 519.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | Sylvan.Data.Excel | 64.85 ms | 400.4 KB |  | Sylvan.Data.Excel | 32.3% faster than OfficeIMO |
| 25000 | speed-comparison | read-range | OfficeIMO.Excel | 95.74 ms | 24.2 MB |  | Sylvan.Data.Excel | Loss +47.6% |
| 25000 | speed-comparison | read-range | ExcelDataReader | 187.10 ms | 60.1 MB |  | Sylvan.Data.Excel | 95.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | MiniExcel | 196.40 ms | 180.7 MB |  | Sylvan.Data.Excel | 105.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | EPPlus 4.5.3.3 | 230.63 ms | 0 B |  | Sylvan.Data.Excel | 140.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | EPPlus | 280.39 ms | 169.8 MB |  | Sylvan.Data.Excel | 192.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | ClosedXML | 513.96 ms | 191.6 MB |  | Sylvan.Data.Excel | 436.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | Sylvan.Data.Excel | 78.91 ms | 1.3 MB |  | Sylvan.Data.Excel | 15.1% faster than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | OfficeIMO.Excel | 92.90 ms | 24.8 MB |  | Sylvan.Data.Excel | Loss +17.7% |
| 25000 | speed-comparison | read-range-decimal | MiniExcel | 192.45 ms | 180.7 MB |  | Sylvan.Data.Excel | 107.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | ExcelDataReader | 192.58 ms | 60.1 MB |  | Sylvan.Data.Excel | 107.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | EPPlus | 287.81 ms | 169.8 MB |  | Sylvan.Data.Excel | 209.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | ClosedXML | 512.92 ms | 191.6 MB |  | Sylvan.Data.Excel | 452.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | Sylvan.Data.Excel | 44.91 ms | 400.4 KB |  | Sylvan.Data.Excel | 15.7% faster than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | OfficeIMO.Excel | 53.27 ms | 25.0 MB |  | Sylvan.Data.Excel | Loss +18.6% |
| 25000 | speed-comparison | read-range-stream | ExcelDataReader | 112.51 ms | 60.1 MB |  | Sylvan.Data.Excel | 111.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | MiniExcel | 125.39 ms | 180.7 MB |  | Sylvan.Data.Excel | 135.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | EPPlus | 205.66 ms | 169.8 MB |  | Sylvan.Data.Excel | 286.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | EPPlus 4.5.3.3 | 218.89 ms | 0 B |  | Sylvan.Data.Excel | 310.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | ClosedXML | 339.06 ms | 191.6 MB |  | Sylvan.Data.Excel | 536.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | Sylvan.Data.Excel | 0.44 ms | 367.3 KB |  | Sylvan.Data.Excel | 77.4% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range | MiniExcel | 1.01 ms | 973.1 KB |  | Sylvan.Data.Excel | 47.6% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range | OfficeIMO.Excel | 1.92 ms | 447.6 KB |  | Sylvan.Data.Excel | Loss +342.4% |
| 25000 | speed-comparison | read-top-range | ExcelDataReader | 38.24 ms | 16.8 MB |  | Sylvan.Data.Excel | 1886.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | EPPlus | 173.93 ms | 89.9 MB |  | Sylvan.Data.Excel | 8935.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | EPPlus 4.5.3.3 | 182.26 ms | 0 B |  | Sylvan.Data.Excel | 9368.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | ClosedXML | 353.32 ms | 177.7 MB |  | Sylvan.Data.Excel | 18254.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | Sylvan.Data.Excel | 0.47 ms | 367.3 KB |  | Sylvan.Data.Excel | 77.9% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | MiniExcel | 0.83 ms | 959.8 KB |  | Sylvan.Data.Excel | 61.0% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | OfficeIMO.Excel | 2.14 ms | 443.0 KB |  | Sylvan.Data.Excel | Loss +353.0% |
| 25000 | speed-comparison | read-top-range-stream | ExcelDataReader | 38.84 ms | 16.8 MB |  | Sylvan.Data.Excel | 1717.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | EPPlus | 155.37 ms | 89.9 MB |  | Sylvan.Data.Excel | 7171.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | EPPlus 4.5.3.3 | 177.92 ms | 0 B |  | Sylvan.Data.Excel | 8226.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | ClosedXML | 355.00 ms | 177.7 MB |  | Sylvan.Data.Excel | 16513.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | Sylvan.Data.Excel | 0.46 ms | 367.3 KB |  | Sylvan.Data.Excel | 77.9% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | MiniExcel | 0.75 ms | 959.8 KB |  | Sylvan.Data.Excel | 64.0% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | OfficeIMO.Excel | 2.08 ms | 443.7 KB |  | Sylvan.Data.Excel | Loss +352.9% |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | ExcelDataReader | 37.65 ms | 16.8 MB |  | Sylvan.Data.Excel | 1712.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | EPPlus | 152.04 ms | 89.9 MB |  | Sylvan.Data.Excel | 7217.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | ClosedXML | 329.59 ms | 177.7 MB |  | Sylvan.Data.Excel | 15762.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-used-range | Sylvan.Data.Excel | 72.76 ms | 400.4 KB |  | Sylvan.Data.Excel | 54.0% faster than OfficeIMO |
| 25000 | speed-comparison | read-used-range | OfficeIMO.Excel | 158.02 ms | 32.0 MB |  | Sylvan.Data.Excel | Loss +117.2% |
| 25000 | speed-comparison | read-used-range | ExcelDataReader | 178.84 ms | 60.1 MB |  | Sylvan.Data.Excel | 13.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-used-range | MiniExcel | 189.95 ms | 180.7 MB |  | Sylvan.Data.Excel | 20.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-used-range | EPPlus | 261.65 ms | 169.8 MB |  | Sylvan.Data.Excel | 65.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-used-range | ClosedXML | 531.27 ms | 191.6 MB |  | Sylvan.Data.Excel | 236.2% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | Sylvan.Data.Excel | 16.81 ms | 1.9 MB |  | Sylvan.Data.Excel | 23.1% faster than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | OfficeIMO.Excel | 21.85 ms | 9.1 MB |  | Sylvan.Data.Excel | Loss +30.0% |
| 25000 | speed-comparison | shared-string-read | ExcelDataReader | 45.09 ms | 24.4 MB |  | Sylvan.Data.Excel | 106.3% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | MiniExcel | 51.19 ms | 72.7 MB |  | Sylvan.Data.Excel | 134.3% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | EPPlus | 97.28 ms | 87.3 MB |  | Sylvan.Data.Excel | 345.2% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | EPPlus 4.5.3.3 | 123.49 ms | 0 B |  | Sylvan.Data.Excel | 465.2% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | ClosedXML | 149.92 ms | 88.3 MB |  | Sylvan.Data.Excel | 586.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-blog-2023-20-string-columns | LargeXlsx | 53.82 ms | 10.5 MB |  | LargeXlsx | 9.5% faster than OfficeIMO |
| 25000 | speed-comparison | write-blog-2023-20-string-columns | OfficeIMO.Excel | 59.47 ms | 12.5 MB |  | LargeXlsx | Loss +10.5% |
| 25000 | speed-comparison | write-blog-2023-20-string-columns | MiniExcel | 201.46 ms | 221.6 MB |  | LargeXlsx | 238.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-blog-2023-20-string-columns | ClosedXML | 1.17 s | 742.0 MB |  | LargeXlsx | 1874.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | OfficeIMO.Excel | 53.23 ms | 12.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-bulk-report | MiniExcel | 96.82 ms | 122.6 MB |  | OfficeIMO.Excel | 81.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | EPPlus | 470.38 ms | 249.0 MB |  | OfficeIMO.Excel | 783.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | EPPlus 4.5.3.3 | 529.04 ms | 0 B |  | OfficeIMO.Excel | 893.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | ClosedXML | 1.09 s | 552.7 MB |  | OfficeIMO.Excel | 1945.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellformula | OfficeIMO.Excel | 33.79 ms | 9.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellformula | EPPlus 4.5.3.3 | 124.66 ms | 0 B |  | OfficeIMO.Excel | 269.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellformula | ClosedXML | 318.84 ms | 111.2 MB |  | OfficeIMO.Excel | 843.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellformula | EPPlus | 499.00 ms | 137.4 MB |  | OfficeIMO.Excel | 1376.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-empty-strings | OfficeIMO.Excel | 17.83 ms | 6.6 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-empty-strings | ClosedXML | 149.02 ms | 90.7 MB |  | OfficeIMO.Excel | 736.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-empty-strings | EPPlus | 149.42 ms | 72.7 MB |  | OfficeIMO.Excel | 738.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-numbers | OfficeIMO.Excel | 21.77 ms | 5.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-numbers | EPPlus 4.5.3.3 | 91.08 ms | 0 B |  | OfficeIMO.Excel | 318.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-numbers | ClosedXML | 149.81 ms | 82.2 MB |  | OfficeIMO.Excel | 588.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-numbers | EPPlus | 177.08 ms | 84.3 MB |  | OfficeIMO.Excel | 713.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-mixed | OfficeIMO.Excel | 29.44 ms | 7.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-object-mixed | EPPlus 4.5.3.3 | 112.75 ms | 0 B |  | OfficeIMO.Excel | 283.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-mixed | EPPlus | 214.18 ms | 110.5 MB |  | OfficeIMO.Excel | 627.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-mixed | ClosedXML | 233.38 ms | 108.5 MB |  | OfficeIMO.Excel | 692.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-sparse | OfficeIMO.Excel | 31.35 ms | 7.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-object-sparse | ClosedXML | 210.09 ms | 102.7 MB |  | OfficeIMO.Excel | 570.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-sparse | EPPlus | 229.40 ms | 103.8 MB |  | OfficeIMO.Excel | 631.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 23.52 ms | 7.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-object-sparse-batch | ClosedXML | 209.63 ms | 102.7 MB |  | OfficeIMO.Excel | 791.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-sparse-batch | EPPlus | 244.44 ms | 103.8 MB |  | OfficeIMO.Excel | 939.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-scalars | OfficeIMO.Excel | 16.68 ms | 5.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-scalars | EPPlus 4.5.3.3 | 92.13 ms | 0 B |  | OfficeIMO.Excel | 452.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-scalars | ClosedXML | 134.47 ms | 80.6 MB |  | OfficeIMO.Excel | 706.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-scalars | EPPlus | 174.97 ms | 83.1 MB |  | OfficeIMO.Excel | 949.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings | OfficeIMO.Excel | 24.55 ms | 14.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-strings | EPPlus 4.5.3.3 | 94.47 ms | 0 B |  | OfficeIMO.Excel | 284.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings | ClosedXML | 140.20 ms | 101.8 MB |  | OfficeIMO.Excel | 471.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings | EPPlus | 146.21 ms | 82.4 MB |  | OfficeIMO.Excel | 495.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings-distinct | OfficeIMO.Excel | 19.01 ms | 13.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-strings-distinct | ClosedXML | 190.38 ms | 128.4 MB |  | OfficeIMO.Excel | 901.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings-distinct | EPPlus | 197.38 ms | 95.4 MB |  | OfficeIMO.Excel | 938.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings-repeated | OfficeIMO.Excel | 18.68 ms | 7.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-strings-repeated | ClosedXML | 124.43 ms | 82.5 MB |  | OfficeIMO.Excel | 566.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings-repeated | EPPlus | 134.78 ms | 68.3 MB |  | OfficeIMO.Excel | 621.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-temporal | OfficeIMO.Excel | 30.31 ms | 7.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-temporal | EPPlus 4.5.3.3 | 106.01 ms | 0 B |  | OfficeIMO.Excel | 249.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-temporal | ClosedXML | 212.00 ms | 87.2 MB |  | OfficeIMO.Excel | 599.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-temporal | EPPlus | 221.41 ms | 101.3 MB |  | OfficeIMO.Excel | 630.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 12.78 ms | 3.4 MB |  | LargeXlsx | 14.5% faster than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 14.95 ms | 7.1 MB |  | LargeXlsx | Loss +16.9% |
| 25000 | speed-comparison | write-cellvalues-headerless-rectangle-direct | ClosedXML | 152.84 ms | 93.8 MB |  | LargeXlsx | 922.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-headerless-rectangle-direct | EPPlus | 166.81 ms | 85.3 MB |  | LargeXlsx | 1016.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | LargeXlsx | 40.23 ms | 5.5 MB |  | LargeXlsx | 20.2% faster than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 50.39 ms | 15.3 MB |  | LargeXlsx | Loss +25.2% |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | MiniExcel | 90.01 ms | 91.1 MB |  | LargeXlsx | 78.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 228.62 ms | 0 B |  | LargeXlsx | 353.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus | 419.29 ms | 206.8 MB |  | LargeXlsx | 732.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | ClosedXML | 422.39 ms | 205.7 MB |  | LargeXlsx | 738.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-sparse-rectangle-direct | OfficeIMO.Excel | 22.22 ms | 7.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalues-sparse-rectangle-direct | ClosedXML | 188.68 ms | 102.7 MB |  | OfficeIMO.Excel | 749.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-sparse-rectangle-direct | EPPlus | 221.98 ms | 103.8 MB |  | OfficeIMO.Excel | 899.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | Sylvan.Data.Excel | 39.84 ms | 5.6 MB |  | Sylvan.Data.Excel | 22.4% faster than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | LargeXlsx | 48.38 ms | 8.1 MB |  | Sylvan.Data.Excel | 5.8% faster than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | OfficeIMO.Excel | 51.35 ms | 12.4 MB |  | Sylvan.Data.Excel | Loss +28.9% |
| 25000 | speed-comparison | write-datareader-plain | MiniExcel | 99.08 ms | 90.0 MB |  | Sylvan.Data.Excel | 93.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | EPPlus 4.5.3.3 | 212.03 ms | 0 B |  | Sylvan.Data.Excel | 312.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | ClosedXML | 375.50 ms | 101.8 MB |  | Sylvan.Data.Excel | 631.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | EPPlus | 382.21 ms | 114.6 MB |  | Sylvan.Data.Excel | 644.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | OfficeIMO.Excel | 49.48 ms | 12.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datareader-table | MiniExcel | 95.70 ms | 90.0 MB |  | OfficeIMO.Excel | 93.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | EPPlus 4.5.3.3 | 217.74 ms | 0 B |  | OfficeIMO.Excel | 340.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | EPPlus | 368.45 ms | 114.6 MB |  | OfficeIMO.Excel | 644.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | ClosedXML | 487.84 ms | 169.3 MB |  | OfficeIMO.Excel | 886.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table-autofit | OfficeIMO.Excel | 49.71 ms | 12.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datareader-table-autofit | MiniExcel | 95.92 ms | 121.6 MB |  | OfficeIMO.Excel | 93.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table-autofit | EPPlus | 380.97 ms | 155.9 MB |  | OfficeIMO.Excel | 666.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table-autofit | EPPlus 4.5.3.3 | 454.89 ms | 0 B |  | OfficeIMO.Excel | 815.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table-autofit | ClosedXML | 1.06 s | 552.9 MB |  | OfficeIMO.Excel | 2039.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-headerless-tables | OfficeIMO.Excel | 48.30 ms | 12.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-headerless-tables | MiniExcel | 108.40 ms | 94.8 MB |  | OfficeIMO.Excel | 124.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-headerless-tables | EPPlus | 480.25 ms | 108.6 MB |  | OfficeIMO.Excel | 894.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-headerless-tables | ClosedXML | 502.48 ms | 168.0 MB |  | OfficeIMO.Excel | 940.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | OfficeIMO.Excel | 42.54 ms | 9.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | LargeXlsx | 49.41 ms | 9.0 MB |  | OfficeIMO.Excel | 16.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | MiniExcel | 125.57 ms | 105.6 MB |  | OfficeIMO.Excel | 195.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | EPPlus | 568.43 ms | 132.5 MB |  | OfficeIMO.Excel | 1236.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | ClosedXML | 760.20 ms | 273.8 MB |  | OfficeIMO.Excel | 1687.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-tables | OfficeIMO.Excel | 53.63 ms | 12.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-sparse-tables | MiniExcel | 126.43 ms | 105.6 MB |  | OfficeIMO.Excel | 135.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-tables | EPPlus | 573.54 ms | 132.5 MB |  | OfficeIMO.Excel | 969.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-tables | ClosedXML | 745.19 ms | 273.8 MB |  | OfficeIMO.Excel | 1289.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | OfficeIMO.Excel | 47.79 ms | 9.6 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-tables | MiniExcel | 98.40 ms | 94.8 MB |  | OfficeIMO.Excel | 105.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | EPPlus 4.5.3.3 | 231.75 ms | 0 B |  | OfficeIMO.Excel | 384.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | EPPlus | 341.38 ms | 108.1 MB |  | OfficeIMO.Excel | 614.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | ClosedXML | 476.93 ms | 168.0 MB |  | OfficeIMO.Excel | 898.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables-autofit | OfficeIMO.Excel | 49.55 ms | 9.6 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-tables-autofit | MiniExcel | 106.52 ms | 125.8 MB |  | OfficeIMO.Excel | 115.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables-autofit | EPPlus | 442.21 ms | 190.7 MB |  | OfficeIMO.Excel | 792.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables-autofit | ClosedXML | 1.04 s | 537.2 MB |  | OfficeIMO.Excel | 2000.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | LargeXlsx | 42.44 ms | 9.3 MB |  | LargeXlsx | 6.2% faster than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | OfficeIMO.Excel | 45.25 ms | 12.1 MB |  | LargeXlsx | Loss +6.6% |
| 25000 | speed-comparison | write-datatable-direct | MiniExcel | 100.03 ms | 90.2 MB |  | LargeXlsx | 121.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | EPPlus 4.5.3.3 | 230.61 ms | 0 B |  | LargeXlsx | 409.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | ClosedXML | 368.54 ms | 101.8 MB |  | LargeXlsx | 714.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | EPPlus | 374.20 ms | 114.6 MB |  | LargeXlsx | 726.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-object-table-direct | OfficeIMO.Excel | 48.83 ms | 9.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datatable-object-table-direct | MiniExcel | 106.06 ms | 87.5 MB |  | OfficeIMO.Excel | 117.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-object-table-direct | EPPlus | 368.87 ms | 111.9 MB |  | OfficeIMO.Excel | 655.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-object-table-direct | ClosedXML | 488.68 ms | 166.7 MB |  | OfficeIMO.Excel | 900.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | OfficeIMO.Excel | 53.05 ms | 12.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datatable-table-direct | MiniExcel | 108.42 ms | 90.2 MB |  | OfficeIMO.Excel | 104.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | EPPlus 4.5.3.3 | 227.28 ms | 0 B |  | OfficeIMO.Excel | 328.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | EPPlus | 372.19 ms | 114.6 MB |  | OfficeIMO.Excel | 601.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | ClosedXML | 512.46 ms | 169.3 MB |  | OfficeIMO.Excel | 866.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-dictionary-objects-table-direct | OfficeIMO.Excel | 63.08 ms | 14.5 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | LargeXlsx | 43.24 ms | 5.5 MB |  | LargeXlsx | 13.7% faster than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 50.10 ms | 12.3 MB |  | LargeXlsx | Loss +15.9% |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | MiniExcel | 105.41 ms | 91.1 MB |  | LargeXlsx | 110.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 277.80 ms | 0 B |  | LargeXlsx | 454.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus | 388.85 ms | 114.7 MB |  | LargeXlsx | 676.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | ClosedXML | 408.35 ms | 101.8 MB |  | LargeXlsx | 715.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 54.90 ms | 12.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 427.07 ms | 156.0 MB |  | OfficeIMO.Excel | 677.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 1.02 s | 485.3 MB |  | OfficeIMO.Excel | 1757.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | LargeXlsx | 43.06 ms | 5.5 MB |  | LargeXlsx | 18.6% faster than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | OfficeIMO.Excel | 52.88 ms | 12.3 MB |  | LargeXlsx | Loss +22.8% |
| 25000 | speed-comparison | write-insertobjects-direct | MiniExcel | 104.11 ms | 91.1 MB |  | LargeXlsx | 96.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | EPPlus 4.5.3.3 | 219.20 ms | 0 B |  | LargeXlsx | 314.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | EPPlus | 404.99 ms | 114.7 MB |  | LargeXlsx | 665.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | ClosedXML | 422.95 ms | 101.8 MB |  | LargeXlsx | 699.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 82.00 ms | 6.9 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 476.10 ms | 156.0 MB |  | OfficeIMO.Excel | 480.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 1.09 s | 485.3 MB |  | OfficeIMO.Excel | 1225.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 52.79 ms | 5.5 MB |  | LargeXlsx | 24.0% faster than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 69.43 ms | 6.9 MB |  | LargeXlsx | Loss +31.5% |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | MiniExcel | 122.05 ms | 91.1 MB |  | LargeXlsx | 75.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | ClosedXML | 451.72 ms | 101.8 MB |  | LargeXlsx | 550.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | EPPlus | 510.98 ms | 114.7 MB |  | LargeXlsx | 636.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 51.91 ms | 12.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 398.70 ms | 135.0 MB |  | OfficeIMO.Excel | 668.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 668.25 ms | 269.0 MB |  | OfficeIMO.Excel | 1187.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | LargeXlsx | 45.88 ms | 5.9 MB |  | LargeXlsx | 15.5% faster than OfficeIMO |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 54.28 ms | 7.1 MB |  | LargeXlsx | Loss +18.3% |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | MiniExcel | 101.37 ms | 111.3 MB |  | LargeXlsx | 86.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | EPPlus | 468.01 ms | 141.5 MB |  | LargeXlsx | 762.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | ClosedXML | 497.42 ms | 175.3 MB |  | LargeXlsx | 816.3% slower than OfficeIMO |
| 300000 | focused-package-profile | write-blog-2023-20-string-columns | LargeXlsx | 636.47 ms | 93.1 MB | 28.6 MB | LargeXlsx | Win |
| 300000 | focused-package-profile | write-blog-2023-20-string-columns | OfficeIMO.Excel | 701.65 ms | 173.4 MB | 26.6 MB | LargeXlsx | 1.10x vs best |
| 300000 | focused-package-profile | write-blog-2023-20-string-columns | MiniExcel | 2.28 s | 2.46 GB | 28.5 MB | LargeXlsx | 3.58x vs best |
| 300000 | focused-package-profile | write-blog-2023-20-string-columns | ClosedXML | 15.89 s | 8.51 GB | 31.0 MB | LargeXlsx | 24.97x vs best |
