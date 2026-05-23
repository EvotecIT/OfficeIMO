# OfficeIMO.Excel Benchmark Report

Generated: 2026-05-23T13:48:50.3578200Z
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
| 2500 | dense-helloworld-comparison | read | Other | 0 | 2 | dense-helloworld-read-stream vs Sylvan.Data.Excel (1.89x) |
| 2500 | package-profile | package | Package size | 12 | 0 |  |
| 2500 | speed-comparison | mutate | AutoFit and mutation | 1 | 0 |  |
| 2500 | speed-comparison | read | Formula write/read | 1 | 0 |  |
| 2500 | speed-comparison | read | Other | 1 | 2 | shared-string-read vs Sylvan.Data.Excel (2.06x) |
| 2500 | speed-comparison | read | Range and table read | 1 | 2 | read-top-range vs Sylvan.Data.Excel (2.48x) |
| 2500 | speed-comparison | read | Streaming read | 0 | 2 | read-top-range-stream vs Sylvan.Data.Excel (3.97x) |
| 2500 | speed-comparison | read | Typed object read | 0 | 2 | read-objects-stream vs Sylvan.Data.Excel (1.83x) |
| 2500 | speed-comparison | write | AutoFit and mutation | 1 | 0 |  |
| 2500 | speed-comparison | write | DataSet table export | 3 | 0 |  |
| 2500 | speed-comparison | write | DataTable table export | 3 | 0 |  |
| 2500 | speed-comparison | write | Formatted report write | 1 | 0 |  |
| 2500 | speed-comparison | write | Plain cell export | 2 | 0 |  |
| 2500 | speed-comparison | write | Shared string write | 1 | 0 |  |
| 2500 | speed-comparison | write | Typed object export | 2 | 0 |  |
| 10000 | focused-package-profile | package | Package size | 1 | 0 |  |
| 25000 | dense-helloworld-comparison | read | Other | 0 | 2 | dense-helloworld-read-stream vs Sylvan.Data.Excel (1.82x) |
| 25000 | package-profile | package | Package size | 10 | 2 | write-bulk-report vs MiniExcel (1.20x) |
| 25000 | speed-comparison | mutate | AutoFit and mutation | 0 | 1 | autofit-existing vs EPPlus (1.03x) |
| 25000 | speed-comparison | read | Formula write/read | 1 | 0 |  |
| 25000 | speed-comparison | read | Other | 0 | 3 | shared-string-read vs Sylvan.Data.Excel (2.33x) |
| 25000 | speed-comparison | read | Range and table read | 0 | 3 | read-range vs Sylvan.Data.Excel (6.59x) |
| 25000 | speed-comparison | read | Streaming read | 0 | 2 | read-top-range-stream vs Sylvan.Data.Excel (5.27x) |
| 25000 | speed-comparison | read | Typed object read | 0 | 2 | read-objects vs Sylvan.Data.Excel (5.49x) |
| 25000 | speed-comparison | write | AutoFit and mutation | 1 | 0 |  |
| 25000 | speed-comparison | write | DataSet table export | 3 | 0 |  |
| 25000 | speed-comparison | write | DataTable table export | 3 | 0 |  |
| 25000 | speed-comparison | write | Formatted report write | 0 | 1 | write-bulk-report vs MiniExcel (1.11x) |
| 25000 | speed-comparison | write | Plain cell export | 2 | 0 |  |
| 25000 | speed-comparison | write | Shared string write | 1 | 0 |  |
| 25000 | speed-comparison | write | Typed object export | 2 | 0 |  |
| 300000 | focused-package-profile | package | Package size | 0 | 1 | write-blog-2023-20-string-columns vs LargeXlsx (1.10x) |

## Rows

| Rows | Artifact | Scenario | Library | Mean | Allocated | Package | Best | Outcome |
| ---: | --- | --- | --- | ---: | ---: | ---: | --- | --- |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | Sylvan.Data.Excel | 5.83 ms | 357.8 KB |  | Sylvan.Data.Excel | 28.7% faster than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | OfficeIMO.Excel | 8.19 ms | 2.7 MB |  | Sylvan.Data.Excel | Loss +40.3% |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | MiniExcel | 20.58 ms | 21.0 MB |  | Sylvan.Data.Excel | 151.4% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | Sylvan.Data.Excel | 5.63 ms | 357.8 KB |  | Sylvan.Data.Excel | 47.2% faster than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | OfficeIMO.Excel | 10.65 ms | 2.7 MB |  | Sylvan.Data.Excel | Loss +89.3% |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | MiniExcel | 22.27 ms | 21.0 MB |  | Sylvan.Data.Excel | 109.1% slower than OfficeIMO |
| 2500 | package-profile | append-plain-rows | OfficeIMO.Excel | 3.97 ms | 3.6 MB | 64.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | append-plain-rows | MiniExcel | 5.76 ms | 19.2 MB | 68.1 KB | OfficeIMO.Excel | 44.9% slower than OfficeIMO |
| 2500 | package-profile | append-plain-rows | ClosedXML | 18.47 ms | 10.9 MB | 59.8 KB | OfficeIMO.Excel | 365.1% slower than OfficeIMO |
| 2500 | package-profile | append-plain-rows | EPPlus | 21.47 ms | 14.0 MB | 56.9 KB | OfficeIMO.Excel | 440.6% slower than OfficeIMO |
| 2500 | package-profile | autofit-existing | OfficeIMO.Excel | 35.07 ms | 13.5 MB | 139.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | autofit-existing | EPPlus | 80.97 ms | 49.5 MB | 115.0 KB | OfficeIMO.Excel | 130.9% slower than OfficeIMO |
| 2500 | package-profile | autofit-existing | ClosedXML | 159.91 ms | 82.6 MB | 121.0 KB | OfficeIMO.Excel | 355.9% slower than OfficeIMO |
| 2500 | package-profile | large-shared-strings | OfficeIMO.Excel | 2.87 ms | 4.1 MB | 52.6 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | large-shared-strings | MiniExcel | 5.62 ms | 20.6 MB | 60.7 KB | OfficeIMO.Excel | 96.1% slower than OfficeIMO |
| 2500 | package-profile | large-shared-strings | ClosedXML | 15.88 ms | 11.0 MB | 50.3 KB | OfficeIMO.Excel | 453.7% slower than OfficeIMO |
| 2500 | package-profile | large-shared-strings | EPPlus | 19.97 ms | 12.5 MB | 48.1 KB | OfficeIMO.Excel | 596.2% slower than OfficeIMO |
| 2500 | package-profile | write-bulk-report | OfficeIMO.Excel | 9.70 ms | 5.7 MB | 139.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-bulk-report | MiniExcel | 11.30 ms | 26.2 MB | 153.8 KB | OfficeIMO.Excel | 16.5% slower than OfficeIMO |
| 2500 | package-profile | write-bulk-report | EPPlus | 67.12 ms | 46.1 MB | 115.0 KB | OfficeIMO.Excel | 591.7% slower than OfficeIMO |
| 2500 | package-profile | write-bulk-report | ClosedXML | 105.30 ms | 57.0 MB | 121.0 KB | OfficeIMO.Excel | 985.1% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 5.21 ms | 3.7 MB | 138.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalues-rectangle-direct | MiniExcel | 9.93 ms | 22.7 MB | 153.7 KB | OfficeIMO.Excel | 90.7% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | ClosedXML | 38.38 ms | 21.7 MB | 120.1 KB | OfficeIMO.Excel | 636.7% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | EPPlus | 42.75 ms | 24.1 MB | 114.1 KB | OfficeIMO.Excel | 720.7% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table | OfficeIMO.Excel | 8.97 ms | 4.9 MB | 138.8 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-datareader-table | MiniExcel | 9.31 ms | 22.5 MB | 153.6 KB | OfficeIMO.Excel | 3.7% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table | EPPlus | 36.78 ms | 16.3 MB | 114.9 KB | OfficeIMO.Excel | 310.0% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table | ClosedXML | 44.03 ms | 18.6 MB | 120.9 KB | OfficeIMO.Excel | 390.9% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables | OfficeIMO.Excel | 5.61 ms | 4.6 MB | 139.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-dataset-tables | MiniExcel | 10.37 ms | 28.0 MB | 156.4 KB | OfficeIMO.Excel | 85.0% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables | EPPlus | 40.31 ms | 18.3 MB | 116.6 KB | OfficeIMO.Excel | 619.0% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables | ClosedXML | 49.97 ms | 18.4 MB | 123.4 KB | OfficeIMO.Excel | 791.2% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables-autofit | OfficeIMO.Excel | 6.05 ms | 4.6 MB | 139.3 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-dataset-tables-autofit | MiniExcel | 10.37 ms | 31.1 MB | 156.6 KB | OfficeIMO.Excel | 71.4% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables-autofit | EPPlus | 66.69 ms | 40.5 MB | 116.9 KB | OfficeIMO.Excel | 1002.3% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables-autofit | ClosedXML | 100.46 ms | 55.4 MB | 123.7 KB | OfficeIMO.Excel | 1560.5% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | OfficeIMO.Excel | 4.99 ms | 3.4 MB | 138.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-datatable-direct | MiniExcel | 9.56 ms | 22.5 MB | 153.7 KB | OfficeIMO.Excel | 91.8% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | ClosedXML | 33.22 ms | 11.3 MB | 120.1 KB | OfficeIMO.Excel | 566.1% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | EPPlus | 35.71 ms | 16.3 MB | 114.9 KB | OfficeIMO.Excel | 616.0% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-table-direct | OfficeIMO.Excel | 5.71 ms | 3.8 MB | 138.8 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-datatable-table-direct | MiniExcel | 9.70 ms | 22.5 MB | 153.7 KB | OfficeIMO.Excel | 69.7% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-table-direct | EPPlus | 36.54 ms | 16.3 MB | 114.9 KB | OfficeIMO.Excel | 539.4% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-table-direct | ClosedXML | 42.08 ms | 18.6 MB | 120.9 KB | OfficeIMO.Excel | 636.4% slower than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 5.40 ms | 3.4 MB | 138.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-fluent-rowsfrom-direct | MiniExcel | 8.66 ms | 22.7 MB | 153.7 KB | OfficeIMO.Excel | 60.4% slower than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | ClosedXML | 33.46 ms | 11.3 MB | 120.1 KB | OfficeIMO.Excel | 520.0% slower than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | EPPlus | 35.41 ms | 16.3 MB | 114.9 KB | OfficeIMO.Excel | 556.3% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | OfficeIMO.Excel | 4.48 ms | 3.4 MB | 138.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-insertobjects-direct | MiniExcel | 10.60 ms | 22.7 MB | 153.7 KB | OfficeIMO.Excel | 136.6% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | EPPlus | 35.97 ms | 16.3 MB | 114.9 KB | OfficeIMO.Excel | 702.6% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | ClosedXML | 36.98 ms | 11.3 MB | 120.1 KB | OfficeIMO.Excel | 725.1% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | OfficeIMO.Excel | 4.95 ms | 3.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | append-plain-rows | MiniExcel | 5.59 ms | 19.2 MB |  | OfficeIMO.Excel | 13.1% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | EPPlus 4.5.3.3 | 15.79 ms | 0 B |  | OfficeIMO.Excel | 219.3% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | ClosedXML | 19.72 ms | 10.9 MB |  | OfficeIMO.Excel | 298.8% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | EPPlus | 21.63 ms | 14.0 MB |  | OfficeIMO.Excel | 337.3% slower than OfficeIMO |
| 2500 | speed-comparison | autofit-existing | OfficeIMO.Excel | 38.28 ms | 13.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | autofit-existing | EPPlus | 84.32 ms | 49.5 MB |  | OfficeIMO.Excel | 120.2% slower than OfficeIMO |
| 2500 | speed-comparison | autofit-existing | EPPlus 4.5.3.3 | 140.94 ms | 0 B |  | OfficeIMO.Excel | 268.1% slower than OfficeIMO |
| 2500 | speed-comparison | autofit-existing | ClosedXML | 172.84 ms | 82.7 MB |  | OfficeIMO.Excel | 351.5% slower than OfficeIMO |
| 2500 | speed-comparison | formula-heavy-read | OfficeIMO.Excel | 4.95 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | formula-heavy-read | EPPlus | 16.37 ms | 8.1 MB |  | OfficeIMO.Excel | 230.4% slower than OfficeIMO |
| 2500 | speed-comparison | formula-heavy-read | EPPlus 4.5.3.3 | 20.61 ms | 0 B |  | OfficeIMO.Excel | 316.0% slower than OfficeIMO |
| 2500 | speed-comparison | formula-heavy-read | ClosedXML | 22.40 ms | 9.1 MB |  | OfficeIMO.Excel | 352.0% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | OfficeIMO.Excel | 3.61 ms | 4.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | large-shared-strings | MiniExcel | 4.99 ms | 20.6 MB |  | OfficeIMO.Excel | 38.5% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | ClosedXML | 15.49 ms | 11.0 MB |  | OfficeIMO.Excel | 329.6% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | EPPlus | 21.40 ms | 12.5 MB |  | OfficeIMO.Excel | 493.6% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | EPPlus 4.5.3.3 | 23.83 ms | 0 B |  | OfficeIMO.Excel | 560.8% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | OfficeIMO.Excel | 3.16 ms | 289.2 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | large-sparse-column-read | Sylvan.Data.Excel | 3.48 ms | 316.5 KB |  | OfficeIMO.Excel | 10.2% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | ClosedXML | 3.92 ms | 4.3 MB |  | OfficeIMO.Excel | 24.2% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | EPPlus 4.5.3.3 | 9.64 ms | 0 B |  | OfficeIMO.Excel | 204.9% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | MiniExcel | 18.75 ms | 45.1 MB |  | OfficeIMO.Excel | 493.2% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | EPPlus | 33.11 ms | 42.1 MB |  | OfficeIMO.Excel | 947.4% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | Sylvan.Data.Excel | 2.18 ms | 316.5 KB |  | Sylvan.Data.Excel | 42.6% faster than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | OfficeIMO.Excel | 3.80 ms | 289.1 KB |  | Sylvan.Data.Excel | Loss +74.1% |
| 2500 | speed-comparison | large-sparse-row-read | ClosedXML | 4.01 ms | 4.3 MB |  | Sylvan.Data.Excel | 5.6% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | EPPlus 4.5.3.3 | 9.42 ms | 0 B |  | Sylvan.Data.Excel | 148.1% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | MiniExcel | 19.45 ms | 45.1 MB |  | Sylvan.Data.Excel | 412.3% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | EPPlus | 32.86 ms | 42.1 MB |  | Sylvan.Data.Excel | 765.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | Sylvan.Data.Excel | 11.75 ms | 1.9 MB |  | Sylvan.Data.Excel | 17.0% faster than OfficeIMO |
| 2500 | speed-comparison | read-datatable | OfficeIMO.Excel | 14.16 ms | 4.1 MB |  | Sylvan.Data.Excel | Loss +20.5% |
| 2500 | speed-comparison | read-datatable | MiniExcel | 23.23 ms | 18.2 MB |  | Sylvan.Data.Excel | 64.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | EPPlus 4.5.3.3 | 36.94 ms | 0 B |  | Sylvan.Data.Excel | 160.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | EPPlus | 39.77 ms | 20.0 MB |  | Sylvan.Data.Excel | 180.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | ClosedXML | 117.49 ms | 21.7 MB |  | Sylvan.Data.Excel | 729.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | Sylvan.Data.Excel | 7.53 ms | 544.4 KB |  | Sylvan.Data.Excel | 30.5% faster than OfficeIMO |
| 2500 | speed-comparison | read-objects | OfficeIMO.Excel | 10.83 ms | 3.1 MB |  | Sylvan.Data.Excel | Loss +43.9% |
| 2500 | speed-comparison | read-objects | MiniExcel | 24.37 ms | 18.3 MB |  | Sylvan.Data.Excel | 125.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | EPPlus | 25.83 ms | 19.6 MB |  | Sylvan.Data.Excel | 138.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | EPPlus 4.5.3.3 | 34.14 ms | 0 B |  | Sylvan.Data.Excel | 215.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | ClosedXML | 51.27 ms | 20.2 MB |  | Sylvan.Data.Excel | 373.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | Sylvan.Data.Excel | 5.77 ms | 544.5 KB |  | Sylvan.Data.Excel | 45.5% faster than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | OfficeIMO.Excel | 10.57 ms | 1.4 MB |  | Sylvan.Data.Excel | Loss +83.4% |
| 2500 | speed-comparison | read-objects-stream | MiniExcel | 22.51 ms | 18.3 MB |  | Sylvan.Data.Excel | 112.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | EPPlus | 29.27 ms | 19.6 MB |  | Sylvan.Data.Excel | 176.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | EPPlus 4.5.3.3 | 33.74 ms | 0 B |  | Sylvan.Data.Excel | 219.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | ClosedXML | 46.12 ms | 20.2 MB |  | Sylvan.Data.Excel | 336.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | OfficeIMO.Excel | 15.31 ms | 2.8 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | read-range | Sylvan.Data.Excel | 18.51 ms | 368.5 KB |  | OfficeIMO.Excel | 20.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | MiniExcel | 30.85 ms | 18.6 MB |  | OfficeIMO.Excel | 101.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | EPPlus | 32.76 ms | 18.5 MB |  | OfficeIMO.Excel | 114.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | EPPlus 4.5.3.3 | 33.80 ms | 0 B |  | OfficeIMO.Excel | 120.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | ClosedXML | 91.23 ms | 20.2 MB |  | OfficeIMO.Excel | 496.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | Sylvan.Data.Excel | 6.14 ms | 368.7 KB |  | Sylvan.Data.Excel | 51.3% faster than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | OfficeIMO.Excel | 12.62 ms | 3.0 MB |  | Sylvan.Data.Excel | Loss +105.4% |
| 2500 | speed-comparison | read-range-stream | MiniExcel | 20.54 ms | 18.6 MB |  | Sylvan.Data.Excel | 62.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | EPPlus | 26.94 ms | 18.5 MB |  | Sylvan.Data.Excel | 113.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | EPPlus 4.5.3.3 | 31.90 ms | 0 B |  | Sylvan.Data.Excel | 152.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | ClosedXML | 80.90 ms | 20.1 MB |  | Sylvan.Data.Excel | 541.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | Sylvan.Data.Excel | 1.14 ms | 365.3 KB |  | Sylvan.Data.Excel | 59.6% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range | MiniExcel | 1.99 ms | 901.9 KB |  | Sylvan.Data.Excel | 29.2% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range | OfficeIMO.Excel | 2.82 ms | 583.6 KB |  | Sylvan.Data.Excel | Loss +147.6% |
| 2500 | speed-comparison | read-top-range | EPPlus 4.5.3.3 | 27.71 ms | 0 B |  | Sylvan.Data.Excel | 884.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | EPPlus | 31.20 ms | 10.9 MB |  | Sylvan.Data.Excel | 1008.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | ClosedXML | 101.50 ms | 18.8 MB |  | Sylvan.Data.Excel | 3505.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | Sylvan.Data.Excel | 0.63 ms | 365.5 KB |  | Sylvan.Data.Excel | 74.8% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | MiniExcel | 1.32 ms | 902.1 KB |  | Sylvan.Data.Excel | 47.2% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | OfficeIMO.Excel | 2.50 ms | 587.3 KB |  | Sylvan.Data.Excel | Loss +297.4% |
| 2500 | speed-comparison | read-top-range-stream | EPPlus | 25.14 ms | 10.9 MB |  | Sylvan.Data.Excel | 904.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | EPPlus 4.5.3.3 | 29.54 ms | 0 B |  | Sylvan.Data.Excel | 1080.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | ClosedXML | 50.58 ms | 18.7 MB |  | Sylvan.Data.Excel | 1921.3% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | Sylvan.Data.Excel | 3.38 ms | 578.8 KB |  | Sylvan.Data.Excel | 51.3% faster than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | OfficeIMO.Excel | 6.95 ms | 2.8 MB |  | Sylvan.Data.Excel | Loss +105.5% |
| 2500 | speed-comparison | shared-string-read | MiniExcel | 7.43 ms | 9.5 MB |  | Sylvan.Data.Excel | 6.9% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | EPPlus | 14.87 ms | 11.0 MB |  | Sylvan.Data.Excel | 114.0% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | ClosedXML | 22.18 ms | 11.4 MB |  | Sylvan.Data.Excel | 219.1% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | EPPlus 4.5.3.3 | 26.11 ms | 0 B |  | Sylvan.Data.Excel | 275.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | OfficeIMO.Excel | 14.90 ms | 5.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-bulk-report | MiniExcel | 16.98 ms | 26.2 MB |  | OfficeIMO.Excel | 14.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | EPPlus 4.5.3.3 | 108.37 ms | 0 B |  | OfficeIMO.Excel | 627.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | EPPlus | 140.98 ms | 48.0 MB |  | OfficeIMO.Excel | 846.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | ClosedXML | 333.61 ms | 57.0 MB |  | OfficeIMO.Excel | 2139.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 7.50 ms | 3.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | MiniExcel | 13.34 ms | 22.7 MB |  | OfficeIMO.Excel | 77.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 32.58 ms | 0 B |  | OfficeIMO.Excel | 334.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | ClosedXML | 44.58 ms | 21.7 MB |  | OfficeIMO.Excel | 494.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus | 51.54 ms | 24.1 MB |  | OfficeIMO.Excel | 586.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | OfficeIMO.Excel | 9.59 ms | 4.9 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datareader-table | MiniExcel | 11.65 ms | 22.5 MB |  | OfficeIMO.Excel | 21.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | EPPlus | 39.47 ms | 16.3 MB |  | OfficeIMO.Excel | 311.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | ClosedXML | 45.39 ms | 18.6 MB |  | OfficeIMO.Excel | 373.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table-direct | EPPlus 4.5.3.3 | 36.89 ms | 0 B |  | EPPlus 4.5.3.3 |  |
| 2500 | speed-comparison | write-dataset-headerless-tables | OfficeIMO.Excel | 4.70 ms | 4.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-headerless-tables | MiniExcel | 16.83 ms | 28.0 MB |  | OfficeIMO.Excel | 258.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-headerless-tables | ClosedXML | 52.53 ms | 18.5 MB |  | OfficeIMO.Excel | 1017.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-headerless-tables | EPPlus | 70.94 ms | 18.0 MB |  | OfficeIMO.Excel | 1409.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-tables | OfficeIMO.Excel | 8.04 ms | 5.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-sparse-tables | MiniExcel | 23.07 ms | 29.0 MB |  | OfficeIMO.Excel | 187.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-tables | EPPlus | 62.41 ms | 21.4 MB |  | OfficeIMO.Excel | 676.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-tables | ClosedXML | 75.15 ms | 26.8 MB |  | OfficeIMO.Excel | 835.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | OfficeIMO.Excel | 7.47 ms | 4.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-tables | MiniExcel | 17.95 ms | 28.0 MB |  | OfficeIMO.Excel | 140.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | EPPlus 4.5.3.3 | 40.44 ms | 0 B |  | OfficeIMO.Excel | 441.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | EPPlus | 86.23 ms | 19.0 MB |  | OfficeIMO.Excel | 1053.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | ClosedXML | 189.98 ms | 18.4 MB |  | OfficeIMO.Excel | 2441.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables-autofit | OfficeIMO.Excel | 7.06 ms | 4.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-tables-autofit | MiniExcel | 19.10 ms | 31.1 MB |  | OfficeIMO.Excel | 170.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables-autofit | EPPlus | 153.87 ms | 42.4 MB |  | OfficeIMO.Excel | 2078.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables-autofit | ClosedXML | 271.37 ms | 55.4 MB |  | OfficeIMO.Excel | 3742.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | OfficeIMO.Excel | 5.61 ms | 3.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datatable-direct | MiniExcel | 13.17 ms | 22.5 MB |  | OfficeIMO.Excel | 134.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | ClosedXML | 37.46 ms | 11.3 MB |  | OfficeIMO.Excel | 567.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | EPPlus | 40.63 ms | 16.3 MB |  | OfficeIMO.Excel | 624.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | EPPlus 4.5.3.3 | 41.55 ms | 0 B |  | OfficeIMO.Excel | 640.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | OfficeIMO.Excel | 5.26 ms | 3.8 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datatable-table-direct | MiniExcel | 12.09 ms | 22.5 MB |  | OfficeIMO.Excel | 129.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | EPPlus | 40.02 ms | 16.3 MB |  | OfficeIMO.Excel | 660.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | EPPlus 4.5.3.3 | 42.18 ms | 0 B |  | OfficeIMO.Excel | 701.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | ClosedXML | 47.84 ms | 18.6 MB |  | OfficeIMO.Excel | 809.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 6.44 ms | 3.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | MiniExcel | 11.46 ms | 22.7 MB |  | OfficeIMO.Excel | 77.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 29.50 ms | 0 B |  | OfficeIMO.Excel | 357.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | ClosedXML | 35.38 ms | 11.3 MB |  | OfficeIMO.Excel | 449.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus | 39.19 ms | 16.3 MB |  | OfficeIMO.Excel | 508.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | OfficeIMO.Excel | 6.74 ms | 3.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-insertobjects-direct | MiniExcel | 13.17 ms | 22.7 MB |  | OfficeIMO.Excel | 95.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | EPPlus 4.5.3.3 | 29.59 ms | 0 B |  | OfficeIMO.Excel | 339.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | EPPlus | 38.47 ms | 16.3 MB |  | OfficeIMO.Excel | 471.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | ClosedXML | 42.38 ms | 11.3 MB |  | OfficeIMO.Excel | 529.0% slower than OfficeIMO |
| 10000 | focused-package-profile | write-blog-2023-20-string-columns | OfficeIMO.Excel | 33.01 ms | 7.6 MB | 880.4 KB | OfficeIMO.Excel | Win |
| 10000 | focused-package-profile | write-blog-2023-20-string-columns | LargeXlsx | 88.91 ms | 3.1 MB | 970.2 KB | OfficeIMO.Excel | 2.69x vs best |
| 10000 | focused-package-profile | write-blog-2023-20-string-columns | MiniExcel | 142.47 ms | 96.2 MB | 957.6 KB | OfficeIMO.Excel | 4.32x vs best |
| 10000 | focused-package-profile | write-blog-2023-20-string-columns | ClosedXML | 725.30 ms | 280.2 MB | 1,015.4 KB | OfficeIMO.Excel | 21.97x vs best |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | Sylvan.Data.Excel | 48.00 ms | 376.2 KB |  | Sylvan.Data.Excel | 23.4% faster than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | OfficeIMO.Excel | 62.70 ms | 25.1 MB |  | Sylvan.Data.Excel | Loss +30.6% |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | MiniExcel | 181.01 ms | 210.3 MB |  | Sylvan.Data.Excel | 188.7% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | Sylvan.Data.Excel | 59.33 ms | 376.2 KB |  | Sylvan.Data.Excel | 45.0% faster than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | OfficeIMO.Excel | 107.85 ms | 26.7 MB |  | Sylvan.Data.Excel | Loss +81.8% |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | MiniExcel | 201.97 ms | 210.3 MB |  | Sylvan.Data.Excel | 87.3% slower than OfficeIMO |
| 25000 | package-profile | append-plain-rows | OfficeIMO.Excel | 27.34 ms | 14.1 MB | 622.6 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | append-plain-rows | MiniExcel | 37.18 ms | 56.9 MB | 642.3 KB | OfficeIMO.Excel | 36.0% slower than OfficeIMO |
| 25000 | package-profile | append-plain-rows | EPPlus | 159.43 ms | 97.9 MB | 525.6 KB | OfficeIMO.Excel | 483.1% slower than OfficeIMO |
| 25000 | package-profile | append-plain-rows | ClosedXML | 161.98 ms | 101.8 MB | 540.6 KB | OfficeIMO.Excel | 492.5% slower than OfficeIMO |
| 25000 | package-profile | autofit-existing | OfficeIMO.Excel | 406.19 ms | 132.8 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | autofit-existing | EPPlus | 478.09 ms | 245.0 MB | 1.1 MB | OfficeIMO.Excel | 17.7% slower than OfficeIMO |
| 25000 | package-profile | autofit-existing | ClosedXML | 1.64 s | 810.1 MB | 1.1 MB | OfficeIMO.Excel | 303.5% slower than OfficeIMO |
| 25000 | package-profile | large-shared-strings | OfficeIMO.Excel | 22.13 ms | 20.0 MB | 520.3 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | large-shared-strings | MiniExcel | 38.35 ms | 72.0 MB | 581.0 KB | OfficeIMO.Excel | 73.3% slower than OfficeIMO |
| 25000 | package-profile | large-shared-strings | EPPlus | 130.40 ms | 82.4 MB | 444.7 KB | OfficeIMO.Excel | 489.4% slower than OfficeIMO |
| 25000 | package-profile | large-shared-strings | ClosedXML | 134.38 ms | 101.8 MB | 460.1 KB | OfficeIMO.Excel | 507.4% slower than OfficeIMO |
| 25000 | package-profile | write-bulk-report | MiniExcel | 82.33 ms | 122.6 MB | 1.5 MB | MiniExcel | 16.8% faster than OfficeIMO |
| 25000 | package-profile | write-bulk-report | OfficeIMO.Excel | 98.89 ms | 32.5 MB | 1.4 MB | MiniExcel | Loss +20.1% |
| 25000 | package-profile | write-bulk-report | EPPlus | 388.91 ms | 248.9 MB | 1.1 MB | MiniExcel | 293.3% slower than OfficeIMO |
| 25000 | package-profile | write-bulk-report | ClosedXML | 973.26 ms | 552.7 MB | 1.1 MB | MiniExcel | 884.2% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 49.45 ms | 18.0 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalues-rectangle-direct | MiniExcel | 81.80 ms | 91.1 MB | 1.5 MB | OfficeIMO.Excel | 65.4% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | EPPlus | 355.05 ms | 206.8 MB | 1.1 MB | OfficeIMO.Excel | 618.0% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | ClosedXML | 380.73 ms | 205.7 MB | 1.1 MB | OfficeIMO.Excel | 669.9% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table | MiniExcel | 86.42 ms | 90.0 MB | 1.5 MB | MiniExcel | 4.8% faster than OfficeIMO |
| 25000 | package-profile | write-datareader-table | OfficeIMO.Excel | 90.74 ms | 25.2 MB | 1.4 MB | MiniExcel | Loss +5.0% |
| 25000 | package-profile | write-datareader-table | EPPlus | 320.77 ms | 114.6 MB | 1.1 MB | MiniExcel | 253.5% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table | ClosedXML | 501.10 ms | 169.3 MB | 1.1 MB | MiniExcel | 452.3% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables | OfficeIMO.Excel | 42.54 ms | 13.3 MB | 1.3 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-dataset-tables | MiniExcel | 89.33 ms | 94.8 MB | 1.5 MB | OfficeIMO.Excel | 110.0% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables | EPPlus | 309.52 ms | 108.1 MB | 1.1 MB | OfficeIMO.Excel | 627.7% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables | ClosedXML | 460.61 ms | 168.0 MB | 1.1 MB | OfficeIMO.Excel | 982.8% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables-autofit | OfficeIMO.Excel | 47.34 ms | 13.3 MB | 1.3 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-dataset-tables-autofit | MiniExcel | 107.16 ms | 125.8 MB | 1.5 MB | OfficeIMO.Excel | 126.4% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables-autofit | EPPlus | 399.57 ms | 190.7 MB | 1.1 MB | OfficeIMO.Excel | 744.0% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables-autofit | ClosedXML | 929.55 ms | 537.2 MB | 1.1 MB | OfficeIMO.Excel | 1863.5% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | OfficeIMO.Excel | 44.92 ms | 14.8 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-datatable-direct | MiniExcel | 99.34 ms | 90.2 MB | 1.5 MB | OfficeIMO.Excel | 121.2% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | EPPlus | 337.50 ms | 114.6 MB | 1.1 MB | OfficeIMO.Excel | 651.4% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | ClosedXML | 353.51 ms | 101.8 MB | 1.1 MB | OfficeIMO.Excel | 687.0% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-table-direct | OfficeIMO.Excel | 49.60 ms | 15.3 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-datatable-table-direct | MiniExcel | 99.11 ms | 90.2 MB | 1.5 MB | OfficeIMO.Excel | 99.8% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-table-direct | EPPlus | 317.06 ms | 114.6 MB | 1.1 MB | OfficeIMO.Excel | 539.2% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-table-direct | ClosedXML | 505.15 ms | 169.3 MB | 1.1 MB | OfficeIMO.Excel | 918.5% slower than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 51.57 ms | 15.2 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-fluent-rowsfrom-direct | MiniExcel | 88.15 ms | 91.1 MB | 1.5 MB | OfficeIMO.Excel | 71.0% slower than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | EPPlus | 324.70 ms | 114.6 MB | 1.1 MB | OfficeIMO.Excel | 529.7% slower than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | ClosedXML | 341.71 ms | 101.8 MB | 1.1 MB | OfficeIMO.Excel | 562.7% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | OfficeIMO.Excel | 45.40 ms | 15.0 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-insertobjects-direct | MiniExcel | 87.89 ms | 91.1 MB | 1.5 MB | OfficeIMO.Excel | 93.6% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | EPPlus | 316.89 ms | 114.6 MB | 1.1 MB | OfficeIMO.Excel | 598.0% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | ClosedXML | 344.21 ms | 101.8 MB | 1.1 MB | OfficeIMO.Excel | 658.2% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | OfficeIMO.Excel | 20.10 ms | 14.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | append-plain-rows | MiniExcel | 32.83 ms | 56.9 MB |  | OfficeIMO.Excel | 63.3% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | EPPlus 4.5.3.3 | 128.73 ms | 0 B |  | OfficeIMO.Excel | 540.4% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | EPPlus | 146.05 ms | 98.0 MB |  | OfficeIMO.Excel | 626.5% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | ClosedXML | 149.54 ms | 101.8 MB |  | OfficeIMO.Excel | 643.9% slower than OfficeIMO |
| 25000 | speed-comparison | autofit-existing | EPPlus | 522.02 ms | 245.1 MB |  | EPPlus | 3.1% faster than OfficeIMO |
| 25000 | speed-comparison | autofit-existing | OfficeIMO.Excel | 538.56 ms | 133.0 MB |  | EPPlus | Loss +3.2% |
| 25000 | speed-comparison | autofit-existing | EPPlus 4.5.3.3 | 736.34 ms | 0 B |  | EPPlus | 36.7% slower than OfficeIMO |
| 25000 | speed-comparison | autofit-existing | ClosedXML | 1.86 s | 810.5 MB |  | EPPlus | 245.8% slower than OfficeIMO |
| 25000 | speed-comparison | formula-heavy-read | OfficeIMO.Excel | 39.09 ms | 15.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | formula-heavy-read | EPPlus 4.5.3.3 | 79.28 ms | 0 B |  | OfficeIMO.Excel | 102.8% slower than OfficeIMO |
| 25000 | speed-comparison | formula-heavy-read | EPPlus | 121.16 ms | 74.0 MB |  | OfficeIMO.Excel | 209.9% slower than OfficeIMO |
| 25000 | speed-comparison | formula-heavy-read | ClosedXML | 265.01 ms | 87.7 MB |  | OfficeIMO.Excel | 577.9% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | OfficeIMO.Excel | 22.12 ms | 20.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | large-shared-strings | MiniExcel | 37.28 ms | 72.0 MB |  | OfficeIMO.Excel | 68.6% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | EPPlus 4.5.3.3 | 105.19 ms | 0 B |  | OfficeIMO.Excel | 375.6% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | ClosedXML | 139.09 ms | 101.8 MB |  | OfficeIMO.Excel | 528.9% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | EPPlus | 140.95 ms | 82.4 MB |  | OfficeIMO.Excel | 537.3% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | Sylvan.Data.Excel | 0.99 ms | 316.5 KB |  | Sylvan.Data.Excel | 39.6% faster than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | OfficeIMO.Excel | 1.64 ms | 289.0 KB |  | Sylvan.Data.Excel | Loss +65.5% |
| 25000 | speed-comparison | large-sparse-column-read | ClosedXML | 3.20 ms | 4.3 MB |  | Sylvan.Data.Excel | 95.1% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | EPPlus 4.5.3.3 | 11.34 ms | 0 B |  | Sylvan.Data.Excel | 592.5% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | MiniExcel | 11.60 ms | 45.1 MB |  | Sylvan.Data.Excel | 608.2% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | EPPlus | 30.02 ms | 42.1 MB |  | Sylvan.Data.Excel | 1732.5% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | Sylvan.Data.Excel | 0.93 ms | 316.5 KB |  | Sylvan.Data.Excel | 46.2% faster than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | OfficeIMO.Excel | 1.72 ms | 289.0 KB |  | Sylvan.Data.Excel | Loss +85.7% |
| 25000 | speed-comparison | large-sparse-row-read | ClosedXML | 3.09 ms | 4.3 MB |  | Sylvan.Data.Excel | 79.4% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | MiniExcel | 11.21 ms | 45.1 MB |  | Sylvan.Data.Excel | 551.0% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | EPPlus 4.5.3.3 | 15.24 ms | 0 B |  | Sylvan.Data.Excel | 784.4% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | EPPlus | 28.12 ms | 42.1 MB |  | Sylvan.Data.Excel | 1532.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | Sylvan.Data.Excel | 53.21 ms | 14.9 MB |  | Sylvan.Data.Excel | 83.3% faster than OfficeIMO |
| 25000 | speed-comparison | read-datatable | MiniExcel | 165.32 ms | 180.5 MB |  | Sylvan.Data.Excel | 48.1% faster than OfficeIMO |
| 25000 | speed-comparison | read-datatable | EPPlus | 190.23 ms | 186.5 MB |  | Sylvan.Data.Excel | 40.2% faster than OfficeIMO |
| 25000 | speed-comparison | read-datatable | EPPlus 4.5.3.3 | 216.64 ms | 0 B |  | Sylvan.Data.Excel | 31.9% faster than OfficeIMO |
| 25000 | speed-comparison | read-datatable | OfficeIMO.Excel | 318.28 ms | 213.0 MB |  | Sylvan.Data.Excel | Loss +498.2% |
| 25000 | speed-comparison | read-datatable | ClosedXML | 368.23 ms | 208.8 MB |  | Sylvan.Data.Excel | 15.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | Sylvan.Data.Excel | 63.73 ms | 2.1 MB |  | Sylvan.Data.Excel | 81.8% faster than OfficeIMO |
| 25000 | speed-comparison | read-objects | EPPlus 4.5.3.3 | 196.60 ms | 0 B |  | Sylvan.Data.Excel | 43.8% faster than OfficeIMO |
| 25000 | speed-comparison | read-objects | MiniExcel | 196.94 ms | 182.3 MB |  | Sylvan.Data.Excel | 43.7% faster than OfficeIMO |
| 25000 | speed-comparison | read-objects | EPPlus | 240.46 ms | 183.7 MB |  | Sylvan.Data.Excel | 31.3% faster than OfficeIMO |
| 25000 | speed-comparison | read-objects | OfficeIMO.Excel | 350.08 ms | 205.8 MB |  | Sylvan.Data.Excel | Loss +449.3% |
| 25000 | speed-comparison | read-objects | ClosedXML | 425.86 ms | 196.0 MB |  | Sylvan.Data.Excel | 21.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | Sylvan.Data.Excel | 56.34 ms | 2.1 MB |  | Sylvan.Data.Excel | 33.0% faster than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | OfficeIMO.Excel | 84.10 ms | 10.1 MB |  | Sylvan.Data.Excel | Loss +49.3% |
| 25000 | speed-comparison | read-objects-stream | MiniExcel | 190.21 ms | 182.3 MB |  | Sylvan.Data.Excel | 126.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | EPPlus 4.5.3.3 | 205.10 ms | 0 B |  | Sylvan.Data.Excel | 143.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | EPPlus | 232.97 ms | 183.7 MB |  | Sylvan.Data.Excel | 177.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | ClosedXML | 467.43 ms | 196.0 MB |  | Sylvan.Data.Excel | 455.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | Sylvan.Data.Excel | 41.41 ms | 398.5 KB |  | Sylvan.Data.Excel | 84.8% faster than OfficeIMO |
| 25000 | speed-comparison | read-range | MiniExcel | 132.04 ms | 185.5 MB |  | Sylvan.Data.Excel | 51.6% faster than OfficeIMO |
| 25000 | speed-comparison | read-range | EPPlus | 163.18 ms | 171.9 MB |  | Sylvan.Data.Excel | 40.2% faster than OfficeIMO |
| 25000 | speed-comparison | read-range | EPPlus 4.5.3.3 | 206.61 ms | 0 B |  | Sylvan.Data.Excel | 24.3% faster than OfficeIMO |
| 25000 | speed-comparison | read-range | OfficeIMO.Excel | 272.88 ms | 204.3 MB |  | Sylvan.Data.Excel | Loss +558.9% |
| 25000 | speed-comparison | read-range | ClosedXML | 343.34 ms | 194.2 MB |  | Sylvan.Data.Excel | 25.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | Sylvan.Data.Excel | 42.95 ms | 398.5 KB |  | Sylvan.Data.Excel | 45.3% faster than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | OfficeIMO.Excel | 78.52 ms | 26.4 MB |  | Sylvan.Data.Excel | Loss +82.8% |
| 25000 | speed-comparison | read-range-stream | MiniExcel | 130.80 ms | 185.5 MB |  | Sylvan.Data.Excel | 66.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | EPPlus | 161.47 ms | 171.9 MB |  | Sylvan.Data.Excel | 105.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | EPPlus 4.5.3.3 | 194.27 ms | 0 B |  | Sylvan.Data.Excel | 147.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | ClosedXML | 342.39 ms | 194.2 MB |  | Sylvan.Data.Excel | 336.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | Sylvan.Data.Excel | 0.43 ms | 365.5 KB |  | Sylvan.Data.Excel | 78.8% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range | MiniExcel | 0.93 ms | 902.0 KB |  | Sylvan.Data.Excel | 54.1% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range | OfficeIMO.Excel | 2.02 ms | 1.1 MB |  | Sylvan.Data.Excel | Loss +370.9% |
| 25000 | speed-comparison | read-top-range | EPPlus | 128.29 ms | 92.0 MB |  | Sylvan.Data.Excel | 6254.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | EPPlus 4.5.3.3 | 168.82 ms | 0 B |  | Sylvan.Data.Excel | 8261.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | ClosedXML | 337.49 ms | 180.4 MB |  | Sylvan.Data.Excel | 16615.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | Sylvan.Data.Excel | 0.40 ms | 365.5 KB |  | Sylvan.Data.Excel | 81.0% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | MiniExcel | 0.81 ms | 901.8 KB |  | Sylvan.Data.Excel | 61.2% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | OfficeIMO.Excel | 2.09 ms | 1.1 MB |  | Sylvan.Data.Excel | Loss +427.2% |
| 25000 | speed-comparison | read-top-range-stream | EPPlus | 126.90 ms | 92.0 MB |  | Sylvan.Data.Excel | 5971.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | EPPlus 4.5.3.3 | 163.48 ms | 0 B |  | Sylvan.Data.Excel | 7721.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | ClosedXML | 329.82 ms | 180.4 MB |  | Sylvan.Data.Excel | 15680.3% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | Sylvan.Data.Excel | 30.52 ms | 2.4 MB |  | Sylvan.Data.Excel | 57.0% faster than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | OfficeIMO.Excel | 71.03 ms | 25.4 MB |  | Sylvan.Data.Excel | Loss +132.7% |
| 25000 | speed-comparison | shared-string-read | MiniExcel | 84.00 ms | 93.8 MB |  | Sylvan.Data.Excel | 18.3% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | EPPlus 4.5.3.3 | 94.46 ms | 0 B |  | Sylvan.Data.Excel | 33.0% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | EPPlus | 111.82 ms | 96.2 MB |  | Sylvan.Data.Excel | 57.4% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | ClosedXML | 266.43 ms | 109.4 MB |  | Sylvan.Data.Excel | 275.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | MiniExcel | 85.78 ms | 122.6 MB |  | MiniExcel | 10.0% faster than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | OfficeIMO.Excel | 95.32 ms | 32.5 MB |  | MiniExcel | Loss +11.1% |
| 25000 | speed-comparison | write-bulk-report | EPPlus | 382.40 ms | 249.0 MB |  | MiniExcel | 301.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | EPPlus 4.5.3.3 | 589.05 ms | 0 B |  | MiniExcel | 518.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | ClosedXML | 949.72 ms | 552.7 MB |  | MiniExcel | 896.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 42.90 ms | 18.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | MiniExcel | 68.98 ms | 91.1 MB |  | OfficeIMO.Excel | 60.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 294.17 ms | 0 B |  | OfficeIMO.Excel | 585.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus | 326.43 ms | 206.9 MB |  | OfficeIMO.Excel | 660.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | ClosedXML | 333.93 ms | 205.7 MB |  | OfficeIMO.Excel | 678.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | OfficeIMO.Excel | 90.01 ms | 25.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datareader-table | MiniExcel | 98.91 ms | 90.0 MB |  | OfficeIMO.Excel | 9.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | EPPlus | 354.24 ms | 114.7 MB |  | OfficeIMO.Excel | 293.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | ClosedXML | 515.32 ms | 169.3 MB |  | OfficeIMO.Excel | 472.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table-direct | EPPlus 4.5.3.3 | 281.55 ms | 0 B |  | EPPlus 4.5.3.3 |  |
| 25000 | speed-comparison | write-dataset-headerless-tables | OfficeIMO.Excel | 49.32 ms | 16.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-headerless-tables | MiniExcel | 94.37 ms | 94.8 MB |  | OfficeIMO.Excel | 91.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-headerless-tables | EPPlus | 423.09 ms | 108.6 MB |  | OfficeIMO.Excel | 757.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-headerless-tables | ClosedXML | 465.36 ms | 168.0 MB |  | OfficeIMO.Excel | 843.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-tables | OfficeIMO.Excel | 63.65 ms | 22.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-sparse-tables | MiniExcel | 125.81 ms | 105.6 MB |  | OfficeIMO.Excel | 97.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-tables | EPPlus | 544.87 ms | 132.5 MB |  | OfficeIMO.Excel | 756.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-tables | ClosedXML | 717.71 ms | 273.8 MB |  | OfficeIMO.Excel | 1027.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | OfficeIMO.Excel | 43.40 ms | 13.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-tables | MiniExcel | 85.18 ms | 94.8 MB |  | OfficeIMO.Excel | 96.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | EPPlus 4.5.3.3 | 279.62 ms | 0 B |  | OfficeIMO.Excel | 544.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | EPPlus | 325.72 ms | 108.2 MB |  | OfficeIMO.Excel | 650.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | ClosedXML | 455.16 ms | 168.0 MB |  | OfficeIMO.Excel | 948.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables-autofit | OfficeIMO.Excel | 45.41 ms | 13.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-tables-autofit | MiniExcel | 111.39 ms | 125.9 MB |  | OfficeIMO.Excel | 145.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables-autofit | EPPlus | 406.05 ms | 190.8 MB |  | OfficeIMO.Excel | 794.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables-autofit | ClosedXML | 908.10 ms | 537.2 MB |  | OfficeIMO.Excel | 1899.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | OfficeIMO.Excel | 46.54 ms | 14.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datatable-direct | MiniExcel | 92.57 ms | 90.2 MB |  | OfficeIMO.Excel | 98.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | EPPlus 4.5.3.3 | 282.34 ms | 0 B |  | OfficeIMO.Excel | 506.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | ClosedXML | 318.75 ms | 101.8 MB |  | OfficeIMO.Excel | 584.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | EPPlus | 319.02 ms | 114.7 MB |  | OfficeIMO.Excel | 585.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | OfficeIMO.Excel | 47.53 ms | 15.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datatable-table-direct | MiniExcel | 102.05 ms | 90.2 MB |  | OfficeIMO.Excel | 114.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | EPPlus 4.5.3.3 | 306.70 ms | 0 B |  | OfficeIMO.Excel | 545.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | EPPlus | 325.46 ms | 114.7 MB |  | OfficeIMO.Excel | 584.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | ClosedXML | 487.97 ms | 169.3 MB |  | OfficeIMO.Excel | 926.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 37.11 ms | 15.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | MiniExcel | 64.01 ms | 91.1 MB |  | OfficeIMO.Excel | 72.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 254.84 ms | 0 B |  | OfficeIMO.Excel | 586.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus | 263.66 ms | 114.7 MB |  | OfficeIMO.Excel | 610.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | ClosedXML | 284.23 ms | 101.8 MB |  | OfficeIMO.Excel | 665.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | OfficeIMO.Excel | 34.40 ms | 15.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-insertobjects-direct | MiniExcel | 70.24 ms | 91.1 MB |  | OfficeIMO.Excel | 104.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | EPPlus | 267.17 ms | 114.7 MB |  | OfficeIMO.Excel | 676.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | EPPlus 4.5.3.3 | 273.84 ms | 0 B |  | OfficeIMO.Excel | 696.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | ClosedXML | 293.01 ms | 101.8 MB |  | OfficeIMO.Excel | 751.9% slower than OfficeIMO |
| 300000 | focused-package-profile | write-blog-2023-20-string-columns | LargeXlsx | 636.47 ms | 93.1 MB | 28.6 MB | LargeXlsx | Win |
| 300000 | focused-package-profile | write-blog-2023-20-string-columns | OfficeIMO.Excel | 701.65 ms | 173.4 MB | 26.6 MB | LargeXlsx | 1.10x vs best |
| 300000 | focused-package-profile | write-blog-2023-20-string-columns | MiniExcel | 2.28 s | 2.46 GB | 28.5 MB | LargeXlsx | 3.58x vs best |
| 300000 | focused-package-profile | write-blog-2023-20-string-columns | ClosedXML | 15.89 s | 8.51 GB | 31.0 MB | LargeXlsx | 24.97x vs best |
