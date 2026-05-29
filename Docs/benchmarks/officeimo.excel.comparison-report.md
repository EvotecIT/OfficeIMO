# OfficeIMO.Excel Benchmark Report

Generated: 2026-05-29T09:47:05.7621264Z
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
| 2500 | dense-helloworld-comparison | read | Other | 0 | 2 | dense-helloworld-read-range vs Sylvan.Data.Excel (1.52x) |
| 2500 | package-profile | package | Package size | 37 | 12 | write-datatable-direct vs LargeXlsx (2.03x) |
| 2500 | speed-comparison | mutate | AutoFit and mutation | 1 | 0 |  |
| 2500 | speed-comparison | other | Object projection | 2 | 0 |  |
| 2500 | speed-comparison | other | Range and table read | 4 | 0 |  |
| 2500 | speed-comparison | other | Real-world feature mix | 6 | 0 |  |
| 2500 | speed-comparison | other | Real-world report | 2 | 0 |  |
| 2500 | speed-comparison | other | Report workbook | 4 | 0 |  |
| 2500 | speed-comparison | read | Formula write/read | 1 | 0 |  |
| 2500 | speed-comparison | read | Other | 0 | 3 | large-sparse-row-read vs Sylvan.Data.Excel (2.40x) |
| 2500 | speed-comparison | read | Range and table read | 1 | 6 | read-top-range vs Sylvan.Data.Excel (3.02x) |
| 2500 | speed-comparison | read | Streaming read | 0 | 4 | read-top-range-stream-small-chunks vs Sylvan.Data.Excel (3.49x) |
| 2500 | speed-comparison | read | Typed object read | 0 | 2 | read-objects-stream vs Sylvan.Data.Excel (1.18x) |
| 2500 | speed-comparison | write | AutoFit and mutation | 5 | 0 |  |
| 2500 | speed-comparison | write | Cell writer | 7 | 0 |  |
| 2500 | speed-comparison | write | DataSet table export | 3 | 0 |  |
| 2500 | speed-comparison | write | DataTable table export | 4 | 0 |  |
| 2500 | speed-comparison | write | Formatted report write | 1 | 0 |  |
| 2500 | speed-comparison | write | Formula write/read | 1 | 0 |  |
| 2500 | speed-comparison | write | Other | 2 | 2 | write-powershell-psobject-mixed-direct vs LargeXlsx (1.24x) |
| 2500 | speed-comparison | write | Plain cell export | 1 | 3 | append-plain-rows vs LargeXlsx (1.81x) |
| 2500 | speed-comparison | write | Plain streaming export | 2 | 0 |  |
| 2500 | speed-comparison | write | Plain string export | 1 | 0 |  |
| 2500 | speed-comparison | write | Shared string write | 4 | 0 |  |
| 2500 | speed-comparison | write | Typed object export | 0 | 3 | write-insertobjects-flat-dictionaries-direct vs LargeXlsx (1.48x) |
| 10000 | focused-package-profile | package | Package size | 1 | 0 |  |
| 25000 | dense-helloworld-comparison | read | Other | 1 | 1 | dense-helloworld-read-stream vs Sylvan.Data.Excel (1.23x) |
| 25000 | package-profile | package | Package size | 37 | 12 | write-insertobjects-legacy-dictionaries-direct vs LargeXlsx (1.66x) |
| 25000 | speed-comparison | mutate | AutoFit and mutation | 1 | 0 |  |
| 25000 | speed-comparison | other | Object projection | 2 | 0 |  |
| 25000 | speed-comparison | other | Range and table read | 4 | 0 |  |
| 25000 | speed-comparison | other | Real-world feature mix | 6 | 0 |  |
| 25000 | speed-comparison | other | Real-world report | 2 | 0 |  |
| 25000 | speed-comparison | other | Report workbook | 4 | 0 |  |
| 25000 | speed-comparison | read | Formula write/read | 1 | 0 |  |
| 25000 | speed-comparison | read | Other | 0 | 3 | large-sparse-row-read vs Sylvan.Data.Excel (1.74x) |
| 25000 | speed-comparison | read | Range and table read | 4 | 3 | read-top-range vs Sylvan.Data.Excel (4.20x) |
| 25000 | speed-comparison | read | Streaming read | 2 | 2 | read-top-range-stream vs Sylvan.Data.Excel (4.13x) |
| 25000 | speed-comparison | read | Typed object read | 0 | 2 | read-objects vs Sylvan.Data.Excel (1.05x) |
| 25000 | speed-comparison | write | AutoFit and mutation | 5 | 0 |  |
| 25000 | speed-comparison | write | Cell writer | 7 | 0 |  |
| 25000 | speed-comparison | write | DataSet table export | 3 | 0 |  |
| 25000 | speed-comparison | write | DataTable table export | 3 | 1 | write-datatable-direct vs LargeXlsx (1.16x) |
| 25000 | speed-comparison | write | Formatted report write | 1 | 0 |  |
| 25000 | speed-comparison | write | Formula write/read | 1 | 0 |  |
| 25000 | speed-comparison | write | Other | 2 | 2 | write-powershell-mixed-objects-direct vs LargeXlsx (1.17x) |
| 25000 | speed-comparison | write | Plain cell export | 2 | 2 | append-plain-rows vs LargeXlsx (1.51x) |
| 25000 | speed-comparison | write | Plain streaming export | 1 | 1 | write-datareader-plain vs Sylvan.Data.Excel (1.28x) |
| 25000 | speed-comparison | write | Plain string export | 0 | 1 | write-blog-2023-20-string-columns vs LargeXlsx (1.07x) |
| 25000 | speed-comparison | write | Shared string write | 4 | 0 |  |
| 25000 | speed-comparison | write | Typed object export | 0 | 3 | write-insertobjects-flat-dictionaries-direct vs LargeXlsx (1.36x) |
| 300000 | focused-package-profile | package | Package size | 0 | 1 | write-blog-2023-20-string-columns vs LargeXlsx (1.10x) |

## Rows

| Rows | Artifact | Scenario | Library | Mean | Allocated | Package | Best | Outcome |
| ---: | --- | --- | --- | ---: | ---: | ---: | --- | --- |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | Sylvan.Data.Excel | 4.04 ms | 362.3 KB |  | Sylvan.Data.Excel | 34.2% faster than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | OfficeIMO.Excel | 6.14 ms | 2.4 MB |  | Sylvan.Data.Excel | Loss +51.9% |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | ExcelDataReader | 12.37 ms | 6.7 MB |  | Sylvan.Data.Excel | 101.6% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | MiniExcel | 15.37 ms | 21.0 MB |  | Sylvan.Data.Excel | 150.6% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | Sylvan.Data.Excel | 4.19 ms | 362.3 KB |  | Sylvan.Data.Excel | 24.5% faster than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | OfficeIMO.Excel | 5.54 ms | 2.5 MB |  | Sylvan.Data.Excel | Loss +32.4% |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | ExcelDataReader | 11.27 ms | 6.7 MB |  | Sylvan.Data.Excel | 103.4% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | MiniExcel | 14.82 ms | 21.0 MB |  | Sylvan.Data.Excel | 167.5% slower than OfficeIMO |
| 2500 | package-profile | append-plain-rows | LargeXlsx | 1.53 ms | 288.4 KB | 63.1 KB | LargeXlsx | 37.6% faster than OfficeIMO |
| 2500 | package-profile | append-plain-rows | OfficeIMO.Excel | 2.45 ms | 1.6 MB | 64.5 KB | LargeXlsx | Loss +60.3% |
| 2500 | package-profile | append-plain-rows | MiniExcel | 4.10 ms | 19.2 MB | 68.1 KB | LargeXlsx | 67.3% slower than OfficeIMO |
| 2500 | package-profile | append-plain-rows | ClosedXML | 15.92 ms | 10.9 MB | 59.8 KB | LargeXlsx | 550.1% slower than OfficeIMO |
| 2500 | package-profile | append-plain-rows | EPPlus | 28.54 ms | 13.9 MB | 56.9 KB | LargeXlsx | 1065.4% slower than OfficeIMO |
| 2500 | package-profile | autofit-existing | OfficeIMO.Excel | 44.54 ms | 13.6 MB | 143.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | autofit-existing | EPPlus | 99.96 ms | 49.5 MB | 115.0 KB | OfficeIMO.Excel | 124.4% slower than OfficeIMO |
| 2500 | package-profile | autofit-existing | ClosedXML | 133.33 ms | 82.4 MB | 121.0 KB | OfficeIMO.Excel | 199.3% slower than OfficeIMO |
| 2500 | package-profile | large-shared-strings | OfficeIMO.Excel | 2.01 ms | 2.1 MB | 55.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | large-shared-strings | MiniExcel | 4.39 ms | 20.6 MB | 60.7 KB | OfficeIMO.Excel | 118.2% slower than OfficeIMO |
| 2500 | package-profile | large-shared-strings | ClosedXML | 11.82 ms | 11.0 MB | 50.3 KB | OfficeIMO.Excel | 487.6% slower than OfficeIMO |
| 2500 | package-profile | large-shared-strings | EPPlus | 21.40 ms | 12.4 MB | 48.1 KB | OfficeIMO.Excel | 963.4% slower than OfficeIMO |
| 2500 | package-profile | realworld-autofilter | OfficeIMO.Excel | 3.57 ms | 1.1 MB | 142.4 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-autofilter | ClosedXML | 32.58 ms | 21.7 MB | 120.2 KB | OfficeIMO.Excel | 811.3% slower than OfficeIMO |
| 2500 | package-profile | realworld-autofilter | EPPlus | 53.85 ms | 24.1 MB | 114.2 KB | OfficeIMO.Excel | 1406.2% slower than OfficeIMO |
| 2500 | package-profile | realworld-charts | OfficeIMO.Excel | 5.25 ms | 1.6 MB | 147.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-charts | EPPlus | 48.96 ms | 26.4 MB | 117.0 KB | OfficeIMO.Excel | 832.8% slower than OfficeIMO |
| 2500 | package-profile | realworld-conditional-formatting | OfficeIMO.Excel | 4.40 ms | 1.2 MB | 142.7 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-conditional-formatting | ClosedXML | 33.31 ms | 21.7 MB | 120.3 KB | OfficeIMO.Excel | 657.5% slower than OfficeIMO |
| 2500 | package-profile | realworld-conditional-formatting | EPPlus | 54.02 ms | 24.1 MB | 114.3 KB | OfficeIMO.Excel | 1128.4% slower than OfficeIMO |
| 2500 | package-profile | realworld-data-validation | OfficeIMO.Excel | 3.54 ms | 1.2 MB | 142.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-data-validation | ClosedXML | 31.64 ms | 21.7 MB | 120.3 KB | OfficeIMO.Excel | 793.7% slower than OfficeIMO |
| 2500 | package-profile | realworld-data-validation | EPPlus | 47.78 ms | 24.1 MB | 114.2 KB | OfficeIMO.Excel | 1249.4% slower than OfficeIMO |
| 2500 | package-profile | realworld-freeze-panes | OfficeIMO.Excel | 3.86 ms | 1.1 MB | 142.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-freeze-panes | ClosedXML | 36.17 ms | 21.7 MB | 120.2 KB | OfficeIMO.Excel | 836.5% slower than OfficeIMO |
| 2500 | package-profile | realworld-freeze-panes | EPPlus | 53.22 ms | 24.1 MB | 114.3 KB | OfficeIMO.Excel | 1278.3% slower than OfficeIMO |
| 2500 | package-profile | realworld-pivot-table | OfficeIMO.Excel | 20.98 ms | 18.2 MB | 203.9 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-pivot-table | EPPlus | 56.12 ms | 28.8 MB | 117.4 KB | OfficeIMO.Excel | 167.5% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-all-in-one | OfficeIMO.Excel | 18.78 ms | 19.0 MB | 210.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-all-in-one | EPPlus | 65.51 ms | 53.2 MB | 121.8 KB | OfficeIMO.Excel | 248.8% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-core | OfficeIMO.Excel | 4.52 ms | 1.3 MB | 143.9 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | realworld-report-core | EPPlus | 68.92 ms | 46.1 MB | 115.5 KB | OfficeIMO.Excel | 1423.1% slower than OfficeIMO |
| 2500 | package-profile | realworld-report-core | ClosedXML | 88.23 ms | 68.2 MB | 121.5 KB | OfficeIMO.Excel | 1849.9% slower than OfficeIMO |
| 2500 | package-profile | report-workbook | OfficeIMO.Excel | 14.44 ms | 11.9 MB | 90.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | report-workbook | EPPlus | 100.13 ms | 75.6 MB | 161.8 KB | OfficeIMO.Excel | 593.5% slower than OfficeIMO |
| 2500 | package-profile | report-workbook-core | OfficeIMO.Excel | 6.18 ms | 2.3 MB | 187.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | report-workbook-core | EPPlus | 117.65 ms | 70.2 MB | 157.2 KB | OfficeIMO.Excel | 1802.5% slower than OfficeIMO |
| 2500 | package-profile | report-workbook-core | ClosedXML | 118.34 ms | 94.9 MB | 165.1 KB | OfficeIMO.Excel | 1813.7% slower than OfficeIMO |
| 2500 | package-profile | report-workbook-datatable | OfficeIMO.Excel | 16.45 ms | 12.2 MB | 90.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | report-workbook-datatable | EPPlus | 98.05 ms | 64.4 MB | 161.8 KB | OfficeIMO.Excel | 495.9% slower than OfficeIMO |
| 2500 | package-profile | report-workbook-datatable-core | OfficeIMO.Excel | 6.33 ms | 2.6 MB | 187.5 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | report-workbook-datatable-core | EPPlus | 97.29 ms | 59.0 MB | 157.2 KB | OfficeIMO.Excel | 1436.6% slower than OfficeIMO |
| 2500 | package-profile | report-workbook-datatable-core | ClosedXML | 112.16 ms | 80.9 MB | 165.1 KB | OfficeIMO.Excel | 1671.5% slower than OfficeIMO |
| 2500 | package-profile | write-blog-2023-20-string-columns | LargeXlsx | 4.37 ms | 849.6 KB | 237.7 KB | LargeXlsx | 12.5% faster than OfficeIMO |
| 2500 | package-profile | write-blog-2023-20-string-columns | OfficeIMO.Excel | 5.00 ms | 1.5 MB | 216.7 KB | LargeXlsx | Loss +14.3% |
| 2500 | package-profile | write-blog-2023-20-string-columns | MiniExcel | 18.06 ms | 35.1 MB | 235.3 KB | LargeXlsx | 261.4% slower than OfficeIMO |
| 2500 | package-profile | write-blog-2023-20-string-columns | ClosedXML | 103.58 ms | 69.8 MB | 257.2 KB | LargeXlsx | 1973.2% slower than OfficeIMO |
| 2500 | package-profile | write-bulk-report | OfficeIMO.Excel | 4.78 ms | 1.2 MB | 143.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-bulk-report | MiniExcel | 9.96 ms | 26.2 MB | 153.8 KB | OfficeIMO.Excel | 108.2% slower than OfficeIMO |
| 2500 | package-profile | write-bulk-report | EPPlus | 76.38 ms | 46.0 MB | 115.0 KB | OfficeIMO.Excel | 1496.4% slower than OfficeIMO |
| 2500 | package-profile | write-bulk-report | ClosedXML | 93.49 ms | 57.0 MB | 121.0 KB | OfficeIMO.Excel | 1853.9% slower than OfficeIMO |
| 2500 | package-profile | write-cellformula | OfficeIMO.Excel | 2.27 ms | 1.1 MB | 66.6 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellformula | ClosedXML | 16.56 ms | 11.7 MB | 70.6 KB | OfficeIMO.Excel | 630.4% slower than OfficeIMO |
| 2500 | package-profile | write-cellformula | EPPlus | 39.21 ms | 17.6 MB | 62.1 KB | OfficeIMO.Excel | 1629.2% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-empty-strings | OfficeIMO.Excel | 2.40 ms | 1.4 MB | 44.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-empty-strings | ClosedXML | 14.84 ms | 9.7 MB | 44.9 KB | OfficeIMO.Excel | 517.8% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-empty-strings | EPPlus | 27.33 ms | 11.4 MB | 42.0 KB | OfficeIMO.Excel | 1037.8% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-numbers | OfficeIMO.Excel | 2.01 ms | 946.8 KB | 47.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-numbers | ClosedXML | 11.14 ms | 9.0 MB | 45.9 KB | OfficeIMO.Excel | 453.9% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-numbers | EPPlus | 23.22 ms | 12.5 MB | 43.7 KB | OfficeIMO.Excel | 1054.3% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-mixed | OfficeIMO.Excel | 3.14 ms | 1.4 MB | 61.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-object-mixed | ClosedXML | 19.34 ms | 11.6 MB | 59.5 KB | OfficeIMO.Excel | 515.7% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-mixed | EPPlus | 31.64 ms | 15.2 MB | 58.9 KB | OfficeIMO.Excel | 907.2% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-sparse | OfficeIMO.Excel | 2.60 ms | 1.2 MB | 62.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-object-sparse | ClosedXML | 18.09 ms | 11.0 MB | 52.5 KB | OfficeIMO.Excel | 596.5% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-sparse | EPPlus | 24.91 ms | 14.5 MB | 54.2 KB | OfficeIMO.Excel | 859.2% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 2.27 ms | 1.2 MB | 62.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-object-sparse-batch | ClosedXML | 14.77 ms | 11.0 MB | 52.5 KB | OfficeIMO.Excel | 551.5% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-object-sparse-batch | EPPlus | 24.83 ms | 14.5 MB | 54.2 KB | OfficeIMO.Excel | 995.4% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-scalars | OfficeIMO.Excel | 1.85 ms | 964.9 KB | 46.3 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-scalars | ClosedXML | 11.94 ms | 8.8 MB | 45.4 KB | OfficeIMO.Excel | 546.9% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-scalars | EPPlus | 24.55 ms | 12.5 MB | 42.4 KB | OfficeIMO.Excel | 1229.5% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings | OfficeIMO.Excel | 2.87 ms | 2.2 MB | 55.1 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-strings | ClosedXML | 12.04 ms | 11.0 MB | 50.3 KB | OfficeIMO.Excel | 319.0% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings | EPPlus | 25.70 ms | 12.4 MB | 48.1 KB | OfficeIMO.Excel | 794.6% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings-distinct | OfficeIMO.Excel | 2.39 ms | 2.2 MB | 51.8 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-strings-distinct | ClosedXML | 16.13 ms | 12.8 MB | 61.9 KB | OfficeIMO.Excel | 576.1% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings-distinct | EPPlus | 24.08 ms | 13.5 MB | 61.5 KB | OfficeIMO.Excel | 909.4% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings-repeated | OfficeIMO.Excel | 2.26 ms | 1.2 MB | 40.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-strings-repeated | ClosedXML | 13.11 ms | 9.0 MB | 38.8 KB | OfficeIMO.Excel | 479.3% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-strings-repeated | EPPlus | 20.78 ms | 11.0 MB | 34.8 KB | OfficeIMO.Excel | 817.9% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-temporal | OfficeIMO.Excel | 2.49 ms | 1.2 MB | 63.3 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-cellvalue-temporal | ClosedXML | 15.12 ms | 9.5 MB | 54.5 KB | OfficeIMO.Excel | 508.6% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalue-temporal | EPPlus | 25.99 ms | 14.3 MB | 53.1 KB | OfficeIMO.Excel | 945.8% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 1.23 ms | 439.0 KB | 47.3 KB | LargeXlsx | 31.4% faster than OfficeIMO |
| 2500 | package-profile | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 1.79 ms | 923.6 KB | 48.2 KB | LargeXlsx | Loss +45.7% |
| 2500 | package-profile | write-cellvalues-headerless-rectangle-direct | ClosedXML | 12.80 ms | 10.0 MB | 53.0 KB | LargeXlsx | 615.0% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-headerless-rectangle-direct | EPPlus | 24.08 ms | 12.7 MB | 52.5 KB | LargeXlsx | 1245.5% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | LargeXlsx | 3.26 ms | 750.2 KB | 138.4 KB | LargeXlsx | 21.4% faster than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 4.15 ms | 1.7 MB | 138.0 KB | LargeXlsx | Loss +27.3% |
| 2500 | package-profile | write-cellvalues-rectangle-direct | MiniExcel | 8.04 ms | 22.7 MB | 153.7 KB | LargeXlsx | 93.8% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | ClosedXML | 32.29 ms | 21.7 MB | 120.1 KB | LargeXlsx | 678.7% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | EPPlus | 51.02 ms | 24.0 MB | 114.1 KB | LargeXlsx | 1130.5% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | Sylvan.Data.Excel | 4.32 ms | 750.7 KB | 78.5 KB | Sylvan.Data.Excel | 9.3% faster than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | OfficeIMO.Excel | 4.76 ms | 1.4 MB | 138.0 KB | Sylvan.Data.Excel | Loss +10.2% |
| 2500 | package-profile | write-datareader-plain | LargeXlsx | 5.08 ms | 1.0 MB | 138.4 KB | Sylvan.Data.Excel | 6.8% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | MiniExcel | 9.66 ms | 22.5 MB | 153.6 KB | Sylvan.Data.Excel | 103.1% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | ClosedXML | 31.62 ms | 11.3 MB | 120.1 KB | Sylvan.Data.Excel | 564.6% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-plain | EPPlus | 44.99 ms | 16.2 MB | 114.9 KB | Sylvan.Data.Excel | 845.7% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table | OfficeIMO.Excel | 7.12 ms | 1.4 MB | 138.8 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-datareader-table | MiniExcel | 7.61 ms | 22.5 MB | 153.6 KB | OfficeIMO.Excel | 6.8% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table | ClosedXML | 36.72 ms | 18.6 MB | 120.9 KB | OfficeIMO.Excel | 415.4% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table | EPPlus | 46.27 ms | 16.2 MB | 114.9 KB | OfficeIMO.Excel | 549.4% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table-autofit | OfficeIMO.Excel | 4.84 ms | 1.4 MB | 138.9 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-datareader-table-autofit | MiniExcel | 8.33 ms | 26.0 MB | 153.8 KB | OfficeIMO.Excel | 72.2% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table-autofit | EPPlus | 68.76 ms | 37.4 MB | 115.1 KB | OfficeIMO.Excel | 1320.6% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table-autofit | ClosedXML | 94.38 ms | 57.0 MB | 121.0 KB | OfficeIMO.Excel | 1850.1% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-direct-export | LargeXlsx | 3.96 ms | 1.1 MB | 164.2 KB | LargeXlsx | 10.4% faster than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-direct-export | OfficeIMO.Excel | 4.42 ms | 1.6 MB | 131.1 KB | LargeXlsx | Loss +11.7% |
| 2500 | package-profile | write-dataset-sparse-direct-export | MiniExcel | 9.48 ms | 29.0 MB | 180.5 KB | LargeXlsx | 114.5% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-direct-export | EPPlus | 54.50 ms | 21.3 MB | 144.5 KB | LargeXlsx | 1132.8% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-direct-export | ClosedXML | 64.55 ms | 26.8 MB | 159.4 KB | LargeXlsx | 1360.2% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-tables | OfficeIMO.Excel | 4.90 ms | 2.3 MB | 176.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-dataset-sparse-tables | MiniExcel | 10.77 ms | 29.0 MB | 180.5 KB | OfficeIMO.Excel | 119.8% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-tables | EPPlus | 58.49 ms | 21.3 MB | 144.5 KB | OfficeIMO.Excel | 1094.3% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-sparse-tables | ClosedXML | 65.46 ms | 26.8 MB | 159.4 KB | OfficeIMO.Excel | 1236.6% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables | OfficeIMO.Excel | 3.86 ms | 1.5 MB | 138.9 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-dataset-tables | MiniExcel | 7.89 ms | 28.0 MB | 156.4 KB | OfficeIMO.Excel | 104.3% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables | EPPlus | 38.26 ms | 18.2 MB | 116.6 KB | OfficeIMO.Excel | 891.0% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables | ClosedXML | 40.02 ms | 18.4 MB | 123.4 KB | OfficeIMO.Excel | 936.7% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables-autofit | OfficeIMO.Excel | 4.73 ms | 1.6 MB | 139.2 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-dataset-tables-autofit | MiniExcel | 9.20 ms | 31.0 MB | 156.6 KB | OfficeIMO.Excel | 94.7% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables-autofit | EPPlus | 74.88 ms | 40.4 MB | 116.9 KB | OfficeIMO.Excel | 1483.8% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables-autofit | ClosedXML | 79.77 ms | 55.4 MB | 123.7 KB | OfficeIMO.Excel | 1587.2% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | LargeXlsx | 3.54 ms | 1.1 MB | 138.4 KB | LargeXlsx | 50.6% faster than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | OfficeIMO.Excel | 7.16 ms | 1.4 MB | 138.0 KB | LargeXlsx | Loss +102.5% |
| 2500 | package-profile | write-datatable-direct | MiniExcel | 8.41 ms | 22.5 MB | 153.7 KB | LargeXlsx | 17.5% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | ClosedXML | 33.78 ms | 11.3 MB | 120.1 KB | LargeXlsx | 371.7% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | EPPlus | 40.32 ms | 16.2 MB | 114.9 KB | LargeXlsx | 463.2% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-table-direct | OfficeIMO.Excel | 4.56 ms | 1.4 MB | 138.8 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-datatable-table-direct | MiniExcel | 8.42 ms | 22.5 MB | 153.7 KB | OfficeIMO.Excel | 84.7% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-table-direct | ClosedXML | 39.21 ms | 18.6 MB | 120.9 KB | OfficeIMO.Excel | 759.9% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-table-direct | EPPlus | 39.60 ms | 16.2 MB | 114.9 KB | OfficeIMO.Excel | 768.4% slower than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | LargeXlsx | 3.38 ms | 750.2 KB | 138.4 KB | LargeXlsx | 25.3% faster than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 4.52 ms | 1.4 MB | 138.0 KB | LargeXlsx | Loss +33.9% |
| 2500 | package-profile | write-fluent-rowsfrom-direct | MiniExcel | 11.68 ms | 22.7 MB | 153.7 KB | LargeXlsx | 158.1% slower than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | ClosedXML | 43.17 ms | 11.3 MB | 120.1 KB | LargeXlsx | 854.3% slower than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | EPPlus | 44.85 ms | 16.2 MB | 114.9 KB | LargeXlsx | 891.5% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.99 ms | 1.2 MB | 142.4 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 70.27 ms | 49.7 MB | 120.2 KB | OfficeIMO.Excel | 1308.3% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 72.92 ms | 37.4 MB | 115.1 KB | OfficeIMO.Excel | 1361.4% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | LargeXlsx | 3.70 ms | 750.2 KB | 138.4 KB | LargeXlsx | 8.0% faster than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | OfficeIMO.Excel | 4.02 ms | 1.1 MB | 142.3 KB | LargeXlsx | Loss +8.7% |
| 2500 | package-profile | write-insertobjects-direct | MiniExcel | 7.46 ms | 22.7 MB | 153.7 KB | LargeXlsx | 85.4% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | ClosedXML | 29.77 ms | 11.3 MB | 120.1 KB | LargeXlsx | 640.0% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | EPPlus | 42.31 ms | 16.2 MB | 114.9 KB | LargeXlsx | 952.0% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 5.71 ms | 1.1 MB | 138.1 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 73.08 ms | 37.4 MB | 115.1 KB | OfficeIMO.Excel | 1180.2% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 87.93 ms | 49.7 MB | 120.2 KB | OfficeIMO.Excel | 1440.3% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 3.05 ms | 750.2 KB | 138.4 KB | LargeXlsx | 27.0% faster than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 4.18 ms | 1.1 MB | 138.0 KB | LargeXlsx | Loss +36.9% |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | MiniExcel | 7.26 ms | 22.7 MB | 153.7 KB | LargeXlsx | 73.7% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | ClosedXML | 27.30 ms | 11.3 MB | 120.1 KB | LargeXlsx | 553.2% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-flat-dictionaries-direct | EPPlus | 40.65 ms | 16.2 MB | 114.9 KB | LargeXlsx | 872.9% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | LargeXlsx | 3.48 ms | 750.2 KB | 138.4 KB | LargeXlsx | 42.0% faster than OfficeIMO |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | OfficeIMO.Excel | 5.99 ms | 1.6 MB | 142.3 KB | LargeXlsx | Loss +72.3% |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | MiniExcel | 9.62 ms | 22.7 MB | 153.7 KB | LargeXlsx | 60.6% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | ClosedXML | 28.70 ms | 11.3 MB | 120.1 KB | LargeXlsx | 379.1% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-legacy-dictionaries-direct | EPPlus | 39.24 ms | 16.2 MB | 114.9 KB | LargeXlsx | 555.0% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 3.68 ms | 1.2 MB | 142.4 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 41.77 ms | 27.9 MB | 120.2 KB | OfficeIMO.Excel | 1035.7% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 56.45 ms | 26.6 MB | 115.0 KB | OfficeIMO.Excel | 1434.8% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-mixed-objects-direct | LargeXlsx | 4.05 ms | 794.5 KB | 182.6 KB | LargeXlsx | 32.2% faster than OfficeIMO |
| 2500 | package-profile | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 5.97 ms | 2.0 MB | 183.1 KB | LargeXlsx | Loss +47.6% |
| 2500 | package-profile | write-powershell-mixed-objects-direct | MiniExcel | 8.42 ms | 24.6 MB | 194.0 KB | LargeXlsx | 40.9% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-mixed-objects-direct | EPPlus | 53.40 ms | 19.6 MB | 152.1 KB | LargeXlsx | 793.9% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-mixed-objects-direct | ClosedXML | 54.32 ms | 16.6 MB | 161.0 KB | LargeXlsx | 809.3% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 4.46 ms | 1.3 MB | 182.4 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-powershell-psobject-mixed-direct | LargeXlsx | 4.73 ms | 794.5 KB | 182.6 KB | OfficeIMO.Excel | 6.1% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-mixed-direct | MiniExcel | 9.33 ms | 24.6 MB | 194.0 KB | OfficeIMO.Excel | 109.3% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-mixed-direct | ClosedXML | 45.15 ms | 16.6 MB | 161.0 KB | OfficeIMO.Excel | 912.7% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-mixed-direct | EPPlus | 52.58 ms | 19.6 MB | 152.1 KB | OfficeIMO.Excel | 1079.3% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 20.71 ms | 4.2 MB | 651.0 KB | OfficeIMO.Excel | Win |
| 2500 | package-profile | write-powershell-psobject-wide-direct | LargeXlsx | 25.80 ms | 2.7 MB | 644.6 KB | OfficeIMO.Excel | 24.6% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-wide-direct | MiniExcel | 40.03 ms | 47.3 MB | 674.4 KB | OfficeIMO.Excel | 93.3% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-wide-direct | ClosedXML | 141.94 ms | 50.4 MB | 615.5 KB | OfficeIMO.Excel | 585.4% slower than OfficeIMO |
| 2500 | package-profile | write-powershell-psobject-wide-direct | EPPlus | 181.38 ms | 67.5 MB | 548.9 KB | OfficeIMO.Excel | 775.8% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | LargeXlsx | 1.37 ms | 288.4 KB |  | LargeXlsx | 44.6% faster than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | OfficeIMO.Excel | 2.47 ms | 1.6 MB |  | LargeXlsx | Loss +80.6% |
| 2500 | speed-comparison | append-plain-rows | MiniExcel | 3.88 ms | 19.2 MB |  | LargeXlsx | 56.9% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | ClosedXML | 13.41 ms | 10.9 MB |  | LargeXlsx | 442.7% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | EPPlus 4.5.3.3 | 20.76 ms | 0 B |  | LargeXlsx | 740.1% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | EPPlus | 22.93 ms | 13.9 MB |  | LargeXlsx | 827.6% slower than OfficeIMO |
| 2500 | speed-comparison | autofit-existing | OfficeIMO.Excel | 28.07 ms | 13.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | autofit-existing | EPPlus | 70.35 ms | 49.5 MB |  | OfficeIMO.Excel | 150.6% slower than OfficeIMO |
| 2500 | speed-comparison | autofit-existing | EPPlus 4.5.3.3 | 101.48 ms | 0 B |  | OfficeIMO.Excel | 261.5% slower than OfficeIMO |
| 2500 | speed-comparison | autofit-existing | ClosedXML | 116.48 ms | 82.6 MB |  | OfficeIMO.Excel | 314.9% slower than OfficeIMO |
| 2500 | speed-comparison | build-object-datatable-dictionaries | OfficeIMO.Excel | 1.57 ms | 564.2 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | build-object-datatable-typed | OfficeIMO.Excel | 1.34 ms | 856.9 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-cells | OfficeIMO.Excel | 7.92 ms | 2.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-cells | EPPlus | 27.96 ms | 19.7 MB |  | OfficeIMO.Excel | 252.9% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-cells | ClosedXML | 29.81 ms | 16.6 MB |  | OfficeIMO.Excel | 276.3% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-first-column-from-wide-sheet | OfficeIMO.Excel | 4.83 ms | 643.4 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-first-column-from-wide-sheet | EPPlus | 23.00 ms | 12.8 MB |  | OfficeIMO.Excel | 375.9% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-first-column-from-wide-sheet | ClosedXML | 30.15 ms | 15.1 MB |  | OfficeIMO.Excel | 523.8% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-range | OfficeIMO.Excel | 8.36 ms | 2.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-range | EPPlus | 27.50 ms | 19.7 MB |  | OfficeIMO.Excel | 229.0% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-range | ClosedXML | 28.87 ms | 16.6 MB |  | OfficeIMO.Excel | 245.4% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-top-range | OfficeIMO.Excel | 1.69 ms | 402.6 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | enumerate-top-range | EPPlus | 23.03 ms | 12.1 MB |  | OfficeIMO.Excel | 1261.5% slower than OfficeIMO |
| 2500 | speed-comparison | enumerate-top-range | ClosedXML | 28.74 ms | 15.0 MB |  | OfficeIMO.Excel | 1599.0% slower than OfficeIMO |
| 2500 | speed-comparison | formula-heavy-read | OfficeIMO.Excel | 3.33 ms | 777.4 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | formula-heavy-read | ClosedXML | 16.25 ms | 8.1 MB |  | OfficeIMO.Excel | 388.1% slower than OfficeIMO |
| 2500 | speed-comparison | formula-heavy-read | EPPlus 4.5.3.3 | 17.44 ms | 0 B |  | OfficeIMO.Excel | 423.9% slower than OfficeIMO |
| 2500 | speed-comparison | formula-heavy-read | EPPlus | 18.86 ms | 7.5 MB |  | OfficeIMO.Excel | 466.5% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | OfficeIMO.Excel | 2.09 ms | 2.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | large-shared-strings | MiniExcel | 3.79 ms | 20.6 MB |  | OfficeIMO.Excel | 81.1% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | ClosedXML | 11.41 ms | 11.0 MB |  | OfficeIMO.Excel | 444.7% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | EPPlus 4.5.3.3 | 13.84 ms | 0 B |  | OfficeIMO.Excel | 560.8% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | EPPlus | 20.79 ms | 12.4 MB |  | OfficeIMO.Excel | 892.6% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | Sylvan.Data.Excel | 1.01 ms | 316.6 KB |  | Sylvan.Data.Excel | 42.0% faster than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | ExcelDataReader | 1.51 ms | 4.0 MB |  | Sylvan.Data.Excel | 13.0% faster than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | OfficeIMO.Excel | 1.73 ms | 248.8 KB |  | Sylvan.Data.Excel | Loss +72.4% |
| 2500 | speed-comparison | large-sparse-column-read | ClosedXML | 3.27 ms | 4.3 MB |  | Sylvan.Data.Excel | 88.9% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | MiniExcel | 10.85 ms | 45.1 MB |  | Sylvan.Data.Excel | 526.1% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | EPPlus 4.5.3.3 | 11.98 ms | 0 B |  | Sylvan.Data.Excel | 591.7% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | EPPlus | 38.09 ms | 42.1 MB |  | Sylvan.Data.Excel | 2098.2% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | Sylvan.Data.Excel | 0.99 ms | 316.6 KB |  | Sylvan.Data.Excel | 58.3% faster than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | ExcelDataReader | 1.48 ms | 4.0 MB |  | Sylvan.Data.Excel | 38.0% faster than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | OfficeIMO.Excel | 2.38 ms | 248.9 KB |  | Sylvan.Data.Excel | Loss +139.6% |
| 2500 | speed-comparison | large-sparse-row-read | ClosedXML | 3.20 ms | 4.3 MB |  | Sylvan.Data.Excel | 34.1% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | EPPlus 4.5.3.3 | 10.94 ms | 0 B |  | Sylvan.Data.Excel | 358.8% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | MiniExcel | 11.27 ms | 45.1 MB |  | Sylvan.Data.Excel | 372.7% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | EPPlus | 39.79 ms | 42.1 MB |  | Sylvan.Data.Excel | 1568.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | Sylvan.Data.Excel | 4.43 ms | 655.2 KB |  | Sylvan.Data.Excel | 11.0% faster than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | OfficeIMO.Excel | 4.98 ms | 494.6 KB |  | Sylvan.Data.Excel | Loss +12.4% |
| 2500 | speed-comparison | read-bottom-range | ExcelDataReader | 10.87 ms | 5.9 MB |  | Sylvan.Data.Excel | 118.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | MiniExcel | 13.55 ms | 18.2 MB |  | Sylvan.Data.Excel | 171.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | EPPlus | 23.27 ms | 12.1 MB |  | Sylvan.Data.Excel | 366.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range | ClosedXML | 28.54 ms | 15.0 MB |  | Sylvan.Data.Excel | 472.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | Sylvan.Data.Excel | 4.37 ms | 655.2 KB |  | Sylvan.Data.Excel | 11.3% faster than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | OfficeIMO.Excel | 4.93 ms | 497.9 KB |  | Sylvan.Data.Excel | Loss +12.7% |
| 2500 | speed-comparison | read-bottom-range-stream | ExcelDataReader | 10.58 ms | 5.9 MB |  | Sylvan.Data.Excel | 114.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | MiniExcel | 13.35 ms | 18.2 MB |  | Sylvan.Data.Excel | 170.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | EPPlus | 22.38 ms | 12.1 MB |  | Sylvan.Data.Excel | 354.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-bottom-range-stream | ClosedXML | 29.16 ms | 15.0 MB |  | Sylvan.Data.Excel | 491.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | Sylvan.Data.Excel | 6.36 ms | 2.2 MB |  | Sylvan.Data.Excel | 25.0% faster than OfficeIMO |
| 2500 | speed-comparison | read-datatable | OfficeIMO.Excel | 8.48 ms | 3.6 MB |  | Sylvan.Data.Excel | Loss +33.3% |
| 2500 | speed-comparison | read-datatable | ExcelDataReader | 12.54 ms | 7.5 MB |  | Sylvan.Data.Excel | 47.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | MiniExcel | 13.53 ms | 17.8 MB |  | Sylvan.Data.Excel | 59.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | ClosedXML | 31.55 ms | 17.9 MB |  | Sylvan.Data.Excel | 272.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | EPPlus | 32.84 ms | 21.2 MB |  | Sylvan.Data.Excel | 287.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | EPPlus 4.5.3.3 | 38.04 ms | 0 B |  | Sylvan.Data.Excel | 348.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | Sylvan.Data.Excel | 4.81 ms | 733.5 KB |  | Sylvan.Data.Excel | 2.6% faster than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | OfficeIMO.Excel | 4.93 ms | 663.0 KB |  | Sylvan.Data.Excel | Loss +2.7% |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | MiniExcel | 10.33 ms | 15.5 MB |  | Sylvan.Data.Excel | 109.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | ExcelDataReader | 10.76 ms | 5.9 MB |  | Sylvan.Data.Excel | 118.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | EPPlus | 21.94 ms | 12.8 MB |  | Sylvan.Data.Excel | 344.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-first-column-from-wide-sheet | ClosedXML | 29.16 ms | 15.1 MB |  | Sylvan.Data.Excel | 490.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | Sylvan.Data.Excel | 6.80 ms | 895.3 KB |  | Sylvan.Data.Excel | 14.0% faster than OfficeIMO |
| 2500 | speed-comparison | read-objects | OfficeIMO.Excel | 7.90 ms | 2.5 MB |  | Sylvan.Data.Excel | Loss +16.2% |
| 2500 | speed-comparison | read-objects | ExcelDataReader | 12.76 ms | 6.2 MB |  | Sylvan.Data.Excel | 61.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | MiniExcel | 13.90 ms | 18.0 MB |  | Sylvan.Data.Excel | 75.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | ClosedXML | 29.64 ms | 16.5 MB |  | Sylvan.Data.Excel | 275.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | EPPlus | 30.86 ms | 20.9 MB |  | Sylvan.Data.Excel | 290.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | EPPlus 4.5.3.3 | 33.35 ms | 0 B |  | Sylvan.Data.Excel | 322.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | Sylvan.Data.Excel | 4.94 ms | 831.0 KB |  | Sylvan.Data.Excel | 15.1% faster than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | OfficeIMO.Excel | 5.81 ms | 2.5 MB |  | Sylvan.Data.Excel | Loss +17.8% |
| 2500 | speed-comparison | read-objects-stream | ExcelDataReader | 11.16 ms | 6.1 MB |  | Sylvan.Data.Excel | 92.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | MiniExcel | 13.42 ms | 18.0 MB |  | Sylvan.Data.Excel | 130.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | EPPlus | 28.75 ms | 20.8 MB |  | Sylvan.Data.Excel | 394.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | ClosedXML | 29.20 ms | 16.5 MB |  | Sylvan.Data.Excel | 402.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | EPPlus 4.5.3.3 | 32.83 ms | 0 B |  | Sylvan.Data.Excel | 464.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | OfficeIMO.Excel | 11.83 ms | 2.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | read-range | Sylvan.Data.Excel | 14.26 ms | 654.9 KB |  | OfficeIMO.Excel | 20.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | MiniExcel | 25.82 ms | 18.2 MB |  | OfficeIMO.Excel | 118.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | ExcelDataReader | 26.47 ms | 5.9 MB |  | OfficeIMO.Excel | 123.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | EPPlus | 34.88 ms | 19.7 MB |  | OfficeIMO.Excel | 194.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | EPPlus 4.5.3.3 | 35.77 ms | 0 B |  | OfficeIMO.Excel | 202.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | ClosedXML | 73.80 ms | 16.4 MB |  | OfficeIMO.Excel | 523.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | Sylvan.Data.Excel | 5.65 ms | 750.3 KB |  | Sylvan.Data.Excel | 30.9% faster than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | OfficeIMO.Excel | 8.17 ms | 2.8 MB |  | Sylvan.Data.Excel | Loss +44.6% |
| 2500 | speed-comparison | read-range-decimal | ExcelDataReader | 10.78 ms | 5.9 MB |  | Sylvan.Data.Excel | 31.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | MiniExcel | 12.62 ms | 18.2 MB |  | Sylvan.Data.Excel | 54.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | EPPlus | 27.98 ms | 19.7 MB |  | Sylvan.Data.Excel | 242.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-decimal | ClosedXML | 30.65 ms | 16.3 MB |  | Sylvan.Data.Excel | 275.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | Sylvan.Data.Excel | 4.94 ms | 655.2 KB |  | Sylvan.Data.Excel | 43.2% faster than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | OfficeIMO.Excel | 8.69 ms | 2.8 MB |  | Sylvan.Data.Excel | Loss +76.0% |
| 2500 | speed-comparison | read-range-stream | ExcelDataReader | 11.25 ms | 5.9 MB |  | Sylvan.Data.Excel | 29.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | MiniExcel | 12.31 ms | 18.2 MB |  | Sylvan.Data.Excel | 41.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | EPPlus | 29.21 ms | 19.7 MB |  | Sylvan.Data.Excel | 236.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | ClosedXML | 30.93 ms | 16.3 MB |  | Sylvan.Data.Excel | 256.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | EPPlus 4.5.3.3 | 32.64 ms | 0 B |  | Sylvan.Data.Excel | 275.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | Sylvan.Data.Excel | 0.52 ms | 348.4 KB |  | Sylvan.Data.Excel | 66.9% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range | MiniExcel | 0.90 ms | 858.3 KB |  | Sylvan.Data.Excel | 42.1% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range | OfficeIMO.Excel | 1.56 ms | 416.1 KB |  | Sylvan.Data.Excel | Loss +201.9% |
| 2500 | speed-comparison | read-top-range | ExcelDataReader | 4.61 ms | 1.9 MB |  | Sylvan.Data.Excel | 195.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | EPPlus | 26.65 ms | 12.1 MB |  | Sylvan.Data.Excel | 1608.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | ClosedXML | 30.08 ms | 15.0 MB |  | Sylvan.Data.Excel | 1828.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | EPPlus 4.5.3.3 | 31.81 ms | 0 B |  | Sylvan.Data.Excel | 1939.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | Sylvan.Data.Excel | 0.44 ms | 348.5 KB |  | Sylvan.Data.Excel | 70.7% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | MiniExcel | 0.78 ms | 858.3 KB |  | Sylvan.Data.Excel | 47.6% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | OfficeIMO.Excel | 1.49 ms | 419.5 KB |  | Sylvan.Data.Excel | Loss +241.8% |
| 2500 | speed-comparison | read-top-range-stream | ExcelDataReader | 4.27 ms | 1.9 MB |  | Sylvan.Data.Excel | 185.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | EPPlus | 22.65 ms | 12.1 MB |  | Sylvan.Data.Excel | 1416.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | ClosedXML | 28.80 ms | 15.0 MB |  | Sylvan.Data.Excel | 1827.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | EPPlus 4.5.3.3 | 29.40 ms | 0 B |  | Sylvan.Data.Excel | 1868.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | Sylvan.Data.Excel | 0.46 ms | 348.5 KB |  | Sylvan.Data.Excel | 71.4% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | MiniExcel | 0.73 ms | 858.3 KB |  | Sylvan.Data.Excel | 54.2% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | OfficeIMO.Excel | 1.60 ms | 420.2 KB |  | Sylvan.Data.Excel | Loss +249.3% |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | ExcelDataReader | 4.30 ms | 1.9 MB |  | Sylvan.Data.Excel | 167.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | EPPlus | 21.95 ms | 12.1 MB |  | Sylvan.Data.Excel | 1267.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream-small-chunks | ClosedXML | 28.60 ms | 15.0 MB |  | Sylvan.Data.Excel | 1681.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-used-range | Sylvan.Data.Excel | 5.00 ms | 655.2 KB |  | Sylvan.Data.Excel | 65.6% faster than OfficeIMO |
| 2500 | speed-comparison | read-used-range | MiniExcel | 13.80 ms | 18.2 MB |  | Sylvan.Data.Excel | 5.2% faster than OfficeIMO |
| 2500 | speed-comparison | read-used-range | OfficeIMO.Excel | 14.55 ms | 3.5 MB |  | Sylvan.Data.Excel | Loss +191.0% |
| 2500 | speed-comparison | read-used-range | ExcelDataReader | 14.99 ms | 5.9 MB |  | Sylvan.Data.Excel | 3.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-used-range | EPPlus | 33.20 ms | 19.7 MB |  | Sylvan.Data.Excel | 128.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-used-range | ClosedXML | 70.35 ms | 16.4 MB |  | Sylvan.Data.Excel | 383.4% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-autofilter | OfficeIMO.Excel | 3.69 ms | 1.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-autofilter | ClosedXML | 28.95 ms | 21.7 MB |  | OfficeIMO.Excel | 683.4% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-autofilter | EPPlus 4.5.3.3 | 39.92 ms | 0 B |  | OfficeIMO.Excel | 980.4% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-autofilter | EPPlus | 44.38 ms | 24.1 MB |  | OfficeIMO.Excel | 1101.2% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-charts | OfficeIMO.Excel | 5.01 ms | 1.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-charts | EPPlus 4.5.3.3 | 39.41 ms | 0 B |  | OfficeIMO.Excel | 687.4% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-charts | EPPlus | 51.29 ms | 26.4 MB |  | OfficeIMO.Excel | 924.7% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-conditional-formatting | OfficeIMO.Excel | 3.93 ms | 1.2 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-conditional-formatting | ClosedXML | 30.67 ms | 21.7 MB |  | OfficeIMO.Excel | 679.9% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-conditional-formatting | EPPlus 4.5.3.3 | 39.96 ms | 0 B |  | OfficeIMO.Excel | 916.3% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-conditional-formatting | EPPlus | 47.48 ms | 24.1 MB |  | OfficeIMO.Excel | 1107.7% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-data-validation | OfficeIMO.Excel | 3.67 ms | 1.2 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-data-validation | ClosedXML | 29.63 ms | 21.7 MB |  | OfficeIMO.Excel | 706.9% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-data-validation | EPPlus 4.5.3.3 | 42.20 ms | 0 B |  | OfficeIMO.Excel | 1049.3% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-data-validation | EPPlus | 46.03 ms | 24.1 MB |  | OfficeIMO.Excel | 1153.7% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-freeze-panes | OfficeIMO.Excel | 3.46 ms | 1.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-freeze-panes | ClosedXML | 28.29 ms | 21.7 MB |  | OfficeIMO.Excel | 717.1% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-freeze-panes | EPPlus 4.5.3.3 | 38.40 ms | 0 B |  | OfficeIMO.Excel | 1009.3% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-freeze-panes | EPPlus | 44.71 ms | 24.1 MB |  | OfficeIMO.Excel | 1191.6% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-pivot-table | OfficeIMO.Excel | 16.56 ms | 18.2 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-pivot-table | EPPlus 4.5.3.3 | 37.41 ms | 0 B |  | OfficeIMO.Excel | 125.9% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-pivot-table | EPPlus | 50.71 ms | 28.8 MB |  | OfficeIMO.Excel | 206.1% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-all-in-one | OfficeIMO.Excel | 18.59 ms | 19.0 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-all-in-one | EPPlus | 58.40 ms | 53.2 MB |  | OfficeIMO.Excel | 214.2% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-all-in-one | EPPlus 4.5.3.3 | 76.95 ms | 0 B |  | OfficeIMO.Excel | 313.9% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-core | OfficeIMO.Excel | 4.15 ms | 1.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | realworld-report-core | EPPlus | 61.69 ms | 46.1 MB |  | OfficeIMO.Excel | 1384.8% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-core | EPPlus 4.5.3.3 | 73.14 ms | 0 B |  | OfficeIMO.Excel | 1660.3% slower than OfficeIMO |
| 2500 | speed-comparison | realworld-report-core | ClosedXML | 75.83 ms | 68.2 MB |  | OfficeIMO.Excel | 1725.2% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook | OfficeIMO.Excel | 15.21 ms | 11.9 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | report-workbook | EPPlus | 92.76 ms | 75.6 MB |  | OfficeIMO.Excel | 509.9% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook | EPPlus 4.5.3.3 | 105.14 ms | 0 B |  | OfficeIMO.Excel | 591.3% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-core | OfficeIMO.Excel | 5.79 ms | 2.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | report-workbook-core | EPPlus | 79.26 ms | 70.2 MB |  | OfficeIMO.Excel | 1269.4% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-core | ClosedXML | 98.36 ms | 94.9 MB |  | OfficeIMO.Excel | 1599.4% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-core | EPPlus 4.5.3.3 | 98.91 ms | 0 B |  | OfficeIMO.Excel | 1609.0% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-datatable | OfficeIMO.Excel | 13.57 ms | 12.2 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | report-workbook-datatable | EPPlus | 78.84 ms | 64.4 MB |  | OfficeIMO.Excel | 480.9% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-datatable | EPPlus 4.5.3.3 | 95.85 ms | 0 B |  | OfficeIMO.Excel | 606.2% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-datatable-core | OfficeIMO.Excel | 5.55 ms | 2.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | report-workbook-datatable-core | EPPlus | 74.43 ms | 59.0 MB |  | OfficeIMO.Excel | 1241.5% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-datatable-core | EPPlus 4.5.3.3 | 85.02 ms | 0 B |  | OfficeIMO.Excel | 1432.2% slower than OfficeIMO |
| 2500 | speed-comparison | report-workbook-datatable-core | ClosedXML | 94.28 ms | 80.9 MB |  | OfficeIMO.Excel | 1599.3% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | Sylvan.Data.Excel | 2.08 ms | 518.6 KB |  | Sylvan.Data.Excel | 40.2% faster than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | OfficeIMO.Excel | 3.47 ms | 1.1 MB |  | Sylvan.Data.Excel | Loss +67.1% |
| 2500 | speed-comparison | shared-string-read | ExcelDataReader | 4.73 ms | 2.5 MB |  | Sylvan.Data.Excel | 36.3% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | MiniExcel | 5.80 ms | 7.3 MB |  | Sylvan.Data.Excel | 67.0% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | EPPlus 4.5.3.3 | 14.28 ms | 0 B |  | Sylvan.Data.Excel | 311.1% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | ClosedXML | 15.21 ms | 9.3 MB |  | Sylvan.Data.Excel | 338.1% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | EPPlus | 17.88 ms | 10.1 MB |  | Sylvan.Data.Excel | 414.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-blog-2023-20-string-columns | OfficeIMO.Excel | 5.04 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-blog-2023-20-string-columns | LargeXlsx | 5.22 ms | 849.6 KB |  | OfficeIMO.Excel | 3.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-blog-2023-20-string-columns | MiniExcel | 17.52 ms | 35.1 MB |  | OfficeIMO.Excel | 247.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-blog-2023-20-string-columns | ClosedXML | 96.02 ms | 69.8 MB |  | OfficeIMO.Excel | 1804.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | OfficeIMO.Excel | 7.21 ms | 1.2 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-bulk-report | MiniExcel | 17.56 ms | 26.2 MB |  | OfficeIMO.Excel | 143.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | EPPlus 4.5.3.3 | 112.17 ms | 0 B |  | OfficeIMO.Excel | 1455.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | EPPlus | 152.33 ms | 47.9 MB |  | OfficeIMO.Excel | 2013.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | ClosedXML | 277.96 ms | 57.0 MB |  | OfficeIMO.Excel | 3755.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellformula | OfficeIMO.Excel | 2.81 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellformula | EPPlus 4.5.3.3 | 20.04 ms | 0 B |  | OfficeIMO.Excel | 613.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellformula | ClosedXML | 21.99 ms | 11.7 MB |  | OfficeIMO.Excel | 682.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellformula | EPPlus | 41.27 ms | 17.6 MB |  | OfficeIMO.Excel | 1367.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-empty-strings | OfficeIMO.Excel | 2.91 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-empty-strings | ClosedXML | 16.46 ms | 9.7 MB |  | OfficeIMO.Excel | 466.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-empty-strings | EPPlus | 24.85 ms | 11.4 MB |  | OfficeIMO.Excel | 755.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-numbers | OfficeIMO.Excel | 2.98 ms | 946.8 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-numbers | ClosedXML | 11.93 ms | 9.0 MB |  | OfficeIMO.Excel | 299.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-numbers | EPPlus 4.5.3.3 | 15.35 ms | 0 B |  | OfficeIMO.Excel | 414.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-numbers | EPPlus | 27.95 ms | 12.5 MB |  | OfficeIMO.Excel | 836.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-mixed | OfficeIMO.Excel | 2.90 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-object-mixed | EPPlus 4.5.3.3 | 18.26 ms | 0 B |  | OfficeIMO.Excel | 529.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-mixed | ClosedXML | 19.84 ms | 11.6 MB |  | OfficeIMO.Excel | 583.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-mixed | EPPlus | 29.60 ms | 15.2 MB |  | OfficeIMO.Excel | 920.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-sparse | OfficeIMO.Excel | 4.69 ms | 1.2 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-object-sparse | ClosedXML | 16.30 ms | 11.0 MB |  | OfficeIMO.Excel | 247.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-sparse | EPPlus | 30.71 ms | 14.5 MB |  | OfficeIMO.Excel | 554.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 3.26 ms | 1.2 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-object-sparse-batch | ClosedXML | 15.47 ms | 11.0 MB |  | OfficeIMO.Excel | 374.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-object-sparse-batch | EPPlus | 27.73 ms | 14.5 MB |  | OfficeIMO.Excel | 750.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-scalars | OfficeIMO.Excel | 2.65 ms | 964.9 KB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-scalars | ClosedXML | 11.44 ms | 8.8 MB |  | OfficeIMO.Excel | 331.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-scalars | EPPlus 4.5.3.3 | 16.51 ms | 0 B |  | OfficeIMO.Excel | 523.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-scalars | EPPlus | 26.49 ms | 12.5 MB |  | OfficeIMO.Excel | 899.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings | OfficeIMO.Excel | 2.99 ms | 2.2 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-strings | ClosedXML | 16.22 ms | 11.0 MB |  | OfficeIMO.Excel | 441.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings | EPPlus 4.5.3.3 | 17.27 ms | 0 B |  | OfficeIMO.Excel | 476.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings | EPPlus | 25.78 ms | 12.4 MB |  | OfficeIMO.Excel | 760.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings-distinct | OfficeIMO.Excel | 3.11 ms | 2.2 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-strings-distinct | ClosedXML | 20.50 ms | 12.8 MB |  | OfficeIMO.Excel | 558.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings-distinct | EPPlus | 28.35 ms | 13.5 MB |  | OfficeIMO.Excel | 810.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings-repeated | OfficeIMO.Excel | 2.57 ms | 1.2 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-strings-repeated | ClosedXML | 12.18 ms | 9.0 MB |  | OfficeIMO.Excel | 374.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-strings-repeated | EPPlus | 19.46 ms | 11.0 MB |  | OfficeIMO.Excel | 657.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-temporal | OfficeIMO.Excel | 3.28 ms | 1.2 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalue-temporal | ClosedXML | 15.48 ms | 9.5 MB |  | OfficeIMO.Excel | 371.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-temporal | EPPlus 4.5.3.3 | 17.51 ms | 0 B |  | OfficeIMO.Excel | 433.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalue-temporal | EPPlus | 25.73 ms | 14.3 MB |  | OfficeIMO.Excel | 684.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 1.30 ms | 439.0 KB |  | LargeXlsx | 21.9% faster than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 1.66 ms | 923.6 KB |  | LargeXlsx | Loss +28.0% |
| 2500 | speed-comparison | write-cellvalues-headerless-rectangle-direct | ClosedXML | 15.63 ms | 10.0 MB |  | LargeXlsx | 840.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-headerless-rectangle-direct | EPPlus | 24.31 ms | 12.7 MB |  | LargeXlsx | 1362.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | LargeXlsx | 3.65 ms | 750.2 KB |  | LargeXlsx | 21.2% faster than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 4.64 ms | 1.7 MB |  | LargeXlsx | Loss +27.0% |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | MiniExcel | 10.89 ms | 22.7 MB |  | LargeXlsx | 134.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 35.16 ms | 0 B |  | LargeXlsx | 658.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | ClosedXML | 36.01 ms | 21.7 MB |  | LargeXlsx | 676.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus | 53.97 ms | 24.0 MB |  | LargeXlsx | 1063.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-sparse-rectangle-direct | OfficeIMO.Excel | 2.43 ms | 1.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-cellvalues-sparse-rectangle-direct | ClosedXML | 17.48 ms | 11.0 MB |  | OfficeIMO.Excel | 619.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-sparse-rectangle-direct | EPPlus | 27.79 ms | 14.5 MB |  | OfficeIMO.Excel | 1043.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | OfficeIMO.Excel | 5.34 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datareader-plain | LargeXlsx | 7.74 ms | 1.0 MB |  | OfficeIMO.Excel | 44.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | Sylvan.Data.Excel | 7.91 ms | 750.5 KB |  | OfficeIMO.Excel | 48.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | MiniExcel | 8.71 ms | 22.5 MB |  | OfficeIMO.Excel | 63.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | EPPlus 4.5.3.3 | 33.37 ms | 0 B |  | OfficeIMO.Excel | 524.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | ClosedXML | 34.33 ms | 11.3 MB |  | OfficeIMO.Excel | 542.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-plain | EPPlus | 44.66 ms | 16.2 MB |  | OfficeIMO.Excel | 735.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | OfficeIMO.Excel | 4.40 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datareader-table | MiniExcel | 10.69 ms | 22.5 MB |  | OfficeIMO.Excel | 142.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | EPPlus 4.5.3.3 | 32.13 ms | 0 B |  | OfficeIMO.Excel | 630.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | ClosedXML | 37.82 ms | 18.6 MB |  | OfficeIMO.Excel | 759.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | EPPlus | 39.41 ms | 16.2 MB |  | OfficeIMO.Excel | 795.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table-autofit | OfficeIMO.Excel | 5.40 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datareader-table-autofit | MiniExcel | 13.86 ms | 26.0 MB |  | OfficeIMO.Excel | 156.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table-autofit | EPPlus | 71.32 ms | 37.4 MB |  | OfficeIMO.Excel | 1221.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table-autofit | EPPlus 4.5.3.3 | 73.95 ms | 0 B |  | OfficeIMO.Excel | 1270.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table-autofit | ClosedXML | 90.59 ms | 57.0 MB |  | OfficeIMO.Excel | 1579.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-headerless-tables | OfficeIMO.Excel | 5.53 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-headerless-tables | MiniExcel | 15.07 ms | 28.4 MB |  | OfficeIMO.Excel | 172.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-headerless-tables | ClosedXML | 44.46 ms | 18.5 MB |  | OfficeIMO.Excel | 704.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-headerless-tables | EPPlus | 58.18 ms | 17.2 MB |  | OfficeIMO.Excel | 953.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | OfficeIMO.Excel | 5.41 ms | 1.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | LargeXlsx | 10.34 ms | 1.1 MB |  | OfficeIMO.Excel | 91.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | MiniExcel | 12.31 ms | 29.0 MB |  | OfficeIMO.Excel | 127.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | ClosedXML | 58.39 ms | 26.8 MB |  | OfficeIMO.Excel | 980.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-direct-export | EPPlus | 65.19 ms | 21.3 MB |  | OfficeIMO.Excel | 1106.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-tables | OfficeIMO.Excel | 8.58 ms | 2.3 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-sparse-tables | MiniExcel | 12.68 ms | 29.0 MB |  | OfficeIMO.Excel | 47.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-tables | ClosedXML | 57.84 ms | 26.8 MB |  | OfficeIMO.Excel | 574.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-tables | EPPlus | 58.03 ms | 21.3 MB |  | OfficeIMO.Excel | 576.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | OfficeIMO.Excel | 5.44 ms | 1.5 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-tables | MiniExcel | 18.88 ms | 28.4 MB |  | OfficeIMO.Excel | 246.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | EPPlus 4.5.3.3 | 42.06 ms | 0 B |  | OfficeIMO.Excel | 672.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | ClosedXML | 86.38 ms | 18.4 MB |  | OfficeIMO.Excel | 1486.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | EPPlus | 89.88 ms | 18.9 MB |  | OfficeIMO.Excel | 1550.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables-autofit | OfficeIMO.Excel | 6.35 ms | 1.6 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-dataset-tables-autofit | MiniExcel | 17.92 ms | 31.6 MB |  | OfficeIMO.Excel | 182.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables-autofit | EPPlus | 116.48 ms | 42.3 MB |  | OfficeIMO.Excel | 1733.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables-autofit | ClosedXML | 231.05 ms | 55.4 MB |  | OfficeIMO.Excel | 3536.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | OfficeIMO.Excel | 5.02 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datatable-direct | LargeXlsx | 9.94 ms | 1.1 MB |  | OfficeIMO.Excel | 98.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | MiniExcel | 11.92 ms | 22.5 MB |  | OfficeIMO.Excel | 137.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | ClosedXML | 40.89 ms | 11.3 MB |  | OfficeIMO.Excel | 714.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | EPPlus 4.5.3.3 | 41.55 ms | 0 B |  | OfficeIMO.Excel | 727.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | EPPlus | 51.81 ms | 16.2 MB |  | OfficeIMO.Excel | 932.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-object-table-direct | OfficeIMO.Excel | 5.05 ms | 1.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datatable-object-table-direct | MiniExcel | 9.14 ms | 22.2 MB |  | OfficeIMO.Excel | 80.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-object-table-direct | ClosedXML | 43.54 ms | 18.3 MB |  | OfficeIMO.Excel | 761.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-object-table-direct | EPPlus | 49.24 ms | 15.9 MB |  | OfficeIMO.Excel | 874.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | OfficeIMO.Excel | 4.82 ms | 1.4 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-datatable-table-direct | MiniExcel | 9.28 ms | 22.5 MB |  | OfficeIMO.Excel | 92.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | EPPlus 4.5.3.3 | 42.06 ms | 0 B |  | OfficeIMO.Excel | 773.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | ClosedXML | 43.22 ms | 18.6 MB |  | OfficeIMO.Excel | 797.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | EPPlus | 51.11 ms | 16.2 MB |  | OfficeIMO.Excel | 961.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-dictionary-objects-table-direct | OfficeIMO.Excel | 7.91 ms | 1.7 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | LargeXlsx | 2.95 ms | 750.2 KB |  | LargeXlsx | 18.1% faster than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 3.60 ms | 1.4 MB |  | LargeXlsx | Loss +22.1% |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | MiniExcel | 7.76 ms | 22.7 MB |  | LargeXlsx | 115.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | ClosedXML | 25.77 ms | 11.3 MB |  | LargeXlsx | 615.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus | 34.21 ms | 16.2 MB |  | LargeXlsx | 849.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 35.19 ms | 0 B |  | LargeXlsx | 876.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.42 ms | 1.2 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 73.80 ms | 37.4 MB |  | OfficeIMO.Excel | 1568.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 91.44 ms | 49.7 MB |  | OfficeIMO.Excel | 1967.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | LargeXlsx | 3.50 ms | 750.2 KB |  | LargeXlsx | 24.0% faster than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | OfficeIMO.Excel | 4.60 ms | 1.1 MB |  | LargeXlsx | Loss +31.5% |
| 2500 | speed-comparison | write-insertobjects-direct | MiniExcel | 8.79 ms | 22.7 MB |  | LargeXlsx | 90.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | ClosedXML | 31.75 ms | 11.3 MB |  | LargeXlsx | 589.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | EPPlus 4.5.3.3 | 32.91 ms | 0 B |  | LargeXlsx | 614.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | EPPlus | 45.59 ms | 16.2 MB |  | LargeXlsx | 890.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.45 ms | 1.1 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 60.80 ms | 37.4 MB |  | OfficeIMO.Excel | 1266.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 68.27 ms | 49.7 MB |  | OfficeIMO.Excel | 1434.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 3.15 ms | 750.2 KB |  | LargeXlsx | 32.4% faster than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 4.66 ms | 1.1 MB |  | LargeXlsx | Loss +47.8% |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | MiniExcel | 9.41 ms | 22.7 MB |  | LargeXlsx | 101.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | ClosedXML | 28.99 ms | 11.3 MB |  | LargeXlsx | 521.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-flat-dictionaries-direct | EPPlus | 41.06 ms | 16.2 MB |  | LargeXlsx | 780.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.23 ms | 1.2 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 49.91 ms | 27.9 MB |  | OfficeIMO.Excel | 1080.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 58.72 ms | 26.6 MB |  | OfficeIMO.Excel | 1288.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | LargeXlsx | 4.58 ms | 794.5 KB |  | LargeXlsx | 18.4% faster than OfficeIMO |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 5.61 ms | 2.0 MB |  | LargeXlsx | Loss +22.5% |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | MiniExcel | 8.18 ms | 24.6 MB |  | LargeXlsx | 45.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | ClosedXML | 39.16 ms | 16.6 MB |  | LargeXlsx | 597.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-mixed-objects-direct | EPPlus | 47.42 ms | 19.6 MB |  | LargeXlsx | 745.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-mixed-direct | LargeXlsx | 4.33 ms | 794.5 KB |  | LargeXlsx | 19.6% faster than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 5.39 ms | 1.3 MB |  | LargeXlsx | Loss +24.3% |
| 2500 | speed-comparison | write-powershell-psobject-mixed-direct | MiniExcel | 9.12 ms | 24.6 MB |  | LargeXlsx | 69.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-mixed-direct | ClosedXML | 38.16 ms | 16.6 MB |  | LargeXlsx | 608.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-mixed-direct | EPPlus | 48.34 ms | 19.6 MB |  | LargeXlsx | 797.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 20.40 ms | 4.2 MB |  | OfficeIMO.Excel | Win |
| 2500 | speed-comparison | write-powershell-psobject-wide-direct | LargeXlsx | 21.01 ms | 2.7 MB |  | OfficeIMO.Excel | 3.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-wide-direct | MiniExcel | 35.38 ms | 47.3 MB |  | OfficeIMO.Excel | 73.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-wide-direct | ClosedXML | 122.84 ms | 50.4 MB |  | OfficeIMO.Excel | 502.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-powershell-psobject-wide-direct | EPPlus | 174.10 ms | 67.5 MB |  | OfficeIMO.Excel | 753.6% slower than OfficeIMO |
| 10000 | focused-package-profile | write-blog-2023-20-string-columns | OfficeIMO.Excel | 33.01 ms | 7.6 MB | 880.4 KB | OfficeIMO.Excel | Win |
| 10000 | focused-package-profile | write-blog-2023-20-string-columns | LargeXlsx | 88.91 ms | 3.1 MB | 970.2 KB | OfficeIMO.Excel | 2.69x vs best |
| 10000 | focused-package-profile | write-blog-2023-20-string-columns | MiniExcel | 142.47 ms | 96.2 MB | 957.6 KB | OfficeIMO.Excel | 4.32x vs best |
| 10000 | focused-package-profile | write-blog-2023-20-string-columns | ClosedXML | 725.30 ms | 280.2 MB | 1,015.4 KB | OfficeIMO.Excel | 21.97x vs best |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | OfficeIMO.Excel | 60.77 ms | 23.1 MB |  | OfficeIMO.Excel, Sylvan.Data.Excel | Win |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | Sylvan.Data.Excel | 61.82 ms | 394.1 KB |  | OfficeIMO.Excel, Sylvan.Data.Excel | Tie vs OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | ExcelDataReader | 164.62 ms | 67.9 MB |  | OfficeIMO.Excel, Sylvan.Data.Excel | 170.9% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | MiniExcel | 208.69 ms | 210.3 MB |  | OfficeIMO.Excel, Sylvan.Data.Excel | 243.4% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | Sylvan.Data.Excel | 47.40 ms | 394.1 KB |  | Sylvan.Data.Excel | 18.7% faster than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | OfficeIMO.Excel | 58.28 ms | 23.9 MB |  | Sylvan.Data.Excel | Loss +22.9% |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | ExcelDataReader | 154.40 ms | 67.9 MB |  | Sylvan.Data.Excel | 165.0% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | MiniExcel | 202.82 ms | 210.3 MB |  | Sylvan.Data.Excel | 248.0% slower than OfficeIMO |
| 25000 | package-profile | append-plain-rows | LargeXlsx | 14.09 ms | 2.7 MB | 605.0 KB | LargeXlsx | 32.5% faster than OfficeIMO |
| 25000 | package-profile | append-plain-rows | OfficeIMO.Excel | 20.87 ms | 11.4 MB | 622.5 KB | LargeXlsx | Loss +48.1% |
| 25000 | package-profile | append-plain-rows | MiniExcel | 40.13 ms | 56.9 MB | 642.3 KB | LargeXlsx | 92.3% slower than OfficeIMO |
| 25000 | package-profile | append-plain-rows | ClosedXML | 146.77 ms | 101.8 MB | 540.6 KB | LargeXlsx | 603.2% slower than OfficeIMO |
| 25000 | package-profile | append-plain-rows | EPPlus | 223.35 ms | 97.9 MB | 525.6 KB | LargeXlsx | 970.1% slower than OfficeIMO |
| 25000 | package-profile | autofit-existing | OfficeIMO.Excel | 342.58 ms | 135.1 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | autofit-existing | EPPlus | 464.86 ms | 245.0 MB | 1.1 MB | OfficeIMO.Excel | 35.7% slower than OfficeIMO |
| 25000 | package-profile | autofit-existing | ClosedXML | 1.36 s | 810.1 MB | 1.1 MB | OfficeIMO.Excel | 296.0% slower than OfficeIMO |
| 25000 | package-profile | large-shared-strings | OfficeIMO.Excel | 14.99 ms | 15.1 MB | 529.7 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | large-shared-strings | MiniExcel | 31.72 ms | 72.0 MB | 581.0 KB | OfficeIMO.Excel | 111.6% slower than OfficeIMO |
| 25000 | package-profile | large-shared-strings | ClosedXML | 118.65 ms | 101.8 MB | 460.1 KB | OfficeIMO.Excel | 691.5% slower than OfficeIMO |
| 25000 | package-profile | large-shared-strings | EPPlus | 197.43 ms | 82.4 MB | 444.7 KB | OfficeIMO.Excel | 1217.0% slower than OfficeIMO |
| 25000 | package-profile | realworld-autofilter | OfficeIMO.Excel | 30.84 ms | 11.1 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-autofilter | ClosedXML | 286.03 ms | 205.7 MB | 1.1 MB | OfficeIMO.Excel | 827.4% slower than OfficeIMO |
| 25000 | package-profile | realworld-autofilter | EPPlus | 336.19 ms | 206.8 MB | 1.1 MB | OfficeIMO.Excel | 990.0% slower than OfficeIMO |
| 25000 | package-profile | realworld-charts | OfficeIMO.Excel | 41.92 ms | 12.0 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-charts | EPPlus | 428.40 ms | 209.8 MB | 1.1 MB | OfficeIMO.Excel | 922.0% slower than OfficeIMO |
| 25000 | package-profile | realworld-conditional-formatting | OfficeIMO.Excel | 30.63 ms | 11.1 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-conditional-formatting | ClosedXML | 297.68 ms | 205.8 MB | 1.1 MB | OfficeIMO.Excel | 871.8% slower than OfficeIMO |
| 25000 | package-profile | realworld-conditional-formatting | EPPlus | 362.33 ms | 206.9 MB | 1.1 MB | OfficeIMO.Excel | 1082.8% slower than OfficeIMO |
| 25000 | package-profile | realworld-data-validation | OfficeIMO.Excel | 39.78 ms | 11.1 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-data-validation | ClosedXML | 352.17 ms | 205.7 MB | 1.1 MB | OfficeIMO.Excel | 785.2% slower than OfficeIMO |
| 25000 | package-profile | realworld-data-validation | EPPlus | 438.54 ms | 206.8 MB | 1.1 MB | OfficeIMO.Excel | 1002.3% slower than OfficeIMO |
| 25000 | package-profile | realworld-freeze-panes | OfficeIMO.Excel | 31.47 ms | 11.1 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-freeze-panes | ClosedXML | 289.57 ms | 205.7 MB | 1.1 MB | OfficeIMO.Excel | 820.0% slower than OfficeIMO |
| 25000 | package-profile | realworld-freeze-panes | EPPlus | 345.28 ms | 206.9 MB | 1.1 MB | OfficeIMO.Excel | 997.0% slower than OfficeIMO |
| 25000 | package-profile | realworld-pivot-table | OfficeIMO.Excel | 117.92 ms | 98.5 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-pivot-table | EPPlus | 439.09 ms | 225.3 MB | 1.1 MB | OfficeIMO.Excel | 272.4% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-all-in-one | OfficeIMO.Excel | 117.37 ms | 99.9 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-all-in-one | EPPlus | 474.97 ms | 270.6 MB | 1.1 MB | OfficeIMO.Excel | 304.7% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-core | OfficeIMO.Excel | 38.60 ms | 11.2 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | realworld-report-core | EPPlus | 400.17 ms | 249.1 MB | 1.1 MB | OfficeIMO.Excel | 936.7% slower than OfficeIMO |
| 25000 | package-profile | realworld-report-core | ClosedXML | 887.72 ms | 664.2 MB | 1.1 MB | OfficeIMO.Excel | 2199.9% slower than OfficeIMO |
| 25000 | package-profile | report-workbook | OfficeIMO.Excel | 48.14 ms | 14.0 MB | 1.8 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | report-workbook | EPPlus | 545.30 ms | 356.1 MB | 1.5 MB | OfficeIMO.Excel | 1032.9% slower than OfficeIMO |
| 25000 | package-profile | report-workbook-core | OfficeIMO.Excel | 55.61 ms | 10.6 MB | 1.8 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | report-workbook-core | EPPlus | 673.82 ms | 334.7 MB | 1.5 MB | OfficeIMO.Excel | 1111.6% slower than OfficeIMO |
| 25000 | package-profile | report-workbook-core | ClosedXML | 1.34 s | 952.9 MB | 1.5 MB | OfficeIMO.Excel | 2312.8% slower than OfficeIMO |
| 25000 | package-profile | report-workbook-datatable | OfficeIMO.Excel | 148.11 ms | 16.7 MB | 1.8 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | report-workbook-datatable | EPPlus | 1.34 s | 242.0 MB | 1.5 MB | OfficeIMO.Excel | 802.3% slower than OfficeIMO |
| 25000 | package-profile | report-workbook-datatable-core | OfficeIMO.Excel | 61.19 ms | 13.2 MB | 1.8 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | report-workbook-datatable-core | EPPlus | 676.56 ms | 220.7 MB | 1.5 MB | OfficeIMO.Excel | 1005.7% slower than OfficeIMO |
| 25000 | package-profile | report-workbook-datatable-core | ClosedXML | 1.46 s | 812.7 MB | 1.5 MB | OfficeIMO.Excel | 2279.6% slower than OfficeIMO |
| 25000 | package-profile | write-blog-2023-20-string-columns | LargeXlsx | 39.31 ms | 10.5 MB | 2.4 MB | LargeXlsx | 9.6% faster than OfficeIMO |
| 25000 | package-profile | write-blog-2023-20-string-columns | OfficeIMO.Excel | 43.48 ms | 11.3 MB | 2.2 MB | LargeXlsx | Loss +10.6% |
| 25000 | package-profile | write-blog-2023-20-string-columns | MiniExcel | 148.34 ms | 221.6 MB | 2.4 MB | LargeXlsx | 241.2% slower than OfficeIMO |
| 25000 | package-profile | write-blog-2023-20-string-columns | ClosedXML | 926.37 ms | 742.0 MB | 2.5 MB | LargeXlsx | 2030.6% slower than OfficeIMO |
| 25000 | package-profile | write-bulk-report | OfficeIMO.Excel | 34.49 ms | 11.1 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-bulk-report | MiniExcel | 64.38 ms | 122.6 MB | 1.5 MB | OfficeIMO.Excel | 86.7% slower than OfficeIMO |
| 25000 | package-profile | write-bulk-report | EPPlus | 382.74 ms | 248.9 MB | 1.1 MB | OfficeIMO.Excel | 1009.7% slower than OfficeIMO |
| 25000 | package-profile | write-bulk-report | ClosedXML | 761.72 ms | 552.7 MB | 1.1 MB | OfficeIMO.Excel | 2108.5% slower than OfficeIMO |
| 25000 | package-profile | write-cellformula | OfficeIMO.Excel | 25.87 ms | 9.3 MB | 670.3 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellformula | ClosedXML | 221.62 ms | 111.2 MB | 643.2 KB | OfficeIMO.Excel | 756.7% slower than OfficeIMO |
| 25000 | package-profile | write-cellformula | EPPlus | 436.26 ms | 137.4 MB | 593.9 KB | OfficeIMO.Excel | 1586.3% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-empty-strings | OfficeIMO.Excel | 12.62 ms | 6.6 MB | 451.4 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-empty-strings | ClosedXML | 109.53 ms | 90.7 MB | 398.1 KB | OfficeIMO.Excel | 767.6% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-empty-strings | EPPlus | 162.80 ms | 72.7 MB | 390.6 KB | OfficeIMO.Excel | 1189.5% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-numbers | OfficeIMO.Excel | 15.94 ms | 5.7 MB | 462.6 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-numbers | ClosedXML | 104.55 ms | 82.2 MB | 411.4 KB | OfficeIMO.Excel | 555.9% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-numbers | EPPlus | 202.05 ms | 84.4 MB | 406.5 KB | OfficeIMO.Excel | 1167.6% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-mixed | OfficeIMO.Excel | 18.84 ms | 7.8 MB | 585.1 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-object-mixed | ClosedXML | 159.22 ms | 108.5 MB | 532.9 KB | OfficeIMO.Excel | 745.2% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-mixed | EPPlus | 274.52 ms | 110.6 MB | 544.3 KB | OfficeIMO.Excel | 1357.3% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-sparse | OfficeIMO.Excel | 18.93 ms | 7.0 MB | 607.1 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-object-sparse | ClosedXML | 145.15 ms | 102.8 MB | 468.0 KB | OfficeIMO.Excel | 666.8% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-sparse | EPPlus | 234.91 ms | 103.8 MB | 494.4 KB | OfficeIMO.Excel | 1140.9% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 20.86 ms | 7.0 MB | 607.1 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-object-sparse-batch | ClosedXML | 160.37 ms | 102.8 MB | 468.0 KB | OfficeIMO.Excel | 668.8% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-object-sparse-batch | EPPlus | 279.38 ms | 103.8 MB | 494.4 KB | OfficeIMO.Excel | 1239.3% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-scalars | OfficeIMO.Excel | 11.71 ms | 5.8 MB | 441.9 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-scalars | ClosedXML | 96.98 ms | 80.7 MB | 394.9 KB | OfficeIMO.Excel | 727.9% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-scalars | EPPlus | 236.22 ms | 83.1 MB | 379.3 KB | OfficeIMO.Excel | 1916.6% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings | OfficeIMO.Excel | 18.58 ms | 14.7 MB | 527.8 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-strings | ClosedXML | 141.87 ms | 101.8 MB | 460.1 KB | OfficeIMO.Excel | 663.5% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings | EPPlus | 223.16 ms | 82.4 MB | 444.7 KB | OfficeIMO.Excel | 1101.1% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings-distinct | OfficeIMO.Excel | 14.93 ms | 13.3 MB | 499.0 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-strings-distinct | ClosedXML | 159.99 ms | 128.4 MB | 555.3 KB | OfficeIMO.Excel | 971.3% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings-distinct | EPPlus | 215.95 ms | 95.4 MB | 565.1 KB | OfficeIMO.Excel | 1346.1% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings-repeated | OfficeIMO.Excel | 14.38 ms | 7.0 MB | 376.0 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-strings-repeated | ClosedXML | 117.96 ms | 82.5 MB | 331.8 KB | OfficeIMO.Excel | 720.5% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-strings-repeated | EPPlus | 244.28 ms | 68.4 MB | 300.8 KB | OfficeIMO.Excel | 1599.1% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-temporal | OfficeIMO.Excel | 20.22 ms | 7.1 MB | 620.5 KB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-cellvalue-temporal | ClosedXML | 145.82 ms | 87.2 MB | 483.0 KB | OfficeIMO.Excel | 621.2% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalue-temporal | EPPlus | 215.91 ms | 101.4 MB | 495.1 KB | OfficeIMO.Excel | 967.8% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 9.36 ms | 3.4 MB | 443.4 KB | LargeXlsx | 17.4% faster than OfficeIMO |
| 25000 | package-profile | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 11.34 ms | 6.6 MB | 455.5 KB | LargeXlsx | Loss +21.1% |
| 25000 | package-profile | write-cellvalues-headerless-rectangle-direct | ClosedXML | 118.84 ms | 93.8 MB | 467.5 KB | LargeXlsx | 948.3% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-headerless-rectangle-direct | EPPlus | 193.96 ms | 85.3 MB | 484.1 KB | LargeXlsx | 1611.0% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | LargeXlsx | 27.10 ms | 5.5 MB | 1.4 MB | LargeXlsx | 22.6% faster than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 35.01 ms | 15.3 MB | 1.4 MB | LargeXlsx | Loss +29.2% |
| 25000 | package-profile | write-cellvalues-rectangle-direct | MiniExcel | 66.46 ms | 91.1 MB | 1.5 MB | LargeXlsx | 89.8% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | ClosedXML | 309.83 ms | 205.7 MB | 1.1 MB | LargeXlsx | 784.9% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | EPPlus | 376.69 ms | 206.8 MB | 1.1 MB | LargeXlsx | 975.8% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | Sylvan.Data.Excel | 28.40 ms | 5.6 MB | 755.4 KB | Sylvan.Data.Excel | 23.6% faster than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | LargeXlsx | 34.40 ms | 8.1 MB | 1.4 MB | Sylvan.Data.Excel | 7.5% faster than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | OfficeIMO.Excel | 37.20 ms | 12.4 MB | 1.4 MB | Sylvan.Data.Excel | Loss +31.0% |
| 25000 | package-profile | write-datareader-plain | MiniExcel | 76.60 ms | 90.0 MB | 1.5 MB | Sylvan.Data.Excel | 105.9% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | ClosedXML | 279.07 ms | 101.8 MB | 1.1 MB | Sylvan.Data.Excel | 650.3% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-plain | EPPlus | 346.14 ms | 114.6 MB | 1.1 MB | Sylvan.Data.Excel | 830.6% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table | OfficeIMO.Excel | 34.36 ms | 12.4 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-datareader-table | MiniExcel | 64.09 ms | 90.0 MB | 1.5 MB | OfficeIMO.Excel | 86.5% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table | EPPlus | 315.29 ms | 114.6 MB | 1.1 MB | OfficeIMO.Excel | 817.6% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table | ClosedXML | 353.26 ms | 169.3 MB | 1.1 MB | OfficeIMO.Excel | 928.1% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table-autofit | OfficeIMO.Excel | 38.07 ms | 12.4 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-datareader-table-autofit | MiniExcel | 71.32 ms | 121.6 MB | 1.5 MB | OfficeIMO.Excel | 87.3% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table-autofit | EPPlus | 362.42 ms | 155.9 MB | 1.1 MB | OfficeIMO.Excel | 851.9% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table-autofit | ClosedXML | 769.32 ms | 552.9 MB | 1.1 MB | OfficeIMO.Excel | 1920.7% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-direct-export | OfficeIMO.Excel | 31.59 ms | 9.3 MB | 1.3 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-dataset-sparse-direct-export | LargeXlsx | 35.56 ms | 9.0 MB | 1.6 MB | OfficeIMO.Excel | 12.6% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-direct-export | MiniExcel | 92.72 ms | 105.6 MB | 1.8 MB | OfficeIMO.Excel | 193.5% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-direct-export | EPPlus | 471.40 ms | 132.5 MB | 1.4 MB | OfficeIMO.Excel | 1392.3% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-direct-export | ClosedXML | 523.90 ms | 273.8 MB | 1.5 MB | OfficeIMO.Excel | 1558.5% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-tables | OfficeIMO.Excel | 40.66 ms | 12.8 MB | 1.8 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-dataset-sparse-tables | MiniExcel | 94.25 ms | 105.6 MB | 1.8 MB | OfficeIMO.Excel | 131.8% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-tables | EPPlus | 472.69 ms | 132.5 MB | 1.4 MB | OfficeIMO.Excel | 1062.5% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-sparse-tables | ClosedXML | 520.92 ms | 273.8 MB | 1.5 MB | OfficeIMO.Excel | 1181.1% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables | OfficeIMO.Excel | 32.65 ms | 9.6 MB | 1.3 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-dataset-tables | MiniExcel | 73.08 ms | 94.8 MB | 1.5 MB | OfficeIMO.Excel | 123.8% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables | EPPlus | 309.60 ms | 108.1 MB | 1.1 MB | OfficeIMO.Excel | 848.1% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables | ClosedXML | 348.81 ms | 168.0 MB | 1.1 MB | OfficeIMO.Excel | 968.2% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables-autofit | OfficeIMO.Excel | 36.92 ms | 9.6 MB | 1.3 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-dataset-tables-autofit | MiniExcel | 79.01 ms | 125.8 MB | 1.5 MB | OfficeIMO.Excel | 114.0% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables-autofit | EPPlus | 383.06 ms | 190.7 MB | 1.1 MB | OfficeIMO.Excel | 937.6% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables-autofit | ClosedXML | 719.33 ms | 537.2 MB | 1.1 MB | OfficeIMO.Excel | 1848.5% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | LargeXlsx | 30.50 ms | 9.3 MB | 1.4 MB | LargeXlsx | 8.4% faster than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | OfficeIMO.Excel | 33.29 ms | 12.1 MB | 1.4 MB | LargeXlsx | Loss +9.2% |
| 25000 | package-profile | write-datatable-direct | MiniExcel | 75.13 ms | 90.2 MB | 1.5 MB | LargeXlsx | 125.7% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | ClosedXML | 259.77 ms | 101.8 MB | 1.1 MB | LargeXlsx | 680.3% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | EPPlus | 321.63 ms | 114.6 MB | 1.1 MB | LargeXlsx | 866.1% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-table-direct | OfficeIMO.Excel | 33.76 ms | 12.1 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-datatable-table-direct | MiniExcel | 74.23 ms | 90.2 MB | 1.5 MB | OfficeIMO.Excel | 119.8% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-table-direct | EPPlus | 319.32 ms | 114.6 MB | 1.1 MB | OfficeIMO.Excel | 845.7% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-table-direct | ClosedXML | 359.12 ms | 169.3 MB | 1.1 MB | OfficeIMO.Excel | 963.6% slower than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | LargeXlsx | 37.93 ms | 5.5 MB | 1.4 MB | LargeXlsx | 22.9% faster than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 49.21 ms | 12.3 MB | 1.4 MB | LargeXlsx | Loss +29.7% |
| 25000 | package-profile | write-fluent-rowsfrom-direct | MiniExcel | 97.96 ms | 91.1 MB | 1.5 MB | LargeXlsx | 99.1% slower than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | EPPlus | 401.65 ms | 114.6 MB | 1.1 MB | LargeXlsx | 716.2% slower than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | ClosedXML | 407.00 ms | 101.8 MB | 1.1 MB | LargeXlsx | 727.1% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 60.06 ms | 11.1 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 690.87 ms | 156.0 MB | 1.1 MB | OfficeIMO.Excel | 1050.4% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 1.15 s | 485.3 MB | 1.1 MB | OfficeIMO.Excel | 1809.9% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | LargeXlsx | 33.91 ms | 5.5 MB | 1.4 MB | LargeXlsx | 14.4% faster than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | OfficeIMO.Excel | 39.60 ms | 11.1 MB | 1.4 MB | LargeXlsx | Loss +16.8% |
| 25000 | package-profile | write-insertobjects-direct | MiniExcel | 82.37 ms | 91.1 MB | 1.5 MB | LargeXlsx | 108.0% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | ClosedXML | 344.89 ms | 101.8 MB | 1.1 MB | LargeXlsx | 770.8% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | EPPlus | 462.65 ms | 114.7 MB | 1.1 MB | LargeXlsx | 1068.2% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 41.82 ms | 9.6 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 394.55 ms | 156.0 MB | 1.1 MB | OfficeIMO.Excel | 843.5% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 717.96 ms | 485.3 MB | 1.1 MB | OfficeIMO.Excel | 1616.9% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 30.66 ms | 5.5 MB | 1.4 MB | LargeXlsx | 24.8% faster than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 40.79 ms | 9.6 MB | 1.4 MB | LargeXlsx | Loss +33.0% |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | MiniExcel | 69.39 ms | 91.1 MB | 1.5 MB | LargeXlsx | 70.1% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | ClosedXML | 293.04 ms | 101.8 MB | 1.1 MB | LargeXlsx | 618.4% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-flat-dictionaries-direct | EPPlus | 346.45 ms | 114.7 MB | 1.1 MB | LargeXlsx | 749.4% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | LargeXlsx | 28.10 ms | 5.5 MB | 1.4 MB | LargeXlsx | 39.6% faster than OfficeIMO |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | OfficeIMO.Excel | 46.52 ms | 15.3 MB | 1.4 MB | LargeXlsx | Loss +65.6% |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | MiniExcel | 69.20 ms | 91.1 MB | 1.5 MB | LargeXlsx | 48.7% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | ClosedXML | 328.48 ms | 101.8 MB | 1.1 MB | LargeXlsx | 606.0% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-legacy-dictionaries-direct | EPPlus | 386.19 ms | 114.7 MB | 1.1 MB | LargeXlsx | 730.1% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 36.82 ms | 11.1 MB | 1.4 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 422.33 ms | 135.1 MB | 1.1 MB | OfficeIMO.Excel | 1047.0% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 479.24 ms | 269.0 MB | 1.1 MB | OfficeIMO.Excel | 1201.5% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-mixed-objects-direct | LargeXlsx | 49.06 ms | 5.9 MB | 1.8 MB | LargeXlsx | 16.0% faster than OfficeIMO |
| 25000 | package-profile | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 58.38 ms | 10.2 MB | 1.8 MB | LargeXlsx | Loss +19.0% |
| 25000 | package-profile | write-powershell-mixed-objects-direct | MiniExcel | 108.24 ms | 111.3 MB | 1.9 MB | LargeXlsx | 85.4% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-mixed-objects-direct | ClosedXML | 453.78 ms | 175.3 MB | 1.5 MB | LargeXlsx | 677.3% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-mixed-objects-direct | EPPlus | 495.40 ms | 141.5 MB | 1.4 MB | LargeXlsx | 748.6% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-mixed-direct | LargeXlsx | 48.20 ms | 5.9 MB | 1.8 MB | LargeXlsx | 7.7% faster than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 52.24 ms | 9.6 MB | 1.8 MB | LargeXlsx | Loss +8.4% |
| 25000 | package-profile | write-powershell-psobject-mixed-direct | MiniExcel | 106.60 ms | 111.3 MB | 1.9 MB | LargeXlsx | 104.1% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-mixed-direct | ClosedXML | 468.45 ms | 175.3 MB | 1.5 MB | LargeXlsx | 796.7% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-mixed-direct | EPPlus | 521.63 ms | 141.5 MB | 1.4 MB | LargeXlsx | 898.5% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 213.38 ms | 35.1 MB | 6.6 MB | OfficeIMO.Excel | Win |
| 25000 | package-profile | write-powershell-psobject-wide-direct | LargeXlsx | 219.17 ms | 22.7 MB | 6.5 MB | OfficeIMO.Excel | 2.7% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-wide-direct | MiniExcel | 334.59 ms | 339.8 MB | 6.8 MB | OfficeIMO.Excel | 56.8% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-wide-direct | ClosedXML | 1.28 s | 476.0 MB | 6.0 MB | OfficeIMO.Excel | 500.0% slower than OfficeIMO |
| 25000 | package-profile | write-powershell-psobject-wide-direct | EPPlus | 1.61 s | 549.7 MB | 5.3 MB | OfficeIMO.Excel | 655.0% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | LargeXlsx | 11.90 ms | 2.7 MB |  | LargeXlsx | 33.9% faster than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | OfficeIMO.Excel | 18.02 ms | 11.4 MB |  | LargeXlsx | Loss +51.4% |
| 25000 | speed-comparison | append-plain-rows | MiniExcel | 33.10 ms | 56.9 MB |  | LargeXlsx | 83.7% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | EPPlus 4.5.3.3 | 110.08 ms | 0 B |  | LargeXlsx | 510.8% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | ClosedXML | 138.31 ms | 101.8 MB |  | LargeXlsx | 667.5% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | EPPlus | 215.93 ms | 97.9 MB |  | LargeXlsx | 1098.2% slower than OfficeIMO |
| 25000 | speed-comparison | autofit-existing | OfficeIMO.Excel | 345.37 ms | 135.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | autofit-existing | EPPlus | 434.91 ms | 245.0 MB |  | OfficeIMO.Excel | 25.9% slower than OfficeIMO |
| 25000 | speed-comparison | autofit-existing | EPPlus 4.5.3.3 | 551.81 ms | 0 B |  | OfficeIMO.Excel | 59.8% slower than OfficeIMO |
| 25000 | speed-comparison | autofit-existing | ClosedXML | 1.28 s | 810.1 MB |  | OfficeIMO.Excel | 271.2% slower than OfficeIMO |
| 25000 | speed-comparison | build-object-datatable-dictionaries | OfficeIMO.Excel | 9.59 ms | 5.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | build-object-datatable-typed | OfficeIMO.Excel | 7.00 ms | 7.9 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-cells | OfficeIMO.Excel | 49.58 ms | 24.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-cells | EPPlus | 240.62 ms | 183.0 MB |  | OfficeIMO.Excel | 385.3% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-cells | ClosedXML | 311.54 ms | 162.6 MB |  | OfficeIMO.Excel | 528.3% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-first-column-from-wide-sheet | OfficeIMO.Excel | 35.38 ms | 3.9 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-first-column-from-wide-sheet | EPPlus | 207.50 ms | 112.8 MB |  | OfficeIMO.Excel | 486.6% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-first-column-from-wide-sheet | ClosedXML | 291.84 ms | 147.4 MB |  | OfficeIMO.Excel | 725.0% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-range | OfficeIMO.Excel | 48.79 ms | 24.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-range | EPPlus | 254.69 ms | 183.0 MB |  | OfficeIMO.Excel | 422.0% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-range | ClosedXML | 311.93 ms | 162.6 MB |  | OfficeIMO.Excel | 539.3% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-top-range | OfficeIMO.Excel | 1.90 ms | 402.7 KB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | enumerate-top-range | EPPlus | 202.01 ms | 103.1 MB |  | OfficeIMO.Excel | 10508.2% slower than OfficeIMO |
| 25000 | speed-comparison | enumerate-top-range | ClosedXML | 283.85 ms | 145.9 MB |  | OfficeIMO.Excel | 14805.5% slower than OfficeIMO |
| 25000 | speed-comparison | formula-heavy-read | OfficeIMO.Excel | 21.56 ms | 6.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | formula-heavy-read | EPPlus 4.5.3.3 | 72.88 ms | 0 B |  | OfficeIMO.Excel | 238.1% slower than OfficeIMO |
| 25000 | speed-comparison | formula-heavy-read | EPPlus | 144.77 ms | 69.2 MB |  | OfficeIMO.Excel | 571.5% slower than OfficeIMO |
| 25000 | speed-comparison | formula-heavy-read | ClosedXML | 171.49 ms | 77.6 MB |  | OfficeIMO.Excel | 695.5% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | OfficeIMO.Excel | 15.26 ms | 15.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | large-shared-strings | MiniExcel | 29.36 ms | 72.0 MB |  | OfficeIMO.Excel | 92.4% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | EPPlus 4.5.3.3 | 95.69 ms | 0 B |  | OfficeIMO.Excel | 527.1% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | ClosedXML | 108.89 ms | 101.8 MB |  | OfficeIMO.Excel | 613.6% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | EPPlus | 173.88 ms | 82.4 MB |  | OfficeIMO.Excel | 1039.5% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | Sylvan.Data.Excel | 0.96 ms | 316.6 KB |  | Sylvan.Data.Excel | 40.4% faster than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | ExcelDataReader | 1.39 ms | 4.0 MB |  | Sylvan.Data.Excel | 14.0% faster than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | OfficeIMO.Excel | 1.61 ms | 248.8 KB |  | Sylvan.Data.Excel | Loss +67.8% |
| 25000 | speed-comparison | large-sparse-column-read | ClosedXML | 3.24 ms | 4.3 MB |  | Sylvan.Data.Excel | 101.0% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | MiniExcel | 10.93 ms | 45.1 MB |  | Sylvan.Data.Excel | 577.5% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | EPPlus 4.5.3.3 | 16.81 ms | 0 B |  | Sylvan.Data.Excel | 941.8% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | EPPlus | 89.83 ms | 42.1 MB |  | Sylvan.Data.Excel | 5466.0% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | Sylvan.Data.Excel | 0.96 ms | 316.6 KB |  | Sylvan.Data.Excel | 42.6% faster than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | ExcelDataReader | 1.38 ms | 4.0 MB |  | Sylvan.Data.Excel | 17.3% faster than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | OfficeIMO.Excel | 1.67 ms | 248.9 KB |  | Sylvan.Data.Excel | Loss +74.3% |
| 25000 | speed-comparison | large-sparse-row-read | ClosedXML | 3.22 ms | 4.3 MB |  | Sylvan.Data.Excel | 92.8% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | MiniExcel | 11.22 ms | 45.1 MB |  | Sylvan.Data.Excel | 572.1% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | EPPlus 4.5.3.3 | 16.58 ms | 0 B |  | Sylvan.Data.Excel | 893.0% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | EPPlus | 90.18 ms | 42.1 MB |  | Sylvan.Data.Excel | 5300.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | OfficeIMO.Excel | 34.26 ms | 1.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-bottom-range | Sylvan.Data.Excel | 42.52 ms | 3.5 MB |  | OfficeIMO.Excel | 24.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | ExcelDataReader | 104.28 ms | 59.8 MB |  | OfficeIMO.Excel | 204.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | MiniExcel | 121.68 ms | 182.0 MB |  | OfficeIMO.Excel | 255.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | EPPlus | 204.67 ms | 103.1 MB |  | OfficeIMO.Excel | 497.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range | ClosedXML | 285.77 ms | 145.9 MB |  | OfficeIMO.Excel | 734.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | OfficeIMO.Excel | 36.68 ms | 1.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-bottom-range-stream | Sylvan.Data.Excel | 43.76 ms | 3.5 MB |  | OfficeIMO.Excel | 19.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | ExcelDataReader | 107.47 ms | 59.8 MB |  | OfficeIMO.Excel | 193.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | MiniExcel | 116.86 ms | 182.0 MB |  | OfficeIMO.Excel | 218.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | EPPlus | 204.02 ms | 103.1 MB |  | OfficeIMO.Excel | 456.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-bottom-range-stream | ClosedXML | 285.75 ms | 145.9 MB |  | OfficeIMO.Excel | 679.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | OfficeIMO.Excel | 56.83 ms | 34.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-datatable | Sylvan.Data.Excel | 58.57 ms | 18.0 MB |  | OfficeIMO.Excel | 3.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | ExcelDataReader | 123.90 ms | 74.3 MB |  | OfficeIMO.Excel | 118.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | MiniExcel | 141.87 ms | 177.0 MB |  | OfficeIMO.Excel | 149.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | EPPlus 4.5.3.3 | 202.05 ms | 0 B |  | OfficeIMO.Excel | 255.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | EPPlus | 252.00 ms | 197.5 MB |  | OfficeIMO.Excel | 343.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | ClosedXML | 316.21 ms | 174.3 MB |  | OfficeIMO.Excel | 456.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | OfficeIMO.Excel | 36.37 ms | 4.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | Sylvan.Data.Excel | 46.00 ms | 4.2 MB |  | OfficeIMO.Excel | 26.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | MiniExcel | 95.20 ms | 154.9 MB |  | OfficeIMO.Excel | 161.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | ExcelDataReader | 104.41 ms | 59.8 MB |  | OfficeIMO.Excel | 187.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | EPPlus | 211.90 ms | 112.8 MB |  | OfficeIMO.Excel | 482.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-first-column-from-wide-sheet | ClosedXML | 304.83 ms | 147.4 MB |  | OfficeIMO.Excel | 738.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | Sylvan.Data.Excel | 47.00 ms | 5.7 MB |  | Sylvan.Data.Excel | 5.0% faster than OfficeIMO |
| 25000 | speed-comparison | read-objects | OfficeIMO.Excel | 49.48 ms | 23.1 MB |  | Sylvan.Data.Excel | Loss +5.3% |
| 25000 | speed-comparison | read-objects | ExcelDataReader | 113.78 ms | 62.0 MB |  | Sylvan.Data.Excel | 130.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | MiniExcel | 140.40 ms | 179.3 MB |  | Sylvan.Data.Excel | 183.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | EPPlus 4.5.3.3 | 160.99 ms | 0 B |  | Sylvan.Data.Excel | 225.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | EPPlus | 246.55 ms | 194.9 MB |  | Sylvan.Data.Excel | 398.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | ClosedXML | 302.83 ms | 161.7 MB |  | Sylvan.Data.Excel | 512.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | Sylvan.Data.Excel | 45.79 ms | 5.2 MB |  | Sylvan.Data.Excel | 2.0% faster than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | OfficeIMO.Excel | 46.73 ms | 23.1 MB |  | Sylvan.Data.Excel | Loss +2.1% |
| 25000 | speed-comparison | read-objects-stream | ExcelDataReader | 102.85 ms | 61.5 MB |  | Sylvan.Data.Excel | 120.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | MiniExcel | 121.23 ms | 178.8 MB |  | Sylvan.Data.Excel | 159.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | EPPlus 4.5.3.3 | 161.62 ms | 0 B |  | Sylvan.Data.Excel | 245.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | EPPlus | 239.12 ms | 194.7 MB |  | Sylvan.Data.Excel | 411.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | ClosedXML | 296.22 ms | 161.5 MB |  | Sylvan.Data.Excel | 533.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | Sylvan.Data.Excel | 51.30 ms | 3.5 MB |  | Sylvan.Data.Excel | 3.6% faster than OfficeIMO |
| 25000 | speed-comparison | read-range | OfficeIMO.Excel | 53.24 ms | 25.6 MB |  | Sylvan.Data.Excel | Loss +3.8% |
| 25000 | speed-comparison | read-range | ExcelDataReader | 129.90 ms | 59.8 MB |  | Sylvan.Data.Excel | 144.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | MiniExcel | 130.41 ms | 182.0 MB |  | Sylvan.Data.Excel | 145.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | EPPlus 4.5.3.3 | 178.03 ms | 0 B |  | Sylvan.Data.Excel | 234.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | EPPlus | 260.23 ms | 183.0 MB |  | Sylvan.Data.Excel | 388.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | ClosedXML | 347.41 ms | 159.8 MB |  | Sylvan.Data.Excel | 552.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | OfficeIMO.Excel | 53.65 ms | 26.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-range-decimal | Sylvan.Data.Excel | 55.13 ms | 4.4 MB |  | OfficeIMO.Excel | 2.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | ExcelDataReader | 111.19 ms | 59.8 MB |  | OfficeIMO.Excel | 107.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | MiniExcel | 127.15 ms | 182.0 MB |  | OfficeIMO.Excel | 137.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | EPPlus | 246.53 ms | 183.0 MB |  | OfficeIMO.Excel | 359.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-decimal | ClosedXML | 346.89 ms | 159.8 MB |  | OfficeIMO.Excel | 546.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | OfficeIMO.Excel | 45.91 ms | 26.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | read-range-stream | Sylvan.Data.Excel | 47.62 ms | 3.5 MB |  | OfficeIMO.Excel | 3.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | ExcelDataReader | 106.17 ms | 59.8 MB |  | OfficeIMO.Excel | 131.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | MiniExcel | 124.97 ms | 182.0 MB |  | OfficeIMO.Excel | 172.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | EPPlus 4.5.3.3 | 166.86 ms | 0 B |  | OfficeIMO.Excel | 263.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | EPPlus | 240.42 ms | 183.0 MB |  | OfficeIMO.Excel | 423.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | ClosedXML | 304.41 ms | 159.8 MB |  | OfficeIMO.Excel | 563.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | Sylvan.Data.Excel | 0.45 ms | 348.5 KB |  | Sylvan.Data.Excel | 76.2% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range | MiniExcel | 0.71 ms | 858.3 KB |  | Sylvan.Data.Excel | 62.5% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range | OfficeIMO.Excel | 1.90 ms | 416.1 KB |  | Sylvan.Data.Excel | Loss +319.8% |
| 25000 | speed-comparison | read-top-range | ExcelDataReader | 40.42 ms | 16.7 MB |  | Sylvan.Data.Excel | 2025.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | EPPlus 4.5.3.3 | 129.14 ms | 0 B |  | Sylvan.Data.Excel | 6691.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | EPPlus | 205.92 ms | 103.1 MB |  | Sylvan.Data.Excel | 10728.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | ClosedXML | 292.07 ms | 145.9 MB |  | Sylvan.Data.Excel | 15258.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | Sylvan.Data.Excel | 0.44 ms | 348.5 KB |  | Sylvan.Data.Excel | 75.8% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | MiniExcel | 0.77 ms | 858.3 KB |  | Sylvan.Data.Excel | 57.6% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | OfficeIMO.Excel | 1.81 ms | 419.5 KB |  | Sylvan.Data.Excel | Loss +313.2% |
| 25000 | speed-comparison | read-top-range-stream | ExcelDataReader | 38.28 ms | 16.7 MB |  | Sylvan.Data.Excel | 2017.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | EPPlus 4.5.3.3 | 134.20 ms | 0 B |  | Sylvan.Data.Excel | 7324.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | EPPlus | 217.44 ms | 103.1 MB |  | Sylvan.Data.Excel | 11928.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | ClosedXML | 292.44 ms | 145.9 MB |  | Sylvan.Data.Excel | 16078.3% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | Sylvan.Data.Excel | 0.47 ms | 348.5 KB |  | Sylvan.Data.Excel | 75.3% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | MiniExcel | 0.71 ms | 858.3 KB |  | Sylvan.Data.Excel | 62.8% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | OfficeIMO.Excel | 1.91 ms | 420.2 KB |  | Sylvan.Data.Excel | Loss +305.3% |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | ExcelDataReader | 37.01 ms | 16.7 MB |  | Sylvan.Data.Excel | 1841.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | EPPlus | 200.20 ms | 103.1 MB |  | Sylvan.Data.Excel | 10402.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream-small-chunks | ClosedXML | 281.38 ms | 145.9 MB |  | Sylvan.Data.Excel | 14661.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-used-range | Sylvan.Data.Excel | 48.19 ms | 3.5 MB |  | Sylvan.Data.Excel | 44.4% faster than OfficeIMO |
| 25000 | speed-comparison | read-used-range | OfficeIMO.Excel | 86.61 ms | 33.4 MB |  | Sylvan.Data.Excel | Loss +79.7% |
| 25000 | speed-comparison | read-used-range | ExcelDataReader | 115.29 ms | 59.8 MB |  | Sylvan.Data.Excel | 33.1% slower than OfficeIMO |
| 25000 | speed-comparison | read-used-range | MiniExcel | 126.48 ms | 182.0 MB |  | Sylvan.Data.Excel | 46.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-used-range | EPPlus | 249.07 ms | 183.0 MB |  | Sylvan.Data.Excel | 187.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-used-range | ClosedXML | 321.54 ms | 159.8 MB |  | Sylvan.Data.Excel | 271.3% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-autofilter | OfficeIMO.Excel | 31.29 ms | 11.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-autofilter | EPPlus 4.5.3.3 | 237.76 ms | 0 B |  | OfficeIMO.Excel | 660.0% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-autofilter | ClosedXML | 290.20 ms | 205.7 MB |  | OfficeIMO.Excel | 827.6% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-autofilter | EPPlus | 357.32 ms | 206.8 MB |  | OfficeIMO.Excel | 1042.1% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-charts | OfficeIMO.Excel | 32.35 ms | 12.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-charts | EPPlus 4.5.3.3 | 238.29 ms | 0 B |  | OfficeIMO.Excel | 636.5% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-charts | EPPlus | 343.32 ms | 209.8 MB |  | OfficeIMO.Excel | 961.2% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-conditional-formatting | OfficeIMO.Excel | 30.43 ms | 11.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-conditional-formatting | EPPlus 4.5.3.3 | 237.67 ms | 0 B |  | OfficeIMO.Excel | 680.9% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-conditional-formatting | ClosedXML | 303.59 ms | 205.8 MB |  | OfficeIMO.Excel | 897.5% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-conditional-formatting | EPPlus | 346.54 ms | 206.9 MB |  | OfficeIMO.Excel | 1038.7% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-data-validation | OfficeIMO.Excel | 30.51 ms | 11.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-data-validation | EPPlus 4.5.3.3 | 246.27 ms | 0 B |  | OfficeIMO.Excel | 707.3% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-data-validation | ClosedXML | 285.27 ms | 205.7 MB |  | OfficeIMO.Excel | 835.1% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-data-validation | EPPlus | 341.09 ms | 206.8 MB |  | OfficeIMO.Excel | 1018.1% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-freeze-panes | OfficeIMO.Excel | 30.76 ms | 11.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-freeze-panes | EPPlus 4.5.3.3 | 239.74 ms | 0 B |  | OfficeIMO.Excel | 679.5% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-freeze-panes | ClosedXML | 288.10 ms | 205.7 MB |  | OfficeIMO.Excel | 836.7% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-freeze-panes | EPPlus | 350.41 ms | 206.9 MB |  | OfficeIMO.Excel | 1039.3% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-pivot-table | OfficeIMO.Excel | 88.29 ms | 98.5 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-pivot-table | EPPlus 4.5.3.3 | 240.67 ms | 0 B |  | OfficeIMO.Excel | 172.6% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-pivot-table | EPPlus | 388.13 ms | 225.3 MB |  | OfficeIMO.Excel | 339.6% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-all-in-one | OfficeIMO.Excel | 98.88 ms | 99.9 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-all-in-one | EPPlus | 398.97 ms | 270.5 MB |  | OfficeIMO.Excel | 303.5% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-all-in-one | EPPlus 4.5.3.3 | 471.99 ms | 0 B |  | OfficeIMO.Excel | 377.3% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-core | OfficeIMO.Excel | 34.16 ms | 11.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | realworld-report-core | EPPlus | 365.15 ms | 249.0 MB |  | OfficeIMO.Excel | 969.1% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-core | EPPlus 4.5.3.3 | 456.65 ms | 0 B |  | OfficeIMO.Excel | 1236.9% slower than OfficeIMO |
| 25000 | speed-comparison | realworld-report-core | ClosedXML | 781.87 ms | 664.2 MB |  | OfficeIMO.Excel | 2189.1% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook | OfficeIMO.Excel | 47.84 ms | 14.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | report-workbook | EPPlus | 540.39 ms | 356.1 MB |  | OfficeIMO.Excel | 1029.5% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook | EPPlus 4.5.3.3 | 634.90 ms | 0 B |  | OfficeIMO.Excel | 1227.1% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-core | OfficeIMO.Excel | 44.99 ms | 10.6 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | report-workbook-core | EPPlus | 509.77 ms | 334.7 MB |  | OfficeIMO.Excel | 1033.0% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-core | EPPlus 4.5.3.3 | 628.83 ms | 0 B |  | OfficeIMO.Excel | 1297.6% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-core | ClosedXML | 1.06 s | 952.9 MB |  | OfficeIMO.Excel | 2259.4% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-datatable | OfficeIMO.Excel | 50.01 ms | 16.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | report-workbook-datatable | EPPlus | 523.07 ms | 241.9 MB |  | OfficeIMO.Excel | 946.0% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-datatable | EPPlus 4.5.3.3 | 638.00 ms | 0 B |  | OfficeIMO.Excel | 1175.8% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-datatable-core | OfficeIMO.Excel | 47.36 ms | 13.2 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | report-workbook-datatable-core | EPPlus | 495.55 ms | 220.6 MB |  | OfficeIMO.Excel | 946.5% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-datatable-core | EPPlus 4.5.3.3 | 627.62 ms | 0 B |  | OfficeIMO.Excel | 1225.3% slower than OfficeIMO |
| 25000 | speed-comparison | report-workbook-datatable-core | ClosedXML | 1.02 s | 812.7 MB |  | OfficeIMO.Excel | 2053.6% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | Sylvan.Data.Excel | 17.58 ms | 1.9 MB |  | Sylvan.Data.Excel | 15.2% faster than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | OfficeIMO.Excel | 20.74 ms | 9.1 MB |  | Sylvan.Data.Excel | Loss +18.0% |
| 25000 | speed-comparison | shared-string-read | ExcelDataReader | 44.01 ms | 24.4 MB |  | Sylvan.Data.Excel | 112.2% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | MiniExcel | 49.00 ms | 72.7 MB |  | Sylvan.Data.Excel | 136.3% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | EPPlus 4.5.3.3 | 83.68 ms | 0 B |  | Sylvan.Data.Excel | 303.5% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | EPPlus | 145.48 ms | 87.3 MB |  | Sylvan.Data.Excel | 601.5% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | ClosedXML | 147.84 ms | 88.3 MB |  | Sylvan.Data.Excel | 612.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-blog-2023-20-string-columns | LargeXlsx | 40.03 ms | 10.5 MB |  | LargeXlsx | 6.1% faster than OfficeIMO |
| 25000 | speed-comparison | write-blog-2023-20-string-columns | OfficeIMO.Excel | 42.64 ms | 11.3 MB |  | LargeXlsx | Loss +6.5% |
| 25000 | speed-comparison | write-blog-2023-20-string-columns | MiniExcel | 146.06 ms | 221.6 MB |  | LargeXlsx | 242.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-blog-2023-20-string-columns | ClosedXML | 926.06 ms | 742.0 MB |  | LargeXlsx | 2071.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | OfficeIMO.Excel | 43.86 ms | 11.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-bulk-report | MiniExcel | 75.86 ms | 122.6 MB |  | OfficeIMO.Excel | 73.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | EPPlus | 451.91 ms | 248.9 MB |  | OfficeIMO.Excel | 930.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | EPPlus 4.5.3.3 | 519.39 ms | 0 B |  | OfficeIMO.Excel | 1084.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | ClosedXML | 966.22 ms | 552.7 MB |  | OfficeIMO.Excel | 2103.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellformula | OfficeIMO.Excel | 21.20 ms | 9.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellformula | EPPlus 4.5.3.3 | 131.18 ms | 0 B |  | OfficeIMO.Excel | 518.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellformula | ClosedXML | 179.99 ms | 111.2 MB |  | OfficeIMO.Excel | 748.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellformula | EPPlus | 323.42 ms | 137.4 MB |  | OfficeIMO.Excel | 1425.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-empty-strings | OfficeIMO.Excel | 12.92 ms | 6.6 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-empty-strings | ClosedXML | 108.12 ms | 90.7 MB |  | OfficeIMO.Excel | 737.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-empty-strings | EPPlus | 163.97 ms | 72.7 MB |  | OfficeIMO.Excel | 1169.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-numbers | OfficeIMO.Excel | 16.92 ms | 5.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-numbers | EPPlus 4.5.3.3 | 100.49 ms | 0 B |  | OfficeIMO.Excel | 493.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-numbers | ClosedXML | 109.60 ms | 82.2 MB |  | OfficeIMO.Excel | 547.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-numbers | EPPlus | 186.60 ms | 84.3 MB |  | OfficeIMO.Excel | 1002.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-mixed | OfficeIMO.Excel | 20.04 ms | 7.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-object-mixed | EPPlus 4.5.3.3 | 115.97 ms | 0 B |  | OfficeIMO.Excel | 478.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-mixed | ClosedXML | 175.43 ms | 108.5 MB |  | OfficeIMO.Excel | 775.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-mixed | EPPlus | 225.75 ms | 110.5 MB |  | OfficeIMO.Excel | 1026.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-sparse | OfficeIMO.Excel | 20.10 ms | 7.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-object-sparse | ClosedXML | 149.88 ms | 102.7 MB |  | OfficeIMO.Excel | 645.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-sparse | EPPlus | 209.57 ms | 103.8 MB |  | OfficeIMO.Excel | 942.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 16.76 ms | 7.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-object-sparse-batch | ClosedXML | 146.78 ms | 102.7 MB |  | OfficeIMO.Excel | 776.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-object-sparse-batch | EPPlus | 212.57 ms | 103.8 MB |  | OfficeIMO.Excel | 1168.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-scalars | OfficeIMO.Excel | 11.89 ms | 5.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-scalars | EPPlus 4.5.3.3 | 96.74 ms | 0 B |  | OfficeIMO.Excel | 713.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-scalars | ClosedXML | 102.20 ms | 80.6 MB |  | OfficeIMO.Excel | 759.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-scalars | EPPlus | 185.08 ms | 83.1 MB |  | OfficeIMO.Excel | 1457.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings | OfficeIMO.Excel | 18.65 ms | 14.7 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-strings | EPPlus 4.5.3.3 | 92.38 ms | 0 B |  | OfficeIMO.Excel | 395.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings | ClosedXML | 116.93 ms | 101.8 MB |  | OfficeIMO.Excel | 527.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings | EPPlus | 192.78 ms | 82.4 MB |  | OfficeIMO.Excel | 933.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings-distinct | OfficeIMO.Excel | 13.14 ms | 13.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-strings-distinct | ClosedXML | 145.31 ms | 128.4 MB |  | OfficeIMO.Excel | 1005.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings-distinct | EPPlus | 199.21 ms | 95.4 MB |  | OfficeIMO.Excel | 1415.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings-repeated | OfficeIMO.Excel | 16.32 ms | 7.0 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-strings-repeated | ClosedXML | 111.50 ms | 82.5 MB |  | OfficeIMO.Excel | 583.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-strings-repeated | EPPlus | 235.21 ms | 68.3 MB |  | OfficeIMO.Excel | 1340.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-temporal | OfficeIMO.Excel | 20.70 ms | 7.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalue-temporal | EPPlus 4.5.3.3 | 102.09 ms | 0 B |  | OfficeIMO.Excel | 393.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-temporal | ClosedXML | 144.39 ms | 87.2 MB |  | OfficeIMO.Excel | 597.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalue-temporal | EPPlus | 199.40 ms | 101.3 MB |  | OfficeIMO.Excel | 863.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 15.73 ms | 6.6 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 16.19 ms | 3.4 MB |  | OfficeIMO.Excel | 2.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-headerless-rectangle-direct | ClosedXML | 202.46 ms | 93.8 MB |  | OfficeIMO.Excel | 1186.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-headerless-rectangle-direct | EPPlus | 258.46 ms | 85.4 MB |  | OfficeIMO.Excel | 1542.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | LargeXlsx | 30.47 ms | 5.5 MB |  | LargeXlsx | 27.4% faster than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 41.96 ms | 15.3 MB |  | LargeXlsx | Loss +37.7% |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | MiniExcel | 66.03 ms | 91.1 MB |  | LargeXlsx | 57.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 237.95 ms | 0 B |  | LargeXlsx | 467.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | ClosedXML | 342.73 ms | 205.7 MB |  | LargeXlsx | 716.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus | 357.01 ms | 206.8 MB |  | LargeXlsx | 750.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-sparse-rectangle-direct | OfficeIMO.Excel | 33.10 ms | 7.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-cellvalues-sparse-rectangle-direct | ClosedXML | 286.93 ms | 102.8 MB |  | OfficeIMO.Excel | 766.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-sparse-rectangle-direct | EPPlus | 318.34 ms | 103.8 MB |  | OfficeIMO.Excel | 861.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | Sylvan.Data.Excel | 27.64 ms | 5.6 MB |  | Sylvan.Data.Excel | 21.6% faster than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | LargeXlsx | 33.94 ms | 8.1 MB |  | Sylvan.Data.Excel | 3.7% faster than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | OfficeIMO.Excel | 35.24 ms | 12.4 MB |  | Sylvan.Data.Excel | Loss +27.5% |
| 25000 | speed-comparison | write-datareader-plain | MiniExcel | 70.92 ms | 90.0 MB |  | Sylvan.Data.Excel | 101.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | EPPlus 4.5.3.3 | 221.32 ms | 0 B |  | Sylvan.Data.Excel | 527.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | ClosedXML | 269.04 ms | 101.8 MB |  | Sylvan.Data.Excel | 663.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-plain | EPPlus | 325.80 ms | 114.6 MB |  | Sylvan.Data.Excel | 824.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | OfficeIMO.Excel | 36.75 ms | 12.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datareader-table | MiniExcel | 69.47 ms | 90.0 MB |  | OfficeIMO.Excel | 89.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | EPPlus 4.5.3.3 | 222.38 ms | 0 B |  | OfficeIMO.Excel | 505.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | EPPlus | 329.10 ms | 114.6 MB |  | OfficeIMO.Excel | 795.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | ClosedXML | 375.34 ms | 169.3 MB |  | OfficeIMO.Excel | 921.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table-autofit | OfficeIMO.Excel | 38.06 ms | 12.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datareader-table-autofit | MiniExcel | 70.60 ms | 121.6 MB |  | OfficeIMO.Excel | 85.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table-autofit | EPPlus | 380.41 ms | 155.9 MB |  | OfficeIMO.Excel | 899.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table-autofit | EPPlus 4.5.3.3 | 444.55 ms | 0 B |  | OfficeIMO.Excel | 1068.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table-autofit | ClosedXML | 808.69 ms | 552.9 MB |  | OfficeIMO.Excel | 2024.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-headerless-tables | OfficeIMO.Excel | 36.09 ms | 12.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-headerless-tables | MiniExcel | 88.55 ms | 94.8 MB |  | OfficeIMO.Excel | 145.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-headerless-tables | ClosedXML | 390.36 ms | 168.0 MB |  | OfficeIMO.Excel | 981.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-headerless-tables | EPPlus | 436.16 ms | 108.6 MB |  | OfficeIMO.Excel | 1108.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | OfficeIMO.Excel | 40.02 ms | 9.3 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | LargeXlsx | 48.19 ms | 9.0 MB |  | OfficeIMO.Excel | 20.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | MiniExcel | 102.42 ms | 105.6 MB |  | OfficeIMO.Excel | 155.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | EPPlus | 548.05 ms | 132.5 MB |  | OfficeIMO.Excel | 1269.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-direct-export | ClosedXML | 563.87 ms | 273.8 MB |  | OfficeIMO.Excel | 1309.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-tables | OfficeIMO.Excel | 45.39 ms | 12.8 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-sparse-tables | MiniExcel | 128.26 ms | 105.6 MB |  | OfficeIMO.Excel | 182.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-tables | EPPlus | 547.82 ms | 132.5 MB |  | OfficeIMO.Excel | 1107.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-tables | ClosedXML | 587.97 ms | 273.8 MB |  | OfficeIMO.Excel | 1195.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | OfficeIMO.Excel | 41.49 ms | 9.6 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-tables | MiniExcel | 95.54 ms | 94.8 MB |  | OfficeIMO.Excel | 130.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | EPPlus 4.5.3.3 | 213.69 ms | 0 B |  | OfficeIMO.Excel | 415.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | EPPlus | 359.86 ms | 108.1 MB |  | OfficeIMO.Excel | 767.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | ClosedXML | 421.48 ms | 168.0 MB |  | OfficeIMO.Excel | 915.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables-autofit | OfficeIMO.Excel | 41.98 ms | 9.6 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-dataset-tables-autofit | MiniExcel | 85.39 ms | 125.8 MB |  | OfficeIMO.Excel | 103.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables-autofit | EPPlus | 407.26 ms | 190.7 MB |  | OfficeIMO.Excel | 870.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables-autofit | ClosedXML | 788.34 ms | 537.2 MB |  | OfficeIMO.Excel | 1778.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | LargeXlsx | 37.99 ms | 9.3 MB |  | LargeXlsx | 13.6% faster than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | OfficeIMO.Excel | 43.95 ms | 12.1 MB |  | LargeXlsx | Loss +15.7% |
| 25000 | speed-comparison | write-datatable-direct | MiniExcel | 81.01 ms | 90.2 MB |  | LargeXlsx | 84.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | EPPlus 4.5.3.3 | 217.29 ms | 0 B |  | LargeXlsx | 394.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | ClosedXML | 294.38 ms | 101.8 MB |  | LargeXlsx | 569.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | EPPlus | 354.88 ms | 114.6 MB |  | LargeXlsx | 707.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-object-table-direct | OfficeIMO.Excel | 35.92 ms | 9.4 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datatable-object-table-direct | MiniExcel | 78.36 ms | 87.6 MB |  | OfficeIMO.Excel | 118.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-object-table-direct | EPPlus | 322.71 ms | 112.0 MB |  | OfficeIMO.Excel | 798.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-object-table-direct | ClosedXML | 370.25 ms | 166.7 MB |  | OfficeIMO.Excel | 930.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | OfficeIMO.Excel | 40.36 ms | 12.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-datatable-table-direct | MiniExcel | 111.29 ms | 90.2 MB |  | OfficeIMO.Excel | 175.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | EPPlus 4.5.3.3 | 223.64 ms | 0 B |  | OfficeIMO.Excel | 454.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | EPPlus | 407.39 ms | 114.7 MB |  | OfficeIMO.Excel | 909.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | ClosedXML | 413.73 ms | 169.3 MB |  | OfficeIMO.Excel | 925.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-dictionary-objects-table-direct | OfficeIMO.Excel | 46.09 ms | 14.5 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | LargeXlsx | 28.34 ms | 5.5 MB |  | LargeXlsx | 14.4% faster than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 33.11 ms | 12.3 MB |  | LargeXlsx | Loss +16.8% |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | MiniExcel | 63.74 ms | 91.1 MB |  | LargeXlsx | 92.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 228.79 ms | 0 B |  | LargeXlsx | 591.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | ClosedXML | 279.93 ms | 101.8 MB |  | LargeXlsx | 745.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus | 348.97 ms | 114.6 MB |  | LargeXlsx | 954.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 36.60 ms | 11.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 371.98 ms | 155.9 MB |  | OfficeIMO.Excel | 916.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 711.64 ms | 485.3 MB |  | OfficeIMO.Excel | 1844.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | LargeXlsx | 28.51 ms | 5.5 MB |  | LargeXlsx | 12.5% faster than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | OfficeIMO.Excel | 32.57 ms | 11.1 MB |  | LargeXlsx | Loss +14.2% |
| 25000 | speed-comparison | write-insertobjects-direct | MiniExcel | 63.60 ms | 91.1 MB |  | LargeXlsx | 95.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | EPPlus 4.5.3.3 | 219.47 ms | 0 B |  | LargeXlsx | 573.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | ClosedXML | 279.24 ms | 101.8 MB |  | LargeXlsx | 757.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | EPPlus | 340.20 ms | 114.6 MB |  | LargeXlsx | 944.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 41.65 ms | 9.6 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 368.05 ms | 155.9 MB |  | OfficeIMO.Excel | 783.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 694.81 ms | 485.3 MB |  | OfficeIMO.Excel | 1568.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 28.54 ms | 5.5 MB |  | LargeXlsx | 26.3% faster than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 38.73 ms | 9.6 MB |  | LargeXlsx | Loss +35.7% |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | MiniExcel | 64.00 ms | 91.1 MB |  | LargeXlsx | 65.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | ClosedXML | 281.39 ms | 101.8 MB |  | LargeXlsx | 626.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-flat-dictionaries-direct | EPPlus | 334.00 ms | 114.6 MB |  | LargeXlsx | 762.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 33.99 ms | 11.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 358.72 ms | 135.0 MB |  | OfficeIMO.Excel | 955.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 438.00 ms | 269.0 MB |  | OfficeIMO.Excel | 1188.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | LargeXlsx | 37.45 ms | 5.9 MB |  | LargeXlsx | 14.4% faster than OfficeIMO |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 43.73 ms | 10.2 MB |  | LargeXlsx | Loss +16.8% |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | MiniExcel | 79.83 ms | 111.3 MB |  | LargeXlsx | 82.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | ClosedXML | 369.08 ms | 175.3 MB |  | LargeXlsx | 744.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-mixed-objects-direct | EPPlus | 440.36 ms | 141.5 MB |  | LargeXlsx | 907.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-mixed-direct | LargeXlsx | 39.21 ms | 5.9 MB |  | LargeXlsx | 10.1% faster than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 43.61 ms | 9.5 MB |  | LargeXlsx | Loss +11.2% |
| 25000 | speed-comparison | write-powershell-psobject-mixed-direct | MiniExcel | 80.87 ms | 111.3 MB |  | LargeXlsx | 85.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-mixed-direct | ClosedXML | 379.99 ms | 175.3 MB |  | LargeXlsx | 771.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-mixed-direct | EPPlus | 444.45 ms | 141.5 MB |  | LargeXlsx | 919.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 188.84 ms | 35.1 MB |  | OfficeIMO.Excel | Win |
| 25000 | speed-comparison | write-powershell-psobject-wide-direct | LargeXlsx | 202.63 ms | 22.7 MB |  | OfficeIMO.Excel | 7.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-wide-direct | MiniExcel | 330.00 ms | 339.8 MB |  | OfficeIMO.Excel | 74.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-wide-direct | ClosedXML | 1.21 s | 476.0 MB |  | OfficeIMO.Excel | 541.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-powershell-psobject-wide-direct | EPPlus | 1.54 s | 549.7 MB |  | OfficeIMO.Excel | 713.8% slower than OfficeIMO |
| 300000 | focused-package-profile | write-blog-2023-20-string-columns | LargeXlsx | 636.47 ms | 93.1 MB | 28.6 MB | LargeXlsx | Win |
| 300000 | focused-package-profile | write-blog-2023-20-string-columns | OfficeIMO.Excel | 701.65 ms | 173.4 MB | 26.6 MB | LargeXlsx | 1.10x vs best |
| 300000 | focused-package-profile | write-blog-2023-20-string-columns | MiniExcel | 2.28 s | 2.46 GB | 28.5 MB | LargeXlsx | 3.58x vs best |
| 300000 | focused-package-profile | write-blog-2023-20-string-columns | ClosedXML | 15.89 s | 8.51 GB | 31.0 MB | LargeXlsx | 24.97x vs best |
