# OfficeIMO.Excel Comparison Summary

This is the suite-level decision table. Mean, standard deviation, standard error, ratios, and allocations come from the lightweight rotated local runner; they are meant for engineering direction. Results within 5% are practical ties. Use the BenchmarkDotNet benchmark classes when a publication-grade `Error` column is required.

## At a glance

| Row count | Artifact | Workload | Category | OfficeIMO wins | OfficeIMO ties | OfficeIMO losses | Biggest loss |
| ---: | --- | --- | --- | ---: | ---: | ---: | --- |
| 25000 | speed-comparison | other | Real-world report | 1 | 0 | 0 |  |
| 25000 | speed-comparison | read | Typed object read | 1 | 0 | 0 |  |
| 25000 | speed-comparison | write | DataTable table export | 2 | 0 | 0 |  |

## OfficeIMO decision table

| Row count | Artifact | Workload | Category | Scenario | OfficeIMO mean | Best | OfficeIMO vs best | Alloc | Package |
| ---: | --- | --- | --- | --- | ---: | --- | ---: | ---: | ---: |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | 39.78 ms | OfficeIMO.Excel | Win | 16582.7 KB |  |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | 25.26 ms | OfficeIMO.Excel | Win | 2033.7 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-compact-package | 21.60 ms | OfficeIMO.Excel | Win | 6061.2 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | 21.18 ms | OfficeIMO.Excel | Win | 6076.1 KB |  |

## Full comparison table

| Row count | Artifact | Workload | Category | Scenario | Library | Mean | StdDev | StdErr | Ratio to OfficeIMO | Ratio to best | Alloc | Alloc ratio | Package | Package ratio | Outcome |
| ---: | --- | --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | OfficeIMO.Excel | 39.78 ms | 1.39 ms | 0.46 ms | 1.00 | 1.00 | 16582.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus | 478.58 ms | 19.24 ms | 6.41 ms | 12.03 | 12.03 | 291819.0 KB | 17.60 |  |  | 1103.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | OfficeIMO.Excel | 25.26 ms | 1.83 ms | 0.61 ms | 1.00 | 1.00 | 2033.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | Sylvan.Data.Excel | 42.94 ms | 2.51 ms | 0.84 ms | 1.70 | 1.70 | 2148.4 KB | 1.06 |  |  | 70.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | ExcelDataReader | 105.28 ms | 7.95 ms | 2.65 ms | 4.17 | 4.17 | 43493.0 KB | 21.39 |  |  | 316.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | MiniExcel | 130.62 ms | 6.07 ms | 2.02 ms | 5.17 | 5.17 | 179550.4 KB | 88.29 |  |  | 417.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | EPPlus | 240.58 ms | 8.66 ms | 2.89 ms | 9.53 | 9.53 | 183069.5 KB | 90.02 |  |  | 852.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | ClosedXML | 344.66 ms | 13.88 ms | 4.63 ms | 13.65 | 13.65 | 195722.6 KB | 96.24 |  |  | 1264.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-compact-package | OfficeIMO.Excel | 21.60 ms | 1.63 ms | 0.54 ms | 1.00 | 1.00 | 6061.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-compact-package | SpreadCheetah | 23.12 ms | 1.86 ms | 0.62 ms | 1.07 | 1.07 | 5541.9 KB | 0.91 |  |  | 7.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-compact-package | LargeXlsx | 28.58 ms | 1.86 ms | 0.62 ms | 1.32 | 1.32 | 5604.6 KB | 0.92 |  |  | 32.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-compact-package | Sylvan.Data.Excel | 30.90 ms | 2.04 ms | 0.68 ms | 1.43 | 1.43 | 5693.4 KB | 0.94 |  |  | 43.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | OfficeIMO.Excel | 21.18 ms | 1.74 ms | 0.58 ms | 1.00 | 1.00 | 6076.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | EPPlus | 324.39 ms | 9.84 ms | 3.28 ms | 15.32 | 15.32 | 117390.1 KB | 19.32 |  |  | 1431.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | ClosedXML | 381.10 ms | 16.67 ms | 5.56 ms | 17.99 | 17.99 | 173393.8 KB | 28.54 |  |  | 1699.4% slower than OfficeIMO |
