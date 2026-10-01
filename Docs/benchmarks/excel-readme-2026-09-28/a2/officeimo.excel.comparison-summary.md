# OfficeIMO.Excel Comparison Summary

This is the suite-level decision table. Mean, standard deviation, standard error, ratios, and allocations come from the lightweight rotated local runner; they are meant for engineering direction. Results within 5% are practical ties. Use the BenchmarkDotNet benchmark classes when a publication-grade `Error` column is required.

## At a glance

| Row count | Artifact | Workload | Category | OfficeIMO wins | OfficeIMO ties | OfficeIMO losses | Biggest loss |
| ---: | --- | --- | --- | ---: | ---: | ---: | --- |
| 25000 | speed-comparison | other | Real-world report | 1 | 0 | 0 |  |
| 25000 | speed-comparison | read | Typed object read | 1 | 0 | 0 |  |
| 25000 | speed-comparison | write | DataTable table export | 1 | 1 | 0 |  |

## OfficeIMO decision table

| Row count | Artifact | Workload | Category | Scenario | OfficeIMO mean | Best | OfficeIMO vs best | Alloc | Package |
| ---: | --- | --- | --- | --- | ---: | --- | ---: | ---: | ---: |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | 56.96 ms | OfficeIMO.Excel | Win | 16582.1 KB |  |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | 31.50 ms | OfficeIMO.Excel | Win | 2033.7 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-compact-package | 26.09 ms | OfficeIMO.Excel, SpreadCheetah | Tie for best | 6061.2 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | 36.01 ms | OfficeIMO.Excel | Win | 6122.2 KB |  |

## Full comparison table

| Row count | Artifact | Workload | Category | Scenario | Library | Mean | StdDev | StdErr | Ratio to OfficeIMO | Ratio to best | Alloc | Alloc ratio | Package | Package ratio | Outcome |
| ---: | --- | --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | OfficeIMO.Excel | 56.96 ms | 14.58 ms | 4.86 ms | 1.00 | 1.00 | 16582.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus | 614.02 ms | 31.98 ms | 10.66 ms | 10.78 | 10.78 | 291819.0 KB | 17.60 |  |  | 977.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | OfficeIMO.Excel | 31.50 ms | 3.45 ms | 1.15 ms | 1.00 | 1.00 | 2033.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | Sylvan.Data.Excel | 62.10 ms | 4.40 ms | 1.47 ms | 1.97 | 1.97 | 2148.4 KB | 1.06 |  |  | 97.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | ExcelDataReader | 146.18 ms | 11.30 ms | 3.77 ms | 4.64 | 4.64 | 43493.0 KB | 21.39 |  |  | 364.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | MiniExcel | 186.17 ms | 12.39 ms | 4.13 ms | 5.91 | 5.91 | 179550.3 KB | 88.29 |  |  | 491.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | EPPlus | 319.20 ms | 22.11 ms | 7.37 ms | 10.13 | 10.13 | 183069.5 KB | 90.02 |  |  | 913.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | ClosedXML | 471.59 ms | 8.93 ms | 2.98 ms | 14.97 | 14.97 | 195722.4 KB | 96.24 |  |  | 1397.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-compact-package | OfficeIMO.Excel | 26.09 ms | 3.47 ms | 1.16 ms | 1.00 | 1.00 | 6061.2 KB | 1.00 |  |  | Tie for best |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-compact-package | SpreadCheetah | 26.76 ms | 3.97 ms | 1.32 ms | 1.03 | 1.03 | 5541.9 KB | 0.91 |  |  | Tie vs OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-compact-package | LargeXlsx | 32.55 ms | 5.73 ms | 1.91 ms | 1.25 | 1.25 | 5604.6 KB | 0.92 |  |  | 24.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-compact-package | Sylvan.Data.Excel | 34.45 ms | 4.21 ms | 1.40 ms | 1.32 | 1.32 | 5693.4 KB | 0.94 |  |  | 32.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | OfficeIMO.Excel | 36.01 ms | 24.09 ms | 8.03 ms | 1.00 | 1.00 | 6122.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | EPPlus | 437.17 ms | 69.57 ms | 23.19 ms | 12.14 | 12.14 | 117420.0 KB | 19.18 |  |  | 1113.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | ClosedXML | 550.24 ms | 101.67 ms | 33.89 ms | 15.28 | 15.28 | 173397.3 KB | 28.32 |  |  | 1427.8% slower than OfficeIMO |
