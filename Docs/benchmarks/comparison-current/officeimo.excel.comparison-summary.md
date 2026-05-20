# OfficeIMO.Excel Comparison Summary

This is the suite-level decision table. Mean, standard deviation, standard error, ratios, and allocations come from the lightweight rotated local runner; they are meant for engineering direction. Use the BenchmarkDotNet benchmark classes when a publication-grade `Error` column is required.

## At a glance

| Row count | Artifact | Workload | OfficeIMO wins | OfficeIMO losses | Biggest loss |
| ---: | --- | --- | ---: | ---: | --- |
| 2500 | dense-helloworld-comparison | read | 0 | 2 | dense-helloworld-read-stream: Loss +89.3% vs Sylvan.Data.Excel |
| 2500 | package-profile | package | 12 | 0 |  |
| 2500 | speed-comparison | mutate | 1 | 0 |  |
| 2500 | speed-comparison | read | 3 | 8 | read-top-range-stream: Loss +297.4% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | write | 13 | 0 |  |
| 25000 | dense-helloworld-comparison | read | 0 | 2 | dense-helloworld-read-stream: Loss +81.8% vs Sylvan.Data.Excel |
| 25000 | package-profile | package | 10 | 2 | write-bulk-report: Loss +20.1% vs MiniExcel |
| 25000 | speed-comparison | mutate | 0 | 1 | autofit-existing: Loss +3.2% vs EPPlus |
| 25000 | speed-comparison | read | 1 | 10 | read-range: Loss +558.9% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | write | 12 | 1 | write-bulk-report: Loss +11.1% vs MiniExcel |

## OfficeIMO decision table

| Row count | Artifact | Workload | Scenario | OfficeIMO mean | Best | OfficeIMO vs best | Alloc | Package |
| ---: | --- | --- | --- | ---: | --- | ---: | ---: | ---: |
| 2500 | dense-helloworld-comparison | read | dense-helloworld-read-range | 8.19 ms | Sylvan.Data.Excel | Loss +40.3% | 2726.3 KB |  |
| 2500 | dense-helloworld-comparison | read | dense-helloworld-read-stream | 10.65 ms | Sylvan.Data.Excel | Loss +89.3% | 2804.8 KB |  |
| 2500 | package-profile | package | append-plain-rows | 3.97 ms | OfficeIMO.Excel | Win | 3721.6 KB | 64.5 KB |
| 2500 | package-profile | package | autofit-existing | 35.07 ms | OfficeIMO.Excel | Win | 13869.2 KB | 139.0 KB |
| 2500 | package-profile | package | large-shared-strings | 2.87 ms | OfficeIMO.Excel | Win | 4228.5 KB | 52.6 KB |
| 2500 | package-profile | package | write-bulk-report | 9.70 ms | OfficeIMO.Excel | Win | 5839.3 KB | 139.0 KB |
| 2500 | package-profile | package | write-cellvalues-rectangle-direct | 5.21 ms | OfficeIMO.Excel | Win | 3810.1 KB | 138.0 KB |
| 2500 | package-profile | package | write-datareader-table | 8.97 ms | OfficeIMO.Excel | Win | 5051.9 KB | 138.8 KB |
| 2500 | package-profile | package | write-dataset-tables | 5.61 ms | OfficeIMO.Excel | Win | 4676.0 KB | 139.0 KB |
| 2500 | package-profile | package | write-dataset-tables-autofit | 6.05 ms | OfficeIMO.Excel | Win | 4676.4 KB | 139.3 KB |
| 2500 | package-profile | package | write-datatable-direct | 4.99 ms | OfficeIMO.Excel | Win | 3477.0 KB | 138.0 KB |
| 2500 | package-profile | package | write-datatable-table-direct | 5.71 ms | OfficeIMO.Excel | Win | 3934.3 KB | 138.8 KB |
| 2500 | package-profile | package | write-fluent-rowsfrom-direct | 5.40 ms | OfficeIMO.Excel | Win | 3518.5 KB | 138.0 KB |
| 2500 | package-profile | package | write-insertobjects-direct | 4.48 ms | OfficeIMO.Excel | Win | 3497.7 KB | 138.0 KB |
| 2500 | speed-comparison | mutate | autofit-existing | 38.28 ms | OfficeIMO.Excel | Win | 13869.2 KB |  |
| 2500 | speed-comparison | read | formula-heavy-read | 4.95 ms | OfficeIMO.Excel | Win | 1704.8 KB |  |
| 2500 | speed-comparison | read | large-sparse-column-read | 3.16 ms | OfficeIMO.Excel | Win | 289.2 KB |  |
| 2500 | speed-comparison | read | large-sparse-row-read | 3.80 ms | Sylvan.Data.Excel | Loss +74.1% | 289.1 KB |  |
| 2500 | speed-comparison | read | read-datatable | 14.16 ms | Sylvan.Data.Excel | Loss +20.5% | 4196.4 KB |  |
| 2500 | speed-comparison | read | read-objects | 10.83 ms | Sylvan.Data.Excel | Loss +43.9% | 3155.2 KB |  |
| 2500 | speed-comparison | read | read-objects-stream | 10.57 ms | Sylvan.Data.Excel | Loss +83.4% | 1441.3 KB |  |
| 2500 | speed-comparison | read | read-range | 15.31 ms | OfficeIMO.Excel | Win | 2839.5 KB |  |
| 2500 | speed-comparison | read | read-range-stream | 12.62 ms | Sylvan.Data.Excel | Loss +105.4% | 3074.8 KB |  |
| 2500 | speed-comparison | read | read-top-range | 2.82 ms | Sylvan.Data.Excel | Loss +147.6% | 583.6 KB |  |
| 2500 | speed-comparison | read | read-top-range-stream | 2.50 ms | Sylvan.Data.Excel | Loss +297.4% | 587.3 KB |  |
| 2500 | speed-comparison | read | shared-string-read | 6.95 ms | Sylvan.Data.Excel | Loss +105.5% | 2861.4 KB |  |
| 2500 | speed-comparison | write | append-plain-rows | 4.95 ms | OfficeIMO.Excel | Win | 3721.6 KB |  |
| 2500 | speed-comparison | write | large-shared-strings | 3.61 ms | OfficeIMO.Excel | Win | 4228.5 KB |  |
| 2500 | speed-comparison | write | write-bulk-report | 14.90 ms | OfficeIMO.Excel | Win | 5839.8 KB |  |
| 2500 | speed-comparison | write | write-cellvalues-rectangle-direct | 7.50 ms | OfficeIMO.Excel | Win | 3810.1 KB |  |
| 2500 | speed-comparison | write | write-datareader-table | 9.59 ms | OfficeIMO.Excel | Win | 5051.9 KB |  |
| 2500 | speed-comparison | write | write-dataset-headerless-tables | 4.70 ms | OfficeIMO.Excel | Win | 4674.4 KB |  |
| 2500 | speed-comparison | write | write-dataset-sparse-tables | 8.04 ms | OfficeIMO.Excel | Win | 5648.3 KB |  |
| 2500 | speed-comparison | write | write-dataset-tables | 7.47 ms | OfficeIMO.Excel | Win | 4676.0 KB |  |
| 2500 | speed-comparison | write | write-dataset-tables-autofit | 7.06 ms | OfficeIMO.Excel | Win | 4676.4 KB |  |
| 2500 | speed-comparison | write | write-datatable-direct | 5.61 ms | OfficeIMO.Excel | Win | 3477.0 KB |  |
| 2500 | speed-comparison | write | write-datatable-table-direct | 5.26 ms | OfficeIMO.Excel | Win | 3934.3 KB |  |
| 2500 | speed-comparison | write | write-fluent-rowsfrom-direct | 6.44 ms | OfficeIMO.Excel | Win | 3518.5 KB |  |
| 2500 | speed-comparison | write | write-insertobjects-direct | 6.74 ms | OfficeIMO.Excel | Win | 3497.7 KB |  |
| 25000 | dense-helloworld-comparison | read | dense-helloworld-read-range | 62.70 ms | Sylvan.Data.Excel | Loss +30.6% | 25687.1 KB |  |
| 25000 | dense-helloworld-comparison | read | dense-helloworld-read-stream | 107.85 ms | Sylvan.Data.Excel | Loss +81.8% | 27316.6 KB |  |
| 25000 | package-profile | package | append-plain-rows | 27.34 ms | OfficeIMO.Excel | Win | 14434.5 KB | 622.6 KB |
| 25000 | package-profile | package | autofit-existing | 406.19 ms | OfficeIMO.Excel | Win | 136001.3 KB | 1385.9 KB |
| 25000 | package-profile | package | large-shared-strings | 22.13 ms | OfficeIMO.Excel | Win | 20468.2 KB | 520.3 KB |
| 25000 | package-profile | package | write-bulk-report | 98.89 ms | MiniExcel | Loss +20.1% | 33258.7 KB | 1385.9 KB |
| 25000 | package-profile | package | write-cellvalues-rectangle-direct | 49.45 ms | OfficeIMO.Excel | Win | 18475.6 KB | 1385.0 KB |
| 25000 | package-profile | package | write-datareader-table | 90.74 ms | MiniExcel | Loss +5.0% | 25791.7 KB | 1385.8 KB |
| 25000 | package-profile | package | write-dataset-tables | 42.54 ms | OfficeIMO.Excel | Win | 13599.9 KB | 1376.4 KB |
| 25000 | package-profile | package | write-dataset-tables-autofit | 47.34 ms | OfficeIMO.Excel | Win | 13600.3 KB | 1376.7 KB |
| 25000 | package-profile | package | write-datatable-direct | 44.92 ms | OfficeIMO.Excel | Win | 15154.1 KB | 1385.0 KB |
| 25000 | package-profile | package | write-datatable-table-direct | 49.60 ms | OfficeIMO.Excel | Win | 15617.5 KB | 1385.8 KB |
| 25000 | package-profile | package | write-fluent-rowsfrom-direct | 51.57 ms | OfficeIMO.Excel | Win | 15547.2 KB | 1385.0 KB |
| 25000 | package-profile | package | write-insertobjects-direct | 45.40 ms | OfficeIMO.Excel | Win | 15350.7 KB | 1385.0 KB |
| 25000 | speed-comparison | mutate | autofit-existing | 538.56 ms | EPPlus | Loss +3.2% | 136146.9 KB |  |
| 25000 | speed-comparison | read | formula-heavy-read | 39.09 ms | OfficeIMO.Excel | Win | 15474.4 KB |  |
| 25000 | speed-comparison | read | large-sparse-column-read | 1.64 ms | Sylvan.Data.Excel | Loss +65.5% | 289.0 KB |  |
| 25000 | speed-comparison | read | large-sparse-row-read | 1.72 ms | Sylvan.Data.Excel | Loss +85.7% | 289.0 KB |  |
| 25000 | speed-comparison | read | read-datatable | 318.28 ms | Sylvan.Data.Excel | Loss +498.2% | 218093.2 KB |  |
| 25000 | speed-comparison | read | read-objects | 350.08 ms | Sylvan.Data.Excel | Loss +449.3% | 210710.2 KB |  |
| 25000 | speed-comparison | read | read-objects-stream | 84.10 ms | Sylvan.Data.Excel | Loss +49.3% | 10347.7 KB |  |
| 25000 | speed-comparison | read | read-range | 272.88 ms | Sylvan.Data.Excel | Loss +558.9% | 209159.5 KB |  |
| 25000 | speed-comparison | read | read-range-stream | 78.52 ms | Sylvan.Data.Excel | Loss +82.8% | 27067.8 KB |  |
| 25000 | speed-comparison | read | read-top-range | 2.02 ms | Sylvan.Data.Excel | Loss +370.9% | 1111.0 KB |  |
| 25000 | speed-comparison | read | read-top-range-stream | 2.09 ms | Sylvan.Data.Excel | Loss +427.2% | 1114.2 KB |  |
| 25000 | speed-comparison | read | shared-string-read | 71.03 ms | Sylvan.Data.Excel | Loss +132.7% | 25985.7 KB |  |
| 25000 | speed-comparison | write | append-plain-rows | 20.10 ms | OfficeIMO.Excel | Win | 14442.5 KB |  |
| 25000 | speed-comparison | write | large-shared-strings | 22.12 ms | OfficeIMO.Excel | Win | 20476.3 KB |  |
| 25000 | speed-comparison | write | write-bulk-report | 95.32 ms | MiniExcel | Loss +11.1% | 33266.8 KB |  |
| 25000 | speed-comparison | write | write-cellvalues-rectangle-direct | 42.90 ms | OfficeIMO.Excel | Win | 18483.6 KB |  |
| 25000 | speed-comparison | write | write-datareader-table | 90.01 ms | OfficeIMO.Excel | Win | 25799.8 KB |  |
| 25000 | speed-comparison | write | write-dataset-headerless-tables | 49.32 ms | OfficeIMO.Excel | Win | 16367.2 KB |  |
| 25000 | speed-comparison | write | write-dataset-sparse-tables | 63.65 ms | OfficeIMO.Excel | Win | 22530.1 KB |  |
| 25000 | speed-comparison | write | write-dataset-tables | 43.40 ms | OfficeIMO.Excel | Win | 13607.9 KB |  |
| 25000 | speed-comparison | write | write-dataset-tables-autofit | 45.41 ms | OfficeIMO.Excel | Win | 13608.7 KB |  |
| 25000 | speed-comparison | write | write-datatable-direct | 46.54 ms | OfficeIMO.Excel | Win | 15162.2 KB |  |
| 25000 | speed-comparison | write | write-datatable-table-direct | 47.53 ms | OfficeIMO.Excel | Win | 15625.6 KB |  |
| 25000 | speed-comparison | write | write-fluent-rowsfrom-direct | 37.11 ms | OfficeIMO.Excel | Win | 15555.2 KB |  |
| 25000 | speed-comparison | write | write-insertobjects-direct | 34.40 ms | OfficeIMO.Excel | Win | 15358.7 KB |  |

## Full competitor table

| Row count | Artifact | Scenario | Library | Mean | StdDev | StdErr | Ratio to OfficeIMO | Ratio to best | Alloc | Alloc ratio | Package | Package ratio | Outcome |
| ---: | --- | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | Sylvan.Data.Excel | 5.83 ms | 1.21 ms | 0.70 ms | 0.71 | 1.00 | 357.8 KB | 0.13 |  |  | 28.7% faster than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | OfficeIMO.Excel | 8.19 ms | 1.53 ms | 0.88 ms | 1.00 | 1.40 | 2726.3 KB | 1.00 |  |  | Loss +40.3% |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-range | MiniExcel | 20.58 ms | 3.37 ms | 1.94 ms | 2.51 | 3.53 | 21507.3 KB | 7.89 |  |  | 151.4% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | Sylvan.Data.Excel | 5.63 ms | 1.03 ms | 0.59 ms | 0.53 | 1.00 | 357.8 KB | 0.13 |  |  | 47.2% faster than OfficeIMO |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | OfficeIMO.Excel | 10.65 ms | 2.35 ms | 1.36 ms | 1.00 | 1.89 | 2804.8 KB | 1.00 |  |  | Loss +89.3% |
| 2500 | dense-helloworld-comparison | dense-helloworld-read-stream | MiniExcel | 22.27 ms | 4.83 ms | 2.79 ms | 2.09 | 3.96 | 21507.3 KB | 7.67 |  |  | 109.1% slower than OfficeIMO |
| 2500 | package-profile | append-plain-rows | OfficeIMO.Excel | 3.97 ms | 0.33 ms | 0.19 ms | 1.00 | 1.00 | 3721.6 KB | 1.00 | 64.5 KB | 1.00 | Win |
| 2500 | package-profile | append-plain-rows | MiniExcel | 5.76 ms | 0.48 ms | 0.28 ms | 1.45 | 1.45 | 19710.7 KB | 5.30 | 68.1 KB | 1.06 | 44.9% slower than OfficeIMO |
| 2500 | package-profile | append-plain-rows | ClosedXML | 18.47 ms | 1.52 ms | 0.88 ms | 4.65 | 4.65 | 11197.4 KB | 3.01 | 59.8 KB | 0.93 | 365.1% slower than OfficeIMO |
| 2500 | package-profile | append-plain-rows | EPPlus | 21.47 ms | 1.68 ms | 0.97 ms | 5.41 | 5.41 | 14365.6 KB | 3.86 | 56.9 KB | 0.88 | 440.6% slower than OfficeIMO |
| 2500 | package-profile | autofit-existing | OfficeIMO.Excel | 35.07 ms | 1.65 ms | 0.95 ms | 1.00 | 1.00 | 13869.2 KB | 1.00 | 139.0 KB | 1.00 | Win |
| 2500 | package-profile | autofit-existing | EPPlus | 80.97 ms | 2.29 ms | 1.32 ms | 2.31 | 2.31 | 50711.7 KB | 3.66 | 115.0 KB | 0.83 | 130.9% slower than OfficeIMO |
| 2500 | package-profile | autofit-existing | ClosedXML | 159.91 ms | 7.39 ms | 4.27 ms | 4.56 | 4.56 | 84562.3 KB | 6.10 | 121.0 KB | 0.87 | 355.9% slower than OfficeIMO |
| 2500 | package-profile | large-shared-strings | OfficeIMO.Excel | 2.87 ms | 0.45 ms | 0.26 ms | 1.00 | 1.00 | 4228.5 KB | 1.00 | 52.6 KB | 1.00 | Win |
| 2500 | package-profile | large-shared-strings | MiniExcel | 5.62 ms | 0.17 ms | 0.10 ms | 1.96 | 1.96 | 21137.5 KB | 5.00 | 60.7 KB | 1.15 | 96.1% slower than OfficeIMO |
| 2500 | package-profile | large-shared-strings | ClosedXML | 15.88 ms | 1.13 ms | 0.65 ms | 5.54 | 5.54 | 11299.2 KB | 2.67 | 50.3 KB | 0.96 | 453.7% slower than OfficeIMO |
| 2500 | package-profile | large-shared-strings | EPPlus | 19.97 ms | 1.44 ms | 0.83 ms | 6.96 | 6.96 | 12804.9 KB | 3.03 | 48.1 KB | 0.91 | 596.2% slower than OfficeIMO |
| 2500 | package-profile | write-bulk-report | OfficeIMO.Excel | 9.70 ms | 1.05 ms | 0.60 ms | 1.00 | 1.00 | 5839.3 KB | 1.00 | 139.0 KB | 1.00 | Win |
| 2500 | package-profile | write-bulk-report | MiniExcel | 11.30 ms | 2.15 ms | 1.24 ms | 1.16 | 1.16 | 26825.4 KB | 4.59 | 153.8 KB | 1.11 | 16.5% slower than OfficeIMO |
| 2500 | package-profile | write-bulk-report | EPPlus | 67.12 ms | 10.01 ms | 5.78 ms | 6.92 | 6.92 | 47193.9 KB | 8.08 | 115.0 KB | 0.83 | 591.7% slower than OfficeIMO |
| 2500 | package-profile | write-bulk-report | ClosedXML | 105.30 ms | 9.37 ms | 5.41 ms | 10.85 | 10.85 | 58343.8 KB | 9.99 | 121.0 KB | 0.87 | 985.1% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 5.21 ms | 0.35 ms | 0.20 ms | 1.00 | 1.00 | 3810.1 KB | 1.00 | 138.0 KB | 1.00 | Win |
| 2500 | package-profile | write-cellvalues-rectangle-direct | MiniExcel | 9.93 ms | 1.58 ms | 0.91 ms | 1.91 | 1.91 | 23222.2 KB | 6.09 | 153.7 KB | 1.11 | 90.7% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | ClosedXML | 38.38 ms | 1.29 ms | 0.74 ms | 7.37 | 7.37 | 22221.3 KB | 5.83 | 120.1 KB | 0.87 | 636.7% slower than OfficeIMO |
| 2500 | package-profile | write-cellvalues-rectangle-direct | EPPlus | 42.75 ms | 1.29 ms | 0.75 ms | 8.21 | 8.21 | 24694.4 KB | 6.48 | 114.1 KB | 0.83 | 720.7% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table | OfficeIMO.Excel | 8.97 ms | 0.92 ms | 0.53 ms | 1.00 | 1.00 | 5051.9 KB | 1.00 | 138.8 KB | 1.00 | Win |
| 2500 | package-profile | write-datareader-table | MiniExcel | 9.31 ms | 1.86 ms | 1.07 ms | 1.04 | 1.04 | 23044.1 KB | 4.56 | 153.6 KB | 1.11 | 3.7% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table | EPPlus | 36.78 ms | 2.09 ms | 1.21 ms | 4.10 | 4.10 | 16646.6 KB | 3.30 | 114.9 KB | 0.83 | 310.0% slower than OfficeIMO |
| 2500 | package-profile | write-datareader-table | ClosedXML | 44.03 ms | 4.09 ms | 2.36 ms | 4.91 | 4.91 | 19008.5 KB | 3.76 | 120.9 KB | 0.87 | 390.9% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables | OfficeIMO.Excel | 5.61 ms | 0.67 ms | 0.38 ms | 1.00 | 1.00 | 4676.0 KB | 1.00 | 139.0 KB | 1.00 | Win |
| 2500 | package-profile | write-dataset-tables | MiniExcel | 10.37 ms | 0.42 ms | 0.24 ms | 1.85 | 1.85 | 28700.3 KB | 6.14 | 156.4 KB | 1.13 | 85.0% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables | EPPlus | 40.31 ms | 1.74 ms | 1.01 ms | 7.19 | 7.19 | 18701.2 KB | 4.00 | 116.6 KB | 0.84 | 619.0% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables | ClosedXML | 49.97 ms | 4.80 ms | 2.77 ms | 8.91 | 8.91 | 18876.2 KB | 4.04 | 123.4 KB | 0.89 | 791.2% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables-autofit | OfficeIMO.Excel | 6.05 ms | 0.43 ms | 0.25 ms | 1.00 | 1.00 | 4676.4 KB | 1.00 | 139.3 KB | 1.00 | Win |
| 2500 | package-profile | write-dataset-tables-autofit | MiniExcel | 10.37 ms | 1.20 ms | 0.70 ms | 1.71 | 1.71 | 31804.3 KB | 6.80 | 156.6 KB | 1.12 | 71.4% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables-autofit | EPPlus | 66.69 ms | 0.66 ms | 0.38 ms | 11.02 | 11.02 | 41456.3 KB | 8.86 | 116.9 KB | 0.84 | 1002.3% slower than OfficeIMO |
| 2500 | package-profile | write-dataset-tables-autofit | ClosedXML | 100.46 ms | 7.20 ms | 4.16 ms | 16.61 | 16.61 | 56707.4 KB | 12.13 | 123.7 KB | 0.89 | 1560.5% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | OfficeIMO.Excel | 4.99 ms | 0.75 ms | 0.43 ms | 1.00 | 1.00 | 3477.0 KB | 1.00 | 138.0 KB | 1.00 | Win |
| 2500 | package-profile | write-datatable-direct | MiniExcel | 9.56 ms | 1.30 ms | 0.75 ms | 1.92 | 1.92 | 23062.5 KB | 6.63 | 153.7 KB | 1.11 | 91.8% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | ClosedXML | 33.22 ms | 5.39 ms | 3.11 ms | 6.66 | 6.66 | 11581.0 KB | 3.33 | 120.1 KB | 0.87 | 566.1% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-direct | EPPlus | 35.71 ms | 3.01 ms | 1.74 ms | 7.16 | 7.16 | 16646.6 KB | 4.79 | 114.9 KB | 0.83 | 616.0% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-table-direct | OfficeIMO.Excel | 5.71 ms | 0.27 ms | 0.16 ms | 1.00 | 1.00 | 3934.3 KB | 1.00 | 138.8 KB | 1.00 | Win |
| 2500 | package-profile | write-datatable-table-direct | MiniExcel | 9.70 ms | 1.42 ms | 0.82 ms | 1.70 | 1.70 | 23062.8 KB | 5.86 | 153.7 KB | 1.11 | 69.7% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-table-direct | EPPlus | 36.54 ms | 2.27 ms | 1.31 ms | 6.39 | 6.39 | 16646.6 KB | 4.23 | 114.9 KB | 0.83 | 539.4% slower than OfficeIMO |
| 2500 | package-profile | write-datatable-table-direct | ClosedXML | 42.08 ms | 4.55 ms | 2.63 ms | 7.36 | 7.36 | 19007.8 KB | 4.83 | 120.9 KB | 0.87 | 636.4% slower than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 5.40 ms | 0.85 ms | 0.49 ms | 1.00 | 1.00 | 3518.5 KB | 1.00 | 138.0 KB | 1.00 | Win |
| 2500 | package-profile | write-fluent-rowsfrom-direct | MiniExcel | 8.66 ms | 0.99 ms | 0.57 ms | 1.60 | 1.60 | 23222.2 KB | 6.60 | 153.7 KB | 1.11 | 60.4% slower than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | ClosedXML | 33.46 ms | 3.68 ms | 2.12 ms | 6.20 | 6.20 | 11581.0 KB | 3.29 | 120.1 KB | 0.87 | 520.0% slower than OfficeIMO |
| 2500 | package-profile | write-fluent-rowsfrom-direct | EPPlus | 35.41 ms | 2.72 ms | 1.57 ms | 6.56 | 6.56 | 16646.6 KB | 4.73 | 114.9 KB | 0.83 | 556.3% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | OfficeIMO.Excel | 4.48 ms | 0.43 ms | 0.25 ms | 1.00 | 1.00 | 3497.7 KB | 1.00 | 138.0 KB | 1.00 | Win |
| 2500 | package-profile | write-insertobjects-direct | MiniExcel | 10.60 ms | 0.46 ms | 0.26 ms | 2.37 | 2.37 | 23222.2 KB | 6.64 | 153.7 KB | 1.11 | 136.6% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | EPPlus | 35.97 ms | 3.57 ms | 2.06 ms | 8.03 | 8.03 | 16646.6 KB | 4.76 | 114.9 KB | 0.83 | 702.6% slower than OfficeIMO |
| 2500 | package-profile | write-insertobjects-direct | ClosedXML | 36.98 ms | 4.56 ms | 2.63 ms | 8.25 | 8.25 | 11581.0 KB | 3.31 | 120.1 KB | 0.87 | 725.1% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | OfficeIMO.Excel | 4.95 ms | 0.26 ms | 0.15 ms | 1.00 | 1.00 | 3721.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | append-plain-rows | MiniExcel | 5.59 ms | 1.31 ms | 0.76 ms | 1.13 | 1.13 | 19710.6 KB | 5.30 |  |  | 13.1% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | EPPlus 4.5.3.3 | 15.79 ms |  |  | 3.19 | 3.19 |  |  |  |  | 219.3% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | ClosedXML | 19.72 ms | 1.46 ms | 0.84 ms | 3.99 | 3.99 | 11197.4 KB | 3.01 |  |  | 298.8% slower than OfficeIMO |
| 2500 | speed-comparison | append-plain-rows | EPPlus | 21.63 ms | 0.44 ms | 0.25 ms | 4.37 | 4.37 | 14365.7 KB | 3.86 |  |  | 337.3% slower than OfficeIMO |
| 2500 | speed-comparison | autofit-existing | OfficeIMO.Excel | 38.28 ms | 1.32 ms | 0.76 ms | 1.00 | 1.00 | 13869.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | autofit-existing | EPPlus | 84.32 ms | 5.76 ms | 3.33 ms | 2.20 | 2.20 | 50711.7 KB | 3.66 |  |  | 120.2% slower than OfficeIMO |
| 2500 | speed-comparison | autofit-existing | EPPlus 4.5.3.3 | 140.94 ms |  |  | 3.68 | 3.68 |  |  |  |  | 268.1% slower than OfficeIMO |
| 2500 | speed-comparison | autofit-existing | ClosedXML | 172.84 ms | 8.25 ms | 4.77 ms | 4.51 | 4.51 | 84677.9 KB | 6.11 |  |  | 351.5% slower than OfficeIMO |
| 2500 | speed-comparison | formula-heavy-read | OfficeIMO.Excel | 4.95 ms | 1.33 ms | 0.77 ms | 1.00 | 1.00 | 1704.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | formula-heavy-read | EPPlus | 16.37 ms | 2.62 ms | 1.51 ms | 3.30 | 3.30 | 8281.8 KB | 4.86 |  |  | 230.4% slower than OfficeIMO |
| 2500 | speed-comparison | formula-heavy-read | EPPlus 4.5.3.3 | 20.61 ms |  |  | 4.16 | 4.16 |  |  |  |  | 316.0% slower than OfficeIMO |
| 2500 | speed-comparison | formula-heavy-read | ClosedXML | 22.40 ms | 1.06 ms | 0.61 ms | 4.52 | 4.52 | 9277.7 KB | 5.44 |  |  | 352.0% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | OfficeIMO.Excel | 3.61 ms | 0.31 ms | 0.18 ms | 1.00 | 1.00 | 4228.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | large-shared-strings | MiniExcel | 4.99 ms | 0.40 ms | 0.23 ms | 1.39 | 1.39 | 21137.5 KB | 5.00 |  |  | 38.5% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | ClosedXML | 15.49 ms | 1.17 ms | 0.67 ms | 4.30 | 4.30 | 11299.2 KB | 2.67 |  |  | 329.6% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | EPPlus | 21.40 ms | 3.77 ms | 2.18 ms | 5.94 | 5.94 | 12804.9 KB | 3.03 |  |  | 493.6% slower than OfficeIMO |
| 2500 | speed-comparison | large-shared-strings | EPPlus 4.5.3.3 | 23.83 ms |  |  | 6.61 | 6.61 |  |  |  |  | 560.8% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | OfficeIMO.Excel | 3.16 ms | 0.87 ms | 0.51 ms | 1.00 | 1.00 | 289.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | large-sparse-column-read | Sylvan.Data.Excel | 3.48 ms | 1.68 ms | 0.97 ms | 1.10 | 1.10 | 316.5 KB | 1.09 |  |  | 10.2% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | ClosedXML | 3.92 ms | 0.74 ms | 0.43 ms | 1.24 | 1.24 | 4392.5 KB | 15.19 |  |  | 24.2% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | EPPlus 4.5.3.3 | 9.64 ms |  |  | 3.05 | 3.05 |  |  |  |  | 204.9% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | MiniExcel | 18.75 ms | 1.38 ms | 0.80 ms | 5.93 | 5.93 | 46194.9 KB | 159.73 |  |  | 493.2% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-column-read | EPPlus | 33.11 ms | 3.01 ms | 1.74 ms | 10.47 | 10.47 | 43070.9 KB | 148.92 |  |  | 947.4% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | Sylvan.Data.Excel | 2.18 ms | 0.53 ms | 0.31 ms | 0.57 | 1.00 | 316.5 KB | 1.09 |  |  | 42.6% faster than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | OfficeIMO.Excel | 3.80 ms | 0.44 ms | 0.26 ms | 1.00 | 1.74 | 289.1 KB | 1.00 |  |  | Loss +74.1% |
| 2500 | speed-comparison | large-sparse-row-read | ClosedXML | 4.01 ms | 0.70 ms | 0.41 ms | 1.06 | 1.84 | 4392.4 KB | 15.19 |  |  | 5.6% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | EPPlus 4.5.3.3 | 9.42 ms |  |  | 2.48 | 4.32 |  |  |  |  | 148.1% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | MiniExcel | 19.45 ms | 3.29 ms | 1.90 ms | 5.12 | 8.92 | 46194.9 KB | 159.79 |  |  | 412.3% slower than OfficeIMO |
| 2500 | speed-comparison | large-sparse-row-read | EPPlus | 32.86 ms | 2.95 ms | 1.70 ms | 8.65 | 15.07 | 43070.9 KB | 148.98 |  |  | 765.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | Sylvan.Data.Excel | 11.75 ms | 7.51 ms | 4.34 ms | 0.83 | 1.00 | 1952.8 KB | 0.47 |  |  | 17.0% faster than OfficeIMO |
| 2500 | speed-comparison | read-datatable | OfficeIMO.Excel | 14.16 ms | 0.16 ms | 0.09 ms | 1.00 | 1.21 | 4196.4 KB | 1.00 |  |  | Loss +20.5% |
| 2500 | speed-comparison | read-datatable | MiniExcel | 23.23 ms | 8.79 ms | 5.07 ms | 1.64 | 1.98 | 18637.8 KB | 4.44 |  |  | 64.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | EPPlus 4.5.3.3 | 36.94 ms |  |  | 2.61 | 3.14 |  |  |  |  | 160.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | EPPlus | 39.77 ms | 4.60 ms | 2.66 ms | 2.81 | 3.38 | 20509.6 KB | 4.89 |  |  | 180.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-datatable | ClosedXML | 117.49 ms | 5.35 ms | 3.09 ms | 8.29 | 10.00 | 22236.9 KB | 5.30 |  |  | 729.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | Sylvan.Data.Excel | 7.53 ms | 1.68 ms | 0.97 ms | 0.69 | 1.00 | 544.4 KB | 0.17 |  |  | 30.5% faster than OfficeIMO |
| 2500 | speed-comparison | read-objects | OfficeIMO.Excel | 10.83 ms | 1.05 ms | 0.61 ms | 1.00 | 1.44 | 3155.2 KB | 1.00 |  |  | Loss +43.9% |
| 2500 | speed-comparison | read-objects | MiniExcel | 24.37 ms | 4.73 ms | 2.73 ms | 2.25 | 3.24 | 18781.0 KB | 5.95 |  |  | 125.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | EPPlus | 25.83 ms | 1.97 ms | 1.14 ms | 2.38 | 3.43 | 20107.5 KB | 6.37 |  |  | 138.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | EPPlus 4.5.3.3 | 34.14 ms |  |  | 3.15 | 4.54 |  |  |  |  | 215.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects | ClosedXML | 51.27 ms | 7.15 ms | 4.13 ms | 4.73 | 6.81 | 20713.1 KB | 6.56 |  |  | 373.4% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | Sylvan.Data.Excel | 5.77 ms | 0.67 ms | 0.39 ms | 0.55 | 1.00 | 544.5 KB | 0.38 |  |  | 45.5% faster than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | OfficeIMO.Excel | 10.57 ms | 1.38 ms | 0.80 ms | 1.00 | 1.83 | 1441.3 KB | 1.00 |  |  | Loss +83.4% |
| 2500 | speed-comparison | read-objects-stream | MiniExcel | 22.51 ms | 4.70 ms | 2.71 ms | 2.13 | 3.90 | 18781.0 KB | 13.03 |  |  | 112.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | EPPlus | 29.27 ms | 5.16 ms | 2.98 ms | 2.77 | 5.08 | 20107.4 KB | 13.95 |  |  | 176.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | EPPlus 4.5.3.3 | 33.74 ms |  |  | 3.19 | 5.85 |  |  |  |  | 219.1% slower than OfficeIMO |
| 2500 | speed-comparison | read-objects-stream | ClosedXML | 46.12 ms | 2.47 ms | 1.42 ms | 4.36 | 8.00 | 20711.5 KB | 14.37 |  |  | 336.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | OfficeIMO.Excel | 15.31 ms | 1.37 ms | 0.79 ms | 1.00 | 1.00 | 2839.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read-range | Sylvan.Data.Excel | 18.51 ms | 0.63 ms | 0.36 ms | 1.21 | 1.21 | 368.5 KB | 0.13 |  |  | 20.9% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | MiniExcel | 30.85 ms | 2.27 ms | 1.31 ms | 2.02 | 2.02 | 19033.3 KB | 6.70 |  |  | 101.6% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | EPPlus | 32.76 ms | 3.81 ms | 2.20 ms | 2.14 | 2.14 | 18925.5 KB | 6.67 |  |  | 114.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | EPPlus 4.5.3.3 | 33.80 ms |  |  | 2.21 | 2.21 |  |  |  |  | 120.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-range | ClosedXML | 91.23 ms | 4.07 ms | 2.35 ms | 5.96 | 5.96 | 20654.1 KB | 7.27 |  |  | 496.0% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | Sylvan.Data.Excel | 6.14 ms | 1.25 ms | 0.72 ms | 0.49 | 1.00 | 368.7 KB | 0.12 |  |  | 51.3% faster than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | OfficeIMO.Excel | 12.62 ms | 1.41 ms | 0.81 ms | 1.00 | 2.05 | 3074.8 KB | 1.00 |  |  | Loss +105.4% |
| 2500 | speed-comparison | read-range-stream | MiniExcel | 20.54 ms | 2.41 ms | 1.39 ms | 1.63 | 3.34 | 19033.6 KB | 6.19 |  |  | 62.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | EPPlus | 26.94 ms | 1.87 ms | 1.08 ms | 2.14 | 4.39 | 18925.6 KB | 6.16 |  |  | 113.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | EPPlus 4.5.3.3 | 31.90 ms |  |  | 2.53 | 5.19 |  |  |  |  | 152.8% slower than OfficeIMO |
| 2500 | speed-comparison | read-range-stream | ClosedXML | 80.90 ms | 32.97 ms | 19.04 ms | 6.41 | 13.17 | 20615.8 KB | 6.70 |  |  | 541.2% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | Sylvan.Data.Excel | 1.14 ms | 0.26 ms | 0.15 ms | 0.40 | 1.00 | 365.3 KB | 0.63 |  |  | 59.6% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range | MiniExcel | 1.99 ms | 0.43 ms | 0.25 ms | 0.71 | 1.75 | 901.9 KB | 1.55 |  |  | 29.2% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range | OfficeIMO.Excel | 2.82 ms | 0.82 ms | 0.47 ms | 1.00 | 2.48 | 583.6 KB | 1.00 |  |  | Loss +147.6% |
| 2500 | speed-comparison | read-top-range | EPPlus 4.5.3.3 | 27.71 ms |  |  | 9.84 | 24.38 |  |  |  |  | 884.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | EPPlus | 31.20 ms | 11.72 ms | 6.76 ms | 11.08 | 27.45 | 11174.9 KB | 19.15 |  |  | 1008.3% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range | ClosedXML | 101.50 ms | 33.27 ms | 19.21 ms | 36.06 | 89.29 | 19285.1 KB | 33.05 |  |  | 3505.7% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | Sylvan.Data.Excel | 0.63 ms | 0.03 ms | 0.02 ms | 0.25 | 1.00 | 365.5 KB | 0.62 |  |  | 74.8% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | MiniExcel | 1.32 ms | 0.06 ms | 0.04 ms | 0.53 | 2.10 | 902.1 KB | 1.54 |  |  | 47.2% faster than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | OfficeIMO.Excel | 2.50 ms | 0.38 ms | 0.22 ms | 1.00 | 3.97 | 587.3 KB | 1.00 |  |  | Loss +297.4% |
| 2500 | speed-comparison | read-top-range-stream | EPPlus | 25.14 ms | 1.92 ms | 1.11 ms | 10.05 | 39.92 | 11175.1 KB | 19.03 |  |  | 904.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | EPPlus 4.5.3.3 | 29.54 ms |  |  | 11.80 | 46.91 |  |  |  |  | 1080.5% slower than OfficeIMO |
| 2500 | speed-comparison | read-top-range-stream | ClosedXML | 50.58 ms | 5.78 ms | 3.34 ms | 20.21 | 80.32 | 19167.1 KB | 32.64 |  |  | 1921.3% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | Sylvan.Data.Excel | 3.38 ms | 0.30 ms | 0.18 ms | 0.49 | 1.00 | 578.8 KB | 0.20 |  |  | 51.3% faster than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | OfficeIMO.Excel | 6.95 ms | 1.21 ms | 0.70 ms | 1.00 | 2.06 | 2861.4 KB | 1.00 |  |  | Loss +105.5% |
| 2500 | speed-comparison | shared-string-read | MiniExcel | 7.43 ms | 0.38 ms | 0.22 ms | 1.07 | 2.20 | 9708.7 KB | 3.39 |  |  | 6.9% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | EPPlus | 14.87 ms | 1.06 ms | 0.61 ms | 2.14 | 4.40 | 11239.7 KB | 3.93 |  |  | 114.0% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | ClosedXML | 22.18 ms | 2.00 ms | 1.15 ms | 3.19 | 6.56 | 11658.5 KB | 4.07 |  |  | 219.1% slower than OfficeIMO |
| 2500 | speed-comparison | shared-string-read | EPPlus 4.5.3.3 | 26.11 ms |  |  | 3.76 | 7.72 |  |  |  |  | 275.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | OfficeIMO.Excel | 14.90 ms | 3.20 ms | 1.85 ms | 1.00 | 1.00 | 5839.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write-bulk-report | MiniExcel | 16.98 ms | 2.02 ms | 1.17 ms | 1.14 | 1.14 | 26825.0 KB | 4.59 |  |  | 14.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | EPPlus 4.5.3.3 | 108.37 ms |  |  | 7.27 | 7.27 |  |  |  |  | 627.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | EPPlus | 140.98 ms | 19.90 ms | 11.49 ms | 9.46 | 9.46 | 49157.1 KB | 8.42 |  |  | 846.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-bulk-report | ClosedXML | 333.61 ms | 95.64 ms | 55.22 ms | 22.40 | 22.40 | 58345.7 KB | 9.99 |  |  | 2139.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 7.50 ms | 3.02 ms | 1.75 ms | 1.00 | 1.00 | 3810.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | MiniExcel | 13.34 ms | 0.60 ms | 0.35 ms | 1.78 | 1.78 | 23221.9 KB | 6.09 |  |  | 77.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 32.58 ms |  |  | 4.34 | 4.34 |  |  |  |  | 334.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | ClosedXML | 44.58 ms | 2.87 ms | 1.66 ms | 5.94 | 5.94 | 22221.3 KB | 5.83 |  |  | 494.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus | 51.54 ms | 15.60 ms | 9.01 ms | 6.87 | 6.87 | 24695.5 KB | 6.48 |  |  | 586.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | OfficeIMO.Excel | 9.59 ms | 0.82 ms | 0.47 ms | 1.00 | 1.00 | 5051.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write-datareader-table | MiniExcel | 11.65 ms | 0.41 ms | 0.24 ms | 1.22 | 1.22 | 23044.1 KB | 4.56 |  |  | 21.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | EPPlus | 39.47 ms | 2.91 ms | 1.68 ms | 4.12 | 4.12 | 16647.5 KB | 3.30 |  |  | 311.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table | ClosedXML | 45.39 ms | 3.81 ms | 2.20 ms | 4.74 | 4.74 | 19008.4 KB | 3.76 |  |  | 373.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datareader-table-direct | EPPlus 4.5.3.3 | 36.89 ms |  |  |  | 1.00 |  |  |  |  |  |
| 2500 | speed-comparison | write-dataset-headerless-tables | OfficeIMO.Excel | 4.70 ms | 0.42 ms | 0.24 ms | 1.00 | 1.00 | 4674.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write-dataset-headerless-tables | MiniExcel | 16.83 ms | 2.06 ms | 1.19 ms | 3.58 | 3.58 | 28694.9 KB | 6.14 |  |  | 258.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-headerless-tables | ClosedXML | 52.53 ms | 1.88 ms | 1.08 ms | 11.18 | 11.18 | 18914.5 KB | 4.05 |  |  | 1017.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-headerless-tables | EPPlus | 70.94 ms | 6.17 ms | 3.56 ms | 15.09 | 15.09 | 18406.8 KB | 3.94 |  |  | 1409.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-tables | OfficeIMO.Excel | 8.04 ms | 2.11 ms | 1.22 ms | 1.00 | 1.00 | 5648.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write-dataset-sparse-tables | MiniExcel | 23.07 ms | 4.83 ms | 2.79 ms | 2.87 | 2.87 | 29746.9 KB | 5.27 |  |  | 187.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-tables | EPPlus | 62.41 ms | 3.01 ms | 1.74 ms | 7.77 | 7.77 | 21891.6 KB | 3.88 |  |  | 676.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-sparse-tables | ClosedXML | 75.15 ms | 3.23 ms | 1.86 ms | 9.35 | 9.35 | 27410.8 KB | 4.85 |  |  | 835.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | OfficeIMO.Excel | 7.47 ms | 1.27 ms | 0.73 ms | 1.00 | 1.00 | 4676.0 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write-dataset-tables | MiniExcel | 17.95 ms | 3.30 ms | 1.90 ms | 2.40 | 2.40 | 28700.3 KB | 6.14 |  |  | 140.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | EPPlus 4.5.3.3 | 40.44 ms |  |  | 5.41 | 5.41 |  |  |  |  | 441.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | EPPlus | 86.23 ms | 18.18 ms | 10.49 ms | 11.54 | 11.54 | 19428.9 KB | 4.16 |  |  | 1053.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables | ClosedXML | 189.98 ms | 12.45 ms | 7.19 ms | 25.42 | 25.42 | 18876.7 KB | 4.04 |  |  | 2441.6% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables-autofit | OfficeIMO.Excel | 7.06 ms | 0.86 ms | 0.50 ms | 1.00 | 1.00 | 4676.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write-dataset-tables-autofit | MiniExcel | 19.10 ms | 1.76 ms | 1.02 ms | 2.70 | 2.70 | 31804.3 KB | 6.80 |  |  | 170.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables-autofit | EPPlus | 153.87 ms | 2.22 ms | 1.28 ms | 21.79 | 21.79 | 43438.3 KB | 9.29 |  |  | 2078.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-dataset-tables-autofit | ClosedXML | 271.37 ms | 157.52 ms | 90.94 ms | 38.42 | 38.42 | 56706.6 KB | 12.13 |  |  | 3742.4% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | OfficeIMO.Excel | 5.61 ms | 0.49 ms | 0.28 ms | 1.00 | 1.00 | 3477.0 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write-datatable-direct | MiniExcel | 13.17 ms | 1.89 ms | 1.09 ms | 2.35 | 2.35 | 23062.6 KB | 6.63 |  |  | 134.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | ClosedXML | 37.46 ms | 2.01 ms | 1.16 ms | 6.68 | 6.68 | 11581.0 KB | 3.33 |  |  | 567.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | EPPlus | 40.63 ms | 3.67 ms | 2.12 ms | 7.24 | 7.24 | 16647.7 KB | 4.79 |  |  | 624.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-direct | EPPlus 4.5.3.3 | 41.55 ms |  |  | 7.41 | 7.41 |  |  |  |  | 640.9% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | OfficeIMO.Excel | 5.26 ms | 0.76 ms | 0.44 ms | 1.00 | 1.00 | 3934.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write-datatable-table-direct | MiniExcel | 12.09 ms | 2.26 ms | 1.30 ms | 2.30 | 2.30 | 23062.8 KB | 5.86 |  |  | 129.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | EPPlus | 40.02 ms | 4.26 ms | 2.46 ms | 7.61 | 7.61 | 16647.5 KB | 4.23 |  |  | 660.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | EPPlus 4.5.3.3 | 42.18 ms |  |  | 8.02 | 8.02 |  |  |  |  | 701.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-datatable-table-direct | ClosedXML | 47.84 ms | 4.93 ms | 2.85 ms | 9.09 | 9.09 | 19007.7 KB | 4.83 |  |  | 809.3% slower than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 6.44 ms | 1.13 ms | 0.65 ms | 1.00 | 1.00 | 3518.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | MiniExcel | 11.46 ms | 1.76 ms | 1.02 ms | 1.78 | 1.78 | 23221.8 KB | 6.60 |  |  | 77.8% slower than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 29.50 ms |  |  | 4.58 | 4.58 |  |  |  |  | 357.7% slower than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | ClosedXML | 35.38 ms | 0.85 ms | 0.49 ms | 5.49 | 5.49 | 11581.0 KB | 3.29 |  |  | 449.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus | 39.19 ms | 3.04 ms | 1.75 ms | 6.08 | 6.08 | 16646.7 KB | 4.73 |  |  | 508.1% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | OfficeIMO.Excel | 6.74 ms | 2.19 ms | 1.26 ms | 1.00 | 1.00 | 3497.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write-insertobjects-direct | MiniExcel | 13.17 ms | 0.25 ms | 0.15 ms | 1.95 | 1.95 | 23221.8 KB | 6.64 |  |  | 95.5% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | EPPlus 4.5.3.3 | 29.59 ms |  |  | 4.39 | 4.39 |  |  |  |  | 339.2% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | EPPlus | 38.47 ms | 0.30 ms | 0.18 ms | 5.71 | 5.71 | 16647.4 KB | 4.76 |  |  | 471.0% slower than OfficeIMO |
| 2500 | speed-comparison | write-insertobjects-direct | ClosedXML | 42.38 ms | 9.77 ms | 5.64 ms | 6.29 | 6.29 | 11581.0 KB | 3.31 |  |  | 529.0% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | Sylvan.Data.Excel | 48.00 ms | 2.64 ms | 1.53 ms | 0.77 | 1.00 | 376.2 KB | 0.01 |  |  | 23.4% faster than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | OfficeIMO.Excel | 62.70 ms | 4.05 ms | 2.34 ms | 1.00 | 1.31 | 25687.1 KB | 1.00 |  |  | Loss +30.6% |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-range | MiniExcel | 181.01 ms | 8.22 ms | 4.75 ms | 2.89 | 3.77 | 215348.8 KB | 8.38 |  |  | 188.7% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | Sylvan.Data.Excel | 59.33 ms | 5.86 ms | 3.38 ms | 0.55 | 1.00 | 376.2 KB | 0.01 |  |  | 45.0% faster than OfficeIMO |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | OfficeIMO.Excel | 107.85 ms | 13.69 ms | 7.90 ms | 1.00 | 1.82 | 27316.6 KB | 1.00 |  |  | Loss +81.8% |
| 25000 | dense-helloworld-comparison | dense-helloworld-read-stream | MiniExcel | 201.97 ms | 3.97 ms | 2.29 ms | 1.87 | 3.40 | 215348.8 KB | 7.88 |  |  | 87.3% slower than OfficeIMO |
| 25000 | package-profile | append-plain-rows | OfficeIMO.Excel | 27.34 ms | 3.68 ms | 2.13 ms | 1.00 | 1.00 | 14434.5 KB | 1.00 | 622.6 KB | 1.00 | Win |
| 25000 | package-profile | append-plain-rows | MiniExcel | 37.18 ms | 2.61 ms | 1.50 ms | 1.36 | 1.36 | 58232.9 KB | 4.03 | 642.3 KB | 1.03 | 36.0% slower than OfficeIMO |
| 25000 | package-profile | append-plain-rows | EPPlus | 159.43 ms | 5.81 ms | 3.36 ms | 5.83 | 5.83 | 100275.6 KB | 6.95 | 525.6 KB | 0.84 | 483.1% slower than OfficeIMO |
| 25000 | package-profile | append-plain-rows | ClosedXML | 161.98 ms | 12.45 ms | 7.19 ms | 5.92 | 5.92 | 104225.1 KB | 7.22 | 540.6 KB | 0.87 | 492.5% slower than OfficeIMO |
| 25000 | package-profile | autofit-existing | OfficeIMO.Excel | 406.19 ms | 6.65 ms | 3.84 ms | 1.00 | 1.00 | 136001.3 KB | 1.00 | 1385.9 KB | 1.00 | Win |
| 25000 | package-profile | autofit-existing | EPPlus | 478.09 ms | 10.61 ms | 6.12 ms | 1.18 | 1.18 | 250886.0 KB | 1.84 | 1091.0 KB | 0.79 | 17.7% slower than OfficeIMO |
| 25000 | package-profile | autofit-existing | ClosedXML | 1639.02 ms | 41.28 ms | 23.84 ms | 4.04 | 4.04 | 829580.7 KB | 6.10 | 1140.9 KB | 0.82 | 303.5% slower than OfficeIMO |
| 25000 | package-profile | large-shared-strings | OfficeIMO.Excel | 22.13 ms | 1.83 ms | 1.06 ms | 1.00 | 1.00 | 20468.2 KB | 1.00 | 520.3 KB | 1.00 | Win |
| 25000 | package-profile | large-shared-strings | MiniExcel | 38.35 ms | 5.57 ms | 3.21 ms | 1.73 | 1.73 | 73751.2 KB | 3.60 | 581.0 KB | 1.12 | 73.3% slower than OfficeIMO |
| 25000 | package-profile | large-shared-strings | EPPlus | 130.40 ms | 11.48 ms | 6.63 ms | 5.89 | 5.89 | 84343.9 KB | 4.12 | 444.7 KB | 0.85 | 489.4% slower than OfficeIMO |
| 25000 | package-profile | large-shared-strings | ClosedXML | 134.38 ms | 15.72 ms | 9.08 ms | 6.07 | 6.07 | 104233.3 KB | 5.09 | 460.1 KB | 0.88 | 507.4% slower than OfficeIMO |
| 25000 | package-profile | write-bulk-report | MiniExcel | 82.33 ms | 5.39 ms | 3.11 ms | 0.83 | 1.00 | 125546.9 KB | 3.77 | 1521.1 KB | 1.10 | 16.8% faster than OfficeIMO |
| 25000 | package-profile | write-bulk-report | OfficeIMO.Excel | 98.89 ms | 10.60 ms | 6.12 ms | 1.00 | 1.20 | 33258.7 KB | 1.00 | 1385.9 KB | 1.00 | Loss +20.1% |
| 25000 | package-profile | write-bulk-report | EPPlus | 388.91 ms | 9.04 ms | 5.22 ms | 3.93 | 4.72 | 254895.7 KB | 7.66 | 1091.0 KB | 0.79 | 293.3% slower than OfficeIMO |
| 25000 | package-profile | write-bulk-report | ClosedXML | 973.26 ms | 27.94 ms | 16.13 ms | 9.84 | 11.82 | 565946.7 KB | 17.02 | 1140.9 KB | 0.82 | 884.2% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 49.45 ms | 3.31 ms | 1.91 ms | 1.00 | 1.00 | 18475.6 KB | 1.00 | 1385.0 KB | 1.00 | Win |
| 25000 | package-profile | write-cellvalues-rectangle-direct | MiniExcel | 81.80 ms | 9.05 ms | 5.22 ms | 1.65 | 1.65 | 93246.9 KB | 5.05 | 1521.0 KB | 1.10 | 65.4% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | EPPlus | 355.05 ms | 11.61 ms | 6.70 ms | 7.18 | 7.18 | 211791.2 KB | 11.46 | 1090.0 KB | 0.79 | 618.0% slower than OfficeIMO |
| 25000 | package-profile | write-cellvalues-rectangle-direct | ClosedXML | 380.73 ms | 5.03 ms | 2.90 ms | 7.70 | 7.70 | 210638.1 KB | 11.40 | 1139.9 KB | 0.82 | 669.9% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table | MiniExcel | 86.42 ms | 10.32 ms | 5.96 ms | 0.95 | 1.00 | 92190.0 KB | 3.57 | 1521.0 KB | 1.10 | 4.8% faster than OfficeIMO |
| 25000 | package-profile | write-datareader-table | OfficeIMO.Excel | 90.74 ms | 3.49 ms | 2.01 ms | 1.00 | 1.05 | 25791.7 KB | 1.00 | 1385.8 KB | 1.00 | Loss +5.0% |
| 25000 | package-profile | write-datareader-table | EPPlus | 320.77 ms | 7.90 ms | 4.56 ms | 3.54 | 3.71 | 117378.4 KB | 4.55 | 1090.8 KB | 0.79 | 253.5% slower than OfficeIMO |
| 25000 | package-profile | write-datareader-table | ClosedXML | 501.10 ms | 17.28 ms | 9.97 ms | 5.52 | 5.80 | 173385.9 KB | 6.72 | 1140.7 KB | 0.82 | 452.3% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables | OfficeIMO.Excel | 42.54 ms | 0.92 ms | 0.53 ms | 1.00 | 1.00 | 13599.9 KB | 1.00 | 1376.4 KB | 1.00 | Win |
| 25000 | package-profile | write-dataset-tables | MiniExcel | 89.33 ms | 3.52 ms | 2.03 ms | 2.10 | 2.10 | 97074.9 KB | 7.14 | 1511.8 KB | 1.10 | 110.0% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables | EPPlus | 309.52 ms | 7.25 ms | 4.19 ms | 7.28 | 7.28 | 110708.9 KB | 8.14 | 1100.6 KB | 0.80 | 627.7% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables | ClosedXML | 460.61 ms | 5.25 ms | 3.03 ms | 10.83 | 10.83 | 171992.0 KB | 12.65 | 1139.0 KB | 0.83 | 982.8% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables-autofit | OfficeIMO.Excel | 47.34 ms | 3.01 ms | 1.74 ms | 1.00 | 1.00 | 13600.3 KB | 1.00 | 1376.7 KB | 1.00 | Win |
| 25000 | package-profile | write-dataset-tables-autofit | MiniExcel | 107.16 ms | 14.57 ms | 8.41 ms | 2.26 | 2.26 | 128865.5 KB | 9.48 | 1512.0 KB | 1.10 | 126.4% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables-autofit | EPPlus | 399.57 ms | 32.78 ms | 18.92 ms | 8.44 | 8.44 | 195298.1 KB | 14.36 | 1100.9 KB | 0.80 | 744.0% slower than OfficeIMO |
| 25000 | package-profile | write-dataset-tables-autofit | ClosedXML | 929.55 ms | 23.67 ms | 13.66 ms | 19.64 | 19.64 | 550084.8 KB | 40.45 | 1139.3 KB | 0.83 | 1863.5% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | OfficeIMO.Excel | 44.92 ms | 1.97 ms | 1.14 ms | 1.00 | 1.00 | 15154.1 KB | 1.00 | 1385.0 KB | 1.00 | Win |
| 25000 | package-profile | write-datatable-direct | MiniExcel | 99.34 ms | 4.71 ms | 2.72 ms | 2.21 | 2.21 | 92384.2 KB | 6.10 | 1521.1 KB | 1.10 | 121.2% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | EPPlus | 337.50 ms | 7.39 ms | 4.27 ms | 7.51 | 7.51 | 117378.3 KB | 7.75 | 1090.8 KB | 0.79 | 651.4% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-direct | ClosedXML | 353.51 ms | 13.44 ms | 7.76 ms | 7.87 | 7.87 | 104197.0 KB | 6.88 | 1139.9 KB | 0.82 | 687.0% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-table-direct | OfficeIMO.Excel | 49.60 ms | 2.34 ms | 1.35 ms | 1.00 | 1.00 | 15617.5 KB | 1.00 | 1385.8 KB | 1.00 | Win |
| 25000 | package-profile | write-datatable-table-direct | MiniExcel | 99.11 ms | 9.42 ms | 5.44 ms | 2.00 | 2.00 | 92384.5 KB | 5.92 | 1521.0 KB | 1.10 | 99.8% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-table-direct | EPPlus | 317.06 ms | 4.72 ms | 2.72 ms | 6.39 | 6.39 | 117378.5 KB | 7.52 | 1090.8 KB | 0.79 | 539.2% slower than OfficeIMO |
| 25000 | package-profile | write-datatable-table-direct | ClosedXML | 505.15 ms | 12.14 ms | 7.01 ms | 10.18 | 10.18 | 173392.7 KB | 11.10 | 1140.7 KB | 0.82 | 918.5% slower than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 51.57 ms | 8.33 ms | 4.81 ms | 1.00 | 1.00 | 15547.2 KB | 1.00 | 1385.0 KB | 1.00 | Win |
| 25000 | package-profile | write-fluent-rowsfrom-direct | MiniExcel | 88.15 ms | 5.65 ms | 3.26 ms | 1.71 | 1.71 | 93246.9 KB | 6.00 | 1521.0 KB | 1.10 | 71.0% slower than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | EPPlus | 324.70 ms | 8.25 ms | 4.76 ms | 6.30 | 6.30 | 117378.4 KB | 7.55 | 1090.8 KB | 0.79 | 529.7% slower than OfficeIMO |
| 25000 | package-profile | write-fluent-rowsfrom-direct | ClosedXML | 341.71 ms | 7.60 ms | 4.39 ms | 6.63 | 6.63 | 104197.0 KB | 6.70 | 1139.9 KB | 0.82 | 562.7% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | OfficeIMO.Excel | 45.40 ms | 3.52 ms | 2.03 ms | 1.00 | 1.00 | 15350.7 KB | 1.00 | 1385.0 KB | 1.00 | Win |
| 25000 | package-profile | write-insertobjects-direct | MiniExcel | 87.89 ms | 4.58 ms | 2.64 ms | 1.94 | 1.94 | 93246.9 KB | 6.07 | 1521.1 KB | 1.10 | 93.6% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | EPPlus | 316.89 ms | 13.34 ms | 7.70 ms | 6.98 | 6.98 | 117378.4 KB | 7.65 | 1090.8 KB | 0.79 | 598.0% slower than OfficeIMO |
| 25000 | package-profile | write-insertobjects-direct | ClosedXML | 344.21 ms | 11.48 ms | 6.63 ms | 7.58 | 7.58 | 104197.0 KB | 6.79 | 1139.9 KB | 0.82 | 658.2% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | OfficeIMO.Excel | 20.10 ms | 0.45 ms | 0.26 ms | 1.00 | 1.00 | 14442.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | append-plain-rows | MiniExcel | 32.83 ms | 4.40 ms | 2.54 ms | 1.63 | 1.63 | 58242.9 KB | 4.03 |  |  | 63.3% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | EPPlus 4.5.3.3 | 128.73 ms |  |  | 6.40 | 6.40 |  |  |  |  | 540.4% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | EPPlus | 146.05 ms | 6.16 ms | 3.56 ms | 7.27 | 7.27 | 100373.9 KB | 6.95 |  |  | 626.5% slower than OfficeIMO |
| 25000 | speed-comparison | append-plain-rows | ClosedXML | 149.54 ms | 12.66 ms | 7.31 ms | 7.44 | 7.44 | 104233.1 KB | 7.22 |  |  | 643.9% slower than OfficeIMO |
| 25000 | speed-comparison | autofit-existing | EPPlus | 522.02 ms | 34.54 ms | 19.94 ms | 0.97 | 1.00 | 250948.8 KB | 1.84 |  |  | 3.1% faster than OfficeIMO |
| 25000 | speed-comparison | autofit-existing | OfficeIMO.Excel | 538.56 ms | 33.63 ms | 19.42 ms | 1.00 | 1.03 | 136146.9 KB | 1.00 |  |  | Loss +3.2% |
| 25000 | speed-comparison | autofit-existing | EPPlus 4.5.3.3 | 736.34 ms |  |  | 1.37 | 1.41 |  |  |  |  | 36.7% slower than OfficeIMO |
| 25000 | speed-comparison | autofit-existing | ClosedXML | 1862.23 ms | 72.25 ms | 41.71 ms | 3.46 | 3.57 | 829986.4 KB | 6.10 |  |  | 245.8% slower than OfficeIMO |
| 25000 | speed-comparison | formula-heavy-read | OfficeIMO.Excel | 39.09 ms | 4.06 ms | 2.35 ms | 1.00 | 1.00 | 15474.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | formula-heavy-read | EPPlus 4.5.3.3 | 79.28 ms |  |  | 2.03 | 2.03 |  |  |  |  | 102.8% slower than OfficeIMO |
| 25000 | speed-comparison | formula-heavy-read | EPPlus | 121.16 ms | 12.12 ms | 7.00 ms | 3.10 | 3.10 | 75783.0 KB | 4.90 |  |  | 209.9% slower than OfficeIMO |
| 25000 | speed-comparison | formula-heavy-read | ClosedXML | 265.01 ms | 2.84 ms | 1.64 ms | 6.78 | 6.78 | 89830.2 KB | 5.81 |  |  | 577.9% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | OfficeIMO.Excel | 22.12 ms | 2.22 ms | 1.28 ms | 1.00 | 1.00 | 20476.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | large-shared-strings | MiniExcel | 37.28 ms | 5.83 ms | 3.37 ms | 1.69 | 1.69 | 73760.2 KB | 3.60 |  |  | 68.6% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | EPPlus 4.5.3.3 | 105.19 ms |  |  | 4.76 | 4.76 |  |  |  |  | 375.6% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | ClosedXML | 139.09 ms | 12.62 ms | 7.29 ms | 6.29 | 6.29 | 104241.3 KB | 5.09 |  |  | 528.9% slower than OfficeIMO |
| 25000 | speed-comparison | large-shared-strings | EPPlus | 140.95 ms | 4.49 ms | 2.59 ms | 6.37 | 6.37 | 84410.3 KB | 4.12 |  |  | 537.3% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | Sylvan.Data.Excel | 0.99 ms | 0.10 ms | 0.06 ms | 0.60 | 1.00 | 316.5 KB | 1.10 |  |  | 39.6% faster than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | OfficeIMO.Excel | 1.64 ms | 0.06 ms | 0.03 ms | 1.00 | 1.65 | 289.0 KB | 1.00 |  |  | Loss +65.5% |
| 25000 | speed-comparison | large-sparse-column-read | ClosedXML | 3.20 ms | 0.11 ms | 0.06 ms | 1.95 | 3.23 | 4392.6 KB | 15.20 |  |  | 95.1% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | EPPlus 4.5.3.3 | 11.34 ms |  |  | 6.92 | 11.46 |  |  |  |  | 592.5% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | MiniExcel | 11.60 ms | 0.58 ms | 0.34 ms | 7.08 | 11.72 | 46194.6 KB | 159.87 |  |  | 608.2% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-column-read | EPPlus | 30.02 ms | 1.13 ms | 0.65 ms | 18.32 | 30.33 | 43070.8 KB | 149.06 |  |  | 1732.5% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | Sylvan.Data.Excel | 0.93 ms | 0.03 ms | 0.02 ms | 0.54 | 1.00 | 316.5 KB | 1.10 |  |  | 46.2% faster than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | OfficeIMO.Excel | 1.72 ms | 0.10 ms | 0.06 ms | 1.00 | 1.86 | 289.0 KB | 1.00 |  |  | Loss +85.7% |
| 25000 | speed-comparison | large-sparse-row-read | ClosedXML | 3.09 ms | 0.13 ms | 0.07 ms | 1.79 | 3.33 | 4392.3 KB | 15.20 |  |  | 79.4% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | MiniExcel | 11.21 ms | 0.37 ms | 0.21 ms | 6.51 | 12.09 | 46194.6 KB | 159.83 |  |  | 551.0% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | EPPlus 4.5.3.3 | 15.24 ms |  |  | 8.84 | 16.42 |  |  |  |  | 784.4% slower than OfficeIMO |
| 25000 | speed-comparison | large-sparse-row-read | EPPlus | 28.12 ms | 1.37 ms | 0.79 ms | 16.32 | 30.32 | 43070.8 KB | 149.02 |  |  | 1532.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-datatable | Sylvan.Data.Excel | 53.21 ms | 0.62 ms | 0.36 ms | 0.17 | 1.00 | 15257.9 KB | 0.07 |  |  | 83.3% faster than OfficeIMO |
| 25000 | speed-comparison | read-datatable | MiniExcel | 165.32 ms | 8.58 ms | 4.95 ms | 0.52 | 3.11 | 184822.9 KB | 0.85 |  |  | 48.1% faster than OfficeIMO |
| 25000 | speed-comparison | read-datatable | EPPlus | 190.23 ms | 18.15 ms | 10.48 ms | 0.60 | 3.58 | 190927.1 KB | 0.88 |  |  | 40.2% faster than OfficeIMO |
| 25000 | speed-comparison | read-datatable | EPPlus 4.5.3.3 | 216.64 ms |  |  | 0.68 | 4.07 |  |  |  |  | 31.9% faster than OfficeIMO |
| 25000 | speed-comparison | read-datatable | OfficeIMO.Excel | 318.28 ms | 11.59 ms | 6.69 ms | 1.00 | 5.98 | 218093.2 KB | 1.00 |  |  | Loss +498.2% |
| 25000 | speed-comparison | read-datatable | ClosedXML | 368.23 ms | 15.14 ms | 8.74 ms | 1.16 | 6.92 | 213768.5 KB | 0.98 |  |  | 15.7% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects | Sylvan.Data.Excel | 63.73 ms | 6.29 ms | 3.63 ms | 0.18 | 1.00 | 2156.3 KB | 0.01 |  |  | 81.8% faster than OfficeIMO |
| 25000 | speed-comparison | read-objects | EPPlus 4.5.3.3 | 196.60 ms |  |  | 0.56 | 3.08 |  |  |  |  | 43.8% faster than OfficeIMO |
| 25000 | speed-comparison | read-objects | MiniExcel | 196.94 ms | 9.15 ms | 5.28 ms | 0.56 | 3.09 | 186681.9 KB | 0.89 |  |  | 43.7% faster than OfficeIMO |
| 25000 | speed-comparison | read-objects | EPPlus | 240.46 ms | 12.83 ms | 7.41 ms | 0.69 | 3.77 | 188089.7 KB | 0.89 |  |  | 31.3% faster than OfficeIMO |
| 25000 | speed-comparison | read-objects | OfficeIMO.Excel | 350.08 ms | 48.37 ms | 27.93 ms | 1.00 | 5.49 | 210710.2 KB | 1.00 |  |  | Loss +449.3% |
| 25000 | speed-comparison | read-objects | ClosedXML | 425.86 ms | 75.84 ms | 43.79 ms | 1.22 | 6.68 | 200662.2 KB | 0.95 |  |  | 21.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | Sylvan.Data.Excel | 56.34 ms | 4.57 ms | 2.64 ms | 0.67 | 1.00 | 2156.3 KB | 0.21 |  |  | 33.0% faster than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | OfficeIMO.Excel | 84.10 ms | 4.22 ms | 2.44 ms | 1.00 | 1.49 | 10347.7 KB | 1.00 |  |  | Loss +49.3% |
| 25000 | speed-comparison | read-objects-stream | MiniExcel | 190.21 ms | 7.15 ms | 4.13 ms | 2.26 | 3.38 | 186679.3 KB | 18.04 |  |  | 126.2% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | EPPlus 4.5.3.3 | 205.10 ms |  |  | 2.44 | 3.64 |  |  |  |  | 143.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | EPPlus | 232.97 ms | 5.17 ms | 2.98 ms | 2.77 | 4.14 | 188089.6 KB | 18.18 |  |  | 177.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-objects-stream | ClosedXML | 467.43 ms | 26.05 ms | 15.04 ms | 5.56 | 8.30 | 200659.7 KB | 19.39 |  |  | 455.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-range | Sylvan.Data.Excel | 41.41 ms | 0.77 ms | 0.45 ms | 0.15 | 1.00 | 398.5 KB | 0.00 |  |  | 84.8% faster than OfficeIMO |
| 25000 | speed-comparison | read-range | MiniExcel | 132.04 ms | 3.48 ms | 2.01 ms | 0.48 | 3.19 | 189958.9 KB | 0.91 |  |  | 51.6% faster than OfficeIMO |
| 25000 | speed-comparison | read-range | EPPlus | 163.18 ms | 5.27 ms | 3.04 ms | 0.60 | 3.94 | 176067.7 KB | 0.84 |  |  | 40.2% faster than OfficeIMO |
| 25000 | speed-comparison | read-range | EPPlus 4.5.3.3 | 206.61 ms |  |  | 0.76 | 4.99 |  |  |  |  | 24.3% faster than OfficeIMO |
| 25000 | speed-comparison | read-range | OfficeIMO.Excel | 272.88 ms | 14.91 ms | 8.61 ms | 1.00 | 6.59 | 209159.5 KB | 1.00 |  |  | Loss +558.9% |
| 25000 | speed-comparison | read-range | ClosedXML | 343.34 ms | 5.85 ms | 3.38 ms | 1.26 | 8.29 | 198908.3 KB | 0.95 |  |  | 25.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | Sylvan.Data.Excel | 42.95 ms | 1.86 ms | 1.07 ms | 0.55 | 1.00 | 398.5 KB | 0.01 |  |  | 45.3% faster than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | OfficeIMO.Excel | 78.52 ms | 1.33 ms | 0.77 ms | 1.00 | 1.83 | 27067.8 KB | 1.00 |  |  | Loss +82.8% |
| 25000 | speed-comparison | read-range-stream | MiniExcel | 130.80 ms | 2.10 ms | 1.21 ms | 1.67 | 3.05 | 189956.2 KB | 7.02 |  |  | 66.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | EPPlus | 161.47 ms | 3.01 ms | 1.74 ms | 2.06 | 3.76 | 176067.7 KB | 6.50 |  |  | 105.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | EPPlus 4.5.3.3 | 194.27 ms |  |  | 2.47 | 4.52 |  |  |  |  | 147.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-range-stream | ClosedXML | 342.39 ms | 10.84 ms | 6.26 ms | 4.36 | 7.97 | 198904.7 KB | 7.35 |  |  | 336.0% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | Sylvan.Data.Excel | 0.43 ms | 0.02 ms | 0.01 ms | 0.21 | 1.00 | 365.5 KB | 0.33 |  |  | 78.8% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range | MiniExcel | 0.93 ms | 0.23 ms | 0.13 ms | 0.46 | 2.16 | 902.0 KB | 0.81 |  |  | 54.1% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range | OfficeIMO.Excel | 2.02 ms | 0.09 ms | 0.05 ms | 1.00 | 4.71 | 1111.0 KB | 1.00 |  |  | Loss +370.9% |
| 25000 | speed-comparison | read-top-range | EPPlus | 128.29 ms | 1.69 ms | 0.98 ms | 63.54 | 299.21 | 94254.6 KB | 84.84 |  |  | 6254.4% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | EPPlus 4.5.3.3 | 168.82 ms |  |  | 83.61 | 393.72 |  |  |  |  | 8261.5% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range | ClosedXML | 337.49 ms | 6.35 ms | 3.67 ms | 167.16 | 787.12 | 184698.1 KB | 166.25 |  |  | 16615.9% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | Sylvan.Data.Excel | 0.40 ms | 0.01 ms | 0.00 ms | 0.19 | 1.00 | 365.5 KB | 0.33 |  |  | 81.0% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | MiniExcel | 0.81 ms | 0.02 ms | 0.01 ms | 0.39 | 2.05 | 901.8 KB | 0.81 |  |  | 61.2% faster than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | OfficeIMO.Excel | 2.09 ms | 0.05 ms | 0.03 ms | 1.00 | 5.27 | 1114.2 KB | 1.00 |  |  | Loss +427.2% |
| 25000 | speed-comparison | read-top-range-stream | EPPlus | 126.90 ms | 2.77 ms | 1.60 ms | 60.72 | 320.11 | 94254.6 KB | 84.59 |  |  | 5971.8% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | EPPlus 4.5.3.3 | 163.48 ms |  |  | 78.22 | 412.37 |  |  |  |  | 7721.6% slower than OfficeIMO |
| 25000 | speed-comparison | read-top-range-stream | ClosedXML | 329.82 ms | 11.26 ms | 6.50 ms | 157.80 | 831.96 | 184703.1 KB | 165.77 |  |  | 15680.3% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | Sylvan.Data.Excel | 30.52 ms | 0.74 ms | 0.43 ms | 0.43 | 1.00 | 2443.8 KB | 0.09 |  |  | 57.0% faster than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | OfficeIMO.Excel | 71.03 ms | 4.59 ms | 2.65 ms | 1.00 | 2.33 | 25985.7 KB | 1.00 |  |  | Loss +132.7% |
| 25000 | speed-comparison | shared-string-read | MiniExcel | 84.00 ms | 3.64 ms | 2.10 ms | 1.18 | 2.75 | 96062.6 KB | 3.70 |  |  | 18.3% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | EPPlus 4.5.3.3 | 94.46 ms |  |  | 1.33 | 3.09 |  |  |  |  | 33.0% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | EPPlus | 111.82 ms | 4.02 ms | 2.32 ms | 1.57 | 3.66 | 98547.5 KB | 3.79 |  |  | 57.4% slower than OfficeIMO |
| 25000 | speed-comparison | shared-string-read | ClosedXML | 266.43 ms | 30.25 ms | 17.46 ms | 3.75 | 8.73 | 111981.4 KB | 4.31 |  |  | 275.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | MiniExcel | 85.78 ms | 1.61 ms | 0.93 ms | 0.90 | 1.00 | 125550.4 KB | 3.77 |  |  | 10.0% faster than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | OfficeIMO.Excel | 95.32 ms | 7.65 ms | 4.42 ms | 1.00 | 1.11 | 33266.8 KB | 1.00 |  |  | Loss +11.1% |
| 25000 | speed-comparison | write-bulk-report | EPPlus | 382.40 ms | 21.26 ms | 12.27 ms | 4.01 | 4.46 | 254959.1 KB | 7.66 |  |  | 301.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | EPPlus 4.5.3.3 | 589.05 ms |  |  | 6.18 | 6.87 |  |  |  |  | 518.0% slower than OfficeIMO |
| 25000 | speed-comparison | write-bulk-report | ClosedXML | 949.72 ms | 38.80 ms | 22.40 ms | 9.96 | 11.07 | 565957.6 KB | 17.01 |  |  | 896.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 42.90 ms | 4.33 ms | 2.50 ms | 1.00 | 1.00 | 18483.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | MiniExcel | 68.98 ms | 3.67 ms | 2.12 ms | 1.61 | 1.61 | 93257.0 KB | 5.05 |  |  | 60.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 294.17 ms |  |  | 6.86 | 6.86 |  |  |  |  | 585.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | EPPlus | 326.43 ms | 16.18 ms | 9.34 ms | 7.61 | 7.61 | 211850.4 KB | 11.46 |  |  | 660.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-cellvalues-rectangle-direct | ClosedXML | 333.93 ms | 14.67 ms | 8.47 ms | 7.78 | 7.78 | 210646.1 KB | 11.40 |  |  | 678.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | OfficeIMO.Excel | 90.01 ms | 5.93 ms | 3.42 ms | 1.00 | 1.00 | 25799.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write-datareader-table | MiniExcel | 98.91 ms | 0.28 ms | 0.16 ms | 1.10 | 1.10 | 92200.0 KB | 3.57 |  |  | 9.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | EPPlus | 354.24 ms | 30.24 ms | 17.46 ms | 3.94 | 3.94 | 117437.8 KB | 4.55 |  |  | 293.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table | ClosedXML | 515.32 ms | 30.21 ms | 17.44 ms | 5.73 | 5.73 | 173399.3 KB | 6.72 |  |  | 472.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-datareader-table-direct | EPPlus 4.5.3.3 | 281.55 ms |  |  |  | 1.00 |  |  |  |  |  |
| 25000 | speed-comparison | write-dataset-headerless-tables | OfficeIMO.Excel | 49.32 ms | 4.27 ms | 2.47 ms | 1.00 | 1.00 | 16367.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write-dataset-headerless-tables | MiniExcel | 94.37 ms | 9.83 ms | 5.67 ms | 1.91 | 1.91 | 97086.9 KB | 5.93 |  |  | 91.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-headerless-tables | EPPlus | 423.09 ms | 21.29 ms | 12.29 ms | 8.58 | 8.58 | 111246.0 KB | 6.80 |  |  | 757.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-headerless-tables | ClosedXML | 465.36 ms | 35.24 ms | 20.35 ms | 9.44 | 9.44 | 172020.3 KB | 10.51 |  |  | 843.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-tables | OfficeIMO.Excel | 63.65 ms | 3.05 ms | 1.76 ms | 1.00 | 1.00 | 22530.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write-dataset-sparse-tables | MiniExcel | 125.81 ms | 4.79 ms | 2.76 ms | 1.98 | 1.98 | 108129.1 KB | 4.80 |  |  | 97.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-tables | EPPlus | 544.87 ms | 12.64 ms | 7.30 ms | 8.56 | 8.56 | 135724.0 KB | 6.02 |  |  | 756.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-sparse-tables | ClosedXML | 717.71 ms | 67.95 ms | 39.23 ms | 11.28 | 11.28 | 280381.4 KB | 12.44 |  |  | 1027.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | OfficeIMO.Excel | 43.40 ms | 4.50 ms | 2.60 ms | 1.00 | 1.00 | 13607.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write-dataset-tables | MiniExcel | 85.18 ms | 3.69 ms | 2.13 ms | 1.96 | 1.96 | 97084.5 KB | 7.13 |  |  | 96.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | EPPlus 4.5.3.3 | 279.62 ms |  |  | 6.44 | 6.44 |  |  |  |  | 544.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | EPPlus | 325.72 ms | 9.25 ms | 5.34 ms | 7.50 | 7.50 | 110816.2 KB | 8.14 |  |  | 650.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables | ClosedXML | 455.16 ms | 38.59 ms | 22.28 ms | 10.49 | 10.49 | 172003.9 KB | 12.64 |  |  | 948.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables-autofit | OfficeIMO.Excel | 45.41 ms | 1.66 ms | 0.96 ms | 1.00 | 1.00 | 13608.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write-dataset-tables-autofit | MiniExcel | 111.39 ms | 4.61 ms | 2.66 ms | 2.45 | 2.45 | 128873.9 KB | 9.47 |  |  | 145.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables-autofit | EPPlus | 406.05 ms | 8.26 ms | 4.77 ms | 8.94 | 8.94 | 195407.7 KB | 14.36 |  |  | 794.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-dataset-tables-autofit | ClosedXML | 908.10 ms | 64.96 ms | 37.51 ms | 20.00 | 20.00 | 550094.9 KB | 40.42 |  |  | 1899.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | OfficeIMO.Excel | 46.54 ms | 6.55 ms | 3.78 ms | 1.00 | 1.00 | 15162.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write-datatable-direct | MiniExcel | 92.57 ms | 2.62 ms | 1.51 ms | 1.99 | 1.99 | 92394.2 KB | 6.09 |  |  | 98.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | EPPlus 4.5.3.3 | 282.34 ms |  |  | 6.07 | 6.07 |  |  |  |  | 506.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | ClosedXML | 318.75 ms | 7.67 ms | 4.43 ms | 6.85 | 6.85 | 104205.0 KB | 6.87 |  |  | 584.9% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-direct | EPPlus | 319.02 ms | 17.04 ms | 9.84 ms | 6.85 | 6.85 | 117437.7 KB | 7.75 |  |  | 585.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | OfficeIMO.Excel | 47.53 ms | 5.83 ms | 3.37 ms | 1.00 | 1.00 | 15625.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write-datatable-table-direct | MiniExcel | 102.05 ms | 11.51 ms | 6.65 ms | 2.15 | 2.15 | 92394.5 KB | 5.91 |  |  | 114.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | EPPlus 4.5.3.3 | 306.70 ms |  |  | 6.45 | 6.45 |  |  |  |  | 545.3% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | EPPlus | 325.46 ms | 15.87 ms | 9.16 ms | 6.85 | 6.85 | 117437.7 KB | 7.52 |  |  | 584.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-datatable-table-direct | ClosedXML | 487.97 ms | 20.02 ms | 11.56 ms | 10.27 | 10.27 | 173399.7 KB | 11.10 |  |  | 926.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 37.11 ms | 2.55 ms | 1.47 ms | 1.00 | 1.00 | 15555.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | MiniExcel | 64.01 ms | 2.73 ms | 1.58 ms | 1.72 | 1.72 | 93257.0 KB | 6.00 |  |  | 72.5% slower than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 254.84 ms |  |  | 6.87 | 6.87 |  |  |  |  | 586.6% slower than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | EPPlus | 263.66 ms | 8.48 ms | 4.89 ms | 7.10 | 7.10 | 117437.7 KB | 7.55 |  |  | 610.4% slower than OfficeIMO |
| 25000 | speed-comparison | write-fluent-rowsfrom-direct | ClosedXML | 284.23 ms | 17.79 ms | 10.27 ms | 7.66 | 7.66 | 104205.0 KB | 6.70 |  |  | 665.8% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | OfficeIMO.Excel | 34.40 ms | 0.82 ms | 0.47 ms | 1.00 | 1.00 | 15358.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write-insertobjects-direct | MiniExcel | 70.24 ms | 1.89 ms | 1.09 ms | 2.04 | 2.04 | 93257.0 KB | 6.07 |  |  | 104.2% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | EPPlus | 267.17 ms | 4.87 ms | 2.81 ms | 7.77 | 7.77 | 117437.7 KB | 7.65 |  |  | 676.7% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | EPPlus 4.5.3.3 | 273.84 ms |  |  | 7.96 | 7.96 |  |  |  |  | 696.1% slower than OfficeIMO |
| 25000 | speed-comparison | write-insertobjects-direct | ClosedXML | 293.01 ms | 18.45 ms | 10.65 ms | 8.52 | 8.52 | 104205.0 KB | 6.78 |  |  | 751.9% slower than OfficeIMO |
