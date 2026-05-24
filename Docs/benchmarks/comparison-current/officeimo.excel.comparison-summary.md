# OfficeIMO.Excel Comparison Summary

This is the suite-level decision table. Mean, standard deviation, standard error, ratios, and allocations come from the lightweight rotated local runner; they are meant for engineering direction. Use the BenchmarkDotNet benchmark classes when a publication-grade `Error` column is required.

## At a glance

| Row count | Artifact | Workload | Category | OfficeIMO wins | OfficeIMO losses | Biggest loss |
| ---: | --- | --- | --- | ---: | ---: | --- |
| 2500 | dense-helloworld-comparison | read | Other | 0 | 2 | dense-helloworld-read-range: Loss +51.3% vs Sylvan.Data.Excel |
| 2500 | package-profile | package | Package size | 27 | 8 | write-cellvalues-rectangle-direct: Loss +56.9% vs LargeXlsx |
| 2500 | speed-comparison | mutate | AutoFit and mutation | 1 | 0 |  |
| 2500 | speed-comparison | other | Object projection | 2 | 0 |  |
| 2500 | speed-comparison | other | Range and table read | 4 | 0 |  |
| 2500 | speed-comparison | read | Formula write/read | 1 | 0 |  |
| 2500 | speed-comparison | read | Other | 0 | 3 | large-sparse-row-read: Loss +165.5% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | read | Range and table read | 1 | 6 | read-top-range: Loss +322.8% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | read | Streaming read | 0 | 4 | read-top-range-stream-small-chunks: Loss +279.7% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | read | Typed object read | 0 | 2 | read-objects-stream: Loss +28.6% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | write | AutoFit and mutation | 5 | 0 |  |
| 2500 | speed-comparison | write | Cell writer | 7 | 0 |  |
| 2500 | speed-comparison | write | DataSet table export | 3 | 0 |  |
| 2500 | speed-comparison | write | DataTable table export | 4 | 0 |  |
| 2500 | speed-comparison | write | Formatted report write | 1 | 0 |  |
| 2500 | speed-comparison | write | Formula write/read | 1 | 0 |  |
| 2500 | speed-comparison | write | Other | 1 | 1 | write-powershell-mixed-objects-direct: Loss +8.0% vs LargeXlsx |
| 2500 | speed-comparison | write | Plain cell export | 1 | 3 | write-cellvalues-headerless-rectangle-direct: Loss +153.8% vs LargeXlsx |
| 2500 | speed-comparison | write | Plain streaming export | 1 | 1 | write-datareader-plain: Loss +37.0% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | write | Plain string export | 0 | 1 | write-blog-2023-20-string-columns: Loss +13.2% vs LargeXlsx |
| 2500 | speed-comparison | write | Shared string write | 4 | 0 |  |
| 2500 | speed-comparison | write | Typed object export | 0 | 3 | write-insertobjects-flat-dictionaries-direct: Loss +34.7% vs LargeXlsx |
| 25000 | dense-helloworld-comparison | read | Other | 0 | 2 | dense-helloworld-read-range: Loss +24.4% vs Sylvan.Data.Excel |
| 25000 | package-profile | package | Package size | 24 | 11 | append-plain-rows: Loss +52.4% vs LargeXlsx |
| 25000 | speed-comparison | mutate | AutoFit and mutation | 1 | 0 |  |
| 25000 | speed-comparison | other | Object projection | 2 | 0 |  |
| 25000 | speed-comparison | other | Range and table read | 4 | 0 |  |
| 25000 | speed-comparison | read | Formula write/read | 1 | 0 |  |
| 25000 | speed-comparison | read | Other | 0 | 3 | large-sparse-column-read: Loss +72.4% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | read | Range and table read | 2 | 5 | read-top-range: Loss +342.4% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | read | Streaming read | 1 | 3 | read-top-range-stream: Loss +353.0% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | read | Typed object read | 0 | 2 | read-objects-stream: Loss +13.3% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | write | AutoFit and mutation | 5 | 0 |  |
| 25000 | speed-comparison | write | Cell writer | 7 | 0 |  |
| 25000 | speed-comparison | write | DataSet table export | 3 | 0 |  |
| 25000 | speed-comparison | write | DataTable table export | 3 | 1 | write-datatable-direct: Loss +6.6% vs LargeXlsx |
| 25000 | speed-comparison | write | Formatted report write | 1 | 0 |  |
| 25000 | speed-comparison | write | Formula write/read | 1 | 0 |  |
| 25000 | speed-comparison | write | Other | 1 | 1 | write-powershell-mixed-objects-direct: Loss +18.3% vs LargeXlsx |
| 25000 | speed-comparison | write | Plain cell export | 1 | 3 | append-plain-rows: Loss +55.9% vs LargeXlsx |
| 25000 | speed-comparison | write | Plain streaming export | 1 | 1 | write-datareader-plain: Loss +28.9% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | write | Plain string export | 0 | 1 | write-blog-2023-20-string-columns: Loss +10.5% vs LargeXlsx |
| 25000 | speed-comparison | write | Shared string write | 4 | 0 |  |
| 25000 | speed-comparison | write | Typed object export | 0 | 3 | write-insertobjects-flat-dictionaries-direct: Loss +31.5% vs LargeXlsx |

## OfficeIMO decision table

| Row count | Artifact | Workload | Category | Scenario | OfficeIMO mean | Best | OfficeIMO vs best | Alloc | Package |
| ---: | --- | --- | --- | --- | ---: | --- | ---: | ---: | ---: |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | 9.70 ms | Sylvan.Data.Excel | Loss +51.3% | 2496.7 KB |  |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | 9.12 ms | Sylvan.Data.Excel | Loss +46.0% | 2575.0 KB |  |
| 2500 | package-profile | package | Package size | append-plain-rows | 3.17 ms | LargeXlsx | Loss +44.7% | 1658.1 KB | 64.5 KB |
| 2500 | package-profile | package | Package size | autofit-existing | 43.26 ms | OfficeIMO.Excel | Win | 13884.7 KB | 138.9 KB |
| 2500 | package-profile | package | Package size | large-shared-strings | 3.14 ms | OfficeIMO.Excel | Win | 2112.6 KB | 55.2 KB |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | 6.54 ms | LargeXlsx | Loss +20.8% | 2020.1 KB | 216.7 KB |
| 2500 | package-profile | package | Package size | write-bulk-report | 6.85 ms | OfficeIMO.Excel | Win | 1498.6 KB | 138.9 KB |
| 2500 | package-profile | package | Package size | write-cellformula | 4.43 ms | OfficeIMO.Excel | Win | 1173.1 KB | 66.6 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-empty-strings | 3.02 ms | OfficeIMO.Excel | Win | 1446.7 KB | 44.2 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-numbers | 2.37 ms | OfficeIMO.Excel | Win | 939.5 KB | 47.0 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-object-mixed | 4.60 ms | OfficeIMO.Excel | Win | 1424.1 KB | 61.2 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse | 5.74 ms | OfficeIMO.Excel | Win | 1263.7 KB | 62.0 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | 4.42 ms | OfficeIMO.Excel | Win | 1271.8 KB | 62.0 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-scalars | 2.83 ms | OfficeIMO.Excel | Win | 957.6 KB | 46.3 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-strings | 4.70 ms | OfficeIMO.Excel | Win | 2276.5 KB | 55.1 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-distinct | 3.01 ms | OfficeIMO.Excel | Win | 2198.8 KB | 51.8 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-repeated | 2.40 ms | OfficeIMO.Excel | Win | 1239.4 KB | 40.0 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-temporal | 5.14 ms | OfficeIMO.Excel | Win | 1255.3 KB | 63.3 KB |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | 2.04 ms | LargeXlsx | Loss +13.1% | 1126.8 KB | 48.2 KB |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | 4.90 ms | LargeXlsx | Loss +56.9% | 1745.4 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-datareader-plain | 4.00 ms | Sylvan.Data.Excel | Loss +14.3% | 1427.2 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-datareader-table | 4.24 ms | OfficeIMO.Excel | Win | 1439.1 KB | 138.8 KB |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | 4.37 ms | OfficeIMO.Excel | Win | 1445.4 KB | 138.9 KB |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | 5.18 ms | OfficeIMO.Excel | Win | 1645.6 KB | 131.1 KB |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | 5.80 ms | OfficeIMO.Excel | Win | 2384.8 KB | 176.0 KB |
| 2500 | package-profile | package | Package size | write-dataset-tables | 5.77 ms | OfficeIMO.Excel | Win | 1570.8 KB | 138.9 KB |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | 5.70 ms | OfficeIMO.Excel | Win | 1583.4 KB | 139.2 KB |
| 2500 | package-profile | package | Package size | write-datatable-direct | 4.01 ms | OfficeIMO.Excel | Win | 1412.9 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | 3.88 ms | OfficeIMO.Excel | Win | 1424.9 KB | 138.8 KB |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | 5.36 ms | OfficeIMO.Excel | Win | 1441.5 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | 8.61 ms | OfficeIMO.Excel | Win | 1450.8 KB | 138.1 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | 7.60 ms | LargeXlsx | Loss +4.5% | 1443.1 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | 7.17 ms | OfficeIMO.Excel | Win | 749.9 KB | 142.4 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | 5.36 ms | LargeXlsx | Loss +6.9% | 742.1 KB | 142.3 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | 7.21 ms | OfficeIMO.Excel | Win | 1602.4 KB | 142.3 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | 6.90 ms | OfficeIMO.Excel | Win | 1450.1 KB | 138.1 KB |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | 6.46 ms | LargeXlsx | Loss +5.4% | 1085.0 KB | 182.4 KB |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | 30.04 ms | OfficeIMO.Excel | Win | 13748.7 KB |  |
| 2500 | speed-comparison | other | Object projection | build-object-datatable-dictionaries | 1.85 ms | OfficeIMO.Excel | Win | 564.2 KB |  |
| 2500 | speed-comparison | other | Object projection | build-object-datatable-typed | 1.25 ms | OfficeIMO.Excel | Win | 856.9 KB |  |
| 2500 | speed-comparison | other | Range and table read | enumerate-cells | 7.24 ms | OfficeIMO.Excel | Win | 2539.6 KB |  |
| 2500 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | 4.89 ms | OfficeIMO.Excel | Win | 679.5 KB |  |
| 2500 | speed-comparison | other | Range and table read | enumerate-range | 7.12 ms | OfficeIMO.Excel | Win | 2539.2 KB |  |
| 2500 | speed-comparison | other | Range and table read | enumerate-top-range | 1.68 ms | OfficeIMO.Excel | Win | 433.9 KB |  |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | 3.23 ms | OfficeIMO.Excel | Win | 777.5 KB |  |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | 1.78 ms | Sylvan.Data.Excel | Loss +60.7% | 248.9 KB |  |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | 2.66 ms | Sylvan.Data.Excel | Loss +165.5% | 249.0 KB |  |
| 2500 | speed-comparison | read | Other | shared-string-read | 3.31 ms | Sylvan.Data.Excel | Loss +61.8% | 1133.6 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | 4.62 ms | Sylvan.Data.Excel | Loss +10.8% | 516.7 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-datatable | 8.32 ms | Sylvan.Data.Excel | Loss +22.5% | 3596.9 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | 5.32 ms | Sylvan.Data.Excel | Loss +19.0% | 699.1 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-range | 12.34 ms | OfficeIMO.Excel | Win | 2694.9 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | 10.59 ms | Sylvan.Data.Excel | Loss +107.5% | 2753.4 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-top-range | 1.84 ms | Sylvan.Data.Excel | Loss +322.8% | 439.5 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-used-range | 15.44 ms | Sylvan.Data.Excel | Loss +99.7% | 3416.5 KB |  |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | 4.67 ms | Sylvan.Data.Excel | Loss +8.1% | 520.1 KB |  |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | 9.81 ms | Sylvan.Data.Excel | Loss +112.1% | 2773.8 KB |  |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | 1.61 ms | Sylvan.Data.Excel | Loss +271.8% | 442.9 KB |  |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | 1.57 ms | Sylvan.Data.Excel | Loss +279.7% | 443.7 KB |  |
| 2500 | speed-comparison | read | Typed object read | read-objects | 8.41 ms | Sylvan.Data.Excel | Loss +13.4% | 2444.5 KB |  |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | 6.48 ms | Sylvan.Data.Excel | Loss +28.6% | 2444.8 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | 6.85 ms | OfficeIMO.Excel | Win | 1445.4 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | 7.21 ms | OfficeIMO.Excel | Win | 1584.5 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | 5.75 ms | OfficeIMO.Excel | Win | 1442.8 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | 4.56 ms | OfficeIMO.Excel | Win | 741.8 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | 5.55 ms | OfficeIMO.Excel | Win | 1442.1 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | 3.70 ms | OfficeIMO.Excel | Win | 1446.7 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | 3.43 ms | OfficeIMO.Excel | Win | 939.5 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | 3.98 ms | OfficeIMO.Excel | Win | 1423.9 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | 3.47 ms | OfficeIMO.Excel | Win | 1263.5 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | 3.60 ms | OfficeIMO.Excel | Win | 1263.6 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | 2.89 ms | OfficeIMO.Excel | Win | 957.6 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | 3.74 ms | OfficeIMO.Excel | Win | 1255.1 KB |  |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | 6.68 ms | OfficeIMO.Excel | Win | 1569.1 KB |  |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | 8.30 ms | OfficeIMO.Excel | Win | 2384.8 KB |  |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | 6.48 ms | OfficeIMO.Excel | Win | 1572.4 KB |  |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | 5.97 ms | OfficeIMO.Excel | Win | 1439.1 KB |  |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | 5.01 ms | OfficeIMO.Excel | Win | 1412.9 KB |  |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | 5.74 ms | OfficeIMO.Excel | Win | 1151.5 KB |  |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | 5.71 ms | OfficeIMO.Excel | Win | 1424.9 KB |  |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | 5.87 ms | OfficeIMO.Excel | Win | 1500.1 KB |  |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | 3.98 ms | OfficeIMO.Excel | Win | 1164.5 KB |  |
| 2500 | speed-comparison | write | Other | write-dictionary-objects-table-direct | 7.62 ms | OfficeIMO.Excel | Win | 1715.7 KB |  |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | 4.50 ms | LargeXlsx | Loss +8.0% | 1077.0 KB |  |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | 2.53 ms | LargeXlsx | Loss +55.7% | 1650.0 KB |  |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | 4.64 ms | LargeXlsx | Loss +153.8% | 1126.8 KB |  |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | 6.33 ms | LargeXlsx | Loss +3.6% | 1745.4 KB |  |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | 3.58 ms | OfficeIMO.Excel | Win | 1158.2 KB |  |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | 6.44 ms | Sylvan.Data.Excel | Loss +37.0% | 1427.2 KB |  |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | 7.75 ms | OfficeIMO.Excel | Win | 1645.6 KB |  |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | 7.61 ms | LargeXlsx | Loss +13.2% | 2020.1 KB |  |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | 1.94 ms | OfficeIMO.Excel | Win | 2104.6 KB |  |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | 3.88 ms | OfficeIMO.Excel | Win | 2276.5 KB |  |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | 2.83 ms | OfficeIMO.Excel | Win | 2198.8 KB |  |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | 2.89 ms | OfficeIMO.Excel | Win | 1239.4 KB |  |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | 3.58 ms | LargeXlsx | Loss +19.6% | 1433.4 KB |  |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | 5.30 ms | LargeXlsx | Loss +17.2% | 1435.1 KB |  |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | 5.39 ms | LargeXlsx | Loss +34.7% | 734.1 KB |  |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | 68.71 ms | Sylvan.Data.Excel | Loss +24.4% | 23699.7 KB |  |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | 70.63 ms | Sylvan.Data.Excel | Loss +22.0% | 24482.0 KB |  |
| 25000 | package-profile | package | Package size | append-plain-rows | 24.44 ms | LargeXlsx | Loss +52.4% | 11664.3 KB | 622.5 KB |
| 25000 | package-profile | package | Package size | autofit-existing | 474.61 ms | OfficeIMO.Excel | Win | 136016.3 KB | 1385.9 KB |
| 25000 | package-profile | package | Package size | large-shared-strings | 17.93 ms | OfficeIMO.Excel | Win | 15409.2 KB | 529.7 KB |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | 61.99 ms | LargeXlsx | Loss +6.3% | 12754.6 KB | 2228.8 KB |
| 25000 | package-profile | package | Package size | write-bulk-report | 56.75 ms | OfficeIMO.Excel | Win | 12647.3 KB | 1385.9 KB |
| 25000 | package-profile | package | Package size | write-cellformula | 28.86 ms | OfficeIMO.Excel | Win | 9541.6 KB | 670.3 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-empty-strings | 17.57 ms | OfficeIMO.Excel | Win | 6716.0 KB | 451.4 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-numbers | 22.31 ms | OfficeIMO.Excel | Win | 5790.4 KB | 462.6 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-object-mixed | 27.22 ms | OfficeIMO.Excel | Win | 7993.7 KB | 585.1 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse | 29.93 ms | OfficeIMO.Excel | Win | 7173.0 KB | 607.1 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | 23.41 ms | OfficeIMO.Excel | Win | 7173.1 KB | 607.1 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-scalars | 16.95 ms | OfficeIMO.Excel | Win | 5964.2 KB | 441.9 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-strings | 23.06 ms | OfficeIMO.Excel | Win | 15019.9 KB | 527.8 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-distinct | 20.39 ms | OfficeIMO.Excel | Win | 13643.7 KB | 499.0 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-repeated | 18.04 ms | OfficeIMO.Excel | Win | 7184.9 KB | 376.0 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-temporal | 30.65 ms | OfficeIMO.Excel | Win | 7302.4 KB | 620.5 KB |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | 14.93 ms | LargeXlsx | Loss +10.4% | 7226.8 KB | 455.5 KB |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | 51.42 ms | LargeXlsx | Loss +19.0% | 15700.7 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-datareader-plain | 52.68 ms | Sylvan.Data.Excel | Loss +27.5% | 12666.6 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-datareader-table | 54.87 ms | OfficeIMO.Excel | Win | 12684.6 KB | 1385.7 KB |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | 58.90 ms | OfficeIMO.Excel | Win | 12690.9 KB | 1385.9 KB |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | 51.16 ms | OfficeIMO.Excel | Win | 9484.6 KB | 1329.2 KB |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | 59.60 ms | OfficeIMO.Excel | Win | 13123.2 KB | 1795.1 KB |
| 25000 | package-profile | package | Package size | write-dataset-tables | 55.79 ms | OfficeIMO.Excel | Win | 9792.8 KB | 1376.4 KB |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | 56.87 ms | OfficeIMO.Excel | Win | 9805.4 KB | 1376.7 KB |
| 25000 | package-profile | package | Package size | write-datatable-direct | 51.63 ms | LargeXlsx | Loss +11.0% | 12380.0 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | 54.26 ms | OfficeIMO.Excel | Win | 12398.1 KB | 1385.7 KB |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | 47.42 ms | LargeXlsx | Loss +24.6% | 12576.3 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | 52.64 ms | OfficeIMO.Excel | Win | 12585.6 KB | 1385.1 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | 48.22 ms | LargeXlsx | Loss +25.1% | 12577.9 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | 58.29 ms | OfficeIMO.Excel | Win | 7029.1 KB | 1428.5 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | 51.67 ms | LargeXlsx | Loss +34.0% | 7021.4 KB | 1428.4 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | 59.39 ms | LargeXlsx | Loss +46.0% | 15616.0 KB | 1428.4 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | 49.22 ms | OfficeIMO.Excel | Win | 12584.9 KB | 1385.0 KB |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | 60.97 ms | LargeXlsx | Loss +13.4% | 7226.3 KB | 1828.0 KB |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | 425.06 ms | OfficeIMO.Excel | Win | 136016.3 KB |  |
| 25000 | speed-comparison | other | Object projection | build-object-datatable-dictionaries | 14.84 ms | OfficeIMO.Excel | Win | 5164.3 KB |  |
| 25000 | speed-comparison | other | Object projection | build-object-datatable-typed | 10.41 ms | OfficeIMO.Excel | Win | 8093.8 KB |  |
| 25000 | speed-comparison | other | Range and table read | enumerate-cells | 56.30 ms | OfficeIMO.Excel | Win | 23226.7 KB |  |
| 25000 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | 144.98 ms | OfficeIMO.Excel | Win | 4000.4 KB |  |
| 25000 | speed-comparison | other | Range and table read | enumerate-range | 78.13 ms | OfficeIMO.Excel | Win | 23226.8 KB |  |
| 25000 | speed-comparison | other | Range and table read | enumerate-top-range | 2.65 ms | OfficeIMO.Excel | Win | 445.0 KB |  |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | 21.12 ms | OfficeIMO.Excel | Win | 6287.1 KB |  |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | 1.87 ms | Sylvan.Data.Excel | Loss +72.4% | 249.0 KB |  |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | 1.71 ms | Sylvan.Data.Excel | Loss +66.0% | 249.1 KB |  |
| 25000 | speed-comparison | read | Other | shared-string-read | 21.85 ms | Sylvan.Data.Excel | Loss +30.0% | 9295.1 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | 35.43 ms | OfficeIMO.Excel | Win | 1263.6 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-datatable | 60.55 ms | Sylvan.Data.Excel | Loss +9.2% | 33325.6 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | 45.05 ms | Sylvan.Data.Excel, OfficeIMO.Excel | Win | 4192.9 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-range | 95.74 ms | Sylvan.Data.Excel | Loss +47.6% | 24786.3 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | 92.90 ms | Sylvan.Data.Excel | Loss +17.7% | 25372.2 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-top-range | 1.92 ms | Sylvan.Data.Excel | Loss +342.4% | 447.6 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-used-range | 158.02 ms | Sylvan.Data.Excel | Loss +117.2% | 32779.2 KB |  |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | 36.18 ms | OfficeIMO.Excel | Win | 1261.7 KB |  |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | 53.27 ms | Sylvan.Data.Excel | Loss +18.6% | 25565.2 KB |  |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | 2.14 ms | Sylvan.Data.Excel | Loss +353.0% | 443.0 KB |  |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | 2.08 ms | Sylvan.Data.Excel | Loss +352.9% | 443.7 KB |  |
| 25000 | speed-comparison | read | Typed object read | read-objects | 50.18 ms | Sylvan.Data.Excel | Loss +13.2% | 22242.0 KB |  |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | 72.25 ms | Sylvan.Data.Excel | Loss +13.3% | 22242.4 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | 49.71 ms | OfficeIMO.Excel | Win | 12690.9 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | 49.55 ms | OfficeIMO.Excel | Win | 9805.4 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | 54.90 ms | OfficeIMO.Excel | Win | 12588.3 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | 82.00 ms | OfficeIMO.Excel | Win | 7037.1 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | 51.91 ms | OfficeIMO.Excel | Win | 12584.9 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | 17.83 ms | OfficeIMO.Excel | Win | 6716.0 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | 21.77 ms | OfficeIMO.Excel | Win | 5790.4 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | 29.44 ms | OfficeIMO.Excel | Win | 7993.7 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | 31.35 ms | OfficeIMO.Excel | Win | 7173.0 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | 23.52 ms | OfficeIMO.Excel | Win | 7173.1 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | 16.68 ms | OfficeIMO.Excel | Win | 5964.2 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | 30.31 ms | OfficeIMO.Excel | Win | 7302.4 KB |  |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | 48.30 ms | OfficeIMO.Excel | Win | 12544.3 KB |  |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | 53.63 ms | OfficeIMO.Excel | Win | 13123.2 KB |  |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | 47.79 ms | OfficeIMO.Excel | Win | 9792.8 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | 49.48 ms | OfficeIMO.Excel | Win | 12684.6 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | 45.25 ms | LargeXlsx | Loss +6.6% | 12380.0 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | 48.83 ms | OfficeIMO.Excel | Win | 9663.7 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | 53.05 ms | OfficeIMO.Excel | Win | 12398.1 KB |  |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | 53.23 ms | OfficeIMO.Excel | Win | 12650.0 KB |  |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | 33.79 ms | OfficeIMO.Excel | Win | 9547.3 KB |  |
| 25000 | speed-comparison | write | Other | write-dictionary-objects-table-direct | 63.08 ms | OfficeIMO.Excel | Win | 14828.0 KB |  |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | 54.28 ms | LargeXlsx | Loss +18.3% | 7234.3 KB |  |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | 27.80 ms | LargeXlsx | Loss +55.9% | 11672.4 KB |  |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | 14.95 ms | LargeXlsx | Loss +16.9% | 7226.8 KB |  |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | 50.39 ms | LargeXlsx | Loss +25.2% | 15700.7 KB |  |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | 22.22 ms | OfficeIMO.Excel | Win | 7530.4 KB |  |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | 51.35 ms | Sylvan.Data.Excel | Loss +28.9% | 12666.6 KB |  |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | 42.54 ms | OfficeIMO.Excel | Win | 9484.6 KB |  |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | 59.47 ms | LargeXlsx | Loss +10.5% | 12754.6 KB |  |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | 17.94 ms | OfficeIMO.Excel | Win | 15409.2 KB |  |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | 24.55 ms | OfficeIMO.Excel | Win | 15019.9 KB |  |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | 19.01 ms | OfficeIMO.Excel | Win | 13643.7 KB |  |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | 18.68 ms | OfficeIMO.Excel | Win | 7184.9 KB |  |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | 50.10 ms | LargeXlsx | Loss +15.9% | 12584.3 KB |  |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | 52.88 ms | LargeXlsx | Loss +22.8% | 12585.9 KB |  |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | 69.43 ms | LargeXlsx | Loss +31.5% | 7029.4 KB |  |

## Full comparison table

| Row count | Artifact | Workload | Category | Scenario | Library | Mean | StdDev | StdErr | Ratio to OfficeIMO | Ratio to best | Alloc | Alloc ratio | Package | Package ratio | Outcome |
| ---: | --- | --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | Sylvan.Data.Excel | 6.41 ms | 0.86 ms | 0.50 ms | 0.66 | 1.00 | 362.3 KB | 0.15 |  |  | 33.9% faster than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | OfficeIMO.Excel | 9.70 ms | 1.62 ms | 0.93 ms | 1.00 | 1.51 | 2496.7 KB | 1.00 |  |  | Loss +51.3% |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | ExcelDataReader | 17.83 ms | 1.56 ms | 0.90 ms | 1.84 | 2.78 | 6887.4 KB | 2.76 |  |  | 83.9% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | MiniExcel | 25.26 ms | 3.14 ms | 1.81 ms | 2.60 | 3.94 | 21507.3 KB | 8.61 |  |  | 160.5% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | Sylvan.Data.Excel | 6.24 ms | 0.69 ms | 0.40 ms | 0.68 | 1.00 | 362.3 KB | 0.14 |  |  | 31.5% faster than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | OfficeIMO.Excel | 9.12 ms | 0.99 ms | 0.57 ms | 1.00 | 1.46 | 2575.0 KB | 1.00 |  |  | Loss +46.0% |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | ExcelDataReader | 19.66 ms | 2.46 ms | 1.42 ms | 2.16 | 3.15 | 6887.4 KB | 2.67 |  |  | 115.6% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | MiniExcel | 25.47 ms | 1.56 ms | 0.90 ms | 2.79 | 4.08 | 21507.3 KB | 8.35 |  |  | 179.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | append-plain-rows | LargeXlsx | 2.19 ms | 0.40 ms | 0.23 ms | 0.69 | 1.00 | 296.4 KB | 0.18 | 63.1 KB | 0.98 | 30.9% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | append-plain-rows | OfficeIMO.Excel | 3.17 ms | 0.67 ms | 0.39 ms | 1.00 | 1.45 | 1658.1 KB | 1.00 | 64.5 KB | 1.00 | Loss +44.7% |
| 2500 | package-profile | package | Package size | append-plain-rows | MiniExcel | 6.53 ms | 0.40 ms | 0.23 ms | 2.06 | 2.99 | 19710.7 KB | 11.89 | 68.1 KB | 1.06 | 106.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | append-plain-rows | ClosedXML | 23.06 ms | 3.51 ms | 2.02 ms | 7.29 | 10.54 | 11197.4 KB | 6.75 | 59.8 KB | 0.93 | 628.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | append-plain-rows | EPPlus | 38.87 ms | 7.06 ms | 4.08 ms | 12.28 | 17.77 | 14365.2 KB | 8.66 | 56.9 KB | 0.88 | 1127.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | autofit-existing | OfficeIMO.Excel | 43.26 ms | 3.91 ms | 2.26 ms | 1.00 | 1.00 | 13884.7 KB | 1.00 | 138.9 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | autofit-existing | EPPlus | 96.11 ms | 3.16 ms | 1.82 ms | 2.22 | 2.22 | 50712.0 KB | 3.65 | 115.0 KB | 0.83 | 122.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | autofit-existing | ClosedXML | 186.71 ms | 9.97 ms | 5.76 ms | 4.32 | 4.32 | 84561.3 KB | 6.09 | 121.0 KB | 0.87 | 331.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | large-shared-strings | OfficeIMO.Excel | 3.14 ms | 0.17 ms | 0.10 ms | 1.00 | 1.00 | 2112.6 KB | 1.00 | 55.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | large-shared-strings | MiniExcel | 6.06 ms | 0.17 ms | 0.10 ms | 1.93 | 1.93 | 21137.5 KB | 10.01 | 60.7 KB | 1.10 | 92.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | large-shared-strings | ClosedXML | 19.81 ms | 2.13 ms | 1.23 ms | 6.30 | 6.30 | 11299.2 KB | 5.35 | 50.3 KB | 0.91 | 530.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | large-shared-strings | EPPlus | 27.62 ms | 5.49 ms | 3.17 ms | 8.79 | 8.79 | 12804.4 KB | 6.06 | 48.1 KB | 0.87 | 779.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | LargeXlsx | 5.41 ms | 0.94 ms | 0.54 ms | 0.83 | 1.00 | 849.6 KB | 0.42 | 237.7 KB | 1.10 | 17.2% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | OfficeIMO.Excel | 6.54 ms | 0.75 ms | 0.44 ms | 1.00 | 1.21 | 2020.1 KB | 1.00 | 216.7 KB | 1.00 | Loss +20.8% |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | MiniExcel | 25.19 ms | 3.56 ms | 2.06 ms | 3.86 | 4.66 | 35910.9 KB | 17.78 | 235.3 KB | 1.09 | 285.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | ClosedXML | 124.41 ms | 7.84 ms | 4.53 ms | 19.04 | 22.99 | 71470.2 KB | 35.38 | 257.2 KB | 1.19 | 1803.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-bulk-report | OfficeIMO.Excel | 6.85 ms | 1.54 ms | 0.89 ms | 1.00 | 1.00 | 1498.6 KB | 1.00 | 138.9 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-bulk-report | MiniExcel | 10.85 ms | 1.87 ms | 1.08 ms | 1.58 | 1.58 | 26816.4 KB | 17.89 | 153.8 KB | 1.11 | 58.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-bulk-report | EPPlus | 70.33 ms | 1.46 ms | 0.84 ms | 10.27 | 10.27 | 47121.2 KB | 31.44 | 115.0 KB | 0.83 | 926.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-bulk-report | ClosedXML | 98.93 ms | 3.21 ms | 1.85 ms | 14.44 | 14.44 | 58337.3 KB | 38.93 | 121.0 KB | 0.87 | 1344.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellformula | OfficeIMO.Excel | 4.43 ms | 1.17 ms | 0.68 ms | 1.00 | 1.00 | 1173.1 KB | 1.00 | 66.6 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellformula | ClosedXML | 30.50 ms | 3.11 ms | 1.79 ms | 6.88 | 6.88 | 12039.8 KB | 10.26 | 70.6 KB | 1.06 | 588.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellformula | EPPlus | 64.98 ms | 9.75 ms | 5.63 ms | 14.67 | 14.67 | 18110.5 KB | 15.44 | 62.1 KB | 0.93 | 1366.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-empty-strings | OfficeIMO.Excel | 3.02 ms | 0.38 ms | 0.22 ms | 1.00 | 1.00 | 1446.7 KB | 1.00 | 44.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-empty-strings | ClosedXML | 16.30 ms | 0.54 ms | 0.31 ms | 5.40 | 5.40 | 9951.5 KB | 6.88 | 44.9 KB | 1.02 | 440.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-empty-strings | EPPlus | 22.71 ms | 1.13 ms | 0.65 ms | 7.52 | 7.52 | 11703.7 KB | 8.09 | 42.0 KB | 0.95 | 652.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-numbers | OfficeIMO.Excel | 2.37 ms | 0.38 ms | 0.22 ms | 1.00 | 1.00 | 939.5 KB | 1.00 | 47.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-numbers | ClosedXML | 15.40 ms | 0.42 ms | 0.24 ms | 6.50 | 6.50 | 9169.1 KB | 9.76 | 45.9 KB | 0.98 | 549.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-numbers | EPPlus | 27.83 ms | 1.65 ms | 0.95 ms | 11.74 | 11.74 | 12829.3 KB | 13.66 | 43.7 KB | 0.93 | 1074.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-mixed | OfficeIMO.Excel | 4.60 ms | 1.22 ms | 0.71 ms | 1.00 | 1.00 | 1424.1 KB | 1.00 | 61.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-object-mixed | ClosedXML | 29.75 ms | 5.95 ms | 3.43 ms | 6.46 | 6.46 | 11879.0 KB | 8.34 | 59.5 KB | 0.97 | 546.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-mixed | EPPlus | 39.47 ms | 6.21 ms | 3.58 ms | 8.57 | 8.57 | 15577.2 KB | 10.94 | 58.9 KB | 0.96 | 757.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse | OfficeIMO.Excel | 5.74 ms | 1.27 ms | 0.73 ms | 1.00 | 1.00 | 1263.7 KB | 1.00 | 62.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse | ClosedXML | 26.35 ms | 5.26 ms | 3.04 ms | 4.59 | 4.59 | 11288.3 KB | 8.93 | 52.5 KB | 0.85 | 358.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse | EPPlus | 49.07 ms | 9.77 ms | 5.64 ms | 8.54 | 8.54 | 14894.0 KB | 11.79 | 54.2 KB | 0.88 | 754.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 4.42 ms | 0.92 ms | 0.53 ms | 1.00 | 1.00 | 1271.8 KB | 1.00 | 62.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | ClosedXML | 37.10 ms | 12.10 ms | 6.99 ms | 8.40 | 8.40 | 11296.3 KB | 8.88 | 52.5 KB | 0.85 | 739.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | EPPlus | 67.81 ms | 30.34 ms | 17.52 ms | 15.35 | 15.35 | 14960.2 KB | 11.76 | 54.2 KB | 0.88 | 1434.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-scalars | OfficeIMO.Excel | 2.83 ms | 0.59 ms | 0.34 ms | 1.00 | 1.00 | 957.6 KB | 1.00 | 46.3 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-scalars | ClosedXML | 23.96 ms | 9.44 ms | 5.45 ms | 8.47 | 8.47 | 9013.2 KB | 9.41 | 45.4 KB | 0.98 | 747.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-scalars | EPPlus | 29.94 ms | 3.97 ms | 2.29 ms | 10.59 | 10.59 | 12761.5 KB | 13.33 | 42.4 KB | 0.91 | 959.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings | OfficeIMO.Excel | 4.70 ms | 2.15 ms | 1.24 ms | 1.00 | 1.00 | 2276.5 KB | 1.00 | 55.1 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-strings | ClosedXML | 15.79 ms | 1.46 ms | 0.84 ms | 3.36 | 3.36 | 11291.2 KB | 4.96 | 50.3 KB | 0.91 | 235.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings | EPPlus | 25.02 ms | 3.87 ms | 2.23 ms | 5.32 | 5.32 | 12730.2 KB | 5.59 | 48.1 KB | 0.87 | 432.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-distinct | OfficeIMO.Excel | 3.01 ms | 0.16 ms | 0.09 ms | 1.00 | 1.00 | 2198.8 KB | 1.00 | 51.8 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-distinct | ClosedXML | 23.31 ms | 2.08 ms | 1.20 ms | 7.75 | 7.75 | 13119.1 KB | 5.97 | 61.9 KB | 1.19 | 675.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-distinct | EPPlus | 30.89 ms | 2.86 ms | 1.65 ms | 10.27 | 10.27 | 13793.7 KB | 6.27 | 61.5 KB | 1.19 | 927.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-repeated | OfficeIMO.Excel | 2.40 ms | 0.21 ms | 0.12 ms | 1.00 | 1.00 | 1239.4 KB | 1.00 | 40.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-repeated | ClosedXML | 14.94 ms | 0.73 ms | 0.42 ms | 6.22 | 6.22 | 9218.5 KB | 7.44 | 38.8 KB | 0.97 | 522.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-repeated | EPPlus | 18.98 ms | 1.39 ms | 0.80 ms | 7.90 | 7.90 | 11265.7 KB | 9.09 | 34.8 KB | 0.87 | 690.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-temporal | OfficeIMO.Excel | 5.14 ms | 1.40 ms | 0.81 ms | 1.00 | 1.00 | 1255.3 KB | 1.00 | 63.3 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-temporal | ClosedXML | 29.19 ms | 1.88 ms | 1.08 ms | 5.67 | 5.67 | 9703.1 KB | 7.73 | 54.5 KB | 0.86 | 467.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-temporal | EPPlus | 38.17 ms | 3.81 ms | 2.20 ms | 7.42 | 7.42 | 14654.6 KB | 11.67 | 53.1 KB | 0.84 | 641.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 1.81 ms | 0.03 ms | 0.02 ms | 0.88 | 1.00 | 439.0 KB | 0.39 | 47.3 KB | 0.98 | 11.6% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 2.04 ms | 0.42 ms | 0.24 ms | 1.00 | 1.13 | 1126.8 KB | 1.00 | 48.2 KB | 1.00 | Loss +13.1% |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | ClosedXML | 17.38 ms | 2.41 ms | 1.39 ms | 8.50 | 9.62 | 10227.8 KB | 9.08 | 53.0 KB | 1.10 | 750.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | EPPlus | 27.36 ms | 2.88 ms | 1.66 ms | 13.38 | 15.14 | 12985.4 KB | 11.52 | 52.5 KB | 1.09 | 1238.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | LargeXlsx | 3.12 ms | 0.01 ms | 0.01 ms | 0.64 | 1.00 | 750.2 KB | 0.43 | 138.4 KB | 1.00 | 36.3% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 4.90 ms | 0.91 ms | 0.52 ms | 1.00 | 1.57 | 1745.4 KB | 1.00 | 138.0 KB | 1.00 | Loss +56.9% |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | MiniExcel | 9.49 ms | 0.61 ms | 0.35 ms | 1.94 | 3.04 | 23213.2 KB | 13.30 | 153.7 KB | 1.11 | 93.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | ClosedXML | 31.22 ms | 1.26 ms | 0.73 ms | 6.37 | 9.99 | 22213.3 KB | 12.73 | 120.1 KB | 0.87 | 536.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | EPPlus | 39.70 ms | 1.04 ms | 0.60 ms | 8.10 | 12.71 | 24626.9 KB | 14.11 | 114.1 KB | 0.83 | 709.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | Sylvan.Data.Excel | 3.50 ms | 0.25 ms | 0.14 ms | 0.87 | 1.00 | 750.7 KB | 0.53 | 78.5 KB | 0.57 | 12.5% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | OfficeIMO.Excel | 4.00 ms | 0.06 ms | 0.03 ms | 1.00 | 1.14 | 1427.2 KB | 1.00 | 138.0 KB | 1.00 | Loss +14.3% |
| 2500 | package-profile | package | Package size | write-datareader-plain | LargeXlsx | 4.00 ms | 0.60 ms | 0.35 ms | 1.00 | 1.14 | 1024.5 KB | 0.72 | 138.4 KB | 1.00 | Tie vs OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | MiniExcel | 7.69 ms | 0.44 ms | 0.25 ms | 1.92 | 2.20 | 23034.8 KB | 16.14 | 153.6 KB | 1.11 | 92.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | ClosedXML | 27.28 ms | 0.20 ms | 0.12 ms | 6.82 | 7.80 | 11573.0 KB | 8.11 | 120.1 KB | 0.87 | 582.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | EPPlus | 41.36 ms | 4.50 ms | 2.60 ms | 10.35 | 11.82 | 16579.3 KB | 11.62 | 114.9 KB | 0.83 | 934.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table | OfficeIMO.Excel | 4.24 ms | 0.24 ms | 0.14 ms | 1.00 | 1.00 | 1439.1 KB | 1.00 | 138.8 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-datareader-table | MiniExcel | 7.65 ms | 0.83 ms | 0.48 ms | 1.80 | 1.80 | 23035.0 KB | 16.01 | 153.6 KB | 1.11 | 80.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table | ClosedXML | 36.52 ms | 1.50 ms | 0.87 ms | 8.61 | 8.61 | 19001.0 KB | 13.20 | 120.9 KB | 0.87 | 761.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table | EPPlus | 38.20 ms | 3.35 ms | 1.93 ms | 9.01 | 9.01 | 16579.3 KB | 11.52 | 114.9 KB | 0.83 | 800.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | OfficeIMO.Excel | 4.37 ms | 0.15 ms | 0.08 ms | 1.00 | 1.00 | 1445.4 KB | 1.00 | 138.9 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | MiniExcel | 7.87 ms | 0.49 ms | 0.28 ms | 1.80 | 1.80 | 26638.2 KB | 18.43 | 153.8 KB | 1.11 | 80.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | EPPlus | 54.69 ms | 0.70 ms | 0.41 ms | 12.53 | 12.53 | 38271.3 KB | 26.48 | 115.1 KB | 0.83 | 1152.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | ClosedXML | 77.71 ms | 2.04 ms | 1.18 ms | 17.80 | 17.80 | 58352.5 KB | 40.37 | 121.0 KB | 0.87 | 1680.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | OfficeIMO.Excel | 5.18 ms | 0.23 ms | 0.13 ms | 1.00 | 1.00 | 1645.6 KB | 1.00 | 131.1 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | LargeXlsx | 5.40 ms | 0.48 ms | 0.28 ms | 1.04 | 1.04 | 1115.8 KB | 0.68 | 164.2 KB | 1.25 | 4.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | MiniExcel | 12.76 ms | 0.74 ms | 0.43 ms | 2.46 | 2.46 | 29737.8 KB | 18.07 | 180.5 KB | 1.38 | 146.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | EPPlus | 63.95 ms | 2.68 ms | 1.55 ms | 12.33 | 12.33 | 21822.9 KB | 13.26 | 144.5 KB | 1.10 | 1133.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | ClosedXML | 71.82 ms | 1.89 ms | 1.09 ms | 13.85 | 13.85 | 27402.1 KB | 16.65 | 159.4 KB | 1.22 | 1285.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | OfficeIMO.Excel | 5.80 ms | 0.53 ms | 0.30 ms | 1.00 | 1.00 | 2384.8 KB | 1.00 | 176.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | MiniExcel | 13.36 ms | 1.48 ms | 0.85 ms | 2.30 | 2.30 | 29737.8 KB | 12.47 | 180.5 KB | 1.03 | 130.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | EPPlus | 67.15 ms | 2.97 ms | 1.72 ms | 11.58 | 11.58 | 21822.9 KB | 9.15 | 144.5 KB | 0.82 | 1057.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | ClosedXML | 71.30 ms | 4.06 ms | 2.34 ms | 12.29 | 12.29 | 27402.3 KB | 11.49 | 159.4 KB | 0.91 | 1129.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables | OfficeIMO.Excel | 5.77 ms | 1.39 ms | 0.80 ms | 1.00 | 1.00 | 1570.8 KB | 1.00 | 138.9 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-dataset-tables | MiniExcel | 10.01 ms | 0.56 ms | 0.33 ms | 1.74 | 1.74 | 28691.3 KB | 18.27 | 156.4 KB | 1.13 | 73.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables | ClosedXML | 43.43 ms | 3.10 ms | 1.79 ms | 7.53 | 7.53 | 18868.6 KB | 12.01 | 123.4 KB | 0.89 | 653.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables | EPPlus | 44.07 ms | 0.57 ms | 0.33 ms | 7.64 | 7.64 | 18633.8 KB | 11.86 | 116.6 KB | 0.84 | 664.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | OfficeIMO.Excel | 5.70 ms | 0.70 ms | 0.40 ms | 1.00 | 1.00 | 1583.4 KB | 1.00 | 139.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | MiniExcel | 11.97 ms | 0.46 ms | 0.26 ms | 2.10 | 2.10 | 31789.4 KB | 20.08 | 156.6 KB | 1.13 | 109.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | EPPlus | 76.12 ms | 2.45 ms | 1.41 ms | 13.35 | 13.35 | 41385.5 KB | 26.14 | 116.9 KB | 0.84 | 1235.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | ClosedXML | 95.13 ms | 7.22 ms | 4.17 ms | 16.69 | 16.69 | 56699.7 KB | 35.81 | 123.7 KB | 0.89 | 1568.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-direct | OfficeIMO.Excel | 4.01 ms | 0.16 ms | 0.09 ms | 1.00 | 1.00 | 1412.9 KB | 1.00 | 138.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-datatable-direct | LargeXlsx | 4.14 ms | 1.06 ms | 0.61 ms | 1.03 | 1.03 | 1141.0 KB | 0.81 | 138.4 KB | 1.00 | 3.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-direct | MiniExcel | 8.63 ms | 0.93 ms | 0.54 ms | 2.15 | 2.15 | 23053.5 KB | 16.32 | 153.7 KB | 1.11 | 115.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-direct | ClosedXML | 30.31 ms | 3.68 ms | 2.13 ms | 7.57 | 7.57 | 11573.0 KB | 8.19 | 120.1 KB | 0.87 | 656.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-direct | EPPlus | 40.48 ms | 3.11 ms | 1.80 ms | 10.11 | 10.11 | 16579.3 KB | 11.73 | 114.9 KB | 0.83 | 910.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | OfficeIMO.Excel | 3.88 ms | 0.13 ms | 0.08 ms | 1.00 | 1.00 | 1424.9 KB | 1.00 | 138.8 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | MiniExcel | 7.80 ms | 0.27 ms | 0.16 ms | 2.01 | 2.01 | 23053.8 KB | 16.18 | 153.7 KB | 1.11 | 101.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | ClosedXML | 36.21 ms | 1.46 ms | 0.84 ms | 9.34 | 9.34 | 18999.8 KB | 13.33 | 120.9 KB | 0.87 | 833.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | EPPlus | 36.41 ms | 2.54 ms | 1.47 ms | 9.39 | 9.39 | 16579.3 KB | 11.64 | 114.9 KB | 0.83 | 838.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 5.36 ms | 0.21 ms | 0.12 ms | 1.00 | 1.00 | 1441.5 KB | 1.00 | 138.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | LargeXlsx | 12.29 ms | 13.01 ms | 7.51 ms | 2.29 | 2.29 | 758.3 KB | 0.53 | 138.4 KB | 1.00 | 129.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | MiniExcel | 12.79 ms | 2.18 ms | 1.26 ms | 2.39 | 2.39 | 23222.3 KB | 16.11 | 153.7 KB | 1.11 | 138.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | ClosedXML | 53.29 ms | 22.58 ms | 13.03 ms | 9.95 | 9.95 | 11581.0 KB | 8.03 | 120.1 KB | 0.87 | 894.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | EPPlus | 82.54 ms | 16.36 ms | 9.44 ms | 15.41 | 15.41 | 16646.4 KB | 11.55 | 114.9 KB | 0.83 | 1441.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 8.61 ms | 2.94 ms | 1.70 ms | 1.00 | 1.00 | 1450.8 KB | 1.00 | 138.1 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 128.74 ms | 13.73 ms | 7.93 ms | 14.95 | 14.95 | 38340.4 KB | 26.43 | 115.1 KB | 0.83 | 1395.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 165.08 ms | 34.55 ms | 19.94 ms | 19.17 | 19.17 | 50927.5 KB | 35.10 | 120.2 KB | 0.87 | 1817.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | LargeXlsx | 7.27 ms | 3.91 ms | 2.25 ms | 0.96 | 1.00 | 758.3 KB | 0.53 | 138.4 KB | 1.00 | 4.3% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | OfficeIMO.Excel | 7.60 ms | 2.03 ms | 1.17 ms | 1.00 | 1.04 | 1443.1 KB | 1.00 | 138.0 KB | 1.00 | Loss +4.5% |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | MiniExcel | 22.07 ms | 15.22 ms | 8.79 ms | 2.90 | 3.03 | 23222.3 KB | 16.09 | 153.7 KB | 1.11 | 190.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | ClosedXML | 46.94 ms | 0.99 ms | 0.57 ms | 6.18 | 6.45 | 11581.0 KB | 8.03 | 120.1 KB | 0.87 | 517.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | EPPlus | 68.39 ms | 18.00 ms | 10.39 ms | 9.00 | 9.40 | 16646.0 KB | 11.54 | 114.9 KB | 0.83 | 799.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 7.17 ms | 1.56 ms | 0.90 ms | 1.00 | 1.00 | 749.9 KB | 1.00 | 142.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 122.54 ms | 49.92 ms | 28.82 ms | 17.09 | 17.09 | 38340.4 KB | 51.13 | 115.1 KB | 0.81 | 1608.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 143.73 ms | 54.27 ms | 31.33 ms | 20.04 | 20.04 | 50927.5 KB | 67.92 | 120.2 KB | 0.84 | 1904.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 5.01 ms | 0.91 ms | 0.52 ms | 0.94 | 1.00 | 758.3 KB | 1.02 | 138.4 KB | 0.97 | 6.4% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 5.36 ms | 0.82 ms | 0.48 ms | 1.00 | 1.07 | 742.1 KB | 1.00 | 142.3 KB | 1.00 | Loss +6.9% |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | MiniExcel | 11.45 ms | 1.43 ms | 0.83 ms | 2.14 | 2.28 | 23222.3 KB | 31.29 | 153.7 KB | 1.08 | 113.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | ClosedXML | 40.06 ms | 1.61 ms | 0.93 ms | 7.48 | 7.99 | 11581.0 KB | 15.61 | 120.1 KB | 0.84 | 647.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | EPPlus | 49.66 ms | 0.66 ms | 0.38 ms | 9.27 | 9.91 | 16646.0 KB | 22.43 | 114.9 KB | 0.81 | 827.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | OfficeIMO.Excel | 7.21 ms | 0.41 ms | 0.24 ms | 1.00 | 1.00 | 1602.4 KB | 1.00 | 142.3 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | LargeXlsx | 7.68 ms | 5.59 ms | 3.23 ms | 1.06 | 1.06 | 758.3 KB | 0.47 | 138.4 KB | 0.97 | 6.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | MiniExcel | 12.53 ms | 0.63 ms | 0.36 ms | 1.74 | 1.74 | 23222.3 KB | 14.49 | 153.7 KB | 1.08 | 73.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | ClosedXML | 42.51 ms | 1.15 ms | 0.66 ms | 5.89 | 5.89 | 11581.0 KB | 7.23 | 120.1 KB | 0.84 | 489.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | EPPlus | 53.99 ms | 4.50 ms | 2.60 ms | 7.48 | 7.48 | 16646.0 KB | 10.39 | 114.9 KB | 0.81 | 648.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 6.90 ms | 0.71 ms | 0.41 ms | 1.00 | 1.00 | 1450.1 KB | 1.00 | 138.1 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 66.54 ms | 3.97 ms | 2.29 ms | 9.64 | 9.64 | 28540.6 KB | 19.68 | 120.2 KB | 0.87 | 864.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 73.83 ms | 6.36 ms | 3.67 ms | 10.70 | 10.70 | 27305.4 KB | 18.83 | 115.0 KB | 0.83 | 969.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | LargeXlsx | 6.13 ms | 0.63 ms | 0.37 ms | 0.95 | 1.00 | 802.5 KB | 0.74 | 182.6 KB | 1.00 | 5.1% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 6.46 ms | 0.77 ms | 0.44 ms | 1.00 | 1.05 | 1085.0 KB | 1.00 | 182.4 KB | 1.00 | Loss +5.4% |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | MiniExcel | 12.70 ms | 1.34 ms | 0.77 ms | 1.97 | 2.07 | 25190.5 KB | 23.22 | 194.0 KB | 1.06 | 96.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | ClosedXML | 53.74 ms | 3.77 ms | 2.17 ms | 8.32 | 8.77 | 16973.5 KB | 15.64 | 161.0 KB | 0.88 | 732.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | EPPlus | 63.75 ms | 2.50 ms | 1.44 ms | 9.87 | 10.40 | 20105.1 KB | 18.53 | 152.1 KB | 0.83 | 886.9% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | OfficeIMO.Excel | 30.04 ms | 1.81 ms | 1.05 ms | 1.00 | 1.00 | 13748.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | EPPlus | 71.87 ms | 2.82 ms | 1.63 ms | 2.39 | 2.39 | 50639.2 KB | 3.68 |  |  | 139.2% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | ClosedXML | 127.67 ms | 5.44 ms | 3.14 ms | 4.25 | 4.25 | 84737.4 KB | 6.16 |  |  | 325.0% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | EPPlus 4.5.3.3 | 136.44 ms |  |  | 4.54 | 4.54 |  |  |  |  | 354.2% slower than OfficeIMO |
| 2500 | speed-comparison | other | Object projection | build-object-datatable-dictionaries | OfficeIMO.Excel | 1.85 ms | 0.30 ms | 0.17 ms | 1.00 | 1.00 | 564.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Object projection | build-object-datatable-typed | OfficeIMO.Excel | 1.25 ms | 0.13 ms | 0.08 ms | 1.00 | 1.00 | 856.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-cells | OfficeIMO.Excel | 7.24 ms | 0.22 ms | 0.13 ms | 1.00 | 1.00 | 2539.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-cells | EPPlus | 23.22 ms | 0.10 ms | 0.06 ms | 3.21 | 3.21 | 18869.5 KB | 7.43 |  |  | 220.5% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-cells | ClosedXML | 31.92 ms | 0.21 ms | 0.12 ms | 4.41 | 4.41 | 20352.7 KB | 8.01 |  |  | 340.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | OfficeIMO.Excel | 4.89 ms | 0.10 ms | 0.05 ms | 1.00 | 1.00 | 679.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | EPPlus | 17.16 ms | 1.30 ms | 0.75 ms | 3.51 | 3.51 | 11822.6 KB | 17.40 |  |  | 250.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | ClosedXML | 34.43 ms | 1.92 ms | 1.11 ms | 7.03 | 7.03 | 18789.8 KB | 27.65 |  |  | 603.4% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-range | OfficeIMO.Excel | 7.12 ms | 0.85 ms | 0.49 ms | 1.00 | 1.00 | 2539.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-range | EPPlus | 24.22 ms | 1.90 ms | 1.10 ms | 3.40 | 3.40 | 18869.5 KB | 7.43 |  |  | 240.2% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-range | ClosedXML | 33.06 ms | 0.90 ms | 0.52 ms | 4.64 | 4.64 | 20351.5 KB | 8.01 |  |  | 364.4% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-top-range | OfficeIMO.Excel | 1.68 ms | 0.12 ms | 0.07 ms | 1.00 | 1.00 | 433.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-top-range | EPPlus | 18.03 ms | 2.12 ms | 1.23 ms | 10.75 | 10.75 | 11119.0 KB | 25.62 |  |  | 974.8% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-top-range | ClosedXML | 32.23 ms | 1.15 ms | 0.67 ms | 19.22 | 19.22 | 18701.6 KB | 43.10 |  |  | 1821.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | OfficeIMO.Excel | 3.23 ms | 0.24 ms | 0.14 ms | 1.00 | 1.00 | 777.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | EPPlus | 12.88 ms | 0.61 ms | 0.35 ms | 3.99 | 3.99 | 7707.5 KB | 9.91 |  |  | 299.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | EPPlus 4.5.3.3 | 16.76 ms |  |  | 5.20 | 5.20 |  |  |  |  | 419.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | ClosedXML | 18.37 ms | 1.43 ms | 0.83 ms | 5.69 | 5.69 | 8271.7 KB | 10.64 |  |  | 469.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | Sylvan.Data.Excel | 1.11 ms | 0.11 ms | 0.06 ms | 0.62 | 1.00 | 316.6 KB | 1.27 |  |  | 37.8% faster than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | ExcelDataReader | 1.45 ms | 0.01 ms | 0.01 ms | 0.81 | 1.31 | 4046.2 KB | 16.26 |  |  | 18.7% faster than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | OfficeIMO.Excel | 1.78 ms | 0.17 ms | 0.10 ms | 1.00 | 1.61 | 248.9 KB | 1.00 |  |  | Loss +60.7% |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | ClosedXML | 3.51 ms | 0.39 ms | 0.22 ms | 1.97 | 3.17 | 4393.1 KB | 17.65 |  |  | 97.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | MiniExcel | 10.88 ms | 0.42 ms | 0.24 ms | 6.11 | 9.82 | 46189.1 KB | 185.57 |  |  | 511.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | EPPlus 4.5.3.3 | 16.91 ms |  |  | 9.50 | 15.26 |  |  |  |  | 849.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | EPPlus | 30.45 ms | 0.55 ms | 0.32 ms | 17.10 | 27.48 | 43070.2 KB | 173.04 |  |  | 1610.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | Sylvan.Data.Excel | 1.00 ms | 0.02 ms | 0.01 ms | 0.38 | 1.00 | 316.6 KB | 1.27 |  |  | 62.3% faster than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | ExcelDataReader | 1.44 ms | 0.04 ms | 0.03 ms | 0.54 | 1.43 | 4046.2 KB | 16.25 |  |  | 46.0% faster than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | OfficeIMO.Excel | 2.66 ms | 0.08 ms | 0.05 ms | 1.00 | 2.65 | 249.0 KB | 1.00 |  |  | Loss +165.5% |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | ClosedXML | 3.18 ms | 0.12 ms | 0.07 ms | 1.19 | 3.17 | 4392.9 KB | 17.64 |  |  | 19.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | MiniExcel | 10.94 ms | 0.37 ms | 0.22 ms | 4.11 | 10.91 | 46189.1 KB | 185.50 |  |  | 310.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | EPPlus 4.5.3.3 | 16.00 ms |  |  | 6.01 | 15.95 |  |  |  |  | 500.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | EPPlus | 30.27 ms | 2.12 ms | 1.22 ms | 11.37 | 30.17 | 43070.2 KB | 172.97 |  |  | 1036.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | Sylvan.Data.Excel | 2.04 ms | 0.24 ms | 0.14 ms | 0.62 | 1.00 | 518.6 KB | 0.46 |  |  | 38.2% faster than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | OfficeIMO.Excel | 3.31 ms | 0.17 ms | 0.10 ms | 1.00 | 1.62 | 1133.6 KB | 1.00 |  |  | Loss +61.8% |
| 2500 | speed-comparison | read | Other | shared-string-read | ExcelDataReader | 4.90 ms | 0.44 ms | 0.25 ms | 1.48 | 2.39 | 2603.0 KB | 2.30 |  |  | 48.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | MiniExcel | 5.27 ms | 0.45 ms | 0.26 ms | 1.59 | 2.58 | 7524.7 KB | 6.64 |  |  | 59.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | EPPlus | 14.32 ms | 3.12 ms | 1.80 ms | 4.33 | 7.00 | 10371.6 KB | 9.15 |  |  | 332.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | ClosedXML | 16.75 ms | 1.28 ms | 0.74 ms | 5.06 | 8.19 | 9497.8 KB | 8.38 |  |  | 406.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | EPPlus 4.5.3.3 | 21.81 ms |  |  | 6.59 | 10.67 |  |  |  |  | 559.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | Sylvan.Data.Excel | 4.17 ms | 0.51 ms | 0.30 ms | 0.90 | 1.00 | 370.6 KB | 0.72 |  |  | 9.8% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | OfficeIMO.Excel | 4.62 ms | 0.16 ms | 0.09 ms | 1.00 | 1.11 | 516.7 KB | 1.00 |  |  | Loss +10.8% |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | ExcelDataReader | 10.65 ms | 0.02 ms | 0.01 ms | 2.30 | 2.55 | 6169.7 KB | 11.94 |  |  | 130.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | MiniExcel | 13.08 ms | 2.31 ms | 1.34 ms | 2.83 | 3.13 | 18611.5 KB | 36.02 |  |  | 182.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | EPPlus | 17.59 ms | 2.57 ms | 1.49 ms | 3.80 | 4.22 | 11141.6 KB | 21.56 |  |  | 280.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | ClosedXML | 37.40 ms | 4.47 ms | 2.58 ms | 8.09 | 8.96 | 18692.5 KB | 36.18 |  |  | 708.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | Sylvan.Data.Excel | 6.79 ms | 1.09 ms | 0.63 ms | 0.82 | 1.00 | 1954.6 KB | 0.54 |  |  | 18.4% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | OfficeIMO.Excel | 8.32 ms | 0.86 ms | 0.49 ms | 1.00 | 1.23 | 3596.9 KB | 1.00 |  |  | Loss +22.5% |
| 2500 | speed-comparison | read | Range and table read | read-datatable | ExcelDataReader | 13.16 ms | 1.37 ms | 0.79 ms | 1.58 | 1.94 | 7753.7 KB | 2.16 |  |  | 58.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | MiniExcel | 16.75 ms | 5.28 ms | 3.05 ms | 2.01 | 2.47 | 18216.3 KB | 5.06 |  |  | 101.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | EPPlus | 28.47 ms | 4.14 ms | 2.39 ms | 3.42 | 4.19 | 20451.1 KB | 5.69 |  |  | 242.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | ClosedXML | 37.78 ms | 2.53 ms | 1.46 ms | 4.54 | 5.56 | 21646.2 KB | 6.02 |  |  | 354.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | EPPlus 4.5.3.3 | 40.68 ms |  |  | 4.89 | 5.99 |  |  |  |  | 389.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | Sylvan.Data.Excel | 4.47 ms | 0.38 ms | 0.22 ms | 0.84 | 1.00 | 442.8 KB | 0.63 |  |  | 16.0% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | OfficeIMO.Excel | 5.32 ms | 0.36 ms | 0.21 ms | 1.00 | 1.19 | 699.1 KB | 1.00 |  |  | Loss +19.0% |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | MiniExcel | 11.04 ms | 1.00 ms | 0.58 ms | 2.08 | 2.47 | 15805.3 KB | 22.61 |  |  | 107.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | ExcelDataReader | 12.00 ms | 0.80 ms | 0.46 ms | 2.26 | 2.68 | 6161.7 KB | 8.81 |  |  | 125.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | EPPlus | 18.18 ms | 1.07 ms | 0.62 ms | 3.42 | 4.07 | 11822.6 KB | 16.91 |  |  | 241.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | ClosedXML | 33.55 ms | 1.58 ms | 0.91 ms | 6.31 | 7.51 | 18789.0 KB | 26.88 |  |  | 530.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | OfficeIMO.Excel | 12.34 ms | 1.04 ms | 0.60 ms | 1.00 | 1.00 | 2694.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Range and table read | read-range | Sylvan.Data.Excel | 14.60 ms | 1.31 ms | 0.76 ms | 1.18 | 1.18 | 370.3 KB | 0.14 |  |  | 18.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | MiniExcel | 20.39 ms | 3.18 ms | 1.83 ms | 1.65 | 1.65 | 18611.8 KB | 6.91 |  |  | 65.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | ExcelDataReader | 25.61 ms | 5.57 ms | 3.21 ms | 2.08 | 2.08 | 6169.3 KB | 2.29 |  |  | 107.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | EPPlus | 25.87 ms | 0.61 ms | 0.35 ms | 2.10 | 2.10 | 18867.1 KB | 7.00 |  |  | 109.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | EPPlus 4.5.3.3 | 39.46 ms |  |  | 3.20 | 3.20 |  |  |  |  | 219.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | ClosedXML | 77.04 ms | 16.48 ms | 9.52 ms | 6.24 | 6.24 | 20176.1 KB | 7.49 |  |  | 524.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | Sylvan.Data.Excel | 5.11 ms | 0.13 ms | 0.08 ms | 0.48 | 1.00 | 465.6 KB | 0.17 |  |  | 51.8% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | OfficeIMO.Excel | 10.59 ms | 3.10 ms | 1.79 ms | 1.00 | 2.07 | 2753.4 KB | 1.00 |  |  | Loss +107.5% |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | ExcelDataReader | 11.65 ms | 0.70 ms | 0.40 ms | 1.10 | 2.28 | 6169.7 KB | 2.24 |  |  | 10.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | MiniExcel | 16.78 ms | 5.60 ms | 3.23 ms | 1.58 | 3.29 | 18612.1 KB | 6.76 |  |  | 58.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | EPPlus | 25.55 ms | 1.94 ms | 1.12 ms | 2.41 | 5.00 | 18867.2 KB | 6.85 |  |  | 141.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | ClosedXML | 34.81 ms | 1.38 ms | 0.80 ms | 3.29 | 6.82 | 20098.0 KB | 7.30 |  |  | 228.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | Sylvan.Data.Excel | 0.43 ms | 0.01 ms | 0.00 ms | 0.24 | 1.00 | 367.3 KB | 0.84 |  |  | 76.3% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | MiniExcel | 0.77 ms | 0.02 ms | 0.01 ms | 0.42 | 1.78 | 959.8 KB | 2.18 |  |  | 57.9% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | OfficeIMO.Excel | 1.84 ms | 0.21 ms | 0.12 ms | 1.00 | 4.23 | 439.5 KB | 1.00 |  |  | Loss +322.8% |
| 2500 | speed-comparison | read | Range and table read | read-top-range | ExcelDataReader | 4.38 ms | 0.37 ms | 0.21 ms | 2.39 | 10.09 | 1984.4 KB | 4.51 |  |  | 138.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | EPPlus | 17.86 ms | 2.11 ms | 1.22 ms | 9.72 | 41.11 | 11116.6 KB | 25.29 |  |  | 872.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | EPPlus 4.5.3.3 | 32.52 ms |  |  | 17.70 | 74.84 |  |  |  |  | 1670.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | ClosedXML | 35.24 ms | 4.00 ms | 2.31 ms | 19.18 | 81.09 | 18692.2 KB | 42.53 |  |  | 1817.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | Sylvan.Data.Excel | 7.73 ms | 4.92 ms | 2.84 ms | 0.50 | 1.00 | 370.6 KB | 0.11 |  |  | 49.9% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | ExcelDataReader | 11.68 ms | 0.42 ms | 0.24 ms | 0.76 | 1.51 | 6169.7 KB | 1.81 |  |  | 24.3% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | OfficeIMO.Excel | 15.44 ms | 1.30 ms | 0.75 ms | 1.00 | 2.00 | 3416.5 KB | 1.00 |  |  | Loss +99.7% |
| 2500 | speed-comparison | read | Range and table read | read-used-range | MiniExcel | 20.96 ms | 0.69 ms | 0.40 ms | 1.36 | 2.71 | 18612.2 KB | 5.45 |  |  | 35.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | EPPlus | 27.86 ms | 2.38 ms | 1.37 ms | 1.80 | 3.60 | 18867.3 KB | 5.52 |  |  | 80.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | ClosedXML | 83.58 ms | 19.06 ms | 11.00 ms | 5.41 | 10.81 | 20177.3 KB | 5.91 |  |  | 441.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | Sylvan.Data.Excel | 4.32 ms | 0.75 ms | 0.43 ms | 0.93 | 1.00 | 370.6 KB | 0.71 |  |  | 7.5% faster than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | OfficeIMO.Excel | 4.67 ms | 0.22 ms | 0.13 ms | 1.00 | 1.08 | 520.1 KB | 1.00 |  |  | Loss +8.1% |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | ExcelDataReader | 11.11 ms | 0.77 ms | 0.45 ms | 2.38 | 2.57 | 6169.9 KB | 11.86 |  |  | 137.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | MiniExcel | 13.18 ms | 0.90 ms | 0.52 ms | 2.82 | 3.05 | 18611.6 KB | 35.78 |  |  | 182.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | EPPlus | 17.51 ms | 0.98 ms | 0.57 ms | 3.75 | 4.05 | 11141.6 KB | 21.42 |  |  | 275.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | ClosedXML | 37.30 ms | 5.78 ms | 3.34 ms | 7.99 | 8.63 | 18693.0 KB | 35.94 |  |  | 698.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | Sylvan.Data.Excel | 4.63 ms | 0.22 ms | 0.12 ms | 0.47 | 1.00 | 370.6 KB | 0.13 |  |  | 52.9% faster than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | OfficeIMO.Excel | 9.81 ms | 3.22 ms | 1.86 ms | 1.00 | 2.12 | 2773.8 KB | 1.00 |  |  | Loss +112.1% |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | ExcelDataReader | 11.27 ms | 0.44 ms | 0.25 ms | 1.15 | 2.44 | 6169.7 KB | 2.22 |  |  | 14.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | MiniExcel | 12.66 ms | 0.55 ms | 0.31 ms | 1.29 | 2.74 | 18612.1 KB | 6.71 |  |  | 29.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | EPPlus | 24.76 ms | 1.56 ms | 0.90 ms | 2.52 | 5.35 | 18867.1 KB | 6.80 |  |  | 152.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | EPPlus 4.5.3.3 | 33.86 ms |  |  | 3.45 | 7.32 |  |  |  |  | 245.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | ClosedXML | 34.78 ms | 3.13 ms | 1.81 ms | 3.54 | 7.52 | 20060.7 KB | 7.23 |  |  | 254.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | Sylvan.Data.Excel | 0.43 ms | 0.03 ms | 0.02 ms | 0.27 | 1.00 | 367.4 KB | 0.83 |  |  | 73.1% faster than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | MiniExcel | 0.79 ms | 0.03 ms | 0.02 ms | 0.49 | 1.82 | 959.8 KB | 2.17 |  |  | 51.2% faster than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | OfficeIMO.Excel | 1.61 ms | 0.05 ms | 0.03 ms | 1.00 | 3.72 | 442.9 KB | 1.00 |  |  | Loss +271.8% |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | ExcelDataReader | 4.58 ms | 0.64 ms | 0.37 ms | 2.83 | 10.54 | 1984.5 KB | 4.48 |  |  | 183.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | EPPlus | 17.20 ms | 0.78 ms | 0.45 ms | 10.66 | 39.63 | 11116.6 KB | 25.10 |  |  | 965.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | EPPlus 4.5.3.3 | 28.47 ms |  |  | 17.64 | 65.59 |  |  |  |  | 1663.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | ClosedXML | 34.78 ms | 3.43 ms | 1.98 ms | 21.55 | 80.13 | 18692.9 KB | 42.20 |  |  | 2054.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | Sylvan.Data.Excel | 0.41 ms | 0.01 ms | 0.00 ms | 0.26 | 1.00 | 367.4 KB | 0.83 |  |  | 73.7% faster than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | MiniExcel | 0.80 ms | 0.08 ms | 0.04 ms | 0.51 | 1.94 | 959.8 KB | 2.16 |  |  | 48.9% faster than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | OfficeIMO.Excel | 1.57 ms | 0.05 ms | 0.03 ms | 1.00 | 3.80 | 443.7 KB | 1.00 |  |  | Loss +279.7% |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | ExcelDataReader | 4.28 ms | 0.21 ms | 0.12 ms | 2.72 | 10.34 | 1984.5 KB | 4.47 |  |  | 172.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | EPPlus | 18.36 ms | 1.54 ms | 0.89 ms | 11.68 | 44.37 | 11116.6 KB | 25.06 |  |  | 1068.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | ClosedXML | 32.50 ms | 1.60 ms | 0.92 ms | 20.68 | 78.51 | 18691.8 KB | 42.13 |  |  | 1967.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | Sylvan.Data.Excel | 7.42 ms | 4.25 ms | 2.46 ms | 0.88 | 1.00 | 610.6 KB | 0.25 |  |  | 11.8% faster than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | OfficeIMO.Excel | 8.41 ms | 1.99 ms | 1.15 ms | 1.00 | 1.13 | 2444.5 KB | 1.00 |  |  | Loss +13.4% |
| 2500 | speed-comparison | read | Typed object read | read-objects | MiniExcel | 14.39 ms | 1.01 ms | 0.58 ms | 1.71 | 1.94 | 18423.8 KB | 7.54 |  |  | 71.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | ExcelDataReader | 14.50 ms | 4.32 ms | 2.49 ms | 1.72 | 1.95 | 6409.9 KB | 2.62 |  |  | 72.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | EPPlus | 29.62 ms | 4.24 ms | 2.45 ms | 3.52 | 3.99 | 20068.8 KB | 8.21 |  |  | 252.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | ClosedXML | 35.40 ms | 1.41 ms | 0.81 ms | 4.21 | 4.77 | 20255.8 KB | 8.29 |  |  | 320.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | EPPlus 4.5.3.3 | 37.12 ms |  |  | 4.41 | 5.00 |  |  |  |  | 341.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | Sylvan.Data.Excel | 5.04 ms | 0.31 ms | 0.18 ms | 0.78 | 1.00 | 546.4 KB | 0.22 |  |  | 22.2% faster than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | OfficeIMO.Excel | 6.48 ms | 0.18 ms | 0.11 ms | 1.00 | 1.29 | 2444.8 KB | 1.00 |  |  | Loss +28.6% |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | ExcelDataReader | 11.96 ms | 0.40 ms | 0.23 ms | 1.85 | 2.37 | 6345.6 KB | 2.60 |  |  | 84.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | MiniExcel | 13.91 ms | 0.35 ms | 0.20 ms | 2.15 | 2.76 | 18359.5 KB | 7.51 |  |  | 114.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | EPPlus | 25.91 ms | 1.83 ms | 1.06 ms | 4.00 | 5.14 | 20049.2 KB | 8.20 |  |  | 299.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | ClosedXML | 34.03 ms | 0.86 ms | 0.50 ms | 5.25 | 6.76 | 20235.2 KB | 8.28 |  |  | 425.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | EPPlus 4.5.3.3 | 35.63 ms |  |  | 5.50 | 7.07 |  |  |  |  | 449.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | OfficeIMO.Excel | 6.85 ms | 0.56 ms | 0.32 ms | 1.00 | 1.00 | 1445.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | MiniExcel | 11.95 ms | 0.47 ms | 0.27 ms | 1.74 | 1.74 | 26638.3 KB | 18.43 |  |  | 74.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | EPPlus | 71.54 ms | 1.74 ms | 1.01 ms | 10.44 | 10.44 | 38272.2 KB | 26.48 |  |  | 944.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | ClosedXML | 103.45 ms | 7.08 ms | 4.09 ms | 15.10 | 15.10 | 58353.7 KB | 40.37 |  |  | 1409.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | EPPlus 4.5.3.3 | 117.67 ms |  |  | 17.17 | 17.17 |  |  |  |  | 1617.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | OfficeIMO.Excel | 7.21 ms | 0.15 ms | 0.09 ms | 1.00 | 1.00 | 1584.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | MiniExcel | 20.69 ms | 2.67 ms | 1.54 ms | 2.87 | 2.87 | 32142.9 KB | 20.29 |  |  | 186.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | EPPlus | 180.90 ms | 7.94 ms | 4.58 ms | 25.08 | 25.08 | 43370.5 KB | 27.37 |  |  | 2408.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | ClosedXML | 253.05 ms | 39.90 ms | 23.03 ms | 35.08 | 35.08 | 56700.7 KB | 35.79 |  |  | 3408.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 5.75 ms | 0.39 ms | 0.22 ms | 1.00 | 1.00 | 1442.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 69.22 ms | 5.26 ms | 3.04 ms | 12.03 | 12.03 | 38271.3 KB | 26.53 |  |  | 1102.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 90.57 ms | 8.24 ms | 4.76 ms | 15.74 | 15.74 | 50919.7 KB | 35.29 |  |  | 1473.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.56 ms | 0.23 ms | 0.13 ms | 1.00 | 1.00 | 741.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 53.46 ms | 1.45 ms | 0.84 ms | 11.73 | 11.73 | 38271.3 KB | 51.59 |  |  | 1072.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 63.56 ms | 1.10 ms | 0.64 ms | 13.95 | 13.95 | 50919.3 KB | 68.64 |  |  | 1294.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 5.55 ms | 0.53 ms | 0.30 ms | 1.00 | 1.00 | 1442.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 57.93 ms | 1.81 ms | 1.04 ms | 10.44 | 10.44 | 28532.3 KB | 19.79 |  |  | 943.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 58.10 ms | 2.05 ms | 1.18 ms | 10.47 | 10.47 | 27236.1 KB | 18.89 |  |  | 947.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | OfficeIMO.Excel | 3.70 ms | 0.44 ms | 0.26 ms | 1.00 | 1.00 | 1446.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | ClosedXML | 15.47 ms | 1.00 ms | 0.58 ms | 4.18 | 4.18 | 9951.5 KB | 6.88 |  |  | 318.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | EPPlus | 19.34 ms | 3.64 ms | 2.10 ms | 5.23 | 5.23 | 11703.5 KB | 8.09 |  |  | 423.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | OfficeIMO.Excel | 3.43 ms | 0.12 ms | 0.07 ms | 1.00 | 1.00 | 939.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | ClosedXML | 16.32 ms | 0.70 ms | 0.41 ms | 4.76 | 4.76 | 9169.1 KB | 9.76 |  |  | 375.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | EPPlus 4.5.3.3 | 17.93 ms |  |  | 5.22 | 5.22 |  |  |  |  | 422.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | EPPlus | 25.89 ms | 1.49 ms | 0.86 ms | 7.54 | 7.54 | 12829.1 KB | 13.66 |  |  | 654.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | OfficeIMO.Excel | 3.98 ms | 0.74 ms | 0.43 ms | 1.00 | 1.00 | 1423.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | EPPlus 4.5.3.3 | 21.09 ms |  |  | 5.30 | 5.30 |  |  |  |  | 429.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | ClosedXML | 22.23 ms | 1.41 ms | 0.82 ms | 5.58 | 5.58 | 11879.0 KB | 8.34 |  |  | 458.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | EPPlus | 30.59 ms | 4.93 ms | 2.85 ms | 7.68 | 7.68 | 15577.1 KB | 10.94 |  |  | 668.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | OfficeIMO.Excel | 3.47 ms | 0.31 ms | 0.18 ms | 1.00 | 1.00 | 1263.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | ClosedXML | 19.16 ms | 2.58 ms | 1.49 ms | 5.52 | 5.52 | 11288.3 KB | 8.93 |  |  | 452.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | EPPlus | 31.69 ms | 3.18 ms | 1.84 ms | 9.14 | 9.14 | 14894.0 KB | 11.79 |  |  | 813.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 3.60 ms | 0.53 ms | 0.31 ms | 1.00 | 1.00 | 1263.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | ClosedXML | 21.15 ms | 0.80 ms | 0.46 ms | 5.88 | 5.88 | 11288.3 KB | 8.93 |  |  | 488.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | EPPlus | 31.36 ms | 1.62 ms | 0.93 ms | 8.72 | 8.72 | 14894.0 KB | 11.79 |  |  | 772.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | OfficeIMO.Excel | 2.89 ms | 0.35 ms | 0.20 ms | 1.00 | 1.00 | 957.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | ClosedXML | 16.03 ms | 0.53 ms | 0.31 ms | 5.56 | 5.56 | 9013.2 KB | 9.41 |  |  | 455.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | EPPlus 4.5.3.3 | 16.64 ms |  |  | 5.77 | 5.77 |  |  |  |  | 476.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | EPPlus | 26.96 ms | 1.31 ms | 0.76 ms | 9.34 | 9.34 | 12761.3 KB | 13.33 |  |  | 834.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | OfficeIMO.Excel | 3.74 ms | 0.36 ms | 0.21 ms | 1.00 | 1.00 | 1255.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | EPPlus 4.5.3.3 | 19.78 ms |  |  | 5.28 | 5.28 |  |  |  |  | 428.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | ClosedXML | 22.72 ms | 1.48 ms | 0.85 ms | 6.07 | 6.07 | 9703.1 KB | 7.73 |  |  | 506.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | EPPlus | 28.06 ms | 5.19 ms | 3.00 ms | 7.49 | 7.49 | 14654.5 KB | 11.68 |  |  | 649.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | OfficeIMO.Excel | 6.68 ms | 0.44 ms | 0.25 ms | 1.00 | 1.00 | 1569.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | MiniExcel | 22.39 ms | 0.58 ms | 0.34 ms | 3.35 | 3.35 | 29214.5 KB | 18.62 |  |  | 235.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | ClosedXML | 66.77 ms | 5.41 ms | 3.12 ms | 9.99 | 9.99 | 18905.4 KB | 12.05 |  |  | 898.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | EPPlus | 127.33 ms | 10.10 ms | 5.83 ms | 19.05 | 19.05 | 18347.9 KB | 11.69 |  |  | 1804.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | OfficeIMO.Excel | 8.30 ms | 1.31 ms | 0.76 ms | 1.00 | 1.00 | 2384.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | MiniExcel | 23.44 ms | 2.15 ms | 1.24 ms | 2.82 | 2.82 | 30501.4 KB | 12.79 |  |  | 182.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | ClosedXML | 79.00 ms | 6.50 ms | 3.75 ms | 9.52 | 9.52 | 27402.4 KB | 11.49 |  |  | 852.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | EPPlus | 84.27 ms | 15.58 ms | 8.99 ms | 10.16 | 10.16 | 22284.1 KB | 9.34 |  |  | 915.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | OfficeIMO.Excel | 6.48 ms | 1.22 ms | 0.70 ms | 1.00 | 1.00 | 1572.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | MiniExcel | 16.96 ms | 3.08 ms | 1.78 ms | 2.62 | 2.62 | 28691.2 KB | 18.25 |  |  | 162.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | EPPlus 4.5.3.3 | 33.43 ms |  |  | 5.16 | 5.16 |  |  |  |  | 416.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | ClosedXML | 77.04 ms | 11.95 ms | 6.90 ms | 11.90 | 11.90 | 18870.3 KB | 12.00 |  |  | 1089.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | EPPlus | 79.97 ms | 13.93 ms | 8.04 ms | 12.35 | 12.35 | 19364.2 KB | 12.32 |  |  | 1134.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | OfficeIMO.Excel | 5.97 ms | 0.85 ms | 0.49 ms | 1.00 | 1.00 | 1439.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | MiniExcel | 12.52 ms | 1.77 ms | 1.02 ms | 2.10 | 2.10 | 23035.1 KB | 16.01 |  |  | 109.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | EPPlus 4.5.3.3 | 34.35 ms |  |  | 5.75 | 5.75 |  |  |  |  | 475.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | ClosedXML | 47.05 ms | 2.53 ms | 1.46 ms | 7.88 | 7.88 | 19001.4 KB | 13.20 |  |  | 688.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | EPPlus | 49.34 ms | 6.78 ms | 3.92 ms | 8.27 | 8.27 | 16580.4 KB | 11.52 |  |  | 726.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | OfficeIMO.Excel | 5.01 ms | 0.53 ms | 0.30 ms | 1.00 | 1.00 | 1412.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | LargeXlsx | 8.88 ms | 1.15 ms | 0.66 ms | 1.77 | 1.77 | 1141.0 KB | 0.81 |  |  | 77.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | MiniExcel | 12.64 ms | 1.13 ms | 0.65 ms | 2.52 | 2.52 | 23053.6 KB | 16.32 |  |  | 152.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | EPPlus 4.5.3.3 | 32.64 ms |  |  | 6.51 | 6.51 |  |  |  |  | 551.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | ClosedXML | 33.45 ms | 0.88 ms | 0.51 ms | 6.67 | 6.67 | 11573.0 KB | 8.19 |  |  | 567.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | EPPlus | 41.43 ms | 1.53 ms | 0.88 ms | 8.26 | 8.26 | 16581.8 KB | 11.74 |  |  | 726.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | OfficeIMO.Excel | 5.74 ms | 0.62 ms | 0.36 ms | 1.00 | 1.00 | 1151.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | MiniExcel | 10.49 ms | 0.75 ms | 0.43 ms | 1.83 | 1.83 | 22780.4 KB | 19.78 |  |  | 82.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | EPPlus | 42.38 ms | 2.43 ms | 1.40 ms | 7.38 | 7.38 | 16306.9 KB | 14.16 |  |  | 638.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | ClosedXML | 45.32 ms | 2.12 ms | 1.22 ms | 7.89 | 7.89 | 18726.0 KB | 16.26 |  |  | 689.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | OfficeIMO.Excel | 5.71 ms | 0.94 ms | 0.54 ms | 1.00 | 1.00 | 1424.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | MiniExcel | 12.63 ms | 0.95 ms | 0.55 ms | 2.21 | 2.21 | 23053.8 KB | 16.18 |  |  | 121.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | EPPlus 4.5.3.3 | 34.34 ms |  |  | 6.01 | 6.01 |  |  |  |  | 501.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | EPPlus | 41.67 ms | 1.14 ms | 0.66 ms | 7.29 | 7.29 | 16580.6 KB | 11.64 |  |  | 629.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | ClosedXML | 48.05 ms | 6.05 ms | 3.49 ms | 8.41 | 8.41 | 18999.7 KB | 13.33 |  |  | 741.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | OfficeIMO.Excel | 5.87 ms | 0.56 ms | 0.33 ms | 1.00 | 1.00 | 1500.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | MiniExcel | 16.04 ms | 2.03 ms | 1.17 ms | 2.73 | 2.73 | 26816.0 KB | 17.88 |  |  | 173.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | EPPlus 4.5.3.3 | 103.42 ms |  |  | 17.62 | 17.62 |  |  |  |  | 1661.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | EPPlus | 126.49 ms | 8.34 ms | 4.81 ms | 21.55 | 21.55 | 49086.2 KB | 32.72 |  |  | 2054.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | ClosedXML | 218.79 ms | 47.56 ms | 27.46 ms | 37.27 | 37.27 | 58343.1 KB | 38.89 |  |  | 3627.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | OfficeIMO.Excel | 3.98 ms | 0.20 ms | 0.12 ms | 1.00 | 1.00 | 1164.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | EPPlus 4.5.3.3 | 23.31 ms |  |  | 5.86 | 5.86 |  |  |  |  | 485.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | ClosedXML | 27.80 ms | 4.31 ms | 2.49 ms | 6.99 | 6.99 | 12031.2 KB | 10.33 |  |  | 598.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | EPPlus | 49.06 ms | 3.16 ms | 1.83 ms | 12.33 | 12.33 | 18036.4 KB | 15.49 |  |  | 1132.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-dictionary-objects-table-direct | OfficeIMO.Excel | 7.62 ms | 0.24 ms | 0.14 ms | 1.00 | 1.00 | 1715.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | LargeXlsx | 4.17 ms | 0.32 ms | 0.18 ms | 0.93 | 1.00 | 794.5 KB | 0.74 |  |  | 7.4% faster than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 4.50 ms | 0.07 ms | 0.04 ms | 1.00 | 1.08 | 1077.0 KB | 1.00 |  |  | Loss +8.0% |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | MiniExcel | 8.69 ms | 0.50 ms | 0.29 ms | 1.93 | 2.08 | 25181.5 KB | 23.38 |  |  | 92.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | ClosedXML | 34.65 ms | 1.34 ms | 0.78 ms | 7.69 | 8.31 | 16965.4 KB | 15.75 |  |  | 669.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | EPPlus | 44.33 ms | 0.33 ms | 0.19 ms | 9.85 | 10.64 | 20030.7 KB | 18.60 |  |  | 884.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | LargeXlsx | 1.62 ms | 0.39 ms | 0.23 ms | 0.64 | 1.00 | 288.4 KB | 0.17 |  |  | 35.8% faster than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | OfficeIMO.Excel | 2.53 ms | 0.17 ms | 0.10 ms | 1.00 | 1.56 | 1650.0 KB | 1.00 |  |  | Loss +55.7% |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | MiniExcel | 4.18 ms | 0.20 ms | 0.12 ms | 1.66 | 2.58 | 19701.8 KB | 11.94 |  |  | 65.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | ClosedXML | 13.91 ms | 1.22 ms | 0.70 ms | 5.51 | 8.57 | 11189.4 KB | 6.78 |  |  | 450.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | EPPlus | 20.81 ms | 0.50 ms | 0.29 ms | 8.24 | 12.82 | 14283.5 KB | 8.66 |  |  | 723.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | EPPlus 4.5.3.3 | 21.10 ms |  |  | 8.35 | 13.00 |  |  |  |  | 735.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 1.83 ms | 0.21 ms | 0.12 ms | 0.39 | 1.00 | 439.0 KB | 0.39 |  |  | 60.6% faster than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 4.64 ms | 3.34 ms | 1.93 ms | 1.00 | 2.54 | 1126.8 KB | 1.00 |  |  | Loss +153.8% |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | ClosedXML | 17.16 ms | 0.41 ms | 0.24 ms | 3.70 | 9.39 | 10227.8 KB | 9.08 |  |  | 270.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | EPPlus | 27.88 ms | 2.69 ms | 1.56 ms | 6.01 | 15.26 | 12985.2 KB | 11.52 |  |  | 501.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | LargeXlsx | 6.11 ms | 2.70 ms | 1.56 ms | 0.97 | 1.00 | 750.2 KB | 0.43 |  |  | 3.5% faster than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 6.33 ms | 0.18 ms | 0.10 ms | 1.00 | 1.04 | 1745.4 KB | 1.00 |  |  | Loss +3.6% |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | MiniExcel | 12.78 ms | 1.28 ms | 0.74 ms | 2.02 | 2.09 | 23212.9 KB | 13.30 |  |  | 102.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 41.29 ms |  |  | 6.53 | 6.76 |  |  |  |  | 552.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | ClosedXML | 45.37 ms | 2.46 ms | 1.42 ms | 7.17 | 7.43 | 22213.3 KB | 12.73 |  |  | 617.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | EPPlus | 54.31 ms | 1.94 ms | 1.12 ms | 8.58 | 8.89 | 24626.7 KB | 14.11 |  |  | 758.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | OfficeIMO.Excel | 3.58 ms | 0.52 ms | 0.30 ms | 1.00 | 1.00 | 1158.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | ClosedXML | 21.00 ms | 0.68 ms | 0.40 ms | 5.87 | 5.87 | 11288.3 KB | 9.75 |  |  | 487.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | EPPlus | 29.21 ms | 1.69 ms | 0.98 ms | 8.17 | 8.17 | 14893.8 KB | 12.86 |  |  | 716.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | Sylvan.Data.Excel | 4.70 ms | 0.53 ms | 0.31 ms | 0.73 | 1.00 | 750.5 KB | 0.53 |  |  | 27.0% faster than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | OfficeIMO.Excel | 6.44 ms | 0.11 ms | 0.07 ms | 1.00 | 1.37 | 1427.2 KB | 1.00 |  |  | Loss +37.0% |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | MiniExcel | 11.30 ms | 0.29 ms | 0.17 ms | 1.76 | 2.41 | 23034.8 KB | 16.14 |  |  | 75.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | LargeXlsx | 11.46 ms | 0.35 ms | 0.20 ms | 1.78 | 2.44 | 1024.5 KB | 0.72 |  |  | 78.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | EPPlus 4.5.3.3 | 37.95 ms |  |  | 5.90 | 8.08 |  |  |  |  | 489.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | ClosedXML | 41.06 ms | 0.33 ms | 0.19 ms | 6.38 | 8.74 | 11573.0 KB | 8.11 |  |  | 538.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | EPPlus | 49.16 ms | 3.90 ms | 2.25 ms | 7.64 | 10.47 | 16579.2 KB | 11.62 |  |  | 663.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | OfficeIMO.Excel | 7.75 ms | 0.23 ms | 0.13 ms | 1.00 | 1.00 | 1645.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | LargeXlsx | 12.00 ms | 1.30 ms | 0.75 ms | 1.55 | 1.55 | 1115.8 KB | 0.68 |  |  | 54.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | MiniExcel | 17.73 ms | 3.82 ms | 2.21 ms | 2.29 | 2.29 | 29992.4 KB | 18.23 |  |  | 128.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | EPPlus | 69.02 ms | 2.63 ms | 1.52 ms | 8.91 | 8.91 | 21825.9 KB | 13.26 |  |  | 790.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | ClosedXML | 81.29 ms | 5.15 ms | 2.97 ms | 10.49 | 10.49 | 27402.9 KB | 16.65 |  |  | 949.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | LargeXlsx | 6.72 ms | 0.17 ms | 0.10 ms | 0.88 | 1.00 | 849.6 KB | 0.42 |  |  | 11.7% faster than OfficeIMO |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | OfficeIMO.Excel | 7.61 ms | 0.19 ms | 0.11 ms | 1.00 | 1.13 | 2020.1 KB | 1.00 |  |  | Loss +13.2% |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | MiniExcel | 23.39 ms | 1.45 ms | 0.83 ms | 3.07 | 3.48 | 35909.8 KB | 17.78 |  |  | 207.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | ClosedXML | 117.38 ms | 6.62 ms | 3.82 ms | 15.42 | 17.46 | 71470.2 KB | 35.38 |  |  | 1441.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | OfficeIMO.Excel | 1.94 ms | 0.03 ms | 0.02 ms | 1.00 | 1.00 | 2104.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | MiniExcel | 3.84 ms | 0.15 ms | 0.09 ms | 1.98 | 1.98 | 21128.5 KB | 10.04 |  |  | 98.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | ClosedXML | 11.64 ms | 1.10 ms | 0.63 ms | 6.01 | 6.01 | 11291.2 KB | 5.37 |  |  | 500.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | EPPlus | 17.12 ms | 2.17 ms | 1.25 ms | 8.83 | 8.83 | 12730.2 KB | 6.05 |  |  | 783.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | EPPlus 4.5.3.3 | 17.82 ms |  |  | 9.19 | 9.19 |  |  |  |  | 819.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | OfficeIMO.Excel | 3.88 ms | 0.59 ms | 0.34 ms | 1.00 | 1.00 | 2276.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | ClosedXML | 17.35 ms | 0.69 ms | 0.40 ms | 4.47 | 4.47 | 11291.2 KB | 4.96 |  |  | 346.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | EPPlus 4.5.3.3 | 17.63 ms |  |  | 4.54 | 4.54 |  |  |  |  | 354.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | EPPlus | 22.25 ms | 1.15 ms | 0.66 ms | 5.73 | 5.73 | 12729.9 KB | 5.59 |  |  | 473.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | OfficeIMO.Excel | 2.83 ms | 0.53 ms | 0.31 ms | 1.00 | 1.00 | 2198.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | ClosedXML | 18.93 ms | 3.96 ms | 2.28 ms | 6.70 | 6.70 | 13119.1 KB | 5.97 |  |  | 569.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | EPPlus | 23.95 ms | 2.38 ms | 1.37 ms | 8.48 | 8.48 | 13793.5 KB | 6.27 |  |  | 747.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | OfficeIMO.Excel | 2.89 ms | 0.24 ms | 0.14 ms | 1.00 | 1.00 | 1239.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | ClosedXML | 14.53 ms | 2.56 ms | 1.48 ms | 5.03 | 5.03 | 9218.5 KB | 7.44 |  |  | 402.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | EPPlus | 20.54 ms | 3.19 ms | 1.84 ms | 7.11 | 7.11 | 11265.5 KB | 9.09 |  |  | 610.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | LargeXlsx | 2.99 ms | 0.05 ms | 0.03 ms | 0.84 | 1.00 | 750.2 KB | 0.52 |  |  | 16.4% faster than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 3.58 ms | 0.02 ms | 0.01 ms | 1.00 | 1.20 | 1433.4 KB | 1.00 |  |  | Loss +19.6% |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | MiniExcel | 7.49 ms | 0.08 ms | 0.05 ms | 2.09 | 2.50 | 23213.4 KB | 16.19 |  |  | 109.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | ClosedXML | 27.39 ms | 0.56 ms | 0.33 ms | 7.65 | 9.15 | 11573.0 KB | 8.07 |  |  | 664.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | EPPlus | 35.04 ms | 2.08 ms | 1.20 ms | 9.79 | 11.70 | 16579.3 KB | 11.57 |  |  | 878.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 38.25 ms |  |  | 10.68 | 12.77 |  |  |  |  | 968.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | LargeXlsx | 4.52 ms | 0.17 ms | 0.10 ms | 0.85 | 1.00 | 750.2 KB | 0.52 |  |  | 14.7% faster than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | OfficeIMO.Excel | 5.30 ms | 0.03 ms | 0.02 ms | 1.00 | 1.17 | 1435.1 KB | 1.00 |  |  | Loss +17.2% |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | MiniExcel | 13.19 ms | 2.32 ms | 1.34 ms | 2.49 | 2.92 | 23213.4 KB | 16.18 |  |  | 149.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | EPPlus 4.5.3.3 | 39.78 ms |  |  | 7.51 | 8.80 |  |  |  |  | 651.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | ClosedXML | 40.29 ms | 1.00 ms | 0.58 ms | 7.61 | 8.91 | 11573.0 KB | 8.06 |  |  | 660.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | EPPlus | 50.40 ms | 0.89 ms | 0.51 ms | 9.51 | 11.15 | 16579.2 KB | 11.55 |  |  | 851.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 4.00 ms | 0.66 ms | 0.38 ms | 0.74 | 1.00 | 750.2 KB | 1.02 |  |  | 25.8% faster than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 5.39 ms | 0.99 ms | 0.57 ms | 1.00 | 1.35 | 734.1 KB | 1.00 |  |  | Loss +34.7% |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | MiniExcel | 9.00 ms | 1.14 ms | 0.66 ms | 1.67 | 2.25 | 23213.4 KB | 31.62 |  |  | 67.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | ClosedXML | 33.13 ms | 3.63 ms | 2.09 ms | 6.15 | 8.29 | 11573.0 KB | 15.76 |  |  | 515.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | EPPlus | 40.21 ms | 2.20 ms | 1.27 ms | 7.47 | 10.06 | 16579.3 KB | 22.58 |  |  | 646.6% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | Sylvan.Data.Excel | 55.25 ms | 7.37 ms | 4.25 ms | 0.80 | 1.00 | 394.1 KB | 0.02 |  |  | 19.6% faster than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | OfficeIMO.Excel | 68.71 ms | 9.09 ms | 5.25 ms | 1.00 | 1.24 | 23699.7 KB | 1.00 |  |  | Loss +24.4% |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | ExcelDataReader | 163.40 ms | 6.67 ms | 3.85 ms | 2.38 | 2.96 | 69517.4 KB | 2.93 |  |  | 137.8% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | MiniExcel | 190.66 ms | 5.56 ms | 3.21 ms | 2.77 | 3.45 | 215349.0 KB | 9.09 |  |  | 177.5% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | Sylvan.Data.Excel | 57.91 ms | 6.73 ms | 3.88 ms | 0.82 | 1.00 | 394.1 KB | 0.02 |  |  | 18.0% faster than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | OfficeIMO.Excel | 70.63 ms | 4.06 ms | 2.34 ms | 1.00 | 1.22 | 24482.0 KB | 1.00 |  |  | Loss +22.0% |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | ExcelDataReader | 158.86 ms | 7.48 ms | 4.32 ms | 2.25 | 2.74 | 69517.4 KB | 2.84 |  |  | 124.9% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | MiniExcel | 214.64 ms | 5.89 ms | 3.40 ms | 3.04 | 3.71 | 215349.0 KB | 8.80 |  |  | 203.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | append-plain-rows | LargeXlsx | 16.04 ms | 0.50 ms | 0.29 ms | 0.66 | 1.00 | 2763.0 KB | 0.24 | 605.0 KB | 0.97 | 34.4% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | append-plain-rows | OfficeIMO.Excel | 24.44 ms | 0.32 ms | 0.19 ms | 1.00 | 1.52 | 11664.3 KB | 1.00 | 622.5 KB | 1.00 | Loss +52.4% |
| 25000 | package-profile | package | Package size | append-plain-rows | MiniExcel | 44.05 ms | 2.12 ms | 1.22 ms | 1.80 | 2.75 | 58233.0 KB | 4.99 | 642.3 KB | 1.03 | 80.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | append-plain-rows | ClosedXML | 184.98 ms | 3.85 ms | 2.23 ms | 7.57 | 11.53 | 104225.1 KB | 8.94 | 540.6 KB | 0.87 | 656.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | append-plain-rows | EPPlus | 201.04 ms | 9.32 ms | 5.38 ms | 8.23 | 12.53 | 100275.4 KB | 8.60 | 525.6 KB | 0.84 | 722.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | autofit-existing | OfficeIMO.Excel | 474.61 ms | 12.77 ms | 7.37 ms | 1.00 | 1.00 | 136016.3 KB | 1.00 | 1385.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | autofit-existing | EPPlus | 579.77 ms | 13.08 ms | 7.55 ms | 1.22 | 1.22 | 250878.4 KB | 1.84 | 1091.0 KB | 0.79 | 22.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | autofit-existing | ClosedXML | 1855.16 ms | 30.11 ms | 17.38 ms | 3.91 | 3.91 | 829579.5 KB | 6.10 | 1140.9 KB | 0.82 | 290.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | large-shared-strings | OfficeIMO.Excel | 17.93 ms | 0.55 ms | 0.32 ms | 1.00 | 1.00 | 15409.2 KB | 1.00 | 529.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | large-shared-strings | MiniExcel | 37.33 ms | 2.61 ms | 1.51 ms | 2.08 | 2.08 | 73751.2 KB | 4.79 | 581.0 KB | 1.10 | 108.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | large-shared-strings | ClosedXML | 136.10 ms | 6.24 ms | 3.60 ms | 7.59 | 7.59 | 104233.3 KB | 6.76 | 460.1 KB | 0.87 | 659.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | large-shared-strings | EPPlus | 160.97 ms | 7.22 ms | 4.17 ms | 8.98 | 8.98 | 84343.7 KB | 5.47 | 444.7 KB | 0.84 | 797.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | LargeXlsx | 58.30 ms | 0.65 ms | 0.37 ms | 0.94 | 1.00 | 10787.2 KB | 0.85 | 2444.6 KB | 1.10 | 6.0% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | OfficeIMO.Excel | 61.99 ms | 1.88 ms | 1.09 ms | 1.00 | 1.06 | 12754.6 KB | 1.00 | 2228.8 KB | 1.00 | Loss +6.3% |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | MiniExcel | 206.70 ms | 3.59 ms | 2.07 ms | 3.33 | 3.55 | 226867.4 KB | 17.79 | 2410.6 KB | 1.08 | 233.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | ClosedXML | 1270.96 ms | 23.53 ms | 13.59 ms | 20.50 | 21.80 | 759812.1 KB | 59.57 | 2581.2 KB | 1.16 | 1950.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-bulk-report | OfficeIMO.Excel | 56.75 ms | 0.67 ms | 0.38 ms | 1.00 | 1.00 | 12647.3 KB | 1.00 | 1385.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-bulk-report | MiniExcel | 102.61 ms | 3.82 ms | 2.21 ms | 1.81 | 1.81 | 125541.5 KB | 9.93 | 1521.1 KB | 1.10 | 80.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-bulk-report | EPPlus | 493.29 ms | 1.32 ms | 0.76 ms | 8.69 | 8.69 | 254887.6 KB | 20.15 | 1091.0 KB | 0.79 | 769.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-bulk-report | ClosedXML | 1214.25 ms | 25.02 ms | 14.44 ms | 21.40 | 21.40 | 565943.0 KB | 44.75 | 1140.9 KB | 0.82 | 2039.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellformula | OfficeIMO.Excel | 28.86 ms | 1.44 ms | 0.83 ms | 1.00 | 1.00 | 9541.6 KB | 1.00 | 670.3 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellformula | ClosedXML | 244.63 ms | 2.48 ms | 1.43 ms | 8.48 | 8.48 | 113844.9 KB | 11.93 | 643.2 KB | 0.96 | 747.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellformula | EPPlus | 353.87 ms | 3.02 ms | 1.75 ms | 12.26 | 12.26 | 140665.9 KB | 14.74 | 593.9 KB | 0.89 | 1126.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-empty-strings | OfficeIMO.Excel | 17.57 ms | 0.33 ms | 0.19 ms | 1.00 | 1.00 | 6716.0 KB | 1.00 | 451.4 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-empty-strings | ClosedXML | 153.52 ms | 7.07 ms | 4.08 ms | 8.74 | 8.74 | 92894.1 KB | 13.83 | 398.1 KB | 0.88 | 774.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-empty-strings | EPPlus | 160.59 ms | 1.52 ms | 0.88 ms | 9.14 | 9.14 | 74425.8 KB | 11.08 | 390.6 KB | 0.87 | 814.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-numbers | OfficeIMO.Excel | 22.31 ms | 0.83 ms | 0.48 ms | 1.00 | 1.00 | 5790.4 KB | 1.00 | 462.6 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-numbers | ClosedXML | 148.09 ms | 2.80 ms | 1.62 ms | 6.64 | 6.64 | 84198.7 KB | 14.54 | 411.4 KB | 0.89 | 563.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-numbers | EPPlus | 198.06 ms | 7.03 ms | 4.06 ms | 8.88 | 8.88 | 86279.5 KB | 14.90 | 406.5 KB | 0.88 | 787.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-mixed | OfficeIMO.Excel | 27.22 ms | 0.92 ms | 0.53 ms | 1.00 | 1.00 | 7993.7 KB | 1.00 | 585.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-object-mixed | EPPlus | 213.77 ms | 5.39 ms | 3.11 ms | 7.85 | 7.85 | 113162.8 KB | 14.16 | 544.3 KB | 0.93 | 685.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-mixed | ClosedXML | 219.55 ms | 4.48 ms | 2.59 ms | 8.07 | 8.07 | 111110.6 KB | 13.90 | 532.9 KB | 0.91 | 706.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse | OfficeIMO.Excel | 29.93 ms | 2.89 ms | 1.67 ms | 1.00 | 1.00 | 7173.0 KB | 1.00 | 607.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse | ClosedXML | 193.29 ms | 1.12 ms | 0.65 ms | 6.46 | 6.46 | 105215.9 KB | 14.67 | 468.0 KB | 0.77 | 545.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse | EPPlus | 233.63 ms | 1.33 ms | 0.77 ms | 7.81 | 7.81 | 106250.4 KB | 14.81 | 494.4 KB | 0.81 | 680.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 23.41 ms | 1.33 ms | 0.77 ms | 1.00 | 1.00 | 7173.1 KB | 1.00 | 607.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | ClosedXML | 200.12 ms | 10.36 ms | 5.98 ms | 8.55 | 8.55 | 105215.9 KB | 14.67 | 468.0 KB | 0.77 | 754.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | EPPlus | 240.93 ms | 3.42 ms | 1.98 ms | 10.29 | 10.29 | 106250.4 KB | 14.81 | 494.4 KB | 0.81 | 929.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-scalars | OfficeIMO.Excel | 16.95 ms | 1.10 ms | 0.63 ms | 1.00 | 1.00 | 5964.2 KB | 1.00 | 441.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-scalars | ClosedXML | 144.48 ms | 2.73 ms | 1.58 ms | 8.53 | 8.53 | 82583.3 KB | 13.85 | 394.9 KB | 0.89 | 752.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-scalars | EPPlus | 186.69 ms | 5.33 ms | 3.08 ms | 11.02 | 11.02 | 85057.4 KB | 14.26 | 379.3 KB | 0.86 | 1001.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings | OfficeIMO.Excel | 23.06 ms | 1.62 ms | 0.93 ms | 1.00 | 1.00 | 15019.9 KB | 1.00 | 527.8 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-strings | ClosedXML | 152.25 ms | 0.58 ms | 0.34 ms | 6.60 | 6.60 | 104233.3 KB | 6.94 | 460.1 KB | 0.87 | 560.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings | EPPlus | 167.28 ms | 4.29 ms | 2.48 ms | 7.26 | 7.26 | 84343.7 KB | 5.62 | 444.7 KB | 0.84 | 625.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-distinct | OfficeIMO.Excel | 20.39 ms | 1.28 ms | 0.74 ms | 1.00 | 1.00 | 13643.7 KB | 1.00 | 499.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-distinct | ClosedXML | 204.76 ms | 3.87 ms | 2.24 ms | 10.04 | 10.04 | 131493.2 KB | 9.64 | 555.3 KB | 1.11 | 904.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-distinct | EPPlus | 206.26 ms | 3.92 ms | 2.26 ms | 10.11 | 10.11 | 97646.6 KB | 7.16 | 565.1 KB | 1.13 | 911.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-repeated | OfficeIMO.Excel | 18.04 ms | 1.14 ms | 0.66 ms | 1.00 | 1.00 | 7184.9 KB | 1.00 | 376.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-repeated | ClosedXML | 136.24 ms | 2.97 ms | 1.71 ms | 7.55 | 7.55 | 84512.0 KB | 11.76 | 331.8 KB | 0.88 | 655.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-repeated | EPPlus | 140.99 ms | 3.77 ms | 2.18 ms | 7.81 | 7.81 | 69934.9 KB | 9.73 | 300.8 KB | 0.80 | 681.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-temporal | OfficeIMO.Excel | 30.65 ms | 1.31 ms | 0.76 ms | 1.00 | 1.00 | 7302.4 KB | 1.00 | 620.5 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-temporal | ClosedXML | 201.16 ms | 3.08 ms | 1.78 ms | 6.56 | 6.56 | 89315.7 KB | 12.23 | 483.0 KB | 0.78 | 556.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-temporal | EPPlus | 220.50 ms | 2.25 ms | 1.30 ms | 7.19 | 7.19 | 103733.9 KB | 14.21 | 495.1 KB | 0.80 | 619.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 13.52 ms | 0.75 ms | 0.43 ms | 0.91 | 1.00 | 3436.3 KB | 0.48 | 443.4 KB | 0.97 | 9.4% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 14.93 ms | 0.53 ms | 0.30 ms | 1.00 | 1.10 | 7226.8 KB | 1.00 | 455.5 KB | 1.00 | Loss +10.4% |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | ClosedXML | 156.14 ms | 2.19 ms | 1.27 ms | 10.46 | 11.55 | 96007.6 KB | 13.28 | 467.5 KB | 1.03 | 946.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | EPPlus | 176.07 ms | 1.09 ms | 0.63 ms | 11.80 | 13.03 | 87396.2 KB | 12.09 | 484.1 KB | 1.06 | 1079.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | LargeXlsx | 43.22 ms | 3.90 ms | 2.25 ms | 0.84 | 1.00 | 5606.0 KB | 0.36 | 1386.5 KB | 1.00 | 16.0% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 51.42 ms | 1.65 ms | 0.95 ms | 1.00 | 1.19 | 15700.7 KB | 1.00 | 1384.9 KB | 1.00 | Loss +19.0% |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | MiniExcel | 94.29 ms | 2.90 ms | 1.68 ms | 1.83 | 2.18 | 93247.0 KB | 5.94 | 1521.1 KB | 1.10 | 83.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | EPPlus | 421.55 ms | 4.80 ms | 2.77 ms | 8.20 | 9.75 | 211783.2 KB | 13.49 | 1090.0 KB | 0.79 | 719.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | ClosedXML | 426.81 ms | 1.64 ms | 0.95 ms | 8.30 | 9.88 | 210638.1 KB | 13.42 | 1139.9 KB | 0.82 | 730.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | Sylvan.Data.Excel | 41.31 ms | 4.36 ms | 2.52 ms | 0.78 | 1.00 | 5692.3 KB | 0.45 | 755.4 KB | 0.55 | 21.6% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | LargeXlsx | 52.29 ms | 0.83 ms | 0.48 ms | 0.99 | 1.27 | 8341.2 KB | 0.66 | 1386.5 KB | 1.00 | Tie vs OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | OfficeIMO.Excel | 52.68 ms | 2.04 ms | 1.18 ms | 1.00 | 1.28 | 12666.6 KB | 1.00 | 1384.9 KB | 1.00 | Loss +27.5% |
| 25000 | package-profile | package | Package size | write-datareader-plain | MiniExcel | 105.90 ms | 2.54 ms | 1.47 ms | 2.01 | 2.56 | 92189.6 KB | 7.28 | 1521.0 KB | 1.10 | 101.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | ClosedXML | 394.06 ms | 5.97 ms | 3.44 ms | 7.48 | 9.54 | 104197.0 KB | 8.23 | 1139.9 KB | 0.82 | 648.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | EPPlus | 409.65 ms | 24.53 ms | 14.16 ms | 7.78 | 9.92 | 117370.5 KB | 9.27 | 1090.8 KB | 0.79 | 677.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table | OfficeIMO.Excel | 54.87 ms | 3.54 ms | 2.04 ms | 1.00 | 1.00 | 12684.6 KB | 1.00 | 1385.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-datareader-table | MiniExcel | 105.42 ms | 4.51 ms | 2.61 ms | 1.92 | 1.92 | 92190.0 KB | 7.27 | 1521.0 KB | 1.10 | 92.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table | EPPlus | 410.15 ms | 23.83 ms | 13.76 ms | 7.47 | 7.47 | 117370.5 KB | 9.25 | 1090.8 KB | 0.79 | 647.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table | ClosedXML | 551.06 ms | 6.50 ms | 3.75 ms | 10.04 | 10.04 | 173390.0 KB | 13.67 | 1140.7 KB | 0.82 | 904.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | OfficeIMO.Excel | 58.90 ms | 2.69 ms | 1.56 ms | 1.00 | 1.00 | 12690.9 KB | 1.00 | 1385.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | MiniExcel | 119.78 ms | 11.85 ms | 6.84 ms | 2.03 | 2.03 | 124485.4 KB | 9.81 | 1521.1 KB | 1.10 | 103.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | EPPlus | 466.69 ms | 17.08 ms | 9.86 ms | 7.92 | 7.92 | 159670.6 KB | 12.58 | 1091.0 KB | 0.79 | 692.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | ClosedXML | 1204.10 ms | 25.31 ms | 14.61 ms | 20.44 | 20.44 | 566135.0 KB | 44.61 | 1140.9 KB | 0.82 | 1944.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | OfficeIMO.Excel | 51.16 ms | 0.30 ms | 0.18 ms | 1.00 | 1.00 | 9484.6 KB | 1.00 | 1329.2 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | LargeXlsx | 61.41 ms | 5.72 ms | 3.30 ms | 1.20 | 1.20 | 9257.9 KB | 0.98 | 1680.0 KB | 1.26 | 20.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | MiniExcel | 145.19 ms | 2.73 ms | 1.58 ms | 2.84 | 2.84 | 108118.7 KB | 11.40 | 1819.7 KB | 1.37 | 183.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | EPPlus | 672.92 ms | 15.91 ms | 9.18 ms | 13.15 | 13.15 | 135640.3 KB | 14.30 | 1390.4 KB | 1.05 | 1215.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | ClosedXML | 807.52 ms | 16.12 ms | 9.31 ms | 15.78 | 15.78 | 280361.4 KB | 29.56 | 1519.9 KB | 1.14 | 1478.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | OfficeIMO.Excel | 59.60 ms | 1.87 ms | 1.08 ms | 1.00 | 1.00 | 13123.2 KB | 1.00 | 1795.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | MiniExcel | 143.73 ms | 0.96 ms | 0.56 ms | 2.41 | 2.41 | 108118.7 KB | 8.24 | 1819.7 KB | 1.01 | 141.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | EPPlus | 650.81 ms | 0.92 ms | 0.53 ms | 10.92 | 10.92 | 135640.3 KB | 10.34 | 1390.4 KB | 0.77 | 991.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | ClosedXML | 811.59 ms | 12.56 ms | 7.25 ms | 13.62 | 13.62 | 280361.3 KB | 21.36 | 1519.9 KB | 0.85 | 1261.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables | OfficeIMO.Excel | 55.79 ms | 4.26 ms | 2.46 ms | 1.00 | 1.00 | 9792.8 KB | 1.00 | 1376.4 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-dataset-tables | MiniExcel | 117.54 ms | 3.03 ms | 1.75 ms | 2.11 | 2.11 | 97074.9 KB | 9.91 | 1511.8 KB | 1.10 | 110.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables | EPPlus | 382.34 ms | 8.83 ms | 5.10 ms | 6.85 | 6.85 | 110708.7 KB | 11.31 | 1100.6 KB | 0.80 | 585.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables | ClosedXML | 523.64 ms | 66.95 ms | 38.65 ms | 9.39 | 9.39 | 171992.8 KB | 17.56 | 1139.0 KB | 0.83 | 838.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | OfficeIMO.Excel | 56.87 ms | 0.73 ms | 0.42 ms | 1.00 | 1.00 | 9805.4 KB | 1.00 | 1376.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | MiniExcel | 119.78 ms | 3.95 ms | 2.28 ms | 2.11 | 2.11 | 128864.4 KB | 13.14 | 1512.0 KB | 1.10 | 110.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | EPPlus | 490.82 ms | 11.74 ms | 6.78 ms | 8.63 | 8.63 | 195297.9 KB | 19.92 | 1100.9 KB | 0.80 | 763.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | ClosedXML | 1106.11 ms | 8.72 ms | 5.03 ms | 19.45 | 19.45 | 550085.9 KB | 56.10 | 1139.3 KB | 0.83 | 1845.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-direct | LargeXlsx | 46.50 ms | 2.31 ms | 1.34 ms | 0.90 | 1.00 | 9512.4 KB | 0.77 | 1386.5 KB | 1.00 | 9.9% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-direct | OfficeIMO.Excel | 51.63 ms | 2.32 ms | 1.34 ms | 1.00 | 1.11 | 12380.0 KB | 1.00 | 1384.9 KB | 1.00 | Loss +11.0% |
| 25000 | package-profile | package | Package size | write-datatable-direct | MiniExcel | 113.69 ms | 4.43 ms | 2.56 ms | 2.20 | 2.44 | 92384.2 KB | 7.46 | 1521.1 KB | 1.10 | 120.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-direct | EPPlus | 393.57 ms | 19.96 ms | 11.52 ms | 7.62 | 8.46 | 117370.5 KB | 9.48 | 1090.8 KB | 0.79 | 662.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-direct | ClosedXML | 393.82 ms | 4.05 ms | 2.34 ms | 7.63 | 8.47 | 104197.0 KB | 8.42 | 1139.9 KB | 0.82 | 662.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | OfficeIMO.Excel | 54.26 ms | 1.62 ms | 0.93 ms | 1.00 | 1.00 | 12398.1 KB | 1.00 | 1385.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | MiniExcel | 117.75 ms | 0.49 ms | 0.28 ms | 2.17 | 2.17 | 92384.5 KB | 7.45 | 1521.1 KB | 1.10 | 117.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | EPPlus | 417.49 ms | 9.70 ms | 5.60 ms | 7.69 | 7.69 | 117370.5 KB | 9.47 | 1090.8 KB | 0.79 | 669.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | ClosedXML | 567.64 ms | 12.56 ms | 7.25 ms | 10.46 | 10.46 | 173388.1 KB | 13.99 | 1140.7 KB | 0.82 | 946.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | LargeXlsx | 38.05 ms | 0.63 ms | 0.36 ms | 0.80 | 1.00 | 5606.0 KB | 0.45 | 1386.5 KB | 1.00 | 19.8% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 47.42 ms | 1.26 ms | 0.72 ms | 1.00 | 1.25 | 12576.3 KB | 1.00 | 1384.9 KB | 1.00 | Loss +24.6% |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | MiniExcel | 88.33 ms | 2.71 ms | 1.56 ms | 1.86 | 2.32 | 93247.0 KB | 7.41 | 1521.1 KB | 1.10 | 86.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | EPPlus | 370.59 ms | 13.94 ms | 8.05 ms | 7.82 | 9.74 | 117370.5 KB | 9.33 | 1090.8 KB | 0.79 | 681.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | ClosedXML | 372.14 ms | 6.08 ms | 3.51 ms | 7.85 | 9.78 | 104197.0 KB | 8.29 | 1139.9 KB | 0.82 | 684.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 52.64 ms | 1.31 ms | 0.76 ms | 1.00 | 1.00 | 12585.6 KB | 1.00 | 1385.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 428.17 ms | 10.83 ms | 6.25 ms | 8.13 | 8.13 | 159670.6 KB | 12.69 | 1091.0 KB | 0.79 | 713.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 947.32 ms | 8.07 ms | 4.66 ms | 18.00 | 18.00 | 496948.9 KB | 39.49 | 1140.1 KB | 0.82 | 1699.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | LargeXlsx | 38.55 ms | 0.31 ms | 0.18 ms | 0.80 | 1.00 | 5606.0 KB | 0.45 | 1386.5 KB | 1.00 | 20.0% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | OfficeIMO.Excel | 48.22 ms | 0.30 ms | 0.17 ms | 1.00 | 1.25 | 12577.9 KB | 1.00 | 1384.9 KB | 1.00 | Loss +25.1% |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | MiniExcel | 91.32 ms | 0.65 ms | 0.38 ms | 1.89 | 2.37 | 93247.0 KB | 7.41 | 1521.1 KB | 1.10 | 89.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | EPPlus | 369.82 ms | 11.54 ms | 6.66 ms | 7.67 | 9.59 | 117370.5 KB | 9.33 | 1090.8 KB | 0.79 | 667.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | ClosedXML | 373.40 ms | 6.25 ms | 3.61 ms | 7.74 | 9.69 | 104197.0 KB | 8.28 | 1139.9 KB | 0.82 | 674.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 58.29 ms | 3.16 ms | 1.82 ms | 1.00 | 1.00 | 7029.1 KB | 1.00 | 1428.5 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 418.95 ms | 3.80 ms | 2.20 ms | 7.19 | 7.19 | 159670.6 KB | 22.72 | 1091.0 KB | 0.76 | 618.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 954.95 ms | 14.14 ms | 8.17 ms | 16.38 | 16.38 | 496948.9 KB | 70.70 | 1140.1 KB | 0.80 | 1538.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 38.55 ms | 1.90 ms | 1.09 ms | 0.75 | 1.00 | 5606.0 KB | 0.80 | 1386.5 KB | 0.97 | 25.4% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 51.67 ms | 2.66 ms | 1.54 ms | 1.00 | 1.34 | 7021.4 KB | 1.00 | 1428.4 KB | 1.00 | Loss +34.0% |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | MiniExcel | 89.23 ms | 1.29 ms | 0.74 ms | 1.73 | 2.31 | 93247.0 KB | 13.28 | 1521.1 KB | 1.06 | 72.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | ClosedXML | 366.84 ms | 2.79 ms | 1.61 ms | 7.10 | 9.52 | 104197.0 KB | 14.84 | 1139.9 KB | 0.80 | 610.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | EPPlus | 370.36 ms | 17.58 ms | 10.15 ms | 7.17 | 9.61 | 117370.5 KB | 16.72 | 1090.8 KB | 0.76 | 616.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | LargeXlsx | 40.68 ms | 0.89 ms | 0.52 ms | 0.68 | 1.00 | 5606.0 KB | 0.36 | 1386.5 KB | 0.97 | 31.5% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | OfficeIMO.Excel | 59.39 ms | 4.40 ms | 2.54 ms | 1.00 | 1.46 | 15616.0 KB | 1.00 | 1428.4 KB | 1.00 | Loss +46.0% |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | MiniExcel | 94.16 ms | 2.11 ms | 1.22 ms | 1.59 | 2.31 | 93247.0 KB | 5.97 | 1521.1 KB | 1.06 | 58.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | EPPlus | 379.73 ms | 14.66 ms | 8.46 ms | 6.39 | 9.34 | 117370.5 KB | 7.52 | 1090.8 KB | 0.76 | 539.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | ClosedXML | 380.01 ms | 5.51 ms | 3.18 ms | 6.40 | 9.34 | 104197.0 KB | 6.67 | 1139.9 KB | 0.80 | 539.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 49.22 ms | 0.63 ms | 0.36 ms | 1.00 | 1.00 | 12584.9 KB | 1.00 | 1385.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 396.09 ms | 1.54 ms | 0.89 ms | 8.05 | 8.05 | 138290.0 KB | 10.99 | 1091.0 KB | 0.79 | 704.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 593.53 ms | 5.71 ms | 3.30 ms | 12.06 | 12.06 | 275414.3 KB | 21.88 | 1140.1 KB | 0.82 | 1105.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | LargeXlsx | 53.76 ms | 0.96 ms | 0.55 ms | 0.88 | 1.00 | 6035.9 KB | 0.84 | 1816.3 KB | 0.99 | 11.8% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 60.97 ms | 3.17 ms | 1.83 ms | 1.00 | 1.13 | 7226.3 KB | 1.00 | 1828.0 KB | 1.00 | Loss +13.4% |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | MiniExcel | 115.84 ms | 1.62 ms | 0.94 ms | 1.90 | 2.15 | 113964.2 KB | 15.77 | 1936.7 KB | 1.06 | 90.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | ClosedXML | 509.25 ms | 22.45 ms | 12.96 ms | 8.35 | 9.47 | 179544.5 KB | 24.85 | 1555.2 KB | 0.85 | 735.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | EPPlus | 524.91 ms | 2.84 ms | 1.64 ms | 8.61 | 9.76 | 144853.2 KB | 20.05 | 1473.0 KB | 0.81 | 760.9% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | OfficeIMO.Excel | 425.06 ms | 99.08 ms | 57.20 ms | 1.00 | 1.00 | 136016.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | EPPlus | 498.57 ms | 90.07 ms | 52.00 ms | 1.17 | 1.17 | 250878.4 KB | 1.84 |  |  | 17.3% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | EPPlus 4.5.3.3 | 909.70 ms |  |  | 2.14 | 2.14 |  |  |  |  | 114.0% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | ClosedXML | 1721.62 ms | 310.51 ms | 179.27 ms | 4.05 | 4.05 | 829578.3 KB | 6.10 |  |  | 305.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Object projection | build-object-datatable-dictionaries | OfficeIMO.Excel | 14.84 ms | 1.22 ms | 0.71 ms | 1.00 | 1.00 | 5164.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Object projection | build-object-datatable-typed | OfficeIMO.Excel | 10.41 ms | 1.29 ms | 0.75 ms | 1.00 | 1.00 | 8093.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-cells | OfficeIMO.Excel | 56.30 ms | 5.98 ms | 3.45 ms | 1.00 | 1.00 | 23226.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-cells | EPPlus | 190.25 ms | 11.99 ms | 6.92 ms | 3.38 | 3.38 | 173827.0 KB | 7.48 |  |  | 237.9% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-cells | ClosedXML | 373.06 ms | 7.67 ms | 4.43 ms | 6.63 | 6.63 | 199121.1 KB | 8.57 |  |  | 562.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | OfficeIMO.Excel | 144.98 ms | 189.07 ms | 109.16 ms | 1.00 | 1.00 | 4000.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | EPPlus | 178.72 ms | 0.29 ms | 0.17 ms | 1.23 | 1.23 | 101975.4 KB | 25.49 |  |  | 23.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | ClosedXML | 358.72 ms | 29.94 ms | 17.28 ms | 2.47 | 2.47 | 183492.7 KB | 45.87 |  |  | 147.4% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-range | OfficeIMO.Excel | 78.13 ms | 2.85 ms | 1.65 ms | 1.00 | 1.00 | 23226.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-range | EPPlus | 263.45 ms | 12.45 ms | 7.19 ms | 3.37 | 3.37 | 173827.0 KB | 7.48 |  |  | 237.2% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-range | ClosedXML | 552.71 ms | 60.18 ms | 34.74 ms | 7.07 | 7.07 | 199119.6 KB | 8.57 |  |  | 607.4% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-top-range | OfficeIMO.Excel | 2.65 ms | 0.27 ms | 0.15 ms | 1.00 | 1.00 | 445.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-top-range | EPPlus | 231.31 ms | 8.97 ms | 5.18 ms | 87.27 | 87.27 | 92013.8 KB | 206.79 |  |  | 8627.4% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-top-range | ClosedXML | 610.72 ms | 219.12 ms | 126.51 ms | 230.42 | 230.42 | 181999.7 KB | 409.03 |  |  | 22942.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | OfficeIMO.Excel | 21.12 ms | 0.54 ms | 0.31 ms | 1.00 | 1.00 | 6287.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | EPPlus 4.5.3.3 | 111.90 ms |  |  | 5.30 | 5.30 |  |  |  |  | 429.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | EPPlus | 114.41 ms | 3.65 ms | 2.11 ms | 5.42 | 5.42 | 70813.8 KB | 11.26 |  |  | 441.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | ClosedXML | 175.75 ms | 4.12 ms | 2.38 ms | 8.32 | 8.32 | 79507.8 KB | 12.65 |  |  | 732.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | Sylvan.Data.Excel | 1.09 ms | 0.12 ms | 0.07 ms | 0.58 | 1.00 | 316.6 KB | 1.27 |  |  | 42.0% faster than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | ExcelDataReader | 1.72 ms | 0.53 ms | 0.31 ms | 0.92 | 1.58 | 4046.1 KB | 16.25 |  |  | 8.3% faster than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | OfficeIMO.Excel | 1.87 ms | 0.07 ms | 0.04 ms | 1.00 | 1.72 | 249.0 KB | 1.00 |  |  | Loss +72.4% |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | ClosedXML | 3.52 ms | 0.03 ms | 0.02 ms | 1.88 | 3.24 | 4392.9 KB | 17.64 |  |  | 87.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | MiniExcel | 11.18 ms | 0.47 ms | 0.27 ms | 5.97 | 10.28 | 46189.1 KB | 185.51 |  |  | 496.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | EPPlus 4.5.3.3 | 13.45 ms |  |  | 7.18 | 12.37 |  |  |  |  | 617.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | EPPlus | 58.52 ms | 3.36 ms | 1.94 ms | 31.23 | 53.84 | 43070.2 KB | 172.98 |  |  | 3023.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | Sylvan.Data.Excel | 1.03 ms | 0.09 ms | 0.05 ms | 0.60 | 1.00 | 316.6 KB | 1.27 |  |  | 39.8% faster than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | ExcelDataReader | 1.42 ms | 0.07 ms | 0.04 ms | 0.83 | 1.38 | 4046.1 KB | 16.25 |  |  | 16.9% faster than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | OfficeIMO.Excel | 1.71 ms | 0.06 ms | 0.04 ms | 1.00 | 1.66 | 249.1 KB | 1.00 |  |  | Loss +66.0% |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | ClosedXML | 3.39 ms | 0.04 ms | 0.02 ms | 1.98 | 3.29 | 4392.9 KB | 17.64 |  |  | 98.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | MiniExcel | 11.12 ms | 0.31 ms | 0.18 ms | 6.52 | 10.82 | 46189.1 KB | 185.46 |  |  | 551.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | EPPlus 4.5.3.3 | 15.64 ms |  |  | 9.16 | 15.22 |  |  |  |  | 816.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | EPPlus | 57.65 ms | 3.87 ms | 2.24 ms | 33.77 | 56.07 | 43070.2 KB | 172.93 |  |  | 3276.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | Sylvan.Data.Excel | 16.81 ms | 0.31 ms | 0.18 ms | 0.77 | 1.00 | 1936.7 KB | 0.21 |  |  | 23.1% faster than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | OfficeIMO.Excel | 21.85 ms | 0.71 ms | 0.41 ms | 1.00 | 1.30 | 9295.1 KB | 1.00 |  |  | Loss +30.0% |
| 25000 | speed-comparison | read | Other | shared-string-read | ExcelDataReader | 45.09 ms | 0.94 ms | 0.54 ms | 2.06 | 2.68 | 25004.8 KB | 2.69 |  |  | 106.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | MiniExcel | 51.19 ms | 3.38 ms | 1.95 ms | 2.34 | 3.05 | 74398.5 KB | 8.00 |  |  | 134.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | EPPlus | 97.28 ms | 1.93 ms | 1.11 ms | 4.45 | 5.79 | 89345.5 KB | 9.61 |  |  | 345.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | EPPlus 4.5.3.3 | 123.49 ms |  |  | 5.65 | 7.35 |  |  |  |  | 465.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | ClosedXML | 149.92 ms | 2.72 ms | 1.57 ms | 6.86 | 8.92 | 90414.0 KB | 9.73 |  |  | 586.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | OfficeIMO.Excel | 35.43 ms | 2.09 ms | 1.21 ms | 1.00 | 1.00 | 1263.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | Sylvan.Data.Excel | 40.18 ms | 2.18 ms | 1.26 ms | 1.13 | 1.13 | 400.4 KB | 0.32 |  |  | 13.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | ExcelDataReader | 117.72 ms | 6.53 ms | 3.77 ms | 3.32 | 3.32 | 61544.7 KB | 48.71 |  |  | 232.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | MiniExcel | 154.33 ms | 41.76 ms | 24.11 ms | 4.36 | 4.36 | 185059.6 KB | 146.46 |  |  | 335.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | EPPlus | 174.65 ms | 9.24 ms | 5.33 ms | 4.93 | 4.93 | 92042.7 KB | 72.84 |  |  | 393.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | ClosedXML | 349.20 ms | 18.93 ms | 10.93 ms | 9.86 | 9.86 | 181986.1 KB | 144.03 |  |  | 885.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | Sylvan.Data.Excel | 55.46 ms | 0.90 ms | 0.52 ms | 0.92 | 1.00 | 15259.8 KB | 0.46 |  |  | 8.4% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | OfficeIMO.Excel | 60.55 ms | 5.14 ms | 2.97 ms | 1.00 | 1.09 | 33325.6 KB | 1.00 |  |  | Loss +9.2% |
| 25000 | speed-comparison | read | Range and table read | read-datatable | ExcelDataReader | 134.07 ms | 6.80 ms | 3.92 ms | 2.21 | 2.42 | 76398.6 KB | 2.29 |  |  | 121.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | MiniExcel | 149.64 ms | 3.07 ms | 1.77 ms | 2.47 | 2.70 | 179910.9 KB | 5.40 |  |  | 147.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | EPPlus | 206.42 ms | 10.60 ms | 6.12 ms | 3.41 | 3.72 | 188683.4 KB | 5.66 |  |  | 240.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | EPPlus 4.5.3.3 | 260.12 ms |  |  | 4.30 | 4.69 |  |  |  |  | 329.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | ClosedXML | 364.70 ms | 2.92 ms | 1.69 ms | 6.02 | 6.58 | 211044.7 KB | 6.33 |  |  | 502.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | Sylvan.Data.Excel | 44.30 ms | 1.42 ms | 0.82 ms | 0.98 | 1.00 | 1175.8 KB | 0.28 |  |  | Tie vs OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | OfficeIMO.Excel | 45.05 ms | 10.96 ms | 6.33 ms | 1.00 | 1.02 | 4192.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | MiniExcel | 107.30 ms | 9.24 ms | 5.33 ms | 2.38 | 2.42 | 157250.4 KB | 37.50 |  |  | 138.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | ExcelDataReader | 117.17 ms | 5.35 ms | 3.09 ms | 2.60 | 2.64 | 61547.4 KB | 14.68 |  |  | 160.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | EPPlus | 205.84 ms | 47.28 ms | 27.30 ms | 4.57 | 4.65 | 101975.4 KB | 24.32 |  |  | 356.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | ClosedXML | 411.75 ms | 82.78 ms | 47.80 ms | 9.14 | 9.29 | 183491.8 KB | 43.76 |  |  | 814.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | Sylvan.Data.Excel | 64.85 ms | 11.15 ms | 6.44 ms | 0.68 | 1.00 | 400.4 KB | 0.02 |  |  | 32.3% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | OfficeIMO.Excel | 95.74 ms | 2.01 ms | 1.16 ms | 1.00 | 1.48 | 24786.3 KB | 1.00 |  |  | Loss +47.6% |
| 25000 | speed-comparison | read | Range and table read | read-range | ExcelDataReader | 187.10 ms | 6.49 ms | 3.74 ms | 1.95 | 2.89 | 61547.4 KB | 2.48 |  |  | 95.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | MiniExcel | 196.40 ms | 17.50 ms | 10.10 ms | 2.05 | 3.03 | 185060.7 KB | 7.47 |  |  | 105.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | EPPlus 4.5.3.3 | 230.63 ms |  |  | 2.41 | 3.56 |  |  |  |  | 140.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | EPPlus | 280.39 ms | 5.43 ms | 3.14 ms | 2.93 | 4.32 | 173824.6 KB | 7.01 |  |  | 192.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | ClosedXML | 513.96 ms | 6.37 ms | 3.68 ms | 5.37 | 7.93 | 196187.3 KB | 7.92 |  |  | 436.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | Sylvan.Data.Excel | 78.91 ms | 3.50 ms | 2.02 ms | 0.85 | 1.00 | 1350.5 KB | 0.05 |  |  | 15.1% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | OfficeIMO.Excel | 92.90 ms | 3.54 ms | 2.04 ms | 1.00 | 1.18 | 25372.2 KB | 1.00 |  |  | Loss +17.7% |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | MiniExcel | 192.45 ms | 9.20 ms | 5.31 ms | 2.07 | 2.44 | 185060.7 KB | 7.29 |  |  | 107.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | ExcelDataReader | 192.58 ms | 3.43 ms | 1.98 ms | 2.07 | 2.44 | 61547.4 KB | 2.43 |  |  | 107.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | EPPlus | 287.81 ms | 27.99 ms | 16.16 ms | 3.10 | 3.65 | 173824.6 KB | 6.85 |  |  | 209.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | ClosedXML | 512.92 ms | 22.17 ms | 12.80 ms | 5.52 | 6.50 | 196186.3 KB | 7.73 |  |  | 452.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | Sylvan.Data.Excel | 0.44 ms | 0.00 ms | 0.00 ms | 0.23 | 1.00 | 367.3 KB | 0.82 |  |  | 77.4% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | MiniExcel | 1.01 ms | 0.28 ms | 0.16 ms | 0.52 | 2.32 | 973.1 KB | 2.17 |  |  | 47.6% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | OfficeIMO.Excel | 1.92 ms | 0.13 ms | 0.08 ms | 1.00 | 4.42 | 447.6 KB | 1.00 |  |  | Loss +342.4% |
| 25000 | speed-comparison | read | Range and table read | read-top-range | ExcelDataReader | 38.24 ms | 1.00 ms | 0.58 ms | 19.87 | 87.90 | 17173.8 KB | 38.37 |  |  | 1886.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | EPPlus | 173.93 ms | 9.88 ms | 5.70 ms | 90.36 | 399.78 | 92011.5 KB | 205.55 |  |  | 8935.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | EPPlus 4.5.3.3 | 182.26 ms |  |  | 94.68 | 418.91 |  |  |  |  | 9368.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | ClosedXML | 353.32 ms | 3.80 ms | 2.19 ms | 183.55 | 812.10 | 181986.3 KB | 406.55 |  |  | 18254.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | Sylvan.Data.Excel | 72.76 ms | 4.05 ms | 2.34 ms | 0.46 | 1.00 | 400.4 KB | 0.01 |  |  | 54.0% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | OfficeIMO.Excel | 158.02 ms | 0.60 ms | 0.34 ms | 1.00 | 2.17 | 32779.2 KB | 1.00 |  |  | Loss +117.2% |
| 25000 | speed-comparison | read | Range and table read | read-used-range | ExcelDataReader | 178.84 ms | 6.23 ms | 3.60 ms | 1.13 | 2.46 | 61547.4 KB | 1.88 |  |  | 13.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | MiniExcel | 189.95 ms | 13.55 ms | 7.82 ms | 1.20 | 2.61 | 185060.7 KB | 5.65 |  |  | 20.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | EPPlus | 261.65 ms | 18.02 ms | 10.40 ms | 1.66 | 3.60 | 173824.6 KB | 5.30 |  |  | 65.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | ClosedXML | 531.27 ms | 50.22 ms | 29.00 ms | 3.36 | 7.30 | 196188.2 KB | 5.99 |  |  | 236.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | OfficeIMO.Excel | 36.18 ms | 4.54 ms | 2.62 ms | 1.00 | 1.00 | 1261.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | Sylvan.Data.Excel | 38.91 ms | 1.18 ms | 0.68 ms | 1.08 | 1.08 | 400.4 KB | 0.32 |  |  | 7.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | ExcelDataReader | 115.37 ms | 3.73 ms | 2.16 ms | 3.19 | 3.19 | 61539.4 KB | 48.78 |  |  | 218.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | MiniExcel | 128.03 ms | 2.40 ms | 1.38 ms | 3.54 | 3.54 | 185043.9 KB | 146.67 |  |  | 253.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | EPPlus | 158.12 ms | 8.83 ms | 5.10 ms | 4.37 | 4.37 | 92042.2 KB | 72.95 |  |  | 337.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | ClosedXML | 353.33 ms | 24.60 ms | 14.21 ms | 9.77 | 9.77 | 181989.7 KB | 144.25 |  |  | 876.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | Sylvan.Data.Excel | 44.91 ms | 2.61 ms | 1.51 ms | 0.84 | 1.00 | 400.4 KB | 0.02 |  |  | 15.7% faster than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | OfficeIMO.Excel | 53.27 ms | 4.16 ms | 2.40 ms | 1.00 | 1.19 | 25565.2 KB | 1.00 |  |  | Loss +18.6% |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | ExcelDataReader | 112.51 ms | 0.94 ms | 0.54 ms | 2.11 | 2.51 | 61539.4 KB | 2.41 |  |  | 111.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | MiniExcel | 125.39 ms | 2.90 ms | 1.68 ms | 2.35 | 2.79 | 185044.5 KB | 7.24 |  |  | 135.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | EPPlus | 205.66 ms | 10.11 ms | 5.84 ms | 3.86 | 4.58 | 173824.1 KB | 6.80 |  |  | 286.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | EPPlus 4.5.3.3 | 218.89 ms |  |  | 4.11 | 4.87 |  |  |  |  | 310.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | ClosedXML | 339.06 ms | 6.26 ms | 3.62 ms | 6.37 | 7.55 | 196191.4 KB | 7.67 |  |  | 536.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | Sylvan.Data.Excel | 0.47 ms | 0.03 ms | 0.02 ms | 0.22 | 1.00 | 367.3 KB | 0.83 |  |  | 77.9% faster than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | MiniExcel | 0.83 ms | 0.05 ms | 0.03 ms | 0.39 | 1.77 | 959.8 KB | 2.17 |  |  | 61.0% faster than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | OfficeIMO.Excel | 2.14 ms | 0.16 ms | 0.09 ms | 1.00 | 4.53 | 443.0 KB | 1.00 |  |  | Loss +353.0% |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | ExcelDataReader | 38.84 ms | 2.29 ms | 1.32 ms | 18.17 | 82.33 | 17165.8 KB | 38.75 |  |  | 1717.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | EPPlus | 155.37 ms | 5.57 ms | 3.22 ms | 72.71 | 329.38 | 92011.0 KB | 207.70 |  |  | 7171.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | EPPlus 4.5.3.3 | 177.92 ms |  |  | 83.26 | 377.19 |  |  |  |  | 8226.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | ClosedXML | 355.00 ms | 14.97 ms | 8.65 ms | 166.13 | 752.60 | 181987.4 KB | 410.81 |  |  | 16513.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | Sylvan.Data.Excel | 0.46 ms | 0.06 ms | 0.04 ms | 0.22 | 1.00 | 367.3 KB | 0.83 |  |  | 77.9% faster than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | MiniExcel | 0.75 ms | 0.04 ms | 0.02 ms | 0.36 | 1.63 | 959.8 KB | 2.16 |  |  | 64.0% faster than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | OfficeIMO.Excel | 2.08 ms | 0.18 ms | 0.10 ms | 1.00 | 4.53 | 443.7 KB | 1.00 |  |  | Loss +352.9% |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | ExcelDataReader | 37.65 ms | 1.64 ms | 0.95 ms | 18.12 | 82.07 | 17165.8 KB | 38.69 |  |  | 1712.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | EPPlus | 152.04 ms | 6.19 ms | 3.58 ms | 73.17 | 331.41 | 92011.0 KB | 207.38 |  |  | 7217.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | ClosedXML | 329.59 ms | 11.06 ms | 6.39 ms | 158.62 | 718.42 | 181984.3 KB | 410.17 |  |  | 15762.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | Sylvan.Data.Excel | 44.32 ms | 1.37 ms | 0.79 ms | 0.88 | 1.00 | 2670.6 KB | 0.12 |  |  | 11.7% faster than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | OfficeIMO.Excel | 50.18 ms | 2.01 ms | 1.16 ms | 1.00 | 1.13 | 22242.0 KB | 1.00 |  |  | Loss +13.2% |
| 25000 | speed-comparison | read | Typed object read | read-objects | ExcelDataReader | 113.13 ms | 2.22 ms | 1.28 ms | 2.25 | 2.55 | 63809.5 KB | 2.87 |  |  | 125.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | MiniExcel | 131.88 ms | 3.12 ms | 1.80 ms | 2.63 | 2.98 | 182282.9 KB | 8.20 |  |  | 162.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | EPPlus | 194.35 ms | 2.17 ms | 1.25 ms | 3.87 | 4.39 | 186041.4 KB | 8.36 |  |  | 287.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | EPPlus 4.5.3.3 | 226.02 ms |  |  | 4.50 | 5.10 |  |  |  |  | 350.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | ClosedXML | 346.02 ms | 9.51 ms | 5.49 ms | 6.90 | 7.81 | 198139.6 KB | 8.91 |  |  | 589.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | Sylvan.Data.Excel | 63.79 ms | 18.95 ms | 10.94 ms | 0.88 | 1.00 | 2158.2 KB | 0.10 |  |  | 11.7% faster than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | OfficeIMO.Excel | 72.25 ms | 17.15 ms | 9.90 ms | 1.00 | 1.13 | 22242.4 KB | 1.00 |  |  | Loss +13.3% |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | ExcelDataReader | 148.07 ms | 32.20 ms | 18.59 ms | 2.05 | 2.32 | 63297.2 KB | 2.85 |  |  | 104.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | MiniExcel | 166.45 ms | 33.20 ms | 19.17 ms | 2.30 | 2.61 | 181770.6 KB | 8.17 |  |  | 130.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | EPPlus 4.5.3.3 | 216.02 ms |  |  | 2.99 | 3.39 |  |  |  |  | 199.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | EPPlus | 230.93 ms | 32.58 ms | 18.81 ms | 3.20 | 3.62 | 185846.0 KB | 8.36 |  |  | 219.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | ClosedXML | 447.80 ms | 74.46 ms | 42.99 ms | 6.20 | 7.02 | 197944.0 KB | 8.90 |  |  | 519.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | OfficeIMO.Excel | 49.71 ms | 2.76 ms | 1.59 ms | 1.00 | 1.00 | 12690.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | MiniExcel | 95.92 ms | 3.89 ms | 2.25 ms | 1.93 | 1.93 | 124485.4 KB | 9.81 |  |  | 93.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | EPPlus | 380.97 ms | 13.57 ms | 7.83 ms | 7.66 | 7.66 | 159670.6 KB | 12.58 |  |  | 666.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | EPPlus 4.5.3.3 | 454.89 ms |  |  | 9.15 | 9.15 |  |  |  |  | 815.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | ClosedXML | 1063.26 ms | 44.14 ms | 25.48 ms | 21.39 | 21.39 | 566137.7 KB | 44.61 |  |  | 2039.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | OfficeIMO.Excel | 49.55 ms | 3.18 ms | 1.83 ms | 1.00 | 1.00 | 9805.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | MiniExcel | 106.52 ms | 4.11 ms | 2.37 ms | 2.15 | 2.15 | 128864.4 KB | 13.14 |  |  | 115.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | EPPlus | 442.21 ms | 9.52 ms | 5.49 ms | 8.92 | 8.92 | 195297.9 KB | 19.92 |  |  | 792.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | ClosedXML | 1040.89 ms | 18.43 ms | 10.64 ms | 21.01 | 21.01 | 550081.6 KB | 56.10 |  |  | 2000.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 54.90 ms | 2.32 ms | 1.34 ms | 1.00 | 1.00 | 12588.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 427.07 ms | 18.56 ms | 10.72 ms | 7.78 | 7.78 | 159694.5 KB | 12.69 |  |  | 677.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 1019.70 ms | 26.12 ms | 15.08 ms | 18.57 | 18.57 | 496954.3 KB | 39.48 |  |  | 1757.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 82.00 ms | 14.17 ms | 8.18 ms | 1.00 | 1.00 | 7037.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 476.10 ms | 92.32 ms | 53.30 ms | 5.81 | 5.81 | 159742.3 KB | 22.70 |  |  | 480.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 1087.29 ms | 110.95 ms | 64.06 ms | 13.26 | 13.26 | 496956.9 KB | 70.62 |  |  | 1225.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 51.91 ms | 2.38 ms | 1.38 ms | 1.00 | 1.00 | 12584.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 398.70 ms | 3.01 ms | 1.74 ms | 7.68 | 7.68 | 138290.0 KB | 10.99 |  |  | 668.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 668.25 ms | 75.75 ms | 43.73 ms | 12.87 | 12.87 | 275414.3 KB | 21.88 |  |  | 1187.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | OfficeIMO.Excel | 17.83 ms | 1.52 ms | 0.88 ms | 1.00 | 1.00 | 6716.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | ClosedXML | 149.02 ms | 2.78 ms | 1.60 ms | 8.36 | 8.36 | 92894.1 KB | 13.83 |  |  | 736.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | EPPlus | 149.42 ms | 5.70 ms | 3.29 ms | 8.38 | 8.38 | 74425.8 KB | 11.08 |  |  | 738.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | OfficeIMO.Excel | 21.77 ms | 2.83 ms | 1.63 ms | 1.00 | 1.00 | 5790.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | EPPlus 4.5.3.3 | 91.08 ms |  |  | 4.18 | 4.18 |  |  |  |  | 318.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | ClosedXML | 149.81 ms | 7.49 ms | 4.33 ms | 6.88 | 6.88 | 84198.7 KB | 14.54 |  |  | 588.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | EPPlus | 177.08 ms | 9.08 ms | 5.24 ms | 8.14 | 8.14 | 86279.5 KB | 14.90 |  |  | 713.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | OfficeIMO.Excel | 29.44 ms | 0.66 ms | 0.38 ms | 1.00 | 1.00 | 7993.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | EPPlus 4.5.3.3 | 112.75 ms |  |  | 3.83 | 3.83 |  |  |  |  | 283.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | EPPlus | 214.18 ms | 10.08 ms | 5.82 ms | 7.28 | 7.28 | 113162.8 KB | 14.16 |  |  | 627.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | ClosedXML | 233.38 ms | 17.94 ms | 10.36 ms | 7.93 | 7.93 | 111110.6 KB | 13.90 |  |  | 692.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | OfficeIMO.Excel | 31.35 ms | 0.45 ms | 0.26 ms | 1.00 | 1.00 | 7173.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | ClosedXML | 210.09 ms | 7.42 ms | 4.28 ms | 6.70 | 6.70 | 105215.9 KB | 14.67 |  |  | 570.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | EPPlus | 229.40 ms | 4.31 ms | 2.49 ms | 7.32 | 7.32 | 106250.4 KB | 14.81 |  |  | 631.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 23.52 ms | 2.99 ms | 1.73 ms | 1.00 | 1.00 | 7173.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | ClosedXML | 209.63 ms | 9.45 ms | 5.46 ms | 8.91 | 8.91 | 105215.9 KB | 14.67 |  |  | 791.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | EPPlus | 244.44 ms | 4.21 ms | 2.43 ms | 10.39 | 10.39 | 106250.4 KB | 14.81 |  |  | 939.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | OfficeIMO.Excel | 16.68 ms | 2.35 ms | 1.36 ms | 1.00 | 1.00 | 5964.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | EPPlus 4.5.3.3 | 92.13 ms |  |  | 5.52 | 5.52 |  |  |  |  | 452.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | ClosedXML | 134.47 ms | 4.19 ms | 2.42 ms | 8.06 | 8.06 | 82583.3 KB | 13.85 |  |  | 706.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | EPPlus | 174.97 ms | 9.26 ms | 5.35 ms | 10.49 | 10.49 | 85057.4 KB | 14.26 |  |  | 949.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | OfficeIMO.Excel | 30.31 ms | 2.12 ms | 1.22 ms | 1.00 | 1.00 | 7302.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | EPPlus 4.5.3.3 | 106.01 ms |  |  | 3.50 | 3.50 |  |  |  |  | 249.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | ClosedXML | 212.00 ms | 16.66 ms | 9.62 ms | 6.99 | 6.99 | 89315.7 KB | 12.23 |  |  | 599.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | EPPlus | 221.41 ms | 12.80 ms | 7.39 ms | 7.30 | 7.30 | 103733.9 KB | 14.21 |  |  | 630.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | OfficeIMO.Excel | 48.30 ms | 3.70 ms | 2.14 ms | 1.00 | 1.00 | 12544.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | MiniExcel | 108.40 ms | 2.97 ms | 1.71 ms | 2.24 | 2.24 | 97077.8 KB | 7.74 |  |  | 124.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | EPPlus | 480.25 ms | 14.96 ms | 8.64 ms | 9.94 | 9.94 | 111170.8 KB | 8.86 |  |  | 894.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | ClosedXML | 502.48 ms | 15.21 ms | 8.78 ms | 10.40 | 10.40 | 172008.1 KB | 13.71 |  |  | 940.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | OfficeIMO.Excel | 53.63 ms | 3.02 ms | 1.74 ms | 1.00 | 1.00 | 13123.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | MiniExcel | 126.43 ms | 6.07 ms | 3.50 ms | 2.36 | 2.36 | 108118.7 KB | 8.24 |  |  | 135.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | EPPlus | 573.54 ms | 3.08 ms | 1.78 ms | 10.69 | 10.69 | 135640.3 KB | 10.34 |  |  | 969.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | ClosedXML | 745.19 ms | 7.59 ms | 4.38 ms | 13.89 | 13.89 | 280358.3 KB | 21.36 |  |  | 1289.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | OfficeIMO.Excel | 47.79 ms | 5.67 ms | 3.28 ms | 1.00 | 1.00 | 9792.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | MiniExcel | 98.40 ms | 1.84 ms | 1.06 ms | 2.06 | 2.06 | 97074.9 KB | 9.91 |  |  | 105.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | EPPlus 4.5.3.3 | 231.75 ms |  |  | 4.85 | 4.85 |  |  |  |  | 384.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | EPPlus | 341.38 ms | 19.01 ms | 10.98 ms | 7.14 | 7.14 | 110708.7 KB | 11.31 |  |  | 614.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | ClosedXML | 476.93 ms | 25.16 ms | 14.52 ms | 9.98 | 9.98 | 171990.1 KB | 17.56 |  |  | 898.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | OfficeIMO.Excel | 49.48 ms | 2.52 ms | 1.46 ms | 1.00 | 1.00 | 12684.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | MiniExcel | 95.70 ms | 3.61 ms | 2.08 ms | 1.93 | 1.93 | 92190.0 KB | 7.27 |  |  | 93.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | EPPlus 4.5.3.3 | 217.74 ms |  |  | 4.40 | 4.40 |  |  |  |  | 340.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | EPPlus | 368.45 ms | 16.58 ms | 9.57 ms | 7.45 | 7.45 | 117370.5 KB | 9.25 |  |  | 644.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | ClosedXML | 487.84 ms | 16.82 ms | 9.71 ms | 9.86 | 9.86 | 173394.0 KB | 13.67 |  |  | 886.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | LargeXlsx | 42.44 ms | 1.83 ms | 1.06 ms | 0.94 | 1.00 | 9512.4 KB | 0.77 |  |  | 6.2% faster than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | OfficeIMO.Excel | 45.25 ms | 1.96 ms | 1.13 ms | 1.00 | 1.07 | 12380.0 KB | 1.00 |  |  | Loss +6.6% |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | MiniExcel | 100.03 ms | 7.75 ms | 4.48 ms | 2.21 | 2.36 | 92384.2 KB | 7.46 |  |  | 121.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | EPPlus 4.5.3.3 | 230.61 ms |  |  | 5.10 | 5.43 |  |  |  |  | 409.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | ClosedXML | 368.54 ms | 5.47 ms | 3.16 ms | 8.14 | 8.68 | 104197.0 KB | 8.42 |  |  | 714.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | EPPlus | 374.20 ms | 8.82 ms | 5.09 ms | 8.27 | 8.82 | 117370.5 KB | 9.48 |  |  | 726.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | OfficeIMO.Excel | 48.83 ms | 1.39 ms | 0.80 ms | 1.00 | 1.00 | 9663.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | MiniExcel | 106.06 ms | 8.98 ms | 5.19 ms | 2.17 | 2.17 | 89650.1 KB | 9.28 |  |  | 117.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | EPPlus | 368.87 ms | 20.56 ms | 11.87 ms | 7.55 | 7.55 | 114636.2 KB | 11.86 |  |  | 655.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | ClosedXML | 488.68 ms | 20.70 ms | 11.95 ms | 10.01 | 10.01 | 170650.6 KB | 17.66 |  |  | 900.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | OfficeIMO.Excel | 53.05 ms | 1.40 ms | 0.81 ms | 1.00 | 1.00 | 12398.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | MiniExcel | 108.42 ms | 6.19 ms | 3.58 ms | 2.04 | 2.04 | 92384.5 KB | 7.45 |  |  | 104.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | EPPlus 4.5.3.3 | 227.28 ms |  |  | 4.28 | 4.28 |  |  |  |  | 328.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | EPPlus | 372.19 ms | 5.56 ms | 3.21 ms | 7.02 | 7.02 | 117370.5 KB | 9.47 |  |  | 601.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | ClosedXML | 512.46 ms | 13.95 ms | 8.05 ms | 9.66 | 9.66 | 173393.5 KB | 13.99 |  |  | 866.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | OfficeIMO.Excel | 53.23 ms | 4.87 ms | 2.81 ms | 1.00 | 1.00 | 12650.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | MiniExcel | 96.82 ms | 6.84 ms | 3.95 ms | 1.82 | 1.82 | 125545.0 KB | 9.92 |  |  | 81.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | EPPlus | 470.38 ms | 24.56 ms | 14.18 ms | 8.84 | 8.84 | 254934.4 KB | 20.15 |  |  | 783.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | EPPlus 4.5.3.3 | 529.04 ms |  |  | 9.94 | 9.94 |  |  |  |  | 893.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | ClosedXML | 1088.95 ms | 29.08 ms | 16.79 ms | 20.46 | 20.46 | 565952.9 KB | 44.74 |  |  | 1945.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | OfficeIMO.Excel | 33.79 ms | 6.79 ms | 3.92 ms | 1.00 | 1.00 | 9547.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | EPPlus 4.5.3.3 | 124.66 ms |  |  | 3.69 | 3.69 |  |  |  |  | 269.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | ClosedXML | 318.84 ms | 28.91 ms | 16.69 ms | 9.44 | 9.44 | 113847.8 KB | 11.92 |  |  | 843.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | EPPlus | 499.00 ms | 32.50 ms | 18.76 ms | 14.77 | 14.77 | 140687.9 KB | 14.74 |  |  | 1376.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-dictionary-objects-table-direct | OfficeIMO.Excel | 63.08 ms | 2.71 ms | 1.56 ms | 1.00 | 1.00 | 14828.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | LargeXlsx | 45.88 ms | 2.89 ms | 1.67 ms | 0.85 | 1.00 | 6043.9 KB | 0.84 |  |  | 15.5% faster than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 54.28 ms | 2.41 ms | 1.39 ms | 1.00 | 1.18 | 7234.3 KB | 1.00 |  |  | Loss +18.3% |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | MiniExcel | 101.37 ms | 5.61 ms | 3.24 ms | 1.87 | 2.21 | 113974.3 KB | 15.75 |  |  | 86.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | EPPlus | 468.01 ms | 15.71 ms | 9.07 ms | 8.62 | 10.20 | 144919.9 KB | 20.03 |  |  | 762.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | ClosedXML | 497.42 ms | 56.50 ms | 32.62 ms | 9.16 | 10.84 | 179552.5 KB | 24.82 |  |  | 816.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | LargeXlsx | 17.83 ms | 1.75 ms | 1.01 ms | 0.64 | 1.00 | 2771.0 KB | 0.24 |  |  | 35.8% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | OfficeIMO.Excel | 27.80 ms | 1.52 ms | 0.88 ms | 1.00 | 1.56 | 11672.4 KB | 1.00 |  |  | Loss +55.9% |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | MiniExcel | 53.39 ms | 2.46 ms | 1.42 ms | 1.92 | 2.99 | 58242.9 KB | 4.99 |  |  | 92.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | EPPlus 4.5.3.3 | 131.78 ms |  |  | 4.74 | 7.39 |  |  |  |  | 374.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | ClosedXML | 204.47 ms | 5.25 ms | 3.03 ms | 7.36 | 11.47 | 104233.1 KB | 8.93 |  |  | 635.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | EPPlus | 217.72 ms | 11.51 ms | 6.64 ms | 7.83 | 12.21 | 100373.4 KB | 8.60 |  |  | 683.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 12.78 ms | 0.37 ms | 0.22 ms | 0.86 | 1.00 | 3436.3 KB | 0.48 |  |  | 14.5% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 14.95 ms | 1.20 ms | 0.69 ms | 1.00 | 1.17 | 7226.8 KB | 1.00 |  |  | Loss +16.9% |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | ClosedXML | 152.84 ms | 14.20 ms | 8.20 ms | 10.23 | 11.95 | 96007.6 KB | 13.28 |  |  | 922.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | EPPlus | 166.81 ms | 6.64 ms | 3.83 ms | 11.16 | 13.05 | 87396.2 KB | 12.09 |  |  | 1016.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | LargeXlsx | 40.23 ms | 2.27 ms | 1.31 ms | 0.80 | 1.00 | 5606.0 KB | 0.36 |  |  | 20.2% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 50.39 ms | 1.43 ms | 0.83 ms | 1.00 | 1.25 | 15700.7 KB | 1.00 |  |  | Loss +25.2% |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | MiniExcel | 90.01 ms | 5.10 ms | 2.94 ms | 1.79 | 2.24 | 93247.0 KB | 5.94 |  |  | 78.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 228.62 ms |  |  | 4.54 | 5.68 |  |  |  |  | 353.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | EPPlus | 419.29 ms | 23.16 ms | 13.37 ms | 8.32 | 10.42 | 211783.2 KB | 13.49 |  |  | 732.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | ClosedXML | 422.39 ms | 9.01 ms | 5.20 ms | 8.38 | 10.50 | 210638.1 KB | 13.42 |  |  | 738.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | OfficeIMO.Excel | 22.22 ms | 2.28 ms | 1.32 ms | 1.00 | 1.00 | 7530.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | ClosedXML | 188.68 ms | 2.18 ms | 1.26 ms | 8.49 | 8.49 | 105215.9 KB | 13.97 |  |  | 749.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | EPPlus | 221.98 ms | 1.78 ms | 1.03 ms | 9.99 | 9.99 | 106250.4 KB | 14.11 |  |  | 899.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | Sylvan.Data.Excel | 39.84 ms | 0.55 ms | 0.32 ms | 0.78 | 1.00 | 5692.3 KB | 0.45 |  |  | 22.4% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | LargeXlsx | 48.38 ms | 3.70 ms | 2.14 ms | 0.94 | 1.21 | 8341.2 KB | 0.66 |  |  | 5.8% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | OfficeIMO.Excel | 51.35 ms | 3.04 ms | 1.76 ms | 1.00 | 1.29 | 12666.6 KB | 1.00 |  |  | Loss +28.9% |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | MiniExcel | 99.08 ms | 2.13 ms | 1.23 ms | 1.93 | 2.49 | 92189.6 KB | 7.28 |  |  | 93.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | EPPlus 4.5.3.3 | 212.03 ms |  |  | 4.13 | 5.32 |  |  |  |  | 312.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | ClosedXML | 375.50 ms | 17.42 ms | 10.06 ms | 7.31 | 9.42 | 104197.0 KB | 8.23 |  |  | 631.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | EPPlus | 382.21 ms | 12.66 ms | 7.31 ms | 7.44 | 9.59 | 117370.5 KB | 9.27 |  |  | 644.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | OfficeIMO.Excel | 42.54 ms | 1.97 ms | 1.14 ms | 1.00 | 1.00 | 9484.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | LargeXlsx | 49.41 ms | 6.17 ms | 3.56 ms | 1.16 | 1.16 | 9257.9 KB | 0.98 |  |  | 16.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | MiniExcel | 125.57 ms | 5.40 ms | 3.12 ms | 2.95 | 2.95 | 108118.7 KB | 11.40 |  |  | 195.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | EPPlus | 568.43 ms | 18.24 ms | 10.53 ms | 13.36 | 13.36 | 135640.3 KB | 14.30 |  |  | 1236.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | ClosedXML | 760.20 ms | 27.89 ms | 16.10 ms | 17.87 | 17.87 | 280365.7 KB | 29.56 |  |  | 1687.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | LargeXlsx | 53.82 ms | 4.04 ms | 2.33 ms | 0.91 | 1.00 | 10787.2 KB | 0.85 |  |  | 9.5% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | OfficeIMO.Excel | 59.47 ms | 4.29 ms | 2.48 ms | 1.00 | 1.10 | 12754.6 KB | 1.00 |  |  | Loss +10.5% |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | MiniExcel | 201.46 ms | 5.86 ms | 3.38 ms | 3.39 | 3.74 | 226867.4 KB | 17.79 |  |  | 238.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | ClosedXML | 1174.32 ms | 54.41 ms | 31.41 ms | 19.75 | 21.82 | 759810.8 KB | 59.57 |  |  | 1874.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | OfficeIMO.Excel | 17.94 ms | 1.36 ms | 0.78 ms | 1.00 | 1.00 | 15409.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | MiniExcel | 29.31 ms | 0.77 ms | 0.45 ms | 1.63 | 1.63 | 73751.2 KB | 4.79 |  |  | 63.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | ClosedXML | 112.12 ms | 2.62 ms | 1.51 ms | 6.25 | 6.25 | 104233.3 KB | 6.76 |  |  | 525.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | EPPlus | 132.55 ms | 10.33 ms | 5.96 ms | 7.39 | 7.39 | 84343.7 KB | 5.47 |  |  | 638.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | EPPlus 4.5.3.3 | 272.45 ms |  |  | 15.19 | 15.19 |  |  |  |  | 1418.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | OfficeIMO.Excel | 24.55 ms | 2.44 ms | 1.41 ms | 1.00 | 1.00 | 15019.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | EPPlus 4.5.3.3 | 94.47 ms |  |  | 3.85 | 3.85 |  |  |  |  | 284.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | ClosedXML | 140.20 ms | 11.38 ms | 6.57 ms | 5.71 | 5.71 | 104233.3 KB | 6.94 |  |  | 471.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | EPPlus | 146.21 ms | 7.71 ms | 4.45 ms | 5.96 | 5.96 | 84343.7 KB | 5.62 |  |  | 495.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | OfficeIMO.Excel | 19.01 ms | 1.52 ms | 0.88 ms | 1.00 | 1.00 | 13643.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | ClosedXML | 190.38 ms | 13.79 ms | 7.96 ms | 10.02 | 10.02 | 131493.2 KB | 9.64 |  |  | 901.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | EPPlus | 197.38 ms | 4.98 ms | 2.88 ms | 10.38 | 10.38 | 97646.6 KB | 7.16 |  |  | 938.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | OfficeIMO.Excel | 18.68 ms | 0.35 ms | 0.20 ms | 1.00 | 1.00 | 7184.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | ClosedXML | 124.43 ms | 6.83 ms | 3.94 ms | 6.66 | 6.66 | 84512.0 KB | 11.76 |  |  | 566.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | EPPlus | 134.78 ms | 8.84 ms | 5.11 ms | 7.22 | 7.22 | 69934.9 KB | 9.73 |  |  | 621.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | LargeXlsx | 43.24 ms | 2.27 ms | 1.31 ms | 0.86 | 1.00 | 5614.1 KB | 0.45 |  |  | 13.7% faster than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 50.10 ms | 0.66 ms | 0.38 ms | 1.00 | 1.16 | 12584.3 KB | 1.00 |  |  | Loss +15.9% |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | MiniExcel | 105.41 ms | 10.11 ms | 5.83 ms | 2.10 | 2.44 | 93257.1 KB | 7.41 |  |  | 110.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 277.80 ms |  |  | 5.55 | 6.42 |  |  |  |  | 454.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | EPPlus | 388.85 ms | 5.39 ms | 3.11 ms | 7.76 | 8.99 | 117437.6 KB | 9.33 |  |  | 676.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | ClosedXML | 408.35 ms | 5.28 ms | 3.05 ms | 8.15 | 9.44 | 104205.0 KB | 8.28 |  |  | 715.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | LargeXlsx | 43.06 ms | 8.13 ms | 4.69 ms | 0.81 | 1.00 | 5614.1 KB | 0.45 |  |  | 18.6% faster than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | OfficeIMO.Excel | 52.88 ms | 4.48 ms | 2.59 ms | 1.00 | 1.23 | 12585.9 KB | 1.00 |  |  | Loss +22.8% |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | MiniExcel | 104.11 ms | 1.28 ms | 0.74 ms | 1.97 | 2.42 | 93256.1 KB | 7.41 |  |  | 96.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | EPPlus 4.5.3.3 | 219.20 ms |  |  | 4.15 | 5.09 |  |  |  |  | 314.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | EPPlus | 404.99 ms | 23.22 ms | 13.41 ms | 7.66 | 9.40 | 117437.3 KB | 9.33 |  |  | 665.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | ClosedXML | 422.95 ms | 35.10 ms | 20.27 ms | 8.00 | 9.82 | 104205.0 KB | 8.28 |  |  | 699.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 52.79 ms | 11.10 ms | 6.41 ms | 0.76 | 1.00 | 5614.1 KB | 0.80 |  |  | 24.0% faster than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 69.43 ms | 13.25 ms | 7.65 ms | 1.00 | 1.32 | 7029.4 KB | 1.00 |  |  | Loss +31.5% |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | MiniExcel | 122.05 ms | 16.32 ms | 9.42 ms | 1.76 | 2.31 | 93256.8 KB | 13.27 |  |  | 75.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | ClosedXML | 451.72 ms | 31.73 ms | 18.32 ms | 6.51 | 8.56 | 104205.0 KB | 14.82 |  |  | 550.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | EPPlus | 510.98 ms | 68.93 ms | 39.80 ms | 7.36 | 9.68 | 117437.3 KB | 16.71 |  |  | 636.0% slower than OfficeIMO |
