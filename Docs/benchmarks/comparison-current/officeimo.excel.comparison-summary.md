# OfficeIMO.Excel Comparison Summary

This is the suite-level decision table. Mean, standard deviation, standard error, ratios, and allocations come from the lightweight rotated local runner; they are meant for engineering direction. Use the BenchmarkDotNet benchmark classes when a publication-grade `Error` column is required.

## At a glance

| Row count | Artifact | Workload | Category | OfficeIMO wins | OfficeIMO losses | Biggest loss |
| ---: | --- | --- | --- | ---: | ---: | --- |
| 2500 | dense-helloworld-comparison | read | Other | 0 | 2 | dense-helloworld-read-range: Loss +51.9% vs Sylvan.Data.Excel |
| 2500 | package-profile | package | Package size | 37 | 12 | write-datatable-direct: Loss +102.5% vs LargeXlsx |
| 2500 | speed-comparison | mutate | AutoFit and mutation | 1 | 0 |  |
| 2500 | speed-comparison | other | DataTable table export | 2 | 0 |  |
| 2500 | speed-comparison | other | Object projection | 2 | 0 |  |
| 2500 | speed-comparison | other | Other | 10 | 0 |  |
| 2500 | speed-comparison | other | Range and table read | 4 | 0 |  |
| 2500 | speed-comparison | read | Formula write/read | 1 | 0 |  |
| 2500 | speed-comparison | read | Other | 0 | 3 | large-sparse-row-read: Loss +139.6% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | read | Range and table read | 1 | 6 | read-top-range: Loss +201.9% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | read | Streaming read | 0 | 4 | read-top-range-stream-small-chunks: Loss +249.3% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | read | Typed object read | 0 | 2 | read-objects-stream: Loss +17.8% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | write | AutoFit and mutation | 5 | 0 |  |
| 2500 | speed-comparison | write | Cell writer | 7 | 0 |  |
| 2500 | speed-comparison | write | DataSet table export | 3 | 0 |  |
| 2500 | speed-comparison | write | DataTable table export | 4 | 0 |  |
| 2500 | speed-comparison | write | Formatted report write | 1 | 0 |  |
| 2500 | speed-comparison | write | Formula write/read | 1 | 0 |  |
| 2500 | speed-comparison | write | Other | 2 | 2 | write-powershell-psobject-mixed-direct: Loss +24.3% vs LargeXlsx |
| 2500 | speed-comparison | write | Plain cell export | 1 | 3 | append-plain-rows: Loss +80.6% vs LargeXlsx |
| 2500 | speed-comparison | write | Plain streaming export | 2 | 0 |  |
| 2500 | speed-comparison | write | Plain string export | 1 | 0 |  |
| 2500 | speed-comparison | write | Shared string write | 4 | 0 |  |
| 2500 | speed-comparison | write | Typed object export | 0 | 3 | write-insertobjects-flat-dictionaries-direct: Loss +47.8% vs LargeXlsx |
| 25000 | dense-helloworld-comparison | read | Other | 1 | 1 | dense-helloworld-read-stream: Loss +22.9% vs Sylvan.Data.Excel |
| 25000 | package-profile | package | Package size | 37 | 12 | write-insertobjects-legacy-dictionaries-direct: Loss +65.6% vs LargeXlsx |
| 25000 | speed-comparison | mutate | AutoFit and mutation | 1 | 0 |  |
| 25000 | speed-comparison | other | DataTable table export | 2 | 0 |  |
| 25000 | speed-comparison | other | Object projection | 2 | 0 |  |
| 25000 | speed-comparison | other | Other | 10 | 0 |  |
| 25000 | speed-comparison | other | Range and table read | 4 | 0 |  |
| 25000 | speed-comparison | read | Formula write/read | 1 | 0 |  |
| 25000 | speed-comparison | read | Other | 0 | 3 | large-sparse-row-read: Loss +74.3% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | read | Range and table read | 4 | 3 | read-top-range: Loss +319.8% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | read | Streaming read | 2 | 2 | read-top-range-stream: Loss +313.2% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | read | Typed object read | 0 | 2 | read-objects: Loss +5.3% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | write | AutoFit and mutation | 5 | 0 |  |
| 25000 | speed-comparison | write | Cell writer | 7 | 0 |  |
| 25000 | speed-comparison | write | DataSet table export | 3 | 0 |  |
| 25000 | speed-comparison | write | DataTable table export | 3 | 1 | write-datatable-direct: Loss +15.7% vs LargeXlsx |
| 25000 | speed-comparison | write | Formatted report write | 1 | 0 |  |
| 25000 | speed-comparison | write | Formula write/read | 1 | 0 |  |
| 25000 | speed-comparison | write | Other | 2 | 2 | write-powershell-mixed-objects-direct: Loss +16.8% vs LargeXlsx |
| 25000 | speed-comparison | write | Plain cell export | 2 | 2 | append-plain-rows: Loss +51.4% vs LargeXlsx |
| 25000 | speed-comparison | write | Plain streaming export | 1 | 1 | write-datareader-plain: Loss +27.5% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | write | Plain string export | 0 | 1 | write-blog-2023-20-string-columns: Loss +6.5% vs LargeXlsx |
| 25000 | speed-comparison | write | Shared string write | 4 | 0 |  |
| 25000 | speed-comparison | write | Typed object export | 0 | 3 | write-insertobjects-flat-dictionaries-direct: Loss +35.7% vs LargeXlsx |

## OfficeIMO decision table

| Row count | Artifact | Workload | Category | Scenario | OfficeIMO mean | Best | OfficeIMO vs best | Alloc | Package |
| ---: | --- | --- | --- | --- | ---: | --- | ---: | ---: | ---: |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | 6.14 ms | Sylvan.Data.Excel | Loss +51.9% | 2488.6 KB |  |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | 5.54 ms | Sylvan.Data.Excel | Loss +32.4% | 2566.9 KB |  |
| 2500 | package-profile | package | Package size | append-plain-rows | 2.45 ms | LargeXlsx | Loss +60.3% | 1657.3 KB | 64.5 KB |
| 2500 | package-profile | package | Package size | autofit-existing | 44.54 ms | OfficeIMO.Excel | Win | 13971.2 KB | 143.2 KB |
| 2500 | package-profile | package | Package size | large-shared-strings | 2.01 ms | OfficeIMO.Excel | Win | 2111.9 KB | 55.2 KB |
| 2500 | package-profile | package | Package size | realworld-autofilter | 3.57 ms | OfficeIMO.Excel | Win | 1171.9 KB | 142.4 KB |
| 2500 | package-profile | package | Package size | realworld-charts | 5.25 ms | OfficeIMO.Excel | Win | 1688.9 KB | 147.5 KB |
| 2500 | package-profile | package | Package size | realworld-conditional-formatting | 4.40 ms | OfficeIMO.Excel | Win | 1236.8 KB | 142.7 KB |
| 2500 | package-profile | package | Package size | realworld-data-validation | 3.54 ms | OfficeIMO.Excel | Win | 1187.6 KB | 142.5 KB |
| 2500 | package-profile | package | Package size | realworld-freeze-panes | 3.86 ms | OfficeIMO.Excel | Win | 1174.0 KB | 142.5 KB |
| 2500 | package-profile | package | Package size | realworld-pivot-table | 20.98 ms | OfficeIMO.Excel | Win | 18628.0 KB | 203.9 KB |
| 2500 | package-profile | package | Package size | realworld-report-all-in-one | 18.78 ms | OfficeIMO.Excel | Win | 19412.8 KB | 210.0 KB |
| 2500 | package-profile | package | Package size | realworld-report-core | 4.52 ms | OfficeIMO.Excel | Win | 1319.1 KB | 143.9 KB |
| 2500 | package-profile | package | Package size | report-workbook | 14.44 ms | OfficeIMO.Excel | Win | 12228.4 KB | 90.2 KB |
| 2500 | package-profile | package | Package size | report-workbook-core | 6.18 ms | OfficeIMO.Excel | Win | 2375.5 KB | 187.5 KB |
| 2500 | package-profile | package | Package size | report-workbook-datatable | 16.45 ms | OfficeIMO.Excel | Win | 12500.1 KB | 90.2 KB |
| 2500 | package-profile | package | Package size | report-workbook-datatable-core | 6.33 ms | OfficeIMO.Excel | Win | 2647.2 KB | 187.5 KB |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | 5.00 ms | LargeXlsx | Loss +14.3% | 1508.4 KB | 216.7 KB |
| 2500 | package-profile | package | Package size | write-bulk-report | 4.78 ms | OfficeIMO.Excel | Win | 1233.1 KB | 143.2 KB |
| 2500 | package-profile | package | Package size | write-cellformula | 2.27 ms | OfficeIMO.Excel | Win | 1171.8 KB | 66.6 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-empty-strings | 2.40 ms | OfficeIMO.Excel | Win | 1454.0 KB | 44.2 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-numbers | 2.01 ms | OfficeIMO.Excel | Win | 946.8 KB | 47.0 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-object-mixed | 3.14 ms | OfficeIMO.Excel | Win | 1431.4 KB | 61.2 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse | 2.60 ms | OfficeIMO.Excel | Win | 1271.0 KB | 62.0 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | 2.27 ms | OfficeIMO.Excel | Win | 1271.1 KB | 62.0 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-scalars | 1.85 ms | OfficeIMO.Excel | Win | 964.9 KB | 46.3 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-strings | 2.87 ms | OfficeIMO.Excel | Win | 2283.8 KB | 55.1 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-distinct | 2.39 ms | OfficeIMO.Excel | Win | 2206.1 KB | 51.8 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-repeated | 2.26 ms | OfficeIMO.Excel | Win | 1246.7 KB | 40.0 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-temporal | 2.49 ms | OfficeIMO.Excel | Win | 1262.6 KB | 63.3 KB |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | 1.79 ms | LargeXlsx | Loss +45.7% | 923.6 KB | 48.2 KB |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | 4.15 ms | LargeXlsx | Loss +27.3% | 1752.7 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-datareader-plain | 4.76 ms | Sylvan.Data.Excel | Loss +10.2% | 1434.6 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-datareader-table | 7.12 ms | OfficeIMO.Excel | Win | 1446.4 KB | 138.8 KB |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | 4.84 ms | OfficeIMO.Excel | Win | 1452.7 KB | 138.9 KB |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | 4.42 ms | LargeXlsx | Loss +11.7% | 1652.6 KB | 131.1 KB |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | 4.90 ms | OfficeIMO.Excel | Win | 2392.1 KB | 176.0 KB |
| 2500 | package-profile | package | Package size | write-dataset-tables | 3.86 ms | OfficeIMO.Excel | Win | 1578.0 KB | 138.9 KB |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | 4.73 ms | OfficeIMO.Excel | Win | 1590.6 KB | 139.2 KB |
| 2500 | package-profile | package | Package size | write-datatable-direct | 7.16 ms | LargeXlsx | Loss +102.5% | 1420.2 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | 4.56 ms | OfficeIMO.Excel | Win | 1432.2 KB | 138.8 KB |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | 4.52 ms | LargeXlsx | Loss +33.9% | 1440.7 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | 4.99 ms | OfficeIMO.Excel | Win | 1178.6 KB | 142.4 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | 4.02 ms | LargeXlsx | Loss +8.7% | 1170.9 KB | 142.3 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | 5.71 ms | OfficeIMO.Excel | Win | 1176.8 KB | 138.1 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | 4.18 ms | LargeXlsx | Loss +36.9% | 1169.1 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | 5.99 ms | LargeXlsx | Loss +72.3% | 1601.6 KB | 142.3 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | 3.68 ms | OfficeIMO.Excel | Win | 1177.9 KB | 142.4 KB |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | 5.97 ms | LargeXlsx | Loss +47.6% | 2013.3 KB | 183.1 KB |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | 4.46 ms | OfficeIMO.Excel | Win | 1339.3 KB | 182.4 KB |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | 20.71 ms | OfficeIMO.Excel | Win | 4333.9 KB | 651.0 KB |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | 28.07 ms | OfficeIMO.Excel | Win | 13971.2 KB |  |
| 2500 | speed-comparison | other | DataTable table export | report-workbook-datatable | 13.57 ms | OfficeIMO.Excel | Win | 12500.1 KB |  |
| 2500 | speed-comparison | other | DataTable table export | report-workbook-datatable-core | 5.55 ms | OfficeIMO.Excel | Win | 2647.2 KB |  |
| 2500 | speed-comparison | other | Object projection | build-object-datatable-dictionaries | 1.57 ms | OfficeIMO.Excel | Win | 564.2 KB |  |
| 2500 | speed-comparison | other | Object projection | build-object-datatable-typed | 1.34 ms | OfficeIMO.Excel | Win | 856.9 KB |  |
| 2500 | speed-comparison | other | Other | realworld-autofilter | 3.69 ms | OfficeIMO.Excel | Win | 1171.9 KB |  |
| 2500 | speed-comparison | other | Other | realworld-charts | 5.01 ms | OfficeIMO.Excel | Win | 1689.0 KB |  |
| 2500 | speed-comparison | other | Other | realworld-conditional-formatting | 3.93 ms | OfficeIMO.Excel | Win | 1236.8 KB |  |
| 2500 | speed-comparison | other | Other | realworld-data-validation | 3.67 ms | OfficeIMO.Excel | Win | 1187.6 KB |  |
| 2500 | speed-comparison | other | Other | realworld-freeze-panes | 3.46 ms | OfficeIMO.Excel | Win | 1174.1 KB |  |
| 2500 | speed-comparison | other | Other | realworld-pivot-table | 16.56 ms | OfficeIMO.Excel | Win | 18627.9 KB |  |
| 2500 | speed-comparison | other | Other | realworld-report-all-in-one | 18.59 ms | OfficeIMO.Excel | Win | 19412.9 KB |  |
| 2500 | speed-comparison | other | Other | realworld-report-core | 4.15 ms | OfficeIMO.Excel | Win | 1319.1 KB |  |
| 2500 | speed-comparison | other | Other | report-workbook | 15.21 ms | OfficeIMO.Excel | Win | 12228.4 KB |  |
| 2500 | speed-comparison | other | Other | report-workbook-core | 5.79 ms | OfficeIMO.Excel | Win | 2375.5 KB |  |
| 2500 | speed-comparison | other | Range and table read | enumerate-cells | 7.92 ms | OfficeIMO.Excel | Win | 2649.2 KB |  |
| 2500 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | 4.83 ms | OfficeIMO.Excel | Win | 643.4 KB |  |
| 2500 | speed-comparison | other | Range and table read | enumerate-range | 8.36 ms | OfficeIMO.Excel | Win | 2649.2 KB |  |
| 2500 | speed-comparison | other | Range and table read | enumerate-top-range | 1.69 ms | OfficeIMO.Excel | Win | 402.6 KB |  |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | 3.33 ms | OfficeIMO.Excel | Win | 777.4 KB |  |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | 1.73 ms | Sylvan.Data.Excel | Loss +72.4% | 248.8 KB |  |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | 2.38 ms | Sylvan.Data.Excel | Loss +139.6% | 248.9 KB |  |
| 2500 | speed-comparison | read | Other | shared-string-read | 3.47 ms | Sylvan.Data.Excel | Loss +67.1% | 1133.5 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | 4.98 ms | Sylvan.Data.Excel | Loss +12.4% | 494.6 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-datatable | 8.48 ms | Sylvan.Data.Excel | Loss +33.3% | 3714.5 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | 4.93 ms | Sylvan.Data.Excel | Loss +2.7% | 663.0 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-range | 11.83 ms | OfficeIMO.Excel | Win | 2812.6 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | 8.17 ms | Sylvan.Data.Excel | Loss +44.6% | 2871.4 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-top-range | 1.56 ms | Sylvan.Data.Excel | Loss +201.9% | 416.1 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-used-range | 14.55 ms | Sylvan.Data.Excel | Loss +191.0% | 3535.1 KB |  |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | 4.93 ms | Sylvan.Data.Excel | Loss +12.7% | 497.9 KB |  |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | 8.69 ms | Sylvan.Data.Excel | Loss +76.0% | 2891.5 KB |  |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | 1.49 ms | Sylvan.Data.Excel | Loss +241.8% | 419.5 KB |  |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | 1.60 ms | Sylvan.Data.Excel | Loss +249.3% | 420.2 KB |  |
| 2500 | speed-comparison | read | Typed object read | read-objects | 7.90 ms | Sylvan.Data.Excel | Loss +16.2% | 2562.1 KB |  |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | 5.81 ms | Sylvan.Data.Excel | Loss +17.8% | 2562.4 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | 5.40 ms | OfficeIMO.Excel | Win | 1452.7 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | 6.35 ms | OfficeIMO.Excel | Win | 1591.1 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | 4.42 ms | OfficeIMO.Excel | Win | 1178.6 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | 4.45 ms | OfficeIMO.Excel | Win | 1176.8 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | 4.23 ms | OfficeIMO.Excel | Win | 1177.9 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | 2.91 ms | OfficeIMO.Excel | Win | 1454.0 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | 2.98 ms | OfficeIMO.Excel | Win | 946.8 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | 2.90 ms | OfficeIMO.Excel | Win | 1431.4 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | 4.69 ms | OfficeIMO.Excel | Win | 1271.0 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | 3.26 ms | OfficeIMO.Excel | Win | 1271.1 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | 2.65 ms | OfficeIMO.Excel | Win | 964.9 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | 3.28 ms | OfficeIMO.Excel | Win | 1262.6 KB |  |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | 5.53 ms | OfficeIMO.Excel | Win | 1576.3 KB |  |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | 8.58 ms | OfficeIMO.Excel | Win | 2392.1 KB |  |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | 5.44 ms | OfficeIMO.Excel | Win | 1579.6 KB |  |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | 4.40 ms | OfficeIMO.Excel | Win | 1446.4 KB |  |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | 5.02 ms | OfficeIMO.Excel | Win | 1420.2 KB |  |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | 5.05 ms | OfficeIMO.Excel | Win | 1158.8 KB |  |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | 4.82 ms | OfficeIMO.Excel | Win | 1432.2 KB |  |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | 7.21 ms | OfficeIMO.Excel | Win | 1234.7 KB |  |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | 2.81 ms | OfficeIMO.Excel | Win | 1409.1 KB |  |
| 2500 | speed-comparison | write | Other | write-dictionary-objects-table-direct | 7.91 ms | OfficeIMO.Excel | Win | 1723.0 KB |  |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | 5.61 ms | LargeXlsx | Loss +22.5% | 2013.3 KB |  |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | 5.39 ms | LargeXlsx | Loss +24.3% | 1339.3 KB |  |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | 20.40 ms | OfficeIMO.Excel | Win | 4333.9 KB |  |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | 2.47 ms | LargeXlsx | Loss +80.6% | 1657.3 KB |  |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | 1.66 ms | LargeXlsx | Loss +28.0% | 923.6 KB |  |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | 4.64 ms | LargeXlsx | Loss +27.0% | 1752.7 KB |  |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | 2.43 ms | OfficeIMO.Excel | Win | 1165.6 KB |  |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | 5.34 ms | OfficeIMO.Excel | Win | 1434.6 KB |  |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | 5.41 ms | OfficeIMO.Excel | Win | 1652.6 KB |  |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | 5.04 ms | OfficeIMO.Excel | Win | 1508.4 KB |  |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | 2.09 ms | OfficeIMO.Excel | Win | 2111.9 KB |  |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | 2.99 ms | OfficeIMO.Excel | Win | 2283.8 KB |  |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | 3.11 ms | OfficeIMO.Excel | Win | 2206.1 KB |  |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | 2.57 ms | OfficeIMO.Excel | Win | 1246.7 KB |  |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | 3.60 ms | LargeXlsx | Loss +22.1% | 1440.7 KB |  |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | 4.60 ms | LargeXlsx | Loss +31.5% | 1170.9 KB |  |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | 4.66 ms | LargeXlsx | Loss +47.8% | 1169.1 KB |  |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | 60.77 ms | OfficeIMO.Excel, Sylvan.Data.Excel | Win | 23699.6 KB |  |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | 58.28 ms | Sylvan.Data.Excel | Loss +22.9% | 24481.9 KB |  |
| 25000 | package-profile | package | Package size | append-plain-rows | 20.87 ms | LargeXlsx | Loss +48.1% | 11671.6 KB | 622.5 KB |
| 25000 | package-profile | package | Package size | autofit-existing | 342.58 ms | OfficeIMO.Excel | Win | 138378.6 KB | 1429.3 KB |
| 25000 | package-profile | package | Package size | large-shared-strings | 14.99 ms | OfficeIMO.Excel | Win | 15416.5 KB | 529.7 KB |
| 25000 | package-profile | package | Package size | realworld-autofilter | 30.84 ms | OfficeIMO.Excel | Win | 11326.3 KB | 1428.5 KB |
| 25000 | package-profile | package | Package size | realworld-charts | 41.92 ms | OfficeIMO.Excel | Win | 12323.4 KB | 1433.6 KB |
| 25000 | package-profile | package | Package size | realworld-conditional-formatting | 30.63 ms | OfficeIMO.Excel | Win | 11391.3 KB | 1428.8 KB |
| 25000 | package-profile | package | Package size | realworld-data-validation | 39.78 ms | OfficeIMO.Excel | Win | 11342.1 KB | 1428.6 KB |
| 25000 | package-profile | package | Package size | realworld-freeze-panes | 31.47 ms | OfficeIMO.Excel | Win | 11331.2 KB | 1428.6 KB |
| 25000 | package-profile | package | Package size | realworld-pivot-table | 117.92 ms | OfficeIMO.Excel | Win | 100856.7 KB | 1431.1 KB |
| 25000 | package-profile | package | Package size | realworld-report-all-in-one | 117.37 ms | OfficeIMO.Excel | Win | 102341.4 KB | 1437.1 KB |
| 25000 | package-profile | package | Package size | realworld-report-core | 38.60 ms | OfficeIMO.Excel | Win | 11487.6 KB | 1430.0 KB |
| 25000 | package-profile | package | Package size | report-workbook | 48.14 ms | OfficeIMO.Excel | Win | 14316.9 KB | 1857.1 KB |
| 25000 | package-profile | package | Package size | report-workbook-core | 55.61 ms | OfficeIMO.Excel | Win | 10803.9 KB | 1850.9 KB |
| 25000 | package-profile | package | Package size | report-workbook-datatable | 148.11 ms | OfficeIMO.Excel | Win | 17067.8 KB | 1857.1 KB |
| 25000 | package-profile | package | Package size | report-workbook-datatable-core | 61.19 ms | OfficeIMO.Excel | Win | 13557.9 KB | 1850.9 KB |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | 43.48 ms | LargeXlsx | Loss +10.6% | 11539.8 KB | 2228.8 KB |
| 25000 | package-profile | package | Package size | write-bulk-report | 34.49 ms | OfficeIMO.Excel | Win | 11393.2 KB | 1429.3 KB |
| 25000 | package-profile | package | Package size | write-cellformula | 25.87 ms | OfficeIMO.Excel | Win | 9557.4 KB | 670.3 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-empty-strings | 12.62 ms | OfficeIMO.Excel | Win | 6731.3 KB | 451.4 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-numbers | 15.94 ms | OfficeIMO.Excel | Win | 5805.8 KB | 462.6 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-object-mixed | 18.84 ms | OfficeIMO.Excel | Win | 8009.0 KB | 585.1 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse | 18.93 ms | OfficeIMO.Excel | Win | 7188.3 KB | 607.1 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | 20.86 ms | OfficeIMO.Excel | Win | 7188.4 KB | 607.1 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-scalars | 11.71 ms | OfficeIMO.Excel | Win | 5979.5 KB | 441.9 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-strings | 18.58 ms | OfficeIMO.Excel | Win | 15027.2 KB | 527.8 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-distinct | 14.93 ms | OfficeIMO.Excel | Win | 13659.0 KB | 499.0 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-repeated | 14.38 ms | OfficeIMO.Excel | Win | 7197.6 KB | 376.0 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-temporal | 20.22 ms | OfficeIMO.Excel | Win | 7317.8 KB | 620.5 KB |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | 11.34 ms | LargeXlsx | Loss +21.1% | 6793.3 KB | 455.5 KB |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | 35.01 ms | LargeXlsx | Loss +29.2% | 15708.0 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-datareader-plain | 37.20 ms | Sylvan.Data.Excel | Loss +31.0% | 12673.9 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-datareader-table | 34.36 ms | OfficeIMO.Excel | Win | 12691.9 KB | 1385.7 KB |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | 38.07 ms | OfficeIMO.Excel | Win | 12698.2 KB | 1385.9 KB |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | 31.59 ms | OfficeIMO.Excel | Win | 9491.6 KB | 1329.2 KB |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | 40.66 ms | OfficeIMO.Excel | Win | 13130.4 KB | 1795.1 KB |
| 25000 | package-profile | package | Package size | write-dataset-tables | 32.65 ms | OfficeIMO.Excel | Win | 9800.0 KB | 1376.4 KB |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | 36.92 ms | OfficeIMO.Excel | Win | 9812.6 KB | 1376.7 KB |
| 25000 | package-profile | package | Package size | write-datatable-direct | 33.29 ms | LargeXlsx | Loss +9.2% | 12387.3 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | 33.76 ms | OfficeIMO.Excel | Win | 12405.4 KB | 1385.7 KB |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | 49.21 ms | LargeXlsx | Loss +29.7% | 12583.6 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | 60.06 ms | OfficeIMO.Excel | Win | 11341.1 KB | 1428.5 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | 39.60 ms | LargeXlsx | Loss +16.8% | 11333.4 KB | 1428.4 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | 41.82 ms | OfficeIMO.Excel | Win | 9866.7 KB | 1385.1 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | 40.79 ms | LargeXlsx | Loss +33.0% | 9859.0 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | 46.52 ms | LargeXlsx | Loss +65.6% | 15631.3 KB | 1428.4 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | 36.82 ms | OfficeIMO.Excel | Win | 11340.4 KB | 1428.5 KB |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | 58.38 ms | LargeXlsx | Loss +19.0% | 10416.8 KB | 1828.0 KB |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | 52.24 ms | LargeXlsx | Loss +8.4% | 9781.8 KB | 1828.0 KB |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | 213.38 ms | OfficeIMO.Excel | Win | 35984.4 KB | 6725.6 KB |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | 345.37 ms | OfficeIMO.Excel | Win | 138378.6 KB |  |
| 25000 | speed-comparison | other | DataTable table export | report-workbook-datatable | 50.01 ms | OfficeIMO.Excel | Win | 17062.3 KB |  |
| 25000 | speed-comparison | other | DataTable table export | report-workbook-datatable-core | 47.36 ms | OfficeIMO.Excel | Win | 13549.5 KB |  |
| 25000 | speed-comparison | other | Object projection | build-object-datatable-dictionaries | 9.59 ms | OfficeIMO.Excel | Win | 5164.3 KB |  |
| 25000 | speed-comparison | other | Object projection | build-object-datatable-typed | 7.00 ms | OfficeIMO.Excel | Win | 8093.8 KB |  |
| 25000 | speed-comparison | other | Other | realworld-autofilter | 31.29 ms | OfficeIMO.Excel | Win | 11326.3 KB |  |
| 25000 | speed-comparison | other | Other | realworld-charts | 32.35 ms | OfficeIMO.Excel | Win | 12323.0 KB |  |
| 25000 | speed-comparison | other | Other | realworld-conditional-formatting | 30.43 ms | OfficeIMO.Excel | Win | 11391.3 KB |  |
| 25000 | speed-comparison | other | Other | realworld-data-validation | 30.51 ms | OfficeIMO.Excel | Win | 11342.1 KB |  |
| 25000 | speed-comparison | other | Other | realworld-freeze-panes | 30.76 ms | OfficeIMO.Excel | Win | 11328.5 KB |  |
| 25000 | speed-comparison | other | Other | realworld-pivot-table | 88.29 ms | OfficeIMO.Excel | Win | 100855.7 KB |  |
| 25000 | speed-comparison | other | Other | realworld-report-all-in-one | 98.88 ms | OfficeIMO.Excel | Win | 102332.3 KB |  |
| 25000 | speed-comparison | other | Other | realworld-report-core | 34.16 ms | OfficeIMO.Excel | Win | 11479.2 KB |  |
| 25000 | speed-comparison | other | Other | report-workbook | 47.84 ms | OfficeIMO.Excel | Win | 14317.1 KB |  |
| 25000 | speed-comparison | other | Other | report-workbook-core | 44.99 ms | OfficeIMO.Excel | Win | 10803.9 KB |  |
| 25000 | speed-comparison | other | Range and table read | enumerate-cells | 49.58 ms | OfficeIMO.Excel | Win | 24648.3 KB |  |
| 25000 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | 35.38 ms | OfficeIMO.Excel | Win | 3959.5 KB |  |
| 25000 | speed-comparison | other | Range and table read | enumerate-range | 48.79 ms | OfficeIMO.Excel | Win | 24648.4 KB |  |
| 25000 | speed-comparison | other | Range and table read | enumerate-top-range | 1.90 ms | OfficeIMO.Excel | Win | 402.7 KB |  |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | 21.56 ms | OfficeIMO.Excel | Win | 6287.0 KB |  |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | 1.61 ms | Sylvan.Data.Excel | Loss +67.8% | 248.8 KB |  |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | 1.67 ms | Sylvan.Data.Excel | Loss +74.3% | 248.9 KB |  |
| 25000 | speed-comparison | read | Other | shared-string-read | 20.74 ms | Sylvan.Data.Excel | Loss +18.0% | 9295.0 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | 34.26 ms | OfficeIMO.Excel | Win | 1242.4 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-datatable | 56.83 ms | OfficeIMO.Excel | Win | 34766.1 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | 36.37 ms | OfficeIMO.Excel | Win | 4154.6 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-range | 53.24 ms | Sylvan.Data.Excel | Loss +3.8% | 26218.6 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | 53.65 ms | OfficeIMO.Excel | Win | 26804.5 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-top-range | 1.90 ms | Sylvan.Data.Excel | Loss +319.8% | 416.1 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-used-range | 86.61 ms | Sylvan.Data.Excel | Loss +79.7% | 34214.6 KB |  |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | 36.68 ms | OfficeIMO.Excel | Win | 1245.8 KB |  |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | 45.91 ms | OfficeIMO.Excel | Win | 27005.5 KB |  |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | 1.81 ms | Sylvan.Data.Excel | Loss +313.2% | 419.5 KB |  |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | 1.91 ms | Sylvan.Data.Excel | Loss +305.3% | 420.2 KB |  |
| 25000 | speed-comparison | read | Typed object read | read-objects | 49.48 ms | Sylvan.Data.Excel | Loss +5.3% | 23682.5 KB |  |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | 46.73 ms | Sylvan.Data.Excel | Loss +2.1% | 23682.8 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | 38.06 ms | OfficeIMO.Excel | Win | 12698.2 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | 41.98 ms | OfficeIMO.Excel | Win | 9812.6 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | 36.60 ms | OfficeIMO.Excel | Win | 11333.1 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | 41.65 ms | OfficeIMO.Excel | Win | 9858.7 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | 33.99 ms | OfficeIMO.Excel | Win | 11332.4 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | 12.92 ms | OfficeIMO.Excel | Win | 6723.3 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | 16.92 ms | OfficeIMO.Excel | Win | 5797.7 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | 20.04 ms | OfficeIMO.Excel | Win | 8001.0 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | 20.10 ms | OfficeIMO.Excel | Win | 7180.3 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | 16.76 ms | OfficeIMO.Excel | Win | 7180.4 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | 11.89 ms | OfficeIMO.Excel | Win | 5971.5 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | 20.70 ms | OfficeIMO.Excel | Win | 7309.7 KB |  |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | 36.09 ms | OfficeIMO.Excel | Win | 12551.5 KB |  |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | 45.39 ms | OfficeIMO.Excel | Win | 13130.4 KB |  |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | 41.49 ms | OfficeIMO.Excel | Win | 9800.0 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | 36.75 ms | OfficeIMO.Excel | Win | 12691.9 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | 43.95 ms | LargeXlsx | Loss +15.7% | 12387.3 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | 35.92 ms | OfficeIMO.Excel | Win | 9673.7 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | 40.36 ms | OfficeIMO.Excel | Win | 12410.7 KB |  |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | 43.86 ms | OfficeIMO.Excel | Win | 11393.2 KB |  |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | 21.20 ms | OfficeIMO.Excel | Win | 9548.9 KB |  |
| 25000 | speed-comparison | write | Other | write-dictionary-objects-table-direct | 46.09 ms | OfficeIMO.Excel | Win | 14835.3 KB |  |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | 43.73 ms | LargeXlsx | Loss +16.8% | 10408.8 KB |  |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | 43.61 ms | LargeXlsx | Loss +11.2% | 9773.8 KB |  |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | 188.84 ms | OfficeIMO.Excel | Win | 35981.7 KB |  |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | 18.02 ms | LargeXlsx | Loss +51.4% | 11671.6 KB |  |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | 15.73 ms | OfficeIMO.Excel | Win | 6801.3 KB |  |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | 41.96 ms | LargeXlsx | Loss +37.7% | 15708.0 KB |  |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | 33.10 ms | OfficeIMO.Excel | Win | 7543.0 KB |  |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | 35.24 ms | Sylvan.Data.Excel | Loss +27.5% | 12673.9 KB |  |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | 40.02 ms | OfficeIMO.Excel | Win | 9491.6 KB |  |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | 42.64 ms | LargeXlsx | Loss +6.5% | 11547.8 KB |  |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | 15.26 ms | OfficeIMO.Excel | Win | 15416.5 KB |  |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | 18.65 ms | OfficeIMO.Excel | Win | 15027.2 KB |  |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | 13.14 ms | OfficeIMO.Excel | Win | 13651.0 KB |  |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | 16.32 ms | OfficeIMO.Excel | Win | 7192.2 KB |  |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | 33.11 ms | LargeXlsx | Loss +16.8% | 12583.6 KB |  |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | 32.57 ms | LargeXlsx | Loss +14.2% | 11325.4 KB |  |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | 38.73 ms | LargeXlsx | Loss +35.7% | 9851.0 KB |  |

## Full comparison table

| Row count | Artifact | Workload | Category | Scenario | Library | Mean | StdDev | StdErr | Ratio to OfficeIMO | Ratio to best | Alloc | Alloc ratio | Package | Package ratio | Outcome |
| ---: | --- | --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | Sylvan.Data.Excel | 4.04 ms | 0.17 ms | 0.10 ms | 0.66 | 1.00 | 362.3 KB | 0.15 |  |  | 34.2% faster than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | OfficeIMO.Excel | 6.14 ms | 0.35 ms | 0.20 ms | 1.00 | 1.52 | 2488.6 KB | 1.00 |  |  | Loss +51.9% |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | ExcelDataReader | 12.37 ms | 1.37 ms | 0.79 ms | 2.02 | 3.06 | 6874.1 KB | 2.76 |  |  | 101.6% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | MiniExcel | 15.37 ms | 1.17 ms | 0.67 ms | 2.51 | 3.81 | 21502.8 KB | 8.64 |  |  | 150.6% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | Sylvan.Data.Excel | 4.19 ms | 0.48 ms | 0.28 ms | 0.76 | 1.00 | 362.3 KB | 0.14 |  |  | 24.5% faster than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | OfficeIMO.Excel | 5.54 ms | 0.09 ms | 0.05 ms | 1.00 | 1.32 | 2566.9 KB | 1.00 |  |  | Loss +32.4% |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | ExcelDataReader | 11.27 ms | 0.19 ms | 0.11 ms | 2.03 | 2.69 | 6874.1 KB | 2.68 |  |  | 103.4% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | MiniExcel | 14.82 ms | 0.88 ms | 0.51 ms | 2.67 | 3.54 | 21502.8 KB | 8.38 |  |  | 167.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | append-plain-rows | LargeXlsx | 1.53 ms | 0.19 ms | 0.11 ms | 0.62 | 1.00 | 288.4 KB | 0.17 | 63.1 KB | 0.98 | 37.6% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | append-plain-rows | OfficeIMO.Excel | 2.45 ms | 0.29 ms | 0.17 ms | 1.00 | 1.60 | 1657.3 KB | 1.00 | 64.5 KB | 1.00 | Loss +60.3% |
| 2500 | package-profile | package | Package size | append-plain-rows | MiniExcel | 4.10 ms | 0.11 ms | 0.06 ms | 1.67 | 2.68 | 19701.7 KB | 11.89 | 68.1 KB | 1.06 | 67.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | append-plain-rows | ClosedXML | 15.92 ms | 1.97 ms | 1.14 ms | 6.50 | 10.42 | 11189.4 KB | 6.75 | 59.8 KB | 0.93 | 550.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | append-plain-rows | EPPlus | 28.54 ms | 3.26 ms | 1.88 ms | 11.65 | 18.68 | 14283.5 KB | 8.62 | 56.9 KB | 0.88 | 1065.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | autofit-existing | OfficeIMO.Excel | 44.54 ms | 21.22 ms | 12.25 ms | 1.00 | 1.00 | 13971.2 KB | 1.00 | 143.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | autofit-existing | EPPlus | 99.96 ms | 15.05 ms | 8.69 ms | 2.24 | 2.24 | 50639.2 KB | 3.62 | 115.0 KB | 0.80 | 124.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | autofit-existing | ClosedXML | 133.33 ms | 4.84 ms | 2.79 ms | 2.99 | 2.99 | 84424.1 KB | 6.04 | 121.0 KB | 0.84 | 199.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | large-shared-strings | OfficeIMO.Excel | 2.01 ms | 0.06 ms | 0.03 ms | 1.00 | 1.00 | 2111.9 KB | 1.00 | 55.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | large-shared-strings | MiniExcel | 4.39 ms | 0.55 ms | 0.32 ms | 2.18 | 2.18 | 21128.5 KB | 10.00 | 60.7 KB | 1.10 | 118.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | large-shared-strings | ClosedXML | 11.82 ms | 0.96 ms | 0.56 ms | 5.88 | 5.88 | 11291.2 KB | 5.35 | 50.3 KB | 0.91 | 487.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | large-shared-strings | EPPlus | 21.40 ms | 1.00 ms | 0.58 ms | 10.63 | 10.63 | 12730.2 KB | 6.03 | 48.1 KB | 0.87 | 963.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-autofilter | OfficeIMO.Excel | 3.57 ms | 0.16 ms | 0.09 ms | 1.00 | 1.00 | 1171.9 KB | 1.00 | 142.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-autofilter | ClosedXML | 32.58 ms | 2.93 ms | 1.69 ms | 9.11 | 9.11 | 22218.8 KB | 18.96 | 120.2 KB | 0.84 | 811.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-autofilter | EPPlus | 53.85 ms | 5.21 ms | 3.01 ms | 15.06 | 15.06 | 24647.8 KB | 21.03 | 114.2 KB | 0.80 | 1406.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-charts | OfficeIMO.Excel | 5.25 ms | 0.84 ms | 0.49 ms | 1.00 | 1.00 | 1688.9 KB | 1.00 | 147.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-charts | EPPlus | 48.96 ms | 2.15 ms | 1.24 ms | 9.33 | 9.33 | 27072.7 KB | 16.03 | 117.0 KB | 0.79 | 832.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-conditional-formatting | OfficeIMO.Excel | 4.40 ms | 0.82 ms | 0.47 ms | 1.00 | 1.00 | 1236.8 KB | 1.00 | 142.7 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-conditional-formatting | ClosedXML | 33.31 ms | 1.74 ms | 1.01 ms | 7.58 | 7.58 | 22265.8 KB | 18.00 | 120.3 KB | 0.84 | 657.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-conditional-formatting | EPPlus | 54.02 ms | 2.91 ms | 1.68 ms | 12.28 | 12.28 | 24690.0 KB | 19.96 | 114.3 KB | 0.80 | 1128.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-data-validation | OfficeIMO.Excel | 3.54 ms | 0.19 ms | 0.11 ms | 1.00 | 1.00 | 1187.6 KB | 1.00 | 142.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-data-validation | ClosedXML | 31.64 ms | 1.94 ms | 1.12 ms | 8.94 | 8.94 | 22239.9 KB | 18.73 | 120.3 KB | 0.84 | 793.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-data-validation | EPPlus | 47.78 ms | 0.99 ms | 0.57 ms | 13.49 | 13.49 | 24632.9 KB | 20.74 | 114.2 KB | 0.80 | 1249.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-freeze-panes | OfficeIMO.Excel | 3.86 ms | 0.07 ms | 0.04 ms | 1.00 | 1.00 | 1174.0 KB | 1.00 | 142.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-freeze-panes | ClosedXML | 36.17 ms | 5.61 ms | 3.24 ms | 9.37 | 9.37 | 22214.0 KB | 18.92 | 120.2 KB | 0.84 | 836.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-freeze-panes | EPPlus | 53.22 ms | 8.27 ms | 4.77 ms | 13.78 | 13.78 | 24660.4 KB | 21.01 | 114.3 KB | 0.80 | 1278.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-pivot-table | OfficeIMO.Excel | 20.98 ms | 5.81 ms | 3.36 ms | 1.00 | 1.00 | 18628.0 KB | 1.00 | 203.9 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-pivot-table | EPPlus | 56.12 ms | 2.54 ms | 1.46 ms | 2.68 | 2.68 | 29469.7 KB | 1.58 | 117.4 KB | 0.58 | 167.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-all-in-one | OfficeIMO.Excel | 18.78 ms | 0.49 ms | 0.28 ms | 1.00 | 1.00 | 19412.8 KB | 1.00 | 210.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-all-in-one | EPPlus | 65.51 ms | 1.89 ms | 1.09 ms | 3.49 | 3.49 | 54522.6 KB | 2.81 | 121.8 KB | 0.58 | 248.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-core | OfficeIMO.Excel | 4.52 ms | 0.51 ms | 0.29 ms | 1.00 | 1.00 | 1319.1 KB | 1.00 | 143.9 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-core | EPPlus | 68.92 ms | 2.10 ms | 1.21 ms | 15.23 | 15.23 | 47227.1 KB | 35.80 | 115.5 KB | 0.80 | 1423.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-core | ClosedXML | 88.23 ms | 5.30 ms | 3.06 ms | 19.50 | 19.50 | 69825.5 KB | 52.93 | 121.5 KB | 0.84 | 1849.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook | OfficeIMO.Excel | 14.44 ms | 1.53 ms | 0.88 ms | 1.00 | 1.00 | 12228.4 KB | 1.00 | 90.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | report-workbook | EPPlus | 100.13 ms | 0.51 ms | 0.30 ms | 6.93 | 6.93 | 77408.8 KB | 6.33 | 161.8 KB | 1.79 | 593.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook-core | OfficeIMO.Excel | 6.18 ms | 0.10 ms | 0.06 ms | 1.00 | 1.00 | 2375.5 KB | 1.00 | 187.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | report-workbook-core | EPPlus | 117.65 ms | 25.03 ms | 14.45 ms | 19.03 | 19.03 | 71893.6 KB | 30.26 | 157.2 KB | 0.84 | 1802.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook-core | ClosedXML | 118.34 ms | 10.24 ms | 5.91 ms | 19.14 | 19.14 | 97210.5 KB | 40.92 | 165.1 KB | 0.88 | 1813.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook-datatable | OfficeIMO.Excel | 16.45 ms | 0.30 ms | 0.17 ms | 1.00 | 1.00 | 12500.1 KB | 1.00 | 90.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | report-workbook-datatable | EPPlus | 98.05 ms | 8.90 ms | 5.14 ms | 5.96 | 5.96 | 65913.8 KB | 5.27 | 161.8 KB | 1.79 | 495.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook-datatable-core | OfficeIMO.Excel | 6.33 ms | 0.57 ms | 0.33 ms | 1.00 | 1.00 | 2647.2 KB | 1.00 | 187.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | report-workbook-datatable-core | EPPlus | 97.29 ms | 7.45 ms | 4.30 ms | 15.37 | 15.37 | 60398.6 KB | 22.82 | 157.2 KB | 0.84 | 1436.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook-datatable-core | ClosedXML | 112.16 ms | 6.08 ms | 3.51 ms | 17.71 | 17.71 | 82852.7 KB | 31.30 | 165.1 KB | 0.88 | 1671.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | LargeXlsx | 4.37 ms | 0.51 ms | 0.29 ms | 0.88 | 1.00 | 849.6 KB | 0.56 | 237.7 KB | 1.10 | 12.5% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | OfficeIMO.Excel | 5.00 ms | 0.84 ms | 0.49 ms | 1.00 | 1.14 | 1508.4 KB | 1.00 | 216.7 KB | 1.00 | Loss +14.3% |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | MiniExcel | 18.06 ms | 2.10 ms | 1.21 ms | 3.61 | 4.13 | 35911.1 KB | 23.81 | 235.3 KB | 1.09 | 261.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | ClosedXML | 103.58 ms | 3.25 ms | 1.88 ms | 20.73 | 23.69 | 71470.2 KB | 47.38 | 257.2 KB | 1.19 | 1973.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-bulk-report | OfficeIMO.Excel | 4.78 ms | 0.66 ms | 0.38 ms | 1.00 | 1.00 | 1233.1 KB | 1.00 | 143.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-bulk-report | MiniExcel | 9.96 ms | 1.01 ms | 0.58 ms | 2.08 | 2.08 | 26816.2 KB | 21.75 | 153.8 KB | 1.07 | 108.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-bulk-report | EPPlus | 76.38 ms | 6.03 ms | 3.48 ms | 15.96 | 15.96 | 47121.2 KB | 38.21 | 115.0 KB | 0.80 | 1496.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-bulk-report | ClosedXML | 93.49 ms | 6.86 ms | 3.96 ms | 19.54 | 19.54 | 58336.8 KB | 47.31 | 121.0 KB | 0.84 | 1853.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellformula | OfficeIMO.Excel | 2.27 ms | 0.02 ms | 0.01 ms | 1.00 | 1.00 | 1171.8 KB | 1.00 | 66.6 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellformula | ClosedXML | 16.56 ms | 0.59 ms | 0.34 ms | 7.30 | 7.30 | 12031.2 KB | 10.27 | 70.6 KB | 1.06 | 630.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellformula | EPPlus | 39.21 ms | 6.43 ms | 3.71 ms | 17.29 | 17.29 | 18036.5 KB | 15.39 | 62.1 KB | 0.93 | 1629.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-empty-strings | OfficeIMO.Excel | 2.40 ms | 0.59 ms | 0.34 ms | 1.00 | 1.00 | 1454.0 KB | 1.00 | 44.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-empty-strings | ClosedXML | 14.84 ms | 2.95 ms | 1.70 ms | 6.18 | 6.18 | 9951.5 KB | 6.84 | 44.9 KB | 1.02 | 517.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-empty-strings | EPPlus | 27.33 ms | 4.56 ms | 2.63 ms | 11.38 | 11.38 | 11703.7 KB | 8.05 | 42.0 KB | 0.95 | 1037.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-numbers | OfficeIMO.Excel | 2.01 ms | 0.08 ms | 0.05 ms | 1.00 | 1.00 | 946.8 KB | 1.00 | 47.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-numbers | ClosedXML | 11.14 ms | 0.58 ms | 0.33 ms | 5.54 | 5.54 | 9169.1 KB | 9.68 | 45.9 KB | 0.98 | 453.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-numbers | EPPlus | 23.22 ms | 0.41 ms | 0.24 ms | 11.54 | 11.54 | 12829.3 KB | 13.55 | 43.7 KB | 0.93 | 1054.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-mixed | OfficeIMO.Excel | 3.14 ms | 1.11 ms | 0.64 ms | 1.00 | 1.00 | 1431.4 KB | 1.00 | 61.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-object-mixed | ClosedXML | 19.34 ms | 1.92 ms | 1.11 ms | 6.16 | 6.16 | 11879.0 KB | 8.30 | 59.5 KB | 0.97 | 515.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-mixed | EPPlus | 31.64 ms | 1.98 ms | 1.14 ms | 10.07 | 10.07 | 15577.2 KB | 10.88 | 58.9 KB | 0.96 | 907.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse | OfficeIMO.Excel | 2.60 ms | 0.14 ms | 0.08 ms | 1.00 | 1.00 | 1271.0 KB | 1.00 | 62.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse | ClosedXML | 18.09 ms | 4.52 ms | 2.61 ms | 6.96 | 6.96 | 11288.3 KB | 8.88 | 52.5 KB | 0.85 | 596.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse | EPPlus | 24.91 ms | 2.15 ms | 1.24 ms | 9.59 | 9.59 | 14894.0 KB | 11.72 | 54.2 KB | 0.88 | 859.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 2.27 ms | 0.04 ms | 0.02 ms | 1.00 | 1.00 | 1271.1 KB | 1.00 | 62.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | ClosedXML | 14.77 ms | 1.31 ms | 0.76 ms | 6.51 | 6.51 | 11288.3 KB | 8.88 | 52.5 KB | 0.85 | 551.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | EPPlus | 24.83 ms | 0.66 ms | 0.38 ms | 10.95 | 10.95 | 14894.0 KB | 11.72 | 54.2 KB | 0.88 | 995.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-scalars | OfficeIMO.Excel | 1.85 ms | 0.09 ms | 0.05 ms | 1.00 | 1.00 | 964.9 KB | 1.00 | 46.3 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-scalars | ClosedXML | 11.94 ms | 1.18 ms | 0.68 ms | 6.47 | 6.47 | 9013.2 KB | 9.34 | 45.4 KB | 0.98 | 546.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-scalars | EPPlus | 24.55 ms | 2.36 ms | 1.36 ms | 13.29 | 13.29 | 12761.5 KB | 13.23 | 42.4 KB | 0.91 | 1229.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings | OfficeIMO.Excel | 2.87 ms | 0.27 ms | 0.15 ms | 1.00 | 1.00 | 2283.8 KB | 1.00 | 55.1 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-strings | ClosedXML | 12.04 ms | 0.87 ms | 0.50 ms | 4.19 | 4.19 | 11291.2 KB | 4.94 | 50.3 KB | 0.91 | 319.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings | EPPlus | 25.70 ms | 1.31 ms | 0.76 ms | 8.95 | 8.95 | 12730.2 KB | 5.57 | 48.1 KB | 0.87 | 794.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-distinct | OfficeIMO.Excel | 2.39 ms | 0.24 ms | 0.14 ms | 1.00 | 1.00 | 2206.1 KB | 1.00 | 51.8 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-distinct | ClosedXML | 16.13 ms | 2.33 ms | 1.35 ms | 6.76 | 6.76 | 13119.1 KB | 5.95 | 61.9 KB | 1.19 | 576.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-distinct | EPPlus | 24.08 ms | 1.18 ms | 0.68 ms | 10.09 | 10.09 | 13793.7 KB | 6.25 | 61.5 KB | 1.19 | 909.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-repeated | OfficeIMO.Excel | 2.26 ms | 0.59 ms | 0.34 ms | 1.00 | 1.00 | 1246.7 KB | 1.00 | 40.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-repeated | ClosedXML | 13.11 ms | 2.47 ms | 1.43 ms | 5.79 | 5.79 | 9218.5 KB | 7.39 | 38.8 KB | 0.97 | 479.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-repeated | EPPlus | 20.78 ms | 1.37 ms | 0.79 ms | 9.18 | 9.18 | 11265.7 KB | 9.04 | 34.8 KB | 0.87 | 817.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-temporal | OfficeIMO.Excel | 2.49 ms | 0.03 ms | 0.02 ms | 1.00 | 1.00 | 1262.6 KB | 1.00 | 63.3 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-temporal | ClosedXML | 15.12 ms | 1.33 ms | 0.77 ms | 6.09 | 6.09 | 9703.1 KB | 7.69 | 54.5 KB | 0.86 | 508.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-temporal | EPPlus | 25.99 ms | 2.81 ms | 1.62 ms | 10.46 | 10.46 | 14654.6 KB | 11.61 | 53.1 KB | 0.84 | 945.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 1.23 ms | 0.08 ms | 0.05 ms | 0.69 | 1.00 | 439.0 KB | 0.48 | 47.3 KB | 0.98 | 31.4% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 1.79 ms | 0.22 ms | 0.13 ms | 1.00 | 1.46 | 923.6 KB | 1.00 | 48.2 KB | 1.00 | Loss +45.7% |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | ClosedXML | 12.80 ms | 1.34 ms | 0.77 ms | 7.15 | 10.42 | 10227.8 KB | 11.07 | 53.0 KB | 1.10 | 615.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | EPPlus | 24.08 ms | 2.19 ms | 1.27 ms | 13.45 | 19.60 | 12985.4 KB | 14.06 | 52.5 KB | 1.09 | 1245.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | LargeXlsx | 3.26 ms | 0.20 ms | 0.12 ms | 0.79 | 1.00 | 750.2 KB | 0.43 | 138.4 KB | 1.00 | 21.4% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 4.15 ms | 0.17 ms | 0.10 ms | 1.00 | 1.27 | 1752.7 KB | 1.00 | 138.0 KB | 1.00 | Loss +27.3% |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | MiniExcel | 8.04 ms | 0.31 ms | 0.18 ms | 1.94 | 2.47 | 23213.1 KB | 13.24 | 153.7 KB | 1.11 | 93.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | ClosedXML | 32.29 ms | 1.88 ms | 1.09 ms | 7.79 | 9.91 | 22213.3 KB | 12.67 | 120.1 KB | 0.87 | 678.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | EPPlus | 51.02 ms | 4.02 ms | 2.32 ms | 12.31 | 15.66 | 24626.9 KB | 14.05 | 114.1 KB | 0.83 | 1130.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | Sylvan.Data.Excel | 4.32 ms | 1.84 ms | 1.06 ms | 0.91 | 1.00 | 750.7 KB | 0.52 | 78.5 KB | 0.57 | 9.3% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | OfficeIMO.Excel | 4.76 ms | 1.35 ms | 0.78 ms | 1.00 | 1.10 | 1434.6 KB | 1.00 | 138.0 KB | 1.00 | Loss +10.2% |
| 2500 | package-profile | package | Package size | write-datareader-plain | LargeXlsx | 5.08 ms | 0.87 ms | 0.50 ms | 1.07 | 1.18 | 1024.5 KB | 0.71 | 138.4 KB | 1.00 | 6.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | MiniExcel | 9.66 ms | 3.02 ms | 1.74 ms | 2.03 | 2.24 | 23034.8 KB | 16.06 | 153.6 KB | 1.11 | 103.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | ClosedXML | 31.62 ms | 3.51 ms | 2.03 ms | 6.65 | 7.33 | 11573.0 KB | 8.07 | 120.1 KB | 0.87 | 564.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | EPPlus | 44.99 ms | 4.99 ms | 2.88 ms | 9.46 | 10.42 | 16579.3 KB | 11.56 | 114.9 KB | 0.83 | 845.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table | OfficeIMO.Excel | 7.12 ms | 4.83 ms | 2.79 ms | 1.00 | 1.00 | 1446.4 KB | 1.00 | 138.8 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-datareader-table | MiniExcel | 7.61 ms | 1.03 ms | 0.60 ms | 1.07 | 1.07 | 23035.0 KB | 15.93 | 153.6 KB | 1.11 | 6.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table | ClosedXML | 36.72 ms | 0.67 ms | 0.39 ms | 5.15 | 5.15 | 18999.5 KB | 13.14 | 120.9 KB | 0.87 | 415.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table | EPPlus | 46.27 ms | 8.77 ms | 5.06 ms | 6.49 | 6.49 | 16579.3 KB | 11.46 | 114.9 KB | 0.83 | 549.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | OfficeIMO.Excel | 4.84 ms | 0.24 ms | 0.14 ms | 1.00 | 1.00 | 1452.7 KB | 1.00 | 138.9 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | MiniExcel | 8.33 ms | 1.07 ms | 0.62 ms | 1.72 | 1.72 | 26638.2 KB | 18.34 | 153.8 KB | 1.11 | 72.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | EPPlus | 68.76 ms | 6.91 ms | 3.99 ms | 14.21 | 14.21 | 38271.3 KB | 26.34 | 115.1 KB | 0.83 | 1320.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | ClosedXML | 94.38 ms | 11.31 ms | 6.53 ms | 19.50 | 19.50 | 58353.1 KB | 40.17 | 121.0 KB | 0.87 | 1850.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | LargeXlsx | 3.96 ms | 0.06 ms | 0.04 ms | 0.90 | 1.00 | 1115.8 KB | 0.68 | 164.2 KB | 1.25 | 10.4% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | OfficeIMO.Excel | 4.42 ms | 0.73 ms | 0.42 ms | 1.00 | 1.12 | 1652.6 KB | 1.00 | 131.1 KB | 1.00 | Loss +11.7% |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | MiniExcel | 9.48 ms | 0.29 ms | 0.17 ms | 2.14 | 2.40 | 29737.8 KB | 17.99 | 180.5 KB | 1.38 | 114.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | EPPlus | 54.50 ms | 1.42 ms | 0.82 ms | 12.33 | 13.77 | 21822.9 KB | 13.21 | 144.5 KB | 1.10 | 1132.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | ClosedXML | 64.55 ms | 5.28 ms | 3.05 ms | 14.60 | 16.30 | 27401.4 KB | 16.58 | 159.4 KB | 1.22 | 1360.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | OfficeIMO.Excel | 4.90 ms | 0.85 ms | 0.49 ms | 1.00 | 1.00 | 2392.1 KB | 1.00 | 176.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | MiniExcel | 10.77 ms | 0.79 ms | 0.46 ms | 2.20 | 2.20 | 29737.8 KB | 12.43 | 180.5 KB | 1.03 | 119.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | EPPlus | 58.49 ms | 7.90 ms | 4.56 ms | 11.94 | 11.94 | 21822.9 KB | 9.12 | 144.5 KB | 0.82 | 1094.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | ClosedXML | 65.46 ms | 11.90 ms | 6.87 ms | 13.37 | 13.37 | 27403.3 KB | 11.46 | 159.4 KB | 0.91 | 1236.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables | OfficeIMO.Excel | 3.86 ms | 0.25 ms | 0.15 ms | 1.00 | 1.00 | 1578.0 KB | 1.00 | 138.9 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-dataset-tables | MiniExcel | 7.89 ms | 0.44 ms | 0.26 ms | 2.04 | 2.04 | 28691.3 KB | 18.18 | 156.4 KB | 1.13 | 104.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables | EPPlus | 38.26 ms | 1.23 ms | 0.71 ms | 9.91 | 9.91 | 18633.8 KB | 11.81 | 116.6 KB | 0.84 | 891.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables | ClosedXML | 40.02 ms | 2.98 ms | 1.72 ms | 10.37 | 10.37 | 18868.8 KB | 11.96 | 123.4 KB | 0.89 | 936.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | OfficeIMO.Excel | 4.73 ms | 0.68 ms | 0.40 ms | 1.00 | 1.00 | 1590.6 KB | 1.00 | 139.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | MiniExcel | 9.20 ms | 0.81 ms | 0.47 ms | 1.95 | 1.95 | 31789.4 KB | 19.99 | 156.6 KB | 1.13 | 94.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | EPPlus | 74.88 ms | 6.35 ms | 3.66 ms | 15.84 | 15.84 | 41385.5 KB | 26.02 | 116.9 KB | 0.84 | 1483.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | ClosedXML | 79.77 ms | 6.89 ms | 3.98 ms | 16.87 | 16.87 | 56700.1 KB | 35.65 | 123.7 KB | 0.89 | 1587.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-direct | LargeXlsx | 3.54 ms | 0.04 ms | 0.02 ms | 0.49 | 1.00 | 1141.0 KB | 0.80 | 138.4 KB | 1.00 | 50.6% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-direct | OfficeIMO.Excel | 7.16 ms | 5.31 ms | 3.06 ms | 1.00 | 2.03 | 1420.2 KB | 1.00 | 138.0 KB | 1.00 | Loss +102.5% |
| 2500 | package-profile | package | Package size | write-datatable-direct | MiniExcel | 8.41 ms | 0.53 ms | 0.31 ms | 1.18 | 2.38 | 23053.5 KB | 16.23 | 153.7 KB | 1.11 | 17.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-direct | ClosedXML | 33.78 ms | 10.36 ms | 5.98 ms | 4.72 | 9.55 | 11573.0 KB | 8.15 | 120.1 KB | 0.87 | 371.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-direct | EPPlus | 40.32 ms | 2.88 ms | 1.66 ms | 5.63 | 11.41 | 16579.3 KB | 11.67 | 114.9 KB | 0.83 | 463.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | OfficeIMO.Excel | 4.56 ms | 0.61 ms | 0.35 ms | 1.00 | 1.00 | 1432.2 KB | 1.00 | 138.8 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | MiniExcel | 8.42 ms | 0.28 ms | 0.16 ms | 1.85 | 1.85 | 23053.8 KB | 16.10 | 153.7 KB | 1.11 | 84.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | ClosedXML | 39.21 ms | 2.66 ms | 1.54 ms | 8.60 | 8.60 | 18999.9 KB | 13.27 | 120.9 KB | 0.87 | 759.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | EPPlus | 39.60 ms | 3.13 ms | 1.81 ms | 8.68 | 8.68 | 16579.3 KB | 11.58 | 114.9 KB | 0.83 | 768.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | LargeXlsx | 3.38 ms | 0.36 ms | 0.21 ms | 0.75 | 1.00 | 750.2 KB | 0.52 | 138.4 KB | 1.00 | 25.3% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 4.52 ms | 0.67 ms | 0.38 ms | 1.00 | 1.34 | 1440.7 KB | 1.00 | 138.0 KB | 1.00 | Loss +33.9% |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | MiniExcel | 11.68 ms | 3.71 ms | 2.14 ms | 2.58 | 3.46 | 23213.1 KB | 16.11 | 153.7 KB | 1.11 | 158.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | ClosedXML | 43.17 ms | 12.47 ms | 7.20 ms | 9.54 | 12.78 | 11573.0 KB | 8.03 | 120.1 KB | 0.87 | 854.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | EPPlus | 44.85 ms | 1.54 ms | 0.89 ms | 9.92 | 13.28 | 16579.3 KB | 11.51 | 114.9 KB | 0.83 | 891.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.99 ms | 2.09 ms | 1.20 ms | 1.00 | 1.00 | 1178.6 KB | 1.00 | 142.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 70.27 ms | 10.02 ms | 5.79 ms | 14.08 | 14.08 | 50919.5 KB | 43.20 | 120.2 KB | 0.84 | 1308.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 72.92 ms | 12.19 ms | 7.04 ms | 14.61 | 14.61 | 38271.3 KB | 32.47 | 115.1 KB | 0.81 | 1361.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | LargeXlsx | 3.70 ms | 0.64 ms | 0.37 ms | 0.92 | 1.00 | 750.2 KB | 0.64 | 138.4 KB | 0.97 | 8.0% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | OfficeIMO.Excel | 4.02 ms | 1.01 ms | 0.58 ms | 1.00 | 1.09 | 1170.9 KB | 1.00 | 142.3 KB | 1.00 | Loss +8.7% |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | MiniExcel | 7.46 ms | 0.73 ms | 0.42 ms | 1.85 | 2.01 | 23213.1 KB | 19.83 | 153.7 KB | 1.08 | 85.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | ClosedXML | 29.77 ms | 2.31 ms | 1.34 ms | 7.40 | 8.04 | 11573.0 KB | 9.88 | 120.1 KB | 0.84 | 640.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | EPPlus | 42.31 ms | 7.27 ms | 4.20 ms | 10.52 | 11.43 | 16579.3 KB | 14.16 | 114.9 KB | 0.81 | 952.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 5.71 ms | 1.32 ms | 0.76 ms | 1.00 | 1.00 | 1176.8 KB | 1.00 | 138.1 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 73.08 ms | 7.76 ms | 4.48 ms | 12.80 | 12.80 | 38271.3 KB | 32.52 | 115.1 KB | 0.83 | 1180.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 87.93 ms | 17.24 ms | 9.95 ms | 15.40 | 15.40 | 50919.5 KB | 43.27 | 120.2 KB | 0.87 | 1440.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 3.05 ms | 0.18 ms | 0.11 ms | 0.73 | 1.00 | 750.2 KB | 0.64 | 138.4 KB | 1.00 | 27.0% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 4.18 ms | 0.46 ms | 0.27 ms | 1.00 | 1.37 | 1169.1 KB | 1.00 | 138.0 KB | 1.00 | Loss +36.9% |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | MiniExcel | 7.26 ms | 0.19 ms | 0.11 ms | 1.74 | 2.38 | 23213.1 KB | 19.86 | 153.7 KB | 1.11 | 73.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | ClosedXML | 27.30 ms | 0.76 ms | 0.44 ms | 6.53 | 8.94 | 11573.0 KB | 9.90 | 120.1 KB | 0.87 | 553.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | EPPlus | 40.65 ms | 4.42 ms | 2.55 ms | 9.73 | 13.32 | 16579.3 KB | 14.18 | 114.9 KB | 0.83 | 872.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | LargeXlsx | 3.48 ms | 0.42 ms | 0.24 ms | 0.58 | 1.00 | 750.2 KB | 0.47 | 138.4 KB | 0.97 | 42.0% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | OfficeIMO.Excel | 5.99 ms | 1.83 ms | 1.05 ms | 1.00 | 1.72 | 1601.6 KB | 1.00 | 142.3 KB | 1.00 | Loss +72.3% |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | MiniExcel | 9.62 ms | 2.51 ms | 1.45 ms | 1.61 | 2.77 | 23213.1 KB | 14.49 | 153.7 KB | 1.08 | 60.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | ClosedXML | 28.70 ms | 1.39 ms | 0.80 ms | 4.79 | 8.26 | 11573.0 KB | 7.23 | 120.1 KB | 0.84 | 379.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | EPPlus | 39.24 ms | 1.12 ms | 0.64 ms | 6.55 | 11.29 | 16579.3 KB | 10.35 | 114.9 KB | 0.81 | 555.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 3.68 ms | 0.12 ms | 0.07 ms | 1.00 | 1.00 | 1177.9 KB | 1.00 | 142.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 41.77 ms | 2.27 ms | 1.31 ms | 11.36 | 11.36 | 28532.5 KB | 24.22 | 120.2 KB | 0.84 | 1035.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 56.45 ms | 4.50 ms | 2.60 ms | 15.35 | 15.35 | 27236.1 KB | 23.12 | 115.0 KB | 0.81 | 1434.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | LargeXlsx | 4.05 ms | 0.13 ms | 0.08 ms | 0.68 | 1.00 | 794.5 KB | 0.39 | 182.6 KB | 1.00 | 32.2% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 5.97 ms | 0.62 ms | 0.36 ms | 1.00 | 1.48 | 2013.3 KB | 1.00 | 183.1 KB | 1.00 | Loss +47.6% |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | MiniExcel | 8.42 ms | 0.34 ms | 0.20 ms | 1.41 | 2.08 | 25181.4 KB | 12.51 | 194.0 KB | 1.06 | 40.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | EPPlus | 53.40 ms | 8.37 ms | 4.84 ms | 8.94 | 13.19 | 20030.7 KB | 9.95 | 152.1 KB | 0.83 | 793.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | ClosedXML | 54.32 ms | 18.16 ms | 10.49 ms | 9.09 | 13.42 | 16965.4 KB | 8.43 | 161.0 KB | 0.88 | 809.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 4.46 ms | 0.32 ms | 0.18 ms | 1.00 | 1.00 | 1339.3 KB | 1.00 | 182.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | LargeXlsx | 4.73 ms | 1.38 ms | 0.80 ms | 1.06 | 1.06 | 794.5 KB | 0.59 | 182.6 KB | 1.00 | 6.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | MiniExcel | 9.33 ms | 0.55 ms | 0.32 ms | 2.09 | 2.09 | 25181.4 KB | 18.80 | 194.0 KB | 1.06 | 109.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | ClosedXML | 45.15 ms | 7.57 ms | 4.37 ms | 10.13 | 10.13 | 16965.4 KB | 12.67 | 161.0 KB | 0.88 | 912.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | EPPlus | 52.58 ms | 6.05 ms | 3.49 ms | 11.79 | 11.79 | 20030.7 KB | 14.96 | 152.1 KB | 0.83 | 1079.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 20.71 ms | 1.20 ms | 0.69 ms | 1.00 | 1.00 | 4333.9 KB | 1.00 | 651.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | LargeXlsx | 25.80 ms | 6.41 ms | 3.70 ms | 1.25 | 1.25 | 2802.7 KB | 0.65 | 644.6 KB | 0.99 | 24.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | MiniExcel | 40.03 ms | 7.79 ms | 4.50 ms | 1.93 | 1.93 | 48404.7 KB | 11.17 | 674.4 KB | 1.04 | 93.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | ClosedXML | 141.94 ms | 29.11 ms | 16.81 ms | 6.85 | 6.85 | 51639.0 KB | 11.92 | 615.5 KB | 0.95 | 585.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | EPPlus | 181.38 ms | 7.80 ms | 4.50 ms | 8.76 | 8.76 | 69073.3 KB | 15.94 | 548.9 KB | 0.84 | 775.8% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | OfficeIMO.Excel | 28.07 ms | 0.50 ms | 0.29 ms | 1.00 | 1.00 | 13971.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | EPPlus | 70.35 ms | 1.06 ms | 0.61 ms | 2.51 | 2.51 | 50639.2 KB | 3.62 |  |  | 150.6% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | EPPlus 4.5.3.3 | 101.48 ms |  |  | 3.61 | 3.61 |  |  |  |  | 261.5% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | ClosedXML | 116.48 ms | 2.97 ms | 1.71 ms | 4.15 | 4.15 | 84601.4 KB | 6.06 |  |  | 314.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | DataTable table export | report-workbook-datatable | OfficeIMO.Excel | 13.57 ms | 2.52 ms | 1.45 ms | 1.00 | 1.00 | 12500.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | DataTable table export | report-workbook-datatable | EPPlus | 78.84 ms | 0.32 ms | 0.19 ms | 5.81 | 5.81 | 65913.9 KB | 5.27 |  |  | 480.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | DataTable table export | report-workbook-datatable | EPPlus 4.5.3.3 | 95.85 ms |  |  | 7.06 | 7.06 |  |  |  |  | 606.2% slower than OfficeIMO |
| 2500 | speed-comparison | other | DataTable table export | report-workbook-datatable-core | OfficeIMO.Excel | 5.55 ms | 0.37 ms | 0.21 ms | 1.00 | 1.00 | 2647.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | DataTable table export | report-workbook-datatable-core | EPPlus | 74.43 ms | 1.10 ms | 0.64 ms | 13.41 | 13.41 | 60398.6 KB | 22.82 |  |  | 1241.5% slower than OfficeIMO |
| 2500 | speed-comparison | other | DataTable table export | report-workbook-datatable-core | EPPlus 4.5.3.3 | 85.02 ms |  |  | 15.32 | 15.32 |  |  |  |  | 1432.2% slower than OfficeIMO |
| 2500 | speed-comparison | other | DataTable table export | report-workbook-datatable-core | ClosedXML | 94.28 ms | 1.94 ms | 1.12 ms | 16.99 | 16.99 | 82850.8 KB | 31.30 |  |  | 1599.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Object projection | build-object-datatable-dictionaries | OfficeIMO.Excel | 1.57 ms | 0.16 ms | 0.09 ms | 1.00 | 1.00 | 564.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Object projection | build-object-datatable-typed | OfficeIMO.Excel | 1.34 ms | 0.25 ms | 0.14 ms | 1.00 | 1.00 | 856.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Other | realworld-autofilter | OfficeIMO.Excel | 3.69 ms | 0.38 ms | 0.22 ms | 1.00 | 1.00 | 1171.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Other | realworld-autofilter | ClosedXML | 28.95 ms | 0.67 ms | 0.38 ms | 7.83 | 7.83 | 22218.8 KB | 18.96 |  |  | 683.4% slower than OfficeIMO |
| 2500 | speed-comparison | other | Other | realworld-autofilter | EPPlus 4.5.3.3 | 39.92 ms |  |  | 10.80 | 10.80 |  |  |  |  | 980.4% slower than OfficeIMO |
| 2500 | speed-comparison | other | Other | realworld-autofilter | EPPlus | 44.38 ms | 1.15 ms | 0.67 ms | 12.01 | 12.01 | 24647.8 KB | 21.03 |  |  | 1101.2% slower than OfficeIMO |
| 2500 | speed-comparison | other | Other | realworld-charts | OfficeIMO.Excel | 5.01 ms | 0.26 ms | 0.15 ms | 1.00 | 1.00 | 1689.0 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Other | realworld-charts | EPPlus 4.5.3.3 | 39.41 ms |  |  | 7.87 | 7.87 |  |  |  |  | 687.4% slower than OfficeIMO |
| 2500 | speed-comparison | other | Other | realworld-charts | EPPlus | 51.29 ms | 4.51 ms | 2.60 ms | 10.25 | 10.25 | 27072.7 KB | 16.03 |  |  | 924.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Other | realworld-conditional-formatting | OfficeIMO.Excel | 3.93 ms | 0.23 ms | 0.14 ms | 1.00 | 1.00 | 1236.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Other | realworld-conditional-formatting | ClosedXML | 30.67 ms | 0.65 ms | 0.37 ms | 7.80 | 7.80 | 22265.8 KB | 18.00 |  |  | 679.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Other | realworld-conditional-formatting | EPPlus 4.5.3.3 | 39.96 ms |  |  | 10.16 | 10.16 |  |  |  |  | 916.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Other | realworld-conditional-formatting | EPPlus | 47.48 ms | 0.18 ms | 0.10 ms | 12.08 | 12.08 | 24690.0 KB | 19.96 |  |  | 1107.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Other | realworld-data-validation | OfficeIMO.Excel | 3.67 ms | 0.15 ms | 0.09 ms | 1.00 | 1.00 | 1187.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Other | realworld-data-validation | ClosedXML | 29.63 ms | 0.37 ms | 0.22 ms | 8.07 | 8.07 | 22239.9 KB | 18.73 |  |  | 706.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Other | realworld-data-validation | EPPlus 4.5.3.3 | 42.20 ms |  |  | 11.49 | 11.49 |  |  |  |  | 1049.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Other | realworld-data-validation | EPPlus | 46.03 ms | 0.78 ms | 0.45 ms | 12.54 | 12.54 | 24632.9 KB | 20.74 |  |  | 1153.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Other | realworld-freeze-panes | OfficeIMO.Excel | 3.46 ms | 0.04 ms | 0.02 ms | 1.00 | 1.00 | 1174.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Other | realworld-freeze-panes | ClosedXML | 28.29 ms | 0.30 ms | 0.17 ms | 8.17 | 8.17 | 22214.0 KB | 18.92 |  |  | 717.1% slower than OfficeIMO |
| 2500 | speed-comparison | other | Other | realworld-freeze-panes | EPPlus 4.5.3.3 | 38.40 ms |  |  | 11.09 | 11.09 |  |  |  |  | 1009.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Other | realworld-freeze-panes | EPPlus | 44.71 ms | 0.52 ms | 0.30 ms | 12.92 | 12.92 | 24660.4 KB | 21.00 |  |  | 1191.6% slower than OfficeIMO |
| 2500 | speed-comparison | other | Other | realworld-pivot-table | OfficeIMO.Excel | 16.56 ms | 0.05 ms | 0.03 ms | 1.00 | 1.00 | 18627.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Other | realworld-pivot-table | EPPlus 4.5.3.3 | 37.41 ms |  |  | 2.26 | 2.26 |  |  |  |  | 125.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Other | realworld-pivot-table | EPPlus | 50.71 ms | 0.60 ms | 0.35 ms | 3.06 | 3.06 | 29469.8 KB | 1.58 |  |  | 206.1% slower than OfficeIMO |
| 2500 | speed-comparison | other | Other | realworld-report-all-in-one | OfficeIMO.Excel | 18.59 ms | 0.36 ms | 0.21 ms | 1.00 | 1.00 | 19412.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Other | realworld-report-all-in-one | EPPlus | 58.40 ms | 1.98 ms | 1.15 ms | 3.14 | 3.14 | 54522.6 KB | 2.81 |  |  | 214.2% slower than OfficeIMO |
| 2500 | speed-comparison | other | Other | realworld-report-all-in-one | EPPlus 4.5.3.3 | 76.95 ms |  |  | 4.14 | 4.14 |  |  |  |  | 313.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Other | realworld-report-core | OfficeIMO.Excel | 4.15 ms | 0.05 ms | 0.03 ms | 1.00 | 1.00 | 1319.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Other | realworld-report-core | EPPlus | 61.69 ms | 0.47 ms | 0.27 ms | 14.85 | 14.85 | 47227.1 KB | 35.80 |  |  | 1384.8% slower than OfficeIMO |
| 2500 | speed-comparison | other | Other | realworld-report-core | EPPlus 4.5.3.3 | 73.14 ms |  |  | 17.60 | 17.60 |  |  |  |  | 1660.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Other | realworld-report-core | ClosedXML | 75.83 ms | 1.23 ms | 0.71 ms | 18.25 | 18.25 | 69825.6 KB | 52.93 |  |  | 1725.2% slower than OfficeIMO |
| 2500 | speed-comparison | other | Other | report-workbook | OfficeIMO.Excel | 15.21 ms | 1.13 ms | 0.65 ms | 1.00 | 1.00 | 12228.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Other | report-workbook | EPPlus | 92.76 ms | 3.31 ms | 1.91 ms | 6.10 | 6.10 | 77408.8 KB | 6.33 |  |  | 509.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Other | report-workbook | EPPlus 4.5.3.3 | 105.14 ms |  |  | 6.91 | 6.91 |  |  |  |  | 591.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Other | report-workbook-core | OfficeIMO.Excel | 5.79 ms | 0.02 ms | 0.01 ms | 1.00 | 1.00 | 2375.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Other | report-workbook-core | EPPlus | 79.26 ms | 3.75 ms | 2.17 ms | 13.69 | 13.69 | 71893.6 KB | 30.26 |  |  | 1269.4% slower than OfficeIMO |
| 2500 | speed-comparison | other | Other | report-workbook-core | ClosedXML | 98.36 ms | 2.94 ms | 1.70 ms | 16.99 | 16.99 | 97210.4 KB | 40.92 |  |  | 1599.4% slower than OfficeIMO |
| 2500 | speed-comparison | other | Other | report-workbook-core | EPPlus 4.5.3.3 | 98.91 ms |  |  | 17.09 | 17.09 |  |  |  |  | 1609.0% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-cells | OfficeIMO.Excel | 7.92 ms | 0.08 ms | 0.04 ms | 1.00 | 1.00 | 2649.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-cells | EPPlus | 27.96 ms | 4.98 ms | 2.88 ms | 3.53 | 3.53 | 20154.4 KB | 7.61 |  |  | 252.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-cells | ClosedXML | 29.81 ms | 0.66 ms | 0.38 ms | 3.76 | 3.76 | 17018.2 KB | 6.42 |  |  | 276.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | OfficeIMO.Excel | 4.83 ms | 0.10 ms | 0.06 ms | 1.00 | 1.00 | 643.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | EPPlus | 23.00 ms | 0.31 ms | 0.18 ms | 4.76 | 4.76 | 13107.5 KB | 20.37 |  |  | 375.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | ClosedXML | 30.15 ms | 1.00 ms | 0.58 ms | 6.24 | 6.24 | 15455.9 KB | 24.02 |  |  | 523.8% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-range | OfficeIMO.Excel | 8.36 ms | 0.36 ms | 0.21 ms | 1.00 | 1.00 | 2649.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-range | EPPlus | 27.50 ms | 4.47 ms | 2.58 ms | 3.29 | 3.29 | 20154.4 KB | 7.61 |  |  | 229.0% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-range | ClosedXML | 28.87 ms | 0.31 ms | 0.18 ms | 3.45 | 3.45 | 17018.9 KB | 6.42 |  |  | 245.4% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-top-range | OfficeIMO.Excel | 1.69 ms | 0.06 ms | 0.03 ms | 1.00 | 1.00 | 402.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-top-range | EPPlus | 23.03 ms | 0.57 ms | 0.33 ms | 13.62 | 13.62 | 12403.9 KB | 30.81 |  |  | 1261.5% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-top-range | ClosedXML | 28.74 ms | 0.52 ms | 0.30 ms | 16.99 | 16.99 | 15368.8 KB | 38.17 |  |  | 1599.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | OfficeIMO.Excel | 3.33 ms | 0.12 ms | 0.07 ms | 1.00 | 1.00 | 777.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | ClosedXML | 16.25 ms | 0.36 ms | 0.21 ms | 4.88 | 4.88 | 8271.4 KB | 10.64 |  |  | 388.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | EPPlus 4.5.3.3 | 17.44 ms |  |  | 5.24 | 5.24 |  |  |  |  | 423.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | EPPlus | 18.86 ms | 1.53 ms | 0.89 ms | 5.67 | 5.67 | 7707.5 KB | 9.91 |  |  | 466.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | Sylvan.Data.Excel | 1.01 ms | 0.03 ms | 0.02 ms | 0.58 | 1.00 | 316.6 KB | 1.27 |  |  | 42.0% faster than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | ExcelDataReader | 1.51 ms | 0.13 ms | 0.07 ms | 0.87 | 1.50 | 4046.2 KB | 16.26 |  |  | 13.0% faster than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | OfficeIMO.Excel | 1.73 ms | 0.21 ms | 0.12 ms | 1.00 | 1.72 | 248.8 KB | 1.00 |  |  | Loss +72.4% |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | ClosedXML | 3.27 ms | 0.06 ms | 0.03 ms | 1.89 | 3.26 | 4392.9 KB | 17.65 |  |  | 88.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | MiniExcel | 10.85 ms | 0.28 ms | 0.16 ms | 6.26 | 10.79 | 46189.1 KB | 185.63 |  |  | 526.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | EPPlus 4.5.3.3 | 11.98 ms |  |  | 6.92 | 11.92 |  |  |  |  | 591.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | EPPlus | 38.09 ms | 1.35 ms | 0.78 ms | 21.98 | 37.89 | 43070.2 KB | 173.09 |  |  | 2098.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | Sylvan.Data.Excel | 0.99 ms | 0.04 ms | 0.02 ms | 0.42 | 1.00 | 316.6 KB | 1.27 |  |  | 58.3% faster than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | ExcelDataReader | 1.48 ms | 0.07 ms | 0.04 ms | 0.62 | 1.48 | 4046.2 KB | 16.26 |  |  | 38.0% faster than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | OfficeIMO.Excel | 2.38 ms | 0.63 ms | 0.36 ms | 1.00 | 2.40 | 248.9 KB | 1.00 |  |  | Loss +139.6% |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | ClosedXML | 3.20 ms | 0.05 ms | 0.03 ms | 1.34 | 3.21 | 4392.7 KB | 17.65 |  |  | 34.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | EPPlus 4.5.3.3 | 10.94 ms |  |  | 4.59 | 10.99 |  |  |  |  | 358.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | MiniExcel | 11.27 ms | 1.32 ms | 0.76 ms | 4.73 | 11.33 | 46189.1 KB | 185.57 |  |  | 372.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | EPPlus | 39.79 ms | 3.37 ms | 1.95 ms | 16.69 | 39.99 | 43070.2 KB | 173.04 |  |  | 1568.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | Sylvan.Data.Excel | 2.08 ms | 0.19 ms | 0.11 ms | 0.60 | 1.00 | 518.6 KB | 0.46 |  |  | 40.2% faster than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | OfficeIMO.Excel | 3.47 ms | 0.06 ms | 0.04 ms | 1.00 | 1.67 | 1133.5 KB | 1.00 |  |  | Loss +67.1% |
| 2500 | speed-comparison | read | Other | shared-string-read | ExcelDataReader | 4.73 ms | 0.44 ms | 0.26 ms | 1.36 | 2.28 | 2603.0 KB | 2.30 |  |  | 36.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | MiniExcel | 5.80 ms | 0.65 ms | 0.37 ms | 1.67 | 2.79 | 7524.7 KB | 6.64 |  |  | 67.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | EPPlus 4.5.3.3 | 14.28 ms |  |  | 4.11 | 6.87 |  |  |  |  | 311.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | ClosedXML | 15.21 ms | 1.97 ms | 1.14 ms | 4.38 | 7.32 | 9498.0 KB | 8.38 |  |  | 338.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | EPPlus | 17.88 ms | 1.78 ms | 1.03 ms | 5.15 | 8.61 | 10371.6 KB | 9.15 |  |  | 414.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | Sylvan.Data.Excel | 4.43 ms | 0.21 ms | 0.12 ms | 0.89 | 1.00 | 655.2 KB | 1.32 |  |  | 11.0% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | OfficeIMO.Excel | 4.98 ms | 0.30 ms | 0.18 ms | 1.00 | 1.12 | 494.6 KB | 1.00 |  |  | Loss +12.4% |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | ExcelDataReader | 10.87 ms | 0.78 ms | 0.45 ms | 2.18 | 2.45 | 6081.3 KB | 12.30 |  |  | 118.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | MiniExcel | 13.55 ms | 1.38 ms | 0.80 ms | 2.72 | 3.06 | 18651.2 KB | 37.71 |  |  | 171.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | EPPlus | 23.27 ms | 1.72 ms | 1.00 ms | 4.67 | 5.25 | 12426.5 KB | 25.13 |  |  | 366.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | ClosedXML | 28.54 ms | 0.64 ms | 0.37 ms | 5.73 | 6.44 | 15356.9 KB | 31.05 |  |  | 472.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | Sylvan.Data.Excel | 6.36 ms | 0.19 ms | 0.11 ms | 0.75 | 1.00 | 2239.3 KB | 0.60 |  |  | 25.0% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | OfficeIMO.Excel | 8.48 ms | 1.59 ms | 0.92 ms | 1.00 | 1.33 | 3714.5 KB | 1.00 |  |  | Loss +33.3% |
| 2500 | speed-comparison | read | Range and table read | read-datatable | ExcelDataReader | 12.54 ms | 0.58 ms | 0.33 ms | 1.48 | 1.97 | 7665.3 KB | 2.06 |  |  | 47.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | MiniExcel | 13.53 ms | 1.19 ms | 0.69 ms | 1.59 | 2.13 | 18255.8 KB | 4.91 |  |  | 59.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | ClosedXML | 31.55 ms | 1.04 ms | 0.60 ms | 3.72 | 4.96 | 18311.6 KB | 4.93 |  |  | 272.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | EPPlus | 32.84 ms | 4.63 ms | 2.67 ms | 3.87 | 5.16 | 21736.0 KB | 5.85 |  |  | 287.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | EPPlus 4.5.3.3 | 38.04 ms |  |  | 4.48 | 5.98 |  |  |  |  | 348.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | Sylvan.Data.Excel | 4.81 ms | 0.03 ms | 0.02 ms | 0.97 | 1.00 | 733.5 KB | 1.11 |  |  | 2.6% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | OfficeIMO.Excel | 4.93 ms | 0.12 ms | 0.07 ms | 1.00 | 1.03 | 663.0 KB | 1.00 |  |  | Loss +2.7% |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | MiniExcel | 10.33 ms | 0.86 ms | 0.50 ms | 2.09 | 2.15 | 15842.3 KB | 23.89 |  |  | 109.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | ExcelDataReader | 10.76 ms | 0.24 ms | 0.14 ms | 2.18 | 2.24 | 6081.3 KB | 9.17 |  |  | 118.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | EPPlus | 21.94 ms | 0.21 ms | 0.12 ms | 4.45 | 4.57 | 13107.6 KB | 19.77 |  |  | 344.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | ClosedXML | 29.16 ms | 0.44 ms | 0.26 ms | 5.91 | 6.07 | 15456.0 KB | 23.31 |  |  | 490.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | OfficeIMO.Excel | 11.83 ms | 0.72 ms | 0.42 ms | 1.00 | 1.00 | 2812.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Range and table read | read-range | Sylvan.Data.Excel | 14.26 ms | 0.90 ms | 0.52 ms | 1.21 | 1.21 | 654.9 KB | 0.23 |  |  | 20.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | MiniExcel | 25.82 ms | 5.60 ms | 3.23 ms | 2.18 | 2.18 | 18651.4 KB | 6.63 |  |  | 118.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | ExcelDataReader | 26.47 ms | 4.35 ms | 2.51 ms | 2.24 | 2.24 | 6081.0 KB | 2.16 |  |  | 123.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | EPPlus | 34.88 ms | 7.21 ms | 4.16 ms | 2.95 | 2.95 | 20152.0 KB | 7.16 |  |  | 194.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | EPPlus 4.5.3.3 | 35.77 ms |  |  | 3.02 | 3.02 |  |  |  |  | 202.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | ClosedXML | 73.80 ms | 19.95 ms | 11.52 ms | 6.24 | 6.24 | 16843.3 KB | 5.99 |  |  | 523.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | Sylvan.Data.Excel | 5.65 ms | 0.05 ms | 0.03 ms | 0.69 | 1.00 | 750.3 KB | 0.26 |  |  | 30.9% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | OfficeIMO.Excel | 8.17 ms | 2.32 ms | 1.34 ms | 1.00 | 1.45 | 2871.4 KB | 1.00 |  |  | Loss +44.6% |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | ExcelDataReader | 10.78 ms | 0.06 ms | 0.04 ms | 1.32 | 1.91 | 6081.3 KB | 2.12 |  |  | 31.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | MiniExcel | 12.62 ms | 0.46 ms | 0.26 ms | 1.54 | 2.23 | 18651.7 KB | 6.50 |  |  | 54.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | EPPlus | 27.98 ms | 4.10 ms | 2.37 ms | 3.42 | 4.95 | 20152.1 KB | 7.02 |  |  | 242.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | ClosedXML | 30.65 ms | 1.19 ms | 0.69 ms | 3.75 | 5.42 | 16726.0 KB | 5.83 |  |  | 275.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | Sylvan.Data.Excel | 0.52 ms | 0.11 ms | 0.07 ms | 0.33 | 1.00 | 348.4 KB | 0.84 |  |  | 66.9% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | MiniExcel | 0.90 ms | 0.18 ms | 0.10 ms | 0.58 | 1.75 | 858.3 KB | 2.06 |  |  | 42.1% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | OfficeIMO.Excel | 1.56 ms | 0.05 ms | 0.03 ms | 1.00 | 3.02 | 416.1 KB | 1.00 |  |  | Loss +201.9% |
| 2500 | speed-comparison | read | Range and table read | read-top-range | ExcelDataReader | 4.61 ms | 0.31 ms | 0.18 ms | 2.95 | 8.92 | 1923.6 KB | 4.62 |  |  | 195.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | EPPlus | 26.65 ms | 5.30 ms | 3.06 ms | 17.08 | 51.56 | 12401.5 KB | 29.81 |  |  | 1608.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | ClosedXML | 30.08 ms | 1.48 ms | 0.86 ms | 19.28 | 58.20 | 15357.4 KB | 36.91 |  |  | 1828.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | EPPlus 4.5.3.3 | 31.81 ms |  |  | 20.39 | 61.55 |  |  |  |  | 1939.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | Sylvan.Data.Excel | 5.00 ms | 0.01 ms | 0.01 ms | 0.34 | 1.00 | 655.2 KB | 0.19 |  |  | 65.6% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | MiniExcel | 13.80 ms | 2.81 ms | 1.62 ms | 0.95 | 2.76 | 18651.7 KB | 5.28 |  |  | 5.2% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | OfficeIMO.Excel | 14.55 ms | 0.43 ms | 0.25 ms | 1.00 | 2.91 | 3535.1 KB | 1.00 |  |  | Loss +191.0% |
| 2500 | speed-comparison | read | Range and table read | read-used-range | ExcelDataReader | 14.99 ms | 7.06 ms | 4.07 ms | 1.03 | 3.00 | 6081.3 KB | 1.72 |  |  | 3.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | EPPlus | 33.20 ms | 4.96 ms | 2.86 ms | 2.28 | 6.64 | 20152.1 KB | 5.70 |  |  | 128.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | ClosedXML | 70.35 ms | 34.47 ms | 19.90 ms | 4.83 | 14.07 | 16805.2 KB | 4.75 |  |  | 383.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | Sylvan.Data.Excel | 4.37 ms | 0.01 ms | 0.01 ms | 0.89 | 1.00 | 655.2 KB | 1.32 |  |  | 11.3% faster than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | OfficeIMO.Excel | 4.93 ms | 0.09 ms | 0.05 ms | 1.00 | 1.13 | 497.9 KB | 1.00 |  |  | Loss +12.7% |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | ExcelDataReader | 10.58 ms | 0.08 ms | 0.05 ms | 2.15 | 2.42 | 6081.5 KB | 12.21 |  |  | 114.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | MiniExcel | 13.35 ms | 1.56 ms | 0.90 ms | 2.71 | 3.05 | 18651.2 KB | 37.46 |  |  | 170.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | EPPlus | 22.38 ms | 0.32 ms | 0.19 ms | 4.54 | 5.12 | 12426.5 KB | 24.96 |  |  | 354.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | ClosedXML | 29.16 ms | 1.30 ms | 0.75 ms | 5.91 | 6.67 | 15357.9 KB | 30.84 |  |  | 491.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | Sylvan.Data.Excel | 4.94 ms | 0.11 ms | 0.06 ms | 0.57 | 1.00 | 655.2 KB | 0.23 |  |  | 43.2% faster than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | OfficeIMO.Excel | 8.69 ms | 2.65 ms | 1.53 ms | 1.00 | 1.76 | 2891.5 KB | 1.00 |  |  | Loss +76.0% |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | ExcelDataReader | 11.25 ms | 0.53 ms | 0.31 ms | 1.30 | 2.28 | 6081.3 KB | 2.10 |  |  | 29.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | MiniExcel | 12.31 ms | 0.59 ms | 0.34 ms | 1.42 | 2.49 | 18651.7 KB | 6.45 |  |  | 41.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | EPPlus | 29.21 ms | 5.14 ms | 2.97 ms | 3.36 | 5.92 | 20152.0 KB | 6.97 |  |  | 236.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | ClosedXML | 30.93 ms | 1.41 ms | 0.81 ms | 3.56 | 6.27 | 16725.7 KB | 5.78 |  |  | 256.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | EPPlus 4.5.3.3 | 32.64 ms |  |  | 3.76 | 6.61 |  |  |  |  | 275.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | Sylvan.Data.Excel | 0.44 ms | 0.00 ms | 0.00 ms | 0.29 | 1.00 | 348.5 KB | 0.83 |  |  | 70.7% faster than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | MiniExcel | 0.78 ms | 0.05 ms | 0.03 ms | 0.52 | 1.79 | 858.3 KB | 2.05 |  |  | 47.6% faster than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | OfficeIMO.Excel | 1.49 ms | 0.05 ms | 0.03 ms | 1.00 | 3.42 | 419.5 KB | 1.00 |  |  | Loss +241.8% |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | ExcelDataReader | 4.27 ms | 0.08 ms | 0.04 ms | 2.86 | 9.76 | 1923.7 KB | 4.59 |  |  | 185.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | EPPlus | 22.65 ms | 1.10 ms | 0.63 ms | 15.16 | 51.82 | 12401.5 KB | 29.56 |  |  | 1416.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | ClosedXML | 28.80 ms | 1.81 ms | 1.05 ms | 19.28 | 65.88 | 15357.8 KB | 36.61 |  |  | 1827.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | EPPlus 4.5.3.3 | 29.40 ms |  |  | 19.68 | 67.26 |  |  |  |  | 1868.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | Sylvan.Data.Excel | 0.46 ms | 0.03 ms | 0.01 ms | 0.29 | 1.00 | 348.5 KB | 0.83 |  |  | 71.4% faster than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | MiniExcel | 0.73 ms | 0.00 ms | 0.00 ms | 0.46 | 1.60 | 858.3 KB | 2.04 |  |  | 54.2% faster than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | OfficeIMO.Excel | 1.60 ms | 0.06 ms | 0.03 ms | 1.00 | 3.49 | 420.2 KB | 1.00 |  |  | Loss +249.3% |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | ExcelDataReader | 4.30 ms | 0.05 ms | 0.03 ms | 2.68 | 9.36 | 1923.7 KB | 4.58 |  |  | 167.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | EPPlus | 21.95 ms | 1.13 ms | 0.65 ms | 13.68 | 47.77 | 12401.5 KB | 29.51 |  |  | 1267.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | ClosedXML | 28.60 ms | 1.10 ms | 0.64 ms | 17.82 | 62.24 | 15357.4 KB | 36.54 |  |  | 1681.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | Sylvan.Data.Excel | 6.80 ms | 3.32 ms | 1.92 ms | 0.86 | 1.00 | 895.3 KB | 0.35 |  |  | 14.0% faster than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | OfficeIMO.Excel | 7.90 ms | 1.84 ms | 1.06 ms | 1.00 | 1.16 | 2562.1 KB | 1.00 |  |  | Loss +16.2% |
| 2500 | speed-comparison | read | Typed object read | read-objects | ExcelDataReader | 12.76 ms | 3.59 ms | 2.07 ms | 1.61 | 1.88 | 6321.5 KB | 2.47 |  |  | 61.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | MiniExcel | 13.90 ms | 0.71 ms | 0.41 ms | 1.76 | 2.04 | 18463.3 KB | 7.21 |  |  | 75.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | ClosedXML | 29.64 ms | 1.60 ms | 0.92 ms | 3.75 | 4.36 | 16922.5 KB | 6.61 |  |  | 275.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | EPPlus | 30.86 ms | 8.69 ms | 5.02 ms | 3.91 | 4.54 | 21353.7 KB | 8.33 |  |  | 290.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | EPPlus 4.5.3.3 | 33.35 ms |  |  | 4.22 | 4.91 |  |  |  |  | 322.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | Sylvan.Data.Excel | 4.94 ms | 0.17 ms | 0.10 ms | 0.85 | 1.00 | 831.0 KB | 0.32 |  |  | 15.1% faster than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | OfficeIMO.Excel | 5.81 ms | 0.16 ms | 0.09 ms | 1.00 | 1.18 | 2562.4 KB | 1.00 |  |  | Loss +17.8% |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | ExcelDataReader | 11.16 ms | 0.39 ms | 0.22 ms | 1.92 | 2.26 | 6257.2 KB | 2.44 |  |  | 92.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | MiniExcel | 13.42 ms | 0.51 ms | 0.29 ms | 2.31 | 2.72 | 18399.1 KB | 7.18 |  |  | 130.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | EPPlus | 28.75 ms | 5.57 ms | 3.21 ms | 4.95 | 5.82 | 21334.1 KB | 8.33 |  |  | 394.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | ClosedXML | 29.20 ms | 1.49 ms | 0.86 ms | 5.02 | 5.91 | 16903.3 KB | 6.60 |  |  | 402.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | EPPlus 4.5.3.3 | 32.83 ms |  |  | 5.65 | 6.65 |  |  |  |  | 464.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | OfficeIMO.Excel | 5.40 ms | 0.87 ms | 0.50 ms | 1.00 | 1.00 | 1452.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | MiniExcel | 13.86 ms | 5.13 ms | 2.96 ms | 2.57 | 2.57 | 26638.3 KB | 18.34 |  |  | 156.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | EPPlus | 71.32 ms | 2.61 ms | 1.51 ms | 13.22 | 13.22 | 38271.2 KB | 26.34 |  |  | 1221.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | EPPlus 4.5.3.3 | 73.95 ms |  |  | 13.71 | 13.71 |  |  |  |  | 1270.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | ClosedXML | 90.59 ms | 3.46 ms | 2.00 ms | 16.79 | 16.79 | 58353.7 KB | 40.17 |  |  | 1579.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | OfficeIMO.Excel | 6.35 ms | 0.40 ms | 0.23 ms | 1.00 | 1.00 | 1591.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | MiniExcel | 17.92 ms | 1.72 ms | 0.99 ms | 2.82 | 2.82 | 32319.6 KB | 20.31 |  |  | 182.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | EPPlus | 116.48 ms | 18.62 ms | 10.75 ms | 18.33 | 18.33 | 43362.1 KB | 27.25 |  |  | 1733.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | ClosedXML | 231.05 ms | 28.08 ms | 16.21 ms | 36.36 | 36.36 | 56700.2 KB | 35.63 |  |  | 3536.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.42 ms | 0.38 ms | 0.22 ms | 1.00 | 1.00 | 1178.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 73.80 ms | 9.17 ms | 5.29 ms | 16.68 | 16.68 | 38271.3 KB | 32.47 |  |  | 1568.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 91.44 ms | 29.22 ms | 16.87 ms | 20.67 | 20.67 | 50919.4 KB | 43.20 |  |  | 1967.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.45 ms | 0.20 ms | 0.11 ms | 1.00 | 1.00 | 1176.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 60.80 ms | 2.51 ms | 1.45 ms | 13.66 | 13.66 | 38271.3 KB | 32.52 |  |  | 1266.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 68.27 ms | 5.21 ms | 3.01 ms | 15.34 | 15.34 | 50919.3 KB | 43.27 |  |  | 1434.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.23 ms | 0.38 ms | 0.22 ms | 1.00 | 1.00 | 1177.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 49.91 ms | 7.10 ms | 4.10 ms | 11.80 | 11.80 | 28532.3 KB | 24.22 |  |  | 1080.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 58.72 ms | 3.30 ms | 1.91 ms | 13.89 | 13.89 | 27236.1 KB | 23.12 |  |  | 1288.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | OfficeIMO.Excel | 2.91 ms | 0.11 ms | 0.06 ms | 1.00 | 1.00 | 1454.0 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | ClosedXML | 16.46 ms | 4.25 ms | 2.45 ms | 5.67 | 5.67 | 9951.5 KB | 6.84 |  |  | 466.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | EPPlus | 24.85 ms | 3.91 ms | 2.26 ms | 8.55 | 8.55 | 11703.6 KB | 8.05 |  |  | 755.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | OfficeIMO.Excel | 2.98 ms | 0.92 ms | 0.53 ms | 1.00 | 1.00 | 946.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | ClosedXML | 11.93 ms | 1.06 ms | 0.61 ms | 4.00 | 4.00 | 9169.1 KB | 9.68 |  |  | 299.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | EPPlus 4.5.3.3 | 15.35 ms |  |  | 5.14 | 5.14 |  |  |  |  | 414.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | EPPlus | 27.95 ms | 2.33 ms | 1.35 ms | 9.37 | 9.37 | 12829.2 KB | 13.55 |  |  | 836.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | OfficeIMO.Excel | 2.90 ms | 0.03 ms | 0.02 ms | 1.00 | 1.00 | 1431.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | EPPlus 4.5.3.3 | 18.26 ms |  |  | 6.29 | 6.29 |  |  |  |  | 529.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | ClosedXML | 19.84 ms | 2.85 ms | 1.65 ms | 6.84 | 6.84 | 11879.0 KB | 8.30 |  |  | 583.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | EPPlus | 29.60 ms | 2.38 ms | 1.37 ms | 10.20 | 10.20 | 15577.1 KB | 10.88 |  |  | 920.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | OfficeIMO.Excel | 4.69 ms | 1.65 ms | 0.95 ms | 1.00 | 1.00 | 1271.0 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | ClosedXML | 16.30 ms | 2.41 ms | 1.39 ms | 3.47 | 3.47 | 11288.3 KB | 8.88 |  |  | 247.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | EPPlus | 30.71 ms | 5.22 ms | 3.01 ms | 6.55 | 6.55 | 14894.0 KB | 11.72 |  |  | 554.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 3.26 ms | 0.42 ms | 0.24 ms | 1.00 | 1.00 | 1271.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | ClosedXML | 15.47 ms | 2.05 ms | 1.18 ms | 4.74 | 4.74 | 11288.3 KB | 8.88 |  |  | 374.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | EPPlus | 27.73 ms | 1.60 ms | 0.92 ms | 8.50 | 8.50 | 14894.0 KB | 11.72 |  |  | 750.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | OfficeIMO.Excel | 2.65 ms | 0.43 ms | 0.25 ms | 1.00 | 1.00 | 964.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | ClosedXML | 11.44 ms | 0.37 ms | 0.21 ms | 4.32 | 4.32 | 9013.2 KB | 9.34 |  |  | 331.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | EPPlus 4.5.3.3 | 16.51 ms |  |  | 6.23 | 6.23 |  |  |  |  | 523.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | EPPlus | 26.49 ms | 5.76 ms | 3.33 ms | 10.00 | 10.00 | 12761.4 KB | 13.23 |  |  | 899.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | OfficeIMO.Excel | 3.28 ms | 0.23 ms | 0.13 ms | 1.00 | 1.00 | 1262.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | ClosedXML | 15.48 ms | 0.24 ms | 0.14 ms | 4.72 | 4.72 | 9703.1 KB | 7.69 |  |  | 371.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | EPPlus 4.5.3.3 | 17.51 ms |  |  | 5.34 | 5.34 |  |  |  |  | 433.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | EPPlus | 25.73 ms | 1.57 ms | 0.90 ms | 7.84 | 7.84 | 14654.6 KB | 11.61 |  |  | 684.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | OfficeIMO.Excel | 5.53 ms | 0.33 ms | 0.19 ms | 1.00 | 1.00 | 1576.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | MiniExcel | 15.07 ms | 4.70 ms | 2.71 ms | 2.73 | 2.73 | 29038.3 KB | 18.42 |  |  | 172.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | ClosedXML | 44.46 ms | 6.13 ms | 3.54 ms | 8.05 | 8.05 | 18905.4 KB | 11.99 |  |  | 704.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | EPPlus | 58.18 ms | 5.69 ms | 3.29 ms | 10.53 | 10.53 | 17634.2 KB | 11.19 |  |  | 953.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | OfficeIMO.Excel | 8.58 ms | 1.50 ms | 0.87 ms | 1.00 | 1.00 | 2392.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | MiniExcel | 12.68 ms | 0.78 ms | 0.45 ms | 1.48 | 1.48 | 29737.9 KB | 12.43 |  |  | 47.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | ClosedXML | 57.84 ms | 4.08 ms | 2.36 ms | 6.74 | 6.74 | 27402.4 KB | 11.46 |  |  | 574.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | EPPlus | 58.03 ms | 7.74 ms | 4.47 ms | 6.76 | 6.76 | 21826.0 KB | 9.12 |  |  | 576.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | OfficeIMO.Excel | 5.44 ms | 0.56 ms | 0.32 ms | 1.00 | 1.00 | 1579.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | MiniExcel | 18.88 ms | 5.27 ms | 3.04 ms | 3.47 | 3.47 | 29043.9 KB | 18.39 |  |  | 246.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | EPPlus 4.5.3.3 | 42.06 ms |  |  | 7.73 | 7.73 |  |  |  |  | 672.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | ClosedXML | 86.38 ms | 16.14 ms | 9.32 ms | 15.87 | 15.87 | 18870.3 KB | 11.95 |  |  | 1486.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | EPPlus | 89.88 ms | 3.72 ms | 2.15 ms | 16.51 | 16.51 | 19364.2 KB | 12.26 |  |  | 1550.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | OfficeIMO.Excel | 4.40 ms | 0.34 ms | 0.20 ms | 1.00 | 1.00 | 1446.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | MiniExcel | 10.69 ms | 1.93 ms | 1.11 ms | 2.43 | 2.43 | 23035.1 KB | 15.93 |  |  | 142.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | EPPlus 4.5.3.3 | 32.13 ms |  |  | 7.30 | 7.30 |  |  |  |  | 630.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | ClosedXML | 37.82 ms | 4.98 ms | 2.87 ms | 8.59 | 8.59 | 19001.4 KB | 13.14 |  |  | 759.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | EPPlus | 39.41 ms | 2.22 ms | 1.28 ms | 8.95 | 8.95 | 16579.4 KB | 11.46 |  |  | 795.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | OfficeIMO.Excel | 5.02 ms | 0.29 ms | 0.17 ms | 1.00 | 1.00 | 1420.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | LargeXlsx | 9.94 ms | 2.08 ms | 1.20 ms | 1.98 | 1.98 | 1141.0 KB | 0.80 |  |  | 98.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | MiniExcel | 11.92 ms | 3.35 ms | 1.94 ms | 2.37 | 2.37 | 23053.6 KB | 16.23 |  |  | 137.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | ClosedXML | 40.89 ms | 10.64 ms | 6.14 ms | 8.15 | 8.15 | 11573.0 KB | 8.15 |  |  | 714.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | EPPlus 4.5.3.3 | 41.55 ms |  |  | 8.28 | 8.28 |  |  |  |  | 727.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | EPPlus | 51.81 ms | 9.18 ms | 5.30 ms | 10.32 | 10.32 | 16580.6 KB | 11.67 |  |  | 932.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | OfficeIMO.Excel | 5.05 ms | 0.97 ms | 0.56 ms | 1.00 | 1.00 | 1158.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | MiniExcel | 9.14 ms | 1.90 ms | 1.10 ms | 1.81 | 1.81 | 22780.4 KB | 19.66 |  |  | 80.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | ClosedXML | 43.54 ms | 5.32 ms | 3.07 ms | 8.62 | 8.62 | 18726.0 KB | 16.16 |  |  | 761.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | EPPlus | 49.24 ms | 12.90 ms | 7.45 ms | 9.74 | 9.74 | 16307.0 KB | 14.07 |  |  | 874.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | OfficeIMO.Excel | 4.82 ms | 0.28 ms | 0.16 ms | 1.00 | 1.00 | 1432.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | MiniExcel | 9.28 ms | 1.09 ms | 0.63 ms | 1.93 | 1.93 | 23053.9 KB | 16.10 |  |  | 92.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | EPPlus 4.5.3.3 | 42.06 ms |  |  | 8.73 | 8.73 |  |  |  |  | 773.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | ClosedXML | 43.22 ms | 2.13 ms | 1.23 ms | 8.97 | 8.97 | 18999.7 KB | 13.27 |  |  | 797.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | EPPlus | 51.11 ms | 0.98 ms | 0.57 ms | 10.61 | 10.61 | 16580.5 KB | 11.58 |  |  | 961.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | OfficeIMO.Excel | 7.21 ms | 1.00 ms | 0.58 ms | 1.00 | 1.00 | 1234.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | MiniExcel | 17.56 ms | 4.52 ms | 2.61 ms | 2.44 | 2.44 | 26815.9 KB | 21.72 |  |  | 143.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | EPPlus 4.5.3.3 | 112.17 ms |  |  | 15.56 | 15.56 |  |  |  |  | 1455.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | EPPlus | 152.33 ms | 6.80 ms | 3.92 ms | 21.13 | 21.13 | 49086.2 KB | 39.76 |  |  | 2013.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | ClosedXML | 277.96 ms | 24.07 ms | 13.90 ms | 38.56 | 38.56 | 58343.1 KB | 47.25 |  |  | 3755.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | OfficeIMO.Excel | 2.81 ms | 0.13 ms | 0.07 ms | 1.00 | 1.00 | 1409.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | EPPlus 4.5.3.3 | 20.04 ms |  |  | 7.13 | 7.13 |  |  |  |  | 613.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | ClosedXML | 21.99 ms | 6.20 ms | 3.58 ms | 7.82 | 7.82 | 12031.2 KB | 8.54 |  |  | 682.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | EPPlus | 41.27 ms | 1.74 ms | 1.00 ms | 14.68 | 14.68 | 18036.5 KB | 12.80 |  |  | 1367.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-dictionary-objects-table-direct | OfficeIMO.Excel | 7.91 ms | 2.10 ms | 1.21 ms | 1.00 | 1.00 | 1723.0 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | LargeXlsx | 4.58 ms | 1.07 ms | 0.62 ms | 0.82 | 1.00 | 794.5 KB | 0.39 |  |  | 18.4% faster than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 5.61 ms | 0.58 ms | 0.33 ms | 1.00 | 1.22 | 2013.3 KB | 1.00 |  |  | Loss +22.5% |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | MiniExcel | 8.18 ms | 0.23 ms | 0.13 ms | 1.46 | 1.79 | 25181.4 KB | 12.51 |  |  | 45.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | ClosedXML | 39.16 ms | 2.36 ms | 1.36 ms | 6.98 | 8.55 | 16965.4 KB | 8.43 |  |  | 597.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | EPPlus | 47.42 ms | 4.62 ms | 2.67 ms | 8.45 | 10.35 | 20030.7 KB | 9.95 |  |  | 745.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | LargeXlsx | 4.33 ms | 0.40 ms | 0.23 ms | 0.80 | 1.00 | 794.5 KB | 0.59 |  |  | 19.6% faster than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 5.39 ms | 0.48 ms | 0.28 ms | 1.00 | 1.24 | 1339.3 KB | 1.00 |  |  | Loss +24.3% |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | MiniExcel | 9.12 ms | 1.44 ms | 0.83 ms | 1.69 | 2.10 | 25181.4 KB | 18.80 |  |  | 69.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | ClosedXML | 38.16 ms | 4.16 ms | 2.40 ms | 7.08 | 8.81 | 16965.4 KB | 12.67 |  |  | 608.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | EPPlus | 48.34 ms | 5.86 ms | 3.39 ms | 8.97 | 11.16 | 20030.7 KB | 14.96 |  |  | 797.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 20.40 ms | 2.88 ms | 1.66 ms | 1.00 | 1.00 | 4333.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | LargeXlsx | 21.01 ms | 2.53 ms | 1.46 ms | 1.03 | 1.03 | 2802.7 KB | 0.65 |  |  | 3.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | MiniExcel | 35.38 ms | 2.80 ms | 1.62 ms | 1.73 | 1.73 | 48404.7 KB | 11.17 |  |  | 73.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | ClosedXML | 122.84 ms | 8.17 ms | 4.72 ms | 6.02 | 6.02 | 51639.0 KB | 11.92 |  |  | 502.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | EPPlus | 174.10 ms | 5.06 ms | 2.92 ms | 8.54 | 8.54 | 69073.3 KB | 15.94 |  |  | 753.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | LargeXlsx | 1.37 ms | 0.05 ms | 0.03 ms | 0.55 | 1.00 | 288.4 KB | 0.17 |  |  | 44.6% faster than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | OfficeIMO.Excel | 2.47 ms | 0.05 ms | 0.03 ms | 1.00 | 1.81 | 1657.3 KB | 1.00 |  |  | Loss +80.6% |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | MiniExcel | 3.88 ms | 0.02 ms | 0.01 ms | 1.57 | 2.83 | 19701.8 KB | 11.89 |  |  | 56.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | ClosedXML | 13.41 ms | 0.40 ms | 0.23 ms | 5.43 | 9.80 | 11189.4 KB | 6.75 |  |  | 442.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | EPPlus 4.5.3.3 | 20.76 ms |  |  | 8.40 | 15.17 |  |  |  |  | 740.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | EPPlus | 22.93 ms | 1.40 ms | 0.81 ms | 9.28 | 16.75 | 14283.5 KB | 8.62 |  |  | 827.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 1.30 ms | 0.06 ms | 0.03 ms | 0.78 | 1.00 | 439.0 KB | 0.48 |  |  | 21.9% faster than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 1.66 ms | 0.02 ms | 0.01 ms | 1.00 | 1.28 | 923.6 KB | 1.00 |  |  | Loss +28.0% |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | ClosedXML | 15.63 ms | 4.68 ms | 2.70 ms | 9.40 | 12.03 | 10227.8 KB | 11.07 |  |  | 840.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | EPPlus | 24.31 ms | 6.28 ms | 3.63 ms | 14.62 | 18.71 | 12985.2 KB | 14.06 |  |  | 1362.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | LargeXlsx | 3.65 ms | 0.21 ms | 0.12 ms | 0.79 | 1.00 | 750.2 KB | 0.43 |  |  | 21.2% faster than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 4.64 ms | 0.42 ms | 0.24 ms | 1.00 | 1.27 | 1752.7 KB | 1.00 |  |  | Loss +27.0% |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | MiniExcel | 10.89 ms | 1.15 ms | 0.66 ms | 2.35 | 2.98 | 23212.8 KB | 13.24 |  |  | 134.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 35.16 ms |  |  | 7.58 | 9.63 |  |  |  |  | 658.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | ClosedXML | 36.01 ms | 7.83 ms | 4.52 ms | 7.76 | 9.86 | 22213.3 KB | 12.67 |  |  | 676.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | EPPlus | 53.97 ms | 4.39 ms | 2.53 ms | 11.64 | 14.77 | 24626.6 KB | 14.05 |  |  | 1063.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | OfficeIMO.Excel | 2.43 ms | 0.05 ms | 0.03 ms | 1.00 | 1.00 | 1165.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | ClosedXML | 17.48 ms | 2.43 ms | 1.40 ms | 7.20 | 7.20 | 11288.3 KB | 9.68 |  |  | 619.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | EPPlus | 27.79 ms | 3.55 ms | 2.05 ms | 11.44 | 11.44 | 14893.8 KB | 12.78 |  |  | 1043.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | OfficeIMO.Excel | 5.34 ms | 0.99 ms | 0.57 ms | 1.00 | 1.00 | 1434.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | LargeXlsx | 7.74 ms | 3.30 ms | 1.91 ms | 1.45 | 1.45 | 1024.5 KB | 0.71 |  |  | 44.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | Sylvan.Data.Excel | 7.91 ms | 7.25 ms | 4.19 ms | 1.48 | 1.48 | 750.5 KB | 0.52 |  |  | 48.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | MiniExcel | 8.71 ms | 1.19 ms | 0.69 ms | 1.63 | 1.63 | 23034.8 KB | 16.06 |  |  | 63.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | EPPlus 4.5.3.3 | 33.37 ms |  |  | 6.24 | 6.24 |  |  |  |  | 524.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | ClosedXML | 34.33 ms | 7.08 ms | 4.09 ms | 6.42 | 6.42 | 11573.0 KB | 8.07 |  |  | 542.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | EPPlus | 44.66 ms | 8.81 ms | 5.09 ms | 8.36 | 8.36 | 16579.0 KB | 11.56 |  |  | 735.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | OfficeIMO.Excel | 5.41 ms | 0.46 ms | 0.27 ms | 1.00 | 1.00 | 1652.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | LargeXlsx | 10.34 ms | 0.73 ms | 0.42 ms | 1.91 | 1.91 | 1115.8 KB | 0.68 |  |  | 91.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | MiniExcel | 12.31 ms | 0.63 ms | 0.36 ms | 2.28 | 2.28 | 29737.9 KB | 17.99 |  |  | 127.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | ClosedXML | 58.39 ms | 7.33 ms | 4.23 ms | 10.80 | 10.80 | 27402.9 KB | 16.58 |  |  | 980.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | EPPlus | 65.19 ms | 4.47 ms | 2.58 ms | 12.06 | 12.06 | 21826.0 KB | 13.21 |  |  | 1106.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | OfficeIMO.Excel | 5.04 ms | 0.47 ms | 0.27 ms | 1.00 | 1.00 | 1508.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | LargeXlsx | 5.22 ms | 1.39 ms | 0.80 ms | 1.04 | 1.04 | 849.6 KB | 0.56 |  |  | 3.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | MiniExcel | 17.52 ms | 1.92 ms | 1.11 ms | 3.48 | 3.48 | 35910.0 KB | 23.81 |  |  | 247.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | ClosedXML | 96.02 ms | 3.24 ms | 1.87 ms | 19.04 | 19.04 | 71470.2 KB | 47.38 |  |  | 1804.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | OfficeIMO.Excel | 2.09 ms | 0.09 ms | 0.05 ms | 1.00 | 1.00 | 2111.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | MiniExcel | 3.79 ms | 0.24 ms | 0.14 ms | 1.81 | 1.81 | 21128.5 KB | 10.00 |  |  | 81.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | ClosedXML | 11.41 ms | 0.54 ms | 0.31 ms | 5.45 | 5.45 | 11291.2 KB | 5.35 |  |  | 444.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | EPPlus 4.5.3.3 | 13.84 ms |  |  | 6.61 | 6.61 |  |  |  |  | 560.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | EPPlus | 20.79 ms | 0.50 ms | 0.29 ms | 9.93 | 9.93 | 12730.2 KB | 6.03 |  |  | 892.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | OfficeIMO.Excel | 2.99 ms | 0.10 ms | 0.06 ms | 1.00 | 1.00 | 2283.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | ClosedXML | 16.22 ms | 4.82 ms | 2.78 ms | 5.42 | 5.42 | 11291.2 KB | 4.94 |  |  | 441.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | EPPlus 4.5.3.3 | 17.27 ms |  |  | 5.77 | 5.77 |  |  |  |  | 476.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | EPPlus | 25.78 ms | 4.54 ms | 2.62 ms | 8.61 | 8.61 | 12730.1 KB | 5.57 |  |  | 760.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | OfficeIMO.Excel | 3.11 ms | 0.64 ms | 0.37 ms | 1.00 | 1.00 | 2206.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | ClosedXML | 20.50 ms | 3.64 ms | 2.10 ms | 6.58 | 6.58 | 13119.1 KB | 5.95 |  |  | 558.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | EPPlus | 28.35 ms | 2.03 ms | 1.17 ms | 9.10 | 9.10 | 13793.7 KB | 6.25 |  |  | 810.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | OfficeIMO.Excel | 2.57 ms | 0.62 ms | 0.36 ms | 1.00 | 1.00 | 1246.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | ClosedXML | 12.18 ms | 1.73 ms | 1.00 ms | 4.74 | 4.74 | 9218.5 KB | 7.39 |  |  | 374.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | EPPlus | 19.46 ms | 0.54 ms | 0.31 ms | 7.58 | 7.58 | 11265.6 KB | 9.04 |  |  | 657.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | LargeXlsx | 2.95 ms | 0.05 ms | 0.03 ms | 0.82 | 1.00 | 750.2 KB | 0.52 |  |  | 18.1% faster than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 3.60 ms | 0.18 ms | 0.10 ms | 1.00 | 1.22 | 1440.7 KB | 1.00 |  |  | Loss +22.1% |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | MiniExcel | 7.76 ms | 0.68 ms | 0.39 ms | 2.15 | 2.63 | 23213.3 KB | 16.11 |  |  | 115.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | ClosedXML | 25.77 ms | 0.27 ms | 0.15 ms | 7.15 | 8.73 | 11573.0 KB | 8.03 |  |  | 615.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | EPPlus | 34.21 ms | 0.18 ms | 0.10 ms | 9.50 | 11.60 | 16579.3 KB | 11.51 |  |  | 849.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 35.19 ms |  |  | 9.77 | 11.93 |  |  |  |  | 876.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | LargeXlsx | 3.50 ms | 0.14 ms | 0.08 ms | 0.76 | 1.00 | 750.2 KB | 0.64 |  |  | 24.0% faster than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | OfficeIMO.Excel | 4.60 ms | 1.02 ms | 0.59 ms | 1.00 | 1.32 | 1170.9 KB | 1.00 |  |  | Loss +31.5% |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | MiniExcel | 8.79 ms | 0.68 ms | 0.39 ms | 1.91 | 2.51 | 23213.3 KB | 19.83 |  |  | 90.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | ClosedXML | 31.75 ms | 1.39 ms | 0.81 ms | 6.90 | 9.07 | 11573.0 KB | 9.88 |  |  | 589.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | EPPlus 4.5.3.3 | 32.91 ms |  |  | 7.15 | 9.40 |  |  |  |  | 614.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | EPPlus | 45.59 ms | 7.64 ms | 4.41 ms | 9.90 | 13.02 | 16579.3 KB | 14.16 |  |  | 890.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 3.15 ms | 0.15 ms | 0.09 ms | 0.68 | 1.00 | 750.2 KB | 0.64 |  |  | 32.4% faster than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 4.66 ms | 0.12 ms | 0.07 ms | 1.00 | 1.48 | 1169.1 KB | 1.00 |  |  | Loss +47.8% |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | MiniExcel | 9.41 ms | 2.30 ms | 1.33 ms | 2.02 | 2.98 | 23213.3 KB | 19.86 |  |  | 101.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | ClosedXML | 28.99 ms | 2.16 ms | 1.25 ms | 6.22 | 9.19 | 11573.0 KB | 9.90 |  |  | 521.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | EPPlus | 41.06 ms | 3.13 ms | 1.81 ms | 8.81 | 13.02 | 16579.3 KB | 14.18 |  |  | 780.6% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | OfficeIMO.Excel | 60.77 ms | 10.97 ms | 6.33 ms | 1.00 | 1.00 | 23699.6 KB | 1.00 |  |  | Win |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | Sylvan.Data.Excel | 61.82 ms | 4.31 ms | 2.49 ms | 1.02 | 1.02 | 394.1 KB | 0.02 |  |  | Tie vs OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | ExcelDataReader | 164.62 ms | 37.31 ms | 21.54 ms | 2.71 | 2.71 | 69517.4 KB | 2.93 |  |  | 170.9% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | MiniExcel | 208.69 ms | 44.53 ms | 25.71 ms | 3.43 | 3.43 | 215349.0 KB | 9.09 |  |  | 243.4% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | Sylvan.Data.Excel | 47.40 ms | 14.26 ms | 8.24 ms | 0.81 | 1.00 | 394.1 KB | 0.02 |  |  | 18.7% faster than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | OfficeIMO.Excel | 58.28 ms | 13.32 ms | 7.69 ms | 1.00 | 1.23 | 24481.9 KB | 1.00 |  |  | Loss +22.9% |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | ExcelDataReader | 154.40 ms | 34.84 ms | 20.12 ms | 2.65 | 3.26 | 69517.4 KB | 2.84 |  |  | 165.0% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | MiniExcel | 202.82 ms | 45.80 ms | 26.44 ms | 3.48 | 4.28 | 215349.0 KB | 8.80 |  |  | 248.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | append-plain-rows | LargeXlsx | 14.09 ms | 3.29 ms | 1.90 ms | 0.68 | 1.00 | 2763.0 KB | 0.24 | 605.0 KB | 0.97 | 32.5% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | append-plain-rows | OfficeIMO.Excel | 20.87 ms | 7.00 ms | 4.04 ms | 1.00 | 1.48 | 11671.6 KB | 1.00 | 622.5 KB | 1.00 | Loss +48.1% |
| 25000 | package-profile | package | Package size | append-plain-rows | MiniExcel | 40.13 ms | 16.84 ms | 9.72 ms | 1.92 | 2.85 | 58233.0 KB | 4.99 | 642.3 KB | 1.03 | 92.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | append-plain-rows | ClosedXML | 146.77 ms | 8.08 ms | 4.67 ms | 7.03 | 10.41 | 104225.1 KB | 8.93 | 540.6 KB | 0.87 | 603.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | append-plain-rows | EPPlus | 223.35 ms | 6.99 ms | 4.04 ms | 10.70 | 15.85 | 100275.4 KB | 8.59 | 525.6 KB | 0.84 | 970.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | autofit-existing | OfficeIMO.Excel | 342.58 ms | 13.12 ms | 7.58 ms | 1.00 | 1.00 | 138378.6 KB | 1.00 | 1429.3 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | autofit-existing | EPPlus | 464.86 ms | 24.15 ms | 13.94 ms | 1.36 | 1.36 | 250878.4 KB | 1.81 | 1091.0 KB | 0.76 | 35.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | autofit-existing | ClosedXML | 1356.48 ms | 124.28 ms | 71.75 ms | 3.96 | 3.96 | 829585.3 KB | 6.00 | 1140.9 KB | 0.80 | 296.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | large-shared-strings | OfficeIMO.Excel | 14.99 ms | 0.38 ms | 0.22 ms | 1.00 | 1.00 | 15416.5 KB | 1.00 | 529.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | large-shared-strings | MiniExcel | 31.72 ms | 2.11 ms | 1.22 ms | 2.12 | 2.12 | 73751.2 KB | 4.78 | 581.0 KB | 1.10 | 111.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | large-shared-strings | ClosedXML | 118.65 ms | 16.50 ms | 9.53 ms | 7.91 | 7.91 | 104233.3 KB | 6.76 | 460.1 KB | 0.87 | 691.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | large-shared-strings | EPPlus | 197.43 ms | 14.93 ms | 8.62 ms | 13.17 | 13.17 | 84343.7 KB | 5.47 | 444.7 KB | 0.84 | 1217.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-autofilter | OfficeIMO.Excel | 30.84 ms | 0.98 ms | 0.57 ms | 1.00 | 1.00 | 11326.3 KB | 1.00 | 1428.5 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-autofilter | ClosedXML | 286.03 ms | 1.36 ms | 0.78 ms | 9.27 | 9.27 | 210655.8 KB | 18.60 | 1140.0 KB | 0.80 | 827.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-autofilter | EPPlus | 336.19 ms | 3.33 ms | 1.92 ms | 10.90 | 10.90 | 211804.2 KB | 18.70 | 1090.1 KB | 0.76 | 990.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-charts | OfficeIMO.Excel | 41.92 ms | 2.52 ms | 1.45 ms | 1.00 | 1.00 | 12323.4 KB | 1.00 | 1433.6 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-charts | EPPlus | 428.40 ms | 7.50 ms | 4.33 ms | 10.22 | 10.22 | 214836.4 KB | 17.43 | 1092.9 KB | 0.76 | 922.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-conditional-formatting | OfficeIMO.Excel | 30.63 ms | 0.49 ms | 0.28 ms | 1.00 | 1.00 | 11391.3 KB | 1.00 | 1428.8 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-conditional-formatting | ClosedXML | 297.68 ms | 9.74 ms | 5.62 ms | 9.72 | 9.72 | 210703.7 KB | 18.50 | 1140.1 KB | 0.80 | 871.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-conditional-formatting | EPPlus | 362.33 ms | 6.22 ms | 3.59 ms | 11.83 | 11.83 | 211845.7 KB | 18.60 | 1090.2 KB | 0.76 | 1082.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-data-validation | OfficeIMO.Excel | 39.78 ms | 7.11 ms | 4.11 ms | 1.00 | 1.00 | 11342.1 KB | 1.00 | 1428.6 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-data-validation | ClosedXML | 352.17 ms | 40.89 ms | 23.61 ms | 8.85 | 8.85 | 210664.6 KB | 18.57 | 1140.1 KB | 0.80 | 785.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-data-validation | EPPlus | 438.54 ms | 6.83 ms | 3.94 ms | 11.02 | 11.02 | 211789.9 KB | 18.67 | 1090.1 KB | 0.76 | 1002.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-freeze-panes | OfficeIMO.Excel | 31.47 ms | 2.32 ms | 1.34 ms | 1.00 | 1.00 | 11331.2 KB | 1.00 | 1428.6 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-freeze-panes | ClosedXML | 289.57 ms | 7.40 ms | 4.27 ms | 9.20 | 9.20 | 210641.4 KB | 18.59 | 1140.0 KB | 0.80 | 820.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-freeze-panes | EPPlus | 345.28 ms | 6.16 ms | 3.55 ms | 10.97 | 10.97 | 211838.4 KB | 18.70 | 1090.2 KB | 0.76 | 997.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-pivot-table | OfficeIMO.Excel | 117.92 ms | 11.87 ms | 6.85 ms | 1.00 | 1.00 | 100856.7 KB | 1.00 | 1431.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-pivot-table | EPPlus | 439.09 ms | 47.93 ms | 27.67 ms | 3.72 | 3.72 | 230731.6 KB | 2.29 | 1093.4 KB | 0.76 | 272.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-all-in-one | OfficeIMO.Excel | 117.37 ms | 2.33 ms | 1.35 ms | 1.00 | 1.00 | 102341.4 KB | 1.00 | 1437.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-all-in-one | EPPlus | 474.97 ms | 4.99 ms | 2.88 ms | 4.05 | 4.05 | 277073.9 KB | 2.71 | 1097.8 KB | 0.76 | 304.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-core | OfficeIMO.Excel | 38.60 ms | 6.20 ms | 3.58 ms | 1.00 | 1.00 | 11487.6 KB | 1.00 | 1430.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-core | EPPlus | 400.17 ms | 20.39 ms | 11.77 ms | 10.37 | 10.37 | 255063.6 KB | 22.20 | 1091.5 KB | 0.76 | 936.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-core | ClosedXML | 887.72 ms | 90.31 ms | 52.14 ms | 23.00 | 23.00 | 680111.7 KB | 59.20 | 1141.3 KB | 0.80 | 2199.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook | OfficeIMO.Excel | 48.14 ms | 0.65 ms | 0.37 ms | 1.00 | 1.00 | 14316.9 KB | 1.00 | 1857.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | report-workbook | EPPlus | 545.30 ms | 3.85 ms | 2.22 ms | 11.33 | 11.33 | 364641.3 KB | 25.47 | 1517.2 KB | 0.82 | 1032.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook-core | OfficeIMO.Excel | 55.61 ms | 8.40 ms | 4.85 ms | 1.00 | 1.00 | 10803.9 KB | 1.00 | 1850.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | report-workbook-core | EPPlus | 673.82 ms | 45.65 ms | 26.36 ms | 12.12 | 12.12 | 342774.1 KB | 31.73 | 1512.6 KB | 0.82 | 1111.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook-core | ClosedXML | 1341.84 ms | 187.61 ms | 108.32 ms | 24.13 | 24.13 | 975769.4 KB | 90.32 | 1579.8 KB | 0.85 | 2312.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook-datatable | OfficeIMO.Excel | 148.11 ms | 57.77 ms | 33.35 ms | 1.00 | 1.00 | 17067.8 KB | 1.00 | 1857.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | report-workbook-datatable | EPPlus | 1336.40 ms | 199.68 ms | 115.28 ms | 9.02 | 9.02 | 247775.1 KB | 14.52 | 1517.2 KB | 0.82 | 802.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook-datatable-core | OfficeIMO.Excel | 61.19 ms | 17.52 ms | 10.11 ms | 1.00 | 1.00 | 13557.9 KB | 1.00 | 1850.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | report-workbook-datatable-core | EPPlus | 676.56 ms | 157.08 ms | 90.69 ms | 11.06 | 11.06 | 225953.6 KB | 16.67 | 1512.6 KB | 0.82 | 1005.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook-datatable-core | ClosedXML | 1455.98 ms | 214.48 ms | 123.83 ms | 23.80 | 23.80 | 832229.4 KB | 61.38 | 1579.8 KB | 0.85 | 2279.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | LargeXlsx | 39.31 ms | 1.07 ms | 0.62 ms | 0.90 | 1.00 | 10787.2 KB | 0.93 | 2444.6 KB | 1.10 | 9.6% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | OfficeIMO.Excel | 43.48 ms | 2.25 ms | 1.30 ms | 1.00 | 1.11 | 11539.8 KB | 1.00 | 2228.8 KB | 1.00 | Loss +10.6% |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | MiniExcel | 148.34 ms | 3.01 ms | 1.74 ms | 3.41 | 3.77 | 226867.6 KB | 19.66 | 2410.6 KB | 1.08 | 241.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | ClosedXML | 926.37 ms | 38.09 ms | 21.99 ms | 21.31 | 23.56 | 759810.2 KB | 65.84 | 2581.2 KB | 1.16 | 2030.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-bulk-report | OfficeIMO.Excel | 34.49 ms | 1.01 ms | 0.59 ms | 1.00 | 1.00 | 11393.2 KB | 1.00 | 1429.3 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-bulk-report | MiniExcel | 64.38 ms | 0.34 ms | 0.20 ms | 1.87 | 1.87 | 125541.4 KB | 11.02 | 1521.1 KB | 1.06 | 86.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-bulk-report | EPPlus | 382.74 ms | 5.81 ms | 3.35 ms | 11.10 | 11.10 | 254887.6 KB | 22.37 | 1091.0 KB | 0.76 | 1009.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-bulk-report | ClosedXML | 761.72 ms | 3.93 ms | 2.27 ms | 22.09 | 22.09 | 565944.6 KB | 49.67 | 1140.9 KB | 0.80 | 2108.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellformula | OfficeIMO.Excel | 25.87 ms | 1.64 ms | 0.95 ms | 1.00 | 1.00 | 9557.4 KB | 1.00 | 670.3 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellformula | ClosedXML | 221.62 ms | 11.35 ms | 6.55 ms | 8.57 | 8.57 | 113853.5 KB | 11.91 | 643.2 KB | 0.96 | 756.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellformula | EPPlus | 436.26 ms | 7.13 ms | 4.12 ms | 16.86 | 16.86 | 140731.8 KB | 14.72 | 593.9 KB | 0.89 | 1586.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-empty-strings | OfficeIMO.Excel | 12.62 ms | 0.44 ms | 0.25 ms | 1.00 | 1.00 | 6731.3 KB | 1.00 | 451.4 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-empty-strings | ClosedXML | 109.53 ms | 5.12 ms | 2.96 ms | 8.68 | 8.68 | 92902.1 KB | 13.80 | 398.1 KB | 0.88 | 767.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-empty-strings | EPPlus | 162.80 ms | 1.19 ms | 0.68 ms | 12.89 | 12.89 | 74492.7 KB | 11.07 | 390.6 KB | 0.87 | 1189.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-numbers | OfficeIMO.Excel | 15.94 ms | 0.92 ms | 0.53 ms | 1.00 | 1.00 | 5805.8 KB | 1.00 | 462.6 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-numbers | ClosedXML | 104.55 ms | 4.61 ms | 2.66 ms | 6.56 | 6.56 | 84206.7 KB | 14.50 | 411.4 KB | 0.89 | 555.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-numbers | EPPlus | 202.05 ms | 31.05 ms | 17.92 ms | 12.68 | 12.68 | 86377.4 KB | 14.88 | 406.5 KB | 0.88 | 1167.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-mixed | OfficeIMO.Excel | 18.84 ms | 1.70 ms | 0.98 ms | 1.00 | 1.00 | 8009.0 KB | 1.00 | 585.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-object-mixed | ClosedXML | 159.22 ms | 7.24 ms | 4.18 ms | 8.45 | 8.45 | 111118.7 KB | 13.87 | 532.9 KB | 0.91 | 745.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-mixed | EPPlus | 274.52 ms | 29.21 ms | 16.86 ms | 14.57 | 14.57 | 113245.0 KB | 14.14 | 544.3 KB | 0.93 | 1357.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse | OfficeIMO.Excel | 18.93 ms | 0.17 ms | 0.10 ms | 1.00 | 1.00 | 7188.3 KB | 1.00 | 607.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse | ClosedXML | 145.15 ms | 4.49 ms | 2.59 ms | 7.67 | 7.67 | 105223.9 KB | 14.64 | 468.0 KB | 0.77 | 666.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse | EPPlus | 234.91 ms | 27.43 ms | 15.84 ms | 12.41 | 12.41 | 106316.9 KB | 14.79 | 494.4 KB | 0.81 | 1140.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 20.86 ms | 3.52 ms | 2.03 ms | 1.00 | 1.00 | 7188.4 KB | 1.00 | 607.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | ClosedXML | 160.37 ms | 17.19 ms | 9.93 ms | 7.69 | 7.69 | 105223.9 KB | 14.64 | 468.0 KB | 0.77 | 668.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | EPPlus | 279.38 ms | 63.52 ms | 36.67 ms | 13.39 | 13.39 | 106316.9 KB | 14.79 | 494.4 KB | 0.81 | 1239.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-scalars | OfficeIMO.Excel | 11.71 ms | 0.12 ms | 0.07 ms | 1.00 | 1.00 | 5979.5 KB | 1.00 | 441.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-scalars | ClosedXML | 96.98 ms | 2.25 ms | 1.30 ms | 8.28 | 8.28 | 82591.3 KB | 13.81 | 394.9 KB | 0.89 | 727.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-scalars | EPPlus | 236.22 ms | 6.87 ms | 3.97 ms | 20.17 | 20.17 | 85127.3 KB | 14.24 | 379.3 KB | 0.86 | 1916.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings | OfficeIMO.Excel | 18.58 ms | 2.18 ms | 1.26 ms | 1.00 | 1.00 | 15027.2 KB | 1.00 | 527.8 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-strings | ClosedXML | 141.87 ms | 17.21 ms | 9.94 ms | 7.64 | 7.64 | 104233.3 KB | 6.94 | 460.1 KB | 0.87 | 663.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings | EPPlus | 223.16 ms | 18.55 ms | 10.71 ms | 12.01 | 12.01 | 84343.7 KB | 5.61 | 444.7 KB | 0.84 | 1101.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-distinct | OfficeIMO.Excel | 14.93 ms | 1.87 ms | 1.08 ms | 1.00 | 1.00 | 13659.0 KB | 1.00 | 499.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-distinct | ClosedXML | 159.99 ms | 16.08 ms | 9.28 ms | 10.71 | 10.71 | 131501.7 KB | 9.63 | 555.3 KB | 1.11 | 971.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-distinct | EPPlus | 215.95 ms | 4.92 ms | 2.84 ms | 14.46 | 14.46 | 97729.5 KB | 7.15 | 565.1 KB | 1.13 | 1346.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-repeated | OfficeIMO.Excel | 14.38 ms | 2.29 ms | 1.32 ms | 1.00 | 1.00 | 7197.6 KB | 1.00 | 376.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-repeated | ClosedXML | 117.96 ms | 8.85 ms | 5.11 ms | 8.20 | 8.20 | 84517.3 KB | 11.74 | 331.8 KB | 0.88 | 720.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-repeated | EPPlus | 244.28 ms | 49.37 ms | 28.50 ms | 16.99 | 16.99 | 70033.2 KB | 9.73 | 300.8 KB | 0.80 | 1599.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-temporal | OfficeIMO.Excel | 20.22 ms | 0.45 ms | 0.26 ms | 1.00 | 1.00 | 7317.8 KB | 1.00 | 620.5 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-temporal | ClosedXML | 145.82 ms | 1.44 ms | 0.83 ms | 7.21 | 7.21 | 89323.7 KB | 12.21 | 483.0 KB | 0.78 | 621.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-temporal | EPPlus | 215.91 ms | 28.99 ms | 16.74 ms | 10.68 | 10.68 | 103799.9 KB | 14.18 | 495.1 KB | 0.80 | 967.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 9.36 ms | 0.29 ms | 0.17 ms | 0.83 | 1.00 | 3436.3 KB | 0.51 | 443.4 KB | 0.97 | 17.4% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 11.34 ms | 0.32 ms | 0.19 ms | 1.00 | 1.21 | 6793.3 KB | 1.00 | 455.5 KB | 1.00 | Loss +21.1% |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | ClosedXML | 118.84 ms | 3.06 ms | 1.77 ms | 10.48 | 12.70 | 96007.6 KB | 14.13 | 467.5 KB | 1.03 | 948.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | EPPlus | 193.96 ms | 6.92 ms | 4.00 ms | 17.11 | 20.72 | 87396.2 KB | 12.87 | 484.1 KB | 1.06 | 1611.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | LargeXlsx | 27.10 ms | 0.51 ms | 0.29 ms | 0.77 | 1.00 | 5606.0 KB | 0.36 | 1386.5 KB | 1.00 | 22.6% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 35.01 ms | 0.44 ms | 0.25 ms | 1.00 | 1.29 | 15708.0 KB | 1.00 | 1384.9 KB | 1.00 | Loss +29.2% |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | MiniExcel | 66.46 ms | 3.06 ms | 1.77 ms | 1.90 | 2.45 | 93246.9 KB | 5.94 | 1521.1 KB | 1.10 | 89.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | ClosedXML | 309.83 ms | 3.31 ms | 1.91 ms | 8.85 | 11.43 | 210638.1 KB | 13.41 | 1139.9 KB | 0.82 | 784.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | EPPlus | 376.69 ms | 6.00 ms | 3.46 ms | 10.76 | 13.90 | 211783.2 KB | 13.48 | 1090.0 KB | 0.79 | 975.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | Sylvan.Data.Excel | 28.40 ms | 2.30 ms | 1.33 ms | 0.76 | 1.00 | 5692.3 KB | 0.45 | 755.4 KB | 0.55 | 23.6% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | LargeXlsx | 34.40 ms | 1.02 ms | 0.59 ms | 0.92 | 1.21 | 8341.2 KB | 0.66 | 1386.5 KB | 1.00 | 7.5% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | OfficeIMO.Excel | 37.20 ms | 2.36 ms | 1.36 ms | 1.00 | 1.31 | 12673.9 KB | 1.00 | 1384.9 KB | 1.00 | Loss +31.0% |
| 25000 | package-profile | package | Package size | write-datareader-plain | MiniExcel | 76.60 ms | 7.08 ms | 4.09 ms | 2.06 | 2.70 | 92189.6 KB | 7.27 | 1521.0 KB | 1.10 | 105.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | ClosedXML | 279.07 ms | 3.39 ms | 1.96 ms | 7.50 | 9.83 | 104197.0 KB | 8.22 | 1139.9 KB | 0.82 | 650.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | EPPlus | 346.14 ms | 18.87 ms | 10.90 ms | 9.31 | 12.19 | 117370.5 KB | 9.26 | 1090.8 KB | 0.79 | 830.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table | OfficeIMO.Excel | 34.36 ms | 0.44 ms | 0.26 ms | 1.00 | 1.00 | 12691.9 KB | 1.00 | 1385.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-datareader-table | MiniExcel | 64.09 ms | 1.17 ms | 0.68 ms | 1.87 | 1.87 | 92190.0 KB | 7.26 | 1521.0 KB | 1.10 | 86.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table | EPPlus | 315.29 ms | 1.85 ms | 1.07 ms | 9.18 | 9.18 | 117370.5 KB | 9.25 | 1090.8 KB | 0.79 | 817.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table | ClosedXML | 353.26 ms | 3.88 ms | 2.24 ms | 10.28 | 10.28 | 173388.8 KB | 13.66 | 1140.7 KB | 0.82 | 928.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | OfficeIMO.Excel | 38.07 ms | 0.41 ms | 0.24 ms | 1.00 | 1.00 | 12698.2 KB | 1.00 | 1385.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | MiniExcel | 71.32 ms | 2.49 ms | 1.44 ms | 1.87 | 1.87 | 124485.4 KB | 9.80 | 1521.1 KB | 1.10 | 87.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | EPPlus | 362.42 ms | 1.26 ms | 0.73 ms | 9.52 | 9.52 | 159670.6 KB | 12.57 | 1091.0 KB | 0.79 | 851.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | ClosedXML | 769.32 ms | 11.16 ms | 6.45 ms | 20.21 | 20.21 | 566133.4 KB | 44.58 | 1140.9 KB | 0.82 | 1920.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | OfficeIMO.Excel | 31.59 ms | 1.60 ms | 0.92 ms | 1.00 | 1.00 | 9491.6 KB | 1.00 | 1329.2 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | LargeXlsx | 35.56 ms | 0.92 ms | 0.53 ms | 1.13 | 1.13 | 9257.9 KB | 0.98 | 1680.0 KB | 1.26 | 12.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | MiniExcel | 92.72 ms | 1.89 ms | 1.09 ms | 2.94 | 2.94 | 108118.7 KB | 11.39 | 1819.7 KB | 1.37 | 193.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | EPPlus | 471.40 ms | 5.43 ms | 3.13 ms | 14.92 | 14.92 | 135640.3 KB | 14.29 | 1390.4 KB | 1.05 | 1392.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | ClosedXML | 523.90 ms | 22.94 ms | 13.24 ms | 16.58 | 16.58 | 280365.8 KB | 29.54 | 1519.9 KB | 1.14 | 1558.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | OfficeIMO.Excel | 40.66 ms | 2.68 ms | 1.55 ms | 1.00 | 1.00 | 13130.4 KB | 1.00 | 1795.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | MiniExcel | 94.25 ms | 2.49 ms | 1.44 ms | 2.32 | 2.32 | 108118.7 KB | 8.23 | 1819.7 KB | 1.01 | 131.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | EPPlus | 472.69 ms | 9.79 ms | 5.65 ms | 11.62 | 11.62 | 135640.3 KB | 10.33 | 1390.4 KB | 0.77 | 1062.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | ClosedXML | 520.92 ms | 21.12 ms | 12.19 ms | 12.81 | 12.81 | 280363.3 KB | 21.35 | 1519.9 KB | 0.85 | 1181.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables | OfficeIMO.Excel | 32.65 ms | 0.32 ms | 0.19 ms | 1.00 | 1.00 | 9800.0 KB | 1.00 | 1376.4 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-dataset-tables | MiniExcel | 73.08 ms | 0.87 ms | 0.50 ms | 2.24 | 2.24 | 97074.9 KB | 9.91 | 1511.8 KB | 1.10 | 123.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables | EPPlus | 309.60 ms | 1.77 ms | 1.02 ms | 9.48 | 9.48 | 110708.7 KB | 11.30 | 1100.6 KB | 0.80 | 848.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables | ClosedXML | 348.81 ms | 8.18 ms | 4.73 ms | 10.68 | 10.68 | 171997.0 KB | 17.55 | 1139.0 KB | 0.83 | 968.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | OfficeIMO.Excel | 36.92 ms | 1.21 ms | 0.70 ms | 1.00 | 1.00 | 9812.6 KB | 1.00 | 1376.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | MiniExcel | 79.01 ms | 3.39 ms | 1.96 ms | 2.14 | 2.14 | 128864.4 KB | 13.13 | 1512.0 KB | 1.10 | 114.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | EPPlus | 383.06 ms | 5.20 ms | 3.00 ms | 10.38 | 10.38 | 195297.9 KB | 19.90 | 1100.9 KB | 0.80 | 937.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | ClosedXML | 719.33 ms | 4.54 ms | 2.62 ms | 19.48 | 19.48 | 550083.4 KB | 56.06 | 1139.3 KB | 0.83 | 1848.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-direct | LargeXlsx | 30.50 ms | 1.40 ms | 0.81 ms | 0.92 | 1.00 | 9512.4 KB | 0.77 | 1386.5 KB | 1.00 | 8.4% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-direct | OfficeIMO.Excel | 33.29 ms | 1.06 ms | 0.61 ms | 1.00 | 1.09 | 12387.3 KB | 1.00 | 1384.9 KB | 1.00 | Loss +9.2% |
| 25000 | package-profile | package | Package size | write-datatable-direct | MiniExcel | 75.13 ms | 3.09 ms | 1.78 ms | 2.26 | 2.46 | 92384.2 KB | 7.46 | 1521.0 KB | 1.10 | 125.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-direct | ClosedXML | 259.77 ms | 4.00 ms | 2.31 ms | 7.80 | 8.52 | 104197.0 KB | 8.41 | 1139.9 KB | 0.82 | 680.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-direct | EPPlus | 321.63 ms | 8.93 ms | 5.16 ms | 9.66 | 10.55 | 117370.5 KB | 9.48 | 1090.8 KB | 0.79 | 866.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | OfficeIMO.Excel | 33.76 ms | 0.49 ms | 0.28 ms | 1.00 | 1.00 | 12405.4 KB | 1.00 | 1385.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | MiniExcel | 74.23 ms | 0.96 ms | 0.55 ms | 2.20 | 2.20 | 92384.5 KB | 7.45 | 1521.1 KB | 1.10 | 119.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | EPPlus | 319.32 ms | 2.34 ms | 1.35 ms | 9.46 | 9.46 | 117370.5 KB | 9.46 | 1090.8 KB | 0.79 | 845.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | ClosedXML | 359.12 ms | 2.04 ms | 1.18 ms | 10.64 | 10.64 | 173395.3 KB | 13.98 | 1140.7 KB | 0.82 | 963.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | LargeXlsx | 37.93 ms | 3.94 ms | 2.27 ms | 0.77 | 1.00 | 5606.0 KB | 0.45 | 1386.5 KB | 1.00 | 22.9% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 49.21 ms | 15.96 ms | 9.21 ms | 1.00 | 1.30 | 12583.6 KB | 1.00 | 1384.9 KB | 1.00 | Loss +29.7% |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | MiniExcel | 97.96 ms | 4.03 ms | 2.33 ms | 1.99 | 2.58 | 93246.9 KB | 7.41 | 1521.1 KB | 1.10 | 99.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | EPPlus | 401.65 ms | 58.47 ms | 33.76 ms | 8.16 | 10.59 | 117370.5 KB | 9.33 | 1090.8 KB | 0.79 | 716.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | ClosedXML | 407.00 ms | 113.89 ms | 65.76 ms | 8.27 | 10.73 | 104197.0 KB | 8.28 | 1139.9 KB | 0.82 | 727.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 60.06 ms | 4.14 ms | 2.39 ms | 1.00 | 1.00 | 11341.1 KB | 1.00 | 1428.5 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 690.87 ms | 134.23 ms | 77.50 ms | 11.50 | 11.50 | 159742.3 KB | 14.09 | 1091.0 KB | 0.76 | 1050.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 1147.01 ms | 93.88 ms | 54.20 ms | 19.10 | 19.10 | 496956.9 KB | 43.82 | 1140.1 KB | 0.80 | 1809.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | LargeXlsx | 33.91 ms | 1.01 ms | 0.58 ms | 0.86 | 1.00 | 5614.1 KB | 0.50 | 1386.5 KB | 0.97 | 14.4% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | OfficeIMO.Excel | 39.60 ms | 3.01 ms | 1.74 ms | 1.00 | 1.17 | 11333.4 KB | 1.00 | 1428.4 KB | 1.00 | Loss +16.8% |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | MiniExcel | 82.37 ms | 5.84 ms | 3.37 ms | 2.08 | 2.43 | 93257.0 KB | 8.23 | 1521.0 KB | 1.06 | 108.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | ClosedXML | 344.89 ms | 9.09 ms | 5.25 ms | 8.71 | 10.17 | 104205.0 KB | 9.19 | 1139.9 KB | 0.80 | 770.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | EPPlus | 462.65 ms | 9.35 ms | 5.40 ms | 11.68 | 13.65 | 117437.3 KB | 10.36 | 1090.8 KB | 0.76 | 1068.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 41.82 ms | 2.56 ms | 1.48 ms | 1.00 | 1.00 | 9866.7 KB | 1.00 | 1385.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 394.55 ms | 15.81 ms | 9.13 ms | 9.43 | 9.43 | 159742.1 KB | 16.19 | 1091.0 KB | 0.79 | 843.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 717.96 ms | 30.56 ms | 17.65 ms | 17.17 | 17.17 | 496956.9 KB | 50.37 | 1140.1 KB | 0.82 | 1616.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 30.66 ms | 3.00 ms | 1.73 ms | 0.75 | 1.00 | 5614.1 KB | 0.57 | 1386.5 KB | 1.00 | 24.8% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 40.79 ms | 2.06 ms | 1.19 ms | 1.00 | 1.33 | 9859.0 KB | 1.00 | 1384.9 KB | 1.00 | Loss +33.0% |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | MiniExcel | 69.39 ms | 6.06 ms | 3.50 ms | 1.70 | 2.26 | 93257.0 KB | 9.46 | 1521.1 KB | 1.10 | 70.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | ClosedXML | 293.04 ms | 6.30 ms | 3.64 ms | 7.18 | 9.56 | 104205.0 KB | 10.57 | 1139.9 KB | 0.82 | 618.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | EPPlus | 346.45 ms | 5.41 ms | 3.13 ms | 8.49 | 11.30 | 117437.3 KB | 11.91 | 1090.8 KB | 0.79 | 749.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | LargeXlsx | 28.10 ms | 0.34 ms | 0.19 ms | 0.60 | 1.00 | 5614.1 KB | 0.36 | 1386.5 KB | 0.97 | 39.6% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | OfficeIMO.Excel | 46.52 ms | 4.96 ms | 2.86 ms | 1.00 | 1.66 | 15631.3 KB | 1.00 | 1428.4 KB | 1.00 | Loss +65.6% |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | MiniExcel | 69.20 ms | 5.26 ms | 3.03 ms | 1.49 | 2.46 | 93257.2 KB | 5.97 | 1521.1 KB | 1.06 | 48.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | ClosedXML | 328.48 ms | 37.50 ms | 21.65 ms | 7.06 | 11.69 | 104205.0 KB | 6.67 | 1139.9 KB | 0.80 | 606.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | EPPlus | 386.19 ms | 28.20 ms | 16.28 ms | 8.30 | 13.75 | 117437.3 KB | 7.51 | 1090.8 KB | 0.76 | 730.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 36.82 ms | 0.92 ms | 0.53 ms | 1.00 | 1.00 | 11340.4 KB | 1.00 | 1428.5 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 422.33 ms | 56.17 ms | 32.43 ms | 11.47 | 11.47 | 138360.3 KB | 12.20 | 1091.0 KB | 0.76 | 1047.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 479.24 ms | 7.32 ms | 4.23 ms | 13.02 | 13.02 | 275422.3 KB | 24.29 | 1140.1 KB | 0.80 | 1201.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | LargeXlsx | 49.06 ms | 8.40 ms | 4.85 ms | 0.84 | 1.00 | 6043.9 KB | 0.58 | 1816.3 KB | 0.99 | 16.0% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 58.38 ms | 10.60 ms | 6.12 ms | 1.00 | 1.19 | 10416.8 KB | 1.00 | 1828.0 KB | 1.00 | Loss +19.0% |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | MiniExcel | 108.24 ms | 9.99 ms | 5.77 ms | 1.85 | 2.21 | 113974.3 KB | 10.94 | 1936.7 KB | 1.06 | 85.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | ClosedXML | 453.78 ms | 35.58 ms | 20.54 ms | 7.77 | 9.25 | 179552.5 KB | 17.24 | 1555.2 KB | 0.85 | 677.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | EPPlus | 495.40 ms | 15.28 ms | 8.82 ms | 8.49 | 10.10 | 144919.9 KB | 13.91 | 1473.0 KB | 0.81 | 748.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | LargeXlsx | 48.20 ms | 4.57 ms | 2.64 ms | 0.92 | 1.00 | 6043.9 KB | 0.62 | 1816.3 KB | 0.99 | 7.7% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 52.24 ms | 8.47 ms | 4.89 ms | 1.00 | 1.08 | 9781.8 KB | 1.00 | 1828.0 KB | 1.00 | Loss +8.4% |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | MiniExcel | 106.60 ms | 16.72 ms | 9.65 ms | 2.04 | 2.21 | 113973.6 KB | 11.65 | 1936.7 KB | 1.06 | 104.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | ClosedXML | 468.45 ms | 77.47 ms | 44.73 ms | 8.97 | 9.72 | 179552.7 KB | 18.36 | 1555.2 KB | 0.85 | 796.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | EPPlus | 521.63 ms | 35.03 ms | 20.22 ms | 9.99 | 10.82 | 144919.9 KB | 14.82 | 1473.0 KB | 0.81 | 898.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 213.38 ms | 23.65 ms | 13.65 ms | 1.00 | 1.00 | 35984.4 KB | 1.00 | 6725.6 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | LargeXlsx | 219.17 ms | 10.21 ms | 5.89 ms | 1.03 | 1.03 | 23206.1 KB | 0.64 | 6614.8 KB | 0.98 | 2.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | MiniExcel | 334.59 ms | 5.26 ms | 3.04 ms | 1.57 | 1.57 | 347919.8 KB | 9.67 | 6949.8 KB | 1.03 | 56.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | ClosedXML | 1280.22 ms | 31.61 ms | 18.25 ms | 6.00 | 6.00 | 487444.1 KB | 13.55 | 6165.9 KB | 0.92 | 500.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | EPPlus | 1611.05 ms | 74.21 ms | 42.85 ms | 7.55 | 7.55 | 562893.4 KB | 15.64 | 5441.6 KB | 0.81 | 655.0% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | OfficeIMO.Excel | 345.37 ms | 5.96 ms | 3.44 ms | 1.00 | 1.00 | 138378.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | EPPlus | 434.91 ms | 7.94 ms | 4.58 ms | 1.26 | 1.26 | 250878.4 KB | 1.81 |  |  | 25.9% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | EPPlus 4.5.3.3 | 551.81 ms |  |  | 1.60 | 1.60 |  |  |  |  | 59.8% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | ClosedXML | 1282.08 ms | 19.86 ms | 11.47 ms | 3.71 | 3.71 | 829583.9 KB | 6.00 |  |  | 271.2% slower than OfficeIMO |
| 25000 | speed-comparison | other | DataTable table export | report-workbook-datatable | OfficeIMO.Excel | 50.01 ms | 0.86 ms | 0.49 ms | 1.00 | 1.00 | 17062.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | DataTable table export | report-workbook-datatable | EPPlus | 523.07 ms | 2.55 ms | 1.47 ms | 10.46 | 10.46 | 247752.2 KB | 14.52 |  |  | 946.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | DataTable table export | report-workbook-datatable | EPPlus 4.5.3.3 | 638.00 ms |  |  | 12.76 | 12.76 |  |  |  |  | 1175.8% slower than OfficeIMO |
| 25000 | speed-comparison | other | DataTable table export | report-workbook-datatable-core | OfficeIMO.Excel | 47.36 ms | 0.69 ms | 0.40 ms | 1.00 | 1.00 | 13549.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | DataTable table export | report-workbook-datatable-core | EPPlus | 495.55 ms | 3.00 ms | 1.73 ms | 10.46 | 10.46 | 225884.3 KB | 16.67 |  |  | 946.5% slower than OfficeIMO |
| 25000 | speed-comparison | other | DataTable table export | report-workbook-datatable-core | EPPlus 4.5.3.3 | 627.62 ms |  |  | 13.25 | 13.25 |  |  |  |  | 1225.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | DataTable table export | report-workbook-datatable-core | ClosedXML | 1019.86 ms | 12.53 ms | 7.23 ms | 21.54 | 21.54 | 832219.8 KB | 61.42 |  |  | 2053.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Object projection | build-object-datatable-dictionaries | OfficeIMO.Excel | 9.59 ms | 0.00 ms | 0.00 ms | 1.00 | 1.00 | 5164.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Object projection | build-object-datatable-typed | OfficeIMO.Excel | 7.00 ms | 0.06 ms | 0.04 ms | 1.00 | 1.00 | 8093.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Other | realworld-autofilter | OfficeIMO.Excel | 31.29 ms | 0.95 ms | 0.55 ms | 1.00 | 1.00 | 11326.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Other | realworld-autofilter | EPPlus 4.5.3.3 | 237.76 ms |  |  | 7.60 | 7.60 |  |  |  |  | 660.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Other | realworld-autofilter | ClosedXML | 290.20 ms | 7.12 ms | 4.11 ms | 9.28 | 9.28 | 210655.8 KB | 18.60 |  |  | 827.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Other | realworld-autofilter | EPPlus | 357.32 ms | 10.26 ms | 5.92 ms | 11.42 | 11.42 | 211804.2 KB | 18.70 |  |  | 1042.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Other | realworld-charts | OfficeIMO.Excel | 32.35 ms | 0.83 ms | 0.48 ms | 1.00 | 1.00 | 12323.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Other | realworld-charts | EPPlus 4.5.3.3 | 238.29 ms |  |  | 7.37 | 7.37 |  |  |  |  | 636.5% slower than OfficeIMO |
| 25000 | speed-comparison | other | Other | realworld-charts | EPPlus | 343.32 ms | 1.92 ms | 1.11 ms | 10.61 | 10.61 | 214836.4 KB | 17.43 |  |  | 961.2% slower than OfficeIMO |
| 25000 | speed-comparison | other | Other | realworld-conditional-formatting | OfficeIMO.Excel | 30.43 ms | 0.16 ms | 0.09 ms | 1.00 | 1.00 | 11391.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Other | realworld-conditional-formatting | EPPlus 4.5.3.3 | 237.67 ms |  |  | 7.81 | 7.81 |  |  |  |  | 680.9% slower than OfficeIMO |
| 25000 | speed-comparison | other | Other | realworld-conditional-formatting | ClosedXML | 303.59 ms | 17.40 ms | 10.04 ms | 9.98 | 9.98 | 210703.7 KB | 18.50 |  |  | 897.5% slower than OfficeIMO |
| 25000 | speed-comparison | other | Other | realworld-conditional-formatting | EPPlus | 346.54 ms | 9.04 ms | 5.22 ms | 11.39 | 11.39 | 211845.7 KB | 18.60 |  |  | 1038.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Other | realworld-data-validation | OfficeIMO.Excel | 30.51 ms | 0.17 ms | 0.10 ms | 1.00 | 1.00 | 11342.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Other | realworld-data-validation | EPPlus 4.5.3.3 | 246.27 ms |  |  | 8.07 | 8.07 |  |  |  |  | 707.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Other | realworld-data-validation | ClosedXML | 285.27 ms | 2.38 ms | 1.38 ms | 9.35 | 9.35 | 210664.6 KB | 18.57 |  |  | 835.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Other | realworld-data-validation | EPPlus | 341.09 ms | 3.30 ms | 1.91 ms | 11.18 | 11.18 | 211789.9 KB | 18.67 |  |  | 1018.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Other | realworld-freeze-panes | OfficeIMO.Excel | 30.76 ms | 0.86 ms | 0.50 ms | 1.00 | 1.00 | 11328.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Other | realworld-freeze-panes | EPPlus 4.5.3.3 | 239.74 ms |  |  | 7.79 | 7.79 |  |  |  |  | 679.5% slower than OfficeIMO |
| 25000 | speed-comparison | other | Other | realworld-freeze-panes | ClosedXML | 288.10 ms | 3.80 ms | 2.20 ms | 9.37 | 9.37 | 210638.7 KB | 18.59 |  |  | 836.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Other | realworld-freeze-panes | EPPlus | 350.41 ms | 14.19 ms | 8.19 ms | 11.39 | 11.39 | 211816.0 KB | 18.70 |  |  | 1039.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Other | realworld-pivot-table | OfficeIMO.Excel | 88.29 ms | 2.19 ms | 1.26 ms | 1.00 | 1.00 | 100855.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Other | realworld-pivot-table | EPPlus 4.5.3.3 | 240.67 ms |  |  | 2.73 | 2.73 |  |  |  |  | 172.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Other | realworld-pivot-table | EPPlus | 388.13 ms | 4.34 ms | 2.50 ms | 4.40 | 4.40 | 230731.4 KB | 2.29 |  |  | 339.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Other | realworld-report-all-in-one | OfficeIMO.Excel | 98.88 ms | 0.59 ms | 0.34 ms | 1.00 | 1.00 | 102332.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Other | realworld-report-all-in-one | EPPlus | 398.97 ms | 4.72 ms | 2.72 ms | 4.03 | 4.03 | 277004.6 KB | 2.71 |  |  | 303.5% slower than OfficeIMO |
| 25000 | speed-comparison | other | Other | realworld-report-all-in-one | EPPlus 4.5.3.3 | 471.99 ms |  |  | 4.77 | 4.77 |  |  |  |  | 377.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Other | realworld-report-core | OfficeIMO.Excel | 34.16 ms | 1.67 ms | 0.96 ms | 1.00 | 1.00 | 11479.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Other | realworld-report-core | EPPlus | 365.15 ms | 2.85 ms | 1.64 ms | 10.69 | 10.69 | 254994.4 KB | 22.21 |  |  | 969.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Other | realworld-report-core | EPPlus 4.5.3.3 | 456.65 ms |  |  | 13.37 | 13.37 |  |  |  |  | 1236.9% slower than OfficeIMO |
| 25000 | speed-comparison | other | Other | realworld-report-core | ClosedXML | 781.87 ms | 20.70 ms | 11.95 ms | 22.89 | 22.89 | 680108.9 KB | 59.25 |  |  | 2189.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Other | report-workbook | OfficeIMO.Excel | 47.84 ms | 0.39 ms | 0.22 ms | 1.00 | 1.00 | 14317.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Other | report-workbook | EPPlus | 540.39 ms | 7.64 ms | 4.41 ms | 11.30 | 11.30 | 364641.3 KB | 25.47 |  |  | 1029.5% slower than OfficeIMO |
| 25000 | speed-comparison | other | Other | report-workbook | EPPlus 4.5.3.3 | 634.90 ms |  |  | 13.27 | 13.27 |  |  |  |  | 1227.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Other | report-workbook-core | OfficeIMO.Excel | 44.99 ms | 0.11 ms | 0.06 ms | 1.00 | 1.00 | 10803.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Other | report-workbook-core | EPPlus | 509.77 ms | 8.09 ms | 4.67 ms | 11.33 | 11.33 | 342774.1 KB | 31.73 |  |  | 1033.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Other | report-workbook-core | EPPlus 4.5.3.3 | 628.83 ms |  |  | 13.98 | 13.98 |  |  |  |  | 1297.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Other | report-workbook-core | ClosedXML | 1061.55 ms | 19.06 ms | 11.01 ms | 23.59 | 23.59 | 975766.2 KB | 90.32 |  |  | 2259.4% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-cells | OfficeIMO.Excel | 49.58 ms | 3.94 ms | 2.27 ms | 1.00 | 1.00 | 24648.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-cells | EPPlus | 240.62 ms | 4.21 ms | 2.43 ms | 4.85 | 4.85 | 187392.6 KB | 7.60 |  |  | 385.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-cells | ClosedXML | 311.54 ms | 3.49 ms | 2.01 ms | 6.28 | 6.28 | 166510.3 KB | 6.76 |  |  | 528.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | OfficeIMO.Excel | 35.38 ms | 0.69 ms | 0.40 ms | 1.00 | 1.00 | 3959.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | EPPlus | 207.50 ms | 4.90 ms | 2.83 ms | 5.87 | 5.87 | 115541.0 KB | 29.18 |  |  | 486.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | ClosedXML | 291.84 ms | 2.17 ms | 1.25 ms | 8.25 | 8.25 | 150890.2 KB | 38.11 |  |  | 725.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-range | OfficeIMO.Excel | 48.79 ms | 1.94 ms | 1.12 ms | 1.00 | 1.00 | 24648.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-range | EPPlus | 254.69 ms | 6.13 ms | 3.54 ms | 5.22 | 5.22 | 187392.6 KB | 7.60 |  |  | 422.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-range | ClosedXML | 311.93 ms | 2.02 ms | 1.17 ms | 6.39 | 6.39 | 166516.9 KB | 6.76 |  |  | 539.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-top-range | OfficeIMO.Excel | 1.90 ms | 0.01 ms | 0.01 ms | 1.00 | 1.00 | 402.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-top-range | EPPlus | 202.01 ms | 6.55 ms | 3.78 ms | 106.08 | 106.08 | 105579.4 KB | 262.16 |  |  | 10508.2% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-top-range | ClosedXML | 283.85 ms | 8.22 ms | 4.75 ms | 149.05 | 149.05 | 149397.6 KB | 370.96 |  |  | 14805.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | OfficeIMO.Excel | 21.56 ms | 1.90 ms | 1.09 ms | 1.00 | 1.00 | 6287.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | EPPlus 4.5.3.3 | 72.88 ms |  |  | 3.38 | 3.38 |  |  |  |  | 238.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | EPPlus | 144.77 ms | 1.90 ms | 1.10 ms | 6.72 | 6.72 | 70813.8 KB | 11.26 |  |  | 571.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | ClosedXML | 171.49 ms | 15.27 ms | 8.82 ms | 7.95 | 7.95 | 79507.6 KB | 12.65 |  |  | 695.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | Sylvan.Data.Excel | 0.96 ms | 0.01 ms | 0.00 ms | 0.60 | 1.00 | 316.6 KB | 1.27 |  |  | 40.4% faster than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | ExcelDataReader | 1.39 ms | 0.02 ms | 0.01 ms | 0.86 | 1.44 | 4046.1 KB | 16.26 |  |  | 14.0% faster than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | OfficeIMO.Excel | 1.61 ms | 0.02 ms | 0.01 ms | 1.00 | 1.68 | 248.8 KB | 1.00 |  |  | Loss +67.8% |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | ClosedXML | 3.24 ms | 0.07 ms | 0.04 ms | 2.01 | 3.37 | 4392.9 KB | 17.65 |  |  | 101.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | MiniExcel | 10.93 ms | 0.57 ms | 0.33 ms | 6.77 | 11.37 | 46189.1 KB | 185.63 |  |  | 577.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | EPPlus 4.5.3.3 | 16.81 ms |  |  | 10.42 | 17.48 |  |  |  |  | 941.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | EPPlus | 89.83 ms | 1.89 ms | 1.09 ms | 55.66 | 93.40 | 43070.2 KB | 173.09 |  |  | 5466.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | Sylvan.Data.Excel | 0.96 ms | 0.01 ms | 0.01 ms | 0.57 | 1.00 | 316.6 KB | 1.27 |  |  | 42.6% faster than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | ExcelDataReader | 1.38 ms | 0.02 ms | 0.01 ms | 0.83 | 1.44 | 4046.1 KB | 16.26 |  |  | 17.3% faster than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | OfficeIMO.Excel | 1.67 ms | 0.08 ms | 0.04 ms | 1.00 | 1.74 | 248.9 KB | 1.00 |  |  | Loss +74.3% |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | ClosedXML | 3.22 ms | 0.20 ms | 0.12 ms | 1.93 | 3.36 | 4392.9 KB | 17.65 |  |  | 92.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | MiniExcel | 11.22 ms | 0.11 ms | 0.06 ms | 6.72 | 11.72 | 46189.1 KB | 185.57 |  |  | 572.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | EPPlus 4.5.3.3 | 16.58 ms |  |  | 9.93 | 17.31 |  |  |  |  | 893.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | EPPlus | 90.18 ms | 2.07 ms | 1.20 ms | 54.01 | 94.15 | 43070.2 KB | 173.04 |  |  | 5300.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | Sylvan.Data.Excel | 17.58 ms | 0.47 ms | 0.27 ms | 0.85 | 1.00 | 1936.7 KB | 0.21 |  |  | 15.2% faster than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | OfficeIMO.Excel | 20.74 ms | 1.19 ms | 0.69 ms | 1.00 | 1.18 | 9295.0 KB | 1.00 |  |  | Loss +18.0% |
| 25000 | speed-comparison | read | Other | shared-string-read | ExcelDataReader | 44.01 ms | 0.74 ms | 0.43 ms | 2.12 | 2.50 | 25004.8 KB | 2.69 |  |  | 112.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | MiniExcel | 49.00 ms | 0.43 ms | 0.25 ms | 2.36 | 2.79 | 74398.5 KB | 8.00 |  |  | 136.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | EPPlus 4.5.3.3 | 83.68 ms |  |  | 4.04 | 4.76 |  |  |  |  | 303.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | EPPlus | 145.48 ms | 2.05 ms | 1.18 ms | 7.02 | 8.28 | 89345.5 KB | 9.61 |  |  | 601.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | ClosedXML | 147.84 ms | 5.88 ms | 3.40 ms | 7.13 | 8.41 | 90414.2 KB | 9.73 |  |  | 612.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | OfficeIMO.Excel | 34.26 ms | 1.02 ms | 0.59 ms | 1.00 | 1.00 | 1242.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | Sylvan.Data.Excel | 42.52 ms | 3.43 ms | 1.98 ms | 1.24 | 1.24 | 3534.8 KB | 2.85 |  |  | 24.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | ExcelDataReader | 104.28 ms | 3.21 ms | 1.85 ms | 3.04 | 3.04 | 61193.9 KB | 49.25 |  |  | 204.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | MiniExcel | 121.68 ms | 4.91 ms | 2.83 ms | 3.55 | 3.55 | 186406.3 KB | 150.03 |  |  | 255.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | EPPlus | 204.67 ms | 7.43 ms | 4.29 ms | 5.97 | 5.97 | 105608.3 KB | 85.00 |  |  | 497.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | ClosedXML | 285.77 ms | 1.03 ms | 0.59 ms | 8.34 | 8.34 | 149384.2 KB | 120.24 |  |  | 734.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | OfficeIMO.Excel | 56.83 ms | 1.98 ms | 1.14 ms | 1.00 | 1.00 | 34766.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Range and table read | read-datatable | Sylvan.Data.Excel | 58.57 ms | 1.69 ms | 0.98 ms | 1.03 | 1.03 | 18394.2 KB | 0.53 |  |  | 3.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | ExcelDataReader | 123.90 ms | 1.63 ms | 0.94 ms | 2.18 | 2.18 | 76053.3 KB | 2.19 |  |  | 118.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | MiniExcel | 141.87 ms | 3.03 ms | 1.75 ms | 2.50 | 2.50 | 181273.2 KB | 5.21 |  |  | 149.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | EPPlus 4.5.3.3 | 202.05 ms |  |  | 3.56 | 3.56 |  |  |  |  | 255.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | EPPlus | 252.00 ms | 1.38 ms | 0.80 ms | 4.43 | 4.43 | 202249.6 KB | 5.82 |  |  | 343.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | ClosedXML | 316.21 ms | 4.67 ms | 2.70 ms | 5.56 | 5.56 | 178452.0 KB | 5.13 |  |  | 456.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | OfficeIMO.Excel | 36.37 ms | 3.14 ms | 1.81 ms | 1.00 | 1.00 | 4154.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | Sylvan.Data.Excel | 46.00 ms | 3.05 ms | 1.76 ms | 1.26 | 1.26 | 4316.2 KB | 1.04 |  |  | 26.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | MiniExcel | 95.20 ms | 5.53 ms | 3.19 ms | 2.62 | 2.62 | 158604.8 KB | 38.18 |  |  | 161.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | ExcelDataReader | 104.41 ms | 1.39 ms | 0.80 ms | 2.87 | 2.87 | 61193.9 KB | 14.73 |  |  | 187.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | EPPlus | 211.90 ms | 7.15 ms | 4.13 ms | 5.83 | 5.83 | 115541.0 KB | 27.81 |  |  | 482.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | ClosedXML | 304.83 ms | 17.47 ms | 10.08 ms | 8.38 | 8.38 | 150891.1 KB | 36.32 |  |  | 738.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | Sylvan.Data.Excel | 51.30 ms | 3.48 ms | 2.01 ms | 0.96 | 1.00 | 3534.8 KB | 0.13 |  |  | 3.6% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | OfficeIMO.Excel | 53.24 ms | 3.18 ms | 1.84 ms | 1.00 | 1.04 | 26218.6 KB | 1.00 |  |  | Loss +3.8% |
| 25000 | speed-comparison | read | Range and table read | read-range | ExcelDataReader | 129.90 ms | 7.22 ms | 4.17 ms | 2.44 | 2.53 | 61193.9 KB | 2.33 |  |  | 144.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | MiniExcel | 130.41 ms | 5.35 ms | 3.09 ms | 2.45 | 2.54 | 186406.4 KB | 7.11 |  |  | 145.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | EPPlus 4.5.3.3 | 178.03 ms |  |  | 3.34 | 3.47 |  |  |  |  | 234.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | EPPlus | 260.23 ms | 11.78 ms | 6.80 ms | 4.89 | 5.07 | 187390.2 KB | 7.15 |  |  | 388.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | ClosedXML | 347.41 ms | 14.99 ms | 8.65 ms | 6.53 | 6.77 | 163590.3 KB | 6.24 |  |  | 552.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | OfficeIMO.Excel | 53.65 ms | 3.69 ms | 2.13 ms | 1.00 | 1.00 | 26804.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | Sylvan.Data.Excel | 55.13 ms | 3.74 ms | 2.16 ms | 1.03 | 1.03 | 4484.9 KB | 0.17 |  |  | 2.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | ExcelDataReader | 111.19 ms | 5.82 ms | 3.36 ms | 2.07 | 2.07 | 61193.9 KB | 2.28 |  |  | 107.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | MiniExcel | 127.15 ms | 4.81 ms | 2.78 ms | 2.37 | 2.37 | 186406.4 KB | 6.95 |  |  | 137.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | EPPlus | 246.53 ms | 6.44 ms | 3.72 ms | 4.60 | 4.60 | 187390.2 KB | 6.99 |  |  | 359.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | ClosedXML | 346.89 ms | 32.62 ms | 18.84 ms | 6.47 | 6.47 | 163587.3 KB | 6.10 |  |  | 546.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | Sylvan.Data.Excel | 0.45 ms | 0.05 ms | 0.03 ms | 0.24 | 1.00 | 348.5 KB | 0.84 |  |  | 76.2% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | MiniExcel | 0.71 ms | 0.01 ms | 0.01 ms | 0.38 | 1.58 | 858.3 KB | 2.06 |  |  | 62.5% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | OfficeIMO.Excel | 1.90 ms | 0.13 ms | 0.08 ms | 1.00 | 4.20 | 416.1 KB | 1.00 |  |  | Loss +319.8% |
| 25000 | speed-comparison | read | Range and table read | read-top-range | ExcelDataReader | 40.42 ms | 1.95 ms | 1.13 ms | 21.25 | 89.23 | 17107.2 KB | 41.11 |  |  | 2025.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | EPPlus 4.5.3.3 | 129.14 ms |  |  | 67.91 | 285.11 |  |  |  |  | 6691.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | EPPlus | 205.92 ms | 4.03 ms | 2.33 ms | 108.28 | 454.60 | 105577.1 KB | 253.73 |  |  | 10728.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | ClosedXML | 292.07 ms | 11.52 ms | 6.65 ms | 153.58 | 644.80 | 149382.8 KB | 359.01 |  |  | 15258.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | Sylvan.Data.Excel | 48.19 ms | 1.24 ms | 0.72 ms | 0.56 | 1.00 | 3534.8 KB | 0.10 |  |  | 44.4% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | OfficeIMO.Excel | 86.61 ms | 1.22 ms | 0.71 ms | 1.00 | 1.80 | 34214.6 KB | 1.00 |  |  | Loss +79.7% |
| 25000 | speed-comparison | read | Range and table read | read-used-range | ExcelDataReader | 115.29 ms | 5.81 ms | 3.36 ms | 1.33 | 2.39 | 61193.9 KB | 1.79 |  |  | 33.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | MiniExcel | 126.48 ms | 4.98 ms | 2.88 ms | 1.46 | 2.62 | 186406.4 KB | 5.45 |  |  | 46.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | EPPlus | 249.07 ms | 4.50 ms | 2.60 ms | 2.88 | 5.17 | 187390.2 KB | 5.48 |  |  | 187.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | ClosedXML | 321.54 ms | 11.01 ms | 6.35 ms | 3.71 | 6.67 | 163589.5 KB | 4.78 |  |  | 271.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | OfficeIMO.Excel | 36.68 ms | 3.77 ms | 2.18 ms | 1.00 | 1.00 | 1245.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | Sylvan.Data.Excel | 43.76 ms | 3.18 ms | 1.84 ms | 1.19 | 1.19 | 3534.8 KB | 2.84 |  |  | 19.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | ExcelDataReader | 107.47 ms | 2.79 ms | 1.61 ms | 2.93 | 2.93 | 61193.9 KB | 49.12 |  |  | 193.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | MiniExcel | 116.86 ms | 0.47 ms | 0.27 ms | 3.19 | 3.19 | 186406.3 KB | 149.63 |  |  | 218.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | EPPlus | 204.02 ms | 4.24 ms | 2.45 ms | 5.56 | 5.56 | 105608.3 KB | 84.77 |  |  | 456.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | ClosedXML | 285.75 ms | 8.46 ms | 4.89 ms | 7.79 | 7.79 | 149389.0 KB | 119.91 |  |  | 679.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | OfficeIMO.Excel | 45.91 ms | 0.59 ms | 0.34 ms | 1.00 | 1.00 | 27005.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | Sylvan.Data.Excel | 47.62 ms | 0.22 ms | 0.13 ms | 1.04 | 1.04 | 3534.8 KB | 0.13 |  |  | 3.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | ExcelDataReader | 106.17 ms | 2.41 ms | 1.39 ms | 2.31 | 2.31 | 61193.9 KB | 2.27 |  |  | 131.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | MiniExcel | 124.97 ms | 6.83 ms | 3.94 ms | 2.72 | 2.72 | 186406.4 KB | 6.90 |  |  | 172.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | EPPlus 4.5.3.3 | 166.86 ms |  |  | 3.63 | 3.63 |  |  |  |  | 263.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | EPPlus | 240.42 ms | 2.44 ms | 1.41 ms | 5.24 | 5.24 | 187390.2 KB | 6.94 |  |  | 423.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | ClosedXML | 304.41 ms | 1.05 ms | 0.60 ms | 6.63 | 6.63 | 163592.4 KB | 6.06 |  |  | 563.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | Sylvan.Data.Excel | 0.44 ms | 0.01 ms | 0.01 ms | 0.24 | 1.00 | 348.5 KB | 0.83 |  |  | 75.8% faster than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | MiniExcel | 0.77 ms | 0.12 ms | 0.07 ms | 0.42 | 1.75 | 858.3 KB | 2.05 |  |  | 57.6% faster than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | OfficeIMO.Excel | 1.81 ms | 0.04 ms | 0.03 ms | 1.00 | 4.13 | 419.5 KB | 1.00 |  |  | Loss +313.2% |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | ExcelDataReader | 38.28 ms | 0.71 ms | 0.41 ms | 21.18 | 87.49 | 17107.2 KB | 40.78 |  |  | 2017.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | EPPlus 4.5.3.3 | 134.20 ms |  |  | 74.24 | 306.74 |  |  |  |  | 7324.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | EPPlus | 217.44 ms | 14.00 ms | 8.08 ms | 120.29 | 496.99 | 105577.1 KB | 251.68 |  |  | 11928.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | ClosedXML | 292.44 ms | 9.28 ms | 5.36 ms | 161.78 | 668.43 | 149389.5 KB | 356.12 |  |  | 16078.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | Sylvan.Data.Excel | 0.47 ms | 0.06 ms | 0.03 ms | 0.25 | 1.00 | 348.5 KB | 0.83 |  |  | 75.3% faster than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | MiniExcel | 0.71 ms | 0.04 ms | 0.03 ms | 0.37 | 1.51 | 858.3 KB | 2.04 |  |  | 62.8% faster than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | OfficeIMO.Excel | 1.91 ms | 0.19 ms | 0.11 ms | 1.00 | 4.05 | 420.2 KB | 1.00 |  |  | Loss +305.3% |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | ExcelDataReader | 37.01 ms | 0.30 ms | 0.17 ms | 19.41 | 78.69 | 17107.2 KB | 40.71 |  |  | 1841.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | EPPlus | 200.20 ms | 2.40 ms | 1.39 ms | 105.03 | 425.71 | 105577.1 KB | 251.22 |  |  | 10402.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | ClosedXML | 281.38 ms | 9.79 ms | 5.65 ms | 147.62 | 598.34 | 149384.3 KB | 355.47 |  |  | 14661.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | Sylvan.Data.Excel | 47.00 ms | 1.44 ms | 0.83 ms | 0.95 | 1.00 | 5805.0 KB | 0.25 |  |  | 5.0% faster than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | OfficeIMO.Excel | 49.48 ms | 2.09 ms | 1.21 ms | 1.00 | 1.05 | 23682.5 KB | 1.00 |  |  | Loss +5.3% |
| 25000 | speed-comparison | read | Typed object read | read-objects | ExcelDataReader | 113.78 ms | 2.69 ms | 1.55 ms | 2.30 | 2.42 | 63464.0 KB | 2.68 |  |  | 130.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | MiniExcel | 140.40 ms | 6.12 ms | 3.53 ms | 2.84 | 2.99 | 183645.3 KB | 7.75 |  |  | 183.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | EPPlus 4.5.3.3 | 160.99 ms |  |  | 3.25 | 3.43 |  |  |  |  | 225.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | EPPlus | 246.55 ms | 5.23 ms | 3.02 ms | 4.98 | 5.25 | 199607.5 KB | 8.43 |  |  | 398.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | ClosedXML | 302.83 ms | 7.88 ms | 4.55 ms | 6.12 | 6.44 | 165538.0 KB | 6.99 |  |  | 512.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | Sylvan.Data.Excel | 45.79 ms | 0.26 ms | 0.15 ms | 0.98 | 1.00 | 5292.6 KB | 0.22 |  |  | 2.0% faster than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | OfficeIMO.Excel | 46.73 ms | 0.58 ms | 0.34 ms | 1.00 | 1.02 | 23682.8 KB | 1.00 |  |  | Loss +2.1% |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | ExcelDataReader | 102.85 ms | 0.98 ms | 0.57 ms | 2.20 | 2.25 | 62951.7 KB | 2.66 |  |  | 120.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | MiniExcel | 121.23 ms | 1.60 ms | 0.92 ms | 2.59 | 2.65 | 183133.0 KB | 7.73 |  |  | 159.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | EPPlus 4.5.3.3 | 161.62 ms |  |  | 3.46 | 3.53 |  |  |  |  | 245.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | EPPlus | 239.12 ms | 1.97 ms | 1.14 ms | 5.12 | 5.22 | 199412.2 KB | 8.42 |  |  | 411.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | ClosedXML | 296.22 ms | 4.07 ms | 2.35 ms | 6.34 | 6.47 | 165345.7 KB | 6.98 |  |  | 533.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | OfficeIMO.Excel | 38.06 ms | 0.93 ms | 0.53 ms | 1.00 | 1.00 | 12698.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | MiniExcel | 70.60 ms | 3.81 ms | 2.20 ms | 1.85 | 1.85 | 124485.4 KB | 9.80 |  |  | 85.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | EPPlus | 380.41 ms | 33.25 ms | 19.20 ms | 10.00 | 10.00 | 159670.6 KB | 12.57 |  |  | 899.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | EPPlus 4.5.3.3 | 444.55 ms |  |  | 11.68 | 11.68 |  |  |  |  | 1068.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | ClosedXML | 808.69 ms | 45.19 ms | 26.09 ms | 21.25 | 21.25 | 566135.6 KB | 44.58 |  |  | 2024.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | OfficeIMO.Excel | 41.98 ms | 4.81 ms | 2.77 ms | 1.00 | 1.00 | 9812.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | MiniExcel | 85.39 ms | 3.67 ms | 2.12 ms | 2.03 | 2.03 | 128864.4 KB | 13.13 |  |  | 103.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | EPPlus | 407.26 ms | 23.33 ms | 13.47 ms | 9.70 | 9.70 | 195297.9 KB | 19.90 |  |  | 870.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | ClosedXML | 788.34 ms | 49.07 ms | 28.33 ms | 18.78 | 18.78 | 550087.1 KB | 56.06 |  |  | 1778.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 36.60 ms | 1.15 ms | 0.66 ms | 1.00 | 1.00 | 11333.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 371.98 ms | 8.98 ms | 5.18 ms | 10.16 | 10.16 | 159670.6 KB | 14.09 |  |  | 916.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 711.64 ms | 14.31 ms | 8.26 ms | 19.44 | 19.44 | 496948.9 KB | 43.85 |  |  | 1844.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 41.65 ms | 0.76 ms | 0.44 ms | 1.00 | 1.00 | 9858.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 368.05 ms | 3.44 ms | 1.98 ms | 8.84 | 8.84 | 159670.6 KB | 16.20 |  |  | 783.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 694.81 ms | 12.92 ms | 7.46 ms | 16.68 | 16.68 | 496948.9 KB | 50.41 |  |  | 1568.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 33.99 ms | 1.45 ms | 0.84 ms | 1.00 | 1.00 | 11332.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 358.72 ms | 11.29 ms | 6.52 ms | 10.55 | 10.55 | 138290.0 KB | 12.20 |  |  | 955.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 438.00 ms | 11.06 ms | 6.38 ms | 12.89 | 12.89 | 275414.3 KB | 24.30 |  |  | 1188.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | OfficeIMO.Excel | 12.92 ms | 0.95 ms | 0.55 ms | 1.00 | 1.00 | 6723.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | ClosedXML | 108.12 ms | 3.46 ms | 2.00 ms | 8.37 | 8.37 | 92894.1 KB | 13.82 |  |  | 737.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | EPPlus | 163.97 ms | 4.81 ms | 2.78 ms | 12.69 | 12.69 | 74425.8 KB | 11.07 |  |  | 1169.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | OfficeIMO.Excel | 16.92 ms | 1.74 ms | 1.00 ms | 1.00 | 1.00 | 5797.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | EPPlus 4.5.3.3 | 100.49 ms |  |  | 5.94 | 5.94 |  |  |  |  | 493.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | ClosedXML | 109.60 ms | 6.22 ms | 3.59 ms | 6.48 | 6.48 | 84198.7 KB | 14.52 |  |  | 547.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | EPPlus | 186.60 ms | 9.37 ms | 5.41 ms | 11.03 | 11.03 | 86279.5 KB | 14.88 |  |  | 1002.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | OfficeIMO.Excel | 20.04 ms | 2.80 ms | 1.62 ms | 1.00 | 1.00 | 8001.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | EPPlus 4.5.3.3 | 115.97 ms |  |  | 5.79 | 5.79 |  |  |  |  | 478.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | ClosedXML | 175.43 ms | 7.68 ms | 4.43 ms | 8.75 | 8.75 | 111110.6 KB | 13.89 |  |  | 775.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | EPPlus | 225.75 ms | 8.50 ms | 4.91 ms | 11.26 | 11.26 | 113162.8 KB | 14.14 |  |  | 1026.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | OfficeIMO.Excel | 20.10 ms | 0.38 ms | 0.22 ms | 1.00 | 1.00 | 7180.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | ClosedXML | 149.88 ms | 4.66 ms | 2.69 ms | 7.46 | 7.46 | 105215.9 KB | 14.65 |  |  | 645.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | EPPlus | 209.57 ms | 2.76 ms | 1.59 ms | 10.43 | 10.43 | 106250.4 KB | 14.80 |  |  | 942.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 16.76 ms | 0.62 ms | 0.36 ms | 1.00 | 1.00 | 7180.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | ClosedXML | 146.78 ms | 4.36 ms | 2.52 ms | 8.76 | 8.76 | 105215.9 KB | 14.65 |  |  | 776.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | EPPlus | 212.57 ms | 8.63 ms | 4.98 ms | 12.69 | 12.69 | 106250.4 KB | 14.80 |  |  | 1168.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | OfficeIMO.Excel | 11.89 ms | 0.33 ms | 0.19 ms | 1.00 | 1.00 | 5971.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | EPPlus 4.5.3.3 | 96.74 ms |  |  | 8.14 | 8.14 |  |  |  |  | 713.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | ClosedXML | 102.20 ms | 2.51 ms | 1.45 ms | 8.60 | 8.60 | 82583.3 KB | 13.83 |  |  | 759.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | EPPlus | 185.08 ms | 6.04 ms | 3.49 ms | 15.57 | 15.57 | 85057.4 KB | 14.24 |  |  | 1457.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | OfficeIMO.Excel | 20.70 ms | 0.26 ms | 0.15 ms | 1.00 | 1.00 | 7309.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | EPPlus 4.5.3.3 | 102.09 ms |  |  | 4.93 | 4.93 |  |  |  |  | 393.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | ClosedXML | 144.39 ms | 2.42 ms | 1.40 ms | 6.97 | 6.97 | 89315.7 KB | 12.22 |  |  | 597.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | EPPlus | 199.40 ms | 6.78 ms | 3.92 ms | 9.63 | 9.63 | 103733.9 KB | 14.19 |  |  | 863.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | OfficeIMO.Excel | 36.09 ms | 2.18 ms | 1.26 ms | 1.00 | 1.00 | 12551.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | MiniExcel | 88.55 ms | 12.54 ms | 7.24 ms | 2.45 | 2.45 | 97077.8 KB | 7.73 |  |  | 145.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | ClosedXML | 390.36 ms | 17.92 ms | 10.34 ms | 10.82 | 10.82 | 172008.1 KB | 13.70 |  |  | 981.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | EPPlus | 436.16 ms | 22.20 ms | 12.82 ms | 12.09 | 12.09 | 111170.8 KB | 8.86 |  |  | 1108.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | OfficeIMO.Excel | 45.39 ms | 3.96 ms | 2.29 ms | 1.00 | 1.00 | 13130.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | MiniExcel | 128.26 ms | 21.40 ms | 12.36 ms | 2.83 | 2.83 | 108118.7 KB | 8.23 |  |  | 182.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | EPPlus | 547.82 ms | 54.78 ms | 31.63 ms | 12.07 | 12.07 | 135640.3 KB | 10.33 |  |  | 1107.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | ClosedXML | 587.97 ms | 3.38 ms | 1.95 ms | 12.95 | 12.95 | 280364.4 KB | 21.35 |  |  | 1195.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | OfficeIMO.Excel | 41.49 ms | 3.55 ms | 2.05 ms | 1.00 | 1.00 | 9800.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | MiniExcel | 95.54 ms | 4.64 ms | 2.68 ms | 2.30 | 2.30 | 97074.9 KB | 9.91 |  |  | 130.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | EPPlus 4.5.3.3 | 213.69 ms |  |  | 5.15 | 5.15 |  |  |  |  | 415.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | EPPlus | 359.86 ms | 16.70 ms | 9.64 ms | 8.67 | 8.67 | 110708.7 KB | 11.30 |  |  | 767.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | ClosedXML | 421.48 ms | 31.83 ms | 18.38 ms | 10.16 | 10.16 | 171989.8 KB | 17.55 |  |  | 915.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | OfficeIMO.Excel | 36.75 ms | 1.23 ms | 0.71 ms | 1.00 | 1.00 | 12691.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | MiniExcel | 69.47 ms | 1.97 ms | 1.14 ms | 1.89 | 1.89 | 92190.0 KB | 7.26 |  |  | 89.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | EPPlus 4.5.3.3 | 222.38 ms |  |  | 6.05 | 6.05 |  |  |  |  | 505.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | EPPlus | 329.10 ms | 9.98 ms | 5.76 ms | 8.95 | 8.95 | 117370.5 KB | 9.25 |  |  | 795.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | ClosedXML | 375.34 ms | 9.19 ms | 5.31 ms | 10.21 | 10.21 | 173392.3 KB | 13.66 |  |  | 921.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | LargeXlsx | 37.99 ms | 5.33 ms | 3.08 ms | 0.86 | 1.00 | 9512.4 KB | 0.77 |  |  | 13.6% faster than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | OfficeIMO.Excel | 43.95 ms | 13.61 ms | 7.86 ms | 1.00 | 1.16 | 12387.3 KB | 1.00 |  |  | Loss +15.7% |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | MiniExcel | 81.01 ms | 1.53 ms | 0.88 ms | 1.84 | 2.13 | 92384.2 KB | 7.46 |  |  | 84.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | EPPlus 4.5.3.3 | 217.29 ms |  |  | 4.94 | 5.72 |  |  |  |  | 394.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | ClosedXML | 294.38 ms | 15.43 ms | 8.91 ms | 6.70 | 7.75 | 104197.0 KB | 8.41 |  |  | 569.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | EPPlus | 354.88 ms | 15.85 ms | 9.15 ms | 8.07 | 9.34 | 117370.5 KB | 9.48 |  |  | 707.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | OfficeIMO.Excel | 35.92 ms | 0.97 ms | 0.56 ms | 1.00 | 1.00 | 9673.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | MiniExcel | 78.36 ms | 0.84 ms | 0.48 ms | 2.18 | 2.18 | 89653.1 KB | 9.27 |  |  | 118.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | EPPlus | 322.71 ms | 7.90 ms | 4.56 ms | 8.98 | 8.98 | 114680.7 KB | 11.85 |  |  | 798.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | ClosedXML | 370.25 ms | 8.30 ms | 4.79 ms | 10.31 | 10.31 | 170658.5 KB | 17.64 |  |  | 930.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | OfficeIMO.Excel | 40.36 ms | 4.26 ms | 2.46 ms | 1.00 | 1.00 | 12410.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | MiniExcel | 111.29 ms | 18.42 ms | 10.64 ms | 2.76 | 2.76 | 92390.5 KB | 7.44 |  |  | 175.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | EPPlus 4.5.3.3 | 223.64 ms |  |  | 5.54 | 5.54 |  |  |  |  | 454.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | EPPlus | 407.39 ms | 100.68 ms | 58.13 ms | 10.09 | 10.09 | 117415.0 KB | 9.46 |  |  | 909.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | ClosedXML | 413.73 ms | 19.89 ms | 11.49 ms | 10.25 | 10.25 | 173392.1 KB | 13.97 |  |  | 925.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | OfficeIMO.Excel | 43.86 ms | 3.72 ms | 2.15 ms | 1.00 | 1.00 | 11393.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | MiniExcel | 75.86 ms | 2.93 ms | 1.69 ms | 1.73 | 1.73 | 125541.4 KB | 11.02 |  |  | 73.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | EPPlus | 451.91 ms | 11.43 ms | 6.60 ms | 10.30 | 10.30 | 254887.6 KB | 22.37 |  |  | 930.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | EPPlus 4.5.3.3 | 519.39 ms |  |  | 11.84 | 11.84 |  |  |  |  | 1084.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | ClosedXML | 966.22 ms | 40.17 ms | 23.19 ms | 22.03 | 22.03 | 565942.2 KB | 49.67 |  |  | 2103.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | OfficeIMO.Excel | 21.20 ms | 1.29 ms | 0.74 ms | 1.00 | 1.00 | 9548.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | EPPlus 4.5.3.3 | 131.18 ms |  |  | 6.19 | 6.19 |  |  |  |  | 518.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | ClosedXML | 179.99 ms | 15.13 ms | 8.74 ms | 8.49 | 8.49 | 113844.9 KB | 11.92 |  |  | 748.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | EPPlus | 323.42 ms | 18.56 ms | 10.72 ms | 15.25 | 15.25 | 140665.9 KB | 14.73 |  |  | 1425.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-dictionary-objects-table-direct | OfficeIMO.Excel | 46.09 ms | 0.74 ms | 0.43 ms | 1.00 | 1.00 | 14835.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | LargeXlsx | 37.45 ms | 0.51 ms | 0.30 ms | 0.86 | 1.00 | 6035.9 KB | 0.58 |  |  | 14.4% faster than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 43.73 ms | 1.71 ms | 0.99 ms | 1.00 | 1.17 | 10408.8 KB | 1.00 |  |  | Loss +16.8% |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | MiniExcel | 79.83 ms | 1.90 ms | 1.10 ms | 1.83 | 2.13 | 113964.2 KB | 10.95 |  |  | 82.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | ClosedXML | 369.08 ms | 4.46 ms | 2.57 ms | 8.44 | 9.86 | 179544.5 KB | 17.25 |  |  | 744.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | EPPlus | 440.36 ms | 4.21 ms | 2.43 ms | 10.07 | 11.76 | 144853.2 KB | 13.92 |  |  | 907.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | LargeXlsx | 39.21 ms | 0.22 ms | 0.13 ms | 0.90 | 1.00 | 6035.9 KB | 0.62 |  |  | 10.1% faster than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 43.61 ms | 0.84 ms | 0.49 ms | 1.00 | 1.11 | 9773.8 KB | 1.00 |  |  | Loss +11.2% |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | MiniExcel | 80.87 ms | 4.07 ms | 2.35 ms | 1.85 | 2.06 | 113964.2 KB | 11.66 |  |  | 85.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | ClosedXML | 379.99 ms | 6.45 ms | 3.72 ms | 8.71 | 9.69 | 179544.5 KB | 18.37 |  |  | 771.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | EPPlus | 444.45 ms | 5.74 ms | 3.32 ms | 10.19 | 11.34 | 144853.2 KB | 14.82 |  |  | 919.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 188.84 ms | 3.24 ms | 1.87 ms | 1.00 | 1.00 | 35981.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | LargeXlsx | 202.63 ms | 2.41 ms | 1.39 ms | 1.07 | 1.07 | 23203.4 KB | 0.64 |  |  | 7.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | MiniExcel | 330.00 ms | 3.53 ms | 2.04 ms | 1.75 | 1.75 | 347916.6 KB | 9.67 |  |  | 74.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | ClosedXML | 1211.33 ms | 7.00 ms | 4.04 ms | 6.41 | 6.41 | 487439.1 KB | 13.55 |  |  | 541.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | EPPlus | 1536.70 ms | 16.38 ms | 9.46 ms | 8.14 | 8.14 | 562848.8 KB | 15.64 |  |  | 713.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | LargeXlsx | 11.90 ms | 0.67 ms | 0.39 ms | 0.66 | 1.00 | 2763.0 KB | 0.24 |  |  | 33.9% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | OfficeIMO.Excel | 18.02 ms | 1.46 ms | 0.84 ms | 1.00 | 1.51 | 11671.6 KB | 1.00 |  |  | Loss +51.4% |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | MiniExcel | 33.10 ms | 2.75 ms | 1.59 ms | 1.84 | 2.78 | 58233.0 KB | 4.99 |  |  | 83.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | EPPlus 4.5.3.3 | 110.08 ms |  |  | 6.11 | 9.25 |  |  |  |  | 510.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | ClosedXML | 138.31 ms | 0.78 ms | 0.45 ms | 7.67 | 11.62 | 104225.1 KB | 8.93 |  |  | 667.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | EPPlus | 215.93 ms | 7.04 ms | 4.06 ms | 11.98 | 18.14 | 100275.4 KB | 8.59 |  |  | 1098.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 15.73 ms | 6.47 ms | 3.74 ms | 1.00 | 1.00 | 6801.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 16.19 ms | 6.11 ms | 3.53 ms | 1.03 | 1.03 | 3444.4 KB | 0.51 |  |  | 2.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | ClosedXML | 202.46 ms | 55.72 ms | 32.17 ms | 12.87 | 12.87 | 96015.7 KB | 14.12 |  |  | 1186.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | EPPlus | 258.46 ms | 27.21 ms | 15.71 ms | 16.43 | 16.43 | 87467.0 KB | 12.86 |  |  | 1542.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | LargeXlsx | 30.47 ms | 2.77 ms | 1.60 ms | 0.73 | 1.00 | 5606.0 KB | 0.36 |  |  | 27.4% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 41.96 ms | 9.06 ms | 5.23 ms | 1.00 | 1.38 | 15708.0 KB | 1.00 |  |  | Loss +37.7% |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | MiniExcel | 66.03 ms | 3.98 ms | 2.30 ms | 1.57 | 2.17 | 93246.9 KB | 5.94 |  |  | 57.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 237.95 ms |  |  | 5.67 | 7.81 |  |  |  |  | 467.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | ClosedXML | 342.73 ms | 67.26 ms | 38.83 ms | 8.17 | 11.25 | 210638.1 KB | 13.41 |  |  | 716.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | EPPlus | 357.01 ms | 8.79 ms | 5.08 ms | 8.51 | 11.72 | 211783.2 KB | 13.48 |  |  | 750.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | OfficeIMO.Excel | 33.10 ms | 14.46 ms | 8.35 ms | 1.00 | 1.00 | 7543.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | ClosedXML | 286.93 ms | 102.12 ms | 58.96 ms | 8.67 | 8.67 | 105218.5 KB | 13.95 |  |  | 766.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | EPPlus | 318.34 ms | 35.38 ms | 20.43 ms | 9.62 | 9.62 | 106294.3 KB | 14.09 |  |  | 861.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | Sylvan.Data.Excel | 27.64 ms | 1.63 ms | 0.94 ms | 0.78 | 1.00 | 5692.3 KB | 0.45 |  |  | 21.6% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | LargeXlsx | 33.94 ms | 0.81 ms | 0.47 ms | 0.96 | 1.23 | 8341.2 KB | 0.66 |  |  | 3.7% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | OfficeIMO.Excel | 35.24 ms | 0.30 ms | 0.17 ms | 1.00 | 1.28 | 12673.9 KB | 1.00 |  |  | Loss +27.5% |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | MiniExcel | 70.92 ms | 2.80 ms | 1.62 ms | 2.01 | 2.57 | 92189.6 KB | 7.27 |  |  | 101.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | EPPlus 4.5.3.3 | 221.32 ms |  |  | 6.28 | 8.01 |  |  |  |  | 527.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | ClosedXML | 269.04 ms | 3.70 ms | 2.13 ms | 7.63 | 9.73 | 104197.0 KB | 8.22 |  |  | 663.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | EPPlus | 325.80 ms | 2.38 ms | 1.37 ms | 9.24 | 11.79 | 117370.5 KB | 9.26 |  |  | 824.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | OfficeIMO.Excel | 40.02 ms | 2.06 ms | 1.19 ms | 1.00 | 1.00 | 9491.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | LargeXlsx | 48.19 ms | 6.12 ms | 3.53 ms | 1.20 | 1.20 | 9257.9 KB | 0.98 |  |  | 20.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | MiniExcel | 102.42 ms | 3.15 ms | 1.82 ms | 2.56 | 2.56 | 108118.7 KB | 11.39 |  |  | 155.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | EPPlus | 548.05 ms | 17.40 ms | 10.04 ms | 13.69 | 13.69 | 135640.3 KB | 14.29 |  |  | 1269.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | ClosedXML | 563.87 ms | 21.13 ms | 12.20 ms | 14.09 | 14.09 | 280363.8 KB | 29.54 |  |  | 1309.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | LargeXlsx | 40.03 ms | 1.71 ms | 0.99 ms | 0.94 | 1.00 | 10795.2 KB | 0.93 |  |  | 6.1% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | OfficeIMO.Excel | 42.64 ms | 1.72 ms | 0.99 ms | 1.00 | 1.07 | 11547.8 KB | 1.00 |  |  | Loss +6.5% |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | MiniExcel | 146.06 ms | 4.11 ms | 2.38 ms | 3.43 | 3.65 | 226876.1 KB | 19.65 |  |  | 242.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | ClosedXML | 926.06 ms | 3.81 ms | 2.20 ms | 21.72 | 23.13 | 759818.3 KB | 65.80 |  |  | 2071.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | OfficeIMO.Excel | 15.26 ms | 0.79 ms | 0.45 ms | 1.00 | 1.00 | 15416.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | MiniExcel | 29.36 ms | 1.84 ms | 1.06 ms | 1.92 | 1.92 | 73751.2 KB | 4.78 |  |  | 92.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | EPPlus 4.5.3.3 | 95.69 ms |  |  | 6.27 | 6.27 |  |  |  |  | 527.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | ClosedXML | 108.89 ms | 5.04 ms | 2.91 ms | 7.14 | 7.14 | 104233.3 KB | 6.76 |  |  | 613.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | EPPlus | 173.88 ms | 8.97 ms | 5.18 ms | 11.40 | 11.40 | 84343.7 KB | 5.47 |  |  | 1039.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | OfficeIMO.Excel | 18.65 ms | 2.47 ms | 1.43 ms | 1.00 | 1.00 | 15027.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | EPPlus 4.5.3.3 | 92.38 ms |  |  | 4.95 | 4.95 |  |  |  |  | 395.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | ClosedXML | 116.93 ms | 13.91 ms | 8.03 ms | 6.27 | 6.27 | 104233.3 KB | 6.94 |  |  | 527.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | EPPlus | 192.78 ms | 26.96 ms | 15.57 ms | 10.34 | 10.34 | 84343.7 KB | 5.61 |  |  | 933.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | OfficeIMO.Excel | 13.14 ms | 0.12 ms | 0.07 ms | 1.00 | 1.00 | 13651.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | ClosedXML | 145.31 ms | 5.37 ms | 3.10 ms | 11.05 | 11.05 | 131493.2 KB | 9.63 |  |  | 1005.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | EPPlus | 199.21 ms | 2.61 ms | 1.51 ms | 15.16 | 15.16 | 97646.6 KB | 7.15 |  |  | 1415.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | OfficeIMO.Excel | 16.32 ms | 4.93 ms | 2.84 ms | 1.00 | 1.00 | 7192.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | ClosedXML | 111.50 ms | 19.68 ms | 11.36 ms | 6.83 | 6.83 | 84512.0 KB | 11.75 |  |  | 583.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | EPPlus | 235.21 ms | 48.39 ms | 27.94 ms | 14.41 | 14.41 | 69934.9 KB | 9.72 |  |  | 1340.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | LargeXlsx | 28.34 ms | 0.58 ms | 0.34 ms | 0.86 | 1.00 | 5606.0 KB | 0.45 |  |  | 14.4% faster than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 33.11 ms | 2.72 ms | 1.57 ms | 1.00 | 1.17 | 12583.6 KB | 1.00 |  |  | Loss +16.8% |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | MiniExcel | 63.74 ms | 0.64 ms | 0.37 ms | 1.93 | 2.25 | 93246.9 KB | 7.41 |  |  | 92.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 228.79 ms |  |  | 6.91 | 8.07 |  |  |  |  | 591.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | ClosedXML | 279.93 ms | 5.49 ms | 3.17 ms | 8.46 | 9.88 | 104197.0 KB | 8.28 |  |  | 745.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | EPPlus | 348.97 ms | 11.72 ms | 6.76 ms | 10.54 | 12.31 | 117370.5 KB | 9.33 |  |  | 954.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | LargeXlsx | 28.51 ms | 0.09 ms | 0.05 ms | 0.88 | 1.00 | 5606.0 KB | 0.49 |  |  | 12.5% faster than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | OfficeIMO.Excel | 32.57 ms | 1.62 ms | 0.94 ms | 1.00 | 1.14 | 11325.4 KB | 1.00 |  |  | Loss +14.2% |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | MiniExcel | 63.60 ms | 1.85 ms | 1.07 ms | 1.95 | 2.23 | 93246.9 KB | 8.23 |  |  | 95.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | EPPlus 4.5.3.3 | 219.47 ms |  |  | 6.74 | 7.70 |  |  |  |  | 573.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | ClosedXML | 279.24 ms | 4.47 ms | 2.58 ms | 8.57 | 9.80 | 104197.0 KB | 9.20 |  |  | 757.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | EPPlus | 340.20 ms | 18.08 ms | 10.44 ms | 10.45 | 11.93 | 117370.5 KB | 10.36 |  |  | 944.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 28.54 ms | 0.61 ms | 0.35 ms | 0.74 | 1.00 | 5606.0 KB | 0.57 |  |  | 26.3% faster than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 38.73 ms | 1.66 ms | 0.96 ms | 1.00 | 1.36 | 9851.0 KB | 1.00 |  |  | Loss +35.7% |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | MiniExcel | 64.00 ms | 1.06 ms | 0.61 ms | 1.65 | 2.24 | 93246.9 KB | 9.47 |  |  | 65.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | ClosedXML | 281.39 ms | 14.37 ms | 8.29 ms | 7.27 | 9.86 | 104197.0 KB | 10.58 |  |  | 626.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | EPPlus | 334.00 ms | 5.85 ms | 3.38 ms | 8.62 | 11.70 | 117370.5 KB | 11.91 |  |  | 762.4% slower than OfficeIMO |
