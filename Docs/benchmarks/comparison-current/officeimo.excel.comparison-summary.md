# OfficeIMO.Excel Comparison Summary

This is the suite-level decision table. Mean, standard deviation, standard error, ratios, and allocations come from the lightweight rotated local runner; they are meant for engineering direction. Use the BenchmarkDotNet benchmark classes when a publication-grade `Error` column is required.

## At a glance

| Row count | Artifact | Workload | Category | OfficeIMO wins | OfficeIMO losses | Biggest loss |
| ---: | --- | --- | --- | ---: | ---: | --- |
| 2500 | dense-helloworld-comparison | read | Other | 1 | 1 | dense-helloworld-read-range: Loss +34.6% vs Sylvan.Data.Excel |
| 2500 | package-profile | package | Package size | 44 | 10 | write-insertobjects-legacy-dictionaries-direct: Loss +52.0% vs LargeXlsx |
| 2500 | speed-comparison | mutate | AutoFit and mutation | 2 | 0 |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | 4 | 0 |  |
| 2500 | speed-comparison | other | Object projection | 2 | 0 |  |
| 2500 | speed-comparison | other | Range and table read | 4 | 0 |  |
| 2500 | speed-comparison | other | Real-world feature mix | 6 | 0 |  |
| 2500 | speed-comparison | other | Real-world report | 2 | 0 |  |
| 2500 | speed-comparison | other | Report workbook | 4 | 0 |  |
| 2500 | speed-comparison | read | Formula write/read | 1 | 0 |  |
| 2500 | speed-comparison | read | Other | 1 | 2 | large-sparse-row-read: Loss +52.4% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | read | Range and table read | 4 | 3 | read-used-range: Loss +188.4% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | read | Streaming read | 1 | 3 | read-top-range-stream-small-chunks: Loss +61.5% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | read | Typed object read | 2 | 0 |  |
| 2500 | speed-comparison | write | AutoFit and mutation | 5 | 0 |  |
| 2500 | speed-comparison | write | Cell writer | 7 | 0 |  |
| 2500 | speed-comparison | write | DataSet table export | 3 | 0 |  |
| 2500 | speed-comparison | write | DataTable table export | 4 | 0 |  |
| 2500 | speed-comparison | write | Formatted report write | 1 | 0 |  |
| 2500 | speed-comparison | write | Formula write/read | 1 | 0 |  |
| 2500 | speed-comparison | write | Other | 2 | 2 | write-powershell-psobject-mixed-direct: Loss +38.2% vs LargeXlsx |
| 2500 | speed-comparison | write | Plain cell export | 2 | 2 | append-plain-rows: Loss +62.7% vs LargeXlsx |
| 2500 | speed-comparison | write | Plain streaming export | 1 | 1 | write-datareader-plain: Loss +2.9% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | write | Plain string export | 0 | 1 | write-blog-2023-20-string-columns: Loss +3.2% vs LargeXlsx |
| 2500 | speed-comparison | write | Shared string write | 4 | 0 |  |
| 2500 | speed-comparison | write | Typed object export | 0 | 3 | write-insertobjects-flat-dictionaries-direct: Loss +46.3% vs LargeXlsx |
| 25000 | dense-helloworld-comparison | read | Other | 0 | 2 | dense-helloworld-read-range: Loss +21.2% vs Sylvan.Data.Excel |
| 25000 | package-profile | package | Package size | 42 | 12 | write-insertobjects-legacy-dictionaries-direct: Loss +61.4% vs LargeXlsx |
| 25000 | speed-comparison | mutate | AutoFit and mutation | 2 | 0 |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | 4 | 0 |  |
| 25000 | speed-comparison | other | Object projection | 2 | 0 |  |
| 25000 | speed-comparison | other | Range and table read | 4 | 0 |  |
| 25000 | speed-comparison | other | Real-world feature mix | 6 | 0 |  |
| 25000 | speed-comparison | other | Real-world report | 2 | 0 |  |
| 25000 | speed-comparison | other | Report workbook | 4 | 0 |  |
| 25000 | speed-comparison | read | Formula write/read | 1 | 0 |  |
| 25000 | speed-comparison | read | Other | 2 | 1 | shared-string-read: Loss +19.7% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | read | Range and table read | 2 | 5 | read-used-range: Loss +92.7% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | read | Streaming read | 1 | 3 | read-top-range-stream-small-chunks: Loss +23.9% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | read | Typed object read | 1 | 1 | read-objects: Loss +19.0% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | write | AutoFit and mutation | 5 | 0 |  |
| 25000 | speed-comparison | write | Cell writer | 7 | 0 |  |
| 25000 | speed-comparison | write | DataSet table export | 3 | 0 |  |
| 25000 | speed-comparison | write | DataTable table export | 4 | 0 |  |
| 25000 | speed-comparison | write | Formatted report write | 1 | 0 |  |
| 25000 | speed-comparison | write | Formula write/read | 1 | 0 |  |
| 25000 | speed-comparison | write | Other | 2 | 2 | write-powershell-mixed-objects-direct: Loss +12.3% vs LargeXlsx |
| 25000 | speed-comparison | write | Plain cell export | 2 | 2 | append-plain-rows: Loss +45.2% vs LargeXlsx |
| 25000 | speed-comparison | write | Plain streaming export | 1 | 1 | write-datareader-plain: Loss +48.9% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | write | Plain string export | 0 | 1 | write-blog-2023-20-string-columns: Loss +26.0% vs LargeXlsx |
| 25000 | speed-comparison | write | Shared string write | 4 | 0 |  |
| 25000 | speed-comparison | write | Typed object export | 0 | 3 | write-insertobjects-flat-dictionaries-direct: Loss +34.0% vs LargeXlsx |

## OfficeIMO decision table

| Row count | Artifact | Workload | Category | Scenario | OfficeIMO mean | Best | OfficeIMO vs best | Alloc | Package |
| ---: | --- | --- | --- | --- | ---: | --- | ---: | ---: | ---: |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | 5.98 ms | Sylvan.Data.Excel | Loss +34.6% | 2411.1 KB |  |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | 5.67 ms | OfficeIMO.Excel | Win | 2489.5 KB |  |
| 2500 | package-profile | package | Package size | append-plain-rows | 2.04 ms | LargeXlsx | Loss +31.2% | 1576.9 KB | 63.0 KB |
| 2500 | package-profile | package | Package size | autofit-existing | 8.95 ms | OfficeIMO.Excel | Win | 1895.5 KB | 143.2 KB |
| 2500 | package-profile | package | Package size | large-shared-strings | 2.13 ms | OfficeIMO.Excel | Win | 2440.3 KB | 55.2 KB |
| 2500 | package-profile | package | Package size | realworld-autofilter | 3.93 ms | OfficeIMO.Excel | Win | 1340.4 KB | 142.4 KB |
| 2500 | package-profile | package | Package size | realworld-charts | 5.01 ms | OfficeIMO.Excel | Win | 1892.9 KB | 147.6 KB |
| 2500 | package-profile | package | Package size | realworld-conditional-formatting | 3.98 ms | OfficeIMO.Excel | Win | 1405.8 KB | 142.7 KB |
| 2500 | package-profile | package | Package size | realworld-data-validation | 4.66 ms | OfficeIMO.Excel | Win | 1356.1 KB | 142.5 KB |
| 2500 | package-profile | package | Package size | realworld-freeze-panes | 4.21 ms | OfficeIMO.Excel | Win | 1342.8 KB | 142.5 KB |
| 2500 | package-profile | package | Package size | realworld-pivot-table | 40.43 ms | OfficeIMO.Excel | Win | 15676.8 KB | 200.4 KB |
| 2500 | package-profile | package | Package size | realworld-report-all-in-one | 18.79 ms | OfficeIMO.Excel | Win | 16478.1 KB | 206.5 KB |
| 2500 | package-profile | package | Package size | realworld-report-chart-first | 17.90 ms | OfficeIMO.Excel | Win | 7452.6 KB | 206.5 KB |
| 2500 | package-profile | package | Package size | realworld-report-core | 6.40 ms | OfficeIMO.Excel | Win | 1488.5 KB | 143.9 KB |
| 2500 | package-profile | package | Package size | realworld-report-extra-column | 24.24 ms | OfficeIMO.Excel | Win | 17641.2 KB | 219.1 KB |
| 2500 | package-profile | package | Package size | realworld-report-no-autofit | 18.99 ms | OfficeIMO.Excel | Win | 16466.7 KB | 206.4 KB |
| 2500 | package-profile | package | Package size | realworld-report-post-mutation | 19.99 ms | OfficeIMO.Excel | Win | 16487.6 KB | 206.6 KB |
| 2500 | package-profile | package | Package size | realworld-report-shuffled-columns | 18.80 ms | OfficeIMO.Excel | Win | 16485.8 KB | 211.2 KB |
| 2500 | package-profile | package | Package size | report-workbook | 24.52 ms | OfficeIMO.Excel | Win | 20494.1 KB | 275.6 KB |
| 2500 | package-profile | package | Package size | report-workbook-core | 6.63 ms | OfficeIMO.Excel | Win | 2711.1 KB | 187.5 KB |
| 2500 | package-profile | package | Package size | report-workbook-datatable | 22.31 ms | OfficeIMO.Excel | Win | 20765.6 KB | 275.6 KB |
| 2500 | package-profile | package | Package size | report-workbook-datatable-core | 6.43 ms | OfficeIMO.Excel | Win | 2982.7 KB | 187.5 KB |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | 4.71 ms | LargeXlsx | Loss +17.5% | 1676.8 KB | 216.7 KB |
| 2500 | package-profile | package | Package size | write-bulk-report | 4.71 ms | OfficeIMO.Excel | Win | 1401.7 KB | 143.2 KB |
| 2500 | package-profile | package | Package size | write-cellformula | 2.60 ms | OfficeIMO.Excel | Win | 1383.3 KB | 66.6 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-empty-strings | 1.98 ms | OfficeIMO.Excel | Win | 1787.1 KB | 44.2 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-numbers | 2.19 ms | OfficeIMO.Excel | Win | 1119.9 KB | 47.0 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-object-mixed | 2.58 ms | OfficeIMO.Excel | Win | 1763.1 KB | 61.2 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse | 3.03 ms | OfficeIMO.Excel | Win | 1506.7 KB | 62.0 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | 3.09 ms | OfficeIMO.Excel | Win | 1506.8 KB | 62.0 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-scalars | 1.92 ms | OfficeIMO.Excel | Win | 1138.1 KB | 46.3 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-strings | 3.43 ms | OfficeIMO.Excel | Win | 2617.0 KB | 55.1 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-distinct | 2.19 ms | OfficeIMO.Excel | Win | 2379.2 KB | 51.8 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-repeated | 1.87 ms | OfficeIMO.Excel | Win | 1579.8 KB | 40.0 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-temporal | 3.23 ms | OfficeIMO.Excel | Win | 1435.5 KB | 63.3 KB |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | 1.59 ms | LargeXlsx | Loss +11.8% | 1092.0 KB | 48.2 KB |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | 4.32 ms | LargeXlsx | Loss +37.4% | 2081.1 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-datareader-plain | 4.51 ms | Sylvan.Data.Excel | Loss +48.4% | 1763.0 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-datareader-table | 4.96 ms | OfficeIMO.Excel | Win | 1774.9 KB | 138.8 KB |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | 4.81 ms | OfficeIMO.Excel | Win | 1781.2 KB | 138.9 KB |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | 4.53 ms | LargeXlsx | Loss +4.0% | 2140.6 KB | 131.1 KB |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | 4.57 ms | OfficeIMO.Excel | Win | 2880.2 KB | 176.0 KB |
| 2500 | package-profile | package | Package size | write-dataset-tables | 3.97 ms | OfficeIMO.Excel | Win | 2066.1 KB | 138.9 KB |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | 4.42 ms | OfficeIMO.Excel | Win | 2078.7 KB | 139.2 KB |
| 2500 | package-profile | package | Package size | write-datatable-direct | 4.44 ms | OfficeIMO.Excel | Win | 1748.6 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | 3.95 ms | OfficeIMO.Excel | Win | 1760.7 KB | 138.8 KB |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | 4.06 ms | LargeXlsx | Loss +23.9% | 1769.2 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | 5.08 ms | OfficeIMO.Excel | Win | 1347.1 KB | 142.4 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | 4.15 ms | OfficeIMO.Excel | Win | 1339.3 KB | 142.3 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | 5.56 ms | OfficeIMO.Excel | Win | 1505.3 KB | 138.1 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | 5.07 ms | LargeXlsx | Loss +25.8% | 1497.5 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | 5.12 ms | LargeXlsx | Loss +52.0% | 1770.1 KB | 142.3 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | 4.66 ms | OfficeIMO.Excel | Win | 1346.4 KB | 142.4 KB |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | 5.92 ms | OfficeIMO.Excel | Win | 2341.7 KB | 183.1 KB |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | 4.99 ms | LargeXlsx | Loss +13.1% | 1507.7 KB | 182.4 KB |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | 22.03 ms | OfficeIMO.Excel | Win | 4502.3 KB | 651.0 KB |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | 9.58 ms | OfficeIMO.Excel | Win | 1895.3 KB |  |
| 2500 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | 15.80 ms | OfficeIMO.Excel | Win | 16466.9 KB |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | 12.31 ms | OfficeIMO.Excel | Win | 7452.5 KB |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | 18.17 ms | OfficeIMO.Excel | Win | 17641.5 KB |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | 16.10 ms | OfficeIMO.Excel | Win | 16487.8 KB |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | 22.37 ms | OfficeIMO.Excel | Win | 16482.7 KB |  |
| 2500 | speed-comparison | other | Object projection | build-object-datatable-dictionaries | 1.58 ms | OfficeIMO.Excel | Win | 564.2 KB |  |
| 2500 | speed-comparison | other | Object projection | build-object-datatable-typed | 1.24 ms | OfficeIMO.Excel | Win | 856.9 KB |  |
| 2500 | speed-comparison | other | Range and table read | enumerate-cells | 6.05 ms | OfficeIMO.Excel | Win | 2531.8 KB |  |
| 2500 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | 3.93 ms | OfficeIMO.Excel | Win | 526.1 KB |  |
| 2500 | speed-comparison | other | Range and table read | enumerate-range | 5.77 ms | OfficeIMO.Excel | Win | 2531.9 KB |  |
| 2500 | speed-comparison | other | Range and table read | enumerate-top-range | 0.74 ms | OfficeIMO.Excel | Win | 285.5 KB |  |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-autofilter | 3.76 ms | OfficeIMO.Excel | Win | 1340.4 KB |  |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-charts | 5.08 ms | OfficeIMO.Excel | Win | 1893.1 KB |  |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | 3.99 ms | OfficeIMO.Excel | Win | 1405.8 KB |  |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-data-validation | 3.79 ms | OfficeIMO.Excel | Win | 1356.1 KB |  |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | 3.54 ms | OfficeIMO.Excel | Win | 1342.9 KB |  |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | 15.54 ms | OfficeIMO.Excel | Win | 15676.8 KB |  |
| 2500 | speed-comparison | other | Real-world report | realworld-report-all-in-one | 18.80 ms | OfficeIMO.Excel | Win | 16477.8 KB |  |
| 2500 | speed-comparison | other | Real-world report | realworld-report-core | 4.45 ms | OfficeIMO.Excel | Win | 1488.6 KB |  |
| 2500 | speed-comparison | other | Report workbook | report-workbook | 31.04 ms | OfficeIMO.Excel | Win | 20450.8 KB |  |
| 2500 | speed-comparison | other | Report workbook | report-workbook-core | 7.25 ms | OfficeIMO.Excel | Win | 2711.1 KB |  |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable | 27.85 ms | OfficeIMO.Excel | Win | 20765.7 KB |  |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable-core | 6.52 ms | OfficeIMO.Excel | Win | 2982.8 KB |  |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | 2.66 ms | OfficeIMO.Excel | Win | 709.4 KB |  |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | 1.20 ms | OfficeIMO.Excel | Win | 177.4 KB |  |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | 2.64 ms | Sylvan.Data.Excel | Loss +52.4% | 177.5 KB |  |
| 2500 | speed-comparison | read | Other | shared-string-read | 3.02 ms | Sylvan.Data.Excel | Loss +28.5% | 1056.7 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | 4.63 ms | OfficeIMO.Excel | Win | 374.5 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-datatable | 7.97 ms | Sylvan.Data.Excel | Loss +24.7% | 3594.5 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | 4.69 ms | Sylvan.Data.Excel, OfficeIMO.Excel | Win | 551.0 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-range | 9.55 ms | OfficeIMO.Excel | Win | 2692.7 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | 6.14 ms | OfficeIMO.Excel | Win | 2751.4 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-top-range | 0.61 ms | Sylvan.Data.Excel | Loss +14.5% | 296.0 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-used-range | 13.46 ms | Sylvan.Data.Excel | Loss +188.4% | 3472.6 KB |  |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | 3.53 ms | OfficeIMO.Excel | Win | 377.8 KB |  |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | 5.94 ms | Sylvan.Data.Excel | Loss +31.8% | 2771.3 KB |  |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | 0.53 ms | Sylvan.Data.Excel | Loss +27.0% | 299.4 KB |  |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | 0.67 ms | Sylvan.Data.Excel | Loss +61.5% | 300.1 KB |  |
| 2500 | speed-comparison | read | Typed object read | read-objects | 13.24 ms | OfficeIMO.Excel | Win | 2441.9 KB |  |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | 7.59 ms | OfficeIMO.Excel | Win | 2422.8 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | 7.58 ms | OfficeIMO.Excel | Win | 1781.2 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | 5.97 ms | OfficeIMO.Excel | Win | 2080.3 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | 15.06 ms | OfficeIMO.Excel | Win | 1347.1 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | 7.52 ms | OfficeIMO.Excel | Win | 1505.3 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | 7.62 ms | OfficeIMO.Excel | Win | 1346.4 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | 2.63 ms | OfficeIMO.Excel | Win | 1787.1 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | 3.20 ms | OfficeIMO.Excel | Win | 1119.9 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | 3.41 ms | OfficeIMO.Excel | Win | 1763.1 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | 3.72 ms | OfficeIMO.Excel | Win | 1506.7 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | 3.08 ms | OfficeIMO.Excel | Win | 1506.8 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | 2.57 ms | OfficeIMO.Excel | Win | 1138.1 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | 3.64 ms | OfficeIMO.Excel | Win | 1435.5 KB |  |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | 4.95 ms | OfficeIMO.Excel | Win | 2065.5 KB |  |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | 7.34 ms | OfficeIMO.Excel | Win | 2880.2 KB |  |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | 5.55 ms | OfficeIMO.Excel | Win | 2067.7 KB |  |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | 4.67 ms | OfficeIMO.Excel | Win | 1774.9 KB |  |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | 7.70 ms | OfficeIMO.Excel | Win | 1748.6 KB |  |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | 5.28 ms | OfficeIMO.Excel | Win | 1487.2 KB |  |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | 4.80 ms | OfficeIMO.Excel | Win | 1760.7 KB |  |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | 5.80 ms | OfficeIMO.Excel | Win | 1403.3 KB |  |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | 4.36 ms | OfficeIMO.Excel | Win | 1620.6 KB |  |
| 2500 | speed-comparison | write | Other | write-dictionary-objects-table-direct | 7.13 ms | OfficeIMO.Excel | Win | 2051.4 KB |  |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | 6.04 ms | LargeXlsx | Loss +33.1% | 2341.7 KB |  |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | 6.71 ms | LargeXlsx | Loss +38.2% | 1507.7 KB |  |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | 22.25 ms | OfficeIMO.Excel | Win | 4502.3 KB |  |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | 2.71 ms | LargeXlsx | Loss +62.7% | 1576.9 KB |  |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | 1.72 ms | LargeXlsx | Loss +20.0% | 1092.0 KB |  |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | 4.25 ms | OfficeIMO.Excel | Win | 2081.1 KB |  |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | 2.46 ms | OfficeIMO.Excel | Win | 1494.0 KB |  |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | 4.84 ms | Sylvan.Data.Excel | Loss +2.9% | 1763.0 KB |  |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | 7.33 ms | OfficeIMO.Excel | Win | 2140.6 KB |  |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | 4.45 ms | LargeXlsx | Loss +3.2% | 1676.8 KB |  |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | 2.05 ms | OfficeIMO.Excel | Win | 2440.3 KB |  |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | 2.85 ms | OfficeIMO.Excel | Win | 2617.0 KB |  |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | 2.50 ms | OfficeIMO.Excel | Win | 2379.2 KB |  |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | 2.06 ms | OfficeIMO.Excel | Win | 1579.8 KB |  |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | 4.14 ms | LargeXlsx | Loss +22.5% | 1769.2 KB |  |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | 5.91 ms | LargeXlsx | Loss +44.3% | 1339.3 KB |  |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | 7.23 ms | LargeXlsx | Loss +46.3% | 1497.5 KB |  |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | 73.99 ms | Sylvan.Data.Excel | Loss +21.2% | 23622.1 KB |  |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | 68.98 ms | Sylvan.Data.Excel | Loss +3.5% | 24404.2 KB |  |
| 25000 | package-profile | package | Package size | append-plain-rows | 20.45 ms | LargeXlsx | Loss +42.7% | 10843.1 KB | 610.4 KB |
| 25000 | package-profile | package | Package size | autofit-existing | 92.81 ms | OfficeIMO.Excel | Win | 15708.5 KB | 1429.3 KB |
| 25000 | package-profile | package | Package size | large-shared-strings | 20.06 ms | OfficeIMO.Excel | Win | 15744.9 KB | 529.7 KB |
| 25000 | package-profile | package | Package size | realworld-autofilter | 43.84 ms | OfficeIMO.Excel | Win | 11494.9 KB | 1428.5 KB |
| 25000 | package-profile | package | Package size | realworld-charts | 45.87 ms | OfficeIMO.Excel | Win | 12553.6 KB | 1433.7 KB |
| 25000 | package-profile | package | Package size | realworld-conditional-formatting | 42.72 ms | OfficeIMO.Excel | Win | 11560.2 KB | 1428.8 KB |
| 25000 | package-profile | package | Package size | realworld-data-validation | 42.86 ms | OfficeIMO.Excel | Win | 11510.5 KB | 1428.6 KB |
| 25000 | package-profile | package | Package size | realworld-freeze-panes | 42.53 ms | OfficeIMO.Excel | Win | 11497.3 KB | 1428.6 KB |
| 25000 | package-profile | package | Package size | realworld-pivot-table | 411.32 ms | OfficeIMO.Excel | Win | 143628.5 KB | 1979.8 KB |
| 25000 | package-profile | package | Package size | realworld-report-all-in-one | 307.01 ms | OfficeIMO.Excel | Win | 145144.6 KB | 1985.9 KB |
| 25000 | package-profile | package | Package size | realworld-report-chart-first | 119.04 ms | OfficeIMO.Excel | Win | 55261.1 KB | 1985.9 KB |
| 25000 | package-profile | package | Package size | realworld-report-core | 46.76 ms | OfficeIMO.Excel | Win | 11648.7 KB | 1430.0 KB |
| 25000 | package-profile | package | Package size | realworld-report-extra-column | 325.78 ms | OfficeIMO.Excel | Win | 156862.3 KB | 2110.8 KB |
| 25000 | package-profile | package | Package size | realworld-report-no-autofit | 320.07 ms | OfficeIMO.Excel | Win | 145135.1 KB | 1985.8 KB |
| 25000 | package-profile | package | Package size | realworld-report-post-mutation | 336.55 ms | OfficeIMO.Excel | Win | 145153.9 KB | 1986.0 KB |
| 25000 | package-profile | package | Package size | realworld-report-shuffled-columns | 325.90 ms | OfficeIMO.Excel | Win | 145205.6 KB | 2046.1 KB |
| 25000 | package-profile | package | Package size | report-workbook | 443.58 ms | OfficeIMO.Excel | Win | 196356.9 KB | 2672.0 KB |
| 25000 | package-profile | package | Package size | report-workbook-core | 63.24 ms | OfficeIMO.Excel | Win | 10979.4 KB | 1850.9 KB |
| 25000 | package-profile | package | Package size | report-workbook-datatable | 467.76 ms | OfficeIMO.Excel | Win | 199104.2 KB | 2672.0 KB |
| 25000 | package-profile | package | Package size | report-workbook-datatable-core | 76.78 ms | OfficeIMO.Excel | Win | 13725.0 KB | 1850.9 KB |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | 62.30 ms | LargeXlsx | Loss +13.1% | 11708.2 KB | 2228.8 KB |
| 25000 | package-profile | package | Package size | write-bulk-report | 41.76 ms | OfficeIMO.Excel | Win | 11561.8 KB | 1429.3 KB |
| 25000 | package-profile | package | Package size | write-cellformula | 31.55 ms | OfficeIMO.Excel | Win | 10112.0 KB | 670.3 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-empty-strings | 16.17 ms | OfficeIMO.Excel | Win | 6896.4 KB | 451.4 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-numbers | 17.61 ms | OfficeIMO.Excel | Win | 5970.9 KB | 462.6 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-object-mixed | 22.69 ms | OfficeIMO.Excel | Win | 8332.7 KB | 585.1 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse | 25.91 ms | OfficeIMO.Excel | Win | 7416.0 KB | 607.1 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | 24.04 ms | OfficeIMO.Excel | Win | 7416.1 KB | 607.1 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-scalars | 15.41 ms | OfficeIMO.Excel | Win | 6144.6 KB | 441.9 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-strings | 21.93 ms | OfficeIMO.Excel | Win | 15360.4 KB | 527.8 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-distinct | 17.75 ms | OfficeIMO.Excel | Win | 13824.1 KB | 499.0 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-repeated | 16.46 ms | OfficeIMO.Excel | Win | 7525.3 KB | 376.0 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-temporal | 28.28 ms | OfficeIMO.Excel | Win | 7482.6 KB | 620.5 KB |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | 14.67 ms | LargeXlsx | Loss +16.5% | 6961.7 KB | 455.5 KB |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | 46.60 ms | LargeXlsx | Loss +22.1% | 16036.5 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-datareader-plain | 58.95 ms | Sylvan.Data.Excel | Loss +47.1% | 13002.3 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-datareader-table | 46.01 ms | OfficeIMO.Excel | Win | 13020.3 KB | 1385.7 KB |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | 55.69 ms | OfficeIMO.Excel | Win | 13026.6 KB | 1385.9 KB |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | 42.31 ms | OfficeIMO.Excel | Win | 9819.7 KB | 1329.2 KB |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | 49.85 ms | OfficeIMO.Excel | Win | 13458.5 KB | 1795.1 KB |
| 25000 | package-profile | package | Package size | write-dataset-tables | 49.25 ms | OfficeIMO.Excel | Win | 10288.1 KB | 1376.4 KB |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | 69.82 ms | OfficeIMO.Excel | Win | 10300.7 KB | 1376.7 KB |
| 25000 | package-profile | package | Package size | write-datatable-direct | 43.62 ms | LargeXlsx | Loss +9.5% | 12715.7 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | 46.22 ms | OfficeIMO.Excel | Win | 12733.8 KB | 1385.7 KB |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | 41.97 ms | LargeXlsx | Loss +29.0% | 12912.0 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | 54.71 ms | OfficeIMO.Excel | Win | 11501.6 KB | 1428.5 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | 42.85 ms | LargeXlsx | Loss +22.0% | 11493.8 KB | 1428.4 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | 59.48 ms | OfficeIMO.Excel | Win | 10187.2 KB | 1385.1 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | 51.95 ms | LargeXlsx | Loss +43.5% | 10179.4 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | 62.41 ms | LargeXlsx | Loss +61.4% | 15791.7 KB | 1428.4 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | 46.77 ms | OfficeIMO.Excel | Win | 11500.9 KB | 1428.5 KB |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | 57.93 ms | LargeXlsx | Loss +13.5% | 10577.2 KB | 1828.0 KB |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | 54.45 ms | LargeXlsx | Loss +10.8% | 9942.2 KB | 1828.0 KB |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | 271.73 ms | OfficeIMO.Excel, LargeXlsx | Win | 36150.1 KB | 6725.6 KB |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | 85.60 ms | OfficeIMO.Excel | Win | 15708.5 KB |  |
| 25000 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | 234.38 ms | OfficeIMO.Excel | Win | 145134.4 KB |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | 93.06 ms | OfficeIMO.Excel | Win | 55262.1 KB |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | 356.27 ms | OfficeIMO.Excel | Win | 156857.7 KB |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | 320.89 ms | OfficeIMO.Excel | Win | 145151.3 KB |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | 253.35 ms | OfficeIMO.Excel | Win | 145205.7 KB |  |
| 25000 | speed-comparison | other | Object projection | build-object-datatable-dictionaries | 8.95 ms | OfficeIMO.Excel | Win | 5164.3 KB |  |
| 25000 | speed-comparison | other | Object projection | build-object-datatable-typed | 8.15 ms | OfficeIMO.Excel | Win | 8093.8 KB |  |
| 25000 | speed-comparison | other | Range and table read | enumerate-cells | 52.17 ms | OfficeIMO.Excel | Win | 24530.9 KB |  |
| 25000 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | 35.19 ms | OfficeIMO.Excel | Win | 3844.6 KB |  |
| 25000 | speed-comparison | other | Range and table read | enumerate-range | 49.51 ms | OfficeIMO.Excel | Win | 24531.0 KB |  |
| 25000 | speed-comparison | other | Range and table read | enumerate-top-range | 0.61 ms | OfficeIMO.Excel | Win | 285.3 KB |  |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-autofilter | 43.12 ms | OfficeIMO.Excel | Win | 11494.9 KB |  |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-charts | 44.45 ms | OfficeIMO.Excel | Win | 12553.0 KB |  |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | 42.96 ms | OfficeIMO.Excel | Win | 11560.2 KB |  |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-data-validation | 43.52 ms | OfficeIMO.Excel | Win | 11510.5 KB |  |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | 41.39 ms | OfficeIMO.Excel | Win | 11497.3 KB |  |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | 299.58 ms | EPPlus 4.5.3.3, OfficeIMO.Excel | Win | 143621.8 KB |  |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | 226.59 ms | OfficeIMO.Excel | Win | 145143.8 KB |  |
| 25000 | speed-comparison | other | Real-world report | realworld-report-core | 46.34 ms | OfficeIMO.Excel | Win | 11648.7 KB |  |
| 25000 | speed-comparison | other | Report workbook | report-workbook | 342.77 ms | OfficeIMO.Excel | Win | 196317.1 KB |  |
| 25000 | speed-comparison | other | Report workbook | report-workbook-core | 51.45 ms | OfficeIMO.Excel | Win | 10979.4 KB |  |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable | 334.91 ms | OfficeIMO.Excel | Win | 199103.0 KB |  |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable-core | 50.73 ms | OfficeIMO.Excel | Win | 13725.0 KB |  |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | 28.98 ms | OfficeIMO.Excel | Win | 6219.1 KB |  |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | 0.86 ms | OfficeIMO.Excel | Win | 177.5 KB |  |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | 0.86 ms | OfficeIMO.Excel | Win | 177.6 KB |  |
| 25000 | speed-comparison | read | Other | shared-string-read | 31.30 ms | Sylvan.Data.Excel | Loss +19.7% | 9218.1 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | 32.46 ms | OfficeIMO.Excel | Win | 1122.4 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-datatable | 62.47 ms | Sylvan.Data.Excel | Loss +6.6% | 34645.8 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | 36.67 ms | OfficeIMO.Excel | Win | 4042.5 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-range | 52.68 ms | Sylvan.Data.Excel | Loss +20.5% | 26098.2 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | 54.25 ms | Sylvan.Data.Excel | Loss +5.0% | 26684.2 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-top-range | 0.57 ms | Sylvan.Data.Excel | Loss +13.4% | 296.0 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-used-range | 86.94 ms | Sylvan.Data.Excel | Loss +92.7% | 34151.6 KB |  |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | 36.09 ms | OfficeIMO.Excel | Win | 1125.6 KB |  |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | 51.84 ms | Sylvan.Data.Excel | Loss +14.3% | 26883.8 KB |  |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | 0.54 ms | Sylvan.Data.Excel | Loss +19.6% | 299.3 KB |  |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | 0.54 ms | Sylvan.Data.Excel | Loss +23.9% | 300.1 KB |  |
| 25000 | speed-comparison | read | Typed object read | read-objects | 56.11 ms | Sylvan.Data.Excel | Loss +19.0% | 23562.3 KB |  |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | 45.62 ms | OfficeIMO.Excel | Win | 23367.3 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | 42.59 ms | OfficeIMO.Excel | Win | 13026.6 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | 43.41 ms | OfficeIMO.Excel | Win | 10300.7 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | 39.51 ms | OfficeIMO.Excel | Win | 11501.6 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | 42.96 ms | OfficeIMO.Excel | Win | 10187.2 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | 37.23 ms | OfficeIMO.Excel | Win | 11500.9 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | 12.34 ms | OfficeIMO.Excel | Win | 6896.4 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | 14.87 ms | OfficeIMO.Excel | Win | 5970.9 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | 17.26 ms | OfficeIMO.Excel | Win | 8332.7 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | 19.43 ms | OfficeIMO.Excel | Win | 7416.0 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | 17.85 ms | OfficeIMO.Excel | Win | 7416.1 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | 10.97 ms | OfficeIMO.Excel | Win | 6144.6 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | 26.79 ms | OfficeIMO.Excel | Win | 7482.6 KB |  |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | 39.94 ms | OfficeIMO.Excel | Win | 13039.6 KB |  |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | 45.31 ms | OfficeIMO.Excel | Win | 13458.5 KB |  |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | 38.33 ms | OfficeIMO.Excel | Win | 10288.1 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | 37.14 ms | OfficeIMO.Excel | Win | 13020.3 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | 36.47 ms | LargeXlsx, OfficeIMO.Excel | Win | 12715.7 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | 37.27 ms | OfficeIMO.Excel | Win | 9999.4 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | 35.86 ms | OfficeIMO.Excel | Win | 12733.8 KB |  |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | 41.74 ms | OfficeIMO.Excel | Win | 11561.8 KB |  |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | 21.59 ms | OfficeIMO.Excel | Win | 10112.1 KB |  |
| 25000 | speed-comparison | write | Other | write-dictionary-objects-table-direct | 47.26 ms | OfficeIMO.Excel | Win | 15163.8 KB |  |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | 46.21 ms | LargeXlsx | Loss +12.3% | 10577.2 KB |  |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | 48.71 ms | LargeXlsx | Loss +4.9% | 9942.2 KB |  |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | 216.40 ms | OfficeIMO.Excel | Win | 36150.1 KB |  |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | 16.80 ms | LargeXlsx | Loss +45.2% | 10843.1 KB |  |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | 11.45 ms | OfficeIMO.Excel, LargeXlsx | Win | 6961.7 KB |  |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | 34.62 ms | LargeXlsx | Loss +17.5% | 16036.5 KB |  |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | 21.86 ms | OfficeIMO.Excel | Win | 7866.1 KB |  |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | 41.67 ms | Sylvan.Data.Excel | Loss +48.9% | 13002.3 KB |  |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | 34.90 ms | OfficeIMO.Excel | Win | 9819.7 KB |  |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | 48.91 ms | LargeXlsx | Loss +26.0% | 11708.2 KB |  |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | 20.03 ms | OfficeIMO.Excel | Win | 15744.9 KB |  |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | 17.16 ms | OfficeIMO.Excel | Win | 15360.4 KB |  |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | 12.68 ms | OfficeIMO.Excel | Win | 13824.1 KB |  |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | 12.81 ms | OfficeIMO.Excel | Win | 7525.3 KB |  |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | 35.25 ms | LargeXlsx | Loss +22.1% | 12912.0 KB |  |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | 35.22 ms | LargeXlsx | Loss +24.2% | 11493.8 KB |  |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | 39.65 ms | LargeXlsx | Loss +34.0% | 10179.4 KB |  |

## Full comparison table

| Row count | Artifact | Workload | Category | Scenario | Library | Mean | StdDev | StdErr | Ratio to OfficeIMO | Ratio to best | Alloc | Alloc ratio | Package | Package ratio | Outcome |
| ---: | --- | --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | Sylvan.Data.Excel | 4.44 ms | 0.52 ms | 0.30 ms | 0.74 | 1.00 | 362.3 KB | 0.15 |  |  | 25.7% faster than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | OfficeIMO.Excel | 5.98 ms | 0.08 ms | 0.05 ms | 1.00 | 1.35 | 2411.1 KB | 1.00 |  |  | Loss +34.6% |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | ExcelDataReader | 11.13 ms | 0.28 ms | 0.16 ms | 1.86 | 2.51 | 6887.4 KB | 2.86 |  |  | 86.1% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | MiniExcel | 16.89 ms | 1.86 ms | 1.07 ms | 2.82 | 3.80 | 21507.3 KB | 8.92 |  |  | 182.4% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | OfficeIMO.Excel | 5.67 ms | 0.65 ms | 0.37 ms | 1.00 | 1.00 | 2489.5 KB | 1.00 |  |  | Win |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | Sylvan.Data.Excel | 6.72 ms | 4.53 ms | 2.61 ms | 1.19 | 1.19 | 362.3 KB | 0.15 |  |  | 18.5% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | ExcelDataReader | 10.82 ms | 0.19 ms | 0.11 ms | 1.91 | 1.91 | 6887.4 KB | 2.77 |  |  | 91.0% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | MiniExcel | 21.82 ms | 6.93 ms | 4.00 ms | 3.85 | 3.85 | 21507.3 KB | 8.64 |  |  | 285.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | append-plain-rows | LargeXlsx | 1.55 ms | 0.08 ms | 0.05 ms | 0.76 | 1.00 | 296.4 KB | 0.19 | 63.1 KB | 1.00 | 23.8% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | append-plain-rows | OfficeIMO.Excel | 2.04 ms | 0.10 ms | 0.06 ms | 1.00 | 1.31 | 1576.9 KB | 1.00 | 63.0 KB | 1.00 | Loss +31.2% |
| 2500 | package-profile | package | Package size | append-plain-rows | MiniExcel | 5.08 ms | 0.68 ms | 0.39 ms | 2.49 | 3.27 | 19710.8 KB | 12.50 | 68.1 KB | 1.08 | 149.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | append-plain-rows | ClosedXML | 17.06 ms | 2.84 ms | 1.64 ms | 8.37 | 10.99 | 11197.4 KB | 7.10 | 59.8 KB | 0.95 | 737.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | append-plain-rows | EPPlus | 30.39 ms | 2.28 ms | 1.31 ms | 14.91 | 19.57 | 14365.5 KB | 9.11 | 56.9 KB | 0.90 | 1391.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | autofit-existing | OfficeIMO.Excel | 8.95 ms | 0.54 ms | 0.31 ms | 1.00 | 1.00 | 1895.5 KB | 1.00 | 143.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | autofit-existing | EPPlus | 85.46 ms | 6.88 ms | 3.97 ms | 9.55 | 9.55 | 50712.1 KB | 26.75 | 115.0 KB | 0.80 | 854.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | autofit-existing | ClosedXML | 142.40 ms | 9.45 ms | 5.45 ms | 15.91 | 15.91 | 84562.9 KB | 44.61 | 121.0 KB | 0.84 | 1490.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | large-shared-strings | OfficeIMO.Excel | 2.13 ms | 0.11 ms | 0.07 ms | 1.00 | 1.00 | 2440.3 KB | 1.00 | 55.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | large-shared-strings | MiniExcel | 4.75 ms | 0.49 ms | 0.28 ms | 2.23 | 2.23 | 21137.5 KB | 8.66 | 60.7 KB | 1.10 | 122.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | large-shared-strings | ClosedXML | 13.59 ms | 2.26 ms | 1.30 ms | 6.37 | 6.37 | 11299.2 KB | 4.63 | 50.3 KB | 0.91 | 536.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | large-shared-strings | EPPlus | 33.29 ms | 13.54 ms | 7.82 ms | 15.59 | 15.59 | 12804.8 KB | 5.25 | 48.1 KB | 0.87 | 1459.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-autofilter | OfficeIMO.Excel | 3.93 ms | 0.20 ms | 0.12 ms | 1.00 | 1.00 | 1340.4 KB | 1.00 | 142.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-autofilter | ClosedXML | 31.75 ms | 2.59 ms | 1.50 ms | 8.08 | 8.08 | 22226.8 KB | 16.58 | 120.2 KB | 0.84 | 708.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-autofilter | EPPlus | 45.17 ms | 5.88 ms | 3.40 ms | 11.50 | 11.50 | 24715.8 KB | 18.44 | 114.2 KB | 0.80 | 1049.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-charts | OfficeIMO.Excel | 5.01 ms | 0.24 ms | 0.14 ms | 1.00 | 1.00 | 1892.9 KB | 1.00 | 147.6 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-charts | EPPlus | 45.03 ms | 2.83 ms | 1.63 ms | 8.98 | 8.98 | 27142.7 KB | 14.34 | 117.0 KB | 0.79 | 798.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-conditional-formatting | OfficeIMO.Excel | 3.98 ms | 0.15 ms | 0.09 ms | 1.00 | 1.00 | 1405.8 KB | 1.00 | 142.7 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-conditional-formatting | ClosedXML | 32.06 ms | 1.71 ms | 0.99 ms | 8.06 | 8.06 | 22273.8 KB | 15.84 | 120.3 KB | 0.84 | 705.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-conditional-formatting | EPPlus | 41.71 ms | 1.44 ms | 0.83 ms | 10.48 | 10.48 | 24757.8 KB | 17.61 | 114.3 KB | 0.80 | 948.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-data-validation | OfficeIMO.Excel | 4.66 ms | 1.37 ms | 0.79 ms | 1.00 | 1.00 | 1356.1 KB | 1.00 | 142.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-data-validation | ClosedXML | 39.65 ms | 14.95 ms | 8.63 ms | 8.52 | 8.52 | 22247.9 KB | 16.41 | 120.3 KB | 0.84 | 751.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-data-validation | EPPlus | 43.46 ms | 1.96 ms | 1.13 ms | 9.33 | 9.33 | 24701.8 KB | 18.22 | 114.2 KB | 0.80 | 833.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-freeze-panes | OfficeIMO.Excel | 4.21 ms | 0.35 ms | 0.20 ms | 1.00 | 1.00 | 1342.8 KB | 1.00 | 142.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-freeze-panes | ClosedXML | 51.06 ms | 28.75 ms | 16.60 ms | 12.13 | 12.13 | 22222.0 KB | 16.55 | 120.2 KB | 0.84 | 1113.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-freeze-panes | EPPlus | 51.43 ms | 5.79 ms | 3.34 ms | 12.22 | 12.22 | 24730.3 KB | 18.42 | 114.3 KB | 0.80 | 1121.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-pivot-table | OfficeIMO.Excel | 40.43 ms | 25.40 ms | 14.67 ms | 1.00 | 1.00 | 15676.8 KB | 1.00 | 200.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-pivot-table | EPPlus | 75.85 ms | 46.04 ms | 26.58 ms | 1.88 | 1.88 | 29538.0 KB | 1.88 | 117.4 KB | 0.59 | 87.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-all-in-one | OfficeIMO.Excel | 18.79 ms | 0.62 ms | 0.36 ms | 1.00 | 1.00 | 16478.1 KB | 1.00 | 206.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-all-in-one | EPPlus | 109.93 ms | 49.86 ms | 28.79 ms | 5.85 | 5.85 | 54595.6 KB | 3.31 | 121.8 KB | 0.59 | 484.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-chart-first | OfficeIMO.Excel | 17.90 ms | 9.83 ms | 5.67 ms | 1.00 | 1.00 | 7452.6 KB | 1.00 | 206.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-chart-first | EPPlus | 92.37 ms | 10.89 ms | 6.29 ms | 5.16 | 5.16 | 54595.0 KB | 7.33 | 121.8 KB | 0.59 | 416.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-core | OfficeIMO.Excel | 6.40 ms | 1.61 ms | 0.93 ms | 1.00 | 1.00 | 1488.5 KB | 1.00 | 143.9 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-core | EPPlus | 95.81 ms | 21.13 ms | 12.20 ms | 14.98 | 14.98 | 47300.2 KB | 31.78 | 115.6 KB | 0.80 | 1398.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-core | ClosedXML | 125.19 ms | 49.92 ms | 28.82 ms | 19.57 | 19.57 | 69836.4 KB | 46.92 | 121.5 KB | 0.84 | 1857.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-extra-column | OfficeIMO.Excel | 24.24 ms | 5.88 ms | 3.39 ms | 1.00 | 1.00 | 17641.2 KB | 1.00 | 219.1 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-extra-column | EPPlus | 110.91 ms | 23.30 ms | 13.45 ms | 4.58 | 4.58 | 59227.3 KB | 3.36 | 128.4 KB | 0.59 | 357.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-no-autofit | OfficeIMO.Excel | 18.99 ms | 2.59 ms | 1.50 ms | 1.00 | 1.00 | 16466.7 KB | 1.00 | 206.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-no-autofit | EPPlus | 55.14 ms | 2.86 ms | 1.65 ms | 2.90 | 2.90 | 32907.5 KB | 2.00 | 121.8 KB | 0.59 | 190.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-post-mutation | OfficeIMO.Excel | 19.99 ms | 3.06 ms | 1.77 ms | 1.00 | 1.00 | 16487.6 KB | 1.00 | 206.6 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-post-mutation | EPPlus | 89.44 ms | 13.30 ms | 7.68 ms | 4.47 | 4.47 | 54595.5 KB | 3.31 | 121.9 KB | 0.59 | 347.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-shuffled-columns | OfficeIMO.Excel | 18.80 ms | 1.26 ms | 0.73 ms | 1.00 | 1.00 | 16485.8 KB | 1.00 | 211.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-shuffled-columns | EPPlus | 124.37 ms | 13.79 ms | 7.96 ms | 6.62 | 6.62 | 54592.2 KB | 3.31 | 124.3 KB | 0.59 | 561.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook | OfficeIMO.Excel | 24.52 ms | 0.58 ms | 0.33 ms | 1.00 | 1.00 | 20494.1 KB | 1.00 | 275.6 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | report-workbook | EPPlus | 97.75 ms | 1.81 ms | 1.04 ms | 3.99 | 3.99 | 77486.6 KB | 3.78 | 161.8 KB | 0.59 | 298.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook-core | OfficeIMO.Excel | 6.63 ms | 0.50 ms | 0.29 ms | 1.00 | 1.00 | 2711.1 KB | 1.00 | 187.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | report-workbook-core | EPPlus | 102.38 ms | 0.59 ms | 0.34 ms | 15.45 | 15.45 | 71970.9 KB | 26.55 | 157.2 KB | 0.84 | 1444.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook-core | ClosedXML | 108.22 ms | 4.92 ms | 2.84 ms | 16.33 | 16.33 | 97220.0 KB | 35.86 | 165.1 KB | 0.88 | 1532.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook-datatable | OfficeIMO.Excel | 22.31 ms | 0.34 ms | 0.20 ms | 1.00 | 1.00 | 20765.6 KB | 1.00 | 275.6 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | report-workbook-datatable | EPPlus | 108.48 ms | 6.70 ms | 3.87 ms | 4.86 | 4.86 | 65995.8 KB | 3.18 | 161.8 KB | 0.59 | 386.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook-datatable-core | OfficeIMO.Excel | 6.43 ms | 0.62 ms | 0.36 ms | 1.00 | 1.00 | 2982.7 KB | 1.00 | 187.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | report-workbook-datatable-core | EPPlus | 108.72 ms | 1.94 ms | 1.12 ms | 16.91 | 16.91 | 60480.4 KB | 20.28 | 157.2 KB | 0.84 | 1591.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook-datatable-core | ClosedXML | 136.50 ms | 26.69 ms | 15.41 ms | 21.23 | 21.23 | 82860.8 KB | 27.78 | 165.1 KB | 0.88 | 2023.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | LargeXlsx | 4.01 ms | 0.07 ms | 0.04 ms | 0.85 | 1.00 | 857.6 KB | 0.51 | 237.7 KB | 1.10 | 14.9% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | OfficeIMO.Excel | 4.71 ms | 0.28 ms | 0.16 ms | 1.00 | 1.17 | 1676.8 KB | 1.00 | 216.7 KB | 1.00 | Loss +17.5% |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | MiniExcel | 17.78 ms | 1.42 ms | 0.82 ms | 3.78 | 4.44 | 35919.4 KB | 21.42 | 235.3 KB | 1.09 | 277.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | ClosedXML | 93.21 ms | 3.94 ms | 2.27 ms | 19.80 | 23.26 | 71478.2 KB | 42.63 | 257.2 KB | 1.19 | 1879.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-bulk-report | OfficeIMO.Excel | 4.71 ms | 0.76 ms | 0.44 ms | 1.00 | 1.00 | 1401.7 KB | 1.00 | 143.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-bulk-report | MiniExcel | 9.16 ms | 2.17 ms | 1.26 ms | 1.94 | 1.94 | 26825.4 KB | 19.14 | 153.8 KB | 1.07 | 94.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-bulk-report | EPPlus | 74.44 ms | 1.46 ms | 0.84 ms | 15.80 | 15.80 | 47194.2 KB | 33.67 | 115.0 KB | 0.80 | 1480.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-bulk-report | ClosedXML | 74.63 ms | 4.45 ms | 2.57 ms | 15.84 | 15.84 | 58348.8 KB | 41.63 | 121.0 KB | 0.84 | 1484.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellformula | OfficeIMO.Excel | 2.60 ms | 0.06 ms | 0.03 ms | 1.00 | 1.00 | 1383.3 KB | 1.00 | 66.6 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellformula | ClosedXML | 22.39 ms | 1.51 ms | 0.87 ms | 8.62 | 8.62 | 12039.8 KB | 8.70 | 70.6 KB | 1.06 | 761.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellformula | EPPlus | 46.10 ms | 6.52 ms | 3.76 ms | 17.74 | 17.74 | 18110.8 KB | 13.09 | 62.1 KB | 0.93 | 1674.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-empty-strings | OfficeIMO.Excel | 1.98 ms | 0.13 ms | 0.08 ms | 1.00 | 1.00 | 1787.1 KB | 1.00 | 44.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-empty-strings | ClosedXML | 12.24 ms | 0.20 ms | 0.11 ms | 6.17 | 6.17 | 9959.5 KB | 5.57 | 44.9 KB | 1.02 | 517.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-empty-strings | EPPlus | 31.05 ms | 2.64 ms | 1.52 ms | 15.67 | 15.67 | 11773.4 KB | 6.59 | 42.0 KB | 0.95 | 1466.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-numbers | OfficeIMO.Excel | 2.19 ms | 0.12 ms | 0.07 ms | 1.00 | 1.00 | 1119.9 KB | 1.00 | 47.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-numbers | ClosedXML | 12.28 ms | 0.47 ms | 0.27 ms | 5.61 | 5.61 | 9177.1 KB | 8.19 | 45.9 KB | 0.98 | 461.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-numbers | EPPlus | 28.87 ms | 0.71 ms | 0.41 ms | 13.19 | 13.19 | 12895.6 KB | 11.51 | 43.7 KB | 0.93 | 1219.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-mixed | OfficeIMO.Excel | 2.58 ms | 0.11 ms | 0.06 ms | 1.00 | 1.00 | 1763.1 KB | 1.00 | 61.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-object-mixed | ClosedXML | 18.52 ms | 0.86 ms | 0.49 ms | 7.19 | 7.19 | 11887.0 KB | 6.74 | 59.5 KB | 0.97 | 618.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-mixed | EPPlus | 31.22 ms | 1.83 ms | 1.06 ms | 12.11 | 12.11 | 15643.7 KB | 8.87 | 58.9 KB | 0.96 | 1111.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse | OfficeIMO.Excel | 3.03 ms | 0.34 ms | 0.20 ms | 1.00 | 1.00 | 1506.7 KB | 1.00 | 62.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse | ClosedXML | 16.95 ms | 1.07 ms | 0.62 ms | 5.59 | 5.59 | 11296.3 KB | 7.50 | 52.5 KB | 0.85 | 458.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse | EPPlus | 31.93 ms | 1.19 ms | 0.69 ms | 10.53 | 10.53 | 14960.7 KB | 9.93 | 54.2 KB | 0.88 | 952.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 3.09 ms | 0.53 ms | 0.31 ms | 1.00 | 1.00 | 1506.8 KB | 1.00 | 62.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | ClosedXML | 17.17 ms | 1.21 ms | 0.70 ms | 5.55 | 5.55 | 11296.3 KB | 7.50 | 52.5 KB | 0.85 | 455.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | EPPlus | 32.60 ms | 2.24 ms | 1.30 ms | 10.54 | 10.54 | 14960.7 KB | 9.93 | 54.2 KB | 0.88 | 954.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-scalars | OfficeIMO.Excel | 1.92 ms | 0.09 ms | 0.05 ms | 1.00 | 1.00 | 1138.1 KB | 1.00 | 46.3 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-scalars | ClosedXML | 12.18 ms | 0.99 ms | 0.57 ms | 6.33 | 6.33 | 9021.2 KB | 7.93 | 45.4 KB | 0.98 | 532.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-scalars | EPPlus | 28.89 ms | 1.40 ms | 0.81 ms | 15.01 | 15.01 | 12827.9 KB | 11.27 | 42.4 KB | 0.91 | 1401.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings | OfficeIMO.Excel | 3.43 ms | 0.51 ms | 0.30 ms | 1.00 | 1.00 | 2617.0 KB | 1.00 | 55.1 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-strings | ClosedXML | 17.24 ms | 4.39 ms | 2.53 ms | 5.02 | 5.02 | 11299.2 KB | 4.32 | 50.3 KB | 0.91 | 402.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings | EPPlus | 39.97 ms | 6.92 ms | 4.00 ms | 11.64 | 11.64 | 12805.3 KB | 4.89 | 48.1 KB | 0.87 | 1064.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-distinct | OfficeIMO.Excel | 2.19 ms | 0.07 ms | 0.04 ms | 1.00 | 1.00 | 2379.2 KB | 1.00 | 51.8 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-distinct | ClosedXML | 17.77 ms | 0.88 ms | 0.51 ms | 8.10 | 8.10 | 13127.1 KB | 5.52 | 61.9 KB | 1.19 | 709.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-distinct | EPPlus | 32.79 ms | 6.21 ms | 3.58 ms | 14.95 | 14.95 | 13893.4 KB | 5.84 | 61.5 KB | 1.19 | 1394.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-repeated | OfficeIMO.Excel | 1.87 ms | 0.13 ms | 0.07 ms | 1.00 | 1.00 | 1579.8 KB | 1.00 | 40.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-repeated | ClosedXML | 12.02 ms | 0.24 ms | 0.14 ms | 6.42 | 6.42 | 9226.5 KB | 5.84 | 38.8 KB | 0.97 | 541.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-repeated | EPPlus | 29.12 ms | 6.57 ms | 3.79 ms | 15.55 | 15.55 | 11332.9 KB | 7.17 | 34.8 KB | 0.87 | 1454.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-temporal | OfficeIMO.Excel | 3.23 ms | 0.46 ms | 0.26 ms | 1.00 | 1.00 | 1435.5 KB | 1.00 | 63.3 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-temporal | ClosedXML | 17.84 ms | 1.50 ms | 0.86 ms | 5.52 | 5.52 | 9711.1 KB | 6.76 | 54.5 KB | 0.86 | 451.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-temporal | EPPlus | 33.32 ms | 5.32 ms | 3.07 ms | 10.30 | 10.30 | 14723.0 KB | 10.26 | 53.1 KB | 0.84 | 930.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 1.42 ms | 0.13 ms | 0.08 ms | 0.89 | 1.00 | 447.0 KB | 0.41 | 47.3 KB | 0.98 | 10.6% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 1.59 ms | 0.08 ms | 0.05 ms | 1.00 | 1.12 | 1092.0 KB | 1.00 | 48.2 KB | 1.00 | Loss +11.8% |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | ClosedXML | 17.41 ms | 7.39 ms | 4.26 ms | 10.93 | 12.22 | 10235.8 KB | 9.37 | 53.0 KB | 1.10 | 992.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | EPPlus | 28.68 ms | 5.38 ms | 3.11 ms | 18.00 | 20.13 | 13052.5 KB | 11.95 | 52.5 KB | 1.09 | 1700.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | LargeXlsx | 3.15 ms | 0.08 ms | 0.05 ms | 0.73 | 1.00 | 758.3 KB | 0.36 | 138.4 KB | 1.00 | 27.2% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 4.32 ms | 0.20 ms | 0.12 ms | 1.00 | 1.37 | 2081.1 KB | 1.00 | 138.0 KB | 1.00 | Loss +37.4% |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | MiniExcel | 9.36 ms | 0.96 ms | 0.55 ms | 2.17 | 2.98 | 23222.2 KB | 11.16 | 153.7 KB | 1.11 | 116.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | ClosedXML | 35.29 ms | 5.20 ms | 3.00 ms | 8.17 | 11.22 | 22221.3 KB | 10.68 | 120.1 KB | 0.87 | 716.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | EPPlus | 50.31 ms | 7.08 ms | 4.09 ms | 11.64 | 15.99 | 24694.3 KB | 11.87 | 114.1 KB | 0.83 | 1063.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | Sylvan.Data.Excel | 3.04 ms | 0.02 ms | 0.01 ms | 0.67 | 1.00 | 758.7 KB | 0.43 | 78.5 KB | 0.57 | 32.6% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | OfficeIMO.Excel | 4.51 ms | 0.37 ms | 0.21 ms | 1.00 | 1.48 | 1763.0 KB | 1.00 | 138.0 KB | 1.00 | Loss +48.4% |
| 2500 | package-profile | package | Package size | write-datareader-plain | LargeXlsx | 4.54 ms | 0.98 ms | 0.57 ms | 1.01 | 1.49 | 1032.5 KB | 0.59 | 138.4 KB | 1.00 | Tie vs OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | MiniExcel | 10.24 ms | 1.65 ms | 0.95 ms | 2.27 | 3.37 | 23043.8 KB | 13.07 | 153.6 KB | 1.11 | 127.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | ClosedXML | 29.12 ms | 1.09 ms | 0.63 ms | 6.46 | 9.58 | 11581.0 KB | 6.57 | 120.1 KB | 0.87 | 545.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | EPPlus | 47.50 ms | 6.75 ms | 3.90 ms | 10.53 | 15.63 | 16646.8 KB | 9.44 | 114.9 KB | 0.83 | 953.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table | OfficeIMO.Excel | 4.96 ms | 1.03 ms | 0.59 ms | 1.00 | 1.00 | 1774.9 KB | 1.00 | 138.8 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-datareader-table | MiniExcel | 9.06 ms | 0.96 ms | 0.55 ms | 1.83 | 1.83 | 23044.1 KB | 12.98 | 153.6 KB | 1.11 | 82.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table | ClosedXML | 37.73 ms | 1.50 ms | 0.87 ms | 7.60 | 7.60 | 19007.9 KB | 10.71 | 120.9 KB | 0.87 | 660.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table | EPPlus | 45.48 ms | 3.74 ms | 2.16 ms | 9.17 | 9.17 | 16646.5 KB | 9.38 | 114.9 KB | 0.83 | 816.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | OfficeIMO.Excel | 4.81 ms | 0.36 ms | 0.21 ms | 1.00 | 1.00 | 1781.2 KB | 1.00 | 138.9 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | MiniExcel | 8.70 ms | 0.72 ms | 0.41 ms | 1.81 | 1.81 | 26647.2 KB | 14.96 | 153.8 KB | 1.11 | 80.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | EPPlus | 58.63 ms | 0.22 ms | 0.13 ms | 12.19 | 12.19 | 38344.0 KB | 21.53 | 115.1 KB | 0.83 | 1119.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | ClosedXML | 80.64 ms | 10.32 ms | 5.96 ms | 16.77 | 16.77 | 58361.4 KB | 32.77 | 121.0 KB | 0.87 | 1576.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | LargeXlsx | 4.35 ms | 0.83 ms | 0.48 ms | 0.96 | 1.00 | 1123.9 KB | 0.53 | 164.2 KB | 1.25 | 3.9% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | OfficeIMO.Excel | 4.53 ms | 0.73 ms | 0.42 ms | 1.00 | 1.04 | 2140.6 KB | 1.00 | 131.1 KB | 1.00 | Loss +4.0% |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | MiniExcel | 12.93 ms | 5.22 ms | 3.02 ms | 2.86 | 2.97 | 29746.9 KB | 13.90 | 180.5 KB | 1.38 | 185.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | ClosedXML | 53.63 ms | 2.87 ms | 1.66 ms | 11.84 | 12.32 | 27410.3 KB | 12.80 | 159.4 KB | 1.22 | 1084.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | EPPlus | 57.77 ms | 8.12 ms | 4.69 ms | 12.76 | 13.27 | 21890.1 KB | 10.23 | 144.5 KB | 1.10 | 1175.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | OfficeIMO.Excel | 4.57 ms | 0.16 ms | 0.09 ms | 1.00 | 1.00 | 2880.2 KB | 1.00 | 176.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | MiniExcel | 9.87 ms | 0.01 ms | 0.01 ms | 2.16 | 2.16 | 29746.9 KB | 10.33 | 180.5 KB | 1.03 | 116.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | ClosedXML | 53.58 ms | 4.70 ms | 2.71 ms | 11.73 | 11.73 | 27409.3 KB | 9.52 | 159.4 KB | 0.91 | 1073.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | EPPlus | 56.04 ms | 3.64 ms | 2.10 ms | 12.27 | 12.27 | 21890.1 KB | 7.60 | 144.5 KB | 0.82 | 1127.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables | OfficeIMO.Excel | 3.97 ms | 0.13 ms | 0.07 ms | 1.00 | 1.00 | 2066.1 KB | 1.00 | 138.9 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-dataset-tables | MiniExcel | 9.22 ms | 1.20 ms | 0.69 ms | 2.32 | 2.32 | 28700.4 KB | 13.89 | 156.4 KB | 1.13 | 132.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables | ClosedXML | 35.17 ms | 1.66 ms | 0.96 ms | 8.87 | 8.87 | 18876.9 KB | 9.14 | 123.4 KB | 0.89 | 786.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables | EPPlus | 40.60 ms | 5.17 ms | 2.98 ms | 10.23 | 10.23 | 18701.1 KB | 9.05 | 116.6 KB | 0.84 | 923.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | OfficeIMO.Excel | 4.42 ms | 0.30 ms | 0.17 ms | 1.00 | 1.00 | 2078.7 KB | 1.00 | 139.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | MiniExcel | 8.60 ms | 0.05 ms | 0.03 ms | 1.95 | 1.95 | 31798.5 KB | 15.30 | 156.6 KB | 1.13 | 94.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | EPPlus | 60.89 ms | 0.63 ms | 0.36 ms | 13.78 | 13.78 | 41456.2 KB | 19.94 | 116.9 KB | 0.84 | 1277.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | ClosedXML | 77.37 ms | 6.17 ms | 3.56 ms | 17.51 | 17.51 | 56708.2 KB | 27.28 | 123.7 KB | 0.89 | 1650.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-direct | OfficeIMO.Excel | 4.44 ms | 0.79 ms | 0.46 ms | 1.00 | 1.00 | 1748.6 KB | 1.00 | 138.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-datatable-direct | LargeXlsx | 4.70 ms | 1.06 ms | 0.61 ms | 1.06 | 1.06 | 1149.0 KB | 0.66 | 138.4 KB | 1.00 | 6.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-direct | MiniExcel | 9.60 ms | 0.66 ms | 0.38 ms | 2.16 | 2.16 | 23062.5 KB | 13.19 | 153.7 KB | 1.11 | 116.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-direct | ClosedXML | 30.59 ms | 1.68 ms | 0.97 ms | 6.90 | 6.90 | 11581.0 KB | 6.62 | 120.1 KB | 0.87 | 589.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-direct | EPPlus | 52.17 ms | 8.82 ms | 5.09 ms | 11.76 | 11.76 | 16646.5 KB | 9.52 | 114.9 KB | 0.83 | 1076.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | OfficeIMO.Excel | 3.95 ms | 0.07 ms | 0.04 ms | 1.00 | 1.00 | 1760.7 KB | 1.00 | 138.8 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | MiniExcel | 12.17 ms | 6.38 ms | 3.69 ms | 3.08 | 3.08 | 23062.8 KB | 13.10 | 153.7 KB | 1.11 | 208.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | ClosedXML | 38.19 ms | 2.54 ms | 1.46 ms | 9.68 | 9.68 | 19008.3 KB | 10.80 | 120.9 KB | 0.87 | 867.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | EPPlus | 49.41 ms | 4.69 ms | 2.71 ms | 12.52 | 12.52 | 16646.5 KB | 9.45 | 114.9 KB | 0.83 | 1151.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | LargeXlsx | 3.27 ms | 0.06 ms | 0.03 ms | 0.81 | 1.00 | 758.3 KB | 0.43 | 138.4 KB | 1.00 | 19.3% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 4.06 ms | 0.29 ms | 0.16 ms | 1.00 | 1.24 | 1769.2 KB | 1.00 | 138.0 KB | 1.00 | Loss +23.9% |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | MiniExcel | 8.66 ms | 0.23 ms | 0.13 ms | 2.13 | 2.64 | 23222.2 KB | 13.13 | 153.7 KB | 1.11 | 113.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | ClosedXML | 29.25 ms | 0.95 ms | 0.55 ms | 7.21 | 8.94 | 11581.0 KB | 6.55 | 120.1 KB | 0.87 | 621.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | EPPlus | 49.20 ms | 6.05 ms | 3.50 ms | 12.13 | 15.03 | 16646.8 KB | 9.41 | 114.9 KB | 0.83 | 1113.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 5.08 ms | 1.27 ms | 0.73 ms | 1.00 | 1.00 | 1347.1 KB | 1.00 | 142.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 74.15 ms | 8.33 ms | 4.81 ms | 14.59 | 14.59 | 38344.3 KB | 28.46 | 115.1 KB | 0.81 | 1358.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 78.42 ms | 0.89 ms | 0.52 ms | 15.43 | 15.43 | 50927.5 KB | 37.80 | 120.2 KB | 0.84 | 1442.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | OfficeIMO.Excel | 4.15 ms | 0.43 ms | 0.25 ms | 1.00 | 1.00 | 1339.3 KB | 1.00 | 142.3 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | LargeXlsx | 4.39 ms | 1.67 ms | 0.96 ms | 1.06 | 1.06 | 758.3 KB | 0.57 | 138.4 KB | 0.97 | 5.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | MiniExcel | 9.07 ms | 0.32 ms | 0.18 ms | 2.18 | 2.18 | 23222.2 KB | 17.34 | 153.7 KB | 1.08 | 118.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | ClosedXML | 37.29 ms | 9.59 ms | 5.54 ms | 8.98 | 8.98 | 11581.0 KB | 8.65 | 120.1 KB | 0.84 | 797.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | EPPlus | 50.65 ms | 3.10 ms | 1.79 ms | 12.19 | 12.19 | 16646.5 KB | 12.43 | 114.9 KB | 0.81 | 1119.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 5.56 ms | 0.87 ms | 0.50 ms | 1.00 | 1.00 | 1505.3 KB | 1.00 | 138.1 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 62.64 ms | 3.86 ms | 2.23 ms | 11.26 | 11.26 | 38344.3 KB | 25.47 | 115.1 KB | 0.83 | 1025.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 68.76 ms | 2.41 ms | 1.39 ms | 12.36 | 12.36 | 50927.5 KB | 33.83 | 120.2 KB | 0.87 | 1136.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 4.03 ms | 0.66 ms | 0.38 ms | 0.79 | 1.00 | 758.3 KB | 0.51 | 138.4 KB | 1.00 | 20.5% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 5.07 ms | 0.43 ms | 0.25 ms | 1.00 | 1.26 | 1497.5 KB | 1.00 | 138.0 KB | 1.00 | Loss +25.8% |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | MiniExcel | 10.09 ms | 0.87 ms | 0.50 ms | 1.99 | 2.50 | 23222.2 KB | 15.51 | 153.7 KB | 1.11 | 98.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | ClosedXML | 33.31 ms | 5.84 ms | 3.37 ms | 6.56 | 8.26 | 11581.0 KB | 7.73 | 120.1 KB | 0.87 | 556.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | EPPlus | 48.27 ms | 8.46 ms | 4.89 ms | 9.51 | 11.97 | 16646.5 KB | 11.12 | 114.9 KB | 0.83 | 851.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | LargeXlsx | 3.36 ms | 0.19 ms | 0.11 ms | 0.66 | 1.00 | 758.3 KB | 0.43 | 138.4 KB | 0.97 | 34.2% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | OfficeIMO.Excel | 5.12 ms | 0.15 ms | 0.09 ms | 1.00 | 1.52 | 1770.1 KB | 1.00 | 142.3 KB | 1.00 | Loss +52.0% |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | MiniExcel | 10.49 ms | 1.68 ms | 0.97 ms | 2.05 | 3.12 | 23222.2 KB | 13.12 | 153.7 KB | 1.08 | 105.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | ClosedXML | 31.38 ms | 2.18 ms | 1.26 ms | 6.13 | 9.33 | 11581.0 KB | 6.54 | 120.1 KB | 0.84 | 513.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | EPPlus | 44.74 ms | 6.78 ms | 3.92 ms | 8.75 | 13.29 | 16646.5 KB | 9.40 | 114.9 KB | 0.81 | 774.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.66 ms | 0.44 ms | 0.25 ms | 1.00 | 1.00 | 1346.4 KB | 1.00 | 142.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 56.10 ms | 1.52 ms | 0.88 ms | 12.04 | 12.04 | 28540.6 KB | 21.20 | 120.2 KB | 0.84 | 1103.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 62.00 ms | 2.13 ms | 1.23 ms | 13.31 | 13.31 | 27306.2 KB | 20.28 | 115.0 KB | 0.81 | 1230.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 5.92 ms | 0.78 ms | 0.45 ms | 1.00 | 1.00 | 2341.7 KB | 1.00 | 183.1 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | LargeXlsx | 6.24 ms | 3.21 ms | 1.85 ms | 1.05 | 1.05 | 802.5 KB | 0.34 | 182.6 KB | 1.00 | 5.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | MiniExcel | 9.50 ms | 0.56 ms | 0.32 ms | 1.60 | 1.60 | 25190.5 KB | 10.76 | 194.0 KB | 1.06 | 60.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | ClosedXML | 39.28 ms | 3.25 ms | 1.88 ms | 6.63 | 6.63 | 16973.5 KB | 7.25 | 161.0 KB | 0.88 | 563.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | EPPlus | 56.36 ms | 7.15 ms | 4.13 ms | 9.51 | 9.51 | 20105.6 KB | 8.59 | 152.1 KB | 0.83 | 851.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | LargeXlsx | 4.41 ms | 0.25 ms | 0.14 ms | 0.88 | 1.00 | 802.5 KB | 0.53 | 182.6 KB | 1.00 | 11.6% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 4.99 ms | 0.53 ms | 0.30 ms | 1.00 | 1.13 | 1507.7 KB | 1.00 | 182.4 KB | 1.00 | Loss +13.1% |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | MiniExcel | 9.44 ms | 0.60 ms | 0.35 ms | 1.89 | 2.14 | 25190.5 KB | 16.71 | 194.0 KB | 1.06 | 89.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | ClosedXML | 37.90 ms | 2.92 ms | 1.69 ms | 7.59 | 8.59 | 16973.5 KB | 11.26 | 161.0 KB | 0.88 | 659.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | EPPlus | 54.63 ms | 6.16 ms | 3.56 ms | 10.94 | 12.38 | 20105.6 KB | 13.33 | 152.1 KB | 0.83 | 994.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 22.03 ms | 0.48 ms | 0.28 ms | 1.00 | 1.00 | 4502.3 KB | 1.00 | 651.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | LargeXlsx | 25.70 ms | 6.44 ms | 3.72 ms | 1.17 | 1.17 | 2810.7 KB | 0.62 | 644.6 KB | 0.99 | 16.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | MiniExcel | 46.09 ms | 3.57 ms | 2.06 ms | 2.09 | 2.09 | 48414.8 KB | 10.75 | 674.4 KB | 1.04 | 109.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | ClosedXML | 141.65 ms | 6.48 ms | 3.74 ms | 6.43 | 6.43 | 51647.0 KB | 11.47 | 615.5 KB | 0.95 | 543.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | EPPlus | 197.30 ms | 7.19 ms | 4.15 ms | 8.96 | 8.96 | 69140.0 KB | 15.36 | 548.9 KB | 0.84 | 795.8% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | OfficeIMO.Excel | 9.58 ms | 1.29 ms | 0.74 ms | 1.00 | 1.00 | 1895.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | EPPlus 4.5.3.3 | 96.76 ms |  |  | 10.10 | 10.10 |  |  |  |  | 909.8% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | EPPlus | 153.26 ms | 48.90 ms | 28.23 ms | 15.99 | 15.99 | 50712.1 KB | 26.76 |  |  | 1499.4% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | ClosedXML | 246.28 ms | 77.70 ms | 44.86 ms | 25.70 | 25.70 | 84579.6 KB | 44.63 |  |  | 2470.1% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | OfficeIMO.Excel | 15.80 ms | 0.23 ms | 0.13 ms | 1.00 | 1.00 | 16466.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus 4.5.3.3 | 41.12 ms |  |  | 2.60 | 2.60 |  |  |  |  | 160.2% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus | 52.21 ms | 9.36 ms | 5.40 ms | 3.30 | 3.30 | 32907.5 KB | 2.00 |  |  | 230.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | OfficeIMO.Excel | 12.31 ms | 0.75 ms | 0.43 ms | 1.00 | 1.00 | 7452.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus 4.5.3.3 | 71.50 ms |  |  | 5.81 | 5.81 |  |  |  |  | 480.8% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus | 72.95 ms | 4.70 ms | 2.71 ms | 5.93 | 5.93 | 54594.9 KB | 7.33 |  |  | 492.6% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | OfficeIMO.Excel | 18.17 ms | 2.02 ms | 1.17 ms | 1.00 | 1.00 | 17641.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus | 75.44 ms | 1.49 ms | 0.86 ms | 4.15 | 4.15 | 59227.2 KB | 3.36 |  |  | 315.2% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus 4.5.3.3 | 83.91 ms |  |  | 4.62 | 4.62 |  |  |  |  | 361.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | OfficeIMO.Excel | 16.10 ms | 0.62 ms | 0.36 ms | 1.00 | 1.00 | 16487.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus 4.5.3.3 | 69.99 ms |  |  | 4.35 | 4.35 |  |  |  |  | 334.8% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus | 77.29 ms | 6.98 ms | 4.03 ms | 4.80 | 4.80 | 54595.6 KB | 3.31 |  |  | 380.1% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | OfficeIMO.Excel | 22.37 ms | 7.69 ms | 4.44 ms | 1.00 | 1.00 | 16482.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus 4.5.3.3 | 70.10 ms |  |  | 3.13 | 3.13 |  |  |  |  | 213.4% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus | 70.13 ms | 1.50 ms | 0.86 ms | 3.13 | 3.13 | 54592.2 KB | 3.31 |  |  | 213.5% slower than OfficeIMO |
| 2500 | speed-comparison | other | Object projection | build-object-datatable-dictionaries | OfficeIMO.Excel | 1.58 ms | 0.28 ms | 0.16 ms | 1.00 | 1.00 | 564.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Object projection | build-object-datatable-typed | OfficeIMO.Excel | 1.24 ms | 0.05 ms | 0.03 ms | 1.00 | 1.00 | 856.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-cells | OfficeIMO.Excel | 6.05 ms | 0.59 ms | 0.34 ms | 1.00 | 1.00 | 2531.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-cells | ClosedXML | 32.07 ms | 3.09 ms | 1.78 ms | 5.30 | 5.30 | 17022.5 KB | 6.72 |  |  | 430.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-cells | EPPlus | 34.27 ms | 4.22 ms | 2.44 ms | 5.67 | 5.67 | 20155.0 KB | 7.96 |  |  | 466.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | OfficeIMO.Excel | 3.93 ms | 0.53 ms | 0.31 ms | 1.00 | 1.00 | 526.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | EPPlus | 23.08 ms | 0.76 ms | 0.44 ms | 5.88 | 5.88 | 13108.2 KB | 24.91 |  |  | 487.8% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | ClosedXML | 31.65 ms | 3.18 ms | 1.84 ms | 8.06 | 8.06 | 15458.2 KB | 29.38 |  |  | 705.8% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-range | OfficeIMO.Excel | 5.77 ms | 0.95 ms | 0.55 ms | 1.00 | 1.00 | 2531.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-range | EPPlus | 34.80 ms | 12.04 ms | 6.95 ms | 6.03 | 6.03 | 20155.0 KB | 7.96 |  |  | 502.8% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-range | ClosedXML | 36.64 ms | 6.13 ms | 3.54 ms | 6.35 | 6.35 | 17021.0 KB | 6.72 |  |  | 534.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-top-range | OfficeIMO.Excel | 0.74 ms | 0.07 ms | 0.04 ms | 1.00 | 1.00 | 285.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-top-range | EPPlus | 26.30 ms | 1.30 ms | 0.75 ms | 35.75 | 35.75 | 12404.5 KB | 43.45 |  |  | 3474.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-top-range | ClosedXML | 38.74 ms | 10.02 ms | 5.79 ms | 52.66 | 52.66 | 15370.8 KB | 53.85 |  |  | 5165.6% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-autofilter | OfficeIMO.Excel | 3.76 ms | 0.36 ms | 0.21 ms | 1.00 | 1.00 | 1340.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-autofilter | ClosedXML | 30.06 ms | 0.87 ms | 0.51 ms | 7.99 | 7.99 | 22226.8 KB | 16.58 |  |  | 698.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-autofilter | EPPlus 4.5.3.3 | 31.90 ms |  |  | 8.48 | 8.48 |  |  |  |  | 747.5% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-autofilter | EPPlus | 39.23 ms | 0.51 ms | 0.30 ms | 10.42 | 10.42 | 24715.8 KB | 18.44 |  |  | 942.2% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-charts | OfficeIMO.Excel | 5.08 ms | 0.28 ms | 0.16 ms | 1.00 | 1.00 | 1893.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-charts | EPPlus 4.5.3.3 | 34.43 ms |  |  | 6.78 | 6.78 |  |  |  |  | 578.4% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-charts | EPPlus | 44.67 ms | 3.42 ms | 1.98 ms | 8.80 | 8.80 | 27142.7 KB | 14.34 |  |  | 780.2% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | OfficeIMO.Excel | 3.99 ms | 0.79 ms | 0.46 ms | 1.00 | 1.00 | 1405.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | ClosedXML | 31.01 ms | 3.52 ms | 2.03 ms | 7.77 | 7.77 | 22273.8 KB | 15.84 |  |  | 676.8% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | EPPlus 4.5.3.3 | 35.13 ms |  |  | 8.80 | 8.80 |  |  |  |  | 779.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | EPPlus | 40.54 ms | 0.76 ms | 0.44 ms | 10.15 | 10.15 | 24757.8 KB | 17.61 |  |  | 915.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-data-validation | OfficeIMO.Excel | 3.79 ms | 0.05 ms | 0.03 ms | 1.00 | 1.00 | 1356.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-data-validation | EPPlus 4.5.3.3 | 31.43 ms |  |  | 8.30 | 8.30 |  |  |  |  | 729.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-data-validation | ClosedXML | 31.48 ms | 1.31 ms | 0.76 ms | 8.31 | 8.31 | 22247.9 KB | 16.41 |  |  | 731.2% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-data-validation | EPPlus | 47.86 ms | 6.16 ms | 3.56 ms | 12.64 | 12.64 | 24701.8 KB | 18.22 |  |  | 1163.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | OfficeIMO.Excel | 3.54 ms | 0.04 ms | 0.02 ms | 1.00 | 1.00 | 1342.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | ClosedXML | 30.17 ms | 0.97 ms | 0.56 ms | 8.52 | 8.52 | 22222.0 KB | 16.55 |  |  | 752.0% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | EPPlus 4.5.3.3 | 32.39 ms |  |  | 9.15 | 9.15 |  |  |  |  | 814.5% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | EPPlus | 46.01 ms | 4.84 ms | 2.80 ms | 12.99 | 12.99 | 24730.4 KB | 18.42 |  |  | 1199.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | OfficeIMO.Excel | 15.54 ms | 2.91 ms | 1.68 ms | 1.00 | 1.00 | 15676.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | EPPlus 4.5.3.3 | 35.76 ms |  |  | 2.30 | 2.30 |  |  |  |  | 130.2% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | EPPlus | 50.67 ms | 5.85 ms | 3.38 ms | 3.26 | 3.26 | 29537.9 KB | 1.88 |  |  | 226.1% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world report | realworld-report-all-in-one | OfficeIMO.Excel | 18.80 ms | 2.70 ms | 1.56 ms | 1.00 | 1.00 | 16477.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus | 78.40 ms | 6.37 ms | 3.68 ms | 4.17 | 4.17 | 54595.5 KB | 3.31 |  |  | 317.0% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus 4.5.3.3 | 81.58 ms |  |  | 4.34 | 4.34 |  |  |  |  | 333.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world report | realworld-report-core | OfficeIMO.Excel | 4.45 ms | 0.16 ms | 0.09 ms | 1.00 | 1.00 | 1488.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world report | realworld-report-core | EPPlus | 66.04 ms | 0.38 ms | 0.22 ms | 14.83 | 14.83 | 47300.2 KB | 31.77 |  |  | 1383.1% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world report | realworld-report-core | EPPlus 4.5.3.3 | 67.08 ms |  |  | 15.06 | 15.06 |  |  |  |  | 1406.4% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world report | realworld-report-core | ClosedXML | 87.76 ms | 4.86 ms | 2.81 ms | 19.71 | 19.71 | 69834.2 KB | 46.91 |  |  | 1870.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook | OfficeIMO.Excel | 31.04 ms | 4.10 ms | 2.36 ms | 1.00 | 1.00 | 20450.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Report workbook | report-workbook | EPPlus 4.5.3.3 | 95.11 ms |  |  | 3.06 | 3.06 |  |  |  |  | 206.5% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook | EPPlus | 118.61 ms | 23.08 ms | 13.32 ms | 3.82 | 3.82 | 77486.6 KB | 3.79 |  |  | 282.2% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-core | OfficeIMO.Excel | 7.25 ms | 0.62 ms | 0.36 ms | 1.00 | 1.00 | 2711.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Report workbook | report-workbook-core | EPPlus 4.5.3.3 | 82.84 ms |  |  | 11.42 | 11.42 |  |  |  |  | 1041.8% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-core | EPPlus | 117.98 ms | 15.62 ms | 9.02 ms | 16.26 | 16.26 | 71970.9 KB | 26.55 |  |  | 1526.1% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-core | ClosedXML | 119.32 ms | 10.58 ms | 6.11 ms | 16.45 | 16.45 | 97220.1 KB | 35.86 |  |  | 1544.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable | OfficeIMO.Excel | 27.85 ms | 2.69 ms | 1.55 ms | 1.00 | 1.00 | 20765.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable | EPPlus 4.5.3.3 | 84.67 ms |  |  | 3.04 | 3.04 |  |  |  |  | 204.0% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable | EPPlus | 109.46 ms | 7.29 ms | 4.21 ms | 3.93 | 3.93 | 65995.8 KB | 3.18 |  |  | 293.1% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable-core | OfficeIMO.Excel | 6.52 ms | 0.95 ms | 0.55 ms | 1.00 | 1.00 | 2982.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable-core | EPPlus 4.5.3.3 | 84.15 ms |  |  | 12.91 | 12.91 |  |  |  |  | 1191.0% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable-core | EPPlus | 100.44 ms | 0.95 ms | 0.55 ms | 15.41 | 15.41 | 60480.4 KB | 20.28 |  |  | 1440.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable-core | ClosedXML | 103.43 ms | 7.09 ms | 4.09 ms | 15.87 | 15.87 | 82858.9 KB | 27.78 |  |  | 1486.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | OfficeIMO.Excel | 2.66 ms | 0.62 ms | 0.36 ms | 1.00 | 1.00 | 709.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | EPPlus 4.5.3.3 | 13.45 ms |  |  | 5.06 | 5.06 |  |  |  |  | 405.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | ClosedXML | 16.37 ms | 1.08 ms | 0.62 ms | 6.16 | 6.16 | 8274.0 KB | 11.66 |  |  | 515.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | EPPlus | 22.42 ms | 6.29 ms | 3.63 ms | 8.43 | 8.43 | 7708.1 KB | 10.87 |  |  | 743.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | OfficeIMO.Excel | 1.20 ms | 0.36 ms | 0.21 ms | 1.00 | 1.00 | 177.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | Sylvan.Data.Excel | 1.47 ms | 0.28 ms | 0.16 ms | 1.22 | 1.22 | 316.6 KB | 1.78 |  |  | 22.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | ExcelDataReader | 2.40 ms | 0.37 ms | 0.21 ms | 2.00 | 2.00 | 4062.2 KB | 22.90 |  |  | 99.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | ClosedXML | 5.24 ms | 1.85 ms | 1.07 ms | 4.35 | 4.35 | 4392.6 KB | 24.76 |  |  | 335.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | EPPlus 4.5.3.3 | 10.55 ms |  |  | 8.77 | 8.77 |  |  |  |  | 776.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | MiniExcel | 16.38 ms | 2.50 ms | 1.44 ms | 13.61 | 13.61 | 46194.9 KB | 260.36 |  |  | 1261.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | EPPlus | 43.15 ms | 6.06 ms | 3.50 ms | 35.87 | 35.87 | 43071.1 KB | 242.76 |  |  | 3487.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | Sylvan.Data.Excel | 1.73 ms | 0.20 ms | 0.12 ms | 0.66 | 1.00 | 316.6 KB | 1.78 |  |  | 34.4% faster than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | OfficeIMO.Excel | 2.64 ms | 0.50 ms | 0.29 ms | 1.00 | 1.52 | 177.5 KB | 1.00 |  |  | Loss +52.4% |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | ExcelDataReader | 2.88 ms | 0.27 ms | 0.16 ms | 1.09 | 1.66 | 4062.2 KB | 22.89 |  |  | 9.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | EPPlus 4.5.3.3 | 14.09 ms |  |  | 5.34 | 8.14 |  |  |  |  | 434.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | MiniExcel | 18.39 ms | 2.44 ms | 1.41 ms | 6.98 | 10.63 | 46194.9 KB | 260.29 |  |  | 597.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | EPPlus | 78.22 ms | 30.89 ms | 17.83 ms | 29.67 | 45.22 | 43071.1 KB | 242.69 |  |  | 2867.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | ClosedXML | 130.35 ms | 201.60 ms | 116.39 ms | 49.44 | 75.35 | 4392.6 KB | 24.75 |  |  | 4844.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | Sylvan.Data.Excel | 2.35 ms | 0.69 ms | 0.40 ms | 0.78 | 1.00 | 518.6 KB | 0.49 |  |  | 22.2% faster than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | OfficeIMO.Excel | 3.02 ms | 0.45 ms | 0.26 ms | 1.00 | 1.29 | 1056.7 KB | 1.00 |  |  | Loss +28.5% |
| 2500 | speed-comparison | read | Other | shared-string-read | ExcelDataReader | 5.47 ms | 1.88 ms | 1.09 ms | 1.81 | 2.33 | 2619.1 KB | 2.48 |  |  | 81.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | MiniExcel | 8.74 ms | 3.57 ms | 2.06 ms | 2.89 | 3.72 | 7530.1 KB | 7.13 |  |  | 189.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | EPPlus 4.5.3.3 | 15.48 ms |  |  | 5.13 | 6.59 |  |  |  |  | 412.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | ClosedXML | 15.49 ms | 1.35 ms | 0.78 ms | 5.13 | 6.59 | 9497.7 KB | 8.99 |  |  | 412.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | EPPlus | 20.58 ms | 2.86 ms | 1.65 ms | 6.81 | 8.76 | 10372.3 KB | 9.82 |  |  | 581.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | OfficeIMO.Excel | 4.63 ms | 1.16 ms | 0.67 ms | 1.00 | 1.00 | 374.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | Sylvan.Data.Excel | 5.01 ms | 1.29 ms | 0.74 ms | 1.08 | 1.08 | 655.2 KB | 1.75 |  |  | 8.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | ExcelDataReader | 10.92 ms | 0.85 ms | 0.49 ms | 2.36 | 2.36 | 6089.5 KB | 16.26 |  |  | 136.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | MiniExcel | 16.73 ms | 2.66 ms | 1.53 ms | 3.62 | 3.62 | 18661.8 KB | 49.83 |  |  | 261.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | EPPlus | 25.85 ms | 0.82 ms | 0.47 ms | 5.59 | 5.59 | 12427.1 KB | 33.18 |  |  | 458.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | ClosedXML | 32.55 ms | 2.88 ms | 1.67 ms | 7.04 | 7.04 | 15361.3 KB | 41.02 |  |  | 603.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | Sylvan.Data.Excel | 6.39 ms | 0.24 ms | 0.14 ms | 0.80 | 1.00 | 2239.3 KB | 0.62 |  |  | 19.8% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | OfficeIMO.Excel | 7.97 ms | 2.65 ms | 1.53 ms | 1.00 | 1.25 | 3594.5 KB | 1.00 |  |  | Loss +24.7% |
| 2500 | speed-comparison | read | Range and table read | read-datatable | MiniExcel | 13.58 ms | 0.80 ms | 0.46 ms | 1.70 | 2.12 | 18266.6 KB | 5.08 |  |  | 70.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | ExcelDataReader | 13.74 ms | 2.47 ms | 1.43 ms | 1.72 | 2.15 | 7673.5 KB | 2.13 |  |  | 72.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | EPPlus | 34.54 ms | 4.96 ms | 2.86 ms | 4.33 | 5.41 | 21736.6 KB | 6.05 |  |  | 333.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | ClosedXML | 34.64 ms | 3.52 ms | 2.03 ms | 4.35 | 5.42 | 18314.0 KB | 5.10 |  |  | 334.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | EPPlus 4.5.3.3 | 37.35 ms |  |  | 4.69 | 5.84 |  |  |  |  | 368.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | Sylvan.Data.Excel | 4.62 ms | 0.31 ms | 0.18 ms | 0.98 | 1.00 | 733.5 KB | 1.33 |  |  | Tie vs OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | OfficeIMO.Excel | 4.69 ms | 1.10 ms | 0.64 ms | 1.00 | 1.02 | 551.0 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | MiniExcel | 10.02 ms | 0.82 ms | 0.47 ms | 2.14 | 2.17 | 15847.6 KB | 28.76 |  |  | 113.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | ExcelDataReader | 10.20 ms | 0.09 ms | 0.05 ms | 2.17 | 2.21 | 6089.5 KB | 11.05 |  |  | 117.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | EPPlus | 25.44 ms | 1.63 ms | 0.94 ms | 5.42 | 5.51 | 13108.2 KB | 23.79 |  |  | 442.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | ClosedXML | 40.08 ms | 2.38 ms | 1.37 ms | 8.54 | 8.68 | 15459.9 KB | 28.06 |  |  | 754.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | OfficeIMO.Excel | 9.55 ms | 0.49 ms | 0.29 ms | 1.00 | 1.00 | 2692.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Range and table read | read-range | Sylvan.Data.Excel | 12.29 ms | 5.49 ms | 3.17 ms | 1.29 | 1.29 | 655.0 KB | 0.24 |  |  | 28.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | MiniExcel | 20.15 ms | 2.90 ms | 1.67 ms | 2.11 | 2.11 | 18662.2 KB | 6.93 |  |  | 111.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | ExcelDataReader | 21.67 ms | 9.81 ms | 5.66 ms | 2.27 | 2.27 | 6089.2 KB | 2.26 |  |  | 126.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | EPPlus 4.5.3.3 | 34.03 ms |  |  | 3.56 | 3.56 |  |  |  |  | 256.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | EPPlus | 40.39 ms | 11.30 ms | 6.53 ms | 4.23 | 4.23 | 20152.6 KB | 7.48 |  |  | 323.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | ClosedXML | 79.44 ms | 22.41 ms | 12.94 ms | 8.32 | 8.32 | 16846.4 KB | 6.26 |  |  | 732.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | OfficeIMO.Excel | 6.14 ms | 1.86 ms | 1.07 ms | 1.00 | 1.00 | 2751.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | Sylvan.Data.Excel | 6.26 ms | 0.54 ms | 0.31 ms | 1.02 | 1.02 | 750.3 KB | 0.27 |  |  | 2.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | ExcelDataReader | 13.25 ms | 2.45 ms | 1.41 ms | 2.16 | 2.16 | 6089.5 KB | 2.21 |  |  | 115.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | MiniExcel | 14.05 ms | 1.47 ms | 0.85 ms | 2.29 | 2.29 | 18662.4 KB | 6.78 |  |  | 128.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | ClosedXML | 30.94 ms | 1.82 ms | 1.05 ms | 5.04 | 5.04 | 16728.5 KB | 6.08 |  |  | 404.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | EPPlus | 33.17 ms | 7.33 ms | 4.23 ms | 5.40 | 5.40 | 20152.7 KB | 7.32 |  |  | 440.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | Sylvan.Data.Excel | 0.53 ms | 0.10 ms | 0.06 ms | 0.87 | 1.00 | 348.5 KB | 1.18 |  |  | 12.6% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | OfficeIMO.Excel | 0.61 ms | 0.03 ms | 0.02 ms | 1.00 | 1.14 | 296.0 KB | 1.00 |  |  | Loss +14.5% |
| 2500 | speed-comparison | read | Range and table read | read-top-range | MiniExcel | 0.84 ms | 0.10 ms | 0.06 ms | 1.39 | 1.59 | 869.0 KB | 2.94 |  |  | 38.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | ExcelDataReader | 4.70 ms | 0.55 ms | 0.32 ms | 7.72 | 8.84 | 1931.8 KB | 6.53 |  |  | 672.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | EPPlus | 27.47 ms | 3.86 ms | 2.23 ms | 45.15 | 51.69 | 12402.1 KB | 41.89 |  |  | 4415.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | EPPlus 4.5.3.3 | 30.54 ms |  |  | 50.19 | 57.45 |  |  |  |  | 4918.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | ClosedXML | 34.56 ms | 4.35 ms | 2.51 ms | 56.80 | 65.02 | 15360.4 KB | 51.89 |  |  | 5579.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | Sylvan.Data.Excel | 4.67 ms | 0.23 ms | 0.13 ms | 0.35 | 1.00 | 655.2 KB | 0.19 |  |  | 65.3% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | ExcelDataReader | 10.48 ms | 0.14 ms | 0.08 ms | 0.78 | 2.25 | 6089.5 KB | 1.75 |  |  | 22.1% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | MiniExcel | 12.41 ms | 0.24 ms | 0.14 ms | 0.92 | 2.66 | 18662.4 KB | 5.37 |  |  | 7.8% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | OfficeIMO.Excel | 13.46 ms | 5.54 ms | 3.20 ms | 1.00 | 2.88 | 3472.6 KB | 1.00 |  |  | Loss +188.4% |
| 2500 | speed-comparison | read | Range and table read | read-used-range | EPPlus | 37.20 ms | 6.79 ms | 3.92 ms | 2.76 | 7.97 | 20152.8 KB | 5.80 |  |  | 176.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | ClosedXML | 52.00 ms | 35.19 ms | 20.32 ms | 3.86 | 11.14 | 16767.9 KB | 4.83 |  |  | 286.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | OfficeIMO.Excel | 3.53 ms | 0.05 ms | 0.03 ms | 1.00 | 1.00 | 377.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | Sylvan.Data.Excel | 4.00 ms | 0.23 ms | 0.13 ms | 1.13 | 1.13 | 655.2 KB | 1.73 |  |  | 13.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | ExcelDataReader | 10.07 ms | 0.70 ms | 0.41 ms | 2.85 | 2.85 | 6089.5 KB | 16.12 |  |  | 184.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | MiniExcel | 12.40 ms | 0.89 ms | 0.51 ms | 3.51 | 3.51 | 18661.8 KB | 49.40 |  |  | 251.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | EPPlus | 24.50 ms | 1.51 ms | 0.87 ms | 6.93 | 6.93 | 12427.1 KB | 32.89 |  |  | 593.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | ClosedXML | 28.47 ms | 0.41 ms | 0.24 ms | 8.06 | 8.06 | 15359.5 KB | 40.66 |  |  | 705.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | Sylvan.Data.Excel | 4.51 ms | 0.23 ms | 0.13 ms | 0.76 | 1.00 | 655.2 KB | 0.24 |  |  | 24.1% faster than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | OfficeIMO.Excel | 5.94 ms | 2.07 ms | 1.19 ms | 1.00 | 1.32 | 2771.3 KB | 1.00 |  |  | Loss +31.8% |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | ExcelDataReader | 12.22 ms | 1.16 ms | 0.67 ms | 2.06 | 2.71 | 6089.5 KB | 2.20 |  |  | 105.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | MiniExcel | 13.20 ms | 1.98 ms | 1.14 ms | 2.22 | 2.93 | 18662.4 KB | 6.73 |  |  | 122.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | EPPlus | 30.90 ms | 6.80 ms | 3.93 ms | 5.20 | 6.86 | 20152.6 KB | 7.27 |  |  | 420.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | EPPlus 4.5.3.3 | 31.47 ms |  |  | 5.30 | 6.98 |  |  |  |  | 429.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | ClosedXML | 32.20 ms | 2.98 ms | 1.72 ms | 5.42 | 7.14 | 16729.5 KB | 6.04 |  |  | 442.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | Sylvan.Data.Excel | 0.42 ms | 0.00 ms | 0.00 ms | 0.79 | 1.00 | 348.5 KB | 1.16 |  |  | 21.2% faster than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | OfficeIMO.Excel | 0.53 ms | 0.02 ms | 0.01 ms | 1.00 | 1.27 | 299.4 KB | 1.00 |  |  | Loss +27.0% |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | MiniExcel | 0.74 ms | 0.06 ms | 0.03 ms | 1.40 | 1.77 | 869.0 KB | 2.90 |  |  | 39.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | ExcelDataReader | 4.63 ms | 0.72 ms | 0.41 ms | 8.71 | 11.05 | 1931.8 KB | 6.45 |  |  | 770.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | EPPlus | 24.44 ms | 1.63 ms | 0.94 ms | 45.98 | 58.38 | 12402.1 KB | 41.43 |  |  | 4498.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | EPPlus 4.5.3.3 | 28.84 ms |  |  | 54.27 | 68.89 |  |  |  |  | 5326.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | ClosedXML | 29.86 ms | 0.56 ms | 0.32 ms | 56.19 | 71.33 | 15361.0 KB | 51.31 |  |  | 5518.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | Sylvan.Data.Excel | 0.42 ms | 0.01 ms | 0.01 ms | 0.62 | 1.00 | 348.5 KB | 1.16 |  |  | 38.1% faster than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | OfficeIMO.Excel | 0.67 ms | 0.21 ms | 0.12 ms | 1.00 | 1.62 | 300.1 KB | 1.00 |  |  | Loss +61.5% |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | MiniExcel | 0.87 ms | 0.18 ms | 0.11 ms | 1.30 | 2.09 | 869.0 KB | 2.90 |  |  | 29.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | ExcelDataReader | 5.22 ms | 0.54 ms | 0.31 ms | 7.74 | 12.51 | 1931.8 KB | 6.44 |  |  | 674.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | EPPlus | 26.20 ms | 0.93 ms | 0.54 ms | 38.83 | 62.71 | 12402.1 KB | 41.33 |  |  | 3782.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | ClosedXML | 30.94 ms | 1.91 ms | 1.10 ms | 45.85 | 74.06 | 15360.7 KB | 51.19 |  |  | 4484.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | OfficeIMO.Excel | 13.24 ms | 1.88 ms | 1.09 ms | 1.00 | 1.00 | 2441.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Typed object read | read-objects | Sylvan.Data.Excel | 18.49 ms | 11.55 ms | 6.67 ms | 1.40 | 1.40 | 895.3 KB | 0.37 |  |  | 39.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | EPPlus 4.5.3.3 | 31.66 ms |  |  | 2.39 | 2.39 |  |  |  |  | 139.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | ExcelDataReader | 40.24 ms | 8.25 ms | 4.77 ms | 3.04 | 3.04 | 6329.5 KB | 2.59 |  |  | 204.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | MiniExcel | 44.25 ms | 20.81 ms | 12.01 ms | 3.34 | 3.34 | 18473.9 KB | 7.57 |  |  | 234.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | EPPlus | 49.69 ms | 4.67 ms | 2.70 ms | 3.75 | 3.75 | 21354.3 KB | 8.74 |  |  | 275.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | ClosedXML | 67.44 ms | 15.15 ms | 8.75 ms | 5.10 | 5.10 | 16925.5 KB | 6.93 |  |  | 409.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | OfficeIMO.Excel | 7.59 ms | 2.58 ms | 1.49 ms | 1.00 | 1.00 | 2422.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | Sylvan.Data.Excel | 8.39 ms | 3.51 ms | 2.03 ms | 1.11 | 1.11 | 831.0 KB | 0.34 |  |  | 10.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | ExcelDataReader | 15.25 ms | 4.13 ms | 2.39 ms | 2.01 | 2.01 | 6265.3 KB | 2.59 |  |  | 100.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | MiniExcel | 19.76 ms | 3.53 ms | 2.04 ms | 2.60 | 2.60 | 18409.7 KB | 7.60 |  |  | 160.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | EPPlus 4.5.3.3 | 28.45 ms |  |  | 3.75 | 3.75 |  |  |  |  | 274.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | EPPlus | 43.60 ms | 12.19 ms | 7.04 ms | 5.74 | 5.74 | 21334.7 KB | 8.81 |  |  | 474.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | ClosedXML | 50.37 ms | 4.54 ms | 2.62 ms | 6.63 | 6.63 | 16904.0 KB | 6.98 |  |  | 563.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | OfficeIMO.Excel | 7.58 ms | 3.14 ms | 1.81 ms | 1.00 | 1.00 | 1781.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | MiniExcel | 9.51 ms | 1.47 ms | 0.85 ms | 1.26 | 1.26 | 26647.4 KB | 14.96 |  |  | 25.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | EPPlus | 59.00 ms | 5.71 ms | 3.30 ms | 7.79 | 7.79 | 38345.1 KB | 21.53 |  |  | 678.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | EPPlus 4.5.3.3 | 78.02 ms |  |  | 10.30 | 10.30 |  |  |  |  | 929.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | ClosedXML | 165.11 ms | 19.21 ms | 11.09 ms | 21.80 | 21.80 | 58360.0 KB | 32.77 |  |  | 2079.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | OfficeIMO.Excel | 5.97 ms | 0.36 ms | 0.21 ms | 1.00 | 1.00 | 2080.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | MiniExcel | 15.95 ms | 1.26 ms | 0.73 ms | 2.67 | 2.67 | 31859.9 KB | 15.31 |  |  | 167.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | EPPlus | 151.95 ms | 8.25 ms | 4.76 ms | 25.46 | 25.46 | 43440.2 KB | 20.88 |  |  | 2446.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | ClosedXML | 182.55 ms | 8.63 ms | 4.98 ms | 30.59 | 30.59 | 56708.6 KB | 27.26 |  |  | 2958.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 15.06 ms | 16.55 ms | 9.56 ms | 1.00 | 1.00 | 1347.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 85.75 ms | 1.24 ms | 0.72 ms | 5.69 | 5.69 | 38344.5 KB | 28.46 |  |  | 469.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 118.60 ms | 6.68 ms | 3.86 ms | 7.87 | 7.87 | 50927.7 KB | 37.80 |  |  | 687.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 7.52 ms | 2.94 ms | 1.70 ms | 1.00 | 1.00 | 1505.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 70.60 ms | 3.41 ms | 1.97 ms | 9.39 | 9.39 | 38344.5 KB | 25.47 |  |  | 838.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 95.09 ms | 12.13 ms | 7.00 ms | 12.64 | 12.64 | 50927.3 KB | 33.83 |  |  | 1164.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 7.62 ms | 1.31 ms | 0.76 ms | 1.00 | 1.00 | 1346.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 71.20 ms | 15.81 ms | 9.13 ms | 9.35 | 9.35 | 28540.4 KB | 21.20 |  |  | 834.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 78.97 ms | 3.54 ms | 2.04 ms | 10.37 | 10.37 | 27306.2 KB | 20.28 |  |  | 936.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | OfficeIMO.Excel | 2.63 ms | 0.46 ms | 0.26 ms | 1.00 | 1.00 | 1787.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | ClosedXML | 15.57 ms | 5.33 ms | 3.08 ms | 5.92 | 5.92 | 9959.5 KB | 5.57 |  |  | 492.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | EPPlus | 22.53 ms | 0.89 ms | 0.51 ms | 8.57 | 8.57 | 11773.4 KB | 6.59 |  |  | 756.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | OfficeIMO.Excel | 3.20 ms | 0.72 ms | 0.42 ms | 1.00 | 1.00 | 1119.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | ClosedXML | 12.05 ms | 1.22 ms | 0.70 ms | 3.76 | 3.76 | 9177.1 KB | 8.19 |  |  | 276.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | EPPlus 4.5.3.3 | 17.22 ms |  |  | 5.38 | 5.38 |  |  |  |  | 437.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | EPPlus | 25.62 ms | 0.08 ms | 0.04 ms | 8.00 | 8.00 | 12895.6 KB | 11.51 |  |  | 700.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | OfficeIMO.Excel | 3.41 ms | 0.55 ms | 0.32 ms | 1.00 | 1.00 | 1763.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | ClosedXML | 17.03 ms | 0.40 ms | 0.23 ms | 5.00 | 5.00 | 11887.0 KB | 6.74 |  |  | 399.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | EPPlus 4.5.3.3 | 17.30 ms |  |  | 5.08 | 5.08 |  |  |  |  | 407.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | EPPlus | 27.64 ms | 0.89 ms | 0.52 ms | 8.12 | 8.12 | 15643.8 KB | 8.87 |  |  | 711.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | OfficeIMO.Excel | 3.72 ms | 1.03 ms | 0.60 ms | 1.00 | 1.00 | 1506.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | ClosedXML | 15.37 ms | 0.97 ms | 0.56 ms | 4.13 | 4.13 | 11296.3 KB | 7.50 |  |  | 312.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | EPPlus | 28.52 ms | 2.37 ms | 1.37 ms | 7.66 | 7.66 | 14960.7 KB | 9.93 |  |  | 665.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 3.08 ms | 0.13 ms | 0.08 ms | 1.00 | 1.00 | 1506.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | ClosedXML | 15.05 ms | 0.51 ms | 0.30 ms | 4.89 | 4.89 | 11296.3 KB | 7.50 |  |  | 389.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | EPPlus | 26.93 ms | 3.63 ms | 2.09 ms | 8.75 | 8.75 | 14960.7 KB | 9.93 |  |  | 775.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | OfficeIMO.Excel | 2.57 ms | 0.44 ms | 0.26 ms | 1.00 | 1.00 | 1138.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | ClosedXML | 12.12 ms | 1.43 ms | 0.83 ms | 4.71 | 4.71 | 9021.2 KB | 7.93 |  |  | 371.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | EPPlus 4.5.3.3 | 15.86 ms |  |  | 6.17 | 6.17 |  |  |  |  | 517.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | EPPlus | 25.37 ms | 0.99 ms | 0.57 ms | 9.87 | 9.87 | 12827.9 KB | 11.27 |  |  | 886.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | OfficeIMO.Excel | 3.64 ms | 0.76 ms | 0.44 ms | 1.00 | 1.00 | 1435.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | EPPlus 4.5.3.3 | 17.59 ms |  |  | 4.83 | 4.83 |  |  |  |  | 383.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | ClosedXML | 17.89 ms | 5.02 ms | 2.90 ms | 4.91 | 4.91 | 9711.1 KB | 6.76 |  |  | 391.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | EPPlus | 29.04 ms | 3.07 ms | 1.77 ms | 7.98 | 7.98 | 14723.0 KB | 10.26 |  |  | 697.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | OfficeIMO.Excel | 4.95 ms | 0.27 ms | 0.16 ms | 1.00 | 1.00 | 2065.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | MiniExcel | 14.88 ms | 0.52 ms | 0.30 ms | 3.01 | 3.01 | 29223.6 KB | 14.15 |  |  | 200.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | ClosedXML | 66.79 ms | 1.56 ms | 0.90 ms | 13.48 | 13.48 | 18914.3 KB | 9.16 |  |  | 1248.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | EPPlus | 101.58 ms | 3.45 ms | 1.99 ms | 20.51 | 20.51 | 18415.4 KB | 8.92 |  |  | 1950.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | OfficeIMO.Excel | 7.34 ms | 0.21 ms | 0.12 ms | 1.00 | 1.00 | 2880.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | MiniExcel | 19.25 ms | 2.58 ms | 1.49 ms | 2.62 | 2.62 | 30510.6 KB | 10.59 |  |  | 162.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | ClosedXML | 93.18 ms | 5.89 ms | 3.40 ms | 12.70 | 12.70 | 27410.7 KB | 9.52 |  |  | 1170.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | EPPlus | 105.75 ms | 4.35 ms | 2.51 ms | 14.42 | 14.42 | 22605.3 KB | 7.85 |  |  | 1341.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | OfficeIMO.Excel | 5.55 ms | 0.22 ms | 0.13 ms | 1.00 | 1.00 | 2067.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | MiniExcel | 13.89 ms | 0.77 ms | 0.44 ms | 2.50 | 2.50 | 28700.3 KB | 13.88 |  |  | 150.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | EPPlus 4.5.3.3 | 54.07 ms |  |  | 9.75 | 9.75 |  |  |  |  | 875.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | ClosedXML | 63.06 ms | 2.14 ms | 1.23 ms | 11.37 | 11.37 | 18878.2 KB | 9.13 |  |  | 1037.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | EPPlus | 70.75 ms | 7.44 ms | 4.30 ms | 12.76 | 12.76 | 19431.0 KB | 9.40 |  |  | 1175.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | OfficeIMO.Excel | 4.67 ms | 0.10 ms | 0.06 ms | 1.00 | 1.00 | 1774.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | MiniExcel | 8.66 ms | 0.40 ms | 0.23 ms | 1.85 | 1.85 | 23044.2 KB | 12.98 |  |  | 85.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | ClosedXML | 38.22 ms | 0.55 ms | 0.32 ms | 8.18 | 8.18 | 19008.4 KB | 10.71 |  |  | 718.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | EPPlus 4.5.3.3 | 39.58 ms |  |  | 8.47 | 8.47 |  |  |  |  | 747.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | EPPlus | 44.60 ms | 5.38 ms | 3.10 ms | 9.55 | 9.55 | 16647.3 KB | 9.38 |  |  | 854.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | OfficeIMO.Excel | 7.70 ms | 2.13 ms | 1.23 ms | 1.00 | 1.00 | 1748.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | LargeXlsx | 7.97 ms | 0.26 ms | 0.15 ms | 1.03 | 1.03 | 1149.0 KB | 0.66 |  |  | 3.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | MiniExcel | 10.21 ms | 0.59 ms | 0.34 ms | 1.33 | 1.33 | 23062.6 KB | 13.19 |  |  | 32.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | ClosedXML | 37.21 ms | 10.99 ms | 6.34 ms | 4.83 | 4.83 | 11581.0 KB | 6.62 |  |  | 383.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | EPPlus | 43.91 ms | 4.47 ms | 2.58 ms | 5.70 | 5.70 | 16648.7 KB | 9.52 |  |  | 470.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | EPPlus 4.5.3.3 | 45.86 ms |  |  | 5.95 | 5.95 |  |  |  |  | 495.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | OfficeIMO.Excel | 5.28 ms | 0.49 ms | 0.28 ms | 1.00 | 1.00 | 1487.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | MiniExcel | 11.52 ms | 0.41 ms | 0.23 ms | 2.18 | 2.18 | 22789.5 KB | 15.32 |  |  | 118.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | ClosedXML | 41.38 ms | 2.16 ms | 1.25 ms | 7.83 | 7.83 | 18735.1 KB | 12.60 |  |  | 683.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | EPPlus | 46.89 ms | 5.08 ms | 2.93 ms | 8.88 | 8.88 | 16374.5 KB | 11.01 |  |  | 787.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | OfficeIMO.Excel | 4.80 ms | 0.26 ms | 0.15 ms | 1.00 | 1.00 | 1760.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | MiniExcel | 10.10 ms | 0.81 ms | 0.47 ms | 2.11 | 2.11 | 23062.9 KB | 13.10 |  |  | 110.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | EPPlus 4.5.3.3 | 42.21 ms |  |  | 8.80 | 8.80 |  |  |  |  | 780.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | EPPlus | 48.37 ms | 7.49 ms | 4.33 ms | 10.09 | 10.09 | 16648.8 KB | 9.46 |  |  | 908.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | ClosedXML | 49.33 ms | 16.16 ms | 9.33 ms | 10.29 | 10.29 | 19008.7 KB | 10.80 |  |  | 928.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | OfficeIMO.Excel | 5.80 ms | 0.26 ms | 0.15 ms | 1.00 | 1.00 | 1403.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | MiniExcel | 13.37 ms | 0.79 ms | 0.45 ms | 2.30 | 2.30 | 26825.0 KB | 19.12 |  |  | 130.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | EPPlus 4.5.3.3 | 104.37 ms |  |  | 17.98 | 17.98 |  |  |  |  | 1698.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | EPPlus | 127.91 ms | 4.82 ms | 2.78 ms | 22.04 | 22.04 | 49158.1 KB | 35.03 |  |  | 2103.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | ClosedXML | 219.04 ms | 55.76 ms | 32.19 ms | 37.74 | 37.74 | 58382.2 KB | 41.60 |  |  | 3674.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | OfficeIMO.Excel | 4.36 ms | 0.39 ms | 0.23 ms | 1.00 | 1.00 | 1620.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | EPPlus 4.5.3.3 | 19.69 ms |  |  | 4.52 | 4.52 |  |  |  |  | 352.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | ClosedXML | 26.18 ms | 2.35 ms | 1.35 ms | 6.01 | 6.01 | 12039.8 KB | 7.43 |  |  | 501.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | EPPlus | 51.94 ms | 3.20 ms | 1.85 ms | 11.92 | 11.92 | 18110.8 KB | 11.18 |  |  | 1092.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-dictionary-objects-table-direct | OfficeIMO.Excel | 7.13 ms | 1.81 ms | 1.04 ms | 1.00 | 1.00 | 2051.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | LargeXlsx | 4.54 ms | 0.33 ms | 0.19 ms | 0.75 | 1.00 | 802.5 KB | 0.34 |  |  | 24.9% faster than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 6.04 ms | 0.87 ms | 0.50 ms | 1.00 | 1.33 | 2341.7 KB | 1.00 |  |  | Loss +33.1% |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | MiniExcel | 9.72 ms | 0.69 ms | 0.40 ms | 1.61 | 2.14 | 25190.5 KB | 10.76 |  |  | 61.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | ClosedXML | 49.04 ms | 15.17 ms | 8.76 ms | 8.12 | 10.81 | 16973.5 KB | 7.25 |  |  | 712.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | EPPlus | 59.34 ms | 9.14 ms | 5.28 ms | 9.83 | 13.08 | 20105.6 KB | 8.59 |  |  | 882.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | LargeXlsx | 4.85 ms | 1.12 ms | 0.65 ms | 0.72 | 1.00 | 802.5 KB | 0.53 |  |  | 27.6% faster than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 6.71 ms | 0.81 ms | 0.47 ms | 1.00 | 1.38 | 1507.7 KB | 1.00 |  |  | Loss +38.2% |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | MiniExcel | 14.01 ms | 5.94 ms | 3.43 ms | 2.09 | 2.89 | 25190.5 KB | 16.71 |  |  | 108.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | ClosedXML | 66.92 ms | 45.51 ms | 26.28 ms | 9.97 | 13.78 | 16973.5 KB | 11.26 |  |  | 897.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | EPPlus | 71.05 ms | 35.98 ms | 20.77 ms | 10.59 | 14.63 | 20105.6 KB | 13.33 |  |  | 958.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 22.25 ms | 0.74 ms | 0.43 ms | 1.00 | 1.00 | 4502.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | LargeXlsx | 41.84 ms | 26.32 ms | 15.19 ms | 1.88 | 1.88 | 2810.7 KB | 0.62 |  |  | 88.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | MiniExcel | 55.69 ms | 26.81 ms | 15.48 ms | 2.50 | 2.50 | 48414.8 KB | 10.75 |  |  | 150.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | ClosedXML | 128.12 ms | 6.98 ms | 4.03 ms | 5.76 | 5.76 | 51647.0 KB | 11.47 |  |  | 475.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | EPPlus | 176.60 ms | 7.47 ms | 4.31 ms | 7.94 | 7.94 | 69140.0 KB | 15.36 |  |  | 693.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | LargeXlsx | 1.66 ms | 0.28 ms | 0.16 ms | 0.61 | 1.00 | 296.4 KB | 0.19 |  |  | 38.5% faster than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | OfficeIMO.Excel | 2.71 ms | 0.46 ms | 0.27 ms | 1.00 | 1.63 | 1576.9 KB | 1.00 |  |  | Loss +62.7% |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | MiniExcel | 4.91 ms | 0.94 ms | 0.54 ms | 1.81 | 2.95 | 19710.9 KB | 12.50 |  |  | 81.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | ClosedXML | 15.74 ms | 1.31 ms | 0.76 ms | 5.81 | 9.46 | 11197.4 KB | 7.10 |  |  | 481.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | EPPlus 4.5.3.3 | 16.50 ms |  |  | 6.10 | 9.92 |  |  |  |  | 509.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | EPPlus | 28.16 ms | 4.54 ms | 2.62 ms | 10.40 | 16.92 | 14365.5 KB | 9.11 |  |  | 940.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 1.43 ms | 0.13 ms | 0.08 ms | 0.83 | 1.00 | 447.0 KB | 0.41 |  |  | 16.7% faster than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 1.72 ms | 0.10 ms | 0.06 ms | 1.00 | 1.20 | 1092.0 KB | 1.00 |  |  | Loss +20.0% |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | ClosedXML | 17.27 ms | 5.71 ms | 3.30 ms | 10.06 | 12.07 | 10235.8 KB | 9.37 |  |  | 905.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | EPPlus | 24.65 ms | 2.26 ms | 1.30 ms | 14.35 | 17.23 | 13052.5 KB | 11.95 |  |  | 1335.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 4.25 ms | 0.16 ms | 0.09 ms | 1.00 | 1.00 | 2081.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | LargeXlsx | 4.67 ms | 2.27 ms | 1.31 ms | 1.10 | 1.10 | 758.3 KB | 0.36 |  |  | 9.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | MiniExcel | 9.25 ms | 0.76 ms | 0.44 ms | 2.18 | 2.18 | 23221.8 KB | 11.16 |  |  | 117.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | ClosedXML | 33.28 ms | 1.22 ms | 0.71 ms | 7.84 | 7.84 | 22221.3 KB | 10.68 |  |  | 683.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 35.40 ms |  |  | 8.33 | 8.33 |  |  |  |  | 733.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | EPPlus | 43.02 ms | 1.37 ms | 0.79 ms | 10.13 | 10.13 | 24694.2 KB | 11.87 |  |  | 912.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | OfficeIMO.Excel | 2.46 ms | 0.06 ms | 0.04 ms | 1.00 | 1.00 | 1494.0 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | ClosedXML | 15.15 ms | 0.80 ms | 0.46 ms | 6.17 | 6.17 | 11296.3 KB | 7.56 |  |  | 517.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | EPPlus | 26.15 ms | 1.67 ms | 0.96 ms | 10.65 | 10.65 | 14960.5 KB | 10.01 |  |  | 964.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | Sylvan.Data.Excel | 4.70 ms | 0.42 ms | 0.24 ms | 0.97 | 1.00 | 758.6 KB | 0.43 |  |  | 2.8% faster than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | OfficeIMO.Excel | 4.84 ms | 0.27 ms | 0.16 ms | 1.00 | 1.03 | 1763.0 KB | 1.00 |  |  | Loss +2.9% |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | LargeXlsx | 8.72 ms | 0.31 ms | 0.18 ms | 1.80 | 1.85 | 1032.5 KB | 0.59 |  |  | 80.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | MiniExcel | 9.95 ms | 1.33 ms | 0.77 ms | 2.06 | 2.11 | 23043.9 KB | 13.07 |  |  | 105.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | ClosedXML | 28.01 ms | 1.30 ms | 0.75 ms | 5.79 | 5.96 | 11581.0 KB | 6.57 |  |  | 478.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | EPPlus 4.5.3.3 | 32.76 ms |  |  | 6.77 | 6.97 |  |  |  |  | 576.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | EPPlus | 37.06 ms | 1.05 ms | 0.61 ms | 7.66 | 7.88 | 16647.0 KB | 9.44 |  |  | 665.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | OfficeIMO.Excel | 7.33 ms | 1.34 ms | 0.77 ms | 1.00 | 1.00 | 2140.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | LargeXlsx | 13.53 ms | 3.93 ms | 2.27 ms | 1.85 | 1.85 | 1123.9 KB | 0.53 |  |  | 84.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | MiniExcel | 18.08 ms | 2.37 ms | 1.37 ms | 2.47 | 2.47 | 30388.3 KB | 14.20 |  |  | 146.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | EPPlus | 70.09 ms | 19.16 ms | 11.06 ms | 9.57 | 9.57 | 22358.1 KB | 10.44 |  |  | 856.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | ClosedXML | 89.91 ms | 14.65 ms | 8.46 ms | 12.27 | 12.27 | 27410.6 KB | 12.80 |  |  | 1127.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | LargeXlsx | 4.31 ms | 0.37 ms | 0.21 ms | 0.97 | 1.00 | 857.6 KB | 0.51 |  |  | 3.1% faster than OfficeIMO |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | OfficeIMO.Excel | 4.45 ms | 0.13 ms | 0.07 ms | 1.00 | 1.03 | 1676.8 KB | 1.00 |  |  | Loss +3.2% |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | MiniExcel | 16.45 ms | 0.28 ms | 0.16 ms | 3.70 | 3.82 | 35918.3 KB | 21.42 |  |  | 269.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | ClosedXML | 88.92 ms | 2.49 ms | 1.44 ms | 19.98 | 20.62 | 71478.2 KB | 42.63 |  |  | 1897.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | OfficeIMO.Excel | 2.05 ms | 0.03 ms | 0.02 ms | 1.00 | 1.00 | 2440.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | MiniExcel | 3.97 ms | 0.23 ms | 0.13 ms | 1.94 | 1.94 | 21137.5 KB | 8.66 |  |  | 94.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | ClosedXML | 12.83 ms | 0.88 ms | 0.51 ms | 6.27 | 6.27 | 11299.2 KB | 4.63 |  |  | 527.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | EPPlus 4.5.3.3 | 14.14 ms |  |  | 6.91 | 6.91 |  |  |  |  | 591.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | EPPlus | 25.40 ms | 3.44 ms | 1.99 ms | 12.42 | 12.42 | 12804.8 KB | 5.25 |  |  | 1142.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | OfficeIMO.Excel | 2.85 ms | 0.15 ms | 0.09 ms | 1.00 | 1.00 | 2617.0 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | ClosedXML | 11.80 ms | 0.52 ms | 0.30 ms | 4.14 | 4.14 | 11299.2 KB | 4.32 |  |  | 313.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | EPPlus 4.5.3.3 | 17.38 ms |  |  | 6.10 | 6.10 |  |  |  |  | 509.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | EPPlus | 23.73 ms | 6.39 ms | 3.69 ms | 8.32 | 8.32 | 12805.2 KB | 4.89 |  |  | 732.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | OfficeIMO.Excel | 2.50 ms | 0.26 ms | 0.15 ms | 1.00 | 1.00 | 2379.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | ClosedXML | 14.40 ms | 1.02 ms | 0.59 ms | 5.75 | 5.75 | 13127.1 KB | 5.52 |  |  | 475.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | EPPlus | 23.43 ms | 0.90 ms | 0.52 ms | 9.36 | 9.36 | 13893.4 KB | 5.84 |  |  | 836.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | OfficeIMO.Excel | 2.06 ms | 0.05 ms | 0.03 ms | 1.00 | 1.00 | 1579.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | ClosedXML | 12.47 ms | 1.69 ms | 0.97 ms | 6.04 | 6.04 | 9226.5 KB | 5.84 |  |  | 504.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | EPPlus | 23.44 ms | 7.79 ms | 4.50 ms | 11.36 | 11.36 | 11332.8 KB | 7.17 |  |  | 1035.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | LargeXlsx | 3.38 ms | 0.37 ms | 0.21 ms | 0.82 | 1.00 | 758.3 KB | 0.43 |  |  | 18.4% faster than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 4.14 ms | 0.52 ms | 0.30 ms | 1.00 | 1.23 | 1769.2 KB | 1.00 |  |  | Loss +22.5% |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | MiniExcel | 8.36 ms | 0.27 ms | 0.15 ms | 2.02 | 2.47 | 23222.4 KB | 13.13 |  |  | 101.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | ClosedXML | 30.22 ms | 1.37 ms | 0.79 ms | 7.30 | 8.94 | 11581.0 KB | 6.55 |  |  | 629.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 37.16 ms |  |  | 8.97 | 10.99 |  |  |  |  | 797.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | EPPlus | 39.29 ms | 5.71 ms | 3.30 ms | 9.49 | 11.63 | 16646.8 KB | 9.41 |  |  | 848.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | LargeXlsx | 4.09 ms | 0.14 ms | 0.08 ms | 0.69 | 1.00 | 758.3 KB | 0.57 |  |  | 30.7% faster than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | OfficeIMO.Excel | 5.91 ms | 1.54 ms | 0.89 ms | 1.00 | 1.44 | 1339.3 KB | 1.00 |  |  | Loss +44.3% |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | MiniExcel | 11.66 ms | 1.27 ms | 0.73 ms | 1.97 | 2.85 | 23222.4 KB | 17.34 |  |  | 97.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | EPPlus 4.5.3.3 | 35.44 ms |  |  | 6.00 | 8.65 |  |  |  |  | 499.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | ClosedXML | 59.53 ms | 6.56 ms | 3.79 ms | 10.08 | 14.54 | 11581.0 KB | 8.65 |  |  | 907.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | EPPlus | 64.43 ms | 11.76 ms | 6.79 ms | 10.91 | 15.74 | 16646.5 KB | 12.43 |  |  | 990.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 4.94 ms | 0.08 ms | 0.05 ms | 0.68 | 1.00 | 758.3 KB | 0.51 |  |  | 31.7% faster than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 7.23 ms | 0.36 ms | 0.21 ms | 1.00 | 1.46 | 1497.5 KB | 1.00 |  |  | Loss +46.3% |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | MiniExcel | 16.89 ms | 4.22 ms | 2.44 ms | 2.34 | 3.42 | 23222.4 KB | 15.51 |  |  | 133.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | ClosedXML | 40.65 ms | 9.87 ms | 5.70 ms | 5.63 | 8.23 | 11581.0 KB | 7.73 |  |  | 462.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | EPPlus | 58.71 ms | 7.28 ms | 4.20 ms | 8.13 | 11.89 | 16646.5 KB | 11.12 |  |  | 712.6% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | Sylvan.Data.Excel | 61.06 ms | 2.65 ms | 1.53 ms | 0.83 | 1.00 | 394.1 KB | 0.02 |  |  | 17.5% faster than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | OfficeIMO.Excel | 73.99 ms | 1.60 ms | 0.93 ms | 1.00 | 1.21 | 23622.1 KB | 1.00 |  |  | Loss +21.2% |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | ExcelDataReader | 162.93 ms | 5.91 ms | 3.41 ms | 2.20 | 2.67 | 69530.7 KB | 2.94 |  |  | 120.2% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | MiniExcel | 217.43 ms | 1.85 ms | 1.07 ms | 2.94 | 3.56 | 215349.1 KB | 9.12 |  |  | 193.9% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | Sylvan.Data.Excel | 66.62 ms | 4.01 ms | 2.32 ms | 0.97 | 1.00 | 394.1 KB | 0.02 |  |  | 3.4% faster than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | OfficeIMO.Excel | 68.98 ms | 1.94 ms | 1.12 ms | 1.00 | 1.04 | 24404.2 KB | 1.00 |  |  | Loss +3.5% |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | ExcelDataReader | 194.35 ms | 18.77 ms | 10.84 ms | 2.82 | 2.92 | 69530.7 KB | 2.85 |  |  | 181.7% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | MiniExcel | 244.62 ms | 9.34 ms | 5.39 ms | 3.55 | 3.67 | 215349.1 KB | 8.82 |  |  | 254.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | append-plain-rows | LargeXlsx | 14.33 ms | 1.92 ms | 1.11 ms | 0.70 | 1.00 | 2771.0 KB | 0.26 | 605.0 KB | 0.99 | 29.9% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | append-plain-rows | OfficeIMO.Excel | 20.45 ms | 1.76 ms | 1.02 ms | 1.00 | 1.43 | 10843.1 KB | 1.00 | 610.4 KB | 1.00 | Loss +42.7% |
| 25000 | package-profile | package | Package size | append-plain-rows | MiniExcel | 36.40 ms | 1.78 ms | 1.03 ms | 1.78 | 2.54 | 58242.9 KB | 5.37 | 642.3 KB | 1.05 | 78.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | append-plain-rows | ClosedXML | 163.87 ms | 7.01 ms | 4.05 ms | 8.01 | 11.44 | 104233.1 KB | 9.61 | 540.6 KB | 0.89 | 701.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | append-plain-rows | EPPlus | 244.50 ms | 11.96 ms | 6.90 ms | 11.96 | 17.07 | 100373.9 KB | 9.26 | 525.6 KB | 0.86 | 1095.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | autofit-existing | OfficeIMO.Excel | 92.81 ms | 1.52 ms | 0.88 ms | 1.00 | 1.00 | 15708.5 KB | 1.00 | 1429.3 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | autofit-existing | EPPlus | 594.57 ms | 6.11 ms | 3.53 ms | 6.41 | 6.41 | 250950.0 KB | 15.98 | 1091.0 KB | 0.76 | 540.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | autofit-existing | ClosedXML | 1779.08 ms | 13.17 ms | 7.60 ms | 19.17 | 19.17 | 829717.1 KB | 52.82 | 1140.9 KB | 0.80 | 1816.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | large-shared-strings | OfficeIMO.Excel | 20.06 ms | 2.54 ms | 1.47 ms | 1.00 | 1.00 | 15744.9 KB | 1.00 | 529.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | large-shared-strings | MiniExcel | 39.29 ms | 2.91 ms | 1.68 ms | 1.96 | 1.96 | 73760.2 KB | 4.68 | 581.0 KB | 1.10 | 95.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | large-shared-strings | ClosedXML | 160.95 ms | 5.03 ms | 2.91 ms | 8.02 | 8.02 | 104241.3 KB | 6.62 | 460.1 KB | 0.87 | 702.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | large-shared-strings | EPPlus | 283.73 ms | 1.50 ms | 0.87 ms | 14.14 | 14.14 | 84410.3 KB | 5.36 | 444.7 KB | 0.84 | 1314.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-autofilter | OfficeIMO.Excel | 43.84 ms | 0.80 ms | 0.46 ms | 1.00 | 1.00 | 11494.9 KB | 1.00 | 1428.5 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-autofilter | ClosedXML | 398.52 ms | 7.92 ms | 4.57 ms | 9.09 | 9.09 | 210663.8 KB | 18.33 | 1140.0 KB | 0.80 | 809.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-autofilter | EPPlus | 477.44 ms | 18.95 ms | 10.94 ms | 10.89 | 10.89 | 211871.8 KB | 18.43 | 1090.1 KB | 0.76 | 989.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-charts | OfficeIMO.Excel | 45.87 ms | 1.23 ms | 0.71 ms | 1.00 | 1.00 | 12553.6 KB | 1.00 | 1433.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-charts | EPPlus | 513.82 ms | 41.22 ms | 23.80 ms | 11.20 | 11.20 | 214906.2 KB | 17.12 | 1092.9 KB | 0.76 | 1020.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-conditional-formatting | OfficeIMO.Excel | 42.72 ms | 3.43 ms | 1.98 ms | 1.00 | 1.00 | 11560.2 KB | 1.00 | 1428.8 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-conditional-formatting | ClosedXML | 401.24 ms | 6.54 ms | 3.78 ms | 9.39 | 9.39 | 210711.7 KB | 18.23 | 1140.1 KB | 0.80 | 839.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-conditional-formatting | EPPlus | 522.13 ms | 4.89 ms | 2.82 ms | 12.22 | 12.22 | 211913.3 KB | 18.33 | 1090.2 KB | 0.76 | 1122.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-data-validation | OfficeIMO.Excel | 42.86 ms | 3.52 ms | 2.03 ms | 1.00 | 1.00 | 11510.5 KB | 1.00 | 1428.6 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-data-validation | ClosedXML | 383.78 ms | 18.24 ms | 10.53 ms | 8.95 | 8.95 | 210672.7 KB | 18.30 | 1140.1 KB | 0.80 | 795.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-data-validation | EPPlus | 527.65 ms | 9.86 ms | 5.69 ms | 12.31 | 12.31 | 211857.8 KB | 18.41 | 1090.1 KB | 0.76 | 1131.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-freeze-panes | OfficeIMO.Excel | 42.53 ms | 0.43 ms | 0.25 ms | 1.00 | 1.00 | 11497.3 KB | 1.00 | 1428.6 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-freeze-panes | ClosedXML | 395.05 ms | 10.65 ms | 6.15 ms | 9.29 | 9.29 | 210646.8 KB | 18.32 | 1140.0 KB | 0.80 | 828.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-freeze-panes | EPPlus | 463.69 ms | 5.24 ms | 3.03 ms | 10.90 | 10.90 | 211883.7 KB | 18.43 | 1090.2 KB | 0.76 | 990.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-pivot-table | OfficeIMO.Excel | 411.32 ms | 54.47 ms | 31.45 ms | 1.00 | 1.00 | 143628.5 KB | 1.00 | 1979.8 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-pivot-table | EPPlus | 602.49 ms | 25.10 ms | 14.49 ms | 1.46 | 1.46 | 230801.8 KB | 1.61 | 1093.4 KB | 0.55 | 46.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-all-in-one | OfficeIMO.Excel | 307.01 ms | 8.07 ms | 4.66 ms | 1.00 | 1.00 | 145144.6 KB | 1.00 | 1985.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-all-in-one | EPPlus | 527.21 ms | 3.48 ms | 2.01 ms | 1.72 | 1.72 | 277079.0 KB | 1.91 | 1097.7 KB | 0.55 | 71.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-chart-first | OfficeIMO.Excel | 119.04 ms | 4.48 ms | 2.59 ms | 1.00 | 1.00 | 55261.1 KB | 1.00 | 1985.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-chart-first | EPPlus | 535.98 ms | 4.86 ms | 2.80 ms | 4.50 | 4.50 | 277076.7 KB | 5.01 | 1097.7 KB | 0.55 | 350.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-core | OfficeIMO.Excel | 46.76 ms | 3.96 ms | 2.29 ms | 1.00 | 1.00 | 11648.7 KB | 1.00 | 1430.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-core | EPPlus | 516.28 ms | 3.99 ms | 2.30 ms | 11.04 | 11.04 | 255066.2 KB | 21.90 | 1091.5 KB | 0.76 | 1004.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-core | ClosedXML | 1088.14 ms | 33.82 ms | 19.52 ms | 23.27 | 23.27 | 680116.8 KB | 58.39 | 1141.3 KB | 0.80 | 2227.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-extra-column | OfficeIMO.Excel | 325.78 ms | 5.90 ms | 3.41 ms | 1.00 | 1.00 | 156862.3 KB | 1.00 | 2110.8 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-extra-column | EPPlus | 572.39 ms | 14.55 ms | 8.40 ms | 1.76 | 1.76 | 302761.3 KB | 1.93 | 1166.3 KB | 0.55 | 75.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-no-autofit | OfficeIMO.Excel | 320.07 ms | 3.84 ms | 2.22 ms | 1.00 | 1.00 | 145135.1 KB | 1.00 | 1985.8 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-no-autofit | EPPlus | 498.50 ms | 20.07 ms | 11.59 ms | 1.56 | 1.56 | 234783.8 KB | 1.62 | 1097.7 KB | 0.55 | 55.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-post-mutation | OfficeIMO.Excel | 336.55 ms | 17.18 ms | 9.92 ms | 1.00 | 1.00 | 145153.9 KB | 1.00 | 1986.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-post-mutation | EPPlus | 564.93 ms | 34.29 ms | 19.80 ms | 1.68 | 1.68 | 277079.0 KB | 1.91 | 1097.8 KB | 0.55 | 67.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-shuffled-columns | OfficeIMO.Excel | 325.90 ms | 12.74 ms | 7.36 ms | 1.00 | 1.00 | 145205.6 KB | 1.00 | 2046.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-shuffled-columns | EPPlus | 537.19 ms | 3.90 ms | 2.25 ms | 1.65 | 1.65 | 277071.7 KB | 1.91 | 1098.4 KB | 0.54 | 64.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook | OfficeIMO.Excel | 443.58 ms | 15.19 ms | 8.77 ms | 1.00 | 1.00 | 196356.9 KB | 1.00 | 2672.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | report-workbook | EPPlus | 711.76 ms | 16.19 ms | 9.35 ms | 1.60 | 1.60 | 364710.3 KB | 1.86 | 1517.2 KB | 0.57 | 60.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook-core | OfficeIMO.Excel | 63.24 ms | 2.77 ms | 1.60 ms | 1.00 | 1.00 | 10979.4 KB | 1.00 | 1850.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | report-workbook-core | EPPlus | 673.99 ms | 24.58 ms | 14.19 ms | 10.66 | 10.66 | 342842.6 KB | 31.23 | 1512.6 KB | 0.82 | 965.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook-core | ClosedXML | 1446.27 ms | 11.91 ms | 6.87 ms | 22.87 | 22.87 | 975774.2 KB | 88.87 | 1579.8 KB | 0.85 | 2186.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook-datatable | OfficeIMO.Excel | 467.76 ms | 27.13 ms | 15.66 ms | 1.00 | 1.00 | 199104.2 KB | 1.00 | 2672.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | report-workbook-datatable | EPPlus | 710.89 ms | 13.33 ms | 7.69 ms | 1.52 | 1.52 | 247824.2 KB | 1.24 | 1517.2 KB | 0.57 | 52.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook-datatable-core | OfficeIMO.Excel | 76.78 ms | 8.35 ms | 4.82 ms | 1.00 | 1.00 | 13725.0 KB | 1.00 | 1850.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | report-workbook-datatable-core | EPPlus | 678.16 ms | 18.34 ms | 10.59 ms | 8.83 | 8.83 | 225957.5 KB | 16.46 | 1512.6 KB | 0.82 | 783.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook-datatable-core | ClosedXML | 1422.23 ms | 54.18 ms | 31.28 ms | 18.52 | 18.52 | 832229.0 KB | 60.64 | 1579.8 KB | 0.85 | 1752.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | LargeXlsx | 55.09 ms | 2.69 ms | 1.55 ms | 0.88 | 1.00 | 10795.2 KB | 0.92 | 2444.6 KB | 1.10 | 11.6% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | OfficeIMO.Excel | 62.30 ms | 3.02 ms | 1.75 ms | 1.00 | 1.13 | 11708.2 KB | 1.00 | 2228.8 KB | 1.00 | Loss +13.1% |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | MiniExcel | 199.36 ms | 3.44 ms | 1.98 ms | 3.20 | 3.62 | 226876.0 KB | 19.38 | 2410.6 KB | 1.08 | 220.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | ClosedXML | 1187.01 ms | 15.57 ms | 8.99 ms | 19.05 | 21.55 | 759818.4 KB | 64.90 | 2581.2 KB | 1.16 | 1805.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-bulk-report | OfficeIMO.Excel | 41.76 ms | 3.41 ms | 1.97 ms | 1.00 | 1.00 | 11561.8 KB | 1.00 | 1429.3 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-bulk-report | MiniExcel | 100.58 ms | 19.89 ms | 11.48 ms | 2.41 | 2.41 | 125551.5 KB | 10.86 | 1521.1 KB | 1.06 | 140.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-bulk-report | EPPlus | 507.40 ms | 37.51 ms | 21.66 ms | 12.15 | 12.15 | 254959.4 KB | 22.05 | 1091.0 KB | 0.76 | 1115.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-bulk-report | ClosedXML | 1021.56 ms | 30.08 ms | 17.37 ms | 24.46 | 24.46 | 565953.3 KB | 48.95 | 1140.9 KB | 0.80 | 2346.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellformula | OfficeIMO.Excel | 31.55 ms | 1.25 ms | 0.72 ms | 1.00 | 1.00 | 10112.0 KB | 1.00 | 670.3 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellformula | ClosedXML | 266.92 ms | 13.79 ms | 7.96 ms | 8.46 | 8.46 | 113853.5 KB | 11.26 | 643.2 KB | 0.96 | 746.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellformula | EPPlus | 443.25 ms | 17.98 ms | 10.38 ms | 14.05 | 14.05 | 140732.3 KB | 13.92 | 593.9 KB | 0.89 | 1305.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-empty-strings | OfficeIMO.Excel | 16.17 ms | 1.27 ms | 0.73 ms | 1.00 | 1.00 | 6896.4 KB | 1.00 | 451.4 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-empty-strings | ClosedXML | 154.30 ms | 6.48 ms | 3.74 ms | 9.54 | 9.54 | 92902.1 KB | 13.47 | 398.1 KB | 0.88 | 854.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-empty-strings | EPPlus | 205.95 ms | 18.79 ms | 10.85 ms | 12.74 | 12.74 | 74493.1 KB | 10.80 | 390.6 KB | 0.87 | 1173.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-numbers | OfficeIMO.Excel | 17.61 ms | 1.65 ms | 0.96 ms | 1.00 | 1.00 | 5970.9 KB | 1.00 | 462.6 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-numbers | ClosedXML | 143.04 ms | 3.89 ms | 2.24 ms | 8.12 | 8.12 | 84206.7 KB | 14.10 | 411.4 KB | 0.89 | 712.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-numbers | EPPlus | 231.65 ms | 10.77 ms | 6.22 ms | 13.15 | 13.15 | 86377.9 KB | 14.47 | 406.5 KB | 0.88 | 1215.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-mixed | OfficeIMO.Excel | 22.69 ms | 1.38 ms | 0.80 ms | 1.00 | 1.00 | 8332.7 KB | 1.00 | 585.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-object-mixed | ClosedXML | 206.68 ms | 18.42 ms | 10.64 ms | 9.11 | 9.11 | 111118.7 KB | 13.34 | 532.9 KB | 0.91 | 811.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-mixed | EPPlus | 267.33 ms | 8.96 ms | 5.17 ms | 11.78 | 11.78 | 113245.5 KB | 13.59 | 544.3 KB | 0.93 | 1078.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse | OfficeIMO.Excel | 25.91 ms | 1.96 ms | 1.13 ms | 1.00 | 1.00 | 7416.0 KB | 1.00 | 607.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse | ClosedXML | 181.20 ms | 6.67 ms | 3.85 ms | 6.99 | 6.99 | 105223.9 KB | 14.19 | 468.0 KB | 0.77 | 599.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse | EPPlus | 258.58 ms | 14.27 ms | 8.24 ms | 9.98 | 9.98 | 106317.3 KB | 14.34 | 494.4 KB | 0.81 | 898.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 24.04 ms | 1.86 ms | 1.08 ms | 1.00 | 1.00 | 7416.1 KB | 1.00 | 607.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | ClosedXML | 204.65 ms | 11.49 ms | 6.64 ms | 8.51 | 8.51 | 105223.9 KB | 14.19 | 468.0 KB | 0.77 | 751.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | EPPlus | 291.04 ms | 16.37 ms | 9.45 ms | 12.11 | 12.11 | 106317.3 KB | 14.34 | 494.4 KB | 0.81 | 1110.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-scalars | OfficeIMO.Excel | 15.41 ms | 0.81 ms | 0.47 ms | 1.00 | 1.00 | 6144.6 KB | 1.00 | 441.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-scalars | ClosedXML | 126.06 ms | 3.53 ms | 2.04 ms | 8.18 | 8.18 | 82591.3 KB | 13.44 | 394.9 KB | 0.89 | 718.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-scalars | EPPlus | 226.89 ms | 2.59 ms | 1.50 ms | 14.72 | 14.72 | 85127.8 KB | 13.85 | 379.3 KB | 0.86 | 1372.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings | OfficeIMO.Excel | 21.93 ms | 2.44 ms | 1.41 ms | 1.00 | 1.00 | 15360.4 KB | 1.00 | 527.8 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-strings | ClosedXML | 149.89 ms | 14.23 ms | 8.22 ms | 6.84 | 6.84 | 104241.3 KB | 6.79 | 460.1 KB | 0.87 | 583.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings | EPPlus | 227.21 ms | 7.90 ms | 4.56 ms | 10.36 | 10.36 | 84410.8 KB | 5.50 | 444.7 KB | 0.84 | 936.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-distinct | OfficeIMO.Excel | 17.75 ms | 1.66 ms | 0.96 ms | 1.00 | 1.00 | 13824.1 KB | 1.00 | 499.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-distinct | ClosedXML | 202.70 ms | 11.21 ms | 6.47 ms | 11.42 | 11.42 | 131501.7 KB | 9.51 | 555.3 KB | 1.11 | 1042.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-distinct | EPPlus | 264.92 ms | 4.92 ms | 2.84 ms | 14.93 | 14.93 | 97730.0 KB | 7.07 | 565.1 KB | 1.13 | 1392.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-repeated | OfficeIMO.Excel | 16.46 ms | 0.23 ms | 0.13 ms | 1.00 | 1.00 | 7525.3 KB | 1.00 | 376.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-repeated | ClosedXML | 132.97 ms | 9.66 ms | 5.58 ms | 8.08 | 8.08 | 84520.0 KB | 11.23 | 331.8 KB | 0.88 | 707.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-repeated | EPPlus | 192.12 ms | 0.12 ms | 0.07 ms | 11.67 | 11.67 | 70033.7 KB | 9.31 | 300.8 KB | 0.80 | 1067.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-temporal | OfficeIMO.Excel | 28.28 ms | 3.45 ms | 1.99 ms | 1.00 | 1.00 | 7482.6 KB | 1.00 | 620.5 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-temporal | ClosedXML | 193.06 ms | 20.44 ms | 11.80 ms | 6.83 | 6.83 | 89323.7 KB | 11.94 | 483.0 KB | 0.78 | 582.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-temporal | EPPlus | 248.15 ms | 4.53 ms | 2.62 ms | 8.78 | 8.78 | 103800.4 KB | 13.87 | 495.1 KB | 0.80 | 777.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 12.59 ms | 0.48 ms | 0.28 ms | 0.86 | 1.00 | 3444.4 KB | 0.49 | 443.4 KB | 0.97 | 14.2% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 14.67 ms | 1.03 ms | 0.60 ms | 1.00 | 1.17 | 6961.7 KB | 1.00 | 455.5 KB | 1.00 | Loss +16.5% |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | ClosedXML | 163.44 ms | 4.81 ms | 2.77 ms | 11.14 | 12.98 | 96015.7 KB | 13.79 | 467.5 KB | 1.03 | 1013.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | EPPlus | 236.59 ms | 5.77 ms | 3.33 ms | 16.12 | 18.79 | 87467.3 KB | 12.56 | 484.1 KB | 1.06 | 1512.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | LargeXlsx | 38.15 ms | 1.40 ms | 0.81 ms | 0.82 | 1.00 | 5614.1 KB | 0.35 | 1386.5 KB | 1.00 | 18.1% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 46.60 ms | 3.03 ms | 1.75 ms | 1.00 | 1.22 | 16036.5 KB | 1.00 | 1384.9 KB | 1.00 | Loss +22.1% |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | MiniExcel | 94.53 ms | 17.06 ms | 9.85 ms | 2.03 | 2.48 | 93257.1 KB | 5.82 | 1521.1 KB | 1.10 | 102.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | ClosedXML | 417.65 ms | 47.60 ms | 27.48 ms | 8.96 | 10.95 | 210646.1 KB | 13.14 | 1139.9 KB | 0.82 | 796.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | EPPlus | 502.29 ms | 68.64 ms | 39.63 ms | 10.78 | 13.17 | 211850.3 KB | 13.21 | 1090.0 KB | 0.79 | 977.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | Sylvan.Data.Excel | 40.07 ms | 4.77 ms | 2.75 ms | 0.68 | 1.00 | 5700.3 KB | 0.44 | 755.4 KB | 0.55 | 32.0% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | OfficeIMO.Excel | 58.95 ms | 9.46 ms | 5.46 ms | 1.00 | 1.47 | 13002.3 KB | 1.00 | 1384.9 KB | 1.00 | Loss +47.1% |
| 25000 | package-profile | package | Package size | write-datareader-plain | LargeXlsx | 71.81 ms | 32.96 ms | 19.03 ms | 1.22 | 1.79 | 8349.2 KB | 0.64 | 1386.5 KB | 1.00 | 21.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | MiniExcel | 111.30 ms | 13.05 ms | 7.54 ms | 1.89 | 2.78 | 92199.8 KB | 7.09 | 1521.0 KB | 1.10 | 88.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | ClosedXML | 466.33 ms | 189.31 ms | 109.30 ms | 7.91 | 11.64 | 104205.0 KB | 8.01 | 1139.9 KB | 0.82 | 691.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | EPPlus | 512.75 ms | 73.96 ms | 42.70 ms | 8.70 | 12.80 | 117438.0 KB | 9.03 | 1090.8 KB | 0.79 | 769.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table | OfficeIMO.Excel | 46.01 ms | 1.17 ms | 0.67 ms | 1.00 | 1.00 | 13020.3 KB | 1.00 | 1385.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-datareader-table | MiniExcel | 93.22 ms | 3.14 ms | 1.81 ms | 2.03 | 2.03 | 92200.0 KB | 7.08 | 1521.0 KB | 1.10 | 102.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table | EPPlus | 468.99 ms | 48.40 ms | 27.94 ms | 10.19 | 10.19 | 117437.6 KB | 9.02 | 1090.8 KB | 0.79 | 919.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table | ClosedXML | 492.09 ms | 4.81 ms | 2.78 ms | 10.70 | 10.70 | 173397.5 KB | 13.32 | 1140.7 KB | 0.82 | 969.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | OfficeIMO.Excel | 55.69 ms | 3.09 ms | 1.79 ms | 1.00 | 1.00 | 13026.6 KB | 1.00 | 1385.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | MiniExcel | 103.68 ms | 5.40 ms | 3.12 ms | 1.86 | 1.86 | 124495.5 KB | 9.56 | 1521.1 KB | 1.10 | 86.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | EPPlus | 466.75 ms | 19.07 ms | 11.01 ms | 8.38 | 8.38 | 159742.2 KB | 12.26 | 1091.0 KB | 0.79 | 738.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | ClosedXML | 1090.73 ms | 20.30 ms | 11.72 ms | 19.59 | 19.59 | 566142.3 KB | 43.46 | 1140.9 KB | 0.82 | 1858.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | OfficeIMO.Excel | 42.31 ms | 5.17 ms | 2.99 ms | 1.00 | 1.00 | 9819.7 KB | 1.00 | 1329.2 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | LargeXlsx | 48.29 ms | 1.54 ms | 0.89 ms | 1.14 | 1.14 | 9265.9 KB | 0.94 | 1680.0 KB | 1.26 | 14.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | MiniExcel | 120.23 ms | 1.23 ms | 0.71 ms | 2.84 | 2.84 | 108129.1 KB | 11.01 | 1819.7 KB | 1.37 | 184.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | EPPlus | 590.56 ms | 31.06 ms | 17.93 ms | 13.96 | 13.96 | 135724.0 KB | 13.82 | 1390.4 KB | 1.05 | 1295.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | ClosedXML | 679.04 ms | 44.47 ms | 25.67 ms | 16.05 | 16.05 | 280372.9 KB | 28.55 | 1519.9 KB | 1.14 | 1505.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | OfficeIMO.Excel | 49.85 ms | 3.35 ms | 1.93 ms | 1.00 | 1.00 | 13458.5 KB | 1.00 | 1795.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | MiniExcel | 123.98 ms | 4.78 ms | 2.76 ms | 2.49 | 2.49 | 108129.1 KB | 8.03 | 1819.7 KB | 1.01 | 148.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | EPPlus | 592.86 ms | 27.68 ms | 15.98 ms | 11.89 | 11.89 | 135724.0 KB | 10.08 | 1390.4 KB | 0.77 | 1089.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | ClosedXML | 712.67 ms | 30.33 ms | 17.51 ms | 14.30 | 14.30 | 280371.8 KB | 20.83 | 1519.9 KB | 0.85 | 1329.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables | OfficeIMO.Excel | 49.25 ms | 2.49 ms | 1.44 ms | 1.00 | 1.00 | 10288.1 KB | 1.00 | 1376.4 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-dataset-tables | MiniExcel | 108.39 ms | 1.54 ms | 0.89 ms | 2.20 | 2.20 | 97085.4 KB | 9.44 | 1511.8 KB | 1.10 | 120.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables | EPPlus | 416.33 ms | 11.65 ms | 6.73 ms | 8.45 | 8.45 | 110816.3 KB | 10.77 | 1100.6 KB | 0.80 | 745.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables | ClosedXML | 478.62 ms | 22.03 ms | 12.72 ms | 9.72 | 9.72 | 172003.7 KB | 16.72 | 1139.0 KB | 0.83 | 871.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | OfficeIMO.Excel | 69.82 ms | 10.87 ms | 6.28 ms | 1.00 | 1.00 | 10300.7 KB | 1.00 | 1376.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | MiniExcel | 131.66 ms | 37.08 ms | 21.41 ms | 1.89 | 1.89 | 128874.9 KB | 12.51 | 1512.0 KB | 1.10 | 88.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | EPPlus | 643.58 ms | 58.89 ms | 34.00 ms | 9.22 | 9.22 | 195408.4 KB | 18.97 | 1100.9 KB | 0.80 | 821.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | ClosedXML | 1232.93 ms | 136.32 ms | 78.70 ms | 17.66 | 17.66 | 550095.1 KB | 53.40 | 1139.3 KB | 0.83 | 1665.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-direct | LargeXlsx | 39.84 ms | 4.12 ms | 2.38 ms | 0.91 | 1.00 | 9520.4 KB | 0.75 | 1386.5 KB | 1.00 | 8.7% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-direct | OfficeIMO.Excel | 43.62 ms | 3.03 ms | 1.75 ms | 1.00 | 1.09 | 12715.7 KB | 1.00 | 1384.9 KB | 1.00 | Loss +9.5% |
| 25000 | package-profile | package | Package size | write-datatable-direct | MiniExcel | 97.14 ms | 4.02 ms | 2.32 ms | 2.23 | 2.44 | 92394.2 KB | 7.27 | 1521.1 KB | 1.10 | 122.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-direct | ClosedXML | 340.09 ms | 12.83 ms | 7.41 ms | 7.80 | 8.54 | 104205.0 KB | 8.19 | 1139.9 KB | 0.82 | 679.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-direct | EPPlus | 387.77 ms | 23.56 ms | 13.60 ms | 8.89 | 9.73 | 117437.6 KB | 9.24 | 1090.8 KB | 0.79 | 788.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | OfficeIMO.Excel | 46.22 ms | 5.35 ms | 3.09 ms | 1.00 | 1.00 | 12733.8 KB | 1.00 | 1385.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | MiniExcel | 113.39 ms | 15.93 ms | 9.20 ms | 2.45 | 2.45 | 92394.7 KB | 7.26 | 1521.1 KB | 1.10 | 145.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | EPPlus | 420.29 ms | 17.58 ms | 10.15 ms | 9.09 | 9.09 | 117437.6 KB | 9.22 | 1090.8 KB | 0.79 | 809.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | ClosedXML | 499.93 ms | 44.72 ms | 25.82 ms | 10.82 | 10.82 | 173402.7 KB | 13.62 | 1140.7 KB | 0.82 | 981.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | LargeXlsx | 32.53 ms | 1.94 ms | 1.12 ms | 0.78 | 1.00 | 5614.1 KB | 0.43 | 1386.5 KB | 1.00 | 22.5% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 41.97 ms | 5.01 ms | 2.89 ms | 1.00 | 1.29 | 12912.0 KB | 1.00 | 1384.9 KB | 1.00 | Loss +29.0% |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | MiniExcel | 84.91 ms | 7.66 ms | 4.42 ms | 2.02 | 2.61 | 93257.1 KB | 7.22 | 1521.1 KB | 1.10 | 102.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | ClosedXML | 339.85 ms | 11.35 ms | 6.55 ms | 8.10 | 10.45 | 104205.0 KB | 8.07 | 1139.9 KB | 0.82 | 709.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | EPPlus | 390.85 ms | 4.48 ms | 2.59 ms | 9.31 | 12.02 | 117438.0 KB | 9.10 | 1090.8 KB | 0.79 | 831.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 54.71 ms | 8.44 ms | 4.87 ms | 1.00 | 1.00 | 11501.6 KB | 1.00 | 1428.5 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 629.14 ms | 91.49 ms | 52.82 ms | 11.50 | 11.50 | 159742.5 KB | 13.89 | 1091.0 KB | 0.76 | 1050.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 994.61 ms | 95.19 ms | 54.96 ms | 18.18 | 18.18 | 496956.9 KB | 43.21 | 1140.1 KB | 0.80 | 1718.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | LargeXlsx | 35.13 ms | 3.51 ms | 2.02 ms | 0.82 | 1.00 | 5614.1 KB | 0.49 | 1386.5 KB | 0.97 | 18.0% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | OfficeIMO.Excel | 42.85 ms | 3.76 ms | 2.17 ms | 1.00 | 1.22 | 11493.8 KB | 1.00 | 1428.4 KB | 1.00 | Loss +22.0% |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | MiniExcel | 89.61 ms | 4.60 ms | 2.66 ms | 2.09 | 2.55 | 93257.1 KB | 8.11 | 1521.0 KB | 1.06 | 109.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | ClosedXML | 364.78 ms | 3.99 ms | 2.30 ms | 8.51 | 10.38 | 104205.0 KB | 9.07 | 1139.9 KB | 0.80 | 751.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | EPPlus | 435.73 ms | 11.70 ms | 6.75 ms | 10.17 | 12.40 | 117437.6 KB | 10.22 | 1090.8 KB | 0.76 | 916.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 59.48 ms | 6.18 ms | 3.57 ms | 1.00 | 1.00 | 10187.2 KB | 1.00 | 1385.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 477.41 ms | 15.95 ms | 9.21 ms | 8.03 | 8.03 | 159742.5 KB | 15.68 | 1091.0 KB | 0.79 | 702.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 926.61 ms | 35.20 ms | 20.32 ms | 15.58 | 15.58 | 496956.9 KB | 48.78 | 1140.1 KB | 0.82 | 1457.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 36.21 ms | 1.33 ms | 0.77 ms | 0.70 | 1.00 | 5614.1 KB | 0.55 | 1386.5 KB | 1.00 | 30.3% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 51.95 ms | 1.89 ms | 1.09 ms | 1.00 | 1.43 | 10179.4 KB | 1.00 | 1384.9 KB | 1.00 | Loss +43.5% |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | MiniExcel | 84.91 ms | 5.04 ms | 2.91 ms | 1.63 | 2.34 | 93257.1 KB | 9.16 | 1521.1 KB | 1.10 | 63.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | ClosedXML | 345.72 ms | 6.05 ms | 3.49 ms | 6.66 | 9.55 | 104205.0 KB | 10.24 | 1139.9 KB | 0.82 | 565.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | EPPlus | 424.55 ms | 12.17 ms | 7.03 ms | 8.17 | 11.73 | 117437.6 KB | 11.54 | 1090.8 KB | 0.79 | 717.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | LargeXlsx | 38.66 ms | 0.86 ms | 0.50 ms | 0.62 | 1.00 | 5614.1 KB | 0.36 | 1386.5 KB | 0.97 | 38.1% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | OfficeIMO.Excel | 62.41 ms | 4.17 ms | 2.41 ms | 1.00 | 1.61 | 15791.7 KB | 1.00 | 1428.4 KB | 1.00 | Loss +61.4% |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | MiniExcel | 90.62 ms | 2.09 ms | 1.21 ms | 1.45 | 2.34 | 93257.1 KB | 5.91 | 1521.1 KB | 1.06 | 45.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | ClosedXML | 348.79 ms | 12.70 ms | 7.33 ms | 5.59 | 9.02 | 104205.0 KB | 6.60 | 1139.9 KB | 0.80 | 458.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | EPPlus | 422.81 ms | 9.20 ms | 5.31 ms | 6.78 | 10.94 | 117437.6 KB | 7.44 | 1090.8 KB | 0.76 | 577.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 46.77 ms | 2.12 ms | 1.23 ms | 1.00 | 1.00 | 11500.9 KB | 1.00 | 1428.5 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 424.52 ms | 9.02 ms | 5.21 ms | 9.08 | 9.08 | 138360.7 KB | 12.03 | 1091.0 KB | 0.76 | 807.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 545.26 ms | 22.06 ms | 12.74 ms | 11.66 | 11.66 | 275422.3 KB | 23.95 | 1140.1 KB | 0.80 | 1065.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | LargeXlsx | 51.04 ms | 3.87 ms | 2.23 ms | 0.88 | 1.00 | 6043.9 KB | 0.57 | 1816.3 KB | 0.99 | 11.9% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 57.93 ms | 3.10 ms | 1.79 ms | 1.00 | 1.14 | 10577.2 KB | 1.00 | 1828.0 KB | 1.00 | Loss +13.5% |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | MiniExcel | 109.78 ms | 4.69 ms | 2.71 ms | 1.89 | 2.15 | 113974.3 KB | 10.78 | 1936.7 KB | 1.06 | 89.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | ClosedXML | 488.96 ms | 8.06 ms | 4.65 ms | 8.44 | 9.58 | 179552.5 KB | 16.98 | 1555.2 KB | 0.85 | 744.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | EPPlus | 556.52 ms | 8.72 ms | 5.04 ms | 9.61 | 10.90 | 144920.3 KB | 13.70 | 1473.0 KB | 0.81 | 860.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | LargeXlsx | 49.15 ms | 2.09 ms | 1.20 ms | 0.90 | 1.00 | 6043.9 KB | 0.61 | 1816.3 KB | 0.99 | 9.7% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 54.45 ms | 4.24 ms | 2.45 ms | 1.00 | 1.11 | 9942.2 KB | 1.00 | 1828.0 KB | 1.00 | Loss +10.8% |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | MiniExcel | 108.65 ms | 4.34 ms | 2.50 ms | 2.00 | 2.21 | 113974.5 KB | 11.46 | 1936.7 KB | 1.06 | 99.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | ClosedXML | 473.38 ms | 11.16 ms | 6.44 ms | 8.69 | 9.63 | 179552.5 KB | 18.06 | 1555.2 KB | 0.85 | 769.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | EPPlus | 538.07 ms | 26.71 ms | 15.42 ms | 9.88 | 10.95 | 144920.3 KB | 14.58 | 1473.0 KB | 0.81 | 888.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 271.73 ms | 10.10 ms | 5.83 ms | 1.00 | 1.00 | 36150.1 KB | 1.00 | 6725.6 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | LargeXlsx | 275.39 ms | 11.57 ms | 6.68 ms | 1.01 | 1.01 | 23211.4 KB | 0.64 | 6614.8 KB | 0.98 | Tie vs OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | MiniExcel | 454.22 ms | 18.93 ms | 10.93 ms | 1.67 | 1.67 | 347925.7 KB | 9.62 | 6949.8 KB | 1.03 | 67.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | ClosedXML | 1454.54 ms | 49.74 ms | 28.72 ms | 5.35 | 5.35 | 487446.6 KB | 13.48 | 6165.9 KB | 0.92 | 435.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | EPPlus | 1869.70 ms | 20.52 ms | 11.85 ms | 6.88 | 6.88 | 562916.4 KB | 15.57 | 5441.6 KB | 0.81 | 588.1% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | OfficeIMO.Excel | 85.60 ms | 5.68 ms | 3.28 ms | 1.00 | 1.00 | 15708.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | EPPlus | 460.13 ms | 3.66 ms | 2.11 ms | 5.38 | 5.38 | 250950.0 KB | 15.98 |  |  | 437.5% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | EPPlus 4.5.3.3 | 779.69 ms |  |  | 9.11 | 9.11 |  |  |  |  | 810.9% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | ClosedXML | 1461.51 ms | 93.10 ms | 53.75 ms | 17.07 | 17.07 | 829722.0 KB | 52.82 |  |  | 1607.4% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | OfficeIMO.Excel | 234.38 ms | 10.96 ms | 6.33 ms | 1.00 | 1.00 | 145134.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus 4.5.3.3 | 307.49 ms |  |  | 1.31 | 1.31 |  |  |  |  | 31.2% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus | 405.18 ms | 25.94 ms | 14.98 ms | 1.73 | 1.73 | 234783.8 KB | 1.62 |  |  | 72.9% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | OfficeIMO.Excel | 93.06 ms | 1.06 ms | 0.61 ms | 1.00 | 1.00 | 55262.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus | 434.66 ms | 20.75 ms | 11.98 ms | 4.67 | 4.67 | 277077.7 KB | 5.01 |  |  | 367.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus 4.5.3.3 | 595.63 ms |  |  | 6.40 | 6.40 |  |  |  |  | 540.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | OfficeIMO.Excel | 356.27 ms | 74.06 ms | 42.76 ms | 1.00 | 1.00 | 156857.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus | 528.93 ms | 39.66 ms | 22.90 ms | 1.48 | 1.48 | 302761.3 KB | 1.93 |  |  | 48.5% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus 4.5.3.3 | 1017.06 ms |  |  | 2.85 | 2.85 |  |  |  |  | 185.5% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | OfficeIMO.Excel | 320.89 ms | 28.37 ms | 16.38 ms | 1.00 | 1.00 | 145151.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus | 515.92 ms | 5.59 ms | 3.23 ms | 1.61 | 1.61 | 277079.0 KB | 1.91 |  |  | 60.8% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus 4.5.3.3 | 804.51 ms |  |  | 2.51 | 2.51 |  |  |  |  | 150.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | OfficeIMO.Excel | 253.35 ms | 14.99 ms | 8.66 ms | 1.00 | 1.00 | 145205.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus | 463.37 ms | 7.68 ms | 4.43 ms | 1.83 | 1.83 | 277071.7 KB | 1.91 |  |  | 82.9% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus 4.5.3.3 | 622.17 ms |  |  | 2.46 | 2.46 |  |  |  |  | 145.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Object projection | build-object-datatable-dictionaries | OfficeIMO.Excel | 8.95 ms | 0.58 ms | 0.33 ms | 1.00 | 1.00 | 5164.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Object projection | build-object-datatable-typed | OfficeIMO.Excel | 8.15 ms | 0.10 ms | 0.06 ms | 1.00 | 1.00 | 8093.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-cells | OfficeIMO.Excel | 52.17 ms | 5.83 ms | 3.37 ms | 1.00 | 1.00 | 24530.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-cells | EPPlus | 267.12 ms | 8.64 ms | 4.99 ms | 5.12 | 5.12 | 187393.3 KB | 7.64 |  |  | 412.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-cells | ClosedXML | 348.65 ms | 10.21 ms | 5.90 ms | 6.68 | 6.68 | 166520.9 KB | 6.79 |  |  | 568.2% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | OfficeIMO.Excel | 35.19 ms | 1.27 ms | 0.73 ms | 1.00 | 1.00 | 3844.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | EPPlus | 227.30 ms | 9.76 ms | 5.64 ms | 6.46 | 6.46 | 115541.7 KB | 30.05 |  |  | 545.9% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | ClosedXML | 309.41 ms | 4.13 ms | 2.39 ms | 8.79 | 8.79 | 150895.6 KB | 39.25 |  |  | 779.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-range | OfficeIMO.Excel | 49.51 ms | 2.24 ms | 1.29 ms | 1.00 | 1.00 | 24531.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-range | EPPlus | 279.89 ms | 13.03 ms | 7.52 ms | 5.65 | 5.65 | 187393.3 KB | 7.64 |  |  | 465.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-range | ClosedXML | 359.43 ms | 6.25 ms | 3.61 ms | 7.26 | 7.26 | 166525.5 KB | 6.79 |  |  | 626.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-top-range | OfficeIMO.Excel | 0.61 ms | 0.04 ms | 0.02 ms | 1.00 | 1.00 | 285.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-top-range | EPPlus | 226.94 ms | 6.03 ms | 3.48 ms | 373.14 | 373.14 | 105580.3 KB | 370.06 |  |  | 37213.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-top-range | ClosedXML | 331.90 ms | 4.67 ms | 2.70 ms | 545.71 | 545.71 | 149402.4 KB | 523.66 |  |  | 54470.9% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-autofilter | OfficeIMO.Excel | 43.12 ms | 2.69 ms | 1.56 ms | 1.00 | 1.00 | 11494.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-autofilter | EPPlus 4.5.3.3 | 336.57 ms |  |  | 7.81 | 7.81 |  |  |  |  | 680.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-autofilter | ClosedXML | 394.06 ms | 1.89 ms | 1.09 ms | 9.14 | 9.14 | 210663.8 KB | 18.33 |  |  | 814.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-autofilter | EPPlus | 433.24 ms | 15.46 ms | 8.92 ms | 10.05 | 10.05 | 211871.8 KB | 18.43 |  |  | 904.9% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-charts | OfficeIMO.Excel | 44.45 ms | 5.13 ms | 2.96 ms | 1.00 | 1.00 | 12553.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-charts | EPPlus 4.5.3.3 | 288.53 ms |  |  | 6.49 | 6.49 |  |  |  |  | 549.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-charts | EPPlus | 452.20 ms | 25.29 ms | 14.60 ms | 10.17 | 10.17 | 214906.2 KB | 17.12 |  |  | 917.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | OfficeIMO.Excel | 42.96 ms | 3.34 ms | 1.93 ms | 1.00 | 1.00 | 11560.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | EPPlus 4.5.3.3 | 305.31 ms |  |  | 7.11 | 7.11 |  |  |  |  | 610.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | ClosedXML | 373.78 ms | 14.54 ms | 8.39 ms | 8.70 | 8.70 | 210711.7 KB | 18.23 |  |  | 770.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | EPPlus | 430.02 ms | 16.82 ms | 9.71 ms | 10.01 | 10.01 | 211913.3 KB | 18.33 |  |  | 901.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-data-validation | OfficeIMO.Excel | 43.52 ms | 3.43 ms | 1.98 ms | 1.00 | 1.00 | 11510.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-data-validation | EPPlus 4.5.3.3 | 291.80 ms |  |  | 6.71 | 6.71 |  |  |  |  | 570.5% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-data-validation | ClosedXML | 367.24 ms | 25.70 ms | 14.84 ms | 8.44 | 8.44 | 210672.7 KB | 18.30 |  |  | 743.9% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-data-validation | EPPlus | 428.36 ms | 18.00 ms | 10.39 ms | 9.84 | 9.84 | 211857.8 KB | 18.41 |  |  | 884.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | OfficeIMO.Excel | 41.39 ms | 1.98 ms | 1.15 ms | 1.00 | 1.00 | 11497.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | EPPlus 4.5.3.3 | 317.32 ms |  |  | 7.67 | 7.67 |  |  |  |  | 666.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | ClosedXML | 382.25 ms | 25.69 ms | 14.83 ms | 9.24 | 9.24 | 210646.8 KB | 18.32 |  |  | 823.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | EPPlus | 458.18 ms | 25.22 ms | 14.56 ms | 11.07 | 11.07 | 211883.7 KB | 18.43 |  |  | 1007.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | EPPlus 4.5.3.3 | 295.79 ms |  |  | 0.99 | 1.00 |  |  |  |  | Tie vs OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | OfficeIMO.Excel | 299.58 ms | 5.69 ms | 3.29 ms | 1.00 | 1.01 | 143621.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | EPPlus | 464.34 ms | 5.51 ms | 3.18 ms | 1.55 | 1.57 | 230801.8 KB | 1.61 |  |  | 55.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | OfficeIMO.Excel | 226.59 ms | 10.90 ms | 6.29 ms | 1.00 | 1.00 | 145143.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus | 435.10 ms | 9.77 ms | 5.64 ms | 1.92 | 1.92 | 277079.0 KB | 1.91 |  |  | 92.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus 4.5.3.3 | 644.05 ms |  |  | 2.84 | 2.84 |  |  |  |  | 184.2% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world report | realworld-report-core | OfficeIMO.Excel | 46.34 ms | 3.88 ms | 2.24 ms | 1.00 | 1.00 | 11648.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world report | realworld-report-core | EPPlus | 471.30 ms | 13.87 ms | 8.01 ms | 10.17 | 10.17 | 255066.2 KB | 21.90 |  |  | 917.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world report | realworld-report-core | EPPlus 4.5.3.3 | 736.59 ms |  |  | 15.90 | 15.90 |  |  |  |  | 1489.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world report | realworld-report-core | ClosedXML | 1041.77 ms | 44.25 ms | 25.55 ms | 22.48 | 22.48 | 680116.4 KB | 58.39 |  |  | 2148.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook | OfficeIMO.Excel | 342.77 ms | 14.82 ms | 8.55 ms | 1.00 | 1.00 | 196317.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Report workbook | report-workbook | EPPlus | 571.24 ms | 24.79 ms | 14.31 ms | 1.67 | 1.67 | 364710.3 KB | 1.86 |  |  | 66.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook | EPPlus 4.5.3.3 | 841.42 ms |  |  | 2.45 | 2.45 |  |  |  |  | 145.5% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-core | OfficeIMO.Excel | 51.45 ms | 2.64 ms | 1.53 ms | 1.00 | 1.00 | 10979.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Report workbook | report-workbook-core | EPPlus | 563.86 ms | 4.16 ms | 2.40 ms | 10.96 | 10.96 | 342842.6 KB | 31.23 |  |  | 995.9% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-core | EPPlus 4.5.3.3 | 857.15 ms |  |  | 16.66 | 16.66 |  |  |  |  | 1565.9% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-core | ClosedXML | 1185.69 ms | 14.64 ms | 8.45 ms | 23.04 | 23.04 | 975775.1 KB | 88.87 |  |  | 2204.5% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable | OfficeIMO.Excel | 334.91 ms | 7.78 ms | 4.49 ms | 1.00 | 1.00 | 199103.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable | EPPlus | 560.46 ms | 1.83 ms | 1.06 ms | 1.67 | 1.67 | 247824.3 KB | 1.24 |  |  | 67.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable | EPPlus 4.5.3.3 | 854.84 ms |  |  | 2.55 | 2.55 |  |  |  |  | 155.2% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable-core | OfficeIMO.Excel | 50.73 ms | 2.50 ms | 1.44 ms | 1.00 | 1.00 | 13725.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable-core | EPPlus | 527.51 ms | 31.03 ms | 17.91 ms | 10.40 | 10.40 | 225957.5 KB | 16.46 |  |  | 939.9% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable-core | EPPlus 4.5.3.3 | 872.50 ms |  |  | 17.20 | 17.20 |  |  |  |  | 1620.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable-core | ClosedXML | 1055.08 ms | 26.27 ms | 15.17 ms | 20.80 | 20.80 | 832227.0 KB | 60.64 |  |  | 1979.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | OfficeIMO.Excel | 28.98 ms | 2.74 ms | 1.58 ms | 1.00 | 1.00 | 6219.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | EPPlus 4.5.3.3 | 97.60 ms |  |  | 3.37 | 3.37 |  |  |  |  | 236.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | EPPlus | 193.07 ms | 11.00 ms | 6.35 ms | 6.66 | 6.66 | 70814.6 KB | 11.39 |  |  | 566.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | ClosedXML | 235.57 ms | 17.34 ms | 10.01 ms | 8.13 | 8.13 | 79510.2 KB | 12.78 |  |  | 713.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | OfficeIMO.Excel | 0.86 ms | 0.03 ms | 0.02 ms | 1.00 | 1.00 | 177.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | Sylvan.Data.Excel | 1.22 ms | 0.35 ms | 0.20 ms | 1.42 | 1.42 | 316.6 KB | 1.78 |  |  | 41.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | ExcelDataReader | 1.45 ms | 0.05 ms | 0.03 ms | 1.69 | 1.69 | 4062.2 KB | 22.89 |  |  | 68.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | ClosedXML | 3.71 ms | 0.05 ms | 0.03 ms | 4.31 | 4.31 | 4392.8 KB | 24.75 |  |  | 331.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | MiniExcel | 13.56 ms | 2.39 ms | 1.38 ms | 15.74 | 15.74 | 46194.9 KB | 260.29 |  |  | 1474.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | EPPlus 4.5.3.3 | 28.10 ms |  |  | 32.62 | 32.62 |  |  |  |  | 3162.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | EPPlus | 101.17 ms | 7.04 ms | 4.06 ms | 117.45 | 117.45 | 43071.1 KB | 242.69 |  |  | 11645.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | OfficeIMO.Excel | 0.86 ms | 0.07 ms | 0.04 ms | 1.00 | 1.00 | 177.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | Sylvan.Data.Excel | 1.04 ms | 0.03 ms | 0.02 ms | 1.20 | 1.20 | 316.6 KB | 1.78 |  |  | 20.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | ExcelDataReader | 1.76 ms | 0.52 ms | 0.30 ms | 2.04 | 2.04 | 4062.2 KB | 22.88 |  |  | 103.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | ClosedXML | 4.16 ms | 1.25 ms | 0.72 ms | 4.81 | 4.81 | 4392.8 KB | 24.74 |  |  | 380.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | MiniExcel | 13.58 ms | 1.54 ms | 0.89 ms | 15.70 | 15.70 | 46194.9 KB | 260.17 |  |  | 1470.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | EPPlus 4.5.3.3 | 27.99 ms |  |  | 32.37 | 32.37 |  |  |  |  | 3136.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | EPPlus | 95.81 ms | 3.90 ms | 2.25 ms | 110.81 | 110.81 | 43071.1 KB | 242.58 |  |  | 10980.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | Sylvan.Data.Excel | 26.16 ms | 3.63 ms | 2.09 ms | 0.84 | 1.00 | 1936.7 KB | 0.21 |  |  | 16.4% faster than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | OfficeIMO.Excel | 31.30 ms | 4.76 ms | 2.75 ms | 1.00 | 1.20 | 9218.1 KB | 1.00 |  |  | Loss +19.7% |
| 25000 | speed-comparison | read | Other | shared-string-read | ExcelDataReader | 62.56 ms | 1.56 ms | 0.90 ms | 2.00 | 2.39 | 25020.8 KB | 2.71 |  |  | 99.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | MiniExcel | 64.67 ms | 4.12 ms | 2.38 ms | 2.07 | 2.47 | 74405.3 KB | 8.07 |  |  | 106.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | EPPlus 4.5.3.3 | 111.48 ms |  |  | 3.56 | 4.26 |  |  |  |  | 256.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | EPPlus | 183.69 ms | 12.05 ms | 6.96 ms | 5.87 | 7.02 | 89346.1 KB | 9.69 |  |  | 486.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | ClosedXML | 219.18 ms | 24.53 ms | 14.16 ms | 7.00 | 8.38 | 90414.4 KB | 9.81 |  |  | 600.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | OfficeIMO.Excel | 32.46 ms | 0.48 ms | 0.28 ms | 1.00 | 1.00 | 1122.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | Sylvan.Data.Excel | 38.21 ms | 0.53 ms | 0.31 ms | 1.18 | 1.18 | 3534.8 KB | 3.15 |  |  | 17.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | ExcelDataReader | 104.17 ms | 0.19 ms | 0.11 ms | 3.21 | 3.21 | 61201.9 KB | 54.53 |  |  | 220.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | MiniExcel | 118.85 ms | 3.91 ms | 2.26 ms | 3.66 | 3.66 | 186420.9 KB | 166.10 |  |  | 266.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | EPPlus | 227.00 ms | 8.47 ms | 4.89 ms | 6.99 | 6.99 | 105609.1 KB | 94.10 |  |  | 599.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | ClosedXML | 317.74 ms | 7.04 ms | 4.06 ms | 9.79 | 9.79 | 149387.1 KB | 133.10 |  |  | 878.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | Sylvan.Data.Excel | 58.59 ms | 2.53 ms | 1.46 ms | 0.94 | 1.00 | 18394.2 KB | 0.53 |  |  | 6.2% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | OfficeIMO.Excel | 62.47 ms | 2.26 ms | 1.30 ms | 1.00 | 1.07 | 34645.8 KB | 1.00 |  |  | Loss +6.6% |
| 25000 | speed-comparison | read | Range and table read | read-datatable | ExcelDataReader | 140.43 ms | 4.62 ms | 2.67 ms | 2.25 | 2.40 | 76061.4 KB | 2.20 |  |  | 124.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | MiniExcel | 156.96 ms | 15.81 ms | 9.13 ms | 2.51 | 2.68 | 181285.0 KB | 5.23 |  |  | 151.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | EPPlus 4.5.3.3 | 256.28 ms |  |  | 4.10 | 4.37 |  |  |  |  | 310.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | EPPlus | 278.70 ms | 6.14 ms | 3.54 ms | 4.46 | 4.76 | 202250.3 KB | 5.84 |  |  | 346.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | ClosedXML | 357.61 ms | 30.80 ms | 17.78 ms | 5.72 | 6.10 | 178450.5 KB | 5.15 |  |  | 472.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | OfficeIMO.Excel | 36.67 ms | 2.82 ms | 1.63 ms | 1.00 | 1.00 | 4042.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | Sylvan.Data.Excel | 49.37 ms | 10.89 ms | 6.29 ms | 1.35 | 1.35 | 4316.2 KB | 1.07 |  |  | 34.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | MiniExcel | 97.14 ms | 3.05 ms | 1.76 ms | 2.65 | 2.65 | 158610.2 KB | 39.24 |  |  | 164.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | ExcelDataReader | 115.78 ms | 5.15 ms | 2.98 ms | 3.16 | 3.16 | 61201.9 KB | 15.14 |  |  | 215.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | EPPlus | 237.72 ms | 5.20 ms | 3.00 ms | 6.48 | 6.48 | 115541.7 KB | 28.58 |  |  | 548.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | ClosedXML | 328.79 ms | 16.66 ms | 9.62 ms | 8.97 | 8.97 | 150897.5 KB | 37.33 |  |  | 796.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | Sylvan.Data.Excel | 43.73 ms | 1.15 ms | 0.67 ms | 0.83 | 1.00 | 3534.8 KB | 0.14 |  |  | 17.0% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | OfficeIMO.Excel | 52.68 ms | 0.83 ms | 0.48 ms | 1.00 | 1.20 | 26098.2 KB | 1.00 |  |  | Loss +20.5% |
| 25000 | speed-comparison | read | Range and table read | read-range | ExcelDataReader | 112.35 ms | 2.50 ms | 1.45 ms | 2.13 | 2.57 | 61201.9 KB | 2.35 |  |  | 113.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | MiniExcel | 123.08 ms | 4.40 ms | 2.54 ms | 2.34 | 2.81 | 186421.5 KB | 7.14 |  |  | 133.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | EPPlus | 257.22 ms | 10.78 ms | 6.22 ms | 4.88 | 5.88 | 187390.9 KB | 7.18 |  |  | 388.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | EPPlus 4.5.3.3 | 260.85 ms |  |  | 4.95 | 5.97 |  |  |  |  | 395.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | ClosedXML | 336.19 ms | 25.91 ms | 14.96 ms | 6.38 | 7.69 | 163591.6 KB | 6.27 |  |  | 538.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | Sylvan.Data.Excel | 51.66 ms | 2.83 ms | 1.63 ms | 0.95 | 1.00 | 4484.9 KB | 0.17 |  |  | 4.8% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | OfficeIMO.Excel | 54.25 ms | 2.15 ms | 1.24 ms | 1.00 | 1.05 | 26684.2 KB | 1.00 |  |  | Loss +5.0% |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | ExcelDataReader | 118.38 ms | 7.96 ms | 4.59 ms | 2.18 | 2.29 | 61201.9 KB | 2.29 |  |  | 118.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | MiniExcel | 135.67 ms | 7.19 ms | 4.15 ms | 2.50 | 2.63 | 186421.5 KB | 6.99 |  |  | 150.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | EPPlus | 270.08 ms | 2.90 ms | 1.68 ms | 4.98 | 5.23 | 187390.9 KB | 7.02 |  |  | 397.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | ClosedXML | 366.96 ms | 30.73 ms | 17.74 ms | 6.76 | 7.10 | 163585.9 KB | 6.13 |  |  | 576.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | Sylvan.Data.Excel | 0.51 ms | 0.11 ms | 0.06 ms | 0.88 | 1.00 | 348.5 KB | 1.18 |  |  | 11.8% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | OfficeIMO.Excel | 0.57 ms | 0.02 ms | 0.01 ms | 1.00 | 1.13 | 296.0 KB | 1.00 |  |  | Loss +13.4% |
| 25000 | speed-comparison | read | Range and table read | read-top-range | MiniExcel | 0.79 ms | 0.15 ms | 0.09 ms | 1.38 | 1.57 | 869.0 KB | 2.94 |  |  | 38.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | ExcelDataReader | 40.72 ms | 2.91 ms | 1.68 ms | 70.92 | 80.41 | 17115.3 KB | 57.82 |  |  | 6991.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | EPPlus 4.5.3.3 | 193.50 ms |  |  | 337.00 | 382.09 |  |  |  |  | 33599.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | EPPlus | 244.03 ms | 12.86 ms | 7.42 ms | 425.00 | 481.86 | 105577.8 KB | 356.64 |  |  | 42399.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | ClosedXML | 331.00 ms | 20.22 ms | 11.68 ms | 576.46 | 653.59 | 149390.6 KB | 504.64 |  |  | 57545.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | Sylvan.Data.Excel | 45.12 ms | 2.03 ms | 1.17 ms | 0.52 | 1.00 | 3534.8 KB | 0.10 |  |  | 48.1% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | OfficeIMO.Excel | 86.94 ms | 2.18 ms | 1.26 ms | 1.00 | 1.93 | 34151.6 KB | 1.00 |  |  | Loss +92.7% |
| 25000 | speed-comparison | read | Range and table read | read-used-range | ExcelDataReader | 114.29 ms | 3.88 ms | 2.24 ms | 1.31 | 2.53 | 61201.9 KB | 1.79 |  |  | 31.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | MiniExcel | 127.39 ms | 3.58 ms | 2.07 ms | 1.47 | 2.82 | 186421.5 KB | 5.46 |  |  | 46.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | EPPlus | 273.40 ms | 6.17 ms | 3.56 ms | 3.14 | 6.06 | 187390.9 KB | 5.49 |  |  | 214.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | ClosedXML | 346.12 ms | 0.13 ms | 0.08 ms | 3.98 | 7.67 | 163592.7 KB | 4.79 |  |  | 298.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | OfficeIMO.Excel | 36.09 ms | 3.86 ms | 2.23 ms | 1.00 | 1.00 | 1125.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | Sylvan.Data.Excel | 48.42 ms | 8.00 ms | 4.62 ms | 1.34 | 1.34 | 3534.8 KB | 3.14 |  |  | 34.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | ExcelDataReader | 132.34 ms | 10.15 ms | 5.86 ms | 3.67 | 3.67 | 61201.9 KB | 54.37 |  |  | 266.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | MiniExcel | 136.78 ms | 25.12 ms | 14.50 ms | 3.79 | 3.79 | 186420.9 KB | 165.62 |  |  | 279.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | EPPlus | 333.00 ms | 41.51 ms | 23.96 ms | 9.23 | 9.23 | 105609.1 KB | 93.83 |  |  | 822.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | ClosedXML | 381.75 ms | 48.02 ms | 27.72 ms | 10.58 | 10.58 | 149394.7 KB | 132.73 |  |  | 957.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | Sylvan.Data.Excel | 45.35 ms | 1.84 ms | 1.06 ms | 0.87 | 1.00 | 3534.8 KB | 0.13 |  |  | 12.5% faster than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | OfficeIMO.Excel | 51.84 ms | 1.73 ms | 1.00 ms | 1.00 | 1.14 | 26883.8 KB | 1.00 |  |  | Loss +14.3% |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | ExcelDataReader | 115.16 ms | 6.88 ms | 3.97 ms | 2.22 | 2.54 | 61201.9 KB | 2.28 |  |  | 122.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | MiniExcel | 123.80 ms | 4.67 ms | 2.70 ms | 2.39 | 2.73 | 186421.5 KB | 6.93 |  |  | 138.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | EPPlus 4.5.3.3 | 211.23 ms |  |  | 4.07 | 4.66 |  |  |  |  | 307.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | EPPlus | 257.52 ms | 6.23 ms | 3.60 ms | 4.97 | 5.68 | 187390.9 KB | 6.97 |  |  | 396.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | ClosedXML | 350.82 ms | 10.42 ms | 6.02 ms | 6.77 | 7.74 | 163593.9 KB | 6.09 |  |  | 576.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | Sylvan.Data.Excel | 0.45 ms | 0.07 ms | 0.04 ms | 0.84 | 1.00 | 348.5 KB | 1.16 |  |  | 16.4% faster than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | OfficeIMO.Excel | 0.54 ms | 0.02 ms | 0.01 ms | 1.00 | 1.20 | 299.3 KB | 1.00 |  |  | Loss +19.6% |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | MiniExcel | 1.18 ms | 0.70 ms | 0.41 ms | 2.18 | 2.61 | 869.0 KB | 2.90 |  |  | 117.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | ExcelDataReader | 42.43 ms | 3.45 ms | 1.99 ms | 78.27 | 93.61 | 17115.3 KB | 57.19 |  |  | 7726.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | EPPlus 4.5.3.3 | 192.82 ms |  |  | 355.69 | 425.40 |  |  |  |  | 35468.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | EPPlus | 226.60 ms | 20.43 ms | 11.80 ms | 418.01 | 499.93 | 105577.8 KB | 352.77 |  |  | 41700.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | ClosedXML | 345.53 ms | 48.53 ms | 28.02 ms | 637.39 | 762.31 | 149395.1 KB | 499.18 |  |  | 63639.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | Sylvan.Data.Excel | 0.43 ms | 0.01 ms | 0.01 ms | 0.81 | 1.00 | 348.5 KB | 1.16 |  |  | 19.3% faster than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | OfficeIMO.Excel | 0.54 ms | 0.02 ms | 0.01 ms | 1.00 | 1.24 | 300.1 KB | 1.00 |  |  | Loss +23.9% |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | MiniExcel | 0.72 ms | 0.03 ms | 0.02 ms | 1.33 | 1.65 | 869.0 KB | 2.90 |  |  | 32.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | ExcelDataReader | 40.17 ms | 1.38 ms | 0.80 ms | 74.54 | 92.34 | 17115.3 KB | 57.03 |  |  | 7353.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | EPPlus | 346.90 ms | 146.85 ms | 84.79 ms | 643.72 | 797.54 | 105577.8 KB | 351.82 |  |  | 64272.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | ClosedXML | 358.96 ms | 45.01 ms | 25.98 ms | 666.10 | 825.26 | 149388.8 KB | 497.81 |  |  | 66510.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | Sylvan.Data.Excel | 47.14 ms | 3.45 ms | 1.99 ms | 0.84 | 1.00 | 5805.0 KB | 0.25 |  |  | 16.0% faster than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | OfficeIMO.Excel | 56.11 ms | 5.47 ms | 3.16 ms | 1.00 | 1.19 | 23562.3 KB | 1.00 |  |  | Loss +19.0% |
| 25000 | speed-comparison | read | Typed object read | read-objects | ExcelDataReader | 114.10 ms | 3.54 ms | 2.04 ms | 2.03 | 2.42 | 63472.1 KB | 2.69 |  |  | 103.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | MiniExcel | 138.95 ms | 2.60 ms | 1.50 ms | 2.48 | 2.95 | 183656.4 KB | 7.79 |  |  | 147.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | EPPlus 4.5.3.3 | 239.53 ms |  |  | 4.27 | 5.08 |  |  |  |  | 326.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | EPPlus | 258.29 ms | 11.82 ms | 6.82 ms | 4.60 | 5.48 | 199608.2 KB | 8.47 |  |  | 360.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | ClosedXML | 330.96 ms | 4.30 ms | 2.49 ms | 5.90 | 7.02 | 165542.0 KB | 7.03 |  |  | 489.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | OfficeIMO.Excel | 45.62 ms | 2.36 ms | 1.36 ms | 1.00 | 1.00 | 23367.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | Sylvan.Data.Excel | 46.80 ms | 2.33 ms | 1.34 ms | 1.03 | 1.03 | 5292.6 KB | 0.23 |  |  | 2.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | ExcelDataReader | 107.50 ms | 3.88 ms | 2.24 ms | 2.36 | 2.36 | 62959.7 KB | 2.69 |  |  | 135.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | MiniExcel | 119.80 ms | 1.24 ms | 0.72 ms | 2.63 | 2.63 | 183144.1 KB | 7.84 |  |  | 162.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | EPPlus 4.5.3.3 | 233.78 ms |  |  | 5.12 | 5.12 |  |  |  |  | 412.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | EPPlus | 248.56 ms | 5.99 ms | 3.46 ms | 5.45 | 5.45 | 199412.9 KB | 8.53 |  |  | 444.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | ClosedXML | 322.58 ms | 13.19 ms | 7.62 ms | 7.07 | 7.07 | 165348.4 KB | 7.08 |  |  | 607.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | OfficeIMO.Excel | 42.59 ms | 2.71 ms | 1.57 ms | 1.00 | 1.00 | 13026.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | MiniExcel | 71.88 ms | 1.52 ms | 0.88 ms | 1.69 | 1.69 | 124495.5 KB | 9.56 |  |  | 68.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | EPPlus | 370.44 ms | 8.44 ms | 4.87 ms | 8.70 | 8.70 | 159742.2 KB | 12.26 |  |  | 769.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | EPPlus 4.5.3.3 | 610.52 ms |  |  | 14.33 | 14.33 |  |  |  |  | 1333.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | ClosedXML | 775.61 ms | 26.86 ms | 15.51 ms | 18.21 | 18.21 | 566142.0 KB | 43.46 |  |  | 1721.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | OfficeIMO.Excel | 43.41 ms | 2.97 ms | 1.71 ms | 1.00 | 1.00 | 10300.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | MiniExcel | 89.68 ms | 4.73 ms | 2.73 ms | 2.07 | 2.07 | 128874.9 KB | 12.51 |  |  | 106.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | EPPlus | 522.87 ms | 31.58 ms | 18.23 ms | 12.04 | 12.04 | 195408.4 KB | 18.97 |  |  | 1104.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | ClosedXML | 809.32 ms | 63.10 ms | 36.43 ms | 18.64 | 18.64 | 550095.6 KB | 53.40 |  |  | 1764.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 39.51 ms | 1.39 ms | 0.80 ms | 1.00 | 1.00 | 11501.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 369.72 ms | 11.04 ms | 6.37 ms | 9.36 | 9.36 | 159742.7 KB | 13.89 |  |  | 835.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 680.91 ms | 20.41 ms | 11.78 ms | 17.24 | 17.24 | 496956.9 KB | 43.21 |  |  | 1623.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 42.96 ms | 0.89 ms | 0.51 ms | 1.00 | 1.00 | 10187.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 399.61 ms | 12.29 ms | 7.09 ms | 9.30 | 9.30 | 159742.7 KB | 15.68 |  |  | 830.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 736.55 ms | 32.13 ms | 18.55 ms | 17.14 | 17.14 | 496956.9 KB | 48.78 |  |  | 1614.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 37.23 ms | 5.20 ms | 3.00 ms | 1.00 | 1.00 | 11500.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 359.10 ms | 15.73 ms | 9.08 ms | 9.65 | 9.65 | 138360.7 KB | 12.03 |  |  | 864.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 427.09 ms | 12.20 ms | 7.04 ms | 11.47 | 11.47 | 275422.3 KB | 23.95 |  |  | 1047.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | OfficeIMO.Excel | 12.34 ms | 1.08 ms | 0.62 ms | 1.00 | 1.00 | 6896.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | ClosedXML | 115.57 ms | 4.33 ms | 2.50 ms | 9.36 | 9.36 | 92902.1 KB | 13.47 |  |  | 836.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | EPPlus | 168.93 ms | 1.06 ms | 0.61 ms | 13.69 | 13.69 | 74493.1 KB | 10.80 |  |  | 1268.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | OfficeIMO.Excel | 14.87 ms | 0.08 ms | 0.05 ms | 1.00 | 1.00 | 5970.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | ClosedXML | 106.97 ms | 2.56 ms | 1.48 ms | 7.20 | 7.20 | 84206.7 KB | 14.10 |  |  | 619.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | EPPlus 4.5.3.3 | 123.82 ms |  |  | 8.33 | 8.33 |  |  |  |  | 732.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | EPPlus | 185.49 ms | 5.62 ms | 3.24 ms | 12.48 | 12.48 | 86377.9 KB | 14.47 |  |  | 1147.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | OfficeIMO.Excel | 17.26 ms | 0.25 ms | 0.14 ms | 1.00 | 1.00 | 8332.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | EPPlus 4.5.3.3 | 152.59 ms |  |  | 8.84 | 8.84 |  |  |  |  | 783.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | ClosedXML | 169.43 ms | 6.59 ms | 3.80 ms | 9.81 | 9.81 | 111118.7 KB | 13.34 |  |  | 881.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | EPPlus | 216.54 ms | 4.22 ms | 2.43 ms | 12.54 | 12.54 | 113245.5 KB | 13.59 |  |  | 1154.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | OfficeIMO.Excel | 19.43 ms | 1.25 ms | 0.72 ms | 1.00 | 1.00 | 7416.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | ClosedXML | 146.59 ms | 1.96 ms | 1.13 ms | 7.55 | 7.55 | 105223.9 KB | 14.19 |  |  | 654.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | EPPlus | 221.18 ms | 9.66 ms | 5.58 ms | 11.39 | 11.39 | 106317.3 KB | 14.34 |  |  | 1038.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 17.85 ms | 0.56 ms | 0.32 ms | 1.00 | 1.00 | 7416.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | ClosedXML | 146.67 ms | 3.31 ms | 1.91 ms | 8.22 | 8.22 | 105223.9 KB | 14.19 |  |  | 721.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | EPPlus | 222.85 ms | 13.14 ms | 7.59 ms | 12.48 | 12.48 | 106317.3 KB | 14.34 |  |  | 1148.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | OfficeIMO.Excel | 10.97 ms | 0.68 ms | 0.39 ms | 1.00 | 1.00 | 6144.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | ClosedXML | 103.62 ms | 12.81 ms | 7.40 ms | 9.45 | 9.45 | 82591.3 KB | 13.44 |  |  | 844.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | EPPlus 4.5.3.3 | 126.15 ms |  |  | 11.50 | 11.50 |  |  |  |  | 1050.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | EPPlus | 180.24 ms | 4.12 ms | 2.38 ms | 16.44 | 16.44 | 85127.8 KB | 13.85 |  |  | 1543.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | OfficeIMO.Excel | 26.79 ms | 7.39 ms | 4.27 ms | 1.00 | 1.00 | 7482.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | EPPlus 4.5.3.3 | 137.39 ms |  |  | 5.13 | 5.13 |  |  |  |  | 412.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | ClosedXML | 146.06 ms | 6.71 ms | 3.87 ms | 5.45 | 5.45 | 89323.7 KB | 11.94 |  |  | 445.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | EPPlus | 198.90 ms | 7.63 ms | 4.41 ms | 7.43 | 7.43 | 103800.4 KB | 13.87 |  |  | 642.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | OfficeIMO.Excel | 39.94 ms | 9.65 ms | 5.57 ms | 1.00 | 1.00 | 13039.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | MiniExcel | 86.38 ms | 14.22 ms | 8.21 ms | 2.16 | 2.16 | 97088.3 KB | 7.45 |  |  | 116.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | ClosedXML | 367.90 ms | 23.51 ms | 13.57 ms | 9.21 | 9.21 | 172019.1 KB | 13.19 |  |  | 821.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | EPPlus | 487.20 ms | 20.60 ms | 11.89 ms | 12.20 | 12.20 | 111246.3 KB | 8.53 |  |  | 1119.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | OfficeIMO.Excel | 45.31 ms | 5.97 ms | 3.45 ms | 1.00 | 1.00 | 13458.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | MiniExcel | 122.85 ms | 6.83 ms | 3.94 ms | 2.71 | 2.71 | 108129.1 KB | 8.03 |  |  | 171.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | ClosedXML | 561.97 ms | 31.78 ms | 18.35 ms | 12.40 | 12.40 | 280371.8 KB | 20.83 |  |  | 1140.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | EPPlus | 681.54 ms | 147.05 ms | 84.90 ms | 15.04 | 15.04 | 135724.0 KB | 10.08 |  |  | 1404.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | OfficeIMO.Excel | 38.33 ms | 2.20 ms | 1.27 ms | 1.00 | 1.00 | 10288.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | MiniExcel | 85.48 ms | 7.99 ms | 4.61 ms | 2.23 | 2.23 | 97085.4 KB | 9.44 |  |  | 123.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | EPPlus 4.5.3.3 | 337.04 ms |  |  | 8.79 | 8.79 |  |  |  |  | 779.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | ClosedXML | 365.94 ms | 16.74 ms | 9.66 ms | 9.55 | 9.55 | 171999.1 KB | 16.72 |  |  | 854.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | EPPlus | 395.10 ms | 8.14 ms | 4.70 ms | 10.31 | 10.31 | 110816.3 KB | 10.77 |  |  | 930.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | OfficeIMO.Excel | 37.14 ms | 1.14 ms | 0.66 ms | 1.00 | 1.00 | 13020.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | MiniExcel | 80.41 ms | 8.13 ms | 4.69 ms | 2.17 | 2.17 | 92200.0 KB | 7.08 |  |  | 116.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | EPPlus 4.5.3.3 | 315.77 ms |  |  | 8.50 | 8.50 |  |  |  |  | 750.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | EPPlus | 326.33 ms | 11.93 ms | 6.89 ms | 8.79 | 8.79 | 117437.6 KB | 9.02 |  |  | 778.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | ClosedXML | 377.19 ms | 13.26 ms | 7.66 ms | 10.16 | 10.16 | 173398.1 KB | 13.32 |  |  | 915.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | LargeXlsx | 36.05 ms | 5.68 ms | 3.28 ms | 0.99 | 1.00 | 9520.4 KB | 0.75 |  |  | Tie vs OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | OfficeIMO.Excel | 36.47 ms | 0.81 ms | 0.47 ms | 1.00 | 1.01 | 12715.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | MiniExcel | 85.97 ms | 3.49 ms | 2.01 ms | 2.36 | 2.38 | 92394.2 KB | 7.27 |  |  | 135.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | ClosedXML | 295.14 ms | 22.36 ms | 12.91 ms | 8.09 | 8.19 | 104205.0 KB | 8.19 |  |  | 709.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | EPPlus 4.5.3.3 | 404.09 ms |  |  | 11.08 | 11.21 |  |  |  |  | 1008.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | EPPlus | 419.23 ms | 23.84 ms | 13.76 ms | 11.50 | 11.63 | 117437.6 KB | 9.24 |  |  | 1049.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | OfficeIMO.Excel | 37.27 ms | 3.03 ms | 1.75 ms | 1.00 | 1.00 | 9999.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | MiniExcel | 83.85 ms | 11.48 ms | 6.63 ms | 2.25 | 2.25 | 89659.2 KB | 8.97 |  |  | 125.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | EPPlus | 321.94 ms | 4.48 ms | 2.59 ms | 8.64 | 8.64 | 114703.4 KB | 11.47 |  |  | 763.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | ClosedXML | 378.33 ms | 13.15 ms | 7.59 ms | 10.15 | 10.15 | 170666.2 KB | 17.07 |  |  | 915.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | OfficeIMO.Excel | 35.86 ms | 0.30 ms | 0.18 ms | 1.00 | 1.00 | 12733.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | MiniExcel | 91.28 ms | 5.48 ms | 3.17 ms | 2.55 | 2.55 | 92394.5 KB | 7.26 |  |  | 154.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | EPPlus 4.5.3.3 | 368.88 ms |  |  | 10.29 | 10.29 |  |  |  |  | 928.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | EPPlus | 396.58 ms | 46.29 ms | 26.73 ms | 11.06 | 11.06 | 117437.6 KB | 9.22 |  |  | 1006.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | ClosedXML | 428.25 ms | 11.40 ms | 6.58 ms | 11.94 | 11.94 | 173395.0 KB | 13.62 |  |  | 1094.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | OfficeIMO.Excel | 41.74 ms | 6.93 ms | 4.00 ms | 1.00 | 1.00 | 11561.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | MiniExcel | 69.58 ms | 4.86 ms | 2.80 ms | 1.67 | 1.67 | 125551.5 KB | 10.86 |  |  | 66.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | EPPlus | 458.95 ms | 17.96 ms | 10.37 ms | 11.00 | 11.00 | 254959.3 KB | 22.05 |  |  | 999.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | EPPlus 4.5.3.3 | 602.19 ms |  |  | 14.43 | 14.43 |  |  |  |  | 1342.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | ClosedXML | 868.51 ms | 31.37 ms | 18.11 ms | 20.81 | 20.81 | 565955.0 KB | 48.95 |  |  | 1980.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | OfficeIMO.Excel | 21.59 ms | 1.92 ms | 1.11 ms | 1.00 | 1.00 | 10112.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | EPPlus 4.5.3.3 | 190.19 ms |  |  | 8.81 | 8.81 |  |  |  |  | 781.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | ClosedXML | 191.72 ms | 5.01 ms | 2.89 ms | 8.88 | 8.88 | 113853.5 KB | 11.26 |  |  | 788.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | EPPlus | 343.57 ms | 17.93 ms | 10.35 ms | 15.92 | 15.92 | 140732.3 KB | 13.92 |  |  | 1491.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-dictionary-objects-table-direct | OfficeIMO.Excel | 47.26 ms | 0.44 ms | 0.26 ms | 1.00 | 1.00 | 15163.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | LargeXlsx | 41.13 ms | 3.81 ms | 2.20 ms | 0.89 | 1.00 | 6043.9 KB | 0.57 |  |  | 11.0% faster than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 46.21 ms | 0.55 ms | 0.32 ms | 1.00 | 1.12 | 10577.2 KB | 1.00 |  |  | Loss +12.3% |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | MiniExcel | 89.58 ms | 5.94 ms | 3.43 ms | 1.94 | 2.18 | 113974.3 KB | 10.78 |  |  | 93.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | ClosedXML | 403.61 ms | 18.68 ms | 10.78 ms | 8.73 | 9.81 | 179552.5 KB | 16.98 |  |  | 773.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | EPPlus | 493.62 ms | 14.92 ms | 8.61 ms | 10.68 | 12.00 | 144920.3 KB | 13.70 |  |  | 968.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | LargeXlsx | 46.42 ms | 5.64 ms | 3.26 ms | 0.95 | 1.00 | 6043.9 KB | 0.61 |  |  | 4.7% faster than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 48.71 ms | 2.56 ms | 1.48 ms | 1.00 | 1.05 | 9942.2 KB | 1.00 |  |  | Loss +4.9% |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | MiniExcel | 89.95 ms | 2.82 ms | 1.63 ms | 1.85 | 1.94 | 113974.3 KB | 11.46 |  |  | 84.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | ClosedXML | 410.67 ms | 25.08 ms | 14.48 ms | 8.43 | 8.85 | 179552.5 KB | 18.06 |  |  | 743.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | EPPlus | 487.10 ms | 26.19 ms | 15.12 ms | 10.00 | 10.49 | 144920.3 KB | 14.58 |  |  | 900.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 216.40 ms | 2.64 ms | 1.52 ms | 1.00 | 1.00 | 36150.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | LargeXlsx | 238.08 ms | 11.05 ms | 6.38 ms | 1.10 | 1.10 | 23211.4 KB | 0.64 |  |  | 10.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | MiniExcel | 368.92 ms | 11.44 ms | 6.61 ms | 1.70 | 1.70 | 347925.7 KB | 9.62 |  |  | 70.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | ClosedXML | 1325.18 ms | 50.03 ms | 28.89 ms | 6.12 | 6.12 | 487446.6 KB | 13.48 |  |  | 512.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | EPPlus | 1710.66 ms | 17.46 ms | 10.08 ms | 7.91 | 7.91 | 562916.4 KB | 15.57 |  |  | 690.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | LargeXlsx | 11.57 ms | 0.16 ms | 0.09 ms | 0.69 | 1.00 | 2771.0 KB | 0.26 |  |  | 31.1% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | OfficeIMO.Excel | 16.80 ms | 1.82 ms | 1.05 ms | 1.00 | 1.45 | 10843.1 KB | 1.00 |  |  | Loss +45.2% |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | MiniExcel | 31.96 ms | 0.95 ms | 0.55 ms | 1.90 | 2.76 | 58242.9 KB | 5.37 |  |  | 90.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | EPPlus 4.5.3.3 | 144.17 ms |  |  | 8.58 | 12.46 |  |  |  |  | 758.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | ClosedXML | 144.22 ms | 8.27 ms | 4.77 ms | 8.59 | 12.47 | 104233.1 KB | 9.61 |  |  | 758.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | EPPlus | 212.19 ms | 11.90 ms | 6.87 ms | 12.63 | 18.34 | 100373.9 KB | 9.26 |  |  | 1163.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 11.45 ms | 1.32 ms | 0.76 ms | 1.00 | 1.00 | 6961.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 11.49 ms | 2.18 ms | 1.26 ms | 1.00 | 1.00 | 3444.4 KB | 0.49 |  |  | Tie vs OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | ClosedXML | 122.84 ms | 5.23 ms | 3.02 ms | 10.73 | 10.73 | 96015.7 KB | 13.79 |  |  | 972.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | EPPlus | 184.56 ms | 6.11 ms | 3.53 ms | 16.12 | 16.12 | 87467.5 KB | 12.56 |  |  | 1512.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | LargeXlsx | 29.47 ms | 1.89 ms | 1.09 ms | 0.85 | 1.00 | 5614.1 KB | 0.35 |  |  | 14.9% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 34.62 ms | 0.87 ms | 0.50 ms | 1.00 | 1.17 | 16036.5 KB | 1.00 |  |  | Loss +17.5% |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | MiniExcel | 63.13 ms | 0.93 ms | 0.53 ms | 1.82 | 2.14 | 93257.1 KB | 5.82 |  |  | 82.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 292.85 ms |  |  | 8.46 | 9.94 |  |  |  |  | 745.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | ClosedXML | 293.21 ms | 13.06 ms | 7.54 ms | 8.47 | 9.95 | 210646.1 KB | 13.14 |  |  | 746.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | EPPlus | 353.97 ms | 12.20 ms | 7.04 ms | 10.22 | 12.01 | 211850.3 KB | 13.21 |  |  | 922.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | OfficeIMO.Excel | 21.86 ms | 4.20 ms | 2.42 ms | 1.00 | 1.00 | 7866.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | ClosedXML | 141.69 ms | 3.26 ms | 1.88 ms | 6.48 | 6.48 | 105223.9 KB | 13.38 |  |  | 548.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | EPPlus | 206.85 ms | 10.70 ms | 6.18 ms | 9.46 | 9.46 | 106317.3 KB | 13.52 |  |  | 846.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | Sylvan.Data.Excel | 27.98 ms | 1.45 ms | 0.84 ms | 0.67 | 1.00 | 5700.3 KB | 0.44 |  |  | 32.8% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | LargeXlsx | 35.55 ms | 2.61 ms | 1.50 ms | 0.85 | 1.27 | 8349.2 KB | 0.64 |  |  | 14.7% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | OfficeIMO.Excel | 41.67 ms | 6.80 ms | 3.93 ms | 1.00 | 1.49 | 13002.3 KB | 1.00 |  |  | Loss +48.9% |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | MiniExcel | 71.40 ms | 5.53 ms | 3.19 ms | 1.71 | 2.55 | 92199.8 KB | 7.09 |  |  | 71.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | EPPlus 4.5.3.3 | 269.37 ms |  |  | 6.46 | 9.63 |  |  |  |  | 546.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | ClosedXML | 275.70 ms | 15.23 ms | 8.79 ms | 6.62 | 9.85 | 104205.0 KB | 8.01 |  |  | 561.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | EPPlus | 317.33 ms | 5.84 ms | 3.37 ms | 7.61 | 11.34 | 117438.0 KB | 9.03 |  |  | 661.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | OfficeIMO.Excel | 34.90 ms | 2.95 ms | 1.70 ms | 1.00 | 1.00 | 9819.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | LargeXlsx | 43.39 ms | 4.43 ms | 2.56 ms | 1.24 | 1.24 | 9265.9 KB | 0.94 |  |  | 24.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | MiniExcel | 106.07 ms | 11.12 ms | 6.42 ms | 3.04 | 3.04 | 108129.1 KB | 11.01 |  |  | 203.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | EPPlus | 599.00 ms | 23.90 ms | 13.80 ms | 17.16 | 17.16 | 135724.0 KB | 13.82 |  |  | 1616.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | ClosedXML | 605.72 ms | 53.73 ms | 31.02 ms | 17.36 | 17.36 | 280371.6 KB | 28.55 |  |  | 1635.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | LargeXlsx | 38.84 ms | 1.07 ms | 0.62 ms | 0.79 | 1.00 | 10795.2 KB | 0.92 |  |  | 20.6% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | OfficeIMO.Excel | 48.91 ms | 5.74 ms | 3.31 ms | 1.00 | 1.26 | 11708.2 KB | 1.00 |  |  | Loss +26.0% |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | MiniExcel | 159.38 ms | 15.83 ms | 9.14 ms | 3.26 | 4.10 | 226875.9 KB | 19.38 |  |  | 225.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | ClosedXML | 937.90 ms | 15.81 ms | 9.13 ms | 19.17 | 24.15 | 759818.4 KB | 64.90 |  |  | 1817.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | OfficeIMO.Excel | 20.03 ms | 1.11 ms | 0.64 ms | 1.00 | 1.00 | 15744.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | MiniExcel | 37.70 ms | 0.30 ms | 0.18 ms | 1.88 | 1.88 | 73760.2 KB | 4.68 |  |  | 88.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | EPPlus 4.5.3.3 | 106.29 ms |  |  | 5.31 | 5.31 |  |  |  |  | 430.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | ClosedXML | 146.81 ms | 7.77 ms | 4.48 ms | 7.33 | 7.33 | 104241.3 KB | 6.62 |  |  | 633.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | EPPlus | 220.49 ms | 12.78 ms | 7.38 ms | 11.01 | 11.01 | 84410.3 KB | 5.36 |  |  | 1000.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | OfficeIMO.Excel | 17.16 ms | 0.61 ms | 0.35 ms | 1.00 | 1.00 | 15360.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | ClosedXML | 114.32 ms | 0.66 ms | 0.38 ms | 6.66 | 6.66 | 104241.3 KB | 6.79 |  |  | 566.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | EPPlus 4.5.3.3 | 114.75 ms |  |  | 6.69 | 6.69 |  |  |  |  | 568.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | EPPlus | 191.80 ms | 6.45 ms | 3.72 ms | 11.17 | 11.17 | 84410.8 KB | 5.50 |  |  | 1017.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | OfficeIMO.Excel | 12.68 ms | 0.78 ms | 0.45 ms | 1.00 | 1.00 | 13824.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | ClosedXML | 154.91 ms | 8.91 ms | 5.15 ms | 12.22 | 12.22 | 131501.7 KB | 9.51 |  |  | 1121.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | EPPlus | 209.41 ms | 13.34 ms | 7.70 ms | 16.51 | 16.51 | 97730.0 KB | 7.07 |  |  | 1551.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | OfficeIMO.Excel | 12.81 ms | 2.50 ms | 1.45 ms | 1.00 | 1.00 | 7525.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | ClosedXML | 98.22 ms | 4.76 ms | 2.75 ms | 7.67 | 7.67 | 84520.0 KB | 11.23 |  |  | 667.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | EPPlus | 145.99 ms | 0.87 ms | 0.50 ms | 11.40 | 11.40 | 70033.7 KB | 9.31 |  |  | 1040.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | LargeXlsx | 28.88 ms | 0.80 ms | 0.46 ms | 0.82 | 1.00 | 5614.1 KB | 0.43 |  |  | 18.1% faster than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 35.25 ms | 1.75 ms | 1.01 ms | 1.00 | 1.22 | 12912.0 KB | 1.00 |  |  | Loss +22.1% |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | MiniExcel | 67.04 ms | 2.59 ms | 1.49 ms | 1.90 | 2.32 | 93257.1 KB | 7.22 |  |  | 90.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | ClosedXML | 293.48 ms | 19.30 ms | 11.14 ms | 8.33 | 10.16 | 104205.0 KB | 8.07 |  |  | 732.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 308.31 ms |  |  | 8.75 | 10.68 |  |  |  |  | 774.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | EPPlus | 342.29 ms | 9.12 ms | 5.26 ms | 9.71 | 11.85 | 117438.0 KB | 9.10 |  |  | 871.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | LargeXlsx | 28.36 ms | 0.34 ms | 0.20 ms | 0.81 | 1.00 | 5614.1 KB | 0.49 |  |  | 19.5% faster than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | OfficeIMO.Excel | 35.22 ms | 4.83 ms | 2.79 ms | 1.00 | 1.24 | 11493.8 KB | 1.00 |  |  | Loss +24.2% |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | MiniExcel | 69.38 ms | 5.35 ms | 3.09 ms | 1.97 | 2.45 | 93257.1 KB | 8.11 |  |  | 97.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | ClosedXML | 286.17 ms | 22.78 ms | 13.15 ms | 8.13 | 10.09 | 104205.0 KB | 9.07 |  |  | 712.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | EPPlus 4.5.3.3 | 310.27 ms |  |  | 8.81 | 10.94 |  |  |  |  | 781.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | EPPlus | 351.40 ms | 17.21 ms | 9.94 ms | 9.98 | 12.39 | 117437.6 KB | 10.22 |  |  | 897.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 29.60 ms | 1.24 ms | 0.72 ms | 0.75 | 1.00 | 5614.1 KB | 0.55 |  |  | 25.3% faster than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 39.65 ms | 0.62 ms | 0.36 ms | 1.00 | 1.34 | 10179.4 KB | 1.00 |  |  | Loss +34.0% |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | MiniExcel | 68.24 ms | 3.77 ms | 2.18 ms | 1.72 | 2.31 | 93257.1 KB | 9.16 |  |  | 72.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | ClosedXML | 289.15 ms | 5.69 ms | 3.28 ms | 7.29 | 9.77 | 104205.0 KB | 10.24 |  |  | 629.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | EPPlus | 349.78 ms | 5.19 ms | 3.00 ms | 8.82 | 11.82 | 117437.6 KB | 11.54 |  |  | 782.2% slower than OfficeIMO |
