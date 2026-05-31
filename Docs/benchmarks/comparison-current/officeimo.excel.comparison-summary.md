# OfficeIMO.Excel Comparison Summary

This is the suite-level decision table. Mean, standard deviation, standard error, ratios, and allocations come from the lightweight rotated local runner; they are meant for engineering direction. Use the BenchmarkDotNet benchmark classes when a publication-grade `Error` column is required.

## At a glance

| Row count | Artifact | Workload | Category | OfficeIMO wins | OfficeIMO losses | Biggest loss |
| ---: | --- | --- | --- | ---: | ---: | --- |
| 2500 | dense-helloworld-comparison | read | Other | 0 | 2 | dense-helloworld-read-stream: Loss +47.7% vs Sylvan.Data.Excel |
| 2500 | package-profile | package | Package size | 43 | 11 | write-insertobjects-flat-dictionaries-direct: Loss +62.5% vs LargeXlsx |
| 2500 | speed-comparison | mutate | AutoFit and mutation | 2 | 0 |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | 4 | 0 |  |
| 2500 | speed-comparison | other | Object projection | 2 | 0 |  |
| 2500 | speed-comparison | other | Range and table read | 4 | 0 |  |
| 2500 | speed-comparison | other | Real-world feature mix | 6 | 0 |  |
| 2500 | speed-comparison | other | Real-world report | 2 | 0 |  |
| 2500 | speed-comparison | other | Report workbook | 4 | 0 |  |
| 2500 | speed-comparison | read | Formula write/read | 1 | 0 |  |
| 2500 | speed-comparison | read | Other | 1 | 2 | shared-string-read: Loss +85.9% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | read | Range and table read | 4 | 3 | read-used-range: Loss +188.6% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | read | Streaming read | 1 | 3 | read-range-stream: Loss +30.4% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | read | Typed object read | 0 | 2 | read-objects: Loss +9.4% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | write | AutoFit and mutation | 5 | 0 |  |
| 2500 | speed-comparison | write | Cell writer | 7 | 0 |  |
| 2500 | speed-comparison | write | DataSet table export | 3 | 0 |  |
| 2500 | speed-comparison | write | DataTable table export | 4 | 0 |  |
| 2500 | speed-comparison | write | Formatted report write | 1 | 0 |  |
| 2500 | speed-comparison | write | Formula write/read | 1 | 0 |  |
| 2500 | speed-comparison | write | Other | 2 | 2 | write-powershell-psobject-mixed-direct: Loss +31.1% vs LargeXlsx |
| 2500 | speed-comparison | write | Plain cell export | 1 | 3 | append-plain-rows: Loss +63.6% vs LargeXlsx |
| 2500 | speed-comparison | write | Plain streaming export | 1 | 1 | write-datareader-plain: Loss +22.1% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | write | Plain string export | 1 | 0 |  |
| 2500 | speed-comparison | write | Shared string write | 4 | 0 |  |
| 2500 | speed-comparison | write | Typed object export | 0 | 3 | write-insertobjects-flat-dictionaries-direct: Loss +43.6% vs LargeXlsx |
| 25000 | dense-helloworld-comparison | read | Other | 0 | 2 | dense-helloworld-read-stream: Loss +22.2% vs Sylvan.Data.Excel |
| 25000 | package-profile | package | Package size | 42 | 12 | write-insertobjects-legacy-dictionaries-direct: Loss +52.1% vs LargeXlsx |
| 25000 | speed-comparison | mutate | AutoFit and mutation | 1 | 1 | realworld-report-no-autofit: Loss +15.5% vs EPPlus 4.5.3.3 |
| 25000 | speed-comparison | other | Anti-cheat report variants | 4 | 0 |  |
| 25000 | speed-comparison | other | Object projection | 2 | 0 |  |
| 25000 | speed-comparison | other | Range and table read | 4 | 0 |  |
| 25000 | speed-comparison | other | Real-world feature mix | 6 | 0 |  |
| 25000 | speed-comparison | other | Real-world report | 2 | 0 |  |
| 25000 | speed-comparison | other | Report workbook | 4 | 0 |  |
| 25000 | speed-comparison | read | Formula write/read | 1 | 0 |  |
| 25000 | speed-comparison | read | Other | 2 | 1 | shared-string-read: Loss +29.2% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | read | Range and table read | 3 | 4 | read-used-range: Loss +92.6% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | read | Streaming read | 1 | 3 | read-range-stream: Loss +57.9% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | read | Typed object read | 1 | 1 | read-objects: Loss +12.0% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | write | AutoFit and mutation | 5 | 0 |  |
| 25000 | speed-comparison | write | Cell writer | 7 | 0 |  |
| 25000 | speed-comparison | write | DataSet table export | 3 | 0 |  |
| 25000 | speed-comparison | write | DataTable table export | 3 | 1 | write-datatable-direct: Loss +13.5% vs LargeXlsx |
| 25000 | speed-comparison | write | Formatted report write | 1 | 0 |  |
| 25000 | speed-comparison | write | Formula write/read | 1 | 0 |  |
| 25000 | speed-comparison | write | Other | 2 | 2 | write-powershell-psobject-mixed-direct: Loss +13.2% vs LargeXlsx |
| 25000 | speed-comparison | write | Plain cell export | 2 | 2 | append-plain-rows: Loss +28.8% vs LargeXlsx |
| 25000 | speed-comparison | write | Plain streaming export | 1 | 1 | write-datareader-plain: Loss +26.5% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | write | Plain string export | 0 | 1 | write-blog-2023-20-string-columns: Loss +10.8% vs LargeXlsx |
| 25000 | speed-comparison | write | Shared string write | 4 | 0 |  |
| 25000 | speed-comparison | write | Typed object export | 0 | 3 | write-insertobjects-flat-dictionaries-direct: Loss +35.7% vs LargeXlsx |

## OfficeIMO decision table

| Row count | Artifact | Workload | Category | Scenario | OfficeIMO mean | Best | OfficeIMO vs best | Alloc | Package |
| ---: | --- | --- | --- | --- | ---: | --- | ---: | ---: | ---: |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | 5.52 ms | Sylvan.Data.Excel | Loss +41.6% | 2410.8 KB |  |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | 5.58 ms | Sylvan.Data.Excel | Loss +47.7% | 2489.3 KB |  |
| 2500 | package-profile | package | Package size | append-plain-rows | 1.94 ms | LargeXlsx | Loss +46.5% | 1576.3 KB | 63.0 KB |
| 2500 | package-profile | package | Package size | autofit-existing | 8.13 ms | OfficeIMO.Excel | Win | 1895.4 KB | 143.2 KB |
| 2500 | package-profile | package | Package size | large-shared-strings | 1.99 ms | OfficeIMO.Excel | Win | 2440.3 KB | 55.2 KB |
| 2500 | package-profile | package | Package size | realworld-autofilter | 3.79 ms | OfficeIMO.Excel | Win | 1340.4 KB | 142.4 KB |
| 2500 | package-profile | package | Package size | realworld-charts | 4.88 ms | OfficeIMO.Excel | Win | 1892.9 KB | 147.6 KB |
| 2500 | package-profile | package | Package size | realworld-conditional-formatting | 3.85 ms | OfficeIMO.Excel | Win | 1405.8 KB | 142.7 KB |
| 2500 | package-profile | package | Package size | realworld-data-validation | 3.83 ms | OfficeIMO.Excel | Win | 1356.1 KB | 142.5 KB |
| 2500 | package-profile | package | Package size | realworld-freeze-panes | 3.61 ms | OfficeIMO.Excel | Win | 1342.8 KB | 142.5 KB |
| 2500 | package-profile | package | Package size | realworld-pivot-table | 15.31 ms | OfficeIMO.Excel | Win | 14419.5 KB | 200.4 KB |
| 2500 | package-profile | package | Package size | realworld-report-all-in-one | 16.26 ms | OfficeIMO.Excel | Win | 15220.9 KB | 206.5 KB |
| 2500 | package-profile | package | Package size | realworld-report-chart-first | 10.81 ms | OfficeIMO.Excel | Win | 6196.6 KB | 206.5 KB |
| 2500 | package-profile | package | Package size | realworld-report-core | 4.23 ms | OfficeIMO.Excel | Win | 1488.5 KB | 143.9 KB |
| 2500 | package-profile | package | Package size | realworld-report-extra-column | 17.23 ms | OfficeIMO.Excel | Win | 16350.8 KB | 219.1 KB |
| 2500 | package-profile | package | Package size | realworld-report-no-autofit | 16.02 ms | OfficeIMO.Excel | Win | 15209.7 KB | 206.4 KB |
| 2500 | package-profile | package | Package size | realworld-report-post-mutation | 18.81 ms | OfficeIMO.Excel | Win | 15230.4 KB | 206.6 KB |
| 2500 | package-profile | package | Package size | realworld-report-shuffled-columns | 17.37 ms | OfficeIMO.Excel | Win | 15225.4 KB | 211.2 KB |
| 2500 | package-profile | package | Package size | report-workbook | 22.48 ms | OfficeIMO.Excel | Win | 19112.2 KB | 275.6 KB |
| 2500 | package-profile | package | Package size | report-workbook-core | 6.19 ms | OfficeIMO.Excel | Win | 2711.1 KB | 187.5 KB |
| 2500 | package-profile | package | Package size | report-workbook-datatable | 22.24 ms | OfficeIMO.Excel | Win | 19383.9 KB | 275.6 KB |
| 2500 | package-profile | package | Package size | report-workbook-datatable-core | 5.77 ms | OfficeIMO.Excel | Win | 2982.7 KB | 187.5 KB |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | 4.77 ms | OfficeIMO.Excel, LargeXlsx | Win | 1676.8 KB | 216.7 KB |
| 2500 | package-profile | package | Package size | write-bulk-report | 3.92 ms | OfficeIMO.Excel | Win | 1401.7 KB | 143.2 KB |
| 2500 | package-profile | package | Package size | write-cellformula | 2.28 ms | OfficeIMO.Excel | Win | 1383.3 KB | 66.6 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-empty-strings | 2.08 ms | OfficeIMO.Excel | Win | 1787.1 KB | 44.2 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-numbers | 1.85 ms | OfficeIMO.Excel | Win | 1119.9 KB | 47.0 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-object-mixed | 2.34 ms | OfficeIMO.Excel | Win | 1763.3 KB | 61.2 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse | 2.56 ms | OfficeIMO.Excel | Win | 1506.9 KB | 62.0 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | 2.41 ms | OfficeIMO.Excel | Win | 1507.0 KB | 62.0 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-scalars | 1.77 ms | OfficeIMO.Excel | Win | 1138.1 KB | 46.3 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-strings | 2.45 ms | OfficeIMO.Excel | Win | 2617.0 KB | 55.1 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-distinct | 2.03 ms | OfficeIMO.Excel | Win | 2379.2 KB | 51.8 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-repeated | 1.75 ms | OfficeIMO.Excel | Win | 1579.8 KB | 40.0 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-temporal | 2.71 ms | OfficeIMO.Excel | Win | 1435.7 KB | 63.3 KB |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | 1.89 ms | LargeXlsx | Loss +26.5% | 1092.0 KB | 48.2 KB |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | 4.00 ms | LargeXlsx | Loss +18.7% | 2081.1 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-datareader-plain | 4.09 ms | Sylvan.Data.Excel | Loss +29.5% | 1763.0 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-datareader-table | 4.41 ms | OfficeIMO.Excel | Win | 1774.9 KB | 138.8 KB |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | 4.50 ms | OfficeIMO.Excel | Win | 1781.2 KB | 138.9 KB |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | 3.96 ms | OfficeIMO.Excel, LargeXlsx | Win | 2140.6 KB | 131.1 KB |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | 4.64 ms | OfficeIMO.Excel | Win | 2880.2 KB | 176.0 KB |
| 2500 | package-profile | package | Package size | write-dataset-tables | 4.05 ms | OfficeIMO.Excel | Win | 2066.1 KB | 138.9 KB |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | 4.15 ms | OfficeIMO.Excel | Win | 2078.7 KB | 139.2 KB |
| 2500 | package-profile | package | Package size | write-datatable-direct | 4.15 ms | LargeXlsx | Loss +15.8% | 1748.6 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | 4.21 ms | OfficeIMO.Excel | Win | 1760.7 KB | 138.8 KB |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | 3.94 ms | LargeXlsx | Loss +23.9% | 1769.2 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | 4.11 ms | OfficeIMO.Excel | Win | 1347.1 KB | 142.4 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | 3.83 ms | LargeXlsx | Loss +24.9% | 1339.3 KB | 142.3 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | 5.23 ms | OfficeIMO.Excel | Win | 1505.3 KB | 138.1 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | 4.95 ms | LargeXlsx | Loss +62.5% | 1497.5 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | 4.59 ms | LargeXlsx | Loss +47.9% | 1770.1 KB | 142.3 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | 3.61 ms | OfficeIMO.Excel | Win | 1346.4 KB | 142.4 KB |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | 5.57 ms | LargeXlsx | Loss +30.5% | 2341.7 KB | 183.1 KB |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | 4.47 ms | LargeXlsx | Loss +9.0% | 1507.7 KB | 182.4 KB |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | 19.19 ms | OfficeIMO.Excel, LargeXlsx | Win | 4502.3 KB | 651.0 KB |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | 8.38 ms | OfficeIMO.Excel | Win | 1895.1 KB |  |
| 2500 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | 19.21 ms | OfficeIMO.Excel | Win | 15211.2 KB |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | 11.62 ms | OfficeIMO.Excel | Win | 6198.2 KB |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | 40.21 ms | OfficeIMO.Excel | Win | 16352.1 KB |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | 25.82 ms | OfficeIMO.Excel | Win | 15231.8 KB |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | 18.97 ms | OfficeIMO.Excel | Win | 15225.4 KB |  |
| 2500 | speed-comparison | other | Object projection | build-object-datatable-dictionaries | 1.54 ms | OfficeIMO.Excel | Win | 564.2 KB |  |
| 2500 | speed-comparison | other | Object projection | build-object-datatable-typed | 1.31 ms | OfficeIMO.Excel | Win | 856.9 KB |  |
| 2500 | speed-comparison | other | Range and table read | enumerate-cells | 6.35 ms | OfficeIMO.Excel | Win | 2531.6 KB |  |
| 2500 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | 3.99 ms | OfficeIMO.Excel | Win | 523.4 KB |  |
| 2500 | speed-comparison | other | Range and table read | enumerate-range | 6.68 ms | OfficeIMO.Excel | Win | 2531.8 KB |  |
| 2500 | speed-comparison | other | Range and table read | enumerate-top-range | 0.64 ms | OfficeIMO.Excel | Win | 285.3 KB |  |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-autofilter | 4.46 ms | OfficeIMO.Excel | Win | 1340.4 KB |  |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-charts | 5.07 ms | OfficeIMO.Excel | Win | 1892.8 KB |  |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | 4.24 ms | OfficeIMO.Excel | Win | 1405.8 KB |  |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-data-validation | 3.66 ms | OfficeIMO.Excel | Win | 1356.1 KB |  |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | 4.62 ms | OfficeIMO.Excel | Win | 1342.9 KB |  |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | 14.65 ms | OfficeIMO.Excel | Win | 14419.4 KB |  |
| 2500 | speed-comparison | other | Real-world report | realworld-report-all-in-one | 22.00 ms | OfficeIMO.Excel | Win | 15220.5 KB |  |
| 2500 | speed-comparison | other | Real-world report | realworld-report-core | 5.69 ms | OfficeIMO.Excel | Win | 1488.6 KB |  |
| 2500 | speed-comparison | other | Report workbook | report-workbook | 42.69 ms | OfficeIMO.Excel | Win | 19069.6 KB |  |
| 2500 | speed-comparison | other | Report workbook | report-workbook-core | 7.48 ms | OfficeIMO.Excel | Win | 2711.1 KB |  |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable | 30.54 ms | OfficeIMO.Excel | Win | 19383.8 KB |  |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable-core | 7.50 ms | OfficeIMO.Excel | Win | 2982.8 KB |  |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | 2.11 ms | OfficeIMO.Excel | Win | 706.6 KB |  |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | 0.83 ms | OfficeIMO.Excel | Win | 177.2 KB |  |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | 1.69 ms | Sylvan.Data.Excel | Loss +77.0% | 177.2 KB |  |
| 2500 | speed-comparison | read | Other | shared-string-read | 3.62 ms | Sylvan.Data.Excel | Loss +85.9% | 1056.5 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | 4.49 ms | OfficeIMO.Excel | Win | 374.5 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-datatable | 7.03 ms | Sylvan.Data.Excel | Loss +16.8% | 3594.4 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | 3.98 ms | OfficeIMO.Excel | Win | 542.9 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-range | 11.45 ms | OfficeIMO.Excel | Win | 2692.5 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | 5.71 ms | OfficeIMO.Excel | Win | 2751.5 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-top-range | 0.64 ms | Sylvan.Data.Excel | Loss +17.2% | 296.0 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-used-range | 14.12 ms | Sylvan.Data.Excel | Loss +188.6% | 3472.7 KB |  |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | 3.84 ms | OfficeIMO.Excel | Win | 377.7 KB |  |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | 6.47 ms | Sylvan.Data.Excel | Loss +30.4% | 2771.4 KB |  |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | 0.60 ms | Sylvan.Data.Excel | Loss +29.4% | 299.4 KB |  |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | 0.56 ms | Sylvan.Data.Excel | Loss +27.6% | 300.2 KB |  |
| 2500 | speed-comparison | read | Typed object read | read-objects | 7.55 ms | Sylvan.Data.Excel | Loss +9.4% | 2442.0 KB |  |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | 4.72 ms | Sylvan.Data.Excel | Loss +3.8% | 2422.8 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | 5.13 ms | OfficeIMO.Excel | Win | 1781.2 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | 8.68 ms | OfficeIMO.Excel | Win | 2078.7 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | 4.03 ms | OfficeIMO.Excel | Win | 1347.1 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | 5.32 ms | OfficeIMO.Excel | Win | 1505.3 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | 3.98 ms | OfficeIMO.Excel | Win | 1346.4 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | 2.31 ms | OfficeIMO.Excel | Win | 1787.1 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | 2.41 ms | OfficeIMO.Excel | Win | 1119.9 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | 3.00 ms | OfficeIMO.Excel | Win | 1763.3 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | 2.91 ms | OfficeIMO.Excel | Win | 1506.9 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | 2.85 ms | OfficeIMO.Excel | Win | 1507.0 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | 2.48 ms | OfficeIMO.Excel | Win | 1138.1 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | 3.13 ms | OfficeIMO.Excel | Win | 1435.7 KB |  |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | 7.40 ms | OfficeIMO.Excel | Win | 2064.5 KB |  |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | 8.42 ms | OfficeIMO.Excel | Win | 2880.2 KB |  |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | 6.43 ms | OfficeIMO.Excel | Win | 2067.7 KB |  |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | 4.73 ms | OfficeIMO.Excel | Win | 1774.9 KB |  |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | 5.30 ms | OfficeIMO.Excel | Win | 1748.6 KB |  |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | 4.72 ms | OfficeIMO.Excel | Win | 1487.2 KB |  |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | 4.99 ms | OfficeIMO.Excel | Win | 1760.7 KB |  |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | 6.61 ms | OfficeIMO.Excel | Win | 1403.3 KB |  |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | 2.79 ms | OfficeIMO.Excel | Win | 1541.5 KB |  |
| 2500 | speed-comparison | write | Other | write-dictionary-objects-table-direct | 6.84 ms | OfficeIMO.Excel | Win | 2051.4 KB |  |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | 5.62 ms | LargeXlsx | Loss +28.8% | 2341.7 KB |  |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | 5.59 ms | LargeXlsx | Loss +31.1% | 1507.7 KB |  |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | 20.86 ms | LargeXlsx, OfficeIMO.Excel | Win | 4502.3 KB |  |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | 2.47 ms | LargeXlsx | Loss +63.6% | 1576.3 KB |  |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | 1.70 ms | LargeXlsx | Loss +31.0% | 1092.0 KB |  |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | 4.27 ms | LargeXlsx | Loss +20.5% | 2081.1 KB |  |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | 2.71 ms | OfficeIMO.Excel | Win | 1494.0 KB |  |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | 4.47 ms | Sylvan.Data.Excel | Loss +22.1% | 1763.0 KB |  |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | 7.82 ms | OfficeIMO.Excel | Win | 2140.6 KB |  |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | 4.76 ms | OfficeIMO.Excel | Win | 1676.8 KB |  |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | 2.00 ms | OfficeIMO.Excel | Win | 2440.3 KB |  |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | 3.09 ms | OfficeIMO.Excel | Win | 2617.0 KB |  |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | 2.56 ms | OfficeIMO.Excel | Win | 2379.2 KB |  |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | 2.31 ms | OfficeIMO.Excel | Win | 1579.8 KB |  |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | 4.00 ms | LargeXlsx | Loss +24.5% | 1769.2 KB |  |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | 3.56 ms | LargeXlsx | Loss +13.5% | 1339.3 KB |  |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | 4.39 ms | LargeXlsx | Loss +43.6% | 1497.5 KB |  |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | 45.46 ms | Sylvan.Data.Excel | Loss +18.5% | 23621.9 KB |  |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | 44.90 ms | Sylvan.Data.Excel | Loss +22.2% | 24403.9 KB |  |
| 25000 | package-profile | package | Package size | append-plain-rows | 14.59 ms | LargeXlsx | Loss +32.7% | 10842.5 KB | 610.4 KB |
| 25000 | package-profile | package | Package size | autofit-existing | 79.41 ms | OfficeIMO.Excel | Win | 15708.4 KB | 1429.3 KB |
| 25000 | package-profile | package | Package size | large-shared-strings | 14.86 ms | OfficeIMO.Excel | Win | 15744.9 KB | 529.7 KB |
| 25000 | package-profile | package | Package size | realworld-autofilter | 37.72 ms | OfficeIMO.Excel | Win | 11494.9 KB | 1428.5 KB |
| 25000 | package-profile | package | Package size | realworld-charts | 33.59 ms | OfficeIMO.Excel | Win | 12552.7 KB | 1433.7 KB |
| 25000 | package-profile | package | Package size | realworld-conditional-formatting | 34.63 ms | OfficeIMO.Excel | Win | 11560.2 KB | 1428.8 KB |
| 25000 | package-profile | package | Package size | realworld-data-validation | 31.39 ms | OfficeIMO.Excel | Win | 11510.5 KB | 1428.6 KB |
| 25000 | package-profile | package | Package size | realworld-freeze-panes | 32.83 ms | OfficeIMO.Excel | Win | 11497.3 KB | 1428.6 KB |
| 25000 | package-profile | package | Package size | realworld-pivot-table | 226.20 ms | OfficeIMO.Excel | Win | 131923.0 KB | 1979.8 KB |
| 25000 | package-profile | package | Package size | realworld-report-all-in-one | 301.25 ms | OfficeIMO.Excel | Win | 133446.8 KB | 1985.9 KB |
| 25000 | package-profile | package | Package size | realworld-report-chart-first | 161.07 ms | OfficeIMO.Excel | Win | 43563.0 KB | 1985.9 KB |
| 25000 | package-profile | package | Package size | realworld-report-core | 35.75 ms | OfficeIMO.Excel | Win | 11648.7 KB | 1430.0 KB |
| 25000 | package-profile | package | Package size | realworld-report-extra-column | 352.98 ms | OfficeIMO.Excel | Win | 144823.3 KB | 2110.8 KB |
| 25000 | package-profile | package | Package size | realworld-report-no-autofit | 417.93 ms | OfficeIMO.Excel | Win | 133435.3 KB | 1985.8 KB |
| 25000 | package-profile | package | Package size | realworld-report-post-mutation | 243.05 ms | OfficeIMO.Excel | Win | 133455.0 KB | 1986.0 KB |
| 25000 | package-profile | package | Package size | realworld-report-shuffled-columns | 347.07 ms | OfficeIMO.Excel | Win | 133506.8 KB | 2046.1 KB |
| 25000 | package-profile | package | Package size | report-workbook | 327.22 ms | OfficeIMO.Excel | Win | 175194.8 KB | 2672.0 KB |
| 25000 | package-profile | package | Package size | report-workbook-core | 60.71 ms | OfficeIMO.Excel | Win | 10979.4 KB | 1850.9 KB |
| 25000 | package-profile | package | Package size | report-workbook-datatable | 373.27 ms | OfficeIMO.Excel | Win | 177941.8 KB | 2672.0 KB |
| 25000 | package-profile | package | Package size | report-workbook-datatable-core | 49.96 ms | OfficeIMO.Excel | Win | 13725.0 KB | 1850.9 KB |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | 43.72 ms | LargeXlsx | Loss +11.5% | 11708.2 KB | 2228.8 KB |
| 25000 | package-profile | package | Package size | write-bulk-report | 36.78 ms | OfficeIMO.Excel | Win | 11561.8 KB | 1429.3 KB |
| 25000 | package-profile | package | Package size | write-cellformula | 18.30 ms | OfficeIMO.Excel | Win | 10112.0 KB | 670.3 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-empty-strings | 12.16 ms | OfficeIMO.Excel | Win | 6896.4 KB | 451.4 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-numbers | 14.84 ms | OfficeIMO.Excel | Win | 5970.9 KB | 462.6 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-object-mixed | 16.32 ms | OfficeIMO.Excel | Win | 8332.9 KB | 585.1 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse | 18.01 ms | OfficeIMO.Excel | Win | 7416.2 KB | 607.1 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | 18.94 ms | OfficeIMO.Excel | Win | 7416.3 KB | 607.1 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-scalars | 10.96 ms | OfficeIMO.Excel | Win | 6144.6 KB | 441.9 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-strings | 16.52 ms | OfficeIMO.Excel | Win | 15360.4 KB | 527.8 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-distinct | 12.73 ms | OfficeIMO.Excel | Win | 13824.1 KB | 499.0 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-repeated | 11.60 ms | OfficeIMO.Excel | Win | 7525.3 KB | 376.0 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-temporal | 18.80 ms | OfficeIMO.Excel | Win | 7482.8 KB | 620.5 KB |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | 10.50 ms | LargeXlsx | Loss +3.3% | 6961.7 KB | 455.5 KB |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | 35.85 ms | LargeXlsx | Loss +15.1% | 16036.5 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-datareader-plain | 40.13 ms | Sylvan.Data.Excel | Loss +34.9% | 13002.3 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-datareader-table | 38.33 ms | OfficeIMO.Excel | Win | 13020.3 KB | 1385.7 KB |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | 40.84 ms | OfficeIMO.Excel | Win | 13026.6 KB | 1385.9 KB |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | 34.06 ms | OfficeIMO.Excel | Win | 9819.7 KB | 1329.2 KB |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | 63.10 ms | OfficeIMO.Excel | Win | 13458.5 KB | 1795.1 KB |
| 25000 | package-profile | package | Package size | write-dataset-tables | 36.76 ms | OfficeIMO.Excel | Win | 10288.1 KB | 1376.4 KB |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | 41.77 ms | OfficeIMO.Excel | Win | 10300.7 KB | 1376.7 KB |
| 25000 | package-profile | package | Package size | write-datatable-direct | 35.46 ms | LargeXlsx | Loss +8.1% | 12715.7 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | 39.72 ms | OfficeIMO.Excel | Win | 12733.8 KB | 1385.7 KB |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | 33.04 ms | LargeXlsx | Loss +6.9% | 12912.0 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | 35.17 ms | OfficeIMO.Excel | Win | 11501.6 KB | 1428.5 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | 31.08 ms | LargeXlsx | Loss +18.5% | 11493.8 KB | 1428.4 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | 40.26 ms | OfficeIMO.Excel | Win | 10187.2 KB | 1385.1 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | 37.52 ms | LargeXlsx | Loss +37.0% | 10179.4 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | 39.44 ms | LargeXlsx | Loss +52.1% | 15791.7 KB | 1428.4 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | 32.79 ms | OfficeIMO.Excel | Win | 11500.9 KB | 1428.5 KB |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | 42.04 ms | LargeXlsx | Loss +15.9% | 10577.2 KB | 1828.0 KB |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | 43.20 ms | LargeXlsx | Loss +15.1% | 9942.2 KB | 1828.0 KB |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | 221.70 ms | OfficeIMO.Excel | Win | 36150.1 KB | 6725.6 KB |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | 84.43 ms | OfficeIMO.Excel | Win | 15708.3 KB |  |
| 25000 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | 305.85 ms | EPPlus 4.5.3.3 | Loss +15.5% | 133432.9 KB |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | 99.49 ms | OfficeIMO.Excel | Win | 43560.6 KB |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | 370.85 ms | OfficeIMO.Excel | Win | 144825.0 KB |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | 321.76 ms | OfficeIMO.Excel | Win | 133464.0 KB |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | 304.37 ms | OfficeIMO.Excel | Win | 133505.4 KB |  |
| 25000 | speed-comparison | other | Object projection | build-object-datatable-dictionaries | 9.37 ms | OfficeIMO.Excel | Win | 5164.3 KB |  |
| 25000 | speed-comparison | other | Object projection | build-object-datatable-typed | 7.64 ms | OfficeIMO.Excel | Win | 8093.8 KB |  |
| 25000 | speed-comparison | other | Range and table read | enumerate-cells | 56.17 ms | OfficeIMO.Excel | Win | 24530.8 KB |  |
| 25000 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | 42.80 ms | OfficeIMO.Excel | Win | 3839.1 KB |  |
| 25000 | speed-comparison | other | Range and table read | enumerate-range | 56.08 ms | OfficeIMO.Excel | Win | 24530.9 KB |  |
| 25000 | speed-comparison | other | Range and table read | enumerate-top-range | 0.81 ms | OfficeIMO.Excel | Win | 285.2 KB |  |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-autofilter | 32.00 ms | OfficeIMO.Excel | Win | 11494.9 KB |  |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-charts | 34.42 ms | OfficeIMO.Excel | Win | 12551.8 KB |  |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | 36.56 ms | OfficeIMO.Excel | Win | 11560.2 KB |  |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-data-validation | 34.22 ms | OfficeIMO.Excel | Win | 11510.5 KB |  |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | 31.21 ms | OfficeIMO.Excel | Win | 11497.3 KB |  |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | 185.03 ms | OfficeIMO.Excel | Win | 131923.0 KB |  |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | 358.93 ms | OfficeIMO.Excel | Win | 133443.7 KB |  |
| 25000 | speed-comparison | other | Real-world report | realworld-report-core | 35.30 ms | OfficeIMO.Excel | Win | 11648.7 KB |  |
| 25000 | speed-comparison | other | Report workbook | report-workbook | 339.08 ms | OfficeIMO.Excel | Win | 175151.5 KB |  |
| 25000 | speed-comparison | other | Report workbook | report-workbook-core | 49.50 ms | OfficeIMO.Excel | Win | 10979.4 KB |  |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable | 343.99 ms | OfficeIMO.Excel | Win | 177939.1 KB |  |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable-core | 59.82 ms | OfficeIMO.Excel | Win | 13725.0 KB |  |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | 26.43 ms | OfficeIMO.Excel | Win | 6216.3 KB |  |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | 1.04 ms | OfficeIMO.Excel | Win | 179.8 KB |  |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | 0.84 ms | OfficeIMO.Excel | Win | 177.2 KB |  |
| 25000 | speed-comparison | read | Other | shared-string-read | 25.58 ms | Sylvan.Data.Excel | Loss +29.2% | 9217.9 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | 35.65 ms | OfficeIMO.Excel | Win | 1122.3 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-datatable | 73.16 ms | Sylvan.Data.Excel | Loss +12.8% | 34645.7 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | 41.21 ms | OfficeIMO.Excel | Win | 4034.5 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-range | 77.69 ms | Sylvan.Data.Excel | Loss +5.4% | 26098.2 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | 88.39 ms | OfficeIMO.Excel | Win | 26684.1 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-top-range | 0.68 ms | Sylvan.Data.Excel | Loss +38.7% | 296.0 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-used-range | 280.18 ms | Sylvan.Data.Excel | Loss +92.6% | 34151.7 KB |  |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | 32.60 ms | OfficeIMO.Excel | Win | 1125.6 KB |  |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | 72.39 ms | Sylvan.Data.Excel | Loss +57.9% | 26883.7 KB |  |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | 0.54 ms | Sylvan.Data.Excel | Loss +29.9% | 299.3 KB |  |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | 0.56 ms | Sylvan.Data.Excel | Loss +36.8% | 300.0 KB |  |
| 25000 | speed-comparison | read | Typed object read | read-objects | 53.33 ms | Sylvan.Data.Excel | Loss +12.0% | 23562.2 KB |  |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | 45.81 ms | Sylvan.Data.Excel, OfficeIMO.Excel | Win | 23367.3 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | 38.73 ms | OfficeIMO.Excel | Win | 13026.6 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | 36.71 ms | OfficeIMO.Excel | Win | 10300.7 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | 36.06 ms | OfficeIMO.Excel | Win | 11501.6 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | 40.10 ms | OfficeIMO.Excel | Win | 10187.2 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | 33.77 ms | OfficeIMO.Excel | Win | 11500.9 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | 11.80 ms | OfficeIMO.Excel | Win | 6896.4 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | 14.08 ms | OfficeIMO.Excel | Win | 5970.9 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | 17.23 ms | OfficeIMO.Excel | Win | 8332.9 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | 19.01 ms | OfficeIMO.Excel | Win | 7416.2 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | 17.31 ms | OfficeIMO.Excel | Win | 7416.3 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | 11.20 ms | OfficeIMO.Excel | Win | 6144.6 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | 19.09 ms | OfficeIMO.Excel | Win | 7482.8 KB |  |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | 33.84 ms | OfficeIMO.Excel | Win | 13039.6 KB |  |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | 38.39 ms | OfficeIMO.Excel | Win | 13458.5 KB |  |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | 34.56 ms | OfficeIMO.Excel | Win | 10288.1 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | 35.38 ms | OfficeIMO.Excel | Win | 13020.3 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | 34.85 ms | LargeXlsx | Loss +13.5% | 12715.7 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | 34.13 ms | OfficeIMO.Excel | Win | 9999.4 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | 32.76 ms | OfficeIMO.Excel | Win | 12733.8 KB |  |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | 35.64 ms | OfficeIMO.Excel | Win | 11561.8 KB |  |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | 19.84 ms | OfficeIMO.Excel | Win | 10112.0 KB |  |
| 25000 | speed-comparison | write | Other | write-dictionary-objects-table-direct | 45.43 ms | OfficeIMO.Excel | Win | 15163.8 KB |  |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | 46.09 ms | LargeXlsx | Loss +12.1% | 10577.2 KB |  |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | 45.90 ms | LargeXlsx | Loss +13.2% | 9942.2 KB |  |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | 253.66 ms | OfficeIMO.Excel | Win | 36150.1 KB |  |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | 14.17 ms | LargeXlsx | Loss +28.8% | 10842.5 KB |  |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | 14.30 ms | OfficeIMO.Excel, LargeXlsx | Win | 6961.7 KB |  |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | 34.79 ms | LargeXlsx | Loss +15.3% | 16036.5 KB |  |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | 21.76 ms | OfficeIMO.Excel | Win | 7866.1 KB |  |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | 44.10 ms | Sylvan.Data.Excel | Loss +26.5% | 13002.3 KB |  |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | 41.75 ms | OfficeIMO.Excel | Win | 9819.7 KB |  |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | 43.08 ms | LargeXlsx | Loss +10.8% | 11708.2 KB |  |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | 17.77 ms | OfficeIMO.Excel | Win | 15744.9 KB |  |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | 17.93 ms | OfficeIMO.Excel | Win | 15360.4 KB |  |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | 13.01 ms | OfficeIMO.Excel | Win | 13824.1 KB |  |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | 11.70 ms | OfficeIMO.Excel | Win | 7525.3 KB |  |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | 35.11 ms | LargeXlsx | Loss +15.9% | 12912.0 KB |  |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | 32.62 ms | LargeXlsx | Loss +12.3% | 11493.8 KB |  |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | 38.86 ms | LargeXlsx | Loss +35.7% | 10179.4 KB |  |

## Full comparison table

| Row count | Artifact | Workload | Category | Scenario | Library | Mean | StdDev | StdErr | Ratio to OfficeIMO | Ratio to best | Alloc | Alloc ratio | Package | Package ratio | Outcome |
| ---: | --- | --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | Sylvan.Data.Excel | 3.90 ms | 0.28 ms | 0.16 ms | 0.71 | 1.00 | 362.3 KB | 0.15 |  |  | 29.4% faster than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | OfficeIMO.Excel | 5.52 ms | 0.08 ms | 0.05 ms | 1.00 | 1.42 | 2410.8 KB | 1.00 |  |  | Loss +41.6% |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | ExcelDataReader | 11.23 ms | 0.31 ms | 0.18 ms | 2.04 | 2.88 | 6887.4 KB | 2.86 |  |  | 103.6% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | MiniExcel | 15.89 ms | 2.24 ms | 1.29 ms | 2.88 | 4.08 | 21507.3 KB | 8.92 |  |  | 188.0% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | Sylvan.Data.Excel | 3.78 ms | 0.04 ms | 0.02 ms | 0.68 | 1.00 | 362.3 KB | 0.15 |  |  | 32.3% faster than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | OfficeIMO.Excel | 5.58 ms | 0.05 ms | 0.03 ms | 1.00 | 1.48 | 2489.3 KB | 1.00 |  |  | Loss +47.7% |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | ExcelDataReader | 11.02 ms | 0.20 ms | 0.12 ms | 1.98 | 2.92 | 6887.4 KB | 2.77 |  |  | 97.5% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | MiniExcel | 15.79 ms | 1.66 ms | 0.96 ms | 2.83 | 4.18 | 21507.3 KB | 8.64 |  |  | 183.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | append-plain-rows | LargeXlsx | 1.32 ms | 0.04 ms | 0.02 ms | 0.68 | 1.00 | 296.4 KB | 0.19 | 63.1 KB | 1.00 | 31.7% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | append-plain-rows | OfficeIMO.Excel | 1.94 ms | 0.06 ms | 0.03 ms | 1.00 | 1.46 | 1576.3 KB | 1.00 | 63.0 KB | 1.00 | Loss +46.5% |
| 2500 | package-profile | package | Package size | append-plain-rows | MiniExcel | 4.13 ms | 0.15 ms | 0.09 ms | 2.13 | 3.12 | 19710.7 KB | 12.50 | 68.1 KB | 1.08 | 112.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | append-plain-rows | ClosedXML | 14.60 ms | 0.50 ms | 0.29 ms | 7.52 | 11.02 | 11197.4 KB | 7.10 | 59.8 KB | 0.95 | 652.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | append-plain-rows | EPPlus | 25.32 ms | 0.62 ms | 0.36 ms | 13.05 | 19.12 | 14365.2 KB | 9.11 | 56.9 KB | 0.90 | 1205.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | autofit-existing | OfficeIMO.Excel | 8.13 ms | 0.17 ms | 0.10 ms | 1.00 | 1.00 | 1895.4 KB | 1.00 | 143.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | autofit-existing | EPPlus | 74.38 ms | 1.99 ms | 1.15 ms | 9.15 | 9.15 | 50712.0 KB | 26.76 | 115.0 KB | 0.80 | 815.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | autofit-existing | ClosedXML | 130.45 ms | 4.56 ms | 2.64 ms | 16.05 | 16.05 | 84562.5 KB | 44.62 | 121.0 KB | 0.84 | 1504.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | large-shared-strings | OfficeIMO.Excel | 1.99 ms | 0.06 ms | 0.03 ms | 1.00 | 1.00 | 2440.3 KB | 1.00 | 55.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | large-shared-strings | MiniExcel | 3.97 ms | 0.30 ms | 0.17 ms | 2.00 | 2.00 | 21137.5 KB | 8.66 | 60.7 KB | 1.10 | 100.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | large-shared-strings | ClosedXML | 12.15 ms | 1.05 ms | 0.61 ms | 6.11 | 6.11 | 11299.2 KB | 4.63 | 50.3 KB | 0.91 | 511.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | large-shared-strings | EPPlus | 22.24 ms | 0.54 ms | 0.31 ms | 11.20 | 11.20 | 12804.4 KB | 5.25 | 48.1 KB | 0.87 | 1019.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-autofilter | OfficeIMO.Excel | 3.79 ms | 0.55 ms | 0.32 ms | 1.00 | 1.00 | 1340.4 KB | 1.00 | 142.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-autofilter | ClosedXML | 29.97 ms | 0.80 ms | 0.46 ms | 7.90 | 7.90 | 22226.8 KB | 16.58 | 120.2 KB | 0.84 | 690.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-autofilter | EPPlus | 41.02 ms | 0.96 ms | 0.55 ms | 10.82 | 10.82 | 24715.5 KB | 18.44 | 114.2 KB | 0.80 | 981.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-charts | OfficeIMO.Excel | 4.88 ms | 0.15 ms | 0.09 ms | 1.00 | 1.00 | 1892.9 KB | 1.00 | 147.6 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-charts | EPPlus | 42.93 ms | 1.36 ms | 0.79 ms | 8.79 | 8.79 | 27141.8 KB | 14.34 | 117.0 KB | 0.79 | 779.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-conditional-formatting | OfficeIMO.Excel | 3.85 ms | 0.28 ms | 0.16 ms | 1.00 | 1.00 | 1405.8 KB | 1.00 | 142.7 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-conditional-formatting | ClosedXML | 32.25 ms | 0.64 ms | 0.37 ms | 8.37 | 8.37 | 22273.8 KB | 15.84 | 120.3 KB | 0.84 | 737.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-conditional-formatting | EPPlus | 43.29 ms | 2.29 ms | 1.32 ms | 11.24 | 11.24 | 24757.5 KB | 17.61 | 114.3 KB | 0.80 | 1023.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-data-validation | OfficeIMO.Excel | 3.83 ms | 0.17 ms | 0.10 ms | 1.00 | 1.00 | 1356.1 KB | 1.00 | 142.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-data-validation | ClosedXML | 30.94 ms | 1.39 ms | 0.80 ms | 8.08 | 8.08 | 22247.9 KB | 16.41 | 120.3 KB | 0.84 | 708.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-data-validation | EPPlus | 41.18 ms | 1.43 ms | 0.82 ms | 10.76 | 10.76 | 24701.4 KB | 18.22 | 114.2 KB | 0.80 | 975.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-freeze-panes | OfficeIMO.Excel | 3.61 ms | 0.14 ms | 0.08 ms | 1.00 | 1.00 | 1342.8 KB | 1.00 | 142.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-freeze-panes | ClosedXML | 33.80 ms | 6.02 ms | 3.48 ms | 9.35 | 9.35 | 22222.0 KB | 16.55 | 120.2 KB | 0.84 | 835.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-freeze-panes | EPPlus | 43.01 ms | 3.65 ms | 2.11 ms | 11.90 | 11.90 | 24730.0 KB | 18.42 | 114.3 KB | 0.80 | 1089.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-pivot-table | OfficeIMO.Excel | 15.31 ms | 1.46 ms | 0.84 ms | 1.00 | 1.00 | 14419.5 KB | 1.00 | 200.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-pivot-table | EPPlus | 46.94 ms | 0.35 ms | 0.20 ms | 3.07 | 3.07 | 29537.1 KB | 2.05 | 117.4 KB | 0.59 | 206.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-all-in-one | OfficeIMO.Excel | 16.26 ms | 0.91 ms | 0.52 ms | 1.00 | 1.00 | 15220.9 KB | 1.00 | 206.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-all-in-one | EPPlus | 72.08 ms | 0.83 ms | 0.48 ms | 4.43 | 4.43 | 54594.2 KB | 3.59 | 121.8 KB | 0.59 | 343.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-chart-first | OfficeIMO.Excel | 10.81 ms | 0.09 ms | 0.05 ms | 1.00 | 1.00 | 6196.6 KB | 1.00 | 206.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-chart-first | EPPlus | 71.07 ms | 2.45 ms | 1.41 ms | 6.57 | 6.57 | 54593.4 KB | 8.81 | 121.8 KB | 0.59 | 557.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-core | OfficeIMO.Excel | 4.23 ms | 0.07 ms | 0.04 ms | 1.00 | 1.00 | 1488.5 KB | 1.00 | 143.9 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-core | EPPlus | 66.83 ms | 2.46 ms | 1.42 ms | 15.79 | 15.79 | 47299.8 KB | 31.78 | 115.6 KB | 0.80 | 1478.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-core | ClosedXML | 81.88 ms | 3.80 ms | 2.19 ms | 19.34 | 19.34 | 69836.4 KB | 46.92 | 121.5 KB | 0.84 | 1834.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-extra-column | OfficeIMO.Excel | 17.23 ms | 1.26 ms | 0.73 ms | 1.00 | 1.00 | 16350.8 KB | 1.00 | 219.1 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-extra-column | EPPlus | 78.31 ms | 7.37 ms | 4.25 ms | 4.55 | 4.55 | 59225.9 KB | 3.62 | 128.4 KB | 0.59 | 354.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-no-autofit | OfficeIMO.Excel | 16.02 ms | 0.58 ms | 0.34 ms | 1.00 | 1.00 | 15209.7 KB | 1.00 | 206.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-no-autofit | EPPlus | 50.30 ms | 3.17 ms | 1.83 ms | 3.14 | 3.14 | 32906.1 KB | 2.16 | 121.8 KB | 0.59 | 214.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-post-mutation | OfficeIMO.Excel | 18.81 ms | 3.24 ms | 1.87 ms | 1.00 | 1.00 | 15230.4 KB | 1.00 | 206.6 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-post-mutation | EPPlus | 79.51 ms | 5.64 ms | 3.25 ms | 4.23 | 4.23 | 54594.0 KB | 3.58 | 121.9 KB | 0.59 | 322.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-shuffled-columns | OfficeIMO.Excel | 17.37 ms | 1.01 ms | 0.58 ms | 1.00 | 1.00 | 15225.4 KB | 1.00 | 211.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-shuffled-columns | EPPlus | 69.76 ms | 0.90 ms | 0.52 ms | 4.02 | 4.02 | 54590.6 KB | 3.59 | 124.3 KB | 0.59 | 301.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook | OfficeIMO.Excel | 22.48 ms | 0.73 ms | 0.42 ms | 1.00 | 1.00 | 19112.2 KB | 1.00 | 275.6 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | report-workbook | EPPlus | 90.39 ms | 3.47 ms | 2.00 ms | 4.02 | 4.02 | 77485.3 KB | 4.05 | 161.8 KB | 0.59 | 302.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook-core | OfficeIMO.Excel | 6.19 ms | 0.24 ms | 0.14 ms | 1.00 | 1.00 | 2711.1 KB | 1.00 | 187.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | report-workbook-core | EPPlus | 96.47 ms | 5.02 ms | 2.90 ms | 15.60 | 15.60 | 71970.6 KB | 26.55 | 157.2 KB | 0.84 | 1459.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook-core | ClosedXML | 102.22 ms | 1.67 ms | 0.97 ms | 16.53 | 16.53 | 97220.0 KB | 35.86 | 165.1 KB | 0.88 | 1552.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook-datatable | OfficeIMO.Excel | 22.24 ms | 0.43 ms | 0.25 ms | 1.00 | 1.00 | 19383.9 KB | 1.00 | 275.6 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | report-workbook-datatable | EPPlus | 95.47 ms | 0.96 ms | 0.56 ms | 4.29 | 4.29 | 65994.5 KB | 3.40 | 161.8 KB | 0.59 | 329.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook-datatable-core | OfficeIMO.Excel | 5.77 ms | 0.20 ms | 0.12 ms | 1.00 | 1.00 | 2982.7 KB | 1.00 | 187.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | report-workbook-datatable-core | EPPlus | 89.29 ms | 2.04 ms | 1.18 ms | 15.47 | 15.47 | 60480.1 KB | 20.28 | 157.2 KB | 0.84 | 1446.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook-datatable-core | ClosedXML | 96.65 ms | 1.75 ms | 1.01 ms | 16.74 | 16.74 | 82860.8 KB | 27.78 | 165.1 KB | 0.88 | 1574.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | OfficeIMO.Excel | 4.77 ms | 0.21 ms | 0.12 ms | 1.00 | 1.00 | 1676.8 KB | 1.00 | 216.7 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | LargeXlsx | 4.83 ms | 1.15 ms | 0.67 ms | 1.01 | 1.01 | 857.6 KB | 0.51 | 237.7 KB | 1.10 | Tie vs OfficeIMO |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | MiniExcel | 18.27 ms | 1.64 ms | 0.95 ms | 3.83 | 3.83 | 35918.9 KB | 21.42 | 235.3 KB | 1.09 | 283.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | ClosedXML | 91.84 ms | 5.64 ms | 3.25 ms | 19.27 | 19.27 | 71478.2 KB | 42.63 | 257.2 KB | 1.19 | 1827.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-bulk-report | OfficeIMO.Excel | 3.92 ms | 0.11 ms | 0.07 ms | 1.00 | 1.00 | 1401.7 KB | 1.00 | 143.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-bulk-report | MiniExcel | 8.71 ms | 1.29 ms | 0.75 ms | 2.22 | 2.22 | 26825.3 KB | 19.14 | 153.8 KB | 1.07 | 122.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-bulk-report | EPPlus | 66.39 ms | 2.85 ms | 1.64 ms | 16.92 | 16.92 | 47193.8 KB | 33.67 | 115.0 KB | 0.80 | 1591.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-bulk-report | ClosedXML | 72.27 ms | 4.05 ms | 2.34 ms | 18.42 | 18.42 | 58344.3 KB | 41.62 | 121.0 KB | 0.84 | 1741.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellformula | OfficeIMO.Excel | 2.28 ms | 0.02 ms | 0.01 ms | 1.00 | 1.00 | 1383.3 KB | 1.00 | 66.6 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellformula | ClosedXML | 16.98 ms | 0.21 ms | 0.12 ms | 7.45 | 7.45 | 12039.8 KB | 8.70 | 70.6 KB | 1.06 | 644.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellformula | EPPlus | 35.88 ms | 1.20 ms | 0.69 ms | 15.73 | 15.73 | 18110.5 KB | 13.09 | 62.1 KB | 0.93 | 1473.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-empty-strings | OfficeIMO.Excel | 2.08 ms | 0.10 ms | 0.05 ms | 1.00 | 1.00 | 1787.1 KB | 1.00 | 44.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-empty-strings | ClosedXML | 11.51 ms | 0.19 ms | 0.11 ms | 5.54 | 5.54 | 9959.5 KB | 5.57 | 44.9 KB | 1.02 | 454.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-empty-strings | EPPlus | 21.40 ms | 0.82 ms | 0.48 ms | 10.31 | 10.31 | 11773.0 KB | 6.59 | 42.0 KB | 0.95 | 931.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-numbers | OfficeIMO.Excel | 1.85 ms | 0.01 ms | 0.01 ms | 1.00 | 1.00 | 1119.9 KB | 1.00 | 47.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-numbers | ClosedXML | 11.08 ms | 0.33 ms | 0.19 ms | 6.00 | 6.00 | 9177.1 KB | 8.19 | 45.9 KB | 0.98 | 499.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-numbers | EPPlus | 21.82 ms | 1.45 ms | 0.84 ms | 11.80 | 11.80 | 12895.3 KB | 11.51 | 43.7 KB | 0.93 | 1080.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-mixed | OfficeIMO.Excel | 2.34 ms | 0.03 ms | 0.02 ms | 1.00 | 1.00 | 1763.3 KB | 1.00 | 61.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-object-mixed | ClosedXML | 18.02 ms | 1.03 ms | 0.59 ms | 7.72 | 7.72 | 11887.0 KB | 6.74 | 59.5 KB | 0.97 | 671.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-mixed | EPPlus | 29.09 ms | 3.00 ms | 1.73 ms | 12.46 | 12.46 | 15643.4 KB | 8.87 | 58.9 KB | 0.96 | 1145.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse | OfficeIMO.Excel | 2.56 ms | 0.19 ms | 0.11 ms | 1.00 | 1.00 | 1506.9 KB | 1.00 | 62.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse | ClosedXML | 14.66 ms | 0.40 ms | 0.23 ms | 5.73 | 5.73 | 11296.3 KB | 7.50 | 52.5 KB | 0.85 | 472.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse | EPPlus | 26.12 ms | 1.52 ms | 0.88 ms | 10.21 | 10.21 | 14960.3 KB | 9.93 | 54.2 KB | 0.88 | 920.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 2.41 ms | 0.05 ms | 0.03 ms | 1.00 | 1.00 | 1507.0 KB | 1.00 | 62.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | ClosedXML | 14.07 ms | 0.08 ms | 0.04 ms | 5.84 | 5.84 | 11296.3 KB | 7.50 | 52.5 KB | 0.85 | 484.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | EPPlus | 24.61 ms | 1.40 ms | 0.81 ms | 10.22 | 10.22 | 14960.3 KB | 9.93 | 54.2 KB | 0.88 | 921.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-scalars | OfficeIMO.Excel | 1.77 ms | 0.03 ms | 0.02 ms | 1.00 | 1.00 | 1138.1 KB | 1.00 | 46.3 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-scalars | ClosedXML | 12.12 ms | 1.36 ms | 0.78 ms | 6.84 | 6.84 | 9021.2 KB | 7.93 | 45.4 KB | 0.98 | 584.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-scalars | EPPlus | 21.94 ms | 1.43 ms | 0.82 ms | 12.39 | 12.39 | 12827.5 KB | 11.27 | 42.4 KB | 0.91 | 1138.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings | OfficeIMO.Excel | 2.45 ms | 0.23 ms | 0.13 ms | 1.00 | 1.00 | 2617.0 KB | 1.00 | 55.1 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-strings | ClosedXML | 11.36 ms | 0.39 ms | 0.23 ms | 4.65 | 4.65 | 11299.2 KB | 4.32 | 50.3 KB | 0.91 | 364.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings | EPPlus | 20.94 ms | 0.77 ms | 0.44 ms | 8.56 | 8.56 | 12804.9 KB | 4.89 | 48.1 KB | 0.87 | 756.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-distinct | OfficeIMO.Excel | 2.03 ms | 0.08 ms | 0.04 ms | 1.00 | 1.00 | 2379.2 KB | 1.00 | 51.8 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-distinct | ClosedXML | 16.87 ms | 0.96 ms | 0.55 ms | 8.30 | 8.30 | 13127.1 KB | 5.52 | 61.9 KB | 1.19 | 730.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-distinct | EPPlus | 25.70 ms | 0.79 ms | 0.46 ms | 12.65 | 12.65 | 13893.0 KB | 5.84 | 61.5 KB | 1.19 | 1165.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-repeated | OfficeIMO.Excel | 1.75 ms | 0.07 ms | 0.04 ms | 1.00 | 1.00 | 1579.8 KB | 1.00 | 40.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-repeated | ClosedXML | 10.45 ms | 0.02 ms | 0.01 ms | 5.97 | 5.97 | 9226.5 KB | 5.84 | 38.8 KB | 0.97 | 497.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-repeated | EPPlus | 20.73 ms | 0.95 ms | 0.55 ms | 11.85 | 11.85 | 11332.5 KB | 7.17 | 34.8 KB | 0.87 | 1085.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-temporal | OfficeIMO.Excel | 2.71 ms | 0.30 ms | 0.17 ms | 1.00 | 1.00 | 1435.7 KB | 1.00 | 63.3 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-temporal | ClosedXML | 15.56 ms | 0.88 ms | 0.51 ms | 5.74 | 5.74 | 9711.1 KB | 6.76 | 54.5 KB | 0.86 | 474.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-temporal | EPPlus | 25.15 ms | 2.50 ms | 1.44 ms | 9.28 | 9.28 | 14722.7 KB | 10.25 | 53.1 KB | 0.84 | 828.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 1.49 ms | 0.23 ms | 0.13 ms | 0.79 | 1.00 | 447.0 KB | 0.41 | 47.3 KB | 0.98 | 21.0% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 1.89 ms | 0.24 ms | 0.14 ms | 1.00 | 1.27 | 1092.0 KB | 1.00 | 48.2 KB | 1.00 | Loss +26.5% |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | ClosedXML | 14.29 ms | 2.38 ms | 1.37 ms | 7.57 | 9.57 | 10235.8 KB | 9.37 | 53.0 KB | 1.10 | 656.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | EPPlus | 25.36 ms | 1.63 ms | 0.94 ms | 13.43 | 17.00 | 13052.1 KB | 11.95 | 52.5 KB | 1.09 | 1243.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | LargeXlsx | 3.37 ms | 0.24 ms | 0.14 ms | 0.84 | 1.00 | 758.3 KB | 0.36 | 138.4 KB | 1.00 | 15.8% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 4.00 ms | 0.02 ms | 0.01 ms | 1.00 | 1.19 | 2081.1 KB | 1.00 | 138.0 KB | 1.00 | Loss +18.7% |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | MiniExcel | 7.87 ms | 0.04 ms | 0.02 ms | 1.97 | 2.34 | 23222.1 KB | 11.16 | 153.7 KB | 1.11 | 96.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | ClosedXML | 33.40 ms | 1.25 ms | 0.72 ms | 8.35 | 9.91 | 22221.3 KB | 10.68 | 120.1 KB | 0.87 | 734.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | EPPlus | 43.01 ms | 2.66 ms | 1.54 ms | 10.75 | 12.77 | 24694.0 KB | 11.87 | 114.1 KB | 0.83 | 975.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | Sylvan.Data.Excel | 3.16 ms | 0.23 ms | 0.13 ms | 0.77 | 1.00 | 758.7 KB | 0.43 | 78.5 KB | 0.57 | 22.8% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | OfficeIMO.Excel | 4.09 ms | 0.15 ms | 0.08 ms | 1.00 | 1.30 | 1763.0 KB | 1.00 | 138.0 KB | 1.00 | Loss +29.5% |
| 2500 | package-profile | package | Package size | write-datareader-plain | LargeXlsx | 4.39 ms | 0.76 ms | 0.44 ms | 1.07 | 1.39 | 1032.5 KB | 0.59 | 138.4 KB | 1.00 | 7.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | MiniExcel | 7.73 ms | 0.55 ms | 0.32 ms | 1.89 | 2.45 | 23043.8 KB | 13.07 | 153.6 KB | 1.11 | 89.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | ClosedXML | 28.51 ms | 0.82 ms | 0.48 ms | 6.97 | 9.03 | 11581.0 KB | 6.57 | 120.1 KB | 0.87 | 597.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | EPPlus | 40.25 ms | 3.22 ms | 1.86 ms | 9.84 | 12.75 | 16646.4 KB | 9.44 | 114.9 KB | 0.83 | 884.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table | OfficeIMO.Excel | 4.41 ms | 0.37 ms | 0.22 ms | 1.00 | 1.00 | 1774.9 KB | 1.00 | 138.8 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-datareader-table | MiniExcel | 7.87 ms | 0.80 ms | 0.46 ms | 1.78 | 1.78 | 23044.1 KB | 12.98 | 153.6 KB | 1.11 | 78.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table | ClosedXML | 38.55 ms | 1.72 ms | 0.99 ms | 8.74 | 8.74 | 19007.9 KB | 10.71 | 120.9 KB | 0.87 | 774.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table | EPPlus | 40.92 ms | 4.13 ms | 2.38 ms | 9.28 | 9.28 | 16646.1 KB | 9.38 | 114.9 KB | 0.83 | 827.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | OfficeIMO.Excel | 4.50 ms | 0.08 ms | 0.05 ms | 1.00 | 1.00 | 1781.2 KB | 1.00 | 138.9 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | MiniExcel | 7.94 ms | 0.56 ms | 0.32 ms | 1.77 | 1.77 | 26647.2 KB | 14.96 | 153.8 KB | 1.11 | 76.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | EPPlus | 59.30 ms | 6.72 ms | 3.88 ms | 13.18 | 13.18 | 38343.6 KB | 21.53 | 115.1 KB | 0.83 | 1217.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | ClosedXML | 75.17 ms | 3.02 ms | 1.74 ms | 16.70 | 16.70 | 58361.4 KB | 32.77 | 121.0 KB | 0.87 | 1570.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | OfficeIMO.Excel | 3.96 ms | 0.29 ms | 0.17 ms | 1.00 | 1.00 | 2140.6 KB | 1.00 | 131.1 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | LargeXlsx | 4.00 ms | 0.08 ms | 0.05 ms | 1.01 | 1.01 | 1123.9 KB | 0.53 | 164.2 KB | 1.25 | Tie vs OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | MiniExcel | 10.08 ms | 0.83 ms | 0.48 ms | 2.54 | 2.54 | 29746.9 KB | 13.90 | 180.5 KB | 1.38 | 154.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | ClosedXML | 57.73 ms | 5.80 ms | 3.35 ms | 14.57 | 14.57 | 27410.3 KB | 12.80 | 159.4 KB | 1.22 | 1357.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | EPPlus | 57.77 ms | 5.00 ms | 2.89 ms | 14.58 | 14.58 | 21889.7 KB | 10.23 | 144.5 KB | 1.10 | 1358.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | OfficeIMO.Excel | 4.64 ms | 0.17 ms | 0.10 ms | 1.00 | 1.00 | 2880.2 KB | 1.00 | 176.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | MiniExcel | 9.45 ms | 0.13 ms | 0.07 ms | 2.04 | 2.04 | 29746.9 KB | 10.33 | 180.5 KB | 1.03 | 103.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | ClosedXML | 52.85 ms | 2.59 ms | 1.49 ms | 11.40 | 11.40 | 27409.3 KB | 9.52 | 159.4 KB | 0.91 | 1040.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | EPPlus | 60.78 ms | 2.22 ms | 1.28 ms | 13.11 | 13.11 | 21889.7 KB | 7.60 | 144.5 KB | 0.82 | 1211.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables | OfficeIMO.Excel | 4.05 ms | 0.44 ms | 0.25 ms | 1.00 | 1.00 | 2066.1 KB | 1.00 | 138.9 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-dataset-tables | MiniExcel | 8.28 ms | 0.65 ms | 0.37 ms | 2.05 | 2.05 | 28700.4 KB | 13.89 | 156.4 KB | 1.13 | 104.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables | ClosedXML | 34.47 ms | 1.28 ms | 0.74 ms | 8.51 | 8.51 | 18876.9 KB | 9.14 | 123.4 KB | 0.89 | 751.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables | EPPlus | 38.50 ms | 1.76 ms | 1.01 ms | 9.51 | 9.51 | 18700.6 KB | 9.05 | 116.6 KB | 0.84 | 850.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | OfficeIMO.Excel | 4.15 ms | 0.17 ms | 0.10 ms | 1.00 | 1.00 | 2078.7 KB | 1.00 | 139.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | MiniExcel | 8.49 ms | 1.21 ms | 0.70 ms | 2.04 | 2.04 | 31798.5 KB | 15.30 | 156.6 KB | 1.13 | 104.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | EPPlus | 58.03 ms | 5.40 ms | 3.12 ms | 13.98 | 13.98 | 41455.7 KB | 19.94 | 116.9 KB | 0.84 | 1297.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | ClosedXML | 71.73 ms | 2.75 ms | 1.59 ms | 17.28 | 17.28 | 56708.2 KB | 27.28 | 123.7 KB | 0.89 | 1628.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-direct | LargeXlsx | 3.59 ms | 0.20 ms | 0.12 ms | 0.86 | 1.00 | 1149.0 KB | 0.66 | 138.4 KB | 1.00 | 13.6% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-direct | OfficeIMO.Excel | 4.15 ms | 0.20 ms | 0.11 ms | 1.00 | 1.16 | 1748.6 KB | 1.00 | 138.0 KB | 1.00 | Loss +15.8% |
| 2500 | package-profile | package | Package size | write-datatable-direct | MiniExcel | 7.92 ms | 0.35 ms | 0.20 ms | 1.91 | 2.21 | 23062.5 KB | 13.19 | 153.7 KB | 1.11 | 90.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-direct | ClosedXML | 28.11 ms | 2.32 ms | 1.34 ms | 6.77 | 7.83 | 11581.0 KB | 6.62 | 120.1 KB | 0.87 | 576.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-direct | EPPlus | 37.53 ms | 2.88 ms | 1.66 ms | 9.03 | 10.46 | 16646.1 KB | 9.52 | 114.9 KB | 0.83 | 803.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | OfficeIMO.Excel | 4.21 ms | 0.32 ms | 0.18 ms | 1.00 | 1.00 | 1760.7 KB | 1.00 | 138.8 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | MiniExcel | 8.70 ms | 1.15 ms | 0.66 ms | 2.07 | 2.07 | 23062.8 KB | 13.10 | 153.7 KB | 1.11 | 106.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | EPPlus | 40.58 ms | 0.37 ms | 0.21 ms | 9.64 | 9.64 | 16646.1 KB | 9.45 | 114.9 KB | 0.83 | 863.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | ClosedXML | 43.49 ms | 6.76 ms | 3.90 ms | 10.33 | 10.33 | 19008.3 KB | 10.80 | 120.9 KB | 0.87 | 933.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | LargeXlsx | 3.18 ms | 0.09 ms | 0.05 ms | 0.81 | 1.00 | 758.3 KB | 0.43 | 138.4 KB | 1.00 | 19.3% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 3.94 ms | 0.47 ms | 0.27 ms | 1.00 | 1.24 | 1769.2 KB | 1.00 | 138.0 KB | 1.00 | Loss +23.9% |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | MiniExcel | 8.50 ms | 0.97 ms | 0.56 ms | 2.16 | 2.67 | 23222.1 KB | 13.13 | 153.7 KB | 1.11 | 115.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | ClosedXML | 28.49 ms | 1.36 ms | 0.78 ms | 7.23 | 8.96 | 11581.0 KB | 6.55 | 120.1 KB | 0.87 | 623.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | EPPlus | 38.50 ms | 2.11 ms | 1.22 ms | 9.77 | 12.11 | 16646.4 KB | 9.41 | 114.9 KB | 0.83 | 877.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.11 ms | 0.36 ms | 0.21 ms | 1.00 | 1.00 | 1347.1 KB | 1.00 | 142.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 55.35 ms | 2.31 ms | 1.33 ms | 13.46 | 13.46 | 38343.9 KB | 28.46 | 115.1 KB | 0.81 | 1246.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 63.84 ms | 1.52 ms | 0.88 ms | 15.53 | 15.53 | 50927.5 KB | 37.80 | 120.2 KB | 0.84 | 1453.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | LargeXlsx | 3.07 ms | 0.01 ms | 0.01 ms | 0.80 | 1.00 | 758.3 KB | 0.57 | 138.4 KB | 0.97 | 19.9% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | OfficeIMO.Excel | 3.83 ms | 0.47 ms | 0.27 ms | 1.00 | 1.25 | 1339.3 KB | 1.00 | 142.3 KB | 1.00 | Loss +24.9% |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | MiniExcel | 8.43 ms | 0.56 ms | 0.32 ms | 2.20 | 2.74 | 23222.1 KB | 17.34 | 153.7 KB | 1.08 | 119.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | ClosedXML | 29.38 ms | 1.85 ms | 1.07 ms | 7.66 | 9.57 | 11581.0 KB | 8.65 | 120.1 KB | 0.84 | 666.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | EPPlus | 39.15 ms | 1.76 ms | 1.02 ms | 10.21 | 12.75 | 16646.1 KB | 12.43 | 114.9 KB | 0.81 | 920.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 5.23 ms | 0.94 ms | 0.54 ms | 1.00 | 1.00 | 1505.3 KB | 1.00 | 138.1 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 55.32 ms | 0.89 ms | 0.51 ms | 10.58 | 10.58 | 38343.9 KB | 25.47 | 115.1 KB | 0.83 | 957.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 66.50 ms | 3.04 ms | 1.75 ms | 12.71 | 12.71 | 50927.5 KB | 33.83 | 120.2 KB | 0.87 | 1171.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 3.05 ms | 0.01 ms | 0.01 ms | 0.62 | 1.00 | 758.3 KB | 0.51 | 138.4 KB | 1.00 | 38.4% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 4.95 ms | 0.25 ms | 0.15 ms | 1.00 | 1.62 | 1497.5 KB | 1.00 | 138.0 KB | 1.00 | Loss +62.5% |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | MiniExcel | 7.83 ms | 0.12 ms | 0.07 ms | 1.58 | 2.57 | 23222.1 KB | 15.51 | 153.7 KB | 1.11 | 58.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | ClosedXML | 27.99 ms | 0.95 ms | 0.55 ms | 5.66 | 9.19 | 11581.0 KB | 7.73 | 120.1 KB | 0.87 | 465.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | EPPlus | 37.85 ms | 1.35 ms | 0.78 ms | 7.65 | 12.43 | 16646.1 KB | 11.12 | 114.9 KB | 0.83 | 665.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | LargeXlsx | 3.11 ms | 0.07 ms | 0.04 ms | 0.68 | 1.00 | 758.3 KB | 0.43 | 138.4 KB | 0.97 | 32.4% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | OfficeIMO.Excel | 4.59 ms | 0.15 ms | 0.09 ms | 1.00 | 1.48 | 1770.1 KB | 1.00 | 142.3 KB | 1.00 | Loss +47.9% |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | MiniExcel | 7.83 ms | 0.19 ms | 0.11 ms | 1.70 | 2.52 | 23222.1 KB | 13.12 | 153.7 KB | 1.08 | 70.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | ClosedXML | 27.65 ms | 0.39 ms | 0.23 ms | 6.02 | 8.90 | 11581.0 KB | 6.54 | 120.1 KB | 0.84 | 501.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | EPPlus | 35.97 ms | 0.96 ms | 0.55 ms | 7.83 | 11.57 | 16646.1 KB | 9.40 | 114.9 KB | 0.81 | 682.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 3.61 ms | 0.12 ms | 0.07 ms | 1.00 | 1.00 | 1346.4 KB | 1.00 | 142.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 42.02 ms | 1.45 ms | 0.84 ms | 11.64 | 11.64 | 28540.6 KB | 21.20 | 120.2 KB | 0.84 | 1063.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 44.83 ms | 2.04 ms | 1.18 ms | 12.42 | 12.42 | 27305.8 KB | 20.28 | 115.0 KB | 0.81 | 1141.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | LargeXlsx | 4.27 ms | 0.39 ms | 0.23 ms | 0.77 | 1.00 | 802.5 KB | 0.34 | 182.6 KB | 1.00 | 23.4% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 5.57 ms | 0.61 ms | 0.35 ms | 1.00 | 1.31 | 2341.7 KB | 1.00 | 183.1 KB | 1.00 | Loss +30.5% |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | MiniExcel | 8.33 ms | 0.19 ms | 0.11 ms | 1.49 | 1.95 | 25190.5 KB | 10.76 | 194.0 KB | 1.06 | 49.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | ClosedXML | 35.10 ms | 1.85 ms | 1.07 ms | 6.30 | 8.22 | 16973.5 KB | 7.25 | 161.0 KB | 0.88 | 529.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | EPPlus | 47.57 ms | 3.38 ms | 1.95 ms | 8.54 | 11.15 | 20105.1 KB | 8.59 | 152.1 KB | 0.83 | 753.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | LargeXlsx | 4.10 ms | 0.24 ms | 0.14 ms | 0.92 | 1.00 | 802.5 KB | 0.53 | 182.6 KB | 1.00 | 8.2% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 4.47 ms | 0.15 ms | 0.09 ms | 1.00 | 1.09 | 1507.7 KB | 1.00 | 182.4 KB | 1.00 | Loss +9.0% |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | MiniExcel | 8.80 ms | 0.49 ms | 0.28 ms | 1.97 | 2.14 | 25190.5 KB | 16.71 | 194.0 KB | 1.06 | 96.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | ClosedXML | 36.57 ms | 1.09 ms | 0.63 ms | 8.18 | 8.91 | 16973.5 KB | 11.26 | 161.0 KB | 0.88 | 717.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | EPPlus | 49.63 ms | 2.29 ms | 1.32 ms | 11.10 | 12.10 | 20105.1 KB | 13.33 | 152.1 KB | 0.83 | 1009.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 19.19 ms | 0.79 ms | 0.45 ms | 1.00 | 1.00 | 4502.3 KB | 1.00 | 651.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | LargeXlsx | 19.48 ms | 0.22 ms | 0.13 ms | 1.02 | 1.02 | 2810.7 KB | 0.62 | 644.6 KB | 0.99 | Tie vs OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | MiniExcel | 32.81 ms | 0.28 ms | 0.16 ms | 1.71 | 1.71 | 48414.8 KB | 10.75 | 674.4 KB | 1.04 | 71.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | ClosedXML | 117.33 ms | 4.03 ms | 2.33 ms | 6.11 | 6.11 | 51647.0 KB | 11.47 | 615.5 KB | 0.95 | 511.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | EPPlus | 157.85 ms | 1.69 ms | 0.98 ms | 8.23 | 8.23 | 69139.6 KB | 15.36 | 548.9 KB | 0.84 | 722.6% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | OfficeIMO.Excel | 8.38 ms | 0.34 ms | 0.20 ms | 1.00 | 1.00 | 1895.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | EPPlus | 76.78 ms | 2.49 ms | 1.44 ms | 9.17 | 9.17 | 50712.0 KB | 26.76 |  |  | 816.8% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | EPPlus 4.5.3.3 | 103.66 ms |  |  | 12.38 | 12.38 |  |  |  |  | 1137.7% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | ClosedXML | 136.52 ms | 8.46 ms | 4.88 ms | 16.30 | 16.30 | 84651.6 KB | 44.67 |  |  | 1530.0% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | OfficeIMO.Excel | 19.21 ms | 1.91 ms | 1.10 ms | 1.00 | 1.00 | 15211.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus 4.5.3.3 | 43.60 ms |  |  | 2.27 | 2.27 |  |  |  |  | 126.9% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus | 54.19 ms | 1.07 ms | 0.62 ms | 2.82 | 2.82 | 32906.1 KB | 2.16 |  |  | 182.1% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | OfficeIMO.Excel | 11.62 ms | 0.18 ms | 0.10 ms | 1.00 | 1.00 | 6198.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus | 77.67 ms | 4.37 ms | 2.52 ms | 6.69 | 6.69 | 54593.5 KB | 8.81 |  |  | 568.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus 4.5.3.3 | 83.69 ms |  |  | 7.20 | 7.20 |  |  |  |  | 620.4% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | OfficeIMO.Excel | 40.21 ms | 11.51 ms | 6.65 ms | 1.00 | 1.00 | 16352.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus 4.5.3.3 | 88.63 ms |  |  | 2.20 | 2.20 |  |  |  |  | 120.4% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus | 170.20 ms | 50.35 ms | 29.07 ms | 4.23 | 4.23 | 59225.9 KB | 3.62 |  |  | 323.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | OfficeIMO.Excel | 25.82 ms | 2.01 ms | 1.16 ms | 1.00 | 1.00 | 15231.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus 4.5.3.3 | 76.72 ms |  |  | 2.97 | 2.97 |  |  |  |  | 197.2% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus | 111.68 ms | 11.79 ms | 6.81 ms | 4.33 | 4.33 | 54594.0 KB | 3.58 |  |  | 332.6% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | OfficeIMO.Excel | 18.97 ms | 0.63 ms | 0.36 ms | 1.00 | 1.00 | 15225.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus 4.5.3.3 | 78.88 ms |  |  | 4.16 | 4.16 |  |  |  |  | 315.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus | 89.02 ms | 19.66 ms | 11.35 ms | 4.69 | 4.69 | 54590.6 KB | 3.59 |  |  | 369.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Object projection | build-object-datatable-dictionaries | OfficeIMO.Excel | 1.54 ms | 0.06 ms | 0.03 ms | 1.00 | 1.00 | 564.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Object projection | build-object-datatable-typed | OfficeIMO.Excel | 1.31 ms | 0.03 ms | 0.02 ms | 1.00 | 1.00 | 856.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-cells | OfficeIMO.Excel | 6.35 ms | 1.16 ms | 0.67 ms | 1.00 | 1.00 | 2531.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-cells | EPPlus | 29.73 ms | 5.29 ms | 3.06 ms | 4.68 | 4.68 | 20154.9 KB | 7.96 |  |  | 368.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-cells | ClosedXML | 31.81 ms | 4.10 ms | 2.37 ms | 5.01 | 5.01 | 17022.2 KB | 6.72 |  |  | 400.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | OfficeIMO.Excel | 3.99 ms | 0.22 ms | 0.13 ms | 1.00 | 1.00 | 523.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | EPPlus | 24.83 ms | 1.85 ms | 1.07 ms | 6.22 | 6.22 | 13108.1 KB | 25.04 |  |  | 521.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | ClosedXML | 32.17 ms | 2.02 ms | 1.17 ms | 8.05 | 8.05 | 15463.4 KB | 29.55 |  |  | 705.5% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-range | OfficeIMO.Excel | 6.68 ms | 0.32 ms | 0.19 ms | 1.00 | 1.00 | 2531.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-range | EPPlus | 30.43 ms | 3.96 ms | 2.29 ms | 4.55 | 4.55 | 20154.9 KB | 7.96 |  |  | 355.4% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-range | ClosedXML | 31.78 ms | 2.66 ms | 1.54 ms | 4.76 | 4.76 | 17020.9 KB | 6.72 |  |  | 375.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-top-range | OfficeIMO.Excel | 0.64 ms | 0.05 ms | 0.03 ms | 1.00 | 1.00 | 285.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-top-range | EPPlus | 23.36 ms | 0.75 ms | 0.43 ms | 36.25 | 36.25 | 12404.4 KB | 43.48 |  |  | 3524.8% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-top-range | ClosedXML | 29.03 ms | 0.93 ms | 0.54 ms | 45.06 | 45.06 | 15370.2 KB | 53.88 |  |  | 4406.0% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-autofilter | OfficeIMO.Excel | 4.46 ms | 0.51 ms | 0.30 ms | 1.00 | 1.00 | 1340.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-autofilter | ClosedXML | 36.07 ms | 5.00 ms | 2.89 ms | 8.09 | 8.09 | 22226.8 KB | 16.58 |  |  | 709.0% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-autofilter | EPPlus 4.5.3.3 | 36.75 ms |  |  | 8.24 | 8.24 |  |  |  |  | 724.1% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-autofilter | EPPlus | 49.80 ms | 9.33 ms | 5.39 ms | 11.17 | 11.17 | 24715.5 KB | 18.44 |  |  | 1016.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-charts | OfficeIMO.Excel | 5.07 ms | 0.37 ms | 0.21 ms | 1.00 | 1.00 | 1892.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-charts | EPPlus 4.5.3.3 | 36.97 ms |  |  | 7.29 | 7.29 |  |  |  |  | 628.5% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-charts | EPPlus | 44.00 ms | 1.28 ms | 0.74 ms | 8.67 | 8.67 | 27141.8 KB | 14.34 |  |  | 767.1% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | OfficeIMO.Excel | 4.24 ms | 0.62 ms | 0.36 ms | 1.00 | 1.00 | 1405.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | EPPlus 4.5.3.3 | 35.01 ms |  |  | 8.26 | 8.26 |  |  |  |  | 725.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | ClosedXML | 36.02 ms | 6.97 ms | 4.02 ms | 8.50 | 8.50 | 22273.8 KB | 15.84 |  |  | 749.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | EPPlus | 44.88 ms | 3.86 ms | 2.23 ms | 10.59 | 10.59 | 24757.5 KB | 17.61 |  |  | 958.8% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-data-validation | OfficeIMO.Excel | 3.66 ms | 0.06 ms | 0.03 ms | 1.00 | 1.00 | 1356.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-data-validation | ClosedXML | 32.62 ms | 0.43 ms | 0.25 ms | 8.90 | 8.90 | 22247.9 KB | 16.41 |  |  | 790.2% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-data-validation | EPPlus 4.5.3.3 | 36.82 ms |  |  | 10.05 | 10.05 |  |  |  |  | 904.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-data-validation | EPPlus | 43.28 ms | 2.19 ms | 1.26 ms | 11.81 | 11.81 | 24701.4 KB | 18.22 |  |  | 1081.1% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | OfficeIMO.Excel | 4.62 ms | 0.98 ms | 0.57 ms | 1.00 | 1.00 | 1342.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | EPPlus 4.5.3.3 | 36.27 ms |  |  | 7.85 | 7.85 |  |  |  |  | 685.0% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | ClosedXML | 50.73 ms | 6.31 ms | 3.65 ms | 10.98 | 10.98 | 22222.0 KB | 16.55 |  |  | 997.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | EPPlus | 53.41 ms | 6.08 ms | 3.51 ms | 11.56 | 11.56 | 24730.0 KB | 18.42 |  |  | 1055.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | OfficeIMO.Excel | 14.65 ms | 1.04 ms | 0.60 ms | 1.00 | 1.00 | 14419.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | EPPlus 4.5.3.3 | 37.98 ms |  |  | 2.59 | 2.59 |  |  |  |  | 159.2% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | EPPlus | 46.65 ms | 2.91 ms | 1.68 ms | 3.18 | 3.18 | 29537.1 KB | 2.05 |  |  | 218.4% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world report | realworld-report-all-in-one | OfficeIMO.Excel | 22.00 ms | 3.24 ms | 1.87 ms | 1.00 | 1.00 | 15220.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus 4.5.3.3 | 80.50 ms |  |  | 3.66 | 3.66 |  |  |  |  | 265.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus | 86.01 ms | 4.02 ms | 2.32 ms | 3.91 | 3.91 | 54594.5 KB | 3.59 |  |  | 291.0% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world report | realworld-report-core | OfficeIMO.Excel | 5.69 ms | 0.36 ms | 0.21 ms | 1.00 | 1.00 | 1488.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world report | realworld-report-core | EPPlus 4.5.3.3 | 73.91 ms |  |  | 12.99 | 12.99 |  |  |  |  | 1199.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world report | realworld-report-core | EPPlus | 86.27 ms | 2.81 ms | 1.62 ms | 15.17 | 15.17 | 47299.8 KB | 31.77 |  |  | 1416.5% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world report | realworld-report-core | ClosedXML | 100.27 ms | 5.96 ms | 3.44 ms | 17.63 | 17.63 | 69834.2 KB | 46.91 |  |  | 1662.6% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook | OfficeIMO.Excel | 42.69 ms | 7.07 ms | 4.08 ms | 1.00 | 1.00 | 19069.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Report workbook | report-workbook | EPPlus 4.5.3.3 | 85.84 ms |  |  | 2.01 | 2.01 |  |  |  |  | 101.1% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook | EPPlus | 119.63 ms | 9.52 ms | 5.50 ms | 2.80 | 2.80 | 77486.2 KB | 4.06 |  |  | 180.2% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-core | OfficeIMO.Excel | 7.48 ms | 0.71 ms | 0.41 ms | 1.00 | 1.00 | 2711.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Report workbook | report-workbook-core | EPPlus 4.5.3.3 | 75.29 ms |  |  | 10.07 | 10.07 |  |  |  |  | 907.0% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-core | ClosedXML | 124.29 ms | 22.07 ms | 12.74 ms | 16.62 | 16.62 | 97220.1 KB | 35.86 |  |  | 1562.4% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-core | EPPlus | 130.20 ms | 21.58 ms | 12.46 ms | 17.41 | 17.41 | 71970.6 KB | 26.55 |  |  | 1641.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable | OfficeIMO.Excel | 30.54 ms | 3.87 ms | 2.23 ms | 1.00 | 1.00 | 19383.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable | EPPlus 4.5.3.3 | 75.79 ms |  |  | 2.48 | 2.48 |  |  |  |  | 148.2% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable | EPPlus | 132.53 ms | 8.33 ms | 4.81 ms | 4.34 | 4.34 | 65995.3 KB | 3.40 |  |  | 334.0% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable-core | OfficeIMO.Excel | 7.50 ms | 0.67 ms | 0.39 ms | 1.00 | 1.00 | 2982.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable-core | EPPlus 4.5.3.3 | 70.54 ms |  |  | 9.40 | 9.40 |  |  |  |  | 840.2% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable-core | EPPlus | 125.14 ms | 5.77 ms | 3.33 ms | 16.68 | 16.68 | 60480.1 KB | 20.28 |  |  | 1568.0% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable-core | ClosedXML | 135.90 ms | 10.03 ms | 5.79 ms | 18.11 | 18.11 | 82858.9 KB | 27.78 |  |  | 1711.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | OfficeIMO.Excel | 2.11 ms | 0.03 ms | 0.02 ms | 1.00 | 1.00 | 706.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | EPPlus 4.5.3.3 | 13.28 ms |  |  | 6.31 | 6.31 |  |  |  |  | 530.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | ClosedXML | 18.24 ms | 1.18 ms | 0.68 ms | 8.66 | 8.66 | 8279.3 KB | 11.72 |  |  | 766.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | EPPlus | 18.33 ms | 1.23 ms | 0.71 ms | 8.71 | 8.71 | 7708.0 KB | 10.91 |  |  | 770.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | OfficeIMO.Excel | 0.83 ms | 0.02 ms | 0.01 ms | 1.00 | 1.00 | 177.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | Sylvan.Data.Excel | 1.02 ms | 0.06 ms | 0.03 ms | 1.22 | 1.22 | 316.6 KB | 1.79 |  |  | 22.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | ExcelDataReader | 1.46 ms | 0.03 ms | 0.02 ms | 1.75 | 1.75 | 4062.2 KB | 22.93 |  |  | 74.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | ClosedXML | 3.87 ms | 0.67 ms | 0.39 ms | 4.65 | 4.65 | 4392.6 KB | 24.79 |  |  | 364.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | EPPlus 4.5.3.3 | 11.29 ms |  |  | 13.55 | 13.55 |  |  |  |  | 1255.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | MiniExcel | 11.96 ms | 0.81 ms | 0.47 ms | 14.36 | 14.36 | 46194.9 KB | 260.75 |  |  | 1335.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | EPPlus | 37.89 ms | 0.78 ms | 0.45 ms | 45.50 | 45.50 | 43071.0 KB | 243.11 |  |  | 4450.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | Sylvan.Data.Excel | 0.96 ms | 0.02 ms | 0.01 ms | 0.57 | 1.00 | 316.6 KB | 1.79 |  |  | 43.5% faster than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | ExcelDataReader | 1.47 ms | 0.08 ms | 0.05 ms | 0.87 | 1.54 | 4062.2 KB | 22.92 |  |  | 12.9% faster than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | OfficeIMO.Excel | 1.69 ms | 0.83 ms | 0.48 ms | 1.00 | 1.77 | 177.2 KB | 1.00 |  |  | Loss +77.0% |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | ClosedXML | 3.73 ms | 0.87 ms | 0.50 ms | 2.21 | 3.91 | 4392.5 KB | 24.78 |  |  | 120.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | EPPlus 4.5.3.3 | 10.98 ms |  |  | 6.50 | 11.49 |  |  |  |  | 549.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | MiniExcel | 12.06 ms | 0.41 ms | 0.24 ms | 7.13 | 12.62 | 46194.9 KB | 260.64 |  |  | 613.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | EPPlus | 38.97 ms | 2.49 ms | 1.44 ms | 23.04 | 40.78 | 43071.0 KB | 243.02 |  |  | 2204.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | Sylvan.Data.Excel | 1.95 ms | 0.09 ms | 0.05 ms | 0.54 | 1.00 | 518.6 KB | 0.49 |  |  | 46.2% faster than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | OfficeIMO.Excel | 3.62 ms | 0.19 ms | 0.11 ms | 1.00 | 1.86 | 1056.5 KB | 1.00 |  |  | Loss +85.9% |
| 2500 | speed-comparison | read | Other | shared-string-read | ExcelDataReader | 4.76 ms | 0.27 ms | 0.16 ms | 1.31 | 2.44 | 2619.1 KB | 2.48 |  |  | 31.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | MiniExcel | 6.07 ms | 0.44 ms | 0.26 ms | 1.68 | 3.11 | 7530.1 KB | 7.13 |  |  | 67.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | EPPlus 4.5.3.3 | 12.69 ms |  |  | 3.50 | 6.51 |  |  |  |  | 250.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | ClosedXML | 16.29 ms | 0.57 ms | 0.33 ms | 4.50 | 8.36 | 9498.0 KB | 8.99 |  |  | 349.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | EPPlus | 22.01 ms | 4.51 ms | 2.61 ms | 6.08 | 11.30 | 10372.2 KB | 9.82 |  |  | 507.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | OfficeIMO.Excel | 4.49 ms | 1.18 ms | 0.68 ms | 1.00 | 1.00 | 374.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | Sylvan.Data.Excel | 4.65 ms | 1.16 ms | 0.67 ms | 1.03 | 1.03 | 655.2 KB | 1.75 |  |  | 3.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | ExcelDataReader | 10.62 ms | 0.08 ms | 0.05 ms | 2.37 | 2.37 | 6089.5 KB | 16.26 |  |  | 136.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | MiniExcel | 13.63 ms | 1.00 ms | 0.58 ms | 3.03 | 3.03 | 18661.8 KB | 49.83 |  |  | 203.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | EPPlus | 24.14 ms | 1.33 ms | 0.77 ms | 5.37 | 5.37 | 12427.1 KB | 33.18 |  |  | 437.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | ClosedXML | 30.06 ms | 1.61 ms | 0.93 ms | 6.69 | 6.69 | 15361.1 KB | 41.02 |  |  | 569.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | Sylvan.Data.Excel | 6.02 ms | 0.09 ms | 0.05 ms | 0.86 | 1.00 | 2239.3 KB | 0.62 |  |  | 14.4% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | OfficeIMO.Excel | 7.03 ms | 1.19 ms | 0.69 ms | 1.00 | 1.17 | 3594.4 KB | 1.00 |  |  | Loss +16.8% |
| 2500 | speed-comparison | read | Range and table read | read-datatable | MiniExcel | 13.69 ms | 2.16 ms | 1.25 ms | 1.95 | 2.28 | 18266.6 KB | 5.08 |  |  | 94.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | ExcelDataReader | 14.67 ms | 3.66 ms | 2.11 ms | 2.09 | 2.44 | 7673.5 KB | 2.13 |  |  | 108.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | ClosedXML | 32.63 ms | 1.30 ms | 0.75 ms | 4.64 | 5.42 | 18313.9 KB | 5.10 |  |  | 364.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | EPPlus | 35.24 ms | 7.02 ms | 4.05 ms | 5.01 | 5.85 | 21736.6 KB | 6.05 |  |  | 401.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | EPPlus 4.5.3.3 | 42.02 ms |  |  | 5.98 | 6.98 |  |  |  |  | 497.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | OfficeIMO.Excel | 3.98 ms | 0.15 ms | 0.09 ms | 1.00 | 1.00 | 542.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | Sylvan.Data.Excel | 4.44 ms | 0.11 ms | 0.06 ms | 1.12 | 1.12 | 733.5 KB | 1.35 |  |  | 11.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | MiniExcel | 10.24 ms | 0.96 ms | 0.55 ms | 2.57 | 2.57 | 15850.3 KB | 29.20 |  |  | 157.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | ExcelDataReader | 10.64 ms | 0.23 ms | 0.13 ms | 2.67 | 2.67 | 6089.5 KB | 11.22 |  |  | 167.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | EPPlus | 23.96 ms | 0.99 ms | 0.57 ms | 6.02 | 6.02 | 13108.1 KB | 24.14 |  |  | 501.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | ClosedXML | 30.30 ms | 0.40 ms | 0.23 ms | 7.61 | 7.61 | 15465.1 KB | 28.49 |  |  | 660.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | OfficeIMO.Excel | 11.45 ms | 1.52 ms | 0.88 ms | 1.00 | 1.00 | 2692.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Range and table read | read-range | Sylvan.Data.Excel | 11.70 ms | 5.56 ms | 3.21 ms | 1.02 | 1.02 | 655.0 KB | 0.24 |  |  | 2.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | ExcelDataReader | 21.82 ms | 3.50 ms | 2.02 ms | 1.91 | 1.91 | 6089.2 KB | 2.26 |  |  | 90.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | MiniExcel | 22.84 ms | 4.99 ms | 2.88 ms | 1.99 | 1.99 | 18662.2 KB | 6.93 |  |  | 99.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | EPPlus 4.5.3.3 | 38.35 ms |  |  | 3.35 | 3.35 |  |  |  |  | 235.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | EPPlus | 38.88 ms | 9.03 ms | 5.22 ms | 3.40 | 3.40 | 20152.6 KB | 7.48 |  |  | 239.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | ClosedXML | 80.79 ms | 26.41 ms | 15.25 ms | 7.06 | 7.06 | 16846.5 KB | 6.26 |  |  | 605.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | OfficeIMO.Excel | 5.71 ms | 0.97 ms | 0.56 ms | 1.00 | 1.00 | 2751.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | Sylvan.Data.Excel | 6.22 ms | 0.20 ms | 0.12 ms | 1.09 | 1.09 | 750.3 KB | 0.27 |  |  | 8.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | ExcelDataReader | 11.46 ms | 0.65 ms | 0.38 ms | 2.01 | 2.01 | 6089.5 KB | 2.21 |  |  | 100.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | MiniExcel | 14.88 ms | 2.21 ms | 1.28 ms | 2.61 | 2.61 | 18662.4 KB | 6.78 |  |  | 160.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | ClosedXML | 32.82 ms | 1.34 ms | 0.77 ms | 5.75 | 5.75 | 16728.6 KB | 6.08 |  |  | 475.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | EPPlus | 34.62 ms | 9.53 ms | 5.50 ms | 6.07 | 6.07 | 20152.6 KB | 7.32 |  |  | 506.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | Sylvan.Data.Excel | 0.55 ms | 0.08 ms | 0.05 ms | 0.85 | 1.00 | 348.5 KB | 1.18 |  |  | 14.7% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | OfficeIMO.Excel | 0.64 ms | 0.09 ms | 0.05 ms | 1.00 | 1.17 | 296.0 KB | 1.00 |  |  | Loss +17.2% |
| 2500 | speed-comparison | read | Range and table read | read-top-range | MiniExcel | 0.84 ms | 0.02 ms | 0.01 ms | 1.31 | 1.53 | 869.0 KB | 2.94 |  |  | 30.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | ExcelDataReader | 4.53 ms | 0.17 ms | 0.10 ms | 7.07 | 8.29 | 1931.8 KB | 6.53 |  |  | 607.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | EPPlus | 26.31 ms | 1.87 ms | 1.08 ms | 41.08 | 48.14 | 12402.1 KB | 41.89 |  |  | 4008.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | EPPlus 4.5.3.3 | 29.93 ms |  |  | 46.74 | 54.77 |  |  |  |  | 4574.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | ClosedXML | 36.41 ms | 3.57 ms | 2.06 ms | 56.86 | 66.63 | 15360.3 KB | 51.88 |  |  | 5586.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | Sylvan.Data.Excel | 4.89 ms | 0.36 ms | 0.21 ms | 0.35 | 1.00 | 655.2 KB | 0.19 |  |  | 65.4% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | ExcelDataReader | 11.34 ms | 0.77 ms | 0.45 ms | 0.80 | 2.32 | 6089.5 KB | 1.75 |  |  | 19.7% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | OfficeIMO.Excel | 14.12 ms | 3.13 ms | 1.81 ms | 1.00 | 2.89 | 3472.7 KB | 1.00 |  |  | Loss +188.6% |
| 2500 | speed-comparison | read | Range and table read | read-used-range | MiniExcel | 15.15 ms | 4.50 ms | 2.60 ms | 1.07 | 3.10 | 18662.4 KB | 5.37 |  |  | 7.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | EPPlus | 33.18 ms | 6.74 ms | 3.89 ms | 2.35 | 6.78 | 20152.7 KB | 5.80 |  |  | 134.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | ClosedXML | 63.52 ms | 36.96 ms | 21.34 ms | 4.50 | 12.98 | 16806.9 KB | 4.84 |  |  | 349.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | OfficeIMO.Excel | 3.84 ms | 0.44 ms | 0.26 ms | 1.00 | 1.00 | 377.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | Sylvan.Data.Excel | 4.62 ms | 0.94 ms | 0.54 ms | 1.20 | 1.20 | 655.2 KB | 1.73 |  |  | 20.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | ExcelDataReader | 10.39 ms | 0.84 ms | 0.48 ms | 2.71 | 2.71 | 6089.5 KB | 16.12 |  |  | 170.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | MiniExcel | 12.30 ms | 1.11 ms | 0.64 ms | 3.21 | 3.21 | 18661.8 KB | 49.41 |  |  | 220.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | EPPlus | 24.51 ms | 1.64 ms | 0.94 ms | 6.39 | 6.39 | 12427.1 KB | 32.90 |  |  | 538.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | ClosedXML | 31.13 ms | 1.71 ms | 0.99 ms | 8.11 | 8.11 | 15359.3 KB | 40.66 |  |  | 711.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | Sylvan.Data.Excel | 4.96 ms | 0.28 ms | 0.16 ms | 0.77 | 1.00 | 655.2 KB | 0.24 |  |  | 23.3% faster than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | OfficeIMO.Excel | 6.47 ms | 1.71 ms | 0.99 ms | 1.00 | 1.30 | 2771.4 KB | 1.00 |  |  | Loss +30.4% |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | ExcelDataReader | 11.35 ms | 0.63 ms | 0.36 ms | 1.75 | 2.29 | 6089.5 KB | 2.20 |  |  | 75.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | MiniExcel | 12.23 ms | 0.66 ms | 0.38 ms | 1.89 | 2.47 | 18662.4 KB | 6.73 |  |  | 89.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | EPPlus | 32.63 ms | 3.47 ms | 2.00 ms | 5.04 | 6.58 | 20152.6 KB | 7.27 |  |  | 404.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | ClosedXML | 32.73 ms | 2.15 ms | 1.24 ms | 5.06 | 6.60 | 16729.4 KB | 6.04 |  |  | 405.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | EPPlus 4.5.3.3 | 36.89 ms |  |  | 5.70 | 7.44 |  |  |  |  | 470.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | Sylvan.Data.Excel | 0.46 ms | 0.08 ms | 0.05 ms | 0.77 | 1.00 | 348.5 KB | 1.16 |  |  | 22.7% faster than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | OfficeIMO.Excel | 0.60 ms | 0.12 ms | 0.07 ms | 1.00 | 1.29 | 299.4 KB | 1.00 |  |  | Loss +29.4% |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | MiniExcel | 0.72 ms | 0.02 ms | 0.01 ms | 1.20 | 1.55 | 869.0 KB | 2.90 |  |  | 20.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | ExcelDataReader | 4.57 ms | 0.77 ms | 0.44 ms | 7.62 | 9.86 | 1931.8 KB | 6.45 |  |  | 662.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | EPPlus | 24.16 ms | 0.70 ms | 0.40 ms | 40.25 | 52.06 | 12402.1 KB | 41.43 |  |  | 3924.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | EPPlus 4.5.3.3 | 31.25 ms |  |  | 52.06 | 67.35 |  |  |  |  | 5106.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | ClosedXML | 31.69 ms | 3.02 ms | 1.75 ms | 52.79 | 68.29 | 15360.9 KB | 51.31 |  |  | 5178.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | Sylvan.Data.Excel | 0.44 ms | 0.05 ms | 0.03 ms | 0.78 | 1.00 | 348.5 KB | 1.16 |  |  | 21.6% faster than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | OfficeIMO.Excel | 0.56 ms | 0.03 ms | 0.01 ms | 1.00 | 1.28 | 300.2 KB | 1.00 |  |  | Loss +27.6% |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | MiniExcel | 0.68 ms | 0.00 ms | 0.00 ms | 1.23 | 1.57 | 869.0 KB | 2.89 |  |  | 22.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | ExcelDataReader | 4.35 ms | 0.31 ms | 0.18 ms | 7.81 | 9.96 | 1931.8 KB | 6.44 |  |  | 680.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | EPPlus | 24.60 ms | 0.35 ms | 0.20 ms | 44.18 | 56.36 | 12402.1 KB | 41.31 |  |  | 4317.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | ClosedXML | 31.07 ms | 2.60 ms | 1.50 ms | 55.80 | 71.18 | 15360.5 KB | 51.17 |  |  | 5479.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | Sylvan.Data.Excel | 6.90 ms | 3.63 ms | 2.10 ms | 0.91 | 1.00 | 895.3 KB | 0.37 |  |  | 8.6% faster than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | OfficeIMO.Excel | 7.55 ms | 1.79 ms | 1.03 ms | 1.00 | 1.09 | 2442.0 KB | 1.00 |  |  | Loss +9.4% |
| 2500 | speed-comparison | read | Typed object read | read-objects | ExcelDataReader | 13.57 ms | 4.34 ms | 2.51 ms | 1.80 | 1.97 | 6329.5 KB | 2.59 |  |  | 79.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | MiniExcel | 13.80 ms | 0.59 ms | 0.34 ms | 1.83 | 2.00 | 18474.0 KB | 7.57 |  |  | 82.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | ClosedXML | 32.04 ms | 1.71 ms | 0.98 ms | 4.24 | 4.64 | 16925.4 KB | 6.93 |  |  | 324.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | EPPlus | 33.46 ms | 9.96 ms | 5.75 ms | 4.43 | 4.85 | 21354.2 KB | 8.74 |  |  | 343.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | EPPlus 4.5.3.3 | 35.96 ms |  |  | 4.76 | 5.21 |  |  |  |  | 376.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | Sylvan.Data.Excel | 4.55 ms | 0.07 ms | 0.04 ms | 0.96 | 1.00 | 831.0 KB | 0.34 |  |  | 3.7% faster than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | OfficeIMO.Excel | 4.72 ms | 0.12 ms | 0.07 ms | 1.00 | 1.04 | 2422.8 KB | 1.00 |  |  | Loss +3.8% |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | ExcelDataReader | 10.37 ms | 0.11 ms | 0.07 ms | 2.20 | 2.28 | 6265.3 KB | 2.59 |  |  | 119.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | MiniExcel | 12.33 ms | 0.07 ms | 0.04 ms | 2.61 | 2.71 | 18409.8 KB | 7.60 |  |  | 161.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | ClosedXML | 29.87 ms | 0.77 ms | 0.45 ms | 6.32 | 6.57 | 16903.9 KB | 6.98 |  |  | 532.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | EPPlus | 33.92 ms | 0.57 ms | 0.33 ms | 7.18 | 7.45 | 21334.6 KB | 8.81 |  |  | 617.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | EPPlus 4.5.3.3 | 34.05 ms |  |  | 7.21 | 7.48 |  |  |  |  | 620.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | OfficeIMO.Excel | 5.13 ms | 0.56 ms | 0.32 ms | 1.00 | 1.00 | 1781.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | MiniExcel | 9.47 ms | 1.27 ms | 0.73 ms | 1.85 | 1.85 | 26647.3 KB | 14.96 |  |  | 84.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | EPPlus | 56.55 ms | 2.49 ms | 1.44 ms | 11.03 | 11.03 | 38343.3 KB | 21.53 |  |  | 1003.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | EPPlus 4.5.3.3 | 74.15 ms |  |  | 14.46 | 14.46 |  |  |  |  | 1346.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | ClosedXML | 83.68 ms | 5.63 ms | 3.25 ms | 16.32 | 16.32 | 58360.0 KB | 32.77 |  |  | 1532.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | OfficeIMO.Excel | 8.68 ms | 0.45 ms | 0.26 ms | 1.00 | 1.00 | 2078.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | MiniExcel | 25.40 ms | 8.74 ms | 5.04 ms | 2.93 | 2.93 | 32152.0 KB | 15.47 |  |  | 192.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | EPPlus | 123.19 ms | 38.07 ms | 21.98 ms | 14.19 | 14.19 | 42120.0 KB | 20.26 |  |  | 1318.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | ClosedXML | 341.18 ms | 61.11 ms | 35.28 ms | 39.30 | 39.30 | 56707.1 KB | 27.28 |  |  | 3829.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.03 ms | 0.26 ms | 0.15 ms | 1.00 | 1.00 | 1347.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 54.80 ms | 1.03 ms | 0.59 ms | 13.61 | 13.61 | 38344.1 KB | 28.46 |  |  | 1261.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 63.73 ms | 0.85 ms | 0.49 ms | 15.83 | 15.83 | 50927.5 KB | 37.80 |  |  | 1483.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 5.32 ms | 1.09 ms | 0.63 ms | 1.00 | 1.00 | 1505.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 60.04 ms | 2.74 ms | 1.58 ms | 11.27 | 11.27 | 38344.1 KB | 25.47 |  |  | 1027.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 66.04 ms | 1.61 ms | 0.93 ms | 12.40 | 12.40 | 50927.3 KB | 33.83 |  |  | 1140.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 3.98 ms | 0.18 ms | 0.10 ms | 1.00 | 1.00 | 1346.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 45.30 ms | 2.25 ms | 1.30 ms | 11.39 | 11.39 | 28540.4 KB | 21.20 |  |  | 1039.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 53.16 ms | 8.68 ms | 5.01 ms | 13.37 | 13.37 | 27305.8 KB | 20.28 |  |  | 1236.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | OfficeIMO.Excel | 2.31 ms | 0.08 ms | 0.05 ms | 1.00 | 1.00 | 1787.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | ClosedXML | 11.96 ms | 0.66 ms | 0.38 ms | 5.19 | 5.19 | 9959.5 KB | 5.57 |  |  | 418.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | EPPlus | 20.68 ms | 1.02 ms | 0.59 ms | 8.97 | 8.97 | 11772.9 KB | 6.59 |  |  | 796.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | OfficeIMO.Excel | 2.41 ms | 0.33 ms | 0.19 ms | 1.00 | 1.00 | 1119.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | ClosedXML | 11.72 ms | 0.35 ms | 0.20 ms | 4.87 | 4.87 | 9177.1 KB | 8.19 |  |  | 387.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | EPPlus 4.5.3.3 | 16.63 ms |  |  | 6.91 | 6.91 |  |  |  |  | 591.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | EPPlus | 22.77 ms | 1.67 ms | 0.96 ms | 9.46 | 9.46 | 12895.2 KB | 11.51 |  |  | 846.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | OfficeIMO.Excel | 3.00 ms | 0.26 ms | 0.15 ms | 1.00 | 1.00 | 1763.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | ClosedXML | 17.88 ms | 0.30 ms | 0.17 ms | 5.95 | 5.95 | 11887.0 KB | 6.74 |  |  | 495.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | EPPlus 4.5.3.3 | 21.94 ms |  |  | 7.30 | 7.30 |  |  |  |  | 630.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | EPPlus | 25.36 ms | 0.48 ms | 0.27 ms | 8.44 | 8.44 | 15643.4 KB | 8.87 |  |  | 744.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | OfficeIMO.Excel | 2.91 ms | 0.10 ms | 0.06 ms | 1.00 | 1.00 | 1506.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | ClosedXML | 15.85 ms | 1.37 ms | 0.79 ms | 5.45 | 5.45 | 11296.3 KB | 7.50 |  |  | 445.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | EPPlus | 27.04 ms | 1.27 ms | 0.73 ms | 9.30 | 9.30 | 14960.3 KB | 9.93 |  |  | 830.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 2.85 ms | 0.31 ms | 0.18 ms | 1.00 | 1.00 | 1507.0 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | ClosedXML | 14.63 ms | 0.60 ms | 0.35 ms | 5.13 | 5.13 | 11296.3 KB | 7.50 |  |  | 413.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | EPPlus | 26.72 ms | 0.61 ms | 0.35 ms | 9.37 | 9.37 | 14960.3 KB | 9.93 |  |  | 837.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | OfficeIMO.Excel | 2.48 ms | 0.28 ms | 0.16 ms | 1.00 | 1.00 | 1138.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | ClosedXML | 13.07 ms | 1.60 ms | 0.92 ms | 5.27 | 5.27 | 9021.2 KB | 7.93 |  |  | 426.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | EPPlus 4.5.3.3 | 17.73 ms |  |  | 7.15 | 7.15 |  |  |  |  | 614.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | EPPlus | 24.64 ms | 1.20 ms | 0.69 ms | 9.93 | 9.93 | 12827.4 KB | 11.27 |  |  | 893.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | OfficeIMO.Excel | 3.13 ms | 0.11 ms | 0.06 ms | 1.00 | 1.00 | 1435.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | ClosedXML | 16.16 ms | 0.48 ms | 0.28 ms | 5.16 | 5.16 | 9711.1 KB | 6.76 |  |  | 415.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | EPPlus 4.5.3.3 | 18.25 ms |  |  | 5.82 | 5.82 |  |  |  |  | 482.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | EPPlus | 26.01 ms | 0.60 ms | 0.35 ms | 8.30 | 8.30 | 14722.7 KB | 10.25 |  |  | 729.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | OfficeIMO.Excel | 7.40 ms | 0.38 ms | 0.22 ms | 1.00 | 1.00 | 2064.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | MiniExcel | 13.88 ms | 1.61 ms | 0.93 ms | 1.88 | 1.88 | 28695.1 KB | 13.90 |  |  | 87.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | ClosedXML | 55.46 ms | 10.54 ms | 6.08 ms | 7.50 | 7.50 | 18913.3 KB | 9.16 |  |  | 649.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | EPPlus | 68.09 ms | 9.58 ms | 5.53 ms | 9.21 | 9.21 | 17701.5 KB | 8.57 |  |  | 820.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | OfficeIMO.Excel | 8.42 ms | 0.40 ms | 0.23 ms | 1.00 | 1.00 | 2880.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | MiniExcel | 14.68 ms | 4.53 ms | 2.62 ms | 1.74 | 1.74 | 29747.0 KB | 10.33 |  |  | 74.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | EPPlus | 80.34 ms | 6.67 ms | 3.85 ms | 9.54 | 9.54 | 21891.4 KB | 7.60 |  |  | 854.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | ClosedXML | 86.61 ms | 18.54 ms | 10.70 ms | 10.29 | 10.29 | 27410.7 KB | 9.52 |  |  | 928.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | OfficeIMO.Excel | 6.43 ms | 0.91 ms | 0.52 ms | 1.00 | 1.00 | 2067.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | MiniExcel | 23.18 ms | 1.40 ms | 0.81 ms | 3.61 | 3.61 | 29229.4 KB | 14.14 |  |  | 260.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | EPPlus 4.5.3.3 | 43.50 ms |  |  | 6.77 | 6.77 |  |  |  |  | 577.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | ClosedXML | 102.12 ms | 30.80 ms | 17.79 ms | 15.89 | 15.89 | 18878.2 KB | 9.13 |  |  | 1489.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | EPPlus | 108.94 ms | 5.97 ms | 3.45 ms | 16.96 | 16.96 | 19431.0 KB | 9.40 |  |  | 1595.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | OfficeIMO.Excel | 4.73 ms | 0.23 ms | 0.13 ms | 1.00 | 1.00 | 1774.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | MiniExcel | 8.96 ms | 0.70 ms | 0.40 ms | 1.89 | 1.89 | 23044.1 KB | 12.98 |  |  | 89.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | EPPlus 4.5.3.3 | 35.33 ms |  |  | 7.47 | 7.47 |  |  |  |  | 647.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | EPPlus | 39.86 ms | 3.80 ms | 2.19 ms | 8.43 | 8.43 | 16645.9 KB | 9.38 |  |  | 743.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | ClosedXML | 41.15 ms | 2.42 ms | 1.40 ms | 8.70 | 8.70 | 19008.4 KB | 10.71 |  |  | 770.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | OfficeIMO.Excel | 5.30 ms | 0.41 ms | 0.24 ms | 1.00 | 1.00 | 1748.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | LargeXlsx | 9.57 ms | 0.49 ms | 0.29 ms | 1.81 | 1.81 | 1149.0 KB | 0.66 |  |  | 80.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | MiniExcel | 10.75 ms | 2.03 ms | 1.17 ms | 2.03 | 2.03 | 23062.6 KB | 13.19 |  |  | 102.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | ClosedXML | 38.37 ms | 4.29 ms | 2.48 ms | 7.24 | 7.24 | 11581.0 KB | 6.62 |  |  | 624.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | EPPlus 4.5.3.3 | 40.92 ms |  |  | 7.72 | 7.72 |  |  |  |  | 672.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | EPPlus | 51.25 ms | 5.77 ms | 3.33 ms | 9.67 | 9.67 | 16646.5 KB | 9.52 |  |  | 867.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | OfficeIMO.Excel | 4.72 ms | 0.25 ms | 0.14 ms | 1.00 | 1.00 | 1487.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | MiniExcel | 8.84 ms | 0.60 ms | 0.35 ms | 1.88 | 1.88 | 22789.4 KB | 15.32 |  |  | 87.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | ClosedXML | 39.20 ms | 1.42 ms | 0.82 ms | 8.31 | 8.31 | 18735.1 KB | 12.60 |  |  | 731.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | EPPlus | 39.57 ms | 1.02 ms | 0.59 ms | 8.39 | 8.39 | 16372.7 KB | 11.01 |  |  | 739.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | OfficeIMO.Excel | 4.99 ms | 0.52 ms | 0.30 ms | 1.00 | 1.00 | 1760.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | MiniExcel | 8.24 ms | 0.51 ms | 0.30 ms | 1.65 | 1.65 | 23062.8 KB | 13.10 |  |  | 65.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | ClosedXML | 41.21 ms | 3.72 ms | 2.15 ms | 8.25 | 8.25 | 19008.7 KB | 10.80 |  |  | 725.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | EPPlus 4.5.3.3 | 41.96 ms |  |  | 8.40 | 8.40 |  |  |  |  | 740.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | EPPlus | 42.63 ms | 3.49 ms | 2.02 ms | 8.54 | 8.54 | 16646.2 KB | 9.45 |  |  | 753.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | OfficeIMO.Excel | 6.61 ms | 0.68 ms | 0.39 ms | 1.00 | 1.00 | 1403.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | MiniExcel | 26.19 ms | 11.86 ms | 6.85 ms | 3.96 | 3.96 | 26825.1 KB | 19.12 |  |  | 296.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | EPPlus 4.5.3.3 | 114.98 ms |  |  | 17.39 | 17.39 |  |  |  |  | 1639.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | EPPlus | 228.66 ms | 85.20 ms | 49.19 ms | 34.59 | 34.59 | 49158.1 KB | 35.03 |  |  | 3359.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | ClosedXML | 293.37 ms | 52.78 ms | 30.47 ms | 44.38 | 44.38 | 58350.2 KB | 41.58 |  |  | 4338.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | OfficeIMO.Excel | 2.79 ms | 0.65 ms | 0.37 ms | 1.00 | 1.00 | 1541.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | ClosedXML | 18.59 ms | 0.47 ms | 0.27 ms | 6.66 | 6.66 | 12039.8 KB | 7.81 |  |  | 566.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | EPPlus 4.5.3.3 | 23.82 ms |  |  | 8.54 | 8.54 |  |  |  |  | 753.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | EPPlus | 40.43 ms | 2.70 ms | 1.56 ms | 14.49 | 14.49 | 18110.5 KB | 11.75 |  |  | 1348.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-dictionary-objects-table-direct | OfficeIMO.Excel | 6.84 ms | 1.01 ms | 0.58 ms | 1.00 | 1.00 | 2051.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | LargeXlsx | 4.37 ms | 0.49 ms | 0.28 ms | 0.78 | 1.00 | 802.5 KB | 0.34 |  |  | 22.4% faster than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 5.62 ms | 0.34 ms | 0.20 ms | 1.00 | 1.29 | 2341.7 KB | 1.00 |  |  | Loss +28.8% |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | MiniExcel | 8.92 ms | 0.72 ms | 0.41 ms | 1.59 | 2.04 | 25190.5 KB | 10.76 |  |  | 58.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | ClosedXML | 39.17 ms | 2.04 ms | 1.18 ms | 6.97 | 8.97 | 16973.5 KB | 7.25 |  |  | 596.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | EPPlus | 51.84 ms | 2.57 ms | 1.49 ms | 9.22 | 11.87 | 20105.1 KB | 8.59 |  |  | 821.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | LargeXlsx | 4.27 ms | 0.52 ms | 0.30 ms | 0.76 | 1.00 | 802.5 KB | 0.53 |  |  | 23.7% faster than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 5.59 ms | 0.52 ms | 0.30 ms | 1.00 | 1.31 | 1507.7 KB | 1.00 |  |  | Loss +31.1% |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | MiniExcel | 8.44 ms | 0.04 ms | 0.03 ms | 1.51 | 1.98 | 25190.5 KB | 16.71 |  |  | 51.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | ClosedXML | 39.06 ms | 2.75 ms | 1.59 ms | 6.99 | 9.16 | 16973.5 KB | 11.26 |  |  | 598.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | EPPlus | 52.27 ms | 2.96 ms | 1.71 ms | 9.35 | 12.25 | 20105.1 KB | 13.33 |  |  | 834.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | LargeXlsx | 20.83 ms | 1.09 ms | 0.63 ms | 1.00 | 1.00 | 2810.7 KB | 0.62 |  |  | Tie vs OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 20.86 ms | 2.58 ms | 1.49 ms | 1.00 | 1.00 | 4502.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | MiniExcel | 35.09 ms | 1.46 ms | 0.84 ms | 1.68 | 1.68 | 48414.8 KB | 10.75 |  |  | 68.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | ClosedXML | 120.21 ms | 3.88 ms | 2.24 ms | 5.76 | 5.77 | 51647.0 KB | 11.47 |  |  | 476.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | EPPlus | 159.19 ms | 3.59 ms | 2.07 ms | 7.63 | 7.64 | 69139.6 KB | 15.36 |  |  | 663.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | LargeXlsx | 1.51 ms | 0.03 ms | 0.02 ms | 0.61 | 1.00 | 296.4 KB | 0.19 |  |  | 38.9% faster than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | OfficeIMO.Excel | 2.47 ms | 0.25 ms | 0.15 ms | 1.00 | 1.64 | 1576.3 KB | 1.00 |  |  | Loss +63.6% |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | MiniExcel | 4.64 ms | 0.24 ms | 0.14 ms | 1.88 | 3.08 | 19710.9 KB | 12.50 |  |  | 88.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | ClosedXML | 16.10 ms | 0.34 ms | 0.20 ms | 6.53 | 10.67 | 11197.4 KB | 7.10 |  |  | 552.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | EPPlus 4.5.3.3 | 19.29 ms |  |  | 7.82 | 12.78 |  |  |  |  | 681.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | EPPlus | 25.44 ms | 0.33 ms | 0.19 ms | 10.31 | 16.86 | 14365.2 KB | 9.11 |  |  | 930.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 1.30 ms | 0.06 ms | 0.04 ms | 0.76 | 1.00 | 447.0 KB | 0.41 |  |  | 23.7% faster than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 1.70 ms | 0.06 ms | 0.03 ms | 1.00 | 1.31 | 1092.0 KB | 1.00 |  |  | Loss +31.0% |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | ClosedXML | 14.13 ms | 0.90 ms | 0.52 ms | 8.31 | 10.89 | 10235.8 KB | 9.37 |  |  | 731.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | EPPlus | 23.28 ms | 0.54 ms | 0.31 ms | 13.69 | 17.93 | 13052.2 KB | 11.95 |  |  | 1268.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | LargeXlsx | 3.54 ms | 0.33 ms | 0.19 ms | 0.83 | 1.00 | 758.3 KB | 0.36 |  |  | 17.0% faster than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 4.27 ms | 0.26 ms | 0.15 ms | 1.00 | 1.20 | 2081.1 KB | 1.00 |  |  | Loss +20.5% |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | MiniExcel | 9.58 ms | 0.21 ms | 0.12 ms | 2.25 | 2.70 | 23221.8 KB | 11.16 |  |  | 124.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | ClosedXML | 35.11 ms | 1.36 ms | 0.79 ms | 8.22 | 9.91 | 22221.3 KB | 10.68 |  |  | 722.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 39.94 ms |  |  | 9.36 | 11.27 |  |  |  |  | 835.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | EPPlus | 45.56 ms | 2.80 ms | 1.62 ms | 10.67 | 12.86 | 24693.7 KB | 11.87 |  |  | 967.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | OfficeIMO.Excel | 2.71 ms | 0.60 ms | 0.35 ms | 1.00 | 1.00 | 1494.0 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | ClosedXML | 16.69 ms | 0.74 ms | 0.43 ms | 6.16 | 6.16 | 11296.3 KB | 7.56 |  |  | 516.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | EPPlus | 25.96 ms | 0.91 ms | 0.53 ms | 9.58 | 9.58 | 14960.0 KB | 10.01 |  |  | 858.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | Sylvan.Data.Excel | 3.66 ms | 0.38 ms | 0.22 ms | 0.82 | 1.00 | 758.6 KB | 0.43 |  |  | 18.1% faster than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | LargeXlsx | 3.82 ms | 0.22 ms | 0.13 ms | 0.86 | 1.04 | 1032.5 KB | 0.59 |  |  | 14.5% faster than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | OfficeIMO.Excel | 4.47 ms | 0.30 ms | 0.17 ms | 1.00 | 1.22 | 1763.0 KB | 1.00 |  |  | Loss +22.1% |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | MiniExcel | 9.18 ms | 1.51 ms | 0.87 ms | 2.05 | 2.51 | 23043.8 KB | 13.07 |  |  | 105.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | ClosedXML | 30.40 ms | 1.52 ms | 0.87 ms | 6.80 | 8.30 | 11581.0 KB | 6.57 |  |  | 580.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | EPPlus 4.5.3.3 | 32.21 ms |  |  | 7.21 | 8.80 |  |  |  |  | 620.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | EPPlus | 41.21 ms | 1.97 ms | 1.14 ms | 9.22 | 11.26 | 16646.2 KB | 9.44 |  |  | 821.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | OfficeIMO.Excel | 7.82 ms | 5.05 ms | 2.92 ms | 1.00 | 1.00 | 2140.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | LargeXlsx | 11.68 ms | 3.40 ms | 1.96 ms | 1.49 | 1.49 | 1123.9 KB | 0.53 |  |  | 49.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | MiniExcel | 12.17 ms | 1.06 ms | 0.61 ms | 1.56 | 1.56 | 29747.0 KB | 13.90 |  |  | 55.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | ClosedXML | 69.90 ms | 2.36 ms | 1.36 ms | 8.94 | 8.94 | 27410.6 KB | 12.80 |  |  | 794.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | EPPlus | 74.05 ms | 10.64 ms | 6.14 ms | 9.47 | 9.47 | 21891.0 KB | 10.23 |  |  | 847.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | OfficeIMO.Excel | 4.76 ms | 0.08 ms | 0.05 ms | 1.00 | 1.00 | 1676.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | LargeXlsx | 5.04 ms | 0.82 ms | 0.47 ms | 1.06 | 1.06 | 857.6 KB | 0.51 |  |  | 6.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | MiniExcel | 18.38 ms | 1.82 ms | 1.05 ms | 3.87 | 3.87 | 35917.8 KB | 21.42 |  |  | 286.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | ClosedXML | 89.58 ms | 1.40 ms | 0.81 ms | 18.84 | 18.84 | 71478.2 KB | 42.63 |  |  | 1783.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | OfficeIMO.Excel | 2.00 ms | 0.03 ms | 0.01 ms | 1.00 | 1.00 | 2440.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | MiniExcel | 4.69 ms | 0.94 ms | 0.54 ms | 2.34 | 2.34 | 21137.5 KB | 8.66 |  |  | 134.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | EPPlus 4.5.3.3 | 10.18 ms |  |  | 5.08 | 5.08 |  |  |  |  | 407.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | ClosedXML | 11.30 ms | 0.24 ms | 0.14 ms | 5.64 | 5.64 | 11299.2 KB | 4.63 |  |  | 463.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | EPPlus | 25.44 ms | 3.33 ms | 1.92 ms | 12.70 | 12.70 | 12804.4 KB | 5.25 |  |  | 1169.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | OfficeIMO.Excel | 3.09 ms | 0.07 ms | 0.04 ms | 1.00 | 1.00 | 2617.0 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | ClosedXML | 12.66 ms | 0.84 ms | 0.48 ms | 4.10 | 4.10 | 11299.2 KB | 4.32 |  |  | 309.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | EPPlus 4.5.3.3 | 20.51 ms |  |  | 6.63 | 6.63 |  |  |  |  | 563.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | EPPlus | 22.29 ms | 1.95 ms | 1.13 ms | 7.21 | 7.21 | 12804.9 KB | 4.89 |  |  | 621.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | OfficeIMO.Excel | 2.56 ms | 0.26 ms | 0.15 ms | 1.00 | 1.00 | 2379.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | ClosedXML | 15.67 ms | 1.02 ms | 0.59 ms | 6.13 | 6.13 | 13127.1 KB | 5.52 |  |  | 512.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | EPPlus | 25.22 ms | 1.27 ms | 0.74 ms | 9.86 | 9.86 | 13892.9 KB | 5.84 |  |  | 886.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | OfficeIMO.Excel | 2.31 ms | 0.17 ms | 0.10 ms | 1.00 | 1.00 | 1579.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | ClosedXML | 11.31 ms | 0.62 ms | 0.36 ms | 4.89 | 4.89 | 9226.5 KB | 5.84 |  |  | 388.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | EPPlus | 21.01 ms | 1.00 ms | 0.58 ms | 9.08 | 9.08 | 11332.4 KB | 7.17 |  |  | 807.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | LargeXlsx | 3.22 ms | 0.04 ms | 0.02 ms | 0.80 | 1.00 | 758.3 KB | 0.43 |  |  | 19.7% faster than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 4.00 ms | 0.32 ms | 0.19 ms | 1.00 | 1.25 | 1769.2 KB | 1.00 |  |  | Loss +24.5% |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | MiniExcel | 9.83 ms | 1.89 ms | 1.09 ms | 2.45 | 3.06 | 23222.3 KB | 13.13 |  |  | 145.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | ClosedXML | 30.57 ms | 1.24 ms | 0.71 ms | 7.63 | 9.51 | 11581.0 KB | 6.55 |  |  | 663.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 35.34 ms |  |  | 8.83 | 10.99 |  |  |  |  | 782.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | EPPlus | 43.05 ms | 4.89 ms | 2.83 ms | 10.75 | 13.39 | 16646.4 KB | 9.41 |  |  | 975.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | LargeXlsx | 3.14 ms | 0.14 ms | 0.08 ms | 0.88 | 1.00 | 758.3 KB | 0.57 |  |  | 11.9% faster than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | OfficeIMO.Excel | 3.56 ms | 0.04 ms | 0.02 ms | 1.00 | 1.13 | 1339.3 KB | 1.00 |  |  | Loss +13.5% |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | MiniExcel | 8.48 ms | 0.78 ms | 0.45 ms | 2.38 | 2.70 | 23222.2 KB | 17.34 |  |  | 137.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | ClosedXML | 27.41 ms | 0.43 ms | 0.25 ms | 7.69 | 8.73 | 11581.0 KB | 8.65 |  |  | 669.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | EPPlus 4.5.3.3 | 36.04 ms |  |  | 10.11 | 11.48 |  |  |  |  | 911.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | EPPlus | 41.99 ms | 3.58 ms | 2.07 ms | 11.78 | 13.37 | 16646.1 KB | 12.43 |  |  | 1078.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 3.06 ms | 0.07 ms | 0.04 ms | 0.70 | 1.00 | 758.3 KB | 0.51 |  |  | 30.3% faster than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 4.39 ms | 0.34 ms | 0.20 ms | 1.00 | 1.44 | 1497.5 KB | 1.00 |  |  | Loss +43.6% |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | MiniExcel | 9.37 ms | 1.35 ms | 0.78 ms | 2.13 | 3.06 | 23222.3 KB | 15.51 |  |  | 113.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | ClosedXML | 27.12 ms | 0.69 ms | 0.40 ms | 6.18 | 8.87 | 11581.0 KB | 7.73 |  |  | 517.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | EPPlus | 42.16 ms | 2.00 ms | 1.15 ms | 9.60 | 13.79 | 16646.1 KB | 11.12 |  |  | 860.3% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | Sylvan.Data.Excel | 38.38 ms | 1.54 ms | 0.89 ms | 0.84 | 1.00 | 394.1 KB | 0.02 |  |  | 15.6% faster than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | OfficeIMO.Excel | 45.46 ms | 2.71 ms | 1.57 ms | 1.00 | 1.18 | 23621.9 KB | 1.00 |  |  | Loss +18.5% |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | ExcelDataReader | 114.89 ms | 4.55 ms | 2.63 ms | 2.53 | 2.99 | 69530.7 KB | 2.94 |  |  | 152.7% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | MiniExcel | 145.24 ms | 2.02 ms | 1.17 ms | 3.20 | 3.78 | 215349.0 KB | 9.12 |  |  | 219.5% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | Sylvan.Data.Excel | 36.76 ms | 0.29 ms | 0.17 ms | 0.82 | 1.00 | 394.1 KB | 0.02 |  |  | 18.1% faster than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | OfficeIMO.Excel | 44.90 ms | 1.18 ms | 0.68 ms | 1.00 | 1.22 | 24403.9 KB | 1.00 |  |  | Loss +22.2% |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | ExcelDataReader | 115.71 ms | 6.01 ms | 3.47 ms | 2.58 | 3.15 | 69530.7 KB | 2.85 |  |  | 157.7% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | MiniExcel | 147.87 ms | 4.36 ms | 2.52 ms | 3.29 | 4.02 | 215349.0 KB | 8.82 |  |  | 229.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | append-plain-rows | LargeXlsx | 10.99 ms | 0.30 ms | 0.17 ms | 0.75 | 1.00 | 2771.0 KB | 0.26 | 605.0 KB | 0.99 | 24.6% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | append-plain-rows | OfficeIMO.Excel | 14.59 ms | 1.11 ms | 0.64 ms | 1.00 | 1.33 | 10842.5 KB | 1.00 | 610.4 KB | 1.00 | Loss +32.7% |
| 25000 | package-profile | package | Package size | append-plain-rows | MiniExcel | 33.66 ms | 3.03 ms | 1.75 ms | 2.31 | 3.06 | 58242.9 KB | 5.37 | 642.3 KB | 1.05 | 130.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | append-plain-rows | ClosedXML | 133.64 ms | 4.07 ms | 2.35 ms | 9.16 | 12.15 | 104233.1 KB | 9.61 | 540.6 KB | 0.89 | 816.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | append-plain-rows | EPPlus | 209.65 ms | 2.45 ms | 1.41 ms | 14.37 | 19.07 | 100373.5 KB | 9.26 | 525.6 KB | 0.86 | 1337.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | autofit-existing | OfficeIMO.Excel | 79.41 ms | 0.83 ms | 0.48 ms | 1.00 | 1.00 | 15708.4 KB | 1.00 | 1429.3 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | autofit-existing | EPPlus | 491.40 ms | 51.56 ms | 29.77 ms | 6.19 | 6.19 | 250950.0 KB | 15.98 | 1091.0 KB | 0.76 | 518.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | autofit-existing | ClosedXML | 1386.30 ms | 33.56 ms | 19.37 ms | 17.46 | 17.46 | 829716.8 KB | 52.82 | 1140.9 KB | 0.80 | 1645.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | large-shared-strings | OfficeIMO.Excel | 14.86 ms | 0.42 ms | 0.24 ms | 1.00 | 1.00 | 15744.9 KB | 1.00 | 529.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | large-shared-strings | MiniExcel | 29.02 ms | 0.93 ms | 0.53 ms | 1.95 | 1.95 | 73760.2 KB | 4.68 | 581.0 KB | 1.10 | 95.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | large-shared-strings | ClosedXML | 110.59 ms | 2.25 ms | 1.30 ms | 7.44 | 7.44 | 104241.3 KB | 6.62 | 460.1 KB | 0.87 | 644.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | large-shared-strings | EPPlus | 180.83 ms | 4.59 ms | 2.65 ms | 12.17 | 12.17 | 84410.0 KB | 5.36 | 444.7 KB | 0.84 | 1117.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-autofilter | OfficeIMO.Excel | 37.72 ms | 3.95 ms | 2.28 ms | 1.00 | 1.00 | 11494.9 KB | 1.00 | 1428.5 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-autofilter | ClosedXML | 362.90 ms | 92.66 ms | 53.50 ms | 9.62 | 9.62 | 210663.8 KB | 18.33 | 1140.0 KB | 0.80 | 862.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-autofilter | EPPlus | 435.34 ms | 22.37 ms | 12.92 ms | 11.54 | 11.54 | 211871.5 KB | 18.43 | 1090.1 KB | 0.76 | 1054.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-charts | OfficeIMO.Excel | 33.59 ms | 0.59 ms | 0.34 ms | 1.00 | 1.00 | 12552.7 KB | 1.00 | 1433.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-charts | EPPlus | 350.87 ms | 4.61 ms | 2.66 ms | 10.44 | 10.44 | 214905.3 KB | 17.12 | 1092.9 KB | 0.76 | 944.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-conditional-formatting | OfficeIMO.Excel | 34.63 ms | 2.23 ms | 1.28 ms | 1.00 | 1.00 | 11560.2 KB | 1.00 | 1428.8 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-conditional-formatting | ClosedXML | 313.89 ms | 12.34 ms | 7.12 ms | 9.06 | 9.06 | 210711.7 KB | 18.23 | 1140.1 KB | 0.80 | 806.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-conditional-formatting | EPPlus | 384.78 ms | 3.38 ms | 1.95 ms | 11.11 | 11.11 | 211912.9 KB | 18.33 | 1090.2 KB | 0.76 | 1011.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-data-validation | OfficeIMO.Excel | 31.39 ms | 0.69 ms | 0.40 ms | 1.00 | 1.00 | 11510.5 KB | 1.00 | 1428.6 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-data-validation | ClosedXML | 304.58 ms | 22.40 ms | 12.93 ms | 9.70 | 9.70 | 210672.7 KB | 18.30 | 1140.1 KB | 0.80 | 870.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-data-validation | EPPlus | 369.82 ms | 15.70 ms | 9.07 ms | 11.78 | 11.78 | 211857.4 KB | 18.41 | 1090.1 KB | 0.76 | 1078.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-freeze-panes | OfficeIMO.Excel | 32.83 ms | 1.19 ms | 0.69 ms | 1.00 | 1.00 | 11497.3 KB | 1.00 | 1428.6 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-freeze-panes | ClosedXML | 309.57 ms | 13.27 ms | 7.66 ms | 9.43 | 9.43 | 210646.8 KB | 18.32 | 1140.0 KB | 0.80 | 843.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-freeze-panes | EPPlus | 387.10 ms | 22.09 ms | 12.76 ms | 11.79 | 11.79 | 211883.3 KB | 18.43 | 1090.2 KB | 0.76 | 1079.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-pivot-table | OfficeIMO.Excel | 226.20 ms | 12.17 ms | 7.02 ms | 1.00 | 1.00 | 131923.0 KB | 1.00 | 1979.8 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-pivot-table | EPPlus | 389.18 ms | 15.33 ms | 8.85 ms | 1.72 | 1.72 | 230800.9 KB | 1.75 | 1093.4 KB | 0.55 | 72.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-all-in-one | OfficeIMO.Excel | 301.25 ms | 65.20 ms | 37.64 ms | 1.00 | 1.00 | 133446.8 KB | 1.00 | 1985.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-all-in-one | EPPlus | 601.05 ms | 176.06 ms | 101.65 ms | 2.00 | 2.00 | 277077.5 KB | 2.08 | 1097.7 KB | 0.55 | 99.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-chart-first | OfficeIMO.Excel | 161.07 ms | 11.81 ms | 6.82 ms | 1.00 | 1.00 | 43563.0 KB | 1.00 | 1985.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-chart-first | EPPlus | 788.97 ms | 147.70 ms | 85.28 ms | 4.90 | 4.90 | 277075.4 KB | 6.36 | 1097.7 KB | 0.55 | 389.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-core | OfficeIMO.Excel | 35.75 ms | 0.75 ms | 0.43 ms | 1.00 | 1.00 | 11648.7 KB | 1.00 | 1430.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-core | EPPlus | 389.91 ms | 15.50 ms | 8.95 ms | 10.91 | 10.91 | 255065.8 KB | 21.90 | 1091.5 KB | 0.76 | 990.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-core | ClosedXML | 771.95 ms | 9.29 ms | 5.37 ms | 21.59 | 21.59 | 680116.8 KB | 58.39 | 1141.3 KB | 0.80 | 2059.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-extra-column | OfficeIMO.Excel | 352.98 ms | 10.86 ms | 6.27 ms | 1.00 | 1.00 | 144823.3 KB | 1.00 | 2110.8 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-extra-column | EPPlus | 527.93 ms | 53.94 ms | 31.14 ms | 1.50 | 1.50 | 302759.9 KB | 2.09 | 1166.3 KB | 0.55 | 49.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-no-autofit | OfficeIMO.Excel | 417.93 ms | 125.62 ms | 72.53 ms | 1.00 | 1.00 | 133435.3 KB | 1.00 | 1985.8 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-no-autofit | EPPlus | 559.86 ms | 33.06 ms | 19.09 ms | 1.34 | 1.34 | 234782.5 KB | 1.76 | 1097.7 KB | 0.55 | 34.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-post-mutation | OfficeIMO.Excel | 243.05 ms | 9.08 ms | 5.24 ms | 1.00 | 1.00 | 133455.0 KB | 1.00 | 1986.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-post-mutation | EPPlus | 429.98 ms | 3.91 ms | 2.26 ms | 1.77 | 1.77 | 277077.5 KB | 2.08 | 1097.8 KB | 0.55 | 76.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-shuffled-columns | OfficeIMO.Excel | 347.07 ms | 17.10 ms | 9.87 ms | 1.00 | 1.00 | 133506.8 KB | 1.00 | 2046.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-shuffled-columns | EPPlus | 507.12 ms | 19.96 ms | 11.52 ms | 1.46 | 1.46 | 277070.2 KB | 2.08 | 1098.4 KB | 0.54 | 46.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook | OfficeIMO.Excel | 327.22 ms | 17.34 ms | 10.01 ms | 1.00 | 1.00 | 175194.8 KB | 1.00 | 2672.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | report-workbook | EPPlus | 560.40 ms | 5.82 ms | 3.36 ms | 1.71 | 1.71 | 364709.5 KB | 2.08 | 1517.2 KB | 0.57 | 71.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook-core | OfficeIMO.Excel | 60.71 ms | 13.09 ms | 7.56 ms | 1.00 | 1.00 | 10979.4 KB | 1.00 | 1850.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | report-workbook-core | EPPlus | 669.00 ms | 133.16 ms | 76.88 ms | 11.02 | 11.02 | 342842.9 KB | 31.23 | 1512.6 KB | 0.82 | 1002.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook-core | ClosedXML | 1335.23 ms | 448.29 ms | 258.82 ms | 21.99 | 21.99 | 975774.1 KB | 88.87 | 1579.8 KB | 0.85 | 2099.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook-datatable | OfficeIMO.Excel | 373.27 ms | 27.61 ms | 15.94 ms | 1.00 | 1.00 | 177941.8 KB | 1.00 | 2672.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | report-workbook-datatable | EPPlus | 558.15 ms | 15.15 ms | 8.75 ms | 1.50 | 1.50 | 247823.0 KB | 1.39 | 1517.2 KB | 0.57 | 49.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook-datatable-core | OfficeIMO.Excel | 49.96 ms | 1.70 ms | 0.98 ms | 1.00 | 1.00 | 13725.0 KB | 1.00 | 1850.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | report-workbook-datatable-core | EPPlus | 508.54 ms | 6.77 ms | 3.91 ms | 10.18 | 10.18 | 225957.1 KB | 16.46 | 1512.6 KB | 0.82 | 917.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook-datatable-core | ClosedXML | 1019.36 ms | 47.95 ms | 27.69 ms | 20.40 | 20.40 | 832229.0 KB | 60.64 | 1579.8 KB | 0.85 | 1940.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | LargeXlsx | 39.19 ms | 0.56 ms | 0.32 ms | 0.90 | 1.00 | 10795.2 KB | 0.92 | 2444.6 KB | 1.10 | 10.3% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | OfficeIMO.Excel | 43.72 ms | 1.30 ms | 0.75 ms | 1.00 | 1.12 | 11708.2 KB | 1.00 | 2228.8 KB | 1.00 | Loss +11.5% |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | MiniExcel | 147.10 ms | 1.92 ms | 1.11 ms | 3.36 | 3.75 | 226875.6 KB | 19.38 | 2410.6 KB | 1.08 | 236.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | ClosedXML | 908.18 ms | 1.83 ms | 1.06 ms | 20.77 | 23.17 | 759818.4 KB | 64.90 | 2581.2 KB | 1.16 | 1977.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-bulk-report | OfficeIMO.Excel | 36.78 ms | 1.19 ms | 0.69 ms | 1.00 | 1.00 | 11561.8 KB | 1.00 | 1429.3 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-bulk-report | MiniExcel | 69.37 ms | 1.56 ms | 0.90 ms | 1.89 | 1.89 | 125550.4 KB | 10.86 | 1521.1 KB | 1.06 | 88.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-bulk-report | EPPlus | 412.29 ms | 10.22 ms | 5.90 ms | 11.21 | 11.21 | 254958.3 KB | 22.05 | 1091.0 KB | 0.76 | 1021.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-bulk-report | ClosedXML | 830.33 ms | 7.58 ms | 4.38 ms | 22.58 | 22.58 | 565953.5 KB | 48.95 | 1140.9 KB | 0.80 | 2157.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellformula | OfficeIMO.Excel | 18.30 ms | 0.07 ms | 0.04 ms | 1.00 | 1.00 | 10112.0 KB | 1.00 | 670.3 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellformula | ClosedXML | 171.37 ms | 2.40 ms | 1.38 ms | 9.36 | 9.36 | 113853.5 KB | 11.26 | 643.2 KB | 0.96 | 836.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellformula | EPPlus | 299.14 ms | 4.49 ms | 2.59 ms | 16.34 | 16.34 | 140731.9 KB | 13.92 | 593.9 KB | 0.89 | 1534.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-empty-strings | OfficeIMO.Excel | 12.16 ms | 1.02 ms | 0.59 ms | 1.00 | 1.00 | 6896.4 KB | 1.00 | 451.4 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-empty-strings | ClosedXML | 113.45 ms | 2.05 ms | 1.18 ms | 9.33 | 9.33 | 92902.1 KB | 13.47 | 398.1 KB | 0.88 | 833.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-empty-strings | EPPlus | 176.51 ms | 0.84 ms | 0.48 ms | 14.52 | 14.52 | 74492.8 KB | 10.80 | 390.6 KB | 0.87 | 1351.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-numbers | OfficeIMO.Excel | 14.84 ms | 0.87 ms | 0.50 ms | 1.00 | 1.00 | 5970.9 KB | 1.00 | 462.6 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-numbers | ClosedXML | 110.39 ms | 3.75 ms | 2.16 ms | 7.44 | 7.44 | 84206.7 KB | 14.10 | 411.4 KB | 0.89 | 643.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-numbers | EPPlus | 193.38 ms | 5.61 ms | 3.24 ms | 13.03 | 13.03 | 86377.5 KB | 14.47 | 406.5 KB | 0.88 | 1203.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-mixed | OfficeIMO.Excel | 16.32 ms | 0.09 ms | 0.05 ms | 1.00 | 1.00 | 8332.9 KB | 1.00 | 585.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-object-mixed | ClosedXML | 147.22 ms | 0.53 ms | 0.31 ms | 9.02 | 9.02 | 111118.7 KB | 13.33 | 532.9 KB | 0.91 | 801.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-mixed | EPPlus | 206.18 ms | 2.51 ms | 1.45 ms | 12.63 | 12.63 | 113245.1 KB | 13.59 | 544.3 KB | 0.93 | 1163.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse | OfficeIMO.Excel | 18.01 ms | 0.83 ms | 0.48 ms | 1.00 | 1.00 | 7416.2 KB | 1.00 | 607.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse | ClosedXML | 134.04 ms | 0.75 ms | 0.43 ms | 7.44 | 7.44 | 105223.9 KB | 14.19 | 468.0 KB | 0.77 | 644.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse | EPPlus | 204.02 ms | 1.13 ms | 0.65 ms | 11.33 | 11.33 | 106316.9 KB | 14.34 | 494.4 KB | 0.81 | 1033.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 18.94 ms | 0.82 ms | 0.47 ms | 1.00 | 1.00 | 7416.3 KB | 1.00 | 607.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | ClosedXML | 147.75 ms | 6.94 ms | 4.01 ms | 7.80 | 7.80 | 105223.9 KB | 14.19 | 468.0 KB | 0.77 | 680.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | EPPlus | 212.85 ms | 9.77 ms | 5.64 ms | 11.24 | 11.24 | 106316.9 KB | 14.34 | 494.4 KB | 0.81 | 1023.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-scalars | OfficeIMO.Excel | 10.96 ms | 0.89 ms | 0.52 ms | 1.00 | 1.00 | 6144.6 KB | 1.00 | 441.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-scalars | ClosedXML | 98.87 ms | 6.62 ms | 3.82 ms | 9.02 | 9.02 | 82591.3 KB | 13.44 | 394.9 KB | 0.89 | 801.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-scalars | EPPlus | 184.68 ms | 3.15 ms | 1.82 ms | 16.85 | 16.85 | 85127.4 KB | 13.85 | 379.3 KB | 0.86 | 1584.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings | OfficeIMO.Excel | 16.52 ms | 0.54 ms | 0.31 ms | 1.00 | 1.00 | 15360.4 KB | 1.00 | 527.8 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-strings | ClosedXML | 115.64 ms | 7.98 ms | 4.61 ms | 7.00 | 7.00 | 104241.3 KB | 6.79 | 460.1 KB | 0.87 | 599.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings | EPPlus | 233.76 ms | 2.12 ms | 1.23 ms | 14.15 | 14.15 | 84410.3 KB | 5.50 | 444.7 KB | 0.84 | 1314.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-distinct | OfficeIMO.Excel | 12.73 ms | 0.83 ms | 0.48 ms | 1.00 | 1.00 | 13824.1 KB | 1.00 | 499.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-distinct | ClosedXML | 151.01 ms | 4.73 ms | 2.73 ms | 11.86 | 11.86 | 131501.7 KB | 9.51 | 555.3 KB | 1.11 | 1086.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-distinct | EPPlus | 209.45 ms | 1.88 ms | 1.08 ms | 16.45 | 16.45 | 97729.6 KB | 7.07 | 565.1 KB | 1.13 | 1545.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-repeated | OfficeIMO.Excel | 11.60 ms | 0.17 ms | 0.10 ms | 1.00 | 1.00 | 7525.3 KB | 1.00 | 376.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-repeated | ClosedXML | 97.75 ms | 2.20 ms | 1.27 ms | 8.43 | 8.43 | 84520.0 KB | 11.23 | 331.8 KB | 0.88 | 742.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-repeated | EPPlus | 157.09 ms | 3.45 ms | 1.99 ms | 13.55 | 13.55 | 70033.4 KB | 9.31 | 300.8 KB | 0.80 | 1254.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-temporal | OfficeIMO.Excel | 18.80 ms | 0.30 ms | 0.17 ms | 1.00 | 1.00 | 7482.8 KB | 1.00 | 620.5 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-temporal | ClosedXML | 140.91 ms | 1.90 ms | 1.10 ms | 7.49 | 7.49 | 89323.7 KB | 11.94 | 483.0 KB | 0.78 | 649.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-temporal | EPPlus | 193.51 ms | 3.58 ms | 2.07 ms | 10.29 | 10.29 | 103800.0 KB | 13.87 | 495.1 KB | 0.80 | 929.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 10.16 ms | 0.86 ms | 0.50 ms | 0.97 | 1.00 | 3444.4 KB | 0.49 | 443.4 KB | 0.97 | 3.2% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 10.50 ms | 0.32 ms | 0.19 ms | 1.00 | 1.03 | 6961.7 KB | 1.00 | 455.5 KB | 1.00 | Loss +3.3% |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | ClosedXML | 126.65 ms | 10.51 ms | 6.07 ms | 12.06 | 12.47 | 96015.7 KB | 13.79 | 467.5 KB | 1.03 | 1106.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | EPPlus | 197.12 ms | 0.88 ms | 0.51 ms | 18.78 | 19.40 | 87466.9 KB | 12.56 | 484.1 KB | 1.06 | 1777.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | LargeXlsx | 31.14 ms | 3.77 ms | 2.17 ms | 0.87 | 1.00 | 5614.1 KB | 0.35 | 1386.5 KB | 1.00 | 13.1% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 35.85 ms | 3.50 ms | 2.02 ms | 1.00 | 1.15 | 16036.5 KB | 1.00 | 1384.9 KB | 1.00 | Loss +15.1% |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | MiniExcel | 68.64 ms | 3.08 ms | 1.78 ms | 1.91 | 2.20 | 93257.0 KB | 5.82 | 1521.1 KB | 1.10 | 91.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | ClosedXML | 322.56 ms | 5.14 ms | 2.97 ms | 9.00 | 10.36 | 210646.1 KB | 13.14 | 1139.9 KB | 0.82 | 799.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | EPPlus | 447.57 ms | 23.61 ms | 13.63 ms | 12.48 | 14.37 | 211849.9 KB | 13.21 | 1090.0 KB | 0.79 | 1148.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | Sylvan.Data.Excel | 29.75 ms | 1.46 ms | 0.85 ms | 0.74 | 1.00 | 5700.3 KB | 0.44 | 755.4 KB | 0.55 | 25.9% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | LargeXlsx | 34.57 ms | 0.36 ms | 0.21 ms | 0.86 | 1.16 | 8349.2 KB | 0.64 | 1386.5 KB | 1.00 | 13.9% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | OfficeIMO.Excel | 40.13 ms | 4.04 ms | 2.33 ms | 1.00 | 1.35 | 13002.3 KB | 1.00 | 1384.9 KB | 1.00 | Loss +34.9% |
| 25000 | package-profile | package | Package size | write-datareader-plain | MiniExcel | 76.72 ms | 1.58 ms | 0.91 ms | 1.91 | 2.58 | 92199.8 KB | 7.09 | 1521.0 KB | 1.10 | 91.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | ClosedXML | 293.14 ms | 22.51 ms | 13.00 ms | 7.31 | 9.85 | 104205.0 KB | 8.01 | 1139.9 KB | 0.82 | 630.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | EPPlus | 360.44 ms | 27.35 ms | 15.79 ms | 8.98 | 12.12 | 117437.7 KB | 9.03 | 1090.8 KB | 0.79 | 798.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table | OfficeIMO.Excel | 38.33 ms | 1.97 ms | 1.14 ms | 1.00 | 1.00 | 13020.3 KB | 1.00 | 1385.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-datareader-table | MiniExcel | 71.65 ms | 0.74 ms | 0.43 ms | 1.87 | 1.87 | 92200.0 KB | 7.08 | 1521.0 KB | 1.10 | 86.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table | EPPlus | 348.87 ms | 10.12 ms | 5.85 ms | 9.10 | 9.10 | 117437.3 KB | 9.02 | 1090.8 KB | 0.79 | 810.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table | ClosedXML | 397.09 ms | 31.72 ms | 18.31 ms | 10.36 | 10.36 | 173397.5 KB | 13.32 | 1140.7 KB | 0.82 | 935.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | OfficeIMO.Excel | 40.84 ms | 2.44 ms | 1.41 ms | 1.00 | 1.00 | 13026.6 KB | 1.00 | 1385.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | MiniExcel | 74.04 ms | 2.41 ms | 1.39 ms | 1.81 | 1.81 | 124495.5 KB | 9.56 | 1521.1 KB | 1.10 | 81.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | EPPlus | 373.41 ms | 5.63 ms | 3.25 ms | 9.14 | 9.14 | 159741.8 KB | 12.26 | 1091.0 KB | 0.79 | 814.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | ClosedXML | 790.87 ms | 18.10 ms | 10.45 ms | 19.36 | 19.36 | 566142.3 KB | 43.46 | 1140.9 KB | 0.82 | 1836.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | OfficeIMO.Excel | 34.06 ms | 2.69 ms | 1.55 ms | 1.00 | 1.00 | 9819.7 KB | 1.00 | 1329.2 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | LargeXlsx | 40.18 ms | 4.40 ms | 2.54 ms | 1.18 | 1.18 | 9265.9 KB | 0.94 | 1680.0 KB | 1.26 | 18.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | MiniExcel | 103.24 ms | 11.15 ms | 6.44 ms | 3.03 | 3.03 | 108129.1 KB | 11.01 | 1819.7 KB | 1.37 | 203.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | EPPlus | 533.62 ms | 21.83 ms | 12.61 ms | 15.67 | 15.67 | 135723.5 KB | 13.82 | 1390.4 KB | 1.05 | 1466.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | ClosedXML | 599.66 ms | 4.52 ms | 2.61 ms | 17.61 | 17.61 | 280372.9 KB | 28.55 | 1519.9 KB | 1.14 | 1660.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | OfficeIMO.Excel | 63.10 ms | 23.59 ms | 13.62 ms | 1.00 | 1.00 | 13458.5 KB | 1.00 | 1795.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | MiniExcel | 141.93 ms | 32.52 ms | 18.77 ms | 2.25 | 2.25 | 108128.8 KB | 8.03 | 1819.7 KB | 1.01 | 124.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | EPPlus | 658.67 ms | 117.10 ms | 67.61 ms | 10.44 | 10.44 | 135723.5 KB | 10.08 | 1390.4 KB | 0.77 | 943.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | ClosedXML | 707.99 ms | 120.86 ms | 69.78 ms | 11.22 | 11.22 | 280371.6 KB | 20.83 | 1519.9 KB | 0.85 | 1021.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables | OfficeIMO.Excel | 36.76 ms | 1.63 ms | 0.94 ms | 1.00 | 1.00 | 10288.1 KB | 1.00 | 1376.4 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-dataset-tables | MiniExcel | 82.30 ms | 1.12 ms | 0.65 ms | 2.24 | 2.24 | 97084.0 KB | 9.44 | 1511.8 KB | 1.10 | 123.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables | EPPlus | 343.40 ms | 9.90 ms | 5.71 ms | 9.34 | 9.34 | 110815.6 KB | 10.77 | 1100.6 KB | 0.80 | 834.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables | ClosedXML | 381.65 ms | 13.24 ms | 7.65 ms | 10.38 | 10.38 | 172003.7 KB | 16.72 | 1139.0 KB | 0.83 | 938.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | OfficeIMO.Excel | 41.77 ms | 2.78 ms | 1.61 ms | 1.00 | 1.00 | 10300.7 KB | 1.00 | 1376.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | MiniExcel | 94.28 ms | 6.11 ms | 3.53 ms | 2.26 | 2.26 | 128874.9 KB | 12.51 | 1512.0 KB | 1.10 | 125.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | EPPlus | 439.32 ms | 18.49 ms | 10.67 ms | 10.52 | 10.52 | 195407.9 KB | 18.97 | 1100.9 KB | 0.80 | 951.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | ClosedXML | 835.89 ms | 76.18 ms | 43.98 ms | 20.01 | 20.01 | 550095.1 KB | 53.40 | 1139.3 KB | 0.83 | 1901.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-direct | LargeXlsx | 32.82 ms | 2.37 ms | 1.37 ms | 0.93 | 1.00 | 9520.4 KB | 0.75 | 1386.5 KB | 1.00 | 7.5% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-direct | OfficeIMO.Excel | 35.46 ms | 1.00 ms | 0.58 ms | 1.00 | 1.08 | 12715.7 KB | 1.00 | 1384.9 KB | 1.00 | Loss +8.1% |
| 25000 | package-profile | package | Package size | write-datatable-direct | MiniExcel | 84.58 ms | 5.70 ms | 3.29 ms | 2.38 | 2.58 | 92394.2 KB | 7.27 | 1521.1 KB | 1.10 | 138.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-direct | ClosedXML | 307.30 ms | 31.38 ms | 18.12 ms | 8.67 | 9.36 | 104205.0 KB | 8.19 | 1139.9 KB | 0.82 | 766.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-direct | EPPlus | 377.26 ms | 16.65 ms | 9.62 ms | 10.64 | 11.50 | 117437.3 KB | 9.24 | 1090.8 KB | 0.79 | 963.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | OfficeIMO.Excel | 39.72 ms | 7.19 ms | 4.15 ms | 1.00 | 1.00 | 12733.8 KB | 1.00 | 1385.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | MiniExcel | 83.95 ms | 7.30 ms | 4.22 ms | 2.11 | 2.11 | 92394.5 KB | 7.26 | 1521.1 KB | 1.10 | 111.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | EPPlus | 357.03 ms | 22.67 ms | 13.09 ms | 8.99 | 8.99 | 117437.3 KB | 9.22 | 1090.8 KB | 0.79 | 798.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | ClosedXML | 423.24 ms | 43.86 ms | 25.32 ms | 10.65 | 10.65 | 173402.7 KB | 13.62 | 1140.7 KB | 0.82 | 965.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | LargeXlsx | 30.91 ms | 6.85 ms | 3.96 ms | 0.94 | 1.00 | 5614.1 KB | 0.43 | 1386.5 KB | 1.00 | 6.4% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 33.04 ms | 1.32 ms | 0.76 ms | 1.00 | 1.07 | 12912.0 KB | 1.00 | 1384.9 KB | 1.00 | Loss +6.9% |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | MiniExcel | 70.99 ms | 10.13 ms | 5.85 ms | 2.15 | 2.30 | 93257.0 KB | 7.22 | 1521.0 KB | 1.10 | 114.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | ClosedXML | 281.13 ms | 4.49 ms | 2.59 ms | 8.51 | 9.10 | 104205.0 KB | 8.07 | 1139.9 KB | 0.82 | 751.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | EPPlus | 372.84 ms | 9.21 ms | 5.31 ms | 11.29 | 12.06 | 117437.7 KB | 9.10 | 1090.8 KB | 0.79 | 1028.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 35.17 ms | 0.05 ms | 0.03 ms | 1.00 | 1.00 | 11501.6 KB | 1.00 | 1428.5 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 375.88 ms | 24.15 ms | 13.94 ms | 10.69 | 10.69 | 159742.2 KB | 13.89 | 1091.0 KB | 0.76 | 968.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 677.49 ms | 31.16 ms | 17.99 ms | 19.26 | 19.26 | 496956.9 KB | 43.21 | 1140.1 KB | 0.80 | 1826.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | LargeXlsx | 26.24 ms | 0.35 ms | 0.20 ms | 0.84 | 1.00 | 5614.1 KB | 0.49 | 1386.5 KB | 0.97 | 15.6% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | OfficeIMO.Excel | 31.08 ms | 1.01 ms | 0.58 ms | 1.00 | 1.18 | 11493.8 KB | 1.00 | 1428.4 KB | 1.00 | Loss +18.5% |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | MiniExcel | 62.38 ms | 0.52 ms | 0.30 ms | 2.01 | 2.38 | 93257.0 KB | 8.11 | 1521.0 KB | 1.06 | 100.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | ClosedXML | 259.10 ms | 1.28 ms | 0.74 ms | 8.34 | 9.87 | 104205.0 KB | 9.07 | 1139.9 KB | 0.80 | 733.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | EPPlus | 315.46 ms | 6.15 ms | 3.55 ms | 10.15 | 12.02 | 117437.3 KB | 10.22 | 1090.8 KB | 0.76 | 914.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 40.26 ms | 1.25 ms | 0.72 ms | 1.00 | 1.00 | 10187.2 KB | 1.00 | 1385.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 361.93 ms | 4.48 ms | 2.59 ms | 8.99 | 8.99 | 159742.2 KB | 15.68 | 1091.0 KB | 0.79 | 799.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 657.21 ms | 10.08 ms | 5.82 ms | 16.33 | 16.33 | 496956.9 KB | 48.78 | 1140.1 KB | 0.82 | 1532.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 27.38 ms | 1.70 ms | 0.98 ms | 0.73 | 1.00 | 5614.1 KB | 0.55 | 1386.5 KB | 1.00 | 27.0% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 37.52 ms | 1.21 ms | 0.70 ms | 1.00 | 1.37 | 10179.4 KB | 1.00 | 1384.9 KB | 1.00 | Loss +37.0% |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | MiniExcel | 63.99 ms | 1.86 ms | 1.07 ms | 1.71 | 2.34 | 93257.0 KB | 9.16 | 1521.1 KB | 1.10 | 70.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | ClosedXML | 283.20 ms | 3.95 ms | 2.28 ms | 7.55 | 10.34 | 104205.0 KB | 10.24 | 1139.9 KB | 0.82 | 654.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | EPPlus | 330.53 ms | 5.86 ms | 3.39 ms | 8.81 | 12.07 | 117437.3 KB | 11.54 | 1090.8 KB | 0.79 | 780.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | LargeXlsx | 25.93 ms | 0.18 ms | 0.10 ms | 0.66 | 1.00 | 5614.1 KB | 0.36 | 1386.5 KB | 0.97 | 34.3% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | OfficeIMO.Excel | 39.44 ms | 1.76 ms | 1.02 ms | 1.00 | 1.52 | 15791.7 KB | 1.00 | 1428.4 KB | 1.00 | Loss +52.1% |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | MiniExcel | 65.10 ms | 3.50 ms | 2.02 ms | 1.65 | 2.51 | 93257.0 KB | 5.91 | 1521.1 KB | 1.06 | 65.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | ClosedXML | 267.03 ms | 14.45 ms | 8.34 ms | 6.77 | 10.30 | 104205.0 KB | 6.60 | 1139.9 KB | 0.80 | 577.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | EPPlus | 315.02 ms | 2.05 ms | 1.18 ms | 7.99 | 12.15 | 117437.3 KB | 7.44 | 1090.8 KB | 0.76 | 698.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 32.79 ms | 0.28 ms | 0.16 ms | 1.00 | 1.00 | 11500.9 KB | 1.00 | 1428.5 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 360.78 ms | 13.74 ms | 7.93 ms | 11.00 | 11.00 | 138360.4 KB | 12.03 | 1091.0 KB | 0.76 | 1000.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 431.01 ms | 5.79 ms | 3.34 ms | 13.14 | 13.14 | 275422.3 KB | 23.95 | 1140.1 KB | 0.80 | 1214.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | LargeXlsx | 36.26 ms | 0.54 ms | 0.31 ms | 0.86 | 1.00 | 6043.9 KB | 0.57 | 1816.3 KB | 0.99 | 13.8% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 42.04 ms | 0.35 ms | 0.20 ms | 1.00 | 1.16 | 10577.2 KB | 1.00 | 1828.0 KB | 1.00 | Loss +15.9% |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | MiniExcel | 78.74 ms | 1.39 ms | 0.80 ms | 1.87 | 2.17 | 113974.3 KB | 10.78 | 1936.7 KB | 1.06 | 87.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | ClosedXML | 359.28 ms | 10.43 ms | 6.02 ms | 8.55 | 9.91 | 179552.5 KB | 16.98 | 1555.2 KB | 0.85 | 754.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | EPPlus | 436.99 ms | 11.79 ms | 6.80 ms | 10.39 | 12.05 | 144920.0 KB | 13.70 | 1473.0 KB | 0.81 | 939.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | LargeXlsx | 37.54 ms | 0.93 ms | 0.54 ms | 0.87 | 1.00 | 6043.9 KB | 0.61 | 1816.3 KB | 0.99 | 13.1% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 43.20 ms | 0.78 ms | 0.45 ms | 1.00 | 1.15 | 9942.2 KB | 1.00 | 1828.0 KB | 1.00 | Loss +15.1% |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | MiniExcel | 81.33 ms | 2.63 ms | 1.52 ms | 1.88 | 2.17 | 113974.3 KB | 11.46 | 1936.7 KB | 1.06 | 88.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | ClosedXML | 370.89 ms | 2.87 ms | 1.66 ms | 8.59 | 9.88 | 179553.1 KB | 18.06 | 1555.2 KB | 0.85 | 758.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | EPPlus | 453.83 ms | 11.07 ms | 6.39 ms | 10.51 | 12.09 | 144920.0 KB | 14.58 | 1473.0 KB | 0.81 | 950.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 221.70 ms | 28.50 ms | 16.45 ms | 1.00 | 1.00 | 36150.1 KB | 1.00 | 6725.6 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | LargeXlsx | 252.35 ms | 44.17 ms | 25.50 ms | 1.14 | 1.14 | 23211.4 KB | 0.64 | 6614.8 KB | 0.98 | 13.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | MiniExcel | 401.93 ms | 95.06 ms | 54.88 ms | 1.81 | 1.81 | 347925.7 KB | 9.62 | 6949.8 KB | 1.03 | 81.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | ClosedXML | 1458.48 ms | 154.25 ms | 89.06 ms | 6.58 | 6.58 | 487446.6 KB | 13.48 | 6165.9 KB | 0.92 | 557.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | EPPlus | 2098.58 ms | 614.64 ms | 354.86 ms | 9.47 | 9.47 | 562916.0 KB | 15.57 | 5441.6 KB | 0.81 | 846.6% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | OfficeIMO.Excel | 84.43 ms | 5.35 ms | 3.09 ms | 1.00 | 1.00 | 15708.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | EPPlus | 465.82 ms | 8.74 ms | 5.05 ms | 5.52 | 5.52 | 250950.0 KB | 15.98 |  |  | 451.7% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | EPPlus 4.5.3.3 | 594.23 ms |  |  | 7.04 | 7.04 |  |  |  |  | 603.8% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | ClosedXML | 1389.74 ms | 15.78 ms | 9.11 ms | 16.46 | 16.46 | 829721.7 KB | 52.82 |  |  | 1546.1% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus 4.5.3.3 | 264.87 ms |  |  | 0.87 | 1.00 |  |  |  |  | 13.4% faster than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | OfficeIMO.Excel | 305.85 ms | 31.42 ms | 18.14 ms | 1.00 | 1.15 | 133432.9 KB | 1.00 |  |  | Loss +15.5% |
| 25000 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus | 589.21 ms | 260.63 ms | 150.47 ms | 1.93 | 2.22 | 234782.5 KB | 1.76 |  |  | 92.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | OfficeIMO.Excel | 99.49 ms | 6.99 ms | 4.03 ms | 1.00 | 1.00 | 43560.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus 4.5.3.3 | 461.26 ms |  |  | 4.64 | 4.64 |  |  |  |  | 363.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus | 510.86 ms | 6.45 ms | 3.72 ms | 5.13 | 5.13 | 277076.3 KB | 6.36 |  |  | 413.5% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | OfficeIMO.Excel | 370.85 ms | 37.06 ms | 21.40 ms | 1.00 | 1.00 | 144825.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus 4.5.3.3 | 505.82 ms |  |  | 1.36 | 1.36 |  |  |  |  | 36.4% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus | 519.47 ms | 32.28 ms | 18.64 ms | 1.40 | 1.40 | 302759.9 KB | 2.09 |  |  | 40.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | OfficeIMO.Excel | 321.76 ms | 18.20 ms | 10.51 ms | 1.00 | 1.00 | 133464.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus 4.5.3.3 | 474.57 ms |  |  | 1.47 | 1.47 |  |  |  |  | 47.5% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus | 494.52 ms | 8.72 ms | 5.03 ms | 1.54 | 1.54 | 277077.5 KB | 2.08 |  |  | 53.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | OfficeIMO.Excel | 304.37 ms | 23.81 ms | 13.75 ms | 1.00 | 1.00 | 133505.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus | 474.39 ms | 26.72 ms | 15.42 ms | 1.56 | 1.56 | 277070.2 KB | 2.08 |  |  | 55.9% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus 4.5.3.3 | 517.28 ms |  |  | 1.70 | 1.70 |  |  |  |  | 70.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Object projection | build-object-datatable-dictionaries | OfficeIMO.Excel | 9.37 ms | 0.19 ms | 0.11 ms | 1.00 | 1.00 | 5164.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Object projection | build-object-datatable-typed | OfficeIMO.Excel | 7.64 ms | 0.53 ms | 0.31 ms | 1.00 | 1.00 | 8093.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-cells | OfficeIMO.Excel | 56.17 ms | 4.42 ms | 2.55 ms | 1.00 | 1.00 | 24530.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-cells | EPPlus | 339.28 ms | 18.64 ms | 10.76 ms | 6.04 | 6.04 | 187393.2 KB | 7.64 |  |  | 504.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-cells | ClosedXML | 398.90 ms | 47.81 ms | 27.60 ms | 7.10 | 7.10 | 166520.7 KB | 6.79 |  |  | 610.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | OfficeIMO.Excel | 42.80 ms | 4.18 ms | 2.42 ms | 1.00 | 1.00 | 3839.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | EPPlus | 310.59 ms | 46.57 ms | 26.89 ms | 7.26 | 7.26 | 115541.6 KB | 30.10 |  |  | 625.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | ClosedXML | 366.22 ms | 30.13 ms | 17.39 ms | 8.56 | 8.56 | 150900.8 KB | 39.31 |  |  | 755.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-range | OfficeIMO.Excel | 56.08 ms | 11.37 ms | 6.57 ms | 1.00 | 1.00 | 24530.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-range | EPPlus | 324.06 ms | 15.50 ms | 8.95 ms | 5.78 | 5.78 | 187393.2 KB | 7.64 |  |  | 477.9% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-range | ClosedXML | 484.17 ms | 49.59 ms | 28.63 ms | 8.63 | 8.63 | 166525.3 KB | 6.79 |  |  | 763.4% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-top-range | OfficeIMO.Excel | 0.81 ms | 0.02 ms | 0.01 ms | 1.00 | 1.00 | 285.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-top-range | EPPlus | 354.23 ms | 84.12 ms | 48.57 ms | 439.54 | 439.54 | 105580.1 KB | 370.21 |  |  | 43854.4% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-top-range | ClosedXML | 420.23 ms | 15.74 ms | 9.09 ms | 521.45 | 521.45 | 149402.2 KB | 523.87 |  |  | 52044.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-autofilter | OfficeIMO.Excel | 32.00 ms | 1.06 ms | 0.61 ms | 1.00 | 1.00 | 11494.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-autofilter | EPPlus 4.5.3.3 | 266.67 ms |  |  | 8.33 | 8.33 |  |  |  |  | 733.5% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-autofilter | ClosedXML | 294.02 ms | 15.95 ms | 9.21 ms | 9.19 | 9.19 | 210663.8 KB | 18.33 |  |  | 818.9% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-autofilter | EPPlus | 348.45 ms | 12.69 ms | 7.33 ms | 10.89 | 10.89 | 211871.5 KB | 18.43 |  |  | 989.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-charts | OfficeIMO.Excel | 34.42 ms | 0.93 ms | 0.54 ms | 1.00 | 1.00 | 12551.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-charts | EPPlus 4.5.3.3 | 329.28 ms |  |  | 9.57 | 9.57 |  |  |  |  | 856.8% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-charts | EPPlus | 356.03 ms | 8.75 ms | 5.05 ms | 10.34 | 10.34 | 214905.3 KB | 17.12 |  |  | 934.5% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | OfficeIMO.Excel | 36.56 ms | 4.72 ms | 2.73 ms | 1.00 | 1.00 | 11560.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | EPPlus 4.5.3.3 | 289.19 ms |  |  | 7.91 | 7.91 |  |  |  |  | 691.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | ClosedXML | 324.45 ms | 39.21 ms | 22.64 ms | 8.88 | 8.88 | 210711.7 KB | 18.23 |  |  | 787.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | EPPlus | 396.52 ms | 59.47 ms | 34.33 ms | 10.85 | 10.85 | 211912.9 KB | 18.33 |  |  | 984.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-data-validation | OfficeIMO.Excel | 34.22 ms | 3.25 ms | 1.88 ms | 1.00 | 1.00 | 11510.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-data-validation | ClosedXML | 301.17 ms | 7.43 ms | 4.29 ms | 8.80 | 8.80 | 210672.7 KB | 18.30 |  |  | 780.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-data-validation | EPPlus 4.5.3.3 | 331.44 ms |  |  | 9.68 | 9.68 |  |  |  |  | 868.4% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-data-validation | EPPlus | 362.22 ms | 11.16 ms | 6.44 ms | 10.58 | 10.58 | 211857.4 KB | 18.41 |  |  | 958.4% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | OfficeIMO.Excel | 31.21 ms | 0.63 ms | 0.36 ms | 1.00 | 1.00 | 11497.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | EPPlus 4.5.3.3 | 241.19 ms |  |  | 7.73 | 7.73 |  |  |  |  | 672.9% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | ClosedXML | 301.69 ms | 9.01 ms | 5.20 ms | 9.67 | 9.67 | 210646.8 KB | 18.32 |  |  | 866.8% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | EPPlus | 386.96 ms | 8.90 ms | 5.14 ms | 12.40 | 12.40 | 211883.3 KB | 18.43 |  |  | 1140.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | OfficeIMO.Excel | 185.03 ms | 13.05 ms | 7.53 ms | 1.00 | 1.00 | 131923.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | EPPlus 4.5.3.3 | 282.74 ms |  |  | 1.53 | 1.53 |  |  |  |  | 52.8% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | EPPlus | 374.84 ms | 4.56 ms | 2.63 ms | 2.03 | 2.03 | 230798.6 KB | 1.75 |  |  | 102.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | OfficeIMO.Excel | 358.93 ms | 83.65 ms | 48.29 ms | 1.00 | 1.00 | 133443.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus 4.5.3.3 | 459.78 ms |  |  | 1.28 | 1.28 |  |  |  |  | 28.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus | 628.05 ms | 90.29 ms | 52.13 ms | 1.75 | 1.75 | 277077.5 KB | 2.08 |  |  | 75.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world report | realworld-report-core | OfficeIMO.Excel | 35.30 ms | 0.94 ms | 0.54 ms | 1.00 | 1.00 | 11648.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world report | realworld-report-core | EPPlus | 441.33 ms | 17.22 ms | 9.94 ms | 12.50 | 12.50 | 255065.8 KB | 21.90 |  |  | 1150.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world report | realworld-report-core | EPPlus 4.5.3.3 | 478.68 ms |  |  | 13.56 | 13.56 |  |  |  |  | 1255.8% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world report | realworld-report-core | ClosedXML | 886.05 ms | 52.44 ms | 30.28 ms | 25.10 | 25.10 | 680116.8 KB | 58.39 |  |  | 2409.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook | OfficeIMO.Excel | 339.08 ms | 11.54 ms | 6.66 ms | 1.00 | 1.00 | 175151.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Report workbook | report-workbook | EPPlus | 578.04 ms | 16.97 ms | 9.80 ms | 1.70 | 1.70 | 364709.1 KB | 2.08 |  |  | 70.5% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook | EPPlus 4.5.3.3 | 1097.75 ms |  |  | 3.24 | 3.24 |  |  |  |  | 223.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-core | OfficeIMO.Excel | 49.50 ms | 2.84 ms | 1.64 ms | 1.00 | 1.00 | 10979.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Report workbook | report-workbook-core | EPPlus | 535.90 ms | 14.95 ms | 8.63 ms | 10.83 | 10.83 | 342842.2 KB | 31.23 |  |  | 982.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-core | EPPlus 4.5.3.3 | 788.24 ms |  |  | 15.93 | 15.93 |  |  |  |  | 1492.5% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-core | ClosedXML | 1107.39 ms | 60.82 ms | 35.11 ms | 22.37 | 22.37 | 975775.1 KB | 88.87 |  |  | 2137.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable | OfficeIMO.Excel | 343.99 ms | 5.99 ms | 3.46 ms | 1.00 | 1.00 | 177939.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable | EPPlus | 547.81 ms | 19.44 ms | 11.23 ms | 1.59 | 1.59 | 247823.0 KB | 1.39 |  |  | 59.2% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable | EPPlus 4.5.3.3 | 694.45 ms |  |  | 2.02 | 2.02 |  |  |  |  | 101.9% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable-core | OfficeIMO.Excel | 59.82 ms | 10.83 ms | 6.25 ms | 1.00 | 1.00 | 13725.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable-core | EPPlus | 569.30 ms | 87.95 ms | 50.78 ms | 9.52 | 9.52 | 225957.1 KB | 16.46 |  |  | 851.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable-core | EPPlus 4.5.3.3 | 668.83 ms |  |  | 11.18 | 11.18 |  |  |  |  | 1018.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable-core | ClosedXML | 1224.35 ms | 353.58 ms | 204.14 ms | 20.47 | 20.47 | 832227.0 KB | 60.64 |  |  | 1946.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | OfficeIMO.Excel | 26.43 ms | 9.18 ms | 5.30 ms | 1.00 | 1.00 | 6216.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | EPPlus 4.5.3.3 | 77.38 ms |  |  | 2.93 | 2.93 |  |  |  |  | 192.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | EPPlus | 218.21 ms | 100.71 ms | 58.15 ms | 8.26 | 8.26 | 70814.5 KB | 11.39 |  |  | 725.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | ClosedXML | 284.10 ms | 34.57 ms | 19.96 ms | 10.75 | 10.75 | 79515.5 KB | 12.79 |  |  | 975.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | OfficeIMO.Excel | 1.04 ms | 0.34 ms | 0.20 ms | 1.00 | 1.00 | 179.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | Sylvan.Data.Excel | 1.44 ms | 0.32 ms | 0.19 ms | 1.39 | 1.39 | 316.6 KB | 1.76 |  |  | 38.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | ExcelDataReader | 2.05 ms | 0.66 ms | 0.38 ms | 1.97 | 1.97 | 4062.2 KB | 22.59 |  |  | 96.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | ClosedXML | 3.55 ms | 0.13 ms | 0.08 ms | 3.40 | 3.40 | 4392.7 KB | 24.43 |  |  | 240.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | MiniExcel | 14.46 ms | 0.90 ms | 0.52 ms | 13.89 | 13.89 | 46194.9 KB | 256.87 |  |  | 1288.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | EPPlus 4.5.3.3 | 17.64 ms |  |  | 16.93 | 16.93 |  |  |  |  | 1593.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | EPPlus | 95.10 ms | 0.33 ms | 0.19 ms | 91.31 | 91.31 | 43071.0 KB | 239.50 |  |  | 9030.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | OfficeIMO.Excel | 0.84 ms | 0.01 ms | 0.01 ms | 1.00 | 1.00 | 177.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | Sylvan.Data.Excel | 1.02 ms | 0.04 ms | 0.02 ms | 1.22 | 1.22 | 316.6 KB | 1.79 |  |  | 22.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | ExcelDataReader | 1.72 ms | 0.34 ms | 0.20 ms | 2.05 | 2.05 | 4062.2 KB | 22.92 |  |  | 104.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | ClosedXML | 3.61 ms | 0.10 ms | 0.06 ms | 4.31 | 4.31 | 4392.7 KB | 24.78 |  |  | 330.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | MiniExcel | 15.64 ms | 3.47 ms | 2.00 ms | 18.67 | 18.67 | 46194.9 KB | 260.64 |  |  | 1767.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | EPPlus 4.5.3.3 | 16.88 ms |  |  | 20.15 | 20.15 |  |  |  |  | 1915.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | EPPlus | 100.23 ms | 4.35 ms | 2.51 ms | 119.66 | 119.66 | 43071.0 KB | 243.02 |  |  | 11866.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | Sylvan.Data.Excel | 19.81 ms | 3.05 ms | 1.76 ms | 0.77 | 1.00 | 1936.7 KB | 0.21 |  |  | 22.6% faster than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | OfficeIMO.Excel | 25.58 ms | 3.53 ms | 2.04 ms | 1.00 | 1.29 | 9217.9 KB | 1.00 |  |  | Loss +29.2% |
| 25000 | speed-comparison | read | Other | shared-string-read | ExcelDataReader | 67.07 ms | 12.33 ms | 7.12 ms | 2.62 | 3.39 | 25020.8 KB | 2.71 |  |  | 162.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | MiniExcel | 79.38 ms | 21.99 ms | 12.70 ms | 3.10 | 4.01 | 74405.3 KB | 8.07 |  |  | 210.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | EPPlus 4.5.3.3 | 93.82 ms |  |  | 3.67 | 4.74 |  |  |  |  | 266.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | EPPlus | 186.17 ms | 17.01 ms | 9.82 ms | 7.28 | 9.40 | 89346.0 KB | 9.69 |  |  | 627.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | ClosedXML | 225.16 ms | 72.81 ms | 42.04 ms | 8.80 | 11.37 | 90414.6 KB | 9.81 |  |  | 780.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | OfficeIMO.Excel | 35.65 ms | 3.25 ms | 1.87 ms | 1.00 | 1.00 | 1122.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | Sylvan.Data.Excel | 40.48 ms | 3.65 ms | 2.11 ms | 1.14 | 1.14 | 3534.8 KB | 3.15 |  |  | 13.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | ExcelDataReader | 114.77 ms | 5.90 ms | 3.41 ms | 3.22 | 3.22 | 61201.9 KB | 54.53 |  |  | 221.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | MiniExcel | 124.93 ms | 10.84 ms | 6.26 ms | 3.50 | 3.50 | 186420.9 KB | 166.11 |  |  | 250.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | EPPlus | 233.28 ms | 14.09 ms | 8.14 ms | 6.54 | 6.54 | 105609.0 KB | 94.10 |  |  | 554.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | ClosedXML | 358.20 ms | 32.21 ms | 18.59 ms | 10.05 | 10.05 | 149387.0 KB | 133.11 |  |  | 904.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | Sylvan.Data.Excel | 64.86 ms | 5.42 ms | 3.13 ms | 0.89 | 1.00 | 18394.2 KB | 0.53 |  |  | 11.3% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | OfficeIMO.Excel | 73.16 ms | 5.41 ms | 3.13 ms | 1.00 | 1.13 | 34645.7 KB | 1.00 |  |  | Loss +12.8% |
| 25000 | speed-comparison | read | Range and table read | read-datatable | ExcelDataReader | 164.24 ms | 14.75 ms | 8.52 ms | 2.24 | 2.53 | 76061.4 KB | 2.20 |  |  | 124.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | MiniExcel | 211.57 ms | 54.90 ms | 31.70 ms | 2.89 | 3.26 | 181285.0 KB | 5.23 |  |  | 189.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | EPPlus 4.5.3.3 | 231.49 ms |  |  | 3.16 | 3.57 |  |  |  |  | 216.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | EPPlus | 329.36 ms | 17.70 ms | 10.22 ms | 4.50 | 5.08 | 202250.2 KB | 5.84 |  |  | 350.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | ClosedXML | 418.26 ms | 59.03 ms | 34.08 ms | 5.72 | 6.45 | 178450.3 KB | 5.15 |  |  | 471.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | OfficeIMO.Excel | 41.21 ms | 7.32 ms | 4.23 ms | 1.00 | 1.00 | 4034.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | Sylvan.Data.Excel | 55.52 ms | 11.68 ms | 6.74 ms | 1.35 | 1.35 | 4316.2 KB | 1.07 |  |  | 34.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | MiniExcel | 118.53 ms | 24.64 ms | 14.23 ms | 2.88 | 2.88 | 158612.9 KB | 39.31 |  |  | 187.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | ExcelDataReader | 126.28 ms | 13.52 ms | 7.81 ms | 3.06 | 3.06 | 61201.9 KB | 15.17 |  |  | 206.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | EPPlus | 243.80 ms | 32.05 ms | 18.51 ms | 5.92 | 5.92 | 115541.6 KB | 28.64 |  |  | 491.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | ClosedXML | 351.77 ms | 63.54 ms | 36.69 ms | 8.54 | 8.54 | 150902.7 KB | 37.40 |  |  | 753.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | Sylvan.Data.Excel | 73.75 ms | 3.72 ms | 2.15 ms | 0.95 | 1.00 | 3534.8 KB | 0.14 |  |  | 5.1% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | OfficeIMO.Excel | 77.69 ms | 15.10 ms | 8.72 ms | 1.00 | 1.05 | 26098.2 KB | 1.00 |  |  | Loss +5.4% |
| 25000 | speed-comparison | read | Range and table read | read-range | MiniExcel | 170.99 ms | 29.35 ms | 16.95 ms | 2.20 | 2.32 | 186421.5 KB | 7.14 |  |  | 120.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | EPPlus 4.5.3.3 | 180.13 ms |  |  | 2.32 | 2.44 |  |  |  |  | 131.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | ExcelDataReader | 185.62 ms | 47.17 ms | 27.23 ms | 2.39 | 2.52 | 61201.9 KB | 2.35 |  |  | 138.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | EPPlus | 481.14 ms | 169.35 ms | 97.78 ms | 6.19 | 6.52 | 187390.9 KB | 7.18 |  |  | 519.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | ClosedXML | 515.27 ms | 30.04 ms | 17.34 ms | 6.63 | 6.99 | 163591.4 KB | 6.27 |  |  | 563.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | OfficeIMO.Excel | 88.39 ms | 23.40 ms | 13.51 ms | 1.00 | 1.00 | 26684.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | Sylvan.Data.Excel | 105.24 ms | 12.05 ms | 6.96 ms | 1.19 | 1.19 | 4484.9 KB | 0.17 |  |  | 19.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | ExcelDataReader | 221.67 ms | 94.33 ms | 54.46 ms | 2.51 | 2.51 | 61201.9 KB | 2.29 |  |  | 150.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | MiniExcel | 333.49 ms | 101.61 ms | 58.67 ms | 3.77 | 3.77 | 186421.5 KB | 6.99 |  |  | 277.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | EPPlus | 522.21 ms | 181.18 ms | 104.60 ms | 5.91 | 5.91 | 187390.9 KB | 7.02 |  |  | 490.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | ClosedXML | 698.43 ms | 117.17 ms | 67.65 ms | 7.90 | 7.90 | 163585.8 KB | 6.13 |  |  | 690.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | Sylvan.Data.Excel | 0.49 ms | 0.09 ms | 0.05 ms | 0.72 | 1.00 | 348.5 KB | 1.18 |  |  | 27.9% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | OfficeIMO.Excel | 0.68 ms | 0.16 ms | 0.09 ms | 1.00 | 1.39 | 296.0 KB | 1.00 |  |  | Loss +38.7% |
| 25000 | speed-comparison | read | Range and table read | read-top-range | MiniExcel | 1.10 ms | 0.49 ms | 0.29 ms | 1.62 | 2.24 | 869.0 KB | 2.94 |  |  | 61.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | ExcelDataReader | 47.34 ms | 13.97 ms | 8.07 ms | 69.42 | 96.27 | 17115.3 KB | 57.83 |  |  | 6841.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | EPPlus 4.5.3.3 | 147.60 ms |  |  | 216.44 | 300.15 |  |  |  |  | 21544.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | EPPlus | 303.47 ms | 100.30 ms | 57.91 ms | 444.99 | 617.09 | 105577.7 KB | 356.72 |  |  | 44398.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | ClosedXML | 443.69 ms | 179.09 ms | 103.40 ms | 650.60 | 902.23 | 149390.4 KB | 504.75 |  |  | 64959.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | Sylvan.Data.Excel | 145.45 ms | 89.42 ms | 51.63 ms | 0.52 | 1.00 | 3534.8 KB | 0.10 |  |  | 48.1% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | ExcelDataReader | 217.22 ms | 44.36 ms | 25.61 ms | 0.78 | 1.49 | 61201.9 KB | 1.79 |  |  | 22.5% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | MiniExcel | 230.58 ms | 19.84 ms | 11.45 ms | 0.82 | 1.59 | 186421.5 KB | 5.46 |  |  | 17.7% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | OfficeIMO.Excel | 280.18 ms | 155.92 ms | 90.02 ms | 1.00 | 1.93 | 34151.7 KB | 1.00 |  |  | Loss +92.6% |
| 25000 | speed-comparison | read | Range and table read | read-used-range | EPPlus | 801.83 ms | 413.33 ms | 238.64 ms | 2.86 | 5.51 | 187390.9 KB | 5.49 |  |  | 186.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | ClosedXML | 860.09 ms | 295.80 ms | 170.78 ms | 3.07 | 5.91 | 163592.6 KB | 4.79 |  |  | 207.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | OfficeIMO.Excel | 32.60 ms | 0.84 ms | 0.49 ms | 1.00 | 1.00 | 1125.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | Sylvan.Data.Excel | 37.16 ms | 0.63 ms | 0.36 ms | 1.14 | 1.14 | 3534.8 KB | 3.14 |  |  | 14.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | ExcelDataReader | 105.42 ms | 4.03 ms | 2.33 ms | 3.23 | 3.23 | 61201.9 KB | 54.37 |  |  | 223.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | MiniExcel | 114.12 ms | 0.31 ms | 0.18 ms | 3.50 | 3.50 | 186420.9 KB | 165.62 |  |  | 250.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | EPPlus | 210.05 ms | 2.68 ms | 1.55 ms | 6.44 | 6.44 | 105609.0 KB | 93.83 |  |  | 544.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | ClosedXML | 306.06 ms | 6.63 ms | 3.83 ms | 9.39 | 9.39 | 149394.6 KB | 132.73 |  |  | 838.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | Sylvan.Data.Excel | 45.85 ms | 2.75 ms | 1.59 ms | 0.63 | 1.00 | 3534.8 KB | 0.13 |  |  | 36.7% faster than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | OfficeIMO.Excel | 72.39 ms | 19.18 ms | 11.08 ms | 1.00 | 1.58 | 26883.7 KB | 1.00 |  |  | Loss +57.9% |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | ExcelDataReader | 132.48 ms | 2.82 ms | 1.63 ms | 1.83 | 2.89 | 61201.9 KB | 2.28 |  |  | 83.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | MiniExcel | 136.69 ms | 19.90 ms | 11.49 ms | 1.89 | 2.98 | 186421.5 KB | 6.93 |  |  | 88.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | EPPlus 4.5.3.3 | 166.05 ms |  |  | 2.29 | 3.62 |  |  |  |  | 129.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | EPPlus | 299.44 ms | 13.29 ms | 7.67 ms | 4.14 | 6.53 | 187390.9 KB | 6.97 |  |  | 313.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | ClosedXML | 377.28 ms | 31.54 ms | 18.21 ms | 5.21 | 8.23 | 163593.7 KB | 6.09 |  |  | 421.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | Sylvan.Data.Excel | 0.42 ms | 0.02 ms | 0.01 ms | 0.77 | 1.00 | 348.5 KB | 1.16 |  |  | 23.0% faster than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | OfficeIMO.Excel | 0.54 ms | 0.02 ms | 0.01 ms | 1.00 | 1.30 | 299.3 KB | 1.00 |  |  | Loss +29.9% |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | MiniExcel | 0.79 ms | 0.10 ms | 0.06 ms | 1.46 | 1.90 | 869.0 KB | 2.90 |  |  | 46.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | ExcelDataReader | 37.83 ms | 0.17 ms | 0.10 ms | 69.91 | 90.81 | 17115.3 KB | 57.18 |  |  | 6890.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | EPPlus 4.5.3.3 | 135.07 ms |  |  | 249.62 | 324.27 |  |  |  |  | 24862.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | EPPlus | 217.49 ms | 3.06 ms | 1.77 ms | 401.94 | 522.15 | 105577.7 KB | 352.75 |  |  | 40094.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | ClosedXML | 310.88 ms | 10.21 ms | 5.90 ms | 574.53 | 746.34 | 149392.3 KB | 499.14 |  |  | 57352.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | Sylvan.Data.Excel | 0.41 ms | 0.02 ms | 0.01 ms | 0.73 | 1.00 | 348.5 KB | 1.16 |  |  | 26.9% faster than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | OfficeIMO.Excel | 0.56 ms | 0.04 ms | 0.02 ms | 1.00 | 1.37 | 300.0 KB | 1.00 |  |  | Loss +36.8% |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | MiniExcel | 0.68 ms | 0.02 ms | 0.01 ms | 1.22 | 1.67 | 869.0 KB | 2.90 |  |  | 21.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | ExcelDataReader | 39.13 ms | 1.11 ms | 0.64 ms | 69.79 | 95.47 | 17115.3 KB | 57.06 |  |  | 6878.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | EPPlus | 220.47 ms | 8.26 ms | 4.77 ms | 393.16 | 537.82 | 105577.7 KB | 351.98 |  |  | 39215.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | ClosedXML | 311.29 ms | 16.37 ms | 9.45 ms | 555.11 | 759.37 | 149388.7 KB | 498.04 |  |  | 55411.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | Sylvan.Data.Excel | 47.62 ms | 1.94 ms | 1.12 ms | 0.89 | 1.00 | 5805.0 KB | 0.25 |  |  | 10.7% faster than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | OfficeIMO.Excel | 53.33 ms | 1.53 ms | 0.88 ms | 1.00 | 1.12 | 23562.2 KB | 1.00 |  |  | Loss +12.0% |
| 25000 | speed-comparison | read | Typed object read | read-objects | ExcelDataReader | 118.13 ms | 2.51 ms | 1.45 ms | 2.22 | 2.48 | 63472.1 KB | 2.69 |  |  | 121.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | MiniExcel | 139.60 ms | 2.80 ms | 1.62 ms | 2.62 | 2.93 | 183656.5 KB | 7.79 |  |  | 161.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | EPPlus 4.5.3.3 | 169.30 ms |  |  | 3.17 | 3.56 |  |  |  |  | 217.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | EPPlus | 273.78 ms | 16.13 ms | 9.32 ms | 5.13 | 5.75 | 199608.2 KB | 8.47 |  |  | 413.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | ClosedXML | 356.08 ms | 25.34 ms | 14.63 ms | 6.68 | 7.48 | 165541.8 KB | 7.03 |  |  | 567.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | Sylvan.Data.Excel | 45.12 ms | 2.41 ms | 1.39 ms | 0.98 | 1.00 | 5292.6 KB | 0.23 |  |  | Tie vs OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | OfficeIMO.Excel | 45.81 ms | 0.60 ms | 0.34 ms | 1.00 | 1.02 | 23367.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | ExcelDataReader | 111.86 ms | 4.59 ms | 2.65 ms | 2.44 | 2.48 | 62959.7 KB | 2.69 |  |  | 144.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | MiniExcel | 122.92 ms | 3.29 ms | 1.90 ms | 2.68 | 2.72 | 183144.2 KB | 7.84 |  |  | 168.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | EPPlus 4.5.3.3 | 163.40 ms |  |  | 3.57 | 3.62 |  |  |  |  | 256.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | EPPlus | 245.97 ms | 4.53 ms | 2.62 ms | 5.37 | 5.45 | 199413.0 KB | 8.53 |  |  | 437.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | ClosedXML | 327.12 ms | 4.26 ms | 2.46 ms | 7.14 | 7.25 | 165348.3 KB | 7.08 |  |  | 614.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | OfficeIMO.Excel | 38.73 ms | 0.77 ms | 0.45 ms | 1.00 | 1.00 | 13026.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | MiniExcel | 73.10 ms | 4.32 ms | 2.49 ms | 1.89 | 1.89 | 124495.5 KB | 9.56 |  |  | 88.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | EPPlus | 408.69 ms | 17.04 ms | 9.84 ms | 10.55 | 10.55 | 159741.6 KB | 12.26 |  |  | 955.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | EPPlus 4.5.3.3 | 437.25 ms |  |  | 11.29 | 11.29 |  |  |  |  | 1029.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | ClosedXML | 839.51 ms | 13.26 ms | 7.66 ms | 21.68 | 21.68 | 566143.8 KB | 43.46 |  |  | 2067.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | OfficeIMO.Excel | 36.71 ms | 0.02 ms | 0.01 ms | 1.00 | 1.00 | 10300.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | MiniExcel | 78.03 ms | 2.39 ms | 1.38 ms | 2.13 | 2.13 | 128874.9 KB | 12.51 |  |  | 112.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | EPPlus | 386.42 ms | 7.87 ms | 4.55 ms | 10.53 | 10.53 | 195407.9 KB | 18.97 |  |  | 952.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | ClosedXML | 724.06 ms | 9.87 ms | 5.70 ms | 19.72 | 19.72 | 550095.6 KB | 53.40 |  |  | 1872.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 36.06 ms | 0.88 ms | 0.51 ms | 1.00 | 1.00 | 11501.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 365.29 ms | 9.66 ms | 5.58 ms | 10.13 | 10.13 | 159742.3 KB | 13.89 |  |  | 913.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 683.69 ms | 9.90 ms | 5.72 ms | 18.96 | 18.96 | 496956.9 KB | 43.21 |  |  | 1796.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 40.10 ms | 1.31 ms | 0.75 ms | 1.00 | 1.00 | 10187.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 362.06 ms | 4.14 ms | 2.39 ms | 9.03 | 9.03 | 159742.3 KB | 15.68 |  |  | 803.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 680.80 ms | 22.30 ms | 12.87 ms | 16.98 | 16.98 | 496956.9 KB | 48.78 |  |  | 1597.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 33.77 ms | 0.52 ms | 0.30 ms | 1.00 | 1.00 | 11500.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 362.44 ms | 15.71 ms | 9.07 ms | 10.73 | 10.73 | 138360.4 KB | 12.03 |  |  | 973.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 437.83 ms | 11.74 ms | 6.78 ms | 12.96 | 12.96 | 275422.3 KB | 23.95 |  |  | 1196.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | OfficeIMO.Excel | 11.80 ms | 0.10 ms | 0.06 ms | 1.00 | 1.00 | 6896.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | ClosedXML | 117.84 ms | 7.27 ms | 4.20 ms | 9.99 | 9.99 | 92902.1 KB | 13.47 |  |  | 898.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | EPPlus | 169.41 ms | 8.54 ms | 4.93 ms | 14.36 | 14.36 | 74492.8 KB | 10.80 |  |  | 1335.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | OfficeIMO.Excel | 14.08 ms | 0.26 ms | 0.15 ms | 1.00 | 1.00 | 5970.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | EPPlus 4.5.3.3 | 90.82 ms |  |  | 6.45 | 6.45 |  |  |  |  | 545.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | ClosedXML | 105.01 ms | 3.07 ms | 1.77 ms | 7.46 | 7.46 | 84206.7 KB | 14.10 |  |  | 646.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | EPPlus | 181.09 ms | 4.36 ms | 2.52 ms | 12.87 | 12.87 | 86377.5 KB | 14.47 |  |  | 1186.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | OfficeIMO.Excel | 17.23 ms | 0.71 ms | 0.41 ms | 1.00 | 1.00 | 8332.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | EPPlus 4.5.3.3 | 123.47 ms |  |  | 7.17 | 7.17 |  |  |  |  | 616.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | ClosedXML | 165.09 ms | 7.46 ms | 4.31 ms | 9.58 | 9.58 | 111118.7 KB | 13.33 |  |  | 858.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | EPPlus | 211.89 ms | 0.66 ms | 0.38 ms | 12.30 | 12.30 | 113245.1 KB | 13.59 |  |  | 1130.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | OfficeIMO.Excel | 19.01 ms | 0.70 ms | 0.41 ms | 1.00 | 1.00 | 7416.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | ClosedXML | 146.53 ms | 3.44 ms | 1.99 ms | 7.71 | 7.71 | 105223.9 KB | 14.19 |  |  | 671.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | EPPlus | 210.63 ms | 2.78 ms | 1.60 ms | 11.08 | 11.08 | 106316.9 KB | 14.34 |  |  | 1008.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 17.31 ms | 0.38 ms | 0.22 ms | 1.00 | 1.00 | 7416.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | ClosedXML | 148.91 ms | 4.72 ms | 2.73 ms | 8.60 | 8.60 | 105223.9 KB | 14.19 |  |  | 760.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | EPPlus | 223.17 ms | 16.18 ms | 9.34 ms | 12.89 | 12.89 | 106316.9 KB | 14.34 |  |  | 1189.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | OfficeIMO.Excel | 11.20 ms | 0.45 ms | 0.26 ms | 1.00 | 1.00 | 6144.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | EPPlus 4.5.3.3 | 92.70 ms |  |  | 8.28 | 8.28 |  |  |  |  | 728.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | ClosedXML | 99.54 ms | 2.16 ms | 1.25 ms | 8.89 | 8.89 | 82591.3 KB | 13.44 |  |  | 789.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | EPPlus | 181.22 ms | 2.09 ms | 1.21 ms | 16.19 | 16.19 | 85127.4 KB | 13.85 |  |  | 1518.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | OfficeIMO.Excel | 19.09 ms | 0.26 ms | 0.15 ms | 1.00 | 1.00 | 7482.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | EPPlus 4.5.3.3 | 103.26 ms |  |  | 5.41 | 5.41 |  |  |  |  | 440.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | ClosedXML | 152.16 ms | 7.82 ms | 4.51 ms | 7.97 | 7.97 | 89323.7 KB | 11.94 |  |  | 696.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | EPPlus | 198.66 ms | 6.17 ms | 3.56 ms | 10.40 | 10.40 | 103800.0 KB | 13.87 |  |  | 940.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | OfficeIMO.Excel | 33.84 ms | 0.96 ms | 0.55 ms | 1.00 | 1.00 | 13039.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | MiniExcel | 75.43 ms | 2.14 ms | 1.24 ms | 2.23 | 2.23 | 97088.3 KB | 7.45 |  |  | 122.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | ClosedXML | 358.59 ms | 17.87 ms | 10.32 ms | 10.60 | 10.60 | 172019.1 KB | 13.19 |  |  | 959.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | EPPlus | 421.15 ms | 14.62 ms | 8.44 ms | 12.44 | 12.44 | 111246.0 KB | 8.53 |  |  | 1144.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | OfficeIMO.Excel | 38.39 ms | 0.73 ms | 0.42 ms | 1.00 | 1.00 | 13458.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | MiniExcel | 97.36 ms | 2.95 ms | 1.70 ms | 2.54 | 2.54 | 108129.1 KB | 8.03 |  |  | 153.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | EPPlus | 473.93 ms | 13.64 ms | 7.88 ms | 12.35 | 12.35 | 135723.5 KB | 10.08 |  |  | 1134.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | ClosedXML | 511.60 ms | 16.91 ms | 9.76 ms | 13.33 | 13.33 | 280371.8 KB | 20.83 |  |  | 1232.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | OfficeIMO.Excel | 34.56 ms | 1.08 ms | 0.62 ms | 1.00 | 1.00 | 10288.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | MiniExcel | 79.96 ms | 5.91 ms | 3.41 ms | 2.31 | 2.31 | 97085.4 KB | 9.44 |  |  | 131.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | EPPlus 4.5.3.3 | 213.50 ms |  |  | 6.18 | 6.18 |  |  |  |  | 517.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | EPPlus | 330.24 ms | 4.37 ms | 2.52 ms | 9.55 | 9.55 | 110815.9 KB | 10.77 |  |  | 855.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | ClosedXML | 364.23 ms | 13.63 ms | 7.87 ms | 10.54 | 10.54 | 171999.1 KB | 16.72 |  |  | 953.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | OfficeIMO.Excel | 35.38 ms | 0.45 ms | 0.26 ms | 1.00 | 1.00 | 13020.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | MiniExcel | 69.40 ms | 2.53 ms | 1.46 ms | 1.96 | 1.96 | 92200.0 KB | 7.08 |  |  | 96.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | EPPlus 4.5.3.3 | 224.92 ms |  |  | 6.36 | 6.36 |  |  |  |  | 535.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | ClosedXML | 365.06 ms | 5.72 ms | 3.30 ms | 10.32 | 10.32 | 173398.1 KB | 13.32 |  |  | 931.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | EPPlus | 371.70 ms | 8.81 ms | 5.09 ms | 10.51 | 10.51 | 117437.3 KB | 9.02 |  |  | 950.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | LargeXlsx | 30.70 ms | 1.51 ms | 0.87 ms | 0.88 | 1.00 | 9520.4 KB | 0.75 |  |  | 11.9% faster than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | OfficeIMO.Excel | 34.85 ms | 2.62 ms | 1.51 ms | 1.00 | 1.14 | 12715.7 KB | 1.00 |  |  | Loss +13.5% |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | MiniExcel | 83.37 ms | 15.14 ms | 8.74 ms | 2.39 | 2.72 | 92394.2 KB | 7.27 |  |  | 139.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | EPPlus 4.5.3.3 | 216.60 ms |  |  | 6.21 | 7.06 |  |  |  |  | 521.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | ClosedXML | 275.12 ms | 14.24 ms | 8.22 ms | 7.89 | 8.96 | 104205.0 KB | 8.19 |  |  | 689.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | EPPlus | 372.68 ms | 22.23 ms | 12.83 ms | 10.69 | 12.14 | 117437.3 KB | 9.24 |  |  | 969.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | OfficeIMO.Excel | 34.13 ms | 0.07 ms | 0.04 ms | 1.00 | 1.00 | 9999.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | MiniExcel | 75.33 ms | 1.03 ms | 0.59 ms | 2.21 | 2.21 | 89659.2 KB | 8.97 |  |  | 120.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | EPPlus | 315.89 ms | 14.00 ms | 8.09 ms | 9.25 | 9.25 | 114703.1 KB | 11.47 |  |  | 825.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | ClosedXML | 361.05 ms | 13.24 ms | 7.64 ms | 10.58 | 10.58 | 170666.2 KB | 17.07 |  |  | 957.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | OfficeIMO.Excel | 32.76 ms | 0.34 ms | 0.19 ms | 1.00 | 1.00 | 12733.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | MiniExcel | 74.70 ms | 1.61 ms | 0.93 ms | 2.28 | 2.28 | 92394.5 KB | 7.26 |  |  | 128.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | EPPlus 4.5.3.3 | 239.78 ms |  |  | 7.32 | 7.32 |  |  |  |  | 632.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | ClosedXML | 360.22 ms | 21.04 ms | 12.15 ms | 11.00 | 11.00 | 173395.0 KB | 13.62 |  |  | 999.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | EPPlus | 360.86 ms | 11.38 ms | 6.57 ms | 11.02 | 11.02 | 117437.3 KB | 9.22 |  |  | 1001.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | OfficeIMO.Excel | 35.64 ms | 1.20 ms | 0.69 ms | 1.00 | 1.00 | 11561.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | MiniExcel | 68.60 ms | 4.56 ms | 2.63 ms | 1.93 | 1.93 | 125551.4 KB | 10.86 |  |  | 92.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | EPPlus | 392.93 ms | 2.34 ms | 1.35 ms | 11.03 | 11.03 | 254959.0 KB | 22.05 |  |  | 1002.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | EPPlus 4.5.3.3 | 536.36 ms |  |  | 15.05 | 15.05 |  |  |  |  | 1405.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | ClosedXML | 753.29 ms | 27.33 ms | 15.78 ms | 21.14 | 21.14 | 565950.2 KB | 48.95 |  |  | 2013.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | OfficeIMO.Excel | 19.84 ms | 0.23 ms | 0.13 ms | 1.00 | 1.00 | 10112.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | EPPlus 4.5.3.3 | 159.35 ms |  |  | 8.03 | 8.03 |  |  |  |  | 703.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | ClosedXML | 178.59 ms | 3.72 ms | 2.15 ms | 9.00 | 9.00 | 113853.5 KB | 11.26 |  |  | 800.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | EPPlus | 305.37 ms | 2.13 ms | 1.23 ms | 15.39 | 15.39 | 140731.9 KB | 13.92 |  |  | 1439.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-dictionary-objects-table-direct | OfficeIMO.Excel | 45.43 ms | 1.19 ms | 0.69 ms | 1.00 | 1.00 | 15163.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | LargeXlsx | 41.10 ms | 2.96 ms | 1.71 ms | 0.89 | 1.00 | 6043.9 KB | 0.57 |  |  | 10.8% faster than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 46.09 ms | 0.48 ms | 0.28 ms | 1.00 | 1.12 | 10577.2 KB | 1.00 |  |  | Loss +12.1% |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | MiniExcel | 83.52 ms | 1.03 ms | 0.59 ms | 1.81 | 2.03 | 113974.3 KB | 10.78 |  |  | 81.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | ClosedXML | 390.53 ms | 7.18 ms | 4.15 ms | 8.47 | 9.50 | 179552.5 KB | 16.98 |  |  | 747.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | EPPlus | 470.46 ms | 22.99 ms | 13.27 ms | 10.21 | 11.45 | 144920.0 KB | 13.70 |  |  | 920.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | LargeXlsx | 40.56 ms | 3.39 ms | 1.96 ms | 0.88 | 1.00 | 6043.9 KB | 0.61 |  |  | 11.6% faster than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 45.90 ms | 2.05 ms | 1.18 ms | 1.00 | 1.13 | 9942.2 KB | 1.00 |  |  | Loss +13.2% |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | MiniExcel | 88.65 ms | 6.50 ms | 3.75 ms | 1.93 | 2.19 | 113974.3 KB | 11.46 |  |  | 93.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | ClosedXML | 383.93 ms | 8.21 ms | 4.74 ms | 8.36 | 9.47 | 179552.5 KB | 18.06 |  |  | 736.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | EPPlus | 472.44 ms | 21.78 ms | 12.58 ms | 10.29 | 11.65 | 144920.0 KB | 14.58 |  |  | 929.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 253.66 ms | 90.08 ms | 52.01 ms | 1.00 | 1.00 | 36150.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | LargeXlsx | 263.61 ms | 58.59 ms | 33.82 ms | 1.04 | 1.04 | 23211.4 KB | 0.64 |  |  | 3.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | MiniExcel | 446.74 ms | 96.03 ms | 55.44 ms | 1.76 | 1.76 | 347925.7 KB | 9.62 |  |  | 76.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | ClosedXML | 1340.18 ms | 85.93 ms | 49.61 ms | 5.28 | 5.28 | 487446.6 KB | 13.48 |  |  | 428.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | EPPlus | 2031.33 ms | 560.56 ms | 323.64 ms | 8.01 | 8.01 | 562937.4 KB | 15.57 |  |  | 700.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | LargeXlsx | 11.00 ms | 0.78 ms | 0.45 ms | 0.78 | 1.00 | 2771.0 KB | 0.26 |  |  | 22.3% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | OfficeIMO.Excel | 14.17 ms | 0.39 ms | 0.22 ms | 1.00 | 1.29 | 10842.5 KB | 1.00 |  |  | Loss +28.8% |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | MiniExcel | 29.69 ms | 0.38 ms | 0.22 ms | 2.10 | 2.70 | 58242.9 KB | 5.37 |  |  | 109.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | EPPlus 4.5.3.3 | 107.14 ms |  |  | 7.56 | 9.74 |  |  |  |  | 656.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | ClosedXML | 127.56 ms | 2.61 ms | 1.51 ms | 9.00 | 11.59 | 104233.1 KB | 9.61 |  |  | 800.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | EPPlus | 191.33 ms | 4.92 ms | 2.84 ms | 13.50 | 17.39 | 100373.5 KB | 9.26 |  |  | 1250.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 14.30 ms | 4.65 ms | 2.68 ms | 1.00 | 1.00 | 6961.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 14.37 ms | 2.57 ms | 1.49 ms | 1.00 | 1.00 | 3444.4 KB | 0.49 |  |  | Tie vs OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | ClosedXML | 215.52 ms | 92.57 ms | 53.45 ms | 15.07 | 15.07 | 96015.7 KB | 13.79 |  |  | 1407.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | EPPlus | 238.62 ms | 20.12 ms | 11.62 ms | 16.69 | 16.69 | 87467.1 KB | 12.56 |  |  | 1568.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | LargeXlsx | 30.17 ms | 1.80 ms | 1.04 ms | 0.87 | 1.00 | 5614.1 KB | 0.35 |  |  | 13.3% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 34.79 ms | 1.43 ms | 0.83 ms | 1.00 | 1.15 | 16036.5 KB | 1.00 |  |  | Loss +15.3% |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | MiniExcel | 64.70 ms | 1.79 ms | 1.03 ms | 1.86 | 2.14 | 93257.0 KB | 5.82 |  |  | 85.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 239.21 ms |  |  | 6.88 | 7.93 |  |  |  |  | 587.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | ClosedXML | 303.26 ms | 2.80 ms | 1.62 ms | 8.72 | 10.05 | 210646.1 KB | 13.14 |  |  | 771.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | EPPlus | 369.82 ms | 4.76 ms | 2.75 ms | 10.63 | 12.26 | 211849.9 KB | 13.21 |  |  | 962.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | OfficeIMO.Excel | 21.76 ms | 2.78 ms | 1.61 ms | 1.00 | 1.00 | 7866.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | ClosedXML | 192.13 ms | 37.45 ms | 21.62 ms | 8.83 | 8.83 | 105223.9 KB | 13.38 |  |  | 783.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | EPPlus | 268.93 ms | 61.29 ms | 35.39 ms | 12.36 | 12.36 | 106316.9 KB | 13.52 |  |  | 1136.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | Sylvan.Data.Excel | 34.88 ms | 7.50 ms | 4.33 ms | 0.79 | 1.00 | 5700.3 KB | 0.44 |  |  | 20.9% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | LargeXlsx | 43.45 ms | 7.97 ms | 4.60 ms | 0.99 | 1.25 | 8349.2 KB | 0.64 |  |  | Tie vs OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | OfficeIMO.Excel | 44.10 ms | 7.72 ms | 4.46 ms | 1.00 | 1.26 | 13002.3 KB | 1.00 |  |  | Loss +26.5% |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | MiniExcel | 85.60 ms | 15.16 ms | 8.75 ms | 1.94 | 2.45 | 92199.8 KB | 7.09 |  |  | 94.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | EPPlus 4.5.3.3 | 223.51 ms |  |  | 5.07 | 6.41 |  |  |  |  | 406.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | ClosedXML | 334.79 ms | 56.03 ms | 32.35 ms | 7.59 | 9.60 | 104205.0 KB | 8.01 |  |  | 659.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | EPPlus | 389.85 ms | 11.86 ms | 6.85 ms | 8.84 | 11.18 | 117437.7 KB | 9.03 |  |  | 784.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | OfficeIMO.Excel | 41.75 ms | 7.83 ms | 4.52 ms | 1.00 | 1.00 | 9819.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | LargeXlsx | 57.57 ms | 1.09 ms | 0.63 ms | 1.38 | 1.38 | 9265.9 KB | 0.94 |  |  | 37.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | MiniExcel | 182.53 ms | 55.46 ms | 32.02 ms | 4.37 | 4.37 | 108129.1 KB | 11.01 |  |  | 337.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | ClosedXML | 715.69 ms | 168.99 ms | 97.57 ms | 17.14 | 17.14 | 280371.6 KB | 28.55 |  |  | 1614.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | EPPlus | 730.53 ms | 162.22 ms | 93.66 ms | 17.50 | 17.50 | 135723.5 KB | 13.82 |  |  | 1649.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | LargeXlsx | 38.87 ms | 0.81 ms | 0.47 ms | 0.90 | 1.00 | 10795.2 KB | 0.92 |  |  | 9.8% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | OfficeIMO.Excel | 43.08 ms | 0.56 ms | 0.32 ms | 1.00 | 1.11 | 11708.2 KB | 1.00 |  |  | Loss +10.8% |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | MiniExcel | 143.92 ms | 1.82 ms | 1.05 ms | 3.34 | 3.70 | 226875.4 KB | 19.38 |  |  | 234.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | ClosedXML | 945.66 ms | 13.45 ms | 7.77 ms | 21.95 | 24.33 | 759818.4 KB | 64.90 |  |  | 2095.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | OfficeIMO.Excel | 17.77 ms | 3.85 ms | 2.23 ms | 1.00 | 1.00 | 15744.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | MiniExcel | 31.49 ms | 4.80 ms | 2.77 ms | 1.77 | 1.77 | 73760.2 KB | 4.68 |  |  | 77.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | EPPlus 4.5.3.3 | 100.12 ms |  |  | 5.63 | 5.63 |  |  |  |  | 463.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | ClosedXML | 127.26 ms | 33.39 ms | 19.28 ms | 7.16 | 7.16 | 104241.3 KB | 6.62 |  |  | 616.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | EPPlus | 172.71 ms | 5.81 ms | 3.35 ms | 9.72 | 9.72 | 84410.0 KB | 5.36 |  |  | 871.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | OfficeIMO.Excel | 17.93 ms | 2.01 ms | 1.16 ms | 1.00 | 1.00 | 15360.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | EPPlus 4.5.3.3 | 89.32 ms |  |  | 4.98 | 4.98 |  |  |  |  | 398.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | ClosedXML | 116.92 ms | 1.79 ms | 1.04 ms | 6.52 | 6.52 | 104241.3 KB | 6.79 |  |  | 551.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | EPPlus | 180.55 ms | 5.05 ms | 2.91 ms | 10.07 | 10.07 | 84410.5 KB | 5.50 |  |  | 906.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | OfficeIMO.Excel | 13.01 ms | 0.83 ms | 0.48 ms | 1.00 | 1.00 | 13824.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | ClosedXML | 150.14 ms | 5.31 ms | 3.07 ms | 11.54 | 11.54 | 131501.7 KB | 9.51 |  |  | 1053.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | EPPlus | 206.81 ms | 2.81 ms | 1.62 ms | 15.89 | 15.89 | 97729.6 KB | 7.07 |  |  | 1489.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | OfficeIMO.Excel | 11.70 ms | 0.37 ms | 0.21 ms | 1.00 | 1.00 | 7525.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | ClosedXML | 100.29 ms | 2.01 ms | 1.16 ms | 8.57 | 8.57 | 84520.0 KB | 11.23 |  |  | 757.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | EPPlus | 152.86 ms | 0.99 ms | 0.57 ms | 13.06 | 13.06 | 70033.4 KB | 9.31 |  |  | 1206.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | LargeXlsx | 30.29 ms | 4.43 ms | 2.56 ms | 0.86 | 1.00 | 5614.1 KB | 0.43 |  |  | 13.7% faster than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 35.11 ms | 1.95 ms | 1.12 ms | 1.00 | 1.16 | 12912.0 KB | 1.00 |  |  | Loss +15.9% |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | MiniExcel | 69.65 ms | 4.25 ms | 2.45 ms | 1.98 | 2.30 | 93257.0 KB | 7.22 |  |  | 98.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 248.03 ms |  |  | 7.06 | 8.19 |  |  |  |  | 606.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | ClosedXML | 281.67 ms | 6.00 ms | 3.47 ms | 8.02 | 9.30 | 104205.0 KB | 8.07 |  |  | 702.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | EPPlus | 367.65 ms | 25.95 ms | 14.98 ms | 10.47 | 12.14 | 117437.7 KB | 9.10 |  |  | 947.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | LargeXlsx | 29.04 ms | 2.23 ms | 1.29 ms | 0.89 | 1.00 | 5614.1 KB | 0.49 |  |  | 11.0% faster than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | OfficeIMO.Excel | 32.62 ms | 0.43 ms | 0.25 ms | 1.00 | 1.12 | 11493.8 KB | 1.00 |  |  | Loss +12.3% |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | MiniExcel | 61.94 ms | 0.53 ms | 0.31 ms | 1.90 | 2.13 | 93257.0 KB | 8.11 |  |  | 89.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | ClosedXML | 274.41 ms | 1.85 ms | 1.07 ms | 8.41 | 9.45 | 104205.0 KB | 9.07 |  |  | 741.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | EPPlus 4.5.3.3 | 286.11 ms |  |  | 8.77 | 9.85 |  |  |  |  | 777.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | EPPlus | 332.68 ms | 15.71 ms | 9.07 ms | 10.20 | 11.46 | 117437.3 KB | 10.22 |  |  | 920.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 28.64 ms | 1.24 ms | 0.72 ms | 0.74 | 1.00 | 5614.1 KB | 0.55 |  |  | 26.3% faster than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 38.86 ms | 1.31 ms | 0.76 ms | 1.00 | 1.36 | 10179.4 KB | 1.00 |  |  | Loss +35.7% |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | MiniExcel | 65.42 ms | 3.25 ms | 1.88 ms | 1.68 | 2.28 | 93257.0 KB | 9.16 |  |  | 68.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | ClosedXML | 284.76 ms | 14.74 ms | 8.51 ms | 7.33 | 9.94 | 104205.0 KB | 10.24 |  |  | 632.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | EPPlus | 339.38 ms | 10.39 ms | 6.00 ms | 8.73 | 11.85 | 117437.3 KB | 11.54 |  |  | 773.3% slower than OfficeIMO |
