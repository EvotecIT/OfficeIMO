# OfficeIMO.Excel Comparison Summary

This is the suite-level decision table. Mean, standard deviation, standard error, ratios, and allocations come from the lightweight rotated local runner; they are meant for engineering direction. Use the BenchmarkDotNet benchmark classes when a publication-grade `Error` column is required.

## At a glance

| Row count | Artifact | Workload | Category | OfficeIMO wins | OfficeIMO losses | Biggest loss |
| ---: | --- | --- | --- | ---: | ---: | --- |
| 2500 | dense-helloworld-comparison | read | Other | 0 | 2 | dense-helloworld-read-stream: Loss +52.2% vs Sylvan.Data.Excel |
| 2500 | package-profile | package | Package size | 41 | 13 | write-insertobjects-legacy-dictionaries-direct: Loss +49.7% vs LargeXlsx |
| 2500 | speed-comparison | mutate | AutoFit and mutation | 2 | 0 |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | 4 | 0 |  |
| 2500 | speed-comparison | other | Object projection | 2 | 0 |  |
| 2500 | speed-comparison | other | Range and table read | 4 | 0 |  |
| 2500 | speed-comparison | other | Real-world feature mix | 6 | 0 |  |
| 2500 | speed-comparison | other | Real-world report | 2 | 0 |  |
| 2500 | speed-comparison | other | Report workbook | 4 | 0 |  |
| 2500 | speed-comparison | read | Formula write/read | 1 | 0 |  |
| 2500 | speed-comparison | read | Other | 1 | 2 | shared-string-read: Loss +73.9% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | read | Range and table read | 4 | 3 | read-used-range: Loss +216.7% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | read | Streaming read | 1 | 3 | read-top-range-stream-small-chunks: Loss +29.2% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | read | Typed object read | 2 | 0 |  |
| 2500 | speed-comparison | write | AutoFit and mutation | 5 | 0 |  |
| 2500 | speed-comparison | write | Cell writer | 7 | 0 |  |
| 2500 | speed-comparison | write | DataSet table export | 3 | 0 |  |
| 2500 | speed-comparison | write | DataTable table export | 4 | 0 |  |
| 2500 | speed-comparison | write | Formatted report write | 1 | 0 |  |
| 2500 | speed-comparison | write | Formula write/read | 1 | 0 |  |
| 2500 | speed-comparison | write | Other | 1 | 3 | write-powershell-mixed-objects-direct: Loss +30.7% vs LargeXlsx |
| 2500 | speed-comparison | write | Plain cell export | 3 | 1 | append-plain-rows: Loss +56.7% vs LargeXlsx |
| 2500 | speed-comparison | write | Plain streaming export | 1 | 1 | write-datareader-plain: Loss +29.7% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | write | Plain string export | 1 | 0 |  |
| 2500 | speed-comparison | write | Shared string write | 4 | 0 |  |
| 2500 | speed-comparison | write | Typed object export | 0 | 3 | write-insertobjects-flat-dictionaries-direct: Loss +34.7% vs LargeXlsx |
| 25000 | dense-helloworld-comparison | read | Other | 0 | 2 | dense-helloworld-read-stream: Loss +30.5% vs Sylvan.Data.Excel |
| 25000 | package-profile | package | Package size | 43 | 11 | write-insertobjects-legacy-dictionaries-direct: Loss +52.0% vs LargeXlsx |
| 25000 | speed-comparison | mutate | AutoFit and mutation | 1 | 1 | realworld-report-no-autofit: Loss +19.2% vs EPPlus 4.5.3.3 |
| 25000 | speed-comparison | other | Anti-cheat report variants | 4 | 0 |  |
| 25000 | speed-comparison | other | Object projection | 2 | 0 |  |
| 25000 | speed-comparison | other | Range and table read | 4 | 0 |  |
| 25000 | speed-comparison | other | Real-world feature mix | 6 | 0 |  |
| 25000 | speed-comparison | other | Real-world report | 2 | 0 |  |
| 25000 | speed-comparison | other | Report workbook | 4 | 0 |  |
| 25000 | speed-comparison | read | Formula write/read | 1 | 0 |  |
| 25000 | speed-comparison | read | Other | 2 | 1 | shared-string-read: Loss +13.6% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | read | Range and table read | 3 | 4 | read-used-range: Loss +112.0% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | read | Streaming read | 2 | 2 | read-top-range-stream-small-chunks: Loss +24.6% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | read | Typed object read | 1 | 1 | read-objects: Loss +15.9% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | write | AutoFit and mutation | 5 | 0 |  |
| 25000 | speed-comparison | write | Cell writer | 7 | 0 |  |
| 25000 | speed-comparison | write | DataSet table export | 3 | 0 |  |
| 25000 | speed-comparison | write | DataTable table export | 3 | 1 | write-datatable-direct: Loss +10.7% vs LargeXlsx |
| 25000 | speed-comparison | write | Formatted report write | 1 | 0 |  |
| 25000 | speed-comparison | write | Formula write/read | 1 | 0 |  |
| 25000 | speed-comparison | write | Other | 2 | 2 | write-powershell-psobject-mixed-direct: Loss +14.8% vs LargeXlsx |
| 25000 | speed-comparison | write | Plain cell export | 1 | 3 | append-plain-rows: Loss +42.0% vs LargeXlsx |
| 25000 | speed-comparison | write | Plain streaming export | 0 | 2 | write-datareader-plain: Loss +33.9% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | write | Plain string export | 0 | 1 | write-blog-2023-20-string-columns: Loss +20.1% vs LargeXlsx |
| 25000 | speed-comparison | write | Shared string write | 4 | 0 |  |
| 25000 | speed-comparison | write | Typed object export | 0 | 3 | write-insertobjects-flat-dictionaries-direct: Loss +31.7% vs LargeXlsx |

## OfficeIMO decision table

| Row count | Artifact | Workload | Category | Scenario | OfficeIMO mean | Best | OfficeIMO vs best | Alloc | Package |
| ---: | --- | --- | --- | --- | ---: | --- | ---: | ---: | ---: |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | 5.95 ms | Sylvan.Data.Excel | Loss +32.3% | 2410.9 KB |  |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | 6.31 ms | Sylvan.Data.Excel | Loss +52.2% | 2489.3 KB |  |
| 2500 | package-profile | package | Package size | append-plain-rows | 1.89 ms | LargeXlsx | Loss +29.9% | 1576.3 KB | 63.0 KB |
| 2500 | package-profile | package | Package size | autofit-existing | 8.03 ms | OfficeIMO.Excel | Win | 1895.2 KB | 143.2 KB |
| 2500 | package-profile | package | Package size | large-shared-strings | 2.07 ms | OfficeIMO.Excel | Win | 2440.3 KB | 55.2 KB |
| 2500 | package-profile | package | Package size | realworld-autofilter | 3.85 ms | OfficeIMO.Excel | Win | 1340.4 KB | 142.4 KB |
| 2500 | package-profile | package | Package size | realworld-charts | 4.87 ms | OfficeIMO.Excel | Win | 1892.9 KB | 147.6 KB |
| 2500 | package-profile | package | Package size | realworld-conditional-formatting | 3.88 ms | OfficeIMO.Excel | Win | 1405.8 KB | 142.7 KB |
| 2500 | package-profile | package | Package size | realworld-data-validation | 3.77 ms | OfficeIMO.Excel | Win | 1356.1 KB | 142.5 KB |
| 2500 | package-profile | package | Package size | realworld-freeze-panes | 3.83 ms | OfficeIMO.Excel | Win | 1342.8 KB | 142.5 KB |
| 2500 | package-profile | package | Package size | realworld-pivot-table | 13.50 ms | OfficeIMO.Excel | Win | 14419.5 KB | 200.4 KB |
| 2500 | package-profile | package | Package size | realworld-report-all-in-one | 16.27 ms | OfficeIMO.Excel | Win | 15220.7 KB | 206.5 KB |
| 2500 | package-profile | package | Package size | realworld-report-chart-first | 10.29 ms | OfficeIMO.Excel | Win | 6196.8 KB | 206.5 KB |
| 2500 | package-profile | package | Package size | realworld-report-core | 4.45 ms | OfficeIMO.Excel | Win | 1488.5 KB | 143.9 KB |
| 2500 | package-profile | package | Package size | realworld-report-extra-column | 17.40 ms | OfficeIMO.Excel | Win | 16350.5 KB | 219.1 KB |
| 2500 | package-profile | package | Package size | realworld-report-no-autofit | 14.90 ms | OfficeIMO.Excel | Win | 15211.1 KB | 206.4 KB |
| 2500 | package-profile | package | Package size | realworld-report-post-mutation | 15.91 ms | OfficeIMO.Excel | Win | 15230.0 KB | 206.6 KB |
| 2500 | package-profile | package | Package size | realworld-report-shuffled-columns | 17.47 ms | OfficeIMO.Excel | Win | 15225.5 KB | 211.2 KB |
| 2500 | package-profile | package | Package size | report-workbook | 23.84 ms | OfficeIMO.Excel | Win | 19112.3 KB | 275.6 KB |
| 2500 | package-profile | package | Package size | report-workbook-core | 5.79 ms | OfficeIMO.Excel | Win | 2711.1 KB | 187.5 KB |
| 2500 | package-profile | package | Package size | report-workbook-datatable | 22.56 ms | OfficeIMO.Excel | Win | 19383.9 KB | 275.6 KB |
| 2500 | package-profile | package | Package size | report-workbook-datatable-core | 5.75 ms | OfficeIMO.Excel | Win | 2982.7 KB | 187.5 KB |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | 4.84 ms | LargeXlsx | Loss +22.8% | 1676.8 KB | 216.7 KB |
| 2500 | package-profile | package | Package size | write-bulk-report | 4.52 ms | OfficeIMO.Excel | Win | 1401.7 KB | 143.2 KB |
| 2500 | package-profile | package | Package size | write-cellformula | 2.31 ms | OfficeIMO.Excel | Win | 1383.3 KB | 66.6 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-empty-strings | 1.92 ms | OfficeIMO.Excel | Win | 1787.1 KB | 44.2 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-numbers | 1.85 ms | OfficeIMO.Excel | Win | 1119.9 KB | 47.0 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-object-mixed | 2.54 ms | OfficeIMO.Excel | Win | 1763.3 KB | 61.2 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse | 2.61 ms | OfficeIMO.Excel | Win | 1506.9 KB | 62.0 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | 2.32 ms | OfficeIMO.Excel | Win | 1507.0 KB | 62.0 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-scalars | 1.73 ms | OfficeIMO.Excel | Win | 1138.1 KB | 46.3 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-strings | 2.54 ms | OfficeIMO.Excel | Win | 2617.0 KB | 55.1 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-distinct | 2.41 ms | OfficeIMO.Excel | Win | 2379.2 KB | 51.8 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-repeated | 1.74 ms | OfficeIMO.Excel | Win | 1579.8 KB | 40.0 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-temporal | 3.07 ms | OfficeIMO.Excel | Win | 1435.7 KB | 63.3 KB |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | 1.47 ms | LargeXlsx | Loss +22.9% | 1092.0 KB | 48.2 KB |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | 3.85 ms | LargeXlsx | Loss +21.6% | 2081.1 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-datareader-plain | 3.84 ms | Sylvan.Data.Excel | Loss +17.5% | 1763.0 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-datareader-table | 4.09 ms | OfficeIMO.Excel | Win | 1774.9 KB | 138.8 KB |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | 4.11 ms | OfficeIMO.Excel | Win | 1781.2 KB | 138.9 KB |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | 4.16 ms | LargeXlsx | Loss +4.1% | 2140.6 KB | 131.1 KB |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | 4.48 ms | OfficeIMO.Excel | Win | 2880.2 KB | 176.0 KB |
| 2500 | package-profile | package | Package size | write-dataset-tables | 4.01 ms | OfficeIMO.Excel | Win | 2066.1 KB | 138.9 KB |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | 4.69 ms | OfficeIMO.Excel | Win | 2078.7 KB | 139.2 KB |
| 2500 | package-profile | package | Package size | write-datatable-direct | 3.73 ms | LargeXlsx | Loss +8.5% | 1748.6 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | 3.76 ms | OfficeIMO.Excel | Win | 1760.7 KB | 138.8 KB |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | 4.10 ms | LargeXlsx | Loss +35.4% | 1769.2 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | 4.14 ms | OfficeIMO.Excel | Win | 1347.1 KB | 142.4 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | 3.53 ms | LargeXlsx | Loss +15.6% | 1339.3 KB | 142.3 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | 5.25 ms | OfficeIMO.Excel | Win | 1505.3 KB | 138.1 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | 4.09 ms | LargeXlsx | Loss +26.5% | 1497.5 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | 4.72 ms | LargeXlsx | Loss +49.7% | 1770.1 KB | 142.3 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | 3.71 ms | OfficeIMO.Excel | Win | 1346.4 KB | 142.4 KB |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | 5.53 ms | LargeXlsx | Loss +37.4% | 2341.7 KB | 183.1 KB |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | 4.42 ms | LargeXlsx | Loss +9.9% | 1507.7 KB | 182.4 KB |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | 18.33 ms | OfficeIMO.Excel | Win | 4502.3 KB | 651.0 KB |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | 8.19 ms | OfficeIMO.Excel | Win | 1895.0 KB |  |
| 2500 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | 17.25 ms | OfficeIMO.Excel | Win | 15209.6 KB |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | 10.76 ms | OfficeIMO.Excel | Win | 6195.5 KB |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | 17.92 ms | OfficeIMO.Excel | Win | 16350.4 KB |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | 16.05 ms | OfficeIMO.Excel | Win | 15230.2 KB |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | 18.01 ms | OfficeIMO.Excel | Win | 15225.6 KB |  |
| 2500 | speed-comparison | other | Object projection | build-object-datatable-dictionaries | 1.36 ms | OfficeIMO.Excel | Win | 564.2 KB |  |
| 2500 | speed-comparison | other | Object projection | build-object-datatable-typed | 1.09 ms | OfficeIMO.Excel | Win | 856.9 KB |  |
| 2500 | speed-comparison | other | Range and table read | enumerate-cells | 5.42 ms | OfficeIMO.Excel | Win | 2531.6 KB |  |
| 2500 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | 3.62 ms | OfficeIMO.Excel | Win | 523.4 KB |  |
| 2500 | speed-comparison | other | Range and table read | enumerate-range | 5.88 ms | OfficeIMO.Excel | Win | 2531.7 KB |  |
| 2500 | speed-comparison | other | Range and table read | enumerate-top-range | 0.78 ms | OfficeIMO.Excel | Win | 285.2 KB |  |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-autofilter | 3.62 ms | OfficeIMO.Excel | Win | 1340.4 KB |  |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-charts | 5.03 ms | OfficeIMO.Excel | Win | 1892.8 KB |  |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | 3.59 ms | OfficeIMO.Excel | Win | 1405.8 KB |  |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-data-validation | 3.66 ms | OfficeIMO.Excel | Win | 1356.1 KB |  |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | 3.65 ms | OfficeIMO.Excel | Win | 1342.9 KB |  |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | 14.05 ms | OfficeIMO.Excel | Win | 14419.5 KB |  |
| 2500 | speed-comparison | other | Real-world report | realworld-report-all-in-one | 16.82 ms | OfficeIMO.Excel | Win | 15220.6 KB |  |
| 2500 | speed-comparison | other | Real-world report | realworld-report-core | 4.27 ms | OfficeIMO.Excel | Win | 1488.6 KB |  |
| 2500 | speed-comparison | other | Report workbook | report-workbook | 29.93 ms | OfficeIMO.Excel | Win | 19069.1 KB |  |
| 2500 | speed-comparison | other | Report workbook | report-workbook-core | 6.41 ms | OfficeIMO.Excel | Win | 2711.1 KB |  |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable | 23.96 ms | OfficeIMO.Excel | Win | 19383.5 KB |  |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable-core | 6.28 ms | OfficeIMO.Excel | Win | 2982.8 KB |  |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | 2.14 ms | OfficeIMO.Excel | Win | 706.6 KB |  |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | 0.81 ms | OfficeIMO.Excel | Win | 177.2 KB |  |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | 1.56 ms | Sylvan.Data.Excel | Loss +61.8% | 177.2 KB |  |
| 2500 | speed-comparison | read | Other | shared-string-read | 3.56 ms | Sylvan.Data.Excel | Loss +73.9% | 1056.5 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | 4.52 ms | OfficeIMO.Excel | Win | 374.5 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-datatable | 8.09 ms | Sylvan.Data.Excel | Loss +31.2% | 3594.4 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | 3.86 ms | OfficeIMO.Excel | Win | 542.8 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-range | 9.03 ms | OfficeIMO.Excel | Win | 2692.6 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | 4.93 ms | OfficeIMO.Excel, Sylvan.Data.Excel | Win | 2751.3 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-top-range | 0.53 ms | Sylvan.Data.Excel | Loss +21.6% | 296.0 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-used-range | 14.31 ms | Sylvan.Data.Excel | Loss +216.7% | 3472.7 KB |  |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | 3.99 ms | OfficeIMO.Excel, Sylvan.Data.Excel | Win | 377.7 KB |  |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | 5.52 ms | Sylvan.Data.Excel | Loss +25.4% | 2771.4 KB |  |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | 0.54 ms | Sylvan.Data.Excel | Loss +26.9% | 299.4 KB |  |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | 0.52 ms | Sylvan.Data.Excel | Loss +29.2% | 300.0 KB |  |
| 2500 | speed-comparison | read | Typed object read | read-objects | 7.01 ms | OfficeIMO.Excel | Win | 2441.9 KB |  |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | 4.72 ms | OfficeIMO.Excel, Sylvan.Data.Excel | Win | 2422.9 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | 5.10 ms | OfficeIMO.Excel | Win | 1781.2 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | 5.49 ms | OfficeIMO.Excel | Win | 2079.8 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | 3.96 ms | OfficeIMO.Excel | Win | 1347.1 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | 4.52 ms | OfficeIMO.Excel | Win | 1505.3 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | 4.06 ms | OfficeIMO.Excel | Win | 1346.4 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | 2.50 ms | OfficeIMO.Excel | Win | 1787.1 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | 2.60 ms | OfficeIMO.Excel | Win | 1119.9 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | 3.11 ms | OfficeIMO.Excel | Win | 1763.1 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | 3.15 ms | OfficeIMO.Excel | Win | 1506.7 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | 3.00 ms | OfficeIMO.Excel | Win | 1506.8 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | 2.68 ms | OfficeIMO.Excel | Win | 1138.1 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | 3.26 ms | OfficeIMO.Excel | Win | 1435.5 KB |  |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | 5.26 ms | OfficeIMO.Excel | Win | 2064.5 KB |  |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | 6.70 ms | OfficeIMO.Excel | Win | 2880.2 KB |  |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | 4.93 ms | OfficeIMO.Excel | Win | 2067.7 KB |  |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | 4.39 ms | OfficeIMO.Excel | Win | 1774.9 KB |  |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | 5.04 ms | OfficeIMO.Excel | Win | 1748.6 KB |  |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | 4.53 ms | OfficeIMO.Excel | Win | 1487.2 KB |  |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | 4.40 ms | OfficeIMO.Excel | Win | 1760.7 KB |  |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | 6.22 ms | OfficeIMO.Excel | Win | 1403.3 KB |  |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | 3.13 ms | OfficeIMO.Excel | Win | 1620.6 KB |  |
| 2500 | speed-comparison | write | Other | write-dictionary-objects-table-direct | 5.46 ms | OfficeIMO.Excel | Win | 2051.4 KB |  |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | 5.42 ms | LargeXlsx | Loss +30.7% | 2341.7 KB |  |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | 5.24 ms | LargeXlsx | Loss +26.7% | 1507.7 KB |  |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | 22.20 ms | LargeXlsx | Loss +10.8% | 4502.3 KB |  |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | 2.16 ms | LargeXlsx | Loss +56.7% | 1576.3 KB |  |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | 1.58 ms | OfficeIMO.Excel | Win | 1092.0 KB |  |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | 4.33 ms | OfficeIMO.Excel | Win | 2081.1 KB |  |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | 3.38 ms | OfficeIMO.Excel | Win | 1494.0 KB |  |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | 4.68 ms | Sylvan.Data.Excel | Loss +29.7% | 1763.0 KB |  |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | 5.69 ms | OfficeIMO.Excel | Win | 2140.6 KB |  |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | 4.76 ms | OfficeIMO.Excel | Win | 1676.8 KB |  |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | 2.04 ms | OfficeIMO.Excel | Win | 2440.3 KB |  |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | 3.17 ms | OfficeIMO.Excel | Win | 2617.0 KB |  |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | 2.48 ms | OfficeIMO.Excel | Win | 2379.2 KB |  |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | 2.27 ms | OfficeIMO.Excel | Win | 1579.8 KB |  |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | 3.86 ms | LargeXlsx | Loss +18.1% | 1769.2 KB |  |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | 3.71 ms | LargeXlsx | Loss +17.7% | 1339.3 KB |  |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | 4.33 ms | LargeXlsx | Loss +34.7% | 1497.5 KB |  |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | 81.10 ms | Sylvan.Data.Excel | Loss +10.1% | 23621.9 KB |  |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | 53.33 ms | Sylvan.Data.Excel | Loss +30.5% | 24404.0 KB |  |
| 25000 | package-profile | package | Package size | append-plain-rows | 14.28 ms | LargeXlsx | Loss +20.8% | 10842.5 KB | 610.4 KB |
| 25000 | package-profile | package | Package size | autofit-existing | 81.97 ms | OfficeIMO.Excel | Win | 15710.9 KB | 1429.3 KB |
| 25000 | package-profile | package | Package size | large-shared-strings | 23.48 ms | OfficeIMO.Excel | Win | 15744.9 KB | 529.7 KB |
| 25000 | package-profile | package | Package size | realworld-autofilter | 32.88 ms | OfficeIMO.Excel | Win | 11494.9 KB | 1428.5 KB |
| 25000 | package-profile | package | Package size | realworld-charts | 34.79 ms | OfficeIMO.Excel | Win | 12553.1 KB | 1433.7 KB |
| 25000 | package-profile | package | Package size | realworld-conditional-formatting | 32.38 ms | OfficeIMO.Excel | Win | 11560.2 KB | 1428.8 KB |
| 25000 | package-profile | package | Package size | realworld-data-validation | 35.11 ms | OfficeIMO.Excel | Win | 11510.5 KB | 1428.6 KB |
| 25000 | package-profile | package | Package size | realworld-freeze-panes | 33.44 ms | OfficeIMO.Excel | Win | 11497.3 KB | 1428.6 KB |
| 25000 | package-profile | package | Package size | realworld-pivot-table | 282.43 ms | OfficeIMO.Excel | Win | 131929.1 KB | 1979.8 KB |
| 25000 | package-profile | package | Package size | realworld-report-all-in-one | 259.80 ms | OfficeIMO.Excel | Win | 133444.7 KB | 1985.9 KB |
| 25000 | package-profile | package | Package size | realworld-report-chart-first | 91.62 ms | OfficeIMO.Excel | Win | 43560.6 KB | 1985.9 KB |
| 25000 | package-profile | package | Package size | realworld-report-core | 59.32 ms | OfficeIMO.Excel | Win | 11648.7 KB | 1430.0 KB |
| 25000 | package-profile | package | Package size | realworld-report-extra-column | 358.46 ms | OfficeIMO.Excel | Win | 144827.2 KB | 2110.8 KB |
| 25000 | package-profile | package | Package size | realworld-report-no-autofit | 265.40 ms | OfficeIMO.Excel | Win | 133435.8 KB | 1985.8 KB |
| 25000 | package-profile | package | Package size | realworld-report-post-mutation | 274.82 ms | OfficeIMO.Excel | Win | 133463.2 KB | 1986.0 KB |
| 25000 | package-profile | package | Package size | realworld-report-shuffled-columns | 296.24 ms | OfficeIMO.Excel | Win | 133503.4 KB | 2046.1 KB |
| 25000 | package-profile | package | Package size | report-workbook | 332.42 ms | OfficeIMO.Excel | Win | 175197.5 KB | 2672.0 KB |
| 25000 | package-profile | package | Package size | report-workbook-core | 49.38 ms | OfficeIMO.Excel | Win | 10979.4 KB | 1850.9 KB |
| 25000 | package-profile | package | Package size | report-workbook-datatable | 358.15 ms | OfficeIMO.Excel | Win | 177940.5 KB | 2672.0 KB |
| 25000 | package-profile | package | Package size | report-workbook-datatable-core | 48.24 ms | OfficeIMO.Excel | Win | 13725.0 KB | 1850.9 KB |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | 43.08 ms | LargeXlsx | Loss +12.0% | 11708.2 KB | 2228.8 KB |
| 25000 | package-profile | package | Package size | write-bulk-report | 35.34 ms | OfficeIMO.Excel | Win | 11561.8 KB | 1429.3 KB |
| 25000 | package-profile | package | Package size | write-cellformula | 19.37 ms | OfficeIMO.Excel | Win | 10112.0 KB | 670.3 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-empty-strings | 11.89 ms | OfficeIMO.Excel | Win | 6896.4 KB | 451.4 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-numbers | 17.33 ms | OfficeIMO.Excel | Win | 5970.9 KB | 462.6 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-object-mixed | 17.11 ms | OfficeIMO.Excel | Win | 8332.9 KB | 585.1 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse | 20.02 ms | OfficeIMO.Excel | Win | 7416.2 KB | 607.1 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | 17.81 ms | OfficeIMO.Excel | Win | 7416.3 KB | 607.1 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-scalars | 11.00 ms | OfficeIMO.Excel | Win | 6144.6 KB | 441.9 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-strings | 16.75 ms | OfficeIMO.Excel | Win | 15360.4 KB | 527.8 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-distinct | 12.87 ms | OfficeIMO.Excel | Win | 13824.1 KB | 499.0 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-repeated | 12.00 ms | OfficeIMO.Excel | Win | 7525.3 KB | 376.0 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-temporal | 22.92 ms | OfficeIMO.Excel | Win | 7482.8 KB | 620.5 KB |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | 10.47 ms | LargeXlsx | Loss +7.5% | 6961.7 KB | 455.5 KB |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | 34.60 ms | LargeXlsx | Loss +21.2% | 16036.5 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-datareader-plain | 38.11 ms | Sylvan.Data.Excel | Loss +34.9% | 13002.3 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-datareader-table | 38.27 ms | OfficeIMO.Excel | Win | 13020.3 KB | 1385.7 KB |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | 40.72 ms | OfficeIMO.Excel | Win | 13026.6 KB | 1385.9 KB |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | 33.90 ms | OfficeIMO.Excel | Win | 9819.7 KB | 1329.2 KB |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | 41.08 ms | OfficeIMO.Excel | Win | 13458.5 KB | 1795.1 KB |
| 25000 | package-profile | package | Package size | write-dataset-tables | 36.50 ms | OfficeIMO.Excel | Win | 10288.1 KB | 1376.4 KB |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | 38.99 ms | OfficeIMO.Excel | Win | 10300.7 KB | 1376.7 KB |
| 25000 | package-profile | package | Package size | write-datatable-direct | 35.16 ms | LargeXlsx | Loss +14.1% | 12715.7 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | 35.23 ms | OfficeIMO.Excel | Win | 12733.8 KB | 1385.7 KB |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | 32.40 ms | LargeXlsx | Loss +14.4% | 12912.0 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | 36.13 ms | OfficeIMO.Excel | Win | 11501.6 KB | 1428.5 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | 32.30 ms | OfficeIMO.Excel | Win | 11493.8 KB | 1428.4 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | 40.38 ms | OfficeIMO.Excel | Win | 10187.2 KB | 1385.1 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | 37.68 ms | LargeXlsx | Loss +29.0% | 10179.4 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | 42.03 ms | LargeXlsx | Loss +52.0% | 15791.7 KB | 1428.4 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | 33.96 ms | OfficeIMO.Excel | Win | 11500.9 KB | 1428.5 KB |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | 45.07 ms | LargeXlsx | Loss +9.0% | 10577.2 KB | 1828.0 KB |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | 42.36 ms | LargeXlsx | Loss +14.7% | 9942.2 KB | 1828.0 KB |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | 188.19 ms | OfficeIMO.Excel | Win | 36150.1 KB | 6725.6 KB |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | 109.90 ms | OfficeIMO.Excel | Win | 15708.2 KB |  |
| 25000 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | 289.71 ms | EPPlus 4.5.3.3 | Loss +19.2% | 133435.3 KB |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | 89.72 ms | OfficeIMO.Excel | Win | 43560.5 KB |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | 306.11 ms | OfficeIMO.Excel | Win | 144825.1 KB |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | 258.71 ms | OfficeIMO.Excel | Win | 133461.2 KB |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | 286.40 ms | OfficeIMO.Excel | Win | 133506.4 KB |  |
| 25000 | speed-comparison | other | Object projection | build-object-datatable-dictionaries | 9.68 ms | OfficeIMO.Excel | Win | 5164.3 KB |  |
| 25000 | speed-comparison | other | Object projection | build-object-datatable-typed | 9.18 ms | OfficeIMO.Excel | Win | 8093.8 KB |  |
| 25000 | speed-comparison | other | Range and table read | enumerate-cells | 49.67 ms | OfficeIMO.Excel | Win | 24531.0 KB |  |
| 25000 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | 34.87 ms | OfficeIMO.Excel | Win | 3839.3 KB |  |
| 25000 | speed-comparison | other | Range and table read | enumerate-range | 47.32 ms | OfficeIMO.Excel | Win | 24531.1 KB |  |
| 25000 | speed-comparison | other | Range and table read | enumerate-top-range | 0.66 ms | OfficeIMO.Excel | Win | 285.4 KB |  |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-autofilter | 31.43 ms | OfficeIMO.Excel | Win | 11494.9 KB |  |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-charts | 32.55 ms | OfficeIMO.Excel | Win | 12554.2 KB |  |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | 33.10 ms | OfficeIMO.Excel | Win | 11560.2 KB |  |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-data-validation | 32.92 ms | OfficeIMO.Excel | Win | 11510.5 KB |  |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | 31.83 ms | OfficeIMO.Excel | Win | 11497.3 KB |  |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | 216.45 ms | OfficeIMO.Excel | Win | 131922.9 KB |  |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | 288.90 ms | OfficeIMO.Excel | Win | 133445.5 KB |  |
| 25000 | speed-comparison | other | Real-world report | realworld-report-core | 35.23 ms | OfficeIMO.Excel | Win | 11648.7 KB |  |
| 25000 | speed-comparison | other | Report workbook | report-workbook | 483.21 ms | OfficeIMO.Excel | Win | 175196.4 KB |  |
| 25000 | speed-comparison | other | Report workbook | report-workbook-core | 51.79 ms | OfficeIMO.Excel | Win | 10979.4 KB |  |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable | 434.69 ms | OfficeIMO.Excel | Win | 177944.2 KB |  |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable-core | 51.48 ms | OfficeIMO.Excel | Win | 13725.0 KB |  |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | 18.36 ms | OfficeIMO.Excel | Win | 6219.0 KB |  |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | 0.93 ms | OfficeIMO.Excel | Win | 177.2 KB |  |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | 0.84 ms | OfficeIMO.Excel | Win | 177.3 KB |  |
| 25000 | speed-comparison | read | Other | shared-string-read | 19.32 ms | Sylvan.Data.Excel | Loss +13.6% | 9218.0 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | 32.61 ms | OfficeIMO.Excel | Win | 1122.4 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-datatable | 67.43 ms | Sylvan.Data.Excel | Loss +4.5% | 34645.9 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | 34.81 ms | OfficeIMO.Excel | Win | 4034.6 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-range | 51.75 ms | Sylvan.Data.Excel | Loss +12.9% | 26098.3 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | 54.18 ms | Sylvan.Data.Excel, OfficeIMO.Excel | Win | 26684.2 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-top-range | 0.66 ms | Sylvan.Data.Excel | Loss +26.6% | 296.1 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-used-range | 97.82 ms | Sylvan.Data.Excel | Loss +112.0% | 34151.8 KB |  |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | 34.14 ms | OfficeIMO.Excel | Win | 1125.7 KB |  |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | 88.69 ms | Sylvan.Data.Excel | Loss +16.7% | 26883.9 KB |  |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | 0.61 ms | OfficeIMO.Excel, Sylvan.Data.Excel | Win | 299.5 KB |  |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | 0.72 ms | Sylvan.Data.Excel | Loss +24.6% | 300.2 KB |  |
| 25000 | speed-comparison | read | Typed object read | read-objects | 52.42 ms | Sylvan.Data.Excel | Loss +15.9% | 23562.4 KB |  |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | 113.12 ms | OfficeIMO.Excel | Win | 23367.5 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | 39.05 ms | OfficeIMO.Excel | Win | 13026.6 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | 38.25 ms | OfficeIMO.Excel | Win | 10300.7 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | 41.09 ms | OfficeIMO.Excel | Win | 11501.6 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | 43.12 ms | OfficeIMO.Excel | Win | 10187.2 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | 34.88 ms | OfficeIMO.Excel | Win | 11500.9 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | 12.51 ms | OfficeIMO.Excel | Win | 6896.4 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | 15.63 ms | OfficeIMO.Excel | Win | 5970.9 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | 26.55 ms | OfficeIMO.Excel | Win | 8332.9 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | 29.96 ms | OfficeIMO.Excel | Win | 7416.2 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | 22.18 ms | OfficeIMO.Excel | Win | 7416.3 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | 11.28 ms | OfficeIMO.Excel | Win | 6144.6 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | 30.82 ms | OfficeIMO.Excel | Win | 7482.8 KB |  |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | 75.48 ms | OfficeIMO.Excel | Win | 13039.6 KB |  |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | 68.65 ms | OfficeIMO.Excel | Win | 13458.5 KB |  |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | 33.25 ms | OfficeIMO.Excel | Win | 10288.1 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | 39.84 ms | OfficeIMO.Excel | Win | 13020.3 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | 34.78 ms | LargeXlsx | Loss +10.7% | 12715.7 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | 37.25 ms | OfficeIMO.Excel | Win | 9999.4 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | 34.52 ms | OfficeIMO.Excel | Win | 12733.8 KB |  |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | 38.05 ms | OfficeIMO.Excel | Win | 11561.8 KB |  |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | 20.47 ms | OfficeIMO.Excel | Win | 10112.0 KB |  |
| 25000 | speed-comparison | write | Other | write-dictionary-objects-table-direct | 50.66 ms | OfficeIMO.Excel | Win | 15163.8 KB |  |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | 44.33 ms | LargeXlsx | Loss +6.7% | 10577.2 KB |  |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | 44.61 ms | LargeXlsx | Loss +14.8% | 9942.2 KB |  |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | 198.92 ms | OfficeIMO.Excel | Win | 36150.1 KB |  |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | 16.00 ms | LargeXlsx | Loss +42.0% | 10842.5 KB |  |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | 11.18 ms | LargeXlsx | Loss +10.8% | 6961.7 KB |  |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | 36.05 ms | LargeXlsx | Loss +15.6% | 16036.5 KB |  |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | 15.77 ms | OfficeIMO.Excel | Win | 7866.1 KB |  |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | 37.13 ms | Sylvan.Data.Excel | Loss +33.9% | 13002.3 KB |  |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | 45.54 ms | LargeXlsx | Loss +10.5% | 9819.7 KB |  |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | 47.57 ms | LargeXlsx | Loss +20.1% | 11708.2 KB |  |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | 14.40 ms | OfficeIMO.Excel | Win | 15744.9 KB |  |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | 16.70 ms | OfficeIMO.Excel | Win | 15360.4 KB |  |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | 13.10 ms | OfficeIMO.Excel | Win | 13824.1 KB |  |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | 11.53 ms | OfficeIMO.Excel | Win | 7525.3 KB |  |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | 32.94 ms | LargeXlsx | Loss +22.0% | 12912.0 KB |  |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | 32.90 ms | LargeXlsx | Loss +8.3% | 11493.8 KB |  |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | 37.15 ms | LargeXlsx | Loss +31.7% | 10179.4 KB |  |

## Full comparison table

| Row count | Artifact | Workload | Category | Scenario | Library | Mean | StdDev | StdErr | Ratio to OfficeIMO | Ratio to best | Alloc | Alloc ratio | Package | Package ratio | Outcome |
| ---: | --- | --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | Sylvan.Data.Excel | 4.50 ms | 0.74 ms | 0.43 ms | 0.76 | 1.00 | 362.3 KB | 0.15 |  |  | 24.4% faster than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | OfficeIMO.Excel | 5.95 ms | 0.27 ms | 0.16 ms | 1.00 | 1.32 | 2410.9 KB | 1.00 |  |  | Loss +32.3% |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | ExcelDataReader | 11.52 ms | 0.41 ms | 0.24 ms | 1.94 | 2.56 | 6887.4 KB | 2.86 |  |  | 93.6% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | MiniExcel | 16.15 ms | 2.17 ms | 1.25 ms | 2.71 | 3.59 | 21507.3 KB | 8.92 |  |  | 171.3% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | Sylvan.Data.Excel | 4.14 ms | 0.24 ms | 0.14 ms | 0.66 | 1.00 | 362.3 KB | 0.15 |  |  | 34.3% faster than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | OfficeIMO.Excel | 6.31 ms | 0.35 ms | 0.20 ms | 1.00 | 1.52 | 2489.3 KB | 1.00 |  |  | Loss +52.2% |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | ExcelDataReader | 12.91 ms | 0.93 ms | 0.54 ms | 2.05 | 3.11 | 6887.4 KB | 2.77 |  |  | 104.7% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | MiniExcel | 18.16 ms | 1.19 ms | 0.69 ms | 2.88 | 4.38 | 21507.3 KB | 8.64 |  |  | 188.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | append-plain-rows | LargeXlsx | 1.45 ms | 0.21 ms | 0.12 ms | 0.77 | 1.00 | 296.4 KB | 0.19 | 63.1 KB | 1.00 | 23.0% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | append-plain-rows | OfficeIMO.Excel | 1.89 ms | 0.07 ms | 0.04 ms | 1.00 | 1.30 | 1576.3 KB | 1.00 | 63.0 KB | 1.00 | Loss +29.9% |
| 2500 | package-profile | package | Package size | append-plain-rows | MiniExcel | 3.84 ms | 0.07 ms | 0.04 ms | 2.04 | 2.65 | 19710.8 KB | 12.50 | 68.1 KB | 1.08 | 103.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | append-plain-rows | ClosedXML | 14.83 ms | 1.08 ms | 0.63 ms | 7.87 | 10.22 | 11197.4 KB | 7.10 | 59.8 KB | 0.95 | 686.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | append-plain-rows | EPPlus | 25.58 ms | 0.22 ms | 0.13 ms | 13.56 | 17.63 | 14365.6 KB | 9.11 | 56.9 KB | 0.90 | 1256.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | autofit-existing | OfficeIMO.Excel | 8.03 ms | 0.12 ms | 0.07 ms | 1.00 | 1.00 | 1895.2 KB | 1.00 | 143.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | autofit-existing | EPPlus | 77.65 ms | 2.64 ms | 1.53 ms | 9.67 | 9.67 | 50712.1 KB | 26.76 | 115.0 KB | 0.80 | 866.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | autofit-existing | ClosedXML | 129.85 ms | 3.71 ms | 2.14 ms | 16.17 | 16.17 | 84562.6 KB | 44.62 | 121.0 KB | 0.84 | 1516.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | large-shared-strings | OfficeIMO.Excel | 2.07 ms | 0.07 ms | 0.04 ms | 1.00 | 1.00 | 2440.3 KB | 1.00 | 55.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | large-shared-strings | MiniExcel | 4.65 ms | 1.17 ms | 0.68 ms | 2.24 | 2.24 | 21137.5 KB | 8.66 | 60.7 KB | 1.10 | 124.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | large-shared-strings | ClosedXML | 13.20 ms | 1.48 ms | 0.85 ms | 6.37 | 6.37 | 11299.2 KB | 4.63 | 50.3 KB | 0.91 | 536.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | large-shared-strings | EPPlus | 24.38 ms | 2.97 ms | 1.72 ms | 11.76 | 11.76 | 12804.8 KB | 5.25 | 48.1 KB | 0.87 | 1076.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-autofilter | OfficeIMO.Excel | 3.85 ms | 0.36 ms | 0.21 ms | 1.00 | 1.00 | 1340.4 KB | 1.00 | 142.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-autofilter | ClosedXML | 30.44 ms | 1.25 ms | 0.72 ms | 7.91 | 7.91 | 22226.8 KB | 16.58 | 120.2 KB | 0.84 | 691.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-autofilter | EPPlus | 42.56 ms | 2.68 ms | 1.55 ms | 11.07 | 11.07 | 24715.8 KB | 18.44 | 114.2 KB | 0.80 | 1006.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-charts | OfficeIMO.Excel | 4.87 ms | 0.07 ms | 0.04 ms | 1.00 | 1.00 | 1892.9 KB | 1.00 | 147.6 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-charts | EPPlus | 41.15 ms | 0.90 ms | 0.52 ms | 8.45 | 8.45 | 27142.8 KB | 14.34 | 117.0 KB | 0.79 | 744.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-conditional-formatting | OfficeIMO.Excel | 3.88 ms | 0.54 ms | 0.31 ms | 1.00 | 1.00 | 1405.8 KB | 1.00 | 142.7 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-conditional-formatting | ClosedXML | 31.02 ms | 0.73 ms | 0.42 ms | 8.00 | 8.00 | 22273.8 KB | 15.84 | 120.3 KB | 0.84 | 700.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-conditional-formatting | EPPlus | 42.23 ms | 1.13 ms | 0.66 ms | 10.89 | 10.89 | 24757.8 KB | 17.61 | 114.3 KB | 0.80 | 989.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-data-validation | OfficeIMO.Excel | 3.77 ms | 0.16 ms | 0.09 ms | 1.00 | 1.00 | 1356.1 KB | 1.00 | 142.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-data-validation | ClosedXML | 34.03 ms | 3.40 ms | 1.96 ms | 9.02 | 9.02 | 22247.9 KB | 16.41 | 120.3 KB | 0.84 | 802.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-data-validation | EPPlus | 41.69 ms | 2.49 ms | 1.44 ms | 11.05 | 11.05 | 24701.8 KB | 18.22 | 114.2 KB | 0.80 | 1005.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-freeze-panes | OfficeIMO.Excel | 3.83 ms | 0.42 ms | 0.24 ms | 1.00 | 1.00 | 1342.8 KB | 1.00 | 142.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-freeze-panes | ClosedXML | 31.93 ms | 2.21 ms | 1.27 ms | 8.34 | 8.34 | 22222.0 KB | 16.55 | 120.2 KB | 0.84 | 734.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-freeze-panes | EPPlus | 46.38 ms | 5.15 ms | 2.97 ms | 12.12 | 12.12 | 24730.4 KB | 18.42 | 114.3 KB | 0.80 | 1111.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-pivot-table | OfficeIMO.Excel | 13.50 ms | 0.11 ms | 0.06 ms | 1.00 | 1.00 | 14419.5 KB | 1.00 | 200.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-pivot-table | EPPlus | 45.40 ms | 2.45 ms | 1.42 ms | 3.36 | 3.36 | 29537.8 KB | 2.05 | 117.4 KB | 0.59 | 236.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-all-in-one | OfficeIMO.Excel | 16.27 ms | 1.10 ms | 0.64 ms | 1.00 | 1.00 | 15220.7 KB | 1.00 | 206.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-all-in-one | EPPlus | 73.85 ms | 2.75 ms | 1.59 ms | 4.54 | 4.54 | 54595.6 KB | 3.59 | 121.8 KB | 0.59 | 354.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-chart-first | OfficeIMO.Excel | 10.29 ms | 0.09 ms | 0.05 ms | 1.00 | 1.00 | 6196.8 KB | 1.00 | 206.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-chart-first | EPPlus | 68.14 ms | 0.55 ms | 0.32 ms | 6.63 | 6.63 | 54595.0 KB | 8.81 | 121.8 KB | 0.59 | 562.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-core | OfficeIMO.Excel | 4.45 ms | 0.21 ms | 0.12 ms | 1.00 | 1.00 | 1488.5 KB | 1.00 | 143.9 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-core | EPPlus | 71.66 ms | 3.10 ms | 1.79 ms | 16.10 | 16.10 | 47300.2 KB | 31.78 | 115.6 KB | 0.80 | 1509.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-core | ClosedXML | 80.71 ms | 4.79 ms | 2.77 ms | 18.13 | 18.13 | 69836.4 KB | 46.92 | 121.5 KB | 0.84 | 1713.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-extra-column | OfficeIMO.Excel | 17.40 ms | 1.87 ms | 1.08 ms | 1.00 | 1.00 | 16350.5 KB | 1.00 | 219.1 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-extra-column | EPPlus | 75.60 ms | 2.71 ms | 1.57 ms | 4.35 | 4.35 | 59227.2 KB | 3.62 | 128.4 KB | 0.59 | 334.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-no-autofit | OfficeIMO.Excel | 14.90 ms | 0.12 ms | 0.07 ms | 1.00 | 1.00 | 15211.1 KB | 1.00 | 206.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-no-autofit | EPPlus | 44.06 ms | 2.22 ms | 1.28 ms | 2.96 | 2.96 | 32907.4 KB | 2.16 | 121.8 KB | 0.59 | 195.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-post-mutation | OfficeIMO.Excel | 15.91 ms | 0.84 ms | 0.49 ms | 1.00 | 1.00 | 15230.0 KB | 1.00 | 206.6 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-post-mutation | EPPlus | 73.35 ms | 6.42 ms | 3.71 ms | 4.61 | 4.61 | 54595.7 KB | 3.58 | 121.9 KB | 0.59 | 360.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-shuffled-columns | OfficeIMO.Excel | 17.47 ms | 0.82 ms | 0.47 ms | 1.00 | 1.00 | 15225.5 KB | 1.00 | 211.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-shuffled-columns | EPPlus | 74.73 ms | 5.41 ms | 3.12 ms | 4.28 | 4.28 | 54592.2 KB | 3.59 | 124.3 KB | 0.59 | 327.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook | OfficeIMO.Excel | 23.84 ms | 1.67 ms | 0.97 ms | 1.00 | 1.00 | 19112.3 KB | 1.00 | 275.6 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | report-workbook | EPPlus | 99.75 ms | 5.73 ms | 3.31 ms | 4.18 | 4.18 | 77486.7 KB | 4.05 | 161.8 KB | 0.59 | 318.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook-core | OfficeIMO.Excel | 5.79 ms | 0.09 ms | 0.05 ms | 1.00 | 1.00 | 2711.1 KB | 1.00 | 187.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | report-workbook-core | EPPlus | 93.82 ms | 0.50 ms | 0.29 ms | 16.20 | 16.20 | 71970.9 KB | 26.55 | 157.2 KB | 0.84 | 1519.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook-core | ClosedXML | 99.15 ms | 2.28 ms | 1.32 ms | 17.12 | 17.12 | 97220.0 KB | 35.86 | 165.1 KB | 0.88 | 1611.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook-datatable | OfficeIMO.Excel | 22.56 ms | 0.82 ms | 0.47 ms | 1.00 | 1.00 | 19383.9 KB | 1.00 | 275.6 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | report-workbook-datatable | EPPlus | 99.40 ms | 3.31 ms | 1.91 ms | 4.41 | 4.41 | 65995.9 KB | 3.40 | 161.8 KB | 0.59 | 340.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook-datatable-core | OfficeIMO.Excel | 5.75 ms | 0.34 ms | 0.19 ms | 1.00 | 1.00 | 2982.7 KB | 1.00 | 187.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | report-workbook-datatable-core | EPPlus | 92.65 ms | 3.09 ms | 1.78 ms | 16.11 | 16.11 | 60480.4 KB | 20.28 | 157.2 KB | 0.84 | 1510.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook-datatable-core | ClosedXML | 96.01 ms | 3.90 ms | 2.25 ms | 16.69 | 16.69 | 82860.8 KB | 27.78 | 165.1 KB | 0.88 | 1569.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | LargeXlsx | 3.94 ms | 0.11 ms | 0.06 ms | 0.81 | 1.00 | 857.6 KB | 0.51 | 237.7 KB | 1.10 | 18.6% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | OfficeIMO.Excel | 4.84 ms | 0.50 ms | 0.29 ms | 1.00 | 1.23 | 1676.8 KB | 1.00 | 216.7 KB | 1.00 | Loss +22.8% |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | MiniExcel | 16.17 ms | 0.29 ms | 0.17 ms | 3.34 | 4.10 | 35918.9 KB | 21.42 | 235.3 KB | 1.09 | 234.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | ClosedXML | 92.68 ms | 6.30 ms | 3.64 ms | 19.15 | 23.52 | 71478.2 KB | 42.63 | 257.2 KB | 1.19 | 1815.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-bulk-report | OfficeIMO.Excel | 4.52 ms | 0.37 ms | 0.21 ms | 1.00 | 1.00 | 1401.7 KB | 1.00 | 143.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-bulk-report | MiniExcel | 8.06 ms | 0.38 ms | 0.22 ms | 1.78 | 1.78 | 26825.5 KB | 19.14 | 153.8 KB | 1.07 | 78.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-bulk-report | EPPlus | 64.65 ms | 3.12 ms | 1.80 ms | 14.29 | 14.29 | 47194.2 KB | 33.67 | 115.0 KB | 0.80 | 1329.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-bulk-report | ClosedXML | 72.32 ms | 2.28 ms | 1.32 ms | 15.99 | 15.99 | 58348.8 KB | 41.63 | 121.0 KB | 0.84 | 1498.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellformula | OfficeIMO.Excel | 2.31 ms | 0.07 ms | 0.04 ms | 1.00 | 1.00 | 1383.3 KB | 1.00 | 66.6 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellformula | ClosedXML | 17.18 ms | 0.55 ms | 0.32 ms | 7.45 | 7.45 | 12039.8 KB | 8.70 | 70.6 KB | 1.06 | 644.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellformula | EPPlus | 37.17 ms | 1.40 ms | 0.81 ms | 16.11 | 16.11 | 18110.8 KB | 13.09 | 62.1 KB | 0.93 | 1510.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-empty-strings | OfficeIMO.Excel | 1.92 ms | 0.01 ms | 0.01 ms | 1.00 | 1.00 | 1787.1 KB | 1.00 | 44.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-empty-strings | ClosedXML | 13.09 ms | 1.46 ms | 0.84 ms | 6.80 | 6.80 | 9959.5 KB | 5.57 | 44.9 KB | 1.02 | 580.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-empty-strings | EPPlus | 25.15 ms | 2.12 ms | 1.22 ms | 13.07 | 13.07 | 11773.4 KB | 6.59 | 42.0 KB | 0.95 | 1207.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-numbers | OfficeIMO.Excel | 1.85 ms | 0.07 ms | 0.04 ms | 1.00 | 1.00 | 1119.9 KB | 1.00 | 47.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-numbers | ClosedXML | 10.86 ms | 0.52 ms | 0.30 ms | 5.86 | 5.86 | 9177.1 KB | 8.19 | 45.9 KB | 0.98 | 485.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-numbers | EPPlus | 22.47 ms | 0.43 ms | 0.25 ms | 12.12 | 12.12 | 12895.6 KB | 11.51 | 43.7 KB | 0.93 | 1112.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-mixed | OfficeIMO.Excel | 2.54 ms | 0.11 ms | 0.06 ms | 1.00 | 1.00 | 1763.3 KB | 1.00 | 61.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-object-mixed | ClosedXML | 16.14 ms | 0.41 ms | 0.24 ms | 6.35 | 6.35 | 11887.0 KB | 6.74 | 59.5 KB | 0.97 | 534.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-mixed | EPPlus | 26.30 ms | 2.30 ms | 1.33 ms | 10.34 | 10.34 | 15643.7 KB | 8.87 | 58.9 KB | 0.96 | 934.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse | OfficeIMO.Excel | 2.61 ms | 0.26 ms | 0.15 ms | 1.00 | 1.00 | 1506.9 KB | 1.00 | 62.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse | ClosedXML | 15.77 ms | 1.27 ms | 0.73 ms | 6.05 | 6.05 | 11296.3 KB | 7.50 | 52.5 KB | 0.85 | 504.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse | EPPlus | 25.38 ms | 0.92 ms | 0.53 ms | 9.73 | 9.73 | 14960.7 KB | 9.93 | 54.2 KB | 0.88 | 873.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 2.32 ms | 0.02 ms | 0.01 ms | 1.00 | 1.00 | 1507.0 KB | 1.00 | 62.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | ClosedXML | 15.00 ms | 0.62 ms | 0.36 ms | 6.47 | 6.47 | 11296.3 KB | 7.50 | 52.5 KB | 0.85 | 547.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | EPPlus | 26.06 ms | 1.11 ms | 0.64 ms | 11.24 | 11.24 | 14960.7 KB | 9.93 | 54.2 KB | 0.88 | 1024.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-scalars | OfficeIMO.Excel | 1.73 ms | 0.23 ms | 0.13 ms | 1.00 | 1.00 | 1138.1 KB | 1.00 | 46.3 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-scalars | ClosedXML | 11.16 ms | 0.50 ms | 0.29 ms | 6.44 | 6.44 | 9021.2 KB | 7.93 | 45.4 KB | 0.98 | 543.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-scalars | EPPlus | 23.92 ms | 1.11 ms | 0.64 ms | 13.80 | 13.80 | 12827.9 KB | 11.27 | 42.4 KB | 0.91 | 1280.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings | OfficeIMO.Excel | 2.54 ms | 0.29 ms | 0.17 ms | 1.00 | 1.00 | 2617.0 KB | 1.00 | 55.1 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-strings | ClosedXML | 12.19 ms | 0.76 ms | 0.44 ms | 4.81 | 4.81 | 11299.2 KB | 4.32 | 50.3 KB | 0.91 | 380.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings | EPPlus | 24.36 ms | 2.10 ms | 1.21 ms | 9.61 | 9.61 | 12805.3 KB | 4.89 | 48.1 KB | 0.87 | 861.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-distinct | OfficeIMO.Excel | 2.41 ms | 0.49 ms | 0.28 ms | 1.00 | 1.00 | 2379.2 KB | 1.00 | 51.8 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-distinct | ClosedXML | 15.38 ms | 2.25 ms | 1.30 ms | 6.39 | 6.39 | 13127.1 KB | 5.52 | 61.9 KB | 1.19 | 539.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-distinct | EPPlus | 26.81 ms | 0.41 ms | 0.24 ms | 11.14 | 11.14 | 13893.4 KB | 5.84 | 61.5 KB | 1.19 | 1014.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-repeated | OfficeIMO.Excel | 1.74 ms | 0.04 ms | 0.03 ms | 1.00 | 1.00 | 1579.8 KB | 1.00 | 40.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-repeated | ClosedXML | 10.61 ms | 0.79 ms | 0.46 ms | 6.09 | 6.09 | 9226.5 KB | 5.84 | 38.8 KB | 0.97 | 509.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-repeated | EPPlus | 20.33 ms | 0.87 ms | 0.50 ms | 11.68 | 11.68 | 11332.9 KB | 7.17 | 34.8 KB | 0.87 | 1067.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-temporal | OfficeIMO.Excel | 3.07 ms | 0.68 ms | 0.39 ms | 1.00 | 1.00 | 1435.7 KB | 1.00 | 63.3 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-temporal | ClosedXML | 16.52 ms | 1.90 ms | 1.10 ms | 5.38 | 5.38 | 9711.1 KB | 6.76 | 54.5 KB | 0.86 | 437.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-temporal | EPPlus | 27.36 ms | 1.20 ms | 0.70 ms | 8.91 | 8.91 | 14723.0 KB | 10.25 | 53.1 KB | 0.84 | 790.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 1.19 ms | 0.01 ms | 0.00 ms | 0.81 | 1.00 | 447.0 KB | 0.41 | 47.3 KB | 0.98 | 18.6% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 1.47 ms | 0.02 ms | 0.01 ms | 1.00 | 1.23 | 1092.0 KB | 1.00 | 48.2 KB | 1.00 | Loss +22.9% |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | ClosedXML | 11.73 ms | 0.18 ms | 0.10 ms | 8.01 | 9.84 | 10235.8 KB | 9.37 | 53.0 KB | 1.10 | 700.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | EPPlus | 21.13 ms | 0.89 ms | 0.52 ms | 14.42 | 17.71 | 13052.5 KB | 11.95 | 52.5 KB | 1.09 | 1341.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | LargeXlsx | 3.16 ms | 0.16 ms | 0.09 ms | 0.82 | 1.00 | 758.3 KB | 0.36 | 138.4 KB | 1.00 | 17.8% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 3.85 ms | 0.06 ms | 0.03 ms | 1.00 | 1.22 | 2081.1 KB | 1.00 | 138.0 KB | 1.00 | Loss +21.6% |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | MiniExcel | 7.35 ms | 0.30 ms | 0.17 ms | 1.91 | 2.32 | 23222.2 KB | 11.16 | 153.7 KB | 1.11 | 91.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | ClosedXML | 32.88 ms | 2.15 ms | 1.24 ms | 8.55 | 10.39 | 22221.3 KB | 10.68 | 120.1 KB | 0.87 | 754.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | EPPlus | 40.39 ms | 1.57 ms | 0.91 ms | 10.50 | 12.77 | 24694.3 KB | 11.87 | 114.1 KB | 0.83 | 949.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | Sylvan.Data.Excel | 3.27 ms | 0.39 ms | 0.22 ms | 0.85 | 1.00 | 758.7 KB | 0.43 | 78.5 KB | 0.57 | 14.9% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | LargeXlsx | 3.79 ms | 0.27 ms | 0.16 ms | 0.99 | 1.16 | 1032.5 KB | 0.59 | 138.4 KB | 1.00 | Tie vs OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | OfficeIMO.Excel | 3.84 ms | 0.09 ms | 0.05 ms | 1.00 | 1.18 | 1763.0 KB | 1.00 | 138.0 KB | 1.00 | Loss +17.5% |
| 2500 | package-profile | package | Package size | write-datareader-plain | MiniExcel | 7.28 ms | 0.32 ms | 0.19 ms | 1.89 | 2.23 | 23043.8 KB | 13.07 | 153.6 KB | 1.11 | 89.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | ClosedXML | 26.63 ms | 0.53 ms | 0.30 ms | 6.94 | 8.15 | 11581.0 KB | 6.57 | 120.1 KB | 0.87 | 593.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | EPPlus | 37.52 ms | 1.63 ms | 0.94 ms | 9.77 | 11.48 | 16646.8 KB | 9.44 | 114.9 KB | 0.83 | 877.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table | OfficeIMO.Excel | 4.09 ms | 0.21 ms | 0.12 ms | 1.00 | 1.00 | 1774.9 KB | 1.00 | 138.8 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-datareader-table | MiniExcel | 6.86 ms | 0.27 ms | 0.15 ms | 1.68 | 1.68 | 23044.1 KB | 12.98 | 153.6 KB | 1.11 | 67.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table | EPPlus | 35.40 ms | 0.90 ms | 0.52 ms | 8.65 | 8.65 | 16646.5 KB | 9.38 | 114.9 KB | 0.83 | 765.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table | ClosedXML | 37.70 ms | 6.74 ms | 3.89 ms | 9.21 | 9.21 | 19007.9 KB | 10.71 | 120.9 KB | 0.87 | 821.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | OfficeIMO.Excel | 4.11 ms | 0.05 ms | 0.03 ms | 1.00 | 1.00 | 1781.2 KB | 1.00 | 138.9 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | MiniExcel | 7.22 ms | 0.23 ms | 0.13 ms | 1.76 | 1.76 | 26647.2 KB | 14.96 | 153.8 KB | 1.11 | 75.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | EPPlus | 54.54 ms | 4.25 ms | 2.45 ms | 13.27 | 13.27 | 38344.0 KB | 21.53 | 115.1 KB | 0.83 | 1227.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | ClosedXML | 72.64 ms | 2.70 ms | 1.56 ms | 17.68 | 17.68 | 58361.4 KB | 32.77 | 121.0 KB | 0.87 | 1668.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | LargeXlsx | 3.99 ms | 0.11 ms | 0.06 ms | 0.96 | 1.00 | 1123.9 KB | 0.53 | 164.2 KB | 1.25 | 3.9% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | OfficeIMO.Excel | 4.16 ms | 0.55 ms | 0.32 ms | 1.00 | 1.04 | 2140.6 KB | 1.00 | 131.1 KB | 1.00 | Loss +4.1% |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | MiniExcel | 9.88 ms | 1.12 ms | 0.65 ms | 2.38 | 2.47 | 29746.9 KB | 13.90 | 180.5 KB | 1.38 | 137.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | ClosedXML | 52.16 ms | 1.80 ms | 1.04 ms | 12.55 | 13.06 | 27410.3 KB | 12.80 | 159.4 KB | 1.22 | 1154.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | EPPlus | 58.52 ms | 6.49 ms | 3.75 ms | 14.08 | 14.65 | 21890.1 KB | 10.23 | 144.5 KB | 1.10 | 1307.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | OfficeIMO.Excel | 4.48 ms | 0.09 ms | 0.05 ms | 1.00 | 1.00 | 2880.2 KB | 1.00 | 176.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | MiniExcel | 9.47 ms | 0.30 ms | 0.18 ms | 2.11 | 2.11 | 29746.9 KB | 10.33 | 180.5 KB | 1.03 | 111.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | ClosedXML | 53.16 ms | 2.35 ms | 1.36 ms | 11.86 | 11.86 | 27409.3 KB | 9.52 | 159.4 KB | 0.91 | 1086.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | EPPlus | 58.99 ms | 0.96 ms | 0.55 ms | 13.16 | 13.16 | 21890.1 KB | 7.60 | 144.5 KB | 0.82 | 1216.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables | OfficeIMO.Excel | 4.01 ms | 0.52 ms | 0.30 ms | 1.00 | 1.00 | 2066.1 KB | 1.00 | 138.9 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-dataset-tables | MiniExcel | 8.23 ms | 0.70 ms | 0.40 ms | 2.05 | 2.05 | 28700.4 KB | 13.89 | 156.4 KB | 1.13 | 104.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables | ClosedXML | 34.94 ms | 1.42 ms | 0.82 ms | 8.70 | 8.70 | 18876.9 KB | 9.14 | 123.4 KB | 0.89 | 770.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables | EPPlus | 39.68 ms | 7.08 ms | 4.09 ms | 9.88 | 9.88 | 18701.1 KB | 9.05 | 116.6 KB | 0.84 | 888.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | OfficeIMO.Excel | 4.69 ms | 1.06 ms | 0.61 ms | 1.00 | 1.00 | 2078.7 KB | 1.00 | 139.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | MiniExcel | 7.99 ms | 0.09 ms | 0.05 ms | 1.71 | 1.71 | 31798.5 KB | 15.30 | 156.6 KB | 1.13 | 70.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | EPPlus | 61.04 ms | 2.45 ms | 1.42 ms | 13.03 | 13.03 | 41456.2 KB | 19.94 | 116.9 KB | 0.84 | 1202.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | ClosedXML | 73.00 ms | 1.60 ms | 0.92 ms | 15.58 | 15.58 | 56708.2 KB | 27.28 | 123.7 KB | 0.89 | 1457.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-direct | LargeXlsx | 3.44 ms | 0.06 ms | 0.04 ms | 0.92 | 1.00 | 1149.0 KB | 0.66 | 138.4 KB | 1.00 | 7.8% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-direct | OfficeIMO.Excel | 3.73 ms | 0.11 ms | 0.06 ms | 1.00 | 1.09 | 1748.6 KB | 1.00 | 138.0 KB | 1.00 | Loss +8.5% |
| 2500 | package-profile | package | Package size | write-datatable-direct | MiniExcel | 7.56 ms | 0.13 ms | 0.08 ms | 2.03 | 2.20 | 23062.5 KB | 13.19 | 153.7 KB | 1.11 | 102.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-direct | ClosedXML | 27.52 ms | 0.50 ms | 0.29 ms | 7.37 | 8.00 | 11581.0 KB | 6.62 | 120.1 KB | 0.87 | 637.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-direct | EPPlus | 38.92 ms | 6.02 ms | 3.48 ms | 10.43 | 11.32 | 16646.5 KB | 9.52 | 114.9 KB | 0.83 | 942.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | OfficeIMO.Excel | 3.76 ms | 0.09 ms | 0.05 ms | 1.00 | 1.00 | 1760.7 KB | 1.00 | 138.8 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | MiniExcel | 7.53 ms | 0.36 ms | 0.21 ms | 2.01 | 2.01 | 23062.8 KB | 13.10 | 153.7 KB | 1.11 | 100.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | EPPlus | 37.17 ms | 2.48 ms | 1.43 ms | 9.89 | 9.89 | 16646.5 KB | 9.45 | 114.9 KB | 0.83 | 889.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | ClosedXML | 38.16 ms | 5.00 ms | 2.88 ms | 10.16 | 10.16 | 19008.3 KB | 10.80 | 120.9 KB | 0.87 | 915.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | LargeXlsx | 3.03 ms | 0.09 ms | 0.05 ms | 0.74 | 1.00 | 758.3 KB | 0.43 | 138.4 KB | 1.00 | 26.2% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 4.10 ms | 0.50 ms | 0.29 ms | 1.00 | 1.35 | 1769.2 KB | 1.00 | 138.0 KB | 1.00 | Loss +35.4% |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | MiniExcel | 7.48 ms | 0.73 ms | 0.42 ms | 1.83 | 2.47 | 23222.2 KB | 13.13 | 153.7 KB | 1.11 | 82.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | ClosedXML | 30.08 ms | 4.48 ms | 2.59 ms | 7.34 | 9.94 | 11581.0 KB | 6.55 | 120.1 KB | 0.87 | 633.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | EPPlus | 37.28 ms | 2.43 ms | 1.40 ms | 9.09 | 12.32 | 16646.8 KB | 9.41 | 114.9 KB | 0.83 | 809.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.14 ms | 0.25 ms | 0.14 ms | 1.00 | 1.00 | 1347.1 KB | 1.00 | 142.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 54.28 ms | 0.87 ms | 0.50 ms | 13.10 | 13.10 | 38344.3 KB | 28.46 | 115.1 KB | 0.81 | 1209.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 66.16 ms | 1.35 ms | 0.78 ms | 15.97 | 15.97 | 50927.5 KB | 37.80 | 120.2 KB | 0.84 | 1496.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | LargeXlsx | 3.06 ms | 0.12 ms | 0.07 ms | 0.87 | 1.00 | 758.3 KB | 0.57 | 138.4 KB | 0.97 | 13.5% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | OfficeIMO.Excel | 3.53 ms | 0.03 ms | 0.02 ms | 1.00 | 1.16 | 1339.3 KB | 1.00 | 142.3 KB | 1.00 | Loss +15.6% |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | MiniExcel | 7.53 ms | 0.21 ms | 0.12 ms | 2.13 | 2.46 | 23222.2 KB | 17.34 | 153.7 KB | 1.08 | 113.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | ClosedXML | 29.11 ms | 1.58 ms | 0.91 ms | 8.24 | 9.52 | 11581.0 KB | 8.65 | 120.1 KB | 0.84 | 723.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | EPPlus | 38.49 ms | 0.66 ms | 0.38 ms | 10.89 | 12.59 | 16646.5 KB | 12.43 | 114.9 KB | 0.81 | 989.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 5.25 ms | 0.99 ms | 0.57 ms | 1.00 | 1.00 | 1505.3 KB | 1.00 | 138.1 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 56.38 ms | 2.01 ms | 1.16 ms | 10.74 | 10.74 | 38344.3 KB | 25.47 | 115.1 KB | 0.83 | 974.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 65.83 ms | 3.92 ms | 2.27 ms | 12.54 | 12.54 | 50927.5 KB | 33.83 | 120.2 KB | 0.87 | 1154.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 3.23 ms | 0.16 ms | 0.09 ms | 0.79 | 1.00 | 758.3 KB | 0.51 | 138.4 KB | 1.00 | 20.9% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 4.09 ms | 0.05 ms | 0.03 ms | 1.00 | 1.26 | 1497.5 KB | 1.00 | 138.0 KB | 1.00 | Loss +26.5% |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | MiniExcel | 7.89 ms | 0.50 ms | 0.29 ms | 1.93 | 2.44 | 23222.2 KB | 15.51 | 153.7 KB | 1.11 | 93.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | ClosedXML | 28.56 ms | 1.05 ms | 0.61 ms | 6.98 | 8.83 | 11581.0 KB | 7.73 | 120.1 KB | 0.87 | 598.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | EPPlus | 38.07 ms | 3.20 ms | 1.85 ms | 9.31 | 11.77 | 16646.5 KB | 11.12 | 114.9 KB | 0.83 | 830.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | LargeXlsx | 3.15 ms | 0.35 ms | 0.20 ms | 0.67 | 1.00 | 758.3 KB | 0.43 | 138.4 KB | 0.97 | 33.2% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | OfficeIMO.Excel | 4.72 ms | 0.24 ms | 0.14 ms | 1.00 | 1.50 | 1770.1 KB | 1.00 | 142.3 KB | 1.00 | Loss +49.7% |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | MiniExcel | 7.45 ms | 0.02 ms | 0.01 ms | 1.58 | 2.36 | 23222.2 KB | 13.12 | 153.7 KB | 1.08 | 57.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | ClosedXML | 30.94 ms | 3.49 ms | 2.02 ms | 6.56 | 9.81 | 11581.0 KB | 6.54 | 120.1 KB | 0.84 | 555.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | EPPlus | 36.48 ms | 1.15 ms | 0.66 ms | 7.73 | 11.57 | 16646.5 KB | 9.40 | 114.9 KB | 0.81 | 672.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 3.71 ms | 0.12 ms | 0.07 ms | 1.00 | 1.00 | 1346.4 KB | 1.00 | 142.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 45.10 ms | 2.57 ms | 1.48 ms | 12.17 | 12.17 | 28540.6 KB | 21.20 | 120.2 KB | 0.84 | 1116.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 47.16 ms | 2.40 ms | 1.38 ms | 12.72 | 12.72 | 27306.2 KB | 20.28 | 115.0 KB | 0.81 | 1172.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | LargeXlsx | 4.03 ms | 0.07 ms | 0.04 ms | 0.73 | 1.00 | 802.5 KB | 0.34 | 182.6 KB | 1.00 | 27.2% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 5.53 ms | 0.45 ms | 0.26 ms | 1.00 | 1.37 | 2341.7 KB | 1.00 | 183.1 KB | 1.00 | Loss +37.4% |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | MiniExcel | 8.49 ms | 0.35 ms | 0.20 ms | 1.53 | 2.11 | 25190.5 KB | 10.76 | 194.0 KB | 1.06 | 53.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | ClosedXML | 39.62 ms | 3.47 ms | 2.01 ms | 7.16 | 9.84 | 16973.5 KB | 7.25 | 161.0 KB | 0.88 | 616.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | EPPlus | 51.41 ms | 3.20 ms | 1.85 ms | 9.29 | 12.77 | 20105.6 KB | 8.59 | 152.1 KB | 0.83 | 829.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | LargeXlsx | 4.03 ms | 0.05 ms | 0.03 ms | 0.91 | 1.00 | 802.5 KB | 0.53 | 182.6 KB | 1.00 | 9.0% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 4.42 ms | 0.08 ms | 0.05 ms | 1.00 | 1.10 | 1507.7 KB | 1.00 | 182.4 KB | 1.00 | Loss +9.9% |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | MiniExcel | 8.14 ms | 0.08 ms | 0.05 ms | 1.84 | 2.02 | 25190.5 KB | 16.71 | 194.0 KB | 1.06 | 84.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | ClosedXML | 38.42 ms | 4.56 ms | 2.63 ms | 8.68 | 9.54 | 16973.5 KB | 11.26 | 161.0 KB | 0.88 | 768.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | EPPlus | 47.14 ms | 0.60 ms | 0.34 ms | 10.65 | 11.71 | 20105.6 KB | 13.33 | 152.1 KB | 0.83 | 965.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 18.33 ms | 0.19 ms | 0.11 ms | 1.00 | 1.00 | 4502.3 KB | 1.00 | 651.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | LargeXlsx | 20.07 ms | 0.65 ms | 0.38 ms | 1.10 | 1.10 | 2810.7 KB | 0.62 | 644.6 KB | 0.99 | 9.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | MiniExcel | 35.39 ms | 3.38 ms | 1.95 ms | 1.93 | 1.93 | 48414.8 KB | 10.75 | 674.4 KB | 1.04 | 93.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | ClosedXML | 119.38 ms | 3.77 ms | 2.18 ms | 6.51 | 6.51 | 51647.0 KB | 11.47 | 615.5 KB | 0.95 | 551.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | EPPlus | 154.02 ms | 1.83 ms | 1.06 ms | 8.40 | 8.40 | 69140.0 KB | 15.36 | 548.9 KB | 0.84 | 740.5% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | OfficeIMO.Excel | 8.19 ms | 0.16 ms | 0.09 ms | 1.00 | 1.00 | 1895.0 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | EPPlus | 73.54 ms | 0.19 ms | 0.11 ms | 8.98 | 8.98 | 50712.1 KB | 26.76 |  |  | 797.9% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | EPPlus 4.5.3.3 | 95.04 ms |  |  | 11.60 | 11.60 |  |  |  |  | 1060.4% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | ClosedXML | 126.85 ms | 2.31 ms | 1.33 ms | 15.49 | 15.49 | 84884.9 KB | 44.79 |  |  | 1448.8% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | OfficeIMO.Excel | 17.25 ms | 0.43 ms | 0.25 ms | 1.00 | 1.00 | 15209.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus 4.5.3.3 | 42.81 ms |  |  | 2.48 | 2.48 |  |  |  |  | 148.2% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus | 49.27 ms | 2.95 ms | 1.71 ms | 2.86 | 2.86 | 32907.5 KB | 2.16 |  |  | 185.6% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | OfficeIMO.Excel | 10.76 ms | 0.15 ms | 0.09 ms | 1.00 | 1.00 | 6195.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus | 73.65 ms | 3.15 ms | 1.82 ms | 6.84 | 6.84 | 54594.9 KB | 8.81 |  |  | 584.1% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus 4.5.3.3 | 74.45 ms |  |  | 6.92 | 6.92 |  |  |  |  | 591.6% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | OfficeIMO.Excel | 17.92 ms | 0.48 ms | 0.28 ms | 1.00 | 1.00 | 16350.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus | 73.98 ms | 2.01 ms | 1.16 ms | 4.13 | 4.13 | 59227.2 KB | 3.62 |  |  | 312.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus 4.5.3.3 | 77.86 ms |  |  | 4.34 | 4.34 |  |  |  |  | 334.4% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | OfficeIMO.Excel | 16.05 ms | 0.43 ms | 0.25 ms | 1.00 | 1.00 | 15230.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus 4.5.3.3 | 61.57 ms |  |  | 3.83 | 3.83 |  |  |  |  | 283.5% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus | 68.85 ms | 1.48 ms | 0.86 ms | 4.29 | 4.29 | 54595.6 KB | 3.58 |  |  | 328.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | OfficeIMO.Excel | 18.01 ms | 1.62 ms | 0.94 ms | 1.00 | 1.00 | 15225.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus 4.5.3.3 | 72.43 ms |  |  | 4.02 | 4.02 |  |  |  |  | 302.1% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus | 72.89 ms | 5.01 ms | 2.89 ms | 4.05 | 4.05 | 54592.2 KB | 3.59 |  |  | 304.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Object projection | build-object-datatable-dictionaries | OfficeIMO.Excel | 1.36 ms | 0.20 ms | 0.12 ms | 1.00 | 1.00 | 564.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Object projection | build-object-datatable-typed | OfficeIMO.Excel | 1.09 ms | 0.02 ms | 0.01 ms | 1.00 | 1.00 | 856.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-cells | OfficeIMO.Excel | 5.42 ms | 0.59 ms | 0.34 ms | 1.00 | 1.00 | 2531.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-cells | EPPlus | 31.13 ms | 4.69 ms | 2.71 ms | 5.75 | 5.75 | 20155.0 KB | 7.96 |  |  | 474.8% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-cells | ClosedXML | 31.24 ms | 1.61 ms | 0.93 ms | 5.77 | 5.77 | 17022.2 KB | 6.72 |  |  | 476.8% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | OfficeIMO.Excel | 3.62 ms | 0.13 ms | 0.07 ms | 1.00 | 1.00 | 523.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | EPPlus | 23.93 ms | 0.88 ms | 0.51 ms | 6.61 | 6.61 | 13108.2 KB | 25.05 |  |  | 561.5% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | ClosedXML | 29.69 ms | 0.56 ms | 0.32 ms | 8.21 | 8.21 | 15463.4 KB | 29.55 |  |  | 720.6% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-range | OfficeIMO.Excel | 5.88 ms | 0.81 ms | 0.47 ms | 1.00 | 1.00 | 2531.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-range | EPPlus | 29.07 ms | 5.11 ms | 2.95 ms | 4.94 | 4.94 | 20155.0 KB | 7.96 |  |  | 394.5% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-range | ClosedXML | 31.89 ms | 4.87 ms | 2.81 ms | 5.42 | 5.42 | 17020.7 KB | 6.72 |  |  | 442.5% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-top-range | OfficeIMO.Excel | 0.78 ms | 0.34 ms | 0.20 ms | 1.00 | 1.00 | 285.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-top-range | EPPlus | 24.20 ms | 1.65 ms | 0.95 ms | 31.17 | 31.17 | 12404.5 KB | 43.49 |  |  | 3017.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-top-range | ClosedXML | 31.31 ms | 0.39 ms | 0.23 ms | 40.33 | 40.33 | 15370.2 KB | 53.89 |  |  | 3933.4% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-autofilter | OfficeIMO.Excel | 3.62 ms | 0.09 ms | 0.05 ms | 1.00 | 1.00 | 1340.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-autofilter | ClosedXML | 30.89 ms | 1.85 ms | 1.07 ms | 8.53 | 8.53 | 22226.8 KB | 16.58 |  |  | 752.6% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-autofilter | EPPlus 4.5.3.3 | 35.14 ms |  |  | 9.70 | 9.70 |  |  |  |  | 869.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-autofilter | EPPlus | 41.08 ms | 1.08 ms | 0.62 ms | 11.34 | 11.34 | 24715.8 KB | 18.44 |  |  | 1034.0% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-charts | OfficeIMO.Excel | 5.03 ms | 0.18 ms | 0.10 ms | 1.00 | 1.00 | 1892.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-charts | EPPlus 4.5.3.3 | 37.08 ms |  |  | 7.38 | 7.38 |  |  |  |  | 637.6% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-charts | EPPlus | 41.05 ms | 1.15 ms | 0.66 ms | 8.17 | 8.17 | 27142.7 KB | 14.34 |  |  | 716.6% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | OfficeIMO.Excel | 3.59 ms | 0.02 ms | 0.01 ms | 1.00 | 1.00 | 1405.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | ClosedXML | 31.85 ms | 0.97 ms | 0.56 ms | 8.87 | 8.87 | 22273.8 KB | 15.84 |  |  | 786.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | EPPlus 4.5.3.3 | 36.01 ms |  |  | 10.02 | 10.02 |  |  |  |  | 902.2% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | EPPlus | 40.15 ms | 0.06 ms | 0.03 ms | 11.18 | 11.18 | 24757.8 KB | 17.61 |  |  | 1017.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-data-validation | OfficeIMO.Excel | 3.66 ms | 0.04 ms | 0.02 ms | 1.00 | 1.00 | 1356.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-data-validation | ClosedXML | 30.28 ms | 0.17 ms | 0.10 ms | 8.28 | 8.28 | 22247.9 KB | 16.41 |  |  | 727.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-data-validation | EPPlus 4.5.3.3 | 34.68 ms |  |  | 9.48 | 9.48 |  |  |  |  | 848.2% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-data-validation | EPPlus | 38.32 ms | 1.03 ms | 0.59 ms | 10.48 | 10.48 | 24701.8 KB | 18.22 |  |  | 947.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | OfficeIMO.Excel | 3.65 ms | 0.13 ms | 0.08 ms | 1.00 | 1.00 | 1342.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | ClosedXML | 32.77 ms | 0.91 ms | 0.52 ms | 8.97 | 8.97 | 22222.0 KB | 16.55 |  |  | 796.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | EPPlus 4.5.3.3 | 34.24 ms |  |  | 9.37 | 9.37 |  |  |  |  | 837.1% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | EPPlus | 42.07 ms | 0.30 ms | 0.18 ms | 11.51 | 11.51 | 24730.4 KB | 18.42 |  |  | 1051.4% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | OfficeIMO.Excel | 14.05 ms | 0.85 ms | 0.49 ms | 1.00 | 1.00 | 14419.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | EPPlus 4.5.3.3 | 35.46 ms |  |  | 2.52 | 2.52 |  |  |  |  | 152.5% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | EPPlus | 44.12 ms | 2.39 ms | 1.38 ms | 3.14 | 3.14 | 29538.0 KB | 2.05 |  |  | 214.1% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world report | realworld-report-all-in-one | OfficeIMO.Excel | 16.82 ms | 1.87 ms | 1.08 ms | 1.00 | 1.00 | 15220.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus | 72.15 ms | 4.12 ms | 2.38 ms | 4.29 | 4.29 | 54595.5 KB | 3.59 |  |  | 329.0% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus 4.5.3.3 | 76.93 ms |  |  | 4.57 | 4.57 |  |  |  |  | 357.5% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world report | realworld-report-core | OfficeIMO.Excel | 4.27 ms | 0.06 ms | 0.04 ms | 1.00 | 1.00 | 1488.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world report | realworld-report-core | EPPlus 4.5.3.3 | 60.12 ms |  |  | 14.08 | 14.08 |  |  |  |  | 1308.2% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world report | realworld-report-core | EPPlus | 65.98 ms | 3.37 ms | 1.95 ms | 15.46 | 15.46 | 47300.2 KB | 31.77 |  |  | 1445.5% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world report | realworld-report-core | ClosedXML | 78.52 ms | 4.70 ms | 2.72 ms | 18.39 | 18.39 | 69834.2 KB | 46.91 |  |  | 1739.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook | OfficeIMO.Excel | 29.93 ms | 4.84 ms | 2.80 ms | 1.00 | 1.00 | 19069.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Report workbook | report-workbook | EPPlus 4.5.3.3 | 77.70 ms |  |  | 2.60 | 2.60 |  |  |  |  | 159.6% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook | EPPlus | 94.07 ms | 2.81 ms | 1.62 ms | 3.14 | 3.14 | 77486.6 KB | 4.06 |  |  | 214.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-core | OfficeIMO.Excel | 6.41 ms | 0.39 ms | 0.22 ms | 1.00 | 1.00 | 2711.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Report workbook | report-workbook-core | EPPlus 4.5.3.3 | 70.56 ms |  |  | 11.01 | 11.01 |  |  |  |  | 1000.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-core | EPPlus | 97.70 ms | 1.24 ms | 0.72 ms | 15.24 | 15.24 | 71970.9 KB | 26.55 |  |  | 1423.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-core | ClosedXML | 103.64 ms | 1.49 ms | 0.86 ms | 16.17 | 16.17 | 97220.1 KB | 35.86 |  |  | 1516.6% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable | OfficeIMO.Excel | 23.96 ms | 1.77 ms | 1.02 ms | 1.00 | 1.00 | 19383.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable | EPPlus 4.5.3.3 | 67.32 ms |  |  | 2.81 | 2.81 |  |  |  |  | 180.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable | EPPlus | 102.61 ms | 2.82 ms | 1.63 ms | 4.28 | 4.28 | 65995.8 KB | 3.40 |  |  | 328.2% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable-core | OfficeIMO.Excel | 6.28 ms | 0.47 ms | 0.27 ms | 1.00 | 1.00 | 2982.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable-core | EPPlus 4.5.3.3 | 63.60 ms |  |  | 10.13 | 10.13 |  |  |  |  | 913.4% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable-core | EPPlus | 97.86 ms | 4.55 ms | 2.62 ms | 15.59 | 15.59 | 60480.4 KB | 20.28 |  |  | 1459.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable-core | ClosedXML | 102.13 ms | 1.46 ms | 0.84 ms | 16.27 | 16.27 | 82858.9 KB | 27.78 |  |  | 1527.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | OfficeIMO.Excel | 2.14 ms | 0.05 ms | 0.03 ms | 1.00 | 1.00 | 706.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | EPPlus 4.5.3.3 | 13.21 ms |  |  | 6.16 | 6.16 |  |  |  |  | 516.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | ClosedXML | 17.34 ms | 0.58 ms | 0.34 ms | 8.09 | 8.09 | 8279.5 KB | 11.72 |  |  | 708.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | EPPlus | 19.35 ms | 0.98 ms | 0.56 ms | 9.03 | 9.03 | 7708.1 KB | 10.91 |  |  | 802.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | OfficeIMO.Excel | 0.81 ms | 0.01 ms | 0.01 ms | 1.00 | 1.00 | 177.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | Sylvan.Data.Excel | 0.95 ms | 0.03 ms | 0.02 ms | 1.17 | 1.17 | 316.6 KB | 1.79 |  |  | 17.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | ExcelDataReader | 1.45 ms | 0.02 ms | 0.01 ms | 1.79 | 1.79 | 4062.2 KB | 22.92 |  |  | 78.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | ClosedXML | 3.98 ms | 0.97 ms | 0.56 ms | 4.92 | 4.92 | 4392.7 KB | 24.79 |  |  | 391.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | EPPlus 4.5.3.3 | 11.00 ms |  |  | 13.59 | 13.59 |  |  |  |  | 1258.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | MiniExcel | 12.08 ms | 0.36 ms | 0.21 ms | 14.92 | 14.92 | 46194.9 KB | 260.67 |  |  | 1391.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | EPPlus | 38.48 ms | 2.68 ms | 1.55 ms | 47.50 | 47.50 | 43071.1 KB | 243.04 |  |  | 4650.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | Sylvan.Data.Excel | 0.96 ms | 0.03 ms | 0.02 ms | 0.62 | 1.00 | 316.6 KB | 1.79 |  |  | 38.2% faster than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | ExcelDataReader | 1.44 ms | 0.01 ms | 0.00 ms | 0.92 | 1.49 | 4062.2 KB | 22.93 |  |  | 7.7% faster than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | OfficeIMO.Excel | 1.56 ms | 0.57 ms | 0.33 ms | 1.00 | 1.62 | 177.2 KB | 1.00 |  |  | Loss +61.8% |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | ClosedXML | 4.10 ms | 1.19 ms | 0.69 ms | 2.63 | 4.25 | 4392.7 KB | 24.79 |  |  | 162.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | EPPlus 4.5.3.3 | 10.65 ms |  |  | 6.83 | 11.04 |  |  |  |  | 582.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | MiniExcel | 12.98 ms | 1.25 ms | 0.72 ms | 8.32 | 13.47 | 46194.9 KB | 260.72 |  |  | 732.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | EPPlus | 38.29 ms | 1.56 ms | 0.90 ms | 24.55 | 39.71 | 43071.1 KB | 243.09 |  |  | 2354.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | Sylvan.Data.Excel | 2.04 ms | 0.03 ms | 0.02 ms | 0.58 | 1.00 | 518.6 KB | 0.49 |  |  | 42.5% faster than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | OfficeIMO.Excel | 3.56 ms | 0.31 ms | 0.18 ms | 1.00 | 1.74 | 1056.5 KB | 1.00 |  |  | Loss +73.9% |
| 2500 | speed-comparison | read | Other | shared-string-read | ExcelDataReader | 4.69 ms | 0.18 ms | 0.10 ms | 1.32 | 2.29 | 2619.1 KB | 2.48 |  |  | 31.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | MiniExcel | 5.45 ms | 0.45 ms | 0.26 ms | 1.53 | 2.66 | 7530.1 KB | 7.13 |  |  | 53.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | EPPlus 4.5.3.3 | 11.88 ms |  |  | 3.34 | 5.81 |  |  |  |  | 234.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | ClosedXML | 15.84 ms | 0.54 ms | 0.31 ms | 4.45 | 7.75 | 9498.1 KB | 8.99 |  |  | 345.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | EPPlus | 20.07 ms | 1.01 ms | 0.58 ms | 5.65 | 9.82 | 10372.3 KB | 9.82 |  |  | 464.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | OfficeIMO.Excel | 4.52 ms | 1.23 ms | 0.71 ms | 1.00 | 1.00 | 374.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | Sylvan.Data.Excel | 4.70 ms | 1.19 ms | 0.69 ms | 1.04 | 1.04 | 655.2 KB | 1.75 |  |  | 4.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | ExcelDataReader | 10.34 ms | 0.58 ms | 0.34 ms | 2.29 | 2.29 | 6089.3 KB | 16.26 |  |  | 128.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | MiniExcel | 15.51 ms | 1.76 ms | 1.02 ms | 3.43 | 3.43 | 18661.8 KB | 49.83 |  |  | 243.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | EPPlus | 27.26 ms | 3.35 ms | 1.94 ms | 6.03 | 6.03 | 12427.1 KB | 33.18 |  |  | 503.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | ClosedXML | 36.83 ms | 6.70 ms | 3.87 ms | 8.15 | 8.15 | 15361.2 KB | 41.02 |  |  | 715.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | Sylvan.Data.Excel | 6.17 ms | 0.43 ms | 0.25 ms | 0.76 | 1.00 | 2239.3 KB | 0.62 |  |  | 23.8% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | OfficeIMO.Excel | 8.09 ms | 2.34 ms | 1.35 ms | 1.00 | 1.31 | 3594.4 KB | 1.00 |  |  | Loss +31.2% |
| 2500 | speed-comparison | read | Range and table read | read-datatable | ExcelDataReader | 13.36 ms | 1.27 ms | 0.74 ms | 1.65 | 2.17 | 7673.3 KB | 2.13 |  |  | 65.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | MiniExcel | 15.85 ms | 1.37 ms | 0.79 ms | 1.96 | 2.57 | 18266.6 KB | 5.08 |  |  | 95.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | EPPlus | 35.13 ms | 7.06 ms | 4.08 ms | 4.34 | 5.70 | 21736.6 KB | 6.05 |  |  | 334.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | EPPlus 4.5.3.3 | 35.47 ms |  |  | 4.38 | 5.75 |  |  |  |  | 338.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | ClosedXML | 38.43 ms | 5.04 ms | 2.91 ms | 4.75 | 6.23 | 18313.9 KB | 5.10 |  |  | 375.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | OfficeIMO.Excel | 3.86 ms | 0.08 ms | 0.04 ms | 1.00 | 1.00 | 542.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | Sylvan.Data.Excel | 4.36 ms | 0.08 ms | 0.05 ms | 1.13 | 1.13 | 733.5 KB | 1.35 |  |  | 13.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | ExcelDataReader | 10.50 ms | 0.37 ms | 0.21 ms | 2.72 | 2.72 | 6089.3 KB | 11.22 |  |  | 172.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | MiniExcel | 10.76 ms | 0.26 ms | 0.15 ms | 2.79 | 2.79 | 15850.3 KB | 29.20 |  |  | 178.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | EPPlus | 23.88 ms | 1.29 ms | 0.74 ms | 6.19 | 6.19 | 13108.2 KB | 24.15 |  |  | 518.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | ClosedXML | 30.10 ms | 1.12 ms | 0.65 ms | 7.80 | 7.80 | 15465.0 KB | 28.49 |  |  | 680.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | OfficeIMO.Excel | 9.03 ms | 0.31 ms | 0.18 ms | 1.00 | 1.00 | 2692.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Range and table read | read-range | Sylvan.Data.Excel | 15.52 ms | 2.67 ms | 1.54 ms | 1.72 | 1.72 | 655.0 KB | 0.24 |  |  | 71.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | ExcelDataReader | 23.31 ms | 5.50 ms | 3.18 ms | 2.58 | 2.58 | 6089.2 KB | 2.26 |  |  | 158.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | MiniExcel | 23.87 ms | 11.69 ms | 6.75 ms | 2.64 | 2.64 | 18662.1 KB | 6.93 |  |  | 164.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | EPPlus 4.5.3.3 | 35.79 ms |  |  | 3.96 | 3.96 |  |  |  |  | 296.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | EPPlus | 37.27 ms | 5.09 ms | 2.94 ms | 4.13 | 4.13 | 20152.6 KB | 7.48 |  |  | 312.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | ClosedXML | 78.00 ms | 26.95 ms | 15.56 ms | 8.64 | 8.64 | 16846.0 KB | 6.26 |  |  | 763.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | OfficeIMO.Excel | 4.93 ms | 0.20 ms | 0.11 ms | 1.00 | 1.00 | 2751.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | Sylvan.Data.Excel | 5.02 ms | 0.07 ms | 0.04 ms | 1.02 | 1.02 | 750.3 KB | 0.27 |  |  | Tie vs OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | ExcelDataReader | 10.30 ms | 0.18 ms | 0.11 ms | 2.09 | 2.09 | 6089.3 KB | 2.21 |  |  | 108.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | MiniExcel | 12.48 ms | 0.64 ms | 0.37 ms | 2.53 | 2.53 | 18662.4 KB | 6.78 |  |  | 153.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | EPPlus | 29.09 ms | 3.10 ms | 1.79 ms | 5.90 | 5.90 | 20152.7 KB | 7.32 |  |  | 490.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | ClosedXML | 31.46 ms | 1.23 ms | 0.71 ms | 6.38 | 6.38 | 16728.3 KB | 6.08 |  |  | 538.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | Sylvan.Data.Excel | 0.44 ms | 0.04 ms | 0.02 ms | 0.82 | 1.00 | 348.4 KB | 1.18 |  |  | 17.7% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | OfficeIMO.Excel | 0.53 ms | 0.01 ms | 0.01 ms | 1.00 | 1.22 | 296.0 KB | 1.00 |  |  | Loss +21.6% |
| 2500 | speed-comparison | read | Range and table read | read-top-range | MiniExcel | 0.76 ms | 0.04 ms | 0.03 ms | 1.43 | 1.74 | 869.0 KB | 2.94 |  |  | 42.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | ExcelDataReader | 4.28 ms | 0.09 ms | 0.05 ms | 8.08 | 9.82 | 1931.6 KB | 6.53 |  |  | 707.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | EPPlus | 23.97 ms | 1.23 ms | 0.71 ms | 45.27 | 55.03 | 12402.1 KB | 41.90 |  |  | 4426.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | EPPlus 4.5.3.3 | 28.96 ms |  |  | 54.71 | 66.50 |  |  |  |  | 5370.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | ClosedXML | 31.47 ms | 1.31 ms | 0.76 ms | 59.45 | 72.26 | 15360.2 KB | 51.89 |  |  | 5844.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | Sylvan.Data.Excel | 4.52 ms | 0.15 ms | 0.09 ms | 0.32 | 1.00 | 655.2 KB | 0.19 |  |  | 68.4% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | ExcelDataReader | 10.88 ms | 0.59 ms | 0.34 ms | 0.76 | 2.41 | 6089.4 KB | 1.75 |  |  | 24.0% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | MiniExcel | 11.98 ms | 0.59 ms | 0.34 ms | 0.84 | 2.65 | 18662.4 KB | 5.37 |  |  | 16.3% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | OfficeIMO.Excel | 14.31 ms | 3.04 ms | 1.76 ms | 1.00 | 3.17 | 3472.7 KB | 1.00 |  |  | Loss +216.7% |
| 2500 | speed-comparison | read | Range and table read | read-used-range | EPPlus | 29.62 ms | 3.01 ms | 1.74 ms | 2.07 | 6.55 | 20152.8 KB | 5.80 |  |  | 107.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | ClosedXML | 67.87 ms | 33.33 ms | 19.24 ms | 4.74 | 15.02 | 16806.5 KB | 4.84 |  |  | 374.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | OfficeIMO.Excel | 3.99 ms | 0.59 ms | 0.34 ms | 1.00 | 1.00 | 377.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | Sylvan.Data.Excel | 4.04 ms | 0.33 ms | 0.19 ms | 1.01 | 1.01 | 655.2 KB | 1.73 |  |  | Tie vs OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | ExcelDataReader | 10.09 ms | 0.25 ms | 0.15 ms | 2.53 | 2.53 | 6089.5 KB | 16.12 |  |  | 152.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | MiniExcel | 13.68 ms | 2.04 ms | 1.18 ms | 3.43 | 3.43 | 18661.8 KB | 49.41 |  |  | 242.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | EPPlus | 23.69 ms | 0.75 ms | 0.43 ms | 5.94 | 5.94 | 12427.1 KB | 32.90 |  |  | 493.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | ClosedXML | 30.60 ms | 2.10 ms | 1.21 ms | 7.67 | 7.67 | 15359.4 KB | 40.66 |  |  | 666.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | Sylvan.Data.Excel | 4.40 ms | 0.07 ms | 0.04 ms | 0.80 | 1.00 | 655.2 KB | 0.24 |  |  | 20.3% faster than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | OfficeIMO.Excel | 5.52 ms | 1.61 ms | 0.93 ms | 1.00 | 1.25 | 2771.4 KB | 1.00 |  |  | Loss +25.4% |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | ExcelDataReader | 11.03 ms | 0.67 ms | 0.39 ms | 2.00 | 2.51 | 6089.4 KB | 2.20 |  |  | 99.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | MiniExcel | 12.18 ms | 0.93 ms | 0.54 ms | 2.20 | 2.76 | 18662.4 KB | 6.73 |  |  | 120.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | ClosedXML | 29.61 ms | 0.60 ms | 0.35 ms | 5.36 | 6.72 | 16729.4 KB | 6.04 |  |  | 435.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | EPPlus 4.5.3.3 | 29.77 ms |  |  | 5.39 | 6.76 |  |  |  |  | 438.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | EPPlus | 31.45 ms | 4.24 ms | 2.45 ms | 5.69 | 7.14 | 20152.6 KB | 7.27 |  |  | 469.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | Sylvan.Data.Excel | 0.43 ms | 0.03 ms | 0.01 ms | 0.79 | 1.00 | 348.5 KB | 1.16 |  |  | 21.2% faster than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | OfficeIMO.Excel | 0.54 ms | 0.05 ms | 0.03 ms | 1.00 | 1.27 | 299.4 KB | 1.00 |  |  | Loss +26.9% |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | MiniExcel | 0.79 ms | 0.12 ms | 0.07 ms | 1.45 | 1.84 | 869.0 KB | 2.90 |  |  | 44.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | ExcelDataReader | 4.36 ms | 0.09 ms | 0.05 ms | 8.02 | 10.18 | 1931.8 KB | 6.45 |  |  | 702.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | EPPlus | 23.88 ms | 0.51 ms | 0.29 ms | 43.98 | 55.83 | 12402.1 KB | 41.43 |  |  | 4298.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | EPPlus 4.5.3.3 | 25.78 ms |  |  | 47.48 | 60.26 |  |  |  |  | 4647.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | ClosedXML | 31.10 ms | 2.39 ms | 1.38 ms | 57.28 | 72.70 | 15360.9 KB | 51.31 |  |  | 5628.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | Sylvan.Data.Excel | 0.40 ms | 0.00 ms | 0.00 ms | 0.77 | 1.00 | 348.5 KB | 1.16 |  |  | 22.6% faster than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | OfficeIMO.Excel | 0.52 ms | 0.01 ms | 0.01 ms | 1.00 | 1.29 | 300.0 KB | 1.00 |  |  | Loss +29.2% |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | MiniExcel | 0.97 ms | 0.27 ms | 0.16 ms | 1.87 | 2.41 | 869.0 KB | 2.90 |  |  | 86.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | ExcelDataReader | 4.34 ms | 0.24 ms | 0.14 ms | 8.32 | 10.75 | 1931.8 KB | 6.44 |  |  | 732.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | EPPlus | 23.53 ms | 0.44 ms | 0.26 ms | 45.08 | 58.24 | 12402.1 KB | 41.34 |  |  | 4408.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | ClosedXML | 29.79 ms | 1.04 ms | 0.60 ms | 57.08 | 73.74 | 15360.5 KB | 51.20 |  |  | 5608.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | OfficeIMO.Excel | 7.01 ms | 1.98 ms | 1.14 ms | 1.00 | 1.00 | 2441.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Typed object read | read-objects | Sylvan.Data.Excel | 7.16 ms | 3.21 ms | 1.86 ms | 1.02 | 1.02 | 895.3 KB | 0.37 |  |  | 2.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | ExcelDataReader | 13.68 ms | 4.19 ms | 2.42 ms | 1.95 | 1.95 | 6329.5 KB | 2.59 |  |  | 95.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | MiniExcel | 14.31 ms | 1.21 ms | 0.70 ms | 2.04 | 2.04 | 18474.2 KB | 7.57 |  |  | 104.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | EPPlus 4.5.3.3 | 32.53 ms |  |  | 4.64 | 4.64 |  |  |  |  | 364.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | ClosedXML | 32.62 ms | 0.54 ms | 0.31 ms | 4.65 | 4.65 | 16925.7 KB | 6.93 |  |  | 365.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | EPPlus | 34.50 ms | 5.92 ms | 3.42 ms | 4.92 | 4.92 | 21354.3 KB | 8.74 |  |  | 392.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | OfficeIMO.Excel | 4.72 ms | 0.08 ms | 0.04 ms | 1.00 | 1.00 | 2422.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | Sylvan.Data.Excel | 4.73 ms | 0.34 ms | 0.20 ms | 1.00 | 1.00 | 831.0 KB | 0.34 |  |  | Tie vs OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | ExcelDataReader | 10.48 ms | 0.24 ms | 0.14 ms | 2.22 | 2.22 | 6265.3 KB | 2.59 |  |  | 121.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | MiniExcel | 13.56 ms | 0.86 ms | 0.49 ms | 2.87 | 2.87 | 18410.0 KB | 7.60 |  |  | 187.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | EPPlus 4.5.3.3 | 29.22 ms |  |  | 6.19 | 6.19 |  |  |  |  | 518.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | ClosedXML | 30.23 ms | 0.55 ms | 0.32 ms | 6.40 | 6.40 | 16904.2 KB | 6.98 |  |  | 540.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | EPPlus | 34.75 ms | 0.95 ms | 0.55 ms | 7.36 | 7.36 | 21334.7 KB | 8.81 |  |  | 635.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | OfficeIMO.Excel | 5.10 ms | 0.06 ms | 0.03 ms | 1.00 | 1.00 | 1781.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | MiniExcel | 8.19 ms | 0.12 ms | 0.07 ms | 1.61 | 1.61 | 26647.4 KB | 14.96 |  |  | 60.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | EPPlus | 55.09 ms | 6.09 ms | 3.52 ms | 10.80 | 10.80 | 38345.4 KB | 21.53 |  |  | 980.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | EPPlus 4.5.3.3 | 72.00 ms |  |  | 14.12 | 14.12 |  |  |  |  | 1311.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | ClosedXML | 75.69 ms | 6.62 ms | 3.82 ms | 14.84 | 14.84 | 58360.0 KB | 32.77 |  |  | 1383.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | OfficeIMO.Excel | 5.49 ms | 0.39 ms | 0.22 ms | 1.00 | 1.00 | 2079.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | MiniExcel | 15.64 ms | 1.60 ms | 0.93 ms | 2.85 | 2.85 | 32152.0 KB | 15.46 |  |  | 185.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | EPPlus | 137.37 ms | 4.93 ms | 2.84 ms | 25.04 | 25.04 | 43440.7 KB | 20.89 |  |  | 2403.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | ClosedXML | 190.94 ms | 6.70 ms | 3.87 ms | 34.80 | 34.80 | 56708.1 KB | 27.27 |  |  | 3380.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 3.96 ms | 0.09 ms | 0.05 ms | 1.00 | 1.00 | 1347.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 52.62 ms | 1.45 ms | 0.84 ms | 13.27 | 13.27 | 38344.5 KB | 28.46 |  |  | 1227.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 63.25 ms | 1.85 ms | 1.07 ms | 15.95 | 15.95 | 50927.7 KB | 37.80 |  |  | 1495.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.52 ms | 0.11 ms | 0.07 ms | 1.00 | 1.00 | 1505.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 52.76 ms | 1.33 ms | 0.77 ms | 11.67 | 11.67 | 38344.5 KB | 25.47 |  |  | 1067.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 65.95 ms | 1.45 ms | 0.84 ms | 14.59 | 14.59 | 50927.3 KB | 33.83 |  |  | 1359.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.06 ms | 0.36 ms | 0.21 ms | 1.00 | 1.00 | 1346.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 44.44 ms | 2.19 ms | 1.26 ms | 10.93 | 10.93 | 28540.4 KB | 21.20 |  |  | 993.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 45.60 ms | 2.53 ms | 1.46 ms | 11.22 | 11.22 | 27306.2 KB | 20.28 |  |  | 1021.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | OfficeIMO.Excel | 2.50 ms | 0.20 ms | 0.11 ms | 1.00 | 1.00 | 1787.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | ClosedXML | 12.34 ms | 0.17 ms | 0.10 ms | 4.94 | 4.94 | 9959.5 KB | 5.57 |  |  | 393.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | EPPlus | 23.29 ms | 0.36 ms | 0.21 ms | 9.32 | 9.32 | 11773.2 KB | 6.59 |  |  | 831.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | OfficeIMO.Excel | 2.60 ms | 0.20 ms | 0.11 ms | 1.00 | 1.00 | 1119.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | ClosedXML | 11.62 ms | 0.52 ms | 0.30 ms | 4.46 | 4.46 | 9177.1 KB | 8.19 |  |  | 346.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | EPPlus 4.5.3.3 | 18.22 ms |  |  | 7.00 | 7.00 |  |  |  |  | 599.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | EPPlus | 22.11 ms | 0.68 ms | 0.39 ms | 8.49 | 8.49 | 12895.5 KB | 11.51 |  |  | 749.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | OfficeIMO.Excel | 3.11 ms | 0.13 ms | 0.08 ms | 1.00 | 1.00 | 1763.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | ClosedXML | 17.41 ms | 0.88 ms | 0.51 ms | 5.59 | 5.59 | 11887.0 KB | 6.74 |  |  | 459.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | EPPlus 4.5.3.3 | 22.40 ms |  |  | 7.19 | 7.19 |  |  |  |  | 619.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | EPPlus | 24.72 ms | 1.31 ms | 0.75 ms | 7.94 | 7.94 | 15643.6 KB | 8.87 |  |  | 693.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | OfficeIMO.Excel | 3.15 ms | 0.06 ms | 0.03 ms | 1.00 | 1.00 | 1506.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | ClosedXML | 14.82 ms | 0.89 ms | 0.51 ms | 4.70 | 4.70 | 11296.3 KB | 7.50 |  |  | 370.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | EPPlus | 27.15 ms | 2.97 ms | 1.71 ms | 8.62 | 8.62 | 14960.7 KB | 9.93 |  |  | 761.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 3.00 ms | 0.09 ms | 0.05 ms | 1.00 | 1.00 | 1506.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | ClosedXML | 15.83 ms | 1.40 ms | 0.81 ms | 5.28 | 5.28 | 11296.3 KB | 7.50 |  |  | 427.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | EPPlus | 27.97 ms | 0.45 ms | 0.26 ms | 9.32 | 9.32 | 14960.7 KB | 9.93 |  |  | 832.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | OfficeIMO.Excel | 2.68 ms | 0.16 ms | 0.09 ms | 1.00 | 1.00 | 1138.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | ClosedXML | 10.76 ms | 0.33 ms | 0.19 ms | 4.02 | 4.02 | 9021.2 KB | 7.93 |  |  | 302.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | EPPlus 4.5.3.3 | 18.36 ms |  |  | 6.86 | 6.86 |  |  |  |  | 586.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | EPPlus | 22.72 ms | 0.30 ms | 0.17 ms | 8.49 | 8.49 | 12827.7 KB | 11.27 |  |  | 749.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | OfficeIMO.Excel | 3.26 ms | 0.01 ms | 0.00 ms | 1.00 | 1.00 | 1435.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | ClosedXML | 15.70 ms | 0.53 ms | 0.30 ms | 4.81 | 4.81 | 9711.1 KB | 6.76 |  |  | 381.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | EPPlus 4.5.3.3 | 20.54 ms |  |  | 6.29 | 6.29 |  |  |  |  | 529.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | EPPlus | 24.11 ms | 1.31 ms | 0.76 ms | 7.39 | 7.39 | 14722.8 KB | 10.26 |  |  | 638.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | OfficeIMO.Excel | 5.26 ms | 0.35 ms | 0.20 ms | 1.00 | 1.00 | 2064.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | MiniExcel | 14.35 ms | 0.67 ms | 0.39 ms | 2.73 | 2.73 | 29223.6 KB | 14.16 |  |  | 173.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | ClosedXML | 42.13 ms | 2.03 ms | 1.17 ms | 8.02 | 8.02 | 18913.3 KB | 9.16 |  |  | 701.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | EPPlus | 71.86 ms | 15.42 ms | 8.91 ms | 13.67 | 13.67 | 18410.5 KB | 8.92 |  |  | 1267.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | OfficeIMO.Excel | 6.70 ms | 0.18 ms | 0.11 ms | 1.00 | 1.00 | 2880.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | MiniExcel | 18.22 ms | 1.79 ms | 1.03 ms | 2.72 | 2.72 | 30510.5 KB | 10.59 |  |  | 172.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | ClosedXML | 60.91 ms | 2.08 ms | 1.20 ms | 9.10 | 9.10 | 27410.7 KB | 9.52 |  |  | 809.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | EPPlus | 67.57 ms | 1.37 ms | 0.79 ms | 10.09 | 10.09 | 22574.4 KB | 7.84 |  |  | 909.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | OfficeIMO.Excel | 4.93 ms | 0.15 ms | 0.08 ms | 1.00 | 1.00 | 2067.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | MiniExcel | 14.69 ms | 1.40 ms | 0.81 ms | 2.98 | 2.98 | 28700.3 KB | 13.88 |  |  | 198.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | EPPlus 4.5.3.3 | 40.60 ms |  |  | 8.24 | 8.24 |  |  |  |  | 723.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | ClosedXML | 76.68 ms | 3.44 ms | 1.99 ms | 15.56 | 15.56 | 18878.2 KB | 9.13 |  |  | 1455.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | EPPlus | 78.63 ms | 2.66 ms | 1.53 ms | 15.95 | 15.95 | 19431.0 KB | 9.40 |  |  | 1495.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | OfficeIMO.Excel | 4.39 ms | 0.06 ms | 0.03 ms | 1.00 | 1.00 | 1774.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | MiniExcel | 8.93 ms | 1.63 ms | 0.94 ms | 2.04 | 2.04 | 23044.2 KB | 12.98 |  |  | 103.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | EPPlus | 35.16 ms | 0.75 ms | 0.43 ms | 8.01 | 8.01 | 16647.3 KB | 9.38 |  |  | 701.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | ClosedXML | 35.72 ms | 0.70 ms | 0.41 ms | 8.14 | 8.14 | 19008.4 KB | 10.71 |  |  | 714.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | EPPlus 4.5.3.3 | 38.37 ms |  |  | 8.75 | 8.75 |  |  |  |  | 774.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | OfficeIMO.Excel | 5.04 ms | 0.40 ms | 0.23 ms | 1.00 | 1.00 | 1748.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | LargeXlsx | 6.80 ms | 0.02 ms | 0.01 ms | 1.35 | 1.35 | 1149.0 KB | 0.66 |  |  | 35.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | MiniExcel | 11.17 ms | 2.09 ms | 1.21 ms | 2.22 | 2.22 | 23415.0 KB | 13.39 |  |  | 121.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | ClosedXML | 27.12 ms | 0.38 ms | 0.22 ms | 5.38 | 5.38 | 11581.0 KB | 6.62 |  |  | 438.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | EPPlus | 36.65 ms | 1.45 ms | 0.84 ms | 7.27 | 7.27 | 16648.8 KB | 9.52 |  |  | 627.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | EPPlus 4.5.3.3 | 39.39 ms |  |  | 7.82 | 7.82 |  |  |  |  | 681.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | OfficeIMO.Excel | 4.53 ms | 0.10 ms | 0.06 ms | 1.00 | 1.00 | 1487.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | MiniExcel | 8.76 ms | 0.36 ms | 0.21 ms | 1.93 | 1.93 | 22789.5 KB | 15.32 |  |  | 93.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | EPPlus | 34.01 ms | 0.80 ms | 0.46 ms | 7.51 | 7.51 | 16374.5 KB | 11.01 |  |  | 651.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | ClosedXML | 34.51 ms | 0.77 ms | 0.45 ms | 7.62 | 7.62 | 18735.1 KB | 12.60 |  |  | 661.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | OfficeIMO.Excel | 4.40 ms | 0.16 ms | 0.09 ms | 1.00 | 1.00 | 1760.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | MiniExcel | 9.00 ms | 0.15 ms | 0.09 ms | 2.05 | 2.05 | 23062.9 KB | 13.10 |  |  | 104.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | EPPlus | 35.55 ms | 0.60 ms | 0.35 ms | 8.08 | 8.08 | 16648.8 KB | 9.46 |  |  | 708.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | ClosedXML | 35.62 ms | 0.91 ms | 0.53 ms | 8.10 | 8.10 | 19008.7 KB | 10.80 |  |  | 710.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | EPPlus 4.5.3.3 | 38.88 ms |  |  | 8.84 | 8.84 |  |  |  |  | 784.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | OfficeIMO.Excel | 6.22 ms | 0.46 ms | 0.27 ms | 1.00 | 1.00 | 1403.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | MiniExcel | 15.69 ms | 2.44 ms | 1.41 ms | 2.52 | 2.52 | 26825.0 KB | 19.12 |  |  | 152.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | EPPlus 4.5.3.3 | 110.55 ms |  |  | 17.77 | 17.77 |  |  |  |  | 1676.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | EPPlus | 126.60 ms | 8.05 ms | 4.65 ms | 20.35 | 20.35 | 49158.1 KB | 35.03 |  |  | 1934.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | ClosedXML | 225.25 ms | 20.82 ms | 12.02 ms | 36.20 | 36.20 | 58350.2 KB | 41.58 |  |  | 3519.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | OfficeIMO.Excel | 3.13 ms | 0.23 ms | 0.14 ms | 1.00 | 1.00 | 1620.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | ClosedXML | 18.98 ms | 0.67 ms | 0.39 ms | 6.05 | 6.05 | 12039.8 KB | 7.43 |  |  | 505.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | EPPlus 4.5.3.3 | 22.73 ms |  |  | 7.25 | 7.25 |  |  |  |  | 625.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | EPPlus | 36.93 ms | 2.59 ms | 1.50 ms | 11.78 | 11.78 | 18110.8 KB | 11.18 |  |  | 1078.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-dictionary-objects-table-direct | OfficeIMO.Excel | 5.46 ms | 0.18 ms | 0.10 ms | 1.00 | 1.00 | 2051.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | LargeXlsx | 4.15 ms | 0.02 ms | 0.01 ms | 0.77 | 1.00 | 802.5 KB | 0.34 |  |  | 23.5% faster than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 5.42 ms | 0.10 ms | 0.06 ms | 1.00 | 1.31 | 2341.7 KB | 1.00 |  |  | Loss +30.7% |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | MiniExcel | 8.67 ms | 0.05 ms | 0.03 ms | 1.60 | 2.09 | 25190.5 KB | 10.76 |  |  | 60.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | ClosedXML | 36.99 ms | 0.85 ms | 0.49 ms | 6.83 | 8.92 | 16973.5 KB | 7.25 |  |  | 582.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | EPPlus | 47.39 ms | 1.45 ms | 0.84 ms | 8.75 | 11.43 | 20105.6 KB | 8.59 |  |  | 774.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | LargeXlsx | 4.13 ms | 0.04 ms | 0.02 ms | 0.79 | 1.00 | 802.5 KB | 0.53 |  |  | 21.1% faster than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 5.24 ms | 0.10 ms | 0.06 ms | 1.00 | 1.27 | 1507.7 KB | 1.00 |  |  | Loss +26.7% |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | MiniExcel | 8.41 ms | 0.16 ms | 0.09 ms | 1.61 | 2.04 | 25190.5 KB | 16.71 |  |  | 60.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | ClosedXML | 38.16 ms | 4.10 ms | 2.37 ms | 7.29 | 9.23 | 16973.5 KB | 11.26 |  |  | 628.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | EPPlus | 47.12 ms | 1.31 ms | 0.75 ms | 9.00 | 11.40 | 20105.6 KB | 13.33 |  |  | 799.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | LargeXlsx | 20.03 ms | 0.29 ms | 0.17 ms | 0.90 | 1.00 | 2810.7 KB | 0.62 |  |  | 9.8% faster than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 22.20 ms | 6.17 ms | 3.56 ms | 1.00 | 1.11 | 4502.3 KB | 1.00 |  |  | Loss +10.8% |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | MiniExcel | 34.75 ms | 1.77 ms | 1.02 ms | 1.57 | 1.73 | 48414.8 KB | 10.75 |  |  | 56.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | ClosedXML | 119.19 ms | 2.04 ms | 1.18 ms | 5.37 | 5.95 | 51647.0 KB | 11.47 |  |  | 436.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | EPPlus | 156.73 ms | 1.96 ms | 1.13 ms | 7.06 | 7.82 | 69140.0 KB | 15.36 |  |  | 606.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | LargeXlsx | 1.38 ms | 0.02 ms | 0.01 ms | 0.64 | 1.00 | 296.4 KB | 0.19 |  |  | 36.2% faster than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | OfficeIMO.Excel | 2.16 ms | 0.22 ms | 0.13 ms | 1.00 | 1.57 | 1576.3 KB | 1.00 |  |  | Loss +56.7% |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | MiniExcel | 3.97 ms | 0.10 ms | 0.06 ms | 1.83 | 2.87 | 19710.9 KB | 12.50 |  |  | 83.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | ClosedXML | 16.46 ms | 3.42 ms | 1.97 ms | 7.61 | 11.93 | 11197.4 KB | 7.10 |  |  | 661.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | EPPlus 4.5.3.3 | 20.58 ms |  |  | 9.52 | 14.91 |  |  |  |  | 851.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | EPPlus | 24.70 ms | 0.57 ms | 0.33 ms | 11.43 | 17.90 | 14365.5 KB | 9.11 |  |  | 1042.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 1.58 ms | 0.02 ms | 0.01 ms | 1.00 | 1.00 | 1092.0 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 2.93 ms | 0.21 ms | 0.12 ms | 1.86 | 1.86 | 447.0 KB | 0.41 |  |  | 85.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | ClosedXML | 15.37 ms | 3.86 ms | 2.23 ms | 9.75 | 9.75 | 10235.8 KB | 9.37 |  |  | 875.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | EPPlus | 23.31 ms | 0.69 ms | 0.40 ms | 14.79 | 14.79 | 13052.8 KB | 11.95 |  |  | 1379.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 4.33 ms | 0.59 ms | 0.34 ms | 1.00 | 1.00 | 2081.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | LargeXlsx | 7.05 ms | 0.24 ms | 0.14 ms | 1.63 | 1.63 | 758.3 KB | 0.36 |  |  | 62.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | MiniExcel | 8.87 ms | 0.25 ms | 0.14 ms | 2.05 | 2.05 | 23221.8 KB | 11.16 |  |  | 104.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | ClosedXML | 33.24 ms | 1.95 ms | 1.12 ms | 7.67 | 7.67 | 22221.3 KB | 10.68 |  |  | 667.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 36.64 ms |  |  | 8.46 | 8.46 |  |  |  |  | 745.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | EPPlus | 43.54 ms | 1.36 ms | 0.78 ms | 10.05 | 10.05 | 24694.7 KB | 11.87 |  |  | 904.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | OfficeIMO.Excel | 3.38 ms | 1.88 ms | 1.09 ms | 1.00 | 1.00 | 1494.0 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | ClosedXML | 15.37 ms | 0.37 ms | 0.22 ms | 4.55 | 4.55 | 11296.3 KB | 7.56 |  |  | 354.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | EPPlus | 25.35 ms | 0.48 ms | 0.28 ms | 7.50 | 7.50 | 14960.8 KB | 10.01 |  |  | 650.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | Sylvan.Data.Excel | 3.61 ms | 0.44 ms | 0.25 ms | 0.77 | 1.00 | 758.6 KB | 0.43 |  |  | 22.9% faster than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | OfficeIMO.Excel | 4.68 ms | 0.90 ms | 0.52 ms | 1.00 | 1.30 | 1763.0 KB | 1.00 |  |  | Loss +29.7% |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | MiniExcel | 7.46 ms | 0.10 ms | 0.06 ms | 1.60 | 2.07 | 23044.0 KB | 13.07 |  |  | 59.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | LargeXlsx | 8.21 ms | 0.43 ms | 0.25 ms | 1.76 | 2.28 | 1032.5 KB | 0.59 |  |  | 75.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | ClosedXML | 26.90 ms | 1.20 ms | 0.69 ms | 5.75 | 7.46 | 11581.0 KB | 6.57 |  |  | 475.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | EPPlus 4.5.3.3 | 33.60 ms |  |  | 7.18 | 9.32 |  |  |  |  | 618.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | EPPlus | 34.50 ms | 0.29 ms | 0.17 ms | 7.38 | 9.57 | 16648.1 KB | 9.44 |  |  | 637.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | OfficeIMO.Excel | 5.69 ms | 0.10 ms | 0.06 ms | 1.00 | 1.00 | 2140.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | LargeXlsx | 8.87 ms | 0.28 ms | 0.16 ms | 1.56 | 1.56 | 1123.9 KB | 0.53 |  |  | 55.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | MiniExcel | 15.63 ms | 0.69 ms | 0.40 ms | 2.75 | 2.75 | 30510.6 KB | 14.25 |  |  | 174.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | EPPlus | 56.81 ms | 7.01 ms | 4.05 ms | 9.99 | 9.99 | 22120.2 KB | 10.33 |  |  | 898.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | ClosedXML | 59.53 ms | 1.87 ms | 1.08 ms | 10.46 | 10.46 | 27410.6 KB | 12.80 |  |  | 946.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | OfficeIMO.Excel | 4.76 ms | 0.06 ms | 0.04 ms | 1.00 | 1.00 | 1676.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | LargeXlsx | 7.44 ms | 5.46 ms | 3.15 ms | 1.56 | 1.56 | 857.6 KB | 0.51 |  |  | 56.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | MiniExcel | 18.50 ms | 0.70 ms | 0.41 ms | 3.88 | 3.88 | 35917.8 KB | 21.42 |  |  | 288.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | ClosedXML | 92.91 ms | 3.47 ms | 2.00 ms | 19.50 | 19.50 | 71478.2 KB | 42.63 |  |  | 1850.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | OfficeIMO.Excel | 2.04 ms | 0.09 ms | 0.05 ms | 1.00 | 1.00 | 2440.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | MiniExcel | 3.80 ms | 0.31 ms | 0.18 ms | 1.86 | 1.86 | 21137.5 KB | 8.66 |  |  | 86.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | EPPlus 4.5.3.3 | 9.86 ms |  |  | 4.83 | 4.83 |  |  |  |  | 382.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | ClosedXML | 12.09 ms | 0.87 ms | 0.50 ms | 5.92 | 5.92 | 11299.2 KB | 4.63 |  |  | 492.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | EPPlus | 21.62 ms | 0.50 ms | 0.29 ms | 10.58 | 10.58 | 12804.8 KB | 5.25 |  |  | 958.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | OfficeIMO.Excel | 3.17 ms | 0.15 ms | 0.09 ms | 1.00 | 1.00 | 2617.0 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | ClosedXML | 13.86 ms | 1.33 ms | 0.77 ms | 4.37 | 4.37 | 11299.2 KB | 4.32 |  |  | 337.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | EPPlus 4.5.3.3 | 20.63 ms |  |  | 6.51 | 6.51 |  |  |  |  | 550.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | EPPlus | 20.89 ms | 0.64 ms | 0.37 ms | 6.59 | 6.59 | 12805.2 KB | 4.89 |  |  | 558.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | OfficeIMO.Excel | 2.48 ms | 0.03 ms | 0.02 ms | 1.00 | 1.00 | 2379.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | ClosedXML | 14.71 ms | 0.17 ms | 0.10 ms | 5.92 | 5.92 | 13127.1 KB | 5.52 |  |  | 492.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | EPPlus | 23.99 ms | 0.46 ms | 0.27 ms | 9.66 | 9.66 | 13893.2 KB | 5.84 |  |  | 866.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | OfficeIMO.Excel | 2.27 ms | 0.04 ms | 0.02 ms | 1.00 | 1.00 | 1579.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | ClosedXML | 10.98 ms | 0.14 ms | 0.08 ms | 4.85 | 4.85 | 9226.5 KB | 5.84 |  |  | 384.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | EPPlus | 19.15 ms | 0.49 ms | 0.28 ms | 8.45 | 8.45 | 11332.7 KB | 7.17 |  |  | 745.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | LargeXlsx | 3.27 ms | 0.28 ms | 0.16 ms | 0.85 | 1.00 | 758.3 KB | 0.43 |  |  | 15.3% faster than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 3.86 ms | 0.18 ms | 0.10 ms | 1.00 | 1.18 | 1769.2 KB | 1.00 |  |  | Loss +18.1% |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | MiniExcel | 7.93 ms | 0.46 ms | 0.26 ms | 2.06 | 2.43 | 23222.3 KB | 13.13 |  |  | 105.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | ClosedXML | 27.35 ms | 0.78 ms | 0.45 ms | 7.08 | 8.36 | 11581.0 KB | 6.55 |  |  | 608.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 41.48 ms |  |  | 10.75 | 12.69 |  |  |  |  | 974.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | EPPlus | 41.80 ms | 7.05 ms | 4.07 ms | 10.83 | 12.78 | 16646.8 KB | 9.41 |  |  | 982.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | LargeXlsx | 3.16 ms | 0.02 ms | 0.01 ms | 0.85 | 1.00 | 758.3 KB | 0.57 |  |  | 15.0% faster than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | OfficeIMO.Excel | 3.71 ms | 0.10 ms | 0.06 ms | 1.00 | 1.18 | 1339.3 KB | 1.00 |  |  | Loss +17.7% |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | MiniExcel | 8.79 ms | 1.02 ms | 0.59 ms | 2.37 | 2.79 | 23222.4 KB | 17.34 |  |  | 136.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | ClosedXML | 26.91 ms | 0.12 ms | 0.07 ms | 7.25 | 8.53 | 11581.0 KB | 8.65 |  |  | 624.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | EPPlus 4.5.3.3 | 39.61 ms |  |  | 10.67 | 12.55 |  |  |  |  | 966.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | EPPlus | 39.75 ms | 5.31 ms | 3.07 ms | 10.71 | 12.59 | 16646.5 KB | 12.43 |  |  | 970.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 3.21 ms | 0.22 ms | 0.13 ms | 0.74 | 1.00 | 758.3 KB | 0.51 |  |  | 25.8% faster than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 4.33 ms | 0.12 ms | 0.07 ms | 1.00 | 1.35 | 1497.5 KB | 1.00 |  |  | Loss +34.7% |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | MiniExcel | 7.98 ms | 0.26 ms | 0.15 ms | 1.84 | 2.48 | 23222.4 KB | 15.51 |  |  | 84.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | ClosedXML | 28.38 ms | 0.51 ms | 0.30 ms | 6.56 | 8.83 | 11581.0 KB | 7.73 |  |  | 555.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | EPPlus | 42.13 ms | 7.77 ms | 4.49 ms | 9.73 | 13.11 | 16646.5 KB | 11.12 |  |  | 873.2% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | Sylvan.Data.Excel | 73.68 ms | 36.05 ms | 20.81 ms | 0.91 | 1.00 | 394.1 KB | 0.02 |  |  | 9.2% faster than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | OfficeIMO.Excel | 81.10 ms | 39.80 ms | 22.98 ms | 1.00 | 1.10 | 23621.9 KB | 1.00 |  |  | Loss +10.1% |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | MiniExcel | 194.48 ms | 22.37 ms | 12.91 ms | 2.40 | 2.64 | 215349.0 KB | 9.12 |  |  | 139.8% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | ExcelDataReader | 206.19 ms | 56.48 ms | 32.61 ms | 2.54 | 2.80 | 69530.7 KB | 2.94 |  |  | 154.2% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | Sylvan.Data.Excel | 40.87 ms | 2.18 ms | 1.26 ms | 0.77 | 1.00 | 394.1 KB | 0.02 |  |  | 23.4% faster than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | OfficeIMO.Excel | 53.33 ms | 12.36 ms | 7.14 ms | 1.00 | 1.31 | 24404.0 KB | 1.00 |  |  | Loss +30.5% |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | ExcelDataReader | 124.69 ms | 7.97 ms | 4.60 ms | 2.34 | 3.05 | 69530.7 KB | 2.85 |  |  | 133.8% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | MiniExcel | 165.54 ms | 6.52 ms | 3.76 ms | 3.10 | 4.05 | 215349.0 KB | 8.82 |  |  | 210.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | append-plain-rows | LargeXlsx | 11.83 ms | 0.56 ms | 0.33 ms | 0.83 | 1.00 | 2771.0 KB | 0.26 | 605.0 KB | 0.99 | 17.2% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | append-plain-rows | OfficeIMO.Excel | 14.28 ms | 0.24 ms | 0.14 ms | 1.00 | 1.21 | 10842.5 KB | 1.00 | 610.4 KB | 1.00 | Loss +20.8% |
| 25000 | package-profile | package | Package size | append-plain-rows | MiniExcel | 30.72 ms | 0.74 ms | 0.43 ms | 2.15 | 2.60 | 58242.9 KB | 5.37 | 642.3 KB | 1.05 | 115.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | append-plain-rows | ClosedXML | 136.51 ms | 14.10 ms | 8.14 ms | 9.56 | 11.54 | 104233.1 KB | 9.61 | 540.6 KB | 0.89 | 855.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | append-plain-rows | EPPlus | 206.42 ms | 1.43 ms | 0.83 ms | 14.45 | 17.45 | 100373.9 KB | 9.26 | 525.6 KB | 0.86 | 1345.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | autofit-existing | OfficeIMO.Excel | 81.97 ms | 4.94 ms | 2.85 ms | 1.00 | 1.00 | 15710.9 KB | 1.00 | 1429.3 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | autofit-existing | EPPlus | 451.56 ms | 1.75 ms | 1.01 ms | 5.51 | 5.51 | 250950.0 KB | 15.97 | 1091.0 KB | 0.76 | 450.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | autofit-existing | ClosedXML | 1331.71 ms | 16.88 ms | 9.74 ms | 16.25 | 16.25 | 829716.9 KB | 52.81 | 1140.9 KB | 0.80 | 1524.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | large-shared-strings | OfficeIMO.Excel | 23.48 ms | 8.79 ms | 5.08 ms | 1.00 | 1.00 | 15744.9 KB | 1.00 | 529.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | large-shared-strings | MiniExcel | 50.51 ms | 5.97 ms | 3.45 ms | 2.15 | 2.15 | 73760.2 KB | 4.68 | 581.0 KB | 1.10 | 115.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | large-shared-strings | ClosedXML | 169.37 ms | 40.50 ms | 23.38 ms | 7.21 | 7.21 | 104241.3 KB | 6.62 | 460.1 KB | 0.87 | 621.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | large-shared-strings | EPPlus | 308.49 ms | 41.05 ms | 23.70 ms | 13.14 | 13.14 | 84410.3 KB | 5.36 | 444.7 KB | 0.84 | 1213.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-autofilter | OfficeIMO.Excel | 32.88 ms | 1.69 ms | 0.98 ms | 1.00 | 1.00 | 11494.9 KB | 1.00 | 1428.5 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-autofilter | ClosedXML | 306.11 ms | 2.31 ms | 1.33 ms | 9.31 | 9.31 | 210663.8 KB | 18.33 | 1140.0 KB | 0.80 | 830.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-autofilter | EPPlus | 420.88 ms | 12.77 ms | 7.37 ms | 12.80 | 12.80 | 211871.8 KB | 18.43 | 1090.1 KB | 0.76 | 1179.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-charts | OfficeIMO.Excel | 34.79 ms | 1.08 ms | 0.62 ms | 1.00 | 1.00 | 12553.1 KB | 1.00 | 1433.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-charts | EPPlus | 447.66 ms | 14.26 ms | 8.23 ms | 12.87 | 12.87 | 214906.2 KB | 17.12 | 1092.9 KB | 0.76 | 1186.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-conditional-formatting | OfficeIMO.Excel | 32.38 ms | 0.95 ms | 0.55 ms | 1.00 | 1.00 | 11560.2 KB | 1.00 | 1428.8 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-conditional-formatting | ClosedXML | 311.49 ms | 2.28 ms | 1.32 ms | 9.62 | 9.62 | 210711.7 KB | 18.23 | 1140.1 KB | 0.80 | 861.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-conditional-formatting | EPPlus | 421.72 ms | 13.57 ms | 7.84 ms | 13.02 | 13.02 | 211913.3 KB | 18.33 | 1090.2 KB | 0.76 | 1202.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-data-validation | OfficeIMO.Excel | 35.11 ms | 3.13 ms | 1.81 ms | 1.00 | 1.00 | 11510.5 KB | 1.00 | 1428.6 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-data-validation | ClosedXML | 334.51 ms | 29.45 ms | 17.00 ms | 9.53 | 9.53 | 210672.7 KB | 18.30 | 1140.1 KB | 0.80 | 852.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-data-validation | EPPlus | 428.30 ms | 7.74 ms | 4.47 ms | 12.20 | 12.20 | 211857.8 KB | 18.41 | 1090.1 KB | 0.76 | 1119.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-freeze-panes | OfficeIMO.Excel | 33.44 ms | 0.69 ms | 0.40 ms | 1.00 | 1.00 | 11497.3 KB | 1.00 | 1428.6 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-freeze-panes | ClosedXML | 309.66 ms | 17.56 ms | 10.14 ms | 9.26 | 9.26 | 210646.8 KB | 18.32 | 1140.0 KB | 0.80 | 826.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-freeze-panes | EPPlus | 430.69 ms | 27.51 ms | 15.88 ms | 12.88 | 12.88 | 211883.7 KB | 18.43 | 1090.2 KB | 0.76 | 1188.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-pivot-table | OfficeIMO.Excel | 282.43 ms | 18.47 ms | 10.66 ms | 1.00 | 1.00 | 131929.1 KB | 1.00 | 1979.8 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-pivot-table | EPPlus | 459.68 ms | 16.82 ms | 9.71 ms | 1.63 | 1.63 | 230801.7 KB | 1.75 | 1093.4 KB | 0.55 | 62.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-all-in-one | OfficeIMO.Excel | 259.80 ms | 28.15 ms | 16.25 ms | 1.00 | 1.00 | 133444.7 KB | 1.00 | 1985.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-all-in-one | EPPlus | 462.30 ms | 34.50 ms | 19.92 ms | 1.78 | 1.78 | 277078.9 KB | 2.08 | 1097.7 KB | 0.55 | 77.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-chart-first | OfficeIMO.Excel | 91.62 ms | 3.21 ms | 1.86 ms | 1.00 | 1.00 | 43560.6 KB | 1.00 | 1985.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-chart-first | EPPlus | 460.14 ms | 21.64 ms | 12.49 ms | 5.02 | 5.02 | 277077.7 KB | 6.36 | 1097.8 KB | 0.55 | 402.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-core | OfficeIMO.Excel | 59.32 ms | 21.79 ms | 12.58 ms | 1.00 | 1.00 | 11648.7 KB | 1.00 | 1430.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-core | EPPlus | 683.01 ms | 84.47 ms | 48.77 ms | 11.51 | 11.51 | 255066.2 KB | 21.90 | 1091.5 KB | 0.76 | 1051.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-core | ClosedXML | 1301.18 ms | 135.87 ms | 78.44 ms | 21.93 | 21.93 | 680117.1 KB | 58.39 | 1141.3 KB | 0.80 | 2093.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-extra-column | OfficeIMO.Excel | 358.46 ms | 2.82 ms | 1.63 ms | 1.00 | 1.00 | 144827.2 KB | 1.00 | 2110.8 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-extra-column | EPPlus | 518.72 ms | 6.03 ms | 3.48 ms | 1.45 | 1.45 | 302761.2 KB | 2.09 | 1166.3 KB | 0.55 | 44.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-no-autofit | OfficeIMO.Excel | 265.40 ms | 7.08 ms | 4.08 ms | 1.00 | 1.00 | 133435.8 KB | 1.00 | 1985.8 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-no-autofit | EPPlus | 441.51 ms | 11.05 ms | 6.38 ms | 1.66 | 1.66 | 234783.8 KB | 1.76 | 1097.7 KB | 0.55 | 66.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-post-mutation | OfficeIMO.Excel | 274.82 ms | 4.58 ms | 2.64 ms | 1.00 | 1.00 | 133463.2 KB | 1.00 | 1986.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-post-mutation | EPPlus | 461.37 ms | 5.40 ms | 3.12 ms | 1.68 | 1.68 | 277078.9 KB | 2.08 | 1097.8 KB | 0.55 | 67.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-shuffled-columns | OfficeIMO.Excel | 296.24 ms | 5.42 ms | 3.13 ms | 1.00 | 1.00 | 133503.4 KB | 1.00 | 2046.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-shuffled-columns | EPPlus | 482.21 ms | 6.42 ms | 3.71 ms | 1.63 | 1.63 | 277071.7 KB | 2.08 | 1098.4 KB | 0.54 | 62.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook | OfficeIMO.Excel | 332.42 ms | 8.54 ms | 4.93 ms | 1.00 | 1.00 | 175197.5 KB | 1.00 | 2672.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | report-workbook | EPPlus | 576.90 ms | 10.41 ms | 6.01 ms | 1.74 | 1.74 | 364710.2 KB | 2.08 | 1517.2 KB | 0.57 | 73.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook-core | OfficeIMO.Excel | 49.38 ms | 2.00 ms | 1.15 ms | 1.00 | 1.00 | 10979.4 KB | 1.00 | 1850.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | report-workbook-core | EPPlus | 571.21 ms | 17.05 ms | 9.85 ms | 11.57 | 11.57 | 342842.6 KB | 31.23 | 1512.6 KB | 0.82 | 1056.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook-core | ClosedXML | 1110.68 ms | 15.01 ms | 8.66 ms | 22.49 | 22.49 | 975774.1 KB | 88.87 | 1579.8 KB | 0.85 | 2149.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook-datatable | OfficeIMO.Excel | 358.15 ms | 7.66 ms | 4.42 ms | 1.00 | 1.00 | 177940.5 KB | 1.00 | 2672.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | report-workbook-datatable | EPPlus | 583.72 ms | 17.53 ms | 10.12 ms | 1.63 | 1.63 | 247824.1 KB | 1.39 | 1517.2 KB | 0.57 | 63.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook-datatable-core | OfficeIMO.Excel | 48.24 ms | 1.42 ms | 0.82 ms | 1.00 | 1.00 | 13725.0 KB | 1.00 | 1850.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | report-workbook-datatable-core | EPPlus | 517.17 ms | 5.71 ms | 3.30 ms | 10.72 | 10.72 | 225957.5 KB | 16.46 | 1512.6 KB | 0.82 | 972.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook-datatable-core | ClosedXML | 1026.91 ms | 11.60 ms | 6.70 ms | 21.29 | 21.29 | 832229.0 KB | 60.64 | 1579.8 KB | 0.85 | 2028.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | LargeXlsx | 38.47 ms | 0.86 ms | 0.50 ms | 0.89 | 1.00 | 10795.2 KB | 0.92 | 2444.6 KB | 1.10 | 10.7% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | OfficeIMO.Excel | 43.08 ms | 1.16 ms | 0.67 ms | 1.00 | 1.12 | 11708.2 KB | 1.00 | 2228.8 KB | 1.00 | Loss +12.0% |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | MiniExcel | 145.45 ms | 5.18 ms | 2.99 ms | 3.38 | 3.78 | 226875.5 KB | 19.38 | 2410.6 KB | 1.08 | 237.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | ClosedXML | 889.75 ms | 13.03 ms | 7.52 ms | 20.65 | 23.13 | 759818.4 KB | 64.90 | 2581.2 KB | 1.16 | 1965.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-bulk-report | OfficeIMO.Excel | 35.34 ms | 0.53 ms | 0.30 ms | 1.00 | 1.00 | 11561.8 KB | 1.00 | 1429.3 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-bulk-report | MiniExcel | 68.93 ms | 2.08 ms | 1.20 ms | 1.95 | 1.95 | 125551.5 KB | 10.86 | 1521.1 KB | 1.06 | 95.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-bulk-report | EPPlus | 401.74 ms | 22.10 ms | 12.76 ms | 11.37 | 11.37 | 254959.4 KB | 22.05 | 1091.0 KB | 0.76 | 1036.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-bulk-report | ClosedXML | 789.94 ms | 13.91 ms | 8.03 ms | 22.35 | 22.35 | 565953.3 KB | 48.95 | 1140.9 KB | 0.80 | 2135.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellformula | OfficeIMO.Excel | 19.37 ms | 1.32 ms | 0.76 ms | 1.00 | 1.00 | 10112.0 KB | 1.00 | 670.3 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellformula | ClosedXML | 167.75 ms | 6.09 ms | 3.52 ms | 8.66 | 8.66 | 113853.5 KB | 11.26 | 643.2 KB | 0.96 | 766.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellformula | EPPlus | 301.99 ms | 1.85 ms | 1.07 ms | 15.59 | 15.59 | 140732.3 KB | 13.92 | 593.9 KB | 0.89 | 1459.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-empty-strings | OfficeIMO.Excel | 11.89 ms | 0.55 ms | 0.32 ms | 1.00 | 1.00 | 6896.4 KB | 1.00 | 451.4 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-empty-strings | ClosedXML | 119.57 ms | 6.22 ms | 3.59 ms | 10.06 | 10.06 | 92902.1 KB | 13.47 | 398.1 KB | 0.88 | 906.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-empty-strings | EPPlus | 175.67 ms | 6.65 ms | 3.84 ms | 14.78 | 14.78 | 74493.1 KB | 10.80 | 390.6 KB | 0.87 | 1378.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-numbers | OfficeIMO.Excel | 17.33 ms | 4.16 ms | 2.40 ms | 1.00 | 1.00 | 5970.9 KB | 1.00 | 462.6 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-numbers | ClosedXML | 112.19 ms | 6.75 ms | 3.90 ms | 6.48 | 6.48 | 84206.7 KB | 14.10 | 411.4 KB | 0.89 | 547.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-numbers | EPPlus | 193.00 ms | 5.23 ms | 3.02 ms | 11.14 | 11.14 | 86377.9 KB | 14.47 | 406.5 KB | 0.88 | 1013.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-mixed | OfficeIMO.Excel | 17.11 ms | 0.91 ms | 0.53 ms | 1.00 | 1.00 | 8332.9 KB | 1.00 | 585.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-object-mixed | ClosedXML | 155.97 ms | 4.18 ms | 2.42 ms | 9.12 | 9.12 | 111118.7 KB | 13.33 | 532.9 KB | 0.91 | 811.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-mixed | EPPlus | 202.66 ms | 0.79 ms | 0.46 ms | 11.85 | 11.85 | 113245.5 KB | 13.59 | 544.3 KB | 0.93 | 1084.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse | OfficeIMO.Excel | 20.02 ms | 2.37 ms | 1.37 ms | 1.00 | 1.00 | 7416.2 KB | 1.00 | 607.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse | ClosedXML | 137.51 ms | 2.29 ms | 1.32 ms | 6.87 | 6.87 | 105223.9 KB | 14.19 | 468.0 KB | 0.77 | 586.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse | EPPlus | 214.53 ms | 14.59 ms | 8.42 ms | 10.71 | 10.71 | 106317.3 KB | 14.34 | 494.4 KB | 0.81 | 971.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 17.81 ms | 1.36 ms | 0.78 ms | 1.00 | 1.00 | 7416.3 KB | 1.00 | 607.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | ClosedXML | 137.98 ms | 2.37 ms | 1.37 ms | 7.75 | 7.75 | 105223.9 KB | 14.19 | 468.0 KB | 0.77 | 674.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | EPPlus | 206.62 ms | 4.22 ms | 2.44 ms | 11.60 | 11.60 | 106317.3 KB | 14.34 | 494.4 KB | 0.81 | 1060.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-scalars | OfficeIMO.Excel | 11.00 ms | 0.43 ms | 0.25 ms | 1.00 | 1.00 | 6144.6 KB | 1.00 | 441.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-scalars | ClosedXML | 100.15 ms | 3.45 ms | 1.99 ms | 9.11 | 9.11 | 82591.3 KB | 13.44 | 394.9 KB | 0.89 | 810.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-scalars | EPPlus | 193.67 ms | 5.00 ms | 2.89 ms | 17.61 | 17.61 | 85127.8 KB | 13.85 | 379.3 KB | 0.86 | 1661.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings | OfficeIMO.Excel | 16.75 ms | 0.35 ms | 0.20 ms | 1.00 | 1.00 | 15360.4 KB | 1.00 | 527.8 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-strings | ClosedXML | 113.57 ms | 5.86 ms | 3.38 ms | 6.78 | 6.78 | 104241.3 KB | 6.79 | 460.1 KB | 0.87 | 578.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings | EPPlus | 185.44 ms | 6.96 ms | 4.02 ms | 11.07 | 11.07 | 84410.8 KB | 5.50 | 444.7 KB | 0.84 | 1007.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-distinct | OfficeIMO.Excel | 12.87 ms | 0.39 ms | 0.23 ms | 1.00 | 1.00 | 13824.1 KB | 1.00 | 499.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-distinct | ClosedXML | 151.68 ms | 2.87 ms | 1.66 ms | 11.78 | 11.78 | 131501.7 KB | 9.51 | 555.3 KB | 1.11 | 1078.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-distinct | EPPlus | 215.34 ms | 7.89 ms | 4.55 ms | 16.73 | 16.73 | 97730.0 KB | 7.07 | 565.1 KB | 1.13 | 1572.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-repeated | OfficeIMO.Excel | 12.00 ms | 0.62 ms | 0.36 ms | 1.00 | 1.00 | 7525.3 KB | 1.00 | 376.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-repeated | ClosedXML | 95.79 ms | 4.34 ms | 2.51 ms | 7.98 | 7.98 | 84520.0 KB | 11.23 | 331.8 KB | 0.88 | 697.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-repeated | EPPlus | 151.89 ms | 2.47 ms | 1.43 ms | 12.65 | 12.65 | 70033.7 KB | 9.31 | 300.8 KB | 0.80 | 1165.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-temporal | OfficeIMO.Excel | 22.92 ms | 6.75 ms | 3.90 ms | 1.00 | 1.00 | 7482.8 KB | 1.00 | 620.5 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-temporal | ClosedXML | 150.35 ms | 6.81 ms | 3.93 ms | 6.56 | 6.56 | 89323.7 KB | 11.94 | 483.0 KB | 0.78 | 556.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-temporal | EPPlus | 195.57 ms | 5.39 ms | 3.11 ms | 8.53 | 8.53 | 103800.4 KB | 13.87 | 495.1 KB | 0.80 | 753.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 9.74 ms | 0.22 ms | 0.13 ms | 0.93 | 1.00 | 3444.4 KB | 0.49 | 443.4 KB | 0.97 | 7.0% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 10.47 ms | 0.15 ms | 0.09 ms | 1.00 | 1.08 | 6961.7 KB | 1.00 | 455.5 KB | 1.00 | Loss +7.5% |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | ClosedXML | 124.73 ms | 4.54 ms | 2.62 ms | 11.91 | 12.81 | 96015.7 KB | 13.79 | 467.5 KB | 1.03 | 1090.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | EPPlus | 202.06 ms | 14.39 ms | 8.31 ms | 19.29 | 20.75 | 87467.3 KB | 12.56 | 484.1 KB | 1.06 | 1829.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | LargeXlsx | 28.54 ms | 0.35 ms | 0.20 ms | 0.82 | 1.00 | 5614.1 KB | 0.35 | 1386.5 KB | 1.00 | 17.5% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 34.60 ms | 0.67 ms | 0.39 ms | 1.00 | 1.21 | 16036.5 KB | 1.00 | 1384.9 KB | 1.00 | Loss +21.2% |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | MiniExcel | 68.77 ms | 5.86 ms | 3.39 ms | 1.99 | 2.41 | 93257.0 KB | 5.82 | 1521.0 KB | 1.10 | 98.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | ClosedXML | 311.95 ms | 8.11 ms | 4.68 ms | 9.02 | 10.93 | 210646.1 KB | 13.14 | 1139.9 KB | 0.82 | 801.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | EPPlus | 401.78 ms | 44.04 ms | 25.43 ms | 11.61 | 14.08 | 211850.3 KB | 13.21 | 1090.0 KB | 0.79 | 1061.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | Sylvan.Data.Excel | 28.26 ms | 1.44 ms | 0.83 ms | 0.74 | 1.00 | 5700.3 KB | 0.44 | 755.4 KB | 0.55 | 25.8% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | LargeXlsx | 36.80 ms | 3.38 ms | 1.95 ms | 0.97 | 1.30 | 8349.2 KB | 0.64 | 1386.5 KB | 1.00 | 3.5% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | OfficeIMO.Excel | 38.11 ms | 2.01 ms | 1.16 ms | 1.00 | 1.35 | 13002.3 KB | 1.00 | 1384.9 KB | 1.00 | Loss +34.9% |
| 25000 | package-profile | package | Package size | write-datareader-plain | MiniExcel | 75.73 ms | 4.63 ms | 2.67 ms | 1.99 | 2.68 | 92199.8 KB | 7.09 | 1521.0 KB | 1.10 | 98.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | ClosedXML | 284.04 ms | 12.71 ms | 7.34 ms | 7.45 | 10.05 | 104205.0 KB | 8.01 | 1139.9 KB | 0.82 | 645.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | EPPlus | 346.01 ms | 19.03 ms | 10.98 ms | 9.08 | 12.24 | 117438.0 KB | 9.03 | 1090.8 KB | 0.79 | 807.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table | OfficeIMO.Excel | 38.27 ms | 1.90 ms | 1.10 ms | 1.00 | 1.00 | 13020.3 KB | 1.00 | 1385.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-datareader-table | MiniExcel | 71.35 ms | 2.52 ms | 1.46 ms | 1.86 | 1.86 | 92200.0 KB | 7.08 | 1521.0 KB | 1.10 | 86.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table | EPPlus | 349.62 ms | 7.76 ms | 4.48 ms | 9.14 | 9.14 | 117437.6 KB | 9.02 | 1090.8 KB | 0.79 | 813.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table | ClosedXML | 408.08 ms | 51.80 ms | 29.91 ms | 10.66 | 10.66 | 173397.5 KB | 13.32 | 1140.7 KB | 0.82 | 966.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | OfficeIMO.Excel | 40.72 ms | 2.16 ms | 1.25 ms | 1.00 | 1.00 | 13026.6 KB | 1.00 | 1385.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | MiniExcel | 73.24 ms | 1.88 ms | 1.08 ms | 1.80 | 1.80 | 124495.5 KB | 9.56 | 1521.1 KB | 1.10 | 79.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | EPPlus | 380.69 ms | 4.88 ms | 2.82 ms | 9.35 | 9.35 | 159742.2 KB | 12.26 | 1091.0 KB | 0.79 | 834.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | ClosedXML | 791.18 ms | 9.56 ms | 5.52 ms | 19.43 | 19.43 | 566142.3 KB | 43.46 | 1140.9 KB | 0.82 | 1842.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | OfficeIMO.Excel | 33.90 ms | 1.92 ms | 1.11 ms | 1.00 | 1.00 | 9819.7 KB | 1.00 | 1329.2 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | LargeXlsx | 37.40 ms | 2.06 ms | 1.19 ms | 1.10 | 1.10 | 9265.9 KB | 0.94 | 1680.0 KB | 1.26 | 10.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | MiniExcel | 100.47 ms | 2.80 ms | 1.61 ms | 2.96 | 2.96 | 108129.1 KB | 11.01 | 1819.7 KB | 1.37 | 196.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | EPPlus | 497.14 ms | 6.93 ms | 4.00 ms | 14.66 | 14.66 | 135724.0 KB | 13.82 | 1390.4 KB | 1.05 | 1366.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | ClosedXML | 535.38 ms | 6.25 ms | 3.61 ms | 15.79 | 15.79 | 280372.9 KB | 28.55 | 1519.9 KB | 1.14 | 1479.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | OfficeIMO.Excel | 41.08 ms | 1.57 ms | 0.91 ms | 1.00 | 1.00 | 13458.5 KB | 1.00 | 1795.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | MiniExcel | 99.02 ms | 2.86 ms | 1.65 ms | 2.41 | 2.41 | 108129.1 KB | 8.03 | 1819.7 KB | 1.01 | 141.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | EPPlus | 518.34 ms | 14.72 ms | 8.50 ms | 12.62 | 12.62 | 135724.0 KB | 10.08 | 1390.4 KB | 0.77 | 1161.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | ClosedXML | 543.21 ms | 6.94 ms | 4.01 ms | 13.22 | 13.22 | 280371.8 KB | 20.83 | 1519.9 KB | 0.85 | 1222.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables | OfficeIMO.Excel | 36.50 ms | 1.45 ms | 0.84 ms | 1.00 | 1.00 | 10288.1 KB | 1.00 | 1376.4 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-dataset-tables | MiniExcel | 84.67 ms | 7.87 ms | 4.55 ms | 2.32 | 2.32 | 97085.4 KB | 9.44 | 1511.8 KB | 1.10 | 132.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables | EPPlus | 349.91 ms | 8.34 ms | 4.82 ms | 9.59 | 9.59 | 110816.3 KB | 10.77 | 1100.6 KB | 0.80 | 858.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables | ClosedXML | 372.66 ms | 16.42 ms | 9.48 ms | 10.21 | 10.21 | 172003.7 KB | 16.72 | 1139.0 KB | 0.83 | 921.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | OfficeIMO.Excel | 38.99 ms | 2.23 ms | 1.28 ms | 1.00 | 1.00 | 10300.7 KB | 1.00 | 1376.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | MiniExcel | 83.17 ms | 0.99 ms | 0.57 ms | 2.13 | 2.13 | 128874.9 KB | 12.51 | 1512.0 KB | 1.10 | 113.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | EPPlus | 405.08 ms | 9.74 ms | 5.62 ms | 10.39 | 10.39 | 195408.4 KB | 18.97 | 1100.9 KB | 0.80 | 938.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | ClosedXML | 742.48 ms | 26.18 ms | 15.11 ms | 19.04 | 19.04 | 550095.1 KB | 53.40 | 1139.3 KB | 0.83 | 1804.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-direct | LargeXlsx | 30.82 ms | 0.75 ms | 0.43 ms | 0.88 | 1.00 | 9520.4 KB | 0.75 | 1386.5 KB | 1.00 | 12.3% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-direct | OfficeIMO.Excel | 35.16 ms | 1.81 ms | 1.04 ms | 1.00 | 1.14 | 12715.7 KB | 1.00 | 1384.9 KB | 1.00 | Loss +14.1% |
| 25000 | package-profile | package | Package size | write-datatable-direct | MiniExcel | 76.69 ms | 0.81 ms | 0.47 ms | 2.18 | 2.49 | 92394.2 KB | 7.27 | 1521.1 KB | 1.10 | 118.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-direct | ClosedXML | 282.73 ms | 2.46 ms | 1.42 ms | 8.04 | 9.17 | 104205.0 KB | 8.19 | 1139.9 KB | 0.82 | 704.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-direct | EPPlus | 336.46 ms | 9.58 ms | 5.53 ms | 9.57 | 10.92 | 117437.6 KB | 9.24 | 1090.8 KB | 0.79 | 857.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | OfficeIMO.Excel | 35.23 ms | 1.10 ms | 0.64 ms | 1.00 | 1.00 | 12733.8 KB | 1.00 | 1385.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | MiniExcel | 79.22 ms | 1.49 ms | 0.86 ms | 2.25 | 2.25 | 92394.5 KB | 7.26 | 1521.1 KB | 1.10 | 124.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | EPPlus | 343.84 ms | 7.30 ms | 4.21 ms | 9.76 | 9.76 | 117437.6 KB | 9.22 | 1090.8 KB | 0.79 | 876.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | ClosedXML | 384.98 ms | 0.92 ms | 0.53 ms | 10.93 | 10.93 | 173402.7 KB | 13.62 | 1140.7 KB | 0.82 | 992.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | LargeXlsx | 28.33 ms | 0.74 ms | 0.43 ms | 0.87 | 1.00 | 5614.1 KB | 0.43 | 1386.5 KB | 1.00 | 12.6% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 32.40 ms | 0.77 ms | 0.44 ms | 1.00 | 1.14 | 12912.0 KB | 1.00 | 1384.9 KB | 1.00 | Loss +14.4% |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | MiniExcel | 64.27 ms | 2.78 ms | 1.61 ms | 1.98 | 2.27 | 93257.0 KB | 7.22 | 1521.1 KB | 1.10 | 98.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | ClosedXML | 275.05 ms | 7.25 ms | 4.18 ms | 8.49 | 9.71 | 104205.0 KB | 8.07 | 1139.9 KB | 0.82 | 748.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | EPPlus | 339.85 ms | 5.38 ms | 3.10 ms | 10.49 | 12.00 | 117438.0 KB | 9.10 | 1090.8 KB | 0.79 | 948.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 36.13 ms | 2.13 ms | 1.23 ms | 1.00 | 1.00 | 11501.6 KB | 1.00 | 1428.5 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 375.91 ms | 13.49 ms | 7.79 ms | 10.40 | 10.40 | 159742.5 KB | 13.89 | 1091.0 KB | 0.76 | 940.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 689.41 ms | 19.19 ms | 11.08 ms | 19.08 | 19.08 | 496956.9 KB | 43.21 | 1140.1 KB | 0.80 | 1808.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | OfficeIMO.Excel | 32.30 ms | 1.27 ms | 0.74 ms | 1.00 | 1.00 | 11493.8 KB | 1.00 | 1428.4 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | LargeXlsx | 41.68 ms | 23.60 ms | 13.63 ms | 1.29 | 1.29 | 5614.1 KB | 0.49 | 1386.5 KB | 0.97 | 29.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | MiniExcel | 65.64 ms | 3.19 ms | 1.84 ms | 2.03 | 2.03 | 93257.0 KB | 8.11 | 1521.1 KB | 1.06 | 103.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | ClosedXML | 275.34 ms | 12.29 ms | 7.10 ms | 8.52 | 8.52 | 104205.0 KB | 9.07 | 1139.9 KB | 0.80 | 752.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | EPPlus | 348.44 ms | 25.55 ms | 14.75 ms | 10.79 | 10.79 | 117437.6 KB | 10.22 | 1090.8 KB | 0.76 | 978.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 40.38 ms | 1.39 ms | 0.80 ms | 1.00 | 1.00 | 10187.2 KB | 1.00 | 1385.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 353.09 ms | 6.56 ms | 3.79 ms | 8.74 | 8.74 | 159742.5 KB | 15.68 | 1091.0 KB | 0.79 | 774.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 646.00 ms | 10.01 ms | 5.78 ms | 16.00 | 16.00 | 496956.9 KB | 48.78 | 1140.1 KB | 0.82 | 1499.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 29.20 ms | 3.51 ms | 2.03 ms | 0.78 | 1.00 | 5614.1 KB | 0.55 | 1386.5 KB | 1.00 | 22.5% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 37.68 ms | 2.86 ms | 1.65 ms | 1.00 | 1.29 | 10179.4 KB | 1.00 | 1384.9 KB | 1.00 | Loss +29.0% |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | MiniExcel | 62.43 ms | 0.84 ms | 0.49 ms | 1.66 | 2.14 | 93257.0 KB | 9.16 | 1521.1 KB | 1.10 | 65.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | ClosedXML | 264.44 ms | 3.54 ms | 2.05 ms | 7.02 | 9.06 | 104205.0 KB | 10.24 | 1139.9 KB | 0.82 | 601.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | EPPlus | 324.52 ms | 6.55 ms | 3.78 ms | 8.61 | 11.11 | 117437.6 KB | 11.54 | 1090.8 KB | 0.79 | 761.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | LargeXlsx | 27.64 ms | 0.35 ms | 0.20 ms | 0.66 | 1.00 | 5614.1 KB | 0.36 | 1386.5 KB | 0.97 | 34.2% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | OfficeIMO.Excel | 42.03 ms | 1.82 ms | 1.05 ms | 1.00 | 1.52 | 15791.7 KB | 1.00 | 1428.4 KB | 1.00 | Loss +52.0% |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | MiniExcel | 64.31 ms | 0.96 ms | 0.55 ms | 1.53 | 2.33 | 93257.0 KB | 5.91 | 1521.1 KB | 1.06 | 53.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | ClosedXML | 284.90 ms | 20.28 ms | 11.71 ms | 6.78 | 10.31 | 104205.0 KB | 6.60 | 1139.9 KB | 0.80 | 577.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | EPPlus | 340.14 ms | 13.81 ms | 7.97 ms | 8.09 | 12.31 | 117437.6 KB | 7.44 | 1090.8 KB | 0.76 | 709.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 33.96 ms | 0.94 ms | 0.54 ms | 1.00 | 1.00 | 11500.9 KB | 1.00 | 1428.5 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 342.32 ms | 7.24 ms | 4.18 ms | 10.08 | 10.08 | 138360.7 KB | 12.03 | 1091.0 KB | 0.76 | 908.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 415.95 ms | 16.36 ms | 9.44 ms | 12.25 | 12.25 | 275422.3 KB | 23.95 | 1140.1 KB | 0.80 | 1124.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | LargeXlsx | 41.35 ms | 2.77 ms | 1.60 ms | 0.92 | 1.00 | 6043.9 KB | 0.57 | 1816.3 KB | 0.99 | 8.2% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 45.07 ms | 2.41 ms | 1.39 ms | 1.00 | 1.09 | 10577.2 KB | 1.00 | 1828.0 KB | 1.00 | Loss +9.0% |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | MiniExcel | 84.17 ms | 4.85 ms | 2.80 ms | 1.87 | 2.04 | 113974.3 KB | 10.78 | 1936.7 KB | 1.06 | 86.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | ClosedXML | 373.70 ms | 5.92 ms | 3.42 ms | 8.29 | 9.04 | 179552.5 KB | 16.98 | 1555.2 KB | 0.85 | 729.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | EPPlus | 456.24 ms | 20.48 ms | 11.82 ms | 10.12 | 11.03 | 144920.3 KB | 13.70 | 1473.0 KB | 0.81 | 912.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | LargeXlsx | 36.93 ms | 0.23 ms | 0.13 ms | 0.87 | 1.00 | 6043.9 KB | 0.61 | 1816.3 KB | 0.99 | 12.8% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 42.36 ms | 0.23 ms | 0.13 ms | 1.00 | 1.15 | 9942.2 KB | 1.00 | 1828.0 KB | 1.00 | Loss +14.7% |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | MiniExcel | 79.53 ms | 2.19 ms | 1.27 ms | 1.88 | 2.15 | 113974.3 KB | 11.46 | 1936.7 KB | 1.06 | 87.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | ClosedXML | 363.78 ms | 5.48 ms | 3.16 ms | 8.59 | 9.85 | 179552.5 KB | 18.06 | 1555.2 KB | 0.85 | 758.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | EPPlus | 442.11 ms | 10.43 ms | 6.02 ms | 10.44 | 11.97 | 144920.3 KB | 14.58 | 1473.0 KB | 0.81 | 943.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 188.19 ms | 3.01 ms | 1.74 ms | 1.00 | 1.00 | 36150.1 KB | 1.00 | 6725.6 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | LargeXlsx | 207.43 ms | 4.77 ms | 2.75 ms | 1.10 | 1.10 | 23211.4 KB | 0.64 | 6614.8 KB | 0.98 | 10.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | MiniExcel | 326.91 ms | 8.78 ms | 5.07 ms | 1.74 | 1.74 | 347925.7 KB | 9.62 | 6949.8 KB | 1.03 | 73.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | ClosedXML | 1171.47 ms | 22.72 ms | 13.11 ms | 6.23 | 6.23 | 487446.6 KB | 13.48 | 6165.9 KB | 0.92 | 522.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | EPPlus | 1483.62 ms | 36.18 ms | 20.89 ms | 7.88 | 7.88 | 562916.4 KB | 15.57 | 5441.6 KB | 0.81 | 688.4% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | OfficeIMO.Excel | 109.90 ms | 13.56 ms | 7.83 ms | 1.00 | 1.00 | 15708.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | EPPlus 4.5.3.3 | 630.47 ms |  |  | 5.74 | 5.74 |  |  |  |  | 473.7% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | EPPlus | 673.41 ms | 43.64 ms | 25.20 ms | 6.13 | 6.13 | 250948.6 KB | 15.98 |  |  | 512.7% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | ClosedXML | 2596.79 ms | 613.44 ms | 354.17 ms | 23.63 | 23.63 | 829860.0 KB | 52.83 |  |  | 2262.8% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus 4.5.3.3 | 243.01 ms |  |  | 0.84 | 1.00 |  |  |  |  | 16.1% faster than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | OfficeIMO.Excel | 289.71 ms | 17.43 ms | 10.06 ms | 1.00 | 1.19 | 133435.3 KB | 1.00 |  |  | Loss +19.2% |
| 25000 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus | 439.16 ms | 11.70 ms | 6.76 ms | 1.52 | 1.81 | 234783.8 KB | 1.76 |  |  | 51.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | OfficeIMO.Excel | 89.72 ms | 5.38 ms | 3.10 ms | 1.00 | 1.00 | 43560.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus | 438.47 ms | 22.78 ms | 13.15 ms | 4.89 | 4.89 | 277077.7 KB | 6.36 |  |  | 388.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus 4.5.3.3 | 467.21 ms |  |  | 5.21 | 5.21 |  |  |  |  | 420.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | OfficeIMO.Excel | 306.11 ms | 44.37 ms | 25.62 ms | 1.00 | 1.00 | 144825.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus | 461.40 ms | 19.73 ms | 11.39 ms | 1.51 | 1.51 | 302761.2 KB | 2.09 |  |  | 50.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus 4.5.3.3 | 517.68 ms |  |  | 1.69 | 1.69 |  |  |  |  | 69.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | OfficeIMO.Excel | 258.71 ms | 29.67 ms | 17.13 ms | 1.00 | 1.00 | 133461.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus | 447.48 ms | 8.20 ms | 4.73 ms | 1.73 | 1.73 | 277078.9 KB | 2.08 |  |  | 73.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus 4.5.3.3 | 476.17 ms |  |  | 1.84 | 1.84 |  |  |  |  | 84.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | OfficeIMO.Excel | 286.40 ms | 7.63 ms | 4.41 ms | 1.00 | 1.00 | 133506.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus | 460.49 ms | 15.92 ms | 9.19 ms | 1.61 | 1.61 | 277071.7 KB | 2.08 |  |  | 60.8% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus 4.5.3.3 | 483.91 ms |  |  | 1.69 | 1.69 |  |  |  |  | 69.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Object projection | build-object-datatable-dictionaries | OfficeIMO.Excel | 9.68 ms | 0.72 ms | 0.42 ms | 1.00 | 1.00 | 5164.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Object projection | build-object-datatable-typed | OfficeIMO.Excel | 9.18 ms | 1.46 ms | 0.84 ms | 1.00 | 1.00 | 8093.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-cells | OfficeIMO.Excel | 49.67 ms | 4.57 ms | 2.64 ms | 1.00 | 1.00 | 24531.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-cells | EPPlus | 271.74 ms | 29.07 ms | 16.78 ms | 5.47 | 5.47 | 187393.3 KB | 7.64 |  |  | 447.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-cells | ClosedXML | 323.82 ms | 10.06 ms | 5.81 ms | 6.52 | 6.52 | 166521.0 KB | 6.79 |  |  | 551.9% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | OfficeIMO.Excel | 34.87 ms | 1.20 ms | 0.69 ms | 1.00 | 1.00 | 3839.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | EPPlus | 240.57 ms | 7.29 ms | 4.21 ms | 6.90 | 6.90 | 115541.7 KB | 30.09 |  |  | 589.9% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | ClosedXML | 325.60 ms | 14.65 ms | 8.46 ms | 9.34 | 9.34 | 150901.1 KB | 39.30 |  |  | 833.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-range | OfficeIMO.Excel | 47.32 ms | 1.53 ms | 0.88 ms | 1.00 | 1.00 | 24531.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-range | EPPlus | 274.79 ms | 26.82 ms | 15.48 ms | 5.81 | 5.81 | 187393.3 KB | 7.64 |  |  | 480.8% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-range | ClosedXML | 346.83 ms | 17.34 ms | 10.01 ms | 7.33 | 7.33 | 166525.6 KB | 6.79 |  |  | 633.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-top-range | OfficeIMO.Excel | 0.66 ms | 0.07 ms | 0.04 ms | 1.00 | 1.00 | 285.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-top-range | EPPlus | 233.46 ms | 17.41 ms | 10.05 ms | 353.07 | 353.07 | 105580.2 KB | 369.97 |  |  | 35206.8% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-top-range | ClosedXML | 316.82 ms | 10.28 ms | 5.93 ms | 479.13 | 479.13 | 149402.5 KB | 523.54 |  |  | 47812.8% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-autofilter | OfficeIMO.Excel | 31.43 ms | 0.63 ms | 0.37 ms | 1.00 | 1.00 | 11494.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-autofilter | EPPlus 4.5.3.3 | 233.13 ms |  |  | 7.42 | 7.42 |  |  |  |  | 641.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-autofilter | ClosedXML | 287.67 ms | 4.27 ms | 2.47 ms | 9.15 | 9.15 | 210663.8 KB | 18.33 |  |  | 815.2% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-autofilter | EPPlus | 340.51 ms | 7.54 ms | 4.35 ms | 10.83 | 10.83 | 211871.8 KB | 18.43 |  |  | 983.4% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-charts | OfficeIMO.Excel | 32.55 ms | 0.16 ms | 0.09 ms | 1.00 | 1.00 | 12554.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-charts | EPPlus 4.5.3.3 | 247.47 ms |  |  | 7.60 | 7.60 |  |  |  |  | 660.2% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-charts | EPPlus | 359.42 ms | 10.03 ms | 5.79 ms | 11.04 | 11.04 | 214906.2 KB | 17.12 |  |  | 1004.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | OfficeIMO.Excel | 33.10 ms | 1.25 ms | 0.72 ms | 1.00 | 1.00 | 11560.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | EPPlus 4.5.3.3 | 244.45 ms |  |  | 7.39 | 7.39 |  |  |  |  | 638.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | ClosedXML | 312.23 ms | 9.79 ms | 5.65 ms | 9.43 | 9.43 | 210711.7 KB | 18.23 |  |  | 843.4% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | EPPlus | 367.28 ms | 3.09 ms | 1.78 ms | 11.10 | 11.10 | 211913.3 KB | 18.33 |  |  | 1009.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-data-validation | OfficeIMO.Excel | 32.92 ms | 2.28 ms | 1.32 ms | 1.00 | 1.00 | 11510.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-data-validation | EPPlus 4.5.3.3 | 258.20 ms |  |  | 7.84 | 7.84 |  |  |  |  | 684.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-data-validation | ClosedXML | 310.28 ms | 14.67 ms | 8.47 ms | 9.42 | 9.42 | 210672.7 KB | 18.30 |  |  | 842.4% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-data-validation | EPPlus | 364.82 ms | 10.87 ms | 6.28 ms | 11.08 | 11.08 | 211857.8 KB | 18.41 |  |  | 1008.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | OfficeIMO.Excel | 31.83 ms | 1.99 ms | 1.15 ms | 1.00 | 1.00 | 11497.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | EPPlus 4.5.3.3 | 274.84 ms |  |  | 8.64 | 8.64 |  |  |  |  | 763.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | ClosedXML | 289.98 ms | 7.05 ms | 4.07 ms | 9.11 | 9.11 | 210646.8 KB | 18.32 |  |  | 811.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | EPPlus | 342.74 ms | 4.23 ms | 2.44 ms | 10.77 | 10.77 | 211883.7 KB | 18.43 |  |  | 976.9% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | OfficeIMO.Excel | 216.45 ms | 2.38 ms | 1.38 ms | 1.00 | 1.00 | 131922.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | EPPlus 4.5.3.3 | 249.58 ms |  |  | 1.15 | 1.15 |  |  |  |  | 15.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | EPPlus | 381.95 ms | 8.03 ms | 4.63 ms | 1.76 | 1.76 | 230801.7 KB | 1.75 |  |  | 76.5% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | OfficeIMO.Excel | 288.90 ms | 6.24 ms | 3.60 ms | 1.00 | 1.00 | 133445.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus | 472.62 ms | 9.73 ms | 5.62 ms | 1.64 | 1.64 | 277078.9 KB | 2.08 |  |  | 63.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus 4.5.3.3 | 474.61 ms |  |  | 1.64 | 1.64 |  |  |  |  | 64.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world report | realworld-report-core | OfficeIMO.Excel | 35.23 ms | 0.76 ms | 0.44 ms | 1.00 | 1.00 | 11648.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world report | realworld-report-core | EPPlus | 395.54 ms | 16.02 ms | 9.25 ms | 11.23 | 11.23 | 255066.2 KB | 21.90 |  |  | 1022.8% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world report | realworld-report-core | EPPlus 4.5.3.3 | 467.60 ms |  |  | 13.27 | 13.27 |  |  |  |  | 1227.4% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world report | realworld-report-core | ClosedXML | 796.25 ms | 52.10 ms | 30.08 ms | 22.60 | 22.60 | 680116.4 KB | 58.39 |  |  | 2160.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook | OfficeIMO.Excel | 483.21 ms | 23.31 ms | 13.46 ms | 1.00 | 1.00 | 175196.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Report workbook | report-workbook | EPPlus | 612.18 ms | 20.73 ms | 11.97 ms | 1.27 | 1.27 | 364710.2 KB | 2.08 |  |  | 26.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook | EPPlus 4.5.3.3 | 691.69 ms |  |  | 1.43 | 1.43 |  |  |  |  | 43.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-core | OfficeIMO.Excel | 51.79 ms | 2.34 ms | 1.35 ms | 1.00 | 1.00 | 10979.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Report workbook | report-workbook-core | EPPlus | 588.39 ms | 38.94 ms | 22.48 ms | 11.36 | 11.36 | 342842.6 KB | 31.23 |  |  | 1036.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-core | EPPlus 4.5.3.3 | 674.43 ms |  |  | 13.02 | 13.02 |  |  |  |  | 1202.2% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-core | ClosedXML | 1244.70 ms | 19.59 ms | 11.31 ms | 24.03 | 24.03 | 975776.3 KB | 88.87 |  |  | 2303.2% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable | OfficeIMO.Excel | 434.69 ms | 22.98 ms | 13.27 ms | 1.00 | 1.00 | 177944.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable | EPPlus | 556.91 ms | 9.78 ms | 5.65 ms | 1.28 | 1.28 | 247824.2 KB | 1.39 |  |  | 28.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable | EPPlus 4.5.3.3 | 686.85 ms |  |  | 1.58 | 1.58 |  |  |  |  | 58.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable-core | OfficeIMO.Excel | 51.48 ms | 2.34 ms | 1.35 ms | 1.00 | 1.00 | 13725.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable-core | EPPlus | 587.14 ms | 22.95 ms | 13.25 ms | 11.40 | 11.40 | 225956.2 KB | 16.46 |  |  | 1040.4% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable-core | EPPlus 4.5.3.3 | 653.00 ms |  |  | 12.68 | 12.68 |  |  |  |  | 1168.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable-core | ClosedXML | 1144.56 ms | 39.83 ms | 23.00 ms | 22.23 | 22.23 | 832229.7 KB | 60.64 |  |  | 2123.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | OfficeIMO.Excel | 18.36 ms | 0.27 ms | 0.16 ms | 1.00 | 1.00 | 6219.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | EPPlus 4.5.3.3 | 88.07 ms |  |  | 4.80 | 4.80 |  |  |  |  | 379.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | EPPlus | 147.84 ms | 5.65 ms | 3.26 ms | 8.05 | 8.05 | 70814.6 KB | 11.39 |  |  | 705.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | ClosedXML | 171.08 ms | 0.96 ms | 0.55 ms | 9.32 | 9.32 | 79515.7 KB | 12.79 |  |  | 831.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | OfficeIMO.Excel | 0.93 ms | 0.10 ms | 0.06 ms | 1.00 | 1.00 | 177.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | Sylvan.Data.Excel | 1.01 ms | 0.00 ms | 0.00 ms | 1.08 | 1.08 | 316.6 KB | 1.79 |  |  | 8.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | ExcelDataReader | 1.50 ms | 0.18 ms | 0.10 ms | 1.61 | 1.61 | 4062.2 KB | 22.92 |  |  | 60.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | ClosedXML | 4.18 ms | 0.50 ms | 0.29 ms | 4.47 | 4.47 | 4392.9 KB | 24.78 |  |  | 347.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | MiniExcel | 13.50 ms | 1.16 ms | 0.67 ms | 14.46 | 14.46 | 46194.9 KB | 260.63 |  |  | 1345.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | EPPlus 4.5.3.3 | 21.73 ms |  |  | 23.26 | 23.26 |  |  |  |  | 2226.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | EPPlus | 98.85 ms | 3.85 ms | 2.22 ms | 105.83 | 105.83 | 43071.1 KB | 243.01 |  |  | 10483.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | OfficeIMO.Excel | 0.84 ms | 0.06 ms | 0.03 ms | 1.00 | 1.00 | 177.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | Sylvan.Data.Excel | 0.93 ms | 0.01 ms | 0.00 ms | 1.11 | 1.11 | 316.6 KB | 1.79 |  |  | 10.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | ExcelDataReader | 1.43 ms | 0.01 ms | 0.00 ms | 1.70 | 1.70 | 4062.2 KB | 22.91 |  |  | 69.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | ClosedXML | 5.25 ms | 2.62 ms | 1.51 ms | 6.25 | 6.25 | 4392.9 KB | 24.78 |  |  | 524.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | MiniExcel | 13.88 ms | 0.60 ms | 0.35 ms | 16.51 | 16.51 | 46194.9 KB | 260.53 |  |  | 1550.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | EPPlus 4.5.3.3 | 23.31 ms |  |  | 27.71 | 27.71 |  |  |  |  | 2671.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | EPPlus | 98.70 ms | 8.00 ms | 4.62 ms | 117.35 | 117.35 | 43071.1 KB | 242.91 |  |  | 11634.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | Sylvan.Data.Excel | 17.00 ms | 0.59 ms | 0.34 ms | 0.88 | 1.00 | 1936.7 KB | 0.21 |  |  | 12.0% faster than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | OfficeIMO.Excel | 19.32 ms | 1.28 ms | 0.74 ms | 1.00 | 1.14 | 9218.0 KB | 1.00 |  |  | Loss +13.6% |
| 25000 | speed-comparison | read | Other | shared-string-read | ExcelDataReader | 45.78 ms | 2.47 ms | 1.43 ms | 2.37 | 2.69 | 25020.8 KB | 2.71 |  |  | 137.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | MiniExcel | 52.21 ms | 3.32 ms | 1.91 ms | 2.70 | 3.07 | 74405.3 KB | 8.07 |  |  | 170.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | EPPlus 4.5.3.3 | 95.96 ms |  |  | 4.97 | 5.64 |  |  |  |  | 396.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | EPPlus | 146.58 ms | 4.74 ms | 2.74 ms | 7.59 | 8.62 | 89346.1 KB | 9.69 |  |  | 658.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | ClosedXML | 153.00 ms | 2.78 ms | 1.61 ms | 7.92 | 9.00 | 90414.8 KB | 9.81 |  |  | 692.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | OfficeIMO.Excel | 32.61 ms | 0.44 ms | 0.25 ms | 1.00 | 1.00 | 1122.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | Sylvan.Data.Excel | 39.34 ms | 0.66 ms | 0.38 ms | 1.21 | 1.21 | 3534.8 KB | 3.15 |  |  | 20.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | ExcelDataReader | 116.04 ms | 6.20 ms | 3.58 ms | 3.56 | 3.56 | 61201.9 KB | 54.53 |  |  | 255.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | MiniExcel | 135.54 ms | 8.41 ms | 4.86 ms | 4.16 | 4.16 | 186420.9 KB | 166.08 |  |  | 315.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | EPPlus | 244.19 ms | 18.11 ms | 10.45 ms | 7.49 | 7.49 | 105609.1 KB | 94.09 |  |  | 648.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | ClosedXML | 363.10 ms | 40.96 ms | 23.65 ms | 11.14 | 11.14 | 149387.3 KB | 133.09 |  |  | 1013.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | Sylvan.Data.Excel | 64.53 ms | 7.32 ms | 4.23 ms | 0.96 | 1.00 | 18394.2 KB | 0.53 |  |  | 4.3% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | OfficeIMO.Excel | 67.43 ms | 11.25 ms | 6.49 ms | 1.00 | 1.04 | 34645.9 KB | 1.00 |  |  | Loss +4.5% |
| 25000 | speed-comparison | read | Range and table read | read-datatable | ExcelDataReader | 164.53 ms | 29.36 ms | 16.95 ms | 2.44 | 2.55 | 76061.4 KB | 2.20 |  |  | 144.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | MiniExcel | 178.36 ms | 23.00 ms | 13.28 ms | 2.65 | 2.76 | 181285.0 KB | 5.23 |  |  | 164.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | EPPlus | 306.76 ms | 29.85 ms | 17.24 ms | 4.55 | 4.75 | 202250.3 KB | 5.84 |  |  | 355.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | ClosedXML | 401.63 ms | 8.55 ms | 4.93 ms | 5.96 | 6.22 | 178450.6 KB | 5.15 |  |  | 495.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | EPPlus 4.5.3.3 | 410.01 ms |  |  | 6.08 | 6.35 |  |  |  |  | 508.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | OfficeIMO.Excel | 34.81 ms | 1.23 ms | 0.71 ms | 1.00 | 1.00 | 4034.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | Sylvan.Data.Excel | 45.36 ms | 3.30 ms | 1.91 ms | 1.30 | 1.30 | 4316.2 KB | 1.07 |  |  | 30.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | MiniExcel | 100.80 ms | 10.47 ms | 6.04 ms | 2.90 | 2.90 | 158612.9 KB | 39.31 |  |  | 189.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | ExcelDataReader | 108.24 ms | 4.57 ms | 2.64 ms | 3.11 | 3.11 | 61201.9 KB | 15.17 |  |  | 210.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | EPPlus | 230.17 ms | 11.23 ms | 6.48 ms | 6.61 | 6.61 | 115541.7 KB | 28.64 |  |  | 561.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | ClosedXML | 324.69 ms | 20.90 ms | 12.07 ms | 9.33 | 9.33 | 150903.0 KB | 37.40 |  |  | 832.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | Sylvan.Data.Excel | 45.83 ms | 4.68 ms | 2.70 ms | 0.89 | 1.00 | 3534.8 KB | 0.14 |  |  | 11.4% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | OfficeIMO.Excel | 51.75 ms | 1.44 ms | 0.83 ms | 1.00 | 1.13 | 26098.3 KB | 1.00 |  |  | Loss +12.9% |
| 25000 | speed-comparison | read | Range and table read | read-range | ExcelDataReader | 114.17 ms | 12.41 ms | 7.16 ms | 2.21 | 2.49 | 61201.9 KB | 2.35 |  |  | 120.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | MiniExcel | 127.55 ms | 6.43 ms | 3.71 ms | 2.46 | 2.78 | 186421.5 KB | 7.14 |  |  | 146.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | EPPlus | 253.79 ms | 2.85 ms | 1.64 ms | 4.90 | 5.54 | 187390.9 KB | 7.18 |  |  | 390.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | ClosedXML | 328.36 ms | 16.13 ms | 9.31 ms | 6.35 | 7.16 | 163591.7 KB | 6.27 |  |  | 534.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | EPPlus 4.5.3.3 | 384.90 ms |  |  | 7.44 | 8.40 |  |  |  |  | 643.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | Sylvan.Data.Excel | 53.67 ms | 6.41 ms | 3.70 ms | 0.99 | 1.00 | 4484.9 KB | 0.17 |  |  | Tie vs OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | OfficeIMO.Excel | 54.18 ms | 2.32 ms | 1.34 ms | 1.00 | 1.01 | 26684.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | ExcelDataReader | 116.42 ms | 10.03 ms | 5.79 ms | 2.15 | 2.17 | 61201.9 KB | 2.29 |  |  | 114.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | MiniExcel | 124.69 ms | 7.06 ms | 4.08 ms | 2.30 | 2.32 | 186421.5 KB | 6.99 |  |  | 130.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | EPPlus | 261.30 ms | 31.92 ms | 18.43 ms | 4.82 | 4.87 | 187390.9 KB | 7.02 |  |  | 382.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | ClosedXML | 327.20 ms | 15.24 ms | 8.80 ms | 6.04 | 6.10 | 163586.1 KB | 6.13 |  |  | 503.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | Sylvan.Data.Excel | 0.52 ms | 0.09 ms | 0.05 ms | 0.79 | 1.00 | 348.5 KB | 1.18 |  |  | 21.0% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | OfficeIMO.Excel | 0.66 ms | 0.11 ms | 0.06 ms | 1.00 | 1.27 | 296.1 KB | 1.00 |  |  | Loss +26.6% |
| 25000 | speed-comparison | read | Range and table read | read-top-range | MiniExcel | 0.93 ms | 0.13 ms | 0.07 ms | 1.40 | 1.77 | 869.0 KB | 2.93 |  |  | 39.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | ExcelDataReader | 52.14 ms | 3.97 ms | 2.29 ms | 78.59 | 99.50 | 17115.3 KB | 57.80 |  |  | 7758.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | EPPlus | 328.85 ms | 87.90 ms | 50.75 ms | 495.68 | 627.58 | 105577.8 KB | 356.53 |  |  | 49468.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | ClosedXML | 356.60 ms | 32.57 ms | 18.80 ms | 537.51 | 680.54 | 149390.7 KB | 504.49 |  |  | 53651.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | EPPlus 4.5.3.3 | 570.15 ms |  |  | 859.39 | 1088.06 |  |  |  |  | 85838.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | Sylvan.Data.Excel | 46.15 ms | 4.18 ms | 2.41 ms | 0.47 | 1.00 | 3534.8 KB | 0.10 |  |  | 52.8% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | OfficeIMO.Excel | 97.82 ms | 6.95 ms | 4.01 ms | 1.00 | 2.12 | 34151.8 KB | 1.00 |  |  | Loss +112.0% |
| 25000 | speed-comparison | read | Range and table read | read-used-range | ExcelDataReader | 119.94 ms | 3.73 ms | 2.16 ms | 1.23 | 2.60 | 61201.9 KB | 1.79 |  |  | 22.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | MiniExcel | 126.05 ms | 4.47 ms | 2.58 ms | 1.29 | 2.73 | 186421.5 KB | 5.46 |  |  | 28.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | EPPlus | 262.09 ms | 13.99 ms | 8.08 ms | 2.68 | 5.68 | 187390.9 KB | 5.49 |  |  | 167.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | ClosedXML | 341.54 ms | 25.87 ms | 14.94 ms | 3.49 | 7.40 | 163592.9 KB | 4.79 |  |  | 249.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | OfficeIMO.Excel | 34.14 ms | 1.80 ms | 1.04 ms | 1.00 | 1.00 | 1125.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | Sylvan.Data.Excel | 47.18 ms | 8.11 ms | 4.68 ms | 1.38 | 1.38 | 3534.8 KB | 3.14 |  |  | 38.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | ExcelDataReader | 114.55 ms | 14.24 ms | 8.22 ms | 3.35 | 3.35 | 61201.9 KB | 54.37 |  |  | 235.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | MiniExcel | 126.08 ms | 3.17 ms | 1.83 ms | 3.69 | 3.69 | 186420.9 KB | 165.60 |  |  | 269.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | EPPlus | 239.79 ms | 15.86 ms | 9.16 ms | 7.02 | 7.02 | 105609.1 KB | 93.81 |  |  | 602.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | ClosedXML | 356.22 ms | 20.96 ms | 12.10 ms | 10.43 | 10.43 | 149394.9 KB | 132.71 |  |  | 943.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | Sylvan.Data.Excel | 76.02 ms | 16.18 ms | 9.34 ms | 0.86 | 1.00 | 3534.8 KB | 0.13 |  |  | 14.3% faster than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | OfficeIMO.Excel | 88.69 ms | 18.84 ms | 10.88 ms | 1.00 | 1.17 | 26883.9 KB | 1.00 |  |  | Loss +16.7% |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | MiniExcel | 230.22 ms | 94.99 ms | 54.84 ms | 2.60 | 3.03 | 186421.5 KB | 6.93 |  |  | 159.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | ExcelDataReader | 270.79 ms | 130.14 ms | 75.14 ms | 3.05 | 3.56 | 61201.9 KB | 2.28 |  |  | 205.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | EPPlus 4.5.3.3 | 304.24 ms |  |  | 3.43 | 4.00 |  |  |  |  | 243.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | ClosedXML | 538.78 ms | 123.60 ms | 71.36 ms | 6.08 | 7.09 | 163594.0 KB | 6.09 |  |  | 507.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | EPPlus | 553.94 ms | 135.45 ms | 78.20 ms | 6.25 | 7.29 | 187390.9 KB | 6.97 |  |  | 524.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | OfficeIMO.Excel | 0.61 ms | 0.13 ms | 0.08 ms | 1.00 | 1.00 | 299.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | Sylvan.Data.Excel | 0.62 ms | 0.13 ms | 0.07 ms | 1.00 | 1.00 | 348.5 KB | 1.16 |  |  | Tie vs OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | MiniExcel | 0.77 ms | 0.03 ms | 0.02 ms | 1.25 | 1.25 | 869.0 KB | 2.90 |  |  | 24.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | ExcelDataReader | 42.85 ms | 4.03 ms | 2.33 ms | 69.67 | 69.67 | 17115.3 KB | 57.16 |  |  | 6867.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | EPPlus 4.5.3.3 | 228.05 ms |  |  | 370.81 | 370.81 |  |  |  |  | 36981.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | EPPlus | 248.42 ms | 32.57 ms | 18.81 ms | 403.93 | 403.93 | 105577.8 KB | 352.57 |  |  | 40293.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | ClosedXML | 342.46 ms | 33.56 ms | 19.38 ms | 556.85 | 556.85 | 149392.7 KB | 498.88 |  |  | 55584.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | Sylvan.Data.Excel | 0.58 ms | 0.09 ms | 0.05 ms | 0.80 | 1.00 | 348.5 KB | 1.16 |  |  | 19.8% faster than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | OfficeIMO.Excel | 0.72 ms | 0.16 ms | 0.09 ms | 1.00 | 1.25 | 300.2 KB | 1.00 |  |  | Loss +24.6% |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | MiniExcel | 0.94 ms | 0.10 ms | 0.06 ms | 1.30 | 1.62 | 869.0 KB | 2.89 |  |  | 30.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | ExcelDataReader | 46.85 ms | 8.53 ms | 4.93 ms | 65.27 | 81.34 | 17115.3 KB | 57.02 |  |  | 6426.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | EPPlus | 298.74 ms | 28.67 ms | 16.55 ms | 416.15 | 518.62 | 105577.8 KB | 351.71 |  |  | 41515.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | ClosedXML | 370.70 ms | 37.07 ms | 21.40 ms | 516.40 | 643.55 | 149389.0 KB | 497.65 |  |  | 51539.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | Sylvan.Data.Excel | 45.24 ms | 2.17 ms | 1.25 ms | 0.86 | 1.00 | 5805.0 KB | 0.25 |  |  | 13.7% faster than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | OfficeIMO.Excel | 52.42 ms | 2.29 ms | 1.32 ms | 1.00 | 1.16 | 23562.4 KB | 1.00 |  |  | Loss +15.9% |
| 25000 | speed-comparison | read | Typed object read | read-objects | ExcelDataReader | 117.71 ms | 6.92 ms | 4.00 ms | 2.25 | 2.60 | 63472.1 KB | 2.69 |  |  | 124.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | MiniExcel | 143.83 ms | 12.64 ms | 7.30 ms | 2.74 | 3.18 | 183656.7 KB | 7.79 |  |  | 174.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | EPPlus 4.5.3.3 | 210.03 ms |  |  | 4.01 | 4.64 |  |  |  |  | 300.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | EPPlus | 258.15 ms | 11.05 ms | 6.38 ms | 4.92 | 5.71 | 199608.2 KB | 8.47 |  |  | 392.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | ClosedXML | 316.28 ms | 5.06 ms | 2.92 ms | 6.03 | 6.99 | 165542.1 KB | 7.03 |  |  | 503.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | OfficeIMO.Excel | 113.12 ms | 59.76 ms | 34.50 ms | 1.00 | 1.00 | 23367.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | Sylvan.Data.Excel | 155.32 ms | 56.13 ms | 32.41 ms | 1.37 | 1.37 | 5292.6 KB | 0.23 |  |  | 37.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | EPPlus 4.5.3.3 | 184.71 ms |  |  | 1.63 | 1.63 |  |  |  |  | 63.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | ExcelDataReader | 192.75 ms | 31.43 ms | 18.15 ms | 1.70 | 1.70 | 62959.7 KB | 2.69 |  |  | 70.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | MiniExcel | 339.62 ms | 208.75 ms | 120.52 ms | 3.00 | 3.00 | 183144.4 KB | 7.84 |  |  | 200.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | ClosedXML | 538.91 ms | 179.11 ms | 103.41 ms | 4.76 | 4.76 | 165348.6 KB | 7.08 |  |  | 376.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | EPPlus | 584.59 ms | 308.79 ms | 178.28 ms | 5.17 | 5.17 | 199412.9 KB | 8.53 |  |  | 416.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | OfficeIMO.Excel | 39.05 ms | 0.68 ms | 0.39 ms | 1.00 | 1.00 | 13026.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | MiniExcel | 78.25 ms | 3.93 ms | 2.27 ms | 2.00 | 2.00 | 124495.5 KB | 9.56 |  |  | 100.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | EPPlus | 374.16 ms | 5.68 ms | 3.28 ms | 9.58 | 9.58 | 159742.2 KB | 12.26 |  |  | 858.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | EPPlus 4.5.3.3 | 433.61 ms |  |  | 11.10 | 11.10 |  |  |  |  | 1010.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | ClosedXML | 787.58 ms | 11.50 ms | 6.64 ms | 20.17 | 20.17 | 566142.0 KB | 43.46 |  |  | 1917.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | OfficeIMO.Excel | 38.25 ms | 3.95 ms | 2.28 ms | 1.00 | 1.00 | 10300.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | MiniExcel | 79.15 ms | 4.22 ms | 2.43 ms | 2.07 | 2.07 | 128874.9 KB | 12.51 |  |  | 106.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | EPPlus | 380.13 ms | 5.47 ms | 3.16 ms | 9.94 | 9.94 | 195408.4 KB | 18.97 |  |  | 893.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | ClosedXML | 746.95 ms | 22.88 ms | 13.21 ms | 19.53 | 19.53 | 550095.6 KB | 53.40 |  |  | 1853.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 41.09 ms | 1.37 ms | 0.79 ms | 1.00 | 1.00 | 11501.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 392.24 ms | 16.18 ms | 9.34 ms | 9.55 | 9.55 | 159742.7 KB | 13.89 |  |  | 854.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 777.21 ms | 8.10 ms | 4.67 ms | 18.92 | 18.92 | 496956.9 KB | 43.21 |  |  | 1791.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 43.12 ms | 4.95 ms | 2.86 ms | 1.00 | 1.00 | 10187.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 375.58 ms | 19.92 ms | 11.50 ms | 8.71 | 8.71 | 159742.7 KB | 15.68 |  |  | 770.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 717.46 ms | 8.10 ms | 4.68 ms | 16.64 | 16.64 | 496956.9 KB | 48.78 |  |  | 1563.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 34.88 ms | 3.00 ms | 1.73 ms | 1.00 | 1.00 | 11500.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 421.61 ms | 59.16 ms | 34.15 ms | 12.09 | 12.09 | 138360.7 KB | 12.03 |  |  | 1108.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 434.63 ms | 8.37 ms | 4.83 ms | 12.46 | 12.46 | 275422.3 KB | 23.95 |  |  | 1146.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | OfficeIMO.Excel | 12.51 ms | 1.12 ms | 0.65 ms | 1.00 | 1.00 | 6896.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | ClosedXML | 117.53 ms | 3.07 ms | 1.77 ms | 9.39 | 9.39 | 92902.1 KB | 13.47 |  |  | 839.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | EPPlus | 171.50 ms | 3.81 ms | 2.20 ms | 13.70 | 13.70 | 74493.1 KB | 10.80 |  |  | 1270.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | OfficeIMO.Excel | 15.63 ms | 2.20 ms | 1.27 ms | 1.00 | 1.00 | 5970.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | ClosedXML | 112.58 ms | 2.10 ms | 1.21 ms | 7.20 | 7.20 | 84206.7 KB | 14.10 |  |  | 620.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | EPPlus 4.5.3.3 | 118.17 ms |  |  | 7.56 | 7.56 |  |  |  |  | 655.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | EPPlus | 200.85 ms | 9.84 ms | 5.68 ms | 12.85 | 12.85 | 86377.9 KB | 14.47 |  |  | 1184.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | OfficeIMO.Excel | 26.55 ms | 8.79 ms | 5.07 ms | 1.00 | 1.00 | 8332.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | EPPlus 4.5.3.3 | 155.26 ms |  |  | 5.85 | 5.85 |  |  |  |  | 484.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | ClosedXML | 232.31 ms | 17.78 ms | 10.26 ms | 8.75 | 8.75 | 111118.7 KB | 13.33 |  |  | 775.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | EPPlus | 286.79 ms | 23.14 ms | 13.36 ms | 10.80 | 10.80 | 113245.5 KB | 13.59 |  |  | 980.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | OfficeIMO.Excel | 29.96 ms | 1.95 ms | 1.12 ms | 1.00 | 1.00 | 7416.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | ClosedXML | 221.95 ms | 43.29 ms | 24.99 ms | 7.41 | 7.41 | 105223.9 KB | 14.19 |  |  | 640.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | EPPlus | 363.91 ms | 22.59 ms | 13.04 ms | 12.15 | 12.15 | 106317.3 KB | 14.34 |  |  | 1114.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 22.18 ms | 6.33 ms | 3.65 ms | 1.00 | 1.00 | 7416.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | ClosedXML | 168.72 ms | 5.27 ms | 3.04 ms | 7.61 | 7.61 | 105223.9 KB | 14.19 |  |  | 660.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | EPPlus | 242.52 ms | 38.99 ms | 22.51 ms | 10.94 | 10.94 | 106317.3 KB | 14.34 |  |  | 993.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | OfficeIMO.Excel | 11.28 ms | 0.26 ms | 0.15 ms | 1.00 | 1.00 | 6144.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | ClosedXML | 106.95 ms | 7.23 ms | 4.18 ms | 9.48 | 9.48 | 82591.3 KB | 13.44 |  |  | 848.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | EPPlus 4.5.3.3 | 120.06 ms |  |  | 10.65 | 10.65 |  |  |  |  | 964.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | EPPlus | 203.42 ms | 12.92 ms | 7.46 ms | 18.04 | 18.04 | 85127.8 KB | 13.85 |  |  | 1703.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | OfficeIMO.Excel | 30.82 ms | 0.88 ms | 0.51 ms | 1.00 | 1.00 | 7482.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | EPPlus 4.5.3.3 | 140.73 ms |  |  | 4.57 | 4.57 |  |  |  |  | 356.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | ClosedXML | 243.04 ms | 19.91 ms | 11.49 ms | 7.89 | 7.89 | 89323.7 KB | 11.94 |  |  | 688.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | EPPlus | 301.75 ms | 63.43 ms | 36.62 ms | 9.79 | 9.79 | 103800.4 KB | 13.87 |  |  | 879.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | OfficeIMO.Excel | 75.48 ms | 29.06 ms | 16.78 ms | 1.00 | 1.00 | 13039.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | MiniExcel | 148.96 ms | 55.61 ms | 32.11 ms | 1.97 | 1.97 | 97088.3 KB | 7.45 |  |  | 97.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | ClosedXML | 636.53 ms | 95.20 ms | 54.96 ms | 8.43 | 8.43 | 172019.1 KB | 13.19 |  |  | 743.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | EPPlus | 827.09 ms | 192.77 ms | 111.30 ms | 10.96 | 10.96 | 111246.3 KB | 8.53 |  |  | 995.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | OfficeIMO.Excel | 68.65 ms | 26.84 ms | 15.50 ms | 1.00 | 1.00 | 13458.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | MiniExcel | 157.28 ms | 20.80 ms | 12.01 ms | 2.29 | 2.29 | 108129.1 KB | 8.03 |  |  | 129.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | EPPlus | 909.51 ms | 86.52 ms | 49.95 ms | 13.25 | 13.25 | 135724.0 KB | 10.08 |  |  | 1224.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | ClosedXML | 945.16 ms | 273.18 ms | 157.72 ms | 13.77 | 13.77 | 280371.8 KB | 20.83 |  |  | 1276.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | OfficeIMO.Excel | 33.25 ms | 0.20 ms | 0.12 ms | 1.00 | 1.00 | 10288.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | MiniExcel | 78.32 ms | 5.94 ms | 3.43 ms | 2.36 | 2.36 | 97085.4 KB | 9.44 |  |  | 135.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | EPPlus 4.5.3.3 | 210.73 ms |  |  | 6.34 | 6.34 |  |  |  |  | 533.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | EPPlus | 329.90 ms | 5.52 ms | 3.19 ms | 9.92 | 9.92 | 110816.3 KB | 10.77 |  |  | 892.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | ClosedXML | 340.81 ms | 3.05 ms | 1.76 ms | 10.25 | 10.25 | 171999.1 KB | 16.72 |  |  | 925.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | OfficeIMO.Excel | 39.84 ms | 6.45 ms | 3.73 ms | 1.00 | 1.00 | 13020.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | MiniExcel | 70.82 ms | 3.30 ms | 1.90 ms | 1.78 | 1.78 | 92200.0 KB | 7.08 |  |  | 77.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | EPPlus 4.5.3.3 | 215.71 ms |  |  | 5.41 | 5.41 |  |  |  |  | 441.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | EPPlus | 341.89 ms | 20.61 ms | 11.90 ms | 8.58 | 8.58 | 117437.6 KB | 9.02 |  |  | 758.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | ClosedXML | 374.21 ms | 20.15 ms | 11.63 ms | 9.39 | 9.39 | 173398.1 KB | 13.32 |  |  | 839.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | LargeXlsx | 31.41 ms | 1.35 ms | 0.78 ms | 0.90 | 1.00 | 9520.4 KB | 0.75 |  |  | 9.7% faster than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | OfficeIMO.Excel | 34.78 ms | 1.90 ms | 1.10 ms | 1.00 | 1.11 | 12715.7 KB | 1.00 |  |  | Loss +10.7% |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | MiniExcel | 75.79 ms | 0.34 ms | 0.20 ms | 2.18 | 2.41 | 92394.2 KB | 7.27 |  |  | 117.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | EPPlus 4.5.3.3 | 214.19 ms |  |  | 6.16 | 6.82 |  |  |  |  | 515.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | ClosedXML | 267.94 ms | 10.21 ms | 5.90 ms | 7.70 | 8.53 | 104205.0 KB | 8.19 |  |  | 670.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | EPPlus | 320.36 ms | 7.36 ms | 4.25 ms | 9.21 | 10.20 | 117437.6 KB | 9.24 |  |  | 821.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | OfficeIMO.Excel | 37.25 ms | 3.01 ms | 1.74 ms | 1.00 | 1.00 | 9999.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | MiniExcel | 81.26 ms | 7.11 ms | 4.11 ms | 2.18 | 2.18 | 89659.2 KB | 8.97 |  |  | 118.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | EPPlus | 317.89 ms | 3.97 ms | 2.29 ms | 8.53 | 8.53 | 114703.4 KB | 11.47 |  |  | 753.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | ClosedXML | 374.24 ms | 8.53 ms | 4.93 ms | 10.05 | 10.05 | 170666.2 KB | 17.07 |  |  | 904.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | OfficeIMO.Excel | 34.52 ms | 0.41 ms | 0.24 ms | 1.00 | 1.00 | 12733.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | MiniExcel | 81.29 ms | 2.29 ms | 1.32 ms | 2.35 | 2.35 | 92394.5 KB | 7.26 |  |  | 135.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | EPPlus 4.5.3.3 | 216.54 ms |  |  | 6.27 | 6.27 |  |  |  |  | 527.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | EPPlus | 338.12 ms | 22.67 ms | 13.09 ms | 9.79 | 9.79 | 117437.6 KB | 9.22 |  |  | 879.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | ClosedXML | 368.22 ms | 13.36 ms | 7.71 ms | 10.67 | 10.67 | 173395.0 KB | 13.62 |  |  | 966.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | OfficeIMO.Excel | 38.05 ms | 5.45 ms | 3.15 ms | 1.00 | 1.00 | 11561.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | MiniExcel | 66.51 ms | 0.65 ms | 0.38 ms | 1.75 | 1.75 | 125551.5 KB | 10.86 |  |  | 74.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | EPPlus | 401.29 ms | 28.33 ms | 16.35 ms | 10.55 | 10.55 | 254959.4 KB | 22.05 |  |  | 954.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | EPPlus 4.5.3.3 | 501.28 ms |  |  | 13.17 | 13.17 |  |  |  |  | 1217.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | ClosedXML | 775.84 ms | 17.58 ms | 10.15 ms | 20.39 | 20.39 | 565950.2 KB | 48.95 |  |  | 1938.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | OfficeIMO.Excel | 20.47 ms | 1.64 ms | 0.95 ms | 1.00 | 1.00 | 10112.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | ClosedXML | 184.73 ms | 9.78 ms | 5.65 ms | 9.02 | 9.02 | 113853.5 KB | 11.26 |  |  | 802.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | EPPlus 4.5.3.3 | 197.28 ms |  |  | 9.64 | 9.64 |  |  |  |  | 863.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | EPPlus | 341.28 ms | 12.69 ms | 7.33 ms | 16.67 | 16.67 | 140732.3 KB | 13.92 |  |  | 1567.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-dictionary-objects-table-direct | OfficeIMO.Excel | 50.66 ms | 5.52 ms | 3.19 ms | 1.00 | 1.00 | 15163.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | LargeXlsx | 41.56 ms | 1.61 ms | 0.93 ms | 0.94 | 1.00 | 6043.9 KB | 0.57 |  |  | 6.3% faster than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 44.33 ms | 0.95 ms | 0.55 ms | 1.00 | 1.07 | 10577.2 KB | 1.00 |  |  | Loss +6.7% |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | MiniExcel | 86.31 ms | 2.30 ms | 1.33 ms | 1.95 | 2.08 | 113974.3 KB | 10.78 |  |  | 94.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | ClosedXML | 391.17 ms | 12.16 ms | 7.02 ms | 8.82 | 9.41 | 179552.5 KB | 16.98 |  |  | 782.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | EPPlus | 459.50 ms | 17.23 ms | 9.95 ms | 10.37 | 11.06 | 144920.3 KB | 13.70 |  |  | 936.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | LargeXlsx | 38.84 ms | 1.14 ms | 0.66 ms | 0.87 | 1.00 | 6043.9 KB | 0.61 |  |  | 12.9% faster than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 44.61 ms | 1.74 ms | 1.01 ms | 1.00 | 1.15 | 9942.2 KB | 1.00 |  |  | Loss +14.8% |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | MiniExcel | 83.41 ms | 3.80 ms | 2.19 ms | 1.87 | 2.15 | 113974.3 KB | 11.46 |  |  | 87.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | ClosedXML | 382.51 ms | 3.09 ms | 1.78 ms | 8.57 | 9.85 | 179552.5 KB | 18.06 |  |  | 757.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | EPPlus | 455.91 ms | 19.78 ms | 11.42 ms | 10.22 | 11.74 | 144920.3 KB | 14.58 |  |  | 922.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 198.92 ms | 9.78 ms | 5.65 ms | 1.00 | 1.00 | 36150.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | LargeXlsx | 217.02 ms | 6.55 ms | 3.78 ms | 1.09 | 1.09 | 23211.4 KB | 0.64 |  |  | 9.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | MiniExcel | 337.07 ms | 5.73 ms | 3.31 ms | 1.69 | 1.69 | 347925.7 KB | 9.62 |  |  | 69.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | ClosedXML | 1192.66 ms | 4.37 ms | 2.52 ms | 6.00 | 6.00 | 487446.6 KB | 13.48 |  |  | 499.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | EPPlus | 1523.36 ms | 28.36 ms | 16.37 ms | 7.66 | 7.66 | 562916.4 KB | 15.57 |  |  | 665.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | LargeXlsx | 11.27 ms | 0.62 ms | 0.36 ms | 0.70 | 1.00 | 2771.0 KB | 0.26 |  |  | 29.6% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | OfficeIMO.Excel | 16.00 ms | 1.92 ms | 1.11 ms | 1.00 | 1.42 | 10842.5 KB | 1.00 |  |  | Loss +42.0% |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | MiniExcel | 33.04 ms | 2.12 ms | 1.23 ms | 2.06 | 2.93 | 58242.9 KB | 5.37 |  |  | 106.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | ClosedXML | 137.30 ms | 11.00 ms | 6.35 ms | 8.58 | 12.18 | 104233.1 KB | 9.61 |  |  | 757.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | EPPlus | 206.71 ms | 17.71 ms | 10.23 ms | 12.92 | 18.34 | 100373.9 KB | 9.26 |  |  | 1191.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | EPPlus 4.5.3.3 | 209.43 ms |  |  | 13.09 | 18.58 |  |  |  |  | 1208.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 10.09 ms | 1.21 ms | 0.70 ms | 0.90 | 1.00 | 3444.4 KB | 0.49 |  |  | 9.7% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 11.18 ms | 1.61 ms | 0.93 ms | 1.00 | 1.11 | 6961.7 KB | 1.00 |  |  | Loss +10.8% |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | ClosedXML | 122.21 ms | 5.32 ms | 3.07 ms | 10.93 | 12.11 | 96015.7 KB | 13.79 |  |  | 993.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | EPPlus | 183.59 ms | 4.65 ms | 2.69 ms | 16.42 | 18.19 | 87467.5 KB | 12.56 |  |  | 1542.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | LargeXlsx | 31.20 ms | 4.80 ms | 2.77 ms | 0.87 | 1.00 | 5614.1 KB | 0.35 |  |  | 13.5% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 36.05 ms | 3.39 ms | 1.96 ms | 1.00 | 1.16 | 16036.5 KB | 1.00 |  |  | Loss +15.6% |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | MiniExcel | 68.11 ms | 4.99 ms | 2.88 ms | 1.89 | 2.18 | 93257.0 KB | 5.82 |  |  | 88.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 230.16 ms |  |  | 6.38 | 7.38 |  |  |  |  | 538.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | ClosedXML | 303.69 ms | 2.87 ms | 1.66 ms | 8.42 | 9.73 | 210646.1 KB | 13.14 |  |  | 742.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | EPPlus | 359.14 ms | 5.30 ms | 3.06 ms | 9.96 | 11.51 | 211850.3 KB | 13.21 |  |  | 896.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | OfficeIMO.Excel | 15.77 ms | 0.30 ms | 0.17 ms | 1.00 | 1.00 | 7866.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | ClosedXML | 147.00 ms | 5.94 ms | 3.43 ms | 9.32 | 9.32 | 105223.9 KB | 13.38 |  |  | 832.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | EPPlus | 207.24 ms | 0.65 ms | 0.38 ms | 13.14 | 13.14 | 106317.3 KB | 13.52 |  |  | 1214.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | Sylvan.Data.Excel | 27.72 ms | 0.30 ms | 0.18 ms | 0.75 | 1.00 | 5700.3 KB | 0.44 |  |  | 25.3% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | LargeXlsx | 35.56 ms | 0.86 ms | 0.50 ms | 0.96 | 1.28 | 8349.2 KB | 0.64 |  |  | 4.2% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | OfficeIMO.Excel | 37.13 ms | 1.03 ms | 0.59 ms | 1.00 | 1.34 | 13002.3 KB | 1.00 |  |  | Loss +33.9% |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | MiniExcel | 72.48 ms | 5.33 ms | 3.08 ms | 1.95 | 2.61 | 92199.8 KB | 7.09 |  |  | 95.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | EPPlus 4.5.3.3 | 222.26 ms |  |  | 5.99 | 8.02 |  |  |  |  | 498.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | ClosedXML | 278.61 ms | 5.52 ms | 3.19 ms | 7.50 | 10.05 | 104205.0 KB | 8.01 |  |  | 650.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | EPPlus | 340.77 ms | 8.35 ms | 4.82 ms | 9.18 | 12.29 | 117438.0 KB | 9.03 |  |  | 817.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | LargeXlsx | 41.23 ms | 5.14 ms | 2.97 ms | 0.91 | 1.00 | 9265.9 KB | 0.94 |  |  | 9.5% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | OfficeIMO.Excel | 45.54 ms | 19.62 ms | 11.33 ms | 1.00 | 1.10 | 9819.7 KB | 1.00 |  |  | Loss +10.5% |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | MiniExcel | 99.68 ms | 5.27 ms | 3.04 ms | 2.19 | 2.42 | 108129.1 KB | 11.01 |  |  | 118.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | EPPlus | 591.72 ms | 63.89 ms | 36.89 ms | 12.99 | 14.35 | 135724.0 KB | 13.82 |  |  | 1199.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | ClosedXML | 691.56 ms | 245.51 ms | 141.74 ms | 15.18 | 16.77 | 280371.6 KB | 28.55 |  |  | 1418.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | LargeXlsx | 39.59 ms | 1.17 ms | 0.67 ms | 0.83 | 1.00 | 10795.2 KB | 0.92 |  |  | 16.8% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | OfficeIMO.Excel | 47.57 ms | 4.94 ms | 2.85 ms | 1.00 | 1.20 | 11708.2 KB | 1.00 |  |  | Loss +20.1% |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | MiniExcel | 149.96 ms | 3.51 ms | 2.02 ms | 3.15 | 3.79 | 226875.4 KB | 19.38 |  |  | 215.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | ClosedXML | 917.30 ms | 43.22 ms | 24.95 ms | 19.28 | 23.17 | 759818.4 KB | 64.90 |  |  | 1828.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | OfficeIMO.Excel | 14.40 ms | 0.28 ms | 0.16 ms | 1.00 | 1.00 | 15744.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | MiniExcel | 28.27 ms | 1.78 ms | 1.03 ms | 1.96 | 1.96 | 73760.2 KB | 4.68 |  |  | 96.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | EPPlus 4.5.3.3 | 101.45 ms |  |  | 7.05 | 7.05 |  |  |  |  | 604.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | ClosedXML | 106.16 ms | 0.98 ms | 0.57 ms | 7.37 | 7.37 | 104241.3 KB | 6.62 |  |  | 637.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | EPPlus | 174.77 ms | 2.46 ms | 1.42 ms | 12.14 | 12.14 | 84410.3 KB | 5.36 |  |  | 1114.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | OfficeIMO.Excel | 16.70 ms | 0.31 ms | 0.18 ms | 1.00 | 1.00 | 15360.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | EPPlus 4.5.3.3 | 93.86 ms |  |  | 5.62 | 5.62 |  |  |  |  | 462.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | ClosedXML | 123.65 ms | 14.98 ms | 8.65 ms | 7.40 | 7.40 | 104241.3 KB | 6.79 |  |  | 640.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | EPPlus | 193.53 ms | 11.13 ms | 6.43 ms | 11.59 | 11.59 | 84410.8 KB | 5.50 |  |  | 1058.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | OfficeIMO.Excel | 13.10 ms | 0.40 ms | 0.23 ms | 1.00 | 1.00 | 13824.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | ClosedXML | 149.32 ms | 8.04 ms | 4.64 ms | 11.39 | 11.39 | 131501.7 KB | 9.51 |  |  | 1039.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | EPPlus | 210.80 ms | 5.10 ms | 2.94 ms | 16.09 | 16.09 | 97730.0 KB | 7.07 |  |  | 1508.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | OfficeIMO.Excel | 11.53 ms | 0.20 ms | 0.12 ms | 1.00 | 1.00 | 7525.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | ClosedXML | 96.29 ms | 2.58 ms | 1.49 ms | 8.35 | 8.35 | 84520.0 KB | 11.23 |  |  | 734.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | EPPlus | 154.33 ms | 2.40 ms | 1.38 ms | 13.38 | 13.38 | 70033.7 KB | 9.31 |  |  | 1238.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | LargeXlsx | 26.99 ms | 0.09 ms | 0.05 ms | 0.82 | 1.00 | 5614.1 KB | 0.43 |  |  | 18.1% faster than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 32.94 ms | 1.84 ms | 1.06 ms | 1.00 | 1.22 | 12912.0 KB | 1.00 |  |  | Loss +22.0% |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | MiniExcel | 65.78 ms | 2.10 ms | 1.21 ms | 2.00 | 2.44 | 93257.0 KB | 7.22 |  |  | 99.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | ClosedXML | 273.55 ms | 5.28 ms | 3.05 ms | 8.31 | 10.14 | 104205.0 KB | 8.07 |  |  | 730.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | EPPlus | 364.75 ms | 3.03 ms | 1.75 ms | 11.07 | 13.51 | 117438.0 KB | 9.10 |  |  | 1007.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 405.64 ms |  |  | 12.32 | 15.03 |  |  |  |  | 1131.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | LargeXlsx | 30.38 ms | 1.83 ms | 1.06 ms | 0.92 | 1.00 | 5614.1 KB | 0.49 |  |  | 7.7% faster than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | OfficeIMO.Excel | 32.90 ms | 1.39 ms | 0.80 ms | 1.00 | 1.08 | 11493.8 KB | 1.00 |  |  | Loss +8.3% |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | MiniExcel | 72.21 ms | 3.77 ms | 2.18 ms | 2.20 | 2.38 | 93257.0 KB | 8.11 |  |  | 119.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | ClosedXML | 282.68 ms | 4.85 ms | 2.80 ms | 8.59 | 9.31 | 104205.0 KB | 9.07 |  |  | 759.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | EPPlus | 333.73 ms | 12.45 ms | 7.19 ms | 10.15 | 10.99 | 117437.6 KB | 10.22 |  |  | 914.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | EPPlus 4.5.3.3 | 340.19 ms |  |  | 10.34 | 11.20 |  |  |  |  | 934.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 28.20 ms | 0.56 ms | 0.32 ms | 0.76 | 1.00 | 5614.1 KB | 0.55 |  |  | 24.1% faster than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 37.15 ms | 0.41 ms | 0.23 ms | 1.00 | 1.32 | 10179.4 KB | 1.00 |  |  | Loss +31.7% |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | MiniExcel | 66.86 ms | 3.05 ms | 1.76 ms | 1.80 | 2.37 | 93257.0 KB | 9.16 |  |  | 80.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | ClosedXML | 284.87 ms | 5.79 ms | 3.34 ms | 7.67 | 10.10 | 104205.0 KB | 10.24 |  |  | 666.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | EPPlus | 333.21 ms | 10.41 ms | 6.01 ms | 8.97 | 11.81 | 117437.6 KB | 11.54 |  |  | 796.9% slower than OfficeIMO |
