# OfficeIMO.Excel Comparison Summary

This is the suite-level decision table. Mean, standard deviation, standard error, ratios, and allocations come from the lightweight rotated local runner; they are meant for engineering direction. Use the BenchmarkDotNet benchmark classes when a publication-grade `Error` column is required.

## At a glance

| Row count | Artifact | Workload | Category | OfficeIMO wins | OfficeIMO losses | Biggest loss |
| ---: | --- | --- | --- | ---: | ---: | --- |
| 2500 | dense-helloworld-comparison | read | Other | 0 | 2 | dense-helloworld-read-range: Loss +49.2% vs Sylvan.Data.Excel |
| 2500 | package-profile | package | Package size | 43 | 11 | write-insertobjects-legacy-dictionaries-direct: Loss +51.5% vs LargeXlsx |
| 2500 | speed-comparison | mutate | AutoFit and mutation | 2 | 0 |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | 4 | 0 |  |
| 2500 | speed-comparison | other | Object projection | 2 | 0 |  |
| 2500 | speed-comparison | other | Range and table read | 4 | 0 |  |
| 2500 | speed-comparison | other | Real-world feature mix | 6 | 0 |  |
| 2500 | speed-comparison | other | Real-world report | 2 | 0 |  |
| 2500 | speed-comparison | other | Report workbook | 4 | 0 |  |
| 2500 | speed-comparison | read | Formula write/read | 1 | 0 |  |
| 2500 | speed-comparison | read | Other | 1 | 2 | large-sparse-row-read: Loss +70.3% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | read | Range and table read | 1 | 6 | read-used-range: Loss +113.1% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | read | Streaming read | 1 | 3 | read-top-range-stream: Loss +49.0% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | read | Typed object read | 1 | 1 | read-objects-stream: Loss +6.9% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | write | AutoFit and mutation | 5 | 0 |  |
| 2500 | speed-comparison | write | Cell writer | 7 | 0 |  |
| 2500 | speed-comparison | write | DataSet table export | 3 | 0 |  |
| 2500 | speed-comparison | write | DataTable table export | 4 | 0 |  |
| 2500 | speed-comparison | write | Formatted report write | 1 | 0 |  |
| 2500 | speed-comparison | write | Formula write/read | 1 | 0 |  |
| 2500 | speed-comparison | write | Other | 2 | 2 | write-powershell-mixed-objects-direct: Loss +51.0% vs LargeXlsx |
| 2500 | speed-comparison | write | Plain cell export | 1 | 3 | append-plain-rows: Loss +38.3% vs LargeXlsx |
| 2500 | speed-comparison | write | Plain streaming export | 1 | 1 | write-datareader-plain: Loss +22.0% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | write | Plain string export | 0 | 1 | write-blog-2023-20-string-columns: Loss +2.0% vs LargeXlsx |
| 2500 | speed-comparison | write | Shared string write | 4 | 0 |  |
| 2500 | speed-comparison | write | Typed object export | 0 | 3 | write-fluent-rowsfrom-direct: Loss +49.2% vs LargeXlsx |
| 25000 | dense-helloworld-comparison | read | Other | 1 | 1 | dense-helloworld-read-range: Loss +16.6% vs Sylvan.Data.Excel |
| 25000 | package-profile | package | Package size | 43 | 11 | write-insertobjects-legacy-dictionaries-direct: Loss +55.0% vs LargeXlsx |
| 25000 | speed-comparison | mutate | AutoFit and mutation | 1 | 1 | realworld-report-no-autofit: Loss +7.4% vs EPPlus 4.5.3.3 |
| 25000 | speed-comparison | other | Anti-cheat report variants | 4 | 0 |  |
| 25000 | speed-comparison | other | Object projection | 2 | 0 |  |
| 25000 | speed-comparison | other | Range and table read | 4 | 0 |  |
| 25000 | speed-comparison | other | Real-world feature mix | 6 | 0 |  |
| 25000 | speed-comparison | other | Real-world report | 2 | 0 |  |
| 25000 | speed-comparison | other | Report workbook | 4 | 0 |  |
| 25000 | speed-comparison | read | Formula write/read | 1 | 0 |  |
| 25000 | speed-comparison | read | Other | 1 | 2 | shared-string-read: Loss +8.6% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | read | Range and table read | 2 | 5 | read-used-range: Loss +110.9% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | read | Streaming read | 1 | 3 | read-top-range-stream-small-chunks: Loss +30.5% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | read | Typed object read | 2 | 0 |  |
| 25000 | speed-comparison | write | AutoFit and mutation | 5 | 0 |  |
| 25000 | speed-comparison | write | Cell writer | 7 | 0 |  |
| 25000 | speed-comparison | write | DataSet table export | 3 | 0 |  |
| 25000 | speed-comparison | write | DataTable table export | 3 | 1 | write-datatable-direct: Loss +8.9% vs LargeXlsx |
| 25000 | speed-comparison | write | Formatted report write | 1 | 0 |  |
| 25000 | speed-comparison | write | Formula write/read | 1 | 0 |  |
| 25000 | speed-comparison | write | Other | 2 | 2 | write-powershell-mixed-objects-direct: Loss +24.7% vs LargeXlsx |
| 25000 | speed-comparison | write | Plain cell export | 1 | 3 | append-plain-rows: Loss +30.5% vs LargeXlsx |
| 25000 | speed-comparison | write | Plain streaming export | 1 | 1 | write-datareader-plain: Loss +22.8% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | write | Plain string export | 0 | 1 | write-blog-2023-20-string-columns: Loss +8.1% vs LargeXlsx |
| 25000 | speed-comparison | write | Shared string write | 4 | 0 |  |
| 25000 | speed-comparison | write | Typed object export | 0 | 3 | write-insertobjects-flat-dictionaries-direct: Loss +28.9% vs LargeXlsx |

## OfficeIMO decision table

| Row count | Artifact | Workload | Category | Scenario | OfficeIMO mean | Best | OfficeIMO vs best | Alloc | Package |
| ---: | --- | --- | --- | --- | ---: | --- | ---: | ---: | ---: |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | 6.14 ms | Sylvan.Data.Excel | Loss +49.2% | 2411.1 KB |  |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | 6.02 ms | Sylvan.Data.Excel | Loss +46.1% | 2489.5 KB |  |
| 2500 | package-profile | package | Package size | append-plain-rows | 1.97 ms | LargeXlsx | Loss +37.8% | 1576.3 KB | 63.0 KB |
| 2500 | package-profile | package | Package size | autofit-existing | 8.84 ms | OfficeIMO.Excel | Win | 1895.3 KB | 143.2 KB |
| 2500 | package-profile | package | Package size | large-shared-strings | 2.17 ms | OfficeIMO.Excel | Win | 2440.3 KB | 55.2 KB |
| 2500 | package-profile | package | Package size | realworld-autofilter | 3.67 ms | OfficeIMO.Excel | Win | 1340.4 KB | 142.4 KB |
| 2500 | package-profile | package | Package size | realworld-charts | 5.91 ms | OfficeIMO.Excel | Win | 1892.7 KB | 147.6 KB |
| 2500 | package-profile | package | Package size | realworld-conditional-formatting | 3.74 ms | OfficeIMO.Excel | Win | 1405.8 KB | 142.7 KB |
| 2500 | package-profile | package | Package size | realworld-data-validation | 3.55 ms | OfficeIMO.Excel | Win | 1356.1 KB | 142.5 KB |
| 2500 | package-profile | package | Package size | realworld-freeze-panes | 3.71 ms | OfficeIMO.Excel | Win | 1342.8 KB | 142.5 KB |
| 2500 | package-profile | package | Package size | realworld-pivot-table | 14.18 ms | OfficeIMO.Excel | Win | 14419.5 KB | 200.4 KB |
| 2500 | package-profile | package | Package size | realworld-report-all-in-one | 16.62 ms | OfficeIMO.Excel | Win | 15221.0 KB | 206.5 KB |
| 2500 | package-profile | package | Package size | realworld-report-chart-first | 11.04 ms | OfficeIMO.Excel | Win | 6195.1 KB | 206.5 KB |
| 2500 | package-profile | package | Package size | realworld-report-core | 4.35 ms | OfficeIMO.Excel | Win | 1488.5 KB | 143.9 KB |
| 2500 | package-profile | package | Package size | realworld-report-extra-column | 17.46 ms | OfficeIMO.Excel | Win | 16350.8 KB | 219.0 KB |
| 2500 | package-profile | package | Package size | realworld-report-no-autofit | 17.08 ms | OfficeIMO.Excel | Win | 15209.7 KB | 206.4 KB |
| 2500 | package-profile | package | Package size | realworld-report-post-mutation | 17.79 ms | OfficeIMO.Excel | Win | 15230.2 KB | 206.6 KB |
| 2500 | package-profile | package | Package size | realworld-report-shuffled-columns | 21.14 ms | OfficeIMO.Excel | Win | 15225.5 KB | 211.2 KB |
| 2500 | package-profile | package | Package size | report-workbook | 22.58 ms | OfficeIMO.Excel | Win | 19112.2 KB | 275.6 KB |
| 2500 | package-profile | package | Package size | report-workbook-core | 6.57 ms | OfficeIMO.Excel | Win | 2711.1 KB | 187.5 KB |
| 2500 | package-profile | package | Package size | report-workbook-datatable | 23.43 ms | OfficeIMO.Excel | Win | 19383.7 KB | 275.6 KB |
| 2500 | package-profile | package | Package size | report-workbook-datatable-core | 6.04 ms | OfficeIMO.Excel | Win | 2982.7 KB | 187.5 KB |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | 6.95 ms | LargeXlsx | Loss +11.6% | 1676.8 KB | 216.7 KB |
| 2500 | package-profile | package | Package size | write-bulk-report | 4.24 ms | OfficeIMO.Excel | Win | 1401.7 KB | 143.2 KB |
| 2500 | package-profile | package | Package size | write-cellformula | 2.38 ms | OfficeIMO.Excel | Win | 1383.3 KB | 66.6 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-empty-strings | 3.06 ms | OfficeIMO.Excel | Win | 1787.1 KB | 44.2 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-numbers | 2.39 ms | OfficeIMO.Excel | Win | 1119.9 KB | 47.0 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-object-mixed | 2.58 ms | OfficeIMO.Excel | Win | 1763.3 KB | 61.2 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse | 2.94 ms | OfficeIMO.Excel | Win | 1506.9 KB | 62.0 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | 2.46 ms | OfficeIMO.Excel | Win | 1507.0 KB | 62.0 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-scalars | 2.54 ms | OfficeIMO.Excel | Win | 1138.1 KB | 46.3 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-strings | 3.73 ms | OfficeIMO.Excel | Win | 2617.0 KB | 55.1 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-distinct | 3.19 ms | OfficeIMO.Excel | Win | 2379.2 KB | 51.8 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-repeated | 2.75 ms | OfficeIMO.Excel | Win | 1579.8 KB | 40.0 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-temporal | 3.64 ms | OfficeIMO.Excel | Win | 1435.7 KB | 63.3 KB |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | 1.51 ms | LargeXlsx, OfficeIMO.Excel | Win | 1092.0 KB | 48.2 KB |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | 4.44 ms | LargeXlsx | Loss +39.4% | 2081.1 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-datareader-plain | 4.15 ms | Sylvan.Data.Excel | Loss +32.7% | 1763.0 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-datareader-table | 4.19 ms | OfficeIMO.Excel | Win | 1774.9 KB | 138.8 KB |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | 4.54 ms | OfficeIMO.Excel | Win | 1781.2 KB | 138.9 KB |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | 3.95 ms | OfficeIMO.Excel | Win | 2140.6 KB | 131.1 KB |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | 4.77 ms | OfficeIMO.Excel | Win | 2880.2 KB | 176.0 KB |
| 2500 | package-profile | package | Package size | write-dataset-tables | 3.98 ms | OfficeIMO.Excel | Win | 2066.1 KB | 138.9 KB |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | 4.33 ms | OfficeIMO.Excel | Win | 2078.7 KB | 139.2 KB |
| 2500 | package-profile | package | Package size | write-datatable-direct | 3.99 ms | LargeXlsx | Loss +12.5% | 1748.6 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | 3.92 ms | OfficeIMO.Excel | Win | 1760.7 KB | 138.8 KB |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | 3.90 ms | LargeXlsx | Loss +25.0% | 1769.2 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | 4.05 ms | OfficeIMO.Excel | Win | 1347.1 KB | 142.4 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | 4.08 ms | LargeXlsx | Loss +15.4% | 1339.3 KB | 142.3 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | 5.42 ms | OfficeIMO.Excel | Win | 1505.3 KB | 138.1 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | 4.24 ms | LargeXlsx | Loss +31.5% | 1497.5 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | 4.73 ms | LargeXlsx | Loss +51.5% | 1770.1 KB | 142.3 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | 4.24 ms | OfficeIMO.Excel | Win | 1346.4 KB | 142.4 KB |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | 5.79 ms | LargeXlsx | Loss +34.4% | 2341.7 KB | 183.1 KB |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | 4.47 ms | LargeXlsx | Loss +5.3% | 1507.7 KB | 182.4 KB |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | 20.44 ms | OfficeIMO.Excel | Win | 4502.3 KB | 651.0 KB |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | 9.89 ms | OfficeIMO.Excel | Win | 1895.1 KB |  |
| 2500 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | 15.66 ms | OfficeIMO.Excel | Win | 15209.4 KB |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | 11.72 ms | OfficeIMO.Excel | Win | 6195.2 KB |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | 18.72 ms | OfficeIMO.Excel | Win | 16350.4 KB |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | 16.75 ms | OfficeIMO.Excel | Win | 15230.4 KB |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | 17.87 ms | OfficeIMO.Excel | Win | 15225.7 KB |  |
| 2500 | speed-comparison | other | Object projection | build-object-datatable-dictionaries | 1.35 ms | OfficeIMO.Excel | Win | 564.2 KB |  |
| 2500 | speed-comparison | other | Object projection | build-object-datatable-typed | 1.18 ms | OfficeIMO.Excel | Win | 856.9 KB |  |
| 2500 | speed-comparison | other | Range and table read | enumerate-cells | 6.06 ms | OfficeIMO.Excel | Win | 2531.7 KB |  |
| 2500 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | 4.33 ms | OfficeIMO.Excel | Win | 523.5 KB |  |
| 2500 | speed-comparison | other | Range and table read | enumerate-range | 6.16 ms | OfficeIMO.Excel | Win | 2531.8 KB |  |
| 2500 | speed-comparison | other | Range and table read | enumerate-top-range | 0.69 ms | OfficeIMO.Excel | Win | 285.4 KB |  |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-autofilter | 3.68 ms | OfficeIMO.Excel | Win | 1340.4 KB |  |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-charts | 5.59 ms | OfficeIMO.Excel | Win | 1892.8 KB |  |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | 3.77 ms | OfficeIMO.Excel | Win | 1405.8 KB |  |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-data-validation | 3.64 ms | OfficeIMO.Excel | Win | 1356.1 KB |  |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | 3.89 ms | OfficeIMO.Excel | Win | 1342.9 KB |  |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | 17.61 ms | OfficeIMO.Excel | Win | 14419.3 KB |  |
| 2500 | speed-comparison | other | Real-world report | realworld-report-all-in-one | 17.09 ms | OfficeIMO.Excel | Win | 15220.5 KB |  |
| 2500 | speed-comparison | other | Real-world report | realworld-report-core | 4.44 ms | OfficeIMO.Excel | Win | 1488.6 KB |  |
| 2500 | speed-comparison | other | Report workbook | report-workbook | 30.29 ms | OfficeIMO.Excel | Win | 19069.0 KB |  |
| 2500 | speed-comparison | other | Report workbook | report-workbook-core | 6.31 ms | OfficeIMO.Excel | Win | 2711.1 KB |  |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable | 22.91 ms | OfficeIMO.Excel | Win | 19383.9 KB |  |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable-core | 5.94 ms | OfficeIMO.Excel | Win | 2982.8 KB |  |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | 2.65 ms | OfficeIMO.Excel | Win | 706.8 KB |  |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | 0.96 ms | OfficeIMO.Excel | Win | 177.3 KB |  |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | 1.77 ms | Sylvan.Data.Excel | Loss +70.3% | 177.5 KB |  |
| 2500 | speed-comparison | read | Other | shared-string-read | 4.31 ms | Sylvan.Data.Excel | Loss +30.1% | 1056.7 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | 4.54 ms | Sylvan.Data.Excel | Loss +6.0% | 374.7 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-datatable | 7.66 ms | Sylvan.Data.Excel | Loss +25.4% | 3594.6 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | 5.21 ms | Sylvan.Data.Excel | Loss +9.3% | 543.1 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-range | 18.87 ms | Sylvan.Data.Excel | Loss +4.6% | 2692.7 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | 5.26 ms | OfficeIMO.Excel, Sylvan.Data.Excel | Win | 2751.4 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-top-range | 0.61 ms | Sylvan.Data.Excel | Loss +19.7% | 296.2 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-used-range | 10.71 ms | Sylvan.Data.Excel | Loss +113.1% | 3472.7 KB |  |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | 3.76 ms | OfficeIMO.Excel | Win | 378.0 KB |  |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | 6.19 ms | Sylvan.Data.Excel | Loss +13.6% | 2771.5 KB |  |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | 0.67 ms | Sylvan.Data.Excel | Loss +49.0% | 299.5 KB |  |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | 0.55 ms | Sylvan.Data.Excel | Loss +28.1% | 300.3 KB |  |
| 2500 | speed-comparison | read | Typed object read | read-objects | 7.55 ms | OfficeIMO.Excel | Win | 2442.2 KB |  |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | 5.35 ms | Sylvan.Data.Excel | Loss +6.9% | 2423.1 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | 5.07 ms | OfficeIMO.Excel | Win | 1781.2 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | 7.29 ms | OfficeIMO.Excel | Win | 2079.3 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | 4.02 ms | OfficeIMO.Excel | Win | 1347.1 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | 4.62 ms | OfficeIMO.Excel | Win | 1505.3 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | 3.90 ms | OfficeIMO.Excel | Win | 1346.4 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | 2.30 ms | OfficeIMO.Excel | Win | 1787.1 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | 2.13 ms | OfficeIMO.Excel | Win | 1119.9 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | 3.21 ms | OfficeIMO.Excel | Win | 1763.1 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | 3.24 ms | OfficeIMO.Excel | Win | 1506.7 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | 3.13 ms | OfficeIMO.Excel | Win | 1506.8 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | 2.25 ms | OfficeIMO.Excel | Win | 1138.1 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | 3.93 ms | OfficeIMO.Excel | Win | 1435.5 KB |  |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | 5.45 ms | OfficeIMO.Excel | Win | 2064.5 KB |  |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | 8.68 ms | OfficeIMO.Excel | Win | 2880.2 KB |  |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | 5.50 ms | OfficeIMO.Excel | Win | 2067.7 KB |  |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | 5.60 ms | OfficeIMO.Excel | Win | 1774.9 KB |  |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | 4.22 ms | OfficeIMO.Excel | Win | 1748.6 KB |  |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | 4.88 ms | OfficeIMO.Excel | Win | 1487.2 KB |  |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | 4.94 ms | OfficeIMO.Excel | Win | 1760.7 KB |  |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | 6.56 ms | OfficeIMO.Excel | Win | 1403.3 KB |  |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | 3.05 ms | OfficeIMO.Excel | Win | 1620.6 KB |  |
| 2500 | speed-comparison | write | Other | write-dictionary-objects-table-direct | 5.62 ms | OfficeIMO.Excel | Win | 2051.4 KB |  |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | 7.16 ms | LargeXlsx | Loss +51.0% | 2341.7 KB |  |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | 6.13 ms | LargeXlsx | Loss +17.6% | 1507.7 KB |  |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | 21.68 ms | OfficeIMO.Excel | Win | 4502.3 KB |  |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | 3.26 ms | LargeXlsx | Loss +38.3% | 1576.3 KB |  |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | 1.65 ms | LargeXlsx | Loss +28.4% | 1092.0 KB |  |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | 4.49 ms | LargeXlsx | Loss +34.3% | 2081.1 KB |  |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | 2.39 ms | OfficeIMO.Excel | Win | 1494.0 KB |  |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | 4.51 ms | Sylvan.Data.Excel | Loss +22.0% | 1763.0 KB |  |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | 4.71 ms | OfficeIMO.Excel | Win | 2140.6 KB |  |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | 5.69 ms | LargeXlsx | Loss +2.0% | 1676.8 KB |  |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | 2.82 ms | OfficeIMO.Excel | Win | 2440.3 KB |  |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | 3.58 ms | OfficeIMO.Excel | Win | 2617.0 KB |  |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | 2.99 ms | OfficeIMO.Excel | Win | 2379.2 KB |  |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | 2.59 ms | OfficeIMO.Excel | Win | 1579.8 KB |  |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | 4.88 ms | LargeXlsx | Loss +49.2% | 1769.2 KB |  |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | 4.36 ms | LargeXlsx | Loss +38.2% | 1339.3 KB |  |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | 4.35 ms | LargeXlsx | Loss +35.9% | 1497.5 KB |  |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | 79.91 ms | Sylvan.Data.Excel | Loss +16.6% | 23622.2 KB |  |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | 62.41 ms | OfficeIMO.Excel | Win | 24404.3 KB |  |
| 25000 | package-profile | package | Package size | append-plain-rows | 20.48 ms | LargeXlsx | Loss +34.7% | 10842.5 KB | 610.4 KB |
| 25000 | package-profile | package | Package size | autofit-existing | 99.77 ms | OfficeIMO.Excel | Win | 15708.3 KB | 1429.3 KB |
| 25000 | package-profile | package | Package size | large-shared-strings | 22.57 ms | OfficeIMO.Excel | Win | 15744.9 KB | 529.7 KB |
| 25000 | package-profile | package | Package size | realworld-autofilter | 49.51 ms | OfficeIMO.Excel | Win | 11494.9 KB | 1428.5 KB |
| 25000 | package-profile | package | Package size | realworld-charts | 34.86 ms | OfficeIMO.Excel | Win | 12551.0 KB | 1433.7 KB |
| 25000 | package-profile | package | Package size | realworld-conditional-formatting | 46.99 ms | OfficeIMO.Excel | Win | 11560.2 KB | 1428.8 KB |
| 25000 | package-profile | package | Package size | realworld-data-validation | 43.64 ms | OfficeIMO.Excel | Win | 11510.5 KB | 1428.6 KB |
| 25000 | package-profile | package | Package size | realworld-freeze-panes | 49.71 ms | OfficeIMO.Excel | Win | 11497.3 KB | 1428.6 KB |
| 25000 | package-profile | package | Package size | realworld-pivot-table | 385.23 ms | OfficeIMO.Excel | Win | 131927.7 KB | 1979.8 KB |
| 25000 | package-profile | package | Package size | realworld-report-all-in-one | 279.05 ms | OfficeIMO.Excel | Win | 133444.7 KB | 1985.9 KB |
| 25000 | package-profile | package | Package size | realworld-report-chart-first | 96.51 ms | OfficeIMO.Excel | Win | 43564.8 KB | 1985.9 KB |
| 25000 | package-profile | package | Package size | realworld-report-core | 54.69 ms | OfficeIMO.Excel | Win | 11648.7 KB | 1430.0 KB |
| 25000 | package-profile | package | Package size | realworld-report-extra-column | 477.94 ms | OfficeIMO.Excel | Win | 144827.2 KB | 2110.8 KB |
| 25000 | package-profile | package | Package size | realworld-report-no-autofit | 269.74 ms | OfficeIMO.Excel | Win | 133432.6 KB | 1985.8 KB |
| 25000 | package-profile | package | Package size | realworld-report-post-mutation | 355.35 ms | OfficeIMO.Excel | Win | 133461.9 KB | 1986.0 KB |
| 25000 | package-profile | package | Package size | realworld-report-shuffled-columns | 301.21 ms | OfficeIMO.Excel | Win | 133506.1 KB | 2046.1 KB |
| 25000 | package-profile | package | Package size | report-workbook | 618.76 ms | OfficeIMO.Excel | Win | 175194.0 KB | 2672.0 KB |
| 25000 | package-profile | package | Package size | report-workbook-core | 72.67 ms | OfficeIMO.Excel | Win | 10979.4 KB | 1850.9 KB |
| 25000 | package-profile | package | Package size | report-workbook-datatable | 597.30 ms | OfficeIMO.Excel | Win | 177940.0 KB | 2672.0 KB |
| 25000 | package-profile | package | Package size | report-workbook-datatable-core | 66.63 ms | OfficeIMO.Excel | Win | 13725.0 KB | 1850.9 KB |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | 48.91 ms | LargeXlsx | Loss +15.7% | 11708.2 KB | 2228.8 KB |
| 25000 | package-profile | package | Package size | write-bulk-report | 47.67 ms | OfficeIMO.Excel | Win | 11561.8 KB | 1429.3 KB |
| 25000 | package-profile | package | Package size | write-cellformula | 29.76 ms | OfficeIMO.Excel | Win | 10112.0 KB | 670.3 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-empty-strings | 16.61 ms | OfficeIMO.Excel | Win | 6896.4 KB | 451.4 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-numbers | 19.93 ms | OfficeIMO.Excel | Win | 5970.9 KB | 462.6 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-object-mixed | 25.30 ms | OfficeIMO.Excel | Win | 8332.9 KB | 585.1 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse | 26.63 ms | OfficeIMO.Excel | Win | 7416.2 KB | 607.1 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | 23.82 ms | OfficeIMO.Excel | Win | 7416.3 KB | 607.1 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-scalars | 15.47 ms | OfficeIMO.Excel | Win | 6144.6 KB | 441.9 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-strings | 26.95 ms | OfficeIMO.Excel | Win | 15360.4 KB | 527.8 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-distinct | 16.89 ms | OfficeIMO.Excel | Win | 13824.1 KB | 499.0 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-repeated | 17.02 ms | OfficeIMO.Excel | Win | 7525.3 KB | 376.0 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-temporal | 28.45 ms | OfficeIMO.Excel | Win | 7482.8 KB | 620.5 KB |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | 13.28 ms | LargeXlsx | Loss +24.4% | 6961.7 KB | 455.5 KB |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | 44.56 ms | OfficeIMO.Excel | Win | 16036.5 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-datareader-plain | 50.74 ms | Sylvan.Data.Excel | Loss +18.6% | 13002.3 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-datareader-table | 48.21 ms | OfficeIMO.Excel | Win | 13020.3 KB | 1385.7 KB |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | 55.01 ms | OfficeIMO.Excel | Win | 13026.6 KB | 1385.9 KB |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | 41.29 ms | OfficeIMO.Excel | Win | 9819.7 KB | 1329.2 KB |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | 47.16 ms | OfficeIMO.Excel | Win | 13458.5 KB | 1795.1 KB |
| 25000 | package-profile | package | Package size | write-dataset-tables | 48.11 ms | OfficeIMO.Excel | Win | 10288.1 KB | 1376.4 KB |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | 48.13 ms | OfficeIMO.Excel | Win | 10300.7 KB | 1376.7 KB |
| 25000 | package-profile | package | Package size | write-datatable-direct | 59.54 ms | LargeXlsx | Loss +24.6% | 12715.7 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | 42.27 ms | OfficeIMO.Excel | Win | 12733.8 KB | 1385.7 KB |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | 46.79 ms | LargeXlsx | Loss +23.7% | 12912.0 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | 66.78 ms | OfficeIMO.Excel | Win | 11501.6 KB | 1428.5 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | 42.53 ms | LargeXlsx | Loss +11.4% | 11493.8 KB | 1428.4 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | 60.51 ms | OfficeIMO.Excel | Win | 10187.2 KB | 1385.1 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | 65.00 ms | LargeXlsx | Loss +35.8% | 10179.4 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | 62.00 ms | LargeXlsx | Loss +55.0% | 15791.7 KB | 1428.4 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | 48.05 ms | OfficeIMO.Excel | Win | 11500.9 KB | 1428.5 KB |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | 69.88 ms | LargeXlsx | Loss +12.2% | 10577.2 KB | 1828.0 KB |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | 81.03 ms | LargeXlsx | Loss +37.2% | 9942.2 KB | 1828.0 KB |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | 275.51 ms | OfficeIMO.Excel | Win | 36150.1 KB | 6725.6 KB |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | 101.95 ms | OfficeIMO.Excel | Win | 15708.3 KB |  |
| 25000 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | 271.08 ms | EPPlus 4.5.3.3 | Loss +7.4% | 133435.8 KB |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | 94.45 ms | OfficeIMO.Excel | Win | 43566.1 KB |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | 408.57 ms | OfficeIMO.Excel | Win | 144822.3 KB |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | 289.75 ms | OfficeIMO.Excel | Win | 133462.3 KB |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | 344.54 ms | OfficeIMO.Excel | Win | 133505.9 KB |  |
| 25000 | speed-comparison | other | Object projection | build-object-datatable-dictionaries | 9.68 ms | OfficeIMO.Excel | Win | 5164.3 KB |  |
| 25000 | speed-comparison | other | Object projection | build-object-datatable-typed | 8.13 ms | OfficeIMO.Excel | Win | 8093.8 KB |  |
| 25000 | speed-comparison | other | Range and table read | enumerate-cells | 56.77 ms | OfficeIMO.Excel | Win | 24531.0 KB |  |
| 25000 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | 35.46 ms | OfficeIMO.Excel | Win | 3839.4 KB |  |
| 25000 | speed-comparison | other | Range and table read | enumerate-range | 51.77 ms | OfficeIMO.Excel | Win | 24531.2 KB |  |
| 25000 | speed-comparison | other | Range and table read | enumerate-top-range | 0.64 ms | OfficeIMO.Excel | Win | 285.5 KB |  |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-autofilter | 32.26 ms | OfficeIMO.Excel | Win | 11494.9 KB |  |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-charts | 33.47 ms | OfficeIMO.Excel | Win | 12552.7 KB |  |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | 33.86 ms | OfficeIMO.Excel | Win | 11560.2 KB |  |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-data-validation | 31.38 ms | OfficeIMO.Excel | Win | 11510.5 KB |  |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | 31.51 ms | OfficeIMO.Excel | Win | 11497.3 KB |  |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | 258.35 ms | OfficeIMO.Excel | Win | 131928.7 KB |  |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | 317.70 ms | OfficeIMO.Excel | Win | 133447.0 KB |  |
| 25000 | speed-comparison | other | Real-world report | realworld-report-core | 40.57 ms | OfficeIMO.Excel | Win | 11648.7 KB |  |
| 25000 | speed-comparison | other | Report workbook | report-workbook | 493.30 ms | OfficeIMO.Excel | Win | 175197.4 KB |  |
| 25000 | speed-comparison | other | Report workbook | report-workbook-core | 52.78 ms | OfficeIMO.Excel | Win | 10979.4 KB |  |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable | 450.29 ms | OfficeIMO.Excel | Win | 177942.0 KB |  |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable-core | 53.43 ms | OfficeIMO.Excel | Win | 13725.0 KB |  |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | 19.74 ms | OfficeIMO.Excel | Win | 6216.4 KB |  |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | 1.11 ms | Sylvan.Data.Excel | Loss +8.0% | 177.4 KB |  |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | 1.17 ms | OfficeIMO.Excel | Win | 177.5 KB |  |
| 25000 | speed-comparison | read | Other | shared-string-read | 18.62 ms | Sylvan.Data.Excel | Loss +8.6% | 9218.2 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | 33.96 ms | OfficeIMO.Excel | Win | 1122.6 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-datatable | 69.30 ms | Sylvan.Data.Excel | Loss +7.1% | 34646.0 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | 72.82 ms | OfficeIMO.Excel | Win | 4034.7 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-range | 87.94 ms | Sylvan.Data.Excel | Loss +4.5% | 26098.5 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | 63.12 ms | Sylvan.Data.Excel | Loss +4.6% | 26684.4 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-top-range | 0.61 ms | Sylvan.Data.Excel | Loss +18.5% | 296.3 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-used-range | 127.12 ms | Sylvan.Data.Excel | Loss +110.9% | 34152.1 KB |  |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | 41.65 ms | Sylvan.Data.Excel | Loss +3.9% | 1125.9 KB |  |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | 53.97 ms | Sylvan.Data.Excel | Loss +12.5% | 26884.0 KB |  |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | 0.75 ms | OfficeIMO.Excel | Win | 302.3 KB |  |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | 0.57 ms | Sylvan.Data.Excel | Loss +30.5% | 300.3 KB |  |
| 25000 | speed-comparison | read | Typed object read | read-objects | 98.10 ms | OfficeIMO.Excel | Win | 23562.5 KB |  |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | 50.40 ms | OfficeIMO.Excel | Win | 23367.6 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | 56.22 ms | OfficeIMO.Excel | Win | 13026.6 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | 37.89 ms | OfficeIMO.Excel | Win | 10300.7 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | 49.88 ms | OfficeIMO.Excel | Win | 11501.6 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | 58.50 ms | OfficeIMO.Excel | Win | 10187.2 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | 45.87 ms | OfficeIMO.Excel | Win | 11500.9 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | 17.09 ms | OfficeIMO.Excel | Win | 6896.4 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | 20.11 ms | OfficeIMO.Excel | Win | 5970.9 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | 21.80 ms | OfficeIMO.Excel | Win | 8332.9 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | 23.70 ms | OfficeIMO.Excel | Win | 7416.2 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | 21.95 ms | OfficeIMO.Excel | Win | 7416.3 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | 15.30 ms | OfficeIMO.Excel | Win | 6144.6 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | 23.99 ms | OfficeIMO.Excel | Win | 7482.8 KB |  |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | 38.51 ms | OfficeIMO.Excel | Win | 13039.6 KB |  |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | 48.69 ms | OfficeIMO.Excel | Win | 13458.5 KB |  |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | 34.40 ms | OfficeIMO.Excel | Win | 10288.1 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | 53.21 ms | OfficeIMO.Excel | Win | 13020.3 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | 33.13 ms | LargeXlsx | Loss +8.9% | 12715.7 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | 37.43 ms | OfficeIMO.Excel | Win | 9999.4 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | 34.71 ms | OfficeIMO.Excel | Win | 12733.8 KB |  |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | 38.72 ms | OfficeIMO.Excel | Win | 11561.8 KB |  |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | 27.62 ms | OfficeIMO.Excel | Win | 10112.0 KB |  |
| 25000 | speed-comparison | write | Other | write-dictionary-objects-table-direct | 50.71 ms | OfficeIMO.Excel | Win | 15163.8 KB |  |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | 68.65 ms | LargeXlsx | Loss +24.7% | 10577.2 KB |  |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | 60.67 ms | LargeXlsx | Loss +12.7% | 9942.2 KB |  |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | 284.19 ms | OfficeIMO.Excel | Win | 36150.1 KB |  |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | 15.18 ms | LargeXlsx | Loss +30.5% | 10842.5 KB |  |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | 14.58 ms | LargeXlsx | Loss +13.8% | 6961.7 KB |  |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | 49.20 ms | LargeXlsx | Loss +10.9% | 16036.5 KB |  |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | 21.40 ms | OfficeIMO.Excel | Win | 7866.1 KB |  |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | 50.25 ms | Sylvan.Data.Excel | Loss +22.8% | 13002.3 KB |  |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | 32.65 ms | OfficeIMO.Excel | Win | 9819.7 KB |  |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | 57.79 ms | LargeXlsx | Loss +8.1% | 11708.2 KB |  |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | 14.81 ms | OfficeIMO.Excel | Win | 15744.9 KB |  |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | 22.75 ms | OfficeIMO.Excel | Win | 15360.4 KB |  |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | 18.28 ms | OfficeIMO.Excel | Win | 13824.1 KB |  |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | 15.99 ms | OfficeIMO.Excel | Win | 7525.3 KB |  |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | 45.41 ms | LargeXlsx | Loss +8.0% | 12912.0 KB |  |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | 44.27 ms | LargeXlsx | Loss +5.5% | 11493.8 KB |  |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | 49.23 ms | LargeXlsx | Loss +28.9% | 10179.4 KB |  |

## Full comparison table

| Row count | Artifact | Workload | Category | Scenario | Library | Mean | StdDev | StdErr | Ratio to OfficeIMO | Ratio to best | Alloc | Alloc ratio | Package | Package ratio | Outcome |
| ---: | --- | --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | Sylvan.Data.Excel | 4.11 ms | 0.17 ms | 0.10 ms | 0.67 | 1.00 | 362.3 KB | 0.15 |  |  | 33.0% faster than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | OfficeIMO.Excel | 6.14 ms | 0.29 ms | 0.17 ms | 1.00 | 1.49 | 2411.1 KB | 1.00 |  |  | Loss +49.2% |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | ExcelDataReader | 11.15 ms | 0.63 ms | 0.36 ms | 1.82 | 2.71 | 6887.4 KB | 2.86 |  |  | 81.7% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | MiniExcel | 18.30 ms | 5.58 ms | 3.22 ms | 2.98 | 4.45 | 21507.3 KB | 8.92 |  |  | 198.2% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | Sylvan.Data.Excel | 4.12 ms | 0.14 ms | 0.08 ms | 0.68 | 1.00 | 362.3 KB | 0.15 |  |  | 31.6% faster than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | OfficeIMO.Excel | 6.02 ms | 0.39 ms | 0.23 ms | 1.00 | 1.46 | 2489.5 KB | 1.00 |  |  | Loss +46.1% |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | ExcelDataReader | 10.48 ms | 0.35 ms | 0.20 ms | 1.74 | 2.54 | 6887.4 KB | 2.77 |  |  | 74.0% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | MiniExcel | 16.70 ms | 1.75 ms | 1.01 ms | 2.77 | 4.05 | 21507.3 KB | 8.64 |  |  | 177.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | append-plain-rows | LargeXlsx | 1.43 ms | 0.01 ms | 0.00 ms | 0.73 | 1.00 | 296.4 KB | 0.19 | 63.1 KB | 1.00 | 27.4% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | append-plain-rows | OfficeIMO.Excel | 1.97 ms | 0.01 ms | 0.00 ms | 1.00 | 1.38 | 1576.3 KB | 1.00 | 63.0 KB | 1.00 | Loss +37.8% |
| 2500 | package-profile | package | Package size | append-plain-rows | MiniExcel | 4.47 ms | 0.43 ms | 0.25 ms | 2.27 | 3.13 | 19710.6 KB | 12.50 | 68.1 KB | 1.08 | 126.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | append-plain-rows | ClosedXML | 14.80 ms | 0.36 ms | 0.21 ms | 7.51 | 10.34 | 11197.4 KB | 7.10 | 59.8 KB | 0.95 | 650.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | append-plain-rows | EPPlus | 28.58 ms | 3.17 ms | 1.83 ms | 14.50 | 19.97 | 14365.2 KB | 9.11 | 56.9 KB | 0.90 | 1349.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | autofit-existing | OfficeIMO.Excel | 8.84 ms | 0.55 ms | 0.32 ms | 1.00 | 1.00 | 1895.3 KB | 1.00 | 143.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | autofit-existing | EPPlus | 74.58 ms | 0.70 ms | 0.40 ms | 8.43 | 8.43 | 50712.0 KB | 26.76 | 115.0 KB | 0.80 | 743.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | autofit-existing | ClosedXML | 134.49 ms | 5.14 ms | 2.97 ms | 15.21 | 15.21 | 84562.8 KB | 44.62 | 121.0 KB | 0.84 | 1420.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | large-shared-strings | OfficeIMO.Excel | 2.17 ms | 0.10 ms | 0.06 ms | 1.00 | 1.00 | 2440.3 KB | 1.00 | 55.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | large-shared-strings | MiniExcel | 4.07 ms | 0.04 ms | 0.02 ms | 1.88 | 1.88 | 21137.5 KB | 8.66 | 60.7 KB | 1.10 | 87.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | large-shared-strings | ClosedXML | 13.57 ms | 1.93 ms | 1.12 ms | 6.25 | 6.25 | 11299.2 KB | 4.63 | 50.3 KB | 0.91 | 524.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | large-shared-strings | EPPlus | 22.59 ms | 0.80 ms | 0.46 ms | 10.40 | 10.40 | 12804.4 KB | 5.25 | 48.1 KB | 0.87 | 939.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-autofilter | OfficeIMO.Excel | 3.67 ms | 0.07 ms | 0.04 ms | 1.00 | 1.00 | 1340.4 KB | 1.00 | 142.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-autofilter | ClosedXML | 33.17 ms | 3.49 ms | 2.02 ms | 9.05 | 9.05 | 22226.8 KB | 16.58 | 120.2 KB | 0.84 | 804.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-autofilter | EPPlus | 40.85 ms | 0.31 ms | 0.18 ms | 11.14 | 11.14 | 24715.5 KB | 18.44 | 114.2 KB | 0.80 | 1014.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-charts | OfficeIMO.Excel | 5.91 ms | 0.86 ms | 0.50 ms | 1.00 | 1.00 | 1892.7 KB | 1.00 | 147.6 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-charts | EPPlus | 44.08 ms | 0.77 ms | 0.44 ms | 7.46 | 7.46 | 27141.8 KB | 14.34 | 117.0 KB | 0.79 | 646.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-conditional-formatting | OfficeIMO.Excel | 3.74 ms | 0.18 ms | 0.11 ms | 1.00 | 1.00 | 1405.8 KB | 1.00 | 142.7 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-conditional-formatting | ClosedXML | 30.31 ms | 0.54 ms | 0.31 ms | 8.09 | 8.09 | 22273.8 KB | 15.84 | 120.3 KB | 0.84 | 709.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-conditional-formatting | EPPlus | 43.51 ms | 1.96 ms | 1.13 ms | 11.62 | 11.62 | 24757.5 KB | 17.61 | 114.3 KB | 0.80 | 1062.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-data-validation | OfficeIMO.Excel | 3.55 ms | 0.09 ms | 0.05 ms | 1.00 | 1.00 | 1356.1 KB | 1.00 | 142.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-data-validation | ClosedXML | 30.75 ms | 1.05 ms | 0.61 ms | 8.66 | 8.66 | 22247.9 KB | 16.41 | 120.3 KB | 0.84 | 766.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-data-validation | EPPlus | 40.78 ms | 1.28 ms | 0.74 ms | 11.49 | 11.49 | 24701.4 KB | 18.22 | 114.2 KB | 0.80 | 1049.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-freeze-panes | OfficeIMO.Excel | 3.71 ms | 0.17 ms | 0.10 ms | 1.00 | 1.00 | 1342.8 KB | 1.00 | 142.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-freeze-panes | ClosedXML | 32.11 ms | 3.01 ms | 1.74 ms | 8.66 | 8.66 | 22222.0 KB | 16.55 | 120.2 KB | 0.84 | 765.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-freeze-panes | EPPlus | 43.72 ms | 4.26 ms | 2.46 ms | 11.79 | 11.79 | 24730.0 KB | 18.42 | 114.3 KB | 0.80 | 1078.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-pivot-table | OfficeIMO.Excel | 14.18 ms | 0.96 ms | 0.55 ms | 1.00 | 1.00 | 14419.5 KB | 1.00 | 200.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-pivot-table | EPPlus | 44.79 ms | 2.22 ms | 1.28 ms | 3.16 | 3.16 | 29537.1 KB | 2.05 | 117.4 KB | 0.59 | 215.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-all-in-one | OfficeIMO.Excel | 16.62 ms | 0.96 ms | 0.55 ms | 1.00 | 1.00 | 15221.0 KB | 1.00 | 206.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-all-in-one | EPPlus | 72.49 ms | 1.91 ms | 1.10 ms | 4.36 | 4.36 | 54594.3 KB | 3.59 | 121.8 KB | 0.59 | 336.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-chart-first | OfficeIMO.Excel | 11.04 ms | 0.81 ms | 0.47 ms | 1.00 | 1.00 | 6195.1 KB | 1.00 | 206.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-chart-first | EPPlus | 72.66 ms | 1.67 ms | 0.96 ms | 6.58 | 6.58 | 54593.5 KB | 8.81 | 121.8 KB | 0.59 | 558.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-core | OfficeIMO.Excel | 4.35 ms | 0.20 ms | 0.12 ms | 1.00 | 1.00 | 1488.5 KB | 1.00 | 143.9 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-core | EPPlus | 64.46 ms | 1.53 ms | 0.88 ms | 14.83 | 14.83 | 47299.8 KB | 31.78 | 115.6 KB | 0.80 | 1382.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-core | ClosedXML | 81.73 ms | 6.06 ms | 3.50 ms | 18.80 | 18.80 | 69836.4 KB | 46.92 | 121.5 KB | 0.84 | 1780.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-extra-column | OfficeIMO.Excel | 17.46 ms | 0.60 ms | 0.34 ms | 1.00 | 1.00 | 16350.8 KB | 1.00 | 219.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-extra-column | EPPlus | 81.43 ms | 3.92 ms | 2.26 ms | 4.66 | 4.66 | 59225.9 KB | 3.62 | 128.4 KB | 0.59 | 366.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-no-autofit | OfficeIMO.Excel | 17.08 ms | 1.13 ms | 0.65 ms | 1.00 | 1.00 | 15209.7 KB | 1.00 | 206.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-no-autofit | EPPlus | 48.34 ms | 2.32 ms | 1.34 ms | 2.83 | 2.83 | 32906.1 KB | 2.16 | 121.8 KB | 0.59 | 183.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-post-mutation | OfficeIMO.Excel | 17.79 ms | 1.76 ms | 1.01 ms | 1.00 | 1.00 | 15230.2 KB | 1.00 | 206.6 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-post-mutation | EPPlus | 72.15 ms | 0.90 ms | 0.52 ms | 4.06 | 4.06 | 54594.0 KB | 3.58 | 121.9 KB | 0.59 | 305.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-shuffled-columns | OfficeIMO.Excel | 21.14 ms | 2.49 ms | 1.44 ms | 1.00 | 1.00 | 15225.5 KB | 1.00 | 211.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-shuffled-columns | EPPlus | 79.17 ms | 4.28 ms | 2.47 ms | 3.74 | 3.74 | 54590.6 KB | 3.59 | 124.3 KB | 0.59 | 274.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook | OfficeIMO.Excel | 22.58 ms | 1.16 ms | 0.67 ms | 1.00 | 1.00 | 19112.2 KB | 1.00 | 275.6 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | report-workbook | EPPlus | 89.88 ms | 0.88 ms | 0.51 ms | 3.98 | 3.98 | 77485.5 KB | 4.05 | 161.8 KB | 0.59 | 298.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook-core | OfficeIMO.Excel | 6.57 ms | 0.07 ms | 0.04 ms | 1.00 | 1.00 | 2711.1 KB | 1.00 | 187.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | report-workbook-core | EPPlus | 99.68 ms | 6.38 ms | 3.68 ms | 15.18 | 15.18 | 71970.6 KB | 26.55 | 157.2 KB | 0.84 | 1417.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook-core | ClosedXML | 105.93 ms | 1.27 ms | 0.73 ms | 16.13 | 16.13 | 97220.0 KB | 35.86 | 165.1 KB | 0.88 | 1513.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook-datatable | OfficeIMO.Excel | 23.43 ms | 0.98 ms | 0.57 ms | 1.00 | 1.00 | 19383.7 KB | 1.00 | 275.6 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | report-workbook-datatable | EPPlus | 105.46 ms | 3.05 ms | 1.76 ms | 4.50 | 4.50 | 65994.6 KB | 3.40 | 161.8 KB | 0.59 | 350.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook-datatable-core | OfficeIMO.Excel | 6.04 ms | 0.21 ms | 0.12 ms | 1.00 | 1.00 | 2982.7 KB | 1.00 | 187.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | report-workbook-datatable-core | EPPlus | 92.36 ms | 3.90 ms | 2.25 ms | 15.30 | 15.30 | 60480.1 KB | 20.28 | 157.2 KB | 0.84 | 1430.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook-datatable-core | ClosedXML | 105.59 ms | 5.40 ms | 3.12 ms | 17.49 | 17.49 | 82860.8 KB | 27.78 | 165.1 KB | 0.88 | 1649.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | LargeXlsx | 6.23 ms | 0.43 ms | 0.25 ms | 0.90 | 1.00 | 857.6 KB | 0.51 | 237.7 KB | 1.10 | 10.4% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | OfficeIMO.Excel | 6.95 ms | 1.93 ms | 1.11 ms | 1.00 | 1.12 | 1676.8 KB | 1.00 | 216.7 KB | 1.00 | Loss +11.6% |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | MiniExcel | 30.59 ms | 9.03 ms | 5.21 ms | 4.40 | 4.91 | 35918.7 KB | 21.42 | 235.3 KB | 1.09 | 340.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | ClosedXML | 118.88 ms | 22.28 ms | 12.87 ms | 17.11 | 19.10 | 71478.2 KB | 42.63 | 257.2 KB | 1.19 | 1611.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-bulk-report | OfficeIMO.Excel | 4.24 ms | 0.06 ms | 0.03 ms | 1.00 | 1.00 | 1401.7 KB | 1.00 | 143.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-bulk-report | MiniExcel | 8.01 ms | 0.19 ms | 0.11 ms | 1.89 | 1.89 | 26825.3 KB | 19.14 | 153.8 KB | 1.07 | 88.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-bulk-report | EPPlus | 68.09 ms | 2.83 ms | 1.63 ms | 16.06 | 16.06 | 47193.8 KB | 33.67 | 115.0 KB | 0.80 | 1505.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-bulk-report | ClosedXML | 74.11 ms | 1.07 ms | 0.62 ms | 17.47 | 17.47 | 58348.8 KB | 41.63 | 121.0 KB | 0.84 | 1647.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellformula | OfficeIMO.Excel | 2.38 ms | 0.04 ms | 0.03 ms | 1.00 | 1.00 | 1383.3 KB | 1.00 | 66.6 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellformula | ClosedXML | 19.85 ms | 3.44 ms | 1.99 ms | 8.34 | 8.34 | 12039.8 KB | 8.70 | 70.6 KB | 1.06 | 733.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellformula | EPPlus | 37.48 ms | 0.44 ms | 0.25 ms | 15.74 | 15.74 | 18110.5 KB | 13.09 | 62.1 KB | 0.93 | 1474.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-empty-strings | OfficeIMO.Excel | 3.06 ms | 0.76 ms | 0.44 ms | 1.00 | 1.00 | 1787.1 KB | 1.00 | 44.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-empty-strings | ClosedXML | 20.92 ms | 5.23 ms | 3.02 ms | 6.85 | 6.85 | 9959.5 KB | 5.57 | 44.9 KB | 1.02 | 584.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-empty-strings | EPPlus | 34.60 ms | 5.33 ms | 3.08 ms | 11.33 | 11.33 | 11773.0 KB | 6.59 | 42.0 KB | 0.95 | 1032.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-numbers | OfficeIMO.Excel | 2.39 ms | 0.26 ms | 0.15 ms | 1.00 | 1.00 | 1119.9 KB | 1.00 | 47.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-numbers | ClosedXML | 15.61 ms | 1.96 ms | 1.13 ms | 6.53 | 6.53 | 9177.1 KB | 8.19 | 45.9 KB | 0.98 | 552.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-numbers | EPPlus | 29.60 ms | 4.14 ms | 2.39 ms | 12.38 | 12.38 | 12895.3 KB | 11.51 | 43.7 KB | 0.93 | 1137.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-mixed | OfficeIMO.Excel | 2.58 ms | 0.19 ms | 0.11 ms | 1.00 | 1.00 | 1763.3 KB | 1.00 | 61.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-object-mixed | ClosedXML | 21.69 ms | 2.01 ms | 1.16 ms | 8.42 | 8.42 | 11887.0 KB | 6.74 | 59.5 KB | 0.97 | 742.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-mixed | EPPlus | 31.93 ms | 4.55 ms | 2.63 ms | 12.39 | 12.39 | 15643.4 KB | 8.87 | 58.9 KB | 0.96 | 1139.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse | OfficeIMO.Excel | 2.94 ms | 0.52 ms | 0.30 ms | 1.00 | 1.00 | 1506.9 KB | 1.00 | 62.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse | ClosedXML | 17.41 ms | 3.01 ms | 1.74 ms | 5.93 | 5.93 | 11296.3 KB | 7.50 | 52.5 KB | 0.85 | 492.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse | EPPlus | 26.56 ms | 1.13 ms | 0.65 ms | 9.04 | 9.04 | 14960.3 KB | 9.93 | 54.2 KB | 0.88 | 804.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 2.46 ms | 0.07 ms | 0.04 ms | 1.00 | 1.00 | 1507.0 KB | 1.00 | 62.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | ClosedXML | 16.73 ms | 2.09 ms | 1.20 ms | 6.80 | 6.80 | 11296.3 KB | 7.50 | 52.5 KB | 0.85 | 579.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | EPPlus | 27.12 ms | 0.82 ms | 0.47 ms | 11.02 | 11.02 | 14960.3 KB | 9.93 | 54.2 KB | 0.88 | 1001.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-scalars | OfficeIMO.Excel | 2.54 ms | 0.29 ms | 0.17 ms | 1.00 | 1.00 | 1138.1 KB | 1.00 | 46.3 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-scalars | ClosedXML | 16.62 ms | 2.27 ms | 1.31 ms | 6.53 | 6.53 | 9021.2 KB | 7.93 | 45.4 KB | 0.98 | 553.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-scalars | EPPlus | 37.06 ms | 2.06 ms | 1.19 ms | 14.57 | 14.57 | 12827.5 KB | 11.27 | 42.4 KB | 0.91 | 1356.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings | OfficeIMO.Excel | 3.73 ms | 0.59 ms | 0.34 ms | 1.00 | 1.00 | 2617.0 KB | 1.00 | 55.1 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-strings | EPPlus | 28.71 ms | 2.99 ms | 1.73 ms | 7.70 | 7.70 | 12804.9 KB | 4.89 | 48.1 KB | 0.87 | 669.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings | ClosedXML | 28.87 ms | 12.91 ms | 7.45 ms | 7.74 | 7.74 | 11299.2 KB | 4.32 | 50.3 KB | 0.91 | 673.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-distinct | OfficeIMO.Excel | 3.19 ms | 0.22 ms | 0.13 ms | 1.00 | 1.00 | 2379.2 KB | 1.00 | 51.8 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-distinct | ClosedXML | 25.17 ms | 1.18 ms | 0.68 ms | 7.90 | 7.90 | 13127.1 KB | 5.52 | 61.9 KB | 1.19 | 690.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-distinct | EPPlus | 42.93 ms | 12.59 ms | 7.27 ms | 13.48 | 13.48 | 13893.0 KB | 5.84 | 61.5 KB | 1.19 | 1247.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-repeated | OfficeIMO.Excel | 2.75 ms | 0.55 ms | 0.32 ms | 1.00 | 1.00 | 1579.8 KB | 1.00 | 40.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-repeated | ClosedXML | 17.78 ms | 1.08 ms | 0.62 ms | 6.48 | 6.48 | 9226.5 KB | 5.84 | 38.8 KB | 0.97 | 547.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-repeated | EPPlus | 31.75 ms | 2.50 ms | 1.44 ms | 11.56 | 11.56 | 11332.5 KB | 7.17 | 34.8 KB | 0.87 | 1056.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-temporal | OfficeIMO.Excel | 3.64 ms | 0.28 ms | 0.16 ms | 1.00 | 1.00 | 1435.7 KB | 1.00 | 63.3 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-temporal | ClosedXML | 18.73 ms | 1.17 ms | 0.67 ms | 5.14 | 5.14 | 9711.1 KB | 6.76 | 54.5 KB | 0.86 | 413.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-temporal | EPPlus | 35.39 ms | 4.99 ms | 2.88 ms | 9.71 | 9.71 | 14722.7 KB | 10.25 | 53.1 KB | 0.84 | 871.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 1.51 ms | 0.45 ms | 0.26 ms | 1.00 | 1.00 | 447.0 KB | 0.41 | 47.3 KB | 0.98 | Tie vs OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 1.51 ms | 0.02 ms | 0.01 ms | 1.00 | 1.00 | 1092.0 KB | 1.00 | 48.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | ClosedXML | 13.63 ms | 0.13 ms | 0.07 ms | 9.06 | 9.06 | 10235.8 KB | 9.37 | 53.0 KB | 1.10 | 805.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | EPPlus | 23.86 ms | 0.61 ms | 0.35 ms | 15.86 | 15.86 | 13052.1 KB | 11.95 | 52.5 KB | 1.09 | 1485.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | LargeXlsx | 3.18 ms | 0.06 ms | 0.04 ms | 0.72 | 1.00 | 758.3 KB | 0.36 | 138.4 KB | 1.00 | 28.3% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 4.44 ms | 0.50 ms | 0.29 ms | 1.00 | 1.39 | 2081.1 KB | 1.00 | 138.0 KB | 1.00 | Loss +39.4% |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | MiniExcel | 8.55 ms | 0.38 ms | 0.22 ms | 1.93 | 2.69 | 23222.1 KB | 11.16 | 153.7 KB | 1.11 | 92.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | ClosedXML | 35.02 ms | 2.54 ms | 1.47 ms | 7.89 | 11.00 | 22221.3 KB | 10.68 | 120.1 KB | 0.87 | 688.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | EPPlus | 41.90 ms | 1.15 ms | 0.66 ms | 9.44 | 13.16 | 24694.0 KB | 11.87 | 114.1 KB | 0.83 | 843.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | Sylvan.Data.Excel | 3.13 ms | 0.13 ms | 0.08 ms | 0.75 | 1.00 | 758.7 KB | 0.43 | 78.5 KB | 0.57 | 24.6% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | LargeXlsx | 3.73 ms | 0.09 ms | 0.05 ms | 0.90 | 1.19 | 1032.5 KB | 0.59 | 138.4 KB | 1.00 | 10.2% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | OfficeIMO.Excel | 4.15 ms | 0.11 ms | 0.07 ms | 1.00 | 1.33 | 1763.0 KB | 1.00 | 138.0 KB | 1.00 | Loss +32.7% |
| 2500 | package-profile | package | Package size | write-datareader-plain | MiniExcel | 7.41 ms | 0.07 ms | 0.04 ms | 1.79 | 2.37 | 23043.8 KB | 13.07 | 153.6 KB | 1.11 | 78.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | ClosedXML | 27.71 ms | 0.72 ms | 0.42 ms | 6.68 | 8.86 | 11581.0 KB | 6.57 | 120.1 KB | 0.87 | 567.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | EPPlus | 38.77 ms | 1.52 ms | 0.88 ms | 9.35 | 12.40 | 16646.4 KB | 9.44 | 114.9 KB | 0.83 | 834.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table | OfficeIMO.Excel | 4.19 ms | 0.11 ms | 0.07 ms | 1.00 | 1.00 | 1774.9 KB | 1.00 | 138.8 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-datareader-table | MiniExcel | 7.35 ms | 0.29 ms | 0.17 ms | 1.75 | 1.75 | 23044.0 KB | 12.98 | 153.6 KB | 1.11 | 75.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table | ClosedXML | 35.81 ms | 2.12 ms | 1.22 ms | 8.55 | 8.55 | 19007.9 KB | 10.71 | 120.9 KB | 0.87 | 755.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table | EPPlus | 37.12 ms | 1.05 ms | 0.60 ms | 8.86 | 8.86 | 16646.1 KB | 9.38 | 114.9 KB | 0.83 | 786.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | OfficeIMO.Excel | 4.54 ms | 0.18 ms | 0.11 ms | 1.00 | 1.00 | 1781.2 KB | 1.00 | 138.9 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | MiniExcel | 7.44 ms | 0.10 ms | 0.06 ms | 1.64 | 1.64 | 26647.2 KB | 14.96 | 153.8 KB | 1.11 | 64.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | EPPlus | 57.57 ms | 4.92 ms | 2.84 ms | 12.69 | 12.69 | 38343.6 KB | 21.53 | 115.1 KB | 0.83 | 1169.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | ClosedXML | 73.07 ms | 0.88 ms | 0.51 ms | 16.11 | 16.11 | 58361.4 KB | 32.77 | 121.0 KB | 0.87 | 1510.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | OfficeIMO.Excel | 3.95 ms | 0.07 ms | 0.04 ms | 1.00 | 1.00 | 2140.6 KB | 1.00 | 131.1 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | LargeXlsx | 4.67 ms | 0.51 ms | 0.29 ms | 1.18 | 1.18 | 1123.9 KB | 0.53 | 164.2 KB | 1.25 | 18.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | MiniExcel | 10.26 ms | 0.86 ms | 0.50 ms | 2.60 | 2.60 | 29746.9 KB | 13.90 | 180.5 KB | 1.38 | 160.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | ClosedXML | 53.85 ms | 2.37 ms | 1.37 ms | 13.64 | 13.64 | 27410.3 KB | 12.80 | 159.4 KB | 1.22 | 1264.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | EPPlus | 58.83 ms | 8.36 ms | 4.82 ms | 14.90 | 14.90 | 21889.7 KB | 10.23 | 144.5 KB | 1.10 | 1390.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | OfficeIMO.Excel | 4.77 ms | 0.20 ms | 0.12 ms | 1.00 | 1.00 | 2880.2 KB | 1.00 | 176.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | MiniExcel | 10.26 ms | 0.67 ms | 0.39 ms | 2.15 | 2.15 | 29746.9 KB | 10.33 | 180.5 KB | 1.03 | 115.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | ClosedXML | 53.79 ms | 1.67 ms | 0.96 ms | 11.28 | 11.28 | 27409.3 KB | 9.52 | 159.4 KB | 0.91 | 1027.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | EPPlus | 63.17 ms | 4.47 ms | 2.58 ms | 13.24 | 13.24 | 21889.7 KB | 7.60 | 144.5 KB | 0.82 | 1224.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables | OfficeIMO.Excel | 3.98 ms | 0.12 ms | 0.07 ms | 1.00 | 1.00 | 2066.1 KB | 1.00 | 138.9 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-dataset-tables | MiniExcel | 7.91 ms | 0.04 ms | 0.02 ms | 1.99 | 1.99 | 28700.4 KB | 13.89 | 156.4 KB | 1.13 | 98.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables | ClosedXML | 35.57 ms | 0.96 ms | 0.56 ms | 8.94 | 8.94 | 18876.9 KB | 9.14 | 123.4 KB | 0.89 | 794.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables | EPPlus | 38.03 ms | 1.53 ms | 0.89 ms | 9.56 | 9.56 | 18700.6 KB | 9.05 | 116.6 KB | 0.84 | 855.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | OfficeIMO.Excel | 4.33 ms | 0.23 ms | 0.13 ms | 1.00 | 1.00 | 2078.7 KB | 1.00 | 139.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | MiniExcel | 8.49 ms | 0.49 ms | 0.28 ms | 1.96 | 1.96 | 31798.5 KB | 15.30 | 156.6 KB | 1.13 | 95.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | EPPlus | 60.29 ms | 6.97 ms | 4.02 ms | 13.91 | 13.91 | 41455.7 KB | 19.94 | 116.9 KB | 0.84 | 1291.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | ClosedXML | 76.08 ms | 6.71 ms | 3.88 ms | 17.56 | 17.56 | 56708.2 KB | 27.28 | 123.7 KB | 0.89 | 1655.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-direct | LargeXlsx | 3.55 ms | 0.05 ms | 0.03 ms | 0.89 | 1.00 | 1149.0 KB | 0.66 | 138.4 KB | 1.00 | 11.1% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-direct | OfficeIMO.Excel | 3.99 ms | 0.15 ms | 0.09 ms | 1.00 | 1.12 | 1748.6 KB | 1.00 | 138.0 KB | 1.00 | Loss +12.5% |
| 2500 | package-profile | package | Package size | write-datatable-direct | MiniExcel | 8.08 ms | 0.14 ms | 0.08 ms | 2.03 | 2.28 | 23062.5 KB | 13.19 | 153.7 KB | 1.11 | 102.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-direct | ClosedXML | 27.66 ms | 0.41 ms | 0.24 ms | 6.94 | 7.80 | 11581.0 KB | 6.62 | 120.1 KB | 0.87 | 593.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-direct | EPPlus | 37.18 ms | 0.65 ms | 0.37 ms | 9.32 | 10.49 | 16646.1 KB | 9.52 | 114.9 KB | 0.83 | 832.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | OfficeIMO.Excel | 3.92 ms | 0.16 ms | 0.09 ms | 1.00 | 1.00 | 1760.7 KB | 1.00 | 138.8 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | MiniExcel | 7.99 ms | 0.14 ms | 0.08 ms | 2.04 | 2.04 | 23062.8 KB | 13.10 | 153.7 KB | 1.11 | 103.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | ClosedXML | 36.52 ms | 1.64 ms | 0.95 ms | 9.31 | 9.31 | 19008.3 KB | 10.80 | 120.9 KB | 0.87 | 831.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | EPPlus | 37.11 ms | 0.25 ms | 0.15 ms | 9.46 | 9.46 | 16646.1 KB | 9.45 | 114.9 KB | 0.83 | 846.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | LargeXlsx | 3.12 ms | 0.20 ms | 0.12 ms | 0.80 | 1.00 | 758.3 KB | 0.43 | 138.4 KB | 1.00 | 20.0% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 3.90 ms | 0.14 ms | 0.08 ms | 1.00 | 1.25 | 1769.2 KB | 1.00 | 138.0 KB | 1.00 | Loss +25.0% |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | MiniExcel | 8.20 ms | 0.76 ms | 0.44 ms | 2.10 | 2.63 | 23222.2 KB | 13.13 | 153.7 KB | 1.11 | 110.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | ClosedXML | 30.76 ms | 2.76 ms | 1.59 ms | 7.88 | 9.86 | 11581.0 KB | 6.55 | 120.1 KB | 0.87 | 688.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | EPPlus | 40.22 ms | 1.78 ms | 1.03 ms | 10.31 | 12.89 | 16646.4 KB | 9.41 | 114.9 KB | 0.83 | 930.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.05 ms | 0.03 ms | 0.02 ms | 1.00 | 1.00 | 1347.1 KB | 1.00 | 142.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 56.19 ms | 3.79 ms | 2.19 ms | 13.89 | 13.89 | 38343.9 KB | 28.46 | 115.1 KB | 0.81 | 1289.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 67.02 ms | 0.06 ms | 0.03 ms | 16.57 | 16.57 | 50927.5 KB | 37.80 | 120.2 KB | 0.84 | 1556.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | LargeXlsx | 3.54 ms | 0.48 ms | 0.28 ms | 0.87 | 1.00 | 758.3 KB | 0.57 | 138.4 KB | 0.97 | 13.3% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | OfficeIMO.Excel | 4.08 ms | 0.58 ms | 0.33 ms | 1.00 | 1.15 | 1339.3 KB | 1.00 | 142.3 KB | 1.00 | Loss +15.4% |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | MiniExcel | 8.02 ms | 0.22 ms | 0.13 ms | 1.96 | 2.26 | 23222.2 KB | 17.34 | 153.7 KB | 1.08 | 96.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | ClosedXML | 31.89 ms | 1.49 ms | 0.86 ms | 7.81 | 9.01 | 11581.0 KB | 8.65 | 120.1 KB | 0.84 | 680.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | EPPlus | 38.15 ms | 1.14 ms | 0.66 ms | 9.34 | 10.77 | 16646.1 KB | 12.43 | 114.9 KB | 0.81 | 834.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 5.42 ms | 0.96 ms | 0.56 ms | 1.00 | 1.00 | 1505.3 KB | 1.00 | 138.1 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 56.70 ms | 2.06 ms | 1.19 ms | 10.46 | 10.46 | 38343.9 KB | 25.47 | 115.1 KB | 0.83 | 946.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 68.16 ms | 4.77 ms | 2.75 ms | 12.58 | 12.58 | 50927.5 KB | 33.83 | 120.2 KB | 0.87 | 1157.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 3.23 ms | 0.20 ms | 0.12 ms | 0.76 | 1.00 | 758.3 KB | 0.51 | 138.4 KB | 1.00 | 24.0% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 4.24 ms | 0.28 ms | 0.16 ms | 1.00 | 1.31 | 1497.5 KB | 1.00 | 138.0 KB | 1.00 | Loss +31.5% |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | MiniExcel | 8.31 ms | 0.49 ms | 0.28 ms | 1.96 | 2.58 | 23222.2 KB | 15.51 | 153.7 KB | 1.11 | 95.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | ClosedXML | 28.40 ms | 1.17 ms | 0.68 ms | 6.69 | 8.80 | 11581.0 KB | 7.73 | 120.1 KB | 0.87 | 569.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | EPPlus | 41.11 ms | 4.56 ms | 2.64 ms | 9.69 | 12.74 | 16646.1 KB | 11.12 | 114.9 KB | 0.83 | 869.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | LargeXlsx | 3.12 ms | 0.04 ms | 0.02 ms | 0.66 | 1.00 | 758.3 KB | 0.43 | 138.4 KB | 0.97 | 34.0% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | OfficeIMO.Excel | 4.73 ms | 0.28 ms | 0.16 ms | 1.00 | 1.51 | 1770.1 KB | 1.00 | 142.3 KB | 1.00 | Loss +51.5% |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | MiniExcel | 8.14 ms | 0.53 ms | 0.31 ms | 1.72 | 2.61 | 23222.2 KB | 13.12 | 153.7 KB | 1.08 | 72.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | ClosedXML | 27.60 ms | 1.24 ms | 0.72 ms | 5.83 | 8.83 | 11581.0 KB | 6.54 | 120.1 KB | 0.84 | 483.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | EPPlus | 37.43 ms | 1.48 ms | 0.85 ms | 7.91 | 11.98 | 16646.1 KB | 9.40 | 114.9 KB | 0.81 | 690.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.24 ms | 0.51 ms | 0.29 ms | 1.00 | 1.00 | 1346.4 KB | 1.00 | 142.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 46.87 ms | 3.61 ms | 2.09 ms | 11.05 | 11.05 | 28540.6 KB | 21.20 | 120.2 KB | 0.84 | 1004.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 49.91 ms | 1.57 ms | 0.90 ms | 11.77 | 11.77 | 27305.8 KB | 20.28 | 115.0 KB | 0.81 | 1076.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | LargeXlsx | 4.30 ms | 0.18 ms | 0.10 ms | 0.74 | 1.00 | 802.5 KB | 0.34 | 182.6 KB | 1.00 | 25.6% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 5.79 ms | 0.87 ms | 0.50 ms | 1.00 | 1.34 | 2341.7 KB | 1.00 | 183.1 KB | 1.00 | Loss +34.4% |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | MiniExcel | 8.84 ms | 0.43 ms | 0.25 ms | 1.53 | 2.05 | 25190.5 KB | 10.76 | 194.0 KB | 1.06 | 52.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | ClosedXML | 36.32 ms | 0.50 ms | 0.29 ms | 6.28 | 8.44 | 16973.5 KB | 7.25 | 161.0 KB | 0.88 | 527.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | EPPlus | 48.77 ms | 0.49 ms | 0.29 ms | 8.43 | 11.33 | 20105.1 KB | 8.59 | 152.1 KB | 0.83 | 743.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | LargeXlsx | 4.24 ms | 0.25 ms | 0.15 ms | 0.95 | 1.00 | 802.5 KB | 0.53 | 182.6 KB | 1.00 | 5.1% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 4.47 ms | 0.25 ms | 0.14 ms | 1.00 | 1.05 | 1507.7 KB | 1.00 | 182.4 KB | 1.00 | Loss +5.3% |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | MiniExcel | 8.56 ms | 0.38 ms | 0.22 ms | 1.92 | 2.02 | 25190.5 KB | 16.71 | 194.0 KB | 1.06 | 91.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | ClosedXML | 35.02 ms | 0.63 ms | 0.37 ms | 7.84 | 8.26 | 16973.5 KB | 11.26 | 161.0 KB | 0.88 | 684.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | EPPlus | 48.37 ms | 2.02 ms | 1.17 ms | 10.83 | 11.41 | 20105.1 KB | 13.33 | 152.1 KB | 0.83 | 983.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 20.44 ms | 1.82 ms | 1.05 ms | 1.00 | 1.00 | 4502.3 KB | 1.00 | 651.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | LargeXlsx | 21.67 ms | 1.53 ms | 0.88 ms | 1.06 | 1.06 | 2810.7 KB | 0.62 | 644.6 KB | 0.99 | 6.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | MiniExcel | 36.87 ms | 2.48 ms | 1.43 ms | 1.80 | 1.80 | 48414.8 KB | 10.75 | 674.4 KB | 1.04 | 80.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | ClosedXML | 122.71 ms | 5.69 ms | 3.28 ms | 6.00 | 6.00 | 51647.0 KB | 11.47 | 615.5 KB | 0.95 | 500.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | EPPlus | 161.19 ms | 4.66 ms | 2.69 ms | 7.89 | 7.89 | 69139.6 KB | 15.36 | 548.9 KB | 0.84 | 688.7% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | OfficeIMO.Excel | 9.89 ms | 0.36 ms | 0.21 ms | 1.00 | 1.00 | 1895.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | EPPlus | 76.06 ms | 4.16 ms | 2.40 ms | 7.69 | 7.69 | 50712.0 KB | 26.76 |  |  | 668.7% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | EPPlus 4.5.3.3 | 93.36 ms |  |  | 9.44 | 9.44 |  |  |  |  | 843.5% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | ClosedXML | 129.48 ms | 1.95 ms | 1.13 ms | 13.09 | 13.09 | 84615.4 KB | 44.65 |  |  | 1208.6% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | OfficeIMO.Excel | 15.66 ms | 0.54 ms | 0.31 ms | 1.00 | 1.00 | 15209.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus 4.5.3.3 | 35.21 ms |  |  | 2.25 | 2.25 |  |  |  |  | 124.8% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus | 45.36 ms | 0.62 ms | 0.36 ms | 2.90 | 2.90 | 32906.1 KB | 2.16 |  |  | 189.6% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | OfficeIMO.Excel | 11.72 ms | 0.40 ms | 0.23 ms | 1.00 | 1.00 | 6195.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus | 74.98 ms | 4.72 ms | 2.72 ms | 6.40 | 6.40 | 54593.4 KB | 8.81 |  |  | 539.6% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus 4.5.3.3 | 81.24 ms |  |  | 6.93 | 6.93 |  |  |  |  | 593.0% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | OfficeIMO.Excel | 18.72 ms | 1.59 ms | 0.92 ms | 1.00 | 1.00 | 16350.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus | 81.06 ms | 1.53 ms | 0.89 ms | 4.33 | 4.33 | 59225.9 KB | 3.62 |  |  | 333.0% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus 4.5.3.3 | 86.64 ms |  |  | 4.63 | 4.63 |  |  |  |  | 362.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | OfficeIMO.Excel | 16.75 ms | 0.29 ms | 0.17 ms | 1.00 | 1.00 | 15230.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus | 69.78 ms | 2.42 ms | 1.40 ms | 4.17 | 4.17 | 54594.0 KB | 3.58 |  |  | 316.6% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus 4.5.3.3 | 70.97 ms |  |  | 4.24 | 4.24 |  |  |  |  | 323.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | OfficeIMO.Excel | 17.87 ms | 0.66 ms | 0.38 ms | 1.00 | 1.00 | 15225.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus | 72.01 ms | 2.04 ms | 1.18 ms | 4.03 | 4.03 | 54590.7 KB | 3.59 |  |  | 302.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus 4.5.3.3 | 75.53 ms |  |  | 4.23 | 4.23 |  |  |  |  | 322.6% slower than OfficeIMO |
| 2500 | speed-comparison | other | Object projection | build-object-datatable-dictionaries | OfficeIMO.Excel | 1.35 ms | 0.03 ms | 0.02 ms | 1.00 | 1.00 | 564.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Object projection | build-object-datatable-typed | OfficeIMO.Excel | 1.18 ms | 0.07 ms | 0.04 ms | 1.00 | 1.00 | 856.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-cells | OfficeIMO.Excel | 6.06 ms | 1.47 ms | 0.85 ms | 1.00 | 1.00 | 2531.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-cells | EPPlus | 28.00 ms | 0.57 ms | 0.33 ms | 4.62 | 4.62 | 20154.9 KB | 7.96 |  |  | 362.4% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-cells | ClosedXML | 32.33 ms | 0.58 ms | 0.34 ms | 5.34 | 5.34 | 17022.4 KB | 6.72 |  |  | 433.8% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | OfficeIMO.Excel | 4.33 ms | 0.44 ms | 0.25 ms | 1.00 | 1.00 | 523.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | EPPlus | 28.29 ms | 1.49 ms | 0.86 ms | 6.54 | 6.54 | 13108.1 KB | 25.04 |  |  | 553.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | ClosedXML | 32.57 ms | 1.42 ms | 0.82 ms | 7.53 | 7.53 | 15463.7 KB | 29.54 |  |  | 652.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-range | OfficeIMO.Excel | 6.16 ms | 1.04 ms | 0.60 ms | 1.00 | 1.00 | 2531.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-range | EPPlus | 26.99 ms | 0.92 ms | 0.53 ms | 4.38 | 4.38 | 20154.9 KB | 7.96 |  |  | 338.0% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-range | ClosedXML | 31.77 ms | 1.12 ms | 0.64 ms | 5.16 | 5.16 | 17020.9 KB | 6.72 |  |  | 415.6% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-top-range | OfficeIMO.Excel | 0.69 ms | 0.12 ms | 0.07 ms | 1.00 | 1.00 | 285.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-top-range | EPPlus | 24.05 ms | 0.52 ms | 0.30 ms | 35.11 | 35.11 | 12404.4 KB | 43.47 |  |  | 3410.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-top-range | ClosedXML | 31.39 ms | 0.78 ms | 0.45 ms | 45.82 | 45.82 | 15370.6 KB | 53.86 |  |  | 4482.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-autofilter | OfficeIMO.Excel | 3.68 ms | 0.08 ms | 0.04 ms | 1.00 | 1.00 | 1340.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-autofilter | EPPlus 4.5.3.3 | 26.80 ms |  |  | 7.28 | 7.28 |  |  |  |  | 627.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-autofilter | ClosedXML | 32.55 ms | 0.98 ms | 0.57 ms | 8.84 | 8.84 | 22226.8 KB | 16.58 |  |  | 783.8% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-autofilter | EPPlus | 42.40 ms | 3.35 ms | 1.93 ms | 11.51 | 11.51 | 24715.5 KB | 18.44 |  |  | 1051.1% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-charts | OfficeIMO.Excel | 5.59 ms | 0.42 ms | 0.24 ms | 1.00 | 1.00 | 1892.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-charts | EPPlus 4.5.3.3 | 28.24 ms |  |  | 5.06 | 5.06 |  |  |  |  | 405.6% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-charts | EPPlus | 45.97 ms | 2.55 ms | 1.47 ms | 8.23 | 8.23 | 27141.8 KB | 14.34 |  |  | 723.1% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | OfficeIMO.Excel | 3.77 ms | 0.06 ms | 0.03 ms | 1.00 | 1.00 | 1405.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | EPPlus 4.5.3.3 | 27.43 ms |  |  | 7.27 | 7.27 |  |  |  |  | 626.8% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | ClosedXML | 31.31 ms | 0.76 ms | 0.44 ms | 8.30 | 8.30 | 22273.8 KB | 15.84 |  |  | 729.8% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | EPPlus | 41.11 ms | 0.92 ms | 0.53 ms | 10.89 | 10.89 | 24757.5 KB | 17.61 |  |  | 989.5% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-data-validation | OfficeIMO.Excel | 3.64 ms | 0.04 ms | 0.02 ms | 1.00 | 1.00 | 1356.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-data-validation | EPPlus 4.5.3.3 | 29.17 ms |  |  | 8.02 | 8.02 |  |  |  |  | 702.0% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-data-validation | ClosedXML | 30.93 ms | 0.62 ms | 0.36 ms | 8.51 | 8.51 | 22247.9 KB | 16.41 |  |  | 750.6% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-data-validation | EPPlus | 39.76 ms | 1.12 ms | 0.65 ms | 10.93 | 10.93 | 24701.4 KB | 18.22 |  |  | 993.4% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | OfficeIMO.Excel | 3.89 ms | 0.19 ms | 0.11 ms | 1.00 | 1.00 | 1342.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | EPPlus 4.5.3.3 | 27.35 ms |  |  | 7.02 | 7.02 |  |  |  |  | 602.2% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | ClosedXML | 32.34 ms | 2.12 ms | 1.22 ms | 8.30 | 8.30 | 22222.0 KB | 16.55 |  |  | 730.4% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | EPPlus | 42.87 ms | 5.16 ms | 2.98 ms | 11.01 | 11.01 | 24730.0 KB | 18.42 |  |  | 1000.8% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | OfficeIMO.Excel | 17.61 ms | 1.46 ms | 0.84 ms | 1.00 | 1.00 | 14419.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | EPPlus 4.5.3.3 | 25.77 ms |  |  | 1.46 | 1.46 |  |  |  |  | 46.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | EPPlus | 51.21 ms | 4.03 ms | 2.33 ms | 2.91 | 2.91 | 29537.1 KB | 2.05 |  |  | 190.8% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world report | realworld-report-all-in-one | OfficeIMO.Excel | 17.09 ms | 0.73 ms | 0.42 ms | 1.00 | 1.00 | 15220.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus 4.5.3.3 | 72.25 ms |  |  | 4.23 | 4.23 |  |  |  |  | 322.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus | 72.35 ms | 4.01 ms | 2.31 ms | 4.23 | 4.23 | 54594.5 KB | 3.59 |  |  | 323.5% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world report | realworld-report-core | OfficeIMO.Excel | 4.44 ms | 0.32 ms | 0.18 ms | 1.00 | 1.00 | 1488.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world report | realworld-report-core | EPPlus 4.5.3.3 | 64.30 ms |  |  | 14.49 | 14.49 |  |  |  |  | 1349.1% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world report | realworld-report-core | EPPlus | 67.85 ms | 2.58 ms | 1.49 ms | 15.29 | 15.29 | 47299.8 KB | 31.77 |  |  | 1429.1% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world report | realworld-report-core | ClosedXML | 77.24 ms | 1.88 ms | 1.09 ms | 17.41 | 17.41 | 69834.2 KB | 46.91 |  |  | 1640.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook | OfficeIMO.Excel | 30.29 ms | 3.57 ms | 2.06 ms | 1.00 | 1.00 | 19069.0 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Report workbook | report-workbook | EPPlus 4.5.3.3 | 76.31 ms |  |  | 2.52 | 2.52 |  |  |  |  | 151.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook | EPPlus | 93.24 ms | 3.28 ms | 1.90 ms | 3.08 | 3.08 | 77486.1 KB | 4.06 |  |  | 207.8% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-core | OfficeIMO.Excel | 6.31 ms | 0.12 ms | 0.07 ms | 1.00 | 1.00 | 2711.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Report workbook | report-workbook-core | EPPlus 4.5.3.3 | 66.16 ms |  |  | 10.49 | 10.49 |  |  |  |  | 948.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-core | EPPlus | 97.68 ms | 4.08 ms | 2.35 ms | 15.48 | 15.48 | 71970.6 KB | 26.55 |  |  | 1448.5% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-core | ClosedXML | 105.25 ms | 2.96 ms | 1.71 ms | 16.69 | 16.69 | 97220.1 KB | 35.86 |  |  | 1568.6% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable | OfficeIMO.Excel | 22.91 ms | 0.97 ms | 0.56 ms | 1.00 | 1.00 | 19383.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable | EPPlus 4.5.3.3 | 67.79 ms |  |  | 2.96 | 2.96 |  |  |  |  | 195.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable | EPPlus | 95.98 ms | 1.10 ms | 0.63 ms | 4.19 | 4.19 | 65995.3 KB | 3.40 |  |  | 318.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable-core | OfficeIMO.Excel | 5.94 ms | 0.08 ms | 0.05 ms | 1.00 | 1.00 | 2982.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable-core | EPPlus 4.5.3.3 | 71.02 ms |  |  | 11.95 | 11.95 |  |  |  |  | 1095.1% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable-core | EPPlus | 97.63 ms | 5.67 ms | 3.27 ms | 16.43 | 16.43 | 60480.1 KB | 20.28 |  |  | 1542.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable-core | ClosedXML | 104.58 ms | 1.48 ms | 0.86 ms | 17.60 | 17.60 | 82858.9 KB | 27.78 |  |  | 1660.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | OfficeIMO.Excel | 2.65 ms | 0.40 ms | 0.23 ms | 1.00 | 1.00 | 706.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | EPPlus 4.5.3.3 | 13.21 ms |  |  | 4.98 | 4.98 |  |  |  |  | 398.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | ClosedXML | 18.76 ms | 2.84 ms | 1.64 ms | 7.07 | 7.07 | 8279.6 KB | 11.71 |  |  | 607.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | EPPlus | 24.84 ms | 7.77 ms | 4.48 ms | 9.37 | 9.37 | 7708.0 KB | 10.91 |  |  | 836.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | OfficeIMO.Excel | 0.96 ms | 0.08 ms | 0.04 ms | 1.00 | 1.00 | 177.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | Sylvan.Data.Excel | 1.01 ms | 0.02 ms | 0.01 ms | 1.05 | 1.05 | 316.6 KB | 1.79 |  |  | 5.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | ExcelDataReader | 1.69 ms | 0.35 ms | 0.20 ms | 1.76 | 1.76 | 4062.2 KB | 22.91 |  |  | 76.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | ClosedXML | 3.55 ms | 0.18 ms | 0.11 ms | 3.70 | 3.70 | 4392.9 KB | 24.77 |  |  | 270.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | EPPlus 4.5.3.3 | 11.16 ms |  |  | 11.62 | 11.62 |  |  |  |  | 1062.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | MiniExcel | 11.67 ms | 0.11 ms | 0.06 ms | 12.15 | 12.15 | 46194.9 KB | 260.48 |  |  | 1115.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | EPPlus | 39.05 ms | 1.72 ms | 0.99 ms | 40.65 | 40.65 | 43071.0 KB | 242.86 |  |  | 3965.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | Sylvan.Data.Excel | 1.04 ms | 0.03 ms | 0.02 ms | 0.59 | 1.00 | 316.6 KB | 1.78 |  |  | 41.3% faster than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | ExcelDataReader | 1.56 ms | 0.13 ms | 0.07 ms | 0.88 | 1.50 | 4062.2 KB | 22.89 |  |  | 11.7% faster than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | OfficeIMO.Excel | 1.77 ms | 0.70 ms | 0.41 ms | 1.00 | 1.70 | 177.5 KB | 1.00 |  |  | Loss +70.3% |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | ClosedXML | 3.96 ms | 0.58 ms | 0.33 ms | 2.24 | 3.82 | 4392.8 KB | 24.75 |  |  | 124.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | EPPlus 4.5.3.3 | 11.49 ms |  |  | 6.51 | 11.08 |  |  |  |  | 550.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | MiniExcel | 12.03 ms | 1.03 ms | 0.60 ms | 6.81 | 11.60 | 46194.9 KB | 260.30 |  |  | 581.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | EPPlus | 39.90 ms | 1.84 ms | 1.06 ms | 22.60 | 38.47 | 43071.0 KB | 242.70 |  |  | 2159.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | Sylvan.Data.Excel | 3.31 ms | 0.41 ms | 0.24 ms | 0.77 | 1.00 | 518.6 KB | 0.49 |  |  | 23.2% faster than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | OfficeIMO.Excel | 4.31 ms | 0.10 ms | 0.06 ms | 1.00 | 1.30 | 1056.7 KB | 1.00 |  |  | Loss +30.1% |
| 2500 | speed-comparison | read | Other | shared-string-read | ExcelDataReader | 8.32 ms | 1.86 ms | 1.08 ms | 1.93 | 2.51 | 2619.1 KB | 2.48 |  |  | 93.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | MiniExcel | 8.43 ms | 0.25 ms | 0.15 ms | 1.96 | 2.55 | 7530.0 KB | 7.13 |  |  | 95.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | EPPlus 4.5.3.3 | 13.09 ms |  |  | 3.04 | 3.95 |  |  |  |  | 203.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | ClosedXML | 21.93 ms | 3.40 ms | 1.96 ms | 5.09 | 6.62 | 9498.0 KB | 8.99 |  |  | 409.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | EPPlus | 24.50 ms | 1.22 ms | 0.70 ms | 5.69 | 7.40 | 10372.2 KB | 9.82 |  |  | 468.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | Sylvan.Data.Excel | 4.28 ms | 0.19 ms | 0.11 ms | 0.94 | 1.00 | 655.2 KB | 1.75 |  |  | 5.6% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | OfficeIMO.Excel | 4.54 ms | 1.18 ms | 0.68 ms | 1.00 | 1.06 | 374.7 KB | 1.00 |  |  | Loss +6.0% |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | ExcelDataReader | 10.65 ms | 0.41 ms | 0.24 ms | 2.35 | 2.49 | 6089.3 KB | 16.25 |  |  | 134.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | MiniExcel | 13.34 ms | 0.52 ms | 0.30 ms | 2.94 | 3.11 | 18661.8 KB | 49.81 |  |  | 193.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | EPPlus | 24.49 ms | 2.04 ms | 1.18 ms | 5.39 | 5.72 | 12427.1 KB | 33.17 |  |  | 439.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | ClosedXML | 31.99 ms | 2.38 ms | 1.37 ms | 7.04 | 7.47 | 15361.6 KB | 41.00 |  |  | 604.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | Sylvan.Data.Excel | 6.11 ms | 0.03 ms | 0.02 ms | 0.80 | 1.00 | 2239.3 KB | 0.62 |  |  | 20.3% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | OfficeIMO.Excel | 7.66 ms | 1.89 ms | 1.09 ms | 1.00 | 1.25 | 3594.6 KB | 1.00 |  |  | Loss +25.4% |
| 2500 | speed-comparison | read | Range and table read | read-datatable | ExcelDataReader | 12.83 ms | 0.28 ms | 0.16 ms | 1.68 | 2.10 | 7673.3 KB | 2.13 |  |  | 67.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | MiniExcel | 14.51 ms | 0.27 ms | 0.16 ms | 1.89 | 2.37 | 18266.6 KB | 5.08 |  |  | 89.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | EPPlus | 30.29 ms | 0.99 ms | 0.57 ms | 3.95 | 4.96 | 21736.6 KB | 6.05 |  |  | 295.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | ClosedXML | 34.96 ms | 1.30 ms | 0.75 ms | 4.56 | 5.72 | 18314.3 KB | 5.09 |  |  | 356.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | EPPlus 4.5.3.3 | 36.78 ms |  |  | 4.80 | 6.02 |  |  |  |  | 380.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | Sylvan.Data.Excel | 4.77 ms | 0.20 ms | 0.11 ms | 0.91 | 1.00 | 733.5 KB | 1.35 |  |  | 8.5% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | OfficeIMO.Excel | 5.21 ms | 1.60 ms | 0.92 ms | 1.00 | 1.09 | 543.1 KB | 1.00 |  |  | Loss +9.3% |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | ExcelDataReader | 11.71 ms | 0.35 ms | 0.20 ms | 2.25 | 2.46 | 6089.3 KB | 11.21 |  |  | 124.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | MiniExcel | 13.90 ms | 2.73 ms | 1.58 ms | 2.67 | 2.92 | 15850.3 KB | 29.19 |  |  | 166.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | EPPlus | 26.03 ms | 1.09 ms | 0.63 ms | 4.99 | 5.46 | 13108.1 KB | 24.14 |  |  | 399.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | ClosedXML | 33.42 ms | 3.19 ms | 1.84 ms | 6.41 | 7.01 | 15465.4 KB | 28.48 |  |  | 541.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | Sylvan.Data.Excel | 18.04 ms | 8.31 ms | 4.80 ms | 0.96 | 1.00 | 655.0 KB | 0.24 |  |  | 4.4% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | OfficeIMO.Excel | 18.87 ms | 2.66 ms | 1.54 ms | 1.00 | 1.05 | 2692.7 KB | 1.00 |  |  | Loss +4.6% |
| 2500 | speed-comparison | read | Range and table read | read-range | ExcelDataReader | 30.78 ms | 11.01 ms | 6.36 ms | 1.63 | 1.71 | 6089.2 KB | 2.26 |  |  | 63.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | MiniExcel | 33.01 ms | 6.65 ms | 3.84 ms | 1.75 | 1.83 | 18662.2 KB | 6.93 |  |  | 75.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | EPPlus 4.5.3.3 | 35.98 ms |  |  | 1.91 | 1.99 |  |  |  |  | 90.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | EPPlus | 68.13 ms | 23.23 ms | 13.41 ms | 3.61 | 3.78 | 20152.6 KB | 7.48 |  |  | 261.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | ClosedXML | 125.88 ms | 9.92 ms | 5.73 ms | 6.67 | 6.98 | 16846.3 KB | 6.26 |  |  | 567.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | OfficeIMO.Excel | 5.26 ms | 0.29 ms | 0.17 ms | 1.00 | 1.00 | 2751.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | Sylvan.Data.Excel | 5.36 ms | 0.10 ms | 0.06 ms | 1.02 | 1.02 | 750.3 KB | 0.27 |  |  | Tie vs OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | ExcelDataReader | 10.59 ms | 0.12 ms | 0.07 ms | 2.01 | 2.01 | 6089.3 KB | 2.21 |  |  | 101.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | MiniExcel | 12.90 ms | 0.97 ms | 0.56 ms | 2.45 | 2.45 | 18662.4 KB | 6.78 |  |  | 145.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | EPPlus | 28.19 ms | 1.51 ms | 0.87 ms | 5.35 | 5.35 | 20152.6 KB | 7.32 |  |  | 435.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | ClosedXML | 30.94 ms | 0.41 ms | 0.24 ms | 5.88 | 5.88 | 16728.5 KB | 6.08 |  |  | 487.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | Sylvan.Data.Excel | 0.51 ms | 0.02 ms | 0.01 ms | 0.84 | 1.00 | 348.4 KB | 1.18 |  |  | 16.5% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | OfficeIMO.Excel | 0.61 ms | 0.04 ms | 0.02 ms | 1.00 | 1.20 | 296.2 KB | 1.00 |  |  | Loss +19.7% |
| 2500 | speed-comparison | read | Range and table read | read-top-range | MiniExcel | 0.83 ms | 0.08 ms | 0.05 ms | 1.37 | 1.63 | 869.0 KB | 2.93 |  |  | 36.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | ExcelDataReader | 4.61 ms | 0.25 ms | 0.14 ms | 7.55 | 9.04 | 1931.6 KB | 6.52 |  |  | 655.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | EPPlus | 25.45 ms | 2.50 ms | 1.44 ms | 41.71 | 49.92 | 12402.1 KB | 41.87 |  |  | 4071.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | EPPlus 4.5.3.3 | 27.90 ms |  |  | 45.72 | 54.73 |  |  |  |  | 4472.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | ClosedXML | 32.49 ms | 2.26 ms | 1.31 ms | 53.26 | 63.74 | 15360.6 KB | 51.86 |  |  | 5225.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | Sylvan.Data.Excel | 5.03 ms | 0.31 ms | 0.18 ms | 0.47 | 1.00 | 655.2 KB | 0.19 |  |  | 53.1% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | ExcelDataReader | 10.67 ms | 0.20 ms | 0.11 ms | 1.00 | 2.12 | 6089.3 KB | 1.75 |  |  | Tie vs OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | OfficeIMO.Excel | 10.71 ms | 0.68 ms | 0.39 ms | 1.00 | 2.13 | 3472.7 KB | 1.00 |  |  | Loss +113.1% |
| 2500 | speed-comparison | read | Range and table read | read-used-range | MiniExcel | 13.16 ms | 0.57 ms | 0.33 ms | 1.23 | 2.62 | 18662.4 KB | 5.37 |  |  | 22.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | EPPlus | 30.58 ms | 3.06 ms | 1.77 ms | 2.86 | 6.08 | 20152.8 KB | 5.80 |  |  | 185.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | ClosedXML | 56.90 ms | 41.88 ms | 24.18 ms | 5.31 | 11.32 | 16784.0 KB | 4.83 |  |  | 431.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | OfficeIMO.Excel | 3.76 ms | 0.00 ms | 0.00 ms | 1.00 | 1.00 | 378.0 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | Sylvan.Data.Excel | 4.16 ms | 0.02 ms | 0.01 ms | 1.11 | 1.11 | 655.2 KB | 1.73 |  |  | 10.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | ExcelDataReader | 10.40 ms | 0.17 ms | 0.10 ms | 2.77 | 2.77 | 6089.5 KB | 16.11 |  |  | 176.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | MiniExcel | 12.94 ms | 0.72 ms | 0.41 ms | 3.44 | 3.44 | 18661.8 KB | 49.38 |  |  | 244.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | EPPlus | 24.12 ms | 0.19 ms | 0.11 ms | 6.42 | 6.42 | 12427.1 KB | 32.88 |  |  | 541.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | ClosedXML | 30.29 ms | 0.96 ms | 0.56 ms | 8.06 | 8.06 | 15359.7 KB | 40.64 |  |  | 705.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | Sylvan.Data.Excel | 5.45 ms | 1.09 ms | 0.63 ms | 0.88 | 1.00 | 655.2 KB | 0.24 |  |  | 12.0% faster than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | OfficeIMO.Excel | 6.19 ms | 2.00 ms | 1.16 ms | 1.00 | 1.14 | 2771.5 KB | 1.00 |  |  | Loss +13.6% |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | ExcelDataReader | 11.37 ms | 1.01 ms | 0.58 ms | 1.84 | 2.09 | 6089.4 KB | 2.20 |  |  | 83.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | MiniExcel | 12.77 ms | 0.15 ms | 0.09 ms | 2.06 | 2.35 | 18662.4 KB | 6.73 |  |  | 106.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | EPPlus | 28.07 ms | 2.11 ms | 1.22 ms | 4.54 | 5.15 | 20152.6 KB | 7.27 |  |  | 353.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | EPPlus 4.5.3.3 | 29.80 ms |  |  | 4.82 | 5.47 |  |  |  |  | 381.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | ClosedXML | 31.41 ms | 0.64 ms | 0.37 ms | 5.08 | 5.77 | 16729.8 KB | 6.04 |  |  | 407.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | Sylvan.Data.Excel | 0.45 ms | 0.02 ms | 0.01 ms | 0.67 | 1.00 | 348.5 KB | 1.16 |  |  | 32.9% faster than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | OfficeIMO.Excel | 0.67 ms | 0.12 ms | 0.07 ms | 1.00 | 1.49 | 299.5 KB | 1.00 |  |  | Loss +49.0% |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | MiniExcel | 0.81 ms | 0.09 ms | 0.05 ms | 1.21 | 1.80 | 869.0 KB | 2.90 |  |  | 20.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | ExcelDataReader | 4.51 ms | 0.16 ms | 0.09 ms | 6.74 | 10.04 | 1931.8 KB | 6.45 |  |  | 573.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | EPPlus | 24.47 ms | 2.31 ms | 1.33 ms | 36.59 | 54.53 | 12402.0 KB | 41.40 |  |  | 3558.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | EPPlus 4.5.3.3 | 27.19 ms |  |  | 40.65 | 60.60 |  |  |  |  | 3965.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | ClosedXML | 31.88 ms | 1.84 ms | 1.06 ms | 47.67 | 71.05 | 15363.9 KB | 51.29 |  |  | 4666.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | Sylvan.Data.Excel | 0.43 ms | 0.01 ms | 0.00 ms | 0.78 | 1.00 | 348.5 KB | 1.16 |  |  | 21.9% faster than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | OfficeIMO.Excel | 0.55 ms | 0.02 ms | 0.01 ms | 1.00 | 1.28 | 300.3 KB | 1.00 |  |  | Loss +28.1% |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | MiniExcel | 0.79 ms | 0.10 ms | 0.06 ms | 1.44 | 1.85 | 869.0 KB | 2.89 |  |  | 44.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | ExcelDataReader | 4.44 ms | 0.13 ms | 0.07 ms | 8.06 | 10.32 | 1931.8 KB | 6.43 |  |  | 706.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | EPPlus | 24.52 ms | 1.48 ms | 0.86 ms | 44.54 | 57.04 | 12402.1 KB | 41.30 |  |  | 4354.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | ClosedXML | 32.07 ms | 1.25 ms | 0.72 ms | 58.25 | 74.60 | 15360.9 KB | 51.16 |  |  | 5725.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | OfficeIMO.Excel | 7.55 ms | 2.41 ms | 1.39 ms | 1.00 | 1.00 | 2442.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Typed object read | read-objects | Sylvan.Data.Excel | 7.81 ms | 3.14 ms | 1.82 ms | 1.03 | 1.03 | 895.3 KB | 0.37 |  |  | 3.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | ExcelDataReader | 13.98 ms | 4.36 ms | 2.51 ms | 1.85 | 1.85 | 6329.5 KB | 2.59 |  |  | 85.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | MiniExcel | 16.71 ms | 0.52 ms | 0.30 ms | 2.21 | 2.21 | 18474.1 KB | 7.56 |  |  | 121.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | EPPlus 4.5.3.3 | 30.13 ms |  |  | 3.99 | 3.99 |  |  |  |  | 299.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | EPPlus | 32.08 ms | 4.94 ms | 2.85 ms | 4.25 | 4.25 | 21354.2 KB | 8.74 |  |  | 325.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | ClosedXML | 33.94 ms | 0.69 ms | 0.40 ms | 4.50 | 4.50 | 16925.8 KB | 6.93 |  |  | 349.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | Sylvan.Data.Excel | 5.00 ms | 0.11 ms | 0.06 ms | 0.94 | 1.00 | 831.0 KB | 0.34 |  |  | 6.5% faster than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | OfficeIMO.Excel | 5.35 ms | 0.43 ms | 0.25 ms | 1.00 | 1.07 | 2423.1 KB | 1.00 |  |  | Loss +6.9% |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | ExcelDataReader | 11.00 ms | 1.18 ms | 0.68 ms | 2.06 | 2.20 | 6265.3 KB | 2.59 |  |  | 105.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | MiniExcel | 14.30 ms | 0.99 ms | 0.57 ms | 2.67 | 2.86 | 18409.8 KB | 7.60 |  |  | 167.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | EPPlus | 31.15 ms | 5.35 ms | 3.09 ms | 5.82 | 6.23 | 21334.6 KB | 8.80 |  |  | 482.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | EPPlus 4.5.3.3 | 31.48 ms |  |  | 5.88 | 6.29 |  |  |  |  | 488.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | ClosedXML | 31.76 ms | 1.17 ms | 0.68 ms | 5.94 | 6.35 | 16904.3 KB | 6.98 |  |  | 493.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | OfficeIMO.Excel | 5.07 ms | 0.68 ms | 0.39 ms | 1.00 | 1.00 | 1781.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | MiniExcel | 10.12 ms | 2.38 ms | 1.38 ms | 2.00 | 2.00 | 26647.3 KB | 14.96 |  |  | 99.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | EPPlus | 57.27 ms | 6.18 ms | 3.57 ms | 11.30 | 11.30 | 38343.5 KB | 21.53 |  |  | 1029.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | ClosedXML | 76.39 ms | 2.62 ms | 1.51 ms | 15.07 | 15.07 | 58360.0 KB | 32.77 |  |  | 1407.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | EPPlus 4.5.3.3 | 206.69 ms |  |  | 40.77 | 40.77 |  |  |  |  | 3977.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | OfficeIMO.Excel | 7.29 ms | 0.52 ms | 0.30 ms | 1.00 | 1.00 | 2079.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | MiniExcel | 18.66 ms | 1.25 ms | 0.72 ms | 2.56 | 2.56 | 32328.7 KB | 15.55 |  |  | 155.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | EPPlus | 147.03 ms | 9.96 ms | 5.75 ms | 20.16 | 20.16 | 43440.5 KB | 20.89 |  |  | 1915.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | ClosedXML | 201.58 ms | 12.71 ms | 7.34 ms | 27.64 | 27.64 | 56707.6 KB | 27.27 |  |  | 2663.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.02 ms | 0.07 ms | 0.04 ms | 1.00 | 1.00 | 1347.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 53.84 ms | 2.40 ms | 1.38 ms | 13.39 | 13.39 | 38344.1 KB | 28.46 |  |  | 1239.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 64.62 ms | 1.14 ms | 0.66 ms | 16.07 | 16.07 | 50927.8 KB | 37.81 |  |  | 1507.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.62 ms | 0.05 ms | 0.03 ms | 1.00 | 1.00 | 1505.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 54.67 ms | 3.51 ms | 2.03 ms | 11.84 | 11.84 | 38344.1 KB | 25.47 |  |  | 1084.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 67.78 ms | 4.22 ms | 2.43 ms | 14.68 | 14.68 | 50927.5 KB | 33.83 |  |  | 1368.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 3.90 ms | 0.12 ms | 0.07 ms | 1.00 | 1.00 | 1346.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 44.06 ms | 1.49 ms | 0.86 ms | 11.31 | 11.31 | 28540.6 KB | 21.20 |  |  | 1030.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 49.31 ms | 3.13 ms | 1.81 ms | 12.66 | 12.66 | 27305.8 KB | 20.28 |  |  | 1165.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | OfficeIMO.Excel | 2.30 ms | 0.03 ms | 0.02 ms | 1.00 | 1.00 | 1787.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | ClosedXML | 12.05 ms | 1.01 ms | 0.58 ms | 5.24 | 5.24 | 9959.5 KB | 5.57 |  |  | 423.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | EPPlus | 23.75 ms | 1.90 ms | 1.10 ms | 10.32 | 10.32 | 11772.9 KB | 6.59 |  |  | 932.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | OfficeIMO.Excel | 2.13 ms | 0.04 ms | 0.02 ms | 1.00 | 1.00 | 1119.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | ClosedXML | 12.57 ms | 2.27 ms | 1.31 ms | 5.89 | 5.89 | 9177.1 KB | 8.19 |  |  | 489.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | EPPlus | 21.55 ms | 0.44 ms | 0.25 ms | 10.10 | 10.10 | 12895.2 KB | 11.51 |  |  | 909.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | EPPlus 4.5.3.3 | 31.66 ms |  |  | 14.83 | 14.83 |  |  |  |  | 1383.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | OfficeIMO.Excel | 3.21 ms | 0.12 ms | 0.07 ms | 1.00 | 1.00 | 1763.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | ClosedXML | 16.73 ms | 0.38 ms | 0.22 ms | 5.21 | 5.21 | 11887.0 KB | 6.74 |  |  | 420.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | EPPlus | 25.04 ms | 0.33 ms | 0.19 ms | 7.79 | 7.79 | 15643.3 KB | 8.87 |  |  | 679.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | EPPlus 4.5.3.3 | 46.58 ms |  |  | 14.50 | 14.50 |  |  |  |  | 1349.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | OfficeIMO.Excel | 3.24 ms | 0.08 ms | 0.05 ms | 1.00 | 1.00 | 1506.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | ClosedXML | 15.17 ms | 0.10 ms | 0.06 ms | 4.69 | 4.69 | 11296.3 KB | 7.50 |  |  | 368.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | EPPlus | 24.37 ms | 0.41 ms | 0.24 ms | 7.53 | 7.53 | 14960.2 KB | 9.93 |  |  | 652.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 3.13 ms | 0.09 ms | 0.05 ms | 1.00 | 1.00 | 1506.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | ClosedXML | 15.18 ms | 1.07 ms | 0.62 ms | 4.84 | 4.84 | 11296.3 KB | 7.50 |  |  | 384.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | EPPlus | 25.76 ms | 1.25 ms | 0.72 ms | 8.22 | 8.22 | 14960.2 KB | 9.93 |  |  | 721.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | OfficeIMO.Excel | 2.25 ms | 0.35 ms | 0.20 ms | 1.00 | 1.00 | 1138.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | ClosedXML | 10.94 ms | 0.21 ms | 0.12 ms | 4.86 | 4.86 | 9021.2 KB | 7.93 |  |  | 386.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | EPPlus | 23.77 ms | 2.04 ms | 1.18 ms | 10.57 | 10.57 | 12827.4 KB | 11.27 |  |  | 956.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | EPPlus 4.5.3.3 | 24.86 ms |  |  | 11.05 | 11.05 |  |  |  |  | 1005.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | OfficeIMO.Excel | 3.93 ms | 0.59 ms | 0.34 ms | 1.00 | 1.00 | 1435.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | ClosedXML | 15.73 ms | 1.26 ms | 0.73 ms | 4.00 | 4.00 | 9711.1 KB | 6.76 |  |  | 299.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | EPPlus | 25.20 ms | 2.99 ms | 1.72 ms | 6.41 | 6.41 | 14722.6 KB | 10.26 |  |  | 540.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | EPPlus 4.5.3.3 | 34.88 ms |  |  | 8.87 | 8.87 |  |  |  |  | 786.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | OfficeIMO.Excel | 5.45 ms | 0.32 ms | 0.18 ms | 1.00 | 1.00 | 2064.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | MiniExcel | 15.49 ms | 1.31 ms | 0.76 ms | 2.84 | 2.84 | 29223.6 KB | 14.16 |  |  | 184.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | ClosedXML | 46.14 ms | 1.31 ms | 0.76 ms | 8.46 | 8.46 | 18913.3 KB | 9.16 |  |  | 745.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | EPPlus | 58.12 ms | 3.90 ms | 2.25 ms | 10.65 | 10.65 | 17701.2 KB | 8.57 |  |  | 965.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | OfficeIMO.Excel | 8.68 ms | 1.23 ms | 0.71 ms | 1.00 | 1.00 | 2880.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | MiniExcel | 16.11 ms | 3.68 ms | 2.13 ms | 1.86 | 1.86 | 29968.0 KB | 10.40 |  |  | 85.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | EPPlus | 63.07 ms | 6.99 ms | 4.03 ms | 7.27 | 7.27 | 21892.9 KB | 7.60 |  |  | 626.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | ClosedXML | 77.00 ms | 8.54 ms | 4.93 ms | 8.87 | 8.87 | 27410.7 KB | 9.52 |  |  | 787.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | OfficeIMO.Excel | 5.50 ms | 0.29 ms | 0.17 ms | 1.00 | 1.00 | 2067.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | MiniExcel | 15.88 ms | 0.88 ms | 0.51 ms | 2.89 | 2.89 | 28700.3 KB | 13.88 |  |  | 188.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | EPPlus 4.5.3.3 | 47.36 ms |  |  | 8.61 | 8.61 |  |  |  |  | 761.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | ClosedXML | 83.23 ms | 2.51 ms | 1.45 ms | 15.13 | 15.13 | 18878.2 KB | 9.13 |  |  | 1413.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | EPPlus | 85.71 ms | 4.21 ms | 2.43 ms | 15.58 | 15.58 | 19431.0 KB | 9.40 |  |  | 1458.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | OfficeIMO.Excel | 5.60 ms | 1.90 ms | 1.10 ms | 1.00 | 1.00 | 1774.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | MiniExcel | 8.96 ms | 0.55 ms | 0.32 ms | 1.60 | 1.60 | 23044.1 KB | 12.98 |  |  | 60.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | ClosedXML | 37.88 ms | 4.78 ms | 2.76 ms | 6.77 | 6.77 | 19008.4 KB | 10.71 |  |  | 577.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | EPPlus | 44.08 ms | 9.08 ms | 5.24 ms | 7.88 | 7.88 | 16646.2 KB | 9.38 |  |  | 687.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | EPPlus 4.5.3.3 | 52.50 ms |  |  | 9.38 | 9.38 |  |  |  |  | 838.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | OfficeIMO.Excel | 4.22 ms | 0.02 ms | 0.01 ms | 1.00 | 1.00 | 1748.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | LargeXlsx | 8.70 ms | 1.96 ms | 1.13 ms | 2.06 | 2.06 | 1149.0 KB | 0.66 |  |  | 106.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | MiniExcel | 9.25 ms | 1.18 ms | 0.68 ms | 2.19 | 2.19 | 23062.6 KB | 13.19 |  |  | 119.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | ClosedXML | 29.38 ms | 0.84 ms | 0.48 ms | 6.97 | 6.97 | 11581.0 KB | 6.62 |  |  | 596.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | EPPlus | 41.04 ms | 5.41 ms | 3.12 ms | 9.73 | 9.73 | 16647.6 KB | 9.52 |  |  | 872.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | EPPlus 4.5.3.3 | 42.05 ms |  |  | 9.97 | 9.97 |  |  |  |  | 896.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | OfficeIMO.Excel | 4.88 ms | 0.72 ms | 0.42 ms | 1.00 | 1.00 | 1487.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | MiniExcel | 8.37 ms | 0.79 ms | 0.46 ms | 1.71 | 1.71 | 22789.4 KB | 15.32 |  |  | 71.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | ClosedXML | 36.16 ms | 1.54 ms | 0.89 ms | 7.41 | 7.41 | 18735.1 KB | 12.60 |  |  | 640.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | EPPlus | 37.03 ms | 2.20 ms | 1.27 ms | 7.58 | 7.58 | 16373.5 KB | 11.01 |  |  | 658.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | OfficeIMO.Excel | 4.94 ms | 0.43 ms | 0.25 ms | 1.00 | 1.00 | 1760.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | MiniExcel | 9.11 ms | 1.14 ms | 0.66 ms | 1.84 | 1.84 | 23062.9 KB | 13.10 |  |  | 84.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | ClosedXML | 38.67 ms | 3.37 ms | 1.95 ms | 7.83 | 7.83 | 19008.7 KB | 10.80 |  |  | 683.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | EPPlus | 42.78 ms | 6.74 ms | 3.89 ms | 8.66 | 8.66 | 16647.4 KB | 9.46 |  |  | 766.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | EPPlus 4.5.3.3 | 44.89 ms |  |  | 9.09 | 9.09 |  |  |  |  | 809.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | OfficeIMO.Excel | 6.56 ms | 0.47 ms | 0.27 ms | 1.00 | 1.00 | 1403.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | MiniExcel | 15.10 ms | 1.03 ms | 0.59 ms | 2.30 | 2.30 | 26825.0 KB | 19.12 |  |  | 130.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | EPPlus | 129.84 ms | 8.58 ms | 4.95 ms | 19.80 | 19.80 | 49158.1 KB | 35.03 |  |  | 1880.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | EPPlus 4.5.3.3 | 133.11 ms |  |  | 20.30 | 20.30 |  |  |  |  | 1930.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | ClosedXML | 244.18 ms | 16.71 ms | 9.65 ms | 37.24 | 37.24 | 58350.2 KB | 41.58 |  |  | 3624.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | OfficeIMO.Excel | 3.05 ms | 0.03 ms | 0.01 ms | 1.00 | 1.00 | 1620.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | ClosedXML | 18.24 ms | 1.18 ms | 0.68 ms | 5.97 | 5.97 | 12039.8 KB | 7.43 |  |  | 497.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | EPPlus | 36.88 ms | 0.48 ms | 0.28 ms | 12.08 | 12.08 | 18110.5 KB | 11.17 |  |  | 1107.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | EPPlus 4.5.3.3 | 47.67 ms |  |  | 15.61 | 15.61 |  |  |  |  | 1461.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-dictionary-objects-table-direct | OfficeIMO.Excel | 5.62 ms | 0.34 ms | 0.19 ms | 1.00 | 1.00 | 2051.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | LargeXlsx | 4.74 ms | 0.57 ms | 0.33 ms | 0.66 | 1.00 | 802.5 KB | 0.34 |  |  | 33.8% faster than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 7.16 ms | 0.53 ms | 0.31 ms | 1.00 | 1.51 | 2341.7 KB | 1.00 |  |  | Loss +51.0% |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | MiniExcel | 10.11 ms | 2.43 ms | 1.40 ms | 1.41 | 2.13 | 25190.4 KB | 10.76 |  |  | 41.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | ClosedXML | 45.73 ms | 3.57 ms | 2.06 ms | 6.39 | 9.64 | 16973.5 KB | 7.25 |  |  | 538.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | EPPlus | 60.05 ms | 3.35 ms | 1.94 ms | 8.39 | 12.67 | 20105.1 KB | 8.59 |  |  | 738.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | LargeXlsx | 5.21 ms | 0.74 ms | 0.43 ms | 0.85 | 1.00 | 802.5 KB | 0.53 |  |  | 14.9% faster than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 6.13 ms | 0.81 ms | 0.47 ms | 1.00 | 1.18 | 1507.7 KB | 1.00 |  |  | Loss +17.6% |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | MiniExcel | 10.48 ms | 1.57 ms | 0.91 ms | 1.71 | 2.01 | 25190.4 KB | 16.71 |  |  | 71.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | ClosedXML | 44.07 ms | 2.73 ms | 1.58 ms | 7.19 | 8.46 | 16973.5 KB | 11.26 |  |  | 619.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | EPPlus | 56.45 ms | 3.83 ms | 2.21 ms | 9.21 | 10.83 | 20105.1 KB | 13.33 |  |  | 821.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 21.68 ms | 2.46 ms | 1.42 ms | 1.00 | 1.00 | 4502.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | LargeXlsx | 23.90 ms | 2.25 ms | 1.30 ms | 1.10 | 1.10 | 2810.7 KB | 0.62 |  |  | 10.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | MiniExcel | 37.56 ms | 1.60 ms | 0.92 ms | 1.73 | 1.73 | 48414.7 KB | 10.75 |  |  | 73.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | ClosedXML | 137.17 ms | 5.62 ms | 3.24 ms | 6.33 | 6.33 | 51647.0 KB | 11.47 |  |  | 532.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | EPPlus | 166.39 ms | 5.34 ms | 3.08 ms | 7.68 | 7.68 | 69139.6 KB | 15.36 |  |  | 667.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | LargeXlsx | 2.36 ms | 0.27 ms | 0.16 ms | 0.72 | 1.00 | 296.4 KB | 0.19 |  |  | 27.7% faster than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | OfficeIMO.Excel | 3.26 ms | 0.39 ms | 0.23 ms | 1.00 | 1.38 | 1576.3 KB | 1.00 |  |  | Loss +38.3% |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | MiniExcel | 7.03 ms | 0.71 ms | 0.41 ms | 2.15 | 2.98 | 19710.7 KB | 12.50 |  |  | 115.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | EPPlus 4.5.3.3 | 20.05 ms |  |  | 6.15 | 8.51 |  |  |  |  | 515.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | ClosedXML | 23.02 ms | 0.82 ms | 0.47 ms | 7.06 | 9.77 | 11197.4 KB | 7.10 |  |  | 606.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | EPPlus | 46.35 ms | 7.88 ms | 4.55 ms | 14.21 | 19.66 | 14365.2 KB | 9.11 |  |  | 1321.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 1.29 ms | 0.03 ms | 0.02 ms | 0.78 | 1.00 | 447.0 KB | 0.41 |  |  | 22.1% faster than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 1.65 ms | 0.20 ms | 0.11 ms | 1.00 | 1.28 | 1092.0 KB | 1.00 |  |  | Loss +28.4% |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | ClosedXML | 13.40 ms | 1.95 ms | 1.12 ms | 8.12 | 10.42 | 10235.8 KB | 9.37 |  |  | 711.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | EPPlus | 22.91 ms | 2.01 ms | 1.16 ms | 13.88 | 17.82 | 13052.1 KB | 11.95 |  |  | 1288.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | LargeXlsx | 3.35 ms | 0.14 ms | 0.08 ms | 0.74 | 1.00 | 758.3 KB | 0.36 |  |  | 25.5% faster than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 4.49 ms | 0.12 ms | 0.07 ms | 1.00 | 1.34 | 2081.1 KB | 1.00 |  |  | Loss +34.3% |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | MiniExcel | 9.13 ms | 0.53 ms | 0.31 ms | 2.03 | 2.73 | 23221.8 KB | 11.16 |  |  | 103.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | ClosedXML | 32.78 ms | 1.06 ms | 0.61 ms | 7.29 | 9.79 | 22221.3 KB | 10.68 |  |  | 629.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | EPPlus | 43.32 ms | 1.72 ms | 0.99 ms | 9.64 | 12.94 | 24693.7 KB | 11.87 |  |  | 863.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 89.15 ms |  |  | 19.84 | 26.63 |  |  |  |  | 1883.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | OfficeIMO.Excel | 2.39 ms | 0.08 ms | 0.05 ms | 1.00 | 1.00 | 1494.0 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | ClosedXML | 15.85 ms | 0.29 ms | 0.17 ms | 6.65 | 6.65 | 11296.3 KB | 7.56 |  |  | 564.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | EPPlus | 25.38 ms | 0.96 ms | 0.56 ms | 10.64 | 10.64 | 14960.0 KB | 10.01 |  |  | 963.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | Sylvan.Data.Excel | 3.69 ms | 0.53 ms | 0.30 ms | 0.82 | 1.00 | 758.6 KB | 0.43 |  |  | 18.0% faster than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | OfficeIMO.Excel | 4.51 ms | 0.29 ms | 0.17 ms | 1.00 | 1.22 | 1763.0 KB | 1.00 |  |  | Loss +22.0% |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | LargeXlsx | 5.41 ms | 2.38 ms | 1.37 ms | 1.20 | 1.46 | 1032.5 KB | 0.59 |  |  | 19.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | MiniExcel | 7.68 ms | 0.42 ms | 0.24 ms | 1.70 | 2.08 | 23043.9 KB | 13.07 |  |  | 70.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | ClosedXML | 28.25 ms | 1.36 ms | 0.78 ms | 6.27 | 7.65 | 11581.0 KB | 6.57 |  |  | 526.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | EPPlus | 37.61 ms | 2.33 ms | 1.34 ms | 8.35 | 10.18 | 16646.2 KB | 9.44 |  |  | 734.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | EPPlus 4.5.3.3 | 89.67 ms |  |  | 19.90 | 24.27 |  |  |  |  | 1889.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | OfficeIMO.Excel | 4.71 ms | 0.33 ms | 0.19 ms | 1.00 | 1.00 | 2140.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | LargeXlsx | 9.72 ms | 0.64 ms | 0.37 ms | 2.06 | 2.06 | 1123.9 KB | 0.53 |  |  | 106.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | MiniExcel | 12.64 ms | 0.96 ms | 0.55 ms | 2.68 | 2.68 | 29747.0 KB | 13.90 |  |  | 168.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | ClosedXML | 58.09 ms | 3.22 ms | 1.86 ms | 12.33 | 12.33 | 27410.6 KB | 12.80 |  |  | 1132.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | EPPlus | 62.40 ms | 0.74 ms | 0.42 ms | 13.24 | 13.24 | 21892.0 KB | 10.23 |  |  | 1224.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | LargeXlsx | 5.58 ms | 0.74 ms | 0.43 ms | 0.98 | 1.00 | 857.6 KB | 0.51 |  |  | Tie vs OfficeIMO |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | OfficeIMO.Excel | 5.69 ms | 1.68 ms | 0.97 ms | 1.00 | 1.02 | 1676.8 KB | 1.00 |  |  | Loss +2.0% |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | MiniExcel | 21.43 ms | 3.59 ms | 2.07 ms | 3.77 | 3.84 | 35917.7 KB | 21.42 |  |  | 276.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | ClosedXML | 111.34 ms | 12.29 ms | 7.10 ms | 19.57 | 19.97 | 71478.2 KB | 42.63 |  |  | 1856.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | OfficeIMO.Excel | 2.82 ms | 0.17 ms | 0.10 ms | 1.00 | 1.00 | 2440.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | MiniExcel | 5.09 ms | 1.10 ms | 0.63 ms | 1.80 | 1.80 | 21137.5 KB | 8.66 |  |  | 80.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | EPPlus 4.5.3.3 | 10.20 ms |  |  | 3.62 | 3.62 |  |  |  |  | 261.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | ClosedXML | 16.18 ms | 2.26 ms | 1.31 ms | 5.74 | 5.74 | 11299.2 KB | 4.63 |  |  | 474.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | EPPlus | 25.89 ms | 1.61 ms | 0.93 ms | 9.18 | 9.18 | 12804.4 KB | 5.25 |  |  | 818.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | OfficeIMO.Excel | 3.58 ms | 0.83 ms | 0.48 ms | 1.00 | 1.00 | 2617.0 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | ClosedXML | 17.65 ms | 1.35 ms | 0.78 ms | 4.93 | 4.93 | 11299.2 KB | 4.32 |  |  | 393.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | EPPlus | 22.66 ms | 0.88 ms | 0.51 ms | 6.33 | 6.33 | 12804.8 KB | 4.89 |  |  | 533.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | EPPlus 4.5.3.3 | 43.64 ms |  |  | 12.19 | 12.19 |  |  |  |  | 1119.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | OfficeIMO.Excel | 2.99 ms | 0.59 ms | 0.34 ms | 1.00 | 1.00 | 2379.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | ClosedXML | 18.34 ms | 4.19 ms | 2.42 ms | 6.14 | 6.14 | 13127.1 KB | 5.52 |  |  | 513.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | EPPlus | 29.41 ms | 2.93 ms | 1.69 ms | 9.84 | 9.84 | 13892.9 KB | 5.84 |  |  | 884.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | OfficeIMO.Excel | 2.59 ms | 0.48 ms | 0.28 ms | 1.00 | 1.00 | 1579.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | ClosedXML | 11.66 ms | 0.47 ms | 0.27 ms | 4.50 | 4.50 | 9226.5 KB | 5.84 |  |  | 349.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | EPPlus | 25.15 ms | 9.14 ms | 5.28 ms | 9.70 | 9.70 | 11332.4 KB | 7.17 |  |  | 870.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | LargeXlsx | 3.27 ms | 0.11 ms | 0.07 ms | 0.67 | 1.00 | 758.3 KB | 0.43 |  |  | 33.0% faster than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 4.88 ms | 1.74 ms | 1.01 ms | 1.00 | 1.49 | 1769.2 KB | 1.00 |  |  | Loss +49.2% |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | MiniExcel | 8.66 ms | 0.50 ms | 0.29 ms | 1.78 | 2.65 | 23222.2 KB | 13.13 |  |  | 77.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | ClosedXML | 32.30 ms | 3.58 ms | 2.06 ms | 6.63 | 9.89 | 11581.0 KB | 6.55 |  |  | 562.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | EPPlus | 41.39 ms | 4.98 ms | 2.88 ms | 8.49 | 12.67 | 16646.4 KB | 9.41 |  |  | 749.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 67.00 ms |  |  | 13.74 | 20.51 |  |  |  |  | 1274.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | LargeXlsx | 3.16 ms | 0.07 ms | 0.04 ms | 0.72 | 1.00 | 758.3 KB | 0.57 |  |  | 27.7% faster than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | OfficeIMO.Excel | 4.36 ms | 0.67 ms | 0.38 ms | 1.00 | 1.38 | 1339.3 KB | 1.00 |  |  | Loss +38.2% |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | MiniExcel | 9.87 ms | 1.47 ms | 0.85 ms | 2.26 | 3.13 | 23222.4 KB | 17.34 |  |  | 126.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | ClosedXML | 28.65 ms | 0.92 ms | 0.53 ms | 6.57 | 9.08 | 11581.0 KB | 8.65 |  |  | 556.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | EPPlus | 41.34 ms | 5.32 ms | 3.07 ms | 9.48 | 13.10 | 16646.1 KB | 12.43 |  |  | 847.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | EPPlus 4.5.3.3 | 72.83 ms |  |  | 16.70 | 23.08 |  |  |  |  | 1569.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 3.20 ms | 0.03 ms | 0.02 ms | 0.74 | 1.00 | 758.3 KB | 0.51 |  |  | 26.4% faster than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 4.35 ms | 0.21 ms | 0.12 ms | 1.00 | 1.36 | 1497.5 KB | 1.00 |  |  | Loss +35.9% |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | MiniExcel | 8.26 ms | 0.26 ms | 0.15 ms | 1.90 | 2.58 | 23222.3 KB | 15.51 |  |  | 89.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | ClosedXML | 28.31 ms | 0.24 ms | 0.14 ms | 6.51 | 8.84 | 11581.0 KB | 7.73 |  |  | 550.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | EPPlus | 41.45 ms | 5.41 ms | 3.12 ms | 9.53 | 12.95 | 16646.1 KB | 11.12 |  |  | 852.8% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | Sylvan.Data.Excel | 68.52 ms | 2.03 ms | 1.17 ms | 0.86 | 1.00 | 394.1 KB | 0.02 |  |  | 14.3% faster than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | OfficeIMO.Excel | 79.91 ms | 2.99 ms | 1.73 ms | 1.00 | 1.17 | 23622.2 KB | 1.00 |  |  | Loss +16.6% |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | ExcelDataReader | 179.67 ms | 10.91 ms | 6.30 ms | 2.25 | 2.62 | 69530.7 KB | 2.94 |  |  | 124.8% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | MiniExcel | 247.23 ms | 14.75 ms | 8.52 ms | 3.09 | 3.61 | 215349.0 KB | 9.12 |  |  | 209.4% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | OfficeIMO.Excel | 62.41 ms | 8.74 ms | 5.04 ms | 1.00 | 1.00 | 24404.3 KB | 1.00 |  |  | Win |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | Sylvan.Data.Excel | 63.74 ms | 3.20 ms | 1.85 ms | 1.02 | 1.02 | 394.1 KB | 0.02 |  |  | 2.1% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | ExcelDataReader | 196.90 ms | 66.50 ms | 38.39 ms | 3.16 | 3.16 | 69530.7 KB | 2.85 |  |  | 215.5% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | MiniExcel | 216.80 ms | 2.57 ms | 1.48 ms | 3.47 | 3.47 | 215349.0 KB | 8.82 |  |  | 247.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | append-plain-rows | LargeXlsx | 15.21 ms | 0.70 ms | 0.41 ms | 0.74 | 1.00 | 2771.0 KB | 0.26 | 605.0 KB | 0.99 | 25.8% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | append-plain-rows | OfficeIMO.Excel | 20.48 ms | 2.74 ms | 1.58 ms | 1.00 | 1.35 | 10842.5 KB | 1.00 | 610.4 KB | 1.00 | Loss +34.7% |
| 25000 | package-profile | package | Package size | append-plain-rows | MiniExcel | 41.72 ms | 3.20 ms | 1.85 ms | 2.04 | 2.74 | 58242.8 KB | 5.37 | 642.3 KB | 1.05 | 103.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | append-plain-rows | ClosedXML | 173.50 ms | 5.88 ms | 3.39 ms | 8.47 | 11.41 | 104233.1 KB | 9.61 | 540.6 KB | 0.89 | 747.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | append-plain-rows | EPPlus | 279.73 ms | 32.69 ms | 18.87 ms | 13.66 | 18.40 | 100373.5 KB | 9.26 | 525.6 KB | 0.86 | 1265.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | autofit-existing | OfficeIMO.Excel | 99.77 ms | 3.49 ms | 2.01 ms | 1.00 | 1.00 | 15708.3 KB | 1.00 | 1429.3 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | autofit-existing | EPPlus | 679.38 ms | 19.06 ms | 11.00 ms | 6.81 | 6.81 | 250948.5 KB | 15.98 | 1091.0 KB | 0.76 | 581.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | autofit-existing | ClosedXML | 2145.01 ms | 57.47 ms | 33.18 ms | 21.50 | 21.50 | 829855.9 KB | 52.83 | 1140.9 KB | 0.80 | 2050.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | large-shared-strings | OfficeIMO.Excel | 22.57 ms | 1.87 ms | 1.08 ms | 1.00 | 1.00 | 15744.9 KB | 1.00 | 529.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | large-shared-strings | MiniExcel | 57.58 ms | 18.43 ms | 10.64 ms | 2.55 | 2.55 | 73760.2 KB | 4.68 | 581.0 KB | 1.10 | 155.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | large-shared-strings | ClosedXML | 171.92 ms | 14.08 ms | 8.13 ms | 7.62 | 7.62 | 104241.3 KB | 6.62 | 460.1 KB | 0.87 | 661.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | large-shared-strings | EPPlus | 419.40 ms | 96.72 ms | 55.84 ms | 18.58 | 18.58 | 84410.0 KB | 5.36 | 444.7 KB | 0.84 | 1758.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-autofilter | OfficeIMO.Excel | 49.51 ms | 17.36 ms | 10.02 ms | 1.00 | 1.00 | 11494.9 KB | 1.00 | 1428.5 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-autofilter | ClosedXML | 428.62 ms | 26.99 ms | 15.58 ms | 8.66 | 8.66 | 210663.8 KB | 18.33 | 1140.0 KB | 0.80 | 765.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-autofilter | EPPlus | 547.71 ms | 39.92 ms | 23.05 ms | 11.06 | 11.06 | 211871.5 KB | 18.43 | 1090.1 KB | 0.76 | 1006.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-charts | OfficeIMO.Excel | 34.86 ms | 3.28 ms | 1.90 ms | 1.00 | 1.00 | 12551.0 KB | 1.00 | 1433.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-charts | EPPlus | 423.05 ms | 17.83 ms | 10.30 ms | 12.14 | 12.14 | 214905.3 KB | 17.12 | 1092.9 KB | 0.76 | 1113.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-conditional-formatting | OfficeIMO.Excel | 46.99 ms | 3.48 ms | 2.01 ms | 1.00 | 1.00 | 11560.2 KB | 1.00 | 1428.8 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-conditional-formatting | ClosedXML | 393.79 ms | 9.12 ms | 5.26 ms | 8.38 | 8.38 | 210711.7 KB | 18.23 | 1140.1 KB | 0.80 | 738.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-conditional-formatting | EPPlus | 546.43 ms | 6.29 ms | 3.63 ms | 11.63 | 11.63 | 211912.9 KB | 18.33 | 1090.2 KB | 0.76 | 1062.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-data-validation | OfficeIMO.Excel | 43.64 ms | 5.67 ms | 3.27 ms | 1.00 | 1.00 | 11510.5 KB | 1.00 | 1428.6 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-data-validation | ClosedXML | 379.88 ms | 15.90 ms | 9.18 ms | 8.70 | 8.70 | 210672.7 KB | 18.30 | 1140.1 KB | 0.80 | 770.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-data-validation | EPPlus | 510.66 ms | 19.03 ms | 10.99 ms | 11.70 | 11.70 | 211857.4 KB | 18.41 | 1090.1 KB | 0.76 | 1070.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-freeze-panes | OfficeIMO.Excel | 49.71 ms | 7.86 ms | 4.54 ms | 1.00 | 1.00 | 11497.3 KB | 1.00 | 1428.6 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-freeze-panes | ClosedXML | 426.98 ms | 63.51 ms | 36.67 ms | 8.59 | 8.59 | 210646.8 KB | 18.32 | 1140.0 KB | 0.80 | 758.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-freeze-panes | EPPlus | 642.83 ms | 94.17 ms | 54.37 ms | 12.93 | 12.93 | 211883.3 KB | 18.43 | 1090.2 KB | 0.76 | 1193.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-pivot-table | OfficeIMO.Excel | 385.23 ms | 28.01 ms | 16.17 ms | 1.00 | 1.00 | 131927.7 KB | 1.00 | 1979.8 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-pivot-table | EPPlus | 574.75 ms | 16.17 ms | 9.34 ms | 1.49 | 1.49 | 230800.9 KB | 1.75 | 1093.4 KB | 0.55 | 49.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-all-in-one | OfficeIMO.Excel | 279.05 ms | 31.79 ms | 18.36 ms | 1.00 | 1.00 | 133444.7 KB | 1.00 | 1985.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-all-in-one | EPPlus | 509.88 ms | 32.82 ms | 18.95 ms | 1.83 | 1.83 | 277077.5 KB | 2.08 | 1097.7 KB | 0.55 | 82.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-chart-first | OfficeIMO.Excel | 96.51 ms | 6.99 ms | 4.04 ms | 1.00 | 1.00 | 43564.8 KB | 1.00 | 1985.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-chart-first | EPPlus | 484.82 ms | 20.95 ms | 12.09 ms | 5.02 | 5.02 | 277076.3 KB | 6.36 | 1097.7 KB | 0.55 | 402.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-core | OfficeIMO.Excel | 54.69 ms | 7.66 ms | 4.42 ms | 1.00 | 1.00 | 11648.7 KB | 1.00 | 1430.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-core | EPPlus | 616.79 ms | 3.29 ms | 1.90 ms | 11.28 | 11.28 | 255065.8 KB | 21.90 | 1091.5 KB | 0.76 | 1027.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-core | ClosedXML | 1286.17 ms | 96.12 ms | 55.50 ms | 23.52 | 23.52 | 680117.1 KB | 58.39 | 1141.3 KB | 0.80 | 2251.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-extra-column | OfficeIMO.Excel | 477.94 ms | 45.12 ms | 26.05 ms | 1.00 | 1.00 | 144827.2 KB | 1.00 | 2110.8 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-extra-column | EPPlus | 684.69 ms | 38.86 ms | 22.43 ms | 1.43 | 1.43 | 302759.9 KB | 2.09 | 1166.3 KB | 0.55 | 43.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-no-autofit | OfficeIMO.Excel | 269.74 ms | 11.30 ms | 6.53 ms | 1.00 | 1.00 | 133432.6 KB | 1.00 | 1985.8 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-no-autofit | EPPlus | 455.50 ms | 12.21 ms | 7.05 ms | 1.69 | 1.69 | 234782.5 KB | 1.76 | 1097.7 KB | 0.55 | 68.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-post-mutation | OfficeIMO.Excel | 355.35 ms | 27.10 ms | 15.65 ms | 1.00 | 1.00 | 133461.9 KB | 1.00 | 1986.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-post-mutation | EPPlus | 597.79 ms | 11.80 ms | 6.81 ms | 1.68 | 1.68 | 277077.5 KB | 2.08 | 1097.8 KB | 0.55 | 68.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-shuffled-columns | OfficeIMO.Excel | 301.21 ms | 21.89 ms | 12.64 ms | 1.00 | 1.00 | 133506.1 KB | 1.00 | 2046.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-shuffled-columns | EPPlus | 523.34 ms | 18.78 ms | 10.84 ms | 1.74 | 1.74 | 277070.2 KB | 2.08 | 1098.4 KB | 0.54 | 73.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook | OfficeIMO.Excel | 618.76 ms | 49.47 ms | 28.56 ms | 1.00 | 1.00 | 175194.0 KB | 1.00 | 2672.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | report-workbook | EPPlus | 746.43 ms | 17.69 ms | 10.21 ms | 1.21 | 1.21 | 364709.1 KB | 2.08 | 1517.2 KB | 0.57 | 20.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook-core | OfficeIMO.Excel | 72.67 ms | 3.78 ms | 2.18 ms | 1.00 | 1.00 | 10979.4 KB | 1.00 | 1850.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | report-workbook-core | EPPlus | 755.83 ms | 42.60 ms | 24.59 ms | 10.40 | 10.40 | 342842.2 KB | 31.23 | 1512.6 KB | 0.82 | 940.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook-core | ClosedXML | 1697.17 ms | 137.44 ms | 79.35 ms | 23.36 | 23.36 | 975775.3 KB | 88.87 | 1579.8 KB | 0.85 | 2235.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook-datatable | OfficeIMO.Excel | 597.30 ms | 10.95 ms | 6.32 ms | 1.00 | 1.00 | 177940.0 KB | 1.00 | 2672.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | report-workbook-datatable | EPPlus | 696.60 ms | 27.28 ms | 15.75 ms | 1.17 | 1.17 | 247822.9 KB | 1.39 | 1517.2 KB | 0.57 | 16.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook-datatable-core | OfficeIMO.Excel | 66.63 ms | 4.38 ms | 2.53 ms | 1.00 | 1.00 | 13725.0 KB | 1.00 | 1850.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | report-workbook-datatable-core | EPPlus | 721.72 ms | 45.23 ms | 26.12 ms | 10.83 | 10.83 | 225955.8 KB | 16.46 | 1512.6 KB | 0.82 | 983.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook-datatable-core | ClosedXML | 1511.05 ms | 126.74 ms | 73.17 ms | 22.68 | 22.68 | 832231.7 KB | 60.64 | 1579.8 KB | 0.85 | 2167.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | LargeXlsx | 42.28 ms | 2.03 ms | 1.17 ms | 0.86 | 1.00 | 10795.2 KB | 0.92 | 2444.6 KB | 1.10 | 13.5% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | OfficeIMO.Excel | 48.91 ms | 8.86 ms | 5.11 ms | 1.00 | 1.16 | 11708.2 KB | 1.00 | 2228.8 KB | 1.00 | Loss +15.7% |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | MiniExcel | 157.40 ms | 5.73 ms | 3.31 ms | 3.22 | 3.72 | 226875.4 KB | 19.38 | 2410.6 KB | 1.08 | 221.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | ClosedXML | 923.70 ms | 55.70 ms | 32.16 ms | 18.89 | 21.85 | 759818.4 KB | 64.90 | 2581.2 KB | 1.16 | 1788.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-bulk-report | OfficeIMO.Excel | 47.67 ms | 3.39 ms | 1.96 ms | 1.00 | 1.00 | 11561.8 KB | 1.00 | 1429.3 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-bulk-report | MiniExcel | 96.96 ms | 7.32 ms | 4.22 ms | 2.03 | 2.03 | 125551.4 KB | 10.86 | 1521.1 KB | 1.06 | 103.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-bulk-report | EPPlus | 594.96 ms | 59.53 ms | 34.37 ms | 12.48 | 12.48 | 254959.0 KB | 22.05 | 1091.0 KB | 0.76 | 1148.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-bulk-report | ClosedXML | 1169.12 ms | 24.70 ms | 14.26 ms | 24.52 | 24.52 | 565955.0 KB | 48.95 | 1140.9 KB | 0.80 | 2352.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellformula | OfficeIMO.Excel | 29.76 ms | 1.52 ms | 0.88 ms | 1.00 | 1.00 | 10112.0 KB | 1.00 | 670.3 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellformula | ClosedXML | 265.63 ms | 35.61 ms | 20.56 ms | 8.92 | 8.92 | 113853.5 KB | 11.26 | 643.2 KB | 0.96 | 792.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellformula | EPPlus | 526.43 ms | 9.97 ms | 5.75 ms | 17.69 | 17.69 | 140731.9 KB | 13.92 | 593.9 KB | 0.89 | 1668.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-empty-strings | OfficeIMO.Excel | 16.61 ms | 1.21 ms | 0.70 ms | 1.00 | 1.00 | 6896.4 KB | 1.00 | 451.4 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-empty-strings | ClosedXML | 157.36 ms | 5.29 ms | 3.05 ms | 9.47 | 9.47 | 92902.1 KB | 13.47 | 398.1 KB | 0.88 | 847.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-empty-strings | EPPlus | 211.34 ms | 5.37 ms | 3.10 ms | 12.72 | 12.72 | 74492.8 KB | 10.80 | 390.6 KB | 0.87 | 1172.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-numbers | OfficeIMO.Excel | 19.93 ms | 2.30 ms | 1.33 ms | 1.00 | 1.00 | 5970.9 KB | 1.00 | 462.6 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-numbers | ClosedXML | 138.15 ms | 2.83 ms | 1.63 ms | 6.93 | 6.93 | 84206.7 KB | 14.10 | 411.4 KB | 0.89 | 593.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-numbers | EPPlus | 280.33 ms | 34.98 ms | 20.19 ms | 14.07 | 14.07 | 86377.5 KB | 14.47 | 406.5 KB | 0.88 | 1306.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-mixed | OfficeIMO.Excel | 25.30 ms | 0.20 ms | 0.11 ms | 1.00 | 1.00 | 8332.9 KB | 1.00 | 585.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-object-mixed | ClosedXML | 207.93 ms | 8.42 ms | 4.86 ms | 8.22 | 8.22 | 111118.7 KB | 13.33 | 532.9 KB | 0.91 | 721.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-mixed | EPPlus | 316.62 ms | 14.44 ms | 8.34 ms | 12.51 | 12.51 | 113245.1 KB | 13.59 | 544.3 KB | 0.93 | 1151.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse | OfficeIMO.Excel | 26.63 ms | 1.05 ms | 0.61 ms | 1.00 | 1.00 | 7416.2 KB | 1.00 | 607.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse | ClosedXML | 203.87 ms | 17.00 ms | 9.81 ms | 7.66 | 7.66 | 105223.9 KB | 14.19 | 468.0 KB | 0.77 | 665.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse | EPPlus | 295.96 ms | 42.68 ms | 24.64 ms | 11.11 | 11.11 | 106316.9 KB | 14.34 | 494.4 KB | 0.81 | 1011.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 23.82 ms | 0.39 ms | 0.23 ms | 1.00 | 1.00 | 7416.3 KB | 1.00 | 607.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | ClosedXML | 209.02 ms | 16.99 ms | 9.81 ms | 8.77 | 8.77 | 105223.9 KB | 14.19 | 468.0 KB | 0.77 | 777.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | EPPlus | 319.47 ms | 45.11 ms | 26.04 ms | 13.41 | 13.41 | 106316.9 KB | 14.34 | 494.4 KB | 0.81 | 1241.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-scalars | OfficeIMO.Excel | 15.47 ms | 1.53 ms | 0.88 ms | 1.00 | 1.00 | 6144.6 KB | 1.00 | 441.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-scalars | ClosedXML | 132.52 ms | 5.60 ms | 3.23 ms | 8.57 | 8.57 | 82591.3 KB | 13.44 | 394.9 KB | 0.89 | 756.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-scalars | EPPlus | 309.48 ms | 19.67 ms | 11.36 ms | 20.01 | 20.01 | 85127.4 KB | 13.85 | 379.3 KB | 0.86 | 1900.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings | OfficeIMO.Excel | 26.95 ms | 2.21 ms | 1.28 ms | 1.00 | 1.00 | 15360.4 KB | 1.00 | 527.8 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-strings | ClosedXML | 241.21 ms | 57.21 ms | 33.03 ms | 8.95 | 8.95 | 104241.3 KB | 6.79 | 460.1 KB | 0.87 | 795.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings | EPPlus | 385.67 ms | 40.78 ms | 23.55 ms | 14.31 | 14.31 | 84410.3 KB | 5.50 | 444.7 KB | 0.84 | 1331.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-distinct | OfficeIMO.Excel | 16.89 ms | 3.01 ms | 1.74 ms | 1.00 | 1.00 | 13824.1 KB | 1.00 | 499.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-distinct | ClosedXML | 201.59 ms | 6.23 ms | 3.60 ms | 11.94 | 11.94 | 131501.7 KB | 9.51 | 555.3 KB | 1.11 | 1093.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-distinct | EPPlus | 286.84 ms | 19.71 ms | 11.38 ms | 16.99 | 16.99 | 97729.6 KB | 7.07 | 565.1 KB | 1.13 | 1598.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-repeated | OfficeIMO.Excel | 17.02 ms | 1.42 ms | 0.82 ms | 1.00 | 1.00 | 7525.3 KB | 1.00 | 376.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-repeated | ClosedXML | 139.32 ms | 4.26 ms | 2.46 ms | 8.19 | 8.19 | 84520.0 KB | 11.23 | 331.8 KB | 0.88 | 718.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-repeated | EPPlus | 241.25 ms | 47.19 ms | 27.25 ms | 14.18 | 14.18 | 70033.4 KB | 9.31 | 300.8 KB | 0.80 | 1317.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-temporal | OfficeIMO.Excel | 28.45 ms | 2.53 ms | 1.46 ms | 1.00 | 1.00 | 7482.8 KB | 1.00 | 620.5 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-temporal | ClosedXML | 207.47 ms | 11.02 ms | 6.36 ms | 7.29 | 7.29 | 89323.7 KB | 11.94 | 483.0 KB | 0.78 | 629.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-temporal | EPPlus | 289.33 ms | 38.08 ms | 21.98 ms | 10.17 | 10.17 | 103800.0 KB | 13.87 | 495.1 KB | 0.80 | 917.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 10.67 ms | 0.47 ms | 0.27 ms | 0.80 | 1.00 | 3444.4 KB | 0.49 | 443.4 KB | 0.97 | 19.6% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 13.28 ms | 2.52 ms | 1.46 ms | 1.00 | 1.24 | 6961.7 KB | 1.00 | 455.5 KB | 1.00 | Loss +24.4% |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | ClosedXML | 193.91 ms | 21.01 ms | 12.13 ms | 14.60 | 18.17 | 96015.7 KB | 13.79 | 467.5 KB | 1.03 | 1359.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | EPPlus | 284.85 ms | 64.82 ms | 37.42 ms | 21.44 | 26.68 | 87466.9 KB | 12.56 | 484.1 KB | 1.06 | 2044.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 44.56 ms | 7.73 ms | 4.47 ms | 1.00 | 1.00 | 16036.5 KB | 1.00 | 1384.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | LargeXlsx | 55.26 ms | 24.84 ms | 14.34 ms | 1.24 | 1.24 | 5614.1 KB | 0.35 | 1386.5 KB | 1.00 | 24.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | MiniExcel | 120.41 ms | 25.67 ms | 14.82 ms | 2.70 | 2.70 | 93257.0 KB | 5.82 | 1521.1 KB | 1.10 | 170.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | ClosedXML | 451.48 ms | 130.40 ms | 75.29 ms | 10.13 | 10.13 | 210646.1 KB | 13.14 | 1139.9 KB | 0.82 | 913.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | EPPlus | 655.73 ms | 163.37 ms | 94.32 ms | 14.72 | 14.72 | 211849.9 KB | 13.21 | 1090.0 KB | 0.79 | 1371.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | Sylvan.Data.Excel | 42.77 ms | 9.11 ms | 5.26 ms | 0.84 | 1.00 | 5700.3 KB | 0.44 | 755.4 KB | 0.55 | 15.7% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | LargeXlsx | 46.73 ms | 1.46 ms | 0.84 ms | 0.92 | 1.09 | 8349.2 KB | 0.64 | 1386.5 KB | 1.00 | 7.9% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | OfficeIMO.Excel | 50.74 ms | 1.38 ms | 0.80 ms | 1.00 | 1.19 | 13002.3 KB | 1.00 | 1384.9 KB | 1.00 | Loss +18.6% |
| 25000 | package-profile | package | Package size | write-datareader-plain | MiniExcel | 100.30 ms | 3.87 ms | 2.23 ms | 1.98 | 2.35 | 92199.7 KB | 7.09 | 1521.0 KB | 1.10 | 97.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | ClosedXML | 365.56 ms | 9.77 ms | 5.64 ms | 7.20 | 8.55 | 104205.0 KB | 8.01 | 1139.9 KB | 0.82 | 620.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | EPPlus | 441.71 ms | 14.68 ms | 8.48 ms | 8.71 | 10.33 | 117437.7 KB | 9.03 | 1090.8 KB | 0.79 | 770.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table | OfficeIMO.Excel | 48.21 ms | 4.14 ms | 2.39 ms | 1.00 | 1.00 | 13020.3 KB | 1.00 | 1385.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-datareader-table | MiniExcel | 106.36 ms | 4.57 ms | 2.64 ms | 2.21 | 2.21 | 92200.0 KB | 7.08 | 1521.0 KB | 1.10 | 120.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table | EPPlus | 439.80 ms | 11.06 ms | 6.38 ms | 9.12 | 9.12 | 117437.3 KB | 9.02 | 1090.8 KB | 0.79 | 812.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table | ClosedXML | 485.38 ms | 14.81 ms | 8.55 ms | 10.07 | 10.07 | 173397.5 KB | 13.32 | 1140.7 KB | 0.82 | 906.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | OfficeIMO.Excel | 55.01 ms | 5.22 ms | 3.01 ms | 1.00 | 1.00 | 13026.6 KB | 1.00 | 1385.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | MiniExcel | 101.47 ms | 4.37 ms | 2.52 ms | 1.84 | 1.84 | 124495.5 KB | 9.56 | 1521.1 KB | 1.10 | 84.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | EPPlus | 471.82 ms | 13.34 ms | 7.70 ms | 8.58 | 8.58 | 159741.8 KB | 12.26 | 1091.0 KB | 0.79 | 757.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | ClosedXML | 1037.93 ms | 53.34 ms | 30.80 ms | 18.87 | 18.87 | 566142.3 KB | 43.46 | 1140.9 KB | 0.82 | 1786.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | OfficeIMO.Excel | 41.29 ms | 3.36 ms | 1.94 ms | 1.00 | 1.00 | 9819.7 KB | 1.00 | 1329.2 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | LargeXlsx | 46.59 ms | 6.63 ms | 3.83 ms | 1.13 | 1.13 | 9265.9 KB | 0.94 | 1680.0 KB | 1.26 | 12.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | MiniExcel | 122.25 ms | 11.10 ms | 6.41 ms | 2.96 | 2.96 | 108129.1 KB | 11.01 | 1819.7 KB | 1.37 | 196.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | EPPlus | 585.11 ms | 39.96 ms | 23.07 ms | 14.17 | 14.17 | 135723.5 KB | 13.82 | 1390.4 KB | 1.05 | 1317.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | ClosedXML | 675.09 ms | 37.12 ms | 21.43 ms | 16.35 | 16.35 | 280372.9 KB | 28.55 | 1519.9 KB | 1.14 | 1534.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | OfficeIMO.Excel | 47.16 ms | 1.71 ms | 0.99 ms | 1.00 | 1.00 | 13458.5 KB | 1.00 | 1795.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | MiniExcel | 123.44 ms | 4.29 ms | 2.47 ms | 2.62 | 2.62 | 108129.1 KB | 8.03 | 1819.7 KB | 1.01 | 161.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | EPPlus | 618.63 ms | 34.79 ms | 20.09 ms | 13.12 | 13.12 | 135723.5 KB | 10.08 | 1390.4 KB | 0.77 | 1211.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | ClosedXML | 632.51 ms | 16.95 ms | 9.79 ms | 13.41 | 13.41 | 280371.8 KB | 20.83 | 1519.9 KB | 0.85 | 1241.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables | OfficeIMO.Excel | 48.11 ms | 3.26 ms | 1.88 ms | 1.00 | 1.00 | 10288.1 KB | 1.00 | 1376.4 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-dataset-tables | MiniExcel | 105.90 ms | 9.13 ms | 5.27 ms | 2.20 | 2.20 | 97085.6 KB | 9.44 | 1511.8 KB | 1.10 | 120.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables | EPPlus | 469.53 ms | 6.59 ms | 3.80 ms | 9.76 | 9.76 | 110815.9 KB | 10.77 | 1100.6 KB | 0.80 | 875.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables | ClosedXML | 492.47 ms | 26.23 ms | 15.14 ms | 10.24 | 10.24 | 172003.7 KB | 16.72 | 1139.0 KB | 0.83 | 923.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | OfficeIMO.Excel | 48.13 ms | 4.15 ms | 2.40 ms | 1.00 | 1.00 | 10300.7 KB | 1.00 | 1376.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | MiniExcel | 107.20 ms | 7.16 ms | 4.13 ms | 2.23 | 2.23 | 128874.9 KB | 12.51 | 1512.0 KB | 1.10 | 122.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | EPPlus | 498.11 ms | 25.66 ms | 14.82 ms | 10.35 | 10.35 | 195407.9 KB | 18.97 | 1100.9 KB | 0.80 | 935.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | ClosedXML | 962.62 ms | 7.24 ms | 4.18 ms | 20.00 | 20.00 | 550095.1 KB | 53.40 | 1139.3 KB | 0.83 | 1900.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-direct | LargeXlsx | 47.80 ms | 1.52 ms | 0.88 ms | 0.80 | 1.00 | 9520.4 KB | 0.75 | 1386.5 KB | 1.00 | 19.7% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-direct | OfficeIMO.Excel | 59.54 ms | 7.90 ms | 4.56 ms | 1.00 | 1.25 | 12715.7 KB | 1.00 | 1384.9 KB | 1.00 | Loss +24.6% |
| 25000 | package-profile | package | Package size | write-datatable-direct | MiniExcel | 115.01 ms | 5.72 ms | 3.30 ms | 1.93 | 2.41 | 92394.2 KB | 7.27 | 1521.1 KB | 1.10 | 93.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-direct | ClosedXML | 405.60 ms | 47.05 ms | 27.16 ms | 6.81 | 8.49 | 104205.0 KB | 8.19 | 1139.9 KB | 0.82 | 581.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-direct | EPPlus | 537.54 ms | 12.93 ms | 7.46 ms | 9.03 | 11.25 | 117437.3 KB | 9.24 | 1090.8 KB | 0.79 | 802.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | OfficeIMO.Excel | 42.27 ms | 0.22 ms | 0.12 ms | 1.00 | 1.00 | 12733.8 KB | 1.00 | 1385.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | MiniExcel | 99.61 ms | 4.30 ms | 2.48 ms | 2.36 | 2.36 | 92394.5 KB | 7.26 | 1521.0 KB | 1.10 | 135.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | EPPlus | 435.10 ms | 21.89 ms | 12.64 ms | 10.29 | 10.29 | 117437.3 KB | 9.22 | 1090.8 KB | 0.79 | 929.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | ClosedXML | 465.71 ms | 19.30 ms | 11.14 ms | 11.02 | 11.02 | 173402.7 KB | 13.62 | 1140.7 KB | 0.82 | 1001.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | LargeXlsx | 37.82 ms | 1.30 ms | 0.75 ms | 0.81 | 1.00 | 5614.1 KB | 0.43 | 1386.5 KB | 1.00 | 19.2% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 46.79 ms | 3.18 ms | 1.83 ms | 1.00 | 1.24 | 12912.0 KB | 1.00 | 1384.9 KB | 1.00 | Loss +23.7% |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | MiniExcel | 93.77 ms | 8.34 ms | 4.82 ms | 2.00 | 2.48 | 93257.0 KB | 7.22 | 1521.1 KB | 1.10 | 100.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | ClosedXML | 366.98 ms | 2.94 ms | 1.69 ms | 7.84 | 9.70 | 104205.0 KB | 8.07 | 1139.9 KB | 0.82 | 684.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | EPPlus | 492.49 ms | 27.39 ms | 15.81 ms | 10.53 | 13.02 | 117437.7 KB | 9.10 | 1090.8 KB | 0.79 | 952.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 66.78 ms | 19.86 ms | 11.47 ms | 1.00 | 1.00 | 11501.6 KB | 1.00 | 1428.5 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 745.50 ms | 145.75 ms | 84.15 ms | 11.16 | 11.16 | 159742.3 KB | 13.89 | 1091.0 KB | 0.76 | 1016.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 1103.79 ms | 207.90 ms | 120.03 ms | 16.53 | 16.53 | 496956.9 KB | 43.21 | 1140.1 KB | 0.80 | 1552.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | LargeXlsx | 38.19 ms | 1.21 ms | 0.70 ms | 0.90 | 1.00 | 5614.1 KB | 0.49 | 1386.5 KB | 0.97 | 10.2% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | OfficeIMO.Excel | 42.53 ms | 5.23 ms | 3.02 ms | 1.00 | 1.11 | 11493.8 KB | 1.00 | 1428.4 KB | 1.00 | Loss +11.4% |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | MiniExcel | 91.58 ms | 4.93 ms | 2.85 ms | 2.15 | 2.40 | 93257.0 KB | 8.11 | 1521.1 KB | 1.06 | 115.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | ClosedXML | 352.64 ms | 14.66 ms | 8.46 ms | 8.29 | 9.23 | 104205.0 KB | 9.07 | 1139.9 KB | 0.80 | 729.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | EPPlus | 511.31 ms | 21.92 ms | 12.66 ms | 12.02 | 13.39 | 117437.3 KB | 10.22 | 1090.8 KB | 0.76 | 1102.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 60.51 ms | 3.45 ms | 1.99 ms | 1.00 | 1.00 | 10187.2 KB | 1.00 | 1385.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 533.66 ms | 5.52 ms | 3.19 ms | 8.82 | 8.82 | 159742.2 KB | 15.68 | 1091.0 KB | 0.79 | 781.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 942.44 ms | 23.17 ms | 13.38 ms | 15.57 | 15.57 | 496956.9 KB | 48.78 | 1140.1 KB | 0.82 | 1457.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 47.87 ms | 5.37 ms | 3.10 ms | 0.74 | 1.00 | 5614.1 KB | 0.55 | 1386.5 KB | 1.00 | 26.3% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 65.00 ms | 4.12 ms | 2.38 ms | 1.00 | 1.36 | 10179.4 KB | 1.00 | 1384.9 KB | 1.00 | Loss +35.8% |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | MiniExcel | 112.24 ms | 7.23 ms | 4.18 ms | 1.73 | 2.34 | 93257.0 KB | 9.16 | 1521.0 KB | 1.10 | 72.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | ClosedXML | 417.42 ms | 24.68 ms | 14.25 ms | 6.42 | 8.72 | 104205.0 KB | 10.24 | 1139.9 KB | 0.82 | 542.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | EPPlus | 623.67 ms | 11.31 ms | 6.53 ms | 9.60 | 13.03 | 117437.3 KB | 11.54 | 1090.8 KB | 0.79 | 859.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | LargeXlsx | 39.99 ms | 2.35 ms | 1.36 ms | 0.65 | 1.00 | 5614.1 KB | 0.36 | 1386.5 KB | 0.97 | 35.5% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | OfficeIMO.Excel | 62.00 ms | 2.97 ms | 1.71 ms | 1.00 | 1.55 | 15791.7 KB | 1.00 | 1428.4 KB | 1.00 | Loss +55.0% |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | MiniExcel | 92.30 ms | 3.95 ms | 2.28 ms | 1.49 | 2.31 | 93257.0 KB | 5.91 | 1521.1 KB | 1.06 | 48.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | ClosedXML | 369.70 ms | 22.40 ms | 12.93 ms | 5.96 | 9.24 | 104205.0 KB | 6.60 | 1139.9 KB | 0.80 | 496.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | EPPlus | 507.55 ms | 50.69 ms | 29.27 ms | 8.19 | 12.69 | 117437.3 KB | 7.44 | 1090.8 KB | 0.76 | 718.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 48.05 ms | 1.77 ms | 1.02 ms | 1.00 | 1.00 | 11500.9 KB | 1.00 | 1428.5 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 525.77 ms | 16.95 ms | 9.79 ms | 10.94 | 10.94 | 138360.4 KB | 12.03 | 1091.0 KB | 0.76 | 994.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 599.75 ms | 32.87 ms | 18.98 ms | 12.48 | 12.48 | 275422.3 KB | 23.95 | 1140.1 KB | 0.80 | 1148.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | LargeXlsx | 62.30 ms | 8.24 ms | 4.76 ms | 0.89 | 1.00 | 6043.9 KB | 0.57 | 1816.3 KB | 0.99 | 10.8% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 69.88 ms | 5.69 ms | 3.29 ms | 1.00 | 1.12 | 10577.2 KB | 1.00 | 1828.0 KB | 1.00 | Loss +12.2% |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | MiniExcel | 153.24 ms | 38.74 ms | 22.37 ms | 2.19 | 2.46 | 113974.3 KB | 10.78 | 1936.7 KB | 1.06 | 119.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | ClosedXML | 591.88 ms | 90.09 ms | 52.01 ms | 8.47 | 9.50 | 179552.5 KB | 16.98 | 1555.2 KB | 0.85 | 747.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | EPPlus | 775.13 ms | 129.51 ms | 74.77 ms | 11.09 | 12.44 | 144920.0 KB | 13.70 | 1473.0 KB | 0.81 | 1009.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | LargeXlsx | 59.05 ms | 6.88 ms | 3.97 ms | 0.73 | 1.00 | 6043.9 KB | 0.61 | 1816.3 KB | 0.99 | 27.1% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 81.03 ms | 22.00 ms | 12.70 ms | 1.00 | 1.37 | 9942.2 KB | 1.00 | 1828.0 KB | 1.00 | Loss +37.2% |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | MiniExcel | 150.22 ms | 12.27 ms | 7.09 ms | 1.85 | 2.54 | 113974.3 KB | 11.46 | 1936.7 KB | 1.06 | 85.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | ClosedXML | 608.27 ms | 80.16 ms | 46.28 ms | 7.51 | 10.30 | 179552.5 KB | 18.06 | 1555.2 KB | 0.85 | 650.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | EPPlus | 826.02 ms | 91.98 ms | 53.10 ms | 10.19 | 13.99 | 144920.0 KB | 14.58 | 1473.0 KB | 0.81 | 919.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 275.51 ms | 22.13 ms | 12.78 ms | 1.00 | 1.00 | 36150.1 KB | 1.00 | 6725.6 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | LargeXlsx | 297.83 ms | 10.02 ms | 5.78 ms | 1.08 | 1.08 | 23211.4 KB | 0.64 | 6614.8 KB | 0.98 | 8.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | MiniExcel | 495.23 ms | 42.92 ms | 24.78 ms | 1.80 | 1.80 | 347925.7 KB | 9.62 | 6949.8 KB | 1.03 | 79.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | ClosedXML | 1725.77 ms | 237.02 ms | 136.84 ms | 6.26 | 6.26 | 487446.6 KB | 13.48 | 6165.9 KB | 0.92 | 526.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | EPPlus | 2219.46 ms | 111.45 ms | 64.34 ms | 8.06 | 8.06 | 562980.6 KB | 15.57 | 5441.6 KB | 0.81 | 705.6% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | OfficeIMO.Excel | 101.95 ms | 6.10 ms | 3.52 ms | 1.00 | 1.00 | 15708.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | EPPlus 4.5.3.3 | 613.54 ms |  |  | 6.02 | 6.02 |  |  |  |  | 501.8% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | EPPlus | 700.21 ms | 136.34 ms | 78.72 ms | 6.87 | 6.87 | 250948.5 KB | 15.98 |  |  | 586.9% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | ClosedXML | 1962.27 ms | 58.00 ms | 33.49 ms | 19.25 | 19.25 | 829858.8 KB | 52.83 |  |  | 1824.8% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus 4.5.3.3 | 252.39 ms |  |  | 0.93 | 1.00 |  |  |  |  | 6.9% faster than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | OfficeIMO.Excel | 271.08 ms | 9.10 ms | 5.25 ms | 1.00 | 1.07 | 133435.8 KB | 1.00 |  |  | Loss +7.4% |
| 25000 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus | 435.15 ms | 9.76 ms | 5.64 ms | 1.61 | 1.72 | 234782.5 KB | 1.76 |  |  | 60.5% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | OfficeIMO.Excel | 94.45 ms | 7.48 ms | 4.32 ms | 1.00 | 1.00 | 43566.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus | 465.71 ms | 17.19 ms | 9.92 ms | 4.93 | 4.93 | 277076.3 KB | 6.36 |  |  | 393.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus 4.5.3.3 | 512.61 ms |  |  | 5.43 | 5.43 |  |  |  |  | 442.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | OfficeIMO.Excel | 408.57 ms | 70.30 ms | 40.59 ms | 1.00 | 1.00 | 144822.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus | 552.72 ms | 19.41 ms | 11.20 ms | 1.35 | 1.35 | 302759.8 KB | 2.09 |  |  | 35.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus 4.5.3.3 | 862.35 ms |  |  | 2.11 | 2.11 |  |  |  |  | 111.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | OfficeIMO.Excel | 289.75 ms | 19.62 ms | 11.33 ms | 1.00 | 1.00 | 133462.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus | 478.18 ms | 10.31 ms | 5.95 ms | 1.65 | 1.65 | 277077.5 KB | 2.08 |  |  | 65.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus 4.5.3.3 | 766.56 ms |  |  | 2.65 | 2.65 |  |  |  |  | 164.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | OfficeIMO.Excel | 344.54 ms | 68.50 ms | 39.55 ms | 1.00 | 1.00 | 133505.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus | 559.22 ms | 36.89 ms | 21.30 ms | 1.62 | 1.62 | 277070.2 KB | 2.08 |  |  | 62.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus 4.5.3.3 | 797.85 ms |  |  | 2.32 | 2.32 |  |  |  |  | 131.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Object projection | build-object-datatable-dictionaries | OfficeIMO.Excel | 9.68 ms | 1.46 ms | 0.85 ms | 1.00 | 1.00 | 5164.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Object projection | build-object-datatable-typed | OfficeIMO.Excel | 8.13 ms | 1.67 ms | 0.96 ms | 1.00 | 1.00 | 8093.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-cells | OfficeIMO.Excel | 56.77 ms | 1.67 ms | 0.97 ms | 1.00 | 1.00 | 24531.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-cells | EPPlus | 323.34 ms | 20.77 ms | 11.99 ms | 5.70 | 5.70 | 187393.2 KB | 7.64 |  |  | 469.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-cells | ClosedXML | 386.18 ms | 8.58 ms | 4.95 ms | 6.80 | 6.80 | 166521.1 KB | 6.79 |  |  | 580.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | OfficeIMO.Excel | 35.46 ms | 2.48 ms | 1.43 ms | 1.00 | 1.00 | 3839.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | EPPlus | 330.23 ms | 29.49 ms | 17.03 ms | 9.31 | 9.31 | 115541.6 KB | 30.09 |  |  | 831.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | ClosedXML | 345.81 ms | 16.85 ms | 9.73 ms | 9.75 | 9.75 | 150901.2 KB | 39.30 |  |  | 875.2% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-range | OfficeIMO.Excel | 51.77 ms | 6.86 ms | 3.96 ms | 1.00 | 1.00 | 24531.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-range | ClosedXML | 348.20 ms | 10.28 ms | 5.93 ms | 6.73 | 6.73 | 166525.7 KB | 6.79 |  |  | 572.5% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-range | EPPlus | 360.84 ms | 38.90 ms | 22.46 ms | 6.97 | 6.97 | 187393.2 KB | 7.64 |  |  | 596.9% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-top-range | OfficeIMO.Excel | 0.64 ms | 0.09 ms | 0.05 ms | 1.00 | 1.00 | 285.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-top-range | EPPlus | 274.16 ms | 14.67 ms | 8.47 ms | 431.66 | 431.66 | 105580.1 KB | 369.87 |  |  | 43066.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-top-range | ClosedXML | 323.07 ms | 11.98 ms | 6.92 ms | 508.67 | 508.67 | 149402.6 KB | 523.39 |  |  | 50766.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-autofilter | OfficeIMO.Excel | 32.26 ms | 1.90 ms | 1.09 ms | 1.00 | 1.00 | 11494.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-autofilter | ClosedXML | 294.16 ms | 5.27 ms | 3.04 ms | 9.12 | 9.12 | 210663.8 KB | 18.33 |  |  | 811.8% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-autofilter | EPPlus 4.5.3.3 | 367.46 ms |  |  | 11.39 | 11.39 |  |  |  |  | 1039.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-autofilter | EPPlus | 396.06 ms | 6.93 ms | 4.00 ms | 12.28 | 12.28 | 211871.5 KB | 18.43 |  |  | 1127.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-charts | OfficeIMO.Excel | 33.47 ms | 0.92 ms | 0.53 ms | 1.00 | 1.00 | 12552.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-charts | EPPlus | 414.56 ms | 21.47 ms | 12.40 ms | 12.39 | 12.39 | 214905.3 KB | 17.12 |  |  | 1138.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-charts | EPPlus 4.5.3.3 | 456.99 ms |  |  | 13.65 | 13.65 |  |  |  |  | 1265.4% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | OfficeIMO.Excel | 33.86 ms | 0.74 ms | 0.43 ms | 1.00 | 1.00 | 11560.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | ClosedXML | 311.94 ms | 8.34 ms | 4.82 ms | 9.21 | 9.21 | 210711.7 KB | 18.23 |  |  | 821.2% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | EPPlus | 383.82 ms | 1.33 ms | 0.77 ms | 11.33 | 11.33 | 211912.9 KB | 18.33 |  |  | 1033.5% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | EPPlus 4.5.3.3 | 399.47 ms |  |  | 11.80 | 11.80 |  |  |  |  | 1079.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-data-validation | OfficeIMO.Excel | 31.38 ms | 1.41 ms | 0.81 ms | 1.00 | 1.00 | 11510.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-data-validation | ClosedXML | 296.52 ms | 10.53 ms | 6.08 ms | 9.45 | 9.45 | 210672.7 KB | 18.30 |  |  | 845.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-data-validation | EPPlus | 395.18 ms | 9.70 ms | 5.60 ms | 12.59 | 12.59 | 211857.4 KB | 18.41 |  |  | 1159.5% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-data-validation | EPPlus 4.5.3.3 | 434.49 ms |  |  | 13.85 | 13.85 |  |  |  |  | 1284.8% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | OfficeIMO.Excel | 31.51 ms | 0.85 ms | 0.49 ms | 1.00 | 1.00 | 11497.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | ClosedXML | 293.64 ms | 3.91 ms | 2.26 ms | 9.32 | 9.32 | 210646.8 KB | 18.32 |  |  | 832.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | EPPlus 4.5.3.3 | 370.53 ms |  |  | 11.76 | 11.76 |  |  |  |  | 1076.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | EPPlus | 394.58 ms | 10.77 ms | 6.22 ms | 12.52 | 12.52 | 211883.3 KB | 18.43 |  |  | 1152.4% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | OfficeIMO.Excel | 258.35 ms | 13.17 ms | 7.60 ms | 1.00 | 1.00 | 131928.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | EPPlus 4.5.3.3 | 351.05 ms |  |  | 1.36 | 1.36 |  |  |  |  | 35.9% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | EPPlus | 446.27 ms | 19.93 ms | 11.51 ms | 1.73 | 1.73 | 230800.9 KB | 1.75 |  |  | 72.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | OfficeIMO.Excel | 317.70 ms | 22.70 ms | 13.10 ms | 1.00 | 1.00 | 133447.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus | 474.92 ms | 36.81 ms | 21.25 ms | 1.49 | 1.49 | 277077.5 KB | 2.08 |  |  | 49.5% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus 4.5.3.3 | 511.23 ms |  |  | 1.61 | 1.61 |  |  |  |  | 60.9% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world report | realworld-report-core | OfficeIMO.Excel | 40.57 ms | 5.40 ms | 3.12 ms | 1.00 | 1.00 | 11648.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world report | realworld-report-core | EPPlus | 435.70 ms | 8.21 ms | 4.74 ms | 10.74 | 10.74 | 255065.8 KB | 21.90 |  |  | 974.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world report | realworld-report-core | EPPlus 4.5.3.3 | 799.72 ms |  |  | 19.71 | 19.71 |  |  |  |  | 1871.4% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world report | realworld-report-core | ClosedXML | 914.32 ms | 29.30 ms | 16.92 ms | 22.54 | 22.54 | 680116.8 KB | 58.39 |  |  | 2153.9% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook | OfficeIMO.Excel | 493.30 ms | 29.93 ms | 17.28 ms | 1.00 | 1.00 | 175197.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Report workbook | report-workbook | EPPlus | 722.33 ms | 99.36 ms | 57.37 ms | 1.46 | 1.46 | 364709.1 KB | 2.08 |  |  | 46.4% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook | EPPlus 4.5.3.3 | 1010.64 ms |  |  | 2.05 | 2.05 |  |  |  |  | 104.9% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-core | OfficeIMO.Excel | 52.78 ms | 4.13 ms | 2.38 ms | 1.00 | 1.00 | 10979.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Report workbook | report-workbook-core | EPPlus | 571.33 ms | 21.27 ms | 12.28 ms | 10.82 | 10.82 | 342842.2 KB | 31.23 |  |  | 982.4% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-core | EPPlus 4.5.3.3 | 944.71 ms |  |  | 17.90 | 17.90 |  |  |  |  | 1689.8% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-core | ClosedXML | 1232.73 ms | 25.47 ms | 14.71 ms | 23.35 | 23.35 | 975776.3 KB | 88.87 |  |  | 2235.4% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable | OfficeIMO.Excel | 450.29 ms | 16.44 ms | 9.49 ms | 1.00 | 1.00 | 177942.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable | EPPlus | 581.27 ms | 16.27 ms | 9.39 ms | 1.29 | 1.29 | 247823.0 KB | 1.39 |  |  | 29.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable | EPPlus 4.5.3.3 | 856.43 ms |  |  | 1.90 | 1.90 |  |  |  |  | 90.2% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable-core | OfficeIMO.Excel | 53.43 ms | 2.98 ms | 1.72 ms | 1.00 | 1.00 | 13725.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable-core | EPPlus | 620.91 ms | 41.21 ms | 23.79 ms | 11.62 | 11.62 | 225955.8 KB | 16.46 |  |  | 1062.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable-core | EPPlus 4.5.3.3 | 677.08 ms |  |  | 12.67 | 12.67 |  |  |  |  | 1167.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable-core | ClosedXML | 1219.58 ms | 26.37 ms | 15.23 ms | 22.83 | 22.83 | 832229.7 KB | 60.64 |  |  | 2182.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | OfficeIMO.Excel | 19.74 ms | 2.57 ms | 1.48 ms | 1.00 | 1.00 | 6216.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | EPPlus 4.5.3.3 | 76.43 ms |  |  | 3.87 | 3.87 |  |  |  |  | 287.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | EPPlus | 169.52 ms | 24.05 ms | 13.89 ms | 8.59 | 8.59 | 70814.5 KB | 11.39 |  |  | 758.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | ClosedXML | 171.06 ms | 14.29 ms | 8.25 ms | 8.66 | 8.66 | 79515.8 KB | 12.79 |  |  | 766.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | Sylvan.Data.Excel | 1.03 ms | 0.03 ms | 0.02 ms | 0.93 | 1.00 | 316.6 KB | 1.78 |  |  | 7.4% faster than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | OfficeIMO.Excel | 1.11 ms | 0.38 ms | 0.22 ms | 1.00 | 1.08 | 177.4 KB | 1.00 |  |  | Loss +8.0% |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | ExcelDataReader | 1.66 ms | 0.16 ms | 0.09 ms | 1.49 | 1.61 | 4062.2 KB | 22.90 |  |  | 49.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | ClosedXML | 3.40 ms | 0.19 ms | 0.11 ms | 3.05 | 3.30 | 4393.0 KB | 24.76 |  |  | 205.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | MiniExcel | 13.28 ms | 1.53 ms | 0.88 ms | 11.92 | 12.87 | 46194.9 KB | 260.40 |  |  | 1091.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | EPPlus 4.5.3.3 | 18.84 ms |  |  | 16.91 | 18.26 |  |  |  |  | 1590.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | EPPlus | 95.85 ms | 2.15 ms | 1.24 ms | 86.02 | 92.90 | 43071.0 KB | 242.79 |  |  | 8502.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | OfficeIMO.Excel | 1.17 ms | 0.25 ms | 0.14 ms | 1.00 | 1.00 | 177.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | Sylvan.Data.Excel | 1.22 ms | 0.31 ms | 0.18 ms | 1.04 | 1.04 | 316.6 KB | 1.78 |  |  | 4.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | ExcelDataReader | 1.90 ms | 0.55 ms | 0.32 ms | 1.63 | 1.63 | 4062.2 KB | 22.89 |  |  | 62.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | ClosedXML | 7.32 ms | 4.92 ms | 2.84 ms | 6.26 | 6.26 | 4393.0 KB | 24.75 |  |  | 526.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | MiniExcel | 14.30 ms | 2.46 ms | 1.42 ms | 12.23 | 12.23 | 46194.9 KB | 260.30 |  |  | 1123.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | EPPlus 4.5.3.3 | 17.41 ms |  |  | 14.89 | 14.89 |  |  |  |  | 1389.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | EPPlus | 102.74 ms | 8.91 ms | 5.14 ms | 87.89 | 87.89 | 43071.0 KB | 242.70 |  |  | 8689.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | Sylvan.Data.Excel | 17.14 ms | 1.18 ms | 0.68 ms | 0.92 | 1.00 | 1936.7 KB | 0.21 |  |  | 8.0% faster than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | OfficeIMO.Excel | 18.62 ms | 0.29 ms | 0.16 ms | 1.00 | 1.09 | 9218.2 KB | 1.00 |  |  | Loss +8.6% |
| 25000 | speed-comparison | read | Other | shared-string-read | ExcelDataReader | 41.87 ms | 0.36 ms | 0.21 ms | 2.25 | 2.44 | 25020.8 KB | 2.71 |  |  | 124.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | MiniExcel | 50.58 ms | 2.95 ms | 1.70 ms | 2.72 | 2.95 | 74405.3 KB | 8.07 |  |  | 171.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | EPPlus 4.5.3.3 | 95.07 ms |  |  | 5.11 | 5.55 |  |  |  |  | 410.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | ClosedXML | 148.53 ms | 5.74 ms | 3.31 ms | 7.98 | 8.67 | 90414.7 KB | 9.81 |  |  | 697.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | EPPlus | 151.94 ms | 3.63 ms | 2.10 ms | 8.16 | 8.87 | 89346.0 KB | 9.69 |  |  | 716.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | OfficeIMO.Excel | 33.96 ms | 2.58 ms | 1.49 ms | 1.00 | 1.00 | 1122.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | Sylvan.Data.Excel | 39.51 ms | 1.73 ms | 1.00 ms | 1.16 | 1.16 | 3534.8 KB | 3.15 |  |  | 16.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | ExcelDataReader | 107.22 ms | 1.96 ms | 1.13 ms | 3.16 | 3.16 | 61201.9 KB | 54.52 |  |  | 215.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | MiniExcel | 123.57 ms | 10.82 ms | 6.25 ms | 3.64 | 3.64 | 186420.9 KB | 166.06 |  |  | 263.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | EPPlus | 232.30 ms | 34.54 ms | 19.94 ms | 6.84 | 6.84 | 105609.0 KB | 94.08 |  |  | 584.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | ClosedXML | 314.30 ms | 14.54 ms | 8.40 ms | 9.26 | 9.26 | 149387.3 KB | 133.07 |  |  | 825.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | Sylvan.Data.Excel | 64.73 ms | 8.81 ms | 5.09 ms | 0.93 | 1.00 | 18394.2 KB | 0.53 |  |  | 6.6% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | OfficeIMO.Excel | 69.30 ms | 5.80 ms | 3.35 ms | 1.00 | 1.07 | 34646.0 KB | 1.00 |  |  | Loss +7.1% |
| 25000 | speed-comparison | read | Range and table read | read-datatable | ExcelDataReader | 141.88 ms | 9.54 ms | 5.51 ms | 2.05 | 2.19 | 76061.4 KB | 2.20 |  |  | 104.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | MiniExcel | 159.44 ms | 9.62 ms | 5.55 ms | 2.30 | 2.46 | 181285.0 KB | 5.23 |  |  | 130.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | EPPlus 4.5.3.3 | 216.10 ms |  |  | 3.12 | 3.34 |  |  |  |  | 211.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | EPPlus | 301.85 ms | 24.80 ms | 14.32 ms | 4.36 | 4.66 | 202250.2 KB | 5.84 |  |  | 335.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | ClosedXML | 369.47 ms | 50.92 ms | 29.40 ms | 5.33 | 5.71 | 178450.7 KB | 5.15 |  |  | 433.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | OfficeIMO.Excel | 72.82 ms | 22.39 ms | 12.93 ms | 1.00 | 1.00 | 4034.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | Sylvan.Data.Excel | 122.76 ms | 71.80 ms | 41.45 ms | 1.69 | 1.69 | 4316.2 KB | 1.07 |  |  | 68.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | MiniExcel | 190.82 ms | 16.00 ms | 9.24 ms | 2.62 | 2.62 | 158612.9 KB | 39.31 |  |  | 162.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | ExcelDataReader | 216.25 ms | 9.14 ms | 5.28 ms | 2.97 | 2.97 | 61201.9 KB | 15.17 |  |  | 197.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | ClosedXML | 481.16 ms | 116.05 ms | 67.00 ms | 6.61 | 6.61 | 150903.1 KB | 37.40 |  |  | 560.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | EPPlus | 527.98 ms | 50.20 ms | 28.98 ms | 7.25 | 7.25 | 115541.6 KB | 28.64 |  |  | 625.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | Sylvan.Data.Excel | 84.14 ms | 16.85 ms | 9.73 ms | 0.96 | 1.00 | 3534.8 KB | 0.14 |  |  | 4.3% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | OfficeIMO.Excel | 87.94 ms | 22.24 ms | 12.84 ms | 1.00 | 1.05 | 26098.5 KB | 1.00 |  |  | Loss +4.5% |
| 25000 | speed-comparison | read | Range and table read | read-range | MiniExcel | 179.25 ms | 58.59 ms | 33.83 ms | 2.04 | 2.13 | 186421.5 KB | 7.14 |  |  | 103.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | ExcelDataReader | 183.81 ms | 33.41 ms | 19.29 ms | 2.09 | 2.18 | 61201.9 KB | 2.35 |  |  | 109.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | EPPlus 4.5.3.3 | 187.35 ms |  |  | 2.13 | 2.23 |  |  |  |  | 113.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | ClosedXML | 475.77 ms | 77.99 ms | 45.03 ms | 5.41 | 5.65 | 163591.8 KB | 6.27 |  |  | 441.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | EPPlus | 536.57 ms | 47.82 ms | 27.61 ms | 6.10 | 6.38 | 187390.9 KB | 7.18 |  |  | 510.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | Sylvan.Data.Excel | 60.33 ms | 16.19 ms | 9.35 ms | 0.96 | 1.00 | 4484.9 KB | 0.17 |  |  | 4.4% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | OfficeIMO.Excel | 63.12 ms | 9.88 ms | 5.70 ms | 1.00 | 1.05 | 26684.4 KB | 1.00 |  |  | Loss +4.6% |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | ExcelDataReader | 121.99 ms | 15.90 ms | 9.18 ms | 1.93 | 2.02 | 61201.9 KB | 2.29 |  |  | 93.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | MiniExcel | 136.11 ms | 7.66 ms | 4.42 ms | 2.16 | 2.26 | 186421.5 KB | 6.99 |  |  | 115.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | EPPlus | 280.17 ms | 10.54 ms | 6.09 ms | 4.44 | 4.64 | 187390.9 KB | 7.02 |  |  | 343.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | ClosedXML | 362.61 ms | 16.05 ms | 9.27 ms | 5.75 | 6.01 | 163586.2 KB | 6.13 |  |  | 474.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | Sylvan.Data.Excel | 0.52 ms | 0.12 ms | 0.07 ms | 0.84 | 1.00 | 348.5 KB | 1.18 |  |  | 15.6% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | OfficeIMO.Excel | 0.61 ms | 0.09 ms | 0.05 ms | 1.00 | 1.19 | 296.3 KB | 1.00 |  |  | Loss +18.5% |
| 25000 | speed-comparison | read | Range and table read | read-top-range | MiniExcel | 0.89 ms | 0.29 ms | 0.17 ms | 1.45 | 1.71 | 869.0 KB | 2.93 |  |  | 44.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | ExcelDataReader | 46.69 ms | 10.81 ms | 6.24 ms | 75.93 | 89.99 | 17115.3 KB | 57.77 |  |  | 7493.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | EPPlus 4.5.3.3 | 163.29 ms |  |  | 265.57 | 314.75 |  |  |  |  | 26456.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | EPPlus | 271.67 ms | 25.14 ms | 14.51 ms | 441.84 | 523.66 | 105577.7 KB | 356.34 |  |  | 44084.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | ClosedXML | 371.92 ms | 67.73 ms | 39.11 ms | 604.87 | 716.88 | 149390.8 KB | 504.22 |  |  | 60387.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | Sylvan.Data.Excel | 60.27 ms | 13.56 ms | 7.83 ms | 0.47 | 1.00 | 3534.8 KB | 0.10 |  |  | 52.6% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | OfficeIMO.Excel | 127.12 ms | 8.13 ms | 4.69 ms | 1.00 | 2.11 | 34152.1 KB | 1.00 |  |  | Loss +110.9% |
| 25000 | speed-comparison | read | Range and table read | read-used-range | ExcelDataReader | 143.90 ms | 30.55 ms | 17.64 ms | 1.13 | 2.39 | 61201.9 KB | 1.79 |  |  | 13.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | MiniExcel | 152.95 ms | 36.14 ms | 20.87 ms | 1.20 | 2.54 | 186421.5 KB | 5.46 |  |  | 20.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | EPPlus | 359.84 ms | 58.78 ms | 33.94 ms | 2.83 | 5.97 | 187390.9 KB | 5.49 |  |  | 183.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | ClosedXML | 374.17 ms | 46.25 ms | 26.70 ms | 2.94 | 6.21 | 163593.0 KB | 4.79 |  |  | 194.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | Sylvan.Data.Excel | 40.08 ms | 1.41 ms | 0.81 ms | 0.96 | 1.00 | 3534.8 KB | 3.14 |  |  | 3.8% faster than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | OfficeIMO.Excel | 41.65 ms | 4.85 ms | 2.80 ms | 1.00 | 1.04 | 1125.9 KB | 1.00 |  |  | Loss +3.9% |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | ExcelDataReader | 110.65 ms | 5.05 ms | 2.91 ms | 2.66 | 2.76 | 61201.9 KB | 54.36 |  |  | 165.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | MiniExcel | 130.41 ms | 11.25 ms | 6.49 ms | 3.13 | 3.25 | 186420.9 KB | 165.58 |  |  | 213.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | EPPlus | 293.56 ms | 32.74 ms | 18.90 ms | 7.05 | 7.32 | 105609.0 KB | 93.80 |  |  | 604.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | ClosedXML | 361.26 ms | 37.85 ms | 21.85 ms | 8.67 | 9.01 | 149395.0 KB | 132.69 |  |  | 767.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | Sylvan.Data.Excel | 47.95 ms | 4.59 ms | 2.65 ms | 0.89 | 1.00 | 3534.8 KB | 0.13 |  |  | 11.1% faster than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | OfficeIMO.Excel | 53.97 ms | 1.54 ms | 0.89 ms | 1.00 | 1.13 | 26884.0 KB | 1.00 |  |  | Loss +12.5% |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | ExcelDataReader | 125.44 ms | 7.55 ms | 4.36 ms | 2.32 | 2.62 | 61201.9 KB | 2.28 |  |  | 132.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | MiniExcel | 131.07 ms | 6.04 ms | 3.49 ms | 2.43 | 2.73 | 186421.5 KB | 6.93 |  |  | 142.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | EPPlus 4.5.3.3 | 219.92 ms |  |  | 4.07 | 4.59 |  |  |  |  | 307.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | EPPlus | 276.84 ms | 31.12 ms | 17.97 ms | 5.13 | 5.77 | 187390.9 KB | 6.97 |  |  | 412.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | ClosedXML | 361.79 ms | 10.82 ms | 6.25 ms | 6.70 | 7.54 | 163594.1 KB | 6.09 |  |  | 570.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | OfficeIMO.Excel | 0.75 ms | 0.02 ms | 0.01 ms | 1.00 | 1.00 | 302.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | MiniExcel | 1.06 ms | 0.13 ms | 0.08 ms | 1.41 | 1.41 | 869.0 KB | 2.87 |  |  | 40.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | Sylvan.Data.Excel | 1.15 ms | 0.78 ms | 0.45 ms | 1.52 | 1.52 | 348.5 KB | 1.15 |  |  | 52.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | ExcelDataReader | 56.83 ms | 16.25 ms | 9.38 ms | 75.53 | 75.53 | 17115.3 KB | 56.62 |  |  | 7453.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | EPPlus 4.5.3.3 | 161.00 ms |  |  | 213.98 | 213.98 |  |  |  |  | 21297.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | EPPlus | 301.87 ms | 59.48 ms | 34.34 ms | 401.20 | 401.20 | 105577.7 KB | 349.30 |  |  | 40019.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | ClosedXML | 465.83 ms | 96.95 ms | 55.97 ms | 619.10 | 619.10 | 149392.7 KB | 494.26 |  |  | 61810.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | Sylvan.Data.Excel | 0.44 ms | 0.04 ms | 0.02 ms | 0.77 | 1.00 | 348.5 KB | 1.16 |  |  | 23.4% faster than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | OfficeIMO.Excel | 0.57 ms | 0.05 ms | 0.03 ms | 1.00 | 1.30 | 300.3 KB | 1.00 |  |  | Loss +30.5% |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | MiniExcel | 0.87 ms | 0.28 ms | 0.16 ms | 1.51 | 1.97 | 869.0 KB | 2.89 |  |  | 51.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | ExcelDataReader | 45.96 ms | 7.46 ms | 4.30 ms | 80.06 | 104.46 | 17115.3 KB | 56.99 |  |  | 7905.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | EPPlus | 242.10 ms | 12.83 ms | 7.41 ms | 421.71 | 550.27 | 105577.7 KB | 351.52 |  |  | 42070.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | ClosedXML | 326.86 ms | 6.52 ms | 3.77 ms | 569.34 | 742.91 | 149389.1 KB | 497.39 |  |  | 56833.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | OfficeIMO.Excel | 98.10 ms | 10.92 ms | 6.30 ms | 1.00 | 1.00 | 23562.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Typed object read | read-objects | Sylvan.Data.Excel | 108.46 ms | 25.74 ms | 14.86 ms | 1.11 | 1.11 | 5805.0 KB | 0.25 |  |  | 10.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | EPPlus 4.5.3.3 | 172.90 ms |  |  | 1.76 | 1.76 |  |  |  |  | 76.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | ExcelDataReader | 275.48 ms | 31.42 ms | 18.14 ms | 2.81 | 2.81 | 63472.1 KB | 2.69 |  |  | 180.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | MiniExcel | 398.09 ms | 114.04 ms | 65.84 ms | 4.06 | 4.06 | 183656.5 KB | 7.79 |  |  | 305.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | EPPlus | 555.37 ms | 105.15 ms | 60.71 ms | 5.66 | 5.66 | 199608.2 KB | 8.47 |  |  | 466.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | ClosedXML | 680.52 ms | 93.13 ms | 53.77 ms | 6.94 | 6.94 | 165542.2 KB | 7.03 |  |  | 593.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | OfficeIMO.Excel | 50.40 ms | 2.58 ms | 1.49 ms | 1.00 | 1.00 | 23367.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | Sylvan.Data.Excel | 61.18 ms | 11.19 ms | 6.46 ms | 1.21 | 1.21 | 5292.6 KB | 0.23 |  |  | 21.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | ExcelDataReader | 145.73 ms | 42.67 ms | 24.64 ms | 2.89 | 2.89 | 62959.8 KB | 2.69 |  |  | 189.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | MiniExcel | 148.76 ms | 24.51 ms | 14.15 ms | 2.95 | 2.95 | 183144.2 KB | 7.84 |  |  | 195.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | EPPlus 4.5.3.3 | 185.17 ms |  |  | 3.67 | 3.67 |  |  |  |  | 267.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | EPPlus | 330.94 ms | 50.53 ms | 29.17 ms | 6.57 | 6.57 | 199412.8 KB | 8.53 |  |  | 556.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | ClosedXML | 461.81 ms | 101.35 ms | 58.51 ms | 9.16 | 9.16 | 165348.7 KB | 7.08 |  |  | 816.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | OfficeIMO.Excel | 56.22 ms | 6.16 ms | 3.56 ms | 1.00 | 1.00 | 13026.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | MiniExcel | 118.37 ms | 4.56 ms | 2.63 ms | 2.11 | 2.11 | 124495.5 KB | 9.56 |  |  | 110.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | EPPlus 4.5.3.3 | 489.87 ms |  |  | 8.71 | 8.71 |  |  |  |  | 771.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | EPPlus | 651.63 ms | 167.41 ms | 96.65 ms | 11.59 | 11.59 | 159741.8 KB | 12.26 |  |  | 1059.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | ClosedXML | 1508.68 ms | 50.57 ms | 29.20 ms | 26.84 | 26.84 | 566142.3 KB | 43.46 |  |  | 2583.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | OfficeIMO.Excel | 37.89 ms | 0.27 ms | 0.15 ms | 1.00 | 1.00 | 10300.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | MiniExcel | 77.22 ms | 2.08 ms | 1.20 ms | 2.04 | 2.04 | 128874.9 KB | 12.51 |  |  | 103.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | EPPlus | 374.24 ms | 6.46 ms | 3.73 ms | 9.88 | 9.88 | 195407.9 KB | 18.97 |  |  | 887.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | ClosedXML | 728.51 ms | 11.15 ms | 6.44 ms | 19.23 | 19.23 | 550095.6 KB | 53.40 |  |  | 1822.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 49.88 ms | 4.61 ms | 2.66 ms | 1.00 | 1.00 | 11501.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 517.50 ms | 22.32 ms | 12.88 ms | 10.38 | 10.38 | 159742.3 KB | 13.89 |  |  | 937.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 972.48 ms | 64.15 ms | 37.04 ms | 19.50 | 19.50 | 496956.9 KB | 43.21 |  |  | 1849.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 58.50 ms | 7.24 ms | 4.18 ms | 1.00 | 1.00 | 10187.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 491.66 ms | 20.07 ms | 11.59 ms | 8.40 | 8.40 | 159742.3 KB | 15.68 |  |  | 740.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 1019.37 ms | 58.55 ms | 33.80 ms | 17.43 | 17.43 | 496956.9 KB | 48.78 |  |  | 1642.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 45.87 ms | 3.80 ms | 2.19 ms | 1.00 | 1.00 | 11500.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 513.26 ms | 42.23 ms | 24.38 ms | 11.19 | 11.19 | 138360.4 KB | 12.03 |  |  | 1019.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 585.29 ms | 36.87 ms | 21.29 ms | 12.76 | 12.76 | 275422.3 KB | 23.95 |  |  | 1176.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | OfficeIMO.Excel | 17.09 ms | 2.05 ms | 1.18 ms | 1.00 | 1.00 | 6896.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | ClosedXML | 152.66 ms | 3.06 ms | 1.77 ms | 8.93 | 8.93 | 92902.1 KB | 13.47 |  |  | 793.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | EPPlus | 215.18 ms | 8.56 ms | 4.94 ms | 12.59 | 12.59 | 74492.8 KB | 10.80 |  |  | 1159.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | OfficeIMO.Excel | 20.11 ms | 3.16 ms | 1.82 ms | 1.00 | 1.00 | 5970.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | EPPlus 4.5.3.3 | 106.95 ms |  |  | 5.32 | 5.32 |  |  |  |  | 431.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | ClosedXML | 133.69 ms | 6.73 ms | 3.89 ms | 6.65 | 6.65 | 84206.7 KB | 14.10 |  |  | 564.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | EPPlus | 256.54 ms | 48.20 ms | 27.83 ms | 12.76 | 12.76 | 86377.5 KB | 14.47 |  |  | 1175.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | OfficeIMO.Excel | 21.80 ms | 1.29 ms | 0.74 ms | 1.00 | 1.00 | 8332.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | EPPlus 4.5.3.3 | 118.04 ms |  |  | 5.42 | 5.42 |  |  |  |  | 441.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | ClosedXML | 190.87 ms | 10.80 ms | 6.24 ms | 8.76 | 8.76 | 111118.7 KB | 13.33 |  |  | 775.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | EPPlus | 240.28 ms | 21.46 ms | 12.39 ms | 11.02 | 11.02 | 113245.1 KB | 13.59 |  |  | 1002.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | OfficeIMO.Excel | 23.70 ms | 2.32 ms | 1.34 ms | 1.00 | 1.00 | 7416.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | ClosedXML | 175.28 ms | 11.91 ms | 6.88 ms | 7.40 | 7.40 | 105223.9 KB | 14.19 |  |  | 639.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | EPPlus | 237.84 ms | 10.99 ms | 6.34 ms | 10.03 | 10.03 | 106316.9 KB | 14.34 |  |  | 903.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 21.95 ms | 1.26 ms | 0.73 ms | 1.00 | 1.00 | 7416.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | ClosedXML | 180.85 ms | 7.10 ms | 4.10 ms | 8.24 | 8.24 | 105223.9 KB | 14.19 |  |  | 724.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | EPPlus | 257.43 ms | 5.44 ms | 3.14 ms | 11.73 | 11.73 | 106316.9 KB | 14.34 |  |  | 1072.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | OfficeIMO.Excel | 15.30 ms | 0.26 ms | 0.15 ms | 1.00 | 1.00 | 6144.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | EPPlus 4.5.3.3 | 104.52 ms |  |  | 6.83 | 6.83 |  |  |  |  | 583.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | ClosedXML | 130.43 ms | 10.54 ms | 6.08 ms | 8.53 | 8.53 | 82591.3 KB | 13.44 |  |  | 752.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | EPPlus | 284.44 ms | 12.18 ms | 7.03 ms | 18.59 | 18.59 | 85127.4 KB | 13.85 |  |  | 1759.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | OfficeIMO.Excel | 23.99 ms | 0.26 ms | 0.15 ms | 1.00 | 1.00 | 7482.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | EPPlus 4.5.3.3 | 106.13 ms |  |  | 4.42 | 4.42 |  |  |  |  | 342.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | ClosedXML | 184.93 ms | 10.37 ms | 5.99 ms | 7.71 | 7.71 | 89323.7 KB | 11.94 |  |  | 670.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | EPPlus | 233.42 ms | 12.56 ms | 7.25 ms | 9.73 | 9.73 | 103800.0 KB | 13.87 |  |  | 873.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | OfficeIMO.Excel | 38.51 ms | 3.77 ms | 2.18 ms | 1.00 | 1.00 | 13039.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | MiniExcel | 97.64 ms | 12.63 ms | 7.29 ms | 2.54 | 2.54 | 97088.3 KB | 7.45 |  |  | 153.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | ClosedXML | 425.52 ms | 24.68 ms | 14.25 ms | 11.05 | 11.05 | 172019.1 KB | 13.19 |  |  | 1004.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | EPPlus | 513.80 ms | 35.99 ms | 20.78 ms | 13.34 | 13.34 | 111246.0 KB | 8.53 |  |  | 1234.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | OfficeIMO.Excel | 48.69 ms | 14.47 ms | 8.35 ms | 1.00 | 1.00 | 13458.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | MiniExcel | 104.23 ms | 4.58 ms | 2.64 ms | 2.14 | 2.14 | 108129.1 KB | 8.03 |  |  | 114.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | EPPlus | 509.93 ms | 5.64 ms | 3.26 ms | 10.47 | 10.47 | 135723.5 KB | 10.08 |  |  | 947.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | ClosedXML | 633.19 ms | 142.78 ms | 82.43 ms | 13.01 | 13.01 | 280371.8 KB | 20.83 |  |  | 1200.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | OfficeIMO.Excel | 34.40 ms | 1.43 ms | 0.83 ms | 1.00 | 1.00 | 10288.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | MiniExcel | 77.99 ms | 4.24 ms | 2.45 ms | 2.27 | 2.27 | 97085.4 KB | 9.44 |  |  | 126.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | EPPlus 4.5.3.3 | 227.84 ms |  |  | 6.62 | 6.62 |  |  |  |  | 562.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | EPPlus | 319.12 ms | 3.16 ms | 1.83 ms | 9.28 | 9.28 | 110815.9 KB | 10.77 |  |  | 827.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | ClosedXML | 357.54 ms | 1.82 ms | 1.05 ms | 10.39 | 10.39 | 171999.1 KB | 16.72 |  |  | 939.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | OfficeIMO.Excel | 53.21 ms | 8.57 ms | 4.95 ms | 1.00 | 1.00 | 13020.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | MiniExcel | 119.97 ms | 6.65 ms | 3.84 ms | 2.25 | 2.25 | 92200.0 KB | 7.08 |  |  | 125.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | EPPlus 4.5.3.3 | 229.34 ms |  |  | 4.31 | 4.31 |  |  |  |  | 331.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | ClosedXML | 464.92 ms | 105.68 ms | 61.01 ms | 8.74 | 8.74 | 173398.1 KB | 13.32 |  |  | 773.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | EPPlus | 494.11 ms | 120.82 ms | 69.76 ms | 9.29 | 9.29 | 117437.3 KB | 9.02 |  |  | 828.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | LargeXlsx | 30.42 ms | 0.86 ms | 0.50 ms | 0.92 | 1.00 | 9520.4 KB | 0.75 |  |  | 8.2% faster than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | OfficeIMO.Excel | 33.13 ms | 0.18 ms | 0.10 ms | 1.00 | 1.09 | 12715.7 KB | 1.00 |  |  | Loss +8.9% |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | MiniExcel | 78.08 ms | 6.67 ms | 3.85 ms | 2.36 | 2.57 | 92394.2 KB | 7.27 |  |  | 135.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | EPPlus 4.5.3.3 | 218.81 ms |  |  | 6.60 | 7.19 |  |  |  |  | 560.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | ClosedXML | 269.30 ms | 13.10 ms | 7.57 ms | 8.13 | 8.85 | 104205.0 KB | 8.19 |  |  | 712.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | EPPlus | 317.45 ms | 6.51 ms | 3.76 ms | 9.58 | 10.44 | 117437.3 KB | 9.24 |  |  | 858.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | OfficeIMO.Excel | 37.43 ms | 4.03 ms | 2.33 ms | 1.00 | 1.00 | 9999.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | MiniExcel | 78.90 ms | 0.95 ms | 0.55 ms | 2.11 | 2.11 | 89659.2 KB | 8.97 |  |  | 110.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | EPPlus | 319.82 ms | 8.01 ms | 4.62 ms | 8.54 | 8.54 | 114703.1 KB | 11.47 |  |  | 754.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | ClosedXML | 363.95 ms | 17.32 ms | 10.00 ms | 9.72 | 9.72 | 170666.2 KB | 17.07 |  |  | 872.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | OfficeIMO.Excel | 34.71 ms | 1.37 ms | 0.79 ms | 1.00 | 1.00 | 12733.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | MiniExcel | 79.05 ms | 4.35 ms | 2.51 ms | 2.28 | 2.28 | 92394.5 KB | 7.26 |  |  | 127.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | EPPlus 4.5.3.3 | 226.71 ms |  |  | 6.53 | 6.53 |  |  |  |  | 553.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | EPPlus | 325.33 ms | 7.81 ms | 4.51 ms | 9.37 | 9.37 | 117437.3 KB | 9.22 |  |  | 837.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | ClosedXML | 363.47 ms | 16.52 ms | 9.54 ms | 10.47 | 10.47 | 173395.0 KB | 13.62 |  |  | 947.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | OfficeIMO.Excel | 38.72 ms | 5.84 ms | 3.37 ms | 1.00 | 1.00 | 11561.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | MiniExcel | 70.28 ms | 2.87 ms | 1.66 ms | 1.82 | 1.82 | 125551.4 KB | 10.86 |  |  | 81.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | EPPlus | 399.97 ms | 12.50 ms | 7.21 ms | 10.33 | 10.33 | 254959.0 KB | 22.05 |  |  | 933.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | EPPlus 4.5.3.3 | 531.79 ms |  |  | 13.74 | 13.74 |  |  |  |  | 1273.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | ClosedXML | 789.63 ms | 15.33 ms | 8.85 ms | 20.40 | 20.40 | 565950.2 KB | 48.95 |  |  | 1939.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | OfficeIMO.Excel | 27.62 ms | 3.33 ms | 1.92 ms | 1.00 | 1.00 | 10112.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | EPPlus 4.5.3.3 | 141.51 ms |  |  | 5.12 | 5.12 |  |  |  |  | 412.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | ClosedXML | 232.97 ms | 10.47 ms | 6.04 ms | 8.44 | 8.44 | 113853.5 KB | 11.26 |  |  | 743.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | EPPlus | 385.41 ms | 8.24 ms | 4.76 ms | 13.96 | 13.96 | 140731.9 KB | 13.92 |  |  | 1295.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-dictionary-objects-table-direct | OfficeIMO.Excel | 50.71 ms | 3.14 ms | 1.81 ms | 1.00 | 1.00 | 15163.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | LargeXlsx | 55.05 ms | 4.39 ms | 2.53 ms | 0.80 | 1.00 | 6043.9 KB | 0.57 |  |  | 19.8% faster than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 68.65 ms | 6.15 ms | 3.55 ms | 1.00 | 1.25 | 10577.2 KB | 1.00 |  |  | Loss +24.7% |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | MiniExcel | 120.76 ms | 6.21 ms | 3.59 ms | 1.76 | 2.19 | 113977.1 KB | 10.78 |  |  | 75.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | ClosedXML | 531.28 ms | 29.05 ms | 16.77 ms | 7.74 | 9.65 | 179552.5 KB | 16.98 |  |  | 673.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | EPPlus | 602.15 ms | 6.00 ms | 3.46 ms | 8.77 | 10.94 | 144920.0 KB | 13.70 |  |  | 777.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | LargeXlsx | 53.86 ms | 1.46 ms | 0.84 ms | 0.89 | 1.00 | 6043.9 KB | 0.61 |  |  | 11.2% faster than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 60.67 ms | 2.97 ms | 1.71 ms | 1.00 | 1.13 | 9942.2 KB | 1.00 |  |  | Loss +12.7% |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | MiniExcel | 115.79 ms | 5.60 ms | 3.23 ms | 1.91 | 2.15 | 113974.3 KB | 11.46 |  |  | 90.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | ClosedXML | 509.78 ms | 11.80 ms | 6.81 ms | 8.40 | 9.47 | 179552.5 KB | 18.06 |  |  | 740.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | EPPlus | 566.85 ms | 25.94 ms | 14.98 ms | 9.34 | 10.52 | 144920.0 KB | 14.58 |  |  | 834.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 284.19 ms | 32.07 ms | 18.52 ms | 1.00 | 1.00 | 36150.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | LargeXlsx | 292.23 ms | 86.28 ms | 49.81 ms | 1.03 | 1.03 | 23211.4 KB | 0.64 |  |  | 2.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | MiniExcel | 447.67 ms | 55.59 ms | 32.09 ms | 1.58 | 1.58 | 347925.7 KB | 9.62 |  |  | 57.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | ClosedXML | 1525.06 ms | 226.90 ms | 131.00 ms | 5.37 | 5.37 | 487446.6 KB | 13.48 |  |  | 436.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | EPPlus | 1993.52 ms | 188.34 ms | 108.74 ms | 7.01 | 7.01 | 562959.6 KB | 15.57 |  |  | 601.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | LargeXlsx | 11.63 ms | 0.68 ms | 0.39 ms | 0.77 | 1.00 | 2771.0 KB | 0.26 |  |  | 23.4% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | OfficeIMO.Excel | 15.18 ms | 0.65 ms | 0.37 ms | 1.00 | 1.31 | 10842.5 KB | 1.00 |  |  | Loss +30.5% |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | MiniExcel | 37.69 ms | 5.45 ms | 3.15 ms | 2.48 | 3.24 | 58242.8 KB | 5.37 |  |  | 148.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | EPPlus 4.5.3.3 | 115.51 ms |  |  | 7.61 | 9.93 |  |  |  |  | 660.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | ClosedXML | 142.31 ms | 7.10 ms | 4.10 ms | 9.37 | 12.23 | 104233.1 KB | 9.61 |  |  | 837.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | EPPlus | 224.80 ms | 31.96 ms | 18.45 ms | 14.80 | 19.32 | 100373.5 KB | 9.26 |  |  | 1380.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 12.82 ms | 0.87 ms | 0.50 ms | 0.88 | 1.00 | 3444.4 KB | 0.49 |  |  | 12.1% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 14.58 ms | 2.23 ms | 1.29 ms | 1.00 | 1.14 | 6961.7 KB | 1.00 |  |  | Loss +13.8% |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | ClosedXML | 151.27 ms | 8.78 ms | 5.07 ms | 10.37 | 11.80 | 96015.7 KB | 13.79 |  |  | 937.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | EPPlus | 228.81 ms | 8.54 ms | 4.93 ms | 15.69 | 17.85 | 87467.1 KB | 12.56 |  |  | 1469.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | LargeXlsx | 44.35 ms | 8.57 ms | 4.95 ms | 0.90 | 1.00 | 5614.1 KB | 0.35 |  |  | 9.9% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 49.20 ms | 2.07 ms | 1.19 ms | 1.00 | 1.11 | 16036.5 KB | 1.00 |  |  | Loss +10.9% |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | MiniExcel | 91.57 ms | 7.04 ms | 4.06 ms | 1.86 | 2.06 | 93257.0 KB | 5.82 |  |  | 86.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 272.30 ms |  |  | 5.53 | 6.14 |  |  |  |  | 453.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | ClosedXML | 419.05 ms | 25.95 ms | 14.98 ms | 8.52 | 9.45 | 210646.1 KB | 13.14 |  |  | 751.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | EPPlus | 490.75 ms | 36.07 ms | 20.82 ms | 9.97 | 11.07 | 211849.9 KB | 13.21 |  |  | 897.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | OfficeIMO.Excel | 21.40 ms | 1.12 ms | 0.65 ms | 1.00 | 1.00 | 7866.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | ClosedXML | 183.67 ms | 9.26 ms | 5.35 ms | 8.58 | 8.58 | 105223.9 KB | 13.38 |  |  | 758.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | EPPlus | 251.47 ms | 15.25 ms | 8.80 ms | 11.75 | 11.75 | 106316.9 KB | 13.52 |  |  | 1075.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | Sylvan.Data.Excel | 40.92 ms | 4.07 ms | 2.35 ms | 0.81 | 1.00 | 5700.3 KB | 0.44 |  |  | 18.6% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | LargeXlsx | 46.96 ms | 2.85 ms | 1.65 ms | 0.93 | 1.15 | 8349.2 KB | 0.64 |  |  | 6.5% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | OfficeIMO.Excel | 50.25 ms | 2.78 ms | 1.60 ms | 1.00 | 1.23 | 13002.3 KB | 1.00 |  |  | Loss +22.8% |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | MiniExcel | 112.84 ms | 17.16 ms | 9.91 ms | 2.25 | 2.76 | 92199.7 KB | 7.09 |  |  | 124.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | EPPlus 4.5.3.3 | 229.11 ms |  |  | 4.56 | 5.60 |  |  |  |  | 356.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | ClosedXML | 412.77 ms | 46.25 ms | 26.70 ms | 8.21 | 10.09 | 104205.0 KB | 8.01 |  |  | 721.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | EPPlus | 528.04 ms | 94.98 ms | 54.84 ms | 10.51 | 12.90 | 117437.7 KB | 9.03 |  |  | 950.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | OfficeIMO.Excel | 32.65 ms | 0.68 ms | 0.39 ms | 1.00 | 1.00 | 9819.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | LargeXlsx | 39.32 ms | 4.12 ms | 2.38 ms | 1.20 | 1.20 | 9265.9 KB | 0.94 |  |  | 20.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | MiniExcel | 98.65 ms | 2.89 ms | 1.67 ms | 3.02 | 3.02 | 108129.1 KB | 11.01 |  |  | 202.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | EPPlus | 474.14 ms | 13.91 ms | 8.03 ms | 14.52 | 14.52 | 135723.5 KB | 13.82 |  |  | 1352.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | ClosedXML | 515.34 ms | 6.76 ms | 3.90 ms | 15.78 | 15.78 | 280371.6 KB | 28.55 |  |  | 1478.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | LargeXlsx | 53.45 ms | 4.30 ms | 2.48 ms | 0.92 | 1.00 | 10795.2 KB | 0.92 |  |  | 7.5% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | OfficeIMO.Excel | 57.79 ms | 3.59 ms | 2.07 ms | 1.00 | 1.08 | 11708.2 KB | 1.00 |  |  | Loss +8.1% |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | MiniExcel | 207.19 ms | 18.01 ms | 10.40 ms | 3.59 | 3.88 | 226875.2 KB | 19.38 |  |  | 258.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | ClosedXML | 1146.09 ms | 72.14 ms | 41.65 ms | 19.83 | 21.44 | 759818.4 KB | 64.90 |  |  | 1883.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | OfficeIMO.Excel | 14.81 ms | 0.25 ms | 0.15 ms | 1.00 | 1.00 | 15744.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | MiniExcel | 31.06 ms | 1.17 ms | 0.67 ms | 2.10 | 2.10 | 73760.2 KB | 4.68 |  |  | 109.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | EPPlus 4.5.3.3 | 97.28 ms |  |  | 6.57 | 6.57 |  |  |  |  | 556.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | ClosedXML | 115.98 ms | 4.82 ms | 2.78 ms | 7.83 | 7.83 | 104241.3 KB | 6.62 |  |  | 682.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | EPPlus | 233.73 ms | 2.40 ms | 1.39 ms | 15.78 | 15.78 | 84410.0 KB | 5.36 |  |  | 1477.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | OfficeIMO.Excel | 22.75 ms | 0.73 ms | 0.42 ms | 1.00 | 1.00 | 15360.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | EPPlus 4.5.3.3 | 96.25 ms |  |  | 4.23 | 4.23 |  |  |  |  | 323.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | ClosedXML | 144.35 ms | 5.06 ms | 2.92 ms | 6.35 | 6.35 | 104241.3 KB | 6.79 |  |  | 534.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | EPPlus | 221.35 ms | 12.48 ms | 7.21 ms | 9.73 | 9.73 | 84410.5 KB | 5.50 |  |  | 873.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | OfficeIMO.Excel | 18.28 ms | 0.88 ms | 0.51 ms | 1.00 | 1.00 | 13824.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | ClosedXML | 202.58 ms | 30.10 ms | 17.38 ms | 11.08 | 11.08 | 131501.7 KB | 9.51 |  |  | 1008.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | EPPlus | 222.71 ms | 21.35 ms | 12.32 ms | 12.18 | 12.18 | 97729.6 KB | 7.07 |  |  | 1118.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | OfficeIMO.Excel | 15.99 ms | 0.86 ms | 0.49 ms | 1.00 | 1.00 | 7525.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | ClosedXML | 133.66 ms | 6.76 ms | 3.90 ms | 8.36 | 8.36 | 84520.0 KB | 11.23 |  |  | 736.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | EPPlus | 187.59 ms | 4.91 ms | 2.84 ms | 11.73 | 11.73 | 70033.4 KB | 9.31 |  |  | 1073.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | LargeXlsx | 42.05 ms | 3.26 ms | 1.88 ms | 0.93 | 1.00 | 5614.1 KB | 0.43 |  |  | 7.4% faster than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 45.41 ms | 2.60 ms | 1.50 ms | 1.00 | 1.08 | 12912.0 KB | 1.00 |  |  | Loss +8.0% |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | MiniExcel | 107.83 ms | 19.11 ms | 11.03 ms | 2.37 | 2.56 | 93257.0 KB | 7.22 |  |  | 137.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 247.48 ms |  |  | 5.45 | 5.89 |  |  |  |  | 444.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | ClosedXML | 397.45 ms | 41.87 ms | 24.18 ms | 8.75 | 9.45 | 104205.0 KB | 8.07 |  |  | 775.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | EPPlus | 490.36 ms | 8.37 ms | 4.83 ms | 10.80 | 11.66 | 117437.7 KB | 9.10 |  |  | 979.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | LargeXlsx | 41.97 ms | 4.72 ms | 2.73 ms | 0.95 | 1.00 | 5614.1 KB | 0.49 |  |  | 5.2% faster than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | OfficeIMO.Excel | 44.27 ms | 4.09 ms | 2.36 ms | 1.00 | 1.05 | 11493.8 KB | 1.00 |  |  | Loss +5.5% |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | MiniExcel | 99.77 ms | 7.18 ms | 4.14 ms | 2.25 | 2.38 | 93257.0 KB | 8.11 |  |  | 125.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | EPPlus 4.5.3.3 | 236.34 ms |  |  | 5.34 | 5.63 |  |  |  |  | 433.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | ClosedXML | 356.55 ms | 12.51 ms | 7.22 ms | 8.05 | 8.50 | 104205.0 KB | 9.07 |  |  | 705.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | EPPlus | 455.43 ms | 21.33 ms | 12.32 ms | 10.29 | 10.85 | 117437.3 KB | 10.22 |  |  | 928.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 38.19 ms | 5.52 ms | 3.18 ms | 0.78 | 1.00 | 5614.1 KB | 0.55 |  |  | 22.4% faster than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 49.23 ms | 0.69 ms | 0.40 ms | 1.00 | 1.29 | 10179.4 KB | 1.00 |  |  | Loss +28.9% |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | MiniExcel | 81.82 ms | 4.89 ms | 2.82 ms | 1.66 | 2.14 | 93257.0 KB | 9.16 |  |  | 66.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | ClosedXML | 352.25 ms | 20.12 ms | 11.62 ms | 7.16 | 9.22 | 104205.0 KB | 10.24 |  |  | 615.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | EPPlus | 399.54 ms | 3.29 ms | 1.90 ms | 8.12 | 10.46 | 117437.3 KB | 11.54 |  |  | 711.6% slower than OfficeIMO |
