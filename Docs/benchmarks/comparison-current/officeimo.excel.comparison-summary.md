# OfficeIMO.Excel Comparison Summary

This is the suite-level decision table. Mean, standard deviation, standard error, ratios, and allocations come from the lightweight rotated local runner; they are meant for engineering direction. Use the BenchmarkDotNet benchmark classes when a publication-grade `Error` column is required.

## At a glance

| Row count | Artifact | Workload | Category | OfficeIMO wins | OfficeIMO losses | Biggest loss |
| ---: | --- | --- | --- | ---: | ---: | --- |
| 2500 | dense-helloworld-comparison | read | Other | 0 | 2 | dense-helloworld-read-range: Loss +20.4% vs Sylvan.Data.Excel |
| 2500 | package-profile | package | Package size | 41 | 13 | write-insertobjects-legacy-dictionaries-direct: Loss +53.3% vs LargeXlsx |
| 2500 | speed-comparison | mutate | AutoFit and mutation | 2 | 0 |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | 4 | 0 |  |
| 2500 | speed-comparison | other | Object projection | 2 | 0 |  |
| 2500 | speed-comparison | other | Range and table read | 4 | 0 |  |
| 2500 | speed-comparison | other | Real-world feature mix | 6 | 0 |  |
| 2500 | speed-comparison | other | Real-world report | 2 | 0 |  |
| 2500 | speed-comparison | other | Report workbook | 4 | 0 |  |
| 2500 | speed-comparison | read | Formula write/read | 1 | 0 |  |
| 2500 | speed-comparison | read | Other | 1 | 2 | shared-string-read: Loss +61.4% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | read | Range and table read | 3 | 4 | read-datatable: Loss +95.7% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | read | Streaming read | 1 | 3 | read-range-stream: Loss +142.7% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | read | Typed object read | 0 | 2 | read-objects: Loss +15.0% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | write | AutoFit and mutation | 5 | 0 |  |
| 2500 | speed-comparison | write | Cell writer | 7 | 0 |  |
| 2500 | speed-comparison | write | DataSet table export | 3 | 0 |  |
| 2500 | speed-comparison | write | DataTable table export | 4 | 0 |  |
| 2500 | speed-comparison | write | Formatted report write | 1 | 0 |  |
| 2500 | speed-comparison | write | Formula write/read | 1 | 0 |  |
| 2500 | speed-comparison | write | Other | 2 | 2 | write-powershell-mixed-objects-direct: Loss +30.3% vs LargeXlsx |
| 2500 | speed-comparison | write | Plain cell export | 1 | 3 | append-plain-rows: Loss +43.5% vs LargeXlsx |
| 2500 | speed-comparison | write | Plain streaming export | 1 | 1 | write-datareader-plain: Loss +14.0% vs Sylvan.Data.Excel |
| 2500 | speed-comparison | write | Plain string export | 0 | 1 | write-blog-2023-20-string-columns: Loss +6.9% vs LargeXlsx |
| 2500 | speed-comparison | write | Shared string write | 4 | 0 |  |
| 2500 | speed-comparison | write | Typed object export | 0 | 3 | write-insertobjects-direct: Loss +37.5% vs LargeXlsx |
| 25000 | dense-helloworld-comparison | read | Other | 0 | 2 | dense-helloworld-read-stream: Loss +28.7% vs Sylvan.Data.Excel |
| 25000 | package-profile | package | Package size | 42 | 12 | write-insertobjects-legacy-dictionaries-direct: Loss +51.8% vs LargeXlsx |
| 25000 | speed-comparison | mutate | AutoFit and mutation | 2 | 0 |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | 4 | 0 |  |
| 25000 | speed-comparison | other | Object projection | 2 | 0 |  |
| 25000 | speed-comparison | other | Range and table read | 4 | 0 |  |
| 25000 | speed-comparison | other | Real-world feature mix | 6 | 0 |  |
| 25000 | speed-comparison | other | Real-world report | 2 | 0 |  |
| 25000 | speed-comparison | other | Report workbook | 4 | 0 |  |
| 25000 | speed-comparison | read | Formula write/read | 1 | 0 |  |
| 25000 | speed-comparison | read | Other | 2 | 1 | shared-string-read: Loss +14.4% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | read | Range and table read | 2 | 5 | read-top-range: Loss +31.6% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | read | Streaming read | 1 | 3 | read-top-range-stream-small-chunks: Loss +33.3% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | read | Typed object read | 1 | 1 | read-objects: Loss +29.3% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | write | AutoFit and mutation | 5 | 0 |  |
| 25000 | speed-comparison | write | Cell writer | 7 | 0 |  |
| 25000 | speed-comparison | write | DataSet table export | 3 | 0 |  |
| 25000 | speed-comparison | write | DataTable table export | 3 | 1 | write-datatable-direct: Loss +3.1% vs LargeXlsx |
| 25000 | speed-comparison | write | Formatted report write | 1 | 0 |  |
| 25000 | speed-comparison | write | Formula write/read | 1 | 0 |  |
| 25000 | speed-comparison | write | Other | 2 | 2 | write-powershell-psobject-mixed-direct: Loss +22.9% vs LargeXlsx |
| 25000 | speed-comparison | write | Plain cell export | 1 | 3 | append-plain-rows: Loss +29.4% vs LargeXlsx |
| 25000 | speed-comparison | write | Plain streaming export | 1 | 1 | write-datareader-plain: Loss +24.1% vs Sylvan.Data.Excel |
| 25000 | speed-comparison | write | Plain string export | 0 | 1 | write-blog-2023-20-string-columns: Loss +12.0% vs LargeXlsx |
| 25000 | speed-comparison | write | Shared string write | 4 | 0 |  |
| 25000 | speed-comparison | write | Typed object export | 0 | 3 | write-insertobjects-flat-dictionaries-direct: Loss +42.6% vs LargeXlsx |

## OfficeIMO decision table

| Row count | Artifact | Workload | Category | Scenario | OfficeIMO mean | Best | OfficeIMO vs best | Alloc | Package |
| ---: | --- | --- | --- | --- | ---: | --- | ---: | ---: | ---: |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | 8.12 ms | Sylvan.Data.Excel | Loss +20.4% | 2410.9 KB |  |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | 7.38 ms | Sylvan.Data.Excel | Loss +18.2% | 2489.4 KB |  |
| 2500 | package-profile | package | Package size | append-plain-rows | 2.25 ms | LargeXlsx | Loss +25.3% | 1576.3 KB | 63.0 KB |
| 2500 | package-profile | package | Package size | autofit-existing | 8.84 ms | OfficeIMO.Excel | Win | 1895.3 KB | 143.2 KB |
| 2500 | package-profile | package | Package size | large-shared-strings | 2.36 ms | OfficeIMO.Excel | Win | 2440.3 KB | 55.2 KB |
| 2500 | package-profile | package | Package size | realworld-autofilter | 4.61 ms | OfficeIMO.Excel | Win | 1340.4 KB | 142.4 KB |
| 2500 | package-profile | package | Package size | realworld-charts | 6.37 ms | OfficeIMO.Excel | Win | 1891.6 KB | 147.6 KB |
| 2500 | package-profile | package | Package size | realworld-conditional-formatting | 4.19 ms | OfficeIMO.Excel | Win | 1405.8 KB | 142.7 KB |
| 2500 | package-profile | package | Package size | realworld-data-validation | 4.94 ms | OfficeIMO.Excel | Win | 1356.1 KB | 142.5 KB |
| 2500 | package-profile | package | Package size | realworld-freeze-panes | 4.03 ms | OfficeIMO.Excel | Win | 1342.8 KB | 142.5 KB |
| 2500 | package-profile | package | Package size | realworld-pivot-table | 10.88 ms | OfficeIMO.Excel | Win | 5495.0 KB | 200.4 KB |
| 2500 | package-profile | package | Package size | realworld-report-all-in-one | 14.06 ms | OfficeIMO.Excel | Win | 6197.0 KB | 206.5 KB |
| 2500 | package-profile | package | Package size | realworld-report-chart-first | 13.49 ms | OfficeIMO.Excel | Win | 6198.7 KB | 206.5 KB |
| 2500 | package-profile | package | Package size | realworld-report-core | 4.93 ms | OfficeIMO.Excel | Win | 1488.5 KB | 143.9 KB |
| 2500 | package-profile | package | Package size | realworld-report-extra-column | 14.28 ms | OfficeIMO.Excel | Win | 6389.4 KB | 219.1 KB |
| 2500 | package-profile | package | Package size | realworld-report-no-autofit | 11.95 ms | OfficeIMO.Excel | Win | 6190.2 KB | 206.4 KB |
| 2500 | package-profile | package | Package size | realworld-report-post-mutation | 12.69 ms | OfficeIMO.Excel | Win | 6208.3 KB | 206.6 KB |
| 2500 | package-profile | package | Package size | realworld-report-shuffled-columns | 15.33 ms | OfficeIMO.Excel | Win | 6201.5 KB | 211.2 KB |
| 2500 | package-profile | package | Package size | report-workbook | 17.65 ms | OfficeIMO.Excel | Win | 7277.0 KB | 275.6 KB |
| 2500 | package-profile | package | Package size | report-workbook-core | 7.96 ms | OfficeIMO.Excel | Win | 2711.1 KB | 187.5 KB |
| 2500 | package-profile | package | Package size | report-workbook-datatable | 17.85 ms | OfficeIMO.Excel | Win | 7548.5 KB | 275.6 KB |
| 2500 | package-profile | package | Package size | report-workbook-datatable-core | 7.45 ms | OfficeIMO.Excel | Win | 2982.7 KB | 187.5 KB |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | 5.36 ms | LargeXlsx | Loss +5.8% | 1676.8 KB | 216.7 KB |
| 2500 | package-profile | package | Package size | write-bulk-report | 4.33 ms | OfficeIMO.Excel | Win | 1401.7 KB | 143.2 KB |
| 2500 | package-profile | package | Package size | write-cellformula | 3.39 ms | OfficeIMO.Excel | Win | 1383.3 KB | 66.6 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-empty-strings | 2.37 ms | OfficeIMO.Excel | Win | 1787.1 KB | 44.2 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-numbers | 2.49 ms | OfficeIMO.Excel | Win | 1119.9 KB | 47.0 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-object-mixed | 3.22 ms | OfficeIMO.Excel | Win | 1763.3 KB | 61.2 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse | 3.17 ms | OfficeIMO.Excel | Win | 1506.9 KB | 62.0 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | 3.11 ms | OfficeIMO.Excel | Win | 1507.0 KB | 62.0 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-scalars | 2.13 ms | OfficeIMO.Excel | Win | 1138.1 KB | 46.3 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-strings | 3.87 ms | OfficeIMO.Excel | Win | 2617.0 KB | 55.1 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-distinct | 2.43 ms | OfficeIMO.Excel | Win | 2379.2 KB | 51.8 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-repeated | 2.14 ms | OfficeIMO.Excel | Win | 1579.8 KB | 40.0 KB |
| 2500 | package-profile | package | Package size | write-cellvalue-temporal | 2.80 ms | OfficeIMO.Excel | Win | 1435.7 KB | 63.3 KB |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | 1.71 ms | LargeXlsx | Loss +23.2% | 1092.0 KB | 48.2 KB |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | 4.64 ms | LargeXlsx | Loss +19.8% | 2081.1 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-datareader-plain | 4.46 ms | Sylvan.Data.Excel | Loss +21.0% | 1763.0 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-datareader-table | 4.83 ms | OfficeIMO.Excel | Win | 1774.9 KB | 138.8 KB |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | 5.27 ms | OfficeIMO.Excel | Win | 1781.2 KB | 138.9 KB |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | 4.53 ms | OfficeIMO.Excel | Win | 2140.6 KB | 131.1 KB |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | 5.66 ms | OfficeIMO.Excel | Win | 2880.2 KB | 176.0 KB |
| 2500 | package-profile | package | Package size | write-dataset-tables | 5.01 ms | OfficeIMO.Excel | Win | 2066.1 KB | 138.9 KB |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | 5.08 ms | OfficeIMO.Excel | Win | 2078.7 KB | 139.2 KB |
| 2500 | package-profile | package | Package size | write-datatable-direct | 4.89 ms | LargeXlsx | Loss +22.4% | 1748.6 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | 4.74 ms | OfficeIMO.Excel | Win | 1760.7 KB | 138.8 KB |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | 4.56 ms | LargeXlsx | Loss +36.4% | 1769.2 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | 5.36 ms | OfficeIMO.Excel | Win | 1347.1 KB | 142.4 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | 4.68 ms | LargeXlsx | Loss +16.6% | 1339.3 KB | 142.3 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | 6.49 ms | OfficeIMO.Excel | Win | 1505.3 KB | 138.1 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | 5.71 ms | LargeXlsx | Loss +45.0% | 1497.5 KB | 138.0 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | 6.06 ms | LargeXlsx | Loss +53.3% | 1770.1 KB | 142.3 KB |
| 2500 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | 4.93 ms | OfficeIMO.Excel | Win | 1346.4 KB | 142.4 KB |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | 7.74 ms | LargeXlsx | Loss +45.8% | 2341.7 KB | 183.1 KB |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | 5.26 ms | LargeXlsx | Loss +5.2% | 1507.7 KB | 182.4 KB |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | 25.72 ms | LargeXlsx | Loss +7.5% | 4502.3 KB | 651.0 KB |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | 9.68 ms | OfficeIMO.Excel | Win | 1895.3 KB |  |
| 2500 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | 13.55 ms | OfficeIMO.Excel | Win | 6190.5 KB |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | 13.43 ms | OfficeIMO.Excel | Win | 6197.4 KB |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | 14.14 ms | OfficeIMO.Excel | Win | 6390.9 KB |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | 14.68 ms | OfficeIMO.Excel | Win | 6208.1 KB |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | 15.82 ms | OfficeIMO.Excel | Win | 6204.6 KB |  |
| 2500 | speed-comparison | other | Object projection | build-object-datatable-dictionaries | 1.67 ms | OfficeIMO.Excel | Win | 564.2 KB |  |
| 2500 | speed-comparison | other | Object projection | build-object-datatable-typed | 1.60 ms | OfficeIMO.Excel | Win | 856.9 KB |  |
| 2500 | speed-comparison | other | Range and table read | enumerate-cells | 7.95 ms | OfficeIMO.Excel | Win | 2531.8 KB |  |
| 2500 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | 5.17 ms | OfficeIMO.Excel | Win | 523.5 KB |  |
| 2500 | speed-comparison | other | Range and table read | enumerate-range | 7.62 ms | OfficeIMO.Excel | Win | 2531.9 KB |  |
| 2500 | speed-comparison | other | Range and table read | enumerate-top-range | 0.77 ms | OfficeIMO.Excel | Win | 285.5 KB |  |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-autofilter | 4.26 ms | OfficeIMO.Excel | Win | 1340.4 KB |  |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-charts | 6.57 ms | OfficeIMO.Excel | Win | 1891.6 KB |  |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | 4.53 ms | OfficeIMO.Excel | Win | 1405.8 KB |  |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-data-validation | 4.76 ms | OfficeIMO.Excel | Win | 1356.1 KB |  |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | 4.25 ms | OfficeIMO.Excel | Win | 1342.9 KB |  |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | 10.98 ms | OfficeIMO.Excel | Win | 5494.9 KB |  |
| 2500 | speed-comparison | other | Real-world report | realworld-report-all-in-one | 13.87 ms | OfficeIMO.Excel | Win | 6198.3 KB |  |
| 2500 | speed-comparison | other | Real-world report | realworld-report-core | 5.27 ms | OfficeIMO.Excel | Win | 1488.6 KB |  |
| 2500 | speed-comparison | other | Report workbook | report-workbook | 26.10 ms | OfficeIMO.Excel | Win | 7234.1 KB |  |
| 2500 | speed-comparison | other | Report workbook | report-workbook-core | 7.48 ms | OfficeIMO.Excel | Win | 2711.1 KB |  |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable | 18.81 ms | OfficeIMO.Excel | Win | 7548.8 KB |  |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable-core | 7.58 ms | OfficeIMO.Excel | Win | 2982.8 KB |  |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | 3.23 ms | OfficeIMO.Excel | Win | 706.7 KB |  |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | 1.33 ms | OfficeIMO.Excel | Win | 177.3 KB |  |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | 2.33 ms | Sylvan.Data.Excel | Loss +44.4% | 177.3 KB |  |
| 2500 | speed-comparison | read | Other | shared-string-read | 3.89 ms | Sylvan.Data.Excel | Loss +61.4% | 1056.5 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | 6.31 ms | Sylvan.Data.Excel | Loss +23.2% | 374.5 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-datatable | 21.50 ms | Sylvan.Data.Excel | Loss +95.7% | 3594.5 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | 5.97 ms | OfficeIMO.Excel | Win | 543.0 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-range | 11.52 ms | OfficeIMO.Excel | Win | 2692.5 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | 6.46 ms | OfficeIMO.Excel | Win | 2751.1 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-top-range | 0.77 ms | Sylvan.Data.Excel | Loss +27.8% | 296.0 KB |  |
| 2500 | speed-comparison | read | Range and table read | read-used-range | 8.52 ms | Sylvan.Data.Excel | Loss +25.8% | 2750.3 KB |  |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | 5.85 ms | OfficeIMO.Excel | Win | 377.9 KB |  |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | 16.61 ms | Sylvan.Data.Excel | Loss +142.7% | 2771.4 KB |  |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | 0.69 ms | Sylvan.Data.Excel | Loss +5.7% | 299.4 KB |  |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | 0.72 ms | Sylvan.Data.Excel | Loss +26.9% | 300.2 KB |  |
| 2500 | speed-comparison | read | Typed object read | read-objects | 9.62 ms | Sylvan.Data.Excel | Loss +15.0% | 2442.0 KB |  |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | 7.32 ms | Sylvan.Data.Excel | Loss +11.4% | 2422.9 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | 6.87 ms | OfficeIMO.Excel | Win | 1781.2 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | 6.06 ms | OfficeIMO.Excel | Win | 2079.3 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | 5.95 ms | OfficeIMO.Excel | Win | 1347.1 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | 5.93 ms | OfficeIMO.Excel | Win | 1505.3 KB |  |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | 4.96 ms | OfficeIMO.Excel | Win | 1346.4 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | 2.17 ms | OfficeIMO.Excel | Win | 1787.1 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | 2.01 ms | OfficeIMO.Excel | Win | 1119.9 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | 3.34 ms | OfficeIMO.Excel | Win | 1763.1 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | 3.47 ms | OfficeIMO.Excel | Win | 1506.7 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | 3.65 ms | OfficeIMO.Excel | Win | 1506.8 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | 1.85 ms | OfficeIMO.Excel | Win | 1138.1 KB |  |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | 3.11 ms | OfficeIMO.Excel | Win | 1435.5 KB |  |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | 5.96 ms | OfficeIMO.Excel | Win | 2064.5 KB |  |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | 7.23 ms | OfficeIMO.Excel | Win | 2880.2 KB |  |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | 6.40 ms | OfficeIMO.Excel | Win | 2067.7 KB |  |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | 5.22 ms | OfficeIMO.Excel | Win | 1774.9 KB |  |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | 5.01 ms | OfficeIMO.Excel | Win | 1748.6 KB |  |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | 4.57 ms | OfficeIMO.Excel | Win | 1487.2 KB |  |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | 4.70 ms | OfficeIMO.Excel | Win | 1760.7 KB |  |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | 7.20 ms | OfficeIMO.Excel | Win | 1403.3 KB |  |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | 4.33 ms | OfficeIMO.Excel | Win | 1620.6 KB |  |
| 2500 | speed-comparison | write | Other | write-dictionary-objects-table-direct | 6.69 ms | OfficeIMO.Excel | Win | 2051.4 KB |  |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | 6.99 ms | LargeXlsx | Loss +30.3% | 2341.7 KB |  |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | 7.68 ms | LargeXlsx | Loss +25.2% | 1507.7 KB |  |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | 23.95 ms | OfficeIMO.Excel | Win | 4502.3 KB |  |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | 2.79 ms | LargeXlsx | Loss +43.5% | 1576.3 KB |  |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | 1.57 ms | LargeXlsx | Loss +27.0% | 1092.0 KB |  |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | 4.14 ms | LargeXlsx | Loss +29.7% | 2081.1 KB |  |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | 2.27 ms | OfficeIMO.Excel | Win | 1494.0 KB |  |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | 4.47 ms | Sylvan.Data.Excel | Loss +14.0% | 1763.0 KB |  |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | 6.99 ms | OfficeIMO.Excel | Win | 2140.6 KB |  |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | 4.51 ms | LargeXlsx | Loss +6.9% | 1676.8 KB |  |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | 2.99 ms | OfficeIMO.Excel | Win | 2440.3 KB |  |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | 2.78 ms | OfficeIMO.Excel | Win | 2617.0 KB |  |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | 2.66 ms | OfficeIMO.Excel | Win | 2379.2 KB |  |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | 2.13 ms | OfficeIMO.Excel | Win | 1579.8 KB |  |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | 4.60 ms | LargeXlsx | Loss +26.6% | 1769.2 KB |  |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | 5.65 ms | LargeXlsx | Loss +37.5% | 1339.3 KB |  |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | 5.39 ms | LargeXlsx | Loss +31.8% | 1497.5 KB |  |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | 48.16 ms | Sylvan.Data.Excel | Loss +28.4% | 23622.0 KB |  |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | 47.73 ms | Sylvan.Data.Excel | Loss +28.7% | 24404.4 KB |  |
| 25000 | package-profile | package | Package size | append-plain-rows | 14.25 ms | LargeXlsx | Loss +28.4% | 10842.5 KB | 610.4 KB |
| 25000 | package-profile | package | Package size | autofit-existing | 85.90 ms | OfficeIMO.Excel | Win | 15708.3 KB | 1429.3 KB |
| 25000 | package-profile | package | Package size | large-shared-strings | 14.85 ms | OfficeIMO.Excel | Win | 15744.9 KB | 529.7 KB |
| 25000 | package-profile | package | Package size | realworld-autofilter | 30.76 ms | OfficeIMO.Excel | Win | 11494.9 KB | 1428.5 KB |
| 25000 | package-profile | package | Package size | realworld-charts | 36.44 ms | OfficeIMO.Excel | Win | 12550.9 KB | 1433.6 KB |
| 25000 | package-profile | package | Package size | realworld-conditional-formatting | 30.36 ms | OfficeIMO.Excel | Win | 11560.2 KB | 1428.8 KB |
| 25000 | package-profile | package | Package size | realworld-data-validation | 32.26 ms | OfficeIMO.Excel | Win | 11510.5 KB | 1428.6 KB |
| 25000 | package-profile | package | Package size | realworld-freeze-panes | 31.38 ms | OfficeIMO.Excel | Win | 11497.3 KB | 1428.6 KB |
| 25000 | package-profile | package | Package size | realworld-pivot-table | 81.89 ms | OfficeIMO.Excel | Win | 42218.3 KB | 1979.8 KB |
| 25000 | package-profile | package | Package size | realworld-report-all-in-one | 88.60 ms | OfficeIMO.Excel | Win | 43677.9 KB | 1985.9 KB |
| 25000 | package-profile | package | Package size | realworld-report-chart-first | 90.24 ms | OfficeIMO.Excel | Win | 43564.4 KB | 1985.9 KB |
| 25000 | package-profile | package | Package size | realworld-report-core | 34.89 ms | OfficeIMO.Excel | Win | 11648.7 KB | 1430.0 KB |
| 25000 | package-profile | package | Package size | realworld-report-extra-column | 99.55 ms | OfficeIMO.Excel | Win | 45561.9 KB | 2110.8 KB |
| 25000 | package-profile | package | Package size | realworld-report-no-autofit | 82.57 ms | OfficeIMO.Excel | Win | 43671.9 KB | 1985.8 KB |
| 25000 | package-profile | package | Package size | realworld-report-post-mutation | 91.54 ms | OfficeIMO.Excel | Win | 43687.3 KB | 1986.0 KB |
| 25000 | package-profile | package | Package size | realworld-report-shuffled-columns | 100.92 ms | OfficeIMO.Excel | Win | 43743.0 KB | 2046.1 KB |
| 25000 | package-profile | package | Package size | report-workbook | 114.88 ms | OfficeIMO.Excel | Win | 59187.8 KB | 2672.0 KB |
| 25000 | package-profile | package | Package size | report-workbook-core | 48.41 ms | OfficeIMO.Excel | Win | 10979.4 KB | 1850.9 KB |
| 25000 | package-profile | package | Package size | report-workbook-datatable | 125.75 ms | OfficeIMO.Excel | Win | 61933.5 KB | 2672.0 KB |
| 25000 | package-profile | package | Package size | report-workbook-datatable-core | 48.98 ms | OfficeIMO.Excel | Win | 13725.0 KB | 1850.9 KB |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | 62.44 ms | LargeXlsx | Loss +17.8% | 11708.2 KB | 2228.8 KB |
| 25000 | package-profile | package | Package size | write-bulk-report | 35.39 ms | OfficeIMO.Excel | Win | 11561.8 KB | 1429.3 KB |
| 25000 | package-profile | package | Package size | write-cellformula | 24.74 ms | OfficeIMO.Excel | Win | 10112.0 KB | 670.3 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-empty-strings | 14.96 ms | OfficeIMO.Excel | Win | 6896.4 KB | 451.4 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-numbers | 22.16 ms | OfficeIMO.Excel | Win | 5970.9 KB | 462.6 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-object-mixed | 26.90 ms | OfficeIMO.Excel | Win | 8332.9 KB | 585.1 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse | 24.34 ms | OfficeIMO.Excel | Win | 7416.2 KB | 607.1 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | 20.17 ms | OfficeIMO.Excel | Win | 7416.3 KB | 607.1 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-scalars | 14.54 ms | OfficeIMO.Excel | Win | 6144.6 KB | 441.9 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-strings | 21.69 ms | OfficeIMO.Excel | Win | 15360.4 KB | 527.8 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-distinct | 18.79 ms | OfficeIMO.Excel | Win | 13824.1 KB | 499.0 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-repeated | 15.44 ms | OfficeIMO.Excel | Win | 7525.3 KB | 376.0 KB |
| 25000 | package-profile | package | Package size | write-cellvalue-temporal | 27.28 ms | OfficeIMO.Excel | Win | 7482.8 KB | 620.5 KB |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | 14.79 ms | LargeXlsx | Loss +12.9% | 6961.7 KB | 455.5 KB |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | 44.94 ms | LargeXlsx | Loss +32.8% | 16036.5 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-datareader-plain | 47.47 ms | Sylvan.Data.Excel | Loss +23.2% | 13002.3 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-datareader-table | 51.03 ms | OfficeIMO.Excel | Win | 13020.3 KB | 1385.7 KB |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | 45.06 ms | OfficeIMO.Excel | Win | 13026.6 KB | 1385.9 KB |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | 31.35 ms | OfficeIMO.Excel | Win | 9819.7 KB | 1329.2 KB |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | 37.94 ms | OfficeIMO.Excel | Win | 13458.5 KB | 1795.1 KB |
| 25000 | package-profile | package | Package size | write-dataset-tables | 32.75 ms | OfficeIMO.Excel | Win | 10288.1 KB | 1376.4 KB |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | 37.93 ms | OfficeIMO.Excel | Win | 10300.7 KB | 1376.7 KB |
| 25000 | package-profile | package | Package size | write-datatable-direct | 48.70 ms | LargeXlsx | Loss +4.1% | 12715.7 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | 43.59 ms | OfficeIMO.Excel | Win | 12733.8 KB | 1385.7 KB |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | 31.39 ms | LargeXlsx | Loss +15.6% | 12912.0 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | 41.71 ms | OfficeIMO.Excel | Win | 11501.6 KB | 1428.5 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | 41.89 ms | LargeXlsx | Loss +12.5% | 11493.8 KB | 1428.4 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | 41.91 ms | OfficeIMO.Excel | Win | 10187.2 KB | 1385.1 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | 37.17 ms | LargeXlsx | Loss +31.1% | 10179.4 KB | 1384.9 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | 39.14 ms | LargeXlsx | Loss +51.8% | 15791.7 KB | 1428.4 KB |
| 25000 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | 33.39 ms | OfficeIMO.Excel | Win | 11500.9 KB | 1428.5 KB |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | 45.00 ms | LargeXlsx | Loss +20.6% | 10577.2 KB | 1828.0 KB |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | 42.38 ms | LargeXlsx | Loss +14.5% | 9942.2 KB | 1828.0 KB |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | 190.90 ms | OfficeIMO.Excel | Win | 36150.1 KB | 6725.6 KB |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | 84.25 ms | OfficeIMO.Excel | Win | 15708.5 KB |  |
| 25000 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | 89.46 ms | OfficeIMO.Excel | Win | 43670.5 KB |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | 89.50 ms | OfficeIMO.Excel | Win | 43558.6 KB |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | 98.59 ms | OfficeIMO.Excel | Win | 45565.9 KB |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | 93.32 ms | OfficeIMO.Excel | Win | 43686.0 KB |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | 105.35 ms | OfficeIMO.Excel | Win | 43738.1 KB |  |
| 25000 | speed-comparison | other | Object projection | build-object-datatable-dictionaries | 12.67 ms | OfficeIMO.Excel | Win | 5164.3 KB |  |
| 25000 | speed-comparison | other | Object projection | build-object-datatable-typed | 10.47 ms | OfficeIMO.Excel | Win | 8093.8 KB |  |
| 25000 | speed-comparison | other | Range and table read | enumerate-cells | 55.93 ms | OfficeIMO.Excel | Win | 24530.8 KB |  |
| 25000 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | 41.78 ms | OfficeIMO.Excel | Win | 3839.2 KB |  |
| 25000 | speed-comparison | other | Range and table read | enumerate-range | 67.22 ms | OfficeIMO.Excel | Win | 24531.0 KB |  |
| 25000 | speed-comparison | other | Range and table read | enumerate-top-range | 0.74 ms | OfficeIMO.Excel | Win | 285.3 KB |  |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-autofilter | 32.65 ms | OfficeIMO.Excel | Win | 11494.9 KB |  |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-charts | 32.68 ms | OfficeIMO.Excel | Win | 12549.2 KB |  |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | 31.89 ms | OfficeIMO.Excel | Win | 11560.2 KB |  |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-data-validation | 32.47 ms | OfficeIMO.Excel | Win | 11510.5 KB |  |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | 33.62 ms | OfficeIMO.Excel | Win | 11497.3 KB |  |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | 83.49 ms | OfficeIMO.Excel | Win | 42217.9 KB |  |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | 91.44 ms | OfficeIMO.Excel | Win | 43675.4 KB |  |
| 25000 | speed-comparison | other | Real-world report | realworld-report-core | 34.81 ms | OfficeIMO.Excel | Win | 11648.7 KB |  |
| 25000 | speed-comparison | other | Report workbook | report-workbook | 116.92 ms | OfficeIMO.Excel | Win | 59145.6 KB |  |
| 25000 | speed-comparison | other | Report workbook | report-workbook-core | 46.03 ms | OfficeIMO.Excel | Win | 10979.4 KB |  |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable | 126.00 ms | OfficeIMO.Excel | Win | 61935.3 KB |  |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable-core | 49.99 ms | OfficeIMO.Excel | Win | 13725.0 KB |  |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | 19.08 ms | OfficeIMO.Excel | Win | 6219.0 KB |  |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | 0.84 ms | OfficeIMO.Excel | Win | 177.3 KB |  |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | 0.83 ms | OfficeIMO.Excel | Win | 177.4 KB |  |
| 25000 | speed-comparison | read | Other | shared-string-read | 17.94 ms | Sylvan.Data.Excel | Loss +14.4% | 9218.0 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | 37.82 ms | OfficeIMO.Excel | Win | 1122.3 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-datatable | 82.41 ms | Sylvan.Data.Excel | Loss +2.4% | 34645.8 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | 41.76 ms | OfficeIMO.Excel | Win | 4034.6 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-range | 70.99 ms | Sylvan.Data.Excel | Loss +13.4% | 26098.2 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | 74.37 ms | Sylvan.Data.Excel | Loss +21.1% | 26684.1 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-top-range | 0.67 ms | Sylvan.Data.Excel | Loss +31.6% | 296.0 KB |  |
| 25000 | speed-comparison | read | Range and table read | read-used-range | 72.30 ms | Sylvan.Data.Excel | Loss +8.9% | 26156.0 KB |  |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | 31.74 ms | OfficeIMO.Excel | Win | 1125.7 KB |  |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | 58.57 ms | Sylvan.Data.Excel | Loss +29.3% | 26885.3 KB |  |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | 0.53 ms | Sylvan.Data.Excel | Loss +24.7% | 299.3 KB |  |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | 0.54 ms | Sylvan.Data.Excel | Loss +33.3% | 300.0 KB |  |
| 25000 | speed-comparison | read | Typed object read | read-objects | 56.11 ms | Sylvan.Data.Excel | Loss +29.3% | 23562.3 KB |  |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | 57.68 ms | OfficeIMO.Excel | Win | 23367.3 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | 48.83 ms | OfficeIMO.Excel | Win | 13026.6 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | 47.84 ms | OfficeIMO.Excel | Win | 10300.7 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | 44.97 ms | OfficeIMO.Excel | Win | 11501.6 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | 51.37 ms | OfficeIMO.Excel | Win | 10187.2 KB |  |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | 39.87 ms | OfficeIMO.Excel | Win | 11500.9 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | 18.17 ms | OfficeIMO.Excel | Win | 6896.4 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | 21.13 ms | OfficeIMO.Excel | Win | 5970.9 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | 26.87 ms | OfficeIMO.Excel | Win | 8332.9 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | 25.88 ms | OfficeIMO.Excel | Win | 7416.2 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | 20.35 ms | OfficeIMO.Excel | Win | 7416.3 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | 16.44 ms | OfficeIMO.Excel | Win | 6144.6 KB |  |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | 28.12 ms | OfficeIMO.Excel | Win | 7482.8 KB |  |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | 56.02 ms | OfficeIMO.Excel | Win | 13039.6 KB |  |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | 54.08 ms | OfficeIMO.Excel | Win | 13458.5 KB |  |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | 45.61 ms | OfficeIMO.Excel | Win | 10288.1 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | 46.10 ms | OfficeIMO.Excel | Win | 13020.3 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | 39.58 ms | LargeXlsx | Loss +3.1% | 12715.7 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | 44.19 ms | OfficeIMO.Excel | Win | 9999.4 KB |  |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | 45.13 ms | OfficeIMO.Excel | Win | 12733.8 KB |  |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | 43.87 ms | OfficeIMO.Excel | Win | 11561.8 KB |  |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | 28.29 ms | OfficeIMO.Excel | Win | 10112.1 KB |  |
| 25000 | speed-comparison | write | Other | write-dictionary-objects-table-direct | 57.67 ms | OfficeIMO.Excel | Win | 15163.8 KB |  |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | 52.53 ms | LargeXlsx | Loss +3.7% | 10577.2 KB |  |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | 55.52 ms | LargeXlsx | Loss +22.9% | 9942.2 KB |  |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | 226.59 ms | OfficeIMO.Excel | Win | 36150.1 KB |  |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | 19.42 ms | LargeXlsx | Loss +29.4% | 10842.5 KB |  |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | 14.28 ms | LargeXlsx | Loss +11.2% | 6961.7 KB |  |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | 39.59 ms | LargeXlsx | Loss +11.4% | 16036.5 KB |  |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | 20.83 ms | OfficeIMO.Excel | Win | 7866.1 KB |  |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | 41.47 ms | Sylvan.Data.Excel | Loss +24.1% | 13002.3 KB |  |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | 45.09 ms | OfficeIMO.Excel | Win | 9819.7 KB |  |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | 57.24 ms | LargeXlsx | Loss +12.0% | 11708.2 KB |  |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | 14.54 ms | OfficeIMO.Excel | Win | 15744.9 KB |  |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | 23.58 ms | OfficeIMO.Excel | Win | 15360.4 KB |  |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | 17.30 ms | OfficeIMO.Excel | Win | 13824.1 KB |  |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | 18.09 ms | OfficeIMO.Excel | Win | 7525.3 KB |  |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | 39.68 ms | LargeXlsx | Loss +14.1% | 12912.0 KB |  |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | 39.19 ms | LargeXlsx | Loss +18.2% | 11493.8 KB |  |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | 47.75 ms | LargeXlsx | Loss +42.6% | 10179.4 KB |  |

## Full comparison table

| Row count | Artifact | Workload | Category | Scenario | Library | Mean | StdDev | StdErr | Ratio to OfficeIMO | Ratio to best | Alloc | Alloc ratio | Package | Package ratio | Outcome |
| ---: | --- | --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | Sylvan.Data.Excel | 6.74 ms | 0.38 ms | 0.22 ms | 0.83 | 1.00 | 362.3 KB | 0.15 |  |  | 17.0% faster than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | OfficeIMO.Excel | 8.12 ms | 1.34 ms | 0.78 ms | 1.00 | 1.20 | 2410.9 KB | 1.00 |  |  | Loss +20.4% |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | ExcelDataReader | 16.65 ms | 1.04 ms | 0.60 ms | 2.05 | 2.47 | 6887.4 KB | 2.86 |  |  | 105.1% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | MiniExcel | 20.71 ms | 3.73 ms | 2.15 ms | 2.55 | 3.07 | 21507.3 KB | 8.92 |  |  | 155.1% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | Sylvan.Data.Excel | 6.24 ms | 0.15 ms | 0.09 ms | 0.85 | 1.00 | 362.3 KB | 0.15 |  |  | 15.4% faster than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | OfficeIMO.Excel | 7.38 ms | 1.17 ms | 0.67 ms | 1.00 | 1.18 | 2489.4 KB | 1.00 |  |  | Loss +18.2% |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | ExcelDataReader | 15.69 ms | 2.06 ms | 1.19 ms | 2.13 | 2.51 | 6887.4 KB | 2.77 |  |  | 112.7% slower than OfficeIMO |
| 2500 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | MiniExcel | 20.40 ms | 1.28 ms | 0.74 ms | 2.77 | 3.27 | 21507.3 KB | 8.64 |  |  | 176.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | append-plain-rows | LargeXlsx | 1.79 ms | 0.42 ms | 0.24 ms | 0.80 | 1.00 | 296.4 KB | 0.19 | 63.1 KB | 1.00 | 20.2% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | append-plain-rows | OfficeIMO.Excel | 2.25 ms | 0.47 ms | 0.27 ms | 1.00 | 1.25 | 1576.3 KB | 1.00 | 63.0 KB | 1.00 | Loss +25.3% |
| 2500 | package-profile | package | Package size | append-plain-rows | MiniExcel | 4.52 ms | 0.28 ms | 0.16 ms | 2.01 | 2.52 | 19710.6 KB | 12.50 | 68.1 KB | 1.08 | 101.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | append-plain-rows | ClosedXML | 17.74 ms | 0.54 ms | 0.31 ms | 7.90 | 9.90 | 11197.4 KB | 7.10 | 59.8 KB | 0.95 | 690.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | append-plain-rows | EPPlus | 29.62 ms | 1.48 ms | 0.85 ms | 13.19 | 16.53 | 14365.2 KB | 9.11 | 56.9 KB | 0.90 | 1219.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | autofit-existing | OfficeIMO.Excel | 8.84 ms | 0.53 ms | 0.30 ms | 1.00 | 1.00 | 1895.3 KB | 1.00 | 143.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | autofit-existing | EPPlus | 88.48 ms | 2.96 ms | 1.71 ms | 10.01 | 10.01 | 50712.0 KB | 26.76 | 115.0 KB | 0.80 | 900.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | autofit-existing | ClosedXML | 154.22 ms | 14.42 ms | 8.32 ms | 17.44 | 17.44 | 84561.9 KB | 44.62 | 121.0 KB | 0.84 | 1644.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | large-shared-strings | OfficeIMO.Excel | 2.36 ms | 0.41 ms | 0.24 ms | 1.00 | 1.00 | 2440.3 KB | 1.00 | 55.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | large-shared-strings | MiniExcel | 4.81 ms | 0.62 ms | 0.36 ms | 2.04 | 2.04 | 21137.5 KB | 8.66 | 60.7 KB | 1.10 | 104.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | large-shared-strings | ClosedXML | 13.22 ms | 1.30 ms | 0.75 ms | 5.61 | 5.61 | 11299.2 KB | 4.63 | 50.3 KB | 0.91 | 460.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | large-shared-strings | EPPlus | 24.93 ms | 1.86 ms | 1.07 ms | 10.58 | 10.58 | 12804.4 KB | 5.25 | 48.1 KB | 0.87 | 957.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-autofilter | OfficeIMO.Excel | 4.61 ms | 0.21 ms | 0.12 ms | 1.00 | 1.00 | 1340.4 KB | 1.00 | 142.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-autofilter | ClosedXML | 38.08 ms | 2.23 ms | 1.29 ms | 8.26 | 8.26 | 22226.8 KB | 16.58 | 120.2 KB | 0.84 | 726.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-autofilter | EPPlus | 48.85 ms | 1.64 ms | 0.94 ms | 10.60 | 10.60 | 24715.5 KB | 18.44 | 114.2 KB | 0.80 | 959.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-charts | OfficeIMO.Excel | 6.37 ms | 0.50 ms | 0.29 ms | 1.00 | 1.00 | 1891.6 KB | 1.00 | 147.6 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-charts | EPPlus | 50.42 ms | 1.71 ms | 0.99 ms | 7.92 | 7.92 | 27142.3 KB | 14.35 | 117.0 KB | 0.79 | 692.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-conditional-formatting | OfficeIMO.Excel | 4.19 ms | 0.33 ms | 0.19 ms | 1.00 | 1.00 | 1405.8 KB | 1.00 | 142.7 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-conditional-formatting | ClosedXML | 34.23 ms | 1.30 ms | 0.75 ms | 8.16 | 8.16 | 22273.8 KB | 15.84 | 120.3 KB | 0.84 | 716.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-conditional-formatting | EPPlus | 44.59 ms | 0.55 ms | 0.32 ms | 10.63 | 10.63 | 24757.5 KB | 17.61 | 114.3 KB | 0.80 | 963.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-data-validation | OfficeIMO.Excel | 4.94 ms | 0.23 ms | 0.13 ms | 1.00 | 1.00 | 1356.1 KB | 1.00 | 142.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-data-validation | ClosedXML | 37.64 ms | 1.82 ms | 1.05 ms | 7.62 | 7.62 | 22247.9 KB | 16.41 | 120.3 KB | 0.84 | 661.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-data-validation | EPPlus | 49.86 ms | 2.25 ms | 1.30 ms | 10.09 | 10.09 | 24701.4 KB | 18.22 | 114.2 KB | 0.80 | 909.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-freeze-panes | OfficeIMO.Excel | 4.03 ms | 0.52 ms | 0.30 ms | 1.00 | 1.00 | 1342.8 KB | 1.00 | 142.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-freeze-panes | ClosedXML | 38.72 ms | 4.41 ms | 2.55 ms | 9.60 | 9.60 | 22222.0 KB | 16.55 | 120.2 KB | 0.84 | 860.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-freeze-panes | EPPlus | 49.75 ms | 5.81 ms | 3.35 ms | 12.34 | 12.34 | 24730.0 KB | 18.42 | 114.3 KB | 0.80 | 1133.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-pivot-table | OfficeIMO.Excel | 10.88 ms | 0.54 ms | 0.31 ms | 1.00 | 1.00 | 5495.0 KB | 1.00 | 200.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-pivot-table | EPPlus | 53.85 ms | 2.36 ms | 1.36 ms | 4.95 | 4.95 | 29537.4 KB | 5.38 | 117.4 KB | 0.59 | 394.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-all-in-one | OfficeIMO.Excel | 14.06 ms | 1.01 ms | 0.58 ms | 1.00 | 1.00 | 6197.0 KB | 1.00 | 206.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-all-in-one | EPPlus | 81.69 ms | 2.63 ms | 1.52 ms | 5.81 | 5.81 | 54595.0 KB | 8.81 | 121.8 KB | 0.59 | 481.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-chart-first | OfficeIMO.Excel | 13.49 ms | 1.59 ms | 0.92 ms | 1.00 | 1.00 | 6198.7 KB | 1.00 | 206.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-chart-first | EPPlus | 84.74 ms | 5.35 ms | 3.09 ms | 6.28 | 6.28 | 54594.3 KB | 8.81 | 121.8 KB | 0.59 | 528.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-core | OfficeIMO.Excel | 4.93 ms | 0.70 ms | 0.40 ms | 1.00 | 1.00 | 1488.5 KB | 1.00 | 143.9 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-core | EPPlus | 77.21 ms | 1.71 ms | 0.99 ms | 15.65 | 15.65 | 47299.8 KB | 31.78 | 115.6 KB | 0.80 | 1465.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-core | ClosedXML | 96.60 ms | 6.88 ms | 3.97 ms | 19.59 | 19.59 | 69833.7 KB | 46.91 | 121.5 KB | 0.84 | 1858.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-extra-column | OfficeIMO.Excel | 14.28 ms | 0.19 ms | 0.11 ms | 1.00 | 1.00 | 6389.4 KB | 1.00 | 219.1 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-extra-column | EPPlus | 92.46 ms | 2.94 ms | 1.70 ms | 6.47 | 6.47 | 59226.6 KB | 9.27 | 128.4 KB | 0.59 | 547.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-no-autofit | OfficeIMO.Excel | 11.95 ms | 1.09 ms | 0.63 ms | 1.00 | 1.00 | 6190.2 KB | 1.00 | 206.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-no-autofit | EPPlus | 53.86 ms | 5.43 ms | 3.13 ms | 4.51 | 4.51 | 32906.8 KB | 5.32 | 121.8 KB | 0.59 | 350.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-post-mutation | OfficeIMO.Excel | 12.69 ms | 0.35 ms | 0.20 ms | 1.00 | 1.00 | 6208.3 KB | 1.00 | 206.6 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-post-mutation | EPPlus | 90.40 ms | 5.42 ms | 3.13 ms | 7.12 | 7.12 | 54594.9 KB | 8.79 | 121.9 KB | 0.59 | 612.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-shuffled-columns | OfficeIMO.Excel | 15.33 ms | 0.35 ms | 0.20 ms | 1.00 | 1.00 | 6201.5 KB | 1.00 | 211.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-shuffled-columns | EPPlus | 84.62 ms | 6.61 ms | 3.81 ms | 5.52 | 5.52 | 54591.5 KB | 8.80 | 124.3 KB | 0.59 | 451.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook | OfficeIMO.Excel | 17.65 ms | 0.83 ms | 0.48 ms | 1.00 | 1.00 | 7277.0 KB | 1.00 | 275.6 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | report-workbook | EPPlus | 108.51 ms | 4.68 ms | 2.70 ms | 6.15 | 6.15 | 77486.2 KB | 10.65 | 161.8 KB | 0.59 | 514.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook-core | OfficeIMO.Excel | 7.96 ms | 0.42 ms | 0.24 ms | 1.00 | 1.00 | 2711.1 KB | 1.00 | 187.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | report-workbook-core | ClosedXML | 119.39 ms | 4.19 ms | 2.42 ms | 15.00 | 15.00 | 97218.9 KB | 35.86 | 165.1 KB | 0.88 | 1400.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook-core | EPPlus | 119.81 ms | 4.33 ms | 2.50 ms | 15.06 | 15.06 | 71970.6 KB | 26.55 | 157.2 KB | 0.84 | 1405.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook-datatable | OfficeIMO.Excel | 17.85 ms | 0.60 ms | 0.34 ms | 1.00 | 1.00 | 7548.5 KB | 1.00 | 275.6 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | report-workbook-datatable | EPPlus | 122.78 ms | 6.69 ms | 3.86 ms | 6.88 | 6.88 | 65995.3 KB | 8.74 | 161.8 KB | 0.59 | 587.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook-datatable-core | OfficeIMO.Excel | 7.45 ms | 1.50 ms | 0.87 ms | 1.00 | 1.00 | 2982.7 KB | 1.00 | 187.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | report-workbook-datatable-core | EPPlus | 109.00 ms | 5.27 ms | 3.04 ms | 14.64 | 14.64 | 60480.1 KB | 20.28 | 157.2 KB | 0.84 | 1363.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | report-workbook-datatable-core | ClosedXML | 121.77 ms | 7.19 ms | 4.15 ms | 16.35 | 16.35 | 82861.2 KB | 27.78 | 165.1 KB | 0.88 | 1535.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | LargeXlsx | 5.07 ms | 0.63 ms | 0.36 ms | 0.95 | 1.00 | 857.6 KB | 0.51 | 237.7 KB | 1.10 | 5.5% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | OfficeIMO.Excel | 5.36 ms | 0.65 ms | 0.38 ms | 1.00 | 1.06 | 1676.8 KB | 1.00 | 216.7 KB | 1.00 | Loss +5.8% |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | MiniExcel | 22.32 ms | 0.95 ms | 0.55 ms | 4.16 | 4.40 | 35919.1 KB | 21.42 | 235.3 KB | 1.09 | 316.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-blog-2023-20-string-columns | ClosedXML | 104.61 ms | 4.75 ms | 2.74 ms | 19.50 | 20.63 | 71478.2 KB | 42.63 | 257.2 KB | 1.19 | 1849.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-bulk-report | OfficeIMO.Excel | 4.33 ms | 0.20 ms | 0.11 ms | 1.00 | 1.00 | 1401.7 KB | 1.00 | 143.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-bulk-report | MiniExcel | 9.96 ms | 1.03 ms | 0.59 ms | 2.30 | 2.30 | 26825.4 KB | 19.14 | 153.8 KB | 1.07 | 129.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-bulk-report | EPPlus | 79.07 ms | 8.40 ms | 4.85 ms | 18.25 | 18.25 | 47193.8 KB | 33.67 | 115.0 KB | 0.80 | 1724.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-bulk-report | ClosedXML | 91.62 ms | 4.60 ms | 2.66 ms | 21.14 | 21.14 | 58348.7 KB | 41.63 | 121.0 KB | 0.84 | 2014.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellformula | OfficeIMO.Excel | 3.39 ms | 0.20 ms | 0.12 ms | 1.00 | 1.00 | 1383.3 KB | 1.00 | 66.6 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellformula | ClosedXML | 24.79 ms | 3.77 ms | 2.18 ms | 7.31 | 7.31 | 12039.8 KB | 8.70 | 70.6 KB | 1.06 | 631.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellformula | EPPlus | 44.04 ms | 2.78 ms | 1.61 ms | 12.99 | 12.99 | 18110.5 KB | 13.09 | 62.1 KB | 0.93 | 1198.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-empty-strings | OfficeIMO.Excel | 2.37 ms | 0.11 ms | 0.06 ms | 1.00 | 1.00 | 1787.1 KB | 1.00 | 44.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-empty-strings | ClosedXML | 13.31 ms | 0.96 ms | 0.55 ms | 5.62 | 5.62 | 9959.5 KB | 5.57 | 44.9 KB | 1.02 | 461.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-empty-strings | EPPlus | 24.48 ms | 1.60 ms | 0.92 ms | 10.33 | 10.33 | 11773.0 KB | 6.59 | 42.0 KB | 0.95 | 933.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-numbers | OfficeIMO.Excel | 2.49 ms | 0.42 ms | 0.24 ms | 1.00 | 1.00 | 1119.9 KB | 1.00 | 47.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-numbers | ClosedXML | 12.57 ms | 1.22 ms | 0.71 ms | 5.06 | 5.06 | 9177.1 KB | 8.19 | 45.9 KB | 0.98 | 405.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-numbers | EPPlus | 25.76 ms | 0.60 ms | 0.35 ms | 10.36 | 10.36 | 12895.3 KB | 11.51 | 43.7 KB | 0.93 | 936.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-mixed | OfficeIMO.Excel | 3.22 ms | 0.34 ms | 0.20 ms | 1.00 | 1.00 | 1763.3 KB | 1.00 | 61.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-object-mixed | ClosedXML | 20.79 ms | 0.54 ms | 0.31 ms | 6.46 | 6.46 | 11887.0 KB | 6.74 | 59.5 KB | 0.97 | 546.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-mixed | EPPlus | 32.49 ms | 1.06 ms | 0.61 ms | 10.10 | 10.10 | 15643.4 KB | 8.87 | 58.9 KB | 0.96 | 910.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse | OfficeIMO.Excel | 3.17 ms | 0.70 ms | 0.41 ms | 1.00 | 1.00 | 1506.9 KB | 1.00 | 62.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse | ClosedXML | 17.85 ms | 0.20 ms | 0.12 ms | 5.63 | 5.63 | 11296.3 KB | 7.50 | 52.5 KB | 0.85 | 462.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse | EPPlus | 31.58 ms | 1.47 ms | 0.85 ms | 9.96 | 9.96 | 14960.3 KB | 9.93 | 54.2 KB | 0.88 | 895.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 3.11 ms | 0.33 ms | 0.19 ms | 1.00 | 1.00 | 1507.0 KB | 1.00 | 62.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | ClosedXML | 24.55 ms | 6.67 ms | 3.85 ms | 7.89 | 7.89 | 11296.3 KB | 7.50 | 52.5 KB | 0.85 | 688.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | EPPlus | 32.53 ms | 2.39 ms | 1.38 ms | 10.46 | 10.46 | 14960.3 KB | 9.93 | 54.2 KB | 0.88 | 945.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-scalars | OfficeIMO.Excel | 2.13 ms | 0.54 ms | 0.31 ms | 1.00 | 1.00 | 1138.1 KB | 1.00 | 46.3 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-scalars | ClosedXML | 14.10 ms | 2.36 ms | 1.36 ms | 6.63 | 6.63 | 9021.2 KB | 7.93 | 45.4 KB | 0.98 | 563.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-scalars | EPPlus | 27.22 ms | 0.73 ms | 0.42 ms | 12.80 | 12.80 | 12827.5 KB | 11.27 | 42.4 KB | 0.91 | 1179.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings | OfficeIMO.Excel | 3.87 ms | 0.43 ms | 0.25 ms | 1.00 | 1.00 | 2617.0 KB | 1.00 | 55.1 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-strings | ClosedXML | 14.43 ms | 2.21 ms | 1.27 ms | 3.73 | 3.73 | 11299.2 KB | 4.32 | 50.3 KB | 0.91 | 272.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings | EPPlus | 25.28 ms | 2.35 ms | 1.36 ms | 6.53 | 6.53 | 12804.9 KB | 4.89 | 48.1 KB | 0.87 | 553.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-distinct | OfficeIMO.Excel | 2.43 ms | 0.60 ms | 0.35 ms | 1.00 | 1.00 | 2379.2 KB | 1.00 | 51.8 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-distinct | ClosedXML | 16.10 ms | 1.08 ms | 0.63 ms | 6.63 | 6.63 | 13127.1 KB | 5.52 | 61.9 KB | 1.19 | 562.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-distinct | EPPlus | 28.89 ms | 1.59 ms | 0.92 ms | 11.89 | 11.89 | 13893.0 KB | 5.84 | 61.5 KB | 1.19 | 1089.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-repeated | OfficeIMO.Excel | 2.14 ms | 0.29 ms | 0.17 ms | 1.00 | 1.00 | 1579.8 KB | 1.00 | 40.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-repeated | ClosedXML | 11.76 ms | 0.56 ms | 0.33 ms | 5.50 | 5.50 | 9226.5 KB | 5.84 | 38.8 KB | 0.97 | 450.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-strings-repeated | EPPlus | 23.25 ms | 2.42 ms | 1.40 ms | 10.88 | 10.88 | 11332.5 KB | 7.17 | 34.8 KB | 0.87 | 988.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-temporal | OfficeIMO.Excel | 2.80 ms | 0.18 ms | 0.11 ms | 1.00 | 1.00 | 1435.7 KB | 1.00 | 63.3 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-cellvalue-temporal | ClosedXML | 18.76 ms | 4.55 ms | 2.63 ms | 6.71 | 6.71 | 9711.1 KB | 6.76 | 54.5 KB | 0.86 | 571.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalue-temporal | EPPlus | 30.46 ms | 4.42 ms | 2.55 ms | 10.90 | 10.90 | 14722.7 KB | 10.25 | 53.1 KB | 0.84 | 989.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 1.39 ms | 0.24 ms | 0.14 ms | 0.81 | 1.00 | 447.0 KB | 0.41 | 47.3 KB | 0.98 | 18.9% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 1.71 ms | 0.23 ms | 0.13 ms | 1.00 | 1.23 | 1092.0 KB | 1.00 | 48.2 KB | 1.00 | Loss +23.2% |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | ClosedXML | 13.32 ms | 1.06 ms | 0.61 ms | 7.79 | 9.60 | 10235.8 KB | 9.37 | 53.0 KB | 1.10 | 678.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | EPPlus | 25.99 ms | 0.67 ms | 0.39 ms | 15.20 | 18.73 | 13052.1 KB | 11.95 | 52.5 KB | 1.09 | 1420.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | LargeXlsx | 3.87 ms | 0.37 ms | 0.21 ms | 0.83 | 1.00 | 758.3 KB | 0.36 | 138.4 KB | 1.00 | 16.5% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 4.64 ms | 0.51 ms | 0.29 ms | 1.00 | 1.20 | 2081.1 KB | 1.00 | 138.0 KB | 1.00 | Loss +19.8% |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | MiniExcel | 9.47 ms | 0.59 ms | 0.34 ms | 2.04 | 2.44 | 23222.3 KB | 11.16 | 153.7 KB | 1.11 | 104.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | ClosedXML | 40.20 ms | 4.19 ms | 2.42 ms | 8.66 | 10.38 | 22221.3 KB | 10.68 | 120.1 KB | 0.87 | 766.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-cellvalues-rectangle-direct | EPPlus | 45.62 ms | 2.51 ms | 1.45 ms | 9.83 | 11.78 | 24694.0 KB | 11.87 | 114.1 KB | 0.83 | 883.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | Sylvan.Data.Excel | 3.69 ms | 0.24 ms | 0.14 ms | 0.83 | 1.00 | 758.7 KB | 0.43 | 78.5 KB | 0.57 | 17.3% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | OfficeIMO.Excel | 4.46 ms | 0.12 ms | 0.07 ms | 1.00 | 1.21 | 1763.0 KB | 1.00 | 138.0 KB | 1.00 | Loss +21.0% |
| 2500 | package-profile | package | Package size | write-datareader-plain | LargeXlsx | 4.51 ms | 0.08 ms | 0.05 ms | 1.01 | 1.22 | 1032.5 KB | 0.59 | 138.4 KB | 1.00 | Tie vs OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | MiniExcel | 8.62 ms | 1.04 ms | 0.60 ms | 1.93 | 2.34 | 23043.8 KB | 13.07 | 153.6 KB | 1.11 | 93.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | ClosedXML | 31.74 ms | 3.09 ms | 1.78 ms | 7.11 | 8.61 | 11581.0 KB | 6.57 | 120.1 KB | 0.87 | 611.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-plain | EPPlus | 43.07 ms | 0.52 ms | 0.30 ms | 9.65 | 11.68 | 16646.4 KB | 9.44 | 114.9 KB | 0.83 | 865.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table | OfficeIMO.Excel | 4.83 ms | 0.57 ms | 0.33 ms | 1.00 | 1.00 | 1774.9 KB | 1.00 | 138.8 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-datareader-table | MiniExcel | 8.15 ms | 0.90 ms | 0.52 ms | 1.69 | 1.69 | 23044.1 KB | 12.98 | 153.6 KB | 1.11 | 68.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table | ClosedXML | 42.68 ms | 3.15 ms | 1.82 ms | 8.84 | 8.84 | 19007.4 KB | 10.71 | 120.9 KB | 0.87 | 784.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table | EPPlus | 42.98 ms | 0.91 ms | 0.52 ms | 8.90 | 8.90 | 16646.1 KB | 9.38 | 114.9 KB | 0.83 | 790.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | OfficeIMO.Excel | 5.27 ms | 0.40 ms | 0.23 ms | 1.00 | 1.00 | 1781.2 KB | 1.00 | 138.9 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | MiniExcel | 10.06 ms | 0.24 ms | 0.14 ms | 1.91 | 1.91 | 26647.3 KB | 14.96 | 153.8 KB | 1.11 | 90.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | EPPlus | 66.73 ms | 5.27 ms | 3.04 ms | 12.67 | 12.67 | 38343.6 KB | 21.53 | 115.1 KB | 0.83 | 1166.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datareader-table-autofit | ClosedXML | 87.41 ms | 5.65 ms | 3.26 ms | 16.59 | 16.59 | 58360.1 KB | 32.77 | 121.0 KB | 0.87 | 1559.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | OfficeIMO.Excel | 4.53 ms | 0.40 ms | 0.23 ms | 1.00 | 1.00 | 2140.6 KB | 1.00 | 131.1 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | LargeXlsx | 5.15 ms | 0.21 ms | 0.12 ms | 1.14 | 1.14 | 1123.9 KB | 0.53 | 164.2 KB | 1.25 | 13.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | MiniExcel | 11.67 ms | 0.58 ms | 0.33 ms | 2.57 | 2.57 | 29746.9 KB | 13.90 | 180.5 KB | 1.38 | 157.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | ClosedXML | 67.27 ms | 2.47 ms | 1.43 ms | 14.84 | 14.84 | 27410.8 KB | 12.81 | 159.4 KB | 1.22 | 1384.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-direct-export | EPPlus | 67.79 ms | 9.55 ms | 5.51 ms | 14.96 | 14.96 | 21889.7 KB | 10.23 | 144.5 KB | 1.10 | 1395.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | OfficeIMO.Excel | 5.66 ms | 0.19 ms | 0.11 ms | 1.00 | 1.00 | 2880.2 KB | 1.00 | 176.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | MiniExcel | 11.50 ms | 0.68 ms | 0.40 ms | 2.03 | 2.03 | 29746.9 KB | 10.33 | 180.5 KB | 1.03 | 103.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | ClosedXML | 67.63 ms | 4.63 ms | 2.67 ms | 11.94 | 11.94 | 27410.3 KB | 9.52 | 159.4 KB | 0.91 | 1094.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-sparse-tables | EPPlus | 71.10 ms | 2.40 ms | 1.39 ms | 12.56 | 12.56 | 21889.7 KB | 7.60 | 144.5 KB | 0.82 | 1155.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables | OfficeIMO.Excel | 5.01 ms | 0.35 ms | 0.20 ms | 1.00 | 1.00 | 2066.1 KB | 1.00 | 138.9 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-dataset-tables | MiniExcel | 9.63 ms | 0.99 ms | 0.57 ms | 1.92 | 1.92 | 28700.4 KB | 13.89 | 156.4 KB | 1.13 | 92.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables | EPPlus | 41.45 ms | 1.02 ms | 0.59 ms | 8.28 | 8.28 | 18700.6 KB | 9.05 | 116.6 KB | 0.84 | 727.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables | ClosedXML | 43.41 ms | 1.41 ms | 0.81 ms | 8.67 | 8.67 | 18875.8 KB | 9.14 | 123.4 KB | 0.89 | 766.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | OfficeIMO.Excel | 5.08 ms | 0.06 ms | 0.03 ms | 1.00 | 1.00 | 2078.7 KB | 1.00 | 139.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | MiniExcel | 10.95 ms | 1.54 ms | 0.89 ms | 2.15 | 2.15 | 31798.6 KB | 15.30 | 156.6 KB | 1.13 | 115.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | EPPlus | 74.32 ms | 5.44 ms | 3.14 ms | 14.62 | 14.62 | 41455.7 KB | 19.94 | 116.9 KB | 0.84 | 1361.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-dataset-tables-autofit | ClosedXML | 95.69 ms | 4.10 ms | 2.36 ms | 18.82 | 18.82 | 56707.1 KB | 27.28 | 123.7 KB | 0.89 | 1781.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-direct | LargeXlsx | 3.99 ms | 0.32 ms | 0.19 ms | 0.82 | 1.00 | 1149.0 KB | 0.66 | 138.4 KB | 1.00 | 18.3% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-direct | OfficeIMO.Excel | 4.89 ms | 0.41 ms | 0.24 ms | 1.00 | 1.22 | 1748.6 KB | 1.00 | 138.0 KB | 1.00 | Loss +22.4% |
| 2500 | package-profile | package | Package size | write-datatable-direct | MiniExcel | 9.58 ms | 0.25 ms | 0.14 ms | 1.96 | 2.40 | 23062.5 KB | 13.19 | 153.7 KB | 1.11 | 96.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-direct | ClosedXML | 43.65 ms | 15.64 ms | 9.03 ms | 8.93 | 10.93 | 11581.0 KB | 6.62 | 120.1 KB | 0.87 | 792.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-direct | EPPlus | 44.21 ms | 1.23 ms | 0.71 ms | 9.04 | 11.07 | 16646.1 KB | 9.52 | 114.9 KB | 0.83 | 804.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | OfficeIMO.Excel | 4.74 ms | 0.80 ms | 0.46 ms | 1.00 | 1.00 | 1760.7 KB | 1.00 | 138.8 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | MiniExcel | 9.45 ms | 1.44 ms | 0.83 ms | 2.00 | 2.00 | 23062.8 KB | 13.10 | 153.7 KB | 1.11 | 99.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | ClosedXML | 40.97 ms | 1.51 ms | 0.87 ms | 8.65 | 8.65 | 19007.5 KB | 10.80 | 120.9 KB | 0.87 | 764.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-datatable-table-direct | EPPlus | 42.90 ms | 1.85 ms | 1.07 ms | 9.06 | 9.06 | 16646.1 KB | 9.45 | 114.9 KB | 0.83 | 805.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | LargeXlsx | 3.34 ms | 0.53 ms | 0.31 ms | 0.73 | 1.00 | 758.3 KB | 0.43 | 138.4 KB | 1.00 | 26.7% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 4.56 ms | 0.08 ms | 0.05 ms | 1.00 | 1.36 | 1769.2 KB | 1.00 | 138.0 KB | 1.00 | Loss +36.4% |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | MiniExcel | 10.34 ms | 0.77 ms | 0.45 ms | 2.27 | 3.09 | 23222.3 KB | 13.13 | 153.7 KB | 1.11 | 126.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | ClosedXML | 32.63 ms | 1.59 ms | 0.92 ms | 7.15 | 9.76 | 11581.0 KB | 6.55 | 120.1 KB | 0.87 | 615.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-fluent-rowsfrom-direct | EPPlus | 43.58 ms | 0.48 ms | 0.28 ms | 9.55 | 13.03 | 16646.4 KB | 9.41 | 114.9 KB | 0.83 | 855.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 5.36 ms | 0.27 ms | 0.15 ms | 1.00 | 1.00 | 1347.1 KB | 1.00 | 142.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 70.55 ms | 2.39 ms | 1.38 ms | 13.15 | 13.15 | 38343.9 KB | 28.46 | 115.1 KB | 0.81 | 1215.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 91.83 ms | 4.42 ms | 2.55 ms | 17.12 | 17.12 | 50927.5 KB | 37.80 | 120.2 KB | 0.84 | 1611.7% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | LargeXlsx | 4.02 ms | 0.64 ms | 0.37 ms | 0.86 | 1.00 | 758.3 KB | 0.57 | 138.4 KB | 0.97 | 14.2% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | OfficeIMO.Excel | 4.68 ms | 0.14 ms | 0.08 ms | 1.00 | 1.17 | 1339.3 KB | 1.00 | 142.3 KB | 1.00 | Loss +16.6% |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | MiniExcel | 10.79 ms | 1.26 ms | 0.73 ms | 2.30 | 2.69 | 23222.3 KB | 17.34 | 153.7 KB | 1.08 | 130.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | ClosedXML | 40.84 ms | 8.72 ms | 5.03 ms | 8.72 | 10.16 | 11581.0 KB | 8.65 | 120.1 KB | 0.84 | 771.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-direct | EPPlus | 48.52 ms | 1.15 ms | 0.66 ms | 10.36 | 12.07 | 16646.1 KB | 12.43 | 114.9 KB | 0.81 | 935.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 6.49 ms | 0.99 ms | 0.57 ms | 1.00 | 1.00 | 1505.3 KB | 1.00 | 138.1 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 67.74 ms | 1.39 ms | 0.80 ms | 10.44 | 10.44 | 38343.9 KB | 25.47 | 115.1 KB | 0.83 | 943.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 83.81 ms | 3.44 ms | 1.99 ms | 12.91 | 12.91 | 50927.5 KB | 33.83 | 120.2 KB | 0.87 | 1191.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 3.94 ms | 0.10 ms | 0.06 ms | 0.69 | 1.00 | 758.3 KB | 0.51 | 138.4 KB | 1.00 | 31.0% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 5.71 ms | 0.93 ms | 0.54 ms | 1.00 | 1.45 | 1497.5 KB | 1.00 | 138.0 KB | 1.00 | Loss +45.0% |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | MiniExcel | 9.88 ms | 0.08 ms | 0.04 ms | 1.73 | 2.51 | 23222.3 KB | 15.51 | 153.7 KB | 1.11 | 73.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | ClosedXML | 35.22 ms | 1.43 ms | 0.82 ms | 6.17 | 8.94 | 11581.0 KB | 7.73 | 120.1 KB | 0.87 | 516.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | EPPlus | 44.96 ms | 2.09 ms | 1.21 ms | 7.87 | 11.41 | 16646.1 KB | 11.12 | 114.9 KB | 0.83 | 687.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | LargeXlsx | 3.95 ms | 0.14 ms | 0.08 ms | 0.65 | 1.00 | 758.3 KB | 0.43 | 138.4 KB | 0.97 | 34.8% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | OfficeIMO.Excel | 6.06 ms | 0.87 ms | 0.50 ms | 1.00 | 1.53 | 1770.1 KB | 1.00 | 142.3 KB | 1.00 | Loss +53.3% |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | MiniExcel | 10.60 ms | 0.67 ms | 0.38 ms | 1.75 | 2.68 | 23222.3 KB | 13.12 | 153.7 KB | 1.08 | 74.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | ClosedXML | 36.87 ms | 0.57 ms | 0.33 ms | 6.09 | 9.33 | 11581.0 KB | 6.54 | 120.1 KB | 0.84 | 508.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | EPPlus | 43.50 ms | 0.57 ms | 0.33 ms | 7.18 | 11.00 | 16646.1 KB | 9.40 | 114.9 KB | 0.81 | 617.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.93 ms | 0.21 ms | 0.12 ms | 1.00 | 1.00 | 1346.4 KB | 1.00 | 142.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 58.21 ms | 7.04 ms | 4.07 ms | 11.80 | 11.80 | 28540.6 KB | 21.20 | 120.2 KB | 0.84 | 1080.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 61.51 ms | 1.14 ms | 0.66 ms | 12.47 | 12.47 | 27305.8 KB | 20.28 | 115.0 KB | 0.81 | 1147.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | LargeXlsx | 5.31 ms | 0.99 ms | 0.57 ms | 0.69 | 1.00 | 802.5 KB | 0.34 | 182.6 KB | 1.00 | 31.4% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 7.74 ms | 1.30 ms | 0.75 ms | 1.00 | 1.46 | 2341.7 KB | 1.00 | 183.1 KB | 1.00 | Loss +45.8% |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | MiniExcel | 10.69 ms | 0.47 ms | 0.27 ms | 1.38 | 2.01 | 25190.5 KB | 10.76 | 194.0 KB | 1.06 | 38.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | ClosedXML | 45.32 ms | 3.80 ms | 2.20 ms | 5.86 | 8.54 | 16973.5 KB | 7.25 | 161.0 KB | 0.88 | 485.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-mixed-objects-direct | EPPlus | 59.97 ms | 3.21 ms | 1.85 ms | 7.75 | 11.29 | 20105.1 KB | 8.59 | 152.1 KB | 0.83 | 674.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | LargeXlsx | 5.00 ms | 0.31 ms | 0.18 ms | 0.95 | 1.00 | 802.5 KB | 0.53 | 182.6 KB | 1.00 | 5.0% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 5.26 ms | 0.55 ms | 0.32 ms | 1.00 | 1.05 | 1507.7 KB | 1.00 | 182.4 KB | 1.00 | Loss +5.2% |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | MiniExcel | 10.38 ms | 1.79 ms | 1.03 ms | 1.97 | 2.07 | 25190.5 KB | 16.71 | 194.0 KB | 1.06 | 97.2% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | ClosedXML | 45.53 ms | 4.32 ms | 2.49 ms | 8.65 | 9.10 | 16973.5 KB | 11.26 | 161.0 KB | 0.88 | 764.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | EPPlus | 56.18 ms | 3.51 ms | 2.02 ms | 10.67 | 11.23 | 20105.1 KB | 13.33 | 152.1 KB | 0.83 | 967.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | LargeXlsx | 23.92 ms | 2.16 ms | 1.25 ms | 0.93 | 1.00 | 2810.7 KB | 0.62 | 644.6 KB | 0.99 | 7.0% faster than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 25.72 ms | 2.23 ms | 1.29 ms | 1.00 | 1.08 | 4502.3 KB | 1.00 | 651.0 KB | 1.00 | Loss +7.5% |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | MiniExcel | 41.68 ms | 2.32 ms | 1.34 ms | 1.62 | 1.74 | 48414.8 KB | 10.75 | 674.4 KB | 1.04 | 62.0% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | ClosedXML | 141.20 ms | 5.00 ms | 2.89 ms | 5.49 | 5.90 | 51647.0 KB | 11.47 | 615.5 KB | 0.95 | 448.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | write-powershell-psobject-wide-direct | EPPlus | 182.54 ms | 6.21 ms | 3.58 ms | 7.10 | 7.63 | 69139.6 KB | 15.36 | 548.9 KB | 0.84 | 609.6% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | OfficeIMO.Excel | 9.68 ms | 0.21 ms | 0.12 ms | 1.00 | 1.00 | 1895.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | EPPlus | 104.81 ms | 4.97 ms | 2.87 ms | 10.83 | 10.83 | 50712.0 KB | 26.76 |  |  | 983.2% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | EPPlus 4.5.3.3 | 112.86 ms |  |  | 11.66 | 11.66 |  |  |  |  | 1066.4% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | ClosedXML | 181.55 ms | 10.10 ms | 5.83 ms | 18.76 | 18.76 | 84605.7 KB | 44.64 |  |  | 1776.3% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | OfficeIMO.Excel | 13.55 ms | 1.24 ms | 0.72 ms | 1.00 | 1.00 | 6190.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus 4.5.3.3 | 39.42 ms |  |  | 2.91 | 2.91 |  |  |  |  | 190.9% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus | 54.57 ms | 3.57 ms | 2.06 ms | 4.03 | 4.03 | 32906.8 KB | 5.32 |  |  | 302.6% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | OfficeIMO.Excel | 13.43 ms | 1.00 ms | 0.58 ms | 1.00 | 1.00 | 6197.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus 4.5.3.3 | 81.33 ms |  |  | 6.06 | 6.06 |  |  |  |  | 505.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus | 88.79 ms | 4.35 ms | 2.51 ms | 6.61 | 6.61 | 54594.0 KB | 8.81 |  |  | 561.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | OfficeIMO.Excel | 14.14 ms | 0.73 ms | 0.42 ms | 1.00 | 1.00 | 6390.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus 4.5.3.3 | 88.80 ms |  |  | 6.28 | 6.28 |  |  |  |  | 527.8% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus | 89.14 ms | 1.46 ms | 0.85 ms | 6.30 | 6.30 | 59226.4 KB | 9.27 |  |  | 530.2% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | OfficeIMO.Excel | 14.68 ms | 0.62 ms | 0.36 ms | 1.00 | 1.00 | 6208.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus 4.5.3.3 | 78.95 ms |  |  | 5.38 | 5.38 |  |  |  |  | 437.8% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus | 85.11 ms | 0.28 ms | 0.16 ms | 5.80 | 5.80 | 54594.7 KB | 8.79 |  |  | 479.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | OfficeIMO.Excel | 15.82 ms | 0.65 ms | 0.38 ms | 1.00 | 1.00 | 6204.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus 4.5.3.3 | 81.45 ms |  |  | 5.15 | 5.15 |  |  |  |  | 414.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus | 84.62 ms | 0.61 ms | 0.35 ms | 5.35 | 5.35 | 54591.3 KB | 8.80 |  |  | 434.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Object projection | build-object-datatable-dictionaries | OfficeIMO.Excel | 1.67 ms | 0.40 ms | 0.23 ms | 1.00 | 1.00 | 564.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Object projection | build-object-datatable-typed | OfficeIMO.Excel | 1.60 ms | 0.13 ms | 0.08 ms | 1.00 | 1.00 | 856.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-cells | OfficeIMO.Excel | 7.95 ms | 0.52 ms | 0.30 ms | 1.00 | 1.00 | 2531.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-cells | ClosedXML | 40.45 ms | 1.12 ms | 0.65 ms | 5.09 | 5.09 | 17022.3 KB | 6.72 |  |  | 409.0% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-cells | EPPlus | 40.75 ms | 6.37 ms | 3.68 ms | 5.13 | 5.13 | 20154.9 KB | 7.96 |  |  | 412.8% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | OfficeIMO.Excel | 5.17 ms | 0.06 ms | 0.04 ms | 1.00 | 1.00 | 523.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | EPPlus | 32.06 ms | 1.30 ms | 0.75 ms | 6.20 | 6.20 | 13108.1 KB | 25.04 |  |  | 520.1% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | ClosedXML | 44.98 ms | 1.46 ms | 0.84 ms | 8.70 | 8.70 | 15463.3 KB | 29.54 |  |  | 770.1% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-range | OfficeIMO.Excel | 7.62 ms | 1.94 ms | 1.12 ms | 1.00 | 1.00 | 2531.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-range | EPPlus | 37.58 ms | 5.57 ms | 3.22 ms | 4.93 | 4.93 | 20154.9 KB | 7.96 |  |  | 393.4% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-range | ClosedXML | 38.81 ms | 0.55 ms | 0.32 ms | 5.09 | 5.09 | 17020.5 KB | 6.72 |  |  | 409.5% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-top-range | OfficeIMO.Excel | 0.77 ms | 0.06 ms | 0.04 ms | 1.00 | 1.00 | 285.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Range and table read | enumerate-top-range | EPPlus | 28.35 ms | 3.28 ms | 1.89 ms | 36.92 | 36.92 | 12404.4 KB | 43.45 |  |  | 3592.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Range and table read | enumerate-top-range | ClosedXML | 37.77 ms | 1.18 ms | 0.68 ms | 49.20 | 49.20 | 15370.3 KB | 53.84 |  |  | 4819.5% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-autofilter | OfficeIMO.Excel | 4.26 ms | 0.12 ms | 0.07 ms | 1.00 | 1.00 | 1340.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-autofilter | EPPlus 4.5.3.3 | 28.49 ms |  |  | 6.68 | 6.68 |  |  |  |  | 568.1% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-autofilter | ClosedXML | 42.14 ms | 8.20 ms | 4.74 ms | 9.88 | 9.88 | 22226.8 KB | 16.58 |  |  | 888.2% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-autofilter | EPPlus | 50.65 ms | 4.04 ms | 2.33 ms | 11.88 | 11.88 | 24715.5 KB | 18.44 |  |  | 1087.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-charts | OfficeIMO.Excel | 6.57 ms | 0.44 ms | 0.26 ms | 1.00 | 1.00 | 1891.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-charts | EPPlus 4.5.3.3 | 34.29 ms |  |  | 5.22 | 5.22 |  |  |  |  | 422.1% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-charts | EPPlus | 54.79 ms | 1.79 ms | 1.03 ms | 8.34 | 8.34 | 27142.3 KB | 14.35 |  |  | 734.2% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | OfficeIMO.Excel | 4.53 ms | 0.59 ms | 0.34 ms | 1.00 | 1.00 | 1405.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | EPPlus 4.5.3.3 | 30.75 ms |  |  | 6.79 | 6.79 |  |  |  |  | 578.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | ClosedXML | 37.66 ms | 1.92 ms | 1.11 ms | 8.31 | 8.31 | 22273.8 KB | 15.84 |  |  | 731.1% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | EPPlus | 44.53 ms | 1.00 ms | 0.58 ms | 9.83 | 9.83 | 24757.5 KB | 17.61 |  |  | 882.8% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-data-validation | OfficeIMO.Excel | 4.76 ms | 0.32 ms | 0.19 ms | 1.00 | 1.00 | 1356.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-data-validation | EPPlus 4.5.3.3 | 28.70 ms |  |  | 6.03 | 6.03 |  |  |  |  | 502.6% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-data-validation | ClosedXML | 40.88 ms | 3.01 ms | 1.74 ms | 8.58 | 8.58 | 22247.9 KB | 16.41 |  |  | 758.2% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-data-validation | EPPlus | 47.64 ms | 0.82 ms | 0.47 ms | 10.00 | 10.00 | 24701.4 KB | 18.22 |  |  | 900.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | OfficeIMO.Excel | 4.25 ms | 0.40 ms | 0.23 ms | 1.00 | 1.00 | 1342.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | EPPlus 4.5.3.3 | 28.87 ms |  |  | 6.79 | 6.79 |  |  |  |  | 579.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | ClosedXML | 38.50 ms | 0.78 ms | 0.45 ms | 9.06 | 9.06 | 22222.0 KB | 16.55 |  |  | 806.0% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | EPPlus | 46.20 ms | 2.55 ms | 1.47 ms | 10.87 | 10.87 | 24730.0 KB | 18.42 |  |  | 987.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | OfficeIMO.Excel | 10.98 ms | 0.69 ms | 0.40 ms | 1.00 | 1.00 | 5494.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | EPPlus 4.5.3.3 | 33.03 ms |  |  | 3.01 | 3.01 |  |  |  |  | 200.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | EPPlus | 55.47 ms | 1.23 ms | 0.71 ms | 5.05 | 5.05 | 29537.4 KB | 5.38 |  |  | 405.4% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world report | realworld-report-all-in-one | OfficeIMO.Excel | 13.87 ms | 1.79 ms | 1.03 ms | 1.00 | 1.00 | 6198.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus | 85.20 ms | 5.79 ms | 3.34 ms | 6.14 | 6.14 | 54594.8 KB | 8.81 |  |  | 514.2% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus 4.5.3.3 | 85.69 ms |  |  | 6.18 | 6.18 |  |  |  |  | 517.8% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world report | realworld-report-core | OfficeIMO.Excel | 5.27 ms | 0.92 ms | 0.53 ms | 1.00 | 1.00 | 1488.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world report | realworld-report-core | EPPlus | 78.53 ms | 2.09 ms | 1.21 ms | 14.91 | 14.91 | 47299.8 KB | 31.77 |  |  | 1390.6% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world report | realworld-report-core | EPPlus 4.5.3.3 | 85.01 ms |  |  | 16.14 | 16.14 |  |  |  |  | 1513.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world report | realworld-report-core | ClosedXML | 92.13 ms | 4.02 ms | 2.32 ms | 17.49 | 17.49 | 69835.1 KB | 46.91 |  |  | 1648.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook | OfficeIMO.Excel | 26.10 ms | 1.19 ms | 0.69 ms | 1.00 | 1.00 | 7234.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Report workbook | report-workbook | EPPlus 4.5.3.3 | 108.50 ms |  |  | 4.16 | 4.16 |  |  |  |  | 315.6% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook | EPPlus | 124.26 ms | 5.83 ms | 3.36 ms | 4.76 | 4.76 | 77486.1 KB | 10.71 |  |  | 376.0% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-core | OfficeIMO.Excel | 7.48 ms | 0.63 ms | 0.36 ms | 1.00 | 1.00 | 2711.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Report workbook | report-workbook-core | EPPlus 4.5.3.3 | 103.99 ms |  |  | 13.90 | 13.90 |  |  |  |  | 1290.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-core | EPPlus | 119.77 ms | 3.92 ms | 2.26 ms | 16.01 | 16.01 | 71970.6 KB | 26.55 |  |  | 1501.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-core | ClosedXML | 123.93 ms | 7.36 ms | 4.25 ms | 16.57 | 16.57 | 97217.3 KB | 35.86 |  |  | 1556.8% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable | OfficeIMO.Excel | 18.81 ms | 0.34 ms | 0.19 ms | 1.00 | 1.00 | 7548.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable | EPPlus 4.5.3.3 | 102.31 ms |  |  | 5.44 | 5.44 |  |  |  |  | 443.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable | EPPlus | 125.61 ms | 6.54 ms | 3.77 ms | 6.68 | 6.68 | 65995.1 KB | 8.74 |  |  | 567.8% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable-core | OfficeIMO.Excel | 7.58 ms | 0.57 ms | 0.33 ms | 1.00 | 1.00 | 2982.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable-core | EPPlus 4.5.3.3 | 100.61 ms |  |  | 13.28 | 13.28 |  |  |  |  | 1227.8% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable-core | EPPlus | 114.61 ms | 7.51 ms | 4.34 ms | 15.13 | 15.13 | 60480.1 KB | 20.28 |  |  | 1412.6% slower than OfficeIMO |
| 2500 | speed-comparison | other | Report workbook | report-workbook-datatable-core | ClosedXML | 119.46 ms | 6.01 ms | 3.47 ms | 15.77 | 15.77 | 82860.0 KB | 27.78 |  |  | 1476.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | OfficeIMO.Excel | 3.23 ms | 0.22 ms | 0.13 ms | 1.00 | 1.00 | 706.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | EPPlus 4.5.3.3 | 17.95 ms |  |  | 5.56 | 5.56 |  |  |  |  | 456.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | EPPlus | 23.17 ms | 0.12 ms | 0.07 ms | 7.18 | 7.18 | 7708.0 KB | 10.91 |  |  | 618.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Formula write/read | formula-heavy-read | ClosedXML | 25.61 ms | 0.61 ms | 0.35 ms | 7.94 | 7.94 | 8279.4 KB | 11.72 |  |  | 693.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | OfficeIMO.Excel | 1.33 ms | 0.13 ms | 0.08 ms | 1.00 | 1.00 | 177.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | Sylvan.Data.Excel | 1.48 ms | 0.16 ms | 0.09 ms | 1.11 | 1.11 | 316.6 KB | 1.79 |  |  | 10.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | ExcelDataReader | 2.49 ms | 0.16 ms | 0.09 ms | 1.87 | 1.87 | 4062.2 KB | 22.92 |  |  | 87.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | ClosedXML | 4.79 ms | 0.41 ms | 0.24 ms | 3.59 | 3.59 | 4392.6 KB | 24.78 |  |  | 259.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | EPPlus 4.5.3.3 | 10.56 ms |  |  | 7.92 | 7.92 |  |  |  |  | 692.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | MiniExcel | 17.43 ms | 3.15 ms | 1.82 ms | 13.08 | 13.08 | 46194.9 KB | 260.60 |  |  | 1207.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-column-read | EPPlus | 48.44 ms | 4.96 ms | 2.86 ms | 36.34 | 36.34 | 43071.0 KB | 242.98 |  |  | 3534.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | Sylvan.Data.Excel | 1.61 ms | 0.36 ms | 0.21 ms | 0.69 | 1.00 | 316.6 KB | 1.79 |  |  | 30.8% faster than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | OfficeIMO.Excel | 2.33 ms | 0.91 ms | 0.52 ms | 1.00 | 1.44 | 177.3 KB | 1.00 |  |  | Loss +44.4% |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | ExcelDataReader | 2.73 ms | 0.40 ms | 0.23 ms | 1.17 | 1.69 | 4062.2 KB | 22.91 |  |  | 17.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | ClosedXML | 4.32 ms | 0.19 ms | 0.11 ms | 1.86 | 2.68 | 4392.6 KB | 24.77 |  |  | 85.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | EPPlus 4.5.3.3 | 13.49 ms |  |  | 5.79 | 8.36 |  |  |  |  | 478.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | MiniExcel | 14.22 ms | 1.74 ms | 1.01 ms | 6.10 | 8.81 | 46194.9 KB | 260.50 |  |  | 510.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | large-sparse-row-read | EPPlus | 47.28 ms | 3.92 ms | 2.26 ms | 20.28 | 29.29 | 43071.0 KB | 242.89 |  |  | 1928.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | Sylvan.Data.Excel | 2.41 ms | 0.58 ms | 0.33 ms | 0.62 | 1.00 | 518.6 KB | 0.49 |  |  | 38.0% faster than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | OfficeIMO.Excel | 3.89 ms | 0.39 ms | 0.22 ms | 1.00 | 1.61 | 1056.5 KB | 1.00 |  |  | Loss +61.4% |
| 2500 | speed-comparison | read | Other | shared-string-read | ExcelDataReader | 6.73 ms | 0.65 ms | 0.37 ms | 1.73 | 2.79 | 2619.0 KB | 2.48 |  |  | 73.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | MiniExcel | 7.48 ms | 0.35 ms | 0.20 ms | 1.92 | 3.10 | 7530.0 KB | 7.13 |  |  | 92.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | ClosedXML | 20.53 ms | 2.29 ms | 1.32 ms | 5.27 | 8.51 | 9497.7 KB | 8.99 |  |  | 427.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | EPPlus 4.5.3.3 | 20.53 ms |  |  | 5.27 | 8.51 |  |  |  |  | 427.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Other | shared-string-read | EPPlus | 23.82 ms | 1.14 ms | 0.66 ms | 6.12 | 9.87 | 10372.2 KB | 9.82 |  |  | 511.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | Sylvan.Data.Excel | 5.13 ms | 1.11 ms | 0.64 ms | 0.81 | 1.00 | 655.2 KB | 1.75 |  |  | 18.8% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | OfficeIMO.Excel | 6.31 ms | 1.61 ms | 0.93 ms | 1.00 | 1.23 | 374.5 KB | 1.00 |  |  | Loss +23.2% |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | ExcelDataReader | 14.84 ms | 1.65 ms | 0.95 ms | 2.35 | 2.90 | 6089.5 KB | 16.26 |  |  | 135.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | MiniExcel | 17.45 ms | 1.23 ms | 0.71 ms | 2.76 | 3.41 | 18661.8 KB | 49.83 |  |  | 176.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | EPPlus | 31.90 ms | 1.75 ms | 1.01 ms | 5.05 | 6.23 | 12427.1 KB | 33.18 |  |  | 405.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-bottom-range | ClosedXML | 40.72 ms | 2.99 ms | 1.72 ms | 6.45 | 7.94 | 15360.9 KB | 41.02 |  |  | 544.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | Sylvan.Data.Excel | 10.99 ms | 3.35 ms | 1.94 ms | 0.51 | 1.00 | 2239.3 KB | 0.62 |  |  | 48.9% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | ExcelDataReader | 16.89 ms | 1.75 ms | 1.01 ms | 0.79 | 1.54 | 7673.5 KB | 2.13 |  |  | 21.5% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | MiniExcel | 19.08 ms | 2.56 ms | 1.48 ms | 0.89 | 1.74 | 18266.6 KB | 5.08 |  |  | 11.3% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | OfficeIMO.Excel | 21.50 ms | 18.77 ms | 10.84 ms | 1.00 | 1.96 | 3594.5 KB | 1.00 |  |  | Loss +95.7% |
| 2500 | speed-comparison | read | Range and table read | read-datatable | EPPlus | 44.57 ms | 6.25 ms | 3.61 ms | 2.07 | 4.06 | 21736.6 KB | 6.05 |  |  | 107.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | EPPlus 4.5.3.3 | 45.30 ms |  |  | 2.11 | 4.12 |  |  |  |  | 110.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-datatable | ClosedXML | 46.27 ms | 7.56 ms | 4.37 ms | 2.15 | 4.21 | 18314.1 KB | 5.10 |  |  | 115.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | OfficeIMO.Excel | 5.97 ms | 0.50 ms | 0.29 ms | 1.00 | 1.00 | 543.0 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | Sylvan.Data.Excel | 6.47 ms | 0.10 ms | 0.06 ms | 1.08 | 1.08 | 733.5 KB | 1.35 |  |  | 8.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | MiniExcel | 14.59 ms | 0.79 ms | 0.45 ms | 2.44 | 2.44 | 15850.3 KB | 29.19 |  |  | 144.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | ExcelDataReader | 15.25 ms | 0.73 ms | 0.42 ms | 2.55 | 2.55 | 6089.5 KB | 11.21 |  |  | 155.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | EPPlus | 32.44 ms | 3.45 ms | 1.99 ms | 5.43 | 5.43 | 13108.1 KB | 24.14 |  |  | 443.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | ClosedXML | 43.22 ms | 1.93 ms | 1.11 ms | 7.23 | 7.23 | 15465.0 KB | 28.48 |  |  | 623.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | OfficeIMO.Excel | 11.52 ms | 0.79 ms | 0.46 ms | 1.00 | 1.00 | 2692.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Range and table read | read-range | Sylvan.Data.Excel | 15.83 ms | 6.88 ms | 3.97 ms | 1.37 | 1.37 | 655.0 KB | 0.24 |  |  | 37.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | MiniExcel | 27.33 ms | 3.56 ms | 2.05 ms | 2.37 | 2.37 | 18662.2 KB | 6.93 |  |  | 137.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | ExcelDataReader | 29.17 ms | 12.29 ms | 7.10 ms | 2.53 | 2.53 | 6089.2 KB | 2.26 |  |  | 153.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | EPPlus 4.5.3.3 | 41.98 ms |  |  | 3.65 | 3.65 |  |  |  |  | 264.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | EPPlus | 42.16 ms | 7.02 ms | 4.05 ms | 3.66 | 3.66 | 20152.6 KB | 7.48 |  |  | 266.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range | ClosedXML | 92.11 ms | 19.74 ms | 11.40 ms | 8.00 | 8.00 | 16846.3 KB | 6.26 |  |  | 699.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | OfficeIMO.Excel | 6.46 ms | 0.55 ms | 0.32 ms | 1.00 | 1.00 | 2751.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | Sylvan.Data.Excel | 7.72 ms | 0.86 ms | 0.50 ms | 1.20 | 1.20 | 750.3 KB | 0.27 |  |  | 19.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | ExcelDataReader | 14.77 ms | 2.65 ms | 1.53 ms | 2.29 | 2.29 | 6089.5 KB | 2.21 |  |  | 128.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | MiniExcel | 16.22 ms | 2.05 ms | 1.19 ms | 2.51 | 2.51 | 18662.4 KB | 6.78 |  |  | 151.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | EPPlus | 33.72 ms | 3.57 ms | 2.06 ms | 5.22 | 5.22 | 20152.6 KB | 7.33 |  |  | 421.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-range-decimal | ClosedXML | 38.28 ms | 2.77 ms | 1.60 ms | 5.93 | 5.93 | 16728.2 KB | 6.08 |  |  | 492.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | Sylvan.Data.Excel | 0.60 ms | 0.13 ms | 0.07 ms | 0.78 | 1.00 | 348.5 KB | 1.18 |  |  | 21.8% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | OfficeIMO.Excel | 0.77 ms | 0.12 ms | 0.07 ms | 1.00 | 1.28 | 296.0 KB | 1.00 |  |  | Loss +27.8% |
| 2500 | speed-comparison | read | Range and table read | read-top-range | MiniExcel | 0.99 ms | 0.13 ms | 0.08 ms | 1.29 | 1.65 | 869.0 KB | 2.94 |  |  | 29.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | ExcelDataReader | 6.38 ms | 0.75 ms | 0.43 ms | 8.28 | 10.58 | 1931.8 KB | 6.53 |  |  | 727.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | EPPlus | 33.20 ms | 0.90 ms | 0.52 ms | 43.10 | 55.08 | 12402.1 KB | 41.89 |  |  | 4209.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | EPPlus 4.5.3.3 | 35.67 ms |  |  | 46.30 | 59.18 |  |  |  |  | 4530.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-top-range | ClosedXML | 44.00 ms | 2.71 ms | 1.57 ms | 57.11 | 72.99 | 15360.2 KB | 51.89 |  |  | 5610.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | Sylvan.Data.Excel | 6.77 ms | 0.98 ms | 0.57 ms | 0.79 | 1.00 | 655.2 KB | 0.24 |  |  | 20.5% faster than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | OfficeIMO.Excel | 8.52 ms | 1.92 ms | 1.11 ms | 1.00 | 1.26 | 2750.3 KB | 1.00 |  |  | Loss +25.8% |
| 2500 | speed-comparison | read | Range and table read | read-used-range | ExcelDataReader | 14.77 ms | 2.03 ms | 1.17 ms | 1.73 | 2.18 | 6089.4 KB | 2.21 |  |  | 73.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | MiniExcel | 19.20 ms | 3.74 ms | 2.16 ms | 2.25 | 2.84 | 18662.4 KB | 6.79 |  |  | 125.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | EPPlus | 36.49 ms | 5.62 ms | 3.25 ms | 4.28 | 5.39 | 20152.7 KB | 7.33 |  |  | 328.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Range and table read | read-used-range | ClosedXML | 68.92 ms | 38.36 ms | 22.15 ms | 8.09 | 10.18 | 16806.7 KB | 6.11 |  |  | 709.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | OfficeIMO.Excel | 5.85 ms | 1.17 ms | 0.67 ms | 1.00 | 1.00 | 377.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | Sylvan.Data.Excel | 6.08 ms | 0.13 ms | 0.08 ms | 1.04 | 1.04 | 655.2 KB | 1.73 |  |  | 4.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | ExcelDataReader | 17.06 ms | 0.23 ms | 0.13 ms | 2.92 | 2.92 | 6089.5 KB | 16.12 |  |  | 191.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | MiniExcel | 20.60 ms | 1.05 ms | 0.61 ms | 3.52 | 3.52 | 18661.8 KB | 49.39 |  |  | 252.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | EPPlus | 35.11 ms | 0.48 ms | 0.28 ms | 6.01 | 6.01 | 12427.1 KB | 32.89 |  |  | 500.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-bottom-range-stream | ClosedXML | 44.83 ms | 3.98 ms | 2.30 ms | 7.67 | 7.67 | 15359.3 KB | 40.65 |  |  | 666.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | Sylvan.Data.Excel | 6.84 ms | 0.88 ms | 0.51 ms | 0.41 | 1.00 | 655.2 KB | 0.24 |  |  | 58.8% faster than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | OfficeIMO.Excel | 16.61 ms | 6.52 ms | 3.76 ms | 1.00 | 2.43 | 2771.4 KB | 1.00 |  |  | Loss +142.7% |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | ExcelDataReader | 17.35 ms | 1.57 ms | 0.91 ms | 1.04 | 2.53 | 6089.5 KB | 2.20 |  |  | 4.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | MiniExcel | 19.48 ms | 0.30 ms | 0.17 ms | 1.17 | 2.85 | 18662.4 KB | 6.73 |  |  | 17.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | EPPlus 4.5.3.3 | 38.28 ms |  |  | 2.31 | 5.59 |  |  |  |  | 130.5% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | EPPlus | 45.78 ms | 8.33 ms | 4.81 ms | 2.76 | 6.69 | 20152.6 KB | 7.27 |  |  | 175.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-range-stream | ClosedXML | 46.72 ms | 1.33 ms | 0.77 ms | 2.81 | 6.83 | 16729.3 KB | 6.04 |  |  | 181.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | Sylvan.Data.Excel | 0.66 ms | 0.07 ms | 0.04 ms | 0.95 | 1.00 | 348.5 KB | 1.16 |  |  | 5.4% faster than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | OfficeIMO.Excel | 0.69 ms | 0.03 ms | 0.02 ms | 1.00 | 1.06 | 299.4 KB | 1.00 |  |  | Loss +5.7% |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | MiniExcel | 1.21 ms | 0.36 ms | 0.21 ms | 1.75 | 1.84 | 869.0 KB | 2.90 |  |  | 74.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | ExcelDataReader | 6.59 ms | 1.13 ms | 0.65 ms | 9.50 | 10.04 | 1931.8 KB | 6.45 |  |  | 850.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | EPPlus 4.5.3.3 | 33.54 ms |  |  | 48.33 | 51.08 |  |  |  |  | 4733.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | EPPlus | 35.07 ms | 3.19 ms | 1.84 ms | 50.54 | 53.41 | 12402.1 KB | 41.42 |  |  | 4954.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream | ClosedXML | 47.34 ms | 2.30 ms | 1.33 ms | 68.22 | 72.09 | 15360.8 KB | 51.30 |  |  | 6722.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | Sylvan.Data.Excel | 0.56 ms | 0.05 ms | 0.03 ms | 0.79 | 1.00 | 348.5 KB | 1.16 |  |  | 21.2% faster than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | OfficeIMO.Excel | 0.72 ms | 0.08 ms | 0.04 ms | 1.00 | 1.27 | 300.2 KB | 1.00 |  |  | Loss +26.9% |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | MiniExcel | 1.01 ms | 0.08 ms | 0.04 ms | 1.42 | 1.80 | 869.0 KB | 2.89 |  |  | 41.8% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | ExcelDataReader | 6.03 ms | 0.36 ms | 0.21 ms | 8.43 | 10.70 | 1931.8 KB | 6.44 |  |  | 742.9% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | EPPlus | 31.18 ms | 2.93 ms | 1.69 ms | 43.58 | 55.31 | 12402.1 KB | 41.32 |  |  | 4258.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | ClosedXML | 45.49 ms | 6.53 ms | 3.77 ms | 63.59 | 80.70 | 15360.4 KB | 51.17 |  |  | 6259.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | Sylvan.Data.Excel | 8.37 ms | 4.16 ms | 2.40 ms | 0.87 | 1.00 | 895.3 KB | 0.37 |  |  | 13.1% faster than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | OfficeIMO.Excel | 9.62 ms | 3.34 ms | 1.93 ms | 1.00 | 1.15 | 2442.0 KB | 1.00 |  |  | Loss +15.0% |
| 2500 | speed-comparison | read | Typed object read | read-objects | ExcelDataReader | 16.84 ms | 3.66 ms | 2.11 ms | 1.75 | 2.01 | 6329.5 KB | 2.59 |  |  | 75.0% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | MiniExcel | 19.73 ms | 1.80 ms | 1.04 ms | 2.05 | 2.36 | 18473.9 KB | 7.56 |  |  | 105.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | EPPlus 4.5.3.3 | 39.04 ms |  |  | 4.06 | 4.67 |  |  |  |  | 305.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | ClosedXML | 41.15 ms | 3.70 ms | 2.14 ms | 4.28 | 4.92 | 16925.3 KB | 6.93 |  |  | 327.6% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects | EPPlus | 41.83 ms | 8.83 ms | 5.10 ms | 4.35 | 5.00 | 21354.2 KB | 8.74 |  |  | 334.7% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | Sylvan.Data.Excel | 6.56 ms | 0.35 ms | 0.20 ms | 0.90 | 1.00 | 831.0 KB | 0.34 |  |  | 10.3% faster than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | OfficeIMO.Excel | 7.32 ms | 0.56 ms | 0.32 ms | 1.00 | 1.11 | 2422.9 KB | 1.00 |  |  | Loss +11.4% |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | ExcelDataReader | 16.47 ms | 0.67 ms | 0.39 ms | 2.25 | 2.51 | 6265.3 KB | 2.59 |  |  | 125.2% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | MiniExcel | 19.62 ms | 1.23 ms | 0.71 ms | 2.68 | 2.99 | 18409.7 KB | 7.60 |  |  | 168.3% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | EPPlus 4.5.3.3 | 37.46 ms |  |  | 5.12 | 5.71 |  |  |  |  | 412.1% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | EPPlus | 45.75 ms | 7.85 ms | 4.53 ms | 6.25 | 6.97 | 21334.6 KB | 8.81 |  |  | 525.4% slower than OfficeIMO |
| 2500 | speed-comparison | read | Typed object read | read-objects-stream | ClosedXML | 47.14 ms | 0.29 ms | 0.17 ms | 6.44 | 7.18 | 16903.8 KB | 6.98 |  |  | 544.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | OfficeIMO.Excel | 6.87 ms | 2.38 ms | 1.37 ms | 1.00 | 1.00 | 1781.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | MiniExcel | 8.40 ms | 0.51 ms | 0.30 ms | 1.22 | 1.22 | 26647.3 KB | 14.96 |  |  | 22.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | EPPlus | 54.87 ms | 4.20 ms | 2.43 ms | 7.98 | 7.98 | 38344.3 KB | 21.53 |  |  | 698.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | EPPlus 4.5.3.3 | 63.29 ms |  |  | 9.21 | 9.21 |  |  |  |  | 820.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | ClosedXML | 74.64 ms | 2.17 ms | 1.25 ms | 10.86 | 10.86 | 58360.0 KB | 32.77 |  |  | 985.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | OfficeIMO.Excel | 6.06 ms | 0.55 ms | 0.32 ms | 1.00 | 1.00 | 2079.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | MiniExcel | 16.98 ms | 1.23 ms | 0.71 ms | 2.80 | 2.80 | 32328.7 KB | 15.55 |  |  | 179.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | EPPlus | 144.10 ms | 6.79 ms | 3.92 ms | 23.76 | 23.76 | 43440.5 KB | 20.89 |  |  | 2276.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | ClosedXML | 146.69 ms | 18.92 ms | 10.93 ms | 24.19 | 24.19 | 56707.6 KB | 27.27 |  |  | 2319.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 5.95 ms | 1.59 ms | 0.92 ms | 1.00 | 1.00 | 1347.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 64.25 ms | 3.26 ms | 1.88 ms | 10.79 | 10.79 | 38344.1 KB | 28.46 |  |  | 979.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 82.94 ms | 9.53 ms | 5.50 ms | 13.93 | 13.93 | 50927.7 KB | 37.80 |  |  | 1292.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 5.93 ms | 0.21 ms | 0.12 ms | 1.00 | 1.00 | 1505.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 74.10 ms | 1.74 ms | 1.01 ms | 12.50 | 12.50 | 38344.1 KB | 25.47 |  |  | 1150.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 86.37 ms | 2.68 ms | 1.55 ms | 14.58 | 14.58 | 50927.3 KB | 33.83 |  |  | 1357.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 4.96 ms | 0.15 ms | 0.08 ms | 1.00 | 1.00 | 1346.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 55.27 ms | 0.41 ms | 0.24 ms | 11.14 | 11.14 | 28540.4 KB | 21.20 |  |  | 1013.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 60.14 ms | 3.08 ms | 1.78 ms | 12.12 | 12.12 | 27305.8 KB | 20.28 |  |  | 1111.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | OfficeIMO.Excel | 2.17 ms | 0.06 ms | 0.03 ms | 1.00 | 1.00 | 1787.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | ClosedXML | 11.88 ms | 0.76 ms | 0.44 ms | 5.46 | 5.46 | 9959.5 KB | 5.57 |  |  | 446.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | EPPlus | 20.60 ms | 1.70 ms | 0.98 ms | 9.48 | 9.48 | 11772.9 KB | 6.59 |  |  | 847.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | OfficeIMO.Excel | 2.01 ms | 0.02 ms | 0.01 ms | 1.00 | 1.00 | 1119.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | ClosedXML | 10.90 ms | 0.79 ms | 0.46 ms | 5.44 | 5.44 | 9177.1 KB | 8.19 |  |  | 443.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | EPPlus 4.5.3.3 | 13.09 ms |  |  | 6.53 | 6.53 |  |  |  |  | 552.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-numbers | EPPlus | 21.77 ms | 1.24 ms | 0.72 ms | 10.86 | 10.86 | 12895.2 KB | 11.51 |  |  | 986.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | OfficeIMO.Excel | 3.34 ms | 0.28 ms | 0.16 ms | 1.00 | 1.00 | 1763.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | EPPlus 4.5.3.3 | 15.61 ms |  |  | 4.67 | 4.67 |  |  |  |  | 367.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | ClosedXML | 15.63 ms | 0.29 ms | 0.17 ms | 4.68 | 4.68 | 11887.0 KB | 6.74 |  |  | 368.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | EPPlus | 28.99 ms | 4.18 ms | 2.41 ms | 8.68 | 8.68 | 15643.3 KB | 8.87 |  |  | 767.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | OfficeIMO.Excel | 3.47 ms | 0.42 ms | 0.24 ms | 1.00 | 1.00 | 1506.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | ClosedXML | 14.80 ms | 0.78 ms | 0.45 ms | 4.26 | 4.26 | 11296.3 KB | 7.50 |  |  | 326.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | EPPlus | 27.67 ms | 1.12 ms | 0.65 ms | 7.97 | 7.97 | 14960.2 KB | 9.93 |  |  | 697.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 3.65 ms | 0.74 ms | 0.43 ms | 1.00 | 1.00 | 1506.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | ClosedXML | 16.35 ms | 1.79 ms | 1.03 ms | 4.48 | 4.48 | 11296.3 KB | 7.50 |  |  | 348.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | EPPlus | 28.41 ms | 1.00 ms | 0.58 ms | 7.79 | 7.79 | 14960.2 KB | 9.93 |  |  | 679.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | OfficeIMO.Excel | 1.85 ms | 0.01 ms | 0.01 ms | 1.00 | 1.00 | 1138.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | ClosedXML | 10.26 ms | 0.20 ms | 0.12 ms | 5.53 | 5.53 | 9021.2 KB | 7.93 |  |  | 453.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | EPPlus 4.5.3.3 | 12.40 ms |  |  | 6.69 | 6.69 |  |  |  |  | 568.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-scalars | EPPlus | 22.34 ms | 1.05 ms | 0.61 ms | 12.05 | 12.05 | 12827.4 KB | 11.27 |  |  | 1104.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | OfficeIMO.Excel | 3.11 ms | 0.46 ms | 0.27 ms | 1.00 | 1.00 | 1435.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | EPPlus 4.5.3.3 | 15.51 ms |  |  | 4.99 | 4.99 |  |  |  |  | 399.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | ClosedXML | 15.59 ms | 2.49 ms | 1.44 ms | 5.02 | 5.02 | 9711.1 KB | 6.76 |  |  | 402.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Cell writer | write-cellvalue-temporal | EPPlus | 24.48 ms | 2.02 ms | 1.17 ms | 7.88 | 7.88 | 14722.6 KB | 10.26 |  |  | 688.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | OfficeIMO.Excel | 5.96 ms | 0.59 ms | 0.34 ms | 1.00 | 1.00 | 2064.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | MiniExcel | 15.52 ms | 1.13 ms | 0.65 ms | 2.60 | 2.60 | 29223.6 KB | 14.16 |  |  | 160.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | ClosedXML | 53.71 ms | 7.72 ms | 4.46 ms | 9.00 | 9.00 | 18913.3 KB | 9.16 |  |  | 800.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | EPPlus | 95.66 ms | 19.19 ms | 11.08 ms | 16.04 | 16.04 | 18414.6 KB | 8.92 |  |  | 1503.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | OfficeIMO.Excel | 7.23 ms | 0.06 ms | 0.03 ms | 1.00 | 1.00 | 2880.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | MiniExcel | 17.99 ms | 1.06 ms | 0.61 ms | 2.49 | 2.49 | 30510.5 KB | 10.59 |  |  | 148.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | ClosedXML | 62.65 ms | 0.82 ms | 0.47 ms | 8.66 | 8.66 | 27410.7 KB | 9.52 |  |  | 766.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | EPPlus | 75.41 ms | 7.25 ms | 4.19 ms | 10.43 | 10.43 | 22591.6 KB | 7.84 |  |  | 942.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | OfficeIMO.Excel | 6.40 ms | 0.49 ms | 0.28 ms | 1.00 | 1.00 | 2067.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | MiniExcel | 17.91 ms | 3.74 ms | 2.16 ms | 2.80 | 2.80 | 28700.3 KB | 13.88 |  |  | 179.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | EPPlus 4.5.3.3 | 50.19 ms |  |  | 7.84 | 7.84 |  |  |  |  | 684.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | EPPlus | 93.79 ms | 10.89 ms | 6.29 ms | 14.66 | 14.66 | 19431.0 KB | 9.40 |  |  | 1365.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataSet table export | write-dataset-tables | ClosedXML | 108.60 ms | 8.75 ms | 5.05 ms | 16.97 | 16.97 | 18878.2 KB | 9.13 |  |  | 1597.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | OfficeIMO.Excel | 5.22 ms | 1.45 ms | 0.84 ms | 1.00 | 1.00 | 1774.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | MiniExcel | 8.19 ms | 0.21 ms | 0.12 ms | 1.57 | 1.57 | 23044.2 KB | 12.98 |  |  | 56.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | ClosedXML | 38.75 ms | 1.32 ms | 0.76 ms | 7.42 | 7.42 | 19008.4 KB | 10.71 |  |  | 641.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | EPPlus 4.5.3.3 | 40.00 ms |  |  | 7.66 | 7.66 |  |  |  |  | 665.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datareader-table | EPPlus | 41.39 ms | 10.46 ms | 6.04 ms | 7.92 | 7.92 | 16647.3 KB | 9.38 |  |  | 692.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | OfficeIMO.Excel | 5.01 ms | 0.72 ms | 0.41 ms | 1.00 | 1.00 | 1748.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | LargeXlsx | 7.23 ms | 0.23 ms | 0.13 ms | 1.44 | 1.44 | 1149.0 KB | 0.66 |  |  | 44.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | MiniExcel | 10.41 ms | 0.32 ms | 0.19 ms | 2.08 | 2.08 | 23062.6 KB | 13.19 |  |  | 107.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | ClosedXML | 31.52 ms | 3.05 ms | 1.76 ms | 6.29 | 6.29 | 11581.0 KB | 6.62 |  |  | 528.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | EPPlus | 42.67 ms | 5.72 ms | 3.30 ms | 8.51 | 8.51 | 16648.4 KB | 9.52 |  |  | 750.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-direct | EPPlus 4.5.3.3 | 54.65 ms |  |  | 10.90 | 10.90 |  |  |  |  | 990.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | OfficeIMO.Excel | 4.57 ms | 0.49 ms | 0.28 ms | 1.00 | 1.00 | 1487.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | MiniExcel | 9.61 ms | 0.57 ms | 0.33 ms | 2.10 | 2.10 | 22789.5 KB | 15.32 |  |  | 110.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | ClosedXML | 40.64 ms | 2.56 ms | 1.48 ms | 8.90 | 8.90 | 18735.1 KB | 12.60 |  |  | 789.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | EPPlus | 44.57 ms | 8.75 ms | 5.05 ms | 9.76 | 9.76 | 16374.0 KB | 11.01 |  |  | 876.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | OfficeIMO.Excel | 4.70 ms | 0.53 ms | 0.31 ms | 1.00 | 1.00 | 1760.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | MiniExcel | 10.20 ms | 0.95 ms | 0.55 ms | 2.17 | 2.17 | 23062.9 KB | 13.10 |  |  | 117.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | EPPlus | 42.18 ms | 3.99 ms | 2.30 ms | 8.98 | 8.98 | 16647.3 KB | 9.46 |  |  | 797.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | ClosedXML | 48.36 ms | 9.54 ms | 5.51 ms | 10.29 | 10.29 | 19008.7 KB | 10.80 |  |  | 929.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | DataTable table export | write-datatable-table-direct | EPPlus 4.5.3.3 | 49.12 ms |  |  | 10.46 | 10.46 |  |  |  |  | 945.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | OfficeIMO.Excel | 7.20 ms | 0.68 ms | 0.39 ms | 1.00 | 1.00 | 1403.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | MiniExcel | 14.97 ms | 0.95 ms | 0.55 ms | 2.08 | 2.08 | 26825.1 KB | 19.12 |  |  | 107.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | EPPlus 4.5.3.3 | 106.75 ms |  |  | 14.82 | 14.82 |  |  |  |  | 1382.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | EPPlus | 132.23 ms | 9.55 ms | 5.51 ms | 18.36 | 18.36 | 49158.1 KB | 35.03 |  |  | 1735.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formatted report write | write-bulk-report | ClosedXML | 239.90 ms | 71.54 ms | 41.30 ms | 33.31 | 33.31 | 58350.2 KB | 41.58 |  |  | 3230.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | OfficeIMO.Excel | 4.33 ms | 0.33 ms | 0.19 ms | 1.00 | 1.00 | 1620.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | EPPlus 4.5.3.3 | 17.00 ms |  |  | 3.93 | 3.93 |  |  |  |  | 292.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | ClosedXML | 23.08 ms | 4.76 ms | 2.75 ms | 5.33 | 5.33 | 12039.8 KB | 7.43 |  |  | 433.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Formula write/read | write-cellformula | EPPlus | 52.62 ms | 3.61 ms | 2.09 ms | 12.15 | 12.15 | 18110.5 KB | 11.17 |  |  | 1115.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-dictionary-objects-table-direct | OfficeIMO.Excel | 6.69 ms | 0.95 ms | 0.55 ms | 1.00 | 1.00 | 2051.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | LargeXlsx | 5.36 ms | 0.11 ms | 0.06 ms | 0.77 | 1.00 | 802.5 KB | 0.34 |  |  | 23.2% faster than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 6.99 ms | 0.82 ms | 0.47 ms | 1.00 | 1.30 | 2341.7 KB | 1.00 |  |  | Loss +30.3% |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | MiniExcel | 11.77 ms | 0.82 ms | 0.47 ms | 1.69 | 2.20 | 25190.4 KB | 10.76 |  |  | 68.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | ClosedXML | 49.70 ms | 0.99 ms | 0.57 ms | 7.11 | 9.27 | 16973.5 KB | 7.25 |  |  | 611.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | EPPlus | 58.11 ms | 2.37 ms | 1.37 ms | 8.32 | 10.84 | 20105.1 KB | 8.59 |  |  | 731.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | LargeXlsx | 6.13 ms | 0.35 ms | 0.20 ms | 0.80 | 1.00 | 802.5 KB | 0.53 |  |  | 20.1% faster than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 7.68 ms | 0.71 ms | 0.41 ms | 1.00 | 1.25 | 1507.7 KB | 1.00 |  |  | Loss +25.2% |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | MiniExcel | 11.71 ms | 0.59 ms | 0.34 ms | 1.52 | 1.91 | 25190.4 KB | 16.71 |  |  | 52.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | ClosedXML | 49.52 ms | 4.09 ms | 2.36 ms | 6.45 | 8.08 | 16973.5 KB | 11.26 |  |  | 545.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | EPPlus | 57.50 ms | 3.03 ms | 1.75 ms | 7.49 | 9.38 | 20105.1 KB | 13.33 |  |  | 648.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 23.95 ms | 2.47 ms | 1.43 ms | 1.00 | 1.00 | 4502.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | LargeXlsx | 25.83 ms | 1.56 ms | 0.90 ms | 1.08 | 1.08 | 2810.7 KB | 0.62 |  |  | 7.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | MiniExcel | 43.78 ms | 2.24 ms | 1.29 ms | 1.83 | 1.83 | 48414.8 KB | 10.75 |  |  | 82.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | ClosedXML | 151.06 ms | 12.28 ms | 7.09 ms | 6.31 | 6.31 | 51647.0 KB | 11.47 |  |  | 530.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | EPPlus | 182.97 ms | 3.21 ms | 1.86 ms | 7.64 | 7.64 | 69139.6 KB | 15.36 |  |  | 663.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | LargeXlsx | 1.95 ms | 0.62 ms | 0.36 ms | 0.70 | 1.00 | 296.4 KB | 0.19 |  |  | 30.3% faster than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | OfficeIMO.Excel | 2.79 ms | 0.27 ms | 0.16 ms | 1.00 | 1.43 | 1576.3 KB | 1.00 |  |  | Loss +43.5% |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | MiniExcel | 5.35 ms | 0.64 ms | 0.37 ms | 1.92 | 2.75 | 19710.8 KB | 12.50 |  |  | 91.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | EPPlus 4.5.3.3 | 16.83 ms |  |  | 6.02 | 8.64 |  |  |  |  | 502.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | ClosedXML | 18.25 ms | 0.98 ms | 0.57 ms | 6.53 | 9.37 | 11197.4 KB | 7.10 |  |  | 553.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | append-plain-rows | EPPlus | 30.11 ms | 3.06 ms | 1.77 ms | 10.77 | 15.46 | 14365.2 KB | 9.11 |  |  | 977.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 1.24 ms | 0.09 ms | 0.05 ms | 0.79 | 1.00 | 447.0 KB | 0.41 |  |  | 21.3% faster than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 1.57 ms | 0.04 ms | 0.02 ms | 1.00 | 1.27 | 1092.0 KB | 1.00 |  |  | Loss +27.0% |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | ClosedXML | 12.07 ms | 0.77 ms | 0.44 ms | 7.67 | 9.74 | 10235.8 KB | 9.37 |  |  | 667.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | EPPlus | 23.40 ms | 2.17 ms | 1.25 ms | 14.87 | 18.88 | 13052.1 KB | 11.95 |  |  | 1386.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | LargeXlsx | 3.20 ms | 0.19 ms | 0.11 ms | 0.77 | 1.00 | 758.3 KB | 0.36 |  |  | 22.9% faster than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 4.14 ms | 0.20 ms | 0.12 ms | 1.00 | 1.30 | 2081.1 KB | 1.00 |  |  | Loss +29.7% |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | MiniExcel | 9.01 ms | 0.37 ms | 0.22 ms | 2.18 | 2.82 | 23221.9 KB | 11.16 |  |  | 117.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 31.77 ms |  |  | 7.67 | 9.94 |  |  |  |  | 667.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | ClosedXML | 32.28 ms | 2.25 ms | 1.30 ms | 7.79 | 10.10 | 22221.3 KB | 10.68 |  |  | 679.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | EPPlus | 43.19 ms | 7.47 ms | 4.32 ms | 10.43 | 13.52 | 24693.7 KB | 11.87 |  |  | 942.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | OfficeIMO.Excel | 2.27 ms | 0.10 ms | 0.06 ms | 1.00 | 1.00 | 1494.0 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | ClosedXML | 14.29 ms | 0.50 ms | 0.29 ms | 6.30 | 6.30 | 11296.3 KB | 7.56 |  |  | 530.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | EPPlus | 24.41 ms | 0.58 ms | 0.34 ms | 10.76 | 10.76 | 14960.0 KB | 10.01 |  |  | 976.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | Sylvan.Data.Excel | 3.92 ms | 0.37 ms | 0.21 ms | 0.88 | 1.00 | 758.6 KB | 0.43 |  |  | 12.3% faster than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | OfficeIMO.Excel | 4.47 ms | 0.35 ms | 0.20 ms | 1.00 | 1.14 | 1763.0 KB | 1.00 |  |  | Loss +14.0% |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | LargeXlsx | 8.05 ms | 0.19 ms | 0.11 ms | 1.80 | 2.05 | 1032.5 KB | 0.59 |  |  | 80.0% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | MiniExcel | 8.37 ms | 0.27 ms | 0.16 ms | 1.87 | 2.14 | 23043.9 KB | 13.07 |  |  | 87.3% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | ClosedXML | 26.10 ms | 0.40 ms | 0.23 ms | 5.84 | 6.66 | 11581.0 KB | 6.57 |  |  | 483.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | EPPlus 4.5.3.3 | 27.49 ms |  |  | 6.15 | 7.01 |  |  |  |  | 514.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-datareader-plain | EPPlus | 35.56 ms | 1.19 ms | 0.69 ms | 7.95 | 9.07 | 16646.5 KB | 9.44 |  |  | 695.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | OfficeIMO.Excel | 6.99 ms | 1.07 ms | 0.62 ms | 1.00 | 1.00 | 2140.6 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | LargeXlsx | 11.06 ms | 0.72 ms | 0.42 ms | 1.58 | 1.58 | 1123.9 KB | 0.53 |  |  | 58.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | MiniExcel | 16.74 ms | 4.24 ms | 2.45 ms | 2.39 | 2.39 | 30001.5 KB | 14.02 |  |  | 139.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | EPPlus | 62.15 ms | 8.99 ms | 5.19 ms | 8.89 | 8.89 | 21892.9 KB | 10.23 |  |  | 788.5% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | ClosedXML | 74.89 ms | 16.51 ms | 9.53 ms | 10.71 | 10.71 | 27410.6 KB | 12.80 |  |  | 970.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | LargeXlsx | 4.22 ms | 0.26 ms | 0.15 ms | 0.94 | 1.00 | 857.6 KB | 0.51 |  |  | 6.4% faster than OfficeIMO |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | OfficeIMO.Excel | 4.51 ms | 0.09 ms | 0.05 ms | 1.00 | 1.07 | 1676.8 KB | 1.00 |  |  | Loss +6.9% |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | MiniExcel | 16.08 ms | 0.70 ms | 0.40 ms | 3.57 | 3.81 | 35918.0 KB | 21.42 |  |  | 256.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | ClosedXML | 82.94 ms | 0.53 ms | 0.30 ms | 18.40 | 19.67 | 71478.2 KB | 42.63 |  |  | 1740.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | OfficeIMO.Excel | 2.99 ms | 0.22 ms | 0.13 ms | 1.00 | 1.00 | 2440.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | MiniExcel | 5.00 ms | 0.73 ms | 0.42 ms | 1.67 | 1.67 | 21137.5 KB | 8.66 |  |  | 67.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | EPPlus 4.5.3.3 | 12.65 ms |  |  | 4.23 | 4.23 |  |  |  |  | 323.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | ClosedXML | 16.12 ms | 2.10 ms | 1.21 ms | 5.39 | 5.39 | 11299.2 KB | 4.63 |  |  | 439.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | large-shared-strings | EPPlus | 26.57 ms | 3.51 ms | 2.03 ms | 8.88 | 8.88 | 12804.4 KB | 5.25 |  |  | 788.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | OfficeIMO.Excel | 2.78 ms | 0.07 ms | 0.04 ms | 1.00 | 1.00 | 2617.0 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | ClosedXML | 10.96 ms | 0.28 ms | 0.16 ms | 3.95 | 3.95 | 11299.2 KB | 4.32 |  |  | 294.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | EPPlus 4.5.3.3 | 14.80 ms |  |  | 5.33 | 5.33 |  |  |  |  | 432.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings | EPPlus | 19.78 ms | 0.11 ms | 0.06 ms | 7.12 | 7.12 | 12804.7 KB | 4.89 |  |  | 611.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | OfficeIMO.Excel | 2.66 ms | 0.47 ms | 0.27 ms | 1.00 | 1.00 | 2379.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | ClosedXML | 14.50 ms | 1.04 ms | 0.60 ms | 5.45 | 5.45 | 13127.1 KB | 5.52 |  |  | 444.9% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | EPPlus | 26.15 ms | 0.89 ms | 0.51 ms | 9.83 | 9.83 | 13893.0 KB | 5.84 |  |  | 882.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | OfficeIMO.Excel | 2.13 ms | 0.08 ms | 0.05 ms | 1.00 | 1.00 | 1579.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | ClosedXML | 10.64 ms | 1.08 ms | 0.62 ms | 5.00 | 5.00 | 9226.5 KB | 5.84 |  |  | 399.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | EPPlus | 19.63 ms | 0.48 ms | 0.28 ms | 9.22 | 9.22 | 11332.2 KB | 7.17 |  |  | 821.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | LargeXlsx | 3.63 ms | 0.41 ms | 0.24 ms | 0.79 | 1.00 | 758.3 KB | 0.43 |  |  | 21.0% faster than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 4.60 ms | 0.84 ms | 0.49 ms | 1.00 | 1.27 | 1769.2 KB | 1.00 |  |  | Loss +26.6% |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | MiniExcel | 9.76 ms | 1.38 ms | 0.80 ms | 2.12 | 2.69 | 23222.4 KB | 13.13 |  |  | 112.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 30.21 ms |  |  | 6.57 | 8.32 |  |  |  |  | 556.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | ClosedXML | 33.96 ms | 1.79 ms | 1.03 ms | 7.38 | 9.35 | 11581.0 KB | 6.55 |  |  | 638.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | EPPlus | 46.83 ms | 6.68 ms | 3.86 ms | 10.18 | 12.89 | 16646.4 KB | 9.41 |  |  | 918.1% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | LargeXlsx | 4.11 ms | 0.61 ms | 0.35 ms | 0.73 | 1.00 | 758.3 KB | 0.57 |  |  | 27.3% faster than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | OfficeIMO.Excel | 5.65 ms | 0.20 ms | 0.11 ms | 1.00 | 1.38 | 1339.3 KB | 1.00 |  |  | Loss +37.5% |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | MiniExcel | 11.13 ms | 1.99 ms | 1.15 ms | 1.97 | 2.71 | 23222.5 KB | 17.34 |  |  | 96.8% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | EPPlus 4.5.3.3 | 26.76 ms |  |  | 4.73 | 6.51 |  |  |  |  | 373.2% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | ClosedXML | 41.70 ms | 2.29 ms | 1.32 ms | 7.38 | 10.14 | 11581.0 KB | 8.65 |  |  | 637.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-direct | EPPlus | 54.21 ms | 9.13 ms | 5.27 ms | 9.59 | 13.19 | 16646.1 KB | 12.43 |  |  | 858.7% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 4.09 ms | 0.17 ms | 0.10 ms | 0.76 | 1.00 | 758.3 KB | 0.51 |  |  | 24.1% faster than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 5.39 ms | 0.20 ms | 0.11 ms | 1.00 | 1.32 | 1497.5 KB | 1.00 |  |  | Loss +31.8% |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | MiniExcel | 10.07 ms | 1.18 ms | 0.68 ms | 1.87 | 2.46 | 23222.5 KB | 15.51 |  |  | 86.6% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | ClosedXML | 40.54 ms | 7.52 ms | 4.34 ms | 7.51 | 9.91 | 11581.0 KB | 7.73 |  |  | 651.4% slower than OfficeIMO |
| 2500 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | EPPlus | 49.90 ms | 8.74 ms | 5.05 ms | 9.25 | 12.19 | 16646.1 KB | 11.12 |  |  | 825.0% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | Sylvan.Data.Excel | 37.49 ms | 0.25 ms | 0.15 ms | 0.78 | 1.00 | 394.1 KB | 0.02 |  |  | 22.1% faster than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | OfficeIMO.Excel | 48.16 ms | 3.76 ms | 2.17 ms | 1.00 | 1.28 | 23622.0 KB | 1.00 |  |  | Loss +28.4% |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | ExcelDataReader | 106.96 ms | 3.19 ms | 1.84 ms | 2.22 | 2.85 | 69530.7 KB | 2.94 |  |  | 122.1% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-range | MiniExcel | 142.41 ms | 6.19 ms | 3.57 ms | 2.96 | 3.80 | 215349.1 KB | 9.12 |  |  | 195.7% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | Sylvan.Data.Excel | 37.07 ms | 1.32 ms | 0.76 ms | 0.78 | 1.00 | 394.1 KB | 0.02 |  |  | 22.3% faster than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | OfficeIMO.Excel | 47.73 ms | 5.06 ms | 2.92 ms | 1.00 | 1.29 | 24404.4 KB | 1.00 |  |  | Loss +28.7% |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | ExcelDataReader | 107.61 ms | 0.92 ms | 0.53 ms | 2.25 | 2.90 | 69530.7 KB | 2.85 |  |  | 125.5% slower than OfficeIMO |
| 25000 | dense-helloworld-comparison | read | Other | dense-helloworld-read-stream | MiniExcel | 141.46 ms | 2.54 ms | 1.47 ms | 2.96 | 3.82 | 215349.1 KB | 8.82 |  |  | 196.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | append-plain-rows | LargeXlsx | 11.10 ms | 0.09 ms | 0.05 ms | 0.78 | 1.00 | 2771.0 KB | 0.26 | 605.0 KB | 0.99 | 22.1% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | append-plain-rows | OfficeIMO.Excel | 14.25 ms | 0.75 ms | 0.43 ms | 1.00 | 1.28 | 10842.5 KB | 1.00 | 610.4 KB | 1.00 | Loss +28.4% |
| 25000 | package-profile | package | Package size | append-plain-rows | MiniExcel | 29.97 ms | 1.27 ms | 0.73 ms | 2.10 | 2.70 | 58242.8 KB | 5.37 | 642.3 KB | 1.05 | 110.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | append-plain-rows | ClosedXML | 128.65 ms | 6.28 ms | 3.63 ms | 9.03 | 11.59 | 104233.1 KB | 9.61 | 540.6 KB | 0.89 | 802.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | append-plain-rows | EPPlus | 198.66 ms | 3.29 ms | 1.90 ms | 13.94 | 17.89 | 100373.5 KB | 9.26 | 525.6 KB | 0.86 | 1293.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | autofit-existing | OfficeIMO.Excel | 85.90 ms | 6.12 ms | 3.53 ms | 1.00 | 1.00 | 15708.3 KB | 1.00 | 1429.3 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | autofit-existing | EPPlus | 452.28 ms | 4.53 ms | 2.62 ms | 5.27 | 5.27 | 250950.0 KB | 15.98 | 1091.0 KB | 0.76 | 426.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | autofit-existing | ClosedXML | 1315.37 ms | 19.12 ms | 11.04 ms | 15.31 | 15.31 | 829720.5 KB | 52.82 | 1140.9 KB | 0.80 | 1431.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | large-shared-strings | OfficeIMO.Excel | 14.85 ms | 0.09 ms | 0.05 ms | 1.00 | 1.00 | 15744.9 KB | 1.00 | 529.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | large-shared-strings | MiniExcel | 27.87 ms | 0.37 ms | 0.21 ms | 1.88 | 1.88 | 73760.2 KB | 4.68 | 581.0 KB | 1.10 | 87.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | large-shared-strings | ClosedXML | 108.66 ms | 0.30 ms | 0.17 ms | 7.32 | 7.32 | 104241.3 KB | 6.62 | 460.1 KB | 0.87 | 631.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | large-shared-strings | EPPlus | 184.64 ms | 1.83 ms | 1.06 ms | 12.43 | 12.43 | 84410.0 KB | 5.36 | 444.7 KB | 0.84 | 1143.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-autofilter | OfficeIMO.Excel | 30.76 ms | 0.51 ms | 0.29 ms | 1.00 | 1.00 | 11494.9 KB | 1.00 | 1428.5 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-autofilter | ClosedXML | 294.92 ms | 0.60 ms | 0.35 ms | 9.59 | 9.59 | 210663.8 KB | 18.33 | 1140.0 KB | 0.80 | 858.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-autofilter | EPPlus | 367.81 ms | 12.04 ms | 6.95 ms | 11.96 | 11.96 | 211871.5 KB | 18.43 | 1090.1 KB | 0.76 | 1095.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-charts | OfficeIMO.Excel | 36.44 ms | 3.31 ms | 1.91 ms | 1.00 | 1.00 | 12550.9 KB | 1.00 | 1433.6 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-charts | EPPlus | 374.31 ms | 5.02 ms | 2.90 ms | 10.27 | 10.27 | 214905.8 KB | 17.12 | 1092.9 KB | 0.76 | 927.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-conditional-formatting | OfficeIMO.Excel | 30.36 ms | 0.17 ms | 0.10 ms | 1.00 | 1.00 | 11560.2 KB | 1.00 | 1428.8 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-conditional-formatting | ClosedXML | 286.34 ms | 4.91 ms | 2.83 ms | 9.43 | 9.43 | 210711.7 KB | 18.23 | 1140.1 KB | 0.80 | 843.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-conditional-formatting | EPPlus | 361.78 ms | 4.21 ms | 2.43 ms | 11.92 | 11.92 | 211912.9 KB | 18.33 | 1090.2 KB | 0.76 | 1091.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-data-validation | OfficeIMO.Excel | 32.26 ms | 1.92 ms | 1.11 ms | 1.00 | 1.00 | 11510.5 KB | 1.00 | 1428.6 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-data-validation | ClosedXML | 308.34 ms | 21.42 ms | 12.37 ms | 9.56 | 9.56 | 210672.7 KB | 18.30 | 1140.1 KB | 0.80 | 855.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-data-validation | EPPlus | 377.00 ms | 10.42 ms | 6.02 ms | 11.69 | 11.69 | 211857.4 KB | 18.41 | 1090.1 KB | 0.76 | 1068.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-freeze-panes | OfficeIMO.Excel | 31.38 ms | 1.23 ms | 0.71 ms | 1.00 | 1.00 | 11497.3 KB | 1.00 | 1428.6 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-freeze-panes | ClosedXML | 290.90 ms | 2.18 ms | 1.26 ms | 9.27 | 9.27 | 210646.8 KB | 18.32 | 1140.0 KB | 0.80 | 826.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-freeze-panes | EPPlus | 372.05 ms | 12.88 ms | 7.44 ms | 11.85 | 11.85 | 211883.3 KB | 18.43 | 1090.2 KB | 0.76 | 1085.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-pivot-table | OfficeIMO.Excel | 81.89 ms | 1.35 ms | 0.78 ms | 1.00 | 1.00 | 42218.3 KB | 1.00 | 1979.8 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-pivot-table | EPPlus | 395.47 ms | 1.26 ms | 0.73 ms | 4.83 | 4.83 | 230800.4 KB | 5.47 | 1093.4 KB | 0.55 | 382.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-all-in-one | OfficeIMO.Excel | 88.60 ms | 2.72 ms | 1.57 ms | 1.00 | 1.00 | 43677.9 KB | 1.00 | 1985.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-all-in-one | EPPlus | 422.81 ms | 8.26 ms | 4.77 ms | 4.77 | 4.77 | 277078.0 KB | 6.34 | 1097.7 KB | 0.55 | 377.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-chart-first | OfficeIMO.Excel | 90.24 ms | 4.24 ms | 2.45 ms | 1.00 | 1.00 | 43564.4 KB | 1.00 | 1985.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-chart-first | EPPlus | 424.94 ms | 4.71 ms | 2.72 ms | 4.71 | 4.71 | 277077.1 KB | 6.36 | 1097.7 KB | 0.55 | 370.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-core | OfficeIMO.Excel | 34.89 ms | 0.97 ms | 0.56 ms | 1.00 | 1.00 | 11648.7 KB | 1.00 | 1430.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-core | EPPlus | 413.51 ms | 22.21 ms | 12.82 ms | 11.85 | 11.85 | 255065.8 KB | 21.90 | 1091.5 KB | 0.76 | 1085.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-core | ClosedXML | 784.41 ms | 7.92 ms | 4.57 ms | 22.49 | 22.49 | 680116.8 KB | 58.39 | 1141.3 KB | 0.80 | 2148.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-extra-column | OfficeIMO.Excel | 99.55 ms | 3.67 ms | 2.12 ms | 1.00 | 1.00 | 45561.9 KB | 1.00 | 2110.8 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-extra-column | EPPlus | 469.40 ms | 8.21 ms | 4.74 ms | 4.72 | 4.72 | 302760.6 KB | 6.65 | 1166.3 KB | 0.55 | 371.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-no-autofit | OfficeIMO.Excel | 82.57 ms | 2.10 ms | 1.21 ms | 1.00 | 1.00 | 43671.9 KB | 1.00 | 1985.8 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-no-autofit | EPPlus | 389.25 ms | 14.20 ms | 8.20 ms | 4.71 | 4.71 | 234782.2 KB | 5.38 | 1097.7 KB | 0.55 | 371.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-post-mutation | OfficeIMO.Excel | 91.54 ms | 4.96 ms | 2.86 ms | 1.00 | 1.00 | 43687.3 KB | 1.00 | 1986.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-post-mutation | EPPlus | 433.54 ms | 11.74 ms | 6.78 ms | 4.74 | 4.74 | 277078.0 KB | 6.34 | 1097.8 KB | 0.55 | 373.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-shuffled-columns | OfficeIMO.Excel | 100.92 ms | 4.39 ms | 2.53 ms | 1.00 | 1.00 | 43743.0 KB | 1.00 | 2046.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-shuffled-columns | EPPlus | 435.26 ms | 25.44 ms | 14.69 ms | 4.31 | 4.31 | 277070.8 KB | 6.33 | 1098.4 KB | 0.54 | 331.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook | OfficeIMO.Excel | 114.88 ms | 2.88 ms | 1.66 ms | 1.00 | 1.00 | 59187.8 KB | 1.00 | 2672.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | report-workbook | EPPlus | 569.92 ms | 16.67 ms | 9.62 ms | 4.96 | 4.96 | 364709.8 KB | 6.16 | 1517.2 KB | 0.57 | 396.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook-core | OfficeIMO.Excel | 48.41 ms | 1.74 ms | 1.00 ms | 1.00 | 1.00 | 10979.4 KB | 1.00 | 1850.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | report-workbook-core | EPPlus | 578.88 ms | 22.00 ms | 12.70 ms | 11.96 | 11.96 | 342842.2 KB | 31.23 | 1512.6 KB | 0.82 | 1095.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook-core | ClosedXML | 1185.52 ms | 52.61 ms | 30.38 ms | 24.49 | 24.49 | 975775.9 KB | 88.87 | 1579.8 KB | 0.85 | 2348.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook-datatable | OfficeIMO.Excel | 125.75 ms | 8.17 ms | 4.71 ms | 1.00 | 1.00 | 61933.5 KB | 1.00 | 2672.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | report-workbook-datatable | EPPlus | 589.39 ms | 40.29 ms | 23.26 ms | 4.69 | 4.69 | 247823.6 KB | 4.00 | 1517.2 KB | 0.57 | 368.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook-datatable-core | OfficeIMO.Excel | 48.98 ms | 2.21 ms | 1.28 ms | 1.00 | 1.00 | 13725.0 KB | 1.00 | 1850.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | report-workbook-datatable-core | EPPlus | 519.39 ms | 3.93 ms | 2.27 ms | 10.60 | 10.60 | 225957.1 KB | 16.46 | 1512.6 KB | 0.82 | 960.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | report-workbook-datatable-core | ClosedXML | 1018.35 ms | 24.10 ms | 13.92 ms | 20.79 | 20.79 | 832228.3 KB | 60.64 | 1579.8 KB | 0.85 | 1979.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | LargeXlsx | 52.98 ms | 1.07 ms | 0.62 ms | 0.85 | 1.00 | 10795.2 KB | 0.92 | 2444.6 KB | 1.10 | 15.1% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | OfficeIMO.Excel | 62.44 ms | 7.31 ms | 4.22 ms | 1.00 | 1.18 | 11708.2 KB | 1.00 | 2228.8 KB | 1.00 | Loss +17.8% |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | MiniExcel | 189.99 ms | 6.02 ms | 3.47 ms | 3.04 | 3.59 | 226875.7 KB | 19.38 | 2410.6 KB | 1.08 | 204.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-blog-2023-20-string-columns | ClosedXML | 1113.84 ms | 37.39 ms | 21.59 ms | 17.84 | 21.02 | 759818.4 KB | 64.90 | 2581.2 KB | 1.16 | 1684.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-bulk-report | OfficeIMO.Excel | 35.39 ms | 1.28 ms | 0.74 ms | 1.00 | 1.00 | 11561.8 KB | 1.00 | 1429.3 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-bulk-report | MiniExcel | 66.86 ms | 1.31 ms | 0.76 ms | 1.89 | 1.89 | 125551.6 KB | 10.86 | 1521.1 KB | 1.06 | 88.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-bulk-report | EPPlus | 419.29 ms | 62.19 ms | 35.90 ms | 11.85 | 11.85 | 254959.0 KB | 22.05 | 1091.0 KB | 0.76 | 1084.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-bulk-report | ClosedXML | 718.65 ms | 4.07 ms | 2.35 ms | 20.30 | 20.30 | 565953.7 KB | 48.95 | 1140.9 KB | 0.80 | 1930.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellformula | OfficeIMO.Excel | 24.74 ms | 2.29 ms | 1.32 ms | 1.00 | 1.00 | 10112.0 KB | 1.00 | 670.3 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellformula | ClosedXML | 231.30 ms | 21.19 ms | 12.23 ms | 9.35 | 9.35 | 113853.5 KB | 11.26 | 643.2 KB | 0.96 | 834.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellformula | EPPlus | 407.68 ms | 23.86 ms | 13.77 ms | 16.48 | 16.48 | 140731.9 KB | 13.92 | 593.9 KB | 0.89 | 1547.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-empty-strings | OfficeIMO.Excel | 14.96 ms | 0.78 ms | 0.45 ms | 1.00 | 1.00 | 6896.4 KB | 1.00 | 451.4 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-empty-strings | ClosedXML | 152.22 ms | 12.69 ms | 7.32 ms | 10.18 | 10.18 | 92902.1 KB | 13.47 | 398.1 KB | 0.88 | 917.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-empty-strings | EPPlus | 196.67 ms | 6.50 ms | 3.75 ms | 13.15 | 13.15 | 74492.8 KB | 10.80 | 390.6 KB | 0.87 | 1215.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-numbers | OfficeIMO.Excel | 22.16 ms | 1.20 ms | 0.69 ms | 1.00 | 1.00 | 5970.9 KB | 1.00 | 462.6 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-numbers | ClosedXML | 141.31 ms | 8.77 ms | 5.06 ms | 6.38 | 6.38 | 84206.7 KB | 14.10 | 411.4 KB | 0.89 | 537.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-numbers | EPPlus | 237.33 ms | 15.49 ms | 8.95 ms | 10.71 | 10.71 | 86377.5 KB | 14.47 | 406.5 KB | 0.88 | 971.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-mixed | OfficeIMO.Excel | 26.90 ms | 1.51 ms | 0.87 ms | 1.00 | 1.00 | 8332.9 KB | 1.00 | 585.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-object-mixed | ClosedXML | 204.31 ms | 17.89 ms | 10.33 ms | 7.60 | 7.60 | 111118.7 KB | 13.33 | 532.9 KB | 0.91 | 659.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-mixed | EPPlus | 256.47 ms | 2.73 ms | 1.57 ms | 9.54 | 9.54 | 113245.1 KB | 13.59 | 544.3 KB | 0.93 | 853.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse | OfficeIMO.Excel | 24.34 ms | 2.68 ms | 1.54 ms | 1.00 | 1.00 | 7416.2 KB | 1.00 | 607.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse | ClosedXML | 199.93 ms | 26.11 ms | 15.07 ms | 8.22 | 8.22 | 105223.9 KB | 14.19 | 468.0 KB | 0.77 | 721.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse | EPPlus | 306.84 ms | 71.81 ms | 41.46 ms | 12.61 | 12.61 | 106316.9 KB | 14.34 | 494.4 KB | 0.81 | 1160.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 20.17 ms | 1.16 ms | 0.67 ms | 1.00 | 1.00 | 7416.3 KB | 1.00 | 607.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | ClosedXML | 171.99 ms | 16.04 ms | 9.26 ms | 8.53 | 8.53 | 105223.9 KB | 14.19 | 468.0 KB | 0.77 | 752.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-object-sparse-batch | EPPlus | 245.03 ms | 8.92 ms | 5.15 ms | 12.15 | 12.15 | 106316.9 KB | 14.34 | 494.4 KB | 0.81 | 1114.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-scalars | OfficeIMO.Excel | 14.54 ms | 1.47 ms | 0.85 ms | 1.00 | 1.00 | 6144.6 KB | 1.00 | 441.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-scalars | ClosedXML | 125.06 ms | 8.48 ms | 4.89 ms | 8.60 | 8.60 | 82591.3 KB | 13.44 | 394.9 KB | 0.89 | 760.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-scalars | EPPlus | 227.11 ms | 7.36 ms | 4.25 ms | 15.62 | 15.62 | 85127.4 KB | 13.85 | 379.3 KB | 0.86 | 1462.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings | OfficeIMO.Excel | 21.69 ms | 1.86 ms | 1.07 ms | 1.00 | 1.00 | 15360.4 KB | 1.00 | 527.8 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-strings | ClosedXML | 155.26 ms | 15.13 ms | 8.73 ms | 7.16 | 7.16 | 104241.3 KB | 6.79 | 460.1 KB | 0.87 | 616.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings | EPPlus | 217.36 ms | 10.60 ms | 6.12 ms | 10.02 | 10.02 | 84410.5 KB | 5.50 | 444.7 KB | 0.84 | 902.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-distinct | OfficeIMO.Excel | 18.79 ms | 2.77 ms | 1.60 ms | 1.00 | 1.00 | 13824.1 KB | 1.00 | 499.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-distinct | ClosedXML | 189.65 ms | 8.31 ms | 4.80 ms | 10.09 | 10.09 | 131501.7 KB | 9.51 | 555.3 KB | 1.11 | 909.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-distinct | EPPlus | 262.96 ms | 11.96 ms | 6.91 ms | 14.00 | 14.00 | 97729.6 KB | 7.07 | 565.1 KB | 1.13 | 1299.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-repeated | OfficeIMO.Excel | 15.44 ms | 2.67 ms | 1.54 ms | 1.00 | 1.00 | 7525.3 KB | 1.00 | 376.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-repeated | ClosedXML | 133.40 ms | 7.33 ms | 4.23 ms | 8.64 | 8.64 | 84520.0 KB | 11.23 | 331.8 KB | 0.88 | 764.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-strings-repeated | EPPlus | 214.21 ms | 36.68 ms | 21.18 ms | 13.87 | 13.87 | 70033.4 KB | 9.31 | 300.8 KB | 0.80 | 1287.5% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-temporal | OfficeIMO.Excel | 27.28 ms | 1.81 ms | 1.04 ms | 1.00 | 1.00 | 7482.8 KB | 1.00 | 620.5 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-cellvalue-temporal | ClosedXML | 196.97 ms | 9.41 ms | 5.43 ms | 7.22 | 7.22 | 89323.7 KB | 11.94 | 483.0 KB | 0.78 | 622.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalue-temporal | EPPlus | 273.95 ms | 54.20 ms | 31.29 ms | 10.04 | 10.04 | 103800.0 KB | 13.87 | 495.1 KB | 0.80 | 904.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 13.10 ms | 1.27 ms | 0.73 ms | 0.89 | 1.00 | 3444.4 KB | 0.49 | 443.4 KB | 0.97 | 11.4% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 14.79 ms | 1.40 ms | 0.81 ms | 1.00 | 1.13 | 6961.7 KB | 1.00 | 455.5 KB | 1.00 | Loss +12.9% |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | ClosedXML | 161.16 ms | 10.01 ms | 5.78 ms | 10.90 | 12.30 | 96015.7 KB | 13.79 | 467.5 KB | 1.03 | 989.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-headerless-rectangle-direct | EPPlus | 227.82 ms | 1.53 ms | 0.88 ms | 15.41 | 17.39 | 87466.9 KB | 12.56 | 484.1 KB | 1.06 | 1440.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | LargeXlsx | 33.84 ms | 2.02 ms | 1.16 ms | 0.75 | 1.00 | 5614.1 KB | 0.35 | 1386.5 KB | 1.00 | 24.7% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 44.94 ms | 4.26 ms | 2.46 ms | 1.00 | 1.33 | 16036.5 KB | 1.00 | 1384.9 KB | 1.00 | Loss +32.8% |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | MiniExcel | 89.91 ms | 3.82 ms | 2.20 ms | 2.00 | 2.66 | 93257.1 KB | 5.82 | 1521.0 KB | 1.10 | 100.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | ClosedXML | 356.62 ms | 20.09 ms | 11.60 ms | 7.94 | 10.54 | 210646.1 KB | 13.14 | 1139.9 KB | 0.82 | 693.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-cellvalues-rectangle-direct | EPPlus | 431.05 ms | 30.96 ms | 17.87 ms | 9.59 | 12.74 | 211849.9 KB | 13.21 | 1090.0 KB | 0.79 | 859.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | Sylvan.Data.Excel | 38.53 ms | 1.35 ms | 0.78 ms | 0.81 | 1.00 | 5700.3 KB | 0.44 | 755.4 KB | 0.55 | 18.8% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | LargeXlsx | 45.15 ms | 1.22 ms | 0.71 ms | 0.95 | 1.17 | 8349.2 KB | 0.64 | 1386.5 KB | 1.00 | 4.9% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | OfficeIMO.Excel | 47.47 ms | 2.24 ms | 1.29 ms | 1.00 | 1.23 | 13002.3 KB | 1.00 | 1384.9 KB | 1.00 | Loss +23.2% |
| 25000 | package-profile | package | Package size | write-datareader-plain | MiniExcel | 85.99 ms | 5.20 ms | 3.00 ms | 1.81 | 2.23 | 92199.8 KB | 7.09 | 1521.0 KB | 1.10 | 81.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | ClosedXML | 340.53 ms | 14.89 ms | 8.59 ms | 7.17 | 8.84 | 104205.0 KB | 8.01 | 1139.9 KB | 0.82 | 617.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-plain | EPPlus | 407.27 ms | 11.53 ms | 6.66 ms | 8.58 | 10.57 | 117437.7 KB | 9.03 | 1090.8 KB | 0.79 | 757.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table | OfficeIMO.Excel | 51.03 ms | 1.66 ms | 0.96 ms | 1.00 | 1.00 | 13020.3 KB | 1.00 | 1385.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-datareader-table | MiniExcel | 101.73 ms | 3.37 ms | 1.95 ms | 1.99 | 1.99 | 92200.2 KB | 7.08 | 1521.0 KB | 1.10 | 99.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table | EPPlus | 452.81 ms | 15.09 ms | 8.71 ms | 8.87 | 8.87 | 117437.3 KB | 9.02 | 1090.8 KB | 0.79 | 787.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table | ClosedXML | 512.33 ms | 10.08 ms | 5.82 ms | 10.04 | 10.04 | 173396.8 KB | 13.32 | 1140.7 KB | 0.82 | 904.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | OfficeIMO.Excel | 45.06 ms | 2.56 ms | 1.48 ms | 1.00 | 1.00 | 13026.6 KB | 1.00 | 1385.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | MiniExcel | 94.73 ms | 3.02 ms | 1.75 ms | 2.10 | 2.10 | 124495.5 KB | 9.56 | 1521.1 KB | 1.10 | 110.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | EPPlus | 433.23 ms | 25.84 ms | 14.92 ms | 9.61 | 9.61 | 159741.8 KB | 12.26 | 1091.0 KB | 0.79 | 861.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datareader-table-autofit | ClosedXML | 972.91 ms | 53.64 ms | 30.97 ms | 21.59 | 21.59 | 566146.0 KB | 43.46 | 1140.9 KB | 0.82 | 2059.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | OfficeIMO.Excel | 31.35 ms | 0.23 ms | 0.13 ms | 1.00 | 1.00 | 9819.7 KB | 1.00 | 1329.2 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | LargeXlsx | 34.77 ms | 0.56 ms | 0.32 ms | 1.11 | 1.11 | 9265.9 KB | 0.94 | 1680.0 KB | 1.26 | 10.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | MiniExcel | 94.69 ms | 3.83 ms | 2.21 ms | 3.02 | 3.02 | 108129.1 KB | 11.01 | 1819.7 KB | 1.37 | 202.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | EPPlus | 475.12 ms | 4.85 ms | 2.80 ms | 15.16 | 15.16 | 135723.5 KB | 13.82 | 1390.4 KB | 1.05 | 1415.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-direct-export | ClosedXML | 516.93 ms | 13.00 ms | 7.51 ms | 16.49 | 16.49 | 280375.8 KB | 28.55 | 1519.9 KB | 1.14 | 1549.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | OfficeIMO.Excel | 37.94 ms | 1.77 ms | 1.02 ms | 1.00 | 1.00 | 13458.5 KB | 1.00 | 1795.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | MiniExcel | 95.12 ms | 2.96 ms | 1.71 ms | 2.51 | 2.51 | 108129.1 KB | 8.03 | 1819.7 KB | 1.01 | 150.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | EPPlus | 501.01 ms | 21.78 ms | 12.57 ms | 13.20 | 13.20 | 135723.5 KB | 10.08 | 1390.4 KB | 0.77 | 1220.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-sparse-tables | ClosedXML | 528.34 ms | 11.10 ms | 6.41 ms | 13.92 | 13.92 | 280373.8 KB | 20.83 | 1519.9 KB | 0.85 | 1292.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables | OfficeIMO.Excel | 32.75 ms | 0.20 ms | 0.11 ms | 1.00 | 1.00 | 10288.1 KB | 1.00 | 1376.4 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-dataset-tables | MiniExcel | 74.48 ms | 2.03 ms | 1.17 ms | 2.27 | 2.27 | 97085.4 KB | 9.44 | 1511.8 KB | 1.10 | 127.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables | EPPlus | 315.41 ms | 3.62 ms | 2.09 ms | 9.63 | 9.63 | 110815.9 KB | 10.77 | 1100.6 KB | 0.80 | 863.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables | ClosedXML | 334.31 ms | 1.86 ms | 1.08 ms | 10.21 | 10.21 | 171997.9 KB | 16.72 | 1139.0 KB | 0.83 | 920.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | OfficeIMO.Excel | 37.93 ms | 2.03 ms | 1.17 ms | 1.00 | 1.00 | 10300.7 KB | 1.00 | 1376.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | MiniExcel | 80.85 ms | 0.99 ms | 0.57 ms | 2.13 | 2.13 | 128875.0 KB | 12.51 | 1512.0 KB | 1.10 | 113.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | EPPlus | 413.80 ms | 21.30 ms | 12.30 ms | 10.91 | 10.91 | 195407.9 KB | 18.97 | 1100.9 KB | 0.80 | 990.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-dataset-tables-autofit | ClosedXML | 724.66 ms | 35.67 ms | 20.59 ms | 19.10 | 19.10 | 550096.0 KB | 53.40 | 1139.3 KB | 0.83 | 1810.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-direct | LargeXlsx | 46.80 ms | 2.13 ms | 1.23 ms | 0.96 | 1.00 | 9520.4 KB | 0.75 | 1386.5 KB | 1.00 | 3.9% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-direct | OfficeIMO.Excel | 48.70 ms | 2.46 ms | 1.42 ms | 1.00 | 1.04 | 12715.7 KB | 1.00 | 1384.9 KB | 1.00 | Loss +4.1% |
| 25000 | package-profile | package | Package size | write-datatable-direct | MiniExcel | 110.39 ms | 5.37 ms | 3.10 ms | 2.27 | 2.36 | 92394.2 KB | 7.27 | 1521.0 KB | 1.10 | 126.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-direct | ClosedXML | 367.17 ms | 14.98 ms | 8.65 ms | 7.54 | 7.85 | 104205.0 KB | 8.19 | 1139.9 KB | 0.82 | 653.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-direct | EPPlus | 443.45 ms | 15.59 ms | 9.00 ms | 9.11 | 9.48 | 117437.3 KB | 9.24 | 1090.8 KB | 0.79 | 810.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | OfficeIMO.Excel | 43.59 ms | 1.77 ms | 1.02 ms | 1.00 | 1.00 | 12733.8 KB | 1.00 | 1385.7 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | MiniExcel | 108.29 ms | 0.60 ms | 0.35 ms | 2.48 | 2.48 | 92394.5 KB | 7.26 | 1521.1 KB | 1.10 | 148.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | EPPlus | 437.71 ms | 2.63 ms | 1.52 ms | 10.04 | 10.04 | 117437.3 KB | 9.22 | 1090.8 KB | 0.79 | 904.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-datatable-table-direct | ClosedXML | 501.41 ms | 11.22 ms | 6.48 ms | 11.50 | 11.50 | 173396.4 KB | 13.62 | 1140.7 KB | 0.82 | 1050.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | LargeXlsx | 27.15 ms | 0.31 ms | 0.18 ms | 0.87 | 1.00 | 5614.1 KB | 0.43 | 1386.5 KB | 1.00 | 13.5% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 31.39 ms | 0.46 ms | 0.27 ms | 1.00 | 1.16 | 12912.0 KB | 1.00 | 1384.9 KB | 1.00 | Loss +15.6% |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | MiniExcel | 64.33 ms | 1.44 ms | 0.83 ms | 2.05 | 2.37 | 93257.1 KB | 7.22 | 1521.1 KB | 1.10 | 105.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | ClosedXML | 263.69 ms | 8.82 ms | 5.09 ms | 8.40 | 9.71 | 104205.0 KB | 8.07 | 1139.9 KB | 0.82 | 740.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-fluent-rowsfrom-direct | EPPlus | 323.15 ms | 6.11 ms | 3.53 ms | 10.30 | 11.90 | 117437.7 KB | 9.10 | 1090.8 KB | 0.79 | 929.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 41.71 ms | 1.93 ms | 1.11 ms | 1.00 | 1.00 | 11501.6 KB | 1.00 | 1428.5 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 477.62 ms | 92.65 ms | 53.49 ms | 11.45 | 11.45 | 159742.2 KB | 13.89 | 1091.0 KB | 0.76 | 1045.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 819.12 ms | 52.25 ms | 30.17 ms | 19.64 | 19.64 | 496956.9 KB | 43.21 | 1140.1 KB | 0.80 | 1863.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | LargeXlsx | 37.23 ms | 2.84 ms | 1.64 ms | 0.89 | 1.00 | 5614.1 KB | 0.49 | 1386.5 KB | 0.97 | 11.1% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | OfficeIMO.Excel | 41.89 ms | 2.34 ms | 1.35 ms | 1.00 | 1.13 | 11493.8 KB | 1.00 | 1428.4 KB | 1.00 | Loss +12.5% |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | MiniExcel | 83.62 ms | 3.32 ms | 1.92 ms | 2.00 | 2.25 | 93257.1 KB | 8.11 | 1521.1 KB | 1.06 | 99.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | ClosedXML | 342.11 ms | 19.76 ms | 11.41 ms | 8.17 | 9.19 | 104205.0 KB | 9.07 | 1139.9 KB | 0.80 | 716.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-direct | EPPlus | 462.26 ms | 58.55 ms | 33.81 ms | 11.03 | 12.42 | 117437.3 KB | 10.22 | 1090.8 KB | 0.76 | 1003.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 41.91 ms | 3.30 ms | 1.91 ms | 1.00 | 1.00 | 10187.2 KB | 1.00 | 1385.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 376.33 ms | 5.05 ms | 2.91 ms | 8.98 | 8.98 | 159742.2 KB | 15.68 | 1091.0 KB | 0.79 | 798.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 660.30 ms | 20.85 ms | 12.04 ms | 15.76 | 15.76 | 496956.9 KB | 48.78 | 1140.1 KB | 0.82 | 1475.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 28.36 ms | 2.59 ms | 1.49 ms | 0.76 | 1.00 | 5614.1 KB | 0.55 | 1386.5 KB | 1.00 | 23.7% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 37.17 ms | 0.73 ms | 0.42 ms | 1.00 | 1.31 | 10179.4 KB | 1.00 | 1384.9 KB | 1.00 | Loss +31.1% |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | MiniExcel | 83.02 ms | 33.06 ms | 19.09 ms | 2.23 | 2.93 | 93257.1 KB | 9.16 | 1521.1 KB | 1.10 | 123.3% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | ClosedXML | 276.14 ms | 5.72 ms | 3.30 ms | 7.43 | 9.74 | 104205.0 KB | 10.24 | 1139.9 KB | 0.82 | 642.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-flat-dictionaries-direct | EPPlus | 335.68 ms | 1.03 ms | 0.60 ms | 9.03 | 11.84 | 117437.3 KB | 11.54 | 1090.8 KB | 0.79 | 803.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | LargeXlsx | 25.79 ms | 0.12 ms | 0.07 ms | 0.66 | 1.00 | 5614.1 KB | 0.36 | 1386.5 KB | 0.97 | 34.1% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | OfficeIMO.Excel | 39.14 ms | 2.23 ms | 1.29 ms | 1.00 | 1.52 | 15791.7 KB | 1.00 | 1428.4 KB | 1.00 | Loss +51.8% |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | MiniExcel | 59.78 ms | 0.95 ms | 0.55 ms | 1.53 | 2.32 | 93257.1 KB | 5.91 | 1521.0 KB | 1.06 | 52.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | ClosedXML | 257.84 ms | 4.99 ms | 2.88 ms | 6.59 | 10.00 | 104205.0 KB | 6.60 | 1139.9 KB | 0.80 | 558.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-legacy-dictionaries-direct | EPPlus | 322.95 ms | 1.78 ms | 1.03 ms | 8.25 | 12.52 | 117437.3 KB | 7.44 | 1090.8 KB | 0.76 | 725.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 33.39 ms | 2.38 ms | 1.38 ms | 1.00 | 1.00 | 11500.9 KB | 1.00 | 1428.5 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 373.66 ms | 17.61 ms | 10.17 ms | 11.19 | 11.19 | 138360.4 KB | 12.03 | 1091.0 KB | 0.76 | 1019.0% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 441.73 ms | 45.31 ms | 26.16 ms | 13.23 | 13.23 | 275422.3 KB | 23.95 | 1140.1 KB | 0.80 | 1222.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | LargeXlsx | 37.33 ms | 2.05 ms | 1.19 ms | 0.83 | 1.00 | 6043.9 KB | 0.57 | 1816.3 KB | 0.99 | 17.1% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 45.00 ms | 4.98 ms | 2.88 ms | 1.00 | 1.21 | 10577.2 KB | 1.00 | 1828.0 KB | 1.00 | Loss +20.6% |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | MiniExcel | 79.84 ms | 1.66 ms | 0.96 ms | 1.77 | 2.14 | 113974.3 KB | 10.78 | 1936.7 KB | 1.06 | 77.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | ClosedXML | 359.92 ms | 3.10 ms | 1.79 ms | 8.00 | 9.64 | 179552.5 KB | 16.98 | 1555.2 KB | 0.85 | 699.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-mixed-objects-direct | EPPlus | 438.18 ms | 6.67 ms | 3.85 ms | 9.74 | 11.74 | 144920.0 KB | 13.70 | 1473.0 KB | 0.81 | 873.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | LargeXlsx | 37.01 ms | 1.20 ms | 0.70 ms | 0.87 | 1.00 | 6043.9 KB | 0.61 | 1816.3 KB | 0.99 | 12.7% faster than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 42.38 ms | 0.39 ms | 0.23 ms | 1.00 | 1.15 | 9942.2 KB | 1.00 | 1828.0 KB | 1.00 | Loss +14.5% |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | MiniExcel | 79.29 ms | 2.45 ms | 1.42 ms | 1.87 | 2.14 | 113974.3 KB | 11.46 | 1936.7 KB | 1.06 | 87.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | ClosedXML | 371.02 ms | 13.44 ms | 7.76 ms | 8.75 | 10.03 | 179552.5 KB | 18.06 | 1555.2 KB | 0.85 | 775.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-mixed-direct | EPPlus | 448.05 ms | 16.60 ms | 9.58 ms | 10.57 | 12.11 | 144920.0 KB | 14.58 | 1473.0 KB | 0.81 | 957.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 190.90 ms | 3.93 ms | 2.27 ms | 1.00 | 1.00 | 36150.1 KB | 1.00 | 6725.6 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | LargeXlsx | 198.09 ms | 2.74 ms | 1.58 ms | 1.04 | 1.04 | 23211.4 KB | 0.64 | 6614.8 KB | 0.98 | 3.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | MiniExcel | 336.82 ms | 16.49 ms | 9.52 ms | 1.76 | 1.76 | 347925.7 KB | 9.62 | 6949.8 KB | 1.03 | 76.4% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | ClosedXML | 1166.60 ms | 35.89 ms | 20.72 ms | 6.11 | 6.11 | 487446.6 KB | 13.48 | 6165.9 KB | 0.92 | 511.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | write-powershell-psobject-wide-direct | EPPlus | 1513.98 ms | 21.45 ms | 12.38 ms | 7.93 | 7.93 | 562916.0 KB | 15.57 | 5441.6 KB | 0.81 | 693.1% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | OfficeIMO.Excel | 84.25 ms | 5.10 ms | 2.94 ms | 1.00 | 1.00 | 15708.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | EPPlus | 504.64 ms | 3.25 ms | 1.88 ms | 5.99 | 5.99 | 250948.5 KB | 15.98 |  |  | 499.0% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | EPPlus 4.5.3.3 | 652.49 ms |  |  | 7.74 | 7.74 |  |  |  |  | 674.5% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | autofit-existing | ClosedXML | 1563.09 ms | 59.06 ms | 34.10 ms | 18.55 | 18.55 | 829859.0 KB | 52.83 |  |  | 1755.3% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | OfficeIMO.Excel | 89.46 ms | 3.86 ms | 2.23 ms | 1.00 | 1.00 | 43670.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus 4.5.3.3 | 271.20 ms |  |  | 3.03 | 3.03 |  |  |  |  | 203.2% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus | 403.76 ms | 15.62 ms | 9.02 ms | 4.51 | 4.51 | 234781.2 KB | 5.38 |  |  | 351.4% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | OfficeIMO.Excel | 89.50 ms | 2.25 ms | 1.30 ms | 1.00 | 1.00 | 43558.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus | 434.47 ms | 19.00 ms | 10.97 ms | 4.85 | 4.85 | 277077.1 KB | 6.36 |  |  | 385.4% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus 4.5.3.3 | 514.29 ms |  |  | 5.75 | 5.75 |  |  |  |  | 474.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | OfficeIMO.Excel | 98.59 ms | 4.34 ms | 2.51 ms | 1.00 | 1.00 | 45565.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus | 473.03 ms | 5.93 ms | 3.42 ms | 4.80 | 4.80 | 302760.6 KB | 6.64 |  |  | 379.8% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus 4.5.3.3 | 595.57 ms |  |  | 6.04 | 6.04 |  |  |  |  | 504.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | OfficeIMO.Excel | 93.32 ms | 1.50 ms | 0.87 ms | 1.00 | 1.00 | 43686.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus | 416.94 ms | 7.85 ms | 4.53 ms | 4.47 | 4.47 | 277078.0 KB | 6.34 |  |  | 346.8% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus 4.5.3.3 | 477.41 ms |  |  | 5.12 | 5.12 |  |  |  |  | 411.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | OfficeIMO.Excel | 105.35 ms | 1.97 ms | 1.14 ms | 1.00 | 1.00 | 43738.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus | 434.19 ms | 12.41 ms | 7.17 ms | 4.12 | 4.12 | 277070.8 KB | 6.33 |  |  | 312.2% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus 4.5.3.3 | 488.51 ms |  |  | 4.64 | 4.64 |  |  |  |  | 363.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Object projection | build-object-datatable-dictionaries | OfficeIMO.Excel | 12.67 ms | 0.27 ms | 0.16 ms | 1.00 | 1.00 | 5164.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Object projection | build-object-datatable-typed | OfficeIMO.Excel | 10.47 ms | 0.18 ms | 0.11 ms | 1.00 | 1.00 | 8093.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-cells | OfficeIMO.Excel | 55.93 ms | 5.89 ms | 3.40 ms | 1.00 | 1.00 | 24530.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-cells | EPPlus | 318.84 ms | 12.34 ms | 7.13 ms | 5.70 | 5.70 | 187393.2 KB | 7.64 |  |  | 470.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-cells | ClosedXML | 440.20 ms | 10.86 ms | 6.27 ms | 7.87 | 7.87 | 166527.0 KB | 6.79 |  |  | 687.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | OfficeIMO.Excel | 41.78 ms | 6.52 ms | 3.77 ms | 1.00 | 1.00 | 3839.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | EPPlus | 273.17 ms | 7.87 ms | 4.54 ms | 6.54 | 6.54 | 115541.6 KB | 30.10 |  |  | 553.8% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-first-column-from-wide-sheet | ClosedXML | 394.13 ms | 21.20 ms | 12.24 ms | 9.43 | 9.43 | 150901.0 KB | 39.31 |  |  | 843.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-range | OfficeIMO.Excel | 67.22 ms | 0.61 ms | 0.35 ms | 1.00 | 1.00 | 24531.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-range | EPPlus | 302.40 ms | 23.14 ms | 13.36 ms | 4.50 | 4.50 | 187393.2 KB | 7.64 |  |  | 349.8% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-range | ClosedXML | 390.28 ms | 19.62 ms | 11.33 ms | 5.81 | 5.81 | 166520.9 KB | 6.79 |  |  | 480.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-top-range | OfficeIMO.Excel | 0.74 ms | 0.03 ms | 0.02 ms | 1.00 | 1.00 | 285.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Range and table read | enumerate-top-range | EPPlus | 257.85 ms | 32.44 ms | 18.73 ms | 349.75 | 349.75 | 105580.1 KB | 370.12 |  |  | 34874.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Range and table read | enumerate-top-range | ClosedXML | 396.23 ms | 9.95 ms | 5.74 ms | 537.46 | 537.46 | 149396.1 KB | 523.72 |  |  | 53645.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-autofilter | OfficeIMO.Excel | 32.65 ms | 0.34 ms | 0.20 ms | 1.00 | 1.00 | 11494.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-autofilter | EPPlus 4.5.3.3 | 233.13 ms |  |  | 7.14 | 7.14 |  |  |  |  | 613.9% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-autofilter | ClosedXML | 309.25 ms | 7.11 ms | 4.11 ms | 9.47 | 9.47 | 210663.8 KB | 18.33 |  |  | 847.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-autofilter | EPPlus | 367.66 ms | 1.36 ms | 0.78 ms | 11.26 | 11.26 | 211871.5 KB | 18.43 |  |  | 1025.9% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-charts | OfficeIMO.Excel | 32.68 ms | 1.18 ms | 0.68 ms | 1.00 | 1.00 | 12549.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-charts | EPPlus 4.5.3.3 | 240.13 ms |  |  | 7.35 | 7.35 |  |  |  |  | 634.8% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-charts | EPPlus | 370.85 ms | 6.93 ms | 4.00 ms | 11.35 | 11.35 | 214905.8 KB | 17.13 |  |  | 1034.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | OfficeIMO.Excel | 31.89 ms | 1.05 ms | 0.61 ms | 1.00 | 1.00 | 11560.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | EPPlus 4.5.3.3 | 239.57 ms |  |  | 7.51 | 7.51 |  |  |  |  | 651.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | ClosedXML | 300.39 ms | 10.90 ms | 6.29 ms | 9.42 | 9.42 | 210711.7 KB | 18.23 |  |  | 842.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-conditional-formatting | EPPlus | 362.73 ms | 14.14 ms | 8.16 ms | 11.38 | 11.38 | 211912.9 KB | 18.33 |  |  | 1037.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-data-validation | OfficeIMO.Excel | 32.47 ms | 2.57 ms | 1.48 ms | 1.00 | 1.00 | 11510.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-data-validation | EPPlus 4.5.3.3 | 238.04 ms |  |  | 7.33 | 7.33 |  |  |  |  | 633.2% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-data-validation | ClosedXML | 294.79 ms | 0.76 ms | 0.44 ms | 9.08 | 9.08 | 210672.7 KB | 18.30 |  |  | 808.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-data-validation | EPPlus | 369.52 ms | 18.49 ms | 10.68 ms | 11.38 | 11.38 | 211857.4 KB | 18.41 |  |  | 1038.2% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | OfficeIMO.Excel | 33.62 ms | 1.37 ms | 0.79 ms | 1.00 | 1.00 | 11497.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | EPPlus 4.5.3.3 | 231.67 ms |  |  | 6.89 | 6.89 |  |  |  |  | 589.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | ClosedXML | 298.44 ms | 5.75 ms | 3.32 ms | 8.88 | 8.88 | 210646.8 KB | 18.32 |  |  | 787.7% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-freeze-panes | EPPlus | 372.05 ms | 11.39 ms | 6.58 ms | 11.07 | 11.07 | 211883.3 KB | 18.43 |  |  | 1006.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | OfficeIMO.Excel | 83.49 ms | 0.71 ms | 0.41 ms | 1.00 | 1.00 | 42217.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | EPPlus 4.5.3.3 | 238.99 ms |  |  | 2.86 | 2.86 |  |  |  |  | 186.2% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world feature mix | realworld-pivot-table | EPPlus | 394.56 ms | 8.18 ms | 4.72 ms | 4.73 | 4.73 | 230800.4 KB | 5.47 |  |  | 372.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | OfficeIMO.Excel | 91.44 ms | 5.14 ms | 2.97 ms | 1.00 | 1.00 | 43675.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus | 425.47 ms | 10.12 ms | 5.85 ms | 4.65 | 4.65 | 277078.0 KB | 6.34 |  |  | 365.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus 4.5.3.3 | 616.24 ms |  |  | 6.74 | 6.74 |  |  |  |  | 573.9% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world report | realworld-report-core | OfficeIMO.Excel | 34.81 ms | 0.78 ms | 0.45 ms | 1.00 | 1.00 | 11648.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world report | realworld-report-core | EPPlus | 400.27 ms | 5.38 ms | 3.10 ms | 11.50 | 11.50 | 255065.8 KB | 21.90 |  |  | 1049.9% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world report | realworld-report-core | EPPlus 4.5.3.3 | 470.96 ms |  |  | 13.53 | 13.53 |  |  |  |  | 1253.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world report | realworld-report-core | ClosedXML | 801.68 ms | 8.91 ms | 5.14 ms | 23.03 | 23.03 | 680114.5 KB | 58.39 |  |  | 2203.2% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook | OfficeIMO.Excel | 116.92 ms | 2.95 ms | 1.70 ms | 1.00 | 1.00 | 59145.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Report workbook | report-workbook | EPPlus | 603.69 ms | 46.30 ms | 26.73 ms | 5.16 | 5.16 | 364709.8 KB | 6.17 |  |  | 416.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook | EPPlus 4.5.3.3 | 618.55 ms |  |  | 5.29 | 5.29 |  |  |  |  | 429.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-core | OfficeIMO.Excel | 46.03 ms | 1.92 ms | 1.11 ms | 1.00 | 1.00 | 10979.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Report workbook | report-workbook-core | EPPlus | 527.81 ms | 6.29 ms | 3.63 ms | 11.47 | 11.47 | 342842.2 KB | 31.23 |  |  | 1046.8% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-core | EPPlus 4.5.3.3 | 616.09 ms |  |  | 13.39 | 13.39 |  |  |  |  | 1238.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-core | ClosedXML | 1135.82 ms | 22.97 ms | 13.26 ms | 24.68 | 24.68 | 975773.4 KB | 88.87 |  |  | 2367.8% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable | OfficeIMO.Excel | 126.00 ms | 4.91 ms | 2.84 ms | 1.00 | 1.00 | 61935.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable | EPPlus | 565.03 ms | 7.08 ms | 4.09 ms | 4.48 | 4.48 | 247823.7 KB | 4.00 |  |  | 348.4% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable | EPPlus 4.5.3.3 | 697.91 ms |  |  | 5.54 | 5.54 |  |  |  |  | 453.9% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable-core | OfficeIMO.Excel | 49.99 ms | 0.68 ms | 0.39 ms | 1.00 | 1.00 | 13725.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable-core | EPPlus | 542.99 ms | 3.78 ms | 2.18 ms | 10.86 | 10.86 | 225957.1 KB | 16.46 |  |  | 986.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable-core | EPPlus 4.5.3.3 | 617.24 ms |  |  | 12.35 | 12.35 |  |  |  |  | 1134.6% slower than OfficeIMO |
| 25000 | speed-comparison | other | Report workbook | report-workbook-datatable-core | ClosedXML | 1091.46 ms | 21.77 ms | 12.57 ms | 21.83 | 21.83 | 832225.9 KB | 60.64 |  |  | 2083.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | OfficeIMO.Excel | 19.08 ms | 0.90 ms | 0.52 ms | 1.00 | 1.00 | 6219.0 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | EPPlus 4.5.3.3 | 73.40 ms |  |  | 3.85 | 3.85 |  |  |  |  | 284.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | EPPlus | 151.27 ms | 9.05 ms | 5.23 ms | 7.93 | 7.93 | 70814.5 KB | 11.39 |  |  | 692.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Formula write/read | formula-heavy-read | ClosedXML | 170.22 ms | 9.54 ms | 5.51 ms | 8.92 | 8.92 | 79515.6 KB | 12.79 |  |  | 792.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | OfficeIMO.Excel | 0.84 ms | 0.02 ms | 0.01 ms | 1.00 | 1.00 | 177.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | Sylvan.Data.Excel | 1.04 ms | 0.09 ms | 0.05 ms | 1.24 | 1.24 | 316.6 KB | 1.79 |  |  | 23.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | ExcelDataReader | 1.80 ms | 0.61 ms | 0.35 ms | 2.14 | 2.14 | 4062.2 KB | 22.91 |  |  | 113.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | ClosedXML | 3.70 ms | 0.19 ms | 0.11 ms | 4.39 | 4.39 | 4392.8 KB | 24.77 |  |  | 339.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | MiniExcel | 12.33 ms | 1.26 ms | 0.73 ms | 14.66 | 14.66 | 46194.9 KB | 260.53 |  |  | 1366.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | EPPlus 4.5.3.3 | 32.60 ms |  |  | 38.76 | 38.76 |  |  |  |  | 3775.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-column-read | EPPlus | 93.94 ms | 2.73 ms | 1.58 ms | 111.68 | 111.68 | 43071.0 KB | 242.91 |  |  | 11067.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | OfficeIMO.Excel | 0.83 ms | 0.01 ms | 0.01 ms | 1.00 | 1.00 | 177.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | Sylvan.Data.Excel | 1.02 ms | 0.04 ms | 0.02 ms | 1.23 | 1.23 | 316.6 KB | 1.78 |  |  | 22.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | ExcelDataReader | 1.97 ms | 0.46 ms | 0.27 ms | 2.37 | 2.37 | 4062.2 KB | 22.90 |  |  | 136.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | ClosedXML | 3.36 ms | 0.18 ms | 0.10 ms | 4.04 | 4.04 | 4392.8 KB | 24.76 |  |  | 304.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | MiniExcel | 12.87 ms | 0.77 ms | 0.45 ms | 15.49 | 15.49 | 46194.9 KB | 260.43 |  |  | 1448.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | EPPlus 4.5.3.3 | 36.10 ms |  |  | 43.43 | 43.43 |  |  |  |  | 4243.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | large-sparse-row-read | EPPlus | 94.92 ms | 2.59 ms | 1.50 ms | 114.21 | 114.21 | 43071.0 KB | 242.81 |  |  | 11321.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | Sylvan.Data.Excel | 15.68 ms | 0.19 ms | 0.11 ms | 0.87 | 1.00 | 1936.7 KB | 0.21 |  |  | 12.6% faster than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | OfficeIMO.Excel | 17.94 ms | 0.29 ms | 0.17 ms | 1.00 | 1.14 | 9218.0 KB | 1.00 |  |  | Loss +14.4% |
| 25000 | speed-comparison | read | Other | shared-string-read | ExcelDataReader | 41.52 ms | 0.26 ms | 0.15 ms | 2.31 | 2.65 | 25020.8 KB | 2.71 |  |  | 131.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | MiniExcel | 46.74 ms | 0.65 ms | 0.38 ms | 2.61 | 2.98 | 74405.3 KB | 8.07 |  |  | 160.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | EPPlus 4.5.3.3 | 83.96 ms |  |  | 4.68 | 5.35 |  |  |  |  | 368.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | EPPlus | 146.12 ms | 3.03 ms | 1.75 ms | 8.15 | 9.32 | 89346.0 KB | 9.69 |  |  | 714.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Other | shared-string-read | ClosedXML | 151.07 ms | 6.65 ms | 3.84 ms | 8.42 | 9.63 | 90414.4 KB | 9.81 |  |  | 742.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | OfficeIMO.Excel | 37.82 ms | 7.24 ms | 4.18 ms | 1.00 | 1.00 | 1122.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | Sylvan.Data.Excel | 58.11 ms | 1.70 ms | 0.98 ms | 1.54 | 1.54 | 3534.8 KB | 3.15 |  |  | 53.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | ExcelDataReader | 148.27 ms | 16.27 ms | 9.39 ms | 3.92 | 3.92 | 61201.9 KB | 54.53 |  |  | 292.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | MiniExcel | 165.33 ms | 4.49 ms | 2.59 ms | 4.37 | 4.37 | 186420.9 KB | 166.11 |  |  | 337.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | EPPlus | 276.62 ms | 9.88 ms | 5.71 ms | 7.31 | 7.31 | 105609.0 KB | 94.10 |  |  | 631.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-bottom-range | ClosedXML | 407.13 ms | 24.08 ms | 13.90 ms | 10.76 | 10.76 | 149389.9 KB | 133.11 |  |  | 976.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | Sylvan.Data.Excel | 80.47 ms | 6.71 ms | 3.87 ms | 0.98 | 1.00 | 18394.2 KB | 0.53 |  |  | 2.4% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | OfficeIMO.Excel | 82.41 ms | 8.33 ms | 4.81 ms | 1.00 | 1.02 | 34645.8 KB | 1.00 |  |  | Loss +2.4% |
| 25000 | speed-comparison | read | Range and table read | read-datatable | ExcelDataReader | 177.66 ms | 9.90 ms | 5.71 ms | 2.16 | 2.21 | 76061.4 KB | 2.20 |  |  | 115.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | MiniExcel | 203.27 ms | 1.45 ms | 0.84 ms | 2.47 | 2.53 | 181285.0 KB | 5.23 |  |  | 146.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | EPPlus 4.5.3.3 | 210.09 ms |  |  | 2.55 | 2.61 |  |  |  |  | 154.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | EPPlus | 382.50 ms | 86.76 ms | 50.09 ms | 4.64 | 4.75 | 202250.2 KB | 5.84 |  |  | 364.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-datatable | ClosedXML | 450.16 ms | 8.39 ms | 4.85 ms | 5.46 | 5.59 | 178450.1 KB | 5.15 |  |  | 446.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | OfficeIMO.Excel | 41.76 ms | 5.89 ms | 3.40 ms | 1.00 | 1.00 | 4034.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | Sylvan.Data.Excel | 52.65 ms | 7.00 ms | 4.04 ms | 1.26 | 1.26 | 4316.2 KB | 1.07 |  |  | 26.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | MiniExcel | 131.17 ms | 4.85 ms | 2.80 ms | 3.14 | 3.14 | 158612.9 KB | 39.31 |  |  | 214.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | ExcelDataReader | 153.14 ms | 5.66 ms | 3.27 ms | 3.67 | 3.67 | 61202.0 KB | 15.17 |  |  | 266.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | EPPlus | 279.75 ms | 5.30 ms | 3.06 ms | 6.70 | 6.70 | 115541.6 KB | 28.64 |  |  | 569.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-first-column-from-wide-sheet | ClosedXML | 405.28 ms | 11.01 ms | 6.35 ms | 9.71 | 9.71 | 150898.7 KB | 37.40 |  |  | 870.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | Sylvan.Data.Excel | 62.62 ms | 1.85 ms | 1.07 ms | 0.88 | 1.00 | 3534.8 KB | 0.14 |  |  | 11.8% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | OfficeIMO.Excel | 70.99 ms | 9.88 ms | 5.71 ms | 1.00 | 1.13 | 26098.2 KB | 1.00 |  |  | Loss +13.4% |
| 25000 | speed-comparison | read | Range and table read | read-range | ExcelDataReader | 155.87 ms | 5.04 ms | 2.91 ms | 2.20 | 2.49 | 61201.9 KB | 2.35 |  |  | 119.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | MiniExcel | 159.61 ms | 12.24 ms | 7.06 ms | 2.25 | 2.55 | 186421.5 KB | 7.14 |  |  | 124.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | EPPlus 4.5.3.3 | 235.32 ms |  |  | 3.31 | 3.76 |  |  |  |  | 231.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | EPPlus | 306.92 ms | 4.85 ms | 2.80 ms | 4.32 | 4.90 | 187390.9 KB | 7.18 |  |  | 332.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range | ClosedXML | 402.78 ms | 39.73 ms | 22.94 ms | 5.67 | 6.43 | 163589.1 KB | 6.27 |  |  | 467.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | Sylvan.Data.Excel | 61.40 ms | 7.38 ms | 4.26 ms | 0.83 | 1.00 | 4484.9 KB | 0.17 |  |  | 17.4% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | OfficeIMO.Excel | 74.37 ms | 7.05 ms | 4.07 ms | 1.00 | 1.21 | 26684.1 KB | 1.00 |  |  | Loss +21.1% |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | ExcelDataReader | 136.18 ms | 7.14 ms | 4.12 ms | 1.83 | 2.22 | 61201.9 KB | 2.29 |  |  | 83.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | MiniExcel | 162.06 ms | 5.43 ms | 3.14 ms | 2.18 | 2.64 | 186421.5 KB | 6.99 |  |  | 117.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | EPPlus | 393.26 ms | 12.41 ms | 7.17 ms | 5.29 | 6.41 | 187390.9 KB | 7.02 |  |  | 428.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-range-decimal | ClosedXML | 427.27 ms | 31.06 ms | 17.93 ms | 5.74 | 6.96 | 163589.3 KB | 6.13 |  |  | 474.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | Sylvan.Data.Excel | 0.51 ms | 0.08 ms | 0.05 ms | 0.76 | 1.00 | 348.5 KB | 1.18 |  |  | 24.0% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | OfficeIMO.Excel | 0.67 ms | 0.02 ms | 0.01 ms | 1.00 | 1.32 | 296.0 KB | 1.00 |  |  | Loss +31.6% |
| 25000 | speed-comparison | read | Range and table read | read-top-range | MiniExcel | 0.96 ms | 0.08 ms | 0.05 ms | 1.43 | 1.88 | 869.0 KB | 2.94 |  |  | 42.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | ExcelDataReader | 50.63 ms | 7.30 ms | 4.21 ms | 75.36 | 99.18 | 17115.3 KB | 57.83 |  |  | 7436.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | EPPlus 4.5.3.3 | 259.78 ms |  |  | 386.69 | 508.90 |  |  |  |  | 38568.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | EPPlus | 273.27 ms | 3.65 ms | 2.11 ms | 406.78 | 535.34 | 105577.9 KB | 356.74 |  |  | 40578.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-top-range | ClosedXML | 400.35 ms | 27.25 ms | 15.73 ms | 595.93 | 784.27 | 149391.8 KB | 504.78 |  |  | 59492.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | Sylvan.Data.Excel | 66.41 ms | 4.80 ms | 2.77 ms | 0.92 | 1.00 | 3534.8 KB | 0.14 |  |  | 8.1% faster than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | OfficeIMO.Excel | 72.30 ms | 10.79 ms | 6.23 ms | 1.00 | 1.09 | 26156.0 KB | 1.00 |  |  | Loss +8.9% |
| 25000 | speed-comparison | read | Range and table read | read-used-range | ExcelDataReader | 148.97 ms | 9.01 ms | 5.20 ms | 2.06 | 2.24 | 61201.9 KB | 2.34 |  |  | 106.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | MiniExcel | 158.05 ms | 11.14 ms | 6.43 ms | 2.19 | 2.38 | 186421.5 KB | 7.13 |  |  | 118.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | EPPlus | 297.28 ms | 34.61 ms | 19.98 ms | 4.11 | 4.48 | 187390.9 KB | 7.16 |  |  | 311.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Range and table read | read-used-range | ClosedXML | 450.53 ms | 2.68 ms | 1.55 ms | 6.23 | 6.78 | 163586.6 KB | 6.25 |  |  | 523.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | OfficeIMO.Excel | 31.74 ms | 0.28 ms | 0.16 ms | 1.00 | 1.00 | 1125.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | Sylvan.Data.Excel | 38.23 ms | 0.79 ms | 0.45 ms | 1.20 | 1.20 | 3534.8 KB | 3.14 |  |  | 20.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | ExcelDataReader | 104.82 ms | 1.77 ms | 1.02 ms | 3.30 | 3.30 | 61201.9 KB | 54.37 |  |  | 230.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | MiniExcel | 116.31 ms | 5.01 ms | 2.89 ms | 3.66 | 3.66 | 186420.9 KB | 165.61 |  |  | 266.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | EPPlus | 230.07 ms | 20.29 ms | 11.71 ms | 7.25 | 7.25 | 105609.0 KB | 93.82 |  |  | 624.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-bottom-range-stream | ClosedXML | 310.27 ms | 23.08 ms | 13.32 ms | 9.78 | 9.78 | 149391.3 KB | 132.71 |  |  | 877.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | Sylvan.Data.Excel | 45.29 ms | 3.88 ms | 2.24 ms | 0.77 | 1.00 | 3534.8 KB | 0.13 |  |  | 22.7% faster than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | OfficeIMO.Excel | 58.57 ms | 13.74 ms | 7.93 ms | 1.00 | 1.29 | 26885.3 KB | 1.00 |  |  | Loss +29.3% |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | ExcelDataReader | 116.60 ms | 10.69 ms | 6.17 ms | 1.99 | 2.57 | 61201.9 KB | 2.28 |  |  | 99.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | MiniExcel | 121.85 ms | 11.84 ms | 6.84 ms | 2.08 | 2.69 | 186421.5 KB | 6.93 |  |  | 108.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | EPPlus 4.5.3.3 | 167.71 ms |  |  | 2.86 | 3.70 |  |  |  |  | 186.4% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | EPPlus | 340.04 ms | 56.93 ms | 32.87 ms | 5.81 | 7.51 | 187390.9 KB | 6.97 |  |  | 480.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-range-stream | ClosedXML | 368.83 ms | 43.85 ms | 25.32 ms | 6.30 | 8.14 | 163591.4 KB | 6.08 |  |  | 529.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | Sylvan.Data.Excel | 0.42 ms | 0.02 ms | 0.01 ms | 0.80 | 1.00 | 348.5 KB | 1.16 |  |  | 19.8% faster than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | OfficeIMO.Excel | 0.53 ms | 0.01 ms | 0.01 ms | 1.00 | 1.25 | 299.3 KB | 1.00 |  |  | Loss +24.7% |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | MiniExcel | 0.82 ms | 0.24 ms | 0.14 ms | 1.56 | 1.95 | 869.0 KB | 2.90 |  |  | 56.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | ExcelDataReader | 36.07 ms | 0.46 ms | 0.27 ms | 68.42 | 85.32 | 17115.3 KB | 57.18 |  |  | 6741.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | EPPlus 4.5.3.3 | 169.07 ms |  |  | 320.69 | 399.90 |  |  |  |  | 31968.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | EPPlus | 228.12 ms | 17.90 ms | 10.33 ms | 432.69 | 539.58 | 105577.7 KB | 352.70 |  |  | 43169.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream | ClosedXML | 315.57 ms | 33.05 ms | 19.08 ms | 598.58 | 746.44 | 149390.2 KB | 499.06 |  |  | 59757.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | Sylvan.Data.Excel | 0.41 ms | 0.01 ms | 0.01 ms | 0.75 | 1.00 | 348.5 KB | 1.16 |  |  | 25.0% faster than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | OfficeIMO.Excel | 0.54 ms | 0.02 ms | 0.01 ms | 1.00 | 1.33 | 300.0 KB | 1.00 |  |  | Loss +33.3% |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | MiniExcel | 0.83 ms | 0.21 ms | 0.12 ms | 1.52 | 2.03 | 869.0 KB | 2.90 |  |  | 52.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | ExcelDataReader | 39.64 ms | 3.67 ms | 2.12 ms | 73.10 | 97.44 | 17115.3 KB | 57.05 |  |  | 7210.0% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | EPPlus | 224.36 ms | 3.48 ms | 2.01 ms | 413.72 | 551.48 | 105577.8 KB | 351.90 |  |  | 41271.8% slower than OfficeIMO |
| 25000 | speed-comparison | read | Streaming read | read-top-range-stream-small-chunks | ClosedXML | 317.56 ms | 8.32 ms | 4.81 ms | 585.59 | 780.57 | 149391.0 KB | 497.93 |  |  | 58458.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | Sylvan.Data.Excel | 43.41 ms | 1.98 ms | 1.14 ms | 0.77 | 1.00 | 5805.0 KB | 0.25 |  |  | 22.6% faster than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | OfficeIMO.Excel | 56.11 ms | 7.91 ms | 4.57 ms | 1.00 | 1.29 | 23562.3 KB | 1.00 |  |  | Loss +29.3% |
| 25000 | speed-comparison | read | Typed object read | read-objects | ExcelDataReader | 116.41 ms | 4.56 ms | 2.63 ms | 2.07 | 2.68 | 63472.1 KB | 2.69 |  |  | 107.5% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | MiniExcel | 143.73 ms | 7.06 ms | 4.08 ms | 2.56 | 3.31 | 183656.4 KB | 7.79 |  |  | 156.1% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | EPPlus 4.5.3.3 | 242.79 ms |  |  | 4.33 | 5.59 |  |  |  |  | 332.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | EPPlus | 336.83 ms | 14.85 ms | 8.57 ms | 6.00 | 7.76 | 199608.2 KB | 8.47 |  |  | 500.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects | ClosedXML | 361.17 ms | 17.65 ms | 10.19 ms | 6.44 | 8.32 | 165540.4 KB | 7.03 |  |  | 543.7% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | OfficeIMO.Excel | 57.68 ms | 21.28 ms | 12.29 ms | 1.00 | 1.00 | 23367.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | Sylvan.Data.Excel | 86.11 ms | 38.86 ms | 22.43 ms | 1.49 | 1.49 | 5292.6 KB | 0.23 |  |  | 49.3% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | ExcelDataReader | 152.38 ms | 18.87 ms | 10.89 ms | 2.64 | 2.64 | 62959.7 KB | 2.69 |  |  | 164.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | MiniExcel | 224.89 ms | 37.07 ms | 21.40 ms | 3.90 | 3.90 | 183144.1 KB | 7.84 |  |  | 289.9% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | EPPlus 4.5.3.3 | 244.91 ms |  |  | 4.25 | 4.25 |  |  |  |  | 324.6% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | EPPlus | 403.29 ms | 13.31 ms | 7.68 ms | 6.99 | 6.99 | 199412.8 KB | 8.53 |  |  | 599.2% slower than OfficeIMO |
| 25000 | speed-comparison | read | Typed object read | read-objects-stream | ClosedXML | 470.11 ms | 56.58 ms | 32.67 ms | 8.15 | 8.15 | 165348.9 KB | 7.08 |  |  | 715.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | OfficeIMO.Excel | 48.83 ms | 3.38 ms | 1.95 ms | 1.00 | 1.00 | 13026.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | MiniExcel | 99.23 ms | 1.56 ms | 0.90 ms | 2.03 | 2.03 | 124495.5 KB | 9.56 |  |  | 103.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | EPPlus | 432.69 ms | 5.36 ms | 3.10 ms | 8.86 | 8.86 | 159741.8 KB | 12.26 |  |  | 786.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | EPPlus 4.5.3.3 | 479.10 ms |  |  | 9.81 | 9.81 |  |  |  |  | 881.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-datareader-table-autofit | ClosedXML | 965.39 ms | 25.63 ms | 14.80 ms | 19.77 | 19.77 | 566142.6 KB | 43.46 |  |  | 1877.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | OfficeIMO.Excel | 47.84 ms | 4.37 ms | 2.52 ms | 1.00 | 1.00 | 10300.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | MiniExcel | 105.19 ms | 4.15 ms | 2.40 ms | 2.20 | 2.20 | 128875.0 KB | 12.51 |  |  | 119.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | EPPlus | 512.22 ms | 40.25 ms | 23.24 ms | 10.71 | 10.71 | 195407.9 KB | 18.97 |  |  | 970.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-dataset-tables-autofit | ClosedXML | 953.09 ms | 119.05 ms | 68.73 ms | 19.92 | 19.92 | 550092.4 KB | 53.40 |  |  | 1892.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | OfficeIMO.Excel | 44.97 ms | 4.78 ms | 2.76 ms | 1.00 | 1.00 | 11501.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | EPPlus | 448.91 ms | 6.51 ms | 3.76 ms | 9.98 | 9.98 | 159742.3 KB | 13.89 |  |  | 898.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-autofitcolumnsfor-direct | ClosedXML | 851.90 ms | 23.79 ms | 13.73 ms | 18.94 | 18.94 | 496956.9 KB | 43.21 |  |  | 1794.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | OfficeIMO.Excel | 51.37 ms | 6.09 ms | 3.52 ms | 1.00 | 1.00 | 10187.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | EPPlus | 442.64 ms | 12.74 ms | 7.36 ms | 8.62 | 8.62 | 159742.3 KB | 15.68 |  |  | 761.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct | ClosedXML | 815.38 ms | 26.63 ms | 15.38 ms | 15.87 | 15.87 | 496956.9 KB | 48.78 |  |  | 1487.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | OfficeIMO.Excel | 39.87 ms | 6.21 ms | 3.58 ms | 1.00 | 1.00 | 11500.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | EPPlus | 424.75 ms | 1.01 ms | 0.58 ms | 10.65 | 10.65 | 138360.4 KB | 12.03 |  |  | 965.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | AutoFit and mutation | write-insertobjects-partial-autofitcolumnsfor-direct | ClosedXML | 507.98 ms | 17.97 ms | 10.37 ms | 12.74 | 12.74 | 275422.3 KB | 23.95 |  |  | 1174.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | OfficeIMO.Excel | 18.17 ms | 0.44 ms | 0.26 ms | 1.00 | 1.00 | 6896.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | ClosedXML | 150.23 ms | 2.13 ms | 1.23 ms | 8.27 | 8.27 | 92902.1 KB | 13.47 |  |  | 726.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-empty-strings | EPPlus | 190.35 ms | 10.83 ms | 6.25 ms | 10.48 | 10.48 | 74492.8 KB | 10.80 |  |  | 947.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | OfficeIMO.Excel | 21.13 ms | 0.62 ms | 0.36 ms | 1.00 | 1.00 | 5970.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | EPPlus 4.5.3.3 | 101.78 ms |  |  | 4.82 | 4.82 |  |  |  |  | 381.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | ClosedXML | 126.71 ms | 1.69 ms | 0.98 ms | 6.00 | 6.00 | 84206.7 KB | 14.10 |  |  | 499.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-numbers | EPPlus | 222.06 ms | 14.69 ms | 8.48 ms | 10.51 | 10.51 | 86377.5 KB | 14.47 |  |  | 951.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | OfficeIMO.Excel | 26.87 ms | 0.81 ms | 0.46 ms | 1.00 | 1.00 | 8332.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | EPPlus 4.5.3.3 | 120.69 ms |  |  | 4.49 | 4.49 |  |  |  |  | 349.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | ClosedXML | 209.39 ms | 16.66 ms | 9.62 ms | 7.79 | 7.79 | 111118.7 KB | 13.33 |  |  | 679.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-mixed | EPPlus | 253.86 ms | 5.04 ms | 2.91 ms | 9.45 | 9.45 | 113245.1 KB | 13.59 |  |  | 844.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | OfficeIMO.Excel | 25.88 ms | 2.40 ms | 1.39 ms | 1.00 | 1.00 | 7416.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | ClosedXML | 175.49 ms | 12.21 ms | 7.05 ms | 6.78 | 6.78 | 105223.9 KB | 14.19 |  |  | 578.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse | EPPlus | 243.66 ms | 10.81 ms | 6.24 ms | 9.42 | 9.42 | 106316.9 KB | 14.34 |  |  | 841.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | OfficeIMO.Excel | 20.35 ms | 3.18 ms | 1.84 ms | 1.00 | 1.00 | 7416.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | ClosedXML | 183.44 ms | 4.03 ms | 2.33 ms | 9.01 | 9.01 | 105223.9 KB | 14.19 |  |  | 801.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-object-sparse-batch | EPPlus | 249.45 ms | 12.65 ms | 7.30 ms | 12.26 | 12.26 | 106316.9 KB | 14.34 |  |  | 1125.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | OfficeIMO.Excel | 16.44 ms | 0.46 ms | 0.27 ms | 1.00 | 1.00 | 6144.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | EPPlus 4.5.3.3 | 103.46 ms |  |  | 6.29 | 6.29 |  |  |  |  | 529.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | ClosedXML | 114.72 ms | 10.92 ms | 6.30 ms | 6.98 | 6.98 | 82591.3 KB | 13.44 |  |  | 597.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-scalars | EPPlus | 219.56 ms | 16.92 ms | 9.77 ms | 13.35 | 13.35 | 85127.4 KB | 13.85 |  |  | 1235.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | OfficeIMO.Excel | 28.12 ms | 0.68 ms | 0.39 ms | 1.00 | 1.00 | 7482.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | EPPlus 4.5.3.3 | 121.82 ms |  |  | 4.33 | 4.33 |  |  |  |  | 333.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | ClosedXML | 189.26 ms | 8.90 ms | 5.14 ms | 6.73 | 6.73 | 89323.7 KB | 11.94 |  |  | 573.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Cell writer | write-cellvalue-temporal | EPPlus | 223.14 ms | 18.82 ms | 10.87 ms | 7.94 | 7.94 | 103800.0 KB | 13.87 |  |  | 693.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | OfficeIMO.Excel | 56.02 ms | 7.18 ms | 4.15 ms | 1.00 | 1.00 | 13039.6 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | MiniExcel | 132.28 ms | 4.21 ms | 2.43 ms | 2.36 | 2.36 | 97088.3 KB | 7.45 |  |  | 136.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | ClosedXML | 595.67 ms | 1.44 ms | 0.83 ms | 10.63 | 10.63 | 172016.6 KB | 13.19 |  |  | 963.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-headerless-tables | EPPlus | 679.97 ms | 23.43 ms | 13.53 ms | 12.14 | 12.14 | 111246.0 KB | 8.53 |  |  | 1113.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | OfficeIMO.Excel | 54.08 ms | 1.79 ms | 1.03 ms | 1.00 | 1.00 | 13458.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | MiniExcel | 129.31 ms | 3.80 ms | 2.19 ms | 2.39 | 2.39 | 108129.1 KB | 8.03 |  |  | 139.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | EPPlus | 703.23 ms | 19.25 ms | 11.11 ms | 13.00 | 13.00 | 135723.5 KB | 10.08 |  |  | 1200.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-sparse-tables | ClosedXML | 761.84 ms | 66.22 ms | 38.23 ms | 14.09 | 14.09 | 280371.2 KB | 20.83 |  |  | 1308.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | OfficeIMO.Excel | 45.61 ms | 1.82 ms | 1.05 ms | 1.00 | 1.00 | 10288.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | MiniExcel | 94.87 ms | 1.09 ms | 0.63 ms | 2.08 | 2.08 | 97085.4 KB | 9.44 |  |  | 108.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | EPPlus 4.5.3.3 | 217.44 ms |  |  | 4.77 | 4.77 |  |  |  |  | 376.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | EPPlus | 424.85 ms | 35.97 ms | 20.77 ms | 9.31 | 9.31 | 110815.9 KB | 10.77 |  |  | 831.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataSet table export | write-dataset-tables | ClosedXML | 444.96 ms | 15.40 ms | 8.89 ms | 9.75 | 9.75 | 171999.7 KB | 16.72 |  |  | 875.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | OfficeIMO.Excel | 46.10 ms | 3.95 ms | 2.28 ms | 1.00 | 1.00 | 13020.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | MiniExcel | 91.11 ms | 8.11 ms | 4.68 ms | 1.98 | 1.98 | 92200.0 KB | 7.08 |  |  | 97.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | EPPlus 4.5.3.3 | 227.93 ms |  |  | 4.94 | 4.94 |  |  |  |  | 394.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | EPPlus | 394.89 ms | 16.05 ms | 9.26 ms | 8.57 | 8.57 | 117437.3 KB | 9.02 |  |  | 756.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datareader-table | ClosedXML | 480.31 ms | 14.58 ms | 8.42 ms | 10.42 | 10.42 | 173397.5 KB | 13.32 |  |  | 941.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | LargeXlsx | 38.39 ms | 2.63 ms | 1.52 ms | 0.97 | 1.00 | 9520.4 KB | 0.75 |  |  | 3.0% faster than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | OfficeIMO.Excel | 39.58 ms | 5.30 ms | 3.06 ms | 1.00 | 1.03 | 12715.7 KB | 1.00 |  |  | Loss +3.1% |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | MiniExcel | 99.26 ms | 6.90 ms | 3.99 ms | 2.51 | 2.59 | 92394.2 KB | 7.27 |  |  | 150.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | EPPlus 4.5.3.3 | 227.07 ms |  |  | 5.74 | 5.92 |  |  |  |  | 473.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | ClosedXML | 348.73 ms | 13.86 ms | 8.00 ms | 8.81 | 9.08 | 104205.0 KB | 8.19 |  |  | 781.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-direct | EPPlus | 400.16 ms | 14.53 ms | 8.39 ms | 10.11 | 10.42 | 117437.3 KB | 9.24 |  |  | 911.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | OfficeIMO.Excel | 44.19 ms | 2.73 ms | 1.57 ms | 1.00 | 1.00 | 9999.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | MiniExcel | 103.35 ms | 1.53 ms | 0.88 ms | 2.34 | 2.34 | 89659.2 KB | 8.97 |  |  | 133.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | EPPlus | 377.68 ms | 13.36 ms | 7.72 ms | 8.55 | 8.55 | 114703.1 KB | 11.47 |  |  | 754.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-object-table-direct | ClosedXML | 473.81 ms | 8.83 ms | 5.10 ms | 10.72 | 10.72 | 170665.2 KB | 17.07 |  |  | 972.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | OfficeIMO.Excel | 45.13 ms | 0.37 ms | 0.21 ms | 1.00 | 1.00 | 12733.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | MiniExcel | 100.46 ms | 7.13 ms | 4.12 ms | 2.23 | 2.23 | 92394.5 KB | 7.26 |  |  | 122.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | EPPlus 4.5.3.3 | 227.11 ms |  |  | 5.03 | 5.03 |  |  |  |  | 403.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | EPPlus | 411.34 ms | 1.07 ms | 0.62 ms | 9.11 | 9.11 | 117437.3 KB | 9.22 |  |  | 811.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | DataTable table export | write-datatable-table-direct | ClosedXML | 429.33 ms | 29.44 ms | 16.99 ms | 9.51 | 9.51 | 173400.0 KB | 13.62 |  |  | 851.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | OfficeIMO.Excel | 43.87 ms | 2.77 ms | 1.60 ms | 1.00 | 1.00 | 11561.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | MiniExcel | 86.28 ms | 4.46 ms | 2.58 ms | 1.97 | 1.97 | 125551.6 KB | 10.86 |  |  | 96.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | EPPlus | 461.26 ms | 5.72 ms | 3.30 ms | 10.52 | 10.52 | 254959.0 KB | 22.05 |  |  | 951.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | EPPlus 4.5.3.3 | 515.73 ms |  |  | 11.76 | 11.76 |  |  |  |  | 1075.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formatted report write | write-bulk-report | ClosedXML | 939.50 ms | 19.69 ms | 11.37 ms | 21.42 | 21.42 | 565953.9 KB | 48.95 |  |  | 2041.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | OfficeIMO.Excel | 28.29 ms | 0.28 ms | 0.16 ms | 1.00 | 1.00 | 10112.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | EPPlus 4.5.3.3 | 134.02 ms |  |  | 4.74 | 4.74 |  |  |  |  | 373.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | ClosedXML | 199.57 ms | 28.82 ms | 16.64 ms | 7.05 | 7.05 | 113853.5 KB | 11.26 |  |  | 605.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Formula write/read | write-cellformula | EPPlus | 366.58 ms | 51.86 ms | 29.94 ms | 12.96 | 12.96 | 140731.9 KB | 13.92 |  |  | 1195.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-dictionary-objects-table-direct | OfficeIMO.Excel | 57.67 ms | 3.71 ms | 2.14 ms | 1.00 | 1.00 | 15163.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | LargeXlsx | 50.66 ms | 0.38 ms | 0.22 ms | 0.96 | 1.00 | 6043.9 KB | 0.57 |  |  | 3.6% faster than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | OfficeIMO.Excel | 52.53 ms | 7.01 ms | 4.05 ms | 1.00 | 1.04 | 10577.2 KB | 1.00 |  |  | Loss +3.7% |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | MiniExcel | 107.39 ms | 3.30 ms | 1.91 ms | 2.04 | 2.12 | 113974.3 KB | 10.78 |  |  | 104.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | ClosedXML | 440.11 ms | 7.10 ms | 4.10 ms | 8.38 | 8.69 | 179552.5 KB | 16.98 |  |  | 737.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-mixed-objects-direct | EPPlus | 517.02 ms | 22.19 ms | 12.81 ms | 9.84 | 10.21 | 144920.0 KB | 13.70 |  |  | 884.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | LargeXlsx | 45.19 ms | 6.42 ms | 3.70 ms | 0.81 | 1.00 | 6043.9 KB | 0.61 |  |  | 18.6% faster than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | OfficeIMO.Excel | 55.52 ms | 0.43 ms | 0.25 ms | 1.00 | 1.23 | 9942.2 KB | 1.00 |  |  | Loss +22.9% |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | MiniExcel | 104.10 ms | 7.70 ms | 4.44 ms | 1.87 | 2.30 | 113974.3 KB | 11.46 |  |  | 87.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | ClosedXML | 438.68 ms | 11.74 ms | 6.78 ms | 7.90 | 9.71 | 179552.5 KB | 18.06 |  |  | 690.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-mixed-direct | EPPlus | 509.73 ms | 17.29 ms | 9.98 ms | 9.18 | 11.28 | 144920.0 KB | 14.58 |  |  | 818.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | OfficeIMO.Excel | 226.59 ms | 8.57 ms | 4.95 ms | 1.00 | 1.00 | 36150.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | LargeXlsx | 247.61 ms | 9.62 ms | 5.56 ms | 1.09 | 1.09 | 23211.4 KB | 0.64 |  |  | 9.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | MiniExcel | 437.78 ms | 20.02 ms | 11.56 ms | 1.93 | 1.93 | 347925.7 KB | 9.62 |  |  | 93.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | ClosedXML | 1460.91 ms | 31.61 ms | 18.25 ms | 6.45 | 6.45 | 487446.6 KB | 13.48 |  |  | 544.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Other | write-powershell-psobject-wide-direct | EPPlus | 1817.17 ms | 23.42 ms | 13.52 ms | 8.02 | 8.02 | 562916.0 KB | 15.57 |  |  | 702.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | LargeXlsx | 15.01 ms | 0.07 ms | 0.04 ms | 0.77 | 1.00 | 2771.0 KB | 0.26 |  |  | 22.7% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | OfficeIMO.Excel | 19.42 ms | 1.34 ms | 0.78 ms | 1.00 | 1.29 | 10842.5 KB | 1.00 |  |  | Loss +29.4% |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | MiniExcel | 40.00 ms | 2.57 ms | 1.48 ms | 2.06 | 2.67 | 58242.8 KB | 5.37 |  |  | 106.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | EPPlus 4.5.3.3 | 117.61 ms |  |  | 6.06 | 7.84 |  |  |  |  | 505.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | ClosedXML | 150.45 ms | 12.68 ms | 7.32 ms | 7.75 | 10.03 | 104233.1 KB | 9.61 |  |  | 674.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | append-plain-rows | EPPlus | 226.42 ms | 18.27 ms | 10.55 ms | 11.66 | 15.09 | 100373.5 KB | 9.26 |  |  | 1065.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | LargeXlsx | 12.85 ms | 0.38 ms | 0.22 ms | 0.90 | 1.00 | 3444.4 KB | 0.49 |  |  | 10.0% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | OfficeIMO.Excel | 14.28 ms | 1.18 ms | 0.68 ms | 1.00 | 1.11 | 6961.7 KB | 1.00 |  |  | Loss +11.2% |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | ClosedXML | 146.01 ms | 10.25 ms | 5.92 ms | 10.22 | 11.36 | 96015.7 KB | 13.79 |  |  | 922.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-headerless-rectangle-direct | EPPlus | 220.48 ms | 5.70 ms | 3.29 ms | 15.44 | 17.16 | 87467.1 KB | 12.56 |  |  | 1443.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | LargeXlsx | 35.54 ms | 3.59 ms | 2.07 ms | 0.90 | 1.00 | 5614.1 KB | 0.35 |  |  | 10.2% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | OfficeIMO.Excel | 39.59 ms | 6.04 ms | 3.49 ms | 1.00 | 1.11 | 16036.5 KB | 1.00 |  |  | Loss +11.4% |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | MiniExcel | 85.24 ms | 5.29 ms | 3.06 ms | 2.15 | 2.40 | 93257.1 KB | 5.82 |  |  | 115.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | EPPlus 4.5.3.3 | 247.08 ms |  |  | 6.24 | 6.95 |  |  |  |  | 524.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | ClosedXML | 369.50 ms | 11.46 ms | 6.62 ms | 9.33 | 10.40 | 210646.1 KB | 13.14 |  |  | 833.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-rectangle-direct | EPPlus | 443.19 ms | 9.59 ms | 5.54 ms | 11.20 | 12.47 | 211849.9 KB | 13.21 |  |  | 1019.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | OfficeIMO.Excel | 20.83 ms | 1.35 ms | 0.78 ms | 1.00 | 1.00 | 7866.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | ClosedXML | 178.41 ms | 6.76 ms | 3.90 ms | 8.56 | 8.56 | 105223.9 KB | 13.38 |  |  | 756.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain cell export | write-cellvalues-sparse-rectangle-direct | EPPlus | 250.41 ms | 14.74 ms | 8.51 ms | 12.02 | 12.02 | 106316.9 KB | 13.52 |  |  | 1102.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | Sylvan.Data.Excel | 33.41 ms | 5.53 ms | 3.19 ms | 0.81 | 1.00 | 5700.3 KB | 0.44 |  |  | 19.4% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | OfficeIMO.Excel | 41.47 ms | 4.93 ms | 2.85 ms | 1.00 | 1.24 | 13002.3 KB | 1.00 |  |  | Loss +24.1% |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | LargeXlsx | 43.55 ms | 5.28 ms | 3.05 ms | 1.05 | 1.30 | 8349.2 KB | 0.64 |  |  | 5.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | MiniExcel | 87.81 ms | 8.99 ms | 5.19 ms | 2.12 | 2.63 | 92199.8 KB | 7.09 |  |  | 111.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | EPPlus 4.5.3.3 | 235.54 ms |  |  | 5.68 | 7.05 |  |  |  |  | 468.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | ClosedXML | 347.17 ms | 19.21 ms | 11.09 ms | 8.37 | 10.39 | 104205.0 KB | 8.01 |  |  | 737.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-datareader-plain | EPPlus | 394.28 ms | 14.40 ms | 8.31 ms | 9.51 | 11.80 | 117437.7 KB | 9.03 |  |  | 850.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | OfficeIMO.Excel | 45.09 ms | 5.81 ms | 3.35 ms | 1.00 | 1.00 | 9819.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | LargeXlsx | 51.49 ms | 3.12 ms | 1.80 ms | 1.14 | 1.14 | 9265.9 KB | 0.94 |  |  | 14.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | MiniExcel | 120.11 ms | 10.23 ms | 5.91 ms | 2.66 | 2.66 | 108129.1 KB | 11.01 |  |  | 166.4% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | EPPlus | 659.20 ms | 110.75 ms | 63.94 ms | 14.62 | 14.62 | 135723.5 KB | 13.82 |  |  | 1362.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain streaming export | write-dataset-sparse-direct-export | ClosedXML | 686.28 ms | 71.23 ms | 41.12 ms | 15.22 | 15.22 | 280372.7 KB | 28.55 |  |  | 1422.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | LargeXlsx | 51.12 ms | 7.40 ms | 4.27 ms | 0.89 | 1.00 | 10795.2 KB | 0.92 |  |  | 10.7% faster than OfficeIMO |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | OfficeIMO.Excel | 57.24 ms | 10.48 ms | 6.05 ms | 1.00 | 1.12 | 11708.2 KB | 1.00 |  |  | Loss +12.0% |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | MiniExcel | 191.91 ms | 11.08 ms | 6.40 ms | 3.35 | 3.75 | 226875.6 KB | 19.38 |  |  | 235.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Plain string export | write-blog-2023-20-string-columns | ClosedXML | 1091.01 ms | 48.85 ms | 28.20 ms | 19.06 | 21.34 | 759818.4 KB | 64.90 |  |  | 1806.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | OfficeIMO.Excel | 14.54 ms | 0.22 ms | 0.13 ms | 1.00 | 1.00 | 15744.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | MiniExcel | 28.29 ms | 1.37 ms | 0.79 ms | 1.95 | 1.95 | 73760.2 KB | 4.68 |  |  | 94.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | EPPlus 4.5.3.3 | 89.78 ms |  |  | 6.17 | 6.17 |  |  |  |  | 517.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | ClosedXML | 110.05 ms | 0.69 ms | 0.40 ms | 7.57 | 7.57 | 104241.3 KB | 6.62 |  |  | 656.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | large-shared-strings | EPPlus | 183.41 ms | 5.32 ms | 3.07 ms | 12.61 | 12.61 | 84410.0 KB | 5.36 |  |  | 1161.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | OfficeIMO.Excel | 23.58 ms | 1.09 ms | 0.63 ms | 1.00 | 1.00 | 15360.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | EPPlus 4.5.3.3 | 100.17 ms |  |  | 4.25 | 4.25 |  |  |  |  | 324.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | ClosedXML | 133.86 ms | 15.44 ms | 8.91 ms | 5.68 | 5.68 | 104241.3 KB | 6.79 |  |  | 467.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings | EPPlus | 214.83 ms | 5.61 ms | 3.24 ms | 9.11 | 9.11 | 84410.5 KB | 5.50 |  |  | 811.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | OfficeIMO.Excel | 17.30 ms | 3.91 ms | 2.26 ms | 1.00 | 1.00 | 13824.1 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | ClosedXML | 186.98 ms | 13.62 ms | 7.86 ms | 10.81 | 10.81 | 131501.7 KB | 9.51 |  |  | 980.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-distinct | EPPlus | 249.13 ms | 12.93 ms | 7.47 ms | 14.40 | 14.40 | 97729.6 KB | 7.07 |  |  | 1339.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | OfficeIMO.Excel | 18.09 ms | 0.78 ms | 0.45 ms | 1.00 | 1.00 | 7525.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | ClosedXML | 123.25 ms | 6.83 ms | 3.95 ms | 6.81 | 6.81 | 84520.0 KB | 11.23 |  |  | 581.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Shared string write | write-cellvalue-strings-repeated | EPPlus | 188.32 ms | 2.21 ms | 1.28 ms | 10.41 | 10.41 | 70033.4 KB | 9.31 |  |  | 941.0% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | LargeXlsx | 34.79 ms | 2.18 ms | 1.26 ms | 0.88 | 1.00 | 5614.1 KB | 0.43 |  |  | 12.3% faster than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | OfficeIMO.Excel | 39.68 ms | 4.25 ms | 2.46 ms | 1.00 | 1.14 | 12912.0 KB | 1.00 |  |  | Loss +14.1% |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | MiniExcel | 70.71 ms | 9.03 ms | 5.21 ms | 1.78 | 2.03 | 93257.1 KB | 7.22 |  |  | 78.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | EPPlus 4.5.3.3 | 253.69 ms |  |  | 6.39 | 7.29 |  |  |  |  | 539.3% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | ClosedXML | 323.67 ms | 45.75 ms | 26.41 ms | 8.16 | 9.30 | 104205.0 KB | 8.07 |  |  | 715.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-fluent-rowsfrom-direct | EPPlus | 388.83 ms | 21.89 ms | 12.64 ms | 9.80 | 11.18 | 117437.7 KB | 9.10 |  |  | 879.8% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | LargeXlsx | 33.17 ms | 2.67 ms | 1.54 ms | 0.85 | 1.00 | 5614.1 KB | 0.49 |  |  | 15.4% faster than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | OfficeIMO.Excel | 39.19 ms | 3.78 ms | 2.18 ms | 1.00 | 1.18 | 11493.8 KB | 1.00 |  |  | Loss +18.2% |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | MiniExcel | 74.70 ms | 8.92 ms | 5.15 ms | 1.91 | 2.25 | 93257.1 KB | 8.11 |  |  | 90.6% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | EPPlus 4.5.3.3 | 235.18 ms |  |  | 6.00 | 7.09 |  |  |  |  | 500.1% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | ClosedXML | 347.88 ms | 5.84 ms | 3.37 ms | 8.88 | 10.49 | 104205.0 KB | 9.07 |  |  | 787.7% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-direct | EPPlus | 396.57 ms | 9.39 ms | 5.42 ms | 10.12 | 11.96 | 117437.3 KB | 10.22 |  |  | 911.9% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | LargeXlsx | 33.47 ms | 5.58 ms | 3.22 ms | 0.70 | 1.00 | 5614.1 KB | 0.55 |  |  | 29.9% faster than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | OfficeIMO.Excel | 47.75 ms | 2.07 ms | 1.20 ms | 1.00 | 1.43 | 10179.4 KB | 1.00 |  |  | Loss +42.6% |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | MiniExcel | 79.99 ms | 0.67 ms | 0.39 ms | 1.68 | 2.39 | 93257.1 KB | 9.16 |  |  | 67.5% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | ClosedXML | 337.67 ms | 16.36 ms | 9.44 ms | 7.07 | 10.09 | 104205.0 KB | 10.24 |  |  | 607.2% slower than OfficeIMO |
| 25000 | speed-comparison | write | Typed object export | write-insertobjects-flat-dictionaries-direct | EPPlus | 378.78 ms | 9.07 ms | 5.23 ms | 7.93 | 11.32 | 117437.3 KB | 11.54 |  |  | 693.3% slower than OfficeIMO |
