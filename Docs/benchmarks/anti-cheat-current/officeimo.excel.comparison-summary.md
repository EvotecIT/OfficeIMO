# OfficeIMO.Excel Comparison Summary

This is the suite-level decision table. Mean, standard deviation, standard error, ratios, and allocations come from the lightweight rotated local runner; they are meant for engineering direction. Use the BenchmarkDotNet benchmark classes when a publication-grade `Error` column is required.

## At a glance

| Row count | Artifact | Workload | Category | OfficeIMO wins | OfficeIMO losses | Biggest loss |
| ---: | --- | --- | --- | ---: | ---: | --- |
| 100 | package-profile | package | Package size | 6 | 0 |  |
| 100 | speed-comparison | mutate | AutoFit and mutation | 1 | 0 |  |
| 100 | speed-comparison | other | Anti-cheat report variants | 4 | 0 |  |
| 100 | speed-comparison | other | Real-world report | 1 | 0 |  |
| 2500 | package-profile | package | Package size | 6 | 0 |  |
| 2500 | speed-comparison | mutate | AutoFit and mutation | 1 | 0 |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | 4 | 0 |  |
| 2500 | speed-comparison | other | Real-world report | 1 | 0 |  |
| 25000 | package-profile | package | Package size | 6 | 0 |  |
| 25000 | speed-comparison | mutate | AutoFit and mutation | 1 | 0 |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | 4 | 0 |  |
| 25000 | speed-comparison | other | Real-world report | 1 | 0 |  |

## OfficeIMO decision table

| Row count | Artifact | Workload | Category | Scenario | OfficeIMO mean | Best | OfficeIMO vs best | Alloc | Package |
| ---: | --- | --- | --- | --- | ---: | --- | ---: | ---: | ---: |
| 100 | package-profile | package | Package size | realworld-report-all-in-one | 6.12 ms | OfficeIMO.Excel | Win | 1493.4 KB | 19.8 KB |
| 100 | package-profile | package | Package size | realworld-report-chart-first | 5.19 ms | OfficeIMO.Excel | Win | 1493.6 KB | 19.8 KB |
| 100 | package-profile | package | Package size | realworld-report-extra-column | 4.67 ms | OfficeIMO.Excel | Win | 1535.8 KB | 20.2 KB |
| 100 | package-profile | package | Package size | realworld-report-no-autofit | 5.27 ms | OfficeIMO.Excel | Win | 1486.2 KB | 19.6 KB |
| 100 | package-profile | package | Package size | realworld-report-post-mutation | 3.70 ms | OfficeIMO.Excel | Win | 1502.7 KB | 19.8 KB |
| 100 | package-profile | package | Package size | realworld-report-shuffled-columns | 3.99 ms | OfficeIMO.Excel | Win | 1491.7 KB | 19.7 KB |
| 100 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | 4.82 ms | OfficeIMO.Excel | Win | 1490.4 KB |  |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | 5.73 ms | OfficeIMO.Excel | Win | 1497.7 KB |  |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | 7.79 ms | OfficeIMO.Excel | Win | 1539.9 KB |  |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | 6.91 ms | OfficeIMO.Excel | Win | 1506.8 KB |  |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | 5.68 ms | OfficeIMO.Excel | Win | 1495.9 KB |  |
| 100 | speed-comparison | other | Real-world report | realworld-report-all-in-one | 4.53 ms | OfficeIMO.Excel | Win | 1497.5 KB |  |
| 2500 | package-profile | package | Package size | realworld-report-all-in-one | 23.19 ms | OfficeIMO.Excel | Win | 10672.5 KB | 210.0 KB |
| 2500 | package-profile | package | Package size | realworld-report-chart-first | 20.01 ms | OfficeIMO.Excel | Win | 10672.8 KB | 210.0 KB |
| 2500 | package-profile | package | Package size | realworld-report-extra-column | 18.80 ms | OfficeIMO.Excel | Win | 11609.3 KB | 226.2 KB |
| 2500 | package-profile | package | Package size | realworld-report-no-autofit | 19.15 ms | OfficeIMO.Excel | Win | 10666.4 KB | 209.9 KB |
| 2500 | package-profile | package | Package size | realworld-report-post-mutation | 17.54 ms | OfficeIMO.Excel | Win | 10681.6 KB | 210.1 KB |
| 2500 | package-profile | package | Package size | realworld-report-shuffled-columns | 16.17 ms | OfficeIMO.Excel | Win | 10677.1 KB | 214.5 KB |
| 2500 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | 18.68 ms | OfficeIMO.Excel | Win | 10666.4 KB |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | 24.50 ms | OfficeIMO.Excel | Win | 10672.8 KB |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | 22.92 ms | OfficeIMO.Excel | Win | 11609.5 KB |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | 20.15 ms | OfficeIMO.Excel | Win | 10681.5 KB |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | 19.78 ms | OfficeIMO.Excel | Win | 10677.3 KB |  |
| 2500 | speed-comparison | other | Real-world report | realworld-report-all-in-one | 19.57 ms | OfficeIMO.Excel | Win | 10673.0 KB |  |
| 25000 | package-profile | package | Package size | realworld-report-all-in-one | 47.16 ms | OfficeIMO.Excel | Win | 15836.9 KB | 1437.1 KB |
| 25000 | package-profile | package | Package size | realworld-report-chart-first | 46.06 ms | OfficeIMO.Excel | Win | 15836.5 KB | 1437.1 KB |
| 25000 | package-profile | package | Package size | realworld-report-extra-column | 49.37 ms | OfficeIMO.Excel | Win | 16128.6 KB | 1531.5 KB |
| 25000 | package-profile | package | Package size | realworld-report-no-autofit | 42.54 ms | OfficeIMO.Excel | Win | 15829.3 KB | 1437.0 KB |
| 25000 | package-profile | package | Package size | realworld-report-post-mutation | 49.06 ms | OfficeIMO.Excel | Win | 15846.8 KB | 1437.2 KB |
| 25000 | package-profile | package | Package size | realworld-report-shuffled-columns | 46.97 ms | OfficeIMO.Excel | Win | 15816.2 KB | 1415.3 KB |
| 25000 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | 44.16 ms | OfficeIMO.Excel | Win | 15829.9 KB |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | 49.94 ms | OfficeIMO.Excel | Win | 15837.3 KB |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | 50.50 ms | OfficeIMO.Excel | Win | 16129.4 KB |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | 47.58 ms | OfficeIMO.Excel | Win | 15844.7 KB |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | 45.02 ms | OfficeIMO.Excel | Win | 15815.8 KB |  |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | 47.98 ms | OfficeIMO.Excel | Win | 15837.3 KB |  |

## Full comparison table

| Row count | Artifact | Workload | Category | Scenario | Library | Mean | StdDev | StdErr | Ratio to OfficeIMO | Ratio to best | Alloc | Alloc ratio | Package | Package ratio | Outcome |
| ---: | --- | --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- |
| 100 | package-profile | package | Package size | realworld-report-all-in-one | OfficeIMO.Excel | 6.12 ms | 0.57 ms | 0.25 ms | 1.00 | 1.00 | 1493.4 KB | 1.00 | 19.8 KB | 1.00 | Win |
| 100 | package-profile | package | Package size | realworld-report-all-in-one | EPPlus | 27.20 ms | 3.29 ms | 1.47 ms | 4.44 | 4.44 | 13576.5 KB | 9.09 | 16.6 KB | 0.84 | 344.2% slower than OfficeIMO |
| 100 | package-profile | package | Package size | realworld-report-chart-first | OfficeIMO.Excel | 5.19 ms | 0.85 ms | 0.38 ms | 1.00 | 1.00 | 1493.6 KB | 1.00 | 19.8 KB | 1.00 | Win |
| 100 | package-profile | package | Package size | realworld-report-chart-first | EPPlus | 24.52 ms | 3.36 ms | 1.50 ms | 4.73 | 4.73 | 13575.2 KB | 9.09 | 16.6 KB | 0.84 | 372.5% slower than OfficeIMO |
| 100 | package-profile | package | Package size | realworld-report-extra-column | OfficeIMO.Excel | 4.67 ms | 1.00 ms | 0.45 ms | 1.00 | 1.00 | 1535.8 KB | 1.00 | 20.2 KB | 1.00 | Win |
| 100 | package-profile | package | Package size | realworld-report-extra-column | EPPlus | 25.94 ms | 1.49 ms | 0.67 ms | 5.55 | 5.55 | 13865.8 KB | 9.03 | 17.0 KB | 0.84 | 455.2% slower than OfficeIMO |
| 100 | package-profile | package | Package size | realworld-report-no-autofit | OfficeIMO.Excel | 5.27 ms | 0.68 ms | 0.30 ms | 1.00 | 1.00 | 1486.2 KB | 1.00 | 19.6 KB | 1.00 | Win |
| 100 | package-profile | package | Package size | realworld-report-no-autofit | EPPlus | 18.89 ms | 1.26 ms | 0.57 ms | 3.59 | 3.59 | 12158.9 KB | 8.18 | 16.5 KB | 0.84 | 258.6% slower than OfficeIMO |
| 100 | package-profile | package | Package size | realworld-report-post-mutation | OfficeIMO.Excel | 3.70 ms | 0.25 ms | 0.11 ms | 1.00 | 1.00 | 1502.7 KB | 1.00 | 19.8 KB | 1.00 | Win |
| 100 | package-profile | package | Package size | realworld-report-post-mutation | EPPlus | 19.24 ms | 0.74 ms | 0.33 ms | 5.20 | 5.20 | 13576.0 KB | 9.03 | 16.6 KB | 0.84 | 420.3% slower than OfficeIMO |
| 100 | package-profile | package | Package size | realworld-report-shuffled-columns | OfficeIMO.Excel | 3.99 ms | 0.26 ms | 0.11 ms | 1.00 | 1.00 | 1491.7 KB | 1.00 | 19.7 KB | 1.00 | Win |
| 100 | package-profile | package | Package size | realworld-report-shuffled-columns | EPPlus | 24.37 ms | 1.23 ms | 0.55 ms | 6.11 | 6.11 | 13570.5 KB | 9.10 | 16.6 KB | 0.85 | 511.2% slower than OfficeIMO |
| 100 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | OfficeIMO.Excel | 4.82 ms | 0.30 ms | 0.14 ms | 1.00 | 1.00 | 1490.4 KB | 1.00 |  |  | Win |
| 100 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus 4.5.3.3 | 5.90 ms |  |  | 1.22 | 1.22 |  |  |  |  | 22.4% slower than OfficeIMO |
| 100 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus | 14.60 ms | 1.44 ms | 0.64 ms | 3.03 | 3.03 | 12158.9 KB | 8.16 |  |  | 202.8% slower than OfficeIMO |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | OfficeIMO.Excel | 5.73 ms | 1.47 ms | 0.66 ms | 1.00 | 1.00 | 1497.7 KB | 1.00 |  |  | Win |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus 4.5.3.3 | 9.24 ms |  |  | 1.61 | 1.61 |  |  |  |  | 61.3% slower than OfficeIMO |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus | 17.88 ms | 1.57 ms | 0.70 ms | 3.12 | 3.12 | 13575.3 KB | 9.06 |  |  | 212.1% slower than OfficeIMO |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | OfficeIMO.Excel | 7.79 ms | 3.21 ms | 1.44 ms | 1.00 | 1.00 | 1539.9 KB | 1.00 |  |  | Win |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus 4.5.3.3 | 9.17 ms |  |  | 1.18 | 1.18 |  |  |  |  | 17.7% slower than OfficeIMO |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus | 18.38 ms | 1.55 ms | 0.69 ms | 2.36 | 2.36 | 13865.8 KB | 9.00 |  |  | 135.8% slower than OfficeIMO |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | OfficeIMO.Excel | 6.91 ms | 0.18 ms | 0.08 ms | 1.00 | 1.00 | 1506.8 KB | 1.00 |  |  | Win |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus 4.5.3.3 | 8.20 ms |  |  | 1.19 | 1.19 |  |  |  |  | 18.6% slower than OfficeIMO |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus | 17.24 ms | 1.40 ms | 0.63 ms | 2.49 | 2.49 | 13576.1 KB | 9.01 |  |  | 149.5% slower than OfficeIMO |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | OfficeIMO.Excel | 5.68 ms | 0.52 ms | 0.23 ms | 1.00 | 1.00 | 1495.9 KB | 1.00 |  |  | Win |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus 4.5.3.3 | 8.90 ms |  |  | 1.57 | 1.57 |  |  |  |  | 56.6% slower than OfficeIMO |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus | 18.18 ms | 3.09 ms | 1.38 ms | 3.20 | 3.20 | 13570.6 KB | 9.07 |  |  | 219.9% slower than OfficeIMO |
| 100 | speed-comparison | other | Real-world report | realworld-report-all-in-one | OfficeIMO.Excel | 4.53 ms | 0.21 ms | 0.10 ms | 1.00 | 1.00 | 1497.5 KB | 1.00 |  |  | Win |
| 100 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus 4.5.3.3 | 8.56 ms |  |  | 1.89 | 1.89 |  |  |  |  | 89.2% slower than OfficeIMO |
| 100 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus | 15.59 ms | 0.56 ms | 0.25 ms | 3.44 | 3.44 | 13576.1 KB | 9.07 |  |  | 244.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-all-in-one | OfficeIMO.Excel | 23.19 ms | 1.38 ms | 0.62 ms | 1.00 | 1.00 | 10672.5 KB | 1.00 | 210.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-all-in-one | EPPlus | 76.60 ms | 5.50 ms | 2.46 ms | 3.30 | 3.30 | 54591.8 KB | 5.12 | 121.8 KB | 0.58 | 230.3% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-chart-first | OfficeIMO.Excel | 20.01 ms | 2.58 ms | 1.15 ms | 1.00 | 1.00 | 10672.8 KB | 1.00 | 210.0 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-chart-first | EPPlus | 76.37 ms | 5.98 ms | 2.67 ms | 3.82 | 3.82 | 54591.3 KB | 5.11 | 121.8 KB | 0.58 | 281.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-extra-column | OfficeIMO.Excel | 18.80 ms | 2.54 ms | 1.14 ms | 1.00 | 1.00 | 11609.3 KB | 1.00 | 226.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-extra-column | EPPlus | 72.55 ms | 2.44 ms | 1.09 ms | 3.86 | 3.86 | 59224.4 KB | 5.10 | 128.4 KB | 0.57 | 285.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-no-autofit | OfficeIMO.Excel | 19.15 ms | 2.22 ms | 0.99 ms | 1.00 | 1.00 | 10666.4 KB | 1.00 | 209.9 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-no-autofit | EPPlus | 64.31 ms | 5.05 ms | 2.26 ms | 3.36 | 3.36 | 32905.9 KB | 3.09 | 121.8 KB | 0.58 | 235.8% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-post-mutation | OfficeIMO.Excel | 17.54 ms | 0.90 ms | 0.40 ms | 1.00 | 1.00 | 10681.6 KB | 1.00 | 210.1 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-post-mutation | EPPlus | 72.21 ms | 4.45 ms | 1.99 ms | 4.12 | 4.12 | 54592.0 KB | 5.11 | 121.9 KB | 0.58 | 311.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-shuffled-columns | OfficeIMO.Excel | 16.17 ms | 1.27 ms | 0.57 ms | 1.00 | 1.00 | 10677.1 KB | 1.00 | 214.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-shuffled-columns | EPPlus | 69.48 ms | 3.17 ms | 1.42 ms | 4.30 | 4.30 | 54588.6 KB | 5.11 | 124.3 KB | 0.58 | 329.8% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | OfficeIMO.Excel | 18.68 ms | 1.57 ms | 0.70 ms | 1.00 | 1.00 | 10666.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus 4.5.3.3 | 40.00 ms |  |  | 2.14 | 2.14 |  |  |  |  | 114.1% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus | 66.29 ms | 6.89 ms | 3.08 ms | 3.55 | 3.55 | 32907.1 KB | 3.09 |  |  | 254.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | OfficeIMO.Excel | 24.50 ms | 5.28 ms | 2.36 ms | 1.00 | 1.00 | 10672.8 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus 4.5.3.3 | 79.31 ms |  |  | 3.24 | 3.24 |  |  |  |  | 223.8% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus | 83.02 ms | 9.21 ms | 4.12 ms | 3.39 | 3.39 | 54591.7 KB | 5.12 |  |  | 238.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | OfficeIMO.Excel | 22.92 ms | 3.91 ms | 1.75 ms | 1.00 | 1.00 | 11609.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus | 83.47 ms | 3.40 ms | 1.52 ms | 3.64 | 3.64 | 59224.2 KB | 5.10 |  |  | 264.1% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus 4.5.3.3 | 97.39 ms |  |  | 4.25 | 4.25 |  |  |  |  | 324.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | OfficeIMO.Excel | 20.15 ms | 1.64 ms | 0.73 ms | 1.00 | 1.00 | 10681.5 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus | 72.73 ms | 2.64 ms | 1.18 ms | 3.61 | 3.61 | 54591.9 KB | 5.11 |  |  | 260.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus 4.5.3.3 | 78.84 ms |  |  | 3.91 | 3.91 |  |  |  |  | 291.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | OfficeIMO.Excel | 19.78 ms | 1.86 ms | 0.83 ms | 1.00 | 1.00 | 10677.3 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus | 83.00 ms | 9.53 ms | 4.26 ms | 4.20 | 4.20 | 54588.6 KB | 5.11 |  |  | 319.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus 4.5.3.3 | 102.12 ms |  |  | 5.16 | 5.16 |  |  |  |  | 416.3% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world report | realworld-report-all-in-one | OfficeIMO.Excel | 19.57 ms | 3.34 ms | 1.49 ms | 1.00 | 1.00 | 10673.0 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus | 93.08 ms | 14.28 ms | 6.38 ms | 4.76 | 4.76 | 55115.2 KB | 5.16 |  |  | 375.6% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus 4.5.3.3 | 118.73 ms |  |  | 6.07 | 6.07 |  |  |  |  | 506.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-all-in-one | OfficeIMO.Excel | 47.16 ms | 4.19 ms | 1.88 ms | 1.00 | 1.00 | 15836.9 KB | 1.00 | 1437.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-all-in-one | EPPlus | 491.41 ms | 14.01 ms | 6.27 ms | 10.42 | 10.42 | 277074.5 KB | 17.50 | 1097.7 KB | 0.76 | 942.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-chart-first | OfficeIMO.Excel | 46.06 ms | 4.09 ms | 1.83 ms | 1.00 | 1.00 | 15836.5 KB | 1.00 | 1437.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-chart-first | EPPlus | 478.18 ms | 20.22 ms | 9.04 ms | 10.38 | 10.38 | 277073.0 KB | 17.50 | 1097.7 KB | 0.76 | 938.1% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-extra-column | OfficeIMO.Excel | 49.37 ms | 1.80 ms | 0.80 ms | 1.00 | 1.00 | 16128.6 KB | 1.00 | 1531.5 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-extra-column | EPPlus | 551.81 ms | 19.34 ms | 8.65 ms | 11.18 | 11.18 | 302758.6 KB | 18.77 | 1166.3 KB | 0.76 | 1017.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-no-autofit | OfficeIMO.Excel | 42.54 ms | 1.57 ms | 0.70 ms | 1.00 | 1.00 | 15829.3 KB | 1.00 | 1437.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-no-autofit | EPPlus | 441.19 ms | 8.80 ms | 3.93 ms | 10.37 | 10.37 | 234780.3 KB | 14.83 | 1097.7 KB | 0.76 | 937.2% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-post-mutation | OfficeIMO.Excel | 49.06 ms | 3.56 ms | 1.59 ms | 1.00 | 1.00 | 15846.8 KB | 1.00 | 1437.2 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-post-mutation | EPPlus | 519.87 ms | 41.32 ms | 18.48 ms | 10.60 | 10.60 | 277074.0 KB | 17.48 | 1097.8 KB | 0.76 | 959.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-shuffled-columns | OfficeIMO.Excel | 46.97 ms | 5.01 ms | 2.24 ms | 1.00 | 1.00 | 15816.2 KB | 1.00 | 1415.3 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-shuffled-columns | EPPlus | 511.96 ms | 16.33 ms | 7.30 ms | 10.90 | 10.90 | 277068.6 KB | 17.52 | 1098.4 KB | 0.78 | 990.0% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | OfficeIMO.Excel | 44.16 ms | 2.78 ms | 1.24 ms | 1.00 | 1.00 | 15829.9 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus 4.5.3.3 | 283.91 ms |  |  | 6.43 | 6.43 |  |  |  |  | 543.0% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus | 473.37 ms | 17.30 ms | 7.74 ms | 10.72 | 10.72 | 234780.3 KB | 14.83 |  |  | 972.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | OfficeIMO.Excel | 49.94 ms | 1.74 ms | 0.78 ms | 1.00 | 1.00 | 15837.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus | 510.95 ms | 15.95 ms | 7.14 ms | 10.23 | 10.23 | 277073.0 KB | 17.49 |  |  | 923.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus 4.5.3.3 | 573.46 ms |  |  | 11.48 | 11.48 |  |  |  |  | 1048.2% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | OfficeIMO.Excel | 50.50 ms | 6.37 ms | 2.85 ms | 1.00 | 1.00 | 16129.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus | 551.64 ms | 11.40 ms | 5.10 ms | 10.92 | 10.92 | 302758.9 KB | 18.77 |  |  | 992.4% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus 4.5.3.3 | 632.50 ms |  |  | 12.52 | 12.52 |  |  |  |  | 1152.5% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | OfficeIMO.Excel | 47.58 ms | 3.44 ms | 1.54 ms | 1.00 | 1.00 | 15844.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus | 516.59 ms | 20.50 ms | 9.17 ms | 10.86 | 10.86 | 277074.0 KB | 17.49 |  |  | 985.8% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus 4.5.3.3 | 613.64 ms |  |  | 12.90 | 12.90 |  |  |  |  | 1189.8% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | OfficeIMO.Excel | 45.02 ms | 2.53 ms | 1.13 ms | 1.00 | 1.00 | 15815.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus | 485.37 ms | 8.38 ms | 3.75 ms | 10.78 | 10.78 | 277068.7 KB | 17.52 |  |  | 978.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus 4.5.3.3 | 568.83 ms |  |  | 12.63 | 12.63 |  |  |  |  | 1163.5% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | OfficeIMO.Excel | 47.98 ms | 7.15 ms | 3.20 ms | 1.00 | 1.00 | 15837.3 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus | 518.56 ms | 45.82 ms | 20.49 ms | 10.81 | 10.81 | 277074.1 KB | 17.50 |  |  | 980.8% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus 4.5.3.3 | 565.06 ms |  |  | 11.78 | 11.78 |  |  |  |  | 1077.8% slower than OfficeIMO |
