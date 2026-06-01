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
| 100 | package-profile | package | Package size | realworld-report-all-in-one | 4.27 ms | OfficeIMO.Excel | Win | 1846.4 KB | 19.8 KB |
| 100 | package-profile | package | Package size | realworld-report-chart-first | 4.53 ms | OfficeIMO.Excel | Win | 1859.5 KB | 19.8 KB |
| 100 | package-profile | package | Package size | realworld-report-extra-column | 4.38 ms | OfficeIMO.Excel | Win | 1859.0 KB | 20.2 KB |
| 100 | package-profile | package | Package size | realworld-report-no-autofit | 4.35 ms | OfficeIMO.Excel | Win | 1836.4 KB | 19.6 KB |
| 100 | package-profile | package | Package size | realworld-report-post-mutation | 4.46 ms | OfficeIMO.Excel | Win | 1853.8 KB | 19.8 KB |
| 100 | package-profile | package | Package size | realworld-report-shuffled-columns | 5.23 ms | OfficeIMO.Excel | Win | 1843.3 KB | 19.6 KB |
| 100 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | 3.71 ms | OfficeIMO.Excel | Win | 1837.7 KB |  |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | 5.08 ms | OfficeIMO.Excel | Win | 1858.1 KB |  |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | 3.75 ms | OfficeIMO.Excel | Win | 1859.0 KB |  |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | 3.76 ms | OfficeIMO.Excel | Win | 1853.8 KB |  |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | 5.18 ms | OfficeIMO.Excel | Win | 1843.2 KB |  |
| 100 | speed-comparison | other | Real-world report | realworld-report-all-in-one | 4.16 ms | OfficeIMO.Excel | Win | 1875.6 KB |  |
| 2500 | package-profile | package | Package size | realworld-report-all-in-one | 11.24 ms | OfficeIMO.Excel | Win | 6196.8 KB | 206.5 KB |
| 2500 | package-profile | package | Package size | realworld-report-chart-first | 11.13 ms | OfficeIMO.Excel | Win | 6195.8 KB | 206.5 KB |
| 2500 | package-profile | package | Package size | realworld-report-extra-column | 12.27 ms | OfficeIMO.Excel | Win | 6390.7 KB | 219.1 KB |
| 2500 | package-profile | package | Package size | realworld-report-no-autofit | 11.00 ms | OfficeIMO.Excel | Win | 6190.0 KB | 206.4 KB |
| 2500 | package-profile | package | Package size | realworld-report-post-mutation | 11.96 ms | OfficeIMO.Excel | Win | 6206.5 KB | 206.6 KB |
| 2500 | package-profile | package | Package size | realworld-report-shuffled-columns | 12.70 ms | OfficeIMO.Excel | Win | 6202.9 KB | 211.2 KB |
| 2500 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | 23.06 ms | OfficeIMO.Excel | Win | 6196.2 KB |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | 23.93 ms | OfficeIMO.Excel | Win | 6198.1 KB |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | 25.00 ms | OfficeIMO.Excel | Win | 6389.4 KB |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | 25.24 ms | OfficeIMO.Excel | Win | 6209.7 KB |  |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | 24.51 ms | OfficeIMO.Excel | Win | 6201.9 KB |  |
| 2500 | speed-comparison | other | Real-world report | realworld-report-all-in-one | 20.81 ms | OfficeIMO.Excel | Win | 6201.0 KB |  |
| 25000 | package-profile | package | Package size | realworld-report-all-in-one | 87.86 ms | OfficeIMO.Excel | Win | 43678.5 KB | 1985.9 KB |
| 25000 | package-profile | package | Package size | realworld-report-chart-first | 85.98 ms | OfficeIMO.Excel | Win | 43560.9 KB | 1985.9 KB |
| 25000 | package-profile | package | Package size | realworld-report-extra-column | 97.36 ms | OfficeIMO.Excel | Win | 45563.8 KB | 2110.8 KB |
| 25000 | package-profile | package | Package size | realworld-report-no-autofit | 84.07 ms | OfficeIMO.Excel | Win | 43670.7 KB | 1985.8 KB |
| 25000 | package-profile | package | Package size | realworld-report-post-mutation | 88.70 ms | OfficeIMO.Excel | Win | 43688.6 KB | 1986.0 KB |
| 25000 | package-profile | package | Package size | realworld-report-shuffled-columns | 98.94 ms | OfficeIMO.Excel | Win | 43740.6 KB | 2046.1 KB |
| 25000 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | 85.59 ms | OfficeIMO.Excel | Win | 43670.8 KB |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | 90.64 ms | OfficeIMO.Excel | Win | 43560.4 KB |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | 97.45 ms | OfficeIMO.Excel | Win | 45565.2 KB |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | 100.78 ms | OfficeIMO.Excel | Win | 43689.8 KB |  |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | 104.78 ms | OfficeIMO.Excel | Win | 43736.7 KB |  |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | 87.32 ms | OfficeIMO.Excel | Win | 43679.5 KB |  |

## Full comparison table

| Row count | Artifact | Workload | Category | Scenario | Library | Mean | StdDev | StdErr | Ratio to OfficeIMO | Ratio to best | Alloc | Alloc ratio | Package | Package ratio | Outcome |
| ---: | --- | --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- |
| 100 | package-profile | package | Package size | realworld-report-all-in-one | OfficeIMO.Excel | 4.27 ms | 0.24 ms | 0.14 ms | 1.00 | 1.00 | 1846.4 KB | 1.00 | 19.8 KB | 1.00 | Win |
| 100 | package-profile | package | Package size | realworld-report-all-in-one | EPPlus | 16.38 ms | 1.60 ms | 0.92 ms | 3.84 | 3.84 | 13576.2 KB | 7.35 | 16.6 KB | 0.84 | 284.0% slower than OfficeIMO |
| 100 | package-profile | package | Package size | realworld-report-chart-first | OfficeIMO.Excel | 4.53 ms | 0.19 ms | 0.11 ms | 1.00 | 1.00 | 1859.5 KB | 1.00 | 19.8 KB | 1.00 | Win |
| 100 | package-profile | package | Package size | realworld-report-chart-first | EPPlus | 16.00 ms | 0.46 ms | 0.27 ms | 3.53 | 3.53 | 13575.3 KB | 7.30 | 16.6 KB | 0.84 | 253.5% slower than OfficeIMO |
| 100 | package-profile | package | Package size | realworld-report-extra-column | OfficeIMO.Excel | 4.38 ms | 0.21 ms | 0.12 ms | 1.00 | 1.00 | 1859.0 KB | 1.00 | 20.2 KB | 1.00 | Win |
| 100 | package-profile | package | Package size | realworld-report-extra-column | EPPlus | 16.84 ms | 1.68 ms | 0.97 ms | 3.85 | 3.85 | 13865.9 KB | 7.46 | 17.0 KB | 0.84 | 285.0% slower than OfficeIMO |
| 100 | package-profile | package | Package size | realworld-report-no-autofit | OfficeIMO.Excel | 4.35 ms | 0.11 ms | 0.06 ms | 1.00 | 1.00 | 1836.4 KB | 1.00 | 19.6 KB | 1.00 | Win |
| 100 | package-profile | package | Package size | realworld-report-no-autofit | EPPlus | 14.46 ms | 0.99 ms | 0.57 ms | 3.32 | 3.32 | 12159.0 KB | 6.62 | 16.5 KB | 0.84 | 232.2% slower than OfficeIMO |
| 100 | package-profile | package | Package size | realworld-report-post-mutation | OfficeIMO.Excel | 4.46 ms | 0.28 ms | 0.16 ms | 1.00 | 1.00 | 1853.8 KB | 1.00 | 19.8 KB | 1.00 | Win |
| 100 | package-profile | package | Package size | realworld-report-post-mutation | EPPlus | 16.34 ms | 1.07 ms | 0.62 ms | 3.66 | 3.66 | 13576.1 KB | 7.32 | 16.6 KB | 0.84 | 266.1% slower than OfficeIMO |
| 100 | package-profile | package | Package size | realworld-report-shuffled-columns | OfficeIMO.Excel | 5.23 ms | 0.75 ms | 0.43 ms | 1.00 | 1.00 | 1843.3 KB | 1.00 | 19.6 KB | 1.00 | Win |
| 100 | package-profile | package | Package size | realworld-report-shuffled-columns | EPPlus | 17.07 ms | 0.63 ms | 0.36 ms | 3.26 | 3.26 | 13570.7 KB | 7.36 | 16.6 KB | 0.85 | 226.4% slower than OfficeIMO |
| 100 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | OfficeIMO.Excel | 3.71 ms | 0.26 ms | 0.15 ms | 1.00 | 1.00 | 1837.7 KB | 1.00 |  |  | Win |
| 100 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus 4.5.3.3 | 6.14 ms |  |  | 1.66 | 1.66 |  |  |  |  | 65.8% slower than OfficeIMO |
| 100 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus | 11.76 ms | 0.40 ms | 0.23 ms | 3.17 | 3.17 | 12159.0 KB | 6.62 |  |  | 217.4% slower than OfficeIMO |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | OfficeIMO.Excel | 5.08 ms | 2.35 ms | 1.36 ms | 1.00 | 1.00 | 1858.1 KB | 1.00 |  |  | Win |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus 4.5.3.3 | 9.26 ms |  |  | 1.82 | 1.82 |  |  |  |  | 82.3% slower than OfficeIMO |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus | 14.60 ms | 0.72 ms | 0.41 ms | 2.88 | 2.88 | 13575.3 KB | 7.31 |  |  | 187.6% slower than OfficeIMO |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | OfficeIMO.Excel | 3.75 ms | 0.12 ms | 0.07 ms | 1.00 | 1.00 | 1859.0 KB | 1.00 |  |  | Win |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus 4.5.3.3 | 8.62 ms |  |  | 2.30 | 2.30 |  |  |  |  | 129.7% slower than OfficeIMO |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus | 16.28 ms | 1.20 ms | 0.69 ms | 4.34 | 4.34 | 13865.9 KB | 7.46 |  |  | 334.0% slower than OfficeIMO |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | OfficeIMO.Excel | 3.76 ms | 0.13 ms | 0.07 ms | 1.00 | 1.00 | 1853.8 KB | 1.00 |  |  | Win |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus 4.5.3.3 | 8.61 ms |  |  | 2.29 | 2.29 |  |  |  |  | 129.0% slower than OfficeIMO |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus | 15.99 ms | 2.21 ms | 1.28 ms | 4.26 | 4.26 | 13576.1 KB | 7.32 |  |  | 325.6% slower than OfficeIMO |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | OfficeIMO.Excel | 5.18 ms | 2.05 ms | 1.19 ms | 1.00 | 1.00 | 1843.2 KB | 1.00 |  |  | Win |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus 4.5.3.3 | 8.80 ms |  |  | 1.70 | 1.70 |  |  |  |  | 70.0% slower than OfficeIMO |
| 100 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus | 14.87 ms | 1.00 ms | 0.58 ms | 2.87 | 2.87 | 13570.7 KB | 7.36 |  |  | 187.2% slower than OfficeIMO |
| 100 | speed-comparison | other | Real-world report | realworld-report-all-in-one | OfficeIMO.Excel | 4.16 ms | 0.52 ms | 0.30 ms | 1.00 | 1.00 | 1875.6 KB | 1.00 |  |  | Win |
| 100 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus 4.5.3.3 | 10.83 ms |  |  | 2.61 | 2.61 |  |  |  |  | 160.5% slower than OfficeIMO |
| 100 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus | 16.89 ms | 1.00 ms | 0.58 ms | 4.07 | 4.07 | 13576.2 KB | 7.24 |  |  | 306.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-all-in-one | OfficeIMO.Excel | 11.24 ms | 0.25 ms | 0.14 ms | 1.00 | 1.00 | 6196.8 KB | 1.00 | 206.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-all-in-one | EPPlus | 65.18 ms | 0.37 ms | 0.22 ms | 5.80 | 5.80 | 54595.4 KB | 8.81 | 121.8 KB | 0.59 | 479.9% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-chart-first | OfficeIMO.Excel | 11.13 ms | 0.14 ms | 0.08 ms | 1.00 | 1.00 | 6195.8 KB | 1.00 | 206.5 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-chart-first | EPPlus | 68.60 ms | 1.47 ms | 0.85 ms | 6.16 | 6.16 | 54594.2 KB | 8.81 | 121.8 KB | 0.59 | 516.5% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-extra-column | OfficeIMO.Excel | 12.27 ms | 0.88 ms | 0.51 ms | 1.00 | 1.00 | 6390.7 KB | 1.00 | 219.1 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-extra-column | EPPlus | 70.75 ms | 2.45 ms | 1.42 ms | 5.77 | 5.77 | 59226.4 KB | 9.27 | 128.4 KB | 0.59 | 476.6% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-no-autofit | OfficeIMO.Excel | 11.00 ms | 0.53 ms | 0.31 ms | 1.00 | 1.00 | 6190.0 KB | 1.00 | 206.4 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-no-autofit | EPPlus | 48.68 ms | 4.64 ms | 2.68 ms | 4.42 | 4.42 | 32906.4 KB | 5.32 | 121.8 KB | 0.59 | 342.4% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-post-mutation | OfficeIMO.Excel | 11.96 ms | 0.88 ms | 0.51 ms | 1.00 | 1.00 | 6206.5 KB | 1.00 | 206.6 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-post-mutation | EPPlus | 69.29 ms | 2.49 ms | 1.44 ms | 5.79 | 5.79 | 54594.5 KB | 8.80 | 121.9 KB | 0.59 | 479.1% slower than OfficeIMO |
| 2500 | package-profile | package | Package size | realworld-report-shuffled-columns | OfficeIMO.Excel | 12.70 ms | 0.19 ms | 0.11 ms | 1.00 | 1.00 | 6202.9 KB | 1.00 | 211.2 KB | 1.00 | Win |
| 2500 | package-profile | package | Package size | realworld-report-shuffled-columns | EPPlus | 70.56 ms | 7.26 ms | 4.19 ms | 5.56 | 5.56 | 54591.3 KB | 8.80 | 124.3 KB | 0.59 | 455.7% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | OfficeIMO.Excel | 23.06 ms | 0.36 ms | 0.21 ms | 1.00 | 1.00 | 6196.2 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus 4.5.3.3 | 48.73 ms |  |  | 2.11 | 2.11 |  |  |  |  | 111.3% slower than OfficeIMO |
| 2500 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus | 81.67 ms | 6.24 ms | 3.61 ms | 3.54 | 3.54 | 33347.5 KB | 5.38 |  |  | 254.1% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | OfficeIMO.Excel | 23.93 ms | 0.85 ms | 0.49 ms | 1.00 | 1.00 | 6198.1 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus 4.5.3.3 | 89.12 ms |  |  | 3.72 | 3.72 |  |  |  |  | 272.4% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus | 145.19 ms | 7.32 ms | 4.22 ms | 6.07 | 6.07 | 56285.8 KB | 9.08 |  |  | 506.7% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | OfficeIMO.Excel | 25.00 ms | 0.44 ms | 0.26 ms | 1.00 | 1.00 | 6389.4 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus 4.5.3.3 | 85.71 ms |  |  | 3.43 | 3.43 |  |  |  |  | 242.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus | 170.29 ms | 3.97 ms | 2.29 ms | 6.81 | 6.81 | 61191.6 KB | 9.58 |  |  | 581.2% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | OfficeIMO.Excel | 25.24 ms | 2.99 ms | 1.73 ms | 1.00 | 1.00 | 6209.7 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus 4.5.3.3 | 80.09 ms |  |  | 3.17 | 3.17 |  |  |  |  | 217.4% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus | 156.17 ms | 8.14 ms | 4.70 ms | 6.19 | 6.19 | 56231.7 KB | 9.06 |  |  | 518.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | OfficeIMO.Excel | 24.51 ms | 0.50 ms | 0.29 ms | 1.00 | 1.00 | 6201.9 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus 4.5.3.3 | 76.70 ms |  |  | 3.13 | 3.13 |  |  |  |  | 212.9% slower than OfficeIMO |
| 2500 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus | 160.33 ms | 5.69 ms | 3.28 ms | 6.54 | 6.54 | 56283.3 KB | 9.08 |  |  | 554.1% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world report | realworld-report-all-in-one | OfficeIMO.Excel | 20.81 ms | 1.76 ms | 1.01 ms | 1.00 | 1.00 | 6201.0 KB | 1.00 |  |  | Win |
| 2500 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus 4.5.3.3 | 112.86 ms |  |  | 5.42 | 5.42 |  |  |  |  | 442.4% slower than OfficeIMO |
| 2500 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus | 134.33 ms | 2.58 ms | 1.49 ms | 6.46 | 6.46 | 56287.7 KB | 9.08 |  |  | 545.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-all-in-one | OfficeIMO.Excel | 87.86 ms | 1.21 ms | 0.70 ms | 1.00 | 1.00 | 43678.5 KB | 1.00 | 1985.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-all-in-one | EPPlus | 445.19 ms | 17.17 ms | 9.92 ms | 5.07 | 5.07 | 277076.8 KB | 6.34 | 1097.7 KB | 0.55 | 406.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-chart-first | OfficeIMO.Excel | 85.98 ms | 0.31 ms | 0.18 ms | 1.00 | 1.00 | 43560.9 KB | 1.00 | 1985.9 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-chart-first | EPPlus | 419.31 ms | 3.96 ms | 2.28 ms | 4.88 | 4.88 | 277076.7 KB | 6.36 | 1097.7 KB | 0.55 | 387.7% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-extra-column | OfficeIMO.Excel | 97.36 ms | 0.65 ms | 0.37 ms | 1.00 | 1.00 | 45563.8 KB | 1.00 | 2110.8 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-extra-column | EPPlus | 476.92 ms | 19.49 ms | 11.25 ms | 4.90 | 4.90 | 302760.4 KB | 6.64 | 1166.2 KB | 0.55 | 389.8% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-no-autofit | OfficeIMO.Excel | 84.07 ms | 0.82 ms | 0.47 ms | 1.00 | 1.00 | 43670.7 KB | 1.00 | 1985.8 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-no-autofit | EPPlus | 391.65 ms | 1.33 ms | 0.77 ms | 4.66 | 4.66 | 234782.0 KB | 5.38 | 1097.7 KB | 0.55 | 365.9% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-post-mutation | OfficeIMO.Excel | 88.70 ms | 1.40 ms | 0.81 ms | 1.00 | 1.00 | 43688.6 KB | 1.00 | 1986.0 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-post-mutation | EPPlus | 415.67 ms | 10.60 ms | 6.12 ms | 4.69 | 4.69 | 277077.9 KB | 6.34 | 1097.8 KB | 0.55 | 368.6% slower than OfficeIMO |
| 25000 | package-profile | package | Package size | realworld-report-shuffled-columns | OfficeIMO.Excel | 98.94 ms | 1.86 ms | 1.07 ms | 1.00 | 1.00 | 43740.6 KB | 1.00 | 2046.1 KB | 1.00 | Win |
| 25000 | package-profile | package | Package size | realworld-report-shuffled-columns | EPPlus | 431.61 ms | 8.35 ms | 4.82 ms | 4.36 | 4.36 | 277070.5 KB | 6.33 | 1098.4 KB | 0.54 | 336.2% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | OfficeIMO.Excel | 85.59 ms | 1.85 ms | 1.07 ms | 1.00 | 1.00 | 43670.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus 4.5.3.3 | 250.74 ms |  |  | 2.93 | 2.93 |  |  |  |  | 192.9% slower than OfficeIMO |
| 25000 | speed-comparison | mutate | AutoFit and mutation | realworld-report-no-autofit | EPPlus | 385.57 ms | 12.09 ms | 6.98 ms | 4.50 | 4.50 | 234781.0 KB | 5.38 |  |  | 350.5% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | OfficeIMO.Excel | 90.64 ms | 2.70 ms | 1.56 ms | 1.00 | 1.00 | 43560.4 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus | 431.25 ms | 4.93 ms | 2.85 ms | 4.76 | 4.76 | 277076.7 KB | 6.36 |  |  | 375.8% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-chart-first | EPPlus 4.5.3.3 | 485.20 ms |  |  | 5.35 | 5.35 |  |  |  |  | 435.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | OfficeIMO.Excel | 97.45 ms | 2.32 ms | 1.34 ms | 1.00 | 1.00 | 45565.2 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus | 477.97 ms | 15.32 ms | 8.84 ms | 4.90 | 4.90 | 302760.4 KB | 6.64 |  |  | 390.5% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-extra-column | EPPlus 4.5.3.3 | 538.19 ms |  |  | 5.52 | 5.52 |  |  |  |  | 452.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | OfficeIMO.Excel | 100.78 ms | 7.23 ms | 4.17 ms | 1.00 | 1.00 | 43689.8 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus | 434.62 ms | 19.17 ms | 11.07 ms | 4.31 | 4.31 | 277077.9 KB | 6.34 |  |  | 331.3% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-post-mutation | EPPlus 4.5.3.3 | 514.00 ms |  |  | 5.10 | 5.10 |  |  |  |  | 410.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | OfficeIMO.Excel | 104.78 ms | 3.45 ms | 1.99 ms | 1.00 | 1.00 | 43736.7 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus | 434.86 ms | 13.92 ms | 8.04 ms | 4.15 | 4.15 | 277070.4 KB | 6.33 |  |  | 315.0% slower than OfficeIMO |
| 25000 | speed-comparison | other | Anti-cheat report variants | realworld-report-shuffled-columns | EPPlus 4.5.3.3 | 476.58 ms |  |  | 4.55 | 4.55 |  |  |  |  | 354.8% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | OfficeIMO.Excel | 87.32 ms | 0.93 ms | 0.54 ms | 1.00 | 1.00 | 43679.5 KB | 1.00 |  |  | Win |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus | 410.43 ms | 0.25 ms | 0.14 ms | 4.70 | 4.70 | 277077.7 KB | 6.34 |  |  | 370.1% slower than OfficeIMO |
| 25000 | speed-comparison | other | Real-world report | realworld-report-all-in-one | EPPlus 4.5.3.3 | 530.05 ms |  |  | 6.07 | 6.07 |  |  |  |  | 507.0% slower than OfficeIMO |
