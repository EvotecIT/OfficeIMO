# Native grouped tables: historical-exploratory

These are unmodified BenchmarkDotNet Markdown outputs under explicit capture/domain/class headers. Native FullName/job/parameter identities and all warnings/Actuals are in the exact raw packet and JSON/CSV ledgers. Raw API, mapper, metadata, inference and cold diagnostic-allocation boundaries remain as described in the dated report. A native Baseline ratio does not make different API work equivalent.

## stable36df-domain0-generated

Capture source: 36df28c4abc6a4c4371e09e26a5e40e0c07d30d2; started 2026-10-08T11:24:08.69005+00:00; finished 2026-10-08T11:27:56.5355906+00:00. Exact context: contexts/stable36df-domain0-generated-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.BorrowedRefStructReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-ULQWZY : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-36df28",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
InvocationCount=4  IterationCount=10  LaunchCount=1  
UnrollFactor=1  WarmupCount=6  

```
| Method                     | RowCount | Mean      | Error    | StdDev   | Ratio | RatioSD | Allocated | Alloc Ratio |
|--------------------------- |--------- |----------:|---------:|---------:|------:|--------:|----------:|------------:|
| ExcelReaderMappedRefStruct | 50000    |  9.257 ms | 1.827 ms | 1.208 ms |  1.02 |    0.18 |   5.05 KB |        1.00 |
| OfficeIMOFactoryRefStruct  | 50000    | 11.419 ms | 2.659 ms | 1.759 ms |  1.25 |    0.24 |  77.65 KB |       15.36 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.MaterializedRawAsyncReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-ULQWZY : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-36df28",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
InvocationCount=4  IterationCount=10  LaunchCount=1  
UnrollFactor=1  WarmupCount=6  

```
| Method                              | RowCount | Format | Mean      | Error     | StdDev    | Median    | Ratio | RatioSD | Allocated | Alloc Ratio |
|------------------------------------ |--------- |------- |----------:|----------:|----------:|----------:|------:|--------:|----------:|------------:|
| **ExcelReaderStringsMaterializedAsync** | **50000**    | **Xlsx**   | **13.120 ms** |  **5.475 ms** | **3.6213 ms** | **11.551 ms** |  **1.07** |    **0.38** |   **1.58 MB** |        **1.00** |
| OfficeIMOAsync                      | 50000    | Xlsx   | 16.675 ms | 12.636 ms | 8.3578 ms | 13.606 ms |  1.35 |    0.74 |   3.51 MB |        2.22 |
| SylvanAsync                         | 50000    | Xlsx   | 35.994 ms |  6.126 ms | 4.0517 ms | 35.692 ms |  2.92 |    0.77 |   1.89 MB |        1.20 |
|                                     |          |        |           |           |           |           |       |         |           |             |
| **ExcelReaderStringsMaterializedAsync** | **50000**    | **Xlsb**   |  **5.883 ms** |  **1.195 ms** | **0.7907 ms** |  **5.549 ms** |  **1.02** |    **0.18** |   **1.58 MB** |        **1.00** |
| OfficeIMOAsync                      | 50000    | Xlsb   | 14.419 ms |  2.758 ms | 1.8240 ms | 13.808 ms |  2.49 |    0.43 |    6.7 MB |        4.25 |
| SylvanAsync                         | 50000    | Xlsb   | 11.819 ms | 11.507 ms | 7.6114 ms |  8.590 ms |  2.04 |    1.29 |   1.82 MB |        1.16 |
|                                     |          |        |           |           |           |           |       |         |           |             |
| **ExcelReaderStringsMaterializedAsync** | **50000**    | **Xls**    |  **4.736 ms** |  **1.114 ms** | **0.7367 ms** |  **4.881 ms** |  **1.02** |    **0.23** |   **1.58 MB** |        **1.00** |
| OfficeIMOAsync                      | 50000    | Xls    |  6.320 ms |  1.480 ms | 0.9788 ms |  5.936 ms |  1.37 |    0.30 |  10.25 MB |        6.50 |
| SylvanAsync                         | 50000    | Xls    | 10.343 ms |  9.737 ms | 6.4404 ms |  7.978 ms |  2.24 |    1.39 |   1.68 MB |        1.06 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.MaterializedRawReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-ULQWZY : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-36df28",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
InvocationCount=4  IterationCount=10  LaunchCount=1  
UnrollFactor=1  WarmupCount=6  

```
| Method                         | RowCount | Format | Mean      | Error     | StdDev     | Median    | Ratio | RatioSD | Allocated | Alloc Ratio |
|------------------------------- |--------- |------- |----------:|----------:|-----------:|----------:|------:|--------:|----------:|------------:|
| **ExcelReaderStringsMaterialized** | **50000**    | **Xlsx**   |  **8.788 ms** |  **4.579 ms** |  **3.0285 ms** |  **6.821 ms** |  **1.09** |    **0.47** |   **1.58 MB** |        **1.00** |
| OfficeIMOOriginal              | 50000    | Xlsx   | 22.633 ms | 22.550 ms | 14.9152 ms | 14.673 ms |  2.82 |    1.97 |   3.51 MB |        2.22 |
| Sylvan                         | 50000    | Xlsx   | 39.742 ms | 14.194 ms |  9.3884 ms | 39.608 ms |  4.94 |    1.74 |   1.89 MB |        1.20 |
|                                |          |        |           |           |            |           |       |         |           |             |
| **ExcelReaderStringsMaterialized** | **50000**    | **Xlsb**   |  **5.992 ms** |  **1.327 ms** |  **0.8775 ms** |  **5.920 ms** |  **1.02** |    **0.20** |   **1.58 MB** |        **1.00** |
| OfficeIMOOriginal              | 50000    | Xlsb   | 10.358 ms |  1.814 ms |  1.2001 ms |  9.823 ms |  1.76 |    0.31 |    6.7 MB |        4.25 |
| Sylvan                         | 50000    | Xlsb   | 14.574 ms |  3.366 ms |  2.2262 ms | 13.447 ms |  2.48 |    0.50 |   1.82 MB |        1.15 |
|                                |          |        |           |           |            |           |       |         |           |             |
| **ExcelReaderStringsMaterialized** | **50000**    | **Xls**    |  **5.028 ms** |  **1.074 ms** |  **0.7102 ms** |  **5.101 ms** |  **1.02** |    **0.23** |   **1.58 MB** |        **1.00** |
| OfficeIMOOriginal              | 50000    | Xls    |  8.425 ms |  1.878 ms |  1.2425 ms |  8.784 ms |  1.71 |    0.38 |  10.25 MB |        6.50 |
| Sylvan                         | 50000    | Xls    |  7.582 ms |  2.372 ms |  1.5688 ms |  7.822 ms |  1.54 |    0.41 |   1.68 MB |        1.06 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.RawAsyncReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-ULQWZY : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-36df28",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
InvocationCount=4  IterationCount=10  LaunchCount=1  
UnrollFactor=1  WarmupCount=6  

```
| Method           | RowCount | Format | Mean      | Error      | StdDev     | Median    | Ratio | RatioSD | Allocated   | Alloc Ratio |
|----------------- |--------- |------- |----------:|-----------:|-----------:|----------:|------:|--------:|------------:|------------:|
| **ExcelReaderAsync** | **50000**    | **Xlsx**   | **13.318 ms** |  **3.8138 ms** |  **2.5226 ms** | **13.069 ms** |  **1.03** |    **0.26** |     **3.89 KB** |        **1.00** |
| OfficeIMOAsync   | 50000    | Xlsx   | 19.276 ms | 23.2478 ms | 15.3770 ms | 13.581 ms |  1.49 |    1.18 |  3592.43 KB |      923.36 |
| SylvanAsync      | 50000    | Xlsx   | 38.833 ms |  6.8633 ms |  4.5397 ms | 39.819 ms |  3.01 |    0.62 |  1939.91 KB |      498.61 |
|                  |          |        |           |            |            |           |       |         |             |             |
| **ExcelReaderAsync** | **50000**    | **Xlsb**   |  **5.404 ms** |  **0.9636 ms** |  **0.6374 ms** |  **5.313 ms** |  **1.01** |    **0.16** |      **4.2 KB** |        **1.00** |
| OfficeIMOAsync   | 50000    | Xlsb   | 12.031 ms |  1.9288 ms |  1.2758 ms | 12.109 ms |  2.25 |    0.34 |  6858.49 KB |    1,631.76 |
| SylvanAsync      | 50000    | Xlsb   | 22.536 ms |  3.1932 ms |  2.1121 ms | 22.413 ms |  4.22 |    0.61 |  1866.56 KB |      444.09 |
|                  |          |        |           |            |            |           |       |         |             |             |
| **ExcelReaderAsync** | **50000**    | **Xls**    |  **2.921 ms** |  **0.6067 ms** |  **0.4013 ms** |  **2.742 ms** |  **1.01** |    **0.17** |     **2.91 KB** |        **1.00** |
| OfficeIMOAsync   | 50000    | Xls    |  6.942 ms |  2.0808 ms |  1.3763 ms |  6.850 ms |  2.41 |    0.53 | 10499.45 KB |    3,612.71 |
| SylvanAsync      | 50000    | Xls    |  8.308 ms |  2.2537 ms |  1.4907 ms |  8.777 ms |  2.88 |    0.59 |  1719.08 KB |      591.51 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.RawReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-ULQWZY : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-36df28",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
InvocationCount=4  IterationCount=10  LaunchCount=1  
UnrollFactor=1  WarmupCount=6  

```
| Method              | RowCount | Format | Mean      | Error      | StdDev     | Median    | Ratio | RatioSD | Gen0     | Allocated   | Alloc Ratio |
|-------------------- |--------- |------- |----------:|-----------:|-----------:|----------:|------:|--------:|---------:|------------:|------------:|
| **ExcelReaderOriginal** | **50000**    | **Xlsx**   | **10.081 ms** |  **5.8359 ms** |  **3.8601 ms** | **10.698 ms** |  **1.16** |    **0.64** |        **-** |     **3.82 KB** |        **1.00** |
| OfficeIMOOriginal   | 50000    | Xlsx   | 18.401 ms | 19.6052 ms | 12.9677 ms | 13.840 ms |  2.12 |    1.74 |        - |  3592.42 KB |      940.35 |
| Sylvan              | 50000    | Xlsx   | 33.715 ms | 11.4682 ms |  7.5855 ms | 31.402 ms |  3.88 |    1.77 |        - |     1939 KB |      507.55 |
|                     |          |        |           |            |            |           |       |         |          |             |             |
| **ExcelReaderOriginal** | **50000**    | **Xlsb**   |  **4.747 ms** |  **0.7071 ms** |  **0.4677 ms** |  **4.583 ms** |  **1.01** |    **0.13** |        **-** |     **4.13 KB** |        **1.00** |
| OfficeIMOOriginal   | 50000    | Xlsb   | 17.546 ms |  3.4874 ms |  2.3067 ms | 17.510 ms |  3.72 |    0.56 |        - |  6858.42 KB |    1,659.50 |
| Sylvan              | 50000    | Xlsb   | 16.404 ms |  6.8017 ms |  4.4989 ms | 18.875 ms |  3.48 |    0.96 |        - |  1865.65 KB |      451.42 |
|                     |          |        |           |            |            |           |       |         |          |             |             |
| **ExcelReaderOriginal** | **50000**    | **Xls**    |  **4.742 ms** |  **0.3313 ms** |  **0.2191 ms** |  **4.698 ms** |  **1.00** |    **0.06** |        **-** |    **30.26 KB** |        **1.00** |
| OfficeIMOOriginal   | 50000    | Xls    |  9.160 ms |  0.6787 ms |  0.4489 ms |  9.264 ms |  1.94 |    0.12 | 250.0000 | 11587.52 KB |      382.98 |
| Sylvan              | 50000    | Xls    |  9.619 ms |  1.5098 ms |  0.9986 ms |  9.189 ms |  2.03 |    0.22 |        - |  1725.85 KB |       57.04 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedAsyncReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-ULQWZY : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-36df28",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
InvocationCount=4  IterationCount=10  LaunchCount=1  
UnrollFactor=1  WarmupCount=6  

```
| Method                | RowCount | Shape    | Mean     | Error    | StdDev   | Ratio | RatioSD | Allocated | Alloc Ratio |
|---------------------- |--------- |--------- |---------:|---------:|---------:|------:|--------:|----------:|------------:|
| OfficeIMOTypedAsync   | 50000    | Original | 21.79 ms | 4.107 ms | 2.716 ms |  1.45 |    0.23 |   6.94 MB |        1.80 |
| ExcelReaderTypedAsync | 50000    | Original | 15.19 ms | 2.446 ms | 1.618 ms |  1.01 |    0.14 |   3.87 MB |        1.00 |
| SylvanTypedAsync      | 50000    | Original | 68.53 ms | 7.577 ms | 5.011 ms |  4.56 |    0.56 |  10.48 MB |        2.71 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedModelReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-ULQWZY : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-36df28",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
InvocationCount=4  IterationCount=10  LaunchCount=1  
UnrollFactor=1  WarmupCount=6  

```
| Method               | RowCount | Model  | Mean     | Error     | StdDev   | Ratio | RatioSD | Allocated | Alloc Ratio |
|--------------------- |--------- |------- |---------:|----------:|---------:|------:|--------:|----------:|------------:|
| **ExcelReaderAutomatic** | **50000**    | **Class**  | **12.53 ms** |  **3.266 ms** | **2.160 ms** |  **1.03** |    **0.24** |   **3.87 MB** |        **1.00** |
| OfficeIMOAutomatic   | 50000    | Class  | 25.51 ms | 11.050 ms | 7.309 ms |  2.09 |    0.68 |   6.94 MB |        1.80 |
|                      |          |        |          |           |          |       |         |           |             |
| **ExcelReaderAutomatic** | **50000**    | **Struct** | **12.78 ms** |  **2.229 ms** | **1.474 ms** |  **1.01** |    **0.16** |   **1.58 MB** |        **1.00** |
| OfficeIMOAutomatic   | 50000    | Struct | 23.04 ms |  1.741 ms | 1.152 ms |  1.83 |    0.23 |   4.65 MB |        2.95 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedStreamAsyncReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-ULQWZY : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-36df28",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
InvocationCount=4  IterationCount=10  LaunchCount=1  
UnrollFactor=1  WarmupCount=6  

```
| Method                          | RowCount | Shape    | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0      | Gen1      | Gen2      | Allocated | Alloc Ratio |
|-------------------------------- |--------- |--------- |----------:|----------:|----------:|------:|--------:|----------:|----------:|----------:|----------:|------------:|
| OfficeIMOStreamOpenTypedAsync   | 50000    | Original | 622.55 ms | 46.471 ms | 30.738 ms | 55.48 |    7.27 | 9750.0000 | 6750.0000 | 1250.0000 | 411.01 MB |      106.29 |
| ExcelReaderStreamOpenTypedAsync | 50000    | Original |  11.43 ms |  2.717 ms |  1.797 ms |  1.02 |    0.20 |         - |         - |         - |   3.87 MB |        1.00 |
| SylvanStreamOpenTypedAsync      | 50000    | Original |  60.55 ms |  9.435 ms |  6.241 ms |  5.40 |    0.85 |         - |         - |         - |  10.48 MB |        2.71 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedXlsbAsyncReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-ULQWZY : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-36df28",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
InvocationCount=4  IterationCount=10  LaunchCount=1  
UnrollFactor=1  WarmupCount=6  

```
| Method                    | RowCount | Mean      | Error    | StdDev   | Ratio | RatioSD | Allocated | Alloc Ratio |
|-------------------------- |--------- |----------:|---------:|---------:|------:|--------:|----------:|------------:|
| ExcelReaderTypedXlsbAsync | 50000    |  8.358 ms | 1.607 ms | 1.063 ms |  1.01 |    0.17 |   3.87 MB |        1.00 |
| OfficeIMOTypedXlsbAsync   | 50000    | 22.494 ms | 4.708 ms | 3.114 ms |  2.73 |    0.48 |  10.13 MB |        2.62 |
| SylvanTypedXlsbAsync      | 50000    | 25.629 ms | 3.910 ms | 2.586 ms |  3.11 |    0.47 |   10.4 MB |        2.69 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedXlsbReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-ULQWZY : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-36df28",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
InvocationCount=4  IterationCount=10  LaunchCount=1  
UnrollFactor=1  WarmupCount=6  

```
| Method               | RowCount | Mean      | Error     | StdDev    | Ratio | RatioSD | Allocated | Alloc Ratio |
|--------------------- |--------- |----------:|----------:|----------:|------:|--------:|----------:|------------:|
| ExcelReaderTypedXlsb | 50000    |  8.109 ms | 1.6995 ms | 1.1241 ms |  1.02 |    0.19 |   3.87 MB |        1.00 |
| OfficeIMOTypedXlsb   | 50000    | 21.268 ms | 0.8497 ms | 0.5620 ms |  2.67 |    0.34 |   5.55 MB |        1.44 |
| SylvanTypedXlsb      | 50000    | 26.022 ms | 4.5290 ms | 2.9956 ms |  3.26 |    0.54 |   10.4 MB |        2.69 |


## stable36df-domain0-real

Capture source: 36df28c4abc6a4c4371e09e26a5e40e0c07d30d2; started 2026-10-08T11:27:57.4619446+00:00; finished 2026-10-08T11:33:41.2824802+00:00. Exact context: contexts/stable36df-domain0-real-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.MaterializedRealDataReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-ULQWZY : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-36df28",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
InvocationCount=4  IterationCount=10  LaunchCount=1  
UnrollFactor=1  WarmupCount=6  

```
| Method                         | Format | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated   | Alloc Ratio |
|------------------------------- |------- |----------:|----------:|----------:|------:|--------:|---------:|---------:|---------:|------------:|------------:|
| **ExcelReaderStringsMaterialized** | **Xlsx**   |  **44.20 ms** |  **6.083 ms** |  **4.024 ms** |  **1.01** |    **0.12** |        **-** |        **-** |        **-** |    **17.09 KB** |        **1.00** |
| OfficeIMOBytes                 | Xlsx   |  69.63 ms | 13.581 ms |  8.983 ms |  1.59 |    0.23 | 250.0000 |        - |        - | 13939.88 KB |      815.50 |
| OfficeIMOStream                | Xlsx   |  76.84 ms | 16.720 ms | 11.059 ms |  1.75 |    0.28 | 250.0000 |        - |        - | 23048.04 KB |    1,348.33 |
| Sylvan                         | Xlsx   | 205.29 ms | 59.036 ms | 39.049 ms |  4.68 |    0.92 |        - |        - |        - |   647.55 KB |       37.88 |
|                                |        |           |           |           |       |         |          |          |          |             |             |
| **ExcelReaderStringsMaterialized** | **Xlsm**   |  **44.31 ms** |  **7.338 ms** |  **4.854 ms** |  **1.01** |    **0.14** |        **-** |        **-** |        **-** |   **141.48 KB** |        **1.00** |
| OfficeIMOBytes                 | Xlsm   |  65.59 ms | 11.612 ms |  7.681 ms |  1.50 |    0.22 | 250.0000 |        - |        - | 16484.18 KB |      116.51 |
| OfficeIMOStream                | Xlsm   |  79.36 ms | 15.431 ms | 10.206 ms |  1.81 |    0.28 | 250.0000 |        - |        - | 23048.07 KB |      162.90 |
| Sylvan                         | Xlsm   | 164.20 ms | 45.743 ms | 30.256 ms |  3.74 |    0.75 |        - |        - |        - |   647.63 KB |        4.58 |
|                                |        |           |           |           |       |         |          |          |          |             |             |
| **ExcelReaderStringsMaterialized** | **Xlsb**   |  **27.26 ms** |  **1.683 ms** |  **1.113 ms** |  **1.00** |    **0.06** |        **-** |        **-** |        **-** |   **133.23 KB** |        **1.00** |
| OfficeIMOBytes                 | Xlsb   |  31.96 ms |  2.626 ms |  1.737 ms |  1.17 |    0.08 | 500.0000 | 250.0000 | 250.0000 | 18089.64 KB |      135.78 |
| OfficeIMOStream                | Xlsb   |  31.79 ms |  3.526 ms |  2.332 ms |  1.17 |    0.09 | 500.0000 | 250.0000 | 250.0000 | 25970.95 KB |      194.93 |
| Sylvan                         | Xlsb   |  27.61 ms |  2.128 ms |  1.408 ms |  1.01 |    0.06 |        - |        - |        - |   341.98 KB |        2.57 |
|                                |        |           |           |           |       |         |          |          |          |             |             |
| **ExcelReaderStringsMaterialized** | **Xls**    |  **11.42 ms** |  **3.155 ms** |  **2.087 ms** |  **1.02** |    **0.23** |        **-** |        **-** |        **-** |    **82.96 KB** |        **1.00** |
| OfficeIMOBytes                 | Xls    |  14.85 ms |  3.299 ms |  2.182 ms |  1.33 |    0.27 | 250.0000 |        - |        - | 18150.29 KB |      218.77 |
| OfficeIMOStream                | Xls    |  17.25 ms |  3.799 ms |  2.513 ms |  1.55 |    0.31 | 500.0000 | 250.0000 | 250.0000 | 33859.17 KB |      408.11 |
| Sylvan                         | Xls    |  18.57 ms |  5.101 ms |  3.374 ms |  1.67 |    0.37 |        - |        - |        - |    218.4 KB |        2.63 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.PrefetchedRealDataReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-ULQWZY : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-36df28",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
InvocationCount=4  IterationCount=10  LaunchCount=1  
UnrollFactor=1  WarmupCount=6  

```
| Method                    | Format | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1     | Allocated   | Alloc Ratio |
|-------------------------- |------- |----------:|----------:|----------:|------:|--------:|---------:|---------:|------------:|------------:|
| **ExcelReaderPrefetch**       | **Xlsx**   |  **27.64 ms** |  **3.109 ms** |  **2.056 ms** |  **1.00** |    **0.10** |        **-** |        **-** |    **163.8 KB** |        **1.00** |
| ExcelReaderMemoryPrefetch | Xlsx   |  25.99 ms |  4.228 ms |  2.797 ms |  0.94 |    0.12 |        - |        - |   127.54 KB |        0.78 |
| OfficeIMOBytes            | Xlsx   |  64.15 ms |  3.227 ms |  2.135 ms |  2.33 |    0.18 | 250.0000 |        - | 16484.24 KB |      100.63 |
| Sylvan                    | Xlsx   | 228.11 ms | 33.741 ms | 22.318 ms |  8.29 |    0.97 |        - |        - |   647.55 KB |        3.95 |
|                           |        |           |           |           |       |         |          |          |             |             |
| **ExcelReaderPrefetch**       | **Xlsm**   |  **33.78 ms** |  **4.744 ms** |  **3.138 ms** |  **1.01** |    **0.12** |        **-** |        **-** |   **151.53 KB** |        **1.00** |
| ExcelReaderMemoryPrefetch | Xlsm   |  38.66 ms |  6.179 ms |  4.087 ms |  1.15 |    0.15 |        - |        - |   135.59 KB |        0.89 |
| OfficeIMOBytes            | Xlsm   |  66.14 ms | 10.652 ms |  7.046 ms |  1.97 |    0.26 | 250.0000 |        - | 16484.29 KB |      108.79 |
| Sylvan                    | Xlsm   | 158.62 ms | 19.824 ms | 13.113 ms |  4.73 |    0.54 |        - |        - |   647.63 KB |        4.27 |
|                           |        |           |           |           |       |         |          |          |             |             |
| **ExcelReaderPrefetch**       | **Xlsb**   |  **16.90 ms** |  **1.730 ms** |  **1.144 ms** |  **1.00** |    **0.09** |        **-** |        **-** |   **149.61 KB** |        **1.00** |
| ExcelReaderMemoryPrefetch | Xlsb   |  17.20 ms |  2.442 ms |  1.615 ms |  1.02 |    0.12 |        - |        - |    114.5 KB |        0.77 |
| OfficeIMOBytes            | Xlsb   |  33.02 ms |  5.120 ms |  3.387 ms |  1.96 |    0.23 | 500.0000 | 500.0000 | 18087.12 KB |      120.90 |
| Sylvan                    | Xlsb   |  30.37 ms |  5.457 ms |  3.609 ms |  1.81 |    0.24 |        - |        - |   341.98 KB |        2.29 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.RealDataReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-ULQWZY : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-36df28",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
InvocationCount=4  IterationCount=10  LaunchCount=1  
UnrollFactor=1  WarmupCount=6  

```
| Method              | Format | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated   | Alloc Ratio |
|-------------------- |------- |----------:|----------:|----------:|------:|--------:|---------:|---------:|---------:|------------:|------------:|
| **ExcelReaderOriginal** | **Xlsx**   |  **39.82 ms** |  **5.328 ms** |  **3.524 ms** |  **1.01** |    **0.11** |        **-** |        **-** |        **-** |     **7.63 KB** |        **1.00** |
| ExcelReaderMemory   | Xlsx   |  40.82 ms |  7.375 ms |  4.878 ms |  1.03 |    0.14 |        - |        - |        - |     7.38 KB |        0.97 |
| OfficeIMOBytes      | Xlsx   |  66.04 ms | 18.373 ms | 12.153 ms |  1.67 |    0.32 | 250.0000 |        - |        - | 13939.88 KB |    1,828.18 |
| OfficeIMOStream     | Xlsx   |  63.47 ms | 11.034 ms |  7.298 ms |  1.60 |    0.21 | 250.0000 |        - |        - | 20503.74 KB |    2,689.02 |
| Sylvan              | Xlsx   | 156.56 ms | 13.060 ms |  8.639 ms |  3.96 |    0.36 |        - |        - |        - |   644.52 KB |       84.53 |
|                     |        |           |           |           |       |         |          |          |          |             |             |
| **ExcelReaderOriginal** | **Xlsm**   |  **47.31 ms** | **10.874 ms** |  **7.192 ms** |  **1.02** |    **0.19** |        **-** |        **-** |        **-** |     **7.63 KB** |        **1.00** |
| ExcelReaderMemory   | Xlsm   |  53.97 ms | 10.929 ms |  7.229 ms |  1.16 |    0.21 |        - |        - |        - |    83.75 KB |       10.98 |
| OfficeIMOBytes      | Xlsm   |  83.52 ms | 12.024 ms |  7.953 ms |  1.80 |    0.28 | 250.0000 |        - |        - | 16483.94 KB |    2,161.83 |
| OfficeIMOStream     | Xlsm   |  68.97 ms | 12.204 ms |  8.072 ms |  1.48 |    0.25 | 250.0000 |        - |        - | 23048.07 KB |    3,022.70 |
| Sylvan              | Xlsm   | 194.68 ms | 58.872 ms | 38.940 ms |  4.19 |    0.96 |        - |        - |        - |   647.63 KB |       84.93 |
|                     |        |           |           |           |       |         |          |          |          |             |             |
| **ExcelReaderOriginal** | **Xlsb**   |  **26.58 ms** |  **3.368 ms** |  **2.228 ms** |  **1.01** |    **0.11** |        **-** |        **-** |        **-** |   **123.76 KB** |        **1.00** |
| ExcelReaderMemory   | Xlsb   |  26.49 ms |  3.619 ms |  2.394 ms |  1.00 |    0.12 |        - |        - |        - |    76.88 KB |        0.62 |
| OfficeIMOBytes      | Xlsb   |  34.41 ms |  7.219 ms |  4.775 ms |  1.30 |    0.20 | 500.0000 | 250.0000 | 250.0000 | 18090.28 KB |      146.17 |
| OfficeIMOStream     | Xlsb   |  35.37 ms |  8.070 ms |  5.338 ms |  1.34 |    0.22 | 500.0000 | 250.0000 | 250.0000 | 25971.34 KB |      209.85 |
| Sylvan              | Xlsb   |  30.63 ms |  7.175 ms |  4.746 ms |  1.16 |    0.19 |        - |        - |        - |   341.98 KB |        2.76 |
|                     |        |           |           |           |       |         |          |          |          |             |             |
| **ExcelReaderOriginal** | **Xls**    |  **11.43 ms** |  **2.606 ms** |  **1.724 ms** |  **1.02** |    **0.21** |        **-** |        **-** |        **-** |     **73.5 KB** |        **1.00** |
| ExcelReaderMemory   | Xls    |  11.23 ms |  2.705 ms |  1.789 ms |  1.00 |    0.21 |        - |        - |        - |     57.3 KB |        0.78 |
| OfficeIMOBytes      | Xls    |  14.18 ms |  3.411 ms |  2.256 ms |  1.27 |    0.26 | 250.0000 |        - |        - | 18150.23 KB |      246.96 |
| OfficeIMOStream     | Xls    |  14.96 ms |  2.639 ms |  1.746 ms |  1.34 |    0.24 | 250.0000 |        - |        - | 25537.18 KB |      347.46 |
| Sylvan              | Xls    |  24.28 ms |  6.091 ms |  4.029 ms |  2.17 |    0.46 |        - |        - |        - |   186.34 KB |        2.54 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.RealDataTypedReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-ULQWZY : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-36df28",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
InvocationCount=4  IterationCount=10  LaunchCount=1  
UnrollFactor=1  WarmupCount=6  

```
| Method               | Format | Mean     | Error     | StdDev    | Ratio | RatioSD | Gen0     | Allocated | Alloc Ratio |
|--------------------- |------- |---------:|----------:|----------:|------:|--------:|---------:|----------:|------------:|
| **ExcelReaderAutomatic** | **Xlsx**   | **53.55 ms** | **15.902 ms** | **10.518 ms** |  **1.03** |    **0.25** |        **-** |   **8.02 MB** |        **1.00** |
| OfficeIMOAutomatic   | Xlsx   | 73.12 ms |  9.797 ms |  6.480 ms |  1.41 |    0.25 | 500.0000 |  24.62 MB |        3.07 |
|                      |        |          |           |           |       |         |          |           |             |
| **ExcelReaderAutomatic** | **Xlsb**   | **34.38 ms** |  **5.613 ms** |  **3.713 ms** |  **1.01** |    **0.15** |        **-** |   **8.02 MB** |        **1.00** |
| OfficeIMOAutomatic   | Xlsb   | 38.68 ms |  8.118 ms |  5.370 ms |  1.14 |    0.19 | 500.0000 |  24.66 MB |        3.08 |

