# Native grouped tables: historical-support

These are unmodified BenchmarkDotNet Markdown outputs under explicit capture/domain/class headers. Native FullName/job/parameter identities and all warnings/Actuals are in the exact raw packet and JSON/CSV ledgers. Raw API, mapper, metadata, inference and cold diagnostic-allocation boundaries remain as described in the dated report. A native Baseline ratio does not make different API work equivalent.

## stable36df-domain0-main-current

Capture source: 36df28c4abc6a4c4371e09e26a5e40e0c07d30d2; started 2026-10-08T11:16:45.8084858+00:00; finished 2026-10-08T11:19:39.3664789+00:00. Exact context: contexts/stable36df-domain0-main-current-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-LVIGEW : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-36df28",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
InvocationCount=32  IterationCount=15  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method           | RowCount | Shape         | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1    | Allocated | Alloc Ratio |
|----------------- |--------- |-------------- |----------:|----------:|----------:|------:|--------:|---------:|--------:|----------:|------------:|
| **ExcelReaderTyped** | **50000**    | **Original**      |  **8.150 ms** | **0.4721 ms** | **0.4416 ms** |  **1.00** |    **0.07** |  **62.5000** |       **-** |   **3.87 MB** |        **1.00** |
| OfficeIMOTyped   | 50000    | Original      | 10.429 ms | 1.9733 ms | 1.8459 ms |  1.28 |    0.23 |  31.2500 |       - |   2.36 MB |        0.61 |
| SylvanTyped      | 50000    | Original      | 51.899 ms | 7.9879 ms | 7.4719 ms |  6.38 |    0.94 | 187.5000 | 31.2500 |  10.48 MB |        2.71 |
|                  |          |               |           |           |           |       |         |          |         |           |             |
| **ExcelReaderTyped** | **50000**    | **SharedStrings** | **10.451 ms** | **1.6427 ms** | **1.5366 ms** |  **1.02** |    **0.22** |  **31.2500** |       **-** |   **2.29 MB** |        **1.00** |
| OfficeIMOTyped   | 50000    | SharedStrings |  9.907 ms | 1.4808 ms | 1.3851 ms |  0.97 |    0.21 |  31.2500 |       - |   2.36 MB |        1.03 |
| SylvanTyped      | 50000    | SharedStrings | 46.708 ms | 8.2217 ms | 7.6906 ms |  4.57 |    1.05 | 156.2500 | 31.2500 |   8.91 MB |        3.88 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.WriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-LVIGEW : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-36df28",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
InvocationCount=32  IterationCount=15  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method            | RowCount | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated | Alloc Ratio |
|------------------ |--------- |----------:|----------:|----------:|------:|--------:|---------:|---------:|---------:|----------:|------------:|
| OfficeIMO         | 50000    | 15.511 ms | 0.9985 ms | 0.9340 ms |  1.68 |    0.24 | 312.5000 | 312.5000 | 312.5000 |   4.04 MB |        1.00 |
| ExcelReaderWriter | 50000    |  9.373 ms | 1.3879 ms | 1.2983 ms |  1.02 |    0.19 | 312.5000 | 312.5000 | 312.5000 |   4.02 MB |        1.00 |


## stable36df-domain1-main-current

Capture source: 36df28c4abc6a4c4371e09e26a5e40e0c07d30d2; started 2026-10-08T11:19:39.806171+00:00; finished 2026-10-08T11:22:16.2160645+00:00. Exact context: contexts/stable36df-domain1-main-current-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-WWMGIL : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-36df28",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
InvocationCount=32  IterationCount=15  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method           | RowCount | Shape         | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1    | Allocated | Alloc Ratio |
|----------------- |--------- |-------------- |----------:|----------:|----------:|------:|--------:|---------:|--------:|----------:|------------:|
| **ExcelReaderTyped** | **50000**    | **Original**      |  **9.555 ms** | **1.8363 ms** | **1.7177 ms** |  **1.03** |    **0.24** |  **93.7500** |       **-** |   **3.88 MB** |        **1.00** |
| OfficeIMOTyped   | 50000    | Original      |  8.393 ms | 0.9252 ms | 0.8654 ms |  0.90 |    0.17 |  31.2500 |       - |   2.36 MB |        0.61 |
| SylvanTyped      | 50000    | Original      | 40.082 ms | 0.3166 ms | 0.2962 ms |  4.31 |    0.67 | 187.5000 | 31.2500 |  10.48 MB |        2.70 |
|                  |          |               |           |           |           |       |         |          |         |           |             |
| **ExcelReaderTyped** | **50000**    | **SharedStrings** |  **8.131 ms** | **1.0322 ms** | **0.9656 ms** |  **1.01** |    **0.15** |  **31.2500** |       **-** |   **2.29 MB** |        **1.00** |
| OfficeIMOTyped   | 50000    | SharedStrings |  8.330 ms | 1.0391 ms | 0.9719 ms |  1.04 |    0.15 |  31.2500 |       - |   2.36 MB |        1.03 |
| SylvanTyped      | 50000    | SharedStrings | 41.065 ms | 1.4850 ms | 1.3891 ms |  5.10 |    0.50 | 156.2500 | 31.2500 |   8.91 MB |        3.88 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.WriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-WWMGIL : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-36df28",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
InvocationCount=32  IterationCount=15  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method            | RowCount | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated | Alloc Ratio |
|------------------ |--------- |----------:|----------:|----------:|------:|--------:|---------:|---------:|---------:|----------:|------------:|
| OfficeIMO         | 50000    | 14.485 ms | 0.3789 ms | 0.3544 ms |  1.76 |    0.08 | 312.5000 | 312.5000 | 312.5000 |   4.04 MB |        1.00 |
| ExcelReaderWriter | 50000    |  8.229 ms | 0.3523 ms | 0.3296 ms |  1.00 |    0.05 | 312.5000 | 312.5000 | 312.5000 |   4.02 MB |        1.00 |


## stable36df-domain0-crypto-before

Capture source: 36df28c4abc6a4c4371e09e26a5e40e0c07d30d2; started 2026-10-08T11:22:16.5314284+00:00; finished 2026-10-08T11:22:52.8499887+00:00. Exact context: contexts/stable36df-domain0-crypto-before-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.EncryptedVerifiedReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-WZUHJM : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-ba9",/p:OfficeIMOBenchmarkNewApis=true  
InvocationCount=4  IterationCount=10  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                      | Input             | MemoryInput | Mean     | Error    | StdDev  | Gen0      | Allocated |
|---------------------------- |------------------ |------------ |---------:|---------:|--------:|----------:|----------:|
| **OfficeIMOVerifiedFullFields** | **OriginalSmallXlsx** | **False**       | **135.9 ms** |  **7.11 ms** | **4.70 ms** | **2750.0000** | **137.88 MB** |
| **OfficeIMOVerifiedFullFields** | **OriginalSmallXlsx** | **True**        | **151.0 ms** | **10.10 ms** | **6.68 ms** | **2750.0000** | **137.86 MB** |


## stable36df-domain1-crypto-before

Capture source: 36df28c4abc6a4c4371e09e26a5e40e0c07d30d2; started 2026-10-08T11:23:30.9046629+00:00; finished 2026-10-08T11:24:08.3615132+00:00. Exact context: contexts/stable36df-domain1-crypto-before-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.EncryptedVerifiedReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-JTLREC : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-ba9",/p:OfficeIMOBenchmarkNewApis=true  
InvocationCount=4  IterationCount=10  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                      | Input             | MemoryInput | Mean     | Error    | StdDev   | Gen0      | Allocated |
|---------------------------- |------------------ |------------ |---------:|---------:|---------:|----------:|----------:|
| **OfficeIMOVerifiedFullFields** | **OriginalSmallXlsx** | **False**       | **141.9 ms** | **19.20 ms** | **12.70 ms** | **2750.0000** | **137.88 MB** |
| **OfficeIMOVerifiedFullFields** | **OriginalSmallXlsx** | **True**        | **140.5 ms** |  **9.24 ms** |  **6.11 ms** | **2750.0000** | **137.86 MB** |


## stable36df-domain0-crypto-core-only

Capture source: 36df28c4abc6a4c4371e09e26a5e40e0c07d30d2; started 2026-10-08T11:22:53.2828125+00:00; finished 2026-10-08T11:23:14.3919018+00:00. Exact context: contexts/stable36df-domain0-crypto-core-only-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.EncryptedVerifiedReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-JWYJYT : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\crypto-only-bundle-4ee00b",/p:OfficeIMOBenchmarkNewApis=true  
InvocationCount=4  IterationCount=10  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                      | Input             | MemoryInput | Mean     | Error    | StdDev   | Allocated |
|---------------------------- |------------------ |------------ |---------:|---------:|---------:|----------:|
| **OfficeIMOVerifiedFullFields** | **OriginalSmallXlsx** | **False**       | **29.60 ms** | **6.495 ms** | **4.296 ms** | **559.46 KB** |
| **OfficeIMOVerifiedFullFields** | **OriginalSmallXlsx** | **True**        | **25.24 ms** | **2.213 ms** | **1.464 ms** | **540.56 KB** |


## stable36df-domain1-crypto-core-only

Capture source: 36df28c4abc6a4c4371e09e26a5e40e0c07d30d2; started 2026-10-08T11:23:14.751291+00:00; finished 2026-10-08T11:23:30.5283718+00:00. Exact context: contexts/stable36df-domain1-crypto-core-only-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.EncryptedVerifiedReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KJBLOG : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\crypto-only-bundle-4ee00b",/p:OfficeIMOBenchmarkNewApis=true  
InvocationCount=4  IterationCount=10  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                      | Input             | MemoryInput | Mean     | Error    | StdDev   | Allocated |
|---------------------------- |------------------ |------------ |---------:|---------:|---------:|----------:|
| **OfficeIMOVerifiedFullFields** | **OriginalSmallXlsx** | **False**       | **24.06 ms** | **2.733 ms** | **1.808 ms** | **559.98 KB** |
| **OfficeIMOVerifiedFullFields** | **OriginalSmallXlsx** | **True**        | **23.23 ms** | **2.128 ms** | **1.407 ms** | **540.12 KB** |

