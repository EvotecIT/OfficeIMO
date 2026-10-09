# Native grouped tables: primary

These are unmodified BenchmarkDotNet Markdown outputs under explicit capture/domain/class headers. Native FullName/job/parameter identities and all warnings/Actuals are in the exact raw packet and JSON/CSV ledgers. Raw API, mapper, metadata, inference and cold diagnostic-allocation boundaries remain as described in the dated report. A native Baseline ratio does not make different API work equivalent.

## qualified-830951-v1-domain0-main-current

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T13:15:47.5336876+00:00; finished 2026-10-08T13:19:18.8093148+00:00. Exact context: contexts/qualified-830951-v1-domain0-main-current-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-OSBIKL : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
InvocationCount=32  IterationCount=15  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method           | RowCount | Shape         | Mean     | Error    | StdDev   | Ratio | RatioSD | Gen0     | Gen1    | Allocated | Alloc Ratio |
|----------------- |--------- |-------------- |---------:|---------:|---------:|------:|--------:|---------:|--------:|----------:|------------:|
| **ExcelReaderTyped** | **50000**    | **Original**      | **12.06 ms** | **0.295 ms** | **0.276 ms** |  **1.00** |    **0.03** |  **93.7500** |       **-** |   **3.88 MB** |        **1.00** |
| OfficeIMOTyped   | 50000    | Original      | 12.09 ms | 0.552 ms | 0.516 ms |  1.00 |    0.05 |  62.5000 |       - |   2.44 MB |        0.63 |
| SylvanTyped      | 50000    | Original      | 62.95 ms | 3.003 ms | 2.809 ms |  5.22 |    0.25 | 250.0000 | 62.5000 |  10.48 MB |        2.70 |
|                  |          |               |          |          |          |       |         |          |         |           |             |
| **ExcelReaderTyped** | **50000**    | **SharedStrings** | **11.07 ms** | **0.766 ms** | **0.717 ms** |  **1.00** |    **0.09** |  **62.5000** |       **-** |   **2.31 MB** |        **1.00** |
| OfficeIMOTyped   | 50000    | SharedStrings | 10.89 ms | 0.816 ms | 0.763 ms |  0.99 |    0.09 |  62.5000 |       - |   2.44 MB |        1.06 |
| SylvanTyped      | 50000    | SharedStrings | 64.60 ms | 2.130 ms | 1.993 ms |  5.86 |    0.40 | 218.7500 | 62.5000 |   8.91 MB |        3.86 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.WriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-OSBIKL : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
InvocationCount=32  IterationCount=15  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method            | RowCount | Mean     | Error    | StdDev   | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated | Alloc Ratio |
|------------------ |--------- |---------:|---------:|---------:|------:|--------:|---------:|---------:|---------:|----------:|------------:|
| OfficeIMO         | 50000    | 18.65 ms | 0.939 ms | 0.878 ms |  1.76 |    0.09 | 312.5000 | 312.5000 | 312.5000 |   4.04 MB |        1.00 |
| ExcelReaderWriter | 50000    | 10.61 ms | 0.309 ms | 0.289 ms |  1.00 |    0.04 | 312.5000 | 312.5000 | 312.5000 |   4.02 MB |        1.00 |


## qualified-830951-v1-domain1-main-current

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T13:19:19.2342881+00:00; finished 2026-10-08T13:22:06.8206995+00:00. Exact context: contexts/qualified-830951-v1-domain1-main-current-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KEXHOX : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
InvocationCount=32  IterationCount=15  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method           | RowCount | Shape         | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1    | Allocated | Alloc Ratio |
|----------------- |--------- |-------------- |----------:|----------:|----------:|------:|--------:|---------:|--------:|----------:|------------:|
| **ExcelReaderTyped** | **50000**    | **Original**      |  **9.659 ms** | **0.4714 ms** | **0.4409 ms** |  **1.00** |    **0.06** |  **93.7500** |       **-** |   **3.88 MB** |        **1.00** |
| OfficeIMOTyped   | 50000    | Original      | 10.616 ms | 0.8994 ms | 0.8413 ms |  1.10 |    0.10 |  62.5000 |       - |   2.44 MB |        0.63 |
| SylvanTyped      | 50000    | Original      | 48.564 ms | 2.1612 ms | 2.0216 ms |  5.04 |    0.30 | 250.0000 | 62.5000 |  10.48 MB |        2.70 |
|                  |          |               |           |           |           |       |         |          |         |           |             |
| **ExcelReaderTyped** | **50000**    | **SharedStrings** |  **8.399 ms** | **0.3552 ms** | **0.3323 ms** |  **1.00** |    **0.06** |  **31.2500** |       **-** |   **2.31 MB** |        **1.00** |
| OfficeIMOTyped   | 50000    | SharedStrings |  8.325 ms | 0.5546 ms | 0.5187 ms |  0.99 |    0.07 |  31.2500 |       - |   2.44 MB |        1.06 |
| SylvanTyped      | 50000    | SharedStrings | 45.500 ms | 2.0857 ms | 1.9510 ms |  5.43 |    0.31 | 156.2500 | 31.2500 |   8.91 MB |        3.86 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.WriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KEXHOX : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
InvocationCount=32  IterationCount=15  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method            | RowCount | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated | Alloc Ratio |
|------------------ |--------- |----------:|----------:|----------:|------:|--------:|---------:|---------:|---------:|----------:|------------:|
| OfficeIMO         | 50000    | 15.193 ms | 0.3901 ms | 0.3649 ms |  1.55 |    0.09 | 312.5000 | 312.5000 | 312.5000 |   4.04 MB |        1.00 |
| ExcelReaderWriter | 50000    |  9.854 ms | 0.5480 ms | 0.5126 ms |  1.00 |    0.07 | 312.5000 | 312.5000 | 312.5000 |   4.02 MB |        1.00 |


## qualified-e676-v1-domain0-main-before

Capture source: e676775a59458d1095606386c39506ceb8a9bc33; started 2026-10-08T12:22:03.9752683+00:00; finished 2026-10-08T12:28:13.8122217+00:00. Exact context: contexts/qualified-e676-v1-domain0-main-before-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-NHIUTL : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\baseline-owner-bundle-78bfbd"  
InvocationCount=32  IterationCount=15  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method           | RowCount | Shape         | Mean      | Error    | StdDev   | Ratio | RatioSD | Gen0     | Gen1    | Allocated | Alloc Ratio |
|----------------- |--------- |-------------- |----------:|---------:|---------:|------:|--------:|---------:|--------:|----------:|------------:|
| **ExcelReaderTyped** | **50000**    | **Original**      |  **14.42 ms** | **0.479 ms** | **0.448 ms** |  **1.00** |    **0.04** |  **62.5000** |       **-** |   **3.87 MB** |        **1.00** |
| OfficeIMOTyped   | 50000    | Original      | 112.13 ms | 3.980 ms | 3.723 ms |  7.78 |    0.34 |  31.2500 |       - |   2.42 MB |        0.63 |
| SylvanTyped      | 50000    | Original      |  65.99 ms | 1.845 ms | 1.726 ms |  4.58 |    0.18 | 187.5000 | 31.2500 |  10.48 MB |        2.71 |
|                  |          |               |           |          |          |       |         |          |         |           |             |
| **ExcelReaderTyped** | **50000**    | **SharedStrings** |  **11.66 ms** | **0.465 ms** | **0.435 ms** |  **1.00** |    **0.05** |  **31.2500** |       **-** |   **2.29 MB** |        **1.00** |
| OfficeIMOTyped   | 50000    | SharedStrings |  71.97 ms | 2.431 ms | 2.274 ms |  6.18 |    0.28 |  62.5000 |       - |   3.62 MB |        1.58 |
| SylvanTyped      | 50000    | SharedStrings |  62.38 ms | 3.748 ms | 3.506 ms |  5.35 |    0.35 | 156.2500 | 31.2500 |   8.91 MB |        3.88 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.WriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-NHIUTL : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\baseline-owner-bundle-78bfbd"  
InvocationCount=32  IterationCount=15  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method            | RowCount | Mean     | Error    | StdDev   | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated | Alloc Ratio |
|------------------ |--------- |---------:|---------:|---------:|------:|--------:|---------:|---------:|---------:|----------:|------------:|
| OfficeIMO         | 50000    | 20.21 ms | 0.626 ms | 0.586 ms |  1.85 |    0.12 | 312.5000 | 312.5000 | 312.5000 |   4.04 MB |        1.00 |
| ExcelReaderWriter | 50000    | 10.98 ms | 0.678 ms | 0.634 ms |  1.00 |    0.08 | 312.5000 | 312.5000 | 312.5000 |   4.02 MB |        1.00 |


## qualified-e676-v1-domain0-main-published

Capture source: e676775a59458d1095606386c39506ceb8a9bc33; started 2026-10-08T12:28:15.0153634+00:00; finished 2026-10-08T12:34:34.9879744+00:00. Exact context: contexts/qualified-e676-v1-domain0-main-published-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-RTSYPM : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkPackageVersion=3.4.4  
InvocationCount=32  IterationCount=15  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method           | RowCount | Shape         | Mean      | Error    | StdDev   | Ratio | RatioSD | Gen0     | Gen1    | Allocated | Alloc Ratio |
|----------------- |--------- |-------------- |----------:|---------:|---------:|------:|--------:|---------:|--------:|----------:|------------:|
| **ExcelReaderTyped** | **50000**    | **Original**      |  **12.70 ms** | **0.465 ms** | **0.435 ms** |  **1.00** |    **0.05** |  **62.5000** |       **-** |   **3.87 MB** |        **1.00** |
| OfficeIMOTyped   | 50000    | Original      | 123.64 ms | 3.908 ms | 3.655 ms |  9.75 |    0.43 |  93.7500 |       - |   5.69 MB |        1.47 |
| SylvanTyped      | 50000    | Original      |  66.36 ms | 2.456 ms | 2.297 ms |  5.23 |    0.25 | 250.0000 | 62.5000 |  10.48 MB |        2.71 |
|                  |          |               |           |          |          |       |         |          |         |           |             |
| **ExcelReaderTyped** | **50000**    | **SharedStrings** |  **10.97 ms** | **0.306 ms** | **0.286 ms** |  **1.00** |    **0.04** |  **31.2500** |       **-** |   **2.31 MB** |        **1.00** |
| OfficeIMOTyped   | 50000    | SharedStrings |  80.22 ms | 2.496 ms | 2.335 ms |  7.32 |    0.28 | 125.0000 | 31.2500 |   6.23 MB |        2.70 |
| SylvanTyped      | 50000    | SharedStrings |  61.15 ms | 5.549 ms | 5.191 ms |  5.58 |    0.48 | 156.2500 | 31.2500 |   8.91 MB |        3.86 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.WriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-RTSYPM : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkPackageVersion=3.4.4  
InvocationCount=32  IterationCount=15  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method            | RowCount | Mean     | Error    | StdDev   | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated | Alloc Ratio |
|------------------ |--------- |---------:|---------:|---------:|------:|--------:|---------:|---------:|---------:|----------:|------------:|
| OfficeIMO         | 50000    | 19.94 ms | 0.406 ms | 0.379 ms |  1.82 |    0.07 | 312.5000 | 312.5000 | 312.5000 |   4.02 MB |        1.00 |
| ExcelReaderWriter | 50000    | 10.95 ms | 0.397 ms | 0.371 ms |  1.00 |    0.05 | 312.5000 | 312.5000 | 312.5000 |   4.02 MB |        1.00 |


## qualified-e676-v1-domain1-main-published

Capture source: e676775a59458d1095606386c39506ceb8a9bc33; started 2026-10-08T12:41:42.6403118+00:00; finished 2026-10-08T12:48:03.2338025+00:00. Exact context: contexts/qualified-e676-v1-domain1-main-published-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-WOOUEG : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkPackageVersion=3.4.4  
InvocationCount=32  IterationCount=15  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method           | RowCount | Shape         | Mean      | Error    | StdDev   | Ratio | RatioSD | Gen0     | Gen1    | Allocated | Alloc Ratio |
|----------------- |--------- |-------------- |----------:|---------:|---------:|------:|--------:|---------:|--------:|----------:|------------:|
| **ExcelReaderTyped** | **50000**    | **Original**      |  **12.19 ms** | **1.292 ms** | **1.209 ms** |  **1.01** |    **0.15** |  **93.7500** |       **-** |   **3.88 MB** |        **1.00** |
| OfficeIMOTyped   | 50000    | Original      | 121.82 ms | 6.124 ms | 5.729 ms | 10.09 |    1.19 | 125.0000 | 31.2500 |   5.77 MB |        1.49 |
| SylvanTyped      | 50000    | Original      |  66.77 ms | 2.985 ms | 2.792 ms |  5.53 |    0.64 | 250.0000 | 62.5000 |  10.48 MB |        2.70 |
|                  |          |               |           |          |          |       |         |          |         |           |             |
| **ExcelReaderTyped** | **50000**    | **SharedStrings** |  **11.00 ms** | **0.660 ms** | **0.617 ms** |  **1.00** |    **0.08** |  **62.5000** |       **-** |   **2.31 MB** |        **1.00** |
| OfficeIMOTyped   | 50000    | SharedStrings |  83.64 ms | 4.617 ms | 4.319 ms |  7.63 |    0.56 | 156.2500 | 31.2500 |   6.23 MB |        2.70 |
| SylvanTyped      | 50000    | SharedStrings |  60.74 ms | 2.514 ms | 2.351 ms |  5.54 |    0.36 | 187.5000 | 31.2500 |   8.91 MB |        3.86 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.WriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-WOOUEG : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkPackageVersion=3.4.4  
InvocationCount=32  IterationCount=15  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method            | RowCount | Mean     | Error    | StdDev   | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated | Alloc Ratio |
|------------------ |--------- |---------:|---------:|---------:|------:|--------:|---------:|---------:|---------:|----------:|------------:|
| OfficeIMO         | 50000    | 18.75 ms | 0.675 ms | 0.631 ms |  1.87 |    0.09 | 312.5000 | 312.5000 | 312.5000 |   4.03 MB |        1.00 |
| ExcelReaderWriter | 50000    | 10.04 ms | 0.372 ms | 0.348 ms |  1.00 |    0.05 | 437.5000 | 437.5000 | 281.2500 |   4.02 MB |        1.00 |


## qualified-e676-v1-domain1-main-before

Capture source: e676775a59458d1095606386c39506ceb8a9bc33; started 2026-10-08T12:48:03.8940029+00:00; finished 2026-10-08T12:54:02.5825726+00:00. Exact context: contexts/qualified-e676-v1-domain1-main-before-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-FCNDUJ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\baseline-owner-bundle-78bfbd"  
InvocationCount=32  IterationCount=15  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method           | RowCount | Shape         | Mean      | Error    | StdDev   | Ratio | RatioSD | Gen0     | Gen1    | Allocated | Alloc Ratio |
|----------------- |--------- |-------------- |----------:|---------:|---------:|------:|--------:|---------:|--------:|----------:|------------:|
| **ExcelReaderTyped** | **50000**    | **Original**      |  **12.13 ms** | **0.593 ms** | **0.554 ms** |  **1.00** |    **0.06** |  **93.7500** |       **-** |   **3.88 MB** |        **1.00** |
| OfficeIMOTyped   | 50000    | Original      | 107.82 ms | 7.392 ms | 6.915 ms |  8.91 |    0.69 |  31.2500 |       - |    2.5 MB |        0.64 |
| SylvanTyped      | 50000    | Original      |  67.63 ms | 2.650 ms | 2.479 ms |  5.59 |    0.32 | 250.0000 | 62.5000 |  10.48 MB |        2.70 |
|                  |          |               |           |          |          |       |         |          |         |           |             |
| **ExcelReaderTyped** | **50000**    | **SharedStrings** |  **10.79 ms** | **0.418 ms** | **0.391 ms** |  **1.00** |    **0.05** |  **62.5000** |       **-** |   **2.31 MB** |        **1.00** |
| OfficeIMOTyped   | 50000    | SharedStrings |  69.57 ms | 3.591 ms | 3.359 ms |  6.45 |    0.38 |  93.7500 | 31.2500 |    3.7 MB |        1.60 |
| SylvanTyped      | 50000    | SharedStrings |  60.61 ms | 2.243 ms | 2.098 ms |  5.62 |    0.27 | 218.7500 | 62.5000 |   8.91 MB |        3.86 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.WriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-FCNDUJ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\baseline-owner-bundle-78bfbd"  
InvocationCount=32  IterationCount=15  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method            | RowCount | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated | Alloc Ratio |
|------------------ |--------- |----------:|----------:|----------:|------:|--------:|---------:|---------:|---------:|----------:|------------:|
| OfficeIMO         | 50000    | 18.477 ms | 0.5820 ms | 0.5444 ms |  1.86 |    0.07 | 312.5000 | 312.5000 | 312.5000 |   4.04 MB |        1.01 |
| ExcelReaderWriter | 50000    |  9.918 ms | 0.2580 ms | 0.2414 ms |  1.00 |    0.03 | 312.5000 | 312.5000 | 312.5000 |   4.02 MB |        1.00 |


## qualified-e676-v1-domain0-crypto-before

Capture source: e676775a59458d1095606386c39506ceb8a9bc33; started 2026-10-08T12:54:03.4079753+00:00; finished 2026-10-08T12:55:05.0553576+00:00. Exact context: contexts/qualified-e676-v1-domain0-crypto-before-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.EncryptedVerifiedReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-OVQHXH : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-ba9",/p:OfficeIMOBenchmarkNewApis=true  
IterationCount=15  IterationTime=250ms  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                      | Input             | MemoryInput | Mean     | Error   | StdDev  | Gen0      | Allocated |
|---------------------------- |------------------ |------------ |---------:|--------:|--------:|----------:|----------:|
| **OfficeIMOVerifiedFullFields** | **OriginalSmallXlsx** | **False**       | **189.9 ms** | **8.48 ms** | **7.93 ms** | **3000.0000** | **137.89 MB** |
| **OfficeIMOVerifiedFullFields** | **OriginalSmallXlsx** | **True**        | **182.8 ms** | **7.93 ms** | **7.42 ms** | **3000.0000** | **137.87 MB** |


## qualified-e676-v1-domain0-crypto-core-only

Capture source: e676775a59458d1095606386c39506ceb8a9bc33; started 2026-10-08T12:55:05.8047833+00:00; finished 2026-10-08T12:55:36.8096909+00:00. Exact context: contexts/qualified-e676-v1-domain0-crypto-core-only-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.EncryptedVerifiedReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-VBBITA : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\crypto-only-bundle-4ee00b",/p:OfficeIMOBenchmarkNewApis=true  
IterationCount=15  IterationTime=250ms  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                      | Input             | MemoryInput | Mean     | Error    | StdDev   | Allocated |
|---------------------------- |------------------ |------------ |---------:|---------:|---------:|----------:|
| **OfficeIMOVerifiedFullFields** | **OriginalSmallXlsx** | **False**       | **28.10 ms** | **2.434 ms** | **2.276 ms** | **558.83 KB** |
| **OfficeIMOVerifiedFullFields** | **OriginalSmallXlsx** | **True**        | **27.53 ms** | **1.656 ms** | **1.549 ms** |  **539.6 KB** |


## qualified-e676-v1-domain1-crypto-before

Capture source: e676775a59458d1095606386c39506ceb8a9bc33; started 2026-10-08T12:56:08.979533+00:00; finished 2026-10-08T12:57:03.1648565+00:00. Exact context: contexts/qualified-e676-v1-domain1-crypto-before-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.EncryptedVerifiedReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-AMRTFP : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-ba9",/p:OfficeIMOBenchmarkNewApis=true  
IterationCount=15  IterationTime=250ms  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                      | Input             | MemoryInput | Mean     | Error   | StdDev  | Gen0      | Allocated |
|---------------------------- |------------------ |------------ |---------:|--------:|--------:|----------:|----------:|
| **OfficeIMOVerifiedFullFields** | **OriginalSmallXlsx** | **False**       | **176.5 ms** | **6.89 ms** | **6.45 ms** | **2750.0000** | **137.88 MB** |
| **OfficeIMOVerifiedFullFields** | **OriginalSmallXlsx** | **True**        | **170.4 ms** | **7.52 ms** | **7.04 ms** | **2750.0000** | **137.86 MB** |


## qualified-e676-v1-domain1-crypto-core-only

Capture source: e676775a59458d1095606386c39506ceb8a9bc33; started 2026-10-08T12:55:38.137892+00:00; finished 2026-10-08T12:56:07.896196+00:00. Exact context: contexts/qualified-e676-v1-domain1-crypto-core-only-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.EncryptedVerifiedReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-FDSVGA : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\crypto-only-bundle-4ee00b",/p:OfficeIMOBenchmarkNewApis=true  
IterationCount=15  IterationTime=250ms  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                      | Input             | MemoryInput | Mean     | Error    | StdDev   | Allocated |
|---------------------------- |------------------ |------------ |---------:|---------:|---------:|----------:|
| **OfficeIMOVerifiedFullFields** | **OriginalSmallXlsx** | **False**       | **27.66 ms** | **1.285 ms** | **1.202 ms** | **557.36 KB** |
| **OfficeIMOVerifiedFullFields** | **OriginalSmallXlsx** | **True**        | **27.30 ms** | **1.502 ms** | **1.405 ms** | **537.97 KB** |


## qualified-830951-long-v1-domain0-generated

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T13:56:36.8577632+00:00; finished 2026-10-08T14:27:48.8889543+00:00. Exact context: contexts/qualified-830951-long-v1-domain0-generated-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.BorrowedRefStructReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KJRLRQ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                     | RowCount | Mean      | Error     | StdDev    | Ratio | RatioSD | Allocated | Alloc Ratio |
|--------------------------- |--------- |----------:|----------:|----------:|------:|--------:|----------:|------------:|
| ExcelReaderMappedRefStruct | 50000    |  9.733 ms | 0.5197 ms | 0.4861 ms |  1.00 |    0.07 |   4.36 KB |        1.00 |
| OfficeIMOFactoryRefStruct  | 50000    | 13.280 ms | 0.3677 ms | 0.3439 ms |  1.37 |    0.07 |   77.3 KB |       17.73 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.MaterializedRawAsyncReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KJRLRQ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                              | RowCount | Format | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1   | Allocated | Alloc Ratio |
|------------------------------------ |--------- |------- |----------:|----------:|----------:|------:|--------:|---------:|-------:|----------:|------------:|
| **ExcelReaderStringsMaterializedAsync** | **50000**    | **Xlsx**   |  **9.980 ms** | **0.3401 ms** | **0.3181 ms** |  **1.00** |    **0.04** |  **31.2500** |      **-** |   **1.58 MB** |        **1.00** |
| OfficeIMOAsync                      | 50000    | Xlsx   | 15.774 ms | 0.7774 ms | 0.7272 ms |  1.58 |    0.09 |  65.5738 |      - |   3.51 MB |        2.22 |
| SylvanAsync                         | 50000    | Xlsx   | 41.083 ms | 4.6902 ms | 4.3872 ms |  4.12 |    0.44 |        - |      - |   1.89 MB |        1.20 |
|                                     |          |        |           |           |           |       |         |          |        |           |             |
| **ExcelReaderStringsMaterializedAsync** | **50000**    | **Xlsb**   |  **4.710 ms** | **0.5389 ms** | **0.5041 ms** |  **1.01** |    **0.13** |  **31.2500** |      **-** |   **1.58 MB** |        **1.00** |
| OfficeIMOAsync                      | 50000    | Xlsb   |  8.950 ms | 1.5159 ms | 1.4180 ms |  1.92 |    0.33 | 106.8702 |      - |   5.12 MB |        3.25 |
| SylvanAsync                         | 50000    | Xlsb   |  8.760 ms | 0.3248 ms | 0.3038 ms |  1.88 |    0.16 |  33.8983 |      - |   1.82 MB |        1.16 |
|                                     |          |        |           |           |           |       |         |          |        |           |             |
| **ExcelReaderStringsMaterializedAsync** | **50000**    | **Xls**    |  **5.236 ms** | **0.1466 ms** | **0.1372 ms** |  **1.00** |    **0.04** |  **33.6538** |      **-** |   **1.58 MB** |        **1.00** |
| OfficeIMOAsync                      | 50000    | Xls    |  7.390 ms | 0.9881 ms | 0.9242 ms |  1.41 |    0.17 | 211.5385 |      - |  10.25 MB |        6.50 |
| SylvanAsync                         | 50000    | Xls    |  6.004 ms | 0.3341 ms | 0.3125 ms |  1.15 |    0.06 |  32.4324 | 5.4054 |   1.68 MB |        1.06 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.MaterializedRawReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KJRLRQ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                         | RowCount | Format | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1   | Allocated | Alloc Ratio |
|------------------------------- |--------- |------- |----------:|----------:|----------:|------:|--------:|---------:|-------:|----------:|------------:|
| **ExcelReaderStringsMaterialized** | **50000**    | **Xlsx**   |  **8.701 ms** | **0.3994 ms** | **0.3736 ms** |  **1.00** |    **0.06** |  **26.5487** |      **-** |   **1.58 MB** |        **1.00** |
| OfficeIMOOriginal              | 50000    | Xlsx   | 15.467 ms | 0.5782 ms | 0.5408 ms |  1.78 |    0.09 |  69.4444 |      - |   3.51 MB |        2.22 |
| Sylvan                         | 50000    | Xlsx   | 44.067 ms | 2.3771 ms | 2.2235 ms |  5.07 |    0.32 |        - |      - |   1.89 MB |        1.20 |
|                                |          |        |           |           |           |       |         |          |        |           |             |
| **ExcelReaderStringsMaterialized** | **50000**    | **Xlsb**   |  **5.933 ms** | **0.1926 ms** | **0.1802 ms** |  **1.00** |    **0.04** |  **30.1205** |      **-** |   **1.58 MB** |        **1.00** |
| OfficeIMOOriginal              | 50000    | Xlsb   | 10.318 ms | 0.4381 ms | 0.4098 ms |  1.74 |    0.08 | 106.0606 |      - |   5.12 MB |        3.25 |
| Sylvan                         | 50000    | Xlsb   |  9.194 ms | 0.1894 ms | 0.1772 ms |  1.55 |    0.05 |  42.0168 | 8.4034 |   1.82 MB |        1.15 |
|                                |          |        |           |           |           |       |         |          |        |           |             |
| **ExcelReaderStringsMaterialized** | **50000**    | **Xls**    |  **4.763 ms** | **0.1614 ms** | **0.1510 ms** |  **1.00** |    **0.04** |  **30.0429** |      **-** |   **1.58 MB** |        **1.00** |
| OfficeIMOOriginal              | 50000    | Xls    |  8.068 ms | 0.1814 ms | 0.1697 ms |  1.70 |    0.06 | 209.6774 |      - |  10.25 MB |        6.50 |
| Sylvan                         | 50000    | Xls    |  6.753 ms | 0.3606 ms | 0.3373 ms |  1.42 |    0.08 |  28.3688 |      - |   1.68 MB |        1.06 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.RawAsyncReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KJRLRQ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method           | RowCount | Format | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1    | Allocated   | Alloc Ratio |
|----------------- |--------- |------- |----------:|----------:|----------:|------:|--------:|---------:|--------:|------------:|------------:|
| **ExcelReaderAsync** | **50000**    | **Xlsx**   |  **8.900 ms** | **0.1590 ms** | **0.1487 ms** |  **1.00** |    **0.02** |        **-** |       **-** |     **8.39 KB** |        **1.00** |
| OfficeIMOAsync   | 50000    | Xlsx   | 13.111 ms | 1.7681 ms | 1.6539 ms |  1.47 |    0.18 |  70.5882 | 11.7647 |  3621.97 KB |      431.47 |
| SylvanAsync      | 50000    | Xlsx   | 40.077 ms | 4.8203 ms | 4.5089 ms |  4.50 |    0.50 |  27.7778 |       - |  1940.06 KB |      231.11 |
|                  |          |        |           |           |           |       |         |          |         |             |             |
| **ExcelReaderAsync** | **50000**    | **Xlsb**   |  **5.687 ms** | **0.2105 ms** | **0.1969 ms** |  **1.00** |    **0.05** |        **-** |       **-** |     **6.38 KB** |        **1.00** |
| OfficeIMOAsync   | 50000    | Xlsb   |  8.370 ms | 0.9320 ms | 0.8718 ms |  1.47 |    0.16 |  98.5915 |       - |  5246.66 KB |      822.00 |
| SylvanAsync      | 50000    | Xlsb   |  8.156 ms | 1.1360 ms | 1.0626 ms |  1.44 |    0.19 |  34.4828 |       - |  1866.42 KB |      292.41 |
|                  |          |        |           |           |           |       |         |          |         |             |             |
| **ExcelReaderAsync** | **50000**    | **Xls**    |  **4.291 ms** | **0.3084 ms** | **0.2885 ms** |  **1.00** |    **0.09** |        **-** |       **-** |     **2.91 KB** |        **1.00** |
| OfficeIMOAsync   | 50000    | Xls    |  8.265 ms | 0.3991 ms | 0.3733 ms |  1.93 |    0.15 | 213.6752 |       - | 10499.22 KB |    3,612.64 |
| SylvanAsync      | 50000    | Xls    |  7.953 ms | 0.4456 ms | 0.4168 ms |  1.86 |    0.15 |  35.2113 |  7.0423 |  1719.72 KB |      591.73 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.RawReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KJRLRQ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method              | RowCount | Format | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1    | Gen2   | Allocated   | Alloc Ratio |
|-------------------- |--------- |------- |----------:|----------:|----------:|------:|--------:|---------:|--------:|-------:|------------:|------------:|
| **ExcelReaderOriginal** | **50000**    | **Xlsx**   |  **7.793 ms** | **0.4940 ms** | **0.4621 ms** |  **1.00** |    **0.08** |        **-** |       **-** |      **-** |     **3.82 KB** |        **1.00** |
| OfficeIMOOriginal   | 50000    | Xlsx   | 16.744 ms | 0.3873 ms | 0.3623 ms |  2.16 |    0.12 |  90.9091 | 15.1515 |      - |   3630.6 KB |      950.34 |
| Sylvan              | 50000    | Xlsx   | 46.974 ms | 2.5547 ms | 2.3897 ms |  6.05 |    0.44 |  47.6190 |       - |      - |  1939.73 KB |      507.74 |
|                     |          |        |           |           |           |       |         |          |         |        |             |             |
| **ExcelReaderOriginal** | **50000**    | **Xlsb**   |  **5.355 ms** | **0.1972 ms** | **0.1845 ms** |  **1.00** |    **0.05** |        **-** |       **-** |      **-** |     **6.73 KB** |        **1.00** |
| OfficeIMOOriginal   | 50000    | Xlsb   |  9.265 ms | 1.3454 ms | 1.2585 ms |  1.73 |    0.24 | 120.3704 |  9.2593 | 9.2593 |  5322.57 KB |      791.16 |
| Sylvan              | 50000    | Xlsb   |  6.342 ms | 0.7838 ms | 0.7332 ms |  1.19 |    0.14 |  42.5532 |  7.0922 |      - |   1865.6 KB |      277.31 |
|                     |          |        |           |           |           |       |         |          |         |        |             |             |
| **ExcelReaderOriginal** | **50000**    | **Xls**    |  **3.027 ms** | **0.4884 ms** | **0.4569 ms** |  **1.02** |    **0.20** |        **-** |       **-** |      **-** |     **3.13 KB** |        **1.00** |
| OfficeIMOOriginal   | 50000    | Xls    |  6.126 ms | 0.8375 ms | 0.7834 ms |  2.06 |    0.37 | 216.2162 |       - |      - | 10522.68 KB |    3,356.77 |
| Sylvan              | 50000    | Xls    |  6.410 ms | 1.6293 ms | 1.5240 ms |  2.16 |    0.57 |  40.4624 |  5.7803 |      - |  1717.95 KB |      548.03 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedAsyncReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KJRLRQ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                | RowCount | Shape    | Mean     | Error    | StdDev   | Ratio | RatioSD | Gen0     | Allocated | Alloc Ratio |
|---------------------- |--------- |--------- |---------:|---------:|---------:|------:|--------:|---------:|----------:|------------:|
| OfficeIMOTypedAsync   | 50000    | Original | 16.93 ms | 0.768 ms | 0.719 ms |  1.39 |    0.08 |        - |   2.36 MB |        0.61 |
| ExcelReaderTypedAsync | 50000    | Original | 12.20 ms | 0.585 ms | 0.547 ms |  1.00 |    0.06 |  58.8235 |   3.87 MB |        1.00 |
| SylvanTypedAsync      | 50000    | Original | 55.69 ms | 7.514 ms | 7.028 ms |  4.57 |    0.60 | 176.4706 |  10.48 MB |        2.71 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedModelReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KJRLRQ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method               | RowCount | Model  | Mean     | Error    | StdDev   | Ratio | RatioSD | Gen0    | Gen1    | Allocated  | Alloc Ratio |
|--------------------- |--------- |------- |---------:|---------:|---------:|------:|--------:|--------:|--------:|-----------:|------------:|
| **ExcelReaderAutomatic** | **50000**    | **Class**  | **10.86 ms** | **1.655 ms** | **1.548 ms** |  **1.02** |    **0.21** | **71.4286** |       **-** | **3959.57 KB** |        **1.00** |
| OfficeIMOAutomatic   | 50000    | Class  | 17.67 ms | 0.359 ms | 0.335 ms |  1.66 |    0.26 | 61.5385 | 15.3846 | 2460.59 KB |        0.62 |
|                      |          |        |          |          |          |       |         |         |         |            |             |
| **ExcelReaderAutomatic** | **50000**    | **Struct** | **11.79 ms** | **0.955 ms** | **0.893 ms** |  **1.01** |    **0.11** | **37.5000** |       **-** | **1621.62 KB** |        **1.00** |
| OfficeIMOAutomatic   | 50000    | Struct | 15.93 ms | 0.901 ms | 0.842 ms |  1.36 |    0.12 |       - |       - |  171.94 KB |        0.11 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedStreamAsyncReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KJRLRQ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                          | RowCount | Shape    | Mean     | Error    | StdDev   | Ratio | RatioSD | Gen0     | Gen1    | Gen2    | Allocated | Alloc Ratio |
|-------------------------------- |--------- |--------- |---------:|---------:|---------:|------:|--------:|---------:|--------:|--------:|----------:|------------:|
| OfficeIMOStreamOpenTypedAsync   | 50000    | Original | 17.51 ms | 0.602 ms | 0.563 ms |  1.26 |    0.05 |  54.5455 | 18.1818 | 18.1818 |    3.1 MB |        0.80 |
| ExcelReaderStreamOpenTypedAsync | 50000    | Original | 13.86 ms | 0.374 ms | 0.350 ms |  1.00 |    0.03 |  82.1918 |       - |       - |   3.87 MB |        1.00 |
| SylvanStreamOpenTypedAsync      | 50000    | Original | 71.62 ms | 2.835 ms | 2.652 ms |  5.17 |    0.23 | 250.0000 |       - |       - |  10.49 MB |        2.71 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedXlsbAsyncReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KJRLRQ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                    | RowCount | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1    | Gen2    | Allocated | Alloc Ratio |
|-------------------------- |--------- |----------:|----------:|----------:|------:|--------:|---------:|--------:|--------:|----------:|------------:|
| ExcelReaderTypedXlsbAsync | 50000    |  9.191 ms | 0.6246 ms | 0.5843 ms |  1.00 |    0.09 |  79.6460 |       - |       - |   3.87 MB |        1.00 |
| OfficeIMOTypedXlsbAsync   | 50000    | 17.560 ms | 0.9084 ms | 0.8498 ms |  1.92 |    0.14 | 196.4286 | 17.8571 | 17.8571 |    8.7 MB |        2.25 |
| SylvanTypedXlsbAsync      | 50000    | 26.582 ms | 1.7965 ms | 1.6804 ms |  2.90 |    0.25 | 200.0000 | 28.5714 |       - |   10.4 MB |        2.69 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedXlsbReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KJRLRQ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method               | RowCount | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1    | Gen2    | Allocated | Alloc Ratio |
|--------------------- |--------- |----------:|----------:|----------:|------:|--------:|---------:|--------:|--------:|----------:|------------:|
| ExcelReaderTypedXlsb | 50000    |  9.340 ms | 0.4484 ms | 0.4194 ms |  1.00 |    0.06 |  84.7458 |       - |       - |   3.87 MB |        1.00 |
| OfficeIMOTypedXlsb   | 50000    | 10.096 ms | 0.2624 ms | 0.2455 ms |  1.08 |    0.05 |  90.9091 | 11.3636 | 11.3636 |   4.07 MB |        1.05 |
| SylvanTypedXlsb      | 50000    | 25.635 ms | 0.8305 ms | 0.7769 ms |  2.75 |    0.14 | 205.1282 | 25.6410 |       - |   10.4 MB |        2.69 |


## qualified-830951-long-v1-domain0-real

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T14:27:49.7029032+00:00; finished 2026-10-08T14:55:45.8549671+00:00. Exact context: contexts/qualified-830951-long-v1-domain0-real-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.MaterializedRealDataReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KJRLRQ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                         | Format | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated   | Alloc Ratio |
|------------------------------- |------- |----------:|----------:|----------:|------:|--------:|---------:|---------:|---------:|------------:|------------:|
| **ExcelReaderStringsMaterialized** | **Xlsx**   |  **57.50 ms** |  **3.587 ms** |  **3.355 ms** |  **1.00** |    **0.08** |        **-** |        **-** |        **-** |    **17.09 KB** |        **1.00** |
| OfficeIMOBytes                 | Xlsx   |  82.61 ms |  2.787 ms |  2.607 ms |  1.44 |    0.09 | 250.0000 |        - |        - | 13940.02 KB |      815.50 |
| OfficeIMOStream                | Xlsx   |  87.26 ms |  2.369 ms |  2.216 ms |  1.52 |    0.09 | 272.7273 |        - |        - | 20503.83 KB |    1,199.49 |
| Sylvan                         | Xlsx   | 242.15 ms | 10.928 ms | 10.222 ms |  4.22 |    0.29 |        - |        - |        - |   645.53 KB |       37.76 |
|                                |        |           |           |           |       |         |          |          |          |             |             |
| **ExcelReaderStringsMaterialized** | **Xlsm**   |  **58.11 ms** |  **1.427 ms** |  **1.334 ms** |  **1.00** |    **0.03** |        **-** |        **-** |        **-** |    **17.09 KB** |        **1.00** |
| OfficeIMOBytes                 | Xlsm   |  80.43 ms |  4.403 ms |  4.118 ms |  1.38 |    0.08 | 272.7273 |        - |        - | 13939.97 KB |      815.50 |
| OfficeIMOStream                | Xlsm   |  80.43 ms |  2.391 ms |  2.237 ms |  1.38 |    0.05 | 272.7273 |        - |        - | 20503.87 KB |    1,199.50 |
| Sylvan                         | Xlsm   | 219.00 ms |  9.391 ms |  8.784 ms |  3.77 |    0.17 |        - |        - |        - |    644.6 KB |       37.71 |
|                                |        |           |           |           |       |         |          |          |          |             |             |
| **ExcelReaderStringsMaterialized** | **Xlsb**   |  **30.50 ms** |  **0.773 ms** |  **0.723 ms** |  **1.00** |    **0.03** |        **-** |        **-** |        **-** |    **16.65 KB** |        **1.00** |
| OfficeIMOBytes                 | Xlsb   |  31.72 ms |  1.554 ms |  1.454 ms |  1.04 |    0.05 | 266.6667 |        - |        - | 13985.93 KB |      840.07 |
| OfficeIMOStream                | Xlsb   |  37.48 ms |  0.909 ms |  0.850 ms |  1.23 |    0.04 | 451.6129 | 193.5484 |  32.2581 | 18828.52 KB |    1,130.95 |
| Sylvan                         | Xlsb   |  36.64 ms |  3.359 ms |  3.142 ms |  1.20 |    0.10 |        - |        - |        - |   338.87 KB |       20.35 |
|                                |        |           |           |           |       |         |          |          |          |             |             |
| **ExcelReaderStringsMaterialized** | **Xls**    |  **15.95 ms** |  **0.363 ms** |  **0.340 ms** |  **1.00** |    **0.03** |        **-** |        **-** |        **-** |    **21.48 KB** |        **1.00** |
| OfficeIMOBytes                 | Xls    |  21.60 ms |  0.353 ms |  0.330 ms |  1.35 |    0.03 | 279.0698 |        - |        - | 13990.08 KB |      651.17 |
| OfficeIMOStream                | Xls    |  22.75 ms |  1.347 ms |  1.260 ms |  1.43 |    0.08 | 404.7619 | 142.8571 | 142.8571 | 25538.11 KB |    1,188.68 |
| Sylvan                         | Xls    |  25.97 ms |  1.204 ms |  1.126 ms |  1.63 |    0.08 |        - |        - |        - |   189.67 KB |        8.83 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.PrefetchedRealDataReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KJRLRQ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                    | Format | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1    | Gen2    | Allocated   | Alloc Ratio |
|-------------------------- |------- |----------:|----------:|----------:|------:|--------:|---------:|--------:|--------:|------------:|------------:|
| **ExcelReaderPrefetch**       | **Xlsx**   |  **35.32 ms** |  **1.709 ms** |  **1.599 ms** |  **1.00** |    **0.06** |        **-** |       **-** |       **-** |    **46.02 KB** |        **1.00** |
| ExcelReaderMemoryPrefetch | Xlsx   |  34.10 ms |  1.432 ms |  1.339 ms |  0.97 |    0.06 |        - |       - |       - |    25.33 KB |        0.55 |
| OfficeIMOBytes            | Xlsx   |  84.68 ms |  4.560 ms |  4.266 ms |  2.40 |    0.16 | 250.0000 |       - |       - | 15393.67 KB |      334.50 |
| Sylvan                    | Xlsx   | 242.42 ms | 11.967 ms | 11.194 ms |  6.88 |    0.43 |        - |       - |       - |   647.57 KB |       14.07 |
|                           |        |           |           |           |       |         |          |         |         |             |             |
| **ExcelReaderPrefetch**       | **Xlsm**   |  **44.81 ms** |  **6.554 ms** |  **6.130 ms** |  **1.02** |    **0.18** |        **-** |       **-** |       **-** |       **62 KB** |        **1.00** |
| ExcelReaderMemoryPrefetch | Xlsm   |  39.77 ms |  3.342 ms |  3.126 ms |  0.90 |    0.12 |        - |       - |       - |    45.24 KB |        0.73 |
| OfficeIMOBytes            | Xlsm   |  94.67 ms |  2.668 ms |  2.496 ms |  2.14 |    0.25 | 300.0000 |       - |       - | 14957.65 KB |      241.24 |
| Sylvan                    | Xlsm   | 248.34 ms | 10.654 ms |  9.966 ms |  5.63 |    0.68 |        - |       - |       - |   647.63 KB |       10.45 |
|                           |        |           |           |           |       |         |          |         |         |             |             |
| **ExcelReaderPrefetch**       | **Xlsb**   |  **18.36 ms** |  **2.330 ms** |  **2.179 ms** |  **1.01** |    **0.15** |        **-** |       **-** |       **-** |    **28.67 KB** |        **1.00** |
| ExcelReaderMemoryPrefetch | Xlsb   |  16.67 ms |  0.618 ms |  0.578 ms |  0.92 |    0.09 |        - |       - |       - |    17.54 KB |        0.61 |
| OfficeIMOBytes            | Xlsb   |  35.35 ms |  1.481 ms |  1.386 ms |  1.95 |    0.20 | 333.3333 | 37.0370 | 37.0370 | 14593.86 KB |      509.05 |
| Sylvan                    | Xlsb   |  36.87 ms |  0.641 ms |  0.599 ms |  2.03 |    0.20 |        - |       - |       - |   339.35 KB |       11.84 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.RealDataReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KJRLRQ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method              | Format | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated   | Alloc Ratio |
|-------------------- |------- |----------:|----------:|----------:|------:|--------:|---------:|---------:|---------:|------------:|------------:|
| **ExcelReaderOriginal** | **Xlsx**   |  **57.55 ms** |  **1.242 ms** |  **1.161 ms** |  **1.00** |    **0.03** |        **-** |        **-** |        **-** |    **35.27 KB** |        **1.00** |
| ExcelReaderMemory   | Xlsx   |  53.11 ms |  6.072 ms |  5.680 ms |  0.92 |    0.10 |        - |        - |        - |    23.46 KB |        0.67 |
| OfficeIMOBytes      | Xlsx   |  92.50 ms |  6.087 ms |  5.694 ms |  1.61 |    0.10 | 272.7273 |        - |        - |  14865.1 KB |      421.49 |
| OfficeIMOStream     | Xlsx   |  91.86 ms |  4.586 ms |  4.289 ms |  1.60 |    0.08 | 272.7273 |        - |        - | 21428.96 KB |      607.61 |
| Sylvan              | Xlsx   | 193.65 ms | 35.875 ms | 33.557 ms |  3.37 |    0.57 |        - |        - |        - |   647.55 KB |       18.36 |
|                     |        |           |           |           |       |         |          |          |          |             |             |
| **ExcelReaderOriginal** | **Xlsm**   |  **48.14 ms** |  **6.248 ms** |  **5.844 ms** |  **1.01** |    **0.17** |        **-** |        **-** |        **-** |    **26.76 KB** |        **1.00** |
| ExcelReaderMemory   | Xlsm   |  53.22 ms |  5.236 ms |  4.898 ms |  1.12 |    0.16 |        - |        - |        - |     7.38 KB |        0.28 |
| OfficeIMOBytes      | Xlsm   |  83.85 ms |  9.574 ms |  8.956 ms |  1.77 |    0.28 | 333.3333 |  83.3333 |        - | 14787.98 KB |      552.58 |
| OfficeIMOStream     | Xlsm   |  67.95 ms | 15.439 ms | 14.442 ms |  1.43 |    0.34 | 312.5000 |  62.5000 |        - | 21139.92 KB |      789.93 |
| Sylvan              | Xlsm   | 161.55 ms | 32.703 ms | 30.590 ms |  3.40 |    0.74 |        - |        - |        - |   647.63 KB |       24.20 |
|                     |        |           |           |           |       |         |          |          |          |             |             |
| **ExcelReaderOriginal** | **Xlsb**   |  **25.21 ms** |  **2.173 ms** |  **2.032 ms** |  **1.01** |    **0.10** |        **-** |        **-** |        **-** |    **19.14 KB** |        **1.00** |
| ExcelReaderMemory   | Xlsb   |  28.00 ms |  1.702 ms |  1.592 ms |  1.12 |    0.10 |        - |        - |        - |    16.44 KB |        0.86 |
| OfficeIMOBytes      | Xlsb   |  34.11 ms |  2.911 ms |  2.723 ms |  1.36 |    0.14 | 394.7368 | 157.8947 |  26.3158 | 14418.03 KB |      753.42 |
| OfficeIMOStream     | Xlsb   |  37.78 ms |  1.135 ms |  1.061 ms |  1.51 |    0.11 | 296.2963 |  37.0370 |  37.0370 | 18985.34 KB |      992.09 |
| Sylvan              | Xlsb   |  37.13 ms |  1.793 ms |  1.677 ms |  1.48 |    0.12 |        - |        - |        - |   338.87 KB |       17.71 |
|                     |        |           |           |           |       |         |          |          |          |             |             |
| **ExcelReaderOriginal** | **Xls**    |  **13.82 ms** |  **0.479 ms** |  **0.448 ms** |  **1.00** |    **0.04** |        **-** |        **-** |        **-** |    **12.02 KB** |        **1.00** |
| ExcelReaderMemory   | Xls    |  15.47 ms |  0.337 ms |  0.316 ms |  1.12 |    0.04 |        - |        - |        - |    11.88 KB |        0.99 |
| OfficeIMOBytes      | Xls    |  19.43 ms |  1.043 ms |  0.975 ms |  1.41 |    0.08 | 274.5098 |        - |        - | 14316.37 KB |    1,191.48 |
| OfficeIMOStream     | Xls    |  24.37 ms |  0.912 ms |  0.853 ms |  1.77 |    0.08 | 450.0000 | 175.0000 | 175.0000 | 26408.83 KB |    2,197.87 |
| Sylvan              | Xls    |  26.16 ms |  1.009 ms |  0.944 ms |  1.90 |    0.09 |        - |        - |        - |    186.3 KB |       15.50 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.RealDataTypedReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KJRLRQ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method               | Format | Mean     | Error    | StdDev   | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated | Alloc Ratio |
|--------------------- |------- |---------:|---------:|---------:|------:|--------:|---------:|---------:|---------:|----------:|------------:|
| **ExcelReaderAutomatic** | **Xlsx**   | **66.24 ms** | **2.546 ms** | **2.382 ms** |  **1.00** |    **0.05** | **125.0000** |        **-** |        **-** |   **8.02 MB** |        **1.00** |
| OfficeIMOAutomatic   | Xlsx   | 94.64 ms | 2.018 ms | 1.888 ms |  1.43 |    0.06 | 200.0000 |        - |        - |   9.11 MB |        1.14 |
|                      |        |          |          |          |       |         |          |          |          |           |             |
| **ExcelReaderAutomatic** | **Xlsb**   | **39.79 ms** | **0.991 ms** | **0.927 ms** |  **1.00** |    **0.03** | **166.6667** |        **-** |        **-** |   **8.04 MB** |        **1.00** |
| OfficeIMOAutomatic   | Xlsb   | 56.44 ms | 1.663 ms | 1.556 ms |  1.42 |    0.05 | 750.0000 | 250.0000 | 250.0000 |  28.67 MB |        3.57 |


## qualified-830951-long-v1-domain0-strings

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T14:55:47.4313992+00:00; finished 2026-10-08T15:17:40.9451379+00:00. Exact context: contexts/qualified-830951-long-v1-domain0-strings-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.MaterializedStringHeavyReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-BHAAQB : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                         | Format | RowCount | Mean      | Error    | StdDev   | Ratio | RatioSD | Gen0      | Gen1      | Gen2     | Allocated | Alloc Ratio |
|------------------------------- |------- |--------- |----------:|---------:|---------:|------:|--------:|----------:|----------:|---------:|----------:|------------:|
| **ExcelReaderStringsMaterialized** | **Xlsx**   | **65536**    |  **69.11 ms** | **2.683 ms** | **2.095 ms** |  **1.00** |    **0.04** |  **333.3333** |  **266.6667** |  **66.6667** |  **15.25 MB** |        **1.00** |
| ExcelReaderStringsInterned     | Xlsx   | 65536    |  75.56 ms | 4.258 ms | 3.324 ms |  1.09 |    0.06 |  714.2857 |  642.8571 | 214.2857 |  15.89 MB |        1.04 |
| OfficeIMOBytes                 | Xlsx   | 65536    | 136.58 ms | 8.713 ms | 6.803 ms |  1.98 |    0.11 |  875.0000 |  750.0000 | 250.0000 |  25.23 MB |        1.65 |
| OfficeIMOStream                | Xlsx   | 65536    | 125.66 ms | 9.431 ms | 7.363 ms |  1.82 |    0.12 |  875.0000 |  750.0000 | 250.0000 |  31.57 MB |        2.07 |
| Sylvan                         | Xlsx   | 65536    | 219.55 ms | 9.687 ms | 7.563 ms |  3.18 |    0.14 |  250.0000 |         - |        - |  17.41 MB |        1.14 |
|                                |        |          |           |          |          |       |         |           |           |          |           |             |
| **ExcelReaderStringsMaterialized** | **Xlsb**   | **65536**    |  **62.71 ms** | **4.303 ms** | **3.360 ms** |  **1.00** |    **0.07** |  **333.3333** |  **266.6667** |  **66.6667** |  **15.25 MB** |        **1.00** |
| ExcelReaderStringsInterned     | Xlsb   | 65536    |  72.33 ms | 3.152 ms | 2.461 ms |  1.16 |    0.07 |  687.5000 |  625.0000 | 187.5000 |   17.7 MB |        1.16 |
| OfficeIMOBytes                 | Xlsb   | 65536    |  75.65 ms | 4.051 ms | 3.163 ms |  1.21 |    0.08 | 1076.9231 | 1000.0000 | 307.6923 |  36.48 MB |        2.39 |
| OfficeIMOStream                | Xlsb   | 65536    |  70.82 ms | 2.556 ms | 1.996 ms |  1.13 |    0.06 | 1076.9231 | 1000.0000 | 384.6154 |  46.69 MB |        3.06 |
| Sylvan                         | Xlsb   | 65536    |  63.84 ms | 9.496 ms | 7.414 ms |  1.02 |    0.13 |  733.3333 |  666.6667 | 200.0000 |  17.38 MB |        1.14 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.SharedStringBorrowedScanBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-BHAAQB : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                      | RowCount | Storage  | Mean     | Error    | StdDev   | Ratio | RatioSD | Allocated | Alloc Ratio |
|---------------------------- |--------- |--------- |---------:|---------:|---------:|------:|--------:|----------:|------------:|
| **ExcelReaderBorrowed**         | **65536**    | **Deflated** | **46.11 ms** | **1.356 ms** | **1.059 ms** |  **1.00** |    **0.03** |   **2.56 MB** |        **1.00** |
| ExcelReaderBorrowedPrefetch | 65536    | Deflated | 34.42 ms | 1.046 ms | 0.816 ms |  0.75 |    0.02 |   2.55 MB |        0.99 |
|                             |          |          |          |          |          |       |         |           |             |
| **ExcelReaderBorrowed**         | **65536**    | **Stored**   | **23.71 ms** | **1.057 ms** | **0.826 ms** |  **1.00** |    **0.05** |   **2.38 MB** |        **1.00** |
| ExcelReaderBorrowedPrefetch | 65536    | Stored   | 23.91 ms | 1.252 ms | 0.977 ms |  1.01 |    0.05 |   2.42 MB |        1.02 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.SharedStringFirstRowBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-BHAAQB : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                         | RowCount | Storage  | Mean         | Error        | StdDev       | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated   | Alloc Ratio |
|------------------------------- |--------- |--------- |-------------:|-------------:|-------------:|------:|--------:|---------:|---------:|---------:|------------:|------------:|
| **ExcelReaderOpenThroughFirstRow** | **65536**    | **Deflated** | **16,308.50 μs** |   **297.283 μs** |   **232.099 μs** | **1.000** |    **0.02** |        **-** |        **-** |        **-** |   **886.92 KB** |        **1.00** |
| OfficeIMOOpenThroughFirstRow   | 65536    | Deflated | 93,223.37 μs | 4,794.746 μs | 3,743.423 μs | 5.717 |    0.23 | 636.3636 | 545.4545 | 181.8182 | 20464.23 KB |       23.07 |
| SylvanOpenThroughFirstRow      | 65536    | Deflated |    100.06 μs |     3.869 μs |     3.021 μs | 0.006 |    0.00 |   7.5512 |   4.0661 |   2.9873 |   321.13 KB |        0.36 |
|                                |          |          |              |              |              |       |         |          |          |          |             |             |
| **ExcelReaderOpenThroughFirstRow** | **65536**    | **Stored**   |  **4,215.24 μs** |   **117.410 μs** |    **91.666 μs** |  **1.00** |    **0.03** |   **4.0816** |   **4.0816** |   **4.0816** |   **782.37 KB** |        **1.00** |
| OfficeIMOOpenThroughFirstRow   | 65536    | Stored   | 71,484.95 μs | 2,565.155 μs | 2,002.704 μs | 16.97 |    0.58 | 714.2857 | 642.8571 | 214.2857 | 20716.42 KB |       26.48 |
| SylvanOpenThroughFirstRow      | 65536    | Stored   |     88.62 μs |    15.592 μs |    12.173 μs |  0.02 |    0.00 |   6.2764 |   2.8653 |   1.9102 |   320.09 KB |        0.41 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.SharedStringFullScanBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-BHAAQB : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                         | RowCount | Storage | Mean      | Error    | StdDev   | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated | Alloc Ratio |
|------------------------------- |--------- |-------- |----------:|---------:|---------:|------:|--------:|---------:|---------:|---------:|----------:|------------:|
| ExcelReaderStringsMaterialized | 65536    | Stored  |  43.74 ms | 2.090 ms | 1.632 ms |  1.00 |    0.05 | 320.0000 | 280.0000 |  80.0000 |  15.25 MB |        1.00 |
| OfficeIMOStringsMaterialized   | 65536    | Stored  |  91.61 ms | 3.108 ms | 2.427 ms |  2.10 |    0.09 | 500.0000 | 400.0000 | 100.0000 |  23.06 MB |        1.51 |
| SylvanStringsMaterialized      | 65536    | Stored  | 202.68 ms | 8.764 ms | 6.842 ms |  4.64 |    0.22 | 250.0000 |        - |        - |  17.41 MB |        1.14 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.SharedStringUtf8ReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-BHAAQB : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                      | Format | RowCount | Storage  | Mean      | Error     | StdDev   | Ratio | RatioSD | Gen0      | Gen1      | Gen2     | Allocated | Alloc Ratio |
|---------------------------- |------- |--------- |--------- |----------:|----------:|---------:|------:|--------:|----------:|----------:|---------:|----------:|------------:|
| **ExcelReaderBorrowed**         | **Xlsx**   | **100000**   | **Deflated** |  **66.66 ms** |  **1.743 ms** | **1.361 ms** |  **1.00** |    **0.03** |         **-** |         **-** |        **-** |   **3.99 MB** |        **1.00** |
| ExcelReaderBorrowedPrefetch | Xlsx   | 100000   | Deflated |  44.20 ms |  1.091 ms | 0.852 ms |  0.66 |    0.02 |         - |         - |        - |      3 MB |        0.75 |
| OfficeIMOBorrowedSharedUtf8 | Xlsx   | 100000   | Deflated | 232.66 ms |  5.665 ms | 4.422 ms |  3.49 |    0.09 | 2000.0000 | 1000.0000 | 250.0000 |  89.14 MB |       22.33 |
|                             |        |          |          |           |           |          |       |         |           |           |          |           |             |
| **ExcelReaderBorrowed**         | **Xlsx**   | **100000**   | **Stored**   |  **33.34 ms** |  **1.779 ms** | **1.389 ms** |  **1.00** |    **0.06** |         **-** |         **-** |        **-** |   **2.89 MB** |        **1.00** |
| ExcelReaderBorrowedPrefetch | Xlsx   | 100000   | Stored   |  34.15 ms |  3.765 ms | 2.939 ms |  1.03 |    0.09 |         - |         - |        - |   2.91 MB |        1.01 |
| OfficeIMOBorrowedSharedUtf8 | Xlsx   | 100000   | Stored   | 195.68 ms | 10.104 ms | 7.889 ms |  5.88 |    0.32 | 2250.0000 | 1750.0000 | 500.0000 |  88.74 MB |       30.66 |
|                             |        |          |          |           |           |          |       |         |           |           |          |           |             |
| **ExcelReaderBorrowed**         | **Xlsb**   | **100000**   | **Deflated** |  **55.14 ms** |  **2.043 ms** | **1.595 ms** |  **1.00** |    **0.04** |         **-** |         **-** |        **-** |   **3.84 MB** |        **1.00** |
| ExcelReaderBorrowedPrefetch | Xlsb   | 100000   | Deflated |  37.03 ms |  1.700 ms | 1.328 ms |  0.67 |    0.03 |         - |         - |        - |   3.62 MB |        0.94 |
| OfficeIMOBorrowedSharedUtf8 | Xlsb   | 100000   | Deflated | 154.87 ms | 12.628 ms | 9.859 ms |  2.81 |    0.19 | 3428.5714 | 2285.7143 | 714.2857 | 112.59 MB |       29.35 |
|                             |        |          |          |           |           |          |       |         |           |           |          |           |             |
| **ExcelReaderBorrowed**         | **Xlsb**   | **100000**   | **Stored**   |  **27.69 ms** |  **0.337 ms** | **0.263 ms** |  **1.00** |    **0.01** |         **-** |         **-** |        **-** |   **2.89 MB** |        **1.00** |
| ExcelReaderBorrowedPrefetch | Xlsb   | 100000   | Stored   |  34.03 ms |  4.420 ms | 3.451 ms |  1.23 |    0.12 |         - |         - |        - |   3.38 MB |        1.17 |
| OfficeIMOBorrowedSharedUtf8 | Xlsb   | 100000   | Stored   | 119.64 ms |  7.939 ms | 6.198 ms |  4.32 |    0.22 | 3375.0000 | 2250.0000 | 750.0000 | 111.44 MB |       38.50 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.StringHeavyReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-BHAAQB : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method              | Format | RowCount | Mean      | Error    | StdDev   | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated | Alloc Ratio |
|-------------------- |------- |--------- |----------:|---------:|---------:|------:|--------:|---------:|---------:|---------:|----------:|------------:|
| **ExcelReaderOriginal** | **Xlsx**   | **65536**    |  **46.18 ms** | **1.950 ms** | **1.522 ms** |  **1.00** |    **0.04** |        **-** |        **-** |        **-** |   **2.18 MB** |        **1.00** |
| ExcelReaderPrefetch | Xlsx   | 65536    |  26.08 ms | 2.964 ms | 2.314 ms |  0.57 |    0.05 |        - |        - |        - |   2.26 MB |        1.04 |
| ExcelReaderMemory   | Xlsx   | 65536    |  37.89 ms | 7.451 ms | 5.817 ms |  0.82 |    0.12 |        - |        - |        - |   2.18 MB |        1.00 |
| OfficeIMOBytes      | Xlsx   | 65536    | 101.56 ms | 3.162 ms | 2.469 ms |  2.20 |    0.09 | 250.0000 |        - |        - |  19.11 MB |        8.76 |
| OfficeIMOStream     | Xlsx   | 65536    | 106.82 ms | 3.964 ms | 3.094 ms |  2.32 |    0.10 | 444.4444 | 333.3333 | 111.1111 |  25.45 MB |       11.67 |
| Sylvan              | Xlsx   | 65536    | 219.17 ms | 7.714 ms | 6.022 ms |  4.75 |    0.20 | 250.0000 |        - |        - |  17.41 MB |        7.98 |
|                     |        |          |           |          |          |       |         |          |          |          |           |             |
| **ExcelReaderOriginal** | **Xlsb**   | **65536**    |  **40.01 ms** | **1.830 ms** | **1.428 ms** |  **1.00** |    **0.05** |        **-** |        **-** |        **-** |   **2.18 MB** |        **1.00** |
| ExcelReaderPrefetch | Xlsb   | 65536    |  31.16 ms | 2.287 ms | 1.786 ms |  0.78 |    0.05 |        - |        - |        - |   2.66 MB |        1.22 |
| ExcelReaderMemory   | Xlsb   | 65536    |  39.49 ms | 0.986 ms | 0.770 ms |  0.99 |    0.04 |  41.6667 |  41.6667 |  41.6667 |   7.14 MB |        3.28 |
| OfficeIMOBytes      | Xlsb   | 65536    |  63.53 ms | 2.574 ms | 2.009 ms |  1.59 |    0.07 | 466.6667 | 400.0000 | 133.3333 |  31.55 MB |       14.47 |
| OfficeIMOStream     | Xlsb   | 65536    |  69.33 ms | 2.511 ms | 1.961 ms |  1.73 |    0.08 | 785.7143 | 714.2857 | 357.1429 |  46.16 MB |       21.18 |
| Sylvan              | Xlsb   | 65536    |  68.32 ms | 5.727 ms | 4.471 ms |  1.71 |    0.12 | 250.0000 |        - |        - |  17.38 MB |        7.98 |


## qualified-830951-long-v1-domain0-ado

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T15:17:41.5151971+00:00; finished 2026-10-08T15:22:08.9906084+00:00. Exact context: contexts/qualified-830951-long-v1-domain0-ado-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.AdoReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KJRLRQ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method      | RowCount | Access       | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1    | Gen2    | Allocated  | Alloc Ratio |
|------------ |--------- |------------- |----------:|----------:|----------:|------:|--------:|---------:|--------:|--------:|-----------:|------------:|
| **ExcelReader** | **50000**    | **GetValue**     | **10.359 ms** | **0.6254 ms** | **0.5850 ms** |  **1.00** |    **0.08** | **108.6957** |       **-** |       **-** | **5136.55 KB** |        **1.00** |
| OfficeIMO   | 50000    | GetValue     | 14.931 ms | 0.7828 ms | 0.7322 ms |  1.45 |    0.10 |  52.6316 |       - |       - | 4482.33 KB |        0.87 |
| Sylvan      | 50000    | GetValue     | 37.024 ms | 3.6508 ms | 3.4150 ms |  3.58 |    0.37 | 157.8947 |       - |       - | 8382.87 KB |        1.63 |
|             |          |              |           |           |           |       |         |          |         |         |            |             |
| **ExcelReader** | **50000**    | **TypedGetters** |  **8.511 ms** | **1.1360 ms** | **1.0626 ms** |  **1.02** |    **0.18** |  **36.2319** |       **-** |       **-** | **1619.25 KB** |        **1.00** |
| OfficeIMO   | 50000    | TypedGetters | 12.573 ms | 0.4880 ms | 0.4565 ms |  1.50 |    0.19 |        - |       - |       - |  833.01 KB |        0.51 |
| Sylvan      | 50000    | TypedGetters | 36.204 ms | 7.1431 ms | 6.6817 ms |  4.32 |    0.95 |        - |       - |       - | 1939.62 KB |        1.20 |
|             |          |              |           |           |           |       |         |          |         |         |            |             |
| **ExcelReader** | **50000**    | **Utf8TextCopy** |  **8.130 ms** | **0.5258 ms** | **0.4918 ms** |  **1.00** |    **0.08** |        **-** |       **-** |       **-** |    **4.55 KB** |        **1.00** |
| OfficeIMO   | 50000    | Utf8TextCopy |  9.597 ms | 0.4454 ms | 0.4166 ms |  1.18 |    0.08 |  31.5789 | 31.5789 | 31.5789 |  832.98 KB |      182.88 |
| Sylvan      | 50000    | Utf8TextCopy | 38.674 ms | 1.4343 ms | 1.3417 ms |  4.77 |    0.31 |  76.9231 |       - |       - | 4283.37 KB |      940.43 |


## qualified-830951-long-v2-domain0-writing

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T16:45:44.1934617+00:00; finished 2026-10-08T17:02:32.4584707+00:00. Exact context: contexts/qualified-830951-long-v2-domain0-writing-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.ArrowStringWriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-BHAAQB : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                          | Mean     | Error    | StdDev   | Ratio | RatioSD | Allocated  | Alloc Ratio |
|-------------------------------- |---------:|---------:|---------:|------:|--------:|-----------:|------------:|
| ExcelReaderRecordBatch          | 25.03 ms | 2.092 ms | 1.633 ms |  1.00 |    0.09 |   18.31 KB |        1.00 |
| OfficeIMOPublicRowOrchestration | 56.64 ms | 4.478 ms | 3.496 ms |  2.27 |    0.20 | 1121.51 KB |       61.24 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.CompactWriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-BHAAQB : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                     | RowCount | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated | Alloc Ratio |
|--------------------------- |--------- |----------:|----------:|----------:|------:|--------:|---------:|---------:|---------:|----------:|------------:|
| OfficeIMOWithoutReferences | 50000    | 10.371 ms | 0.3437 ms | 0.2683 ms |  1.08 |    0.06 | 326.9231 | 326.9231 | 326.9231 |   4.04 MB |        1.00 |
| ExcelReaderWriter          | 50000    |  9.610 ms | 0.6535 ms | 0.5102 ms |  1.00 |    0.07 | 329.7872 | 329.7872 | 329.7872 |   4.02 MB |        1.00 |
| ExcelReaderWriterPrefetch  | 50000    |  6.739 ms | 0.4017 ms | 0.3136 ms |  0.70 |    0.05 | 329.1139 | 329.1139 | 329.1139 |   4.06 MB |        1.01 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.ConfiguredXlsbWriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-BHAAQB : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                             | RowCount | Mean       | Error       | StdDev      | Ratio  | RatioSD | Gen0       | Gen1       | Gen2      | Allocated | Alloc Ratio |
|----------------------------------- |--------- |-----------:|------------:|------------:|-------:|--------:|-----------:|-----------:|----------:|----------:|------------:|
| OfficeIMOModelAndSaveInlineStrings | 50000    | 864.295 ms | 110.3692 ms |  86.1690 ms | 120.27 |   16.20 | 14000.0000 | 10000.0000 | 2000.0000 | 439.99 MB |      108.44 |
| OfficeIMOModelAndSaveSharedStrings | 50000    | 894.027 ms | 172.7214 ms | 134.8495 ms | 124.41 |   21.58 |  9500.0000 |  8000.0000 | 2250.0000 | 430.12 MB |      106.00 |
| ExcelReaderXlsbWriterSharedStrings | 50000    |   7.247 ms |   0.8654 ms |   0.6757 ms |   1.01 |    0.13 |   330.6452 |   330.6452 |  330.6452 |   4.06 MB |        1.00 |
| ExcelReaderXlsbWriterPrefetch      | 50000    |   4.681 ms |   0.1547 ms |   0.1208 ms |   0.65 |    0.06 |   333.3333 |   333.3333 |  333.3333 |   4.03 MB |        0.99 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.MappedRecordWriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-BHAAQB : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                    | Format | RowCount | Mean       | Error       | StdDev     | Ratio  | RatioSD | Gen0       | Gen1       | Gen2      | Allocated | Alloc Ratio |
|-------------------------- |------- |--------- |-----------:|------------:|-----------:|-------:|--------:|-----------:|-----------:|----------:|----------:|------------:|
| **OfficeIMOPublicTypedWrite** | **Xlsx**   | **50000**    |  **16.992 ms** |   **1.2190 ms** |  **0.9517 ms** |   **1.62** |    **0.20** |   **246.1538** |   **246.1538** |  **246.1538** |   **4.04 MB** |        **1.00** |
| ExcelReaderMappedRecords  | Xlsx   | 50000    |  10.616 ms |   1.5680 ms |  1.2242 ms |   1.01 |    0.16 |   242.4242 |   242.4242 |  242.4242 |   4.02 MB |        1.00 |
|                           |        |          |            |             |            |        |         |            |            |           |           |             |
| **OfficeIMOPublicTypedWrite** | **Xlsb**   | **50000**    | **910.491 ms** | **127.6255 ms** | **99.6416 ms** | **123.57** |   **14.29** | **14500.0000** | **10750.0000** | **1750.0000** | **439.99 MB** |      **109.43** |
| ExcelReaderMappedRecords  | Xlsb   | 50000    |   7.385 ms |   0.4742 ms |  0.3702 ms |   1.00 |    0.07 |   330.3571 |   330.3571 |  330.3571 |   4.02 MB |        1.00 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.NativeWriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-BHAAQB : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                | Format | RowCount | Mean       | Error      | StdDev     | Ratio | RatioSD | Gen0      | Gen1      | Gen2      | Allocated | Alloc Ratio |
|---------------------- |------- |--------- |-----------:|-----------:|-----------:|------:|--------:|----------:|----------:|----------:|----------:|------------:|
| OfficeIMOModelAndSave | Xlsb   | 50000    | 857.791 ms | 62.2843 ms | 48.6275 ms | 97.27 |    6.26 | 9500.0000 | 8000.0000 | 2250.0000 | 439.72 MB |      109.30 |
| ExcelReaderWriter     | Xlsb   | 50000    |   8.828 ms |  0.3897 ms |  0.3043 ms |  1.00 |    0.05 |  355.9322 |  355.9322 |  313.5593 |   4.02 MB |        1.00 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.RecordWriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-BHAAQB : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                    | Format | RowCount | Mean       | Error       | StdDev     | Ratio  | RatioSD | Gen0       | Gen1       | Gen2      | Allocated | Alloc Ratio |
|-------------------------- |------- |--------- |-----------:|------------:|-----------:|-------:|--------:|-----------:|-----------:|----------:|----------:|------------:|
| **OfficeIMOPublicTypedWrite** | **Xlsx**   | **50000**    |  **22.134 ms** |   **0.4378 ms** |  **0.3418 ms** |   **1.72** |    **0.05** |   **244.4444** |   **244.4444** |  **244.4444** |   **4.04 MB** |        **1.00** |
| ExcelReaderWriteRecords   | Xlsx   | 50000    |  12.838 ms |   0.3817 ms |  0.2980 ms |   1.00 |    0.03 |   243.5897 |   243.5897 |  243.5897 |   4.02 MB |        1.00 |
|                           |        |          |            |             |            |        |         |            |            |           |           |             |
| **OfficeIMOPublicTypedWrite** | **Xlsb**   | **50000**    | **869.726 ms** | **112.3029 ms** | **87.6787 ms** | **130.41** |   **18.09** | **14000.0000** | **10000.0000** | **2000.0000** | **439.99 MB** |      **109.36** |
| ExcelReaderWriteRecords   | Xlsb   | 50000    |   6.732 ms |   0.8590 ms |  0.6707 ms |   1.01 |    0.14 |   330.9859 |   330.9859 |  330.9859 |   4.02 MB |        1.00 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.SharedStringWriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-BHAAQB : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                                  | RowCount | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated | Alloc Ratio |
|---------------------------------------- |--------- |----------:|----------:|----------:|------:|--------:|---------:|---------:|---------:|----------:|------------:|
| OfficeIMOSharedStrings                  | 50000    | 18.568 ms | 0.6220 ms | 0.4856 ms |  1.84 |    0.08 | 327.8689 | 327.8689 | 327.8689 |   4.04 MB |        1.00 |
| OfficeIMOSharedStringsWithoutReferences | 50000    |  9.126 ms | 0.5840 ms | 0.4559 ms |  0.90 |    0.05 | 330.2752 | 330.2752 | 330.2752 |   4.04 MB |        1.00 |
| ExcelReaderWriterSharedStrings          | 50000    | 10.130 ms | 0.4919 ms | 0.3840 ms |  1.00 |    0.05 | 310.3448 | 310.3448 | 310.3448 |   4.06 MB |        1.00 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.StyledRowWriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-BHAAQB : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                        | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0    | Allocated  | Alloc Ratio |
|------------------------------ |----------:|----------:|----------:|------:|--------:|--------:|-----------:|------------:|
| OfficeIMODefaultRowStyle      | 11.407 ms | 0.6634 ms | 0.5179 ms |  1.84 |    0.18 | 45.4545 | 2562.91 KB |      140.86 |
| ExcelReaderOriginalStyledRows |  6.227 ms | 0.6672 ms | 0.5209 ms |  1.01 |    0.12 |       - |    18.2 KB |        1.00 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.Utf8WriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-BHAAQB : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                | Mean     | Error    | StdDev   | Ratio | RatioSD | Allocated | Alloc Ratio |
|---------------------- |---------:|---------:|---------:|------:|--------:|----------:|------------:|
| OfficeIMOUtf8         | 35.35 ms | 3.014 ms | 2.353 ms |  2.18 |    0.17 |  41.96 KB |        2.33 |
| ExcelReaderPublicUtf8 | 16.25 ms | 0.897 ms | 0.700 ms |  1.00 |    0.06 |  18.02 KB |        1.00 |


## qualified-830951-long-v2-domain0-csv

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T17:02:32.9950324+00:00; finished 2026-10-08T17:25:09.3771494+00:00. Exact context: contexts/qualified-830951-long-v2-domain0-csv-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.CsvMaterializedReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KJRLRQ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                  | RowCount | Mean     | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1   | Allocated | Alloc Ratio |
|------------------------ |--------- |---------:|----------:|----------:|------:|--------:|---------:|-------:|----------:|------------:|
| ExcelReaderMaterialized | 50000    | 2.932 ms | 0.2182 ms | 0.2041 ms |  1.00 |    0.09 |  32.0285 |      - |   1.57 MB |        1.00 |
| OfficeIMOMaterialized   | 50000    | 6.412 ms | 0.8281 ms | 0.7746 ms |  2.20 |    0.29 | 183.4320 |      - |   8.63 MB |        5.48 |
| SylvanMaterialized      | 50000    | 3.154 ms | 0.3363 ms | 0.3146 ms |  1.08 |    0.13 |  36.2319 | 3.6232 |   1.61 MB |        1.02 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.CsvRawAsyncBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KJRLRQ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method           | RowCount | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1    | Allocated | Alloc Ratio |
|----------------- |--------- |----------:|----------:|----------:|------:|--------:|---------:|--------:|----------:|------------:|
| ExcelReaderAsync | 50000    |  3.179 ms | 0.5961 ms | 0.5576 ms |  1.03 |    0.26 |  32.1839 |       - |   1.57 MB |        1.00 |
| OfficeIMOAsync   | 50000    | 12.624 ms | 2.0272 ms | 1.8963 ms |  4.10 |    0.97 | 313.2530 | 24.0964 |  14.18 MB |        9.01 |
| SylvanAsync      | 50000    |  4.168 ms | 0.7507 ms | 0.7022 ms |  1.35 |    0.34 |  36.4964 |  3.6496 |   1.62 MB |        1.03 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.CsvRawReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KJRLRQ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                | RowCount | Mean     | Error     | StdDev    | Ratio | RatioSD | Allocated | Alloc Ratio |
|---------------------- |--------- |---------:|----------:|----------:|------:|--------:|----------:|------------:|
| ExcelReaderOriginal   | 50000    | 2.129 ms | 0.3389 ms | 0.3170 ms |  1.02 |    0.20 |     997 B |        1.00 |
| OfficeIMOBorrowedUtf8 | 50000    | 6.641 ms | 0.3740 ms | 0.3499 ms |  3.18 |    0.44 |    3849 B |        3.86 |
| SylvanOriginal        | 50000    | 4.115 ms | 0.3752 ms | 0.3510 ms |  1.97 |    0.30 |   38681 B |       38.80 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.CsvRealDataMaterializedReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KJRLRQ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                  | Mean     | Error    | StdDev   | Ratio | RatioSD | Gen0     | Gen1    | Allocated | Alloc Ratio |
|------------------------ |---------:|---------:|---------:|------:|--------:|---------:|--------:|----------:|------------:|
| ExcelReaderMaterialized | 30.62 ms | 1.703 ms | 1.593 ms |  1.00 |    0.07 | 735.2941 |       - |  35.71 MB |        1.00 |
| OfficeIMOMaterialized   | 20.18 ms | 2.222 ms | 2.078 ms |  0.66 |    0.07 | 720.0000 |       - |  34.21 MB |        0.96 |
| SylvanMaterialized      | 16.16 ms | 0.440 ms | 0.412 ms |  0.53 |    0.03 | 806.4516 | 96.7742 |  35.75 MB |        1.00 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.CsvRealDataReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KJRLRQ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                      | Mean      | Error     | StdDev    | Ratio | RatioSD | Allocated | Alloc Ratio |
|---------------------------- |----------:|----------:|----------:|------:|--------:|----------:|------------:|
| ExcelReaderStream           | 11.049 ms | 0.8898 ms | 0.8323 ms |  1.01 |    0.11 |    1712 B |        1.00 |
| ExcelReaderMemory           | 11.309 ms | 0.8946 ms | 0.8369 ms |  1.03 |    0.11 |     738 B |        0.43 |
| OfficeIMOStreamBorrowedUtf8 |  6.415 ms | 0.4758 ms | 0.4451 ms |  0.58 |    0.06 |    2920 B |        1.71 |
| SylvanFieldSpan             | 10.612 ms | 0.2819 ms | 0.2637 ms |  0.97 |    0.08 |   41226 B |       24.08 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.CsvRecordWriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KJRLRQ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                  | Mapped | RowCount | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated | Alloc Ratio |
|------------------------ |------- |--------- |----------:|----------:|----------:|------:|--------:|---------:|---------:|---------:|----------:|------------:|
| **ExcelReaderRecordLayout** | **False**  | **50000**    |  **5.098 ms** | **0.2976 ms** | **0.2784 ms** |  **1.00** |    **0.08** | **329.7872** | **329.7872** | **329.7872** |      **4 MB** |        **1.00** |
| OfficeIMOWriteObjects   | False  | 50000    | 12.567 ms | 1.2511 ms | 1.1703 ms |  2.47 |    0.26 | 484.5361 | 340.2062 | 340.2062 |  11.26 MB |        2.81 |
|                         |        |          |           |           |           |       |         |          |          |          |           |             |
| **ExcelReaderRecordLayout** | **True**   | **50000**    |  **5.411 ms** | **0.3128 ms** | **0.2926 ms** |  **1.00** |    **0.07** | **331.3953** | **331.3953** | **331.3953** |      **4 MB** |        **1.00** |
| OfficeIMOWriteObjects   | True   | 50000    | 12.820 ms | 0.7471 ms | 0.6988 ms |  2.38 |    0.18 | 486.4865 | 337.8378 | 337.8378 |  11.26 MB |        2.81 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.CsvTypedAsyncBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KJRLRQ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                | RowCount | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1    | Allocated | Alloc Ratio |
|---------------------- |--------- |----------:|----------:|----------:|------:|--------:|---------:|--------:|----------:|------------:|
| ExcelReaderTypedAsync | 50000    |  5.523 ms | 0.1607 ms | 0.1503 ms |  1.00 |    0.04 |  77.7385 |       - |   3.86 MB |        1.00 |
| OfficeIMORowsAsAsync  | 50000    | 13.342 ms | 0.6694 ms | 0.6261 ms |  2.42 |    0.13 | 333.3333 | 17.5439 |  16.47 MB |        4.26 |
| SylvanTypedAsync      | 50000    | 10.099 ms | 2.4265 ms | 2.2697 ms |  1.83 |    0.40 | 223.1405 | 24.7934 |  10.96 MB |        2.84 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.CsvTypedReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KJRLRQ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method           | RowCount | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1    | Allocated | Alloc Ratio |
|----------------- |--------- |----------:|----------:|----------:|------:|--------:|---------:|--------:|----------:|------------:|
| ExcelReaderTyped | 50000    |  3.800 ms | 0.8532 ms | 0.7981 ms |  1.03 |    0.27 |  80.3571 |       - |   3.86 MB |        1.00 |
| OfficeIMORowsAs  | 50000    |  8.014 ms | 0.6568 ms | 0.6143 ms |  2.18 |    0.39 | 226.8908 |       - |  10.92 MB |        2.83 |
| SylvanTyped      | 50000    | 11.577 ms | 0.9166 ms | 0.8573 ms |  3.15 |    0.56 | 252.7473 | 32.9670 |  10.95 MB |        2.83 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.CsvUtf8WriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KJRLRQ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                 | Mean     | Error     | StdDev    | Ratio | RatioSD | Allocated | Alloc Ratio |
|----------------------- |---------:|----------:|----------:|------:|--------:|----------:|------------:|
| OfficeIMOPublicUtf8Row | 8.442 ms | 1.3482 ms | 1.2611 ms |  2.52 |    0.48 |    6134 B |       14.75 |
| ExcelReaderPublicUtf8  | 3.407 ms | 0.5062 ms | 0.4735 ms |  1.02 |    0.19 |     416 B |        1.00 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.CsvWideReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KJRLRQ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method          | RowCount | MaterializeStrings | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0      | Gen1     | Allocated   | Alloc Ratio |
|---------------- |--------- |------------------- |----------:|----------:|----------:|------:|--------:|----------:|---------:|------------:|------------:|
| **ExcelReaderWide** | **50000**    | **False**              |  **6.071 ms** | **1.6035 ms** | **1.4999 ms** |  **1.05** |    **0.33** |         **-** |        **-** |     **1.21 KB** |        **1.00** |
| OfficeIMOWide   | 50000    | False              |  6.348 ms | 0.9906 ms | 0.9266 ms |  1.10 |    0.27 |         - |        - |     4.26 KB |        3.51 |
| SylvanWide      | 50000    | False              |  4.364 ms | 0.4441 ms | 0.4154 ms |  0.75 |    0.16 |         - |        - |    45.29 KB |       37.31 |
|                 |          |                    |           |           |           |       |         |           |          |             |             |
| **ExcelReaderWide** | **50000**    | **True**               | **25.899 ms** | **4.0062 ms** | **3.7474 ms** |  **1.02** |    **0.21** | **1039.2157** |        **-** | **51564.85 KB** |        **1.00** |
| OfficeIMOWide   | 50000    | True               | 12.973 ms | 1.2585 ms | 1.1772 ms |  0.51 |    0.09 | 1048.7805 |  12.1951 | 51569.89 KB |        1.00 |
| SylvanWide      | 50000    | True               | 15.417 ms | 1.3172 ms | 1.2321 ms |  0.61 |    0.10 | 1042.5532 | 159.5745 | 51607.99 KB |        1.00 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.CsvWriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KJRLRQ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                | RowCount | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated | Alloc Ratio |
|---------------------- |--------- |----------:|----------:|----------:|------:|--------:|---------:|---------:|---------:|----------:|------------:|
| ExcelReaderWriter     | 50000    |  4.057 ms | 0.3368 ms | 0.3150 ms |  1.01 |    0.11 | 498.1132 | 498.1132 | 498.1132 |      4 MB |        1.00 |
| OfficeIMOWriteObjects | 50000    | 14.167 ms | 0.9393 ms | 0.8786 ms |  3.51 |    0.36 | 693.3333 | 546.6667 | 546.6667 |  11.26 MB |        2.81 |
| SylvanWriter          | 50000    |  6.674 ms | 0.8882 ms | 0.8308 ms |  1.66 |    0.24 | 500.0000 | 500.0000 | 500.0000 |   4.04 MB |        1.01 |


## qualified-830951-long-v2-domain0-csv-large-dop1

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T17:25:09.8902429+00:00; finished 2026-10-08T17:58:32.6196702+00:00. Exact context: contexts/qualified-830951-long-v2-domain0-csv-large-dop1-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.CsvDirectAggregateBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-THMKXO : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=8  IterationTime=250ms  LaunchCount=1  
UnrollFactor=1  WarmupCount=4  

```
| Method                        | Input                | Dop | Mean       | Error     | StdDev    | Ratio | RatioSD | Gen0       | Gen1       | Gen2      | Allocated     | Alloc Ratio |
|------------------------------ |--------------------- |---- |-----------:|----------:|----------:|------:|--------:|-----------:|-----------:|----------:|--------------:|------------:|
| **ExcelReaderAggregate**          | **Conve(...)00000 [23]** | **1**   |   **696.8 ms** | **267.40 ms** | **139.86 ms** |  **1.04** |    **0.28** |          **-** |          **-** |         **-** |    **1228.73 KB** |        **1.00** |
| OfficeIMOTextOwnedAggregate   | Conve(...)00000 [23] | 1   |   724.7 ms |  55.83 ms |  29.20 ms |  1.08 |    0.21 | 10500.0000 | 10250.0000 | 2250.0000 |   812778.7 KB |      661.48 |
| OfficeIMOAsyncPathAggregate   | Conve(...)00000 [23] | 1   | 2,600.0 ms | 333.90 ms | 174.64 ms |  3.87 |    0.78 | 38750.0000 |   250.0000 |         - | 1903086.21 KB |    1,548.82 |
| OfficeIMOAsyncStreamAggregate | Conve(...)00000 [23] | 1   | 2,196.9 ms | 656.28 ms | 343.25 ms |  3.27 |    0.80 | 39000.0000 |   250.0000 |         - | 1903092.23 KB |    1,548.82 |
|                               |                      |     |            |           |           |       |         |            |            |           |               |             |
| **ExcelReaderAggregate**          | **Conve(...)00000 [25]** | **1**   |   **582.8 ms** | **115.14 ms** |  **60.22 ms** |  **1.01** |    **0.14** |          **-** |          **-** |         **-** |     **929.15 KB** |        **1.00** |
| OfficeIMOTextOwnedAggregate   | Conve(...)00000 [25] | 1   |   499.0 ms |  50.40 ms |  26.36 ms |  0.86 |    0.09 | 16750.0000 | 16500.0000 | 3000.0000 |  567413.63 KB |      610.68 |
| OfficeIMOAsyncPathAggregate   | Conve(...)00000 [25] | 1   | 1,843.2 ms | 526.00 ms | 275.11 ms |  3.19 |    0.54 | 27250.0000 |   250.0000 |         - | 1327730.37 KB |    1,428.97 |
| OfficeIMOAsyncStreamAggregate | Conve(...)00000 [25] | 1   | 2,049.4 ms | 704.87 ms | 368.66 ms |  3.55 |    0.69 | 27000.0000 |   250.0000 |         - | 1327745.75 KB |    1,428.99 |
|                               |                      |     |            |           |           |       |         |            |            |           |               |             |
| **ExcelReaderAggregate**          | **NarrowInt-8000000**    | **1**   |   **586.7 ms** |  **38.53 ms** |  **20.15 ms** |  **1.00** |    **0.04** |          **-** |          **-** |         **-** |    **1281.29 KB** |        **1.00** |
| OfficeIMOTextOwnedAggregate   | NarrowInt-8000000    | 1   |   894.4 ms |  87.03 ms |  45.52 ms |  1.53 |    0.09 | 23000.0000 | 22750.0000 | 3000.0000 |  821055.97 KB |      640.80 |
| OfficeIMOAsyncPathAggregate   | NarrowInt-8000000    | 1   | 3,182.5 ms | 360.70 ms | 188.65 ms |  5.43 |    0.35 | 45250.0000 |   250.0000 |         - | 2207873.59 KB |    1,723.17 |
| OfficeIMOAsyncStreamAggregate | NarrowInt-8000000    | 1   | 3,291.0 ms | 270.04 ms | 141.23 ms |  5.62 |    0.29 | 45000.0000 |   250.0000 |         - | 2207872.72 KB |    1,723.17 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.CsvParallelTypedBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-THMKXO : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=8  IterationTime=250ms  LaunchCount=1  
UnrollFactor=1  WarmupCount=4  

```
| Method                | Input                | Dop | Mean       | Error     | StdDev    | Ratio | RatioSD | Gen0       | Gen1       | Gen2      | Allocated  | Alloc Ratio |
|---------------------- |--------------------- |---- |-----------:|----------:|----------:|------:|--------:|-----------:|-----------:|----------:|-----------:|------------:|
| **ExcelReaderParallel**   | **Conve(...)00000 [23]** | **1**   |   **977.8 ms** | **119.55 ms** |  **62.53 ms** |  **1.00** |    **0.09** | **14250.0000** |   **250.0000** |         **-** |  **671.46 MB** |        **1.00** |
| ExcelReaderSequential | Conve(...)00000 [23] | 1   |   927.1 ms | 111.48 ms |  58.30 ms |  0.95 |    0.08 | 14000.0000 |          - |         - |  669.82 MB |        1.00 |
| OfficeIMOCoreParallel | Conve(...)00000 [23] | 1   |   762.8 ms |  28.02 ms |  14.65 ms |  0.78 |    0.05 | 27000.0000 |          - |         - | 1290.27 MB |        1.92 |
| OfficeIMOTextParallel | Conve(...)00000 [23] | 1   |   936.5 ms |  88.72 ms |  46.40 ms |  0.96 |    0.07 | 35750.0000 | 20750.0000 | 2750.0000 | 1463.86 MB |        2.18 |
| SylvanSequential      | Conve(...)00000 [23] | 1   | 1,222.9 ms | 211.76 ms | 110.75 ms |  1.26 |    0.13 | 27000.0000 |          - |         - | 1292.01 MB |        1.92 |
|                       |                      |     |            |           |           |       |         |            |            |           |            |             |
| **ExcelReaderParallel**   | **Conve(...)00000 [25]** | **1**   |   **700.6 ms** |  **65.45 ms** |  **34.23 ms** |  **1.00** |    **0.07** |  **9750.0000** |          **-** |         **-** |  **468.42 MB** |        **1.00** |
| ExcelReaderSequential | Conve(...)00000 [25] | 1   |   702.5 ms |  34.88 ms |  18.24 ms |  1.00 |    0.05 |  9750.0000 |          - |         - |   467.3 MB |        1.00 |
| OfficeIMOCoreParallel | Conve(...)00000 [25] | 1   |   531.6 ms |  73.31 ms |  38.34 ms |  0.76 |    0.06 | 18750.0000 |          - |         - |  900.14 MB |        1.92 |
| OfficeIMOTextParallel | Conve(...)00000 [25] | 1   |   672.6 ms |  39.62 ms |  20.72 ms |  0.96 |    0.05 | 19250.0000 |  9250.0000 | 2500.0000 | 1021.41 MB |        2.18 |
| SylvanSequential      | Conve(...)00000 [25] | 1   |   987.3 ms | 107.26 ms |  56.10 ms |  1.41 |    0.10 | 18750.0000 |          - |         - |  901.42 MB |        1.92 |
|                       |                      |     |            |           |           |       |         |            |            |           |            |             |
| **ExcelReaderParallel**   | **NarrowInt-8000000**    | **1**   |   **577.0 ms** |  **21.79 ms** |  **11.40 ms** |  **1.00** |    **0.03** |  **5000.0000** |          **-** |         **-** |  **245.75 MB** |        **1.00** |
| ExcelReaderSequential | NarrowInt-8000000    | 1   |   547.3 ms |  39.01 ms |  20.40 ms |  0.95 |    0.04 |  5000.0000 |          - |         - |  244.14 MB |        0.99 |
| OfficeIMOCoreParallel | NarrowInt-8000000    | 1   |   728.1 ms |  52.86 ms |  27.65 ms |  1.26 |    0.05 | 24000.0000 |          - |         - | 1158.54 MB |        4.71 |
| OfficeIMOTextParallel | NarrowInt-8000000    | 1   |   818.6 ms | 100.46 ms |  52.54 ms |  1.42 |    0.09 | 16000.0000 | 11500.0000 | 2500.0000 | 1045.73 MB |        4.26 |
| SylvanSequential      | NarrowInt-8000000    | 1   |   694.1 ms |  63.61 ms |  33.27 ms |  1.20 |    0.06 | 24000.0000 |          - |         - | 1158.58 MB |        4.71 |


## qualified-830951-long-v2-domain0-csv-large-dop4

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T17:58:33.0186425+00:00; finished 2026-10-08T18:22:10.5703073+00:00. Exact context: contexts/qualified-830951-long-v2-domain0-csv-large-dop4-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.CsvDirectAggregateBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-THMKXO : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=8  IterationTime=250ms  LaunchCount=1  
UnrollFactor=1  WarmupCount=4  

```
| Method                        | Input                | Dop | Mean       | Error     | StdDev    | Ratio | RatioSD | Gen0       | Gen1       | Gen2      | Allocated  | Alloc Ratio |
|------------------------------ |--------------------- |---- |-----------:|----------:|----------:|------:|--------:|-----------:|-----------:|----------:|-----------:|------------:|
| **ExcelReaderAggregate**          | **Conve(...)00000 [23]** | **4**   |   **214.1 ms** |   **8.95 ms** |   **4.68 ms** |  **1.00** |    **0.03** |          **-** |          **-** |         **-** |    **1.45 MB** |        **1.00** |
| OfficeIMOTextOwnedAggregate   | Conve(...)00000 [23] | 4   |   622.8 ms |  88.45 ms |  46.26 ms |  2.91 |    0.21 | 11000.0000 | 10750.0000 | 2750.0000 |   794.3 MB |      547.44 |
| OfficeIMOAsyncPathAggregate   | Conve(...)00000 [23] | 4   | 2,501.2 ms | 632.17 ms | 330.64 ms | 11.69 |    1.48 | 38750.0000 |  4500.0000 |         - | 1864.89 MB |    1,285.30 |
| OfficeIMOAsyncStreamAggregate | Conve(...)00000 [23] | 4   | 2,662.9 ms | 500.20 ms | 261.62 ms | 12.44 |    1.18 | 38750.0000 |  4500.0000 |         - | 1864.88 MB |    1,285.30 |
|                               |                      |     |            |           |           |       |         |            |            |           |            |             |
| **ExcelReaderAggregate**          | **Conve(...)00000 [25]** | **4**   |   **137.9 ms** |  **18.89 ms** |   **9.88 ms** |  **1.00** |    **0.09** |          **-** |          **-** |         **-** |    **1.08 MB** |        **1.00** |
| OfficeIMOTextOwnedAggregate   | Conve(...)00000 [25] | 4   |   351.1 ms |  29.58 ms |  15.47 ms |  2.56 |    0.19 | 16750.0000 | 16500.0000 | 3000.0000 |  555.06 MB |      513.64 |
| OfficeIMOAsyncPathAggregate   | Conve(...)00000 [25] | 4   | 1,606.2 ms | 685.03 ms | 358.28 ms | 11.70 |    2.58 | 27000.0000 |  3250.0000 |         - |  1301.1 MB |    1,204.01 |
| OfficeIMOAsyncStreamAggregate | Conve(...)00000 [25] | 4   | 1,626.7 ms | 242.63 ms | 126.90 ms | 11.85 |    1.15 | 27000.0000 |  3000.0000 |         - |  1301.1 MB |    1,204.01 |
|                               |                      |     |            |           |           |       |         |            |            |           |            |             |
| **ExcelReaderAggregate**          | **NarrowInt-8000000**    | **4**   |   **117.0 ms** |  **15.76 ms** |   **8.24 ms** |  **1.00** |    **0.09** |          **-** |          **-** |         **-** |    **1.49 MB** |        **1.00** |
| OfficeIMOTextOwnedAggregate   | NarrowInt-8000000    | 4   |   608.6 ms |  47.69 ms |  24.94 ms |  5.22 |    0.40 | 11250.0000 | 11000.0000 | 3000.0000 |  802.18 MB |      539.97 |
| OfficeIMOAsyncPathAggregate   | NarrowInt-8000000    | 4   | 2,521.2 ms | 646.08 ms | 337.91 ms | 21.64 |    3.10 | 45250.0000 |  3000.0000 |         - | 2168.05 MB |    1,459.38 |
| OfficeIMOAsyncStreamAggregate | NarrowInt-8000000    | 4   | 2,634.0 ms | 432.38 ms | 226.14 ms | 22.61 |    2.38 | 45500.0000 |  3000.0000 |         - | 2168.06 MB |    1,459.39 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.CsvParallelTypedBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-THMKXO : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=8  IterationTime=250ms  LaunchCount=1  
UnrollFactor=1  WarmupCount=4  

```
| Method                | Input                | Dop | Mean       | Error     | StdDev    | Ratio | RatioSD | Gen0       | Gen1       | Gen2      | Allocated  | Alloc Ratio |
|---------------------- |--------------------- |---- |-----------:|----------:|----------:|------:|--------:|-----------:|-----------:|----------:|-----------:|------------:|
| **ExcelReaderParallel**   | **Conve(...)00000 [23]** | **4**   |   **325.8 ms** |  **34.82 ms** |  **18.21 ms** |  **1.00** |    **0.08** | **14250.0000** |   **250.0000** |         **-** |   **680.2 MB** |        **1.00** |
| OfficeIMOCoreParallel | Conve(...)00000 [23] | 4   | 1,256.5 ms | 226.38 ms | 118.40 ms |  3.87 |    0.40 | 38250.0000 | 10000.0000 |         - | 1831.32 MB |        2.69 |
| OfficeIMOTextParallel | Conve(...)00000 [23] | 4   |   562.5 ms | 126.07 ms |  65.93 ms |  1.73 |    0.21 | 57000.0000 | 56250.0000 | 3750.0000 | 1466.62 MB |        2.16 |
|                       |                      |     |            |           |           |       |         |            |            |           |            |             |
| **ExcelReaderParallel**   | **Conve(...)00000 [25]** | **4**   |   **240.2 ms** |  **28.78 ms** |  **15.05 ms** |  **1.00** |    **0.08** | **18750.0000** | **15000.0000** |  **250.0000** |  **474.73 MB** |        **1.00** |
| OfficeIMOCoreParallel | Conve(...)00000 [25] | 4   |   871.3 ms | 144.18 ms |  75.41 ms |  3.64 |    0.36 | 31750.0000 |  8250.0000 |         - | 1277.68 MB |        2.69 |
| OfficeIMOTextParallel | Conve(...)00000 [25] | 4   |   437.4 ms |  43.47 ms |  22.74 ms |  1.83 |    0.14 | 39500.0000 | 38500.0000 | 3000.0000 | 1022.45 MB |        2.15 |
|                       |                      |     |            |           |           |       |         |            |            |           |            |             |
| **ExcelReaderParallel**   | **NarrowInt-8000000**    | **4**   |   **223.2 ms** |   **8.10 ms** |   **4.23 ms** |  **1.00** |    **0.03** |  **5500.0000** |   **250.0000** |         **-** |  **254.81 MB** |        **1.00** |
| OfficeIMOCoreParallel | NarrowInt-8000000    | 4   | 1,951.7 ms | 218.11 ms | 114.07 ms |  8.75 |    0.51 | 41250.0000 |  3750.0000 |         - | 1860.04 MB |        7.30 |
| OfficeIMOTextParallel | NarrowInt-8000000    | 4   |   570.8 ms |  49.88 ms |  26.09 ms |  2.56 |    0.12 | 16000.0000 | 12750.0000 | 2500.0000 | 1046.84 MB |        4.11 |


## qualified-830951-long-v2-domain0-arrow

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T18:22:10.9507805+00:00; finished 2026-10-08T18:26:16.3132345+00:00. Exact context: contexts/qualified-830951-long-v2-domain0-arrow-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.ArrowConversionBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KJRLRQ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method      | RowCount | Scenario     | Mean      | Error    | StdDev    | Ratio | RatioSD | Gen0      | Gen1     | Gen2     | Allocated | Alloc Ratio |
|------------ |--------- |------------- |----------:|---------:|----------:|------:|--------:|----------:|---------:|---------:|----------:|------------:|
| **ExcelReader** | **100000**   | **CsvAllString** | **14.611 ms** | **1.205 ms** | **1.1270 ms** |  **1.01** |    **0.10** |  **647.0588** | **602.9412** | **602.9412** |  **16.27 MB** |        **1.00** |
| OfficeIMO   | 100000   | CsvAllString | 42.747 ms | 3.865 ms | 3.6149 ms |  2.94 |    0.32 | 1269.2308 | 923.0769 | 769.2308 |   49.4 MB |        3.04 |
|             |          |              |           |          |           |       |         |           |          |          |           |             |
| **ExcelReader** | **100000**   | **CsvTyped**     |  **9.691 ms** | **1.007 ms** | **0.9417 ms** |  **1.01** |    **0.14** |  **530.6122** | **520.4082** | **520.4082** |   **8.13 MB** |        **1.00** |
| OfficeIMO   | 100000   | CsvTyped     | 26.870 ms | 1.583 ms | 1.4808 ms |  2.80 |    0.32 |  575.0000 | 275.0000 | 225.0000 |  21.55 MB |        2.65 |
|             |          |              |           |          |           |       |         |           |          |          |           |             |
| **ExcelReader** | **100000**   | **XlsbTyped**    | **15.154 ms** | **1.032 ms** | **0.9650 ms** |  **1.00** |    **0.09** |  **530.3030** | **515.1515** | **515.1515** |   **8.14 MB** |        **1.00** |
| OfficeIMO   | 100000   | XlsbTyped    | 24.170 ms | 2.513 ms | 2.3510 ms |  1.60 |    0.18 |  157.8947 | 105.2632 | 105.2632 |   9.09 MB |        1.12 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.ArrowInferredConversionBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-KJRLRQ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                            | RowCount | Mean     | Error    | StdDev   | Gen0     | Gen1     | Gen2     | Allocated |
|---------------------------------- |--------- |---------:|---------:|---------:|---------:|---------:|---------:|----------:|
| ExcelReader_CsvTypedWithInference | 100000   | 10.95 ms | 0.590 ms | 0.552 ms | 623.5294 | 600.0000 | 588.2353 |  14.89 MB |


## qualified-830951-long-v2-domain0-encrypted

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T18:26:16.8091087+00:00; finished 2026-10-08T18:31:18.8910402+00:00. Exact context: contexts/qualified-830951-long-v2-domain0-encrypted-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.EncryptedVerifiedAsyncReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-BHAAQB : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                         | Input              | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Allocated   | Alloc Ratio |
|------------------------------- |------------------- |----------:|----------:|----------:|------:|--------:|---------:|------------:|------------:|
| **ExcelReaderVerifiedStreamAsync** | **OriginalSmallXlsx**  |  **23.18 ms** |  **1.976 ms** |  **1.543 ms** |  **1.00** |    **0.09** |        **-** |    **63.35 KB** |        **1.00** |
| OfficeIMOVerifiedStreamAsync   | OriginalSmallXlsx  |  22.74 ms |  0.980 ms |  0.765 ms |  0.98 |    0.07 |        - |   262.74 KB |        4.15 |
|                                |                    |           |           |           |       |         |          |             |             |
| **ExcelReaderVerifiedStreamAsync** | **GeneratedLargeXlsx** |  **72.80 ms** |  **6.595 ms** |  **5.149 ms** |  **1.00** |    **0.09** | **111.1111** | **10052.98 KB** |        **1.00** |
| OfficeIMOVerifiedStreamAsync   | GeneratedLargeXlsx | 157.32 ms | 15.324 ms | 11.964 ms |  2.17 |    0.21 | 500.0000 | 66664.58 KB |        6.63 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.EncryptedVerifiedReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-BHAAQB : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                        | Input              | MemoryInput | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Allocated   | Alloc Ratio |
|------------------------------ |------------------- |------------ |----------:|----------:|----------:|------:|--------:|---------:|------------:|------------:|
| **ExcelReaderVerifiedFullFields** | **OriginalSmallXlsx**  | **False**       |  **21.92 ms** |  **1.879 ms** |  **1.467 ms** |  **1.00** |    **0.09** |        **-** |    **40.59 KB** |        **1.00** |
| OfficeIMOVerifiedFullFields   | OriginalSmallXlsx  | False       |  24.06 ms |  1.214 ms |  0.948 ms |  1.10 |    0.08 |        - |   261.44 KB |        6.44 |
|                               |                    |             |           |           |           |       |         |          |             |             |
| **ExcelReaderVerifiedFullFields** | **OriginalSmallXlsx**  | **True**        |  **23.22 ms** |  **1.587 ms** |  **1.239 ms** |  **1.00** |    **0.07** |        **-** |    **53.57 KB** |        **1.00** |
| OfficeIMOVerifiedFullFields   | OriginalSmallXlsx  | True        |  23.87 ms |  2.197 ms |  1.715 ms |  1.03 |    0.09 |        - |   242.14 KB |        4.52 |
|                               |                    |             |           |           |           |       |         |          |             |             |
| **ExcelReaderVerifiedFullFields** | **GeneratedLargeXlsx** | **False**       |  **71.67 ms** |  **6.712 ms** |  **5.240 ms** |  **1.00** |    **0.10** | **100.0000** |  **9699.93 KB** |        **1.00** |
| OfficeIMOVerifiedFullFields   | GeneratedLargeXlsx | False       | 135.53 ms | 17.588 ms | 13.732 ms |  1.90 |    0.23 | 571.4286 | 66663.86 KB |        6.87 |
|                               |                    |             |           |           |           |       |         |          |             |             |
| **ExcelReaderVerifiedFullFields** | **GeneratedLargeXlsx** | **True**        |  **74.23 ms** |  **9.028 ms** |  **7.049 ms** |  **1.01** |    **0.13** | **100.0000** | **18741.63 KB** |        **1.00** |
| OfficeIMOVerifiedFullFields   | GeneratedLargeXlsx | True        | 139.85 ms | 18.592 ms | 14.515 ms |  1.90 |    0.26 | 571.4286 | 62103.77 KB |        3.31 |


## qualified-830951-long-v2-domain0-cold

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T18:31:19.2205343+00:00; finished 2026-10-08T18:31:53.1664909+00:00. Exact context: contexts/qualified-830951-long-v2-domain0-cold-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.ColdStartReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-CNKMYX : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
InvocationCount=1  IterationCount=1  LaunchCount=16  
RunStrategy=ColdStart  UnrollFactor=1  WarmupCount=0  

```
| Method                       | Mean     | Error     | StdDev    | Ratio | RatioSD | Allocated | Alloc Ratio |
|----------------------------- |---------:|----------:|----------:|------:|--------:|----------:|------------:|
| ExcelReaderAttributes        | 29.21 ms |  4.417 ms |  4.338 ms |  1.02 |    0.18 |  21.09 KB |        1.00 |
| OfficeIMOAutomaticMapping    | 78.43 ms | 50.564 ms | 49.660 ms |  2.73 |    1.71 |   93.8 KB |        4.45 |
| ExcelReaderFluentMapping     | 24.89 ms |  2.337 ms |  2.296 ms |  0.87 |    0.12 |  24.03 KB |        1.14 |
| ExcelReaderAttributeFallback | 30.04 ms |  1.857 ms |  1.823 ms |  1.04 |    0.13 |  26.16 KB |        1.24 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.ColdStartWriteBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-CNKMYX : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
InvocationCount=1  IterationCount=1  LaunchCount=16  
RunStrategy=ColdStart  UnrollFactor=1  WarmupCount=0  

```
| Method                     | Mean      | Error    | StdDev   | Ratio | RatioSD | Allocated | Alloc Ratio |
|--------------------------- |----------:|---------:|---------:|------:|--------:|----------:|------------:|
| ExcelReaderAutomaticLayout |  33.94 ms | 0.576 ms | 0.565 ms |  1.00 |    0.02 |  84.25 KB |        1.00 |
| OfficeIMOAutomaticLayout   | 125.11 ms | 2.126 ms | 2.088 ms |  3.69 |    0.08 | 664.86 KB |        7.89 |


## qualified-830951-long-v2-domain1-generated

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T18:31:53.4427565+00:00; finished 2026-10-08T18:58:48.3387884+00:00. Exact context: contexts/qualified-830951-long-v2-domain1-generated-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.BorrowedRefStructReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-CPEWJC : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                     | RowCount | Mean     | Error     | StdDev    | Ratio | RatioSD | Allocated | Alloc Ratio |
|--------------------------- |--------- |---------:|----------:|----------:|------:|--------:|----------:|------------:|
| ExcelReaderMappedRefStruct | 50000    | 7.661 ms | 0.4562 ms | 0.4268 ms |  1.00 |    0.08 |   4.36 KB |        1.00 |
| OfficeIMOFactoryRefStruct  | 50000    | 8.938 ms | 0.4358 ms | 0.4076 ms |  1.17 |    0.08 |   77.3 KB |       17.73 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.MaterializedRawAsyncReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-CPEWJC : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                              | RowCount | Format | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1   | Allocated | Alloc Ratio |
|------------------------------------ |--------- |------- |----------:|----------:|----------:|------:|--------:|---------:|-------:|----------:|------------:|
| **ExcelReaderStringsMaterializedAsync** | **50000**    | **Xlsx**   |  **6.840 ms** | **0.4101 ms** | **0.3836 ms** |  **1.00** |    **0.08** |  **26.4901** |      **-** |   **1.58 MB** |        **1.00** |
| OfficeIMOAsync                      | 50000    | Xlsx   | 11.800 ms | 1.2146 ms | 1.1361 ms |  1.73 |    0.19 |  71.4286 |      - |   3.51 MB |        2.22 |
| SylvanAsync                         | 50000    | Xlsx   | 28.504 ms | 2.0869 ms | 1.9521 ms |  4.18 |    0.36 |  35.7143 |      - |   1.89 MB |        1.20 |
|                                     |          |        |           |           |           |       |         |          |        |           |             |
| **ExcelReaderStringsMaterializedAsync** | **50000**    | **Xlsb**   |  **4.459 ms** | **0.0420 ms** | **0.0393 ms** |  **1.00** |    **0.01** |  **30.9735** |      **-** |   **1.58 MB** |        **1.00** |
| OfficeIMOAsync                      | 50000    | Xlsb   |  7.157 ms | 0.3374 ms | 0.3156 ms |  1.61 |    0.07 | 104.8951 |      - |   5.12 MB |        3.24 |
| SylvanAsync                         | 50000    | Xlsb   |  5.816 ms | 0.0758 ms | 0.0709 ms |  1.30 |    0.02 |  32.4675 |      - |   1.82 MB |        1.15 |
|                                     |          |        |           |           |           |       |         |          |        |           |             |
| **ExcelReaderStringsMaterializedAsync** | **50000**    | **Xls**    |  **2.820 ms** | **0.0170 ms** | **0.0159 ms** |  **1.00** |    **0.01** |  **30.8123** |      **-** |   **1.58 MB** |        **1.00** |
| OfficeIMOAsync                      | 50000    | Xls    |  5.454 ms | 0.1017 ms | 0.0951 ms |  1.93 |    0.03 | 208.7912 |      - |  10.25 MB |        6.50 |
| SylvanAsync                         | 50000    | Xls    |  4.813 ms | 0.1293 ms | 0.1209 ms |  1.71 |    0.04 |  34.3137 | 4.9020 |   1.68 MB |        1.06 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.MaterializedRawReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-CPEWJC : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                         | RowCount | Format | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1   | Allocated | Alloc Ratio |
|------------------------------- |--------- |------- |----------:|----------:|----------:|------:|--------:|---------:|-------:|----------:|------------:|
| **ExcelReaderStringsMaterialized** | **50000**    | **Xlsx**   |  **5.972 ms** | **0.0915 ms** | **0.0856 ms** |  **1.00** |    **0.02** |  **29.9401** |      **-** |   **1.58 MB** |        **1.00** |
| OfficeIMOOriginal              | 50000    | Xlsx   |  9.958 ms | 0.8808 ms | 0.8239 ms |  1.67 |    0.14 |  52.6316 |      - |   3.51 MB |        2.22 |
| Sylvan                         | 50000    | Xlsx   | 25.989 ms | 0.1915 ms | 0.1791 ms |  4.35 |    0.07 |  27.0270 |      - |   1.89 MB |        1.20 |
|                                |          |        |           |           |           |       |         |          |        |           |             |
| **ExcelReaderStringsMaterialized** | **50000**    | **Xlsb**   |  **4.372 ms** | **0.0260 ms** | **0.0243 ms** |  **1.00** |    **0.01** |  **30.5677** |      **-** |   **1.58 MB** |        **1.00** |
| OfficeIMOOriginal              | 50000    | Xlsb   |  6.555 ms | 0.1127 ms | 0.1054 ms |  1.50 |    0.02 | 100.0000 |      - |   5.12 MB |        3.25 |
| Sylvan                         | 50000    | Xlsb   |  5.449 ms | 0.1037 ms | 0.0970 ms |  1.25 |    0.02 |  38.0435 | 5.4348 |   1.82 MB |        1.15 |
|                                |          |        |           |           |           |       |         |          |        |           |             |
| **ExcelReaderStringsMaterialized** | **50000**    | **Xls**    |  **2.960 ms** | **0.1148 ms** | **0.1074 ms** |  **1.00** |    **0.05** |  **32.2581** |      **-** |   **1.58 MB** |        **1.00** |
| OfficeIMOOriginal              | 50000    | Xls    |  4.803 ms | 0.0589 ms | 0.0551 ms |  1.62 |    0.05 | 211.5385 |      - |  10.25 MB |        6.50 |
| Sylvan                         | 50000    | Xls    |  4.268 ms | 0.0298 ms | 0.0279 ms |  1.44 |    0.05 |  34.1880 | 4.2735 |   1.68 MB |        1.06 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.RawAsyncReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-CPEWJC : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method           | RowCount | Format | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1   | Allocated   | Alloc Ratio |
|----------------- |--------- |------- |----------:|----------:|----------:|------:|--------:|---------:|-------:|------------:|------------:|
| **ExcelReaderAsync** | **50000**    | **Xlsx**   |  **5.501 ms** | **0.1047 ms** | **0.0980 ms** |  **1.00** |    **0.02** |        **-** |      **-** |     **3.89 KB** |        **1.00** |
| OfficeIMOAsync   | 50000    | Xlsx   |  9.962 ms | 0.0794 ms | 0.0743 ms |  1.81 |    0.03 |  52.6316 |      - |  3592.14 KB |      923.28 |
| SylvanAsync      | 50000    | Xlsx   | 29.912 ms | 2.6887 ms | 2.5150 ms |  5.44 |    0.45 |  27.0270 |      - |  1940.06 KB |      498.65 |
|                  |          |        |           |           |           |       |         |          |        |             |             |
| **ExcelReaderAsync** | **50000**    | **Xlsb**   |  **4.745 ms** | **0.4350 ms** | **0.4069 ms** |  **1.01** |    **0.12** |        **-** |      **-** |      **4.2 KB** |        **1.00** |
| OfficeIMOAsync   | 50000    | Xlsb   |  6.939 ms | 0.0641 ms | 0.0600 ms |  1.47 |    0.12 | 104.8951 |      - |  5246.52 KB |    1,248.24 |
| SylvanAsync      | 50000    | Xlsb   |  6.201 ms | 0.2891 ms | 0.2705 ms |  1.32 |    0.12 |  22.2222 |      - |  1866.42 KB |      444.06 |
|                  |          |        |           |           |           |       |         |          |        |             |             |
| **ExcelReaderAsync** | **50000**    | **Xls**    |  **2.794 ms** | **0.3912 ms** | **0.3659 ms** |  **1.02** |    **0.18** |        **-** |      **-** |     **2.91 KB** |        **1.00** |
| OfficeIMOAsync   | 50000    | Xls    |  5.266 ms | 0.0754 ms | 0.0705 ms |  1.91 |    0.23 | 213.9037 |      - | 10499.22 KB |    3,612.64 |
| SylvanAsync      | 50000    | Xls    |  5.406 ms | 0.5196 ms | 0.4860 ms |  1.96 |    0.29 |  33.3333 | 5.5556 |  1719.04 KB |      591.50 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.RawReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-CPEWJC : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method              | RowCount | Format | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1   | Allocated   | Alloc Ratio |
|-------------------- |--------- |------- |----------:|----------:|----------:|------:|--------:|---------:|-------:|------------:|------------:|
| **ExcelReaderOriginal** | **50000**    | **Xlsx**   |  **5.322 ms** | **0.0483 ms** | **0.0452 ms** |  **1.00** |    **0.01** |        **-** |      **-** |     **3.82 KB** |        **1.00** |
| OfficeIMOOriginal   | 50000    | Xlsx   | 16.714 ms | 0.3323 ms | 0.3109 ms |  3.14 |    0.06 | 100.0000 |      - |  3719.13 KB |      973.51 |
| Sylvan              | 50000    | Xlsx   | 26.888 ms | 0.5443 ms | 0.5091 ms |  5.05 |    0.10 |  33.3333 |      - |  1939.16 KB |      507.59 |
|                     |          |        |           |           |           |       |         |          |        |             |             |
| **ExcelReaderOriginal** | **50000**    | **Xlsb**   |  **4.010 ms** | **0.1678 ms** | **0.1569 ms** |  **1.00** |    **0.05** |        **-** |      **-** |     **4.13 KB** |        **1.00** |
| OfficeIMOOriginal   | 50000    | Xlsb   |  7.778 ms | 0.7981 ms | 0.7466 ms |  1.94 |    0.19 | 104.0000 |      - |  5246.45 KB |    1,269.46 |
| Sylvan              | 50000    | Xlsb   |  5.571 ms | 0.1658 ms | 0.1551 ms |  1.39 |    0.06 |  43.7158 | 5.4645 |  1865.58 KB |      451.41 |
|                     |          |        |           |           |           |       |         |          |        |             |             |
| **ExcelReaderOriginal** | **50000**    | **Xls**    |  **2.533 ms** | **0.0194 ms** | **0.0181 ms** |  **1.00** |    **0.01** |        **-** |      **-** |     **3.11 KB** |        **1.00** |
| OfficeIMOOriginal   | 50000    | Xls    |  4.878 ms | 0.1437 ms | 0.1344 ms |  1.93 |    0.05 | 212.5604 |      - | 10520.18 KB |    3,382.31 |
| Sylvan              | 50000    | Xls    |  4.437 ms | 0.2843 ms | 0.2659 ms |  1.75 |    0.10 |  30.8370 | 4.4053 |  1717.76 KB |      552.27 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedAsyncReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-CPEWJC : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                | RowCount | Shape    | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Allocated | Alloc Ratio |
|---------------------- |--------- |--------- |----------:|----------:|----------:|------:|--------:|---------:|----------:|------------:|
| OfficeIMOTypedAsync   | 50000    | Original | 11.038 ms | 1.0608 ms | 0.9923 ms |  1.44 |    0.13 |        - |   2.36 MB |        0.61 |
| ExcelReaderTypedAsync | 50000    | Original |  7.680 ms | 0.1508 ms | 0.1411 ms |  1.00 |    0.02 |  75.7576 |   3.87 MB |        1.00 |
| SylvanTypedAsync      | 50000    | Original | 43.822 ms | 1.1125 ms | 1.0406 ms |  5.71 |    0.16 | 200.0000 |  10.48 MB |        2.71 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedModelReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-CPEWJC : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method               | RowCount | Model  | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0    | Allocated  | Alloc Ratio |
|--------------------- |--------- |------- |----------:|----------:|----------:|------:|--------:|--------:|-----------:|------------:|
| **ExcelReaderAutomatic** | **50000**    | **Class**  |  **7.528 ms** | **0.1741 ms** | **0.1628 ms** |  **1.00** |    **0.03** | **75.1880** | **3959.57 KB** |        **1.00** |
| OfficeIMOAutomatic   | 50000    | Class  |  9.978 ms | 0.3203 ms | 0.2996 ms |  1.33 |    0.05 | 40.4040 |  2421.6 KB |        0.61 |
|                      |          |        |           |           |           |       |         |         |            |             |
| **ExcelReaderAutomatic** | **50000**    | **Struct** |  **7.408 ms** | **0.1834 ms** | **0.1716 ms** |  **1.00** |    **0.03** | **29.1971** | **1615.82 KB** |        **1.00** |
| OfficeIMOAutomatic   | 50000    | Struct | 10.144 ms | 0.2793 ms | 0.2612 ms |  1.37 |    0.04 |       - |   77.75 KB |        0.05 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedStreamAsyncReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-CPEWJC : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                          | RowCount | Shape    | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Allocated | Alloc Ratio |
|-------------------------------- |--------- |--------- |----------:|----------:|----------:|------:|--------:|---------:|----------:|------------:|
| OfficeIMOStreamOpenTypedAsync   | 50000    | Original | 11.045 ms | 0.2415 ms | 0.2259 ms |  1.46 |    0.03 |        - |    3.1 MB |        0.80 |
| ExcelReaderStreamOpenTypedAsync | 50000    | Original |  7.544 ms | 0.0438 ms | 0.0409 ms |  1.00 |    0.01 |  75.1880 |   3.87 MB |        1.00 |
| SylvanStreamOpenTypedAsync      | 50000    | Original | 43.647 ms | 0.7136 ms | 0.6675 ms |  5.79 |    0.09 | 181.8182 |  10.48 MB |        2.71 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedXlsbAsyncReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-CPEWJC : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                    | RowCount | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Allocated | Alloc Ratio |
|-------------------------- |--------- |----------:|----------:|----------:|------:|--------:|---------:|----------:|------------:|
| ExcelReaderTypedXlsbAsync | 50000    |  6.162 ms | 0.2952 ms | 0.2762 ms |  1.00 |    0.06 |  76.9231 |   3.87 MB |        1.00 |
| OfficeIMOTypedXlsbAsync   | 50000    | 11.090 ms | 1.0249 ms | 0.9587 ms |  1.80 |    0.17 | 171.8750 |   8.56 MB |        2.21 |
| SylvanTypedXlsbAsync      | 50000    | 23.807 ms | 3.1492 ms | 2.9458 ms |  3.87 |    0.49 | 190.4762 |   10.4 MB |        2.69 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedXlsbReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-CPEWJC : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method               | RowCount | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1    | Allocated | Alloc Ratio |
|--------------------- |--------- |----------:|----------:|----------:|------:|--------:|---------:|--------:|----------:|------------:|
| ExcelReaderTypedXlsb | 50000    |  6.156 ms | 0.1886 ms | 0.1765 ms |  1.00 |    0.04 |  77.4194 |       - |   3.87 MB |        1.00 |
| OfficeIMOTypedXlsb   | 50000    |  6.954 ms | 0.1790 ms | 0.1674 ms |  1.13 |    0.04 |  77.4648 |       - |   3.98 MB |        1.03 |
| SylvanTypedXlsb      | 50000    | 19.545 ms | 1.4768 ms | 1.3814 ms |  3.18 |    0.23 | 208.3333 | 41.6667 |   10.4 MB |        2.69 |


## qualified-830951-long-v2-domain1-real

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T18:58:48.7101444+00:00; finished 2026-10-08T19:26:11.6425595+00:00. Exact context: contexts/qualified-830951-long-v2-domain1-real-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.MaterializedRealDataReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-CPEWJC : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                         | Format | Mean      | Error    | StdDev   | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated   | Alloc Ratio |
|------------------------------- |------- |----------:|---------:|---------:|------:|--------:|---------:|---------:|---------:|------------:|------------:|
| **ExcelReaderStringsMaterialized** | **Xlsx**   |  **42.63 ms** | **2.686 ms** | **2.513 ms** |  **1.00** |    **0.08** |        **-** |        **-** |        **-** |    **17.09 KB** |        **1.00** |
| OfficeIMOBytes                 | Xlsx   |  67.23 ms | 6.448 ms | 6.031 ms |  1.58 |    0.16 | 250.0000 |        - |        - | 13939.88 KB |      815.50 |
| OfficeIMOStream                | Xlsx   |  64.93 ms | 5.817 ms | 5.441 ms |  1.53 |    0.15 | 250.0000 |        - |        - | 20503.75 KB |    1,199.49 |
| Sylvan                         | Xlsx   | 155.75 ms | 7.883 ms | 7.374 ms |  3.66 |    0.27 |        - |        - |        - |   644.56 KB |       37.71 |
|                                |        |           |          |          |       |         |          |          |          |             |             |
| **ExcelReaderStringsMaterialized** | **Xlsm**   |  **41.23 ms** | **1.383 ms** | **1.293 ms** |  **1.00** |    **0.04** |        **-** |        **-** |        **-** |    **17.09 KB** |        **1.00** |
| OfficeIMOBytes                 | Xlsm   |  61.14 ms | 3.517 ms | 3.290 ms |  1.48 |    0.09 | 266.6667 |        - |        - |  13939.9 KB |      815.50 |
| OfficeIMOStream                | Xlsm   |  63.58 ms | 3.760 ms | 3.517 ms |  1.54 |    0.09 | 214.2857 |        - |        - | 20503.73 KB |    1,199.49 |
| Sylvan                         | Xlsm   | 151.19 ms | 4.897 ms | 4.580 ms |  3.67 |    0.15 |        - |        - |        - |   644.64 KB |       37.71 |
|                                |        |           |          |          |       |         |          |          |          |             |             |
| **ExcelReaderStringsMaterialized** | **Xlsb**   |  **25.15 ms** | **0.811 ms** | **0.758 ms** |  **1.00** |    **0.04** |        **-** |        **-** |        **-** |    **16.65 KB** |        **1.00** |
| OfficeIMOBytes                 | Xlsb   |  25.20 ms | 0.771 ms | 0.721 ms |  1.00 |    0.04 | 275.0000 |        - |        - | 13986.21 KB |      840.09 |
| OfficeIMOStream                | Xlsb   |  26.36 ms | 1.326 ms | 1.240 ms |  1.05 |    0.06 | 333.3333 |  51.2821 |  51.2821 | 17770.89 KB |    1,067.42 |
| Sylvan                         | Xlsb   |  26.15 ms | 1.098 ms | 1.027 ms |  1.04 |    0.05 |        - |        - |        - |   338.87 KB |       20.35 |
|                                |        |           |          |          |       |         |          |          |          |             |             |
| **ExcelReaderStringsMaterialized** | **Xls**    |  **10.74 ms** | **0.721 ms** | **0.675 ms** |  **1.00** |    **0.09** |        **-** |        **-** |        **-** |    **21.48 KB** |        **1.00** |
| OfficeIMOBytes                 | Xls    |  14.72 ms | 1.162 ms | 1.087 ms |  1.38 |    0.13 | 271.6049 |  12.3457 |        - | 13990.08 KB |      651.17 |
| OfficeIMOStream                | Xls    |  15.84 ms | 0.825 ms | 0.772 ms |  1.48 |    0.11 | 396.8254 | 126.9841 | 126.9841 | 25538.02 KB |    1,188.68 |
| Sylvan                         | Xls    |  18.23 ms | 0.725 ms | 0.678 ms |  1.70 |    0.12 |        - |        - |        - |   186.29 KB |        8.67 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.PrefetchedRealDataReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-CPEWJC : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                    | Format | Mean      | Error    | StdDev   | Ratio | RatioSD | Gen0     | Allocated   | Alloc Ratio |
|-------------------------- |------- |----------:|---------:|---------:|------:|--------:|---------:|------------:|------------:|
| **ExcelReaderPrefetch**       | **Xlsx**   |  **28.12 ms** | **0.895 ms** | **0.837 ms** |  **1.00** |    **0.04** |        **-** |    **22.64 KB** |        **1.00** |
| ExcelReaderMemoryPrefetch | Xlsx   |  28.13 ms | 1.038 ms | 0.971 ms |  1.00 |    0.04 |        - |    24.53 KB |        1.08 |
| OfficeIMOBytes            | Xlsx   |  60.82 ms | 1.947 ms | 1.821 ms |  2.16 |    0.09 | 266.6667 | 13939.94 KB |      615.73 |
| Sylvan                    | Xlsx   | 156.40 ms | 4.967 ms | 4.646 ms |  5.57 |    0.23 |        - |   644.56 KB |       28.47 |
|                           |        |           |          |          |       |         |          |             |             |
| **ExcelReaderPrefetch**       | **Xlsm**   |  **28.03 ms** | **0.857 ms** | **0.801 ms** |  **1.00** |    **0.04** |        **-** |    **24.76 KB** |        **1.00** |
| ExcelReaderMemoryPrefetch | Xlsm   |  29.15 ms | 1.028 ms | 0.962 ms |  1.04 |    0.04 |        - |     23.2 KB |        0.94 |
| OfficeIMOBytes            | Xlsm   |  63.07 ms | 3.253 ms | 3.043 ms |  2.25 |    0.12 | 214.2857 | 13939.97 KB |      563.01 |
| Sylvan                    | Xlsm   | 149.78 ms | 8.987 ms | 8.407 ms |  5.35 |    0.33 |        - |    644.6 KB |       26.03 |
|                           |        |           |          |          |       |         |          |             |             |
| **ExcelReaderPrefetch**       | **Xlsb**   |  **15.48 ms** | **0.296 ms** | **0.277 ms** |  **1.00** |    **0.02** |        **-** |    **16.23 KB** |        **1.00** |
| ExcelReaderMemoryPrefetch | Xlsb   |  15.60 ms | 0.259 ms | 0.242 ms |  1.01 |    0.02 |        - |    19.23 KB |        1.19 |
| OfficeIMOBytes            | Xlsb   |  24.76 ms | 1.246 ms | 1.166 ms |  1.60 |    0.08 | 282.0513 | 13985.91 KB |      861.97 |
| Sylvan                    | Xlsb   |  26.10 ms | 1.066 ms | 0.997 ms |  1.69 |    0.07 |        - |   338.86 KB |       20.88 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.RealDataReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-CPEWJC : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method              | Format | Mean       | Error      | StdDev    | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated   | Alloc Ratio |
|-------------------- |------- |-----------:|-----------:|----------:|------:|--------:|---------:|---------:|---------:|------------:|------------:|
| **ExcelReaderOriginal** | **Xlsx**   |  **39.014 ms** |  **1.0641 ms** | **0.9953 ms** |  **1.00** |    **0.03** |        **-** |        **-** |        **-** |     **7.63 KB** |        **1.00** |
| ExcelReaderMemory   | Xlsx   |  39.695 ms |  1.0834 ms | 1.0134 ms |  1.02 |    0.03 |        - |        - |        - |     7.38 KB |        0.97 |
| OfficeIMOBytes      | Xlsx   |  63.332 ms |  3.0994 ms | 2.8992 ms |  1.62 |    0.08 | 250.0000 |        - |        - | 13939.94 KB |    1,828.19 |
| OfficeIMOStream     | Xlsx   |  60.637 ms |  7.5709 ms | 7.0818 ms |  1.56 |    0.18 | 250.0000 |        - |        - | 20503.75 KB |    2,689.02 |
| Sylvan              | Xlsx   | 154.145 ms | 10.3071 ms | 9.6413 ms |  3.95 |    0.26 |        - |        - |        - |   644.56 KB |       84.53 |
|                     |        |            |            |           |       |         |          |          |          |             |             |
| **ExcelReaderOriginal** | **Xlsm**   |  **40.535 ms** |  **1.7824 ms** | **1.6672 ms** |  **1.00** |    **0.06** |        **-** |        **-** |        **-** |     **7.63 KB** |        **1.00** |
| ExcelReaderMemory   | Xlsm   |  40.127 ms |  1.3181 ms | 1.2329 ms |  0.99 |    0.05 |        - |        - |        - |     7.38 KB |        0.97 |
| OfficeIMOBytes      | Xlsm   |  62.898 ms |  4.6439 ms | 4.3439 ms |  1.55 |    0.12 | 250.0000 |        - |        - | 13939.89 KB |    1,828.18 |
| OfficeIMOStream     | Xlsm   |  61.788 ms |  4.1383 ms | 3.8710 ms |  1.53 |    0.11 | 250.0000 |        - |        - | 20503.89 KB |    2,689.03 |
| Sylvan              | Xlsm   | 155.175 ms |  9.6963 ms | 9.0699 ms |  3.83 |    0.26 |        - |        - |        - |    644.6 KB |       84.54 |
|                     |        |            |            |           |       |         |          |          |          |             |             |
| **ExcelReaderOriginal** | **Xlsb**   |  **24.147 ms** |  **0.6893 ms** | **0.6448 ms** |  **1.00** |    **0.04** |        **-** |        **-** |        **-** |     **7.18 KB** |        **1.00** |
| ExcelReaderMemory   | Xlsb   |  24.437 ms |  0.7554 ms | 0.7066 ms |  1.01 |    0.04 |        - |        - |        - |     8.88 KB |        1.24 |
| OfficeIMOBytes      | Xlsb   |  26.162 ms |  1.5234 ms | 1.4250 ms |  1.08 |    0.06 | 282.0513 |        - |        - | 13985.91 KB |    1,947.98 |
| OfficeIMOStream     | Xlsb   |  26.537 ms |  1.2263 ms | 1.1471 ms |  1.10 |    0.05 | 388.8889 | 111.1111 | 111.1111 | 17771.28 KB |    2,475.22 |
| Sylvan              | Xlsb   |  30.174 ms |  1.8938 ms | 1.7715 ms |  1.25 |    0.08 |        - |        - |        - |   338.87 KB |       47.20 |
|                     |        |            |            |           |       |         |          |          |          |             |             |
| **ExcelReaderOriginal** | **Xls**    |   **9.089 ms** |  **0.3941 ms** | **0.3686 ms** |  **1.00** |    **0.06** |        **-** |        **-** |        **-** |    **12.02 KB** |        **1.00** |
| ExcelReaderMemory   | Xls    |   9.664 ms |  0.7303 ms | 0.6831 ms |  1.06 |    0.08 |        - |        - |        - |    11.88 KB |        0.99 |
| OfficeIMOBytes      | Xls    |  13.401 ms |  0.6530 ms | 0.6109 ms |  1.48 |    0.09 | 275.3623 |  14.4928 |        - | 13990.09 KB |    1,164.32 |
| OfficeIMOStream     | Xls    |  15.065 ms |  1.6750 ms | 1.5668 ms |  1.66 |    0.18 | 416.6667 | 133.3333 | 133.3333 | 25538.06 KB |    2,125.40 |
| Sylvan              | Xls    |  17.305 ms |  1.3064 ms | 1.2220 ms |  1.91 |    0.15 |        - |        - |        - |   186.29 KB |       15.50 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.RealDataTypedReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-CPEWJC : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method               | Format | Mean     | Error    | StdDev   | Ratio | RatioSD | Gen0     | Gen1    | Allocated | Alloc Ratio |
|--------------------- |------- |---------:|---------:|---------:|------:|--------:|---------:|--------:|----------:|------------:|
| **ExcelReaderAutomatic** | **Xlsx**   | **47.45 ms** | **2.563 ms** | **2.397 ms** |  **1.00** |    **0.07** | **142.8571** |       **-** |   **8.02 MB** |        **1.00** |
| OfficeIMOAutomatic   | Xlsx   | 61.43 ms | 3.842 ms | 3.594 ms |  1.30 |    0.09 |        - |       - |   8.12 MB |        1.01 |
|                      |        |          |          |          |       |         |          |         |           |             |
| **ExcelReaderAutomatic** | **Xlsb**   | **29.90 ms** | **2.012 ms** | **1.882 ms** |  **1.00** |    **0.08** | **166.6667** |       **-** |   **8.02 MB** |        **1.00** |
| OfficeIMOAutomatic   | Xlsb   | 35.61 ms | 1.293 ms | 1.209 ms |  1.20 |    0.08 | 482.7586 | 34.4828 |  24.66 MB |        3.08 |


## qualified-830951-long-v2-domain1-strings

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T19:26:11.9898963+00:00; finished 2026-10-08T19:47:16.9251694+00:00. Exact context: contexts/qualified-830951-long-v2-domain1-strings-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.MaterializedStringHeavyReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-OWHAFV : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                         | Format | RowCount | Mean      | Error     | StdDev   | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated | Alloc Ratio |
|------------------------------- |------- |--------- |----------:|----------:|---------:|------:|--------:|---------:|---------:|---------:|----------:|------------:|
| **ExcelReaderStringsMaterialized** | **Xlsx**   | **65536**    |  **46.95 ms** |  **2.864 ms** | **2.236 ms** |  **1.00** |    **0.06** | **368.4211** | **315.7895** | **105.2632** |  **15.25 MB** |        **1.00** |
| ExcelReaderStringsInterned     | Xlsx   | 65536    |  45.93 ms |  1.602 ms | 1.251 ms |  0.98 |    0.05 | 350.0000 | 300.0000 | 100.0000 |  15.26 MB |        1.00 |
| OfficeIMOBytes                 | Xlsx   | 65536    |  77.03 ms |  6.986 ms | 5.454 ms |  1.64 |    0.13 | 454.5455 | 363.6364 |  90.9091 |  19.11 MB |        1.25 |
| OfficeIMOStream                | Xlsx   | 65536    |  75.29 ms |  6.806 ms | 5.313 ms |  1.61 |    0.13 | 416.6667 | 333.3333 |  83.3333 |  25.45 MB |        1.67 |
| Sylvan                         | Xlsx   | 65536    | 135.31 ms | 10.758 ms | 8.399 ms |  2.89 |    0.21 | 142.8571 |        - |        - |  17.41 MB |        1.14 |
|                                |        |          |           |           |          |       |         |          |          |          |           |             |
| **ExcelReaderStringsMaterialized** | **Xlsb**   | **65536**    |  **46.71 ms** |  **2.900 ms** | **2.264 ms** |  **1.00** |    **0.07** | **368.4211** | **315.7895** | **105.2632** |  **15.25 MB** |        **1.00** |
| ExcelReaderStringsInterned     | Xlsb   | 65536    |  47.28 ms |  2.627 ms | 2.051 ms |  1.01 |    0.06 | 363.6364 | 318.1818 |  90.9091 |  15.26 MB |        1.00 |
| OfficeIMOBytes                 | Xlsb   | 65536    |  49.55 ms |  3.266 ms | 2.550 ms |  1.06 |    0.07 | 526.3158 | 473.6842 | 157.8947 |  31.55 MB |        2.07 |
| OfficeIMOStream                | Xlsb   | 65536    |  47.82 ms |  2.757 ms | 2.152 ms |  1.03 |    0.06 | 500.0000 | 450.0000 | 150.0000 |   39.3 MB |        2.58 |
| Sylvan                         | Xlsb   | 65536    |  57.86 ms |  5.227 ms | 4.081 ms |  1.24 |    0.10 | 388.8889 | 333.3333 | 111.1111 |  17.38 MB |        1.14 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.SharedStringBorrowedScanBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-OWHAFV : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                      | RowCount | Storage  | Mean     | Error    | StdDev   | Ratio | RatioSD | Allocated | Alloc Ratio |
|---------------------------- |--------- |--------- |---------:|---------:|---------:|------:|--------:|----------:|------------:|
| **ExcelReaderBorrowed**         | **65536**    | **Deflated** | **35.73 ms** | **1.793 ms** | **1.400 ms** |  **1.00** |    **0.05** |   **2.18 MB** |        **1.00** |
| ExcelReaderBorrowedPrefetch | 65536    | Deflated | 28.28 ms | 1.028 ms | 0.803 ms |  0.79 |    0.04 |   2.26 MB |        1.04 |
|                             |          |          |          |          |          |       |         |           |             |
| **ExcelReaderBorrowed**         | **65536**    | **Stored**   | **14.93 ms** | **1.193 ms** | **0.931 ms** |  **1.00** |    **0.08** |   **2.18 MB** |        **1.00** |
| ExcelReaderBorrowedPrefetch | 65536    | Stored   | 16.68 ms | 0.847 ms | 0.661 ms |  1.12 |    0.08 |   2.19 MB |        1.01 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.SharedStringFirstRowBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-OWHAFV : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                         | RowCount | Storage  | Mean         | Error        | StdDev       | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated   | Alloc Ratio |
|------------------------------- |--------- |--------- |-------------:|-------------:|-------------:|------:|--------:|---------:|---------:|---------:|------------:|------------:|
| **ExcelReaderOpenThroughFirstRow** | **65536**    | **Deflated** | **13,267.99 μs** |   **497.592 μs** |   **388.487 μs** | **1.001** |    **0.04** |        **-** |        **-** |        **-** |    **747.3 KB** |        **1.00** |
| OfficeIMOOpenThroughFirstRow   | 65536    | Deflated | 68,120.31 μs | 7,081.149 μs | 5,528.496 μs | 5.138 |    0.43 | 333.3333 | 266.6667 |  66.6667 | 14958.48 KB |       20.02 |
| SylvanOpenThroughFirstRow      | 65536    | Deflated |     60.79 μs |     5.306 μs |     4.143 μs | 0.005 |    0.00 |   4.7401 |   1.7086 |   0.8268 |   320.92 KB |        0.43 |
|                                |          |          |              |              |              |       |         |          |          |          |             |             |
| **ExcelReaderOpenThroughFirstRow** | **65536**    | **Stored**   |  **3,228.99 μs** |    **41.500 μs** |    **32.401 μs** |  **1.00** |    **0.01** |        **-** |        **-** |        **-** |   **746.25 KB** |        **1.00** |
| OfficeIMOOpenThroughFirstRow   | 65536    | Stored   | 42,783.72 μs | 3,898.120 μs | 3,043.396 μs | 13.25 |    0.91 | 384.6154 | 346.1538 | 115.3846 | 14956.58 KB |       20.04 |
| SylvanOpenThroughFirstRow      | 65536    | Stored   |     57.84 μs |     6.084 μs |     4.750 μs |  0.02 |    0.00 |   5.8116 |   2.6479 |   1.9257 |   320.09 KB |        0.43 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.SharedStringFullScanBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-OWHAFV : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                         | RowCount | Storage | Mean      | Error    | StdDev   | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated | Alloc Ratio |
|------------------------------- |--------- |-------- |----------:|---------:|---------:|------:|--------:|---------:|---------:|---------:|----------:|------------:|
| ExcelReaderStringsMaterialized | 65536    | Stored  |  39.79 ms | 7.227 ms | 5.642 ms |  1.02 |    0.19 | 892.8571 | 857.1429 | 285.7143 |   15.6 MB |        1.00 |
| OfficeIMOStringsMaterialized   | 65536    | Stored  |  63.89 ms | 4.763 ms | 3.719 ms |  1.63 |    0.22 | 466.6667 | 400.0000 | 133.3333 |  19.11 MB |        1.22 |
| SylvanStringsMaterialized      | 65536    | Stored  | 122.70 ms | 9.804 ms | 7.655 ms |  3.14 |    0.43 | 142.8571 |        - |        - |  17.41 MB |        1.12 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.SharedStringUtf8ReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-OWHAFV : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                      | Format | RowCount | Storage  | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0      | Gen1      | Gen2     | Allocated | Alloc Ratio |
|---------------------------- |------- |--------- |--------- |----------:|----------:|----------:|------:|--------:|----------:|----------:|---------:|----------:|------------:|
| **ExcelReaderBorrowed**         | **Xlsx**   | **100000**   | **Deflated** |  **49.95 ms** |  **2.112 ms** |  **1.649 ms** |  **1.00** |    **0.04** |         **-** |         **-** |        **-** |    **2.9 MB** |        **1.00** |
| ExcelReaderBorrowedPrefetch | Xlsx   | 100000   | Deflated |  40.32 ms |  1.608 ms |  1.255 ms |  0.81 |    0.03 |         - |         - |        - |      3 MB |        1.03 |
| OfficeIMOBorrowedSharedUtf8 | Xlsx   | 100000   | Deflated | 159.41 ms | 11.571 ms |  9.034 ms |  3.19 |    0.20 | 2166.6667 | 1500.0000 | 500.0000 |  88.74 MB |       30.64 |
|                             |        |          |          |           |           |           |       |         |           |           |          |           |             |
| **ExcelReaderBorrowed**         | **Xlsx**   | **100000**   | **Stored**   |  **24.25 ms** |  **1.712 ms** |  **1.336 ms** |  **1.00** |    **0.07** |         **-** |         **-** |        **-** |   **2.89 MB** |        **1.00** |
| ExcelReaderBorrowedPrefetch | Xlsx   | 100000   | Stored   |  24.85 ms |  0.982 ms |  0.767 ms |  1.03 |    0.06 |         - |         - |        - |   2.91 MB |        1.01 |
| OfficeIMOBorrowedSharedUtf8 | Xlsx   | 100000   | Stored   | 204.63 ms | 18.465 ms | 14.416 ms |  8.46 |    0.72 | 3000.0000 | 1750.0000 | 500.0000 |  89.26 MB |       30.83 |
|                             |        |          |          |           |           |           |       |         |           |           |          |           |             |
| **ExcelReaderBorrowed**         | **Xlsb**   | **100000**   | **Deflated** |  **52.85 ms** |  **5.631 ms** |  **4.397 ms** |  **1.01** |    **0.12** |         **-** |         **-** |        **-** |    **2.9 MB** |        **1.00** |
| ExcelReaderBorrowedPrefetch | Xlsb   | 100000   | Deflated |  32.97 ms |  0.396 ms |  0.309 ms |  0.63 |    0.05 |         - |         - |        - |   2.99 MB |        1.03 |
| OfficeIMOBorrowedSharedUtf8 | Xlsb   | 100000   | Deflated | 103.38 ms |  3.624 ms |  2.829 ms |  1.97 |    0.17 | 2142.8571 | 1428.5714 | 428.5714 | 103.44 MB |       35.73 |
|                             |        |          |          |           |           |           |       |         |           |           |          |           |             |
| **ExcelReaderBorrowed**         | **Xlsb**   | **100000**   | **Stored**   |  **20.42 ms** |  **0.635 ms** |  **0.496 ms** |  **1.00** |    **0.03** |         **-** |         **-** |        **-** |   **2.89 MB** |        **1.00** |
| ExcelReaderBorrowedPrefetch | Xlsb   | 100000   | Stored   |  22.94 ms |  1.144 ms |  0.893 ms |  1.12 |    0.05 |         - |         - |        - |   2.91 MB |        1.01 |
| OfficeIMOBorrowedSharedUtf8 | Xlsb   | 100000   | Stored   |  78.48 ms |  3.948 ms |  3.082 ms |  3.85 |    0.17 | 2333.3333 | 1666.6667 | 555.5556 | 103.44 MB |       35.74 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.StringHeavyReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-OWHAFV : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method              | Format | RowCount | Mean      | Error     | StdDev   | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated | Alloc Ratio |
|-------------------- |------- |--------- |----------:|----------:|---------:|------:|--------:|---------:|---------:|---------:|----------:|------------:|
| **ExcelReaderOriginal** | **Xlsx**   | **65536**    |  **35.43 ms** |  **1.035 ms** | **0.808 ms** |  **1.00** |    **0.03** |        **-** |        **-** |        **-** |   **2.18 MB** |        **1.00** |
| ExcelReaderPrefetch | Xlsx   | 65536    |  28.05 ms |  0.449 ms | 0.350 ms |  0.79 |    0.02 |        - |        - |        - |   2.26 MB |        1.04 |
| ExcelReaderMemory   | Xlsx   | 65536    |  34.91 ms |  1.755 ms | 1.370 ms |  0.99 |    0.04 |        - |        - |        - |   2.18 MB |        1.00 |
| OfficeIMOBytes      | Xlsx   | 65536    |  91.32 ms | 12.775 ms | 9.974 ms |  2.58 |    0.28 | 454.5455 | 363.6364 |  90.9091 |  19.11 MB |        8.76 |
| OfficeIMOStream     | Xlsx   | 65536    |  72.77 ms |  2.116 ms | 1.652 ms |  2.05 |    0.06 | 416.6667 | 333.3333 |  83.3333 |  25.45 MB |       11.67 |
| Sylvan              | Xlsx   | 65536    | 134.17 ms | 10.080 ms | 7.870 ms |  3.79 |    0.23 | 166.6667 |        - |        - |  17.41 MB |        7.98 |
|                     |        |          |           |           |          |       |         |          |          |          |           |             |
| **ExcelReaderOriginal** | **Xlsb**   | **65536**    |  **31.99 ms** |  **0.543 ms** | **0.424 ms** |  **1.00** |    **0.02** |        **-** |        **-** |        **-** |   **2.18 MB** |        **1.00** |
| ExcelReaderPrefetch | Xlsb   | 65536    |  20.10 ms |  0.999 ms | 0.780 ms |  0.63 |    0.02 |        - |        - |        - |   2.27 MB |        1.04 |
| ExcelReaderMemory   | Xlsb   | 65536    |  30.18 ms |  0.191 ms | 0.149 ms |  0.94 |    0.01 |        - |        - |        - |   6.47 MB |        2.97 |
| OfficeIMOBytes      | Xlsb   | 65536    |  44.37 ms |  0.508 ms | 0.397 ms |  1.39 |    0.02 | 476.1905 | 428.5714 | 142.8571 |  31.55 MB |       14.47 |
| OfficeIMOStream     | Xlsb   | 65536    |  46.12 ms |  1.263 ms | 0.986 ms |  1.44 |    0.03 | 466.6667 | 400.0000 | 133.3333 |   39.3 MB |       18.03 |
| Sylvan              | Xlsb   | 65536    |  46.12 ms |  1.915 ms | 1.495 ms |  1.44 |    0.05 | 363.6364 | 272.7273 |  90.9091 |  17.38 MB |        7.97 |


## qualified-830951-long-v2-domain1-ado

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T19:47:17.2724566+00:00; finished 2026-10-08T19:52:15.6368503+00:00. Exact context: contexts/qualified-830951-long-v2-domain1-ado-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.AdoReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-CPEWJC : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method      | RowCount | Access       | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1    | Gen2    | Allocated  | Alloc Ratio |
|------------ |--------- |------------- |----------:|----------:|----------:|------:|--------:|---------:|--------:|--------:|-----------:|------------:|
| **ExcelReader** | **50000**    | **GetValue**     |  **6.784 ms** | **0.1010 ms** | **0.0945 ms** |  **1.00** |    **0.02** | **100.0000** |       **-** |       **-** | **5131.51 KB** |        **1.00** |
| OfficeIMO   | 50000    | GetValue     |  9.782 ms | 0.1302 ms | 0.1218 ms |  1.44 |    0.03 | 107.8431 | 39.2157 | 39.2157 | 4348.84 KB |        0.85 |
| Sylvan      | 50000    | GetValue     | 33.418 ms | 0.3170 ms | 0.2966 ms |  4.93 |    0.08 | 137.9310 |       - |       - | 8382.24 KB |        1.63 |
|             |          |              |           |           |           |       |         |          |         |         |            |             |
| **ExcelReader** | **50000**    | **TypedGetters** |  **6.904 ms** | **0.0589 ms** | **0.0551 ms** |  **1.00** |    **0.01** |  **27.7778** |       **-** |       **-** | **1615.88 KB** |        **1.00** |
| OfficeIMO   | 50000    | TypedGetters |  8.398 ms | 0.0696 ms | 0.0651 ms |  1.22 |    0.01 |  61.4035 | 61.4035 | 61.4035 | 1011.23 KB |        0.63 |
| Sylvan      | 50000    | TypedGetters | 25.106 ms | 0.2429 ms | 0.2272 ms |  3.64 |    0.04 |  25.0000 |       - |       - | 1939.92 KB |        1.20 |
|             |          |              |           |           |           |       |         |          |         |         |            |             |
| **ExcelReader** | **50000**    | **Utf8TextCopy** |  **5.141 ms** | **0.0463 ms** | **0.0433 ms** |  **1.00** |    **0.01** |        **-** |       **-** |       **-** |    **6.93 KB** |        **1.00** |
| OfficeIMO   | 50000    | Utf8TextCopy |  6.227 ms | 0.1093 ms | 0.1022 ms |  1.21 |    0.02 |  42.8571 | 42.8571 | 42.8571 |   832.9 KB |      120.11 |
| Sylvan      | 50000    | Utf8TextCopy | 21.799 ms | 0.1943 ms | 0.1817 ms |  4.24 |    0.05 |  66.6667 |       - |       - | 4283.37 KB |      617.68 |


## qualified-830951-long-v2-domain1-writing

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T19:52:15.9367041+00:00; finished 2026-10-08T20:06:20.6911046+00:00. Exact context: contexts/qualified-830951-long-v2-domain1-writing-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.ArrowStringWriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-OWHAFV : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                          | Mean     | Error    | StdDev   | Ratio | RatioSD | Allocated  | Alloc Ratio |
|-------------------------------- |---------:|---------:|---------:|------:|--------:|-----------:|------------:|
| ExcelReaderRecordBatch          | 18.24 ms | 0.219 ms | 0.171 ms |  1.00 |    0.01 |   18.06 KB |        1.00 |
| OfficeIMOPublicRowOrchestration | 41.46 ms | 0.523 ms | 0.408 ms |  2.27 |    0.03 | 1115.39 KB |       61.75 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.CompactWriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-OWHAFV : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                     | RowCount | Mean     | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated | Alloc Ratio |
|--------------------------- |--------- |---------:|----------:|----------:|------:|--------:|---------:|---------:|---------:|----------:|------------:|
| OfficeIMOWithoutReferences | 50000    | 7.309 ms | 0.1404 ms | 0.1097 ms |  0.93 |    0.02 | 328.0000 | 328.0000 | 328.0000 |   4.04 MB |        1.00 |
| ExcelReaderWriter          | 50000    | 7.893 ms | 0.2093 ms | 0.1634 ms |  1.00 |    0.03 | 328.2443 | 328.2443 | 328.2443 |   4.02 MB |        1.00 |
| ExcelReaderWriterPrefetch  | 50000    | 4.441 ms | 0.2375 ms | 0.1854 ms |  0.56 |    0.03 | 333.3333 | 333.3333 | 333.3333 |   4.02 MB |        1.00 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.ConfiguredXlsbWriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-OWHAFV : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                             | RowCount | Mean       | Error      | StdDev     | Ratio | RatioSD | Gen0      | Gen1      | Gen2      | Allocated | Alloc Ratio |
|----------------------------------- |--------- |-----------:|-----------:|-----------:|------:|--------:|----------:|----------:|----------:|----------:|------------:|
| OfficeIMOModelAndSaveInlineStrings | 50000    | 545.752 ms | 17.0112 ms | 13.2812 ms | 98.02 |    3.13 | 9750.0000 | 8250.0000 | 2500.0000 | 439.73 MB |      108.36 |
| OfficeIMOModelAndSaveSharedStrings | 50000    | 543.015 ms | 13.8026 ms | 10.7761 ms | 97.53 |    2.82 | 9500.0000 | 8000.0000 | 2250.0000 | 430.12 MB |      106.00 |
| ExcelReaderXlsbWriterSharedStrings | 50000    |   5.570 ms |  0.1655 ms |  0.1292 ms |  1.00 |    0.03 |  329.2683 |  329.2683 |  329.2683 |   4.06 MB |        1.00 |
| ExcelReaderXlsbWriterPrefetch      | 50000    |   4.345 ms |  0.0586 ms |  0.0457 ms |  0.78 |    0.02 |  331.8966 |  331.8966 |  331.8966 |   4.03 MB |        0.99 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.MappedRecordWriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-OWHAFV : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                    | Format | RowCount | Mean       | Error      | StdDev     | Ratio | RatioSD | Gen0      | Gen1      | Gen2      | Allocated | Alloc Ratio |
|-------------------------- |------- |--------- |-----------:|-----------:|-----------:|------:|--------:|----------:|----------:|----------:|----------:|------------:|
| **OfficeIMOPublicTypedWrite** | **Xlsx**   | **50000**    |  **15.021 ms** |  **0.5552 ms** |  **0.4335 ms** |  **1.52** |    **0.13** |  **246.1538** |  **246.1538** |  **246.1538** |   **4.04 MB** |        **1.00** |
| ExcelReaderMappedRecords  | Xlsx   | 50000    |   9.941 ms |  1.2096 ms |  0.9443 ms |  1.01 |    0.12 |  238.0952 |  238.0952 |  238.0952 |   4.02 MB |        1.00 |
|                           |        |          |            |            |            |       |         |           |           |           |           |             |
| **OfficeIMOPublicTypedWrite** | **Xlsb**   | **50000**    | **620.726 ms** | **55.3748 ms** | **43.2330 ms** | **93.93** |    **7.14** | **9500.0000** | **8000.0000** | **2250.0000** | **439.72 MB** |      **109.36** |
| ExcelReaderMappedRecords  | Xlsb   | 50000    |   6.617 ms |  0.3174 ms |  0.2478 ms |  1.00 |    0.05 |  326.5306 |  326.5306 |  326.5306 |   4.02 MB |        1.00 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.NativeWriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-OWHAFV : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                | Format | RowCount | Mean       | Error      | StdDev     | Ratio  | RatioSD | Gen0      | Gen1      | Gen2      | Allocated | Alloc Ratio |
|---------------------- |------- |--------- |-----------:|-----------:|-----------:|-------:|--------:|----------:|----------:|----------:|----------:|------------:|
| OfficeIMOModelAndSave | Xlsb   | 50000    | 619.232 ms | 28.6043 ms | 22.3323 ms | 107.97 |    5.41 | 9500.0000 | 8000.0000 | 2250.0000 | 439.72 MB |      109.36 |
| ExcelReaderWriter     | Xlsb   | 50000    |   5.743 ms |  0.2886 ms |  0.2253 ms |   1.00 |    0.05 |  333.3333 |  333.3333 |  333.3333 |   4.02 MB |        1.00 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.RecordWriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-OWHAFV : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                    | Format | RowCount | Mean       | Error      | StdDev     | Ratio  | RatioSD | Gen0      | Gen1      | Gen2      | Allocated | Alloc Ratio |
|-------------------------- |------- |--------- |-----------:|-----------:|-----------:|-------:|--------:|----------:|----------:|----------:|----------:|------------:|
| **OfficeIMOPublicTypedWrite** | **Xlsx**   | **50000**    |  **14.743 ms** |  **0.5948 ms** |  **0.4644 ms** |   **1.66** |    **0.09** |  **250.0000** |  **250.0000** |  **250.0000** |   **4.04 MB** |        **1.00** |
| ExcelReaderWriteRecords   | Xlsx   | 50000    |   8.909 ms |  0.5192 ms |  0.4053 ms |   1.00 |    0.06 |  250.0000 |  250.0000 |  250.0000 |   4.02 MB |        1.00 |
|                           |        |          |            |            |            |        |         |           |           |           |           |             |
| **OfficeIMOPublicTypedWrite** | **Xlsb**   | **50000**    | **603.284 ms** | **25.9465 ms** | **20.2573 ms** | **106.42** |    **4.65** | **9750.0000** | **8250.0000** | **2500.0000** | **439.72 MB** |      **109.36** |
| ExcelReaderWriteRecords   | Xlsb   | 50000    |   5.674 ms |  0.2221 ms |  0.1734 ms |   1.00 |    0.04 |  329.2683 |  329.2683 |  329.2683 |   4.02 MB |        1.00 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.SharedStringWriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-OWHAFV : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                                  | RowCount | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated | Alloc Ratio |
|---------------------------------------- |--------- |----------:|----------:|----------:|------:|--------:|---------:|---------:|---------:|----------:|------------:|
| OfficeIMOSharedStrings                  | 50000    | 15.575 ms | 0.5416 ms | 0.4228 ms |  1.94 |    0.06 | 323.0769 | 323.0769 | 323.0769 |   4.04 MB |        1.00 |
| OfficeIMOSharedStringsWithoutReferences | 50000    |  8.630 ms | 0.7475 ms | 0.5836 ms |  1.08 |    0.07 | 333.3333 | 333.3333 | 333.3333 |   4.04 MB |        1.00 |
| ExcelReaderWriterSharedStrings          | 50000    |  8.017 ms | 0.1860 ms | 0.1453 ms |  1.00 |    0.02 | 333.3333 | 333.3333 | 333.3333 |   4.06 MB |        1.00 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.StyledRowWriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-OWHAFV : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                        | Mean     | Error     | StdDev    | Ratio | RatioSD | Gen0    | Allocated  | Alloc Ratio |
|------------------------------ |---------:|----------:|----------:|------:|--------:|--------:|-----------:|------------:|
| OfficeIMODefaultRowStyle      | 8.763 ms | 0.5457 ms | 0.4261 ms |  1.70 |    0.12 | 52.1739 | 2563.11 KB |      140.87 |
| ExcelReaderOriginalStyledRows | 5.171 ms | 0.3944 ms | 0.3079 ms |  1.00 |    0.08 |       - |    18.2 KB |        1.00 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.Utf8WriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-OWHAFV : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                | Mean     | Error    | StdDev   | Ratio | RatioSD | Allocated | Alloc Ratio |
|---------------------- |---------:|---------:|---------:|------:|--------:|----------:|------------:|
| OfficeIMOUtf8         | 28.79 ms | 4.705 ms | 3.673 ms |  1.96 |    0.25 |  38.37 KB |        2.11 |
| ExcelReaderPublicUtf8 | 14.69 ms | 0.559 ms | 0.437 ms |  1.00 |    0.04 |  18.18 KB |        1.00 |


## qualified-830951-long-v2-domain1-csv

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T20:06:21.0503545+00:00; finished 2026-10-08T20:28:19.534304+00:00. Exact context: contexts/qualified-830951-long-v2-domain1-csv-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.CsvMaterializedReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-CPEWJC : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                  | RowCount | Mean     | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1   | Allocated | Alloc Ratio |
|------------------------ |--------- |---------:|----------:|----------:|------:|--------:|---------:|-------:|----------:|------------:|
| ExcelReaderMaterialized | 50000    | 2.407 ms | 0.0869 ms | 0.0812 ms |  1.00 |    0.05 |  30.8057 |      - |   1.57 MB |        1.00 |
| OfficeIMOMaterialized   | 50000    | 5.034 ms | 0.2959 ms | 0.2767 ms |  2.09 |    0.13 | 180.2575 |      - |   8.63 MB |        5.48 |
| SylvanMaterialized      | 50000    | 3.160 ms | 0.1226 ms | 0.1146 ms |  1.31 |    0.06 |  32.6797 | 3.2680 |   1.61 MB |        1.02 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.CsvRawAsyncBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-CPEWJC : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method           | RowCount | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1    | Allocated | Alloc Ratio |
|----------------- |--------- |----------:|----------:|----------:|------:|--------:|---------:|--------:|----------:|------------:|
| ExcelReaderAsync | 50000    |  2.514 ms | 0.1458 ms | 0.1364 ms |  1.00 |    0.07 |  32.1839 |       - |   1.57 MB |        1.00 |
| OfficeIMOAsync   | 50000    | 10.801 ms | 0.4394 ms | 0.4110 ms |  4.31 |    0.28 | 288.8889 | 22.2222 |  14.18 MB |        9.01 |
| SylvanAsync      | 50000    |  3.666 ms | 0.1544 ms | 0.1444 ms |  1.46 |    0.10 |  33.5570 |  3.3557 |   1.62 MB |        1.03 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.CsvRawReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-CPEWJC : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                | RowCount | Mean     | Error     | StdDev    | Ratio | RatioSD | Allocated | Alloc Ratio |
|---------------------- |--------- |---------:|----------:|----------:|------:|--------:|----------:|------------:|
| ExcelReaderOriginal   | 50000    | 1.949 ms | 0.0650 ms | 0.0608 ms |  1.00 |    0.04 |     848 B |        1.00 |
| OfficeIMOBorrowedUtf8 | 50000    | 4.751 ms | 0.2910 ms | 0.2722 ms |  2.44 |    0.15 |    2120 B |        2.50 |
| SylvanOriginal        | 50000    | 2.891 ms | 0.1020 ms | 0.0954 ms |  1.48 |    0.06 |   38681 B |       45.61 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.CsvRealDataMaterializedReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-CPEWJC : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                  | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1     | Allocated | Alloc Ratio |
|------------------------ |----------:|----------:|----------:|------:|--------:|---------:|---------:|----------:|------------:|
| ExcelReaderMaterialized | 23.008 ms | 1.0159 ms | 0.9503 ms |  1.00 |    0.06 | 739.1304 |        - |  35.71 MB |        1.00 |
| OfficeIMOMaterialized   | 13.001 ms | 0.4700 ms | 0.4397 ms |  0.57 |    0.03 | 702.7027 |        - |  34.21 MB |        0.96 |
| SylvanMaterialized      |  9.353 ms | 0.4793 ms | 0.4483 ms |  0.41 |    0.02 | 745.6140 | 105.2632 |  35.75 MB |        1.00 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.CsvRealDataReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-CPEWJC : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                      | Mean     | Error     | StdDev    | Ratio | RatioSD | Allocated | Alloc Ratio |
|---------------------------- |---------:|----------:|----------:|------:|--------:|----------:|------------:|
| ExcelReaderStream           | 6.691 ms | 0.1285 ms | 0.1202 ms |  1.00 |    0.02 |     848 B |        1.00 |
| ExcelReaderMemory           | 7.281 ms | 0.2017 ms | 0.1887 ms |  1.09 |    0.03 |     576 B |        0.68 |
| OfficeIMOStreamBorrowedUtf8 | 4.863 ms | 0.2364 ms | 0.2212 ms |  0.73 |    0.03 |    2920 B |        3.44 |
| SylvanFieldSpan             | 6.380 ms | 0.3409 ms | 0.3189 ms |  0.95 |    0.05 |   41225 B |       48.61 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.CsvRecordWriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-CPEWJC : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                  | Mapped | RowCount | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated | Alloc Ratio |
|------------------------ |------- |--------- |----------:|----------:|----------:|------:|--------:|---------:|---------:|---------:|----------:|------------:|
| **ExcelReaderRecordLayout** | **False**  | **50000**    |  **3.626 ms** | **0.1733 ms** | **0.1621 ms** |  **1.00** |    **0.06** | **333.3333** | **333.3333** | **333.3333** |      **4 MB** |        **1.00** |
| OfficeIMOWriteObjects   | False  | 50000    | 10.275 ms | 0.2333 ms | 0.2183 ms |  2.84 |    0.13 | 489.7959 | 346.9388 | 346.9388 |  11.26 MB |        2.81 |
|                         |        |          |           |           |           |       |         |          |          |          |           |             |
| **ExcelReaderRecordLayout** | **True**   | **50000**    |  **4.321 ms** | **0.1334 ms** | **0.1248 ms** |  **1.00** |    **0.04** | **331.8777** | **331.8777** | **331.8777** |      **4 MB** |        **1.00** |
| OfficeIMOWriteObjects   | True   | 50000    | 10.126 ms | 0.2796 ms | 0.2615 ms |  2.35 |    0.09 | 484.5361 | 340.2062 | 340.2062 |  11.26 MB |        2.81 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.CsvTypedAsyncBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-CPEWJC : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                | RowCount | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1    | Allocated | Alloc Ratio |
|---------------------- |--------- |----------:|----------:|----------:|------:|--------:|---------:|--------:|----------:|------------:|
| ExcelReaderTypedAsync | 50000    |  4.311 ms | 0.4236 ms | 0.3962 ms |  1.01 |    0.13 |  80.0000 |       - |   3.86 MB |        1.00 |
| OfficeIMORowsAsAsync  | 50000    | 13.541 ms | 0.6190 ms | 0.5790 ms |  3.17 |    0.31 | 333.3333 | 24.6914 |  16.47 MB |        4.26 |
| SylvanTypedAsync      | 50000    |  8.721 ms | 0.4959 ms | 0.4639 ms |  2.04 |    0.21 | 222.2222 | 23.8095 |  10.96 MB |        2.84 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.CsvTypedReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-CPEWJC : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method           | RowCount | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1    | Allocated | Alloc Ratio |
|----------------- |--------- |----------:|----------:|----------:|------:|--------:|---------:|--------:|----------:|------------:|
| ExcelReaderTyped | 50000    |  3.835 ms | 0.2172 ms | 0.2031 ms |  1.00 |    0.07 |  80.4196 |       - |   3.86 MB |        1.00 |
| OfficeIMORowsAs  | 50000    |  5.820 ms | 0.3309 ms | 0.3095 ms |  1.52 |    0.11 | 226.7442 |       - |  10.92 MB |        2.83 |
| SylvanTyped      | 50000    | 10.764 ms | 1.1752 ms | 1.0993 ms |  2.81 |    0.31 | 228.5714 | 28.5714 |  10.95 MB |        2.84 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.CsvUtf8WriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-CPEWJC : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                 | Mean     | Error     | StdDev    | Ratio | RatioSD | Allocated | Alloc Ratio |
|----------------------- |---------:|----------:|----------:|------:|--------:|----------:|------------:|
| OfficeIMOPublicUtf8Row | 8.907 ms | 0.4922 ms | 0.4604 ms |  2.82 |    0.21 |    2704 B |        6.50 |
| ExcelReaderPublicUtf8  | 3.168 ms | 0.2075 ms | 0.1941 ms |  1.00 |    0.08 |     416 B |        1.00 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.CsvWideReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-CPEWJC : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method          | RowCount | MaterializeStrings | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0      | Gen1     | Allocated  | Alloc Ratio |
|---------------- |--------- |------------------- |----------:|----------:|----------:|------:|--------:|----------:|---------:|-----------:|------------:|
| **ExcelReaderWide** | **50000**    | **False**              |  **5.487 ms** | **0.2710 ms** | **0.2535 ms** |  **1.00** |    **0.06** |         **-** |        **-** |      **848 B** |        **1.00** |
| OfficeIMOWide   | 50000    | False              |  6.678 ms | 0.4516 ms | 0.4224 ms |  1.22 |    0.09 |         - |        - |     4360 B |        5.14 |
| SylvanWide      | 50000    | False              |  4.650 ms | 0.4114 ms | 0.3848 ms |  0.85 |    0.08 |         - |        - |    46377 B |       54.69 |
|                 |          |                    |           |           |           |       |         |           |          |            |             |
| **ExcelReaderWide** | **50000**    | **True**               | **20.947 ms** | **2.0688 ms** | **1.9351 ms** |  **1.01** |    **0.12** | **1037.7358** |        **-** | **52800848 B** |        **1.00** |
| OfficeIMOWide   | 50000    | True               | 14.023 ms | 0.9493 ms | 0.8880 ms |  0.67 |    0.07 | 1049.3827 |  12.3457 | 52804360 B |        1.00 |
| SylvanWide      | 50000    | True               | 11.755 ms | 0.8562 ms | 0.8009 ms |  0.57 |    0.06 | 1043.9560 | 164.8352 | 52846584 B |        1.00 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.CsvWriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-CPEWJC : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                | RowCount | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated | Alloc Ratio |
|---------------------- |--------- |----------:|----------:|----------:|------:|--------:|---------:|---------:|---------:|----------:|------------:|
| ExcelReaderWriter     | 50000    |  3.721 ms | 0.1842 ms | 0.1723 ms |  1.00 |    0.06 | 500.0000 | 500.0000 | 500.0000 |      4 MB |        1.00 |
| OfficeIMOWriteObjects | 50000    | 10.978 ms | 1.0974 ms | 1.0265 ms |  2.96 |    0.30 | 691.5888 | 542.0561 | 542.0561 |  11.26 MB |        2.81 |
| SylvanWriter          | 50000    |  5.427 ms | 0.3345 ms | 0.3128 ms |  1.46 |    0.10 | 500.0000 | 500.0000 | 500.0000 |   4.04 MB |        1.01 |


## qualified-830951-long-v2-domain1-csv-large-dop1

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T20:28:19.9153196+00:00; finished 2026-10-08T20:55:30.1219312+00:00. Exact context: contexts/qualified-830951-long-v2-domain1-csv-large-dop1-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.CsvDirectAggregateBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-SVEHGN : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=8  IterationTime=250ms  LaunchCount=1  
UnrollFactor=1  WarmupCount=4  

```
| Method                        | Input                | Dop | Mean       | Error     | StdDev    | Ratio | RatioSD | Gen0       | Gen1       | Gen2      | Allocated     | Alloc Ratio |
|------------------------------ |--------------------- |---- |-----------:|----------:|----------:|------:|--------:|-----------:|-----------:|----------:|--------------:|------------:|
| **ExcelReaderAggregate**          | **Conve(...)00000 [23]** | **1**   |   **599.4 ms** |  **69.45 ms** |  **36.32 ms** |  **1.00** |    **0.08** |          **-** |          **-** |         **-** |    **1228.89 KB** |        **1.00** |
| OfficeIMOTextOwnedAggregate   | Conve(...)00000 [23] | 1   |   588.7 ms |  70.30 ms |  36.77 ms |  0.99 |    0.08 | 10500.0000 | 10250.0000 | 2250.0000 |  812776.71 KB |      661.39 |
| OfficeIMOAsyncPathAggregate   | Conve(...)00000 [23] | 1   | 2,197.8 ms | 361.28 ms | 188.96 ms |  3.68 |    0.36 | 38750.0000 |   250.0000 |         - | 1903103.02 KB |    1,548.63 |
| OfficeIMOAsyncStreamAggregate | Conve(...)00000 [23] | 1   | 2,262.4 ms | 185.54 ms |  97.04 ms |  3.79 |    0.25 | 38750.0000 |   250.0000 |         - | 1903113.93 KB |    1,548.64 |
|                               |                      |     |            |           |           |       |         |            |            |           |               |             |
| **ExcelReaderAggregate**          | **Conve(...)00000 [25]** | **1**   |   **445.8 ms** |  **40.86 ms** |  **21.37 ms** |  **1.00** |    **0.06** |          **-** |          **-** |         **-** |     **858.16 KB** |        **1.00** |
| OfficeIMOTextOwnedAggregate   | Conve(...)00000 [25] | 1   |   468.7 ms |  27.02 ms |  14.13 ms |  1.05 |    0.06 |  8000.0000 |  7750.0000 | 2250.0000 |  567066.42 KB |      660.80 |
| OfficeIMOAsyncPathAggregate   | Conve(...)00000 [25] | 1   | 1,551.4 ms | 188.41 ms |  98.54 ms |  3.49 |    0.26 | 27500.0000 |   500.0000 |         - | 1327763.03 KB |    1,547.23 |
| OfficeIMOAsyncStreamAggregate | Conve(...)00000 [25] | 1   | 1,632.9 ms | 152.35 ms |  79.68 ms |  3.67 |    0.24 | 27000.0000 |   250.0000 |         - | 1327760.91 KB |    1,547.23 |
|                               |                      |     |            |           |           |       |         |            |            |           |               |             |
| **ExcelReaderAggregate**          | **NarrowInt-8000000**    | **1**   |   **489.0 ms** |  **56.36 ms** |  **29.48 ms** |  **1.00** |    **0.08** |          **-** |          **-** |         **-** |    **1227.54 KB** |        **1.00** |
| OfficeIMOTextOwnedAggregate   | NarrowInt-8000000    | 1   |   685.1 ms |  47.42 ms |  24.80 ms |  1.41 |    0.09 | 10500.0000 | 10250.0000 | 2250.0000 |   820833.9 KB |      668.68 |
| OfficeIMOAsyncPathAggregate   | NarrowInt-8000000    | 1   | 2,621.9 ms | 294.42 ms | 153.99 ms |  5.38 |    0.43 | 45000.0000 |   250.0000 |         - | 2207893.19 KB |    1,798.64 |
| OfficeIMOAsyncStreamAggregate | NarrowInt-8000000    | 1   | 2,378.8 ms | 491.55 ms | 257.09 ms |  4.88 |    0.57 | 45000.0000 |   250.0000 |         - |  2207895.8 KB |    1,798.64 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.CsvParallelTypedBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-SVEHGN : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=8  IterationTime=250ms  LaunchCount=1  
UnrollFactor=1  WarmupCount=4  

```
| Method                | Input                | Dop | Mean     | Error     | StdDev    | Ratio | RatioSD | Gen0       | Gen1       | Gen2      | Allocated  | Alloc Ratio |
|---------------------- |--------------------- |---- |---------:|----------:|----------:|------:|--------:|-----------:|-----------:|----------:|-----------:|------------:|
| **ExcelReaderParallel**   | **Conve(...)00000 [23]** | **1**   | **618.4 ms** |  **11.22 ms** |   **5.87 ms** |  **1.00** |    **0.01** | **14000.0000** |          **-** |         **-** |  **671.41 MB** |        **1.00** |
| ExcelReaderSequential | Conve(...)00000 [23] | 1   | 834.6 ms |  95.34 ms |  49.86 ms |  1.35 |    0.08 | 13750.0000 |          - |         - |   669.8 MB |        1.00 |
| OfficeIMOCoreParallel | Conve(...)00000 [23] | 1   | 532.4 ms |  14.68 ms |   7.68 ms |  0.86 |    0.01 | 26750.0000 |          - |         - |  1290.2 MB |        1.92 |
| OfficeIMOTextParallel | Conve(...)00000 [23] | 1   | 612.5 ms |  16.80 ms |   8.79 ms |  0.99 |    0.02 | 24750.0000 | 11750.0000 | 2500.0000 | 1463.52 MB |        2.18 |
| SylvanSequential      | Conve(...)00000 [23] | 1   | 890.9 ms |  60.16 ms |  31.46 ms |  1.44 |    0.05 | 26750.0000 |          - |         - | 1292.01 MB |        1.92 |
|                       |                      |     |          |           |           |       |         |            |            |           |            |             |
| **ExcelReaderParallel**   | **Conve(...)00000 [25]** | **1**   | **474.5 ms** |  **42.23 ms** |  **22.09 ms** |  **1.00** |    **0.06** |  **9750.0000** |          **-** |         **-** |  **468.43 MB** |        **1.00** |
| ExcelReaderSequential | Conve(...)00000 [25] | 1   | 427.4 ms |  19.04 ms |   9.96 ms |  0.90 |    0.04 |  9750.0000 |          - |         - |   467.3 MB |        1.00 |
| OfficeIMOCoreParallel | Conve(...)00000 [25] | 1   | 499.7 ms |  59.89 ms |  31.32 ms |  1.06 |    0.08 | 18750.0000 |          - |         - |  900.14 MB |        1.92 |
| OfficeIMOTextParallel | Conve(...)00000 [25] | 1   | 570.5 ms |  33.88 ms |  17.72 ms |  1.20 |    0.06 | 25250.0000 | 14500.0000 | 3000.0000 | 1021.41 MB |        2.18 |
| SylvanSequential      | Conve(...)00000 [25] | 1   | 761.2 ms |  56.80 ms |  29.71 ms |  1.61 |    0.09 | 18750.0000 |          - |         - |  901.42 MB |        1.92 |
|                       |                      |     |          |           |           |       |         |            |            |           |            |             |
| **ExcelReaderParallel**   | **NarrowInt-8000000**    | **1**   | **487.7 ms** |  **36.12 ms** |  **18.89 ms** |  **1.00** |    **0.05** |  **5000.0000** |          **-** |         **-** |  **245.75 MB** |        **1.00** |
| ExcelReaderSequential | NarrowInt-8000000    | 1   | 385.6 ms | 121.76 ms |  63.68 ms |  0.79 |    0.13 |  5000.0000 |          - |         - |  244.14 MB |        0.99 |
| OfficeIMOCoreParallel | NarrowInt-8000000    | 1   | 491.1 ms |  16.87 ms |   8.82 ms |  1.01 |    0.04 | 24000.0000 |          - |         - | 1158.54 MB |        4.71 |
| OfficeIMOTextParallel | NarrowInt-8000000    | 1   | 894.3 ms |  63.25 ms |  33.08 ms |  1.84 |    0.09 | 16000.0000 | 11500.0000 | 2500.0000 | 1045.73 MB |        4.26 |
| SylvanSequential      | NarrowInt-8000000    | 1   | 649.8 ms | 215.79 ms | 112.86 ms |  1.33 |    0.22 | 24000.0000 |          - |         - | 1158.58 MB |        4.71 |


## qualified-830951-long-v2-domain1-csv-large-dop4

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T20:55:30.6642543+00:00; finished 2026-10-08T21:16:56.012155+00:00. Exact context: contexts/qualified-830951-long-v2-domain1-csv-large-dop4-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.CsvDirectAggregateBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-SVEHGN : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=8  IterationTime=250ms  LaunchCount=1  
UnrollFactor=1  WarmupCount=4  

```
| Method                        | Input                | Dop | Mean       | Error     | StdDev    | Ratio | RatioSD | Gen0       | Gen1       | Gen2      | Allocated  | Alloc Ratio |
|------------------------------ |--------------------- |---- |-----------:|----------:|----------:|------:|--------:|-----------:|-----------:|----------:|-----------:|------------:|
| **ExcelReaderAggregate**          | **Conve(...)00000 [23]** | **4**   |   **157.8 ms** |  **35.16 ms** |  **18.39 ms** |  **1.01** |    **0.15** |          **-** |          **-** |         **-** |    **1.45 MB** |        **1.00** |
| OfficeIMOTextOwnedAggregate   | Conve(...)00000 [23] | 4   |   525.7 ms |  39.89 ms |  20.86 ms |  3.37 |    0.38 | 11000.0000 | 10750.0000 | 2750.0000 |   794.3 MB |      547.31 |
| OfficeIMOAsyncPathAggregate   | Conve(...)00000 [23] | 4   | 2,154.5 ms | 851.05 ms | 445.12 ms | 13.81 |    3.07 | 39500.0000 |  4500.0000 |         - | 1864.92 MB |    1,285.02 |
| OfficeIMOAsyncStreamAggregate | Conve(...)00000 [23] | 4   | 1,706.7 ms | 261.12 ms | 136.57 ms | 10.94 |    1.42 | 38750.0000 |  4500.0000 |         - | 1864.92 MB |    1,285.02 |
|                               |                      |     |            |           |           |       |         |            |            |           |            |             |
| **ExcelReaderAggregate**          | **Conve(...)00000 [25]** | **4**   |   **128.4 ms** |  **12.24 ms** |   **6.40 ms** |  **1.00** |    **0.07** |          **-** |          **-** |         **-** |    **1.02 MB** |        **1.00** |
| OfficeIMOTextOwnedAggregate   | Conve(...)00000 [25] | 4   |   409.4 ms | 121.33 ms |  63.46 ms |  3.19 |    0.49 | 16500.0000 | 16250.0000 | 2750.0000 |  555.06 MB |      543.58 |
| OfficeIMOAsyncPathAggregate   | Conve(...)00000 [25] | 4   | 1,261.2 ms | 319.56 ms | 167.14 ms |  9.84 |    1.31 | 27000.0000 |  3000.0000 |         - | 1301.11 MB |    1,274.20 |
| OfficeIMOAsyncStreamAggregate | Conve(...)00000 [25] | 4   | 1,169.4 ms | 206.43 ms | 107.97 ms |  9.12 |    0.90 | 27000.0000 |  3250.0000 |         - | 1301.11 MB |    1,274.21 |
|                               |                      |     |            |           |           |       |         |            |            |           |            |             |
| **ExcelReaderAggregate**          | **NarrowInt-8000000**    | **4**   |   **134.1 ms** |  **25.22 ms** |  **13.19 ms** |  **1.01** |    **0.13** |          **-** |          **-** |         **-** |    **1.61 MB** |        **1.00** |
| OfficeIMOTextOwnedAggregate   | NarrowInt-8000000    | 4   |   825.0 ms | 109.75 ms |  57.40 ms |  6.20 |    0.69 | 22750.0000 | 22000.0000 | 3500.0000 |  802.39 MB |      498.01 |
| OfficeIMOAsyncPathAggregate   | NarrowInt-8000000    | 4   | 2,456.0 ms | 761.79 ms | 398.43 ms | 18.46 |    3.28 | 47250.0000 |  3250.0000 |  500.0000 | 2168.08 MB |    1,345.65 |
| OfficeIMOAsyncStreamAggregate | NarrowInt-8000000    | 4   | 2,780.6 ms | 441.52 ms | 230.92 ms | 20.90 |    2.49 | 45250.0000 |  3250.0000 |         - | 2168.06 MB |    1,345.63 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.CsvParallelTypedBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-SVEHGN : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=8  IterationTime=250ms  LaunchCount=1  
UnrollFactor=1  WarmupCount=4  

```
| Method                | Input                | Dop | Mean       | Error     | StdDev    | Ratio | RatioSD | Gen0       | Gen1       | Gen2      | Allocated  | Alloc Ratio |
|---------------------- |--------------------- |---- |-----------:|----------:|----------:|------:|--------:|-----------:|-----------:|----------:|-----------:|------------:|
| **ExcelReaderParallel**   | **Conve(...)00000 [23]** | **4**   |   **231.0 ms** |  **44.88 ms** |  **23.47 ms** |  **1.01** |    **0.13** | **14250.0000** | **11750.0000** |         **-** |  **680.18 MB** |        **1.00** |
| OfficeIMOCoreParallel | Conve(...)00000 [23] | 4   | 1,242.6 ms | 387.03 ms | 202.42 ms |  5.42 |    0.96 | 39750.0000 | 10500.0000 |         - | 1831.32 MB |        2.69 |
| OfficeIMOTextParallel | Conve(...)00000 [23] | 4   |   385.0 ms |  81.10 ms |  42.42 ms |  1.68 |    0.23 | 25000.0000 | 24750.0000 | 2750.0000 | 1464.38 MB |        2.15 |
|                       |                      |     |            |           |           |       |         |            |            |           |            |             |
| **ExcelReaderParallel**   | **Conve(...)00000 [25]** | **4**   |   **151.9 ms** |  **15.92 ms** |   **8.32 ms** |  **1.00** |    **0.07** |  **9750.0000** |  **8500.0000** |         **-** |  **474.57 MB** |        **1.00** |
| OfficeIMOCoreParallel | Conve(...)00000 [25] | 4   |   756.6 ms |  34.80 ms |  18.20 ms |  4.99 |    0.27 | 26750.0000 |  7250.0000 |         - | 1277.62 MB |        2.69 |
| OfficeIMOTextParallel | Conve(...)00000 [25] | 4   |   372.7 ms |  51.37 ms |  26.87 ms |  2.46 |    0.21 | 18250.0000 | 17000.0000 | 2750.0000 | 1022.27 MB |        2.15 |
|                       |                      |     |            |           |           |       |         |            |            |           |            |             |
| **ExcelReaderParallel**   | **NarrowInt-8000000**    | **4**   |   **153.6 ms** |  **31.98 ms** |  **16.73 ms** |  **1.01** |    **0.14** |  **5250.0000** |   **750.0000** |         **-** |  **254.58 MB** |        **1.00** |
| OfficeIMOCoreParallel | NarrowInt-8000000    | 4   | 1,427.0 ms | 144.31 ms |  75.48 ms |  9.38 |    0.99 | 39250.0000 |  3500.0000 |         - | 1859.97 MB |        7.31 |
| OfficeIMOTextParallel | NarrowInt-8000000    | 4   |   446.1 ms |  93.93 ms |  49.13 ms |  2.93 |    0.41 | 27250.0000 | 21750.0000 | 3250.0000 | 1047.82 MB |        4.12 |


## qualified-830951-long-v2-domain1-arrow

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T21:16:56.3801762+00:00; finished 2026-10-08T21:20:50.0936396+00:00. Exact context: contexts/qualified-830951-long-v2-domain1-arrow-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.ArrowConversionBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-CPEWJC : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method      | RowCount | Scenario     | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0      | Gen1      | Gen2     | Allocated | Alloc Ratio |
|------------ |--------- |------------- |----------:|----------:|----------:|------:|--------:|----------:|----------:|---------:|----------:|------------:|
| **ExcelReader** | **100000**   | **CsvAllString** | **11.420 ms** | **1.0953 ms** | **1.0246 ms** |  **1.01** |    **0.12** |  **588.2353** |  **564.7059** | **541.1765** |  **16.27 MB** |        **1.00** |
| OfficeIMO   | 100000   | CsvAllString | 38.486 ms | 2.2481 ms | 2.1029 ms |  3.39 |    0.32 | 1384.6154 | 1000.0000 | 846.1538 |  49.42 MB |        3.04 |
|             |          |              |           |           |           |       |         |           |           |          |           |             |
| **ExcelReader** | **100000**   | **CsvTyped**     |  **8.074 ms** | **0.8849 ms** | **0.8277 ms** |  **1.01** |    **0.14** |  **488.5496** |  **473.2824** | **473.2824** |   **8.13 MB** |        **1.00** |
| OfficeIMO   | 100000   | CsvTyped     | 22.283 ms | 1.7481 ms | 1.6351 ms |  2.79 |    0.34 |  657.1429 |  314.2857 | 257.1429 |  21.55 MB |        2.65 |
|             |          |              |           |           |           |       |         |           |           |          |           |             |
| **ExcelReader** | **100000**   | **XlsbTyped**    | **12.134 ms** | **0.4051 ms** | **0.3789 ms** |  **1.00** |    **0.04** |  **462.5000** |  **450.0000** | **450.0000** |   **8.19 MB** |        **1.00** |
| OfficeIMO   | 100000   | XlsbTyped    | 17.631 ms | 1.0734 ms | 1.0040 ms |  1.45 |    0.09 |  166.6667 |  100.0000 | 100.0000 |   9.09 MB |        1.11 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.ArrowInferredConversionBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-CPEWJC : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method                            | RowCount | Mean     | Error     | StdDev    | Gen0     | Gen1     | Gen2     | Allocated |
|---------------------------------- |--------- |---------:|----------:|----------:|---------:|---------:|---------:|----------:|
| ExcelReader_CsvTypedWithInference | 100000   | 6.611 ms | 0.4402 ms | 0.4118 ms | 577.9817 | 559.6330 | 559.6330 |  14.89 MB |


## qualified-830951-long-v2-domain1-encrypted

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T21:20:51.6134789+00:00; finished 2026-10-08T21:26:05.4553275+00:00. Exact context: contexts/qualified-830951-long-v2-domain1-encrypted-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.EncryptedVerifiedAsyncReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-OWHAFV : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                         | Input              | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Allocated   | Alloc Ratio |
|------------------------------- |------------------- |----------:|----------:|----------:|------:|--------:|---------:|------------:|------------:|
| **ExcelReaderVerifiedStreamAsync** | **OriginalSmallXlsx**  |  **20.81 ms** |  **0.154 ms** |  **0.121 ms** |  **1.00** |    **0.01** |        **-** |    **59.64 KB** |        **1.00** |
| OfficeIMOVerifiedStreamAsync   | OriginalSmallXlsx  |  22.27 ms |  1.384 ms |  1.081 ms |  1.07 |    0.05 |        - |   263.09 KB |        4.41 |
|                                |                    |           |           |           |       |         |          |             |             |
| **ExcelReaderVerifiedStreamAsync** | **GeneratedLargeXlsx** |  **78.91 ms** | **16.037 ms** | **12.521 ms** |  **1.02** |    **0.22** | **222.2222** | **10265.68 KB** |        **1.00** |
| OfficeIMOVerifiedStreamAsync   | GeneratedLargeXlsx | 136.27 ms |  8.793 ms |  6.865 ms |  1.77 |    0.28 | 750.0000 |  67696.8 KB |        6.59 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.EncryptedVerifiedReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-OWHAFV : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=12  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=8  

```
| Method                        | Input              | MemoryInput | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated   | Alloc Ratio |
|------------------------------ |------------------- |------------ |----------:|----------:|----------:|------:|--------:|---------:|---------:|---------:|------------:|------------:|
| **ExcelReaderVerifiedFullFields** | **OriginalSmallXlsx**  | **False**       |  **22.63 ms** |  **0.774 ms** |  **0.605 ms** |  **1.00** |    **0.04** |        **-** |        **-** |        **-** |    **45.19 KB** |        **1.00** |
| OfficeIMOVerifiedFullFields   | OriginalSmallXlsx  | False       |  21.62 ms |  1.791 ms |  1.398 ms |  0.96 |    0.06 |        - |        - |        - |   261.42 KB |        5.79 |
|                               |                    |             |           |           |           |       |         |          |          |          |             |             |
| **ExcelReaderVerifiedFullFields** | **OriginalSmallXlsx**  | **True**        |  **25.15 ms** |  **3.076 ms** |  **2.402 ms** |  **1.01** |    **0.13** |        **-** |        **-** |        **-** |    **53.61 KB** |        **1.00** |
| OfficeIMOVerifiedFullFields   | OriginalSmallXlsx  | True        |  21.29 ms |  3.246 ms |  2.535 ms |  0.85 |    0.13 |        - |        - |        - |   242.14 KB |        4.52 |
|                               |                    |             |           |           |           |       |         |          |          |          |             |             |
| **ExcelReaderVerifiedFullFields** | **GeneratedLargeXlsx** | **False**       |  **86.09 ms** |  **9.122 ms** |  **7.122 ms** |  **1.01** |    **0.12** | **181.8182** |        **-** |        **-** |  **9749.42 KB** |        **1.00** |
| OfficeIMOVerifiedFullFields   | GeneratedLargeXlsx | False       | 140.36 ms | 23.918 ms | 18.673 ms |  1.64 |    0.25 | 833.3333 | 166.6667 | 166.6667 | 68040.43 KB |        6.98 |
|                               |                    |             |           |           |           |       |         |          |          |          |             |             |
| **ExcelReaderVerifiedFullFields** | **GeneratedLargeXlsx** | **True**        |  **62.24 ms** |  **0.926 ms** |  **0.723 ms** |  **1.00** |    **0.02** | **200.0000** |        **-** |        **-** | **18773.67 KB** |        **1.00** |
| OfficeIMOVerifiedFullFields   | GeneratedLargeXlsx | True        | 134.10 ms | 10.279 ms |  8.025 ms |  2.15 |    0.13 | 777.7778 | 222.2222 | 222.2222 | 62105.49 KB |        3.31 |


## qualified-830951-long-v2-domain1-cold

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T21:26:06.8549975+00:00; finished 2026-10-08T21:27:00.6239217+00:00. Exact context: contexts/qualified-830951-long-v2-domain1-cold-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.ColdStartReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-IJNEND : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
InvocationCount=1  IterationCount=1  LaunchCount=16  
RunStrategy=ColdStart  UnrollFactor=1  WarmupCount=0  

```
| Method                       | Mean      | Error     | StdDev    | Ratio | RatioSD | Allocated | Alloc Ratio |
|----------------------------- |----------:|----------:|----------:|------:|--------:|----------:|------------:|
| ExcelReaderAttributes        |  41.25 ms |  8.912 ms |  8.753 ms |  1.05 |    0.33 | 133.02 KB |        1.00 |
| OfficeIMOAutomaticMapping    | 104.12 ms | 60.643 ms | 59.560 ms |  2.65 |    1.63 | 112.98 KB |        0.85 |
| ExcelReaderFluentMapping     |  28.77 ms |  4.096 ms |  4.023 ms |  0.73 |    0.20 | 136.05 KB |        1.02 |
| ExcelReaderAttributeFallback |  32.21 ms |  1.811 ms |  1.779 ms |  0.82 |    0.19 | 138.09 KB |        1.04 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.ColdStartWriteBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-IJNEND : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
InvocationCount=1  IterationCount=1  LaunchCount=16  
RunStrategy=ColdStart  UnrollFactor=1  WarmupCount=0  

```
| Method                     | Mean      | Error    | StdDev   | Ratio | RatioSD | Allocated | Alloc Ratio |
|--------------------------- |----------:|---------:|---------:|------:|--------:|----------:|------------:|
| ExcelReaderAutomaticLayout |  36.87 ms | 1.663 ms | 1.634 ms |  1.00 |    0.06 |  95.84 KB |        1.00 |
| OfficeIMOAutomaticLayout   | 134.49 ms | 9.483 ms | 9.313 ms |  3.65 |    0.29 | 901.95 KB |        9.41 |

