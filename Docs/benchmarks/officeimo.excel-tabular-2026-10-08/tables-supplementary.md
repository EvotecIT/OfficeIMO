# Native grouped tables: supplementary

These are unmodified BenchmarkDotNet Markdown outputs under explicit capture/domain/class headers. Native FullName/job/parameter identities and all warnings/Actuals are in the exact raw packet and JSON/CSV ledgers. Raw API, mapper, metadata, inference and cold diagnostic-allocation boundaries remain as described in the dated report. A native Baseline ratio does not make different API work equivalent.

## qualified-e676-v1-domain0-main-current

Capture source: e676775a59458d1095606386c39506ceb8a9bc33; started 2026-10-08T12:34:35.3313114+00:00; finished 2026-10-08T12:38:10.18826+00:00. Exact context: contexts/qualified-e676-v1-domain0-main-current-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-PUTKPF : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-e67677",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
InvocationCount=32  IterationCount=15  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method           | RowCount | Shape         | Mean     | Error    | StdDev   | Ratio | RatioSD | Gen0     | Gen1    | Allocated | Alloc Ratio |
|----------------- |--------- |-------------- |---------:|---------:|---------:|------:|--------:|---------:|--------:|----------:|------------:|
| **ExcelReaderTyped** | **50000**    | **Original**      | **11.88 ms** | **0.543 ms** | **0.508 ms** |  **1.00** |    **0.06** |  **62.5000** |       **-** |   **3.87 MB** |        **1.00** |
| OfficeIMOTyped   | 50000    | Original      | 12.82 ms | 0.288 ms | 0.269 ms |  1.08 |    0.05 |  31.2500 |       - |   2.36 MB |        0.61 |
| SylvanTyped      | 50000    | Original      | 65.10 ms | 2.630 ms | 2.460 ms |  5.49 |    0.31 | 187.5000 | 31.2500 |  10.48 MB |        2.71 |
|                  |          |               |          |          |          |       |         |          |         |           |             |
| **ExcelReaderTyped** | **50000**    | **SharedStrings** | **11.70 ms** | **0.637 ms** | **0.595 ms** |  **1.00** |    **0.07** |  **31.2500** |       **-** |   **2.31 MB** |        **1.00** |
| OfficeIMOTyped   | 50000    | SharedStrings | 12.95 ms | 0.245 ms | 0.230 ms |  1.11 |    0.06 |  62.5000 |       - |   2.44 MB |        1.06 |
| SylvanTyped      | 50000    | SharedStrings | 61.29 ms | 2.157 ms | 2.017 ms |  5.25 |    0.31 | 218.7500 | 62.5000 |   8.91 MB |        3.86 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.WriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-PUTKPF : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-e67677",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
InvocationCount=32  IterationCount=15  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method            | RowCount | Mean     | Error    | StdDev   | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated | Alloc Ratio |
|------------------ |--------- |---------:|---------:|---------:|------:|--------:|---------:|---------:|---------:|----------:|------------:|
| OfficeIMO         | 50000    | 17.95 ms | 0.777 ms | 0.727 ms |  1.74 |    0.08 | 312.5000 | 312.5000 | 312.5000 |   4.04 MB |        1.01 |
| ExcelReaderWriter | 50000    | 10.32 ms | 0.233 ms | 0.218 ms |  1.00 |    0.03 | 312.5000 | 312.5000 | 312.5000 |   4.02 MB |        1.00 |


## qualified-e676-v1-domain1-main-current

Capture source: e676775a59458d1095606386c39506ceb8a9bc33; started 2026-10-08T12:38:10.5300404+00:00; finished 2026-10-08T12:41:41.9145561+00:00. Exact context: contexts/qualified-e676-v1-domain1-main-current-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-BJXTQL : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-e67677",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
InvocationCount=32  IterationCount=15  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method           | RowCount | Shape         | Mean     | Error    | StdDev   | Ratio | RatioSD | Gen0     | Gen1    | Allocated | Alloc Ratio |
|----------------- |--------- |-------------- |---------:|---------:|---------:|------:|--------:|---------:|--------:|----------:|------------:|
| **ExcelReaderTyped** | **50000**    | **Original**      | **12.94 ms** | **0.577 ms** | **0.539 ms** |  **1.00** |    **0.06** |  **93.7500** |       **-** |   **3.88 MB** |        **1.00** |
| OfficeIMOTyped   | 50000    | Original      | 12.05 ms | 0.478 ms | 0.447 ms |  0.93 |    0.05 |  62.5000 |       - |   2.44 MB |        0.63 |
| SylvanTyped      | 50000    | Original      | 71.12 ms | 2.083 ms | 1.948 ms |  5.50 |    0.27 | 250.0000 | 62.5000 |  10.48 MB |        2.70 |
|                  |          |               |          |          |          |       |         |          |         |           |             |
| **ExcelReaderTyped** | **50000**    | **SharedStrings** | **11.74 ms** | **0.438 ms** | **0.410 ms** |  **1.00** |    **0.05** |  **62.5000** |       **-** |   **2.31 MB** |        **1.00** |
| OfficeIMOTyped   | 50000    | SharedStrings | 11.83 ms | 0.780 ms | 0.729 ms |  1.01 |    0.07 |  62.5000 |       - |   2.44 MB |        1.06 |
| SylvanTyped      | 50000    | SharedStrings | 58.70 ms | 2.811 ms | 2.629 ms |  5.01 |    0.28 | 218.7500 | 62.5000 |   8.91 MB |        3.86 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.WriterBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-BJXTQL : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-e67677",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
InvocationCount=32  IterationCount=15  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method            | RowCount | Mean     | Error    | StdDev   | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated | Alloc Ratio |
|------------------ |--------- |---------:|---------:|---------:|------:|--------:|---------:|---------:|---------:|----------:|------------:|
| OfficeIMO         | 50000    | 18.61 ms | 0.527 ms | 0.493 ms |  1.86 |    0.07 | 312.5000 | 312.5000 | 312.5000 |   4.04 MB |        1.01 |
| ExcelReaderWriter | 50000    | 10.02 ms | 0.311 ms | 0.291 ms |  1.00 |    0.04 | 312.5000 | 312.5000 | 312.5000 |   4.02 MB |        1.00 |


## qualified-830951-v1-domain0-native-attribution-before

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T13:13:30.5989748+00:00; finished 2026-10-08T13:13:59.2194607+00:00. Exact context: contexts/qualified-830951-v1-domain0-native-attribution-before-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.RawReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-GAYBJO : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-e67677",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=250ms  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method              | RowCount | Format | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0    | Allocated  | Alloc Ratio |
|-------------------- |--------- |------- |----------:|----------:|----------:|------:|--------:|--------:|-----------:|------------:|
| ExcelReaderOriginal | 50000    | Xlsb   |  4.749 ms | 0.5294 ms | 0.4952 ms |  1.01 |    0.15 |       - |    4.13 KB |        1.00 |
| OfficeIMOOriginal   | 50000    | Xlsb   | 11.158 ms | 0.9913 ms | 0.9272 ms |  2.37 |    0.31 | 76.9231 | 6857.92 KB |    1,659.38 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedXlsbReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-GAYBJO : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-e67677",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=250ms  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method               | RowCount | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated | Alloc Ratio |
|--------------------- |--------- |----------:|----------:|----------:|------:|--------:|---------:|---------:|---------:|----------:|------------:|
| ExcelReaderTypedXlsb | 50000    |  9.018 ms | 0.5773 ms | 0.5400 ms |  1.00 |    0.08 | 142.8571 |        - |        - |   3.93 MB |        1.00 |
| OfficeIMOTypedXlsb   | 50000    | 12.027 ms | 1.7218 ms | 1.6106 ms |  1.34 |    0.19 | 222.2222 | 111.1111 | 111.1111 |   6.44 MB |        1.64 |


## qualified-830951-v1-domain0-native-attribution-current

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T13:13:59.9242503+00:00; finished 2026-10-08T13:14:38.96675+00:00. Exact context: contexts/qualified-830951-v1-domain0-native-attribution-current-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.RawReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-GXIQPA : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=250ms  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method              | RowCount | Format | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1    | Gen2    | Allocated  | Alloc Ratio |
|-------------------- |--------- |------- |----------:|----------:|----------:|------:|--------:|---------:|--------:|--------:|-----------:|------------:|
| ExcelReaderOriginal | 50000    | Xlsb   |  5.118 ms | 0.1467 ms | 0.1373 ms |  1.00 |    0.04 |        - |       - |       - |      12 KB |        1.00 |
| OfficeIMOOriginal   | 50000    | Xlsb   | 10.568 ms | 1.4228 ms | 1.3309 ms |  2.07 |    0.26 | 142.8571 | 71.4286 | 71.4286 | 5833.02 KB |      485.93 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedXlsbReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-GXIQPA : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=250ms  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method               | RowCount | Mean     | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1    | Gen2    | Allocated | Alloc Ratio |
|--------------------- |--------- |---------:|----------:|----------:|------:|--------:|---------:|--------:|--------:|----------:|------------:|
| ExcelReaderTypedXlsb | 50000    | 6.239 ms | 0.3178 ms | 0.2973 ms |  1.00 |    0.06 |        - |       - |       - |   3.93 MB |        1.00 |
| OfficeIMOTypedXlsb   | 50000    | 8.934 ms | 0.3909 ms | 0.3656 ms |  1.43 |    0.09 | 108.1081 | 27.0270 | 27.0270 |    4.2 MB |        1.07 |


## qualified-830951-v1-domain1-native-attribution-current

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T13:14:40.1856868+00:00; finished 2026-10-08T13:15:08.6749924+00:00. Exact context: contexts/qualified-830951-v1-domain1-native-attribution-current-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.RawReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-PEWTIJ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=250ms  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method              | RowCount | Format | Mean     | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1    | Gen2    | Allocated  | Alloc Ratio |
|-------------------- |--------- |------- |---------:|----------:|----------:|------:|--------:|---------:|--------:|--------:|-----------:|------------:|
| ExcelReaderOriginal | 50000    | Xlsb   | 5.037 ms | 0.6617 ms | 0.6190 ms |  1.01 |    0.17 |        - |       - |       - |   55.73 KB |        1.00 |
| OfficeIMOOriginal   | 50000    | Xlsb   | 9.274 ms | 0.6803 ms | 0.6363 ms |  1.87 |    0.25 | 153.8462 | 76.9231 | 76.9231 | 5878.96 KB |      105.49 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedXlsbReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-PEWTIJ : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-830951",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=250ms  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method               | RowCount | Mean     | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1    | Gen2    | Allocated | Alloc Ratio |
|--------------------- |--------- |---------:|----------:|----------:|------:|--------:|---------:|--------:|--------:|----------:|------------:|
| ExcelReaderTypedXlsb | 50000    | 6.199 ms | 0.3541 ms | 0.3312 ms |  1.00 |    0.07 | 142.8571 |       - |       - |   3.93 MB |        1.00 |
| OfficeIMOTypedXlsb   | 50000    | 8.874 ms | 0.8772 ms | 0.8206 ms |  1.44 |    0.15 | 153.8462 | 76.9231 | 76.9231 |    4.6 MB |        1.17 |


## qualified-830951-v1-domain1-native-attribution-before

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T13:15:09.0487717+00:00; finished 2026-10-08T13:15:46.6418282+00:00. Exact context: contexts/qualified-830951-v1-domain1-native-attribution-before-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.RawReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-XXLIFY : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-e67677",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=250ms  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method              | RowCount | Format | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1     | Gen2     | Allocated  | Alloc Ratio |
|-------------------- |--------- |------- |----------:|----------:|----------:|------:|--------:|---------:|---------:|---------:|-----------:|------------:|
| ExcelReaderOriginal | 50000    | Xlsb   |  4.772 ms | 0.6027 ms | 0.5638 ms |  1.01 |    0.15 |        - |        - |        - |   11.87 KB |        1.00 |
| OfficeIMOOriginal   | 50000    | Xlsb   | 12.040 ms | 0.9480 ms | 0.8868 ms |  2.55 |    0.31 | 200.0000 | 100.0000 | 100.0000 | 7679.83 KB |      646.83 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedXlsbReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-XXLIFY : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-e67677",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=250ms  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method               | RowCount | Mean     | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1    | Gen2    | Allocated | Alloc Ratio |
|--------------------- |--------- |---------:|----------:|----------:|------:|--------:|---------:|--------:|--------:|----------:|------------:|
| ExcelReaderTypedXlsb | 50000    | 6.831 ms | 0.7839 ms | 0.7333 ms |  1.01 |    0.14 | 142.8571 |       - |       - |   3.93 MB |        1.00 |
| OfficeIMOTypedXlsb   | 50000    | 8.628 ms | 1.0122 ms | 0.9468 ms |  1.28 |    0.18 | 214.2857 | 71.4286 | 71.4286 |   6.13 MB |        1.56 |


## qualified-830951-long-v1-domain0-native-attribution-before

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T13:25:56.7966508+00:00; finished 2026-10-08T13:28:34.7963939+00:00. Exact context: contexts/qualified-830951-long-v1-domain0-native-attribution-before-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.RawReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-UXXZMP : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-e67677",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method              | RowCount | Format | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1    | Gen2    | Allocated  | Alloc Ratio |
|-------------------- |--------- |------- |----------:|----------:|----------:|------:|--------:|---------:|--------:|--------:|-----------:|------------:|
| ExcelReaderOriginal | 50000    | Xlsb   |  5.593 ms | 0.0917 ms | 0.0857 ms |  1.00 |    0.02 |        - |       - |       - |    6.82 KB |        1.00 |
| OfficeIMOOriginal   | 50000    | Xlsb   | 11.561 ms | 0.3811 ms | 0.3565 ms |  2.07 |    0.07 | 149.4253 | 11.4943 | 11.4943 | 6952.34 KB |    1,019.80 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedXlsbReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-UXXZMP : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=00000000000000001111111111111111  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-e67677",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method               | RowCount | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1    | Gen2    | Allocated | Alloc Ratio |
|--------------------- |--------- |----------:|----------:|----------:|------:|--------:|---------:|--------:|--------:|----------:|------------:|
| ExcelReaderTypedXlsb | 50000    |  8.642 ms | 0.2695 ms | 0.2521 ms |  1.00 |    0.04 |  75.6303 |       - |       - |   3.87 MB |        1.00 |
| OfficeIMOTypedXlsb   | 50000    | 10.613 ms | 0.4216 ms | 0.3943 ms |  1.23 |    0.06 | 123.7113 | 10.3093 | 10.3093 |   5.64 MB |        1.46 |


## qualified-830951-long-v1-domain0-native-attribution-current

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T13:28:35.6885308+00:00; finished 2026-10-08T13:31:16.6997464+00:00. Exact context: contexts/qualified-830951-long-v1-domain0-native-attribution-current-context.json.

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
| Method              | RowCount | Format | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1   | Gen2   | Allocated  | Alloc Ratio |
|-------------------- |--------- |------- |----------:|----------:|----------:|------:|--------:|---------:|-------:|-------:|-----------:|------------:|
| ExcelReaderOriginal | 50000    | Xlsb   |  5.411 ms | 0.1430 ms | 0.1338 ms |  1.00 |    0.03 |        - |      - |      - |    6.64 KB |        1.00 |
| OfficeIMOOriginal   | 50000    | Xlsb   | 10.354 ms | 1.3850 ms | 1.2956 ms |  1.91 |    0.24 | 118.8119 | 9.9010 | 9.9010 | 5327.84 KB |      801.96 |


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
| Method               | RowCount | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0    | Gen1   | Gen2   | Allocated | Alloc Ratio |
|--------------------- |--------- |----------:|----------:|----------:|------:|--------:|--------:|-------:|-------:|----------:|------------:|
| ExcelReaderTypedXlsb | 50000    |  8.573 ms | 0.8398 ms | 0.7855 ms |  1.01 |    0.13 | 84.9057 |      - |      - |   3.87 MB |        1.00 |
| OfficeIMOTypedXlsb   | 50000    | 10.689 ms | 0.5030 ms | 0.4705 ms |  1.26 |    0.13 | 96.1538 | 9.6154 | 9.6154 |   4.06 MB |        1.05 |


## qualified-830951-long-v1-domain1-native-attribution-current

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T13:31:21.1389815+00:00; finished 2026-10-08T13:33:50.0879758+00:00. Exact context: contexts/qualified-830951-long-v1-domain1-native-attribution-current-context.json.

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
| Method              | RowCount | Format | Mean     | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1   | Gen2   | Allocated  | Alloc Ratio |
|-------------------- |--------- |------- |---------:|----------:|----------:|------:|--------:|---------:|-------:|-------:|-----------:|------------:|
| ExcelReaderOriginal | 50000    | Xlsb   | 4.469 ms | 0.4252 ms | 0.3977 ms |  1.01 |    0.12 |        - |      - |      - |    6.01 KB |        1.00 |
| OfficeIMOOriginal   | 50000    | Xlsb   | 7.368 ms | 0.2989 ms | 0.2796 ms |  1.66 |    0.14 | 114.7541 | 8.1967 | 8.1967 | 5313.83 KB |      883.77 |


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
| Method               | RowCount | Mean     | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1    | Gen2   | Allocated | Alloc Ratio |
|--------------------- |--------- |---------:|----------:|----------:|------:|--------:|---------:|--------:|-------:|----------:|------------:|
| ExcelReaderTypedXlsb | 50000    | 7.151 ms | 1.4734 ms | 1.3782 ms |  1.03 |    0.26 |  80.7453 |       - |      - |   3.87 MB |        1.00 |
| OfficeIMOTypedXlsb   | 50000    | 7.142 ms | 0.5217 ms | 0.4880 ms |  1.03 |    0.19 | 111.1111 | 41.6667 | 6.9444 |   4.04 MB |        1.04 |


## qualified-830951-long-v1-domain1-native-attribution-before

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T13:33:52.3274394+00:00; finished 2026-10-08T13:36:42.3766949+00:00. Exact context: contexts/qualified-830951-long-v1-domain1-native-attribution-before-context.json.

### OfficeIMO.Excel.ReaderComparison.Benchmarks.RawReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-JONANS : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-e67677",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method              | RowCount | Format | Mean     | Error     | StdDev    | Ratio | RatioSD | Gen0     | Gen1   | Gen2   | Allocated  | Alloc Ratio |
|-------------------- |--------- |------- |---------:|----------:|----------:|------:|--------:|---------:|-------:|-------:|-----------:|------------:|
| ExcelReaderOriginal | 50000    | Xlsb   | 4.056 ms | 0.1420 ms | 0.1328 ms |  1.00 |    0.04 |        - |      - |      - |    6.04 KB |        1.00 |
| OfficeIMOOriginal   | 50000    | Xlsb   | 8.269 ms | 1.1757 ms | 1.0998 ms |  2.04 |    0.27 | 148.9362 | 7.0922 | 7.0922 | 6916.28 KB |    1,145.81 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.TypedXlsbReadBenchmarks-report-github.md

```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26300.9457)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.112
  [Host]     : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4
  Job-JONANS : .NET 10.0.12 (10.0.12, 10.0.1226.42308), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  Affinity=11111111111111110000000000000000  Arguments=/p:UseSharedCompilation=false,/p:OfficeIMOBenchmarkAssemblyDirectory="D:\Dev\Scratch\officeimo-excelreader-performance\current-owner-bundle-e67677",/p:OfficeIMOBenchmarkNewApis=true,/p:OfficeIMOBenchmarkCsv=true,/p:OfficeIMOBenchmarkArrow=true  
IterationCount=15  IterationTime=1s  LaunchCount=1  
UnrollFactor=1  WarmupCount=12  

```
| Method               | RowCount | Mean      | Error    | StdDev   | Ratio | RatioSD | Gen0     | Gen1    | Gen2    | Allocated | Alloc Ratio |
|--------------------- |--------- |----------:|---------:|---------:|------:|--------:|---------:|--------:|--------:|----------:|------------:|
| ExcelReaderTypedXlsb | 50000    |  7.975 ms | 1.100 ms | 1.029 ms |  1.02 |    0.19 |  83.9161 |       - |       - |   3.87 MB |        1.00 |
| OfficeIMOTypedXlsb   | 50000    | 10.281 ms | 1.361 ms | 1.273 ms |  1.31 |    0.24 | 162.7907 | 69.7674 | 11.6279 |   5.65 MB |        1.46 |


## qualified-830951-long-v1-domain0-recipe-probe

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T13:36:45.143763+00:00; finished 2026-10-08T13:45:08.908837+00:00. Exact context: contexts/qualified-830951-long-v1-domain0-recipe-probe-context.json.

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
| **ExcelReader** | **50000**    | **GetValue**     | **11.961 ms** | **0.2544 ms** | **0.2379 ms** |  **1.00** |    **0.03** | **107.1429** |       **-** |       **-** | **5137.03 KB** |        **1.00** |
| OfficeIMO   | 50000    | GetValue     | 18.177 ms | 0.3851 ms | 0.3602 ms |  1.52 |    0.04 | 127.2727 | 54.5455 | 36.3636 | 4487.26 KB |        0.87 |
| Sylvan      | 50000    | GetValue     | 61.266 ms | 2.2958 ms | 2.1475 ms |  5.12 |    0.20 | 187.5000 |       - |       - | 8382.99 KB |        1.63 |
|             |          |              |           |           |           |       |         |          |         |         |            |             |
| **ExcelReader** | **50000**    | **TypedGetters** | **12.381 ms** | **0.5104 ms** | **0.4774 ms** |  **1.00** |    **0.05** |  **36.5854** |       **-** |       **-** | **1621.54 KB** |        **1.00** |
| OfficeIMO   | 50000    | TypedGetters | 13.956 ms | 0.5484 ms | 0.5130 ms |  1.13 |    0.06 |        - |       - |       - | 1014.61 KB |        0.63 |
| Sylvan      | 50000    | TypedGetters | 45.093 ms | 2.5187 ms | 2.3560 ms |  3.65 |    0.23 |  47.6190 |       - |       - | 1940.19 KB |        1.20 |
|             |          |              |           |           |           |       |         |          |         |         |            |             |
| **ExcelReader** | **50000**    | **Utf8TextCopy** |  **8.444 ms** | **0.2315 ms** | **0.2166 ms** |  **1.00** |    **0.04** |        **-** |       **-** |       **-** |    **8.77 KB** |        **1.00** |
| OfficeIMO   | 50000    | Utf8TextCopy | 10.706 ms | 0.3860 ms | 0.3611 ms |  1.27 |    0.05 |        - |       - |       - |  959.91 KB |      109.41 |
| Sylvan      | 50000    | Utf8TextCopy | 40.231 ms | 1.8774 ms | 1.7561 ms |  4.77 |    0.24 |  74.0741 |       - |       - | 4283.37 KB |      488.22 |


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
| ExcelReaderOriginal   | 50000    | 2.760 ms | 0.0703 ms | 0.0658 ms |  1.00 |    0.03 |     848 B |        1.00 |
| OfficeIMOBorrowedUtf8 | 50000    | 7.777 ms | 0.2454 ms | 0.2296 ms |  2.82 |    0.10 |    3970 B |        4.68 |
| SylvanOriginal        | 50000    | 4.460 ms | 0.1085 ms | 0.1014 ms |  1.62 |    0.05 |   38681 B |       45.61 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.SharedStringFirstRowBenchmarks-report-github.md

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
| Method                         | RowCount | Storage | Mean         | Error        | StdDev       | Ratio | RatioSD | Gen0     | Gen1     | Gen2    | Allocated  | Alloc Ratio |
|------------------------------- |--------- |-------- |-------------:|-------------:|-------------:|------:|--------:|---------:|---------:|--------:|-----------:|------------:|
| ExcelReaderOpenThroughFirstRow | 65536    | Stored  |  4,312.61 μs |    70.781 μs |    66.208 μs |  1.00 |    0.02 |        - |        - |       - |  746.25 KB |        1.00 |
| OfficeIMOOpenThroughFirstRow   | 65536    | Stored  | 60,558.08 μs | 1,977.606 μs | 1,849.854 μs | 14.05 |    0.47 | 294.1176 | 235.2941 | 58.8235 | 14956.6 KB |       20.04 |
| SylvanOpenThroughFirstRow      | 65536    | Stored  |     54.79 μs |     5.292 μs |     4.950 μs |  0.01 |    0.00 |   4.6527 |   1.6421 |  0.7663 |  320.08 KB |        0.43 |


## qualified-830951-long-v1-domain1-recipe-probe

Capture source: 83095146282a96143b23a79a7ecca058a24a709b; started 2026-10-08T13:45:10.709728+00:00; finished 2026-10-08T13:53:16.8325155+00:00. Exact context: contexts/qualified-830951-long-v1-domain1-recipe-probe-context.json.

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
| Method      | RowCount | Access       | Mean      | Error     | StdDev    | Ratio | RatioSD | Gen0     | Allocated  | Alloc Ratio |
|------------ |--------- |------------- |----------:|----------:|----------:|------:|--------:|---------:|-----------:|------------:|
| **ExcelReader** | **50000**    | **GetValue**     |  **7.713 ms** | **0.3863 ms** | **0.3614 ms** |  **1.00** |    **0.06** | **100.0000** | **5131.51 KB** |        **1.00** |
| OfficeIMO   | 50000    | GetValue     | 10.877 ms | 1.5583 ms | 1.4576 ms |  1.41 |    0.19 |        - |  4348.5 KB |        0.85 |
| Sylvan      | 50000    | GetValue     | 40.767 ms | 6.7315 ms | 6.2966 ms |  5.30 |    0.82 | 137.9310 | 8382.24 KB |        1.63 |
|             |          |              |           |           |           |       |         |          |            |             |
| **ExcelReader** | **50000**    | **TypedGetters** | **11.934 ms** | **0.2864 ms** | **0.2679 ms** |  **1.00** |    **0.03** |  **36.1446** | **1621.47 KB** |        **1.00** |
| OfficeIMO   | 50000    | TypedGetters | 11.698 ms | 0.7688 ms | 0.7191 ms |  0.98 |    0.06 |        - |  1002.5 KB |        0.62 |
| Sylvan      | 50000    | TypedGetters | 37.534 ms | 6.8721 ms | 6.4281 ms |  3.15 |    0.53 |  35.7143 | 1940.05 KB |        1.20 |
|             |          |              |           |           |           |       |         |          |            |             |
| **ExcelReader** | **50000**    | **Utf8TextCopy** |  **6.397 ms** | **0.9466 ms** | **0.8854 ms** |  **1.02** |    **0.19** |        **-** |    **8.45 KB** |        **1.00** |
| OfficeIMO   | 50000    | Utf8TextCopy |  9.256 ms | 1.4896 ms | 1.3934 ms |  1.47 |    0.29 |        - |  974.04 KB |      115.21 |
| Sylvan      | 50000    | Utf8TextCopy | 41.735 ms | 2.1267 ms | 1.9893 ms |  6.64 |    0.91 |  78.9474 | 4283.68 KB |      506.70 |


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
| ExcelReaderOriginal   | 50000    | 2.930 ms | 0.0957 ms | 0.0895 ms |  1.00 |    0.04 |   1.02 KB |        1.00 |
| OfficeIMOBorrowedUtf8 | 50000    | 6.752 ms | 0.6972 ms | 0.6522 ms |  2.31 |    0.23 |   2.07 KB |        2.02 |
| SylvanOriginal        | 50000    | 4.054 ms | 0.1681 ms | 0.1572 ms |  1.38 |    0.07 |  37.77 KB |       36.94 |


### OfficeIMO.Excel.ReaderComparison.Benchmarks.SharedStringFirstRowBenchmarks-report-github.md

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
| Method                         | RowCount | Storage | Mean         | Error        | StdDev       | Ratio | RatioSD | Gen0     | Gen1     | Gen2    | Allocated   | Alloc Ratio |
|------------------------------- |--------- |-------- |-------------:|-------------:|-------------:|------:|--------:|---------:|---------:|--------:|------------:|------------:|
| ExcelReaderOpenThroughFirstRow | 65536    | Stored  |  4,122.34 μs |   366.796 μs |   343.101 μs |  1.01 |    0.12 |        - |        - |       - |   784.69 KB |        1.00 |
| OfficeIMOOpenThroughFirstRow   | 65536    | Stored  | 70,871.95 μs | 2,357.920 μs | 2,205.599 μs | 17.31 |    1.56 | 384.6154 | 307.6923 | 76.9231 | 18070.33 KB |       23.03 |
| SylvanOpenThroughFirstRow      | 65536    | Stored  |     72.69 μs |     2.773 μs |     2.594 μs |  0.02 |    0.00 |   5.9216 |   2.7286 |  1.9158 |   320.09 KB |        0.41 |

