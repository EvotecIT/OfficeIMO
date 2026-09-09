```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26200.9168/25H2/2025Update/HudsonValley2)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.111
  [Host]              : .NET 10.0.11 (10.0.11, 10.0.1126.37416), X64 RyuJIT x86-64-v4
  Affinity-0xFFFF     : .NET 10.0.11 (10.0.11, 10.0.1126.37416), X64 RyuJIT x86-64-v4
  Affinity-0xFFFF0000 : .NET 10.0.11 (10.0.11, 10.0.1126.37416), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  EnvironmentVariables=OFFICEIMO_BENCHMARK_PROCESS_PRIORITY=Normal  InvocationCount=4
IterationCount=16  LaunchCount=1  UnrollFactor=1
WarmupCount=8

```
| Method         | Job                 | Affinity                         | Mean      | Error     | StdDev    | Rank | Allocated |
|--------------- |-------------------- |--------------------------------- |----------:|----------:|----------:|-----:|----------:|
| OfficeIMO      | Affinity-0xFFFF     | 00000000000000001111111111111111 |  74.82 ms |  3.895 ms |  3.826 ms |    1 | 179.33 KB |
| ExcelReaderNet | Affinity-0xFFFF     | 00000000000000001111111111111111 |  83.97 ms | 22.949 ms | 22.539 ms |    1 |  31.39 KB |
| Sylvan         | Affinity-0xFFFF     | 00000000000000001111111111111111 | 158.30 ms |  7.393 ms |  7.261 ms |    2 | 648.96 KB |
| OfficeIMO      | Affinity-0xFFFF0000 | 11111111111111110000000000000000 |  77.85 ms |  4.896 ms |  4.809 ms |    1 | 179.33 KB |
| ExcelReaderNet | Affinity-0xFFFF0000 | 11111111111111110000000000000000 | 101.23 ms | 27.932 ms | 27.433 ms |    1 |  31.39 KB |
| Sylvan         | Affinity-0xFFFF0000 | 11111111111111110000000000000000 | 191.57 ms | 10.544 ms | 10.356 ms |    3 | 648.96 KB |
