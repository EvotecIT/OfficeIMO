```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26200.9168/25H2/2025Update/HudsonValley2)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.111
  [Host]              : .NET 10.0.11 (10.0.11, 10.0.1126.37416), X64 RyuJIT x86-64-v4
  Affinity-0xFFFF     : .NET 10.0.11 (10.0.11, 10.0.1126.37416), X64 RyuJIT x86-64-v4
  Affinity-0xFFFF0000 : .NET 10.0.11 (10.0.11, 10.0.1126.37416), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  EnvironmentVariables=OFFICEIMO_BENCHMARK_PROCESS_PRIORITY=Normal  InvocationCount=8
IterationCount=16  LaunchCount=1  UnrollFactor=1
WarmupCount=8

```
| Method         | Job                 | Affinity                         | Mean     | Error    | StdDev   | Op/s  | Rank | Gen0     | Allocated |
|--------------- |-------------------- |--------------------------------- |---------:|---------:|---------:|------:|-----:|---------:|----------:|
| OfficeIMO      | Affinity-0xFFFF     | 00000000000000001111111111111111 | 16.22 ms | 1.088 ms | 1.068 ms | 61.66 |    1 | 625.0000 |  34.21 MB |
| ExcelReaderNet | Affinity-0xFFFF     | 00000000000000001111111111111111 | 22.97 ms | 2.036 ms | 2.000 ms | 43.54 |    1 | 625.0000 |  35.71 MB |
| Sep            | Affinity-0xFFFF     | 00000000000000001111111111111111 | 17.84 ms | 2.542 ms | 2.496 ms | 56.07 |    1 | 625.0000 |  34.22 MB |
| Sylvan         | Affinity-0xFFFF     | 00000000000000001111111111111111 | 19.19 ms | 2.016 ms | 1.980 ms | 52.12 |    1 | 625.0000 |  35.75 MB |
| OfficeIMO      | Affinity-0xFFFF0000 | 11111111111111110000000000000000 | 15.72 ms | 0.740 ms | 0.727 ms | 63.60 |    1 | 625.0000 |  34.21 MB |
| ExcelReaderNet | Affinity-0xFFFF0000 | 11111111111111110000000000000000 | 21.41 ms | 1.156 ms | 1.135 ms | 46.70 |    1 | 625.0000 |  35.71 MB |
| Sep            | Affinity-0xFFFF0000 | 11111111111111110000000000000000 | 14.81 ms | 1.792 ms | 1.760 ms | 67.52 |    1 | 625.0000 |  34.22 MB |
| Sylvan         | Affinity-0xFFFF0000 | 11111111111111110000000000000000 | 17.61 ms | 1.664 ms | 1.635 ms | 56.79 |    1 | 625.0000 |  35.75 MB |
