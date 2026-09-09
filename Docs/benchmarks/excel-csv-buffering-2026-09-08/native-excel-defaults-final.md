```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26200.9168/25H2/2025Update/HudsonValley2)
AMD Ryzen 9 9950X3D2 4.30GHz, 1 CPU, 32 logical and 16 physical cores
.NET SDK 10.0.111
  [Host]              : .NET 10.0.11 (10.0.11, 10.0.1126.37416), X64 RyuJIT x86-64-v4
  Affinity-0xFFFF     : .NET 10.0.11 (10.0.11, 10.0.1126.37416), X64 RyuJIT x86-64-v4
  Affinity-0xFFFF0000 : .NET 10.0.11 (10.0.11, 10.0.1126.37416), X64 RyuJIT x86-64-v4

OutlierMode=DontRemove  EnvironmentVariables=OFFICEIMO_BENCHMARK_PROCESS_PRIORITY=Normal  InvocationCount=16
IterationCount=16  LaunchCount=1  UnrollFactor=1
WarmupCount=8

```
| Method    | Job                 | Affinity                         | TextLength | TextShape | Mean       | Error       | StdDev      | Allocated |
|---------- |-------------------- |--------------------------------- |----------- |---------- |-----------:|------------:|------------:|----------:|
| **OfficeIMO** | **Affinity-0xFFFF**     | **00000000000000001111111111111111** | **64**         | **Escaped**   |   **992.4 μs** |    **83.32 μs** |    **81.83 μs** | **389.75 KB** |
| OfficeIMO | Affinity-0xFFFF0000 | 11111111111111110000000000000000 | 64         | Escaped   |   999.6 μs |   259.54 μs |   254.90 μs | 389.75 KB |
| **OfficeIMO** | **Affinity-0xFFFF**     | **00000000000000001111111111111111** | **64**         | **Markup**    | **2,111.0 μs** |   **240.71 μs** |   **236.41 μs** | **389.05 KB** |
| OfficeIMO | Affinity-0xFFFF0000 | 11111111111111110000000000000000 | 64         | Markup    |   885.0 μs |   283.10 μs |   278.04 μs |  389.8 KB |
| **OfficeIMO** | **Affinity-0xFFFF**     | **00000000000000001111111111111111** | **64**         | **Plain**     |   **591.6 μs** |   **107.55 μs** |   **105.63 μs** | **389.52 KB** |
| OfficeIMO | Affinity-0xFFFF0000 | 11111111111111110000000000000000 | 64         | Plain     |   640.6 μs |   156.59 μs |   153.79 μs | 389.52 KB |
| **OfficeIMO** | **Affinity-0xFFFF**     | **00000000000000001111111111111111** | **4096**       | **Escaped**   | **3,474.8 μs** |   **541.01 μs** |   **531.34 μs** | **630.12 KB** |
| OfficeIMO | Affinity-0xFFFF0000 | 11111111111111110000000000000000 | 4096       | Escaped   | 3,618.2 μs |   473.57 μs |   465.11 μs | 630.12 KB |
| **OfficeIMO** | **Affinity-0xFFFF**     | **00000000000000001111111111111111** | **4096**       | **Markup**    | **7,485.9 μs** |   **578.91 μs** |   **568.56 μs** | **950.04 KB** |
| OfficeIMO | Affinity-0xFFFF0000 | 11111111111111110000000000000000 | 4096       | Markup    | 8,469.6 μs | 1,282.59 μs | 1,259.67 μs | 950.04 KB |
| **OfficeIMO** | **Affinity-0xFFFF**     | **00000000000000001111111111111111** | **4096**       | **Plain**     | **3,160.4 μs** |   **721.63 μs** |   **708.74 μs** | **630.22 KB** |
| OfficeIMO | Affinity-0xFFFF0000 | 11111111111111110000000000000000 | 4096       | Plain     | 2,809.8 μs |   333.18 μs |   327.23 μs | 630.22 KB |
