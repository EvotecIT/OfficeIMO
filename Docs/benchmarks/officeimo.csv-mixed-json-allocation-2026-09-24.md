# CSV mixed JSON write allocation — 2026-09-24

The multi-character-delimiter writer now formats common typed fields directly
into the record buffer. It retains the existing string fallback for custom
formatters and long formatted values. The benchmark's complete UTF-8 output
and every decoded field matched the comparison writer before each measurement.

`CsvFileWriteBenchmarks.MixedJson25K` writes 25,000 rows with short JSON,
Unicode text, numbers, UTC dates, booleans, nulls, and a `||` delimiter. It
measures full file creation and disposal in `AsNeeded` and `Always` quote modes.
The source baseline was commit `2180dbbb34d13d83cc54bdffef757c73caa4cc60`.

All runs used Windows 11 on an AMD Ryzen 9 9950X3D2, .NET 10.0.12,
BenchmarkDotNet 0.15.8, 32 invocations per iteration, eight warmups, 16
measured iterations, one launch, and retained outliers. The same machine and
settings were used for a baseline, a candidate, a rotated baseline control,
and a final candidate run.

| Run | AsNeeded mean | AsNeeded allocation | Always mean | Always allocation |
| --- | ---: | ---: | ---: | ---: |
| Baseline | 11.596 ms | 3.99 MB | 9.728 ms | 3.99 MB |
| First candidate | 9.941 ms | 387.4 KB | 6.547 ms | 387.5 KB |
| Rotated baseline | 7.362 ms | 3.99 MB | 6.280 ms | 3.99 MB |
| Final candidate | 7.294 ms | 387.5 KB | 6.781 ms | 387.4 KB |

Allocation fell by about 90% in both candidate runs. Timing varied across the
rotated controls, so these runs do not establish a repeatable elapsed-time
improvement. The broader short-row consistency and Linux/macOS budgets remain
open in the [roadmap](../ROADMAP.md).

To reproduce the lane from a Release build:

```powershell
dotnet run -c Release -f net10.0 --project OfficeIMO.CSV.Benchmarks -- --filter '*CsvFileWriteBenchmarks*MixedJson25K*' --invocationCount 32 --unrollFactor 1 --warmupCount 8 --iterationCount 16 --launchCount 1 --outliers DontRemove
```
