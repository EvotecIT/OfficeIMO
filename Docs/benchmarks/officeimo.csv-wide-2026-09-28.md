# CSV wide read and write snapshot — 2026-09-28

This local run refreshes the 25,000-row, 40-column CSV comparison used by the
benchmark README and website overview. The source library was at `da0d1861d`;
the benchmark harness additionally selected only the 25,000-row parameter.
Each write method was checked against every decoded output value before timing.
Each read method traversed the same data and matched its expected observation.
The compact comparison contains 14 methods; the measurement also included one
UTF-8-stream write method because the original filter matched its name.

Windows 11 build 26200.9457, AMD Ryzen 9 9950X3D2 (16 cores, 32 logical
processors), .NET 8.0.31, and BenchmarkDotNet 0.15.8 were used. Every job ran
with affinity `0xFFFF`, AboveNormal priority, the High performance power plan,
four warmups, eight measured iterations, one launch, and retained outliers.
The comparison measures in-memory text reading or writing; it is not a file-I/O
or peak-resident-memory measurement.

| Contract | OfficeIMO mean | Other measured means | OfficeIMO allocation |
| --- | ---: | --- | ---: |
| Read every field as a span | 3.744 ms | Sep 2.124 ms; Sylvan 2.742 ms | 680 B |
| Write projected object rows | 26.374 ms | CsvHelper 73.661 ms; Dataplat 29.620 ms | 18.15 MB |
| Write from `IDataReader` | 30.421 ms | Sylvan 24.919 ms; Dataplat 35.763 ms | 18.23 MB |
| Write validated text rows | 7.329 ms | Sylvan 11.514 ms; Dataplat 14.419 ms; Sep 25.466 ms; CsvHelper 26.046 ms | 18.15 MB |

The span read is 1.76x Sep's mean in this run, while its allocation is lower.
The DataReader write distribution is noisy (5.882 ms standard deviation), as
are several other lanes. These single-domain results do not establish a
portable ranking or a regression against July's differently controlled run.
The wide span-read follow-up is recorded in the
[read dispatch report](officeimo.csv-wide-read-dispatch-2026-09-28.md); the
DataReader write follow-up is below. Neither establishes a portable budget.
The shorter mixed-JSON file-write
contract is measured separately in
[the mixed-JSON report](officeimo.csv-mixed-json-allocation-2026-09-24.md).

### Rotated wide DataReader write follow-up

A separate .NET 10 run measured the same 25,000-row, 40-column validated
`IDataReader` export against Sylvan's `IDataReader` writer. Both used a fresh
reader over identical rows and produced output checked field by field before
timing. BenchmarkDotNet 0.15.8 ran on the same Windows workstation with
`AboveNormal` priority, both L3 affinity domains, eight fixed invocations,
four warmups, eight measured iterations, one launch, and all outliers retained.
The baseline and a candidate that appended directly to an exact `StringWriter`
buffer were alternated. The final candidate retained the original buffered
behavior for cancellable tokens; this benchmark uses a non-cancellable token.

| Run | `0xFFFF` OfficeIMO / Sylvan | `0xFFFF0000` OfficeIMO / Sylvan | OfficeIMO allocation |
| --- | ---: | ---: | ---: |
| Baseline | 37.71 / 27.69 ms | 36.56 / 32.82 ms | 17.39 MB |
| Direct-buffer candidate | 32.03 / 27.11 ms | 35.90 / 31.00 ms | 17.31 MB |
| Rotated baseline | 30.78 / 26.92 ms | 34.61 / 29.90 ms | 17.39 MB |
| Final candidate | 32.48 / 26.11 ms | 35.07 / 31.15 ms | 17.31 MB |

The direct-buffer candidate saved about 0.08 MB per export but did not show a
repeatable elapsed-time improvement across the rotated runs. It was removed.
A sampled CPU trace of the baseline identified decimal formatting and typed
per-field reader dispatch as the main OfficeIMO work; its percentage totals
also include untimed preflight and should not be read as isolated writer costs.
Further changes need equivalent output and a new rotated measurement across
both processor domains.
The [dated sample data](officeimo.csv-wide-write-samples-2026-09-28.json)
retains every measured iteration, including outliers, and each run's
allocation result.

The command used for each leg was:

```powershell
$env:OFFICEIMO_CSV_WIDE_ROW_COUNT = '25000'
dotnet run -c Release -f net10.0 --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -- --filter '*CsvWideBenchmarks.OfficeIMO_WriteDataReader*' '*CsvWideBenchmarks.Sylvan_WriteProjectedRows*' --affinityMasks '0xFFFF,0xFFFF0000' --priority AboveNormal --invocationCount 8 --unrollFactor 1 --warmupCount 4 --iterationCount 8 --launchCount 1 --outliers DontRemove
```

The committed compact result is
[`readme-current/officeimo.csv.comparison.json`](readme-current/officeimo.csv.comparison.json).
To rerun its selected workload with the same processor placement and iteration
policy, use:

```powershell
.\Build\Benchmarks\Update-BenchmarkReadmes.ps1 -Run Csv -CsvAffinityMasks 0xFFFF -CsvPriority AboveNormal -CsvWarmupCount 4 -CsvIterationCount 8
```

That command regenerates the compact result and CSV benchmark README from a
fresh run. Raw BenchmarkDotNet artifacts remain local and are not required for
the checked-in snapshot.
