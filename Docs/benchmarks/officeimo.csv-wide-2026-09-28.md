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
They identify wide span reading and DataReader writing for rotated, per-domain
follow-up before a performance budget is set. The shorter mixed-JSON file-write
contract is measured separately in
[the mixed-JSON report](officeimo.csv-mixed-json-allocation-2026-09-24.md).

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
