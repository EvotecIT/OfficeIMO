# OfficeIMO CSV Benchmarks

This project compares raw .NET CSV paths without PowerShell object overhead. Use it beside the PSWriteOffice benchmark scoreboard, not as a replacement for it.

## Historical generated workstation snapshot

This single-workstation table is retained so the older focused investigations
remain reproducible. It is not the current cross-platform product ranking.
Lower is faster within a row only; the rows use different contracts and cannot
be combined into one library ranking. Treat differences below 5% as ties. The
snapshot uses three warmups, nine measured iterations, means, and semantic
preflight validation of every typed or prepared value.

Use the hash-pinned library-comparison suite and website matrix below for
current evidence. They keep CSV, XLSX, and XLSB workloads separate and expose
Windows, Linux, and macOS results independently.

<!-- officeimo-csv-benchmark-table:start -->
| Scenario | Variables | Host | Operation | Metric | OfficeIMO.CSV | CsvHelper | Dataplat.Dbatools.Csv | Sep | Sylvan.Data.Csv | Result |
| --- | --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: | --- |
| Wide DataReader CSV write | Contract=IDataReader, Format=CSV, Rows=25,000, Runner=BenchmarkDotNet local, Shape=wide, Snapshot=2026-07-14 | .NET 8 | Format and write rows | MeanMs | 1.00x (27ms) | n/a | 1.74x (47ms) | n/a | 0.99x (26ms) | OfficeIMO.CSV tied with Sylvan.Data.Csv |
| Wide field-span CSV read | Contract=field spans, Format=CSV, Rows=25,000, Runner=BenchmarkDotNet local, Shape=wide, Snapshot=2026-07-14 | .NET 8 | Read every field | MeanMs | 1.00x (2ms) | n/a | n/a | 1.06x (2ms) | 4.47x (9ms) | OfficeIMO.CSV fastest |
| Wide projected-array CSV write | Contract=projected object arrays, Format=CSV, Rows=25,000, Runner=BenchmarkDotNet local, Shape=wide, Snapshot=2026-07-14 | .NET 8 | Format and write rows | MeanMs | 1.00x (31ms) | 2.65x (82ms) | 1.43x (45ms) | n/a | n/a | OfficeIMO.CSV fastest |
| Wide validated text-row CSV write | Contract=preformatted text with escaping, Format=CSV, Rows=25,000, Runner=BenchmarkDotNet local, Shape=wide, Snapshot=2026-07-14 | .NET 8 | Validate and write rows | MeanMs | 1.00x (17ms) | 1.33x (23ms) | 1.25x (21ms) | 1.20x (20ms) | 0.99x (17ms) | OfficeIMO.CSV tied with Sylvan.Data.Csv |
<!-- officeimo-csv-benchmark-table:end -->

## Run

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net8.0 -- --filter *Csv*Benchmarks*
```

Refresh the compact benchmark comparison with one command:

```powershell
.\Build\Benchmarks\Update-BenchmarkReadmes.ps1 -Run Csv
```

The script runs only the focused equivalent lanes, calls PSPublishModule's
`Import-BenchmarkResult` and `Update-BenchmarkDocument`, and replaces the
marker-delimited table. Benchmark execution is local and is not scheduled in CI.

For a write-focused competitor pass:

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net8.0 -- --filter "*Write*" --job short --warmupCount 1 --iterationCount 3
```

For typed DTO reads, keep the materialized document path, the forward-only
`OpenDataReader(...).RowsAs<T>()` path, and CsvHelper's typed reader in the same
run. The document lane intentionally remains visible even when it is slower:

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net10.0 -- --filter "*ReadTyped*"
```

For database-shaped `IDataReader` exports with ordinary, quoted, multiline,
and nullable object-typed values:

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net8.0 -- --filter "*CsvDataReaderWriteBenchmarks*" --job short --warmupCount 3 --iterationCount 9
```

This 25,000-row, 10-column lane complements the numeric-heavy 40-column
headline. OfficeIMO and Sylvan consume fresh readers over the same typed rows,
and an untimed preflight parses every output and checks every value.

Stable Apple M4 snapshot from the command above on 2026-07-14:

| Shape | OfficeIMO.CSV | Sylvan.Data.Csv | Managed allocation |
| --- | ---: | ---: | ---: |
| Mixed values | 2.511 ms | 3.782 ms | 5.51 MB / 5.45 MB |
| Quoted values | 3.784 ms | 4.832 ms | 7.09 MB / 7.03 MB |
| Multiline values | 2.686 ms | 4.378 ms | 6.44 MB / 6.36 MB |

These results use explicit field types, including nullable object-typed text
columns, rather than inferring a convenient schema from the first row.

The suite compares OfficeIMO.CSV object writing, OfficeIMO.CSV projected-row writing, OfficeIMO.CSV trusted text-row writing, OfficeIMO.CSV direct IDataReader writing, OfficeIMO.CSV reusable reads, OfficeIMO.CSV field-span reads, OfficeIMO.CSV in-memory and streaming DataTable materialization with string and inferred-schema columns, OfficeIMO.CSV direct DbDataReader consumption and DbDataReader-to-DataTable loading, CsvHelper typed/projected writes, CsvHelper raw/typed reads, Sylvan raw/string/span field reads and DataTable loading, Dataplat.Dbatools.Csv reader/writer/DataTable paths, and Sep strict reader/writer paths.

Read lanes intentionally touch each field value and return a contract-appropriate checksum. Raw and DataTable lanes checksum field count and text length; typed lanes checksum every projected property and are preflighted against the generated source rows. DataTable lanes materialize the table and then traverse the cells, direct DbDataReader lanes traverse the public reader contract without first materializing a DataTable, and DbDataReader-to-DataTable lanes keep the ADO.NET table-loading path visible. This keeps the comparison honest: a lane cannot win by only counting rows or skipping the field payload.

## Sep feature-parity lanes

Sep's published comparisons separate raw row scans, decoded-column access,
typed object materialization, trimming, unescaping, and ordered parallel
enumeration. OfficeIMO keeps those contracts separate too:

- `CsvBenchmarks` covers validated plain, quoted, escaped, and multiline reads
  and writes. Sep is configured with `Unescape = true` and `Escape = true`
  where decoded values or valid CSV output are required.
- `CsvTrimUnescapeBenchmarks` compares outer ASCII-space trimming plus quote
  unescaping through decoded-string APIs in OfficeIMO, Sep, and CsvHelper.
  `SepStrings` materializes every decoded field so it can be compared with
  OfficeIMO's public `DbDataReader` and CsvHelper's parser without hiding a
  contract difference.
- `CsvTrimUnescapeSpanBenchmarks` separately compares OfficeIMO's internal
  transient-visitor engine with Sep's span reader. It is an engine diagnostic,
  not a public-API ranking. Span and string methods intentionally live in
  different benchmark types, so BenchmarkDotNet never publishes a misleading
  rank or ratio between zero-copy and materializing APIs. Both setups verify
  the full decoded field count, character count, and deterministic checksum
  before timing.
- `CsvTypedSequentialBenchmarks` compares equivalent explicit typed
  materialization through OfficeIMO's `DbDataReader` getters and Sep's
  sequential typed lambda.
- `CsvAutomaticMappingBenchmarks` compares OfficeIMO's public `RowsAs<T>()`
  convenience mapper with its own explicit typed-reader loop. It is not a
  competitor ranking.
- `CsvParallelScalingBenchmarks` compares OfficeIMO's public ordered
  `ReadTextRowsAsParallel<T>` transient-record API with Sep's public ordered
  `ParallelEnumerate`. Both resolve headers once, parse typed fields from
  transient spans, create the same objects, and retain source order. It owns
  the sustained 100,000-row workload.
- `CsvParallelCrossoverBenchmarks` measures the same contract at 25,000 rows.
  Keeping it separate allows enough fixed invocations per iteration without
  multiplying the sustained workload.
- `CsvParallelOfficeTuningBenchmarks` varies only OfficeIMO's public batch size
  and worker limit. It intentionally has no competitor ratio because those are
  OfficeIMO implementation choices, not different semantic contracts.

All typed lanes resolve the same headers, materialize the same objects in
source order, and pass a property-by-property preflight.

Run the focused lanes during development:

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net8.0 -- --filter "*CsvTrimUnescapeBenchmarks*" --job short
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net8.0 -- --filter "*CsvTrimUnescapeSpanBenchmarks*" --job short
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net10.0 -- --filter "*CsvTypedSequentialBenchmarks*" --job short
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net10.0 -- --filter "*CsvAutomaticMappingBenchmarks*" --job short
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net10.0 -- --filter "*CsvParallelCrossoverBenchmarks*" --job short
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net10.0 -- --filter "*CsvParallelScalingBenchmarks*" --job short
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net10.0 -- --filter "*CsvParallelOfficeTuningBenchmarks*" --job short
```

Parallel materialization is not presented as a universal parser ranking. It
uses more memory and thread-pool work, so the two row counts are retained to
show whether the overhead pays for itself on the machine running the suite.

For processors with multiple cache or performance domains, pass explicit
machine-specific affinity masks as separate BenchmarkDotNet jobs:

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net10.0 -- --filter "*CsvTypedSequentialBenchmarks*" --affinityMasks "0x1,0x10000" --invocationCount 5 --unrollFactor 1 --warmupCount 5 --iterationCount 15 --launchCount 1
```

Masks must be derived from the current machine's CPU-set/cache topology and
recorded with the result; the sample values are not portable. Run parallel
lanes again with one whole cache domain and with all intended processors.
Record the active power plan, runtime, GC mode, logical processor count, and
whether a hypervisor is present. Treat rankings that change across cache
domains or fall inside the chosen statistical tolerance as inconclusive.
The comparison axis is the benchmark method within each CPU domain, so the
methods own the baseline and affinity jobs intentionally do not. Supply one
fixed `--invocationCount` with `--unrollFactor 1` across every affinity job;
this avoids asymmetric pilot counts without incorrectly mixing a method
baseline with a second job baseline. Use `--apples` only for a separate run
whose comparison axis is the jobs themselves and which has exactly one job
baseline.

The final fixed-work parallel commands on the dated machine were:

```powershell
# 25,000 rows; 25 invocations keep each measurement iteration long enough.
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net10.0 -- --filter "*CsvParallelCrossoverBenchmarks*" --affinityMasks "0xFFFF,0xFFFF0000" --invocationCount 25 --unrollFactor 1 --warmupCount 5 --iterationCount 10 --launchCount 3

# 100,000 rows; fewer invocations preserve the same fixed work per method/job.
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net10.0 -- --filter "*CsvParallelScalingBenchmarks*" --affinityMasks "0xFFFF,0xFFFF0000" --invocationCount 5 --unrollFactor 1 --warmupCount 5 --iterationCount 10 --launchCount 3

# Loaded-host diagnostic; alternates OfficeIMO/Sep as ABBA then BAAB and records wall and process CPU time.
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net10.0 -- --compare-typed-parallel-paired 60 0xFFFF adaptive 16 16 100000 High
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net10.0 -- --compare-typed-parallel-paired 60 0xFFFF0000 adaptive 16 16 100000 High
```

## Dated Sep feature snapshot (2026-08-08)

The current focused runs used .NET 10, BenchmarkDotNet 0.15.8, workstation GC, the
Windows High performance power plan, and an AMD Ryzen 9 9950X3D2 with 32
logical processors, two measured 16-logical-processor L3 domains, and a
hypervisor present. The earlier sequential/trim/span results below used one
launch, five warmups, fifteen measured iterations, semantic preflight
validation, and a fixed invocation count. The final ordered-parallel results
used three independent launches, five warmups, ten measured iterations, and
fixed invocation counts chosen per row count. CPU 0 and CPU 16 are the first
logical processor in each L3 domain. Both processors showed phase changes or
multimodal results in some methods, so overlapping confidence intervals are
reported as equality.

The final candidate assemblies were SHA-256
`10614E89447791035A6EFCDBCB8D23BF60EF4AD60F1F56ABEB38B1AE4760EBB7`
for `OfficeIMO.CSV.Benchmarks.dll` and
`C29F7E11878A2AD92073B2925CD6F8FE23048705BF7C554AD156826FB64C341C`
for `OfficeIMO.CSV.dll`. These hashes bind the numbers to the measured local
candidate even though the worktree did not have a commit for the changes.

The existing hash-pinned Mark Pflug 65K decoded-string lane was also run as
100 rotating paired samples after ten warmups:

| Affinity | OfficeIMO median | Sep median | Sylvan median | OfficeIMO / Sep paired median | OfficeIMO / Sylvan paired median |
| --- | ---: | ---: | ---: | ---: | ---: |
| CPU 0 | 17.627 ms | 20.347 ms | 20.743 ms | 0.861 (P25 0.800, P75 1.070) | 0.826 (P25 0.774, P75 1.053) |
| CPU 16 | 15.823 ms | 17.202 ms | 17.478 ms | 0.933 (P25 0.874, P75 0.974) | 0.976 (P25 0.869, P75 1.028) |

OfficeIMO has the best median in this full decoded-string lane, but the paired
intervals overlap parity on some domain/library combinations. This is a
workload result, not a universal parser ranking.

Outer-trim plus quote-unescape, 50,000 rows, decoded string contract:

| Affinity | OfficeIMO DataReader | Sep strings | CsvHelper strings | OfficeIMO allocation |
| --- | ---: | ---: | ---: | ---: |
| CPU 0 | 5.154 ms | 6.557 ms | 13.028 ms | 9.54 MB |
| CPU 16 | 7.395 ms | 8.470 ms | 12.348 ms | 9.54 MB |

OfficeIMO's allocation is effectively equal to Sep's 9.53 MB. OfficeIMO leads
on CPU 0; the broad, overlapping confidence intervals on CPU 16 make that
domain statistically equal rather than a defensible win.

The separately ranked transient-span contract produced:

| Affinity | OfficeIMO spans | Sep spans | OfficeIMO allocation | Sep allocation |
| --- | ---: | ---: | ---: | ---: |
| CPU 0 | 5.354 ms | 5.959 ms | 416 B | 728 B |
| CPU 16 | 6.327 ms | 5.548 ms | 416 B | 728 B |

Both span lanes are effectively allocation-free for this workload. The mean
winner changes by CPU domain and the confidence intervals overlap broadly, so
the timing result is equality. OfficeIMO allocates 312 B less per operation.

Equivalent explicit typed materialization produced:

| Rows | Affinity | OfficeIMO typed getters | Sep sequential | OfficeIMO allocation | Sep allocation |
| ---: | --- | ---: | ---: | ---: | ---: |
| 25,000 | CPU 0 | 15.64 ms | 18.18 ms | 10.02 MB | 9.80 MB |
| 25,000 | CPU 16 | 15.52 ms | 18.23 ms | 10.02 MB | 9.80 MB |
| 100,000 | CPU 0 | 44.75 ms | 52.92 ms | 40.06 MB | 39.27 MB |
| 100,000 | CPU 16 | 48.83 ms | 49.14 ms | 40.06 MB | 39.27 MB |

OfficeIMO leads three rows by mean; the 100,000-row CPU-16 result is equal and
strongly affected by OfficeIMO's bimodal distribution. A separate rotating
30-sample run on each domain put OfficeIMO at 0.8495 and 0.8449 of Sep's time,
which supports the improvement without replacing the isolated-process result.

Equivalent public ordered-parallel typed materialization, DOP 16:

| Rows | L3 domain | OfficeIMO mean (99.9% CI) | Sep mean (99.9% CI) | OfficeIMO allocation | Sep allocation | Result |
| ---: | --- | ---: | ---: | ---: | ---: | --- |
| 25,000 | `0xFFFF` | 5.523 ms (5.270-5.776) | 5.776 ms (5.413-6.138) | 9.81 MB | 9.89 MB | equal; OfficeIMO mean 4.4% lower |
| 25,000 | `0xFFFF0000` | 5.883 ms (5.300-6.467) | 5.197 ms (4.897-5.497) | 9.81 MB | 9.89 MB | equal; Sep mean 11.7% lower |
| 100,000 | `0xFFFF` | 24.44 ms (22.69-26.19) | 26.10 ms (24.02-28.18) | 39.29 MB | 39.38 MB | equal; OfficeIMO mean 6.4% lower |
| 100,000 | `0xFFFF0000` | 23.48 ms (21.63-25.34) | 25.19 ms (22.84-27.54) | 39.29 MB | 39.37 MB | equal; OfficeIMO mean 6.8% lower |

OfficeIMO now exposes the missing ordered-parallel capability and is equal to
Sep at both measured sizes on this fixture. OfficeIMO has the lower 100,000-row
mean on both domains, while the 25,000-row mean winner changes by domain. Its
pooled result buffers also put managed allocation slightly below Sep in all
four rows while clearing reference-containing buffers and honoring ordering,
cancellation, exception, and disposal contracts.

This host was also under sustained unrelated CPU load. A fresh 60-sample
symmetric ABBA/BAAB diagnostic reported OfficeIMO/Sep wall-time medians of
1.311 and 1.236 for the two domains, with very broad interquartile ranges of
0.810-1.716 and 0.831-1.502. Process-CPU ratios were 0.849 and 1.000, again
disagreeing with wall time. This confirms that scheduling pressure can change
the apparent winner on this machine. The isolated, fixed-work BenchmarkDotNet
table is the primary comparison, but its multimodal warnings and this paired
diagnostic must accompany any performance claim. Neither result is a universal
parser ranking.

For a SQL-shaped DataTable materialization pass:

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net8.0 -- --filter "*ReadDataTable*" --job short --warmupCount 1 --iterationCount 3
```

For the streaming `CsvDocument.ToDataTable` paths used by thin consumers such as PSWriteOffice:

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net8.0 -- --filter "*ReadStreamingDataTable*" --job short --warmupCount 1 --iterationCount 3
```

For a dbatools.library-shaped CSV reader pass:

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net8.0 -- --filter "*CsvDbatoolsLibraryParityBenchmarks*" --job short --warmupCount 1 --iterationCount 3
```

For the hash-pinned Mark Pflug 65K-record decoded-string comparison:

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net10.0 -- --filter "*MarkPflug65KCsvBenchmarks*"
```

This lane compares OfficeIMO, Sep, Sylvan, CsvHelper, Dataplat/dbatools, and
LumenWorks. Every implementation decodes every field into a string and must
match the same row count, cell count, and payload observation before the run is
accepted. Every library is a peer in the matrix; no implementation is framed as
an opponent or universal baseline. Results are interpreted per workload and
platform; `Build/Run-LibraryComparisonBenchmarks.ps1` keeps Windows, Linux, and
macOS evidence separate.

`CsvDbatoolsLibraryParityBenchmarks` mirrors the published dbatools.library CSV benchmark layout from [dataplat/dbatools.library `benchmarks/CsvBenchmarks`](https://github.com/dataplat/dbatools.library/tree/main/benchmarks/CsvBenchmarks), specifically `CsvReaderBenchmarks.Benchmarks.cs` and `QuickTest.cs`: small, medium, large, wide, quoted, modern medium/large, all-values, and quick-test-style single-column/all-column read lanes. It keeps OfficeIMO in the same file-path reader shape beside Dataplat.Dbatools.Csv, LumenWorks, Sep, Sylvan, and CsvHelper so the raw parser comparison is apples-to-apples. Each parity lane validates the expected row count and deterministic field-length checksum for its input file, so a lane cannot win by silently under-reading or skipping field materialization. The broader `CsvBenchmarks` and `CsvWideBenchmarks` lanes still touch every field and return checksums for stricter payload validation.

Parity check: the class includes all 20 upstream `CsvReaderBenchmarks` methods by benchmark description plus all 10 QuickTest read lanes, then adds matching OfficeIMO lanes beside them. The extra `OfficeIMO-DataReader-QuickTest-GetValues` lane keeps the SQL/bulk-copy-shaped `DbDataReader.GetValues` path visible at the same 100k-row QuickTest size. Dataplat remains the BenchmarkDotNet baseline in this parity class to preserve the upstream comparison frame. `TypeConverterBenchmarks` is intentionally out of scope here because it measures dbatools vector conversion rather than CSV parser throughput, not CSV reader throughput.

The generated table above and the dated sections below record earlier focused
investigations and their reproduction commands. Do not combine their numbers
into a current ranking.

## Dated dbatools.library parity snapshot (2026-07-09)

Archived local short-job run using the 14 QuickTest-shaped parity lanes with row-count and field-length checksum validation enabled. Treat these as direction-finding numbers; rerun before making release claims or performance gates.

QuickTest single-column/all-column read lanes:

| Method | Single column mean | All columns mean | Allocated |
| --- | ---: | ---: | ---: |
| OfficeIMO span reader | 4.57 ms | 4.70 ms | 770 KB / 770 KB |
| OfficeIMO streaming DataReader | 19.68 ms | 21.33 ms | 40.6 MB |
| SEP | 7.14 ms | 15.59 ms | 3.1 MB / 39.4 MB |
| Sylvan | 8.23 ms | 16.34 ms | 3.1 MB / 39.6 MB |
| CsvHelper | 30.15 ms | 45.63 ms | 3.1 MB / 39.6 MB |
| Dataplat.Dbatools.Csv | 26.13 ms | 29.38 ms | 39.9 MB |
| LumenWorks | 118.19 ms | 33.17 ms | 1.58 GB / 39.7 MB |

All-values read lane:

| Method | Mean | Allocated |
| --- | ---: | ---: |
| OfficeIMO span reader | 4.60 ms | 770 KB |
| OfficeIMO streaming DataReader | 22.64 ms | 41.3 MB |
| Dataplat.Dbatools.Csv DataReader | 26.70 ms | 39.9 MB |
| LumenWorks | 32.53 ms | 39.7 MB |

Additional guardrail: `OfficeIMO-DataReader-QuickTest-GetValues` reads the same
100k-row QuickTest file through `DbDataReader.GetValues`; this archived
short run measured 23.10 ms and 40.6 MB. Keep this lane visible when optimizing
the SQL/bulk-copy-shaped reader path.

The span-reader result is the fastest raw parser shape. The streaming DataReader result is the SQL/bulk-copy-shaped path; it now reads reusable parser string rows directly and is faster than Dataplat's DataReader in these short runs, with Dataplat still holding a small allocation edge.

## Dated typed DataReader snapshot (2026-07-09)

Archived local short-job runs using the 25,000-row, 40-column wide payload. Every lane traverses every value. The file lane includes file decoding and uses the public `CsvDocument.OpenDataReader(path, ...)` API used by PSWriteOffice and DbaClientX.

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net8.0 -- --filter "*CsvWideBenchmarks*DataReader*Schema*" --job short --warmupCount 5 --iterationCount 10
```

| Input | Schema | Mean | Allocated |
| --- | --- | ---: | ---: |
| CSV file | Explicit 40-column schema | 103.11 ms | 101.49 MB |
| CSV text | Explicit 40-column schema | 91.41 ms | 66.95 MB |
| CSV text | Inferred from 25,000 rows | 135.78 ms | 66.97 MB |

Explicit typed readers parse numbers, booleans, dates, and GUIDs directly from source spans. Inferred readers inspect spans without retaining sampled rows, then replay the immutable text through the typed reader. String-only file readers stay on the lower-memory streaming path.

## Dated wide read snapshot (2026-07-07)

Archived local short-job run:

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net8.0 -- --filter "*CsvWideBenchmarks*Read*FieldSpan*" --job short --warmupCount 1 --iterationCount 3
```

The table shows the fastest raw field-span read method per wide row-count lane. These lanes touch every field and compare OfficeIMO.CSV against SEP and Sylvan without PowerShell object overhead.

| Shape | Rows | Fastest method | Mean | SEP span read | Sylvan span read |
| --- | ---: | --- | ---: | ---: | ---: |
| Wide | 1000 | OfficeIMO_ReadTextFieldSpanVisitorSkipHeader | 0.06 ms | 0.08 ms | 0.11 ms |
| Wide | 10000 | OfficeIMO_ReadTextFieldSpanVisitorSkipHeader | 0.67 ms | 0.87 ms | 1.05 ms |
| Wide | 25000 | OfficeIMO_ReadTextFieldSpanVisitorSkipHeader | 1.73 ms | 2.09 ms | 2.79 ms |
