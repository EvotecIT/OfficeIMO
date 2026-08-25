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
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net10.0 -- --datareader-write-size-evidence --rows 25000,100000 --json .benchmark-artifacts\csv\datareader-write-size.json
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

### Dated SQL-shaped DataReader write snapshot (2026-08-09)

The same 25,000-row, 10-column contract was measured on .NET 10 with
BenchmarkDotNet 0.15.8, workstation GC, the Windows High performance power
plan, and separate jobs for both 16-logical-processor L3 domains on the AMD
Ryzen 9 9950X3D2. Every benchmark process used `High` priority, twelve warmups,
forty fixed invocations, twenty measured iterations, one launch, and retained
outliers. Target-specific setup validates only the output produced by the
method in that process, so one writer is not primed while measuring the other.

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -- --filter "*CsvDataReaderWriteBenchmarks*" --affinityMasks "0xFFFF,0xFFFF0000" --priority High --invocationCount 40 --unrollFactor 1 --warmupCount 12 --iterationCount 20 --launchCount 1 --outliers DontRemove

dotnet run -c Release -f net10.0 --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj --no-build -- --compare-datareader-write-paired 40 0xFFFF High
dotnet run -c Release -f net10.0 --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj --no-build -- --compare-datareader-write-paired 40 0xFFFF0000 High
```

| Shape | L3 domain | OfficeIMO mean (99.9% CI) | Sylvan mean (99.9% CI) | OfficeIMO / Sylvan allocation |
| --- | --- | ---: | ---: | ---: |
| Mixed | `0xFFFF` | 3.030 ms (2.866-3.194) | 3.742 ms (3.624-3.861) | 5.51 / 5.45 MB |
| Mixed | `0xFFFF0000` | 2.909 ms (2.831-2.987) | 3.661 ms (3.610-3.711) | 5.51 / 5.45 MB |
| Quoted | `0xFFFF` | 5.152 ms (4.788-5.516) | 5.985 ms (5.613-6.357) | 7.09 / 7.03 MB |
| Quoted | `0xFFFF0000` | 4.762 ms (4.328-5.196) | 5.581 ms (5.266-5.896) | 7.09 / 7.03 MB |
| Multiline | `0xFFFF` | 4.299 ms (3.849-4.749) | 5.668 ms (5.377-5.959) | 6.44 / 6.36 MB |
| Multiline | `0xFFFF0000` | 3.444 ms (3.167-3.721) | 4.815 ms (4.581-5.049) | 6.44 / 6.36 MB |

OfficeIMO has the lower isolated-process mean and a wholly lower confidence
interval in all six rows, by 14-29% of Sylvan's mean. Sylvan retains a small
managed-allocation advantage of about 1%.

The companion symmetric ABBA runner confirms clear OfficeIMO wins for mixed
and multiline values on both domains. Its paired OfficeIMO/Sylvan medians were
0.8398 and 0.8323 for mixed values, and 0.7972 and 0.7548 for multiline values.
Quoted values were more sensitive: domain 0 was a tie at 1.0014 (P25-P75
0.9474-1.0718), while domain 1 sat at the predeclared 5% boundary at 0.9499
(0.9351-0.9606). The honest conclusion is that OfficeIMO leads this complete
SQL-shaped workload in the isolated suite and clearly leads two of three
shapes in paired execution; it is not a universal CSV-writer ranking.

### Dated parallel DataReader write snapshot (2026-08-10)

The opt-in parallel writer consumes each `IDataReader` on one thread, formats
detached 4,096-row batches with four workers, and commits buffers in source
order. This keeps provider access safe while exposing parallel formatting for
sustained SQL-shaped exports. The sequential method remains the benchmark
baseline and the product default.

The isolated runs used the same runtime, GC, power plan, affinity jobs,
priority, warmup count, iteration count, and output validation described
above, with eight fixed invocations per iteration. The table shows the
100,000-row crossover; allocation is managed memory per operation.

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -- --filter "*CsvDataReaderWriteBenchmarks.OfficeIMO*" --affinityMasks "0xFFFF,0xFFFF0000" --priority High --invocationCount 8 --unrollFactor 1 --warmupCount 12 --iterationCount 20 --launchCount 1 --outliers DontRemove
dotnet run -c Release -f net10.0 --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj --no-build -- --filter "*CsvDataReaderWriteBenchmarks.Sylvan*" --affinityMasks "0xFFFF,0xFFFF0000" --priority High --invocationCount 8 --unrollFactor 1 --warmupCount 12 --iterationCount 20 --launchCount 1 --outliers DontRemove

dotnet run -c Release -f net10.0 --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj --no-build -- --compare-datareader-parallel-write-paired 30 0xFFFF High 100000 4 4096 4
dotnet run -c Release -f net10.0 --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj --no-build -- --compare-datareader-parallel-write-paired 30 0xFFFF0000 High 100000 4 4096 4
```

| Shape | L3 domain | OfficeIMO sequential | OfficeIMO parallel | Sylvan | Allocation: sequential / parallel / Sylvan |
| --- | --- | ---: | ---: | ---: | ---: |
| Mixed | `0xFFFF` | 18.912 ms | 13.291 ms | 19.161 ms | 21.71 / 24.50 / 21.81 MB |
| Mixed | `0xFFFF0000` | 14.568 ms | 13.629 ms | 17.053 ms | 21.71 / 24.50 / 21.81 MB |
| Quoted | `0xFFFF` | 24.427 ms | 16.492 ms | 22.763 ms | 28.03 / 31.09 / 28.15 MB |
| Quoted | `0xFFFF0000` | 20.529 ms | 14.144 ms | 19.489 ms | 28.03 / 31.08 / 28.15 MB |
| Multiline | `0xFFFF` | 16.610 ms | 13.613 ms | 21.668 ms | 25.37 / 28.17 / 25.47 MB |
| Multiline | `0xFFFF0000` | 14.449 ms | 13.830 ms | 20.698 ms | 25.37 / 28.17 / 25.47 MB |

The isolated parallel mean is 4-32% lower than OfficeIMO sequential in all
six rows. Four of six 99.9% confidence intervals are wholly separated; mixed
and multiline on `0xFFFF0000` overlap. Against Sylvan, every parallel mean and
confidence interval is lower. Managed allocation rises by about 10-13%.

Repeated alternating-order runs confirm clear elapsed-time wins for quoted
and multiline rows on both domains. Mixed rows improve at the median but are
more topology- and run-sensitive: one run's interquartile range crossed parity
and a repeat did not, so mixed formatting should not be treated as a universal
parallel win. Parallel processing also uses roughly 1.5-2.2 times the total process CPU.
At 25,000 rows, isolated results are mixed and quoted rows can be slower.
These results support an explicit large-export option, not a universal default.

### Dated typed parallel DataReader snapshot (2026-08-10)

This read lane uses the same generated 100,000-row UTF-8 file for every
method. Each implementation exposes five typed columns through its public
`DbDataReader`, traverses every row with `GetValues`, and returns an
order-sensitive checksum derived independently while generating the fixture.
The parallel methods use four workers and 4,096-row batches. This models the
reader side of an ordered SQL bulk-copy workflow; it does not rank unrelated
raw-text, span, DTO, or materialized-table contracts.

The run used .NET 10, BenchmarkDotNet 0.15.8, workstation GC, the Windows High
performance power plan, and separate jobs for both 16-logical-processor L3
domains on the AMD Ryzen 9 9950X3D2. Each job used `High` priority, twelve
warmups, twelve fixed invocations, twenty measured iterations, one launch, and
retained outliers.

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net10.0 -- --filter "*CsvParallelDataReaderBenchmarks*" --affinityMasks "0xFFFF,0xFFFF0000" --priority High --invocationCount 12 --unrollFactor 1 --warmupCount 12 --iterationCount 20 --launchCount 1 --outliers DontRemove
```

| L3 domain | Method | Mean (99.9% CI) | Managed allocation |
| --- | --- | ---: | ---: |
| `0xFFFF` | OfficeIMO sequential | 21.829 ms (20.739-22.920) | 31.16 MB |
| `0xFFFF` | OfficeIMO parallel | 12.622 ms (11.608-13.636) | 31.18 MB |
| `0xFFFF` | Dataplat sequential | 72.532 ms (70.336-74.728) | 28.49 MB |
| `0xFFFF` | Dataplat parallel | 36.682 ms (36.075-37.289) | 51.84 MB |
| `0xFFFF0000` | OfficeIMO sequential | 18.999 ms (18.112-19.886) | 31.16 MB |
| `0xFFFF0000` | OfficeIMO parallel | 12.471 ms (11.871-13.070) | 31.18 MB |
| `0xFFFF0000` | Dataplat sequential | 55.984 ms (49.929-62.038) | 28.49 MB |
| `0xFFFF0000` | Dataplat parallel | 40.290 ms (39.330-41.249) | 53.26 MB |

OfficeIMO parallel is 42.2% and 34.4% faster than OfficeIMO sequential on the
two domains. Against Dataplat parallel, it uses 65.6% and 69.0% less time and
about 40-41% less managed memory. The 99.9% confidence intervals are wholly
separated for both comparisons on both domains. BenchmarkDotNet identified a
bimodal distribution for OfficeIMO parallel on `0xFFFF` and a multimodal
Dataplat sequential distribution on `0xFFFF0000`; the table retains every
measured iteration and should be interpreted with that topology sensitivity in
mind.

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
- `CsvParallelOfficeTuningBenchmarks` and
  `CsvParallelCrossoverTuningBenchmarks` vary only OfficeIMO's public batch
  size and worker limit at 100,000 and 25,000 rows respectively. The row
  counts stay in separate benchmark types so ranks never compare different
  workloads. These diagnostics intentionally have no competitor ratio because
  the parameters are OfficeIMO implementation choices.

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
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net10.0 -- --filter "*CsvParallelCrossoverTuningBenchmarks*" --job short
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

## Dated Sep feature snapshots (2026-08-08 and 2026-08-10)

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

The 2026-08-08 candidate assemblies were SHA-256
`10614E89447791035A6EFCDBCB8D23BF60EF4AD60F1F56ABEB38B1AE4760EBB7`
for `OfficeIMO.CSV.Benchmarks.dll` and
`C29F7E11878A2AD92073B2925CD6F8FE23048705BF7C554AD156826FB64C341C`
for `OfficeIMO.CSV.dll`. These hashes bind the 2026-08-08 sequential, trim,
span, and 25,000-row parallel numbers to that measured local candidate.

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
| 100,000 | `0xFFFF` | 23.67 ms (22.00-25.34) | 23.84 ms (22.39-25.29) | 39.29 MB | 39.38 MB | equal; OfficeIMO mean 0.7% lower |
| 100,000 | `0xFFFF0000` | 22.24 ms (20.72-23.76) | 21.74 ms (20.03-23.45) | 39.30 MB | 39.37 MB | equal; Sep mean 2.2% lower |

OfficeIMO now exposes the missing ordered-parallel capability and is equal to
Sep at both measured sizes on this fixture. The 100,000-row results were refreshed
on 2026-08-10 after an OfficeIMO-only fixed-work sweep showed that 2,048-row
batches were 13-16% faster than the previous 3,584-row default on both L3
domains. The final three-launch competitor run remained a tie: the mean winner
changed by domain and all 99.9% intervals overlap. OfficeIMO's pooled result
buffers also put managed allocation slightly below Sep in all four rows while
clearing reference-containing buffers and honoring ordering, cancellation,
exception, and disposal contracts.

The refreshed 2026-08-10 100,000-row results used SHA-256
`059F48B634BA2DC3B591D50B005782767565BC235994BF4257BEF012A413BC20`
for `OfficeIMO.CSV.Benchmarks.dll` and
`FB24E005C297C48B4B3A3B8CB18D512B747150C18E79F5F2C82765CC54E6F97C`
for `OfficeIMO.CSV.dll`. Each active tuning parameter was separately
preflighted against every expected property before its timing run.

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

`CsvDbatoolsLibraryParityBenchmarks` mirrors the published dbatools.library CSV benchmark layout from [dataplat/dbatools.library `benchmarks/CsvBenchmarks`](https://github.com/dataplat/dbatools.library/tree/main/benchmarks/CsvBenchmarks), specifically `CsvReaderBenchmarks.Benchmarks.cs` and `QuickTest.cs`: small, medium, large, wide, quoted, modern medium/large, all-values, and quick-test-style single-column/all-column read lanes. OfficeIMO uses its public `DbDataReader` string contract in this class; its zero-copy field-span API is measured separately. Each implementation opens the same file, materializes the requested strings, and validates the expected row count and deterministic field-length checksum, so a lane cannot win by silently under-reading or skipping the payload. The broader `CsvBenchmarks` and `CsvWideBenchmarks` lanes still touch every field and return checksums for stricter payload validation.

Parity check: the class includes all 20 upstream `CsvReaderBenchmarks` methods by benchmark description plus all 10 QuickTest read lanes, then adds matching OfficeIMO lanes beside them. The extra `OfficeIMO-DataReader-QuickTest-GetValues` lane keeps the SQL/bulk-copy-shaped `DbDataReader.GetValues` path visible at the same 100k-row QuickTest size. BenchmarkDotNet groups results by workload category and uses the matching Dataplat method as that category's baseline; it does not rank a small single-column read against a large or all-column read. `TypeConverterBenchmarks` is intentionally out of scope because it measures dbatools vector conversion rather than CSV reader throughput.

### Dated dbatools.library QuickTest snapshot (2026-08-10)

This .NET 10 run read the same 100,000-row, ten-column UTF-8 file through
each library's public decoded-string API. The single-column lane materialized
column zero; the all-column lane materialized all ten fields. Every invocation
validated the row count and deterministic field-length checksum. The two
16-logical-processor L3 domains were separate jobs with `High` process
priority, twelve warmups, twenty measured iterations, one launch, and retained
outliers. The all-column and `GetValues` lanes used twelve fixed invocations;
the shorter single-column lane used 32 so every measured iteration performed
enough fixed work without BenchmarkDotNet choosing different pilot counts.

The OfficeIMO all-column row below is
`OfficeIMO-QuickTest-AllColumns`, which reads each ordinal with
`DbDataReader.GetValue(i)`. CsvHelper and LumenWorks are included as peers in
both tables. The SQL/bulk-copy-shaped
`OfficeIMO-DataReader-QuickTest-GetValues` lane is reported separately because
the other rows do not call the equivalent bulk API.

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net10.0 -- --filter "*QuickTest*AllColumns*" --affinityMasks "0xFFFF,0xFFFF0000" --priority High --invocationCount 12 --unrollFactor 1 --warmupCount 12 --iterationCount 20 --launchCount 1 --outliers DontRemove
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net10.0 -- --filter "*QuickTest*SingleColumn*" --affinityMasks "0xFFFF,0xFFFF0000" --priority High --invocationCount 32 --unrollFactor 1 --warmupCount 12 --iterationCount 20 --launchCount 1 --outliers DontRemove
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net10.0 -- --filter "*OfficeIMO*QuickTest*GetValues*" --affinityMasks "0xFFFF,0xFFFF0000" --priority High --invocationCount 12 --unrollFactor 1 --warmupCount 12 --iterationCount 20 --launchCount 1 --outliers DontRemove
```

| Contract | Library | L3 domain 0 mean (99.9% CI) | L3 domain 1 mean (99.9% CI) | Managed allocation |
| --- | --- | ---: | ---: | ---: |
| All columns | OfficeIMO.CSV | 12.552 ms (11.193-13.911) | 9.865 ms (9.078-10.652) | 39.40 MB |
| All columns | Sep | 13.307 ms (12.003-14.612) | 10.984 ms (10.275-11.693) | 39.40 MB |
| All columns | Sylvan.Data.Csv | 12.371 ms (11.635-13.107) | 13.212 ms (12.125-14.299) | 39.64 MB |
| All columns | Dataplat.Dbatools.Csv | 23.466 ms (21.815-25.117) | 27.776 ms (25.433-30.119) | 39.86 MB |
| All columns | CsvHelper | 36.090 ms (33.535-38.645) | 37.538 ms (33.926-41.150) | 39.63 MB |
| All columns | LumenWorks.FastCsvReader | 26.631 ms (23.046-30.216) | 29.060 ms (26.170-31.950) | 39.74 MB |
| First column | OfficeIMO.CSV | 3.845 ms (3.651-4.039) | 3.950 ms (3.505-4.395) | 3.06 MB |
| First column | Sep | 6.676 ms (6.379-6.973) | 8.127 ms (7.325-8.929) | 3.06 MB |
| First column | Sylvan.Data.Csv | 6.887 ms (6.724-7.050) | 6.887 ms (6.701-7.073) | 3.09 MB |
| First column | Dataplat.Dbatools.Csv | 19.633 ms (19.167-20.099) | 19.499 ms (18.819-20.179) | 39.86 MB |
| First column | CsvHelper | 24.744 ms (24.082-25.406) | 29.733 ms (28.303-31.163) | 3.08 MB |
| First column | LumenWorks.FastCsvReader | 77.141 ms (75.447-78.835) | 74.565 ms (67.703-81.427) | 1,619.67 MB |

OfficeIMO's single-column confidence intervals are wholly below Sep and Sylvan
on both domains. For all-column reads, Sylvan has a 0.181 ms lower mean on
domain 0 and OfficeIMO has the lowest mean on domain 1; the leading intervals
overlap, so the top all-column result is statistical parity rather than a win.
An earlier isolated run put OfficeIMO's all-column mean first on both domains,
which is another reason not to cherry-pick a CPU-sensitive ranking. In this
final run OfficeIMO is 46-64% faster than Dataplat for all-column reads, about
80% faster for first-column reads, and allocates about 92% less than Dataplat
when only the requested column is materialized. These are workload-specific
results, not a universal parser ranking.

The supplementary `DbDataReader.GetValues` run measured 11.712 ms
(11.195-12.229 ms) on domain 0 and 10.850 ms (10.046-11.653 ms) on domain 1,
allocating 39.40 MB. The domain-1 distribution was multimodal and its detected
outlier was retained. This lane demonstrates the bulk API shape used by SQL
consumers; it is not ranked against the ordinal-access rows because those
contracts are different.

The generated table above and the dated sections below record earlier focused
investigations and their reproduction commands. Do not combine their numbers
into a current ranking.

Run the two QuickTest contracts independently when comparing current code:

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net10.0 -- --filter "*CsvDbatoolsLibraryParityBenchmarks*" --anyCategories QuickTestSingleColumn QuickTestAllColumns --job short --warmupCount 3 --iterationCount 9
```

Do not combine these results with `CsvTrimUnescapeSpanBenchmarks` or another
field-span suite. Those APIs intentionally borrow transient spans and answer a
different performance question.

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
