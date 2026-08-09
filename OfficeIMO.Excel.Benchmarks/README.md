# OfficeIMO.Excel.Benchmarks

`OfficeIMO.Excel.Benchmarks` is the benchmark harness for `OfficeIMO.Excel`. It measures representative workbook read, write, edit, package-size, and real-world feature workloads. It is not a runtime package.

## Run benchmarks

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj
```

Filter a class:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --filter *ExcelWriteBenchmarks*
```

Measure worksheet copy fast paths:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --filter *ExcelWorksheetCopyBenchmarks*
```

## Snapshot and profile artifacts

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --snapshot .\Docs\benchmarks\officeimo.excel.snapshot.json
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- write-profile --rows 25000
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- read-profile --rows 2500
```

On Windows, use a machine-specific affinity mask to keep sequential and
parallel read-policy comparisons inside the intended CPU/cache domain. Repeat
with every domain and with all intended processors; the JSON records the mask,
logical processor count, and GC mode.

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- read-profile .\Ignore\Benchmarks\excel-read-ccd0.json --rows 25000 --warmup 5 --iterations 10 --affinity 0xFFFF
```

Affinity masks are topology-specific. Do not copy the example to another
machine without deriving its CPU sets and cache domains first.

### Dated read-policy snapshot (2026-08-08)

This 25,000-row, eight-column local .NET 8 profile used workstation GC, three
warmups, seven measured samples, and the Windows High performance power plan on
an AMD Ryzen 9 9950X3D2. It measured both 16-logical-processor L3 domains
separately (`0xFFFF` and `0xFFFF0000`). Values are medians in milliseconds.

The final candidate assemblies were SHA-256
`78D36A3BBAB93B013F17768F6D1302260532F9FD6058A32D9F766A72B6B76C61`
for `OfficeIMO.Excel.Benchmarks.dll` and
`5A88E63D531FC7725DAD42EC23E2B04CF2F8218FDA030EA948E1544AA6C8C919`
for `OfficeIMO.Excel.dll`.

| API | Affinity | Automatic | Sequential | Forced parallel |
| --- | --- | ---: | ---: | ---: |
| `ReadObjects` dictionaries | L3 domain 0 | 39.70 | 42.61 | 40.33 |
| `ReadObjects` dictionaries | L3 domain 1 | 40.05 | 40.29 | 40.57 |
| `ReadObjectsAs<T>` | L3 domain 0 | 28.69 | 45.29 | 35.57 |
| `ReadObjectsAs<T>` | L3 domain 1 | 25.06 | 43.84 | 33.09 |
| `ReadRange` | L3 domain 0 | 27.16 | 25.35 | 25.42 |
| `ReadRange` | L3 domain 1 | 25.48 | 25.25 | 25.00 |
| `ReadRangeAsDataTable` | L3 domain 0 | 51.47 | 53.93 | 50.80 |
| `ReadRangeAsDataTable` | L3 domain 1 | 51.22 | 53.72 | 51.43 |
| `ReadRangeStream` | L3 domain 0 | 25.61 | 25.84 | 25.20 |
| `ReadRangeStream` | L3 domain 1 | 24.83 | 24.00 | 24.66 |

Explicit parallel reads no longer bypass the specialized XML and UTF-8
readers. The dense range, DataTable, and stream modes are effectively equal at
this size, while the typed-object automatic and parallel routes retain the
direct typed reader and materially outperform the forced sequential fallback.
`Parallel` remains an execution preference rather than a promise to discard a
faster single-pass reader; diagnostics report the strategy actually selected.

### Dated DataTable execution-mode write snapshot (2026-08-09)

`ExcelDataTableExecutionBenchmarks` creates and saves the same 25,000-row XLSX
package through the public `InsertDataTable` API in Automatic, Sequential, and
Parallel mode. Setup reopens each package and compares every header and cell
before timing. The timed work includes workbook creation, insertion, and save.

This .NET 10 run used BenchmarkDotNet 0.15.8, workstation GC, the Windows High
performance power plan, fixed `High` process priority, and separate jobs for
both 16-logical-processor L3 domains on the AMD Ryzen 9 9950X3D2. Outliers were
retained.

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --filter "*ExcelDataTableExecutionBenchmarks*" --affinityMasks "0xFFFF,0xFFFF0000" --priority High --invocationCount 8 --unrollFactor 1 --warmupCount 12 --iterationCount 20 --launchCount 1 --outliers DontRemove

dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --compare-datatable-execution-paired 40 0xFFFF High
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj --no-build -- --compare-datatable-execution-paired 40 0xFFFF0000 High
```

| Mode | L3 domain 0 mean (99.9% CI) | L3 domain 1 mean (99.9% CI) | Managed allocation |
| --- | ---: | ---: | ---: |
| Automatic | 27.67 ms (25.48-29.86) | 29.02 ms (27.63-30.41) | 11.22 MB |
| Sequential | 29.75 ms (28.38-31.13) | 29.40 ms (27.49-31.32) | 11.22 MB |
| Parallel | 31.45 ms (29.68-33.23) | 26.48 ms (24.17-28.80) | 11.22 MB |

All three modes shared rank 1 in both BenchmarkDotNet jobs. Automatic and
Sequential use the same direct writer here, while the small changes in their
relative means across CPU domains are useful evidence of the machine's temporal
and domain sensitivity rather than a reason to declare a winner from one
favorable job.

The companion runner therefore measures Automatic and Parallel as alternating
ABBA pairs in one pinned process. On domain 0 their medians were 35.639 ms and
35.640 ms; the paired Parallel/Automatic ratio median was 1.0209 (P25 0.9146,
P75 1.1168). On domain 1 their medians were 32.709 ms and 34.738 ms; the paired
ratio median was 1.0352 (P25 0.9541, P75 1.1217). The interquartile ranges
straddle 1 on both domains. That supports practical parity on this workload,
but it is not a statistical-equivalence proof.

Before this fast path, the same Parallel workload took 1,201.16 ms on domain 0
and 946.88 ms on domain 1 while allocating 432.93 MB. The final path is about
31-36 times faster and uses about 36.6 times less managed memory. A Parallel
request now keeps the specialized package writer, eagerly snapshots values so
later source mutations cannot leak into the workbook, and checks cancellation
during snapshot creation. Package serialization itself remains a specialized
single-pass operation rather than parallel work.

### Dated compact DataReader write snapshot (2026-08-09)

This BenchmarkDotNet 0.15.8 run wrote the same prepared 25,000-row,
eight-column `DataTableReader` to a forward-only XLSX package through each
library's public compact writer. Target-specific setup reopens that target's
package and compares every header and cell without priming the other writers in
the same benchmark process. OfficeIMO 3.2.0 was compared with
SpreadCheetah 1.28.0, Sylvan.Data.Excel 0.5.7, and LargeXlsx 2.0.1 on .NET 10
with workstation GC and the Windows High performance power plan.

The two 16-logical-processor L3 domains on the AMD Ryzen 9 9950X3D2 were
measured as separate jobs. Twelve warmups let tiered compilation settle, and
all methods used the same twelve invocations per sample, twenty measured
iterations, and one launch. Every benchmark process also used the same `High`
priority class to reduce interference from unrelated services; affinity and
priority are both recorded in the BenchmarkDotNet evidence.

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --filter "*ExcelDataReaderWriteBenchmarks*" --affinityMasks "0xFFFF,0xFFFF0000" --priority High --invocationCount 12 --unrollFactor 1 --warmupCount 12 --iterationCount 20 --launchCount 1 --outliers DontRemove

dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --compare-datareader-write-paired 40 0xFFFF High
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj --no-build -- --compare-datareader-write-paired 40 0xFFFF0000 High
```

| Library | L3 domain 0 mean (99.9% CI) | L3 domain 1 mean (99.9% CI) | Managed allocation |
| --- | ---: | ---: | ---: |
| OfficeIMO.Excel | 13.30 ms (12.71-13.90) | 13.71 ms (13.27-14.15) | 6.18 MB |
| SpreadCheetah | 14.36 ms (13.73-14.98) | 14.83 ms (14.48-15.18) | 5.62 MB |
| Sylvan.Data.Excel | 22.45 ms (20.26-24.63) | 20.98 ms (20.45-21.51) | 7.52 MB |
| LargeXlsx | 19.62 ms (18.89-20.35) | 18.95 ms (17.86-20.05) | 5.67 MB |

OfficeIMO recorded the lowest mean on both domains. Its interval overlaps
SpreadCheetah's narrowly on domain 0 and is wholly below it on domain 1;
OfficeIMO's intervals are wholly below Sylvan's and LargeXlsx's on both domains.
The companion ABBA run measured OfficeIMO at 16.911 ms versus 18.092 ms on
domain 0 (ratio of medians 0.9347, paired median 0.9496, P25-P75
0.8738-0.9813) and 15.418 ms versus 15.946 ms on domain 1 (ratio 0.9669,
paired median 0.9639, P25-P75 0.9395-1.0289). Under the repository's 5%
threshold this is a domain-0 win and a domain-1 tie, not a universal fastest
claim. SpreadCheetah retains the managed-allocation lead, although OfficeIMO's
pooled package writer reduced this workload from 6.50 MB to 6.18 MB and also
serves typed, asynchronous, and parallel worksheet exports.

### Dated PowerShell PSObject write snapshot (2026-08-09)

This comparison writes 25,000 mixed ten-column rows to an in-memory XLSX
package. OfficeIMO receives PSObject-like rows through the normal
`InsertObjects` API, including property projection, type inference, and package
serialization. LargeXlsx receives equivalent pre-created dictionary rows
through its streaming API. This is therefore a representative PowerShell export
comparison, not an identical input-contract microbenchmark. Setup validates each
package before timing.

The .NET 10 runner used workstation GC, 12 warmups, 60 measured iterations,
rotated execution order, `High` process priority, and separate runs on both
16-logical-processor L3 domains of the AMD Ryzen 9 9950X3D2. Allocation uses
`GC.GetTotalAllocatedBytes(precise: true)`, so worker-thread allocations from
parallel projection are included.

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare .\Ignore\Benchmarks\psobject-domain0.json --rows 25000 --scenario write-powershell-psobject-mixed-direct --library OfficeIMO.Excel --library LargeXlsx --skip-legacy-epplus --warmup 12 --iterations 60 --affinity 0xFFFF --priority High
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj --no-build -- compare .\Ignore\Benchmarks\psobject-domain1.json --rows 25000 --scenario write-powershell-psobject-mixed-direct --library OfficeIMO.Excel --library LargeXlsx --skip-legacy-epplus --warmup 12 --iterations 60 --affinity 0xFFFF0000 --priority High
```

| Library | Domain 0 average / median | Domain 1 average / median | Average managed allocation |
| --- | ---: | ---: | ---: |
| OfficeIMO.Excel | 27.07 / 26.74 ms | 22.33 / 19.86 ms | 12.91 MB |
| LargeXlsx | 39.23 / 39.91 ms | 32.91 / 30.58 ms | 10.53 MB |

OfficeIMO was faster by both average and median on both domains for this user
job. LargeXlsx allocated about 2.4 MB less per export, so OfficeIMO does not
claim the allocation lead. The prior OfficeIMO path measured approximately
52-53 ms median on the same machine; direct bounded MemoryStream output,
compact worksheet cells, and ordered parallel PSObject projection account for
the improvement while preserving valid XLSX output and source-row order.

## Library comparison

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare --rows 2500
```

Each scenario runs only libraries with a directly comparable public API. Legacy
EPPlus runs in a separate process. NPOI comparisons are available through the
opt-in [NPOI runner](../OfficeIMO.Excel.Benchmarks.NPOI/README.md).

The hash-pinned Mark Pflug 65K-record read comparisons are available as focused
BenchmarkDotNet classes:

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --filter "*MarkPflug65KXlsxBenchmarks*"
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --filter "*MarkPflug65KXlsBenchmarks*"
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --filter "*MarkPflug65KXlsbBenchmarks*"
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --filter "*XlsbOfficeIMOPipelineBenchmarks*"
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --profile-markpflug65k-xls-officeimo 100
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --profile-markpflug65k-xls-sylvan 100
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --profile-markpflug65k-xlsb-officeimo 100
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --profile-markpflug65k-xlsb-sylvan 100
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --compare-markpflug65k-xls-paired 40 0x1 High
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --compare-markpflug65k-xlsb-paired 40 0x1 High
```

For one fixed-work run across both measured CPU domains, add affinity jobs and
use the same invocation count for every method and job:

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --filter "*MarkPflug65KXlsxBenchmarks*" --affinityMasks "0x1,0x10000" --invocationCount 1 --unrollFactor 1 --warmupCount 5 --iterationCount 15 --launchCount 1
```

These suites compare library methods within each CPU job, so method baselines
and absolute per-job results are the appropriate shape. Do not combine those
method baselines with an `--apples` job-baseline comparison.

### Dated 65K XLSX snapshot (2026-08-07)

The fixed-work command above was run locally on .NET 10 with workstation GC,
the Windows High performance power plan, and an AMD Ryzen 9 9950X3D2. CPU 0
and CPU 16 are the first logical processor in each measured L3 domain. Every
method read the same hash-pinned workbook and passed the row, cell, and payload
observation before measurement.

| Library | CPU 0 mean | CPU 16 mean | Managed allocation |
| --- | ---: | ---: | ---: |
| OfficeIMO.Excel | 450.1 ms | 481.4 ms | 366.88 KB |
| Sylvan.Data.Excel | 586.2 ms | 484.5 ms | 649 KB |
| ExcelDataReader | 478.6 ms | 457.3 ms | 207,509 KB |
| ClosedXML | 1,288.1 ms | 1,251.5 ms | 725,535-725,691 KB |
| EPPlus | 1,174.5 ms | 888.2 ms | 864,333-864,334 KB |
| MiniExcel | 561.6 ms | 435.3 ms | 666,371 KB |

OfficeIMO has the lowest mean on CPU 0. MiniExcel has the lowest mean on CPU
16, but their 99.9% confidence intervals overlap, so this run does not resolve
a difference on that domain; overlap is not proof of equivalence. OfficeIMO
has the lowest managed allocation on both domains
by a wide margin. Several methods were bimodal or changed performance phase;
the full distributions matter more than selecting one favorable cluster.

### Dated 65K XLS and XLSB snapshot (2026-08-09)

The legacy-format runner uses twelve warmups followed by forty symmetric ABBA
samples at High process priority. Each sample averages two reads per library,
alternates which library runs outside the pair, and rejects any row, cell, or
payload observation mismatch. These results used the same workstation and CPU
domains as the XLSX snapshot above.

| Format | CPU | OfficeIMO median | Sylvan median | Ratio of medians | Paired ratio median (P25-P75) |
| --- | --- | ---: | ---: | ---: | ---: |
| XLS | CPU 0 | 49.494 ms | 50.716 ms | 0.9759 | 0.9847 (0.9498-1.0334) |
| XLS | CPU 16 | 44.355 ms | 46.572 ms | 0.9524 | 0.9502 (0.9425-0.9661) |
| XLSB | CPU 0 | 76.264 ms | 74.601 ms | 1.0223 | 0.9913 (0.9294-1.0411) |
| XLSB | CPU 16 | 55.135 ms | 56.471 ms | 0.9763 | 0.9803 (0.9607-0.9920) |

All four differences are below the repository's predeclared 5% threshold, so
they are classified as ties rather than wins. OfficeIMO had the lower observed
median in both XLS domains and one XLSB domain. A separate BenchmarkDotNet run
with four invocations, twelve warmups, twenty retained-outlier iterations, and
both affinity jobs measured steady-state XLSB allocation at 115.15-115.23 KB
for OfficeIMO, 343.78 KB for Sylvan, and 90,233.52-90,236.31 KB for
ExcelDataReader. Its CPU 16 OfficeIMO timing changed phase dramatically, so the
ABBA runner above is the relative timing evidence; the BenchmarkDotNet run is
used for allocation evidence and its timing outliers remain visible.

Reproduce both timing domains and the allocation run with:

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --compare-markpflug65k-xls-paired 40 0x1 High
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --compare-markpflug65k-xls-paired 40 0x10000 High
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --compare-markpflug65k-xlsb-paired 40 0x1 High
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --compare-markpflug65k-xlsb-paired 40 0x10000 High
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --filter "*MarkPflug65KXlsbBenchmarks*" --affinityMasks "0x1,0x10000" --priority High --invocationCount 4 --unrollFactor 1 --warmupCount 12 --iterationCount 20 --launchCount 1 --outliers DontRemove
```

The four `--profile-markpflug65k-*` commands are lightweight profiling loops,
not publication-grade benchmark runs. The fixture is authenticated during
setup, and every measured library invocation must produce the expected row
count, cell count, and deterministic payload observation. The paired commands
control short-run ordering drift; use the BenchmarkDotNet class filters above
when statistical error and allocation measurements are required.

The XLS lane includes OfficeIMO, Sylvan.Data.Excel, and ExcelDataReader. The
XLSX lane includes OfficeIMO, Sylvan.Data.Excel, ExcelDataReader,
ClosedXML, EPPlus, and MiniExcel. The XLSB lane includes only the compatible
readers: OfficeIMO, Sylvan.Data.Excel, and ExcelDataReader. Every implementation
reads the same fourteen typed columns and must produce the same row count, cell
count, and deterministic payload observation before a measurement is accepted.
Write-focused libraries such as LargeXlsx and SpreadCheetah remain in the write
suites and are not applicable to these read lanes. Every library is a peer in
the matrix; no implementation is framed as an opponent or universal baseline.
The OfficeIMO-only XLSB pipeline diagnostic is not part of that library matrix;
it isolates public multi-sheet dispatch from the underlying worksheet reader
when profiling a demonstrated OfficeIMO bottleneck.

The suite keeps materially different contracts in separate lanes. Compact
writers omit explicit cell references for forward-only throughput, while the
normal OfficeIMO writer preserves the editable worksheet model. Shared-string
reads distinguish forward-only scans from rectangular materialization, and
DataTable reads distinguish automatic type inference from a caller-prepared
typed schema. These lanes should not be collapsed into one ranking.

String-storage readers are compared on separately validated shared-string and
inline-string fixtures. In a 25,000-row Windows run with twenty warmups and nine
measurements, OfficeIMO's forward reader measured 15.98 ms for shared strings
and 20.71 ms for inline strings; Sylvan.Data.Excel measured 20.46 ms and 25.50
ms respectively. These are workstation measurements rather than universal
constants, so rerun the focused scenarios before making a release claim.

For release-style evidence:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- comparison-suite --out-dir .\Docs\benchmarks\comparison-current --row-set 2500,25000 --skip-legacy-epplus --warmup 20 --iterations 9
```

The twenty warmups let tiered compilation settle before the custom rotated
runner starts measuring; fifteen warmups still produced visibly bimodal
isolated read results on ARM64.
The suite writes JSON, CSV, Markdown, and a manifest. Run the focused README
comparison and refresh its generated table locally with:

```powershell
.\Build\Benchmarks\Update-BenchmarkReadmes.ps1 -Run Excel
```

The script selects documented equivalent workloads, emits PSPublishModule's
comparison schema, and calls `Update-BenchmarkDocument` for the
marker-delimited block. Benchmark execution is local and is not scheduled in CI.

Focus the package-copy workbook merge scenario:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare --rows 25000 --scenario copy-worksheet-package --warmup 1 --iterations 3
```

Compare row scanning, selective field access, full `GetValues`, and typed getters:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare .\Ignore\Benchmarks\excel-datareader-25000.json --rows 25000 --scenario read-datareader-readonly,read-datareader-first-column,read-datareader-getvalues,read-datareader-typed --skip-legacy-epplus --warmup 3 --iterations 15
```

Compare the fastest package-native write paths:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare .\Ignore\Benchmarks\excel-write-25000.json --rows 25000 --scenario write-datareader-compact-package,write-typed-rows-compact-package --skip-legacy-epplus --warmup 20 --iterations 31
```

## Historical generated workstation snapshot

This PSPublishModule-managed snapshot is retained for reproducibility, not as
the current cross-platform product ranking. The rows combine raw data paths
with feature-bearing workbook work and only compare libraries that expose a
directly comparable public API. Lower is
faster within a row only; do not combine rows into one library ranking. Results
vary by machine, runtime, package version, workload, warm-up, and options.
Treat differences below 5% as ties. Use the hash-pinned CSV/XLSX/XLSB website
matrix for current platform-separated evidence.

<!-- officeimo-excel-benchmark-table:start -->
| Scenario | Variables | Host | Operation | OfficeIMO.Excel | ClosedXML | EPPlus | LargeXlsx | SpreadCheetah | Sylvan.Data.Excel | Result |
| --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | --- |
| Compact DataReader to XLSX | Format=.xlsx, MeasuredIterations=9, Rows=25,000, Runner=rotated local, Snapshot=2026-07-14, Warmups=20 | .NET 8 | Write | 1.00x (23ms) | n/a | n/a | 1.11x (26ms) | 1.00x (23ms) | 1.11x (26ms) | OfficeIMO.Excel tied with SpreadCheetah |
| Feature-rich report to XLSX | Format=.xlsx, MeasuredIterations=9, Rows=25,000, Runner=rotated local, Snapshot=2026-07-14, Warmups=20 | .NET 8 | Create | 1.00x (37ms) | n/a | 11.12x (409ms) | n/a | n/a | n/a | OfficeIMO.Excel fastest |
| Styled DataReader table to XLSX | Format=.xlsx, MeasuredIterations=9, Rows=25,000, Runner=rotated local, Snapshot=2026-07-14, Warmups=20 | .NET 8 | Write | 1.00x (34ms) | 9.50x (320ms) | 9.76x (329ms) | n/a | n/a | n/a | OfficeIMO.Excel fastest |
| Typed objects streamed from XLSX | Format=.xlsx, MeasuredIterations=9, Rows=25,000, Runner=rotated local, Snapshot=2026-07-14, Warmups=20 | .NET 8 | Read | 1.00x (25ms) | 11.13x (278ms) | 10.08x (252ms) | n/a | n/a | 1.56x (39ms) | OfficeIMO.Excel fastest |
<!-- officeimo-excel-benchmark-table:end -->

`--skip-legacy-epplus` omits only the isolated EPPlus 4.x subprocess; current
EPPlus remains in the comparison. Keep this flag on modern macOS unless
`libgdiplus` is installed because the legacy AutoFit path depends on it:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare --rows 25000 --scenario copy-worksheet-package --skip-legacy-epplus --warmup 1 --iterations 3
```

## Website data

After a comparison-suite run, refresh website/blog benchmark data with:

```powershell
.\Build\Generate-ExcelBenchmarkWebsiteData.ps1 -SummaryPath .\Docs\benchmarks\comparison-current\officeimo.excel.comparison-summary.json -ManifestPath .\Docs\benchmarks\comparison-current\officeimo.excel.comparison-suite-manifest.json -RunMode quick
```
