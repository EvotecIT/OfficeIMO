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

On an AMD Ryzen 9 9950X3D2, use `0xFFFF` and `0xFFFF0000` to measure the two
complete 16-logical-processor cache domains separately. A mask such as `0x1`
measures one logical processor and is not a substitute for a domain-constrained
parallel run.

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

The public forward-only mapping benchmarks keep parser and projection behavior
separate. `ExcelPublicParallelReadBenchmarks` measures cheap automatic property
mapping, where scheduling may cost more than it saves. The
`ExcelPublicParallelProjectionBenchmarks` crossover lane reads through the same
native XLSX, XLSB, or XLS reader in both methods and performs the same validated,
CPU-heavy row projection; only sequential versus ordered-parallel projection
changes. Both lanes enable bounded schema inference, and the CPU-heavy lane
fails setup unless at least two projection workers actually overlap:

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --filter "*ExcelPublicParallelReadBenchmarks*" --affinityMasks "0xFFFF,0xFFFF0000" --priority High --invocationCount 1 --unrollFactor 1 --warmupCount 5 --iterationCount 10 --launchCount 1 --outliers DontRemove
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --filter "*ExcelPublicParallelProjectionBenchmarks*" --affinityMasks "0xFFFF,0xFFFF0000" --priority High --invocationCount 1 --unrollFactor 1 --warmupCount 5 --iterationCount 10 --launchCount 1 --outliers DontRemove
```

Parsing and decompression stay single-owner in both lanes. The parallel API is
for independent projection work; these benchmarks must not be cited as evidence
that the underlying workbook parser runs concurrently.

### Dated ordered-parallel crossover snapshot (2026-08-09)

The commands above were run on .NET 10 with workstation GC, `High` process
priority, retained outliers, five warmups, ten fixed-work iterations, and each
16-logical-processor L3 domain measured as a separate job. The cheap automatic
mapping lane did not establish a general parallel win: XLSX shared rank 1 on
both domains, XLSB ranged from a 4% lower parallel mean to an unstable 83%
higher mean, and XLS ranged from an unresolved noisy result to a 10% lower
parallel mean. Parallel mapping allocated 11% more for XLSX and 18% more for
XLS/XLSB. Keep cheap mappings sequential unless a representative benchmark
shows otherwise.

With 4,096 deterministic mixing rounds in each row projection, ordered-parallel
mapping produced the following means. Setup verified complete output and failed
unless at least two projection workers overlapped.

| Format | L3 domain 0 sequential / parallel | L3 domain 1 sequential / parallel | Parallel allocation cost |
| --- | ---: | ---: | ---: |
| XLSX | 806.11 / 605.86 ms | 794.71 / 568.93 ms | 1.04-1.05x |
| XLSB | 278.05 / 66.49 ms | 277.80 / 89.79 ms | 7.30-7.33x |
| XLS | 278.27 / 50.26 ms | 270.38 / 49.38 ms | 8.15-8.28x |

This is a crossover result, not a claim that ordinary row mapping is CPU-heavy.
The XLSX reader's decompression and XML work dominates more of the total time,
so its 25-28% reduction is smaller than the 68-82% reductions for XLSB and XLS.
The binary formats also make the snapshot and scheduling allocation tradeoff
especially visible. Select parallel mapping for sufficiently expensive,
independent projection work and measure the real consumer workload.

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

Write scenarios keep different capability contracts in separate lanes. The
`write-insertobjects-flat-dictionaries-direct` lane measures editable worksheet
imports, while `write-flat-dictionaries-direct-package` measures forward-only
package writers starting from the same prepared dictionaries. The
`append-plain-rows` lane measures coordinate-cell APIs and therefore excludes
row-only streaming libraries. Input projection is prepared before the timed
delegate unless projection is part of the named contract. Every implementation
must pass an untimed Open XML and semantic-cell equivalence preflight before
its measurements are accepted.

`ExcelExecutionMode.Parallel` permits parallel compute; it does not require
worker fan-out when a specialized single-pass path is faster. In particular,
the rectangular `CellValues` package path snapshots values sequentially. A
focused experiment on the 25,000-row four-column workload made the parallel
snapshot slower, so the benchmark retains the faster production behavior.

### Dated 25K compact DataReader package snapshot (2026-08-10)

This lane writes the same prepared 25,000-row, eight-column `DbDataReader` to a
new XLSX package through each library's forward-only typed-row API. OfficeIMO
uses `ExcelDocument.WriteDataReader` with shared strings and explicit cell
references disabled. The untimed preflight validates the complete semantic
cell grid before any timings are accepted.

The .NET 10 runner used High process priority, twelve warmups, 31 retained
measurements, and three independent processes on each of two logical CPUs from
separate L3 domains. Execution order was rotated within each process, and
differences below 5% were classified as ties.

| CPU | OfficeIMO medians | SpreadCheetah medians | Sylvan medians | LargeXlsx medians | OfficeIMO outcomes (9 comparisons) |
| --- | ---: | ---: | ---: | ---: | --- |
| CPU 0 (`0x1`) | 21.85, 19.16, 20.57 ms | 21.10, 20.97, 19.82 ms | 27.71, 29.41, 26.37 ms | 23.33, 24.29, 23.13 ms | 7 wins, 2 ties |
| CPU 16 (`0x10000`) | 14.41, 15.30, 14.99 ms | 15.40, 15.39, 15.45 ms | 20.15, 21.44, 20.71 ms | 17.68, 18.72, 17.82 ms | 7 wins, 2 ties |

OfficeIMO was fastest or within the 5% tie margin in every launch-library
comparison: six wins against Sylvan, six against LargeXlsx, and two wins plus
four ties against SpreadCheetah. It also produced the smallest package at
907,354 bytes, compared with 934,903 for LargeXlsx, 978,999 for SpreadCheetah,
and 997,938 for Sylvan. SpreadCheetah and LargeXlsx retained the managed
allocation lead at approximately 5.62-5.67 MiB per export; OfficeIMO's median
was 6.18 MiB and Sylvan's was 7.52 MiB.

Use the measured compact contract for a plain SQL-style export:

```csharp
ExcelDocument.WriteDataReader(output, reader, new ExcelTabularWriteOptions {
    IncludeCellReferences = false,
    UseSharedStrings = false
});
```

The default writer deliberately keeps shared strings and explicit references.
It is a different fidelity contract, and options such as table creation or
auto-fit also require richer processing. Reproduce one launch per CPU below;
repeat each command in three independent processes for the table above.

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare .\Ignore\Benchmarks\datareader-compact-cpu0.json --rows 25000 --scenario write-datareader-compact-package --library OfficeIMO.Excel --library Sylvan.Data.Excel --library SpreadCheetah --library LargeXlsx --skip-legacy-epplus --warmup 12 --iterations 31 --affinity 0x1 --priority High
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj --no-build -- compare .\Ignore\Benchmarks\datareader-compact-cpu16.json --rows 25000 --scenario write-datareader-compact-package --library OfficeIMO.Excel --library Sylvan.Data.Excel --library SpreadCheetah --library LargeXlsx --skip-legacy-epplus --warmup 12 --iterations 31 --affinity 0x10000 --priority High
```

The measured binaries were
`OfficeIMO.Excel.Benchmarks.dll` SHA-256
`482816511FB81853B65B1C6A39767F81849ECFF132A2D597D921A5E6249CEEC4` and
`OfficeIMO.Excel.dll` SHA-256
`AD593FF46DFCB175635422C4B74AD42C7660D167F225F543F08BFA06234CA856`.

### Dated 25K CellValues rectangle snapshot (2026-08-10)

This lane writes the same 25,000-row, eight-column sales rectangle and validates
the complete semantic cell grid before accepting either implementation. OfficeIMO
receives prepared coordinate/value tuples through the editable `CellValues` API
and takes its required defensive snapshot. LargeXlsx receives the corresponding
typed rows through its forward-only writer with cell references disabled. This
is an equivalent-output product comparison, not an identical-input
microbenchmark.

The .NET 10 runner used High process priority, twelve warmups, 31 retained
measurements, and three independent processes for each 16-logical-processor L3
domain. OfficeIMO's large contiguous rectangle path writes inline strings and
standards-valid implicit data-cell coordinates; small rectangles retain explicit
references and the existing shared-string policy. Differences below 5% are
classified as ties.

| CPU domain | OfficeIMO medians | LargeXlsx medians | Launch outcomes |
| --- | ---: | ---: | --- |
| Domain 0 (`0xFFFF`) | 15.18, 13.39, 12.24 ms | 15.16, 14.09, 11.91 ms | 1 OfficeIMO win, 2 ties |
| Domain 1 (`0xFFFF0000`) | 10.91, 13.43, 13.78 ms | 11.14, 14.95, 15.00 ms | 2 OfficeIMO wins, 1 tie |

The mean-time classification produced four OfficeIMO wins and two ties across
the same launches; no launch favored LargeXlsx by 5% or more. OfficeIMO produced
a 907,405-byte package versus LargeXlsx's 934,903-byte package. LargeXlsx kept
the allocation advantage at 3,072.0 KB versus OfficeIMO's 5,543.0 KB because
the editable `CellValues` contract snapshots caller-owned values while the
LargeXlsx comparison streams typed rows. Reproduce one launch per domain with:

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare .\Ignore\Benchmarks\cellvalues-domain0.json --rows 25000 --scenario write-cellvalues-rectangle-direct --library OfficeIMO.Excel --library LargeXlsx --skip-legacy-epplus --warmup 12 --iterations 31 --affinity 0xFFFF --priority High
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj --no-build -- compare .\Ignore\Benchmarks\cellvalues-domain1.json --rows 25000 --scenario write-cellvalues-rectangle-direct --library OfficeIMO.Excel --library LargeXlsx --skip-legacy-epplus --warmup 12 --iterations 31 --affinity 0xFFFF0000 --priority High
```

### Dated 25K dictionary-stream snapshot (2026-08-10)

The forward-only dictionary lane was run in three independent processes per
CPU domain on .NET 10, using twelve warmups, 31 retained measurements, High
process priority, and the same 25,000 prepared dictionaries for both writers.
Differences below 5% are classified as ties.

| CPU domain | OfficeIMO medians | LargeXlsx medians | Launch outcomes |
| --- | ---: | ---: | --- |
| CPU 0 (`0x1`) | 24.60, 23.04, 23.26 ms | 24.68, 24.80, 25.17 ms | 2 OfficeIMO wins, 1 tie |
| CPU 16 (`0x10000`) | 20.12, 20.11, 19.78 ms | 22.19, 21.27, 22.07 ms | 3 OfficeIMO wins |

OfficeIMO allocated 7,149.1 KB per invocation and LargeXlsx allocated
6,031.6 KB. The timing evidence therefore favors OfficeIMO in five launches
and ties one, while LargeXlsx retains the allocation advantage. Reproduce one
launch per domain with:

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare .\Ignore\Benchmarks\dictionary-stream-cpu0.json --rows 25000 --scenario write-flat-dictionaries-direct-package --library OfficeIMO.Excel --library LargeXlsx --skip-legacy-epplus --warmup 12 --iterations 31 --affinity 0x1 --priority High
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj --no-build -- compare .\Ignore\Benchmarks\dictionary-stream-cpu16.json --rows 25000 --scenario write-flat-dictionaries-direct-package --library OfficeIMO.Excel --library LargeXlsx --skip-legacy-epplus --warmup 12 --iterations 31 --affinity 0x10000 --priority High
```

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
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --filter "*MarkPflug65KXlsxBenchmarks*" --affinityMasks "0x1,0x10000" --priority High --invocationCount 1 --unrollFactor 1 --warmupCount 5 --iterationCount 15 --launchCount 1 --outliers DontRemove
```

These suites compare library methods within each CPU job, so method baselines
and absolute per-job results are the appropriate shape. Do not combine those
method baselines with an `--apples` job-baseline comparison.

### Dated 65K XLSX snapshot (2026-08-10)

The fixed-work command above was run locally on .NET 10 with workstation GC,
High process priority, the Windows High performance power plan, and an AMD
Ryzen 9 9950X3D2. CPU 0 and CPU 16 are the first logical processor in each
measured L3 domain. Every method read the same hash-pinned workbook and passed
the row, cell, and payload observation before measurement.

| Library | CPU 0 mean (median) | CPU 16 mean (median) | Managed allocation |
| --- | ---: | ---: | ---: |
| OfficeIMO.Excel | 248.4 ms (240.4 ms) | 176.1 ms (185.0 ms) | 366.95-368.04 KB |
| Sylvan.Data.Excel | 473.6 ms (435.6 ms) | 556.4 ms (701.7 ms) | 650.28 KB |
| ExcelDataReader | 844.8 ms (706.6 ms) | 714.8 ms (543.9 ms) | 207,509.89-207,509.97 KB |
| ClosedXML | 1,502.9 ms (1,475.2 ms) | 1,233.5 ms (1,212.0 ms) | 725,534.91-725,536.56 KB |
| EPPlus | 1,379.2 ms (1,132.7 ms) | 862.8 ms (836.4 ms) | 864,345.25-864,434.83 KB |
| MiniExcel | 747.5 ms (602.2 ms) | 623.1 ms (483.6 ms) | 666,370.91-666,372.80 KB |

OfficeIMO has the lowest observed mean and median within both affinity jobs. It
also has the lowest managed allocation by a wide margin. Sylvan,
ExcelDataReader, EPPlus, and MiniExcel changed performance phase during their
isolated processes. Sylvan's broad 99.9% confidence intervals overlap
OfficeIMO on both domains; the run therefore establishes the observed timing
and allocation lead, not a universal statistical separation from every
reader. Outliers were retained, and the full distributions matter more than
selecting one favorable cluster.

The measured assemblies were SHA-256
`1EE9FA523E390F5610352798720D633CB1DA23C4D4DCCDBE7F15777874992CA0`
for `OfficeIMO.Excel.Benchmarks.dll` and
`3D7C721A2DFF57B1505C453A256594C027E13D88F9EB4C6DD2810F75603D89D7`
for `OfficeIMO.Excel.dll`.

### Dated ExcelReader.NET 2.3.0 snapshot (2026-08-31)

This candidate was also compared with ExcelReader.NET 2.3.0 on .NET 10 using
High process priority, workstation GC, the Windows High performance power plan,
and both complete 16-logical-processor cache domains of the same AMD Ryzen 9
9950X3D2 (`0xFFFF` and `0xFFFF0000`). The paired runners use symmetric ABBA
ordering, retain every sample, validate the complete typed read observation, and
validate the generated artifact outside the timed write operation.

For the equivalent 25,000-row XLSX write, OfficeIMO had the lower median on both
domains:

| Cache domain | OfficeIMO median | ExcelReader.NET median | OfficeIMO / ExcelReader.NET |
| --- | ---: | ---: | ---: |
| `0xFFFF` | 4.676 ms | 7.447 ms | 0.6279 |
| `0xFFFF0000` | 4.482 ms | 7.341 ms | 0.6105 |

The native binary write probes deliberately keep non-equivalent competitor
output visible instead of dropping it. ExcelReader.NET's XLS round-tripped its
own values but contained none of the required BIFF8 `DBCell` blocks; its XLSB
round-tripped its own values but omitted the required `BrtWsDim` record. The
runner records those semantic and structural observations and artifact sizes,
but withholds paired timings and ratios until both implementations perform
equivalent work. A future conforming competitor release automatically enters
the timed lane. OfficeIMO setup fails unless every expected BIFF8
`Index`/`DBCell` block is present, so an OfficeIMO regression cannot make this
lane look faster by silently weakening the artifact.

The equivalent 65K read lanes found XLSX parity with prefetch disabled and an
OfficeIMO median about 2% lower on both domains. OfficeIMO's XLSB median was
15-18% lower. XLS stayed close but favored ExcelReader.NET by 2-4% at the
median. Enabling the experimental bounded XLSX worksheet prefetch made this
fixture 19% slower on both domains. Prefetch therefore remains opt-in and
disabled by default; the negative result is retained instead of being presented
as a win.

The measured assemblies were SHA-256
`A76991AE9336A4CFB8A9D6A20F6825C7A6E68896F41FB637C6CEA7F1BE4745F3`
for `OfficeIMO.Excel.Benchmarks.dll` and
`AF8FFC2832A9F19A974D38A26660975B859C5BD513487A2C27DC5A5421CA5374`
for `OfficeIMO.Excel.dll`.

Reproduce the paired lanes on each cache-domain mask with:

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --compare-excelreader-xlsx-paired 40 0xFFFF High
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --compare-excelreader-xlsx-paired 40 0xFFFF High prefetch
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --compare-excelreader-xlsb-paired 40 0xFFFF High
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --compare-excelreader-xls-paired 40 0xFFFF High
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --compare-excelreader-write-paired Xlsx 25000 30 0xFFFF High 8
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --compare-excelreader-write-paired Xlsb 25000 30 0xFFFF High 8
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --compare-excelreader-write-paired Xls 25000 30 0xFFFF High 8
```

Repeat with `0xFFFF0000` for the second cache domain. A one-logical-processor
mask is not equivalent to either domain-constrained parallel measurement.

### Dated 65K XLS and XLSB snapshot (2026-08-10)

The legacy-format runner uses twelve warmups followed by eighty symmetric ABBA
samples at High process priority. Each sample averages two reads per library,
alternates which library runs outside the pair, and rejects any row, cell, or
payload observation mismatch. These results used the same workstation and CPU
domains as the XLSX snapshot above.

| Format | CPU | OfficeIMO median | Sylvan median | Ratio of medians | Paired ratio median (P25-P75) |
| --- | --- | ---: | ---: | ---: | ---: |
| XLS | CPU 0 | 28.334 ms | 27.533 ms | 1.0291 | 1.0238 (0.9679-1.0869) |
| XLS | CPU 16 | 24.271 ms | 23.054 ms | 1.0528 | 1.0459 (0.9845-1.0955) |
| XLSB | CPU 0 | 48.187 ms | 46.466 ms | 1.0370 | 1.0308 (0.9837-1.0804) |
| XLSB | CPU 16 | 37.128 ms | 37.433 ms | 0.9919 | 0.9879 (0.9729-1.0173) |

Three ratio-of-medians results are within the repository's predeclared 5%
threshold. XLS on CPU 16 is borderline: its ratio of medians is 1.0528 while
the paired ratio median is 1.0459 and the interquartile range crosses parity.
The snapshot therefore establishes near parity, not an OfficeIMO timing win.
OfficeIMO had the lower observed median only for XLSB on CPU 16.

The timed observer folds every typed value into the deterministic checksum one
word at a time. This preserves row, cell, type, order, and payload validation
without making byte-at-a-time checksum bookkeeping the dominant workload. A
separate two-reader BenchmarkDotNet run with four invocations, twelve warmups,
twenty retained-outlier iterations, and both affinity jobs measured 115.15 KB
for OfficeIMO and 343.71-343.87 KB for Sylvan. Its isolated method processes
changed performance phase and produced multimodal timing distributions, so the
ABBA runner above is the relative timing evidence; that BenchmarkDotNet run is
used only for allocation evidence.

Reproduce both timing domains and the allocation run with:

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --compare-markpflug65k-xls-paired 80 0x1 High
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --compare-markpflug65k-xls-paired 80 0x10000 High
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --compare-markpflug65k-xlsb-paired 80 0x1 High
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --compare-markpflug65k-xlsb-paired 80 0x10000 High
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --filter "*MarkPflug65KXlsbBenchmarks.OfficeIMO" "*MarkPflug65KXlsbBenchmarks.Sylvan" --affinityMasks "0x1,0x10000" --priority High --invocationCount 4 --unrollFactor 1 --warmupCount 12 --iterationCount 20 --launchCount 1 --outliers DontRemove
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

The `write-datareader-table` lane requires a real worksheet table and
AutoFilter in addition to equivalent cell values. The AutoFit variant also
requires custom column widths. A library that only styles the cell range is
excluded from these lanes instead of being credited with the cheaper contract.

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
