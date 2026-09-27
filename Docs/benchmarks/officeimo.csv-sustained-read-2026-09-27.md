# Sustained CSV projection and sampled memory (2026-09-27)

Incremental input uses substantially less sampled managed heap than snapshot
input for this six-field workload. Sequential typed projection is faster than
four-worker projection on both measured processor domains. These measurements
support keeping parallel conversion opt-in and measuring the application's
conversion cost before selecting it.

## Work and validation

The public snapshot and incremental APIs read the same generated UTF-8 files.
Each row contains an integer ID, distinct name and note strings, a decimal
score, an ISO date/time, and a boolean. Notes are either plain quoted text or
Unicode text containing CRLF, commas, and escaped quotes. There are no null
fields. Input sizes are 100,000 and 1,000,000 rows.

Input creation uses StreamWriter rather than an OfficeIMO writer. Before
timing, both readers must return every expected field in source order.
Measured typed operations construct the same model, consume every field
through an ordered checksum, verify the row count, and retain no result array.
Checksum validation runs after timing. FirstRow opens the reader, awaits
ReadAsync, and consumes all six strings from one row. TypedSequential and
TypedParallel open asynchronously, then enumerate synchronously. This is
**not full asynchronous traversal qualification**.

PowerForge owns warmup, timing, rotation, forced managed collection before each
operation, output validation, summaries, and artifacts. There are two warmups
and seven measured samples for each of 24 cases per lane. All outliers remain.
The two memory and two timing runs contain **672 successful samples** and no
validation failures. Files are read from the warmed operating-system cache.

## Environment and provenance

Windows 11 Pro build 26200, x64, PowerShell 7.6.6, and .NET 10.0.12.
The AMD Ryzen 9 9950X3D2 has 16 cores and 32 logical processors. A fresh native
cache inventory reports two 96 MiB L3 domains:

- A: processor mask `0x0000FFFF`.
- B: processor mask `0xFFFF0000`.

Every run uses Normal process priority and one declared domain. Typed parallel
projection uses four workers and 1,024-row batches. The shared workstation has
background load; this is a workload-specific observation, not a universal
library ranking or a latency service-level guarantee.

The reader implementation is commit
`99283c500ba04d106e1874fcbb0e84e77ed4c905`. The benchmark workload is committed
in `1c8a6884ddd1cf96568476a6b06fa282e030bb2c`. Domain A was measured before the
benchmark commit with a dirty qualification checkout; domain B was measured
from the frozen clean commit. The timed workload and reader binaries are
identical across the accepted runs. Metadata declarations were expanded
between runs. The workload SHA-256 is
`C7A2F3618E4E00351E8F824F797763941AEC67BCC54D99F5E7FE17AD0CDC48E0`.

The PowerForge probe owner is commit
`60a5e52fa661e61cd862bb60a3900d0ef1101986`. Its six focused probe tests and
460 benchmark service tests pass. The engine builds without warnings for
net472, net8.0, and net10.0. Direct public-API smoke checks also succeed under
Windows PowerShell/.NET Framework and .NET 10.0.11 in WSL/Linux; those checks
do not establish portable CSV performance.

| Binary | SHA-256 |
| --- | --- |
| OfficeIMO.CSV | `B0B03F6933DE74915B3A6A042D6451D1EBCF6D1BCFC9C3523B72A5F4D6810369` |
| OfficeIMO.Core | `DA49BDA169116AB83F8A2727F2820BD919F79FC65AA3CFEAEB5AA677A56AC6C2` |
| PowerForge | `517B9882C5FD10D918DEFAE23291CD75D7211F556F174AD9C143DB1B3439E4E1` |
| PowerForge.PowerShell | `28633E6FDA1AA665A1E0FC24748CE3D693B891F6ECDA8239195D5869565968F5` |

The final driver verifies the actually loaded CSV and Core assemblies against
the selected files and records the CSV, Core, probe, runner, and workload
hashes. Earlier domain-A reports record fewer of these metadata fields;
the unchanged workload and selected binary hashes accompany this evidence.

## Uninstrumented typed traversal

Values are median milliseconds from seven retained samples. Timing includes
one PowerShell operation dispatch as well as reader initialization, typed
projection, checksum consumption, and reader disposal.

| Rows | Notes | Reader | Sequential A ms | Parallel A ms | Sequential B ms | Parallel B ms |
| ---: | --- | --- | ---: | ---: | ---: | ---: |
| 100,000 | Plain | Snapshot | 96.56 | 147.33 | 88.81 | 143.06 |
| 100,000 | Plain | Incremental | 81.13 | 117.32 | 82.13 | 115.54 |
| 100,000 | Multiline | Snapshot | 177.18 | 223.46 | 157.84 | 206.13 |
| 100,000 | Multiline | Incremental | 148.86 | 191.23 | 147.60 | 184.09 |
| 1,000,000 | Plain | Snapshot | 1368.23 | 1864.65 | 1261.97 | 1679.49 |
| 1,000,000 | Plain | Incremental | 827.59 | 1184.36 | 808.01 | 1147.56 |
| 1,000,000 | Multiline | Snapshot | 2238.31 | 2512.00 | 1843.51 | 2267.32 |
| 1,000,000 | Multiline | Incremental | 1487.12 | 1967.55 | 1433.39 | 1857.78 |

At one million rows, four-worker incremental projection is about 30–43%
slower than sequential projection. Conversion and checksum work in this
workload are inexpensive enough that scheduling and snapshotting batches
outweigh their benefit. This does not establish how a custom CPU-heavy
converter behaves.

## Sampled managed heap

These are medians of each sample's **increase from its own managed baseline
to its observed peak**, in MiB. They are not total allocation or isolated
library footprint. Managed baselines range from 33.3 to 36.2 MiB after the
runner's collection step.

| Rows | Notes | Reader | Sequential A MiB | Parallel A MiB | Sequential B MiB | Parallel B MiB |
| ---: | --- | --- | ---: | ---: | ---: | ---: |
| 100,000 | Plain | Snapshot | 92.93 | 96.19 | 92.98 | 96.72 |
| 100,000 | Plain | Incremental | 30.94 | 43.44 | 40.47 | 43.53 |
| 100,000 | Multiline | Snapshot | 121.94 | 125.35 | 122.01 | 126.38 |
| 100,000 | Multiline | Incremental | 35.20 | 45.80 | 40.19 | 45.84 |
| 1,000,000 | Plain | Snapshot | 598.69 | 600.48 | 599.90 | 600.10 |
| 1,000,000 | Plain | Incremental | 45.03 | 47.03 | 45.14 | 48.65 |
| 1,000,000 | Multiline | Snapshot | 711.43 | 715.46 | 710.99 | 715.48 |
| 1,000,000 | Multiline | Incremental | 45.77 | 49.56 | 45.76 | 50.22 |

Ten times as many rows increases the incremental typed median from about
31–46 MiB to about 45–50 MiB. Snapshot typed projection grows from about
93–126 MiB to about 599–715 MiB. Incremental projection shows much smaller
sampled heap growth at these two sizes; these points do not prove a bound for every dialect,
schema, field width, or consumer.

The separate first-row observations show the cost of snapshot initialization:

| Rows | Notes | Snapshot A / B increase MiB | Incremental A / B increase KiB |
| ---: | --- | ---: | ---: |
| 100,000 | Plain | 86.71 / 86.85 | 40.04 / 45.63 |
| 100,000 | Multiline | 115.81 / 115.88 | 45.87 / 46.19 |
| 1,000,000 | Plain | 557.67 / 557.58 | 40.12 / 46.28 |
| 1,000,000 | Multiline | 687.50 / 681.51 | 45.55 / 45.63 |

PowerForge samples every requested five milliseconds, with an immediate
baseline and a final observation. Scheduler delays can lengthen intervals,
and brief transients can be missed, so peaks are lower bounds. Probe startup
and sampling consume resources. Managed heap includes uncollected garbage.

The memory lane is instrumented and its timings are not used in the
throughput table. In particular, a short first-row operation can complete
with only baseline and final observations. Use the BenchmarkDotNet async
lane for precise first-row API timings; PowerShell dispatch dominates these
short operations.

Resident baselines range from 226.5 to 1,117.6 MiB because the same process
reuses committed pages after larger snapshot cases. Raw resident baselines
and peaks are retained, but these runs do not establish a clean per-operation
resident-memory budget. Process-isolated observations are needed for that
decision.

## Reproduction and retained evidence

Build OfficeIMO.CSV and the owning PSPublishModule checkout for net10.0.
Set POWERFORGE_ROOT to that tool checkout and use a fresh PowerShell process.
On this machine the commands are:

```powershell
$modulePath = Join-Path $env:POWERFORGE_ROOT 'PSPublishModule/bin/Release/net10.0/PSPublishModule.dll'
foreach ($mask in '0xFFFF', '0xFFFF0000') {
    pwsh -NoProfile -File ./Build/Benchmarks/Run-CsvSustainedReadBenchmarks.ps1 -ModulePath $modulePath -AffinityMask $mask -SampleMemory -OutputRoot "./Ignore/Benchmarks/Memory-$mask"
    pwsh -NoProfile -File ./Build/Benchmarks/Run-CsvSustainedReadBenchmarks.ps1 -ModulePath $modulePath -AffinityMask $mask -OutputRoot "./Ignore/Benchmarks/Time-$mask"
}
```

Discover the machine's topology before selecting masks. Keep source unchanged
during measurement; the shared runner rejects changed source provenance.
Default row counts, worker count, batch size, warmups, and iterations match
this run. The [benchmark README](../../OfficeIMO.CSV.Benchmarks/README.md#sustained-input-and-sampled-memory)
describes overrides and interpretation.

Accepted local artifact sets are retained under Ignore/Benchmarks:

| Lane | Folder | Run |
| --- | --- | --- |
| Memory A | CsvSustainedMemoryDomain1 | 20260927-185034-67209443 |
| Time A | CsvSustainedTimeDomain1 | 20260927-185645-08d1163e |
| Memory B | CsvSustainedMemoryDomain2 | 20260927-190910-ee7246bc |
| Time B | CsvSustainedTimeDomain2 | 20260927-191549-f0201678 |

Each set retains raw samples, shared summaries, environment metadata,
comparisons, and a fixture hash manifest. The four input files are
7,993,725, 11,593,725, 83,935,729, and 119,935,729 bytes; hashes match across
runs. Generated inputs are removed after qualification. A prior second-domain
attempt was rejected by the source-provenance guard and discarded; it is not
included in these figures.

Full ReadAsync traversal, larger-input precise first-row API latency,
process-isolated resident budgets, and Linux/macOS performance remain in the
[product roadmap](../ROADMAP.md#spreadsheet-and-csv-delivery-order).
