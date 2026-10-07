# Reproduce the broad Excel/CSV allocation comparison

The source baseline is `0d0d5a0b6e7bf048accdf8bfc8ca585a3e62e599` and the final
candidate is `8b84a1676c31db2d92a8b80ede7064c72748e554`. Use separate checkouts
and an explicit task-owned scratch directory. The measurements used SDK
10.0.112, .NET 10.0.12, and BenchmarkDotNet 0.15.8.

## Prepare equivalent snapshots

1. Build both candidate benchmark projects in Release for `net10.0`.
2. Copy each complete benchmark output directory into two scratch folders:
   `baseline-csv`/`candidate-csv` and `baseline-excel`/`candidate-excel`.
3. Build the baseline product projects in Release for `net10.0`. Replace the
   product assemblies in each **baseline** snapshot with the corresponding
   baseline build's `OfficeIMO.*.dll` and PDB files. Keep the candidate's benchmark
   assemblies and comparison dependencies in both snapshots. No production
   reference points to a comparison library.
4. Copy the JSON case files and `rotated.ps1` from this directory to the scratch
   root. Put `Program.cs`, `TraceSummary.cs`, and `Compare.csproj.template` in a
   `compare` subdirectory, renaming the template to `Compare.csproj`. Build it in
   Release. Its only package reference is BenchmarkDotNet.
5. Set the environment below and run `--validate` before measurement. Verify
   the processor/cache topology and choose suitable affinity masks for the
   actual host. The recorded Windows masks are not portable defaults.

```powershell
$env:OFFICEIMO_PERFORMANCE_SNAPSHOTS = $taskRoot
$env:OFFICEIMO_COMPARISON_CASES = Join-Path $taskRoot 'comparison-cases.json'
$env:OFFICEIMO_BENCHMARK_DATA = Join-Path $taskRoot 'fixtures'
$env:OFFICEIMO_COMPARISON_MASK = '65535'
dotnet "$taskRoot/compare/bin/Release/net10.0/Compare.dll" --validate
dotnet "$taskRoot/compare/bin/Release/net10.0/Compare.dll" --filter '*' `
  --warmupCount 16 --iterationCount 10 --invocationCount 4 --unrollFactor 1 `
  --outliers DontRemove --artifacts "$taskRoot/before-after"
```

The benchmark classes' setup methods validate complete field values, row counts,
headers, and applicable output semantics before timing. For write/copy cases,
they reopen complete generated packages. The snapshot driver's `ToString`
comparison is only an additional completion/checksum check; it does not establish
byte-array equality or replace those setup validators.

## Case groups and stages

| Case file | Contract | Recorded warmup / iterations / operations |
| --- | --- | --- |
| `comparison-cases.json` | CSV sync/async and 25K shared-string reads | 16 / 10 / 4 |
| `copy-cases.json` | Package and values-only copy/save, 100/2,500/25,000 rows | 8 / 8 / 1 |
| `prefix-comparison-cases.json` | Five shared-string shapes, two sizes, ordinary/prefixed worksheets | 8 / 8 / 4 |
| `values-copy-check-cases.json` | Longer 2,500-row values-copy allocation check | 32 / 12 / 2 |
| `encoding-cases.json` | Rejected encoding experiment, 13 export workloads | 16 / 12 / 2 |

The raw packet keeps intermediate stages. CSV stage 1 is identical to the final
CSV source. Excel stage 1 adds only the shared-string cache changes to the
baseline. Stage 2 adds package-copy/reference and inline-string serialization
changes; stage 3 adds the prefixed reader and XML correctness fixes. Their binary
hashes are recorded in `../provenance.json`. Running every case with the final
candidate measures the final implementation, not those intermediate binaries.

`rejected-encoding-experiment.patch` records the removed nine-line experiment.
Apply it only in a disposable candidate checkout to reproduce that stage. Do
not apply it to the accepted candidate's measurement or correctness run.

## Rotated comparisons and sampled memory

Use a fresh PowerShell process for each processor group. The driver calls the
canonical PowerForge suite runner; it does not implement its own stopwatch,
rotation, warmup, or outlier policy. The retained reproduction adds a final
failed-sample guard; the original measurements were checked explicitly after
each run and all retained samples succeeded.

```powershell
pwsh -NoProfile -File "$taskRoot/rotated.ps1" -IdenticalBaseline `
  -Cases rotated-cases.json -Label aa -Mask 65535
pwsh -NoProfile -File "$taskRoot/rotated.ps1" `
  -Cases prefix-comparison-cases.json -Label prefix -Mask 65535
pwsh -NoProfile -File "$taskRoot/rotated.ps1" `
  -Cases prefix-comparison-cases.json -Label prefix -Mask 4294901760
pwsh -NoProfile -File "$taskRoot/rotated.ps1" `
  -Cases regression-cases.json -Label regression -Mask 65535 -Operations 2
```

Rotated runs retain 24 measured samples per engine after 12 warmups, with four
operations per sample except the two-operation regression follow-up. Run the
same cases on both processor groups and retain unfavorable/control results.
The two recorded regression groups ran sequentially in one PowerShell host;
fresh-host reproduction also avoids carrying that host's retained assembly state.

The memory lane uses the repository's existing
`Build/Benchmarks/Run-CsvSustainedReadBenchmarks.ps1`. Run it in separate fresh
PowerShell processes for baseline and candidate with `-Rows 100000,1000000`,
`-Operation FirstRow,AllRowsAsync`, both engines and shapes, `-WarmupCount 1`,
`-IterationCount 3`, `-AffinityMask 0xFFFF0000`, and `-SampleMemory`. Supply the
appropriate snapshot through `-BinaryRoot` and a separate `-OutputRoot`.

The measured installed module supported rotated timing but lacked
`PowerForge.BenchmarkMemoryProbe`. For memory sampling, `-ModulePath` selected
an existing build of that canonical owner; its binary hashes and checkout head
are recorded in provenance. Build the owning PowerForge source when a containing
published module is unavailable. Do not introduce a replacement local sampler
or turn the observed four-part local build version into a public dependency pin.

Sampling every 5 ms gives lower-bound peaks. It does not measure retained live
objects, and its thread startup dominates very short first-row operations. Keep
those instrumented timings separate from throughput measurements. Delete only
the exact generated fixture directories after retaining the result JSON.

The native benchmark entrypoints accept `--affinityMasks` and `--priority`.
Do not combine `--inProcess` with the explicit affinity jobs: BenchmarkDotNet
adds a separate in-process job instead of applying that toolchain to the existing
jobs. The aborted export run using that combination is excluded from this packet.

Correctness commands and results are recorded in `../validation.json`. Keep
measurement suites opt-in and outside ordinary correctness gates. After retaining
compact evidence, remove disposable traces, generated fixtures, and obsolete
snapshot/build directories without deleting shared caches or another task's files.
