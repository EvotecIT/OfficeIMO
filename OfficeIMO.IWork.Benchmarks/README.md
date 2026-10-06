# iWork runtime evidence

This opt-in workload library measures the shared iWork reader and the Word, Excel
and PowerPoint destination adapters. It stays outside the normal solution;
deterministic native-content and cancellation checks run in a dedicated workflow.
[PowerForge](https://github.com/EvotecIT/PSPublishModule) owns
warmups, rotated iteration order, elapsed and memory counters, artifacts and gates.
The library owns synthetic inputs, pinned native fixtures and semantic validation.

The [2026-10-04 portable run](../Docs/benchmarks/officeimo.iwork-portable-runtime-2026-10-04.md)
retains 6,390 validated samples on Linux x64, macOS arm64 and Windows x64. It covers
two scale/native passes, 50 repeated operations per native case, separate managed/resident
sampling and active I/O cancellation. Its observations do not establish portable ceilings.

The [2026-10-04 macOS comparison](../Docs/benchmarks/officeimo.iwork-runtime-2026-10-04.md)
retains 360 validated samples from a repeated scale matrix, including slower
elapsed cases, allocation changes and whole-process memory observations. It does
not establish portable performance budgets.

Build from the repository root using the pinned SDK:

```powershell
dotnet build OfficeIMO.IWork.Benchmarks/OfficeIMO.IWork.Benchmarks.csproj -c Release
./Build/Benchmarks/Run-IWorkRuntimeBenchmarks.ps1 -Plan
./Build/Benchmarks/Run-IWorkRuntimeBenchmarks.ps1 -OutputRoot $env:EVOTEC_DEV_TEMP/iwork-runtime
```

Use a fresh PowerShell 7 process. The runner requires PowerForge operation allocation
measurement and its operation-memory sampling policy API. It fails when allocation
counters or requested observations are missing. To validate against shared
source, build `PSPublishModule/PSPublishModule.csproj -c Release -f net8.0` in the
PowerForge checkout with its pinned SDK, then pass the built `PSPublishModule.dll`
with `-ModulePath`. `-BinaryRoot` selects the matching OfficeIMO workload and owner
assemblies. Keep other builds and benchmarks idle and record host power mode and
processor placement. Defaults use inherited placement; macOS affinity is unqualified.

| Family | Small | Medium | Large | Verified content |
| --- | ---: | ---: | ---: | --- |
| Pages | 100 | 1,000 | 10,000 | Every paragraph's text and order |
| Numbers | 1,000 | 10,000 | 100,000 | Every numeric cell's coordinate and value, dimensions and sparse count |
| Keynote | 10 | 100 | 1,000 | Every slide title's text and order |

`LoadProject` opens a ZIP and materializes its selected projection. `ConvertSave`
opens, converts and saves through the normal destination API. Input construction
and full semantic readback are outside measurement. Synthetic conversion requires a complete
editable reconstruction report; saved DOCX, XLSX and PPTX packages are reopened and
checked before a sample succeeds. `-Scale`, `-Kind` and `-Operation` select cases;
`-WarmupCount` and `-IterationCount` default to two and five. A quick smoke run uses
`-Scale Small -WarmupCount 0 -IterationCount 1`.

Use `-Scale Native` for `nim-iwork/simple.pages`, `nim-iwork/simple.numbers` and
`native-exports/keynote-colors-v15.4.key` in the [licensed corpus](../OfficeIMO.TestAssets/Documents/IWorkCorpus/README.md).
Their SHA-256 hashes are checked before execution. Pages validates all three body
paragraphs, including the empty paragraph; Numbers validates the complete 3×3 grid,
coordinates, value types and contents. Both operations use the same reader and
editable-conversion owners as the synthetic cases, requiring complete editable
reconstruction for their qualified paragraph defaults. They reject preview fallback
and verify all listed content. The conversion policy is recorded in case variables so these runs
are not compared with earlier strict-policy samples. Saved packages are reopened
before a sample succeeds. Keynote validates both slides, their 1024×768-point canvas,
all title/body text drawables in order and the two opaque background colors against
the controlled Apple export. These small native inputs complement the scale matrix;
they do not establish large native-package budgets or new rendering/font coverage.

Validate the pinned native workload and active cancellation without measuring performance:

```powershell
dotnet run --project OfficeIMO.IWork.Benchmarks/OfficeIMO.IWork.Benchmarks.csproj -c Release -- --validate OfficeIMO.TestAssets/Documents/IWorkCorpus
```

The [iWork runtime evidence workflow](../.github/workflows/iwork-runtime-evidence.yml)
runs this correctness check on Windows, Linux and macOS for workload changes.
Manual dispatch additionally builds a pinned PowerForge source owner, runs the
scale matrix twice, and measures 50 repeated native operations per case and pass
in fresh PowerShell hosts. A separate native pass enables memory sampling.
Measurements stay outside ordinary PR gates.

Artifacts include raw samples, summaries, CSV tables and environment metadata.
Case variables retain deterministic input SHA-256 hashes; metadata retains the
spec, workload and candidate owner assembly hashes. Inspect failures before using
elapsed or allocation results. `AllocatedBytes` includes managed host invocation
and concurrent threads in the same process. `WorkingSetDeltaBytes` is a signed
resident-page difference, not peak or retained memory.

Use `-MeasureRetainedMemory` with a PowerForge source build containing
`BenchmarkManagedMemoryProbe` to record collected managed-heap observations.
Setup releases the previous result and captures a collected baseline; validation
reopens and checks the current output, records its size, releases the projection or
saved bytes, and captures the collected heap again. Collection stays outside timing.
`ManagedBaselineBytes`, `CollectedManagedBytes` and signed
`RetainedManagedDeltaBytes` include live host state and cache changes in the process.
The retained-memory option is recorded in case variables so its runs are distinct
from normal allocation runs. It is off by default and does not need a new runtime
dependency in OfficeIMO packages.

`RepeatedManagedBaselineBytes` retains the first measured setup baseline for each
case/operation. `RepeatedRetainedManagedDeltaBytes` compares later collected heaps
with that fixed baseline, so gradual changes remain visible across iterations.
Use a single case and operation in a fresh host to reduce changes from other cases:

```powershell
./Build/Benchmarks/Run-IWorkRuntimeBenchmarks.ps1 -Scale Native -Kind Keynote -Operation ConvertSave -IterationCount 50 -MeasureRetainedMemory
```

Add `-MemorySamplingIntervalMilliseconds 5` for managed-heap and resident-page
observations during each operation. Zero disables sampling by default; enabled
intervals range from 1 through 1000 milliseconds. Raw metrics retain the baseline,
sampled maximum, maximum-minus-baseline change and sample counts. Sampled maxima are
lower bounds on peaks: short operations may have only the two boundary readings,
and transient peaks can fall between observations. They include observer/host
effects and do not isolate native allocations. The wrapper gives these samples
their own run mode; keep their timing separate from uninstrumented comparisons.

`CancelDuringLoad` and `CancelDuringConvert` request cancellation synchronously
after the caller stream's first nonempty package read. `CancelDuringNativeCopy`
uses an owned one-slide model with 150,000 ASCII characters and requests cancellation
after the first 64 KiB write. Select only `-Scale Native -Kind Keynote -Operation
CancelDuringNativeCopy` for that fixed writer case. Its input hash and size describe
the model's deterministic native encoding. It does not consume a producer fixture.

Every cancellation sample must observe `OperationCanceledException` after real I/O
and leave the caller stream open. Writer copying must stop at 64 KiB. `VerifiedUnits`
is one cancellation boundary in these cases, rather than the fixture's content count.
Metrics retain bytes at the request, total observed I/O and request-to-exception
latency. This qualifies cancellation during package intake and native stream copying;
it does not measure cancellation during projection, destination construction or
native encoding, and it does not qualify atomic path cancellation during staging.

Repeat an equivalent matrix after warmup before selecting a host-specific budget.
Per-iteration deltas do not establish long-run stability, native-memory retention,
peak usage or a portable leak threshold. Keep the native and synthetic workloads
separate when interpreting their results.

Use PowerForge `Test-BenchmarkGate -SummaryPath <summary.json> -BaselinePath
<baseline.json> -Metric MedianMs -Update` to record an intentional host baseline,
then omit `-Update` to compare an equivalent run. Use a separate baseline with
`-Metric AllocatedBytes` for managed allocations. Select tolerances from repeated
runs on the same idle host; do not infer portable ceilings from one measurement.
Keep these gates opt-in.

Synthetic ZIP inputs isolate scaling behavior; the pinned native inputs exercise
independent producers. The lane does not qualify Apple export equivalence, rendering,
native-memory retention, managed/resident peaks, trimming/AOT, sandbox/device
acceptance or portable resource ceilings. Those contracts
remain in [I5](../Docs/ROADMAP.md#i5-runtime-and-apple-host-acceptance).
