# iWork runtime evidence

This opt-in workload library measures the shared iWork reader and the Word, Excel
and PowerPoint destination adapters. It stays outside the normal solution and
correctness CI. [PowerForge](https://github.com/EvotecIT/PSPublishModule) owns
warmups, rotated iteration order, elapsed and memory counters, artifacts and gates.
The library owns deterministic inputs and semantic validation.

Build from the repository root using the pinned SDK:

```powershell
dotnet build OfficeIMO.IWork.Benchmarks/OfficeIMO.IWork.Benchmarks.csproj -c Release
./Build/Benchmarks/Run-IWorkRuntimeBenchmarks.ps1 -Plan
./Build/Benchmarks/Run-IWorkRuntimeBenchmarks.ps1 -OutputRoot $env:EVOTEC_DEV_TEMP/iwork-runtime
```

Use a fresh PowerShell 7 process. The runner requires PowerForge operation allocation
measurement and fails when that counter is missing. To validate against shared
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
and full semantic readback are outside measurement. Conversion requires a complete
editable reconstruction report; saved DOCX, XLSX and PPTX packages are reopened and
checked before a sample succeeds. `-Scale`, `-Kind` and `-Operation` select cases;
`-WarmupCount` and `-IterationCount` default to two and five. A quick smoke run uses
`-Scale Small -WarmupCount 0 -IterationCount 1`.

Artifacts include raw samples, summaries, CSV tables and environment metadata.
Case variables retain deterministic input SHA-256 hashes; metadata retains the
spec, workload and candidate owner assembly hashes. Inspect failures before using
elapsed or allocation results. `AllocatedBytes` includes managed host invocation
and concurrent threads in the same process. `WorkingSetDeltaBytes` is a signed
resident-page difference, not peak or retained memory.

Use PowerForge `Test-BenchmarkGate -SummaryPath <summary.json> -BaselinePath
<baseline.json> -Metric MedianMs -Update` to record an intentional host baseline,
then omit `-Update` to compare an equivalent run. Use a separate baseline with
`-Metric AllocatedBytes` for managed allocations. Select tolerances from repeated
runs on the same idle host; do not infer portable ceilings from one measurement.
Keep these gates opt-in.

These supported-format synthetic ZIP inputs isolate scaling behavior. They do not
qualify native Apple exports, rendering, repeated-open retention, cancellation
latency, managed peaks, trimming/AOT or sandbox/device acceptance. Those contracts
remain in [I5](../Docs/ROADMAP.md#i5-runtime-and-apple-host-acceptance).
