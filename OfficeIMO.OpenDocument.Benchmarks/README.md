# OfficeIMO.OpenDocument benchmarks

This project tracks the performance and allocations of three contracts that are easy to regress: opening and enumerating a 2,000-paragraph ODT, writing an ODS cell at an extreme sparse coordinate, and evaluating a 1,000-cell OpenFormula range.

Run the full benchmark set on the target framework you want to measure:

```powershell
dotnet run --project OfficeIMO.OpenDocument.Benchmarks/OfficeIMO.OpenDocument.Benchmarks.csproj -c Release -f net8.0
```

Use `-f net10.0` for a .NET 10 run. The benchmark classes do not pin a runtime;
the selected target framework is the measured runtime.

BenchmarkDotNet results are machine-specific engineering evidence, not universal throughput guarantees. Keep generated `BenchmarkDotNet.Artifacts` out of source control.

## Workflow evidence

The opt-in evidence runner measures complete ODT, ODS, and ODP create/save,
open/read, and open/edit/save workflows at small, normal, and large scales. Each
measurement runs in a separate child process and records elapsed time, managed
allocations, peak working set, input bytes, output bytes, record counts, and a
content checksum. File-producing lanes reopen and structurally validate the
package after measurement; open/read lanes traverse the complete representative
content and validate the source package after measurement.

Run the normal-scale matrix and keep the JSON under the ignored artifact root:

```powershell
dotnet run --project OfficeIMO.OpenDocument.Benchmarks/OfficeIMO.OpenDocument.Benchmarks.csproj -c Release -f net8.0 -- --evidence --scale Normal --repeat 3 --json .benchmark-artifacts/opendocument/normal.json
```

Use `--format`, `--operation`, or `--scale` to narrow a diagnostic run. The
accepted values are `ODT|ODS|ODP`, `CreateSave|OpenRead|OpenEditSave`, and
`Small|Normal|Large`, respectively. Use `--repeat` to collect multiple cold,
isolated samples per case. The report includes the source commit, dirty-tree
state, runtime, operating system, architecture, and logical processor count so
results are not detached from their environment.

Treat cold elapsed time as directional on a busy workstation. Managed
allocations and output size are normally more stable. Peak working set includes
runtime startup and should only be compared between isolated probes on similar
environments. A `null` peak-working-set value means the runtime did not expose that measurement; it is not a zero-memory result. Establish repeated Windows and non-Windows baselines before
turning these measurements into regression budgets.

## Draw page projection

The opt-in Draw suite projects all pages through `OdgDocument.ToDrawings` at
100, 500, and 1,000 pages. Each document shares one master and page layout. The
two cases use empty pages or a master header with dynamic page-number and
page-count fields. Validation checks every page's dimensions and visible text,
a content checksum, and unchanged content, styles, and metadata XML.

Build the workload, inspect the matrix, then run it in a fresh PowerShell process:

```powershell
dotnet build OfficeIMO.OpenDocument.Benchmarks/OfficeIMO.OpenDocument.Benchmarks.csproj -c Release -f net8.0
pwsh -NoProfile -File Build/Benchmarks/Run-DrawPageProjectionBenchmarks.ps1 -Plan
pwsh -NoProfile -File Build/Benchmarks/Run-DrawPageProjectionBenchmarks.ps1 -OutputRoot Ignore/Benchmarks/DrawPageProjection
```

`PSPublishModule` owns timing, warmups, rotated case order, allocation sampling,
and JSON, CSV, and Markdown artifacts. Use a runner that reports managed
allocation; the wrapper rejects missing allocation or failed validation. Pass
`-ModulePath` to select a source-built `PSPublishModule.dll` or module manifest,
and `-BinaryRoot` to select the workload's build directory. The actual runtime
is the PowerShell host's runtime, which the report records. Selecting a workload
target framework does not select the host runtime.

Use `-PageCount` and `-Content Empty|Fields` to narrow the matrix. Preparation,
source readback, and result validation run outside measurement; operation
samples include PowerShell invocation overhead. Working-set delta is neither
peak nor retained memory. Compare equivalent inputs, runner versions, hosts,
and policies. Processor placement and busy-host timing require separate
qualification before drawing throughput conclusions. These measurements are
outside ordinary correctness CI.
