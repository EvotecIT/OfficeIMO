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
environments. Establish repeated Windows and non-Windows baselines before
turning these measurements into regression budgets.
