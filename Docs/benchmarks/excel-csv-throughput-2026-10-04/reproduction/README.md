# Reproduce the Excel/CSV allocation comparison

Use .NET SDK 10.0.112 and the peer package pins recorded in `../provenance.json`.
Build baseline `dfb5753794bbbe6cbb1a71c8e5d6fe83f4ec4f07` and candidate
`fe9daeb03098217473946affc1b8109714edc49f` in separate clean checkouts.
The application benchmarks are existing opt-in projects; no runtime packages
are added to OfficeIMO to run them.

1. Build `OfficeIMO.CSV.Benchmarks` and `OfficeIMO.Excel.Benchmarks` in Release
   for `net10.0` in both checkouts. Keep outputs separate.
2. Make a task-owned scratch directory, for example
   `Ignore/Benchmarks/excel-csv-throughput` under the candidate checkout. Copy
   each benchmark output directory into `baseline-csv`, `candidate-csv`,
   `baseline-excel`, and `candidate-excel` beneath it.
3. The shared-string benchmark is new. Copy only
   `OfficeIMO.Excel.Benchmarks.*` from the candidate Excel snapshot over the
   same files in the baseline Excel snapshot. Keep the baseline
   `OfficeIMO.Excel.dll`. Both engines then use the identical fixture and
   observation code with their respective product assemblies.
4. Copy this folder's scripts into the scratch directory. Put `Program.cs`
   and `Compare.csproj.template` in a `compare` subdirectory, renaming the
   template to `Compare.csproj`. Build that project in Release.
5. Discover CPU/cache topology and adapt the Windows affinity values in
   `Program.cs` and `qualify.ps1`. The recorded masks are not portable defaults.
   Use a quiet workstation for timing qualification. Keep source, runtime,
   priority, power plan, fixtures, and package versions fixed during measurement.

From the candidate checkout in PowerShell, with `$taskRoot` set to the scratch
directory:

```powershell
$env:OFFICEIMO_PERFORMANCE_SNAPSHOTS = $taskRoot
$env:OFFICEIMO_BENCHMARK_DATA = Join-Path $taskRoot 'fixtures'
$env:OFFICEIMO_COMPARISON_SCENARIOS = 'CsvShortJsonAsNeeded,CsvShortJsonAlways,ExcelSstAscii,ExcelSstUnicodeTail,ExcelSstEntityTail'
dotnet "$taskRoot/compare/bin/Release/net10.0/Compare.dll" --filter '*' `
  --invocationCount 1024 --unrollFactor 1 --warmupCount 6 --iterationCount 12 `
  --launchCount 1 --outliers DontRemove --artifacts "$taskRoot/before-after"
pwsh -NoProfile -File "$taskRoot/qualify.ps1"
pwsh -NoProfile -File "$taskRoot/timing-control.ps1"
$env:OFFICEIMO_COMPARISON_SCENARIOS = 'Csv25K,CsvLongJsonAsNeeded'
dotnet "$taskRoot/compare/bin/Release/net10.0/Compare.dll" --filter '*' `
  --invocationCount 32 --unrollFactor 1 --warmupCount 8 --iterationCount 16 `
  --launchCount 1 --outliers DontRemove --artifacts "$taskRoot/before-after-large"
```

`qualify.ps1` invokes the existing PowerForge `Invoke-BenchmarkSuite` driver,
the repository's native BenchmarkDotNet entrypoints, and netstandard builds.
The measured PowerForge installation reports 3.0.154.2; that is provenance,
not a public dependency pin. `timing-control.ps1` runs identical-baseline A/A
controls followed by before/after A/B runs. Both scripts write only to the
selected scratch directory apart from ordinary project build output. Their
relative checkout discovery assumes the example scratch directory layout.

The packet also retains discarded intermediate stages `82df6ee4a` and
`28464d9e3`. To repeat those experiments, build their CSV snapshots separately
and select their directory prefix with `OFFICEIMO_PERFORMANCE_CSV_CANDIDATE`.
The `qualified-*` native runs measure the final source; the other run names
retain their recorded source stage. The final peer run uses 128 invocations
per iteration and selects short JSON/as-needed plus 25K mixed-value writers.

`Csv25K` is the 25,000-row **Quoted** workload. `CsvLongJsonAsNeeded` writes
1,000 rows with the 4,096 text-length parameter. The native mixed-value writer
comparison is a separate workload; its filter also includes the parallel
OfficeIMO method, which is retained in the evidence but excluded from the
sequential comparison table.

`--alloc ExcelRead 32 65535 baseline` and the corresponding `candidate` command
on `Compare.dll` are warmed allocation diagnostics, not BenchmarkDotNet timing
results. They warm up 32 complete scans and count current-thread allocated
bytes across a further 32. Use the native MemoryDiagnoser results for the
reported comparison and keep any disagreement visible.

Run the test projects using the filters and targets in `../validation.json`.
The existing benchmark setup checks fixture hashes, values, row counts, and
applicable byte equivalence; a zero process exit code alone is insufficient.
Inspect successful benchmark counts and every setup/validation result. Preserve
all samples, including unfavorable and identical-version control samples.
Remove disposable binary snapshots, downloaded fixtures, and build output after
retaining the compact evidence packet.
