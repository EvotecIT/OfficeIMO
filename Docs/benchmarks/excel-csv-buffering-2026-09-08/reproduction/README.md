# Reproducing the buffering measurements

The permanent benchmark projects own the workloads and complete-output validation. This packet adds an isolated before/after assembly loader and the PowerForge measurement driver. It targets the Windows machine used by the report. Keep benchmark artifacts in a dedicated output directory.

1. Create an output directory with a `compare` subdirectory. Copy `Program.cs` there and rename the project template to `Compare.csproj`. Copy both PowerShell scripts and `summarize-writers.py` into the output directory. Keep only one project of that name under the benchmark runner's search root; BenchmarkDotNet rejects ambiguous project names.
2. Build the comparison project in Release for .NET 10. Build the two product libraries from baseline commit `3f2b32a60114514e92af24ba682aae2d4ca3e097` in a separate checkout. Build the current CSV and Excel benchmark projects in Release for .NET 10.
3. Copy each complete benchmark output twice into `CSV-baseline`, `CSV-candidate`, `Excel-baseline`, and `Excel-candidate`, beside the scripts. Replace only the baseline product DLL with the corresponding baseline library. Both sides must use identical benchmark and dependency assemblies. The report's provenance and binary hashes identify the measured candidate.
4. Place the independently validated `65K_Records_Data.csv` and `65K_Records_Data.xlsx` fixtures in `fixtures` for reader measurements. Their sources and checks are defined by `MarkPflug65KFixture` in the benchmark projects. Verify the fixture hashes in the evidence packet.
5. Run the PowerForge suite through an installed PSPublishModule version supporting `Invoke-BenchmarkSuite`, rotated order, case sources, and validation. The measured local engine was `3.0.135.1`; it is recorded as provenance, not a public dependency pin. Freeze the entire checkout during each run.

The original machine uses two cache domains with masks `65535` and `4294901760`. Discover the current topology before reusing these values. The driver applies Normal priority and retains all samples. Update the two masks in the orchestration script for another topology.

```powershell
$writers = 'ExcelLongPlain,ExcelLongEscaped,ExcelLongMarkup,ExcelShortPlain,ExcelDefaultLongPlain,ExcelDefaultLongEscaped,Excel25K,Csv25K,CsvLongAsNeeded,CsvLongJsonAsNeeded,CsvLongQuotesAsNeeded,CsvShortJsonAsNeeded,CsvShortJsonAlways,CsvShortQuotesAsNeeded,CsvShortQuotesAlways,CsvFileShortAsciiAsNeeded,CsvFileShortUnicodeAsNeeded,CsvFileDenseJsonAsNeeded,CsvFileQuoteRunsAsNeeded,CsvFileLongNotesAsNeeded,CsvFileTypedValuesAsNeeded,CsvFileTypedValuesAlways'
./run-comparison.ps1 -Label final-writers -ScenarioList $writers -Iterations 30 -Exports 8 -Repeats 3
python ./summarize-writers.py final-writers

$env:OFFICEIMO_PERFORMANCE_AA = '1'
$calibration = 'Csv25K,CsvShortJsonAsNeeded,CsvShortJsonAlways,CsvLongJsonAsNeeded,CsvFileShortAsciiAsNeeded,CsvFileLongNotesAsNeeded,CsvFileTypedValuesAsNeeded'
./run-comparison.ps1 -Label aa-calibration -ScenarioList $calibration -Iterations 30 -Exports 16 -Repeats 2
python ./summarize-writers.py aa-calibration
$env:OFFICEIMO_PERFORMANCE_AA = $null
```

The source loader changes both sides to the baseline when `OFFICEIMO_PERFORMANCE_AA=1`; the measurement metadata records this mode. Calibration results are never combined with candidate results.

For native reader allocation and short-JSON measurements, set the environment paths and run the comparison project from its directory:

```powershell
$env:OFFICEIMO_PERFORMANCE_SNAPSHOTS = (Resolve-Path .).Path
$env:OFFICEIMO_BENCHMARK_DATA = Join-Path $env:OFFICEIMO_PERFORMANCE_SNAPSHOTS 'fixtures'
$env:OFFICEIMO_BENCHMARK_OUTPUT = Join-Path $env:OFFICEIMO_PERFORMANCE_SNAPSHOTS 'file-output'
$env:OFFICEIMO_COMPARISON_SCENARIOS = 'ExcelRead,CsvRead,CsvShortJsonAsNeeded'
$env:UseAppHost = 'false'
dotnet run --project ./compare/Compare.csproj -c Release -- --filter '*' --invocationCount 4 --unrollFactor 1 --warmupCount 8 --iterationCount 16 --launchCount 1 --outliers DontRemove --artifacts ./native-readers
```

The loader also supports diagnostic modes `--xml-compare` and `--alloc`. The latter uses an untimed allocation counter for profiling; native BenchmarkDotNet `MemoryDiagnoser` owns the reported allocation comparison. Run `csv-string-allocation.py` with the CSV fixture path to reproduce the returned-string allocation estimate.

Keep the native comparison project path and name short. The measured native runner used `rc/RcBench.csproj` with the same `Program.cs`, after a long generated Windows path prevented compilation. The PowerForge loader retained the `compare/Compare.csproj` layout.

Check successful case counts and non-null statistics in native JSON exports. A BenchmarkDotNet process can return exit code zero even when project discovery or compilation prevents any measurements.
