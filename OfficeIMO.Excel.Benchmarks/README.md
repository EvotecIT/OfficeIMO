# OfficeIMO.Excel.Benchmarks

Internal benchmark harness for `OfficeIMO.Excel`.

It measures representative Excel workloads rather than synthetic single-cell operations:

- bulk workbook export
- workbook read/materialization
- load/edit/save round-trips

The built-in comparison baselines are `ClosedXML`, current `EPPlus`, and legacy `EPPlus 4.5.3.3`. The current EPPlus path is an explicit local benchmark command and configures EPPlus for non-commercial local benchmark use; the legacy EPPlus path runs in a separate helper project so the two EPPlus package generations do not share one process. These comparisons are intentionally not wired into CI.

Run all benchmark classes with:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj
```

Filter a specific class with:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --filter *ExcelWriteBenchmarks*
```

Generate a lightweight JSON baseline snapshot with:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --snapshot .\Docs\benchmarks\officeimo.excel.snapshot.json
```

The default snapshot uses 2,500 rows. Use `--rows` to generate larger comparison tiers:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --snapshot .\Docs\benchmarks\officeimo.excel.snapshot-25000.json --rows 25000
```

The snapshot JSON records averages, medians, and raw samples for each scenario.

To refresh the public website benchmark table from the same snapshot run, add `--website-data`:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --snapshot .\Docs\benchmarks\officeimo.excel.snapshot.json --website-data .\Website\data\benchmarks.json
```

Generate a write-stage profile to identify where report-export time is spent:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --profile-write .\Docs\benchmarks\officeimo.excel.write-profile.json
```

Generate a read-stage profile to compare automatic, forced sequential, forced parallel range conversion, and sparse row/column reads:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --profile-read .\Docs\benchmarks\officeimo.excel.read-profile.json
```

The profile and snapshot commands also accept short aliases with a default output path under `Docs\benchmarks`:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- write-profile --rows 25000
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- read-profile --rows 2500
```

The read profile also accepts `--warmup` and `--iterations` when a focused local run needs steadier samples:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- read-profile --rows 2500 --warmup 3 --iterations 9
```

The read profile measures automatic, forced sequential, and forced parallel variants in rotated groups for each read API. This keeps mode comparisons from depending on fixed scenario order and makes first-sample outliers visible in the raw sample list.

Generate a local library comparison against ClosedXML and EPPlus:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare --rows 2500
```

To refresh the committed comparison artifact explicitly, pass the output path immediately after `compare` or with `--out`:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare .\Docs\benchmarks\officeimo.excel.library-comparison.json --rows 2500
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare --out .\Docs\benchmarks\officeimo.excel.library-comparison.json --rows 2500
```

By default this also launches the isolated legacy EPPlus helper. For a faster current-library-only pass, add `--skip-legacy-epplus`. Use `--scenario` to run one or more targeted scenarios during tuning:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare --rows 2500 --scenario read-range --scenario read-objects
```

The comparison command defaults to one warmup and three measured samples so quick checks stay quick. For less noisy local tuning, increase the sample count; the same settings are passed through to the isolated legacy EPPlus helper:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare --rows 2500 --scenario read-range --warmup 2 --iterations 7
```

Current-library comparison scenarios measure OfficeIMO, ClosedXML, and current EPPlus in rotated groups for each scenario so fixed library order does not decide the numbers. Legacy EPPlus still runs in a separate process because it uses a different package generation.

The comparison command covers bulk report writes, append-style writes, dense range reads, bounded top-of-sheet reads, DataTable materialization, streaming range reads, bounded streaming reads, large sparse reads, typed object materialization, AutoFit on an existing workbook, large shared-string payloads, formula text reads, and shared-string reads. Read scenarios record deterministic value checksums as `OutputMetric`, so local comparisons can confirm that each library read equivalent content instead of only touching the same number of rows. The command fails if a read checksum differs across libraries, including legacy EPPlus. Write and AutoFit scenarios keep package-size metrics because each library serializes workbook parts differently. The comparison and read-profile JSON include the benchmark build configuration, and performance artifacts should be produced with `-c Release`. The command writes JSON under `Docs\benchmarks` by default and does not participate in CI.

The write profile also accepts `--rows`, which is the preferred way to investigate 25,000+ row report-export costs without running the full BenchmarkDotNet suite:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --profile-write .\Docs\benchmarks\officeimo.excel.write-profile-25000.json --rows 25000
```

The write profile JSON records averages, medians, and raw samples for each stage so outlier runs are visible. OfficeIMO timing hooks also add AutoFit sub-stages (`BuildPlan`, `CalculateWidths`, and `ApplyWidths`) to make large-sheet tuning more targeted.

The harness configures OfficeIMO with `Execution.SaveWorksheetAfterAutoFit = false`, matching the normal report-export pattern where all worksheet mutations are committed once when the document is saved or disposed.
