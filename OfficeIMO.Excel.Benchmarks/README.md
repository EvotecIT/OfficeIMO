# OfficeIMO.Excel.Benchmarks

Internal benchmark harness for `OfficeIMO.Excel`.

It measures representative Excel workloads rather than synthetic single-cell operations:

- bulk workbook export
- workbook read/materialization
- load/edit/save round-trips

The built-in comparison baseline is `ClosedXML`, because it is easy to restore and run in public repo workflows. The scenario layout is meant to support future EPPlus runs too, but EPPlus-specific comparison is intentionally left as an opt-in local step because license setup is environment-specific.

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

Generate a read-stage profile to compare automatic, forced sequential, and forced parallel range conversion:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --profile-read .\Docs\benchmarks\officeimo.excel.read-profile.json
```

The profile and snapshot commands also accept short aliases with a default output path under `Docs\benchmarks`:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- write-profile --rows 25000
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- read-profile --rows 2500
```

The write profile also accepts `--rows`, which is the preferred way to investigate 25,000+ row report-export costs without running the full BenchmarkDotNet suite:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --profile-write .\Docs\benchmarks\officeimo.excel.write-profile-25000.json --rows 25000
```

The write profile JSON records averages, medians, and raw samples for each stage so outlier runs are visible. OfficeIMO timing hooks also add AutoFit sub-stages (`BuildPlan`, `CalculateWidths`, and `ApplyWidths`) to make large-sheet tuning more targeted.

The harness configures OfficeIMO with `Execution.SaveWorksheetAfterAutoFit = false`, matching the normal report-export pattern where all worksheet mutations are committed once when the document is saved or disposed.
