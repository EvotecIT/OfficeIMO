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

The snapshot JSON records averages, medians, and raw samples for each scenario.

Generate a write-stage profile to identify where report-export time is spent:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --profile-write .\Docs\benchmarks\officeimo.excel.write-profile.json
```

The write profile JSON records averages, medians, and raw samples for each stage so outlier runs are visible.
