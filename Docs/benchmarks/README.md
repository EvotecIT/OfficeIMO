# Excel benchmark artifacts

This folder stores small, committed benchmark artifacts for `OfficeIMO.Excel`.

## Artifact types

- `officeimo.excel.snapshot-YYYY-MM-DD.json`: lightweight scenario snapshot for write, read, and round-trip flows.
- `officeimo.excel.write-profile-YYYY-MM-DD.json`: write-stage breakdown for optimization work.
- `officeimo.excel.read-profile-YYYY-MM-DD.json`: read-stage comparison for automatic, forced sequential, and forced parallel range conversion.
- `officeimo.excel.library-comparison.json`: local opt-in comparison across matching library surfaces.
- `officeimo.excel.npoi-comparison-current.json`: local opt-in NPOI verification for equivalent `.xlsx` row/cell and `.xls` read lanes, including scalar values, formulas, metadata, conditional formatting, AutoFilter range, style signals, and embedded pictures. NPOI stays outside normal solution restore/build.
- `officeimo.excel.npoi-verification-notes.md`: benchmark-only scope notes for the opt-in NPOI runner.
- `comparison-current\officeimo.excel.comparison-suite-manifest.json`: release-style suite manifest.
- `comparison-current\officeimo.excel.comparison-summary.md|csv|json`: one-table decision summary with speed, allocation, and package-size ratios.
- `officeimo.excel.comparison-report.md`: generated website/blog-oriented report distilled from comparison data.
- `Website\data\benchmarks-excel.json|benchmarks-excel-summary.json|benchmarks-excel-index.json`: generated website-facing benchmark data.

## Generate artifacts

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --snapshot .\Docs\benchmarks\officeimo.excel.snapshot-YYYY-MM-DD.json
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --profile-write .\Docs\benchmarks\officeimo.excel.write-profile-YYYY-MM-DD.json
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --profile-read .\Docs\benchmarks\officeimo.excel.read-profile-YYYY-MM-DD.json
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- comparison-suite --out-dir .\Docs\benchmarks\comparison-current --row-set 2500,25000 --warmup 1 --iterations 3
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks.NPOI\OfficeIMO.Excel.Benchmarks.NPOI.csproj -- --rows 2500 --warmup 1 --iterations 3 --out .\Docs\benchmarks\officeimo.excel.npoi-comparison-current.json
```

After a suite run, generate the website/blog data layer:

```powershell
.\Build\Generate-ExcelBenchmarkWebsiteData.ps1 -SummaryPath .\Docs\benchmarks\comparison-current\officeimo.excel.comparison-summary.json -ManifestPath .\Docs\benchmarks\comparison-current\officeimo.excel.comparison-suite-manifest.json -RunMode quick
```
