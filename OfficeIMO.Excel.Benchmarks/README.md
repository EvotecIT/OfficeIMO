# OfficeIMO.Excel.Benchmarks

`OfficeIMO.Excel.Benchmarks` is the benchmark harness for `OfficeIMO.Excel`. It measures representative workbook read, write, edit, package-size, and real-world feature workloads. It is not a runtime package.

## Run benchmarks

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj
```

Filter a class:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --filter *ExcelWriteBenchmarks*
```

Measure worksheet copy fast paths:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --filter *ExcelWorksheetCopyBenchmarks*
```

## Snapshot and profile artifacts

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --snapshot .\Docs\benchmarks\officeimo.excel.snapshot.json
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- write-profile --rows 25000
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- read-profile --rows 2500
```

## Library comparison

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare --rows 2500
```

Each scenario runs only libraries with a directly comparable public API. Legacy
EPPlus runs in a separate process. NPOI comparisons are available through the
opt-in [NPOI runner](../OfficeIMO.Excel.Benchmarks.NPOI/README.md).

For release-style evidence:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- comparison-suite --out-dir .\Docs\benchmarks\comparison-current --row-set 2500,25000 --warmup 1 --iterations 3
```

The suite writes JSON, CSV, Markdown, and a manifest. Run the focused README
comparison and refresh its generated table locally with:

```powershell
.\Build\Benchmarks\Update-BenchmarkReadmes.ps1 -Run Excel
```

The script selects documented equivalent workloads, emits PSPublishModule's
comparison schema, and calls `Update-BenchmarkDocument` for the
marker-delimited blocks. It runs only when a maintainer invokes it locally;
benchmark execution is not scheduled in CI.

Focus the package-copy workbook merge scenario:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare --rows 25000 --scenario copy-worksheet-package --warmup 1 --iterations 3
```

Compare row scanning, selective field access, full `GetValues`, and typed getters:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare .\Ignore\Benchmarks\excel-datareader-25000.json --rows 25000 --scenario read-datareader-readonly,read-datareader-first-column,read-datareader-getvalues,read-datareader-typed --skip-legacy-epplus --warmup 3 --iterations 15
```

Compare the fastest package-native write paths:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare .\Ignore\Benchmarks\excel-write-25000.json --rows 25000 --scenario write-datareader-compact-package,write-typed-rows-compact-package --skip-legacy-epplus --warmup 7 --iterations 31
```

## Current generated headline comparison

The package README uses this same PSPublishModule-managed snapshot. It combines
raw data paths with feature-bearing workbook work and only compares libraries
that expose a directly comparable public API. Lower is faster; results vary by
machine, runtime, package version, workload, warm-up, and options.

<!-- officeimo-excel-benchmark-table:start -->
| Scenario | Variables | Host | Operation | Metric | OfficeIMO.Excel | ClosedXML | EPPlus | LargeXlsx | Sylvan.Data.Excel | Result |
| --- | --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: | --- |
| Feature-rich report to XLSX | Format=.xlsx, Rows=25,000, Runner=rotated local, Snapshot=2026-07-13 | .NET 8 | Create | MeanMs | 1.00x (549ms) | n/a | 0.67x (368ms) | n/a | n/a | OfficeIMO.Excel slower than EPPlus |
| Plain DataReader to XLSX | Format=.xlsx, Rows=25,000, Runner=rotated local, Snapshot=2026-07-13 | .NET 8 | Write | MeanMs | 1.00x (35ms) | n/a | n/a | 1.06x (37ms) | 0.79x (28ms) | OfficeIMO.Excel slower than Sylvan.Data.Excel |
| Styled DataReader table to XLSX | Format=.xlsx, Rows=25,000, Runner=rotated local, Snapshot=2026-07-13 | .NET 8 | Write | MeanMs | 1.00x (39ms) | 13.45x (527ms) | 9.60x (376ms) | n/a | n/a | OfficeIMO.Excel fastest |
| Typed objects streamed from XLSX | Format=.xlsx, Rows=25,000, Runner=rotated local, Snapshot=2026-07-13 | .NET 8 | Read | MeanMs | 1.00x (60ms) | 4.55x (272ms) | 3.44x (205ms) | n/a | 1.17x (70ms) | OfficeIMO.Excel fastest |
<!-- officeimo-excel-benchmark-table:end -->

Use `--skip-legacy-epplus` only when you want a faster local pass without the isolated EPPlus 4.x subprocess:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare --rows 25000 --scenario copy-worksheet-package --skip-legacy-epplus --warmup 1 --iterations 3
```

## Website data

After a comparison-suite run, refresh website/blog benchmark data with:

```powershell
.\Build\Generate-ExcelBenchmarkWebsiteData.ps1 -SummaryPath .\Docs\benchmarks\comparison-current\officeimo.excel.comparison-summary.json -ManifestPath .\Docs\benchmarks\comparison-current\officeimo.excel.comparison-suite-manifest.json -RunMode quick
```
