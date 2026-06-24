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

The comparison harness includes only libraries whose public APIs map to a scenario. It uses separate handling for legacy EPPlus so incompatible package generations do not run in one process.

### Competitor coverage

The comparison suite is intentionally scenario-shaped rather than forcing every library into every workflow:

- Workbook/package edit scenarios, such as `copy-worksheet-package`, compare `OfficeIMO.Excel`, `OfficeIMO.Excel Values`, `ClosedXML`, current `EPPlus`, and legacy `EPPlus 4.5.3.3`.
- Streaming/table export scenarios compare `OfficeIMO.Excel` with `ClosedXML`, current `EPPlus`, legacy `EPPlus 4.5.3.3`, `MiniExcel`, `LargeXlsx`, and `Sylvan.Data.Excel` where those libraries expose a matching write path.
- Read scenarios compare `OfficeIMO.Excel` with `ClosedXML`, current `EPPlus`, legacy `EPPlus 4.5.3.3`, `MiniExcel`, `ExcelDataReader`, and `Sylvan.Data.Excel` where the library exposes a matching read path.
- Libraries that do not expose a natural worksheet-copy or workbook-edit API are omitted from that specific scenario instead of being represented by an artificial row-by-row workaround.

NPOI is intentionally not a default comparison package yet. Current NPOI 2.8.x has a binary EULA / maintenance-fee requirement for revenue-generating users, while NPOI 2.7.6 is no longer latest. Add NPOI through an opt-in benchmark adapter or separate benchmark project once the version and license policy is chosen. Natural NPOI lanes are plain row/cell write and read, DataTable/DataSet-style import/export, formula text/cached value reads, and a separate `.xls` compatibility lane; do not force it into OfficeIMO-specific template, feature-preflight, PDF, package-copy, or fast-package scenarios.

For release-style evidence:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- comparison-suite --out-dir .\Docs\benchmarks\comparison-current --row-set 2500,25000 --warmup 1 --iterations 3
```

Focus the package-copy workbook merge scenario:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare --rows 25000 --scenario copy-worksheet-package --warmup 1 --iterations 3
```

Use `--skip-legacy-epplus` only when you want a faster local pass without the isolated EPPlus 4.x subprocess:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare --rows 25000 --scenario copy-worksheet-package --skip-legacy-epplus --warmup 1 --iterations 3
```

## Website data

After a comparison-suite run, refresh website/blog benchmark data with:

```powershell
.\Build\Generate-ExcelBenchmarkWebsiteData.ps1 -SummaryPath .\Docs\benchmarks\comparison-current\officeimo.excel.comparison-summary.json -ManifestPath .\Docs\benchmarks\comparison-current\officeimo.excel.comparison-suite-manifest.json -RunMode quick
```

## Boundaries

- Benchmark scenarios and opt-in comparison commands belong here.
- Runtime workbook behavior belongs in `OfficeIMO.Excel`.
- Committed benchmark artifact guidance belongs in [Docs/benchmarks](../Docs/benchmarks/README.md).
- Comparison outputs are local evidence and should not be treated as CI gates unless a workflow explicitly opts into them.
