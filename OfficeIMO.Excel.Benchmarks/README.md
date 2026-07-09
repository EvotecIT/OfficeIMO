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

NPOI is intentionally not a default comparison package. Benchmark-only local comparison is fine, but normal solution restore/build should not pull NPOI and OfficeIMO runtime code must not depend on it. Use the opt-in [OfficeIMO.Excel.Benchmarks.NPOI](../OfficeIMO.Excel.Benchmarks.NPOI/README.md) runner for NPOI evidence. Natural NPOI lanes are plain row/cell write and read, DataTable/DataSet-style import/export, formula text/cached value reads, conditional-formatting rule reads, and a separate `.xls` compatibility lane; do not force it into OfficeIMO-specific template, feature-preflight, PDF, package-copy, or fast-package scenarios.

For release-style evidence:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- comparison-suite --out-dir .\Docs\benchmarks\comparison-current --row-set 2500,25000 --warmup 1 --iterations 3
```

Focus the package-copy workbook merge scenario:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare --rows 25000 --scenario copy-worksheet-package --warmup 1 --iterations 3
```

Focus the `IDataReader.GetValues` read path used by bulk-copy style consumers:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare .\Ignore\Benchmarks\excel-getvalues-25000.json --rows 25000 --scenario read-datareader-getvalues --skip-legacy-epplus --warmup 2 --iterations 7
```

Compare row scanning, selective field access, full `GetValues`, and typed getters:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare .\Ignore\Benchmarks\excel-datareader-25000.json --rows 25000 --scenario read-datareader-readonly,read-datareader-first-column,read-datareader-getvalues,read-datareader-typed --skip-legacy-epplus --warmup 3 --iterations 15
```

Local 2026-07-09 results at 25,000 rows, with three warmups and 15 measured
iterations:

| Reader access | OfficeIMO.Excel | Sylvan.Data.Excel | ExcelDataReader | Result |
| --- | ---: | ---: | ---: | --- |
| Rows only | 40.36 ms / 1.0 MB | 44.61 ms / 3.5 MB | 110.41 ms / 42.6 MB | OfficeIMO fastest |
| First column | 42.23 ms / 2.5 MB | 47.37 ms / 3.5 MB | 109.30 ms / 42.6 MB | OfficeIMO fastest |
| `GetValues` | 55.93 ms / 19.7 MB | 67.08 ms / 7.4 MB | 122.95 ms / 42.6 MB | OfficeIMO fastest |
| Typed getters | 57.03 ms / 16.8 MB | 53.65 ms / 3.4 MB | 126.78 ms / 42.6 MB | Sylvan 1.06x faster |

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
