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

Compare setup, row scanning, selective field access, and the full `GetValues` scan:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare .\Ignore\Benchmarks\excel-datareader-diagnostics-25000.json --rows 25000 --scenario read-datareader-open --scenario read-datareader-readonly --scenario read-datareader-first-column --scenario read-datareader-getvalues --skip-legacy-epplus --warmup 2 --iterations 7
```

Check typed getter access separately:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare .\Ignore\Benchmarks\excel-typed-reader-25000.json --rows 25000 --scenario read-datareader-typed --skip-legacy-epplus --warmup 3 --iterations 10
```

Local 2026-07-09 short-run signal for `read-datareader-getvalues` at 25,000
rows after the shared-string index-buffering pass: `OfficeIMO.Excel` averaged
`64.35 ms`, median `66.17 ms`, and `19.7 MB` allocated; `ExcelDataReader`
averaged `121.14 ms`, median `112.67 ms`, and `42.6 MB` allocated;
`Sylvan.Data.Excel` averaged `69.02 ms`, median `70.15 ms`, and `7.4 MB`
allocated. Treat this as a focused signal: OfficeIMO is the fastest full-row
`GetValues` reader in this run and uses less than half of ExcelDataReader's
allocation, while Sylvan still has the lower allocation profile.

The companion diagnostic lanes show where the allocation work belongs. On the
same local run, `read-datareader-open` averaged `0.98 ms` and `221.5 KB` for
OfficeIMO, while `read-datareader-readonly` averaged `53.46 ms` and `1.0 MB`
and `read-datareader-first-column` averaged `40.15 ms` and `3.0 MB`. That means
OfficeIMO now avoids decoding unrequested cells for row-only and selective
field-access consumers; the remaining allocation in the full `GetValues` lane is
the cost of materializing every requested cell value. After the primitive typed
cache and boolean-buffer pass, repeated local 25,000-row `read-datareader-typed`
runs put OfficeIMO around `65-68 ms` average, `58-60 ms` median, and `16.8 MB`
allocated. `ExcelDataReader` was around `129-130 ms` and `42.6 MB` allocated.
`Sylvan.Data.Excel` stayed much leaner at `3.4 MB` allocated and can still win
elapsed time on repeat runs. Treat this as a solid improvement over
ExcelDataReader and over the earlier OfficeIMO allocation profile, not as a final
typed-reader win over Sylvan.

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
