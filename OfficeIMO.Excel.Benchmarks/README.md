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

The suite keeps materially different contracts in separate lanes. Compact
writers omit explicit cell references for forward-only throughput, while the
normal OfficeIMO writer preserves the editable worksheet model. Shared-string
reads distinguish forward-only scans from rectangular materialization, and
DataTable reads distinguish automatic type inference from a caller-prepared
typed schema. These lanes should not be collapsed into one ranking.

One remaining gap is intentionally visible: in a stable 25,000-row Apple M4
forward shared-string scan with twenty warmups and nine measurements,
OfficeIMO measured 19.04 ms and Sylvan.Data.Excel measured 16.56 ms. The
OfficeIMO path retains general shared-string and worksheet compatibility; do
not present its much faster rectangular-materialization result as a win in this
forward-only contract.

For release-style evidence:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- comparison-suite --out-dir .\Docs\benchmarks\comparison-current --row-set 2500,25000 --skip-legacy-epplus --warmup 20 --iterations 9
```

The twenty warmups let tiered compilation settle before the custom rotated
runner starts measuring; fifteen warmups still produced visibly bimodal
isolated read results on ARM64.
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
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare .\Ignore\Benchmarks\excel-write-25000.json --rows 25000 --scenario write-datareader-compact-package,write-typed-rows-compact-package --skip-legacy-epplus --warmup 20 --iterations 31
```

## Current generated headline comparison

The package README uses this same PSPublishModule-managed snapshot. It combines
raw data paths with feature-bearing workbook work and only compares libraries
that expose a directly comparable public API. Lower is faster; results vary by
machine, runtime, package version, workload, warm-up, and options.
Treat differences below 5% as ties rather than ranking claims.

<!-- officeimo-excel-benchmark-table:start -->
| Scenario | Variables | Host | Operation | OfficeIMO.Excel | ClosedXML | EPPlus | LargeXlsx | SpreadCheetah | Sylvan.Data.Excel | Result |
| --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | --- |
| Compact DataReader to XLSX | Format=.xlsx, MeasuredIterations=9, Rows=25,000, Runner=rotated local, Snapshot=2026-07-14, Warmups=20 | .NET 8 | Write | 1.00x (23ms) | n/a | n/a | 1.11x (26ms) | 1.00x (23ms) | 1.11x (26ms) | OfficeIMO.Excel tied with SpreadCheetah |
| Feature-rich report to XLSX | Format=.xlsx, MeasuredIterations=9, Rows=25,000, Runner=rotated local, Snapshot=2026-07-14, Warmups=20 | .NET 8 | Create | 1.00x (37ms) | n/a | 11.12x (409ms) | n/a | n/a | n/a | OfficeIMO.Excel fastest |
| Styled DataReader table to XLSX | Format=.xlsx, MeasuredIterations=9, Rows=25,000, Runner=rotated local, Snapshot=2026-07-14, Warmups=20 | .NET 8 | Write | 1.00x (34ms) | 9.50x (320ms) | 9.76x (329ms) | n/a | n/a | n/a | OfficeIMO.Excel fastest |
| Typed objects streamed from XLSX | Format=.xlsx, MeasuredIterations=9, Rows=25,000, Runner=rotated local, Snapshot=2026-07-14, Warmups=20 | .NET 8 | Read | 1.00x (25ms) | 11.13x (278ms) | 10.08x (252ms) | n/a | n/a | 1.56x (39ms) | OfficeIMO.Excel fastest |
<!-- officeimo-excel-benchmark-table:end -->

`--skip-legacy-epplus` omits only the isolated EPPlus 4.x subprocess; current
EPPlus remains in the comparison. Keep this flag on modern macOS unless
`libgdiplus` is installed because the legacy AutoFit path depends on it:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare --rows 25000 --scenario copy-worksheet-package --skip-legacy-epplus --warmup 1 --iterations 3
```

## Website data

After a comparison-suite run, refresh website/blog benchmark data with:

```powershell
.\Build\Generate-ExcelBenchmarkWebsiteData.ps1 -SummaryPath .\Docs\benchmarks\comparison-current\officeimo.excel.comparison-summary.json -ManifestPath .\Docs\benchmarks\comparison-current\officeimo.excel.comparison-suite-manifest.json -RunMode quick
```
