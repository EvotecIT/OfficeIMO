# OfficeIMO.Excel.Benchmarks

Internal benchmark harness for `OfficeIMO.Excel`.

It measures representative Excel workloads rather than synthetic single-cell operations:

- bulk workbook export
- workbook read/materialization
- load/edit/save round-trips

The built-in comparison baselines are `ClosedXML`, current `EPPlus`, `MiniExcel`, `LargeXlsx`, read-side `ExcelDataReader`, `Sylvan.Data.Excel`, and legacy `EPPlus 4.5.3.3`. The current EPPlus path is an explicit local benchmark command and configures EPPlus for non-commercial local benchmark use; each library is included only where its public surface maps to the scenario being measured; the legacy EPPlus path runs in a separate helper project so the two EPPlus package generations do not share one process. These comparisons are intentionally not wired into CI. `LargeXlsx` participates in plain/streaming writer scenarios where the public API can write equivalent typed worksheet rows without table metadata. `Sylvan.Data.Excel` participates in read scenarios and in the plain `DbDataReader` export scenario; it is intentionally excluded from styled-table, AutoFit, package-edit, and rich-report lanes because its writer is a flat rectangular data writer.

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

Generate a read-stage profile to compare automatic, forced sequential, forced parallel range conversion, streaming typed object reads, and sparse row/column reads:

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

Generate a local library comparison where each library has a comparable public surface:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare --rows 2500
```

To refresh the committed comparison artifact explicitly, pass the output path immediately after `compare` or with `--out`:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare .\Docs\benchmarks\officeimo.excel.library-comparison.json --rows 2500
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare --out .\Docs\benchmarks\officeimo.excel.library-comparison.json --rows 2500
```

By default this also launches the isolated legacy EPPlus helper. The helper accepts the same positional output path and `--out` aliases when run directly. For a faster current-library-only pass, add `--skip-legacy-epplus`. Use `--scenario` to run one or more targeted scenarios during tuning:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare --rows 2500 --scenario read-range --scenario read-objects --scenario read-objects-stream
```

Package size can be profiled separately from the speed comparison. This keeps timed samples clean and then generates one extra workbook per library to break the `.xlsx` ZIP into worksheet, shared-string, style, table, relationship, document-property, and other parts:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- package-profile --out .\Docs\benchmarks\officeimo.excel.package-profile.json --rows 25000 --scenario write-datatable-table-direct --scenario write-datatable-direct --scenario write-cellvalue-strings --scenario large-shared-strings
```

The package profile is intended for opt-in size investigation, especially when speed is already ahead and a smaller package might matter for email attachments, sync, storage, or slow network transfer. Treat it as diagnostic evidence before changing defaults.

For release-style evidence, use the comparison suite command. It runs the normal speed comparison, the package profile, and the dense `HelloWorld` read shape for the same row counts, then writes a manifest that records the artifact paths and benchmark settings:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- comparison-suite --out-dir .\Docs\benchmarks\comparison-current --row-set 2500,25000 --warmup 1 --iterations 3
```

The suite also writes `officeimo.excel.comparison-summary.md`, `.csv`, and `.json`. Those summary files are the decision layer: one table with row count, artifact kind, workload, category, scenario, library, mean, standard deviation, standard error, ratio to OfficeIMO, ratio to best, allocation, allocation ratio, package size, package-size ratio, winner/loss status, and package-part metrics where available. The workload/category columns keep plain streaming export, table export, AutoFit, reads, package size, object projection, and cell-writer lanes from being blended into one score. The standard-deviation and standard-error columns come from the lightweight rotated runner, while allocations use `GC.GetAllocatedBytesForCurrentThread`; use the BenchmarkDotNet benchmark classes when a publication-grade `Error` column is required.

The suite deliberately writes the dense `HelloWorld` benchmark as a separate artifact because it uses a different generated fixture from the normal report/sales scenarios. Pass `--skip-dense-helloworld` only when you want a faster local tuning run.

To refresh website/blog-oriented data after a suite run, generate the Excel benchmark data files:

```powershell
.\Build\Generate-ExcelBenchmarkWebsiteData.ps1 -SummaryPath .\Docs\benchmarks\comparison-current\officeimo.excel.comparison-summary.json -ManifestPath .\Docs\benchmarks\comparison-current\officeimo.excel.comparison-suite-manifest.json -RunMode quick
```

During tuning, use a smaller scenario set before launching the full matrix:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- comparison-suite --out-dir $env:TEMP\officeimo-excel-suite-smoke --row-set 1000 --warmup 1 --iterations 1 --skip-legacy-epplus --scenario write-datatable-direct --scenario read-range --scenario large-shared-strings
```

Focused write-path tuning can target the automatic direct package writer scenarios without changing user-facing API usage:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare --rows 2500 --scenario write-datatable-direct --scenario write-datareader-table --scenario write-datareader-table-autofit --scenario write-datareader-plain --scenario write-cellvalues-rectangle-direct --scenario write-cellvalues-headerless-rectangle-direct --scenario write-insertobjects-autofitcolumnsfor-direct --scenario write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct --scenario write-powershell-mixed-objects-direct --scenario write-cellvalue-strings --scenario write-cellvalue-numbers --scenario write-cellvalue-scalars --scenario write-cellvalue-temporal --scenario write-cellvalue-object-mixed --scenario write-cellformula
```

To compare OfficeIMO's normal direct writer paths against streaming-export libraries such as `LargeXlsx`, use the plain/direct writer subset:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare --rows 25000 --skip-legacy-epplus --scenario write-dataset-sparse-direct-export --scenario write-datatable-direct --scenario write-datareader-plain --scenario write-cellvalues-rectangle-direct --scenario write-insertobjects-direct --scenario write-powershell-mixed-objects-direct --scenario write-fluent-rowsfrom-direct --scenario append-plain-rows
```

The 2023 LargeXlsx/MiniExcel/ClosedXML blog workload is available as a normalized 20-column string DTO export. The original source used string properties `C1` through `C20`; this scenario gives every library the same header row and data rows so package metrics stay comparable:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- package-profile --out .\Docs\benchmarks\officeimo.excel.package-profile-blog-2023-20-string-columns.json --rows 10000 --scenario write-blog-2023-20-string-columns --warmup 1 --iterations 3
```

The comparison harness has opt-in dense and streaming read scenarios for a simple `A1:J(row count)` workbook where every cell contains `HelloWorld`. It generates the workbook shape locally and compares matching read APIs, so it can be run at full scale without also creating the normal sales/report fixtures:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare --out .\Docs\benchmarks\officeimo.excel.dense-helloworld.json --rows 1000000 --warmup 1 --iterations 3 --skip-legacy-epplus --scenario dense-helloworld-read-range --scenario dense-helloworld-read-stream
```

For a quick smoke check before the full proof run, lower `--rows`:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare --out $env:TEMP\officeimo.excel.helloworld-smoke.json --rows 1000 --warmup 1 --iterations 1 --skip-legacy-epplus --scenario dense-helloworld-read-range --scenario dense-helloworld-read-stream
```

The comparison command defaults to one warmup and three measured samples so quick checks stay quick. For less noisy local tuning, increase the sample count; the same settings are passed through to the isolated legacy EPPlus helper:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare --rows 2500 --scenario read-range --warmup 2 --iterations 7
```

Current-library comparison scenarios measure OfficeIMO and the included baselines in rotated groups for each scenario so fixed library order does not decide the numbers. Comparable read scenarios use one canonical generated workbook payload where the library APIs can all read the same file shape. Read-only libraries participate only in read scenarios, and legacy EPPlus still runs in a separate process because it uses a different package generation.

The comparison command covers bulk report writes, automatic direct package writer paths (`InsertDataTable`, `InsertDataTableAsTable`, complete-rectangle `CellValues`, typed `InsertObjects`, PowerShell-like mixed dictionary `InsertObjects`, and fluent `RowsFrom`), styled, AutoFit, and plain streaming `InsertDataReader`, append-style writes, cell-by-cell text assignment, dense range reads, first-column reads from wider sheets, bounded top-of-sheet and bottom-of-sheet reads, DataTable materialization, streaming range reads, bounded streaming reads with default and small chunk sizes, large sparse reads, eager and streaming typed object materialization, AutoFit on an existing workbook, large shared-string payloads, formula text reads, shared-string reads, and the opt-in dense `HelloWorld` grid read. Read scenarios record deterministic value checksums as `OutputMetric`, so local comparisons can confirm that each library read equivalent content instead of only touching the same number of rows. The command fails if a read checksum differs across libraries, including legacy EPPlus. Libraries are skipped for scenarios where their public APIs do not expose an equivalent operation. Write and AutoFit scenarios keep package-size metrics because each library serializes workbook parts differently. The comparison and package-profile JSON include mean, median, standard deviation, standard error, raw timing samples, mean allocation, median allocation, and raw allocation samples. The comparison and read-profile JSON include the benchmark build configuration, and performance artifacts should be produced with `-c Release`. The command writes JSON under `Docs\benchmarks` by default and does not participate in CI.

The write profile also accepts `--rows`, which is the preferred way to investigate 25,000+ row report-export costs without running the full BenchmarkDotNet suite:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --profile-write .\Docs\benchmarks\officeimo.excel.write-profile-25000.json --rows 25000
```

The write profile JSON records averages, medians, and raw samples for each stage so outlier runs are visible. OfficeIMO timing hooks also add AutoFit sub-stages (`BuildPlan`, `CalculateWidths`, and `ApplyWidths`) to make large-sheet tuning more targeted.

The harness configures OfficeIMO with `Execution.SaveWorksheetAfterAutoFit = false`, matching the normal report-export pattern where all worksheet mutations are committed once when the document is saved or disposed.
