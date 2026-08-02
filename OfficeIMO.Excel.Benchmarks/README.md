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

The hash-pinned Mark Pflug 65K-record read comparisons are available as focused
BenchmarkDotNet classes:

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --filter "*MarkPflug65KXlsxBenchmarks*"
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --filter "*MarkPflug65KXlsBenchmarks*"
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --filter "*MarkPflug65KXlsbBenchmarks*"
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --filter "*XlsbOfficeIMOPipelineBenchmarks*"
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --profile-markpflug65k-xlsb-officeimo 100
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --profile-markpflug65k-xlsb-sylvan 100
```

The XLS lane includes OfficeIMO, Sylvan.Data.Excel, and ExcelDataReader. The
XLSX lane includes OfficeIMO, Sylvan.Data.Excel, ExcelDataReader,
ClosedXML, EPPlus, and MiniExcel. The XLSB lane includes only the compatible
readers: OfficeIMO, Sylvan.Data.Excel, and ExcelDataReader. Every implementation
reads the same fourteen typed columns and must produce the same row count, cell
count, and deterministic payload observation before a measurement is accepted.
Write-focused libraries such as LargeXlsx and SpreadCheetah remain in the write
suites and are not applicable to these read lanes. Every library is a peer in
the matrix; no implementation is framed as an opponent or universal baseline.
The OfficeIMO-only XLSB pipeline diagnostic is not part of that library matrix;
it isolates public multi-sheet dispatch from the underlying worksheet reader
when profiling a demonstrated OfficeIMO bottleneck.

The suite keeps materially different contracts in separate lanes. Compact
writers omit explicit cell references for forward-only throughput, while the
normal OfficeIMO writer preserves the editable worksheet model. Shared-string
reads distinguish forward-only scans from rectangular materialization, and
DataTable reads distinguish automatic type inference from a caller-prepared
typed schema. These lanes should not be collapsed into one ranking.

String-storage readers are compared on separately validated shared-string and
inline-string fixtures. In a 25,000-row Windows run with twenty warmups and nine
measurements, OfficeIMO's forward reader measured 15.98 ms for shared strings
and 20.71 ms for inline strings; Sylvan.Data.Excel measured 20.46 ms and 25.50
ms respectively. These are workstation measurements rather than universal
constants, so rerun the focused scenarios before making a release claim.

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
marker-delimited block. Benchmark execution is local and is not scheduled in CI.

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

## Historical generated workstation snapshot

This PSPublishModule-managed snapshot is retained for reproducibility, not as
the current cross-platform product ranking. The rows combine raw data paths
with feature-bearing workbook work and only compare libraries that expose a
directly comparable public API. Lower is
faster within a row only; do not combine rows into one library ranking. Results
vary by machine, runtime, package version, workload, warm-up, and options.
Treat differences below 5% as ties. Use the hash-pinned CSV/XLSX/XLSB website
matrix for current platform-separated evidence.

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
