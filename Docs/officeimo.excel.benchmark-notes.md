# OfficeIMO.Excel Benchmark Notes

The benchmark harness is intended to measure representative Excel workloads across comparable library surfaces. The goal is to keep the suite broad, repeatable, and useful for engineering decisions instead of optimizing around one vendor-specific claim.

## Comparison Suite

Use `comparison-suite` for the broad proof run:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- comparison-suite --out-dir .\Docs\benchmarks\comparison-current --row-set 2500,25000 --skip-legacy-epplus --warmup 20 --iterations 9
```

The release-style runner uses twenty warmups because isolated ARM64 reruns
showed tiered-PGO transitions still occurring after fifteen invocations. This
keeps compilation transitions out of the measured samples for every library.
`--skip-legacy-epplus` keeps current EPPlus coverage and omits
only the isolated EPPlus 4.x process, whose AutoFit path requires `libgdiplus`
on macOS.

The suite writes:

- `officeimo.excel.comparison-speed-<rows>.json`
- `officeimo.excel.comparison-package-<rows>.json`
- `officeimo.excel.comparison-dense-helloworld-<rows>.json`
- `officeimo.excel.comparison-summary.md`
- `officeimo.excel.comparison-summary.csv`
- `officeimo.excel.comparison-summary.json`
- `officeimo.excel.comparison-suite-manifest.json`

The summary artifacts are the preferred wrap-up view. They include one row per scenario/library with mean, median, standard deviation, standard error, ratio to OfficeIMO, ratio to the best result, allocation, allocation ratio, package size, package-size ratio, outcome, and package-part metrics when package profiling is available.

Package-profile lanes reopen the generated workbook package after timing and fail the run if a write scenario produces an empty package, a package without workbook parts, or a package without worksheet rows and cells. This keeps package-size and write-speed comparisons tied to real `.xlsx` output instead of byte-count-only success.

## Scenario Coverage

The suite covers:

- table/report exports with and without AutoFit
- DataSet, DataTable, IDataReader, typed object, fluent row, and cell rectangle writes
- headerless and sparse table-shaped writes
- append-style writes
- dense range reads
- bounded top-of-sheet reads
- streaming range reads
- sparse row and column reads
- typed object materialization
- formula text reads
- repeated shared-string writes and reads
- dense `HelloWorld` grid reads
- package size and package-part breakdowns

## Interpretation

Treat the comparison summary as the engineering decision table. It is intentionally broader than a single benchmark claim and should be used to decide what to optimize next.

Treat elapsed-time differences below 5% as practical ties. Small leads inside that band are not reported as wins.

The lightweight runner reports standard deviation and standard error from local repeated samples. It also captures allocation with `GC.GetAllocatedBytesForCurrentThread`. If a public benchmark table needs BenchmarkDotNet's exact `Error` column, first use the comparison suite to choose the scenarios, then run targeted BenchmarkDotNet jobs for those claims.

## Current Direction

OfficeIMO's identity is fast workbook authoring without giving up normal document features. The implementation should keep fast paths internal to OfficeIMO and preserve the standard APIs for callers. Size work should remain evidence-driven: use the package profile before changing defaults, and prefer adaptive or opt-in size behavior when a change risks the current export performance.
