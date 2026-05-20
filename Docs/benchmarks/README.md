# Excel Benchmark Artifacts

This folder stores small, committed benchmark artifacts for `OfficeIMO.Excel`.

- `officeimo.excel.snapshot-YYYY-MM-DD.json`: lightweight end-to-end scenario snapshot for write, read, and round-trip flows
- `officeimo.excel.write-profile-YYYY-MM-DD.json`: write-stage breakdown intended to highlight where optimization work should focus
- `officeimo.excel.read-profile-YYYY-MM-DD.json`: read-stage comparison for automatic, forced sequential, and forced parallel range conversion
- `officeimo.excel.library-comparison.json`: local opt-in comparison across matching library surfaces
- `comparison-current\officeimo.excel.comparison-suite-manifest.json`: release-style suite manifest that points to the speed, package, dense `HelloWorld`, and summary artifacts
- `comparison-current\officeimo.excel.comparison-summary.md|csv|json`: one-table decision summary with mean, standard deviation, standard error, speed ratios, allocation ratios, package-size ratios, winners, losses, and package-part metrics

Benchmark artifacts now store raw sample lists and medians in addition to averages so noisy runs are easier to spot. Comparison artifacts also include mean/median allocation samples captured with `GC.GetAllocatedBytesForCurrentThread`. Write profiles include OfficeIMO timing-hook sub-stages such as AutoFit plan, width calculation, and width application when those hooks are emitted.
OfficeIMO benchmark runs use the report-export AutoFit mode (`Execution.SaveWorksheetAfterAutoFit = false`) so worksheet changes are committed once at document save/dispose time instead of after each AutoFit operation.

Generate them from the benchmark harness:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --snapshot .\Docs\benchmarks\officeimo.excel.snapshot-YYYY-MM-DD.json
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --profile-write .\Docs\benchmarks\officeimo.excel.write-profile-YYYY-MM-DD.json
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --profile-read .\Docs\benchmarks\officeimo.excel.read-profile-YYYY-MM-DD.json
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- comparison-suite --out-dir .\Docs\benchmarks\comparison-current --row-set 2500,25000 --warmup 1 --iterations 3
```

Commands that write artifacts also accept `--out`, `--output`, or `--output-path` when an explicit output path is clearer than the positional form.

Add `--website-data .\Website\data\benchmarks.json` to a snapshot run when the public benchmark table should be refreshed from the same measured values.

Short aliases can be used when the default `Docs\benchmarks` output path is sufficient:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- snapshot --rows 2500
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- write-profile --rows 25000
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- read-profile --rows 2500
```

Both commands default to 2,500 rows and accept `--rows <count>` for larger tiers, for example:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --snapshot .\Docs\benchmarks\officeimo.excel.snapshot-25000-YYYY-MM-DD.json --rows 25000
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --profile-write .\Docs\benchmarks\officeimo.excel.write-profile-25000-YYYY-MM-DD.json --rows 25000
```
