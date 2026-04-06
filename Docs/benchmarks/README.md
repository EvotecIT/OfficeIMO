# Excel Benchmark Artifacts

This folder stores small, committed benchmark artifacts for `OfficeIMO.Excel`.

- `officeimo.excel.snapshot-YYYY-MM-DD.json`: lightweight end-to-end scenario snapshot for write, read, and round-trip flows
- `officeimo.excel.write-profile-YYYY-MM-DD.json`: write-stage breakdown intended to highlight where optimization work should focus

Both artifact types now store raw sample lists and medians in addition to averages so noisy runs are easier to spot.

Generate them from the benchmark harness:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --snapshot .\Docs\benchmarks\officeimo.excel.snapshot-YYYY-MM-DD.json
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --profile-write .\Docs\benchmarks\officeimo.excel.write-profile-YYYY-MM-DD.json
```
