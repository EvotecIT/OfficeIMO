# OfficeIMO.Excel.Benchmarks.NPOI

Opt-in comparison runner for OfficeIMO.Excel and NPOI.

This project is intentionally not included in `OfficeIMO.sln`. Normal solution
restore and build should not pull NPOI. Run it explicitly when NPOI comparison
evidence is wanted. The project sets NPOI's OSMF EULA acceptance property for
this benchmark-only runner; OfficeIMO runtime projects must not reference NPOI.

```powershell
dotnet run -c Release --project .\OfficeIMO.Excel.Benchmarks.NPOI\OfficeIMO.Excel.Benchmarks.NPOI.csproj -- --rows 2500 --warmup 1 --iterations 3 --out .\Docs\benchmarks\npoi-comparison.json
```

The first lanes are deliberately natural to NPOI:

- `xlsx-write-cellvalues`: plain row/cell writes to `.xlsx`.
- `xlsx-read-cellvalues`: plain row/cell reads from the same `.xlsx` shape.
- `xls-read-cellvalues`: read an HSSF-generated `.xls` workbook through NPOI and
  OfficeIMO's legacy XLS importer.
- `xls-read-formulas`: read formula text and cached values from an
  HSSF-generated `.xls` workbook through NPOI and OfficeIMO's legacy XLS
  importer.

Do not add OfficeIMO-specific template, preflight, PDF, package-copy,
direct-package, or report-workflow scenarios here unless NPOI has an equivalent
native API shape for the same work.
