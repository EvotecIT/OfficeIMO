# OfficeIMO.Excel.Benchmarks.NPOI

Opt-in comparison runner for OfficeIMO.Excel and NPOI.

This project is intentionally not included in `OfficeIMO.sln`. Normal solution
restore and build should not pull NPOI. Run it explicitly when NPOI comparison
evidence is wanted. The project sets NPOI's OSMF EULA acceptance property for
this benchmark-only runner; OfficeIMO runtime projects must not reference NPOI.
The NPOI 2.8.0 OSMF terms target revenue-generating use of NPOI binaries, not a
runtime dependency in OfficeIMO; keep this runner local, explicit, and opt-in.
SkiaSharp is referenced explicitly here because NPOI HSSF comment/drawing reads
load it at runtime while NPOI's package metadata excludes those runtime assets
from the transitive reference.

The JSON `Metric` field is a lightweight anti-elision checksum for each measured
read/write path. Equal metrics are a useful parity signal for the scalar
cell-value, formula, AutoFilter-range, and style-signal lanes; different metrics in richer
metadata lanes should be read as "the benchmark exercised the path and validated
its counts/ranges", not as a full feature-diff verdict.

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
- `xls-read-metadata`: read comments, hyperlinks, merged ranges, and list data
  validations from an HSSF-generated `.xls` workbook through NPOI and
  OfficeIMO's legacy XLS importer.
- `xls-read-conditional-formatting`: read cell-is and formula conditional
  formatting rules from an HSSF-generated `.xls` workbook through NPOI and
  OfficeIMO's legacy XLS importer.
- `xls-read-autofilter-range`: read basic AutoFilter range/drop-down metadata
  from an HSSF-generated `.xls` workbook through NPOI and OfficeIMO's legacy
  XLS importer.
- `xls-read-styles`: read font, fill, border, number-format, and alignment
  style signals from an HSSF-generated `.xls` workbook through NPOI and
  OfficeIMO's legacy XLS importer.
- `xls-read-pictures`: read embedded PNG picture signals from an HSSF-generated
  `.xls` workbook through NPOI and OfficeIMO's preserve-only drawing/image
  metadata.

Do not add OfficeIMO-specific template, preflight, PDF, package-copy,
direct-package, or report-workflow scenarios here unless NPOI has an equivalent
native API shape for the same work.
