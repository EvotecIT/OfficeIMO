# OfficeIMO.Excel.Benchmarks.NPOI

Opt-in benchmark verification runner for OfficeIMO.Excel and NPOI.

This project is not included in `OfficeIMO.sln`, so normal solution restore and
build do not pull NPOI. Run it explicitly when external benchmark verification
evidence is wanted. The project sets NPOI's OSMF EULA acceptance property for
this runner.
SkiaSharp is referenced explicitly here because NPOI HSSF comment/drawing reads
load it at runtime while NPOI's package metadata excludes those runtime assets
from the transitive reference.

The JSON `Metric` field is a lightweight anti-elision checksum for each measured
read/write path. Equal metrics are a useful verification signal for the scalar
cell-value, formula, AutoFilter-range, and style-signal lanes; different metrics in richer
metadata lanes should be read as "the benchmark exercised the path and validated
its counts/ranges", not as a full feature-diff verdict.

```powershell
dotnet run -c Release --project .\OfficeIMO.Excel.Benchmarks.NPOI\OfficeIMO.Excel.Benchmarks.NPOI.csproj -- --rows 2500 --warmup 1 --iterations 3 --out .\Docs\benchmarks\npoi-comparison.json
```

Use the paired runner for competitive XLS write measurements. It validates every
OfficeIMO result through both OfficeIMO and NPOI outside the timed interval,
validates NPOI output through OfficeIMO, and alternates ABBA/BAAB execution order.
Run each processor-affinity domain separately on hybrid or cache-asymmetric CPUs:

```powershell
dotnet run -c Release --project .\OfficeIMO.Excel.Benchmarks.NPOI\OfficeIMO.Excel.Benchmarks.NPOI.csproj -- --paired-xls-write --rows 25000 --warmup 12 --iterations 40 --affinity 0xFFFF --priority High
dotnet run -c Release --project .\OfficeIMO.Excel.Benchmarks.NPOI\OfficeIMO.Excel.Benchmarks.NPOI.csproj -- --paired-xls-write --rows 25000 --warmup 12 --iterations 40 --affinity 0xFFFF0000 --priority High
```

The first lanes are deliberately plain workbook operations:

- `xlsx-write-cellvalues`: plain row/cell writes to `.xlsx`.
- `xls-write-cellvalues`: plain row/cell writes to `.xls`; use the paired runner
  above when drawing a performance conclusion.
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
