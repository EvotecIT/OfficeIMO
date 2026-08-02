# OfficeIMO.Reader.Excel

Excel workbook ingestion for `OfficeIMO.Reader.Core`. Install this package when a reader needs XLSX, XLSM, XLSB, or XLS support, then compose it explicitly:

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Excel;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddExcelHandler()
    .Build();
```

The package also owns the thin Excel/CSV adapter. Parsing and schema inference
come from `OfficeIMO.CSV`; worksheet insertion and saving come from
`OfficeIMO.Excel`:

```csharp
using OfficeIMO.CSV;
using OfficeIMO.Excel;
using OfficeIMO.Reader.Excel;

using var workbook = ExcelDocument.Create();
ExcelCsvImportResult imported = workbook.ImportCsvFile("sales.csv");
workbook[imported.SheetName].SaveAsCsv("sales-roundtrip.csv");
workbook.Save("sales.xlsx");
```

Use `ImportCsvText`, `ImportCsv(Stream)`, or `ImportCsv(CsvDocument)` for other
source shapes. `ExcelCsvImportOptions.LoadOptions` and `ReaderOptions` expose
the canonical CSV option types instead of a second delimiter or conversion
model.
