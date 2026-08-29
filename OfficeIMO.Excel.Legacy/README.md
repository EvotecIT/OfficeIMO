# OfficeIMO.Excel.Legacy

`OfficeIMO.Excel.Legacy` safely reads selected Lotus 1-2-3, Quattro Pro, Multiplan, and Microsoft Works spreadsheet sources into the normal `OfficeIMO.Excel.ExcelDocument` model.

```csharp
using OfficeIMO.Excel.Legacy;

using LegacySpreadsheetImportResult imported = LegacySpreadsheetImporter.Import("archive.wk1");
Console.WriteLine(imported.Report.Quality);
foreach (OfficeCompatibilityFinding finding in imported.Report.Findings) {
    Console.WriteLine($"{finding.Code}: {finding.Message}");
}

imported.Document.Save("archive.xlsx");
```

The importer is deliberately read-only. It never saves back to a legacy format, executes macros, activates embedded objects, or resolves and refreshes external links. Each result states whether recovery was structured or salvage quality and includes explicit feature-level loss diagnostics. The existing OfficeIMO Excel converter packages can export the returned workbook to ODS, CSV, HTML, or PDF.

Supported early WK-family record streams recover cell addresses, text, numbers, cached formula results, basic label alignment, source names as metadata, and chart-record metadata. Unsupported formula token streams are never guessed: the cached value is used and the loss report records that decision. Later compound profiles use bounded salvage and identify the missing workbook structures explicitly.

Path, stream, and byte-array inputs share the same limits, cancellation, detection, and loss-report contracts.
