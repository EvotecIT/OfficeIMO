---
title: Excel and Google Sheets
description: Plan, export, import, and safely replace Google Sheets with OfficeIMO.Excel.
order: 30
---

Install `OfficeIMO.Excel.GoogleSheets`. The translator separates value-heavy writes from structural requests and chunks both paths.

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Excel.GoogleSheets;
using OfficeIMO.GoogleWorkspace;

using ExcelDocument workbook = ExcelDocument.Create();
var summary = workbook.AddWorksheet("Summary");
summary.CellValue(1, 1, "Quarter");
summary.CellValue(1, 2, "Revenue");
summary.CellValue(2, 1, "Q1");
summary.CellValue(2, 2, 125000);

var options = new GoogleSheetsSaveOptions { Title = "Quarterly revenue" };
options.Identity.WriteDeveloperMetadata = true;
GoogleSheetsTranslationPlan plan = workbook.BuildGoogleSheetsPlan(options);
GoogleSpreadsheetReference created = await workbook.ExportToGoogleSheetsAsync(session, options);
```

For an existing spreadsheet, import first and require its Drive version:

```csharp
GoogleSheetsImportResult current = await session.ImportGoogleSheetAsync(
    "spreadsheet-id",
    new GoogleSheetsImportOptions { Mode = GoogleSheetsImportMode.Native, Ranges = new[] { "Summary!A1:Z500" } });

using (current.Document) {
    var replace = new GoogleSheetsSaveOptions {
        Location = new GoogleDriveFileLocation { ExistingFileId = current.Source.FileId },
        Replace = new GoogleSheetsReplaceOptions { ExpectedDriveVersion = current.Source.DriveVersion }
    };
    await workbook.ExportToGoogleSheetsAsync(session, replace);
}
```

Formula compatibility is explicit: error, preserve with warning, or use a cached value. Supported formatting, validation, filters, protection, conditional rules, charts, pivots, outlines, and native tables are discoverable in the [support matrix](/docs/google-workspace/support/). Native import accepts ranges and a Sheets field mask; `DriveExport` converts to XLSX for the broad fallback.
