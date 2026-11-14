---
title: "Export and Import Google Sheets"
description: "Plan an Excel-to-Google-Sheets export offline, create a spreadsheet from an ExcelDocument, import a Google Sheet, and replace an existing one safely."
order: 30
packages:
  - OfficeIMO.Excel.GoogleSheets
  - OfficeIMO.GoogleWorkspace
cmdlets:
  - Export-OfficeExcelGoogleSpreadsheet
  - New-OfficeExcelGoogleSheetsOptions
  - New-OfficeGoogleWorkspaceSession
limits: "Excel print layout has no Google Sheets equivalent and is not exported."
search_terms:
  - excel to google sheets
  - xlsx to google sheets
  - import google sheet
  - google sheets api
verified: 2026-11-14
---

`OfficeIMO.Excel.GoogleSheets` translates an `ExcelDocument` into Google Sheets API requests and imports spreadsheets back. Value-heavy writes are kept apart from structural requests and both are chunked. Formula handling is explicit: fail, keep with a warning, or write the cached value.

## Check what an export would change

### C#

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Excel.GoogleSheets;
using OfficeIMO.GoogleWorkspace;

using ExcelDocument workbook = ExcelDocument.Load("revenue.xlsx");

GoogleSheetsTranslationPlan plan = workbook.BuildGoogleSheetsPlan(new GoogleSheetsSaveOptions { Title = "Quarterly revenue" });

foreach (TranslationNotice notice in plan.Report.Notices) {
    Console.WriteLine($"{notice.Severity} {notice.Code} {notice.Path}: {notice.Message}");
}
```

### PowerShell

```powershell
$options = New-OfficeExcelGoogleSheetsOptions -Title 'Quarterly revenue' -UnsupportedFormula PreserveWithWarning
Export-OfficeExcelGoogleSpreadsheet -Path .\revenue.xlsx -Options $options -Mode Plan
```

## Create a Google Sheet from Excel

### C#

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Excel.GoogleSheets;
using OfficeIMO.GoogleWorkspace;

using var credentials = GoogleServiceAccountCredentialSource.FromFile("service-account.json");
var session = new GoogleWorkspaceSession(credentials, new GoogleWorkspaceSessionOptions { DefaultFolderId = "folder-id" });

using ExcelDocument workbook = ExcelDocument.Create();
ExcelSheet summary = workbook.AddSheet("Summary");
summary.Cell(1, 1).SetValue("Quarter");
summary.Cell(1, 2).SetValue("Revenue");
summary.Cell(2, 1).SetValue("Q1");
summary.Cell(2, 2).SetValue(125000);

var options = new GoogleSheetsSaveOptions { Title = "Quarterly revenue" };
options.Identity.WriteDeveloperMetadata = true;

GoogleSpreadsheetReference created = await workbook.ExportToGoogleSheetsAsync(session, options);
Console.WriteLine(created.WebViewLink);
```

### PowerShell

```powershell
$session = New-OfficeGoogleWorkspaceSession -KeyPath .\Secrets\reporting.json -Scope Spreadsheets, DriveFile
try {
    $options = New-OfficeExcelGoogleSheetsOptions -Title 'Quarterly revenue' -FolderId 'folder-id'
    Export-OfficeExcelGoogleSpreadsheet -Path .\revenue.xlsx -Session $session -Options $options -FailOnLoss
} finally {
    $session.CredentialSource.Dispose()
}
```

## Import a Google Sheet and replace it safely

Native import accepts ranges and a Sheets field mask; `DriveExport` mode converts to XLSX for broader fidelity. For a replacement, pass the Drive version you imported.

### C#

```csharp
using OfficeIMO.Excel.GoogleSheets;
using OfficeIMO.GoogleWorkspace;

using var credentials = GoogleServiceAccountCredentialSource.FromFile("service-account.json");
var session = new GoogleWorkspaceSession(credentials);

GoogleSheetsImportResult current = await session.ImportGoogleSheetAsync(
    "spreadsheet-id",
    new GoogleSheetsImportOptions { Mode = GoogleWorkspaceImportMode.Native, Ranges = new[] { "Summary!A1:Z500" } });

using (current.Document) {
    var replace = new GoogleSheetsSaveOptions {
        Location = new GoogleDriveFileLocation { ExistingFileId = current.Source.FileId },
        Replace = new GoogleSheetsReplaceOptions { ExpectedDriveVersion = current.Source.DriveVersion }
    };
    await current.Document.ExportToGoogleSheetsAsync(session, replace);
}
```

### PowerShell

```powershell
$options = New-OfficeExcelGoogleSheetsOptions -SpreadsheetId $spreadsheetId -ExpectedDriveVersion $driveVersion
Export-OfficeExcelGoogleSpreadsheet -Path .\revenue.xlsx -Session $session -Options $options
```
