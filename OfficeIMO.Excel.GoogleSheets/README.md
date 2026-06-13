# OfficeIMO.Excel.GoogleSheets - Excel to Google Sheets export

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Excel.GoogleSheets)](https://www.nuget.org/packages/OfficeIMO.Excel.GoogleSheets)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Excel.GoogleSheets?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Excel.GoogleSheets)

`OfficeIMO.Excel.GoogleSheets` builds translation plans and export requests for sending `OfficeIMO.Excel` workbooks to Google Sheets through `OfficeIMO.GoogleWorkspace`.

## Install

```powershell
dotnet add package OfficeIMO.Excel.GoogleSheets
```

## Quick start

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Excel.GoogleSheets;
using OfficeIMO.GoogleWorkspace;

using var workbook = ExcelDocument.Create("report.xlsx");
var sheet = workbook.AddWorkSheet("Summary");
sheet.CellValue(1, 1, "Quarter");
sheet.CellValue(1, 2, "Revenue");
sheet.CellValue(2, 1, "Q1");
sheet.CellValue(2, 2, 125000);

var session = new GoogleWorkspaceSession(
    new StaticAccessTokenCredentialSource("<google-access-token>"),
    new GoogleWorkspaceSessionOptions {
        DefaultFolderId = "reports-folder-id",
        MaxRetryCount = 5,
        DiagnosticSink = entry => Console.WriteLine($"{entry.Severity}: {entry.Feature} - {entry.Message}")
    });

var options = new GoogleSheetsSaveOptions {
    Title = "Quarterly revenue"
};

var plan = workbook.CreateGoogleSheetsTranslationPlan(options);
var result = await workbook.ExportToGoogleSheetsAsync(session, options);

Console.WriteLine(result.SpreadsheetId);
Console.WriteLine(result.WebViewLink);
```

## What it does

- Builds a translation plan before network export.
- Exports to a new Google Sheets spreadsheet or replaces an existing spreadsheet through `Location.ExistingFileId`.
- Uses session-level default Drive and folder placement.
- Preserves retry, warning, and failure detail through `TranslationReport`.
- Throws Google Workspace export exceptions that retain failure category and diagnostics.

## Boundaries

- Workbook modeling belongs in `OfficeIMO.Excel`.
- Credentials, sessions, retry, Drive placement, and report primitives belong in `OfficeIMO.GoogleWorkspace`.
- This package owns Excel-to-Google-Sheets mapping and export request construction.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)
