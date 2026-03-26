# OfficeIMO.Excel.GoogleSheets

Excel to Google Sheets planning, batch compilation, and export helpers built on top of `OfficeIMO.Excel` and `OfficeIMO.GoogleWorkspace`.

## Highlights

- Build a translation plan before sending anything to Google Sheets.
- Export to a new spreadsheet or replace an existing spreadsheet via `ExistingFileId`.
- Apply session-level default Drive and folder placement through `GoogleWorkspaceSessionOptions`.
- Automatically retry transient Google API failures and surface successful retries in `TranslationReport`.

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
        RetryBaseDelay = TimeSpan.FromMilliseconds(250),
        RetryMaxDelay = TimeSpan.FromSeconds(10),
        DiagnosticSink = entry => Console.WriteLine($"{entry.Severity}: {entry.Feature} [{entry.FailureKind}] - {entry.Message}"),
    });

var options = new GoogleSheetsSaveOptions {
    Title = "Quarterly revenue",
};

var plan = workbook.CreateGoogleSheetsTranslationPlan(options);
foreach (var notice in plan.Report.Notices) {
    Console.WriteLine($"{notice.Severity}: {notice.Feature} - {notice.Message}");
}

var result = await workbook.ExportToGoogleSheetsAsync(session, options);
Console.WriteLine(result.SpreadsheetId);
Console.WriteLine(result.WebViewLink);
```

`StaticAccessTokenCredentialSource` is provided by [OfficeIMO.GoogleWorkspace](../OfficeIMO.GoogleWorkspace/README.md).

## Handling failures

```csharp
try {
    var result = await workbook.ExportToGoogleSheetsAsync(session, options);
    Console.WriteLine(result.SpreadsheetId);
} catch (GoogleWorkspaceExportException exception) {
    foreach (var entry in exception.ToDiagnosticEntries()) {
        Console.WriteLine($"{entry.Severity}: {entry.Feature} [{entry.FailureKind}] - {entry.Message}");
    }
}
```

## Replace an existing spreadsheet

```csharp
var result = await workbook.ExportToGoogleSheetsAsync(
    session,
    new GoogleSheetsSaveOptions {
        Title = "Quarterly revenue",
        Location = new GoogleDriveFileLocation {
            ExistingFileId = "spreadsheet-id",
        },
    });
```

## Operational notes

- If `Location.FolderId` is omitted, the exporter falls back to `GoogleWorkspaceSessionOptions.DefaultFolderId`.
- If `Location.DriveId` is omitted, the exporter falls back to `GoogleWorkspaceSessionOptions.DefaultDriveId`.
- Successful retries are recorded as `ApiRetries` notices on `GoogleSpreadsheetReference.Report`.
- `GoogleWorkspaceSessionOptions.DiagnosticSink` can stream retry and failure diagnostics while the export is still running.
- Failed exports throw `GoogleWorkspaceExportException`, which preserves `FailureKind` and the collected `TranslationReport`.
- Caller cancellation throws `GoogleWorkspaceExportCanceledException`, so cancellation can still be handled separately from timeout or API instability.
- API failures prefer parsed Google status and reason codes when Google returns a structured JSON error body.
- Use `CreateGoogleSheetsBatch(...)` when you want the provider-neutral request model without making network calls.
