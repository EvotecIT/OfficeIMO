# OfficeIMO.Excel.GoogleSheets - Excel and Google Sheets translation

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Excel.GoogleSheets)](https://www.nuget.org/packages/OfficeIMO.Excel.GoogleSheets)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Excel.GoogleSheets?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Excel.GoogleSheets)

`OfficeIMO.Excel.GoogleSheets` provides bidirectional Excel and Google Sheets translation with formula compatibility policy, sparse/chunked writes, native advanced objects, import, fidelity preflight, and format-specific diff planning.

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
var sheet = workbook.AddWorksheet("Summary");
sheet.CellValue(1, 1, "Quarter");
sheet.CellValue(1, 2, "Revenue");
sheet.CellValue(2, 1, "Q1");
sheet.CellValue(2, 2, 125000);

var receipts = new List<GoogleWorkspaceOperationReceipt>();
var sessionOptions = new GoogleWorkspaceSessionOptions {
    ExpectedAccount = "service-account@project.iam.gserviceaccount.com",
    DefaultFolderId = "reports-folder-id",
    OperationReceiptSink = receipts.Add,
};
sessionOptions.OperationPolicyProvider = context => new GoogleWorkspaceOperationPolicy(
    sessionOptions.ExpectedAccount!, context.RequiredScopes, context.Target,
    context.RevisionPreconditionKind switch {
        GoogleWorkspaceRevisionPreconditionKind.ResourceAbsentCreate => GoogleWorkspaceOperationPolicy.ResourceAbsentForCreateRevision,
        GoogleWorkspaceRevisionPreconditionKind.PayloadRevision => context.AdapterExpectedRevision!,
        GoogleWorkspaceRevisionPreconditionKind.ResumableSessionState => context.AdapterExpectedRevision!,
        GoogleWorkspaceRevisionPreconditionKind.Unavailable => GoogleWorkspaceOperationPolicy.ExplicitlyUnversionedRevision("API exposes no conditional revision"),
        _ => "\"observed-google-etag\"",
    }, context.MaxRetryCount, context.MaxRetryElapsedTime, context.RateLimitPolicy,
    GoogleWorkspaceDataLossDecision.RejectPotentialLoss);
var session = new GoogleWorkspaceSession(
    GoogleServiceAccountCredentialSource.FromFile("service-account.json", sessionOptions),
    sessionOptions);

var options = new GoogleSheetsSaveOptions {
    Title = "Quarterly revenue"
};

var plan = workbook.BuildGoogleSheetsPlan(options);
var result = await workbook.ExportToGoogleSheetsAsync(session, options);

Console.WriteLine(result.SpreadsheetId);
Console.WriteLine(result.WebViewLink);
```

## What it does

- Builds a translation plan before network export.
- Exports to a new Google Sheets spreadsheet or replaces an existing spreadsheet through `Location.ExistingFileId`.
- Imports through native Sheets resources or Drive-exported XLSX.
- Maps supported formulas, styles, filters, validation, protection, conditional formatting, charts, pivots, outlines, and native tables with explicit support boundaries.
- Uses values batching for value-heavy writes and structural batches for formats and objects.
- Supports range/field-mask native import, version evidence, and pre-apply diff planning.
- Uses session-level default Drive and folder placement.
- Preserves retry, warning, and failure detail through `TranslationReport`.
- Throws Google Workspace export exceptions that retain failure category and diagnostics.

## Boundaries

- Workbook modeling belongs in `OfficeIMO.Excel`.
- Credentials, sessions, retry, Drive placement, and report primitives belong in `OfficeIMO.GoogleWorkspace`.
- This package owns Excel/Google Sheets mapping, import, safe replacement, and format-specific diff planning.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`, plus `net472` on Windows.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** `System.Text.Json` plus platform HTTP/cryptography through `OfficeIMO.GoogleWorkspace`; no Google client SDK.
- **OfficeIMO:** `OfficeIMO.Excel` and `OfficeIMO.GoogleWorkspace` own the workbook model, session, translation plan, and reports.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
