# OfficeIMO.Word.GoogleDocs

Word to Google Docs planning, batch compilation, and export helpers built on top of `OfficeIMO.Word` and `OfficeIMO.GoogleWorkspace`.

## Highlights

- Build a translation plan before sending anything to Google Docs.
- Export to a new document or replace an existing document via `ExistingFileId`.
- Apply session-level default Drive and folder placement through `GoogleWorkspaceSessionOptions`.
- Automatically retry transient Google API failures and surface successful retries in `TranslationReport`.

## Quick start

```csharp
using OfficeIMO.GoogleWorkspace;
using OfficeIMO.Word;
using OfficeIMO.Word.GoogleDocs;

using var document = WordDocument.Create("proposal.docx");
document.AddParagraph("Quarterly business review").SetStyle(WordParagraphStyles.Heading1);
document.AddParagraph("Highlights and next steps.");

var session = new GoogleWorkspaceSession(
    new StaticAccessTokenCredentialSource("<google-access-token>"),
    new GoogleWorkspaceSessionOptions {
        DefaultFolderId = "documents-folder-id",
        MaxRetryCount = 5,
        RetryBaseDelay = TimeSpan.FromMilliseconds(250),
        RetryMaxDelay = TimeSpan.FromSeconds(10),
        DiagnosticSink = entry => Console.WriteLine($"{entry.Severity}: {entry.Feature} [{entry.FailureKind}] - {entry.Message}"),
    });

var options = new GoogleDocsSaveOptions {
    Title = "Quarterly business review",
};

var plan = document.CreateGoogleDocsTranslationPlan(options);
foreach (var notice in plan.Report.Notices) {
    Console.WriteLine($"{notice.Severity}: {notice.Feature} - {notice.Message}");
}

var result = await document.ExportToGoogleDocsAsync(session, options);
Console.WriteLine(result.DocumentId);
Console.WriteLine(result.WebViewLink);
```

`StaticAccessTokenCredentialSource` is provided by [OfficeIMO.GoogleWorkspace](../OfficeIMO.GoogleWorkspace/README.md).

## Handling failures

```csharp
try {
    var result = await document.ExportToGoogleDocsAsync(session, options);
    Console.WriteLine(result.DocumentId);
} catch (GoogleWorkspaceExportException exception) {
    foreach (var entry in exception.ToDiagnosticEntries()) {
        Console.WriteLine($"{entry.Severity}: {entry.Feature} [{entry.FailureKind}] - {entry.Message}");
    }
}
```

## Replace an existing document

```csharp
var result = await document.ExportToGoogleDocsAsync(
    session,
    new GoogleDocsSaveOptions {
        Title = "Quarterly business review",
        Location = new GoogleDriveFileLocation {
            ExistingFileId = "document-id",
        },
    });
```

## Operational notes

- If `Location.FolderId` is omitted, the exporter falls back to `GoogleWorkspaceSessionOptions.DefaultFolderId`.
- If `Location.DriveId` is omitted, the exporter falls back to `GoogleWorkspaceSessionOptions.DefaultDriveId`.
- Successful retries are recorded as `ApiRetries` notices on `GoogleDocumentReference.Report`.
- `GoogleWorkspaceSessionOptions.DiagnosticSink` can stream retry and failure diagnostics while the export is still running.
- Failed exports throw `GoogleWorkspaceExportException`, which preserves `FailureKind` and the collected `TranslationReport`.
- Caller cancellation throws `GoogleWorkspaceExportCanceledException`, so cancellation can still be handled separately from timeout or API instability.
- API failures prefer parsed Google status and reason codes when Google returns a structured JSON error body.
- Use `CreateGoogleDocsBatch(...)` when you want the provider-neutral request model without making network calls.
