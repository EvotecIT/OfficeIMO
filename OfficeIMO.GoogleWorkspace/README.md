# OfficeIMO.GoogleWorkspace

Shared session, credential, retry, Drive-location, and translation-report primitives for the Google Workspace extension packages.

## What this package provides

- `IGoogleWorkspaceCredentialSource` so applications can plug in their own OAuth or service-account token acquisition.
- `StaticAccessTokenCredentialSource` for scenarios where the app already has a Google access token.
- `DelegateGoogleWorkspaceCredentialSource` for plugging in an existing token service with a single delegate.
- `GoogleServiceAccountCredentialSource` for built-in service-account JWT bearer exchange, including domain-wide delegation via session options.
- `GoogleWorkspaceSession` as the shared runtime container for credentials and session defaults.
- `GoogleWorkspaceSessionOptions` for app identity, timeout, retry policy, and default Drive placement.
- `GoogleDriveFileLocation` for folder, shared-drive, and existing-file targeting.
- `TranslationReport` / `TranslationNotice` for export diagnostics, warnings, and retry visibility.
- `GoogleWorkspaceExportException` for failed exports that still preserve `TranslationReport` and a high-level failure category.
- `GoogleWorkspaceExportCanceledException` for caller-triggered cancellations that still preserve `TranslationReport`.
- `ToDiagnosticEntries()` helpers for turning reports and export exceptions into structured log-ready entries.
- `GoogleWorkspaceSessionOptions.DiagnosticSink` for streaming retry/auth/API diagnostics while an export is running.

## Quick start

```csharp
using OfficeIMO.GoogleWorkspace;

var session = new GoogleWorkspaceSession(
    new StaticAccessTokenCredentialSource("<google-access-token>"),
    new GoogleWorkspaceSessionOptions {
        ApplicationName = "OfficeIMO Samples",
        DefaultDriveId = "shared-drive-id",
        DefaultFolderId = "reports-folder-id",
        MaxRetryCount = 5,
        RetryBaseDelay = TimeSpan.FromMilliseconds(250),
        RetryMaxDelay = TimeSpan.FromSeconds(10),
        RequestTimeout = TimeSpan.FromSeconds(120),
    });
```

Service account JSON shortcut:

```csharp
using OfficeIMO.GoogleWorkspace;

var sessionOptions = new GoogleWorkspaceSessionOptions {
    SubjectUser = "analyst@example.com",
    UseDomainWideDelegation = true,
    DefaultFolderId = "reports-folder-id",
};

var credentialSource = GoogleServiceAccountCredentialSource.FromFile(
    "service-account.json",
    sessionOptions);

var session = new GoogleWorkspaceSession(credentialSource, sessionOptions);
```

Handling failed exports:

```csharp
try {
    var result = await document.ExportToGoogleDocsAsync(session, options);
    Console.WriteLine(result.WebViewLink);
} catch (GoogleWorkspaceExportException exception) {
    Console.WriteLine(exception.FailureKind);

    foreach (var notice in exception.Report.Notices) {
        Console.WriteLine($"{notice.Severity}: {notice.Feature} - {notice.Message}");
    }
} catch (GoogleWorkspaceExportCanceledException exception) {
    Console.WriteLine(exception.FailureKind);
}
```

Turning diagnostics into log-ready entries:

```csharp
foreach (var entry in exception.ToDiagnosticEntries()) {
    Console.WriteLine($"{entry.Severity} {entry.Feature} {entry.FailureKind}: {entry.Message}");
}
```

Streaming diagnostics during export:

```csharp
var session = new GoogleWorkspaceSession(
    credentialSource,
    new GoogleWorkspaceSessionOptions {
        DiagnosticSink = entry =>
            Console.WriteLine($"{entry.Severity} {entry.Feature} [{entry.FailureKind}]: {entry.Message}")
    });
```

## Session options

- `DefaultFolderId`: used when exporter save options omit `Location.FolderId`.
- `DefaultDriveId`: used when exporter save options omit `Location.DriveId`.
- `MaxRetryCount`: retry budget for transient Google API failures.
- `RetryBaseDelay`: starting point for exponential backoff when no `Retry-After` header is present.
- `RetryMaxDelay`: cap for retry delays.
- `RequestTimeout`: shared `HttpClient` timeout for Google API requests.
- `DiagnosticSink`: optional callback for live retry, authentication, Drive-placement, and API failure diagnostics during export.
- `SubjectUser` and `UseDomainWideDelegation`: available for apps that implement delegated credential flows in their own `IGoogleWorkspaceCredentialSource`.
- `SubjectUser` and `UseDomainWideDelegation`: also consumed by `GoogleServiceAccountCredentialSource` for domain-wide delegation.

## Notes

- `GoogleServiceAccountCredentialSource` currently relies on native PEM import support and is intended for modern runtimes such as `net8.0` and `net10.0`. On older targets, acquire tokens externally and use `StaticAccessTokenCredentialSource` or `DelegateGoogleWorkspaceCredentialSource`.
- The package does not ship an interactive browser OAuth flow. Applications are expected to provide or integrate one through `IGoogleWorkspaceCredentialSource`, although `StaticAccessTokenCredentialSource`, `DelegateGoogleWorkspaceCredentialSource`, and `GoogleServiceAccountCredentialSource` cover common server-side and pre-issued-token scenarios.
- Exporters add `ApiRetries` notices to `TranslationReport` when an operation succeeds after transient Google API failures.
- Exporters throw `GoogleWorkspaceExportException` when token acquisition, Google API execution, or request timeout failures occur, so callers can inspect `FailureKind` and the captured `TranslationReport`.
- Exporters throw `GoogleWorkspaceExportCanceledException` when the caller cancels the export, preserving both cancellation semantics and the captured `TranslationReport`.
- Google API failures now prefer parsed Google JSON error details such as status and reason codes when those are available.
- Replacement flows can target an existing file by setting `GoogleDriveFileLocation.ExistingFileId`.

For concrete exporters, see [OfficeIMO.Word.GoogleDocs](../OfficeIMO.Word.GoogleDocs/README.md) and [OfficeIMO.Excel.GoogleSheets](../OfficeIMO.Excel.GoogleSheets/README.md).
