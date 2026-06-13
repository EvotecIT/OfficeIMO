# OfficeIMO.GoogleWorkspace - shared Google Workspace primitives

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.GoogleWorkspace)](https://www.nuget.org/packages/OfficeIMO.GoogleWorkspace)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.GoogleWorkspace?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.GoogleWorkspace)

`OfficeIMO.GoogleWorkspace` contains shared credential, session, retry, Drive-location, and translation-report primitives for OfficeIMO Google Docs and Google Sheets exporters.

## Install

```powershell
dotnet add package OfficeIMO.GoogleWorkspace
```

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

## What it provides

- `IGoogleWorkspaceCredentialSource` for application-owned OAuth or service-account token acquisition.
- `StaticAccessTokenCredentialSource`, `DelegateGoogleWorkspaceCredentialSource`, and `GoogleServiceAccountCredentialSource`.
- `GoogleWorkspaceSession` and `GoogleWorkspaceSessionOptions`.
- `GoogleDriveFileLocation` for folder, shared-drive, and existing-file targeting.
- `TranslationReport`, `TranslationNotice`, export exceptions, cancellation exceptions, and log-ready diagnostic entries.
- Retry and API failure diagnostics through `GoogleWorkspaceSessionOptions.DiagnosticSink`.

## Service account shortcut

```csharp
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

## Boundaries

- This package owns shared Google Workspace plumbing.
- Word to Google Docs export belongs in `OfficeIMO.Word.GoogleDocs`.
- Excel to Google Sheets export belongs in `OfficeIMO.Excel.GoogleSheets`.
- Applications still own interactive browser OAuth flows unless they supply tokens through the shared credential interface.

## Related packages

- [OfficeIMO.Word.GoogleDocs](../OfficeIMO.Word.GoogleDocs/README.md)
- [OfficeIMO.Excel.GoogleSheets](../OfficeIMO.Excel.GoogleSheets/README.md)

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)
