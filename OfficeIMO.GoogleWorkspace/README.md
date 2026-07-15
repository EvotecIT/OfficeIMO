# OfficeIMO.GoogleWorkspace - shared Google Workspace primitives

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.GoogleWorkspace)](https://www.nuget.org/packages/OfficeIMO.GoogleWorkspace)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.GoogleWorkspace?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.GoogleWorkspace)

`OfficeIMO.GoogleWorkspace` contains the dependency-light credential, session, transport, retry, scope, diagnostics, Drive-location, and translation-report contracts shared by the OfficeIMO Google Docs, Sheets, Slides, Drive, and synchronization packages.

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
- `TranslationReport`, `TranslationNotice`, preflight, conflict/export exceptions, cancellation exceptions, and log-ready diagnostic entries.
- Safety-aware retries, normalized Google API failures, request timeouts, and diagnostic correlation through `GoogleWorkspaceSessionOptions.DiagnosticSink`.
- Minimum-scope catalogs for Docs, Sheets, Slides, and Drive operations.

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
- Word/Google Docs translation belongs in `OfficeIMO.Word.GoogleDocs`.
- Excel/Google Sheets translation belongs in `OfficeIMO.Excel.GoogleSheets`.
- PowerPoint/Google Slides translation belongs in `OfficeIMO.PowerPoint.GoogleSlides`.
- Drive resources belong in `OfficeIMO.GoogleWorkspace.Drive`; change-feed consumption and plan/apply belong in `OfficeIMO.GoogleWorkspace.Sync`.
- Applications own consent and credential policy. `OfficeIMO.GoogleWorkspace.Auth.GoogleApis` is an optional adapter when an application already uses the Google client SDK.

## Related packages

- [OfficeIMO.GoogleWorkspace.Drive](../OfficeIMO.GoogleWorkspace.Drive/README.md)
- [OfficeIMO.GoogleWorkspace.Auth.GoogleApis](../OfficeIMO.GoogleWorkspace.Auth.GoogleApis/README.md)
- [OfficeIMO.GoogleWorkspace.Sync](../OfficeIMO.GoogleWorkspace.Sync/README.md)
- [OfficeIMO.Word.GoogleDocs](../OfficeIMO.Word.GoogleDocs/README.md)
- [OfficeIMO.Excel.GoogleSheets](../OfficeIMO.Excel.GoogleSheets/README.md)
- [OfficeIMO.PowerPoint.GoogleSlides](../OfficeIMO.PowerPoint.GoogleSlides/README.md)

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`, plus `net472` on Windows.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** `System.Text.Json` plus platform HTTP and cryptography; no Google client SDK.
- **OfficeIMO:** Credential abstractions, sessions, retry, Drive placement, failures, and translation reports are first-party.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
