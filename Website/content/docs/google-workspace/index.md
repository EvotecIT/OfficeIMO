---
title: Google Workspace
description: "Choose and configure the OfficeIMO libraries for Google Docs, Sheets, Slides, Drive, and synchronization. Includes examples and package links."
order: 10
---

OfficeIMO exposes one shared Google Workspace foundation and format-specific translators. The format libraries remain thin over `OfficeIMO.Word`, `OfficeIMO.Excel`, and `OfficeIMO.PowerPoint`; they do not introduce a universal intermediate document model.

## Install the parts you use

```powershell
dotnet add package OfficeIMO.GoogleWorkspace
dotnet add package OfficeIMO.GoogleWorkspace.Drive
dotnet add package OfficeIMO.Word.GoogleDocs
dotnet add package OfficeIMO.Excel.GoogleSheets
dotnet add package OfficeIMO.PowerPoint.GoogleSlides
```

Add `OfficeIMO.GoogleWorkspace.Auth.GoogleApis` for Google client-library credentials, or `OfficeIMO.GoogleWorkspace.Sync` for change feeds and plan/apply orchestration.

## Create a session

```csharp
using OfficeIMO.GoogleWorkspace;

var session = new GoogleWorkspaceSession(
    new StaticAccessTokenCredentialSource("<access-token>"),
    new GoogleWorkspaceSessionOptions {
        ApplicationName = "Reporting service",
        DefaultFolderId = "folder-id",
        DefaultDriveId = "shared-drive-id",
        MaxRetryCount = 5,
        DiagnosticSink = entry => Console.WriteLine($"{entry.Severity}: {entry.Message}")
    });
```

A raw static token is useful for read-only samples and short-lived jobs. A caller-entered account name or requested scope is not proof of the token's identity or grants, so it cannot authorize mutations. Mutation-capable applications normally use the built-in service-account source or return `GoogleWorkspaceAccessToken.FromVerifiedCredential` from a delegate or optional Google APIs adapter after checking provider-issued token evidence.

## Common contract

Every translator can build a local plan before it mutates Google. `TranslationReport` entries contain a stable code, source path, severity, selected action, count, and optional target identifier. `GoogleWorkspaceFidelityPolicy` decides whether errors or warnings stop the operation and can accept specific diagnostic codes.

For existing files, import or read first, retain `RevisionId` or `DriveVersion`, and pass it back in the format replacement options. Drive change cursors help applications discover work; they do not replace format-specific diff plans or application-owned persistence.

Continue with [Google Docs](/docs/google-workspace/docs/), [Google Sheets](/docs/google-workspace/sheets/), [Google Slides](/docs/google-workspace/slides/), [Drive](/docs/google-workspace/drive/), [authentication](/docs/google-workspace/authentication/), [shared drives](/docs/google-workspace/shared-drives/), [synchronization](/docs/google-workspace/sync/), or [live tests](/docs/google-workspace/live-tests/).
