---
title: "Export Office Documents to Google Workspace"
description: "Export Word, Excel, and PowerPoint documents to Google Docs, Sheets, and Slides, import them back, and work with Drive through one shared session."
order: 10
packages:
  - OfficeIMO.GoogleWorkspace
  - OfficeIMO.Word.GoogleDocs
  - OfficeIMO.Excel.GoogleSheets
  - OfficeIMO.PowerPoint.GoogleSlides
  - OfficeIMO.GoogleWorkspace.Drive
cmdlets:
  - New-OfficeGoogleWorkspaceSession
  - Export-OfficeWordGoogleDocument
search_terms:
  - google docs
  - google sheets
  - google slides
  - google drive api
  - docx to google docs
verified: 2026-11-14
---

`OfficeIMO.GoogleWorkspace` holds the session, credentials, transport, retries, diagnostics, and fidelity policy. Format packages translate one OfficeIMO document type each: `OfficeIMO.Word.GoogleDocs`, `OfficeIMO.Excel.GoogleSheets`, and `OfficeIMO.PowerPoint.GoogleSlides`. `OfficeIMO.GoogleWorkspace.Drive` covers files and folders, `OfficeIMO.GoogleWorkspace.Auth.GoogleApis` adapts Google client-library credentials, and `OfficeIMO.GoogleWorkspace.Sync` reads change feeds.

| You want to | Page |
|---|---|
| Create a session from a service account, token, or Google APIs credential | [Authenticate to Google Workspace](/docs/google-workspace/authentication/) |
| Export Word to Google Docs or import a Google Doc | [Export and Import Google Docs](/docs/google-workspace/docs/) |
| Export Excel to Google Sheets or import a spreadsheet | [Export and Import Google Sheets](/docs/google-workspace/sheets/) |
| Export PowerPoint to Google Slides | [Export and Import Google Slides](/docs/google-workspace/slides/) |
| List, upload, download, or share Drive files | [Manage Google Drive Files](/docs/google-workspace/drive/) |
| Target a shared drive | [Work with Google Shared Drives](/docs/google-workspace/shared-drives/) |
| Track Drive changes and apply a sync plan | [Track Google Drive Changes](/docs/google-workspace/sync/) |
| Check what each translator supports | [Google Workspace support matrix](/docs/google-workspace/support/) |

## Create a session and export a document

Every translator can build a local plan before it touches Google. `TranslationReport` entries carry a stable code, source path, severity, chosen action, and count, and `GoogleWorkspaceFidelityPolicy` decides whether errors or warnings stop the operation.

### C#

```csharp
using OfficeIMO.GoogleWorkspace;
using OfficeIMO.Word;
using OfficeIMO.Word.GoogleDocs;

using var credentials = GoogleServiceAccountCredentialSource.FromFile("service-account.json");
var session = new GoogleWorkspaceSession(credentials, new GoogleWorkspaceSessionOptions {
    ApplicationName = "Reporting service",
    DefaultFolderId = "folder-id",
    MaxRetryCount = 5,
    DiagnosticSink = entry => Console.WriteLine($"{entry.Severity}: {entry.Message}")
});

using WordDocument document = WordDocument.Load("report.docx");
GoogleDocumentReference created = await document.ExportToGoogleDocsAsync(session,
    new GoogleDocsSaveOptions { Title = "Quarterly report" });

Console.WriteLine(created.FileId);
```

### PowerShell

```powershell
$session = New-OfficeGoogleWorkspaceSession -KeyPath .\Secrets\reporting.json -Scope Documents, DriveFile
try {
    Export-OfficeWordGoogleDocument -Path .\report.docx -Session $session
} finally {
    $session.CredentialSource.Dispose()
}
```

## Update an existing file safely

Import or read the target first, keep its `RevisionId` (Docs, Slides) or `DriveVersion` (Sheets), and pass it back in the replace options. The replacement fails if the file changed in between. Overwrite modes exist, but only as an explicit choice. Drive change feeds tell you what to look at; they do not replace the format-specific replacement checks.
