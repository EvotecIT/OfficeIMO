---
title: "Manage Google Drive Files"
description: "List, upload, download, copy, and share Google Drive files with the typed GoogleDriveClient from OfficeIMO.GoogleWorkspace.Drive."
order: 50
packages:
  - OfficeIMO.GoogleWorkspace.Drive
  - OfficeIMO.GoogleWorkspace
limits: "Multipart uploads are limited to 5 MiB; use UploadResumableAsync for larger files."
search_terms:
  - google drive api
  - upload to google drive
  - download from google drive
  - list drive files
verified: 2026-11-14
---

`GoogleDriveClient` is the Drive layer the format translators use, and applications can call it directly. It covers metadata, folders, shared-drive checks, list, copy, move, delete, permissions, comments and replies, revisions, change tokens, export formats, downloads, multipart and resumable uploads, progress, and cancellation.

`GoogleDriveClient.GetRequiredScopes(operation)` returns the default scope family for an operation, and `GoogleDriveClientOptions` can override read and write scopes. A `drive.file` token only sees files the application created or opened; ask for a broader Drive scope only when the product needs it.

## List files in a folder

### C#

```csharp
using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Drive;

using var credentials = GoogleServiceAccountCredentialSource.FromFile("service-account.json");
var session = new GoogleWorkspaceSession(credentials);

using var drive = new GoogleDriveClient(session);
GoogleDriveFile folder = await drive.ResolveFolderAsync("folder-id");

GoogleDriveFileList files = await drive.ListFilesAsync(new GoogleDriveListOptions {
    DriveId = folder.DriveId,
    Query = $"'{folder.Id}' in parents and trashed = false"
});

foreach (GoogleDriveFile file in files.Files) {
    Console.WriteLine($"{file.Name} ({file.MimeType})");
}
```

## Upload and download a file

### C#

```csharp
using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Drive;

using var credentials = GoogleServiceAccountCredentialSource.FromFile("service-account.json");
var session = new GoogleWorkspaceSession(credentials);
using var drive = new GoogleDriveClient(session);

GoogleDriveFile uploaded = await drive.UploadMultipartAsync(
    File.ReadAllBytes("report.pdf"),
    new GoogleDriveUploadOptions {
        Name = "report.pdf",
        ContentType = "application/pdf",
        ParentId = "folder-id"
    });

byte[] content = await drive.DownloadAsync(uploaded.Id!);
File.WriteAllBytes("report-copy.pdf", content);
```
