---
title: Google Drive
description: Use the shared typed Drive owner for files, conversion, collaboration, media, and change tokens.
order: 50
---

Install `OfficeIMO.GoogleWorkspace.Drive`. The format translators already use this package; applications can use the same owner directly.

```csharp
using OfficeIMO.GoogleWorkspace.Drive;

using var drive = new GoogleDriveClient(session);
GoogleDriveAboutFormats formats = await drive.GetFormatsAsync();
GoogleDriveFile folder = await drive.ResolveFolderAsync("folder-id", "expected-drive-id");
GoogleDriveFileList files = await drive.ListFilesAsync(new GoogleDriveListOptions {
    DriveId = folder.DriveId,
    Query = $"'{folder.Id}' in parents and trashed = false"
});
```

The client covers metadata, folders, shared-drive validation, list/copy/move/delete, permissions, comments/replies, revisions, change tokens, import/export format discovery, downloads, multipart/resumable uploads, progress, cancellation, and temporary content leases.

`GoogleDriveClient.GetRequiredScopes(operation)` exposes the minimum default scope family. Applications can override read/write scopes through `GoogleDriveClientOptions` when their consent model differs. A `drive.file` token only sees files created or explicitly opened by the application; use a broader Drive scope only when the product really needs it.
