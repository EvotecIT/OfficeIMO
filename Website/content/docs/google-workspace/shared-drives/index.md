---
title: Shared drives
description: Validate shared-drive and folder identity, target translators, and track per-drive changes.
order: 65
---

Set both the drive and folder when a workflow targets a shared drive. The folder is the actual placement boundary; a query flag alone does not prove it belongs to the intended drive.

```csharp
var session = new GoogleWorkspaceSession(credentials, new GoogleWorkspaceSessionOptions {
    DefaultDriveId = "shared-drive-id",
    DefaultFolderId = "shared-drive-folder-id"
});

using var drive = new GoogleDriveClient(session);
GoogleSharedDrive sharedDrive = await drive.GetSharedDriveAsync(session.Options.DefaultDriveId!);
GoogleDriveFile folder = await drive.ResolveFolderAsync(
    session.Options.DefaultFolderId!,
    session.Options.DefaultDriveId);
```

All translators resolve `GoogleDriveFileLocation` against session defaults. Override `DriveId` and `FolderId` per operation when needed. `OfficeIMO.GoogleWorkspace.Sync` stores a separate change cursor for the user feed and every configured shared drive, so one failed drive does not advance another drive checkpoint.
