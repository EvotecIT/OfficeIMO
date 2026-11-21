---
title: "Work with Google Shared Drives"
description: "Point exports at a shared drive and folder, check that the folder belongs to that drive, and track changes per shared drive."
order: 65
packages:
  - OfficeIMO.GoogleWorkspace
  - OfficeIMO.GoogleWorkspace.Drive
search_terms:
  - shared drive
  - team drive
  - google drive folder id
verified: 2026-11-14
---

When a workflow writes to a shared drive, set both the drive and the folder. The folder is the real placement boundary; a query flag alone does not prove the folder belongs to the intended drive. All translators resolve `GoogleDriveFileLocation` against the session defaults, and you can override `DriveId` and `FolderId` per operation.

## Check a shared drive and folder

### C#

```csharp
using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Drive;

using var credentials = GoogleServiceAccountCredentialSource.FromFile("service-account.json");
var session = new GoogleWorkspaceSession(credentials, new GoogleWorkspaceSessionOptions {
    DefaultDriveId = "shared-drive-id",
    DefaultFolderId = "shared-drive-folder-id"
});

using var drive = new GoogleDriveClient(session);
GoogleSharedDrive sharedDrive = await drive.GetSharedDriveAsync(session.Options.DefaultDriveId!);

// Throws when the folder is not a folder or is not in the expected shared drive.
GoogleDriveFile folder = await drive.ResolveFolderAsync(
    session.Options.DefaultFolderId!,
    session.Options.DefaultDriveId);

Console.WriteLine($"{sharedDrive.Name}: {folder.Name}");
```

## Track changes per shared drive

`OfficeIMO.GoogleWorkspace.Sync` keeps a separate change cursor for the user feed and for each configured shared drive, so a failure on one drive does not move another drive's checkpoint. See [Track Google Drive Changes](/docs/google-workspace/sync/).
