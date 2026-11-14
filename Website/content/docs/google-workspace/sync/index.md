---
title: "Track Google Drive Changes"
description: "Read Google Drive change feeds for the user and shared drives with checkpoints, and apply sync plans with dry runs, approvals, and conflict reporting."
order: 70
packages:
  - OfficeIMO.GoogleWorkspace.Sync
  - OfficeIMO.GoogleWorkspace
search_terms:
  - google drive changes
  - drive change feed
  - sync google drive
  - drive start page token
---

`GoogleWorkspaceChangeTracker` in `OfficeIMO.GoogleWorkspace.Sync` reads the user change feed and any shared-drive feeds you name. Initialize the cursors once, store the checkpoint, and read complete pages on later runs. The checkpoint holds cursors and optional identity and revision data, never document content. Your application owns checkpoint storage, schedules, approval UI, and conflict decisions.

## Read changes since the last run

Move to `read.NextCheckpoint` only after you have processed the returned changes. Each source cursor moves only when that source reached its new start token, so one failing shared drive does not advance another.

### C#

<!-- docs-check: skip -->
```csharp
using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Sync;

using var credentials = GoogleServiceAccountCredentialSource.FromFile("service-account.json");
var session = new GoogleWorkspaceSession(credentials);

using var tracker = new GoogleWorkspaceChangeTracker(session);
GoogleWorkspaceSyncCheckpoint checkpoint = await tracker.InitializeAsync(new[] { "shared-drive-id" });

var readOptions = new GoogleWorkspaceChangeReadOptions();
readOptions.SharedDriveIds.Add("shared-drive-id");

GoogleWorkspaceChangeReadResult read = await tracker.ReadAsync(checkpoint, readOptions);
foreach (GoogleWorkspaceTrackedChange tracked in read.Changes) {
    Console.WriteLine($"{tracked.Source.Kind} {tracked.Change.FileId} removed={tracked.Change.Removed}");
}

checkpoint = read.NextCheckpoint;
```

## Apply a sync plan

Turn a Docs, Sheets, or Slides diff into `GoogleWorkspaceSyncItem` values and run them through `GoogleWorkspaceSyncExecutor`. It is a dry run by default (`DryRun = true` in `GoogleWorkspaceSyncApplyOptions`), never calls your mutation callback for a conflict, and requires approval by stable ID (`ApprovedLossyItemIds`) for lossy items. Set `DryRun = false` to apply. The result reports each item as planned, applied, skipped, conflict, approval-required, failed, or canceled, and can return the progress already made when the run is canceled.
