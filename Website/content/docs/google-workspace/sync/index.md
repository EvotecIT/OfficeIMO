---
title: Change tracking and synchronization
description: Consume Drive changes and execute explicit dry-run, approval, conflict, cancellation, and partial-failure plans.
order: 70
---

Install `OfficeIMO.GoogleWorkspace.Sync`. Initialize cursors once, persist the minimal checkpoint, then consume complete pages on later runs.

```csharp
using var tracker = new GoogleWorkspaceChangeTracker(session);
GoogleWorkspaceSyncCheckpoint checkpoint = await tracker.InitializeAsync(new[] { "shared-drive-id" });

var readOptions = new GoogleWorkspaceChangeReadOptions();
readOptions.SharedDriveIds.Add("shared-drive-id");
GoogleWorkspaceChangeReadResult read = await tracker.ReadAsync(checkpoint, readOptions);
```

Advance to `read.NextCheckpoint` only after the application processes the returned changes. Each source cursor advances only if that source reaches its new start token. The checkpoint stores cursors and optional stable identity/revision evidence, never document content.

Map a Docs, Sheets, or Slides diff into `GoogleWorkspaceSyncItem` values. `GoogleWorkspaceSyncExecutor` is dry-run by default, never calls the mutation callback for conflicts, and requires stable-ID approval for lossy items. Set `DryRun = false` to apply. The result reports every item as planned, applied, skipped, conflict, approval-required, failed, or canceled and can return already-applied progress on cancellation.
