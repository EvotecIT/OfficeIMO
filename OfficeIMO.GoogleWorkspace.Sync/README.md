# OfficeIMO.GoogleWorkspace.Sync

`OfficeIMO.GoogleWorkspace.Sync` provides the cross-format synchronization mechanics shared by the Google Docs, Sheets, and Slides translators. It consumes Drive change feeds and executes explicit mutation plans; it does not store document content or choose merge policy for an application.

## Install

```powershell
dotnet add package OfficeIMO.GoogleWorkspace.Sync
```

## Track user and shared-drive changes

```csharp
using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Sync;

var session = new GoogleWorkspaceSession(credentialSource);
using var tracker = new GoogleWorkspaceChangeTracker(session);

GoogleWorkspaceSyncCheckpoint checkpoint = await tracker.InitializeAsync(
    new[] { "shared-drive-id" });

// Persist checkpoint. On the next run:
var options = new GoogleWorkspaceChangeReadOptions();
options.SharedDriveIds.Add("shared-drive-id");
GoogleWorkspaceChangeReadResult changes = await tracker.ReadAsync(checkpoint, options);

foreach (GoogleWorkspaceTrackedChange item in changes.Changes) {
    Console.WriteLine($"{item.Source.Key}: {item.Change.FileId}");
}

// Persist only after processing the returned changes successfully.
checkpoint = changes.NextCheckpoint;
```

Each source is paged to completion. Its token advances only when that source succeeds; a failed shared-drive feed keeps its old token while completed sources can advance. A newly added shared drive is initialized at its current cursor and does not invent historical changes.

The checkpoint intentionally contains only the user cursor, per-shared-drive cursors, and optional stable source/file identity plus revision evidence. Applications remain responsible for durable storage, content, schedules, and conflict policy.

## Review, dry-run, and apply

Map a Docs, Sheets, or Slides diff plan into stable shared items, then dry-run it before mutation:

```csharp
var policy = new GoogleWorkspaceOperationPolicy(
    account: "automation@example.com",
    scopes: new[] { GoogleWorkspaceScopeCatalog.DriveFile },
    target: "drive://sync-root",
    expectedRevision: "version-42",
    maxRetryCount: 3,
    maxRetryElapsedTime: TimeSpan.FromMinutes(2),
    rateLimitPolicy: GoogleWorkspaceRateLimitPolicy.HonorRetryAfter,
    dataLossDecision: GoogleWorkspaceDataLossDecision.AcceptSpecifiedLoss,
    acceptedLoss: "chart-1 may be rasterized");

GoogleWorkspaceSyncPlan plan = GoogleWorkspaceSyncPlan.Create(new[] {
    new GoogleWorkspaceSyncItem("summary!A1", GoogleWorkspaceSyncItemKind.SourceChange,
        "sheet/Summary/cell/1:1", "The local value changed.",
        targetResource: "drive://spreadsheet-id/summary!A1", expectedRevision: "version-42",
        googleFileId: "spreadsheet-id"),
    new GoogleWorkspaceSyncItem("chart-1", GoogleWorkspaceSyncItemKind.LossyAction,
        "sheet/Summary/chart/1", "This chart requires a raster fallback.",
        targetResource: "drive://spreadsheet-id/chart-1", expectedRevision: "version-42",
        googleFileId: "spreadsheet-id")
}, policy);

GoogleWorkspaceSyncApplyResult preview = await GoogleWorkspaceSyncExecutor.ApplyAsync(
    plan,
    (item, cancellationToken) => ApplyItemAsync(item, cancellationToken));

var applyOptions = new GoogleWorkspaceSyncApplyOptions { DryRun = false };
applyOptions.ApprovedLossyItemIds.Add("chart-1");
GoogleWorkspaceSyncApplyResult applied = await GoogleWorkspaceSyncExecutor.ApplyAsync(
    plan,
    (item, cancellationToken) => ApplyItemAsync(item, cancellationToken),
    applyOptions,
    cancellationToken);
```

Conflicts are never sent to the operation. Lossy actions require both an `AcceptSpecifiedLoss` plan policy and approval by stable item ID; either guard can refuse the mutation. Every item returns `Planned`, `Applied`, `Conflict`, `ApprovalRequired`, `Failed`, `Skipped`, or `Canceled` with a `GoogleWorkspaceSyncDecisionReceipt`; actual HTTP mutation receipts remain owned by the configured session receipt sink. Cancellation can return the already-applied partial result.

## Boundaries

- `OfficeIMO.GoogleWorkspace.Drive` owns Drive change-token HTTP operations.
- Docs, Sheets, and Slides diff planners own format-specific comparison and fidelity classification.
- This package owns cursor consumption and common plan/apply outcomes.
- The application owns checkpoint persistence, selection/approval UI, schedules, and the actual mutation callback.

Targets: `netstandard2.0`, `net8.0`, `net10.0`, plus `net472` on Windows. License: MIT.
Every `GoogleWorkspaceSyncPlan` requires a `GoogleWorkspaceOperationPolicy`, and every item requires an explicit target resource and expected revision decision. Dry-run remains the default. Apply results include a decision receipt for planned, applied, blocked, failed, and canceled items, so partial progress and policy decisions remain inspectable without misrepresenting those decisions as network receipts.
