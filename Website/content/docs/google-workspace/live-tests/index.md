---
title: Live-test setup
description: Run disposable Google Docs, Sheets, Slides, Drive, shared-drive, and change-feed tests safely.
order: 75
---

Live tests are opt-in and create disposable resources in one dedicated folder. Set:

```powershell
$env:OFFICEIMO_RUN_GOOGLE_WORKSPACE_LIVE = "1"
$env:GOOGLE_WORKSPACE_ACCESS_TOKEN = "<short-lived-token>"
$env:GOOGLE_WORKSPACE_FOLDER_ID = "<disposable-test-folder-id>"
$env:GOOGLE_WORKSPACE_DRIVE_ID = "<optional-shared-drive-id>"
```

Run the format partitions:

```powershell
dotnet test OfficeIMO.Word.Tests -c Release -f net10.0 --filter Category=GoogleWorkspaceLive
dotnet test OfficeIMO.Excel.Tests -c Release -f net10.0 --filter Category=GoogleWorkspaceLive
dotnet test OfficeIMO.PowerPoint.Tests -c Release -f net10.0 --filter Category=GoogleWorkspaceLive
dotnet test OfficeIMO.GoogleWorkspace.Tests -c Release -f net10.0 --filter Category=GoogleWorkspaceLive
```

Each test creates, reads/imports, exports where applicable, and deletes its file in a `finally` path. The change-feed test also verifies that the disposable item becomes visible. Use a non-production folder and principal with only the scopes needed by the lane. If a cleanup failure is reported, remove that named disposable item before rerunning.
