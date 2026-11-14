---
title: "Export and Import Google Slides"
description: "Plan a PowerPoint-to-Google-Slides export, create a presentation from a PowerPointDocument, and import or replace Google Slides safely."
order: 40
packages:
  - OfficeIMO.PowerPoint.GoogleSlides
  - OfficeIMO.GoogleWorkspace
limits: "Inserted images are fetched by Google from temporary public Drive files; if cleanup fails, the file stays public and the report says so."
search_terms:
  - powerpoint to google slides
  - pptx to google slides
  - import google slides
  - google slides api
verified: 2026-11-14
---

`OfficeIMO.PowerPoint.GoogleSlides` translates a `PowerPointDocument` into Google Slides API requests. Text boxes, core run styles, external links, unmerged tables, uncropped PNG/JPEG/GIF pictures, common shapes, solid backgrounds, and speaker notes become editable Slides objects. Slides with merged tables, unsupported pictures or geometry, charts, SmartArt, media, OLE, or connectors are rendered to one PNG per slide by default; `PreferNativeAndReport` skips those elements instead and reports each loss.

## Create a presentation in Google Slides

Export does not change the remote page size; when sizes differ, it scales and centers the source coordinates.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.GoogleWorkspace;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.GoogleSlides;

using var credentials = GoogleServiceAccountCredentialSource.FromFile("service-account.json");
var session = new GoogleWorkspaceSession(credentials, new GoogleWorkspaceSessionOptions { DefaultFolderId = "folder-id" });

using PowerPointDocument deck = PowerPointDocument.Create();
PowerPointSlide slide = deck.AddSlide();
slide.AddTextBox("Quarterly review",
    OfficeLength.Points(30), OfficeLength.Points(40), OfficeLength.Points(500), OfficeLength.Points(80));
slide.Notes.Text = "Discuss the year-over-year change.";

var options = new GoogleSlidesSaveOptions { Title = "Quarterly review" };

GoogleSlidesTranslationPlan plan = deck.BuildGoogleSlidesPlan(options);
foreach (TranslationNotice notice in plan.Report.Notices) {
    Console.WriteLine($"{notice.Severity} {notice.Code}: {notice.Message}");
}

GooglePresentationReference created = await deck.ExportToGoogleSlidesAsync(session, options);
Console.WriteLine(created.WebViewLink);
```

`TemplatePresentationId` copies an existing Drive presentation first and applies the slides to the copy.

## Import Google Slides and replace a presentation safely

Native import returns the Slides revision needed for a guarded replacement. `DriveExport` mode converts to PPTX for broader fidelity. Replacing an existing presentation requires `ExpectedRevisionId`; `OverwriteLatest` is an explicit last-writer-wins mode.

### C#

```csharp
using OfficeIMO.GoogleWorkspace;
using OfficeIMO.PowerPoint.GoogleSlides;

using var credentials = GoogleServiceAccountCredentialSource.FromFile("service-account.json");
var session = new GoogleWorkspaceSession(credentials);

GoogleSlidesImportResult current = await session.ImportGoogleSlidesAsync(
    "presentation-id",
    new GoogleSlidesImportOptions { Mode = GoogleWorkspaceImportMode.Native });

using (current.Document) {
    var replace = new GoogleSlidesSaveOptions {
        Location = new GoogleDriveFileLocation { ExistingFileId = current.Source.FileId },
        Replace = new GoogleSlidesReplaceOptions { ExpectedRevisionId = current.Source.RevisionId }
    };
    await current.Document.ExportToGoogleSlidesAsync(session, replace);
}
```

Google Slides fetches inserted images from URLs, so the exporter creates temporary publicly readable Drive files and tries to delete them on success and failure. A failed deletion is reported as `DRIVE.TEMPORARY_CONTENT.CLEANUP_FAILED` in the returned report and needs follow-up.
