---
title: PowerPoint and Google Slides
description: "Create, import, template, and safely replace Google Slides with OfficeIMO.PowerPoint. Includes examples and package links."
order: 40
---

Install `OfficeIMO.PowerPoint.GoogleSlides`.

```csharp
using OfficeIMO.GoogleWorkspace;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.GoogleSlides;

using PowerPointPresentation deck = PowerPointPresentation.Create();
PowerPointSlide slide = deck.AddSlide();
slide.AddTextBoxPoints("Quarterly review", 30, 40, 500, 80);
slide.Notes.Text = "Discuss the year-over-year change.";

var options = new GoogleSlidesSaveOptions { Title = "Quarterly review" };
GoogleSlidesTranslationPlan plan = deck.BuildGoogleSlidesPlan(options);
GooglePresentationReference created = await deck.ExportToGoogleSlidesAsync(session, options);
```

Supported text boxes, core run styles, external links, unmerged tables, uncropped PNG/JPEG/GIF pictures, common shapes, solid backgrounds, and speaker notes map to editable Slides objects. Export does not change the remote page size; it scales and centers source coordinates when the sizes differ. Complex slides containing merged tables, unsupported picture formats or geometry, charts, SmartArt, media, OLE, or connectors render to a coherent slide PNG by default. `PreferNativeAndReport` instead skips unsupported elements and reports each loss.

Native import returns the Slides revision needed for guarded replacement. `DriveExport` converts to PPTX for broader fidelity. Existing replacement requires `ExpectedRevisionId`; `OverwriteLatest` is an explicit last-writer-wins mode. `TemplatePresentationId` copies a Drive presentation before applying the batch.

Slides fetches inserted images from public URLs. The exporter creates temporary publicly readable Drive files and attempts deletion in success and failure paths. Deletion failures are recorded as `DRIVE.TEMPORARY_CONTENT.CLEANUP_FAILED` in the returned reference's report after a successful export; an undeleted file remains public and needs follow-up.
