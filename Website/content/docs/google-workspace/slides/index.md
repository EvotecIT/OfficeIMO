---
title: PowerPoint and Google Slides
description: Create, import, template, and safely replace Google Slides with OfficeIMO.PowerPoint.
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

Text boxes, core run styles, hyperlinks, tables, pictures, basic shapes, slide size, solid backgrounds, and speaker notes remain editable. Complex slides containing charts, SmartArt, media, OLE, connectors, or unsupported objects render to a coherent slide PNG by default. `PreferNativeAndReport` instead skips unsupported elements and reports each loss.

Native import returns the Slides revision needed for guarded replacement. `DriveExport` converts to PPTX for broader fidelity. Existing replacement requires `ExpectedRevisionId`; `OverwriteLatest` is an explicit last-writer-wins mode. `TemplatePresentationId` copies a Drive presentation before applying the batch.

Slides fetches inserted images from public URLs. The exporter creates short-lived Drive leases and removes their permissions and files in success and failure paths.
