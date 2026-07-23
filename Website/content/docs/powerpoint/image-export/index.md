---
title: "PowerPoint Image Export"
description: "Render PowerPoint slides to PNG, JPEG, TIFF, SVG, or WebP for previews, review, and delivery workflows."
order: 44
meta.seo_title: "Export PowerPoint slides to images in .NET"
---

Image export is useful when a workflow needs review thumbnails, web previews, evidence packs, or visual regression baselines in addition to an editable PPTX file.

## Export slide previews

Use the PowerPoint image-export API to select slides, DPI or scale, destination format, and output naming. The renderer uses OfficeIMO.Drawing and applies bounded raster allocation before creating large pixel buffers.

```csharp
using OfficeIMO.PowerPoint;

using PowerPointPresentation presentation =
    PowerPointPresentation.Load("Quarterly-Review.pptx");

presentation.ToImages()
    .ForSlideRange(1, presentation.Slides.Count)
    .AsPng()
    .Save("Quarterly-Review previews");
```

Pair preview generation with `Preflight()` when text fit, shape bounds, image relationships, and package quality matter:

```csharp
PowerPointDeckPreflightReport report = presentation.InspectPreflight();
File.WriteAllText("Quarterly-Review.preflight.json", report.ToJson());
```

## Practical uses

- Generate a thumbnail index for a document portal.
- Attach slide previews to pull-request or content-review workflows.
- Compare a committed visual baseline after changing a deck template.
- Build a PDF or image evidence pack while retaining the editable PPTX.
- Give downstream services a lightweight preview without installing PowerPoint.

The [showcase](/showcase/) includes committed PowerPoint example renders. See [Designer Decks](/docs/powerpoint/designer/) for the source workflow that creates those outputs.
