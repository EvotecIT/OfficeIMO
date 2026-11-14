---
title: Export PowerPoint Slides to PNG, SVG, and JPEG
description: Render PowerPoint slides to PNG, JPEG, TIFF, SVG, or WebP images for thumbnails, previews, and visual checks with OfficeIMO.PowerPoint and PSWriteOffice.
order: 37
packages:
  - OfficeIMO.PowerPoint
cmdlets:
  - Export-OfficeDocumentImage
  - New-OfficePowerPointImageOptions
limits: Images show the static state of each slide; animations and media playback are not rendered.
search_terms:
  - pptx to png
  - slide thumbnails
  - powerpoint to image
  - slide preview
  - pptx to svg
meta.seo_title: Export PowerPoint slides to images in .NET
verified: 2026-11-14
---

Slide images are useful for thumbnails in a document portal, previews in a review tool, or a visual baseline to compare after a template change. `SaveAsImages` renders each slide with OfficeIMO's own renderer, so PowerPoint is not needed. Large rasters are bounded before pixels are allocated: `MaximumOutputWidth`, `MaximumOutputHeight`, and the raster pixel limits on the options stop a huge slide from exhausting memory.

## Export every slide

`SaveAsImages` writes one file per slide into a folder and returns a report with the files written.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;

using PowerPointDocument presentation = PowerPointDocument.Load("quarterly-review.pptx");
OfficeConversionReport report = presentation.SaveAsImages("previews",
    new PowerPointImageExportOptions { Format = OfficeImageExportFormat.Png, Scale = 2 });
Console.WriteLine($"{report.ItemsWritten} images written");
```

### PowerShell

```powershell
Export-OfficeDocumentImage -Path .\quarterly-review.pptx -Format Png -Scale 2 -OutputDirectory .\previews
```

## Export selected slides as SVG

`SlideNumbers` starts at 1. Hidden slides are skipped unless `IncludeHiddenSlides` is set.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;

using PowerPointDocument presentation = PowerPointDocument.Load("quarterly-review.pptx");
presentation.SaveAsImages("slides", new PowerPointImageExportOptions {
    Format = OfficeImageExportFormat.Svg,
    SlideNumbers = new[] { 1, 2, 3 }
});
```

### PowerShell

```powershell
Export-OfficeDocumentImage -Path .\quarterly-review.pptx -Page 1-3 -Format Svg -OutputDirectory .\slides
```

## Make small thumbnails

Cap the output size instead of picking a scale when slides come in different sizes.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;

using PowerPointDocument presentation = PowerPointDocument.Load("quarterly-review.pptx");
presentation.SaveAsImages("thumbs", new PowerPointImageExportOptions {
    Format = OfficeImageExportFormat.Jpeg,
    MaximumOutputWidth = 320,
    MaximumOutputHeight = 180
});
```

### PowerShell

```powershell
$options = New-OfficePowerPointImageOptions -MaximumOutputWidth 320 -MaximumOutputHeight 180
Export-OfficeDocumentImage -Path .\quarterly-review.pptx -Format Jpeg -Options $options -OutputDirectory .\thumbs
```

## Check slides before publishing

Rendering shows what a slide looks like; `InspectPreflight` reports what is wrong with it: clipped or overflowing text, shapes outside the slide, and broken picture links, each with a stable code.

### C#

```csharp
using OfficeIMO.PowerPoint;

using PowerPointDocument presentation = PowerPointDocument.Load("quarterly-review.pptx");
PowerPointDeckPreflightReport report = presentation.InspectPreflight();
File.WriteAllText("quarterly-review.preflight.json", report.ToJson());
```

### PowerShell

```powershell
Test-OfficePowerPoint -Path .\quarterly-review.pptx
```
