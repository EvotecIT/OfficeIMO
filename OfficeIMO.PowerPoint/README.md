# OfficeIMO.PowerPoint — .NET PowerPoint Utilities

OfficeIMO.PowerPoint focuses on creating and editing .pptx presentations with Open XML.

- Targets: netstandard2.0, net472, net8.0, net9.0
- License: MIT
- NuGet: `OfficeIMO.PowerPoint`
- Dependencies: DocumentFormat.OpenXml

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.PowerPoint)](https://www.nuget.org/packages/OfficeIMO.PowerPoint)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.PowerPoint?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.PowerPoint)

See `OfficeIMO.Examples` for runnable samples. This README hosts PowerPoint‑specific usage and notes.

## Install

```powershell
dotnet add package OfficeIMO.PowerPoint
```

## Quick sample

```csharp
using OfficeIMO.PowerPoint;

using var ppt = PowerPointPresentation.Create("demo.pptx");
var slide = ppt.AddSlide();
slide.AddTitle("Hello PowerPoint");
var box = slide.AddTextBox("Generated with OfficeIMO.PowerPoint");
box.SetPositionCm(2, 2);
box.SetSizeCm(6, 2);
ppt.Save();
```

## Common Tasks by Example

### Title + content
```csharp
var slide = ppt.AddSlide();
slide.AddTitle("Quarterly Review");
slide.AddTextBox("Agenda\n• Intro\n• KPIs\n• Next Steps",
    PowerPointUnits.Cm(1.5), PowerPointUnits.Cm(2.5), PowerPointUnits.Cm(7.5), PowerPointUnits.Cm(3.0));
```

### Bullets (API)
```csharp
var box = slide.AddTextBox("Agenda:");
box.AddBullets(new[] { "Intro", "KPIs", "Next Steps" });
```

### Numbered lists
```csharp
var box = slide.AddTextBox("Plan:");
box.AddNumberedList(new[] { "Discover", "Design", "Deliver" });
```

### Text styles + spacing
```csharp
var box = slide.AddTextBox("Highlights");
box.AddBullets(new[] { "Readable defaults", "Auto spacing", "Consistent styles" });
box.ApplyTextStyle(PowerPointTextStyle.Body.WithColor("1F4E79"));
box.ApplyAutoSpacing(lineSpacingMultiplier: 1.15, spaceAfterPoints: 2);
```

### Text box layout (margins + autofit)
```csharp
using A = DocumentFormat.OpenXml.Drawing;

var box = slide.AddTextBox("Inset text");
box.SetTextMarginsCm(0.3, 0.2, 0.3, 0.2);
box.TextAutoFit = PowerPointTextAutoFit.Normal;
box.TextVerticalAlignment = A.TextAnchoringTypeValues.Center;
```

### Images
```csharp
slide.AddPicture("logo.png",
    PowerPointUnits.Cm(23), PowerPointUnits.Cm(1.2), PowerPointUnits.Cm(5), PowerPointUnits.Cm(2));

using var logoStream = File.OpenRead("logo.png");
slide.AddPicture(logoStream, ImagePartType.Png,
    PowerPointUnits.Cm(23), PowerPointUnits.Cm(1.2), PowerPointUnits.Cm(5), PowerPointUnits.Cm(2));
```

### Background image
```csharp
slide.SetBackgroundImage("hero.png");
```

### Simple shapes
```csharp
slide.AddRectangle(PowerPointUnits.Cm(1), PowerPointUnits.Cm(1),
    PowerPointUnits.Cm(3), PowerPointUnits.Cm(1))
    .Fill("#E7F7FF")
    .Stroke("#007ACC");
```

### Align + distribute shapes
```csharp
slide.AlignShapes(slide.Shapes, PowerPointShapeAlignment.Left);
slide.DistributeShapes(slide.Shapes, PowerPointShapeDistribution.Horizontal);
```

### Slide properties
```csharp
ppt.BuiltinDocumentProperties.Title = "Contoso Review";
ppt.ApplicationProperties.Company = "Contoso";
```

### Slide visibility + duplication
```csharp
var duplicate = ppt.DuplicateSlide(0);
duplicate.Hidden = true;
```

### Import slide from another deck
```csharp
using var source = PowerPointPresentation.Open("source.pptx");
var imported = ppt.ImportSlide(source, sourceIndex: 0);
imported.AddTextBox("Imported content");
```

### Layouts and notes (fluent)
```csharp
using OfficeIMO.PowerPoint.Fluent;

ppt.AsFluent()
   .Slide(masterIndex: 0, layoutIndex: 0, s =>
   {
       s.Title("Fluent Slide");
       s.Bullets("One", "Two", "Three");
       s.Notes("Talking points for the presenter");
   });
```

### Tables (data binding)
```csharp
record SalesRow(string Product, int Q1, int Q2);

var rows = new[] {
    new SalesRow("Alpha", 12, 15),
    new SalesRow("Beta", 9, 11)
};

var columns = new[] {
    PowerPointTableColumn<SalesRow>.Create("Product", r => r.Product).WithWidthCm(4.0),
    PowerPointTableColumn<SalesRow>.Create("Q1", r => r.Q1),
    PowerPointTableColumn<SalesRow>.Create("Q2", r => r.Q2)
};

slide.AddTable(rows, columns, left: PowerPointUnits.Cm(1.5), top: PowerPointUnits.Cm(4),
    width: PowerPointUnits.Cm(20), height: PowerPointUnits.Cm(6));
```

### Tables (merged cells)
```csharp
var table = slide.AddTable(rows: 4, columns: 4, left: PowerPointUnits.Cm(1.5),
    top: PowerPointUnits.Cm(4), width: PowerPointUnits.Cm(20), height: PowerPointUnits.Cm(6));
table.GetCell(0, 0).Text = "Merged header";
table.MergeCells(0, 0, 0, 3);
```

### Placeholders (layout-driven)
```csharp
using DocumentFormat.OpenXml.Presentation;

var slide = ppt.AddSlide(masterIndex: 0, layoutIndex: 1);
var title = slide.GetPlaceholder(PlaceholderValues.Title);
title?.SetTextMarginsCm(0.2, 0.1, 0.2, 0.1);
if (title != null) title.Text = "Layout Placeholder";
```

### Replace text
```csharp
ppt.ReplaceText("FY24", "FY25", includeTables: true, includeNotes: true);
```

## Feature Highlights

- Slides: add, import, duplicate, reorder, hide, and edit slides
- Shapes: basic rectangles/ellipses/lines with fill/stroke; align/distribute
- Images: add images from file/stream (PNG/JPEG/GIF/BMP/TIFF/EMF/WMF/ICO/PCX)
- Properties: set built‑in and application properties
- Themes & transitions: default theme/table styles + slide transitions
- Text boxes: margins, auto-fit, vertical alignment
- Tables: basic styling + merged cells
- Placeholders: read/update layout placeholders
- Backgrounds: set background images
- Text replacement: find/replace across slides

## Feature Matrix (scope today)

- 📽️ Slides
  - ✅ Add slides; ✅ import/duplicate/reorder; ✅ hide/show; ✅ set title; ✅ add text boxes; ✅ basic bullets
- 🖼️ Media & Shapes
  - ✅ Insert images; ✅ basic shapes (rect/ellipse/line) with fill/stroke
- 🗒️ Notes & Layout
  - ✅ Speaker notes; ⚠️ basic layout selection
- 📋 Tables
  - ⚠️ Basic styling + merged cells
- 📊 Charts
  - 🚧 Not yet
- ✨ Themes/Transitions
  - ✅ Default theme + full table styles; ✅ slide transitions (fade/wipe/push/etc.)

> Roadmap: richer shape/text APIs, layout/master controls, charts — tracked in issues.

## Why OfficeIMO.PowerPoint (today)

- Cross‑platform, pure Open XML — no Office automation
- Simple API surface to add slides, titles, text, bullets, and images without repair prompts
- Fluent helpers available for quick demos and templated decks

## Measurements

Positions and sizes are stored in EMUs (English Metric Units). Use `PowerPointUnits` or the `SetPositionCm`/`SetSizeCm`
helpers to work in centimeters, inches, or points.

### Layout helpers
```csharp
var content = ppt.SlideSize.GetContentBoxCm(1.5);
var columns = ppt.SlideSize.GetColumnsCm(2, marginCm: 1.5, gutterCm: 1.0);      
columns[0].ApplyTo(slide.AddTextBox("Left column"));
columns[1].ApplyTo(slide.AddTextBox("Right column"));
```

### Layout boxes as parameters
```csharp
var content = ppt.SlideSize.GetContentBoxCm(1.5);
var columns = content.SplitColumnsCm(2, 1.0);
slide.AddTextBox("Left column", columns[0]);
slide.AddPicture("logo.png", columns[1]);
```

## Dependencies & License

- DocumentFormat.OpenXml: 3.3.x (range [3.3.0, 4.0.0))
- License: MIT

<!-- (No migration notes: these APIs are new additions.) -->
