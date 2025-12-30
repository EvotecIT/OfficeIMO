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
box.SetTextAutoFit(PowerPointTextAutoFit.Normal,
    new PowerPointTextAutoFitOptions(fontScalePercent: 85, lineSpaceReductionPercent: 10));
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

### Lines (dash + arrows)
```csharp
using A = DocumentFormat.OpenXml.Drawing;

var line = slide.AddLine(
    PowerPointUnits.Cm(1), PowerPointUnits.Cm(1),
    PowerPointUnits.Cm(8), PowerPointUnits.Cm(1));
line.OutlineColor = "2F5597";
line.OutlineDash = A.PresetLineDashValues.Dash;
line.SetLineEnds(A.LineEndValues.Triangle, A.LineEndValues.Stealth,
    A.LineEndWidthValues.Medium, A.LineEndLengthValues.Medium);
```

### Shape effects (shadow + glow + blur)
```csharp
var card = slide.AddRectangle(PowerPointUnits.Cm(1), PowerPointUnits.Cm(4),
    PowerPointUnits.Cm(6), PowerPointUnits.Cm(2), "Card");
card.FillColor = "FFFFFF";
card.OutlineColor = "D0D0D0";
card.SetShadow("000000", blurPoints: 6, distancePoints: 4,
    angleDegrees: 270, transparencyPercent: 40);
card.SetGlow("000000", radiusPoints: 3, transparencyPercent: 60);
card.SetSoftEdges(1.5);
card.SetBlur(2, grow: false);
```

### Shape effects (reflection)
```csharp
var logo = slide.AddRectangle(PowerPointUnits.Cm(10), PowerPointUnits.Cm(4),
    PowerPointUnits.Cm(4), PowerPointUnits.Cm(2), "Logo");
logo.SetReflection(blurPoints: 4, distancePoints: 2, directionDegrees: 270,
    fadeDirectionDegrees: 90, startOpacityPercent: 60, endOpacityPercent: 0);
```

### Align + distribute shapes
```csharp
slide.AlignShapes(slide.Shapes, PowerPointShapeAlignment.Left);
slide.DistributeShapes(slide.Shapes, PowerPointShapeDistribution.Horizontal);   

slide.AlignShapesToSlideContent(slide.Shapes, PowerPointShapeAlignment.Left,
    marginEmus: PowerPointUnits.Cm(1));
slide.DistributeShapesToSlideContent(slide.Shapes, PowerPointShapeDistribution.Horizontal,
    marginEmus: PowerPointUnits.Cm(1));
slide.DistributeShapes(slide.Shapes, PowerPointShapeDistribution.Horizontal,
    PowerPointShapeAlignment.Bottom);

slide.DistributeShapesWithSpacing(slide.Shapes, PowerPointShapeDistribution.Horizontal,
    spacingEmus: PowerPointUnits.Cm(0.5), center: true);

slide.DistributeShapesWithSpacing(slide.Shapes, PowerPointShapeDistribution.Horizontal,
    spacingEmus: PowerPointUnits.Cm(0.5), PowerPointShapeAlignment.Right);      
slide.DistributeShapesWithSpacing(slide.Shapes, PowerPointShapeDistribution.Horizontal,
    spacingEmus: PowerPointUnits.Cm(0.5), PowerPointShapeAlignment.Right,
    PowerPointShapeAlignment.Bottom);

slide.DistributeShapesWithSpacing(slide.Shapes, PowerPointShapeDistribution.Horizontal,
    new PowerPointShapeSpacingOptions {
        SpacingEmus = PowerPointUnits.Cm(0.4),
        ClampSpacingToBounds = true,
        Alignment = PowerPointShapeAlignment.Center,
        CrossAxisAlignment = PowerPointShapeAlignment.Bottom
    });
slide.DistributeShapesWithSpacing(slide.Shapes, PowerPointShapeDistribution.Horizontal,
    new PowerPointShapeSpacingOptions {
        SpacingEmus = PowerPointUnits.Cm(0.4),
        ScaleToFitBounds = true,
        PreserveAspect = true
    });

slide.DistributeShapesWithSpacingToSlideContent(slide.Shapes, PowerPointShapeDistribution.Horizontal,
    spacingEmus: PowerPointUnits.Cm(0.5), marginEmus: PowerPointUnits.Cm(1));

slide.StackShapes(slide.Shapes, PowerPointShapeStackDirection.Horizontal,
    spacingEmus: PowerPointUnits.Cm(0.3), PowerPointShapeAlignment.Bottom);
slide.StackShapesToSlideContent(slide.Shapes, PowerPointShapeStackDirection.Vertical,
    spacingEmus: PowerPointUnits.Cm(0.3), marginEmus: PowerPointUnits.Cm(1));

slide.StackShapes(slide.Shapes, PowerPointShapeStackDirection.Horizontal,
    spacingEmus: PowerPointUnits.Cm(0.3), PowerPointShapeStackJustify.Center);

slide.StackShapes(slide.Shapes, PowerPointShapeStackDirection.Horizontal,       
    new PowerPointShapeStackOptions {
        SpacingEmus = PowerPointUnits.Cm(0.4),
        ClampSpacingToBounds = true,
        Justify = PowerPointShapeStackJustify.Center
    });
slide.StackShapes(slide.Shapes, PowerPointShapeStackDirection.Horizontal,
    new PowerPointShapeStackOptions {
        SpacingEmus = PowerPointUnits.Cm(0.4),
        ScaleToFitBounds = true,
        PreserveAspect = true
    });
```

### Resize shapes
```csharp
slide.ResizeShapes(slide.Shapes, PowerPointShapeSizeDimension.Width, PowerPointShapeSizeReference.Largest);
slide.ResizeShapesCm(slide.Shapes, widthCm: 3.0, heightCm: null);
```

### Shape bounds + movement
```csharp
var shape = slide.AddRectangle(0, 0, PowerPointUnits.Cm(3), PowerPointUnits.Cm(2));

shape.Right = PowerPointUnits.Cm(10);
shape.CenterX = PowerPointUnits.Cm(12);
shape.Bounds = PowerPointLayoutBox.FromCentimeters(2, 2, 4, 2);

shape.MoveBy(PowerPointUnits.Cm(0.5), PowerPointUnits.Cm(0.25));
shape.Resize(PowerPointUnits.Cm(5), PowerPointUnits.Cm(2), PowerPointShapeAnchor.Center);
shape.Scale(1.2, PowerPointShapeAnchor.BottomRight);

var inArea = slide.GetShapesInBounds(
    PowerPointLayoutBox.FromCentimeters(1, 1, 5, 3),
    includePartial: true);
```

### Grid layout
```csharp
slide.ArrangeShapesInGrid(slide.Shapes,
    new PowerPointLayoutBox(0, 0, 8000, 4000),
    columns: 4, rows: 2, gutterX: 200, gutterY: 200,
    flow: PowerPointShapeGridFlow.RowMajor);

slide.ArrangeShapesInGridAuto(slide.Shapes,
    new PowerPointLayoutBox(0, 0, 8000, 4000));

slide.ArrangeShapesInGridAuto(slide.Shapes,
    new PowerPointLayoutBox(0, 0, 8000, 4000),
    new PowerPointShapeGridOptions { MinColumns = 2, MaxColumns = 4, TargetCellAspect = 1.0 });

slide.ArrangeShapesInGridToSlideContent(slide.Shapes, columns: 3, rows: 2,
    marginEmus: PowerPointUnits.Cm(1), gutterX: PowerPointUnits.Cm(0.5));
```

### Fit shapes to bounds
```csharp
slide.FitShapesToBounds(slide.Shapes, new PowerPointLayoutBox(0, 0, 8000, 4500));
slide.FitShapesToSlideContentCm(slide.Shapes, marginCm: 1.0, preserveAspect: true, center: true);
```

### Group + ungroup shapes
```csharp
var group = slide.GroupShapes(slide.Shapes);
slide.UngroupShape(group);

slide.AlignGroupChildren(group, PowerPointShapeAlignment.Left);
slide.DistributeGroupChildrenWithSpacing(group, PowerPointShapeDistribution.Horizontal,
    spacingEmus: PowerPointUnits.Cm(0.4));
slide.StackGroupChildren(group, PowerPointShapeStackDirection.Vertical,
    new PowerPointShapeStackOptions { SpacingEmus = PowerPointUnits.Cm(0.3) });
slide.ArrangeGroupChildrenInGrid(group, columns: 2, rows: 2);

var groupTextBoxes = slide.GetGroupTextBoxes(group);
var groupBoundsCm = slide.GetGroupChildBoundsCm(group);
```

### Arrange + duplicate shapes
```csharp
slide.BringForward(shape);
slide.SendBackward(shape);
slide.BringToFront(shape);
slide.SendToBack(shape);

var copy = slide.DuplicateShapeCm(shape, 0.5, 0.5);
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

### Slide layouts + theme helpers
```csharp
using DocumentFormat.OpenXml.Presentation;

var layouts = ppt.GetSlideLayouts();
var titleLayoutIndex = ppt.GetLayoutIndex(SlideLayoutValues.Title);

var slide = ppt.AddSlide(SlideLayoutValues.TitleOnly);
slide.SetLayout("Title and Content");

ppt.SetThemeColor(PowerPointThemeColor.Accent1, "FF0000");
ppt.SetThemeLatinFonts("Aptos", "Calibri");
ppt.SetThemeFonts(new PowerPointThemeFontSet(
    majorLatin: "Aptos",
    minorLatin: "Calibri",
    majorEastAsian: "MS Mincho",
    minorEastAsian: "Yu Gothic",
    majorComplexScript: "Arial",
    minorComplexScript: "Tahoma"));
```

### Import slide from another deck
```csharp
using var source = PowerPointPresentation.Open("source.pptx");
var imported = ppt.ImportSlide(source, sourceIndex: 0);
imported.AddTextBox("Imported content");
```

### Charts (formatting)
```csharp
using C = DocumentFormat.OpenXml.Drawing.Charts;

var chart = slide.AddChart();
chart.SetTitle("Sales Trend")
     .SetLegend(C.LegendPositionValues.Right)
     .SetDataLabels(showValue: true)
     .SetCategoryAxisTitle("Quarter")
     .SetValueAxisTitle("Revenue")
     .SetValueAxisNumberFormat("#,##0.00")
     .SetSeriesFillColor(0, "4472C4")
     .SetSeriesLineColor("Series 2", "ED7D31", widthPoints: 1)
     .SetSeriesMarker(0, C.MarkerStyleValues.Circle, size: 6, fillColor: "FFFFFF", lineColor: "4472C4");
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

### Tables (formatting helpers)
```csharp
using A = DocumentFormat.OpenXml.Drawing;

var table = slide.AddTable(rows: 3, columns: 3, left: PowerPointUnits.Cm(1), top: PowerPointUnits.Cm(5),
    width: PowerPointUnits.Cm(18), height: PowerPointUnits.Cm(5));

table.SetRowHeightsEvenly();
table.SetCellPaddingCm(0.2, 0.1, 0.2, 0.1);
table.SetCellAlignment(A.TextAlignmentTypeValues.Center, A.TextAnchoringTypeValues.Center,
    startRow: 0, endRow: 0, startColumn: 0, endColumn: 2);
table.SetCellBorders(TableCellBorders.All, "E0E0E0", widthPoints: 0.5,
    dash: A.PresetLineDashValues.Dash);
```

### Placeholders (layout-driven)
```csharp
using DocumentFormat.OpenXml.Presentation;

var slide = ppt.AddSlide(masterIndex: 0, layoutIndex: 1);
var title = slide.GetPlaceholder(PlaceholderValues.Title);
title?.SetTextMarginsCm(0.2, 0.1, 0.2, 0.1);
if (title != null) title.Text = "Layout Placeholder";

var layoutPlaceholders = slide.GetLayoutPlaceholders();
var titleBounds = slide.GetLayoutPlaceholderBounds(PlaceholderValues.Title);
if (titleBounds != null) {
    var aligned = slide.AddTextBox("Aligned to layout");
    titleBounds.Value.ApplyTo(aligned);
}
```

### Replace text
```csharp
ppt.ReplaceText("FY24", "FY25", includeTables: true, includeNotes: true);
```

### Sections
```csharp
ppt.AddSection("Intro", startSlideIndex: 0);
ppt.AddSection("Results", startSlideIndex: 2);
ppt.RenameSection("Results", "Deep Dive");

var sections = ppt.GetSections();
```

## Feature Highlights

- Slides: add, import, duplicate, reorder, hide, edit, and section slides
- Sections: add, rename, and list sections
- Shapes: basic rectangles/ellipses/lines with fill/stroke, line styles, shadows/glow/soft edges/blur/reflection; align/distribute; z-order + duplicate
- Images: add images from file/stream (PNG/JPEG/GIF/BMP/TIFF/EMF/WMF/ICO/PCX)   
- Properties: set built‑in and application properties
- Themes & transitions: default theme/table styles + slide transitions
- Text boxes: margins, auto-fit, vertical alignment
- Tables: styling + merged cells + sizing helpers
- Placeholders: read/update layout placeholders
- Backgrounds: set background images
- Text replacement: find/replace across slides
- Charts: add + format titles/legend/labels/series/markers

## Feature Matrix (scope today)

- 📽️ Slides
  - ✅ Add slides; ✅ import/duplicate/reorder; ✅ hide/show; ✅ sections; ✅ set title; ✅ add text boxes; ✅ basic bullets
- 🖼️ Media & Shapes
  - ✅ Insert images; ✅ basic shapes (rect/ellipse/line) with fill/stroke + line styles + shadows/glow/soft edges/blur/reflection; ✅ align/distribute/arrange
- 🗒️ Notes & Layout
  - ✅ Speaker notes; ⚠️ basic layout selection
- 📋 Tables
  - ⚠️ Basic styling + merged cells
- 📊 Charts
  - ✅ Add charts; ✅ title/legend/labels; ✅ axis formatting; ✅ series fill/line/markers
- ✨ Themes/Transitions
  - ✅ Default theme + full table styles; ✅ slide transitions (fade/wipe/push/etc.)

> Roadmap: richer shape/text APIs, layout/master controls, advanced charts — tracked in issues.

## Why OfficeIMO.PowerPoint (today)

- Cross‑platform, pure Open XML — no Office automation
- Simple API surface to add slides, titles, text, bullets, and images without repair prompts
- Fluent helpers available for quick demos and templated decks

## Measurements

Positions and sizes are stored in EMUs (English Metric Units). Use `PowerPointUnits` or the `SetPositionCm`/`SetSizeCm`
helpers to work in centimeters, inches, or points.

### Slide size presets
```csharp
ppt.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);
ppt.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen4x3, portrait: true);
ppt.SlideSize.SetSizeCm(25.4, 14.0); // custom size
```

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
