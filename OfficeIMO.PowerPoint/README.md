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
box.AddBullet("Intro");
box.AddBullet("KPIs");
box.AddBullet("Next Steps");
```

### Images
```csharp
slide.AddPicture("logo.png",
    PowerPointUnits.Cm(23), PowerPointUnits.Cm(1.2), PowerPointUnits.Cm(5), PowerPointUnits.Cm(2));
```

### Simple shapes
```csharp
slide.AddRectangle(PowerPointUnits.Cm(1), PowerPointUnits.Cm(1),
    PowerPointUnits.Cm(3), PowerPointUnits.Cm(1))
    .Fill("#E7F7FF")
    .Stroke("#007ACC");
```

### Slide properties
```csharp
ppt.BuiltinDocumentProperties.Title = "Contoso Review";
ppt.ApplicationProperties.Company = "Contoso";
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

## Feature Highlights

- Slides: add slides and text boxes, titles
- Shapes: basic rectangles/ellipses/lines with fill/stroke
- Images: add images from file/stream
- Properties: set built‑in and application properties

## Feature Matrix (scope today)

- 📽️ Slides
  - ✅ Add slides; ✅ set title; ✅ add text boxes; ✅ basic bullets
- 🖼️ Media & Shapes
  - ✅ Insert images; ✅ basic shapes (rect/ellipse/line) with fill/stroke
- 🗒️ Notes & Layout
  - ✅ Speaker notes; ⚠️ basic layout selection
- 📋 Tables
  - ⚠️ Basic only (where supported)
- 📊 Charts
  - 🚧 Not yet
- ✨ Themes/Transitions
  - 🚧 Not yet

> Roadmap: richer shape/text APIs, layout/mast er controls, charts, transitions — tracked in issues.

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

## Dependencies & License

- DocumentFormat.OpenXml: 3.3.x (range [3.3.0, 4.0.0))
- License: MIT

<!-- (No migration notes: these APIs are new additions.) -->
