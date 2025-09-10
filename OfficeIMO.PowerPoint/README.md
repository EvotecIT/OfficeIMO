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

using var ppt = PowerPointDocument.Create("demo.pptx");
var slide = ppt.AddSlide();
slide.AddTitle("Hello PowerPoint");
slide.AddTextBox(2, 2, 6, 2, "Generated with OfficeIMO.PowerPoint");
ppt.Save();
```

## Common Tasks by Example

### Title + content
```csharp
var slide = ppt.AddSlide();
slide.AddTitle("Quarterly Review");
slide.AddTextBox(1.5, 2.5, 7.5, 3.0, "Agenda\n• Intro\n• KPIs\n• Next Steps");
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
slide.AddImage("logo.png", left: 9.0, top: 0.5, width: 2.0, height: 0.8);
```

### Simple shapes
```csharp
slide.AddRectangle(1,1, 3,1).Fill("#E7F7FF").Stroke("#007ACC");
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

## Dependencies & License

- DocumentFormat.OpenXml: 3.3.x (range [3.3.0, 4.0.0))
- License: MIT

<!-- (No migration notes: these APIs are new additions.) -->

