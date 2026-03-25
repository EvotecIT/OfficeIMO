---
title: PowerPoint Presentations
description: Overview of PowerPoint presentation support in the OfficeIMO ecosystem.
order: 30
---

# PowerPoint Presentations

OfficeIMO provides basic support for creating and manipulating PowerPoint presentations (.pptx) through the Open XML SDK. While the Word and Excel packages offer the most mature APIs, PowerPoint support covers the essential scenarios for programmatic slide generation.

## Current Status

PowerPoint support in OfficeIMO is focused on creating presentations programmatically for reporting and automation scenarios. The API leverages the Open XML SDK's `PresentationDocument` class and wraps it in a simpler interface.

## Basic Usage

```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Word; // Shared helpers

// Create a new presentation
using var presentation = PresentationDocument.Create("slides.pptx", PresentationDocumentType.Presentation);

var presentationPart = presentation.AddPresentationPart();
presentationPart.Presentation = new Presentation();

// Add a slide layout and slide
var slidePart = presentationPart.AddNewPart<SlidePart>();
slidePart.Slide = new Slide(
    new CommonSlideData(
        new ShapeTree()
    )
);

presentation.Save();
```

## Key Concepts

### Slide Structure

A PowerPoint presentation follows this hierarchy:

```
PresentationDocument
  +-- PresentationPart
  |     +-- SlideMasterPart[]
  |     |     +-- SlideLayoutPart[]
  |     +-- SlidePart[]
  |           +-- Shapes, TextBoxes, Images, Charts
```

### Slide Masters and Layouts

Slide masters define the overall theme, color scheme, and font scheme. Slide layouts (Title Slide, Title and Content, Blank, etc.) inherit from the master and provide placeholder positions.

### Shapes and Text

Text in PowerPoint lives inside shapes. Each shape contains a text body with paragraphs and runs, similar to Word:

```csharp
var shape = new Shape();
shape.TextBody = new TextBody(
    new DocumentFormat.OpenXml.Drawing.Paragraph(
        new DocumentFormat.OpenXml.Drawing.Run(
            new DocumentFormat.OpenXml.Drawing.Text("Hello, PowerPoint!")
        )
    )
);
```

## Integration with OfficeIMO.Word

A common workflow is to generate data in OfficeIMO.Word or OfficeIMO.Excel, then embed or reference that content in a PowerPoint presentation:

1. Generate charts or tables with OfficeIMO.Excel.
2. Export charts as images.
3. Insert the images into PowerPoint slides.

## Further Reading

- [Slides](/docs/powerpoint/slides) -- Creating slides with text boxes, shapes, images, and charts.
- For advanced PowerPoint scenarios, consider the [Open XML SDK documentation](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/overview) and the [ShapeCrawler](https://github.com/ShapeCrawler/ShapeCrawler) library, which provides a higher-level API.
