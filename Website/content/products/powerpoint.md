---
title: "OfficeIMO.PowerPoint"
description: "Generate PowerPoint presentations with slides, charts, and shapes. No PowerPoint installation required."
layout: product
product_color: "#dc2626"
install: "dotnet add package OfficeIMO.PowerPoint"
nuget: "OfficeIMO.PowerPoint"
docs_url: "/docs/powerpoint/"
api_url: ""
---

## Why OfficeIMO.PowerPoint?

OfficeIMO.PowerPoint lets you create polished `.pptx` presentations from code. Automate slide decks for reporting pipelines, generate training materials, or build dynamic dashboards -- all without PowerPoint installed on your machine.

## Features

- **Slide creation & management** -- add, remove, reorder, and duplicate slides programmatically
- **Text boxes & bullets** -- rich text with fonts, colors, sizes, and multi-level bullet lists
- **Tables with merged cells** -- rows, columns, horizontal and vertical merges, and per-cell styling
- **Images** -- insert from file path or stream in PNG, JPEG, GIF, BMP, TIFF, EMF, and WMF formats
- **Shapes with fill, stroke & effects** -- rectangles, circles, arrows, callouts, and freeform shapes with shadow, glow, and blur effects
- **Charts with formatting** -- bar, column, pie, line, and area charts with data labels, legends, and axis configuration
- **Slide sections & transitions** -- organize slides into sections and apply transition animations
- **Themes & layouts** -- apply built-in or custom themes and choose from standard slide layouts
- **Speaker notes** -- attach presenter notes to any slide
- **Slide copying & importing** -- copy slides within a presentation or import from another `.pptx` file

## Quick start

```csharp
using OfficeIMO.PowerPoint;

using var presentation = PowerPointDocument.Create("Overview.pptx");

// Title slide
var titleSlide = presentation.AddSlide(SlideLayoutType.Title);
titleSlide.Title.Text = "Product Overview";
titleSlide.Subtitle.Text = "Engineering Team -- March 2026";

// Content slide with bullet points
var contentSlide = presentation.AddSlide(SlideLayoutType.TitleAndContent);
contentSlide.Title.Text = "Key Highlights";
var body = contentSlide.Content;
body.AddParagraph("Revenue grew 18% year-over-year");
body.AddParagraph("Launched 3 new product lines");
body.AddParagraph("Customer satisfaction at 94%");
body.AddParagraph("Expanded to 12 new markets");

// Chart slide
var chartSlide = presentation.AddSlide(SlideLayoutType.Blank);
var chart = chartSlide.AddChart(ChartType.ColumnClustered, 50, 80, 600, 350);
chart.Title.Text = "Revenue by Quarter";
chart.AddSeries("2025", new[] { "Q1", "Q2", "Q3", "Q4" }, new double[] { 3.2, 3.8, 4.1, 4.9 });
chart.AddSeries("2024", new[] { "Q1", "Q2", "Q3", "Q4" }, new double[] { 2.8, 3.1, 3.4, 3.9 });

presentation.Save();
```

## Compatibility

| Target Framework  | Supported |
|-------------------|-----------|
| .NET 10.0         | Yes       |
| .NET 8.0          | Yes       |
| .NET Standard 2.0 | Yes       |
| .NET Framework 4.7.2 | Yes   |

OfficeIMO.PowerPoint runs on Windows, Linux, and macOS. Output files are compatible with Microsoft PowerPoint, LibreOffice Impress, Google Slides, and Keynote.

## Related guides

| Guide | Description |
|-------|-------------|
| [PowerPoint documentation](/docs/powerpoint/) | Start with the package overview and supported presentation workflow. |
| [Slides guide](/docs/powerpoint/slides/) | Build title slides, content slides, charts, and slide layouts. |
| [PSWriteOffice](/products/pswriteoffice/) | Use PowerShell to automate the same presentation scenarios from scripts. |
