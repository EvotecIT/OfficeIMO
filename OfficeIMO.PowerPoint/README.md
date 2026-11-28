# OfficeIMO.PowerPoint

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.PowerPoint)](https://www.nuget.org/packages/OfficeIMO.PowerPoint)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.PowerPoint?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.PowerPoint)

Create, edit, read and convert PowerPoint presentations in .NET without Microsoft PowerPoint or COM. `OfficeIMO.PowerPoint` works with `.pptx`, `.pptm`, templates, shows and add-ins, and PowerPoint 97–2003 `.ppt`, `.pot` and `.pps`. Slides, tables and charts are native PowerPoint objects that stay editable.

## Install

```sh
dotnet add package OfficeIMO.PowerPoint
```

Add [OfficeIMO.PowerPoint.Pdf](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.PowerPoint.Pdf/README.md) only if you need `SaveAsPdf`.

## Quick start

A slide with a title, a table and a native chart from the same data, laid out with layout boxes instead of hand-computed coordinates:

<!-- powerforge:code source="Build/DocumentationExamples/PowerPoint/Program.cs" region="QuickStart" -->
```csharp
using OfficeIMO;
using OfficeIMO.PowerPoint;

Region[] regions = [new("North", 5.7m, 0.36), new("South", 4.1m, 0.31), new("West", 2.9m, 0.27)];

using var deck = PowerPointDocument.Create("Regions.pptx");
PowerPointLayoutBox page = deck.SlideSize.GetContentBox(OfficeLength.Centimeters(1.5));
PowerPointLayoutBox titleBand = page.TakeTop(OfficeLength.Centimeters(2.5));
PowerPointLayoutBox[] halves = page.Inset(OfficeLength.Zero, OfficeLength.Centimeters(3.5), OfficeLength.Zero, OfficeLength.Zero)
    .SplitColumns(2, OfficeLength.Centimeters(1));

PowerPointSlide slide = deck.AddSlide();
slide.AddTitle("Revenue by region", titleBand);
halves[0].ApplyTo(slide.AddTable(regions));
slide.AddChart(OfficeChartType.Column, regions, nameof(Region.Name), new PowerPointChartOptions {
    Title = "Revenue (USD m)",
    Left = halves[1].Left, Top = halves[1].Top, Width = halves[1].Width, Height = halves[1].Height
}, nameof(Region.Revenue));

deck.Save();

record Region(string Name, decimal Revenue, double Margin);
```
<!-- /powerforge:code -->

## What it does

- **Build slides** with titles, text boxes, bullets, images, shapes, connectors, tables, native charts, SmartArt and speaker notes. Positions and sizes are `OfficeLength` values, and `PowerPointLayoutBox` splits a slide into rows, columns and bands.
- **Masters and layouts**: add shapes and placeholders that show on every slide, create layouts and bind content to placeholders with `BindToPlaceholder`.
- **Edit existing decks**: `ReplaceAll` across slides, tables and, on request, speaker notes; reorder, duplicate and remove slides; sections and transitions.
- **Corporate templates and deck plans**: compose designer-style slides from a template and a plan.
- **Read large decks slide by slide** with `PowerPointSlideReader`, which returns titles, shapes, tables, image references and notes without loading the whole presentation, and extracts Markdown chunks.
- **Review**: comments and review threads.
- **Before sharing**: `InspectHiddenData` and `RemoveHiddenData` remove properties, comments, notes, custom XML and image metadata, and optionally hidden slides and off-slide shapes.
- **Render** slides and individual charts to PNG, SVG and other image formats.
- **Load untrusted files** with `PowerPointLoadOptions.UntrustedDefaults`, which rejects macros, embedded objects, ActiveX and external relationships before parsing.

The [PowerPoint guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.powerpoint-guide.md) shows each of these with examples.

## Converting to other formats

| Target | Package |
| --- | --- |
| PDF | [OfficeIMO.PowerPoint.Pdf](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.PowerPoint.Pdf/README.md) |
| HTML | [OfficeIMO.PowerPoint.Html](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.PowerPoint.Html/README.md) |
| OpenDocument presentations | [OfficeIMO.PowerPoint.OpenDocument](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.PowerPoint.OpenDocument/README.md) |
| Google Slides | [OfficeIMO.PowerPoint.GoogleSlides](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.PowerPoint.GoogleSlides/README.md) |
| Apple Keynote | [OfficeIMO.PowerPoint.IWork](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.PowerPoint.IWork/README.md) |
| PDF back to PowerPoint | [OfficeIMO.Pdf.PowerPoint](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Pdf.PowerPoint/README.md) |

Saving to another format returns an `OfficeConversionReport` that lists anything the target couldn't carry.

## Limits

`.pptx` is the native format. Legacy `.ppt` loads into the same model and saves natively within a documented subset. `PowerPointSlideReader` reads text as written on each slide and doesn't read inside SmartArt or charts; encrypted and `.ppt` files need `PowerPointDocument.Load`. The [legacy PPT contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/powerpoint-legacy-ppt.md) and the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md) list exactly what is supported, preserved or rejected.

## Learn more

- [OfficeIMO.PowerPoint on officeimo.com](https://officeimo.com/products/powerpoint/) and the [PowerPoint docs](https://officeimo.com/docs/powerpoint/).
- [PowerPoint guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.powerpoint-guide.md).
- [Examples](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Examples/PowerPoint) and the [migration guide](https://github.com/EvotecIT/OfficeIMO/blob/master/MIGRATION.md).
- PowerShell: [PSWriteOffice](https://github.com/EvotecIT/OfficeIMO/blob/master/PowerShell/PSWriteOffice/README.MD) (`New-OfficePowerPoint`, `Export-OfficePowerPoint`). Command line: `officeimo convert deck.pptx deck.pdf` with [OfficeIMO.Tool](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Tool/README.md).

## Targets and dependencies

- Targets: `netstandard2.0`, `net8.0` and `net10.0`, plus `net472` on Windows.
- Dependencies: the [Open XML SDK](https://github.com/dotnet/Open-XML-SDK) and `OfficeIMO.Core`. Add `OfficeIMO.Security` only to validate package signatures.
- License: MIT.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 48 | 2 | 0 | 10 | 0 | 1 |
| Read | 47 | 0 | 8 | 4 | 0 | 2 |
| Edit | 21 | 0 | 34 | 6 | 0 | 0 |
| Preserve | 19 | 0 | 34 | 4 | 0 | 0 |
| Inspect | 6 | 0 | 0 | 0 | 0 | 0 |
| Validate | 4 | 0 | 0 | 0 | 0 | 2 |
| Remove | 5 | 0 | 0 | 0 | 0 | 0 |
| Convert | 72 | 15 | 0 | 26 | 0 | 0 |
| Export | 5 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.PowerPoint` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
