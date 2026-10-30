# OfficeIMO.PowerPoint

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.PowerPoint)](https://www.nuget.org/packages/OfficeIMO.PowerPoint)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.PowerPoint?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.PowerPoint)

Create, edit, read and save PowerPoint presentations (`.pptx`, and PowerPoint 97–2003 `.ppt`, `.pot` and `.pps`) in .NET, without Microsoft PowerPoint or COM. Slides, tables and charts are native PowerPoint objects that stay editable.

## Install

```sh
dotnet add package OfficeIMO.PowerPoint
dotnet add package OfficeIMO.PowerPoint.Pdf   # only for SaveAsPdf
```

Targets `netstandard2.0`, `net8.0` and `net10.0`, plus `net472` on Windows.

## Create a deck

Positions and sizes are `OfficeLength` values. `SlideSize.GetContentBox` gives the area inside a margin, and a `PowerPointLayoutBox` splits it into rows, columns or a title band, so slides line up without hand-computed coordinates:

```csharp
using OfficeIMO;
using OfficeIMO.PowerPoint;

using var deck = PowerPointDocument.Create("Review.pptx");
PowerPointLayoutBox page = deck.SlideSize.GetContentBox(OfficeLength.Centimeters(1.5));
PowerPointLayoutBox titleBand = page.TakeTop(OfficeLength.Centimeters(2.5));
PowerPointLayoutBox body = page.Inset(OfficeLength.Zero, OfficeLength.Centimeters(3.5), OfficeLength.Zero, OfficeLength.Zero);

deck.AddSlide().AddTitle("Quarterly business review", titleBand);

PowerPointSlide agenda = deck.AddSlide();
agenda.AddTitle("Agenda", titleBand);
agenda.AddTextBox(body).AddBullets(["Results", "Risks", "Decisions", "Next steps"]);
agenda.Notes.Text = "Keep the agenda under two minutes.";

deck.Save();
```

`Create` starts with no slides; each `AddSlide()` adds one.

## Tables and charts from your data

`AddTable` and `AddChart` take any `IEnumerable<T>`, the same way as in Word and Excel, and return the native table or chart:

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

PowerPointTable table = slide.AddTable(regions);
halves[0].ApplyTo(table);

slide.AddChart(OfficeChartType.Column, regions, nameof(Region.Name), new PowerPointChartOptions {
    Title = "Revenue (USD m)",
    Left = halves[1].Left, Top = halves[1].Top, Width = halves[1].Width, Height = halves[1].Height
}, nameof(Region.Revenue));

deck.Save();

record Region(string Name, decimal Revenue, double Margin);
```

## Images

```csharp
using OfficeIMO;
using OfficeIMO.PowerPoint;

using var deck = PowerPointDocument.Create("Logo.pptx");
PowerPointSlide slide = deck.AddSlide();
slide.AddTitle("Northwind Traders");
slide.AddImage("logo.png", left: OfficeLength.Centimeters(2), top: OfficeLength.Centimeters(5),
    width: OfficeLength.Centimeters(6), height: OfficeLength.Centimeters(6));
deck.Save();
```

## Edit an existing deck

```csharp
using OfficeIMO.PowerPoint;

using var deck = PowerPointDocument.Load("Review.pptx");
int replaced = deck.ReplaceText("Quarterly", "Annual", includeNotes: true);
foreach (PowerPointSlide slide in deck.Slides) {
    Console.WriteLine(slide.Notes.Text);
}
deck.Save("Review-annual.pptx");
```

For files from outside your organization, load with `PowerPointLoadOptions.UntrustedDefaults`, which rejects macros, embedded objects, ActiveX and external relationships before parsing.

## Save as PDF or images

```csharp
using OfficeIMO;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;

using var deck = PowerPointDocument.Load("Regions.pptx");
OfficeConversionReport pdf = deck.SaveAsPdf("Regions.pdf");
Console.WriteLine($"{pdf.ItemsWritten} slides, loss: {pdf.HasLoss}");

File.WriteAllBytes("Regions-1.png", deck.Slides[0].ToImage());
```

## More

- [PowerPoint guide](../Docs/officeimo.powerpoint-guide.md): text and paragraph formatting, shapes, the full chart API, designer composition and deck plans, corporate templates, sections, transitions, review threads, SmartArt, accessibility checks, legacy `.ppt` and image export.
- [Generated operation contract](../Docs/Compatibility/generated/package-operations.md): exactly what is supported, preserved or rejected.
- Adapters: [OfficeIMO.PowerPoint.Pdf](../OfficeIMO.PowerPoint.Pdf/README.md), [Html](../OfficeIMO.PowerPoint.Html/README.md), [OpenDocument](../OfficeIMO.PowerPoint.OpenDocument/README.md), [GoogleSlides](../OfficeIMO.PowerPoint.GoogleSlides/README.md) and [IWork](../OfficeIMO.PowerPoint.IWork/README.md). PDF back to PowerPoint lives in [OfficeIMO.Pdf.PowerPoint](../OfficeIMO.Pdf.PowerPoint/README.md).
- [Examples](../OfficeIMO.Examples/PowerPoint) and the [migration guide](../MIGRATION.md). `PowerPointPresentation` is now `PowerPointDocument`.
- PowerShell: [PSWriteOffice](../PowerShell/PSWriteOffice/README.MD) (`New-OfficePowerPoint`, `Export-OfficePowerPoint`).

Dependencies: the [Open XML SDK](https://github.com/dotnet/Open-XML-SDK) and `OfficeIMO.Core`. MIT licensed.

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
| Convert | 72 | 14 | 0 | 26 | 0 | 0 |
| Export | 5 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.PowerPoint` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
