---
title: Slides
description: "Build PowerPoint slides with titles, text boxes, bullets, tables, charts, notes, and slide management helpers. Includes examples and API links."
order: 31
---

This guide covers the main `OfficeIMO.PowerPoint` slide-building workflow. Instead of manually creating `SlidePart` trees, you work with `PowerPointSlide`, `PowerPointTextBox`, `PowerPointTable`, and related helpers directly.

## Start with a title and body region

```csharp
using OfficeIMO;
using OfficeIMO.PowerPoint;
using OfficeIMO.Drawing;

using var presentation = PowerPointDocument.Create("presentation.pptx");
const double marginCm = 1.5;

var content = presentation.SlideSize.GetContentBox(OfficeLength.Centimeters(marginCm));
var slide = presentation.AddSlide();

slide.AddTitle(
    "Quarterly Results",
    OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),
    OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters()),
    OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),
    OfficeIMO.OfficeLength.Centimeters(1.4));

slide.AddTextBox(
    "Generated from the reporting pipeline.",
    OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),
    OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters() + 1.8),
    OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),
    OfficeIMO.OfficeLength.Centimeters(1.0));
```

## Add bullets and formatted text

```csharp
using OfficeIMO;
using OfficeIMO.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;

var text = slide.AddTextBox(
    string.Empty,
    OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),
    OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters() + 3.0),
    OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters() / 2),
    OfficeIMO.OfficeLength.Centimeters(4.5));

text.SetTextMarginsCm(0.3, 0.2, 0.3, 0.2);
text.TextAutoFit = PowerPointTextAutoFit.Normal;

var heading = text.AddParagraph("Highlights", p => {
    p.Alignment = PowerPointTextAlignment.Left;
    p.SpaceAfterPoints = 6;
});
PowerPointTextStyle.Subtitle.WithColor("1F4E79").Apply(heading);

var line = text.AddParagraph();
line.AddText("Ship ");
line.AddFormattedText("faster", bold: true).SetColor("C00000");
line.AddText(" with cleaner automated decks.");

text.AddBullet("Release completed on schedule");
text.AddBullet("Customer guide updated");
text.AddNumberedItem("Review adoption metrics");
text.AddNumberedItem("Prepare next milestone");
text.ApplyAutoSpacing(lineSpacingMultiplier: 1.15, spaceAfterPoints: 2);
```

## Build tables and charts from data

```csharp
using OfficeIMO;
using OfficeIMO.PowerPoint;

var rows = new[] {
    new { Product = "Alpha", Q1 = 12, Q2 = 15, Q3 = 18, Q4 = 20 },
    new { Product = "Beta",  Q1 =  9, Q2 = 11, Q3 = 13, Q4 = 14 },
    new { Product = "Gamma", Q1 =  6, Q2 =  9, Q3 = 12, Q4 = 16 }
};

var tableSlide = presentation.AddSlide();
tableSlide.AddTitle("Sales by Product", OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()), OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters()), OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()), OfficeIMO.OfficeLength.Centimeters(1.4));

var table = tableSlide.AddTable(
    rows,
    options => {
        options.HeaderCase = HeaderCase.Title;
        options.PinFirst("Product");
    },
    includeHeaders: true,
    left: content.Left,
    top: content.Top + OfficeLength.Centimeters(2.0),
    width: content.Width,
    height: OfficeLength.Centimeters(4.5));

table.SetColumnWidthsEvenly();

var chartSlide = presentation.AddSlide();
chartSlide.AddTitle("Quarterly Performance", OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()), OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters()), OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()), OfficeIMO.OfficeLength.Centimeters(1.4));

var chartData = new OfficeChartData(
    rows.Select(r => r.Product),
    new[] {
        new OfficeSeries("Q1", rows.Select(r => (double)r.Q1)),
        new OfficeSeries("Q2", rows.Select(r => (double)r.Q2)),
        new OfficeSeries("Q3", rows.Select(r => (double)r.Q3)),
        new OfficeSeries("Q4", rows.Select(r => (double)r.Q4))
    });

chartSlide.AddChart(
    OfficeChartType.ColumnClustered,
    chartData,
    OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),
    OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters() + 2.0),
    OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),
    OfficeIMO.OfficeLength.Centimeters(4.8));
chartSlide.Notes.Text = "Chart and table share the same source data.";
```

For a native bubble chart, create each series with explicit X, Y, and size values:

```csharp
var bubbleData = new OfficeChartData(
    new[] { "1", "2", "3" },
    new[] {
        OfficeSeries.CreateBubble(
            "Portfolio",
            xValues: new[] { 1D, 2D, 3D },
            yValues: new[] { 12D, 18D, 15D },
            bubbleSizes: new[] { 9D, 36D, 16D })
    });

chartSlide.AddChart(
    OfficeChartType.Bubble,
    bubbleData,
    OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),
    OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters() + 2.0),
    OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),
    OfficeIMO.OfficeLength.Centimeters(4.8));
```

Bubble sizes must be finite and non-negative, with one size for every X/Y point. The chart remains editable in PowerPoint, and `UpdateData(...)` refreshes both cached chart values and the embedded workbook.

## Manage slide order and reuse

```csharp
var intro = presentation.AddSlide();
intro.AddTitle("Intro", OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()), OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters()), OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()), OfficeIMO.OfficeLength.Centimeters(1.4));

var duplicate = presentation.DuplicateSlide(0);
duplicate.Hidden = true;

using var source = PowerPointDocument.Create("source-deck.pptx");
source.AddSlide().AddTitle("Imported slide", OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()), OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters()), OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()), OfficeIMO.OfficeLength.Centimeters(1.4));
source.Save();

presentation.ImportSlide(source, 0, insertAt: 1);
presentation.MoveSlide(2, 0);
```

## Common patterns

- Use `AddTitleCm` and `AddTextBoxCm` when you want a readable layout in source code.
- Use `SetTextMarginsCm`, `ApplyTextStyle`, and `TextAutoFit` early so slides stay legible as content grows.
- Build data-heavy slides from objects with `AddTable` and shared `OfficeChartData` instead of hard-coded cell text.
- Use `Notes.Text` for speaker notes and `DuplicateSlide`, `ImportSlide`, and `MoveSlide` when you need reusable deck templates.

## Related guides

- [PowerPoint overview](/docs/powerpoint/) — Package scope and workflow.
- [Designer Decks](/docs/powerpoint/designer/) — Use higher-level design briefs, semantic plans, and composer primitives when manual placement becomes too repetitive.
- [PSWriteOffice PowerPoint Cmdlets](/docs/pswriteoffice/powerpoint/) — Build slides from PowerShell.
- [OfficeIMO.PowerPoint product page](/products/powerpoint/) — Install and package overview.
