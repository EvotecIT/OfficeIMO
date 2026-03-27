---
title: Slides
description: Build PowerPoint slides with titles, text boxes, bullets, tables, charts, notes, and slide management helpers.
order: 31
---

# Slides

This guide covers the main `OfficeIMO.PowerPoint` slide-building workflow. Instead of manually creating `SlidePart` trees, you work with `PowerPointSlide`, `PowerPointTextBox`, `PowerPointTable`, and related helpers directly.

## Start with a title and body region

```csharp
using OfficeIMO.PowerPoint;

using var presentation = PowerPointPresentation.Create("presentation.pptx");
const double marginCm = 1.5;

var content = presentation.SlideSize.GetContentBoxCm(marginCm);
var slide = presentation.AddSlide();

slide.AddTitleCm(
    "Quarterly Results",
    content.LeftCm,
    content.TopCm,
    content.WidthCm,
    1.4);

slide.AddTextBoxCm(
    "Generated from the reporting pipeline.",
    content.LeftCm,
    content.TopCm + 1.8,
    content.WidthCm,
    1.0);
```

## Add bullets and formatted text

```csharp
using OfficeIMO.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;

var text = slide.AddTextBoxCm(
    string.Empty,
    content.LeftCm,
    content.TopCm + 3.0,
    content.WidthCm / 2,
    4.5);

text.SetTextMarginsCm(0.3, 0.2, 0.3, 0.2);
text.TextAutoFit = PowerPointTextAutoFit.Normal;

var heading = text.AddParagraph("Highlights", p => {
    p.Alignment = A.TextAlignmentTypeValues.Left;
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
using OfficeIMO.PowerPoint;

var rows = new[] {
    new { Product = "Alpha", Q1 = 12, Q2 = 15, Q3 = 18, Q4 = 20 },
    new { Product = "Beta",  Q1 =  9, Q2 = 11, Q3 = 13, Q4 = 14 },
    new { Product = "Gamma", Q1 =  6, Q2 =  9, Q3 = 12, Q4 = 16 }
};

var tableSlide = presentation.AddSlide();
tableSlide.AddTitleCm("Sales by Product", content.LeftCm, content.TopCm, content.WidthCm, 1.4);

var table = tableSlide.AddTableCm(
    rows,
    options => {
        options.HeaderCase = HeaderCase.Title;
        options.PinFirst("Product");
    },
    includeHeaders: true,
    leftCm: content.LeftCm,
    topCm: content.TopCm + 2.0,
    widthCm: content.WidthCm,
    heightCm: 4.5);

table.SetColumnWidthsEvenly();

var chartSlide = presentation.AddSlide();
chartSlide.AddTitleCm("Quarterly Performance", content.LeftCm, content.TopCm, content.WidthCm, 1.4);

var chartData = new PowerPointChartData(
    rows.Select(r => r.Product),
    new[] {
        new PowerPointChartSeries("Q1", rows.Select(r => (double)r.Q1)),
        new PowerPointChartSeries("Q2", rows.Select(r => (double)r.Q2)),
        new PowerPointChartSeries("Q3", rows.Select(r => (double)r.Q3)),
        new PowerPointChartSeries("Q4", rows.Select(r => (double)r.Q4))
    });

chartSlide.AddChartCm(
    chartData,
    content.LeftCm,
    content.TopCm + 2.0,
    content.WidthCm,
    4.8);
chartSlide.Notes.Text = "Chart and table share the same source data.";
```

## Manage slide order and reuse

```csharp
var intro = presentation.AddSlide();
intro.AddTitleCm("Intro", content.LeftCm, content.TopCm, content.WidthCm, 1.4);

var duplicate = presentation.DuplicateSlide(0);
duplicate.Hidden = true;

using var source = PowerPointPresentation.Create("source-deck.pptx");
source.AddSlide().AddTitleCm("Imported slide", content.LeftCm, content.TopCm, content.WidthCm, 1.4);
source.Save();

presentation.ImportSlide(source, 0, insertAt: 1);
presentation.MoveSlide(2, 0);
```

## Common patterns

- Use `AddTitleCm` and `AddTextBoxCm` when you want a readable layout in source code.
- Use `SetTextMarginsCm`, `ApplyTextStyle`, and `TextAutoFit` early so slides stay legible as content grows.
- Build data-heavy slides from objects with `AddTableCm` and `PowerPointChartData` instead of hard-coded cell text.
- Use `Notes.Text` for speaker notes and `DuplicateSlide`, `ImportSlide`, and `MoveSlide` when you need reusable deck templates.

## Related guides

- [PowerPoint overview](/docs/powerpoint/) -- Package scope, workflow, and positioning.
- [PSWriteOffice PowerPoint Cmdlets](/docs/pswriteoffice/powerpoint/) -- Build slides from PowerShell.
- [OfficeIMO.PowerPoint product page](/products/powerpoint/) -- Install and package overview.
