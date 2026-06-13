# OfficeIMO.PowerPoint - PowerPoint presentations for .NET

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.PowerPoint)](https://www.nuget.org/packages/OfficeIMO.PowerPoint)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.PowerPoint?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.PowerPoint)

`OfficeIMO.PowerPoint` creates and edits `.pptx` presentations with Open XML. It is for generating editable decks without COM automation and without Microsoft PowerPoint installed.

If OfficeIMO saves you time, please consider supporting the work through [GitHub Sponsors](https://github.com/sponsors/PrzemyslawKlys) or [PayPal](https://paypal.me/PrzemyslawKlys). PowerShell users should use [PSWriteOffice](https://github.com/EvotecIT/PSWriteOffice) for the PowerShell-facing experience.

## Install

```powershell
dotnet add package OfficeIMO.PowerPoint
```

## Quick start

```csharp
using OfficeIMO.PowerPoint;

using var presentation = PowerPointPresentation.Create("deck.pptx");
presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

var slide = presentation.AddSlide();
slide.AddTitle("OfficeIMO.PowerPoint");

var body = slide.AddTextBox("Generated without PowerPoint automation.");
body.SetPositionCm(2, 2);
body.SetSizeCm(18, 2);

slide.Transition = SlideTransition.Fade;
presentation.Save();
```

## What it does

- Creates and edits PowerPoint presentations, slides, slide size, text boxes, pictures, tables, charts, backgrounds, transitions, notes, and metadata.
- Keeps generated output as editable PowerPoint content instead of screenshots.
- Provides designer composition helpers for theme-aware business decks and repeatable layout alternatives.
- Supports encrypted presentation save/open workflows.
- Uses Open XML directly, making it suitable for services, build agents, desktop apps, and automation hosts.

## Runnable samples

```powershell
dotnet run --project OfficeIMO.Examples/OfficeIMO.Examples.csproj -f net10.0 -- --powerpoint
dotnet run --project OfficeIMO.Examples/OfficeIMO.Examples.csproj -f net10.0 -- --modern-powerpoint
dotnet run --project OfficeIMO.Examples/OfficeIMO.Examples.csproj -f net10.0 -- --powerpoint-design-brief
dotnet run --project OfficeIMO.Examples/OfficeIMO.Examples.csproj -f net10.0 -- --powerpoint-deck-plan
```

## Examples

The quick start creates one simple slide. These examples show the editable deck features that belong in `OfficeIMO.PowerPoint`.

### Title, content, and bullets

```csharp
var slide = presentation.AddSlide();
slide.AddTitle("Quarterly Review");

slide.AddTextBox("Agenda",
    PowerPointUnits.Cm(1.5), PowerPointUnits.Cm(2.0),
    PowerPointUnits.Cm(8.0), PowerPointUnits.Cm(1.0));

var agenda = slide.AddTextBox("Topics",
    PowerPointUnits.Cm(1.5), PowerPointUnits.Cm(3.0),
    PowerPointUnits.Cm(10.0), PowerPointUnits.Cm(3.0));
agenda.AddBullets(new[] { "Intro", "KPIs", "Next steps" });
```

### Images and SVGs

```csharp
slide.AddPicture("logo.png",
    PowerPointUnits.Cm(23), PowerPointUnits.Cm(1.2),
    PowerPointUnits.Cm(5), PowerPointUnits.Cm(2));

using var logo = File.OpenRead("logo.png");
slide.AddPicture(logo, ImagePartType.Png,
    PowerPointUnits.Cm(2), PowerPointUnits.Cm(2),
    PowerPointUnits.Cm(5), PowerPointUnits.Cm(2));

slide.AddPicture("diagram.svg",
    PowerPointUnits.Cm(2), PowerPointUnits.Cm(5),
    PowerPointUnits.Cm(8), PowerPointUnits.Cm(4));
```

### Shapes and layout

```csharp
slide.AddRectangle(
        PowerPointUnits.Cm(1), PowerPointUnits.Cm(1),
        PowerPointUnits.Cm(3), PowerPointUnits.Cm(1))
    .Fill("#E7F7FF")
    .Stroke("#007ACC");

slide.AddLine(
        PowerPointUnits.Cm(1), PowerPointUnits.Cm(3),
        PowerPointUnits.Cm(8), PowerPointUnits.Cm(3))
    .Stroke("#404040");
```

### Tables from data

```csharp
record SalesRow(string Product, int Q1, int Q2);

var rows = new[] {
    new SalesRow("Alpha", 12, 15),
    new SalesRow("Beta", 9, 11)
};

var columns = new[] {
    PowerPointTableColumn<SalesRow>.Create("Product", row => row.Product).WithWidthCm(4.0),
    PowerPointTableColumn<SalesRow>.Create("Q1", row => row.Q1),
    PowerPointTableColumn<SalesRow>.Create("Q2", row => row.Q2)
};

slide.AddTable(rows, columns,
    left: PowerPointUnits.Cm(1.5),
    top: PowerPointUnits.Cm(4),
    width: PowerPointUnits.Cm(20),
    height: PowerPointUnits.Cm(6));
```

### Charts from data

```csharp
using DocumentFormat.OpenXml.Drawing.Charts;

record MetricRow(string Quarter, double Revenue, double Margin);

var metrics = new[] {
    new MetricRow("Q1", 120, 32),
    new MetricRow("Q2", 145, 36),
    new MetricRow("Q3", 172, 41),
    new MetricRow("Q4", 190, 44)
};

var slide = presentation.AddSlide();
slide.AddTitle("Revenue and margin");

slide.AddChartCm(metrics, row => row.Quarter,
        leftCm: 1.4, topCm: 3.0, widthCm: 13.2, heightCm: 8.0,
        new PowerPointChartSeriesDefinition<MetricRow>("Revenue", row => row.Revenue),
        new PowerPointChartSeriesDefinition<MetricRow>("Margin", row => row.Margin))
    .SetTitle("Quarterly performance")
    .SetCategoryAxisTitle("Quarter")
    .SetValueAxisTitle("Value")
    .SetLegend(LegendPositionValues.Bottom)
    .SetChartAreaStyle(fillColor: "FFFFFF", lineColor: "D9E2F3")
    .SetPlotAreaStyle(fillColor: "F8FAFC", lineColor: "D9E2F3");
```

```csharp
var mix = new PowerPointChartData(
    new[] { "Services", "Licenses", "Support" },
    new[] { new PowerPointChartSeries("Share", new[] { 55d, 30d, 15d }) });

slide.AddDoughnutChartCm(mix, leftCm: 15.2, topCm: 3.0, widthCm: 8.0, heightCm: 8.0)
    .SetTitle("Revenue mix")
    .SetLegend(LegendPositionValues.Right);
```

### Table and chart together

```csharp
record SegmentRow(string Segment, int Q1, int Q2, int Q3, int Q4);

var segments = new[] {
    new SegmentRow("Enterprise", 18, 22, 29, 35),
    new SegmentRow("SMB", 12, 14, 18, 21),
    new SegmentRow("Public", 9, 11, 12, 16)
};

var dashboard = presentation.AddSlide();
dashboard.AddTitle("Segment dashboard");

dashboard.AddTableCm(segments, new[] {
        PowerPointTableColumn<SegmentRow>.Create("Segment", row => row.Segment).WithWidthCm(4.0),
        PowerPointTableColumn<SegmentRow>.Create("Q1", row => row.Q1),
        PowerPointTableColumn<SegmentRow>.Create("Q2", row => row.Q2),
        PowerPointTableColumn<SegmentRow>.Create("Q3", row => row.Q3),
        PowerPointTableColumn<SegmentRow>.Create("Q4", row => row.Q4)
    },
    leftCm: 1.4, topCm: 3.0, widthCm: 10.0, heightCm: 5.0);

dashboard.AddLineChartCm(segments, row => row.Segment,
        leftCm: 12.2, topCm: 3.0, widthCm: 12.0, heightCm: 6.5,
        new PowerPointChartSeriesDefinition<SegmentRow>("Q1", row => row.Q1),
        new PowerPointChartSeriesDefinition<SegmentRow>("Q2", row => row.Q2),
        new PowerPointChartSeriesDefinition<SegmentRow>("Q3", row => row.Q3),
        new PowerPointChartSeriesDefinition<SegmentRow>("Q4", row => row.Q4))
    .SetTitle("Segment trend")
    .SetLegend(LegendPositionValues.Bottom);
```

### Slides, notes, and duplication

```csharp
var duplicate = presentation.DuplicateSlide(0);
duplicate.Hidden = true;
duplicate.Notes.Text = "Backup slide for Q&A.";
```

### Designer composition

Use the designer APIs when a deck needs readable business composition without hand-positioning every object:

```csharp
using var presentation = PowerPointPresentation.Create("proposal.pptx");

var brief = PowerPointDesignBrief
    .FromBrand("#008C95", "client-demo", "technical rollout proposal")
    .WithIdentity("Client Theme", footerLeft: "CLIENT", footerRight: "Service deck");

var deck = presentation.UseDesigner(brief, alternativeIndex: 0);
deck.AddSectionSlide("Delivery plan", "Implementation overview");
presentation.Save();
```

### Fluent authoring

```csharp
using OfficeIMO.PowerPoint.Fluent;

presentation.AsFluent()
   .Slide(masterIndex: 0, layoutIndex: 0, slide => {
       slide.Title("Fluent Slide");
       slide.Bullets("One", "Two", "Three");
       slide.Notes("Talking points for the presenter");
   });
```

## Adjacent packages

| Package | Use it for |
| --- | --- |
| [OfficeIMO.PowerPoint.Pdf](../OfficeIMO.PowerPoint.Pdf/README.md) | Export editable PowerPoint slides to PDF and import PDF tables to PowerPoint. |
| [OfficeIMO.Markup.PowerPoint](../OfficeIMO.Markup.PowerPoint/README.md) | Render semantic markup documents to PowerPoint. |

## Boundaries

- `OfficeIMO.PowerPoint` owns editable PowerPoint creation and manipulation.
- PDF export belongs in `OfficeIMO.PowerPoint.Pdf` and shared PDF primitives belong in `OfficeIMO.Pdf`.
- Showcase decks and long design examples belong in `OfficeIMO.Examples` or focused docs, not in the package README.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`; `net472` is included when building on Windows.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)
