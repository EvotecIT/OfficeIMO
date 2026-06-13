# OfficeIMO.Drawing - shared drawing primitives

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Drawing)](https://www.nuget.org/packages/OfficeIMO.Drawing)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Drawing?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Drawing)

`OfficeIMO.Drawing` is the shared first-party drawing layer for OfficeIMO packages. It provides color, image metadata, font, text measurement, vector shape, chart snapshot, and drawing-quality primitives without taking a dependency on a raster imaging library.

## Install

```powershell
dotnet add package OfficeIMO.Drawing
```

## Quick start

### Colors and vector intent

```csharp
using OfficeIMO.Drawing;

OfficeColor accent = OfficeColor.Parse("#336699");
OfficeImageFit fit = OfficeImageFit.Contain;

var badge = OfficeShape.RoundedRectangle(120, 32, 8);
badge.FillColor = OfficeColor.WhiteSmoke;
badge.StrokeColor = accent;
badge.Shadow = new OfficeShadow(OfficeColor.Black, 0.18, 3, 4);
```

### Image metadata without decoding pixels

```csharp
using OfficeIMO.Drawing;

OfficeImageInfo info = OfficeImageReader.Identify("logo.png");
Console.WriteLine($"{info.Width}x{info.Height} {info.MimeType}");

OfficeImageFit fit = OfficeImageFit.Contain;
```

### Deterministic text measurement

```csharp
using OfficeIMO.Drawing;

var measurer = OfficeTextMeasurer.Create();
var style = measurer.CreateStyle(new OfficeFontInfo("Aptos", 11, OfficeFontStyle.Regular));
OfficeTextMetrics metrics = measurer.Measure("Quarterly report", style);

if (metrics.WidthPixels > 240) {
    Console.WriteLine("The label needs wrapping or a smaller font.");
}
```

## Examples

### Build a reusable vector scene

```csharp
using OfficeIMO.Drawing;

var drawing = new OfficeDrawing(width: 420, height: 180)
    .AddShape(new OfficeShape {
        Kind = OfficeShapeKind.Rectangle,
        Width = 420,
        Height = 180,
        FillGradient = OfficeLinearGradient.Horizontal(
            OfficeColor.Parse("#F8FBFF"),
            OfficeColor.Parse("#EAF4FF")),
        StrokeColor = OfficeColor.Parse("#B7D7F5"),
        StrokeWidth = 1
    }, x: 0, y: 0)
    .AddText("OfficeIMO.Drawing", 20, 18, 380, 32,
        new OfficeFontInfo("Aptos", 18, OfficeFontStyle.Bold),
        OfficeColor.Parse("#1F2937"),
        OfficeTextAlignment.Left)
    .AddShape(OfficeShape.RoundedRectangle(140, 44, 10), 20, 86)
    .AddText("Shared vector intent", 34, 98, 240, 24);

OfficeDrawingQualityReport report = OfficeDrawingQualityAnalyzer.Analyze(drawing);
if (report.HasIssues) {
    foreach (var issue in report.Issues) {
        Console.WriteLine($"{issue.Kind}: {issue.Message}");
    }
}
```

### Render a chart snapshot to drawing primitives

```csharp
using OfficeIMO.Drawing;

var snapshot = new OfficeChartSnapshot(
    name: "RevenueChart",
    title: "Revenue by quarter",
    chartKind: OfficeChartKind.ColumnClustered,
    data: new OfficeChartData(
        new[] { "Q1", "Q2", "Q3", "Q4" },
        new[] {
            new OfficeChartSeries("Revenue", new[] { 10d, 18d, 24d, 30d }),
            new OfficeChartSeries("Forecast", new[] { 12d, 19d, 25d, 33d })
        }),
    widthPoints: 420,
    heightPoints: 260);

OfficeChartRenderingResult rendered = OfficeChartDrawingRenderer.RenderWithQuality(snapshot);
OfficeDrawing chartDrawing = rendered.Drawing;

foreach (var issue in rendered.QualityReport.Issues) {
    Console.WriteLine(issue.Message);
}
```

### Read TrueType outlines for renderers

```csharp
using OfficeIMO.Drawing;

OfficeTrueTypeFont? font = OfficeTrueTypeFont.TryLoadDefault(out string? path);
if (font != null) {
    Console.WriteLine($"Loaded {path}");
}
```

## What it provides

- `OfficeColor` immutable RGBA values with named colors and hex parsing.
- `OfficeImageReader` and `OfficeImageInfo` for metadata-only image inspection.
- `OfficeImageFit` for shared stretch, contain, and cover intent.
- `OfficeFontInfo`, `OfficeFontStyle`, `OfficeTextMeasurer`, and `OfficeTextMetrics` for deterministic layout estimates.
- `OfficeTrueTypeFont` for dependency-free font-outline reading when renderers need glyph contours.
- `OfficeShape`, `OfficeDrawing`, gradients, shadows, transforms, clipping, and vector descriptors that format-specific packages can map into their own coordinate systems.
- `OfficeChartSnapshot` and chart rendering primitives shared by PDF and Office exporters.
- Drawing quality diagnostics for canvas bounds and text overlap checks.

## Boundaries

- This package describes reusable drawing intent.
- Word, Excel, PowerPoint, Visio, and PDF packages map that intent into their own file formats.
- Pixel decoding, resizing, transcoding, and full image validation are not part of this runtime package.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)
