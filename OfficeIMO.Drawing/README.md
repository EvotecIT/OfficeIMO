# OfficeIMO.Drawing - shared document and drawing primitives

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Drawing)](https://www.nuget.org/packages/OfficeIMO.Drawing)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Drawing?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Drawing)

`OfficeIMO.Drawing` is the zero-dependency shared foundation for OfficeIMO packages. It owns common document lifecycle contracts plus color conversion, image metadata, font and text measurement, vector scenes, reusable ink and math models, chart snapshots, SVG, raster canvases, PNG/JPEG/TIFF/WebP encoding, and drawing-quality primitives. Format packages keep their file-format behavior while reusing these document-agnostic models and renderers.

## Install

```powershell
dotnet add package OfficeIMO.Drawing
```

## Quick start

### Document lifecycle policy

```csharp
using OfficeIMO.Drawing;

var loadOptions = new DocumentLoadOptions {
    AccessMode = DocumentAccessMode.ReadOnly,
    PersistenceMode = DocumentPersistenceMode.Explicit
};
```

Word, Excel, and PowerPoint expose format-specific options derived from these shared contracts.

### Colors and vector intent

```csharp
using OfficeIMO.Drawing;

OfficeColor accent = OfficeColor.Parse("#336699");
OfficeColor printBlue = OfficeColorSpaceConverter.FromCmyk(1, 0.45, 0, 0.15);
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

// Use content verification when an extension must not identify invalid bytes.
byte[] bytes = File.ReadAllBytes("upload.svg");
bool verified = OfficeImageReader.TryIdentifyByContent(bytes, "upload.svg", out OfficeImageInfo upload);

OfficeImageFit fit = OfficeImageFit.Contain;
```

`TryIdentify(...)` retains the metadata reader's extension fallback. `TryIdentifyByContent(...)`
may use a file name to select the SVG parser, but succeeds only when the bytes match a supported format.

### Encode common raster formats

```csharp
using OfficeIMO.Drawing;

var image = new OfficeRasterImage(320, 180, OfficeColor.White);
var options = new OfficeRasterEncodingOptions {
    Jpeg = new OfficeJpegEncodeOptions { Quality = 90 },
    Tiff = new OfficeTiffEncodeOptions { Compression = OfficeTiffCompression.PackBits }
};

byte[] jpeg = OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Jpeg, options);
byte[] tiff = OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Tiff, options);
byte[] webp = OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Webp, options);
```

WebP output is deterministic, lossless VP8L. TIFF output is a single-page baseline RGBA image with uncompressed or PackBits strips. JPEG uses the existing managed quality, subsampling, progressive, metadata, and transparency-flattening settings.

`OfficeRasterExportPlanner` is the shared pre-allocation owner for image export. It combines the caller's `MaximumRasterPixels` with renderer and encoder dimension/pixel limits, then either reduces scale with `IMAGE_RASTER_SCALE_REDUCED` or throws `OfficeImageExportLimitException`, according to `RasterOverflowBehavior`. The returned plan also owns the effective encoding settings: `CreateEncodingOptions()` reduces encoded density with the raster scale so safety limits preserve the document's physical size. Explicit top-level `DpiX`/`DpiY` values apply across formats; when those values are not assigned, format-specific PNG, JPEG, and TIFF density remains authoritative. Drawing's managed PNG, JPEG, baseline chunky RGB/RGBA TIFF, uncompressed BMP, first-frame GIF, and OfficeIMO literal-lossless WebP decoders enforce encoded-payload and source-pixel guards; the TIFF and WebP readers intentionally reject variants outside those bounded profiles. `OfficeRasterImageFallbackCodec` can wrap an application codec at the final raster boundary. It reports `IMAGE_SOURCE_DECODED_BY_CALLER_CODEC` when that codec succeeds; if neither Drawing nor the application can decode a source image, it returns a visible placeholder and `IMAGE_SOURCE_DECODE_FALLBACK` instead of allowing the renderer to omit the image silently.

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

### Reusable ink

```csharp
using OfficeIMO.Drawing;

var stroke = new OfficeInkStroke {
    Color = OfficeColor.Crimson,
    Width = 2.2,
    Height = 2.2,
    Bias = OfficeInkBias.Handwriting,
    RecognizedText = "hello"
};
stroke.AddPoint(8, 30, 0.35)
      .AddPoint(60, 12, 1.0)
      .AddPoint(120, 36, 0.55);

var ink = new OfficeInkDocument().Add(stroke);
OfficeDrawing inkDrawing = OfficeInkRenderer.Render(ink, width: 140, height: 60);
```

The model keeps sampled geometry, normalized pressure, pen dimensions and tip shape, transforms, handwriting/drawing bias, language, and recognition alternatives. A format adapter decides how those values map to native storage.

### Structured math

```csharp
using OfficeIMO.Drawing;

OfficeMathExpression expression = OfficeMath.Fraction(
    OfficeMath.Row(
        OfficeMath.Identifier("x"),
        OfficeMath.Operator("+"),
        OfficeMath.Number("1")),
    OfficeMath.Radical(OfficeMath.Identifier("y")));

string mathMl = OfficeMathMarkup.ToMathMl(expression);
string latex = OfficeMathMarkup.ToLatex(expression);
OfficeDrawing mathDrawing = OfficeMathRenderer.Render(expression);
```

The same immutable expression tree feeds native OneNote math and Word OMML adapters. The shared model includes right and left scripts, centered upper/lower limits, built-up and slashed fractions, delimiter lists, stacks, matrices, equation arrays, n-ary operators, accents, bars, boxes, and phantoms. OneNote maps all of those structures natively. Word maps the lossless OMML subset; `Stack` and `StretchStack` fail with `NotSupportedException` because OMML has no equivalent, and callers can choose `EquationArray` explicitly when that projection is intended. Drawing owns the AST, portable markup, measurement, and visual layout; each document package owns only its native codec. MathML and LaTeX parsing default to a nesting limit of 128 and expose bounded overloads; excessive nesting fails with `OfficeMathParseException.Code == "DRAWING_MATH_DEPTH"`.

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

- `DocumentAccessMode`, `DocumentPersistenceMode`, `DocumentCreateOptions`, and `DocumentLoadOptions` for one lifecycle vocabulary across document packages.
- `OfficeColor` immutable RGBA values with named colors and hex parsing.
- `OfficeColorSpaceConverter` for dependency-free CMYK, CIE Lab/XYZ, calibrated gray, and calibrated RGB conversion to sRGB.
- `OfficeImageReader` and `OfficeImageInfo` for dependency-free image inspection where supported.
- `OfficeImageFit` for shared stretch, contain, and cover intent.
- `OfficeFontInfo`, `OfficeFontStyle`, `OfficeTextMeasurer`, and `OfficeTextMetrics` for deterministic layout estimates.
- `OfficeTrueTypeFont` for dependency-free font-outline reading when renderers need glyph contours.
- `OfficeInkDocument`, `OfficeInkStroke`, sampled pressure/style metadata, recognition alternatives, and `OfficeInkRenderer` for reusable ink capture and projection.
- `OfficeMathExpression`, `OfficeMath`, MathML/LaTeX conversion, deterministic measurement, and `OfficeMathRenderer` for reusable structured equations.
- `OfficeShape`, `OfficeDrawing`, gradients, shadows, transforms, clipping, and vector descriptors that format-specific packages can map into their own coordinate systems.
- `OfficeChartSnapshot` and chart rendering primitives shared by PDF and Office exporters.
- `OfficeRasterImage`, `OfficeRasterCanvas`, `OfficeRasterRenderTarget`, and `OfficeDrawingRasterRenderer` for shared dependency-free raster rendering.
- `OfficePngReader`, `OfficePngWriter`, and `OfficeJpegCodec` for PNG/JPEG paths that should not be reimplemented by document packages.
- `OfficeTiffCodec`, `OfficeWebpCodec`, and `OfficeRasterImageEncoder` for shared baseline TIFF, lossless WebP, and format-neutral raster output.
- Shared SVG formatting, primitive writing, image projection, text-block rendering, hatch-pattern, data-bar, and sparkline helpers.
- Drawing quality diagnostics for canvas bounds and text overlap checks.

## Boundaries

- This package owns shared lifecycle contracts, persistence mechanics, drawing intent, ink/math models and renderers, raster buffers, SVG and raster encoding primitives, image projection, text layout helpers, chart drawing, and document-agnostic visual diagnostics.
- Word, Excel, PowerPoint, Visio, and PDF packages own source-document semantics: package parsing, layout policy, coordinate systems, style/theme resolution, and user-facing export APIs.
- Document packages should not add private ink or math ASTs, pixel engines, image encoders/decoders, SVG primitive writers, text wrapping engines, or duplicate image-transform loops when the behavior can reasonably live here.
- PDF keeps PDF-stream and page-writer behavior in `OfficeIMO.Pdf`; when it needs generic image-like drawing, vector descriptors, colors, chart snapshots, PNG helpers, or raster visual QA, it should use `OfficeIMO.Drawing`.
- Unsupported or approximate source features belong in stable diagnostics from the adapter, not as silent omissions in a renderer.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** None.
- **OfficeIMO:** This is the shared foundation; it does not depend on another OfficeIMO runtime package.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
