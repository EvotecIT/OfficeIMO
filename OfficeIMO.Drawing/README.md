# OfficeIMO.Drawing

OfficeIMO.Drawing is the shared first-party drawing layer for OfficeIMO packages. It provides small color and image metadata primitives without taking a dependency on a raster imaging library.

## What It Provides

- `OfficeColor`: immutable RGBA color value with named colors and hex parsing.
- `OfficeFontInfo`: immutable font family, size, and style descriptor for Office text features.
- `OfficeFontStyle`: dependency-free font style flags.
- `OfficeTrueTypeFont`: managed TrueType/OpenType outline reader for dependency-free raster text paths.
- `OfficeTextMeasurer`: deterministic text measurement estimates for Office layout decisions.
- `OfficeTextMeasurementStyle` and `OfficeTextMetrics`: normalized measurement inputs and pixel metrics.
- `OfficeImageReader`: header-only image metadata reader for common Office image formats.
- `OfficeImageInfo`: image format, dimensions, DPI, and MIME metadata.
- `OfficeImageFormat`: supported format enum used by OfficeIMO packages.
- `OfficeImageFit`: reusable image fitting intent for stretch, contain, and cover placement.
- `OfficeGradientStop` and `OfficeLinearGradient`: reusable two-stop linear gradient fill descriptors in normalized local coordinates.
- `OfficeShadow`: reusable shape shadow intent with color, opacity, and offset.
- `OfficePoint`, `OfficeTransform`, `OfficePathCommand`, `OfficeShape`, `OfficeShapeKind`, `OfficeStrokeDashStyle`, `OfficeStrokeLineCap`, and `OfficeStrokeLineJoin`: dependency-free vector shape descriptors that format-specific packages can map into their own coordinate systems.
- `OfficeDrawing` and `OfficeDrawingShape`: reusable drawing scenes made from positioned shared shapes.

## Supported Image Metadata

`OfficeImageReader` identifies PNG, JPEG, GIF, BMP, TIFF, ICO, PCX, EMF, placeable WMF, and SVG dimensions from headers or markup. It also maps Office-compatible extension-only formats so callers can still choose the right Open XML part type when dimensions are not available.

The reader is intentionally metadata-only. It does not decode pixels, resize, transcode, or validate complete image payloads.

`OfficeImageFit` gives format-specific renderers a shared way to describe image placement inside a target box:

```csharp
using OfficeIMO.Drawing;

var mode = OfficeImageFit.Contain;
```

`Stretch` fills the box exactly, `Contain` preserves aspect ratio while fitting the whole image, and `Cover` preserves aspect ratio while filling the box and clipping overflow.

## Color Migration

OfficeIMO packages now use `OfficeIMO.Drawing.OfficeColor` instead of external imaging color types.

```csharp
using OfficeIMO.Drawing;

var color = OfficeColor.Parse("#336699");
var accent = OfficeColor.CornflowerBlue;
```

`OfficeColor` accepts named colors, `#RGB`, `#RGBA`, `#RRGGBB`, and `#RRGGBBAA` values.

## Font Descriptors

`OfficeFontInfo` records the font family, point size, and style flags without taking a dependency on a font engine.

```csharp
using OfficeIMO.Drawing;

var font = new OfficeFontInfo("Calibri", 11, OfficeFontStyle.Bold | OfficeFontStyle.Italic | OfficeFontStyle.Underline);
```

## Text Measurement

`OfficeTextMeasurer` provides deterministic Office-oriented estimates for width and line height. It intentionally does not call operating-system font APIs, so results stay stable across machines.

```csharp
using OfficeIMO.Drawing;

var measurer = OfficeTextMeasurer.Create(OfficeFontInfo.Default);
var style = measurer.CreateStyle(new OfficeFontInfo("Aptos", 12, OfficeFontStyle.Bold));
OfficeTextMetrics metrics = measurer.Measure("OfficeIMO", style);
```

Rendering packages can use these estimates for layout planning while keeping public and shared APIs free of font-engine dependencies. Format-specific packages still own their own unit conversions and layout quirks, such as Excel column width units or PDF page coordinates.

## Managed Font Outlines

`OfficeTrueTypeFont` reads TrueType/OpenType font files directly and exposes flattened glyph contours for renderers that need dependency-free raster text. It does not call operating-system graphics or font APIs; callers can request a known font file or let OfficeIMO try common platform font file locations.

```csharp
using OfficeIMO.Drawing;

OfficeTrueTypeFont? font = OfficeTrueTypeFont.TryLoadDefault(out string? resolvedPath);
List<List<OfficePoint>> contours = font?.GetTextContours("OfficeIMO", 0, 0, 18) ?? new List<List<OfficePoint>>();
```

This is a low-level renderer building block, not a replacement for deterministic layout estimates. Use `OfficeTextMeasurer` for layout planning and `OfficeTrueTypeFont` only when a renderer needs actual glyph outlines.

## Vector Shape Descriptors

`OfficeShape` stores simple reusable vector shape intent without choosing a final file format or coordinate system.

```csharp
using OfficeIMO.Drawing;

var badge = OfficeShape.Rectangle(160, 48);
badge.FillColor = OfficeColor.WhiteSmoke;
badge.FillGradient = OfficeLinearGradient.Horizontal(OfficeColor.SteelBlue, OfficeColor.WhiteSmoke);
badge.Shadow = new OfficeShadow(OfficeColor.Black, 0.18, 3, 4);
badge.StrokeColor = OfficeColor.SteelBlue;
badge.StrokeWidth = 1.5;
badge.StrokeDashStyle = OfficeStrokeDashStyle.Dash;
badge.FillOpacity = 0.85;
badge.StrokeOpacity = 0.95;
badge.Transform = OfficeTransform.Translate(4, 8);
badge.ClipPath = OfficeClipPath.Rectangle(120, 36);

var pill = OfficeShape.RoundedRectangle(120, 32, 8);
pill.FillColor = OfficeColor.WhiteSmoke;
pill.StrokeColor = OfficeColor.SteelBlue;

var marker = OfficeShape.Ellipse(48, 24);
marker.FillColor = OfficeColor.SteelBlue;

var connector = OfficeShape.Line(0, 0, 120, 24);
connector.StrokeColor = OfficeColor.SteelBlue;
connector.StrokeWidth = 2;
connector.StrokeDashStyle = OfficeStrokeDashStyle.Dash;
connector.StrokeLineCap = OfficeStrokeLineCap.Round;

var triangle = OfficeShape.Polygon(
    new OfficePoint(0, 40),
    new OfficePoint(40, 0),
    new OfficePoint(80, 40));
triangle.FillColor = OfficeColor.WhiteSmoke;

var curve = OfficeShape.Path(
    OfficePathCommand.MoveTo(0, 40),
    OfficePathCommand.CubicBezierTo(20, 0, 60, 0, 80, 40),
    OfficePathCommand.Close());
curve.StrokeColor = OfficeColor.SteelBlue;
curve.StrokeLineJoin = OfficeStrokeLineJoin.Round;
```

`OfficeLinearGradient` stores two-stop linear fill intent in normalized local coordinates. PDF can map it to axial shading resources, while Open XML renderers can map the same descriptor to native drawing gradients.

```csharp
badge.FillGradient = OfficeLinearGradient.DiagonalDown(OfficeColor.SteelBlue, OfficeColor.WhiteSmoke);
```

`OfficeShadow` stores simple reusable shape-effect intent. Renderers can map it to native shadows, PDF alpha-backed offset geometry, or another format-specific effect model.

```csharp
badge.Shadow = new OfficeShadow(OfficeColor.Black, 0.2, offsetX: 3, offsetY: 4);
```

`OfficeTransform` stores affine transform intent in local top-left coordinates. Renderers can map it into PDF graphics state matrices, Open XML drawing transforms, or other format-native transform models.

```csharp
var rotated = OfficeTransform.RotateDegrees(15, centerX: 60, centerY: 20);
var movedAndScaled = OfficeTransform.Translate(12, 4).Then(OfficeTransform.Scale(1.2, 1.2));
```

`OfficeClipPath` stores reusable local clipping intent for rectangles, rounded rectangles, and freeform paths. Renderers can map it to PDF clipping paths, Open XML masks, or another format-native clipping mechanism.

```csharp
badge.ClipPath = OfficeClipPath.RoundedRectangle(120, 36, 8);
```

`OfficeDrawing` groups positioned shapes into a reusable local canvas. This is useful for logos, badges, simple diagrams, and future Office exporters that need to pass drawing intent into a format-specific renderer.

```csharp
using OfficeIMO.Drawing;

var background = OfficeShape.Rectangle(120, 60);
background.FillColor = OfficeColor.WhiteSmoke;

var connector = OfficeShape.Line(0, 0, 120, 60);
connector.StrokeColor = OfficeColor.SteelBlue;

var marker = OfficeShape.Polygon(
    new OfficePoint(0, 30),
    new OfficePoint(40, 0),
    new OfficePoint(80, 30));
marker.FillColor = OfficeColor.SteelBlue;

var scene = new OfficeDrawing(120, 60)
    .AddShape(background, 0, 0)
    .AddShape(connector, 0, 0)
    .AddShape(marker, 20, 15);
```

PDF, Word, Excel, PowerPoint, and other packages can map these shared descriptors into their own drawing models while keeping serialization details inside the format-specific package.
