# OfficeIMO.Drawing

OfficeIMO.Drawing is the shared first-party drawing layer for OfficeIMO packages. It provides small color and image metadata primitives without taking a dependency on a raster imaging library.

## What It Provides

- `OfficeColor`: immutable RGBA color value with named colors and hex parsing.
- `OfficeFontInfo`: immutable font family, size, and style descriptor for Office text features.
- `OfficeFontStyle`: dependency-free font style flags.
- `OfficeImageReader`: header-only image metadata reader for common Office image formats.
- `OfficeImageInfo`: image format, dimensions, DPI, and MIME metadata.
- `OfficeImageFormat`: supported format enum used by OfficeIMO packages.

## Supported Image Metadata

`OfficeImageReader` identifies PNG, JPEG, GIF, BMP, TIFF, ICO, PCX, EMF, placeable WMF, and SVG dimensions from headers or markup. It also maps Office-compatible extension-only formats so callers can still choose the right Open XML part type when dimensions are not available.

The reader is intentionally metadata-only. It does not decode pixels, resize, transcode, or validate complete image payloads.

## Color Migration

OfficeIMO packages now use `OfficeIMO.Drawing.OfficeColor` instead of external imaging color types.

```csharp
using OfficeIMO.Drawing;

var color = OfficeColor.Parse("#336699");
var accent = OfficeColor.CornflowerBlue;
```

`OfficeColor` accepts named colors, `#RGB`, `#RGBA`, `#RRGGBB`, and `#RRGGBBAA` values.

## Font Descriptors

`OfficeFontInfo` is a metadata contract only. It records the font family, point size, and bold/italic flags without measuring or rendering text.

```csharp
using OfficeIMO.Drawing;

var font = new OfficeFontInfo("Calibri", 11, OfficeFontStyle.Bold | OfficeFontStyle.Italic);
```

Rendering packages can translate this descriptor into their own measurement backend while keeping public and shared APIs free of font-engine dependencies.
