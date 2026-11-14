# OfficeIMO.Drawing.CodeGlyphX

`OfficeIMO.Drawing.CodeGlyphX` is the optional bridge between [CodeGlyphX](https://github.com/EvotecIT/CodeGlyphX) and OfficeIMO. It renders QR codes and barcodes for the document adapters below, and turns CodeGlyphX QR codes, matrix symbols, and linear barcodes into reusable `OfficeDrawing` scenes.

The OfficeIMO engines remain independent of CodeGlyphX. CodeGlyphX produces standard SVG without referencing OfficeIMO, and `OfficeIMO.Drawing` can read that SVG without referencing CodeGlyphX. The bridge requires CodeGlyphX 3.1 or later within version 3.

## QR codes and barcodes in documents

Install the adapter for the format you write; it brings this package and CodeGlyphX with it. Each adds one call in the natural place:

| Package | Call |
| --- | --- |
| [`OfficeIMO.Word.CodeGlyphX`](../OfficeIMO.Word.CodeGlyphX/README.md) | `document.AddQrCode("https://evotec.xyz")`, `paragraph.AddBarcode(SymbolFormat.Code128, "ORD-42")` |
| [`OfficeIMO.Excel.CodeGlyphX`](../OfficeIMO.Excel.CodeGlyphX/README.md) | `sheet.AddQrCode("B2", "https://evotec.xyz")` |
| [`OfficeIMO.PowerPoint.CodeGlyphX`](../OfficeIMO.PowerPoint.CodeGlyphX/README.md) | `slide.AddQrCode("https://evotec.xyz", box)` |
| [`OfficeIMO.Pdf.CodeGlyphX`](../OfficeIMO.Pdf.CodeGlyphX/README.md) | `content.AddQrCode("https://evotec.xyz")`, `canvas.QrCode(...)` |

`AddBarcode` takes any CodeGlyphX `SymbolFormat`. Word, Excel and PowerPoint receive an SVG picture with a PNG fallback, the form Office itself writes; PDF receives vector paths.

### Options

Every call takes an optional `Action<OfficeBarcodeOptions>`:

| Option | Default |
| --- | --- |
| `Width` | 1.25 inches for two-dimensional codes; 0.33 mm per module for linear, stacked and postal codes. With only a width, the symbol keeps its proportions. |
| `Height` | From the proportions. Linear barcodes adjust their bar height to it; other symbols fit inside `Width` and `Height`. Linear bars default to 50 modules. |
| `QuietZone` | 4 modules for two-dimensional codes, 10 for linear barcodes. |
| `Color`, `BackgroundColor` | Black on white. Literal, opaque colors only, so every format shows the same contrast. |
| `ErrorCorrection` | Chosen by CodeGlyphX. QR codes only. |
| `HumanReadableText` | On for linear barcodes, using CodeGlyphX's readable text with check digits. Other symbologies have none. |
| `AlternativeText` | The symbology and the payload, for example `QR Code: https://evotec.xyz`. |

Invalid options and payloads throw before anything is added to the document.

## Install

```powershell
dotnet add package OfficeIMO.Drawing.CodeGlyphX
```

## QR code

```csharp
using CodeGlyphX;
using CodeGlyphX.Rendering.Svg;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.CodeGlyphX;

QrCode qr = QR.Encode("https://evotec.xyz");
OfficeDrawing drawing = qr.ToOfficeDrawing(new QrSvgRenderOptions {
    ModuleSize = 8,
    QuietZone = 4
});
```

## Matrix symbol

```csharp
using CodeGlyphX;
using CodeGlyphX.DataMatrix;
using CodeGlyphX.Rendering.Svg;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.CodeGlyphX;

BitMatrix modules = DataMatrixEncoder.Encode("ORDER-1234");
OfficeDrawing drawing = modules.ToOfficeDrawing(new MatrixSvgRenderOptions());
```

## Linear barcode with searchable text

```csharp
using CodeGlyphX;
using CodeGlyphX.Rendering.Svg;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.CodeGlyphX;

Barcode1D barcode = BarcodeEncoder.Encode(BarcodeType.Code128, "ORDER-1234");
OfficeDrawing drawing = barcode.ToOfficeDrawing(
    out int unsupportedFeatures,
    new BarcodeSvgRenderOptions { LabelText = "ORDER-1234" });

if (unsupportedFeatures != 0) {
    Console.WriteLine($"The SVG import used {unsupportedFeatures} fallback(s).");
}
```

The extension methods use the same neutral route available without this package: render SVG with CodeGlyphX, then pass its UTF-8 bytes to `OfficeSvgDrawingReader.TryRead`.
