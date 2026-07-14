# OfficeIMO.Drawing.CodeGlyphX

`OfficeIMO.Drawing.CodeGlyphX` is an optional convenience package for turning CodeGlyphX QR codes, matrix symbols, and linear barcodes into reusable `OfficeDrawing` scenes.

Both core libraries remain independent. CodeGlyphX produces standard SVG without referencing OfficeIMO, and `OfficeIMO.Drawing` can read that SVG without referencing CodeGlyphX. Install this bridge only when typed extension methods make the handoff more convenient.

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

QrCode qr = QrCode.Encode("https://evotec.xyz");
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
