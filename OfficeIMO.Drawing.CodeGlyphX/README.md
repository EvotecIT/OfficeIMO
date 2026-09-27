# OfficeIMO.Drawing.CodeGlyphX

`OfficeIMO.Drawing.CodeGlyphX` is an optional convenience package for turning CodeGlyphX QR codes, matrix symbols, and linear barcodes into reusable `OfficeDrawing` scenes.

Both core libraries remain independent. CodeGlyphX produces standard SVG without referencing OfficeIMO, and `OfficeIMO.Drawing` can read that SVG without referencing CodeGlyphX. Use this bridge for typed symbol extensions or its optional raster decoder.

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

## Optional raster decoding

`CodeGlyphRasterImageCodec` implements Drawing's image-codec boundary. The managed Drawing decoder handles its built-in formats first; the optional codec can decode inspected payloads outside that subset, including lossy WebP.

```csharp
var options = new OfficeRasterDecodeOptions {
    ImageCodec = new CodeGlyphRasterImageCodec(),
    MaximumDecodedPixels = 8_000_000
};
if (OfficeRasterImageDecoder.TryDecode(imageBytes, options, out var image, out var report)) {
    // report describes any animation or additional frames discarded by the static result.
}
```

The same codec can be supplied through `ImageCodec` on HTML render or PDF options. Decoding preserves the configured resource limits; the synchronous codec call observes export cancellation before and after it runs. Pixel fidelity remains the codec provider's responsibility.
