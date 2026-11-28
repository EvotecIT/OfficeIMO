# OfficeIMO.PowerPoint.CodeGlyphX

QR codes and barcodes for `OfficeIMO.PowerPoint` slides, rendered by [CodeGlyphX](https://github.com/EvotecIT/CodeGlyphX). `OfficeIMO.PowerPoint` itself has no CodeGlyphX dependency; add this package when a deck needs codes.

## Install

```powershell
dotnet add package OfficeIMO.PowerPoint.CodeGlyphX
```

## Add a QR code

Give the code a layout box and it takes the largest size that fits, centered in the box:

```csharp
using OfficeIMO;
using OfficeIMO.PowerPoint;

using var presentation = PowerPointDocument.Create("Talk.pptx");
PowerPointSlide slide = presentation.AddSlide();
PowerPointLayoutBox content = presentation.SlideSize.GetContentBox(OfficeLength.Inches(0.5));
slide.AddQrCode("https://evotec.xyz/talks/officeimo", content);
presentation.Save();
```

To pin a code at a position instead, pass its top-left corner; the size comes from the options or the 1.25-inch default:

```csharp
slide.AddQrCode("https://evotec.xyz", OfficeLength.Centimeters(22), OfficeLength.Centimeters(15),
    options => options.Width = OfficeLength.Centimeters(2.5));
```

Both return the `PowerPointImage` for links, animation and other picture settings.

## Add any barcode

```csharp
using CodeGlyphX;
using OfficeIMO;
using OfficeIMO.PowerPoint;

slide.AddBarcode(SymbolFormat.Code128, "TICKET-7781", box);
```

Any CodeGlyphX `SymbolFormat` works. In a box, linear barcodes fill its width and height; two-dimensional codes keep their proportions. The optional callback sets `Width`, `Height`, `QuietZone`, `Color`, `BackgroundColor`, QR `ErrorCorrection`, linear `HumanReadableText` and `AlternativeText`; the defaults are listed in [OfficeIMO.Drawing.CodeGlyphX](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Drawing.CodeGlyphX/README.md#options).

## Output

The code is an SVG picture with a PNG fallback, the form PowerPoint itself writes. `SaveAsPdf` from `OfficeIMO.PowerPoint.Pdf` places the same module geometry on the PDF page, and the picture's alternative text defaults to the symbology and the encoded text.
