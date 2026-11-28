# OfficeIMO.Word.CodeGlyphX

QR codes and barcodes for `OfficeIMO.Word` documents, rendered by [CodeGlyphX](https://github.com/EvotecIT/CodeGlyphX). `OfficeIMO.Word` itself has no CodeGlyphX dependency; add this package when a document needs codes.

## Install

```powershell
dotnet add package OfficeIMO.Word.CodeGlyphX
```

## Add a QR code

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using var document = WordDocument.Create("Flyer.docx");
document.AddParagraph("Scan for the event agenda:");
document.AddQrCode("https://evotec.xyz/events");
document.Save();
```

`AddQrCode` and `AddBarcode` exist on `WordDocument`, which adds a new paragraph, and on `WordParagraph`, which places the code inline in that paragraph. Both return the `WordImage`, so wrapping, alignment and other picture settings work as usual.

## Add any barcode

Pass a CodeGlyphX `SymbolFormat` for any other symbology: `Code128`, `Gs1Code128`, `Ean`, `UpcA`, `DataMatrix`, `Pdf417`, `Aztec` and the rest of the formats CodeGlyphX encodes.

```csharp
using CodeGlyphX;
using OfficeIMO;
using OfficeIMO.Word;

paragraph.AddBarcode(SymbolFormat.Code128, "ORD-2026-0042", options => {
    options.Width = OfficeLength.Millimeters(60);
    options.Color = OfficeColor.Parse("#1F3864");
});
```

The optional callback sets `Width`, `Height`, `QuietZone`, `Color`, `BackgroundColor`, QR `ErrorCorrection`, linear `HumanReadableText` and `AlternativeText`. The defaults are listed in [OfficeIMO.Drawing.CodeGlyphX](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Drawing.CodeGlyphX/README.md#options).

## Output

The code is an SVG picture with a PNG fallback, the form Word itself writes: Word 2016 and later draw the SVG; older readers show the print-resolution PNG. `SaveAsPdf` from `OfficeIMO.Word.Pdf` places the same module geometry on the PDF page. The picture's alternative text defaults to the symbology and the encoded text, such as `QR Code: https://evotec.xyz/events`.
