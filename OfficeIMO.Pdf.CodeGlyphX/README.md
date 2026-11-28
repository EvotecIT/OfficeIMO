# OfficeIMO.Pdf.CodeGlyphX

QR codes and barcodes for `OfficeIMO.Pdf`, rendered by [CodeGlyphX](https://github.com/EvotecIT/CodeGlyphX) and drawn as vector paths. `OfficeIMO.Pdf` itself has no CodeGlyphX dependency; add this package when a PDF needs codes.

## Install

```powershell
dotnet add package OfficeIMO.Pdf.CodeGlyphX
```

## Add a QR code to the content flow

```csharp
using OfficeIMO;
using OfficeIMO.Pdf;

PdfDocument.Create(pdf => pdf.Content(content => content
        .H1("Invoice 2026-0042")
        .AddParagraph("Pay online:")
        .AddQrCode("https://pay.example.com/inv/2026-0042", align: OfficeAlignment.Right)))
    .Save("Invoice.pdf");
```

## Add any barcode

```csharp
using CodeGlyphX;
using OfficeIMO;
using OfficeIMO.Pdf;

content.AddBarcode(SymbolFormat.Gs1Code128, "(01)09501101530003(17)261231");
content.AddBarcode(SymbolFormat.DataMatrix, "SN-0042-PL", options => options.Width = OfficeLength.Millimeters(15));
```

Any CodeGlyphX `SymbolFormat` works. The optional callback sets `Width`, `Height`, `QuietZone`, `Color`, `BackgroundColor`, QR `ErrorCorrection`, linear `HumanReadableText` and `AlternativeText`; the defaults are listed in [OfficeIMO.Drawing.CodeGlyphX](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Drawing.CodeGlyphX/README.md#options).

## Place a code at a page position

Labels and tickets often need a code at an exact spot. Draw it on a page canvas with top-left page coordinates:

```csharp
PdfDocument.Create(pdf => pdf.Page(page => page.Canvas(canvas => canvas
        .QrCode("TICKET-7781", OfficeLength.Millimeters(20), OfficeLength.Millimeters(30),
            options => options.Width = OfficeLength.Millimeters(25)))))
    .Save("Ticket.pdf");
```

## Output

The code is a vector drawing, so it stays sharp at any zoom or print resolution. Its alternative text, which defaults to the symbology and the encoded text, is written to the tagged PDF.
