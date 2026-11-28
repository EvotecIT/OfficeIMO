# OfficeIMO.Excel.CodeGlyphX

QR codes and barcodes for `OfficeIMO.Excel` worksheets, rendered by [CodeGlyphX](https://github.com/EvotecIT/CodeGlyphX). `OfficeIMO.Excel` itself has no CodeGlyphX dependency; add this package when a workbook needs codes.

## Install

```powershell
dotnet add package OfficeIMO.Excel.CodeGlyphX
```

## Add a QR code

```csharp
using OfficeIMO;
using OfficeIMO.Excel;

using var workbook = ExcelDocument.Create("Catalog.xlsx");
ExcelSheet sheet = workbook.AddSheet("Products");
sheet.Cell(2, 1).SetValue("Contoso Router X2");
sheet.AddQrCode("B2", "https://contoso.example/x2");
workbook.Save();
```

The code is anchored at the given cell, or at a one-based row and column with `AddQrCode(row, column, payload)`. It keeps its own proportions, so the anchor is a single cell rather than a range. The methods return the `ExcelImage` for offsets, placement and other picture settings.

## Add any barcode

```csharp
using CodeGlyphX;
using OfficeIMO;
using OfficeIMO.Excel;

sheet.AddBarcode("C4", SymbolFormat.Ean, "590123412345");
sheet.AddBarcode(10, 2, SymbolFormat.DataMatrix, "SN-0042-PL", options => options.Width = OfficeLength.Millimeters(15));
```

Any CodeGlyphX `SymbolFormat` works. The optional callback sets `Width`, `Height`, `QuietZone`, `Color`, `BackgroundColor`, QR `ErrorCorrection`, linear `HumanReadableText` and `AlternativeText`; the defaults are listed in [OfficeIMO.Drawing.CodeGlyphX](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Drawing.CodeGlyphX/README.md#options).

## Output

The code is an SVG picture with a PNG fallback, the form Excel itself writes. `ExcelImage.ContentType` and `ToBytes()` return the SVG. `SaveAsPdf` from `OfficeIMO.Excel.Pdf` places the same module geometry on the PDF page, and the picture description defaults to the symbology and the encoded text.
