# OfficeIMO.Pdf.Ocr - OCR and searchable PDF integration

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Pdf.Ocr)](https://www.nuget.org/packages/OfficeIMO.Pdf.Ocr)

`OfficeIMO.Pdf.Ocr` connects any `OfficeIMO.Ocr.IOcrEngine` to first-party PDF page rendering, native-text overlap filtering, logical reconstruction, and searchable PDF output. OCR is optional and is not part of the base `OfficeIMO.Pdf` dependency graph.

## Install

Install the PDF integration and one provider. For Tesseract:

```powershell
dotnet add package OfficeIMO.Pdf.Ocr
dotnet add package OfficeIMO.Ocr.Tesseract
```

Tesseract itself remains a separately installed host dependency. A custom or hosted provider only needs the `OfficeIMO.Ocr` contract.

## Read scanned and mixed PDFs

```csharp
using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;

var engine = TesseractOcrEngine.CreateDefault();
PdfDocument pdf = PdfDocument.Load("mixed-report.pdf");

PdfOcrMergeResult result = await pdf.ReadWithOcrAsync(
    engine,
    new PdfOcrMergeOptions {
        Language = "eng+pol",
        Dpi = 180,
        MinimumConfidence = 0.75,
        ReadOptions = new PdfReadOptions {
            LayoutOptions = new PdfTextLayoutOptions {
                ReadingDirection = PdfReadingDirection.Auto
            }
        }
    });

Console.WriteLine(result.Document.Text);
Console.WriteLine($"Accepted OCR words: {result.AcceptedWordCount}");
```

Every selected page is rendered to a bounded raster request. Pixel, point, and normalized provider coordinates are projected into the page's cropped and rotated visual point space. Low-confidence spans and spans overlapping native text are rejected before OCR evidence enters the same language-neutral reading-order, region, list, paragraph, heading, and table pipeline as native positioned text.

`NativeDocument` retains the native-only parse. `Document` is the canonical native-plus-OCR parse and can be passed directly to the existing PDF-to-Word, Excel, PowerPoint, HTML, RTF, or OpenDocument adapters. Page results retain accepted words, provider/model/language evidence, rejections, and diagnostics.

## Add a searchable text layer

```csharp
PdfSearchableOcrResult searchable = await pdf.MakeSearchableAsync(engine);
await searchable.Document.SaveAsync("mixed-report-searchable.pdf");

Console.WriteLine($"Modified pages: {string.Join(", ", searchable.ModifiedPages)}");
Console.WriteLine($"Added words: {searchable.AddedWordCount}");
```

Only pages with accepted OCR words are rewritten. The invisible text layer follows the canonical semantic order, while the returned OCR result records exactly what was added. Signed or otherwise rewrite-sensitive documents remain subject to the base PDF mutation and preservation rules.

`PdfOcrMergeOptions` bounds provider-call duration, rendered pixels, selected pages, inspected spans, accepted OCR words and characters, aggregate raw hierarchy identifiers, provider metadata and diagnostics, native-overlap comparisons, and merged text. Calls use one shared `OcrEngineExecution` per document, so identity and capabilities are stable across pages and the same non-concurrent engine instance cannot overlap across PDF, Reader, or a future integration. Language is provider configuration only; it is never used to infer captions, lists, paragraphs, tables, or continuations.

## Targets and dependency footprint

- Targets: `netstandard2.0`, `net8.0`, `net10.0` (`net472` is also included on Windows builds).
- OfficeIMO dependencies: `OfficeIMO.Ocr` and `OfficeIMO.Pdf`.
- Not dependencies: Reader, Tesseract, process execution, cloud SDKs, or native OCR runtimes.
- License: MIT.

See the [OfficeIMO.Pdf README](../OfficeIMO.Pdf/README.md) for native reading and document operations.
