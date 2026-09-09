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
        MaxConcurrentPages = 2,
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

## Reconstruct columns and mixed-direction text

Provider line hierarchy and logical sequence are preserved by default. If a provider joins two columns into one line or supplies an unsuitable page order, enable geometry-based reconstruction:

```csharp
PdfOcrMergeResult reconstructed = await pdf.ReadWithOcrAsync(engine,
    new PdfOcrMergeOptions {
        ReconstructLayout = true,
        Language = "heb+ara+eng",
        ReadOptions = new PdfReadOptions {
            LayoutOptions = new PdfTextLayoutOptions {
                ReadingDirection = PdfReadingDirection.RightToLeft
            }
        }
    });
```

The shared PDF stages rebuild OCR lines, columns, and aligned tables from accepted words. Mixed-direction fragments use the same logical-order resolver as native text. Dominant quarter-turn layouts are analyzed in a corrected reading frame; returned words, selection rectangles, and table bounds use the original page geometry. Explicit direction is useful for ambiguous pages; `Auto` remains the default.

Reconstruction does not repair recognition errors or infer a figure from its caption vocabulary. A full-page scan can retain caption text without exposing a separate figure region or classified caption. Searchable output preserves the visible scan and supports subsequent line and table reconstruction from its invisible selection boxes. See the [independent layout corpus](../OfficeIMO.TestAssets/MultilingualLayout/README.md) for measured coverage and limits.

## Prepare uneven or rotated scans

Scan cleanup is opt-in and uses the shared `OfficeIMO.Core` image processor. It changes the raster sent to OCR; the source PDF and its visible scans are preserved.

```csharp
using OfficeIMO.Drawing;

PdfSearchableOcrReview review = await pdf.PrepareSearchableOcrAsync(engine,
    new PdfOcrMergeOptions {
        Dpi = 300,
        DetectOrientation = true,
        MinimumOrientationConfidence = 0.75,
        ScanProcessing = new OfficeScanProcessingOptions {
            Deskew = true,
            NormalizeBackground = true,
            ColorMode = OfficeScanColorMode.Grayscale,
            MaximumDimension = 3000,
            MaximumWorkingBytes = 256L * 1024 * 1024
        }
    });

foreach (PdfOcrPageMergeResult page in review.Ocr.Pages) {
    Console.WriteLine($"Page {page.PageNumber}: deskew {page.ScanProcessing?.AppliedDeskewDegrees}");
}
PdfSearchableOcrResult searchable = review.ApplyAll();
```

Orientation detection uses the provider's optional orientation capability and the same timeout, cancellation, and concurrency gate as recognition. Missing or low-confidence evidence retains the source orientation and produces a diagnostic. Tesseract needs its `osd` trained data. An explicit `ClockwiseQuarterTurns` value can supply a caller-reviewed correction; it combines with any accepted provider correction.

Deskew searches a bounded range of small angles; `StraightenDegrees` supplies a manual correction from -15 to 15 degrees. Background normalization estimates local paper brightness. `BlackPoint`, `WhitePoint`, and `Gamma` adjust tonal levels before optional bilevel conversion. Downsampling never enlarges a scan. Blank-page detection reports a suggestion and keeps the page. Curved-page dewarping remains unsupported.

`ScanProcessing` reports applied and skipped operations, buffer estimates, and forward/inverse pixel transforms. Its pixel, buffer, and analysis-work limits reject optional cleanup with an `ocr-scan-limit` diagnostic and retain the original OCR raster; cancellation still propagates. Buffer accounting covers the managed image operation, while encoded PDF/raster and provider-process limits remain separate. `PdfRecognizedWord.Geometry` retains all four corners on the original page, so the invisible text layer follows the original scan's angle after deskew or a quarter-turn. `X`, `Y`, `Width`, and `Height` remain its enclosing visual bounds.

### Review a region and perspective correction

`Regions` accepts one normalized rectangle per page. A nonempty list sends only those pages and pixels to the OCR provider. Region pages must also belong to `ReadOptions.PageSelection` when supplied. `Perspective` specifies four normalized corners relative to the region, or to the full page when no region is selected. Preparation crops first, corrects perspective next, and applies affine scan cleanup last.

```csharp
var options = new PdfOcrMergeOptions {
    Regions = new[] { new PdfOcrPageRegion(1, 0.1, 0.1, 0.8, 0.7) },
    Perspective = new OfficeScanPerspectiveOptions {
        TopLeft = new(0.02, 0.04), TopRight = new(0.98, 0),
        BottomRight = new(1, 1), BottomLeft = new(0, 0.96)
    },
    ScanProcessing = new OfficeScanProcessingOptions {
        Deskew = false, StraightenDegrees = 2, Gamma = 1.1
    }
};
PdfScanPreview preview = await document.PreviewScanAsync(1, options);
byte[] originalPreview = preview.GetSourcePng();
byte[] preparedPreview = preview.GetPreparedPng();
PdfSearchableOcrReview review = await document.PrepareSearchableOcrAsync(engine, options);
```

Preview uses the same preparation code without calling an OCR provider. OCR geometry maps back through all transforms to the original visible page. `preview.CreateImagePdf()` instead creates a separate raster-only PDF of the prepared pixels: native text, forms, links, signatures, and attachments are omitted. Invalid region or perspective settings stop preparation; they do not silently select a different area. `ScanProcessing` describes the affine cleanup relative to the prepared region; its matrix alone does not describe the earlier crop and perspective mapping.

## Discover scanned redaction candidates

Use the same OCR geometry and native-overlap owner to map literal or bounded-regex matches into PDF user-space areas:

```csharp
var search = new PdfRedactionSearchOptions()
    .AddLiteral("Account Secret")
    .AddRegex(@"\b\d{3}-\d{2}-\d{4}\b");

PdfOcrRedactionSearchResult candidates = await pdf
    .SearchRedactionCandidatesWithOcrAsync(engine, search);

foreach (PdfOcrRedactionCandidate candidate in candidates.Candidates) {
    Console.WriteLine($"Page {candidate.Area.PageNumber}: {candidate.Criterion}, confidence {candidate.MinimumConfidence:0.00}");
}
```

Candidate results intentionally omit recognized matched text. They retain the criterion index, geometry, minimum confidence, and provider/model/language evidence needed by a review workflow. Literal and regex search is isolated to provider-declared lines, with bounded geometric line inference only when hierarchy identifiers are unavailable, so candidates are not assembled across unrelated lines or columns. `OfficeIMO.Workflows` can combine these candidates with native matches, persist source-bound decisions, re-run the same provider after destructive application, and publish privacy-safe evidence.

## Add a searchable text layer

```csharp
PdfSearchableOcrResult searchable = await pdf.MakeSearchableAsync(engine);
await searchable.Document.SaveAsync("mixed-report-searchable.pdf");

Console.WriteLine($"Modified pages: {string.Join(", ", searchable.ModifiedPages)}");
Console.WriteLine($"Added words: {searchable.AddedWordCount}");
```

Only pages with accepted OCR words are rewritten. The invisible text layer follows the canonical semantic order. `WrittenWords` records what entered the layer, while `Ocr` retains recognition evidence. Signed or otherwise rewrite-sensitive documents remain subject to the base PDF mutation and preservation rules.

## Review before creating the layer

`PrepareSearchableOcrAsync` captures the source and recognizes its selected pages without changing or saving the PDF. A review interface can display `Ocr.Pages`, including `WordEvidence` for accepted words, low-confidence words, and native-text overlaps. `RenderPage` previews the same source snapshot; `GetPageSize` and word geometry use cropped, rotated visual PDF points.

```csharp
PdfSearchableOcrReview review = await pdf.PrepareSearchableOcrAsync(engine);

// Replace this confidence selection with the eligible word instances chosen in a review interface.
var selected = review.Ocr.Pages.SelectMany(page => page.Words)
    .Where(word => word.Confidence >= 0.90).ToArray();
PdfSearchableOcrResult reviewed = review.Apply(selected);
await reviewed.Document.SaveAsync("reviewed-searchable.pdf");
```

Selections may exclude eligible words but cannot inject words from another review or override a rejection. To change the confidence or overlap policy, prepare a new review with new options. Low-confidence words are rejected before overlap evaluation; invalid geometry remains a diagnostic rather than a selectable word. An empty selection produces an unchanged source copy. `AddedWordCount` and `WrittenWords` describe the actual layer after review exclusions.

`PdfOcrMergeOptions` bounds provider-call duration, rendered pixels, selected pages, inspected spans, accepted OCR words and characters, aggregate raw hierarchy identifiers, provider metadata and diagnostics, native-overlap comparisons, and merged text. Calls use one shared `OcrEngineExecution` per document, so identity and capabilities are stable across pages and the same non-concurrent engine instance cannot overlap across PDF, Reader, or a future integration. Language is provider configuration only; it is never used to infer captions, lists, paragraphs, tables, or continuations.

Use `ApplyCorrections` to correct recognized text after reviewing the page. Include only the eligible words to write, paired with their final text:

```csharp
var corrections = review.Ocr.Pages.SelectMany(page => page.Words)
    .ToDictionary(word => word, word => word.Text);
PdfRecognizedWord selectedWord = review.Ocr.Pages[0].Words[0];
corrections[selectedWord] = "Corrected text";
PdfSearchableOcrResult corrected = review.ApplyCorrections(corrections);
await corrected.Document.SaveAsync("corrected-searchable.pdf");
```

Corrections preserve the selected word's geometry and reading order. `WrittenWords` contains the replacement text, `CorrectedWordCount` counts changed words, and `Ocr` retains the original provider text and confidence. Replacement text must be nonempty and fit the per-page OCR character budget.

## Scan rendering and execution limits

CCITT Group 3 and Group 4 scans use the managed decoder. Packed 1-, 2-, and 4-bit DeviceGray samples pass through the existing decode-array, color, and mask handling. Fax decoding requires a declared row count or image height; uncompressed fax extension mode and damaged-row recovery are outside the supported contract.

Opaque JPEG 2000 images with baseline Gray/sRGB headers or one/three-component codestreams can use `PdfOcrMergeOptions.ImageCodec`, the shared `IOfficeRasterImageCodec` interface. The same codec is used by review previews. A missing decoder or an unprojectable scan causes rendering to fail before that page is sent to OCR. JPEG 2000 embedded or external masks, palette/channel remapping, alternate color spaces, and output-intent normalization remain unsupported. Embedded alpha is rejected even when `SMaskInData` is absent or zero, because those PDF cases require discarding that alpha before rendering. No JPEG 2000 runtime is bundled.

`Pages[i].Diagnostics` includes render warnings as well as provider and normalization diagnostics. Inspect these before treating a result as complete: font substitution and unsupported drawing features can affect recognition even when a page renders.

`MaxConcurrentPages` defaults to one. Raise it to overlap page requests for providers that declare concurrent-request support. Non-concurrent providers remain serialized, and result pages retain the requested order. Parsing and rendering use one producer; only a bounded number of provider requests are retained. `MaxRenderedBytesPerPage` defaults to 64 MiB and limits each encoded PNG. Rendered pages are released as requests complete rather than accumulated for the whole document.

## Targets and dependency footprint

- Targets: `netstandard2.0`, `net8.0`, `net10.0` (`net472` is also included on Windows builds).
- OfficeIMO dependencies: `OfficeIMO.Core`, `OfficeIMO.Ocr`, and `OfficeIMO.Pdf`.
- Not dependencies: Reader, Tesseract, process execution, cloud SDKs, or native OCR runtimes.
- License: MIT.

See the [OfficeIMO.Pdf README](../OfficeIMO.Pdf/README.md) for native reading and document operations.
