# OfficeIMO image export

OfficeIMO image export is a first-party, dependency-free pipeline shared by Drawing, Excel, Word, PowerPoint, HTML, email, EPUB, OneNote, Visio, PDF, and OpenDocument adapters.

## Ownership

`OfficeIMO.Core` owns the reusable contract, exposed through the `OfficeIMO.Drawing` namespace:

- PNG, JPEG, TIFF, SVG, and WebP encoding;
- validated `OfficeImageExportResult` metadata;
- target DPI and raster density metadata;
- caller-supplied TrueType fonts and shared substitution diagnostics;
- source-image decoding and caller-codec fallback;
- per-image and aggregate safety limits;
- diagnostic acceptance policy;
- streaming batches, cancellation, progress, bounded concurrency, filenames, and save conflicts.

Document packages own selection, pagination, layout, and source semantics. ODT/ODS/ODP reuse Word/Excel/PowerPoint; EPUB and email reuse HTML; paged text adapters reuse PDF. An adapter must not introduce another encoder, font resolver, batch engine, or visual layout brain.

## Dependency Rule

Product rendering paths may use the libraries already owned by their document package and `OfficeIMO.Core`. They must not add Office automation, browser screenshots, native PDF rasterizers, System.Drawing, Skia, ImageSharp, commercial renderers, or another output encoder.

External renderers are allowed only as optional test references. They are not part of product output.

## Typical Use

```csharp
OfficeConversionReport preview = presentation.Slides[0].SaveAsImage("preview.png",
    new PowerPointImageExportOptions {
        TargetDpi = 144,
        FailOnLoss = true,
        FileConflictPolicy = OfficeImageExportFileConflictPolicy.CreateUnique,
        Format = OfficeImageExportFormat.Png
    });

Console.WriteLine(preview.OutputPath);
```

For a production batch, save without retaining earlier encoded image payloads:

```csharp
using var cancellation = new CancellationTokenSource();
OfficeConversionReport saved = await document.SaveAsImagesAsync("pages",
    new ExcelImageExportOptions {
        TargetDpi = 300,
        MaximumOutputCount = 500,
        MaximumTotalRasterPixels = 250_000_000,
        MaximumTotalEncodedBytes = 512L * 1024 * 1024,
        MaximumDegreeOfParallelism = 4,
        Progress = new Progress<OfficeImageExportProgress>(progress =>
            Console.WriteLine($"{progress.Stage}: {progress.CompletedCount}")),
        FileConflictPolicy = OfficeImageExportFileConflictPolicy.FailIfExists,
        FailOnLoss = true
    }, cancellation.Token);
```

Add TrueType faces to `options.Fonts` when typography must not depend on installed fonts. A missing requested face reports `font-family-substitution`; `FailOnLoss` rejects the diagnosed approximation. Word uses `WordImageExportOptions` from the `OfficeIMO.Word.Pdf` package, Excel uses `ExcelImageExportOptions`, and PowerPoint uses `PowerPointImageExportOptions`. Tier 2 adapters retain their source-specific selection and rendering APIs.

## Source Bridges

The direct bridges retain the source conversion evidence:

```csharp
IReadOnlyList<OfficeImageExportResult> slides =
    odp.ExportImages(OfficeImageExportFormat.Png);

OfficeConversionReport chapters = await epub.SaveAsImagesAsync("chapters",
    new EpubImageExportOptions {
        Mode = HtmlRenderMode.Paged,
        TargetDpi = 144,
        Format = OfficeImageExportFormat.Png
    });

OfficeConversionResult<byte[]> message = await email.ToImageAsync(
    new EmailImageExportOptions {
        Mode = HtmlRenderMode.Continuous,
        Format = OfficeImageExportFormat.Png
    });
```

Load EPUB with `IncludeRawHtml = true` and `IncludeResourceData = true` for the best result. Email resolves allowed CID/content-location MIME resources through the HTML resource pipeline. ODT/ODS/ODP diagnostics are attached to the resulting Word/Excel/PowerPoint image diagnostics.

## Encoder and SVG output

PNG and lossless WebP default to `OfficeImageEncodingEffort.Balanced`. Set `options.RasterEncoding.Png.Effort` or `.Webp.Effort` to `Fast` for lower encode cost or `Smallest` for exhaustive candidate searches. All levels retain exact pixels. `PdfPageImageExportRequest.RasterEncoding` passes the same settings through Workflows and snapshots them before execution.

PNG optimal encoding selects exact grayscale, grayscale-alpha, RGB, RGBA, or indexed samples and writes sRGB color information. Use the metadata-aware raster encoder for an ICC profile and native density. JPEG defaults to optimized Huffman tables at quality 85 and 4:4:4; progressive and 4:2:0 remain explicit settings. TIFF adds `OfficeTiffCompression.CcittGroup4` for exact opaque black/white sources. WebP streams share the bounded lossless compressor instead of always writing literals. The [Core README](../OfficeIMO.Core/README.md) describes these APIs and their resource limits.

SVG embeds non-interpolated rasters as `<image>` with pixelated sampling hints instead of expanding each pixel into rectangles. Output pixel dimensions follow the raster rounding rule. Soft-mask transfer tables and text outline strokes retain their Drawing paint semantics.

Set `OfficeImageExportOptions.SvgTextAsPaths = true` to export eligible single-line horizontal ASCII text through registered font outlines. Already-painted runs use that path when supported. The SVG keeps logical text as an accessibility label and omits fonts no longer used by `<text>`. Wrapped, decorated, vertical, mixed-fallback, complex-script and color-glyph configurations retain text and the required font. This mode does not subset retained fonts and is not an editable-text export.

## Current Limits

- Word pagination is estimated rather than Microsoft Word-exact.
- PDF rendering is strongest for OfficeIMO-generated PDFs; arbitrary producer features remain capability-diagnosed.
- Static image export does not preserve animation or multi-frame input.
- Complex TIFF/WebP/SVG input variants use a caller codec or a diagnosed visible fallback.
- Exact cross-machine typography requires caller-supplied font data.
- Raster transcoding carries supported ICC and EXIF profiles through `OfficeRasterImageEncoder.EncodeWithMetadata`; this does not imply automatic color conversion or source metadata preservation by every document adapter.

Open fidelity and proof work is tracked in [ROADMAP.md](ROADMAP.md). Additional output formats and package-local copies of the shared engine are outside the current contract.
