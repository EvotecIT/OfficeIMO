# OfficeIMO Image Export

OfficeIMO image export is a first-party, dependency-free pipeline shared by Drawing, Excel, Word, PowerPoint, HTML, email, EPUB, OneNote, Visio, PDF, and OpenDocument adapters.

## Ownership

`OfficeIMO.Drawing` owns the reusable contract:

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

Product rendering paths may use the libraries already owned by their document package and `OfficeIMO.Drawing`. They must not add Office automation, browser screenshots, native PDF rasterizers, System.Drawing, Skia, ImageSharp, commercial renderers, or another output encoder.

External renderers are allowed only as optional test references. They are not part of product output.

## Typical Use

```csharp
OfficeImageExportResult preview = presentation.Slides[0]
    .ToImage()
    .AtDpi(144)
    .WithPolicy(policy => policy.RequireNoOmissions = true)
    .OnFileConflict(OfficeImageExportFileConflictPolicy.CreateUnique)
    .AsPng()
    .Save("preview");

Console.WriteLine(preview.SavedPath);
Console.WriteLine($"{preview.Width}x{preview.Height} at {preview.DpiX:0.#} DPI");
```

For a production batch, stream or save payload-free metadata instead of retaining every encoded image:

```csharp
using var cancellation = new CancellationTokenSource();

OfficeImageExportBatchSaveResult saved = await document
    .ToImages()
    .ForPrint(300)
    .WithBatchLimits(
        maximumOutputCount: 500,
        maximumTotalRasterPixels: 250_000_000,
        maximumTotalEncodedBytes: 512L * 1024 * 1024)
    .WithMaximumConcurrency(4)
    .WithProgress(new Progress<OfficeImageExportProgress>(progress =>
        Console.WriteLine($"{progress.Stage}: {progress.CompletedCount}")))
    .OnFileConflict(OfficeImageExportFileConflictPolicy.FailIfExists)
    .SaveFilesAsync("pages", cancellation.Token);

saved.Report.Require(new OfficeImageExportPolicy {
    RequireNoFailures = true,
    RequireNoOmissions = true
});
```

Use `WithFont(...)` or `WithFonts(...)` when typography must not depend on the machine's installed fonts. A missing requested face produces `IMAGE_FONT_SUBSTITUTED`; it can be promoted to a hard failure through `FailOnDiagnosticCodes`.

## Source Bridges

The direct bridges retain the source conversion evidence:

```csharp
IReadOnlyList<OfficeImageExportResult> slides =
    odp.ExportImages(OfficeImageExportFormat.Png);

OfficeImageExportBatchSaveResult chapters = await epub
    .ToImages()
    .Paged()
    .AtDpi(144)
    .AsPng()
    .SaveFilesAsync("chapters");

OfficeImageExportResult message = await email
    .ToImage()
    .Continuous()
    .AsPng()
    .ExportAsync();
```

Load EPUB with `IncludeRawHtml = true` and `IncludeResourceData = true` for the best result. Email resolves allowed CID/content-location MIME resources through the HTML resource pipeline. ODT/ODS/ODP diagnostics are attached to the resulting Word/Excel/PowerPoint image diagnostics.

## Current Limits

- Word pagination is estimated rather than Microsoft Word-exact.
- PDF rendering is strongest for OfficeIMO-generated PDFs; arbitrary producer features remain capability-diagnosed.
- Static image export does not preserve animation or multi-frame input.
- Complex TIFF/WebP/SVG input variants use a caller codec or a diagnosed visible fallback.
- Exact cross-machine typography requires caller-supplied font data.
- ICC workflows, EXIF preservation, and CMYK color conversion are not part of the current contract.

Further work should improve fidelity and proof for these existing surfaces. It should not add more output formats or package-local copies of the shared engine.
