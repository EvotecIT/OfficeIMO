# OfficeIMO.Core - shared document, security, data, and drawing primitives

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Core)](https://www.nuget.org/packages/OfficeIMO.Core)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Core?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Core)

`OfficeIMO.Core` is the zero-dependency shared foundation for OfficeIMO packages. It owns common document lifecycle and package-safety contracts, neutral tabular row mapping, color conversion, image metadata, font and text measurement, vector scenes, reusable ink and math models, chart snapshots, SVG, raster canvases, PNG/JPEG/TIFF/WebP encoding, and drawing-quality primitives. Format packages keep their file-format behavior while reusing these document-agnostic models and renderers.

The assembly was previously named `OfficeIMO.Drawing`. Drawing became the original shared foundation because keeping these primitives together avoided a separate Core → Drawing → format dependency chain. As lifecycle, security, package, and data contracts accumulated, the package name stopped describing its actual responsibility. That same single zero-dependency foundation is now named `OfficeIMO.Core`; it was not split into another runtime dependency. Actual drawing APIs remain in `OfficeIMO.Drawing`, neutral data and flattening contracts use `OfficeIMO.Data`, security APIs remain in `OfficeIMO.Security`, and cross-document lifecycle, compatibility, capability, and conversion-report contracts use the root `OfficeIMO` namespace.

The managed VP8 decoder is maintained in `OfficeIMO.Core` under OfficeIMO's MIT license. It was ported from the same author's CodeGlyphX implementation. The lossy VP8 encoder is adapted from CodeGlyphX under Apache-2.0 and retains its source notices. Both implementations run inside Core without a CodeGlyphX runtime dependency. The [third-party notices](THIRD-PARTY-NOTICES.md) identify the incorporated code and its license files.

For bounded image sequences, filtering, text drawing, comparison, and visual fingerprints, see [managed raster workflows](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.core-raster-workflows.md). The examples use owned image types and document stream ownership, orientation, and cancellation.

## Install

```powershell
dotnet add package OfficeIMO.Core
```

## Mathematical drawing

`OfficeMathRenderer` renders the owned equation model using caller-supplied fonts.
Register a mathematical face in `OfficeMathRenderOptions.Fonts` and select its family
through `Font`. OpenType MATH constants, glyph variants and assemblies determine
fraction rules, script spacing, stretched delimiters, radicals and accents when
available. Fonts without those tables use the existing geometric fallback.

`UseFontMathMetrics` selects font-derived layout, including available stretched
glyph variants and assemblies. Disable it explicitly when the application requires
the geometric fallback. Rendering cancellation applies during
font measurement and equation construction. These options do not execute scripts or
require an external rendering engine.

## Per-column row mapping

`RowMapper<T>` in `OfficeIMO.Data` provides explicit assignments for any
`DbDataReader`. CSV and Excel use this shared mapper. `FromColumns` tries aliases
in the supplied order; required missing columns and ambiguous duplicate matches
fail before row projection.

```csharp
using OfficeIMO.Data;
using System.Data.Common;
using System.Globalization;

static IEnumerable<Invoice> ReadInvoices(DbDataReader reader) =>
    reader.RowsAs<Invoice>(map => map
        .FromColumns<int>(new[] { "Id", "InvoiceId" },
            (row, value) => { row.Id = value; return row; })
        .FromColumn<decimal>("Amount",
            (row, value) => { row.Amount = value; return row; },
            new RowMappingColumnOptions { Culture = CultureInfo.GetCultureInfo("fr-FR") })
        .FromColumn<DateTime>("Date",
            (row, value) => { row.Date = value; return row; },
            new RowMappingColumnOptions { DateTimeFormats = new[] { "yyyyMMdd" } })
        .FromColumn<string>("Note",
            (row, value) => { row.Note = value; return row; },
            new RowMappingColumnOptions { Optional = true }));

public sealed class Invoice {
    public int Id { get; set; }
    public decimal Amount { get; set; }
    public DateTime Date { get; set; }
    public string? Note { get; set; } = "No note";
}
```

Optional means the source column may be absent. Its assignment is skipped,
preserving the model's initialized value. Present empty or invalid values still
follow the normal conversion and nullability rules. Even when every configured
column is optional and absent, each source row produces a model.

The mapper snapshots each binding's culture and date-format list when it is
configured. Per-column date/time formats replace the reader's format list and
are tried before ordinary culture-based parsing. A per-column `TypeConverter`
replaces the reader's converter for that binding. Returning `(false, null)`
selects built-in conversion; handled results, including null, take precedence.
The converter receives the value exposed by the reader before built-in
conversion. Excel numeric mappings retain the original date serial when a
converter declines.

The same controls apply to `RowsAs`, `RowsAsAsync` and `RowsAsParallel`.
Parallel converters must support concurrent calls. Source-value redaction
continues to follow the reader's mapping-error policy.

## Image export density

`OfficeImageExportOptions.UseQuality(...)` and fluent `WithQuality(...)` select shared density presets: `Preview` is 96 DPI, `Screen` is 192 DPI, and `Print` is 300 DPI. They retain the selected fonts, layout, format, and safety limits. Clear `TargetDpi` when setting `Scale` directly; fluent `WithScale(...)` clears it automatically. Each document adapter defines its logical units per inch.

`OfficeDrawing.ExportImage(format, options)` exports a detached drawing through the same raster limits, density metadata, codecs, deadline, and diagnostic policy. Raster output is rendered at the requested density. SVG retains vector geometry and text; it cannot add detail to embedded raster images. Register regular and bold font faces for consistent measurement and output across machines.

Interactive `OfficeDrawing` links accept HTTP(S), mail, telephone, and local relative or fragment targets. SVG import keeps the painted content but omits links with executable, data, file, or other unsupported schemes; direct `AddLink` calls reject those targets. SVG export retains accepted links as interactive anchors.

Raster strokes preserve fractional widths, caps, joins, miter limits, and dash phase. The renderer paints
overlapping pieces of one stroke together, so an extra point or intersecting subpath does not darken
translucent ink. Affine transforms apply to the stroke outline. Curve detail follows output density,
with a target error of 0.05 device pixels and a bounded segment count. Thin contour details contribute
their area even when they fall between sample rows. Image reduction averages premultiplied colors
before bilinear placement; explicit nearest-neighbor drawing retains its pixel-art behavior.

`OfficeFontInfo.Face` carries numeric weight, stretch, and slant through measurement, raster painting,
and SVG export. Register the faces a document needs and request the descriptor directly:

```csharp
var drawing = new OfficeDrawing(240, 80);
drawing.Fonts.Add("Inter", File.ReadAllBytes("Inter-Regular.ttf"), new OfficeFontFaceDescriptor(400));
drawing.Fonts.Add("Inter", File.ReadAllBytes("Inter-Medium.ttf"), new OfficeFontFaceDescriptor(500));
drawing.AddText("Monthly report", 12, 12, 216, 56,
    new OfficeFontInfo("Inter", 20, new OfficeFontFaceDescriptor(500)));
```

Face matching follows CSS ordering within the requested family. Missing faces use the available
matching face and existing fallback policy. Simulated bold paints its combined outline once,
preserving the requested opacity.

## Prepare scanned images

`OfficeScanProcessor` in `OfficeIMO.Drawing` prepares a separately owned raster for OCR. It supports explicit quarter-turns, manual straightening, confidence-filtered deskew, local paper-brightness normalization, black and white levels, gamma, grayscale or bilevel output, and proportional downsampling:

```csharp
using OfficeIMO.Drawing;

OfficeScanProcessingResult processed = OfficeScanProcessor.Process(sourceImage,
    new OfficeScanProcessingOptions {
        Deskew = true,
        NormalizeBackground = true,
        ColorMode = OfficeScanColorMode.Grayscale,
        MaximumDimension = 3000
    }, cancellationToken);
OfficePoint sourcePoint = processed.Report.ProcessedToSource.TransformPoint(ocrPixelPoint);
```

The source image stays unchanged. The report records transformations, skipped decisions, estimated managed buffers, and a blank-page suggestion; pages are never removed. Pixel, buffer, and analysis-work limits throw `OfficeScanProcessingLimitException`, allowing the caller to retain the original. The operation does not detect quarter-turn orientation itself; an OCR provider or the caller supplies that evidence.

PNG encoding with `OfficePngCompression.Optimal` stores fully opaque black-and-white rasters
as one-bit grayscale images. This reduces the encoded scan payload while preserving every
pixel, dimensions, and requested resolution metadata. Other opaque rasters use eight-bit
RGB samples, avoiding an alpha channel that is uniformly opaque. Transparent rasters
retain eight-bit RGBA samples. The encoder preserves every channel value without
thresholding or quantization.
Byte, stream, and buffer-writer APIs use the same selection. `Stored` compression retains
eight-bit RGBA samples.

`OfficeScanProcessor.CorrectPerspective(image, options)` creates a rectangular raster from four normalized source corners. `OfficeScanPerspectiveOptions` requires a convex, clockwise quadrilateral inside the source image. The result includes a projective mapping in both directions so consumers can place recognized text back on the original. Perspective correction is separate from the affine transform reported by ordinary scan cleanup. Curved-page dewarping remains unsupported.

## Quick start

### Document lifecycle policy

```csharp
using OfficeIMO;

var loadOptions = new DocumentLoadOptions {
    AccessMode = DocumentAccessMode.ReadOnly,
    PersistenceMode = DocumentPersistenceMode.Explicit
};
```

Word, Excel, and PowerPoint expose format-specific options derived from these shared contracts.

### Shared operation and conversion results

OfficeIMO format packages keep their typed diagnostics while sharing a small result contract from
`OfficeIMO.Core`:

```csharp
using OfficeIMO;

IOfficeResult operation = result;
if (!operation.Succeeded) {
    // Inspect the concrete result for its typed failure details.
}

IOfficeResult<MyDocument> documentResult = result;
MyDocument document = documentResult.RequireValue();

IOfficeConversionResult<MyDocument, MyConversionReport> conversion = result;
MyDocument losslessDocument = conversion.RequireNoLoss();
```

`IOfficeResult<T>` standardizes `Succeeded`, `Value`, and `RequireValue()`.
`IOfficeConversionResult<TValue, TReport>` also exposes `Report`, `HasLoss`, and
`RequireNoLoss()`. Concrete results continue to expose format-specific diagnostics, warnings,
or exceptions; the shared interface does not reduce them to an untyped error string.

### Portable AES for restricted hosts

Browser WebAssembly and other hosts without synchronous platform AES can explicitly supply the dependency-free managed
AES-CBC provider to a format API that accepts `IOfficeAesCryptographyProvider`. Desktop and server applications should
normally keep using their native platform AES implementation.

```csharp
using OfficeIMO.Security;

IOfficeAesCryptographyProvider aes = OfficeManagedAesCryptographyProvider.Default;
```

### Colors and vector intent

```csharp
using OfficeIMO.Drawing;

OfficeColor accent = OfficeColor.Parse("#336699");
OfficeColor printBlue = OfficeColorSpaceConverter.FromCmyk(1, 0.45, 0, 0.15);
OfficeImageFit fit = OfficeImageFit.Contain;

var badge = OfficeShape.RoundedRectangle(120, 32, 8);
badge.FillColor = OfficeColor.WhiteSmoke;
badge.StrokeColor = accent;
badge.Shadow = new OfficeShadow(OfficeColor.Black, 0.18, 3, 4);
```

Use `OfficeIccColorProfile` when an application already has embedded ICC profile bytes and needs the
same bounded, dependency-free conversion used by OfficeIMO renderers:

```csharp
using OfficeIMO.Drawing;

byte[] profileBytes = File.ReadAllBytes("display.icc");
if (OfficeIccColorProfile.TryCreate(profileBytes, out OfficeIccColorProfile? profile) &&
    profile.TryConvert(new[] { 0.25D, 0.5D, 0.75D }, out OfficeColor converted)) {
    Console.WriteLine(converted.ToHex());
}

if (profile?.HasOutputTransform == true &&
    profile.TryConvertToDevice(OfficeColor.CornflowerBlue, OfficeIccRenderingIntent.RelativeColorimetric, out double[] deviceColor) &&
    profile.TrySoftProof(OfficeColor.CornflowerBlue, OfficeIccRenderingIntent.RelativeColorimetric, out OfficeColor proofed)) {
    Console.WriteLine($"Device channels: {string.Join(", ", deviceColor)}; proof: {proofed.ToHex()}");
}
```

The managed contract accepts bounded RGB and Gray matrix/TRC input-device and display-device profiles
plus RGB, CMYK, or three-to-eight-channel (`3CLR`–`8CLR`) LUT8 input transforms with a Lab profile
connection space and LUT16 input transforms with an XYZ or Lab profile connection space. For LUT and ICC v4 `mAB` profiles,
conversion selects `A2B1` for relative or absolute colorimetric intent and `A2B2` for saturation intent,
falling back to `A2B0` when the intent-specific transform is absent. It also accepts bounded ICC v4
RGB, CMYK, and `3CLR`–`8CLR` A2B `mAB` input transforms, plus RGB/CMYK B2A `mBA` output transforms using the
specification-defined curve, variable-grid CLUT, matrix, and offset combinations. Multichannel profiles support input conversion only. Output conversion is
available through a valid `B2A0` transform or
the synthesized inverse of a supported RGB matrix/TRC profile; optional intent-specific tags fall back
to `B2A0` when that transform is present. `TryCreate` returns `false` for unsupported
input profile classes and transform types, while malformed or unsupported optional output transforms
leave `HasOutputTransform` false so the caller can choose an explicit color-management provider or
fallback instead of receiving a silent approximation.

### Image metadata and complete-content validation

`OfficeImageReader` exposes `JpegComponentCount` and
`TiffPhotometricInterpretation` on `OfficeImageInfo` as nullable header evidence.
The JPEG value is the frame's declared component count. The TIFF value is the raw
PhotometricInterpretation tag from the first classic-TIFF or BigTIFF directory;
missing, malformed or duplicate tags produce null. TIFF value 2 declares RGB and
5 declares separated samples. Three JPEG components alone do not prove RGB.
Neither field establishes valid pixel data, an ICC profile or color-managed
rendering. Manually constructed metadata and other image formats leave these
fields null. Use the decoder and explicit profile APIs for their separate checks.

`OfficeRasterImageDecoder` preserves encoded color channels and applies supported image orientation.
It does not automatically normalize embedded ICC profiles, PNG gamma, or chromaticities. Applications
that require color-managed pixels must perform that conversion explicitly before using the decoded
image. The profile APIs above provide bounded color conversion; metadata validation alone does not
mean a raster image has been converted to sRGB.

For already unpacked, tightly packed 8-bit RGB, CMYK, or supported multichannel device samples, use the explicit raster
converter with the corresponding embedded profile bytes:

```csharp
OfficeIccRasterConversionStatus status = OfficeIccRasterConverter.TryConvertToSrgb(
    deviceSamples, width, height, profileBytes,
    new OfficeIccRasterConversionOptions {
        MaximumProfileBytes = 1024 * 1024,
        MaximumPixels = 1_000_000
    },
    out OfficeRasterImage? srgbImage);
if (status != OfficeIccRasterConversionStatus.Converted) {
    throw new InvalidDataException($"ICC raster conversion rejected the input: {status}");
}
```

The converter uses the shared ICC matrix/TRC and LUT engine. It rejects malformed or unsupported
profiles and incomplete sample buffers, and accounts for the source, profile parser, and RGBA output
before allocating pixels. The default ceilings are 4 MiB of ICC data, 4 million pixels, and 256 MiB
of accounted managed memory; callers can lower them. The result is opaque sRGB pixels. Alpha,
planar samples, higher bit depths, and extraction of device channels from encoded images are separate
format concerns. The [independent color corpus](../OfficeIMO.Drawing.Tests/TestAssets/IccColorCorpus/SOURCE.md)
checks matrix RGB, ICC v4 LUT RGB, and CMYK LUT swatches against LittleCMS.

`OfficeImageOptimizer` still preserves or reports ICC metadata according to its metadata policy.
It does not call the raster converter or claim that re-encoded pixels were normalized to sRGB.

Routed SVG filters on non-text shape/group content support `SourceGraphic`,
`SourceAlpha`, preceding named results, Gaussian blur, offset, matrix color transforms,
source-over composition and separable blend modes. These operations use managed RGBA
buffers at one pixel per drawing unit, then retain a PNG in the scene. The default
filter color space is linear RGB; a uniform `sRGB` declaration is also supported.
The filter region clips every input/result and the final paint. Simple unrouted
blur/offset/drop-shadow filters retain their existing vector approximation.

Managed filter graphs accept at most 32 primitives, blur deviation up to 64 drawing
pixels per axis, and cumulative document work/intermediate-surface limits. Rotated
or sheared graphs, explicit primitive subregions, mixed filter color spaces,
unsupported primitives and graphs containing text, logical `ActualText`, links or
pattern cells report unsupported features and preserve source geometry. Links on
the filtered container retain their original geometry. Pass
`OfficeSvgDrawingReaderOptions.CancellationToken` to cancel import and filter work.

The SVG drawing reader supports a single rectangle, rounded rectangle, circle, ellipse, polygon, or
path inside a `userSpaceOnUse` clip path, including transforms and even-odd filling. Compound clip
unions, `objectBoundingBox` clips, and referenced or text clip geometry report unsupported features.
Imported SVG text retains its full glyph paint inside an explicit root viewport
clip. Text measurements determine layout, while the viewport and authored clip
paths determine visible paint. Scene inspection should traverse drawing groups
and effect groups to reach their text and shape content.
Shape geometry crossing a nested SVG or symbol viewBox is retained until the viewport clip is applied.
Local symbols without a `viewBox` retain their user coordinates and inherit paint from
their `use` element. Symbol dimensions clip the content; they do not rescale it.
Native import resolves omitted dimensions from the containing viewport. Caller-raster
safety checks require explicit symbol width and height in documents with nested
viewports or symbol references whose viewport context cannot be resolved by that check.

```csharp
using OfficeIMO.Drawing;

OfficeImageInfo info = OfficeImageReader.Identify("logo.png");
Console.WriteLine($"{info.Width}x{info.Height} {info.MimeType}");

// Use content verification when an extension must not identify invalid bytes.
byte[] bytes = File.ReadAllBytes("upload.svg");
bool verified = OfficeImageReader.TryIdentifyByContent(bytes, "upload.svg", out OfficeImageInfo upload);

OfficeImageFit fit = OfficeImageFit.Contain;
```

`TryIdentify(...)` retains the metadata reader's extension fallback. `TryIdentifyByContent(...)`
may use a file name to select the SVG parser, but succeeds only when the bytes match a supported format.

### Text shaping and vertical text

Font resolution, glyph coverage, shaping diagnostics, and baseline placement are shared by the
raster, SVG, and PDF drawing routes. Use `OfficeDrawing.AddVerticalText(...)` for top-to-bottom text.
The optional `OfficeIMO.Drawing.HarfBuzz` package supplies full OpenType shaping and true vertical
advances to raster outline rendering. SVG retains one searchable logical string and explicitly marks
browser-native shaping with vertical writing attributes. PDF uses positioned embedded glyphs when
the provider supplies complete vertical advances and logical coverage and the writer can align the
glyph ink within its text box. Otherwise it retains searchable stacked text, reports a typed
approximation, and strict conversion profiles reject that approximation. The dependency-free
managed provider likewise reports when it cannot supply true vertical shaping.

Use `TryValidateContent(...)` at ingestion and export boundaries that must reject incomplete or
corrupt image bodies. It applies the shared encoded-payload limit, validates the complete known
container, decodes supported raster bodies, and validates every ICO entry. Both byte-array and
stream overloads return the validated metadata; seekable streams are restored to their original
position.

```csharp
using System.IO;
using OfficeIMO.Drawing;

byte[] upload = File.ReadAllBytes("upload.png");
bool validBytes = OfficeImageReader.TryValidateContent(upload, "upload.png", out OfficeImageInfo byteInfo);

using Stream input = File.OpenRead("upload.png");
bool validStream = OfficeImageReader.TryValidateContent(input, "upload.png", out OfficeImageInfo streamInfo);
```

### Bounded SVG safety checks

Use `IsWithinSafetyLimits(...)` before sending untrusted SVG to the ChartForgeX raster fallback. The predicate models the packaged ChartForgeX rasterizer's resource, reference, style, and work behavior; it is not a general safety approval for arbitrary SVG renderers:

```csharp
using OfficeIMO.Drawing;

byte[] svg = File.ReadAllBytes("upload.svg");
var limits = new OfficeSvgDrawingReaderOptions {
    MaximumElements = 10_000,
    MaximumViewportDimension = 8_192,
    MaximumViewportPixels = 16 * 1024 * 1024
};

if (!OfficeSvgDrawingReader.IsWithinSafetyLimits(svg, limits)) {
    throw new InvalidDataException("The SVG exceeds the accepted safety profile.");
}
```

A `true` result means the payload is well-formed SVG and stays within the input, XML nesting, viewport, path-command, element, rendered-reference, rendered-payload, projected raster-paint, and filter-work ceilings enforced for the ChartForgeX fallback. It does not authorize network access or external resource loading by another renderer, and it does not mean every SVG feature can be projected into an `OfficeDrawing`. Apply renderer-specific resource and execution policies before using another SVG engine. `TryRead(...)` performs OfficeIMO's vector projection and reports unsupported features; use it when the drawing result is required.

`MaximumElements`, `MaximumViewportDimension`, and `MaximumViewportPixels` can be lowered for an application policy or raised for trusted input up to their documented hard maxima. They do not relax the fixed 8 MiB input, nesting, path-command, transform, reference-depth, conservative stylesheet/reference, or 256-viewport raster-work checks. Raster work includes projected paint bounds and the estimated cost of blur, morphology, convolution, and turbulence filter parameters.
### Inspect and remove provenance carriers

`OfficeIMO.Provenance` inspects C2PA Content Credentials and IPTC Digital Source Type declarations without loading a cryptographic provider. The bounded parsers understand the format-native carriers used by JPEG, PNG, WebP, GIF, TIFF, SVG, ZIP-based document packages, and structured or variation-selector text.

```csharp
using OfficeIMO.Provenance;

OfficeProvenanceReport report = OfficeProvenanceInspector.InspectFile("generated-image.png");
Console.WriteLine($"C2PA: {report.HasC2paManifest}");
Console.WriteLine($"Generative AI declaration: {report.HasGenerativeAiDeclaration}");

OfficeProvenanceRemovalResult removal = OfficeProvenanceRemover.RemoveFile(
    "generated-image.png",
    "clean-image.png");

foreach (OfficeProvenanceChange change in removal.Changes) {
    Console.WriteLine($"{change.Carrier}: {change.Location}");
}
```

Removal is selective. It removes structurally valid C2PA carriers and AI-specific `trainedAlgorithmicMedia` or `compositeWithTrainedAlgorithmicMedia` declarations while preserving unrelated metadata and non-AI source declarations. The generic Core API blocks signed ZIP packages because rewriting the package invalidates its signatures; callers that own the complete document save must handle signature invalidation separately.

Structural inspection does not claim that a manifest is authentic or trusted. Install `OfficeIMO.Security` and use its optional C2PA verifier when content binding, signature mathematics, and certificate trust must be checked.

For an evidence-oriented result, combine structural carriers, optional cryptographic verification, exact Unicode findings, and vendor-specific detectors without collapsing them into an unreliable universal AI verdict:

```csharp
OfficeProvenanceAssessmentReport assessment =
    OfficeProvenanceAssessment.InspectFile("article.md");

Console.WriteLine($"Verified credential: {assessment.HasVerifiedContentCredential}");
foreach (OfficeProvenanceSignalResult signal in assessment.ProviderSignals) {
    Console.WriteLine($"{signal.ProviderName}: {signal.Status}");
}
```

`IOfficeProvenanceSignalDetector` is deliberately provider-specific. A detector reports its own durable-media watermark, statistical-text watermark, visible disclosure, or deterministic artifact and keeps `NotDetected`, `Inconclusive`, `ProviderUnavailable`, and `Error` distinct. OfficeIMO does not turn the absence of one vendor's signal into “human-authored.”

Existing detector and verifier implementations remain valid. Providers that perform long-running work can additionally implement `ICancellableOfficeProvenanceSignalDetector` or `ICancellableOfficeProvenanceVerifier`; cancellation-aware assessment and workflow calls pass their token to those providers and fall back to the original contracts for compatibility.

Use `OfficeTextIntegrityInspector` to report exact invisible and context-sensitive Unicode code points. It reports offsets, code points, and risk; it does not call those characters an AI watermark. Format content-safety reports also mirror these as selectable `NonPrintingUnicode` findings when the owning adapter can verify and rewrite the exact native text node. Cleanup has no blanket mode: callers pass only reviewed finding IDs, so legitimate joiners, variation selectors, and typographic spaces are not silently normalized.

`OfficeProvenanceAssessmentReport.TextIntegrityStatus`, `VerificationStatus`, and `ProviderSignalsStatus` describe whether each check ran. A disabled or unsupported text check has no report; an absent provider is `NotConfigured`. Provider-specific result statuses retain their separate conclusions.

`OfficeTextIntegrityReview` binds occurrence selections to the exact source text and exports a separate copy without normalization:

```csharp
var review = OfficeTextIntegrityReview.Inspect("invoice\u202E123 · language joiner: a\u200Db");
// After reviewing occurrence 0 (the directional override), remove only that occurrence.
byte[] copy = review.ExportSelected(review.Text, new[] { 0 });
```

For files, pass the encoded bytes to `Inspect(bytes)`. Strict UTF-8 and BOM-declared UTF-16/32 are supported. Export retains the original encoding, BOM, line endings, and all unselected characters. `RemoveSelected(currentText, indices)` rejects changed source text and invalid indices; callers must inspect edits again. The review also exposes hashes of the decoded UTF-16 code units and, for file inputs, the original bytes.

### Inspect concealed content before model ingestion

`OfficeIMO.ContentSafety` is separate from provenance. It reports native hidden text, white-on-white or otherwise low-contrast text, tiny or zero-size text, off-canvas/clipped content, notes/comments/alternative text, and exact Unicode evidence through the format package that understands the file. Concealment can be legitimate accessibility, review, layout, or metadata content; it is not an AI watermark or an authorship verdict.

`OfficeContentInstructionDetector.Analyze(text)` also provides advisory instruction signals for text consumers.
It inspects at most one million source characters, 32 inline Base64 candidates and 32,768 decoded characters by
default. Base64 inspection supports one layer of printable UTF-8 text and line-wrapped tokens; it never executes
or returns decoded instructions. `IsComplete` reports budget coverage, not safety. `Detect(text)` returns the same
signal identifiers for callers that only need a bounded heuristic list. Source text remains unchanged.

```csharp
using OfficeIMO.ContentSafety;
using OfficeIMO.Word;

OfficeContentSafetyReport report = WordDocument.InspectContentSafety("candidate.docx");
foreach (OfficeContentSafetyFinding finding in report.Findings) {
    Console.WriteLine($"{finding.Risk}: {finding.Kind} at {finding.Location}");
}

OfficeContentSafetyFinding[] reviewed = report.Findings
    .Where(item => item.IsInstructionLike)
    .ToArray();

WordDocument.RemoveSelectedContent(
    "candidate.docx",
    "candidate.cleaned.docx",
    new OfficeContentCleanupSelection(reviewed.Select(item => item.Id)));
```

SVG uses the same report and selection contracts through the native drawing owner. The inspection resolves presentation attributes and the bounded stylesheet cascade, then uses the drawing renderer for clipping, compositing, paint-order, and background evidence. Visual comparison findings are report-only: a bounded native render cannot prove browser-equivalent shaping and paint. Exact cleanup remains available for supported source-local structural findings and reviewed Unicode code points. Text in reusable definitions, `use` or `tref` targets, conditional branches, unrecognized elements, and documents with scripts, event handlers, executable links, or animation also remains report-only because removing one source node can alter a different visible instance, renderer-selected branch, or runtime state.

```csharp
using OfficeIMO.Drawing;

OfficeContentSafetyReport svgReport =
    OfficeSvgDrawingReader.InspectContentSafety("diagram.svg");

OfficeContentSafetyFinding[] concealedSvgText = svgReport.Findings
    .Where(item => item.CleanupCapability == OfficeContentCleanupCapability.RemoveText)
    .ToArray();

OfficeSvgDrawingReader.RemoveSelectedContent(
    "diagram.svg",
    "diagram.cleaned.svg",
    new OfficeContentCleanupSelection(concealedSvgText.Select(item => item.Id)));
```

Visual paint-order comparisons use explicit count and cumulative rendered-pixel budgets on `OfficeSvgDrawingReaderOptions`, plus fixed aggregate CSS-match and document-transformation ceilings. Callers can lower the public budgets for hostile-input services or set `MaximumContentSafetyVisualComparisons` to zero for structural-only inspection; the resulting report records that visual evidence was disabled. The safety inspection fails closed when declarations, values, selectors, at-rules, cascade keywords, or CSS work exceed the supported bounded subset, rather than using partial style evidence for removable findings. A changed UTF-8 cleanup result must also remain within the native SVG reader's hard input limit so reopen validation cannot be bypassed by encoding expansion.

Cleanup is always selection-based and stale-evidence checked. The adapter reopens and reinspects the rewritten artifact; it does not provide a blanket “remove anything unusual” switch. See the [content safety support matrix](../Docs/officeimo.content-safety-support-matrix.md) for exact format coverage and renderer boundaries.

Transformations can make an existing Content Credential invalid. `OfficeProvenanceLifecycle.FinalizeFile` makes the disposition explicit:

```csharp
var disposition = new OfficeProvenanceTransformationOptions {
    Policy = OfficeProvenanceTransformationPolicy.PreserveIfUnchanged
};

OfficeProvenanceTransformationResult audit = OfficeProvenanceLifecycle.FinalizeFile(
    sourcePath: "source.png",
    candidatePath: "resized.png",
    outputPath: "final.png",
    options: disposition);
```

The default blocks a changed credentialed source. Applications must explicitly select `RemoveInvalidated` for auditable carrier removal or `SignAsDerived` with an `IOfficeProvenanceSigner`, which records the source as the parent ingredient. Package-owned Word, Excel, PowerPoint, Visio, OpenDocument, EPUB, and PDF removal adapters remain the correct surfaces when package signatures also need safe disposition.

The [provenance support matrix](../Docs/officeimo.provenance-support-matrix.md) records the exact carriers, strict-removal behavior, verification boundary, and known limits.

### Encode common raster formats

```csharp
using OfficeIMO.Drawing;

var image = new OfficeRasterImage(320, 180, OfficeColor.White);
var options = new OfficeRasterEncodingOptions {
    Jpeg = new OfficeJpegEncodeOptions { Quality = 90 },
    Tiff = new OfficeTiffEncodeOptions {
        Compression = OfficeTiffCompression.Lzw,
        Predictor = OfficeTiffPredictor.Horizontal
    }
};

byte[] jpeg = OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Jpeg, options);
byte[] tiff = OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Tiff, options);
byte[] webp = OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Webp, options);
```

The same encoder can write directly to a caller-owned stream when materializing another complete output array would be wasteful. On .NET 8 and later, the codecs also accept `IBufferWriter<byte>` without adding a package dependency to `OfficeIMO.Core`:

```csharp
using System.Buffers;

using (Stream output = File.Create("photo.webp")) {
    OfficeRasterImageEncoder.EncodeTo(image, OfficeImageExportFormat.Webp, output, options);
}

var writer = new ArrayBufferWriter<byte>();
OfficeRasterImageEncoder.EncodeTo(image, OfficeImageExportFormat.Png, writer, options);
ReadOnlyMemory<byte> png = writer.WrittenMemory;
```

The stream overload leaves the destination open. PNG's `Optimal` compression compares adaptive and unfiltered rows in the selected sample layout and writes the smaller compressed form. This preserves pixels, density metadata, cancellation and the encoded-byte ceiling. `Fastest` writes an unfiltered RGBA stream with the platform's fastest Deflate setting. `Stored` writes uncompressed zlib blocks with eight-bit RGBA samples.

Size probes share bounded scanline scratch. An unfiltered candidate that fits in the existing 64 KiB IDAT buffer can be written directly when it wins. Ordinary byte-array exports of non-bilevel images with at least 1 MiB of RGBA pixels can also retain the adaptive candidate in their own output, avoiding recompression when it wins. Retention stops at 4 MiB of compressed payload; the output stream's capacity can exceed that payload limit. Other candidates use a final compression pass. Caller-owned streams, buffer writers and exports with an encoded-byte ceiling use the bounded scanline and IDAT-buffer path.

TIFF output is a classic RGBA image with uncompressed, LZW, PackBits, or Deflate strips; LZW and Deflate use horizontal prediction by default. Use `OfficeTiffCodec.EncodePages(...)` when the output needs more than one page. JPEG uses the managed quality, subsampling, progressive, metadata, and transparency-flattening settings.

TIFF `DpiX` and `DpiY` are physical DPI. To retain native centimeter density or a unitless aspect ratio, assign `OfficeTiffEncodeOptions.Resolution = new OfficeImageResolution(horizontal, vertical, unit)`. That immutable override applies to every encoded page and to streaming output. Meter values normalize to centimeters. The shared encoder retains the override until an explicitly assigned shared `DpiX` or `DpiY` selects physical DPI instead; `metadata.Resolution` supplies a snapshot directly from `OfficeImageMetadata`.

The shared encoder also writes BMP, binary PBM, TGA, and ICO with `OfficeImageExportFormat.Bmp`, `Pbm`, `Tga`, and `Icon`. BMP uses explicit RGBA masks and TGA writes a straight-alpha 32-bit image. PBM composites transparency over white and thresholds luminance into a one-bit image. ICO contains a PNG entry; `OfficeIconEncoder.Encode(images, options, cancellationToken)` accepts several images for a multi-resolution icon, with each image limited to 256 pixels on either axis.

Managed decoding accepts ASCII and binary PBM/PGM/PPM, including sixteen-bit PGM/PPM samples; indexed, grayscale, true-color, and RLE TGA; and PNG or conventional DIB icon entries. Select an ICO entry through `OfficeRasterDecodeOptions.FrameIndex`. Container inspection checks entry dimensions before allocating pixel buffers.

### Choose lossless or lossy WebP

WebP output defaults to lossless VP8L. The byte-array encoder chooses bounded prediction, subtract-green, LZ77, and Huffman coding when that is smaller than the literal form; direct streaming uses the literal form to reduce retained buffers. Set `Webp.Mode` to `Lossy` for VP8 output with an explicit quality:

```csharp
var webpOptions = new OfficeRasterEncodingOptions {
    Webp = new OfficeWebpEncodeOptions {
        Mode = OfficeWebpEncodingMode.Lossy,
        Quality = 85,
        WritePhysicalResolution = true
    }
};

using Stream webpOutput = File.Create("photo.webp");
OfficeRasterImageEncoder.EncodeTo(
    image, OfficeImageExportFormat.Webp, webpOutput, webpOptions);
```

Lossy quality ranges from 1 through 100 and defaults to 85. VP8 uses 4:2:0 chroma sampling, so quality 100 still changes color samples; alpha remains exact at every quality. Transparent output stores filtered, uncompressed alpha. Lossy images accept dimensions up to 16,383 pixels on either axis; lossless VP8L accepts 16,384. Invalid quality or unsupported dimensions fail explicitly. Both modes support cancellation and the shared encoded-byte and managed working-memory limits. The encoder writes one static image; animated WebP encoding is outside its contract.

`OfficeWebpCodec.Encode` and `EncodeTo` accept the same `OfficeWebpEncodeOptions` directly. Direct codec options omit physical resolution metadata by default. Set `WritePhysicalResolution` to `true` and supply `DpiX` and `DpiY` to include EXIF density; format-neutral encoding applies explicitly assigned shared DPI values ahead of the WebP-specific values.

### Inspect and select frames or pages

`OfficeRasterContainerInspector` exposes one bounded inventory for static images, GIF and APNG frames, TIFF pages, and WebP frame metadata. The inventory includes canvas size, count, timing, offsets, disposal, blending, loop count, and the backwards-compatible default image. Decode accepts either bytes or a readable stream; a seekable stream is restored to its original position.

```csharp
using var input = File.OpenRead("pages.tiff");
var decodeOptions = new OfficeRasterDecodeOptions {
    FrameIndex = 1,
    FrameLossPolicy = OfficeRasterFrameLossPolicy.UseSelectedFrame,
    MaximumEncodedBytes = 32 * 1024 * 1024,
    MaximumDecodedPixels = 20_000_000,
    CancellationToken = cancellationToken
};

if (OfficeRasterImageDecoder.TryDecode(input, decodeOptions, out var page, out var decodeInfo)) {
    Console.WriteLine($"Selected {decodeInfo.SelectedFrameIndex + 1} of {decodeInfo.FrameCount}");
}
```

Set `FrameLossPolicy` to `RejectMultipleFrames` when a static result must not discard animation frames or document pages. Animated WebP pixel composition remains a caller-codec boundary, but its frame inventory is still available for a fail-closed decision.

The managed JPEG decoder supports eight-bit baseline and eight/twelve-bit extended sequential and progressive DCT, plus two-through-sixteen-bit Huffman and arithmetic lossless JPEG. Lossless decoding supports all seven predictors, point transforms and row-aligned restart intervals. A nonzero point transform discards source low bits and restores them as zero. Standalone decoding projects to the public eight-bit raster buffer. Lossless JPEG-compressed TIFF retains native two-through-sixteen-bit samples through alpha and color conversion. YCbCr conversion uses the native sample midpoint before eight-bit projection. TIFF retains fractional RGB through alpha unassociation and ICC conversion instead of rounding it back to the source bit depth. Sequential and progressive arithmetic JPEG support eight/twelve-bit samples, conditioning tables and restart intervals. Progressive decoding accepts spectral bands, successive approximation and DC-only previews, retaining missing coefficients as zero. Lossless arithmetic uses two-dimensional difference conditioning and bounded row history; grayscale/RGB fixtures cover all precisions and predictors. DCT JPEG-compressed TIFF retains native eight/twelve-bit samples; packed twelve-bit TIFF supports byte-aligned rows with Predictor 1.

The managed JPEG XR decoder accepts single-image tagged `.jxr`, `.wdp`, and `.hdp` containers with unsigned eight-bit gray/RGB/BGR/BGRA, unsigned sixteen-bit gray/RGB/RGBA, and finite sixteen/thirty-two-bit fixed-point or floating-point gray/RGB/RGBA pixels. It handles 4:4:4, 4:2:2, and 4:2:0 chroma with defined sampling-grid centering, spatial and frequency packet order, lossless and lossy quantization, all three overlap modes, hard and soft tiles, omitted high-frequency bands, trimmed flexbits, and interleaved or separate alpha. Premultiplied alpha becomes straight RGBA at source precision before samples are rounded to the public eight-bit buffer. Fixed-point and floating-point samples default to linear scRGB and use the shared sRGB conversion; out-of-range colors are clipped to the SDR output gamut. The container orientation is applied. Encoded bytes, padded coefficient/sample buffers, output, retained caller data, and cancellation share the Core resource limits. Embedded ICC data is available to the color-management layer; ordinary unsigned raster decoding returns device channel values. The color-management layer also decodes unsigned eight/sixteen-bit CMYK and CMYKDirect with a supplied four-component ICC profile, preserving source precision and alpha. Unprofiled CMYK does not produce approximate RGB pixels. Three-to-eight-channel unsigned eight/sixteen-bit images also require a matching ICC profile. Interleaved alpha can contain the same or fewer frequency bands than the primary plane. Non-finite samples and multiple image directories are outside this decoder contract.


### Read and edit portable image metadata

`OfficeImageMetadata` owns typed Exif values and raw XMP, ICC, and IPTC IIM profile bytes. Its arrays and snapshots own their storage. Standard tags have descriptive names; application-defined tags use their numeric identifier, TIFF representation, and directory. Scalars and arrays use .NET numeric types or `OfficeRational` / `OfficeSignedRational`.

```csharp
byte[] original = File.ReadAllBytes("photo.jpg");
OfficeImageMetadata metadata = OfficeImageMetadata.Read(original);
metadata.SetExifValue(OfficeExifTag.Software, "Photo workflow");
metadata.SetExifValue(OfficeExifTag.ExposureTime, new OfficeRational(1, 125));
metadata.RemoveExifValue(OfficeExifTag.GPSLatitude);
metadata.RemoveExifValue(OfficeExifTag.GPSLongitude);
File.WriteAllBytes("annotated.jpg", OfficeImageMetadata.Apply(original, metadata));

OfficeImageMetadataRemovalResult removal = OfficeImageMetadata.Remove(
    original, OfficeImageMetadataProfileKinds.Exif | OfficeImageMetadataProfileKinds.Xmp);
Console.WriteLine($"Found {removal.PresentProfiles}; removed {removal.RemovedProfiles}");
File.WriteAllBytes("private.jpg", removal.EncodedBytes);
```

Metadata replacement preserves JPEG entropy data, PNG image chunks, WebP image payloads, and TIFF raster encoding fields. TIFF replacement edits the primary image and retains page links; profile removal visits every page. TIFF metadata operations limit the aggregate strip, tile, and JPEG-interchange pixel references across reachable directories to 65,535. The alias check includes child directories and rejects edits that would erase pixel data. GIF supports XMP and ICC application profiles, while BMP supports V5 ICC profiles and physical density. Removing profile families from PBM/PGM/PPM or TGA returns the validated image unchanged because those formats have no defined carriers for these profile families. Unsupported carriers fail explicitly. ICC metadata edits preserve profile bytes; they do not transform pixel colors.

Resolution values retain native aspect-ratio, pixels-per-inch, pixels-per-centimeter, or pixels-per-meter units when the container can represent them. `PhysicalDpiX` and `PhysicalDpiY` convert physical units to DPI and return null for an aspect ratio. JPEG, WebP, and TIFF normalize meter-based density to centimeters. JPEG replacement updates JFIF density, creates that carrier when absent, and synchronizes density in an existing Exif profile. GIF has an aspect-ratio field and no physical-density field: lossless edits require `AspectRatio`, while `PrepareForEncoding` projects resolution to that unit for GIF. Its ratio is rounded to the GIF field's representable increments. TIFF-relative opaque maker notes require the original TIFF container for lossless preservation; `RequiresOriginalTiffContainer` reports that constraint. Exporting those notes to another container requires explicitly removing or replacing the note. Superseded Exif values are erased when their storage is exclusive; aliased ranges that would erase another field or image pixels are rejected.

Use `ParseExifProfile` and `EncodeExifProfile` for metadata-only workflows. They accept and return classic TIFF Exif bytes without JPEG framing, and parsing also accepts the JPEG Exif prefix. Both have cancellation-token overloads. `Clone` preserves independent edits and the original-container constraints. Metadata rewriting bounds retained inputs, profiles, stream backing growth, and the final encoded array against the Core managed working-set limit. Malformed ICC profiles are rejected before replacement. C2PA removal uses the canonical carrier inspector and retains malformed or conflicting carriers rather than treating unrelated bytes as a removable manifest.

When re-encoding pixels into another format, call `PrepareForEncoding(destinationFormat, out omittedProfiles)` before `Apply`. The returned copy contains supported profile families and reports what the destination cannot carry. For example, JPEG IPTC IIM metadata is omitted when producing PNG, while Exif, XMP, and ICC are retained. This explicit projection leaves the source metadata intact; lossless replacement still rejects unsupported profiles.

### Inspect and edit HEIF container metadata

`OfficeHeifMetadataReader` reads HEIF/HEIC brands, primary image dimensions and transforms,
item properties, locations, and references. Its EXIF and XMP methods read each metadata
family independently without decoding image pixels:

```csharp
using OfficeIMO.Drawing;

if (OfficeHeifMetadataReader.TryReadInfo("photo.heic", out OfficeHeifImageInfo? info)) {
    Console.WriteLine($"{info!.Width}x{info.Height}; EXIF: {info.HasExif}");
}

if (OfficeHeifMetadataReader.TryReadExifProfile("photo.heic", out OfficeImageMetadata? exif) &&
    exif != null) {
    exif.SetExifValue(OfficeExifTag.Software, "Photo workflow");
    bool saved = OfficeHeifMetadataReader.TryWriteExifProfile("photo.heic", "annotated.heic", exif);
}
```

Byte-array overloads return independently owned output. Stream readers start at the current
position, restore seekable streams, and leave them open. `HasExifItem` and `HasXmpItem` report
declared items even when no readable payload is located. The writer replaces or clears an
existing single absolute extent within an `mdat` payload; null metadata clears the item.
It does not create missing items or rewrite item-data-box, external, derived, or multi-extent
storage. Shared extents are rejected so clearing old metadata cannot erase image pixels or
another item. Replacement bytes occupy a framed `mdat` box; unchanged sibling payloads are
preserved.

Protected items and XMP items with a nonempty MIME content-encoding declaration remain
visible through the info and presence APIs. Their payload reads and writes, including
clearing, return `false`; independent edits to another family leave their bytes unchanged.
Core does not provide a HEIF decryption or content-decoding fallback.

Input and output are bounded to 128 MiB, individual metadata/property payloads to 16 MiB,
declared item collections to 4,096 entries, and parser work to 65,536 records. Operations
include known caller stream backing in the managed working-set budget and observe
cancellation during parsing and preparation. Rejection or cancellation during
preparation leaves source bytes and an existing output file untouched; the final file write
is synchronous. This metadata API does not provide HEIF pixel decoding or encoding.

### Optimize encoded images for a placement

`OfficeImageOptimizer` resizes and re-encodes a static raster image for the pixel bounds where it will be used:

```csharp
using OfficeIMO.Drawing;

byte[] source = File.ReadAllBytes("photo.jpg");
var request = new OfficeImageOptimizationRequest(1200, 800) {
    OutputFormat = OfficeImageFormat.Jpeg,
    JpegQuality = 82,
    JpegSubsampling = OfficeJpegSubsampling.Y420,
    JpegProgressive = true,
    JpegOptimizeHuffman = true,
    ResamplingMode = OfficeRasterResamplingMode.Lanczos3,
    ResamplingColorSpace = OfficeRasterResamplingColorSpace.LinearLight,
    MetadataPolicy = OfficeImageMetadataPolicy.SelectiveCopy,
    MetadataSelection = OfficeImageMetadataKinds.Exif |
                        OfficeImageMetadataKinds.Xmp |
                        OfficeImageMetadataKinds.Icc |
                        OfficeImageMetadataKinds.Orientation |
                        OfficeImageMetadataKinds.Resolution
};

OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(source, request, "photo.jpg");
File.WriteAllBytes("photo.optimized.jpg", result.Bytes);
Console.WriteLine($"{result.Status}: {result.BytesSaved} bytes saved");
Console.WriteLine($"Metadata lost: {result.Metadata.Lost}");
```

The request preserves aspect ratio, avoids upscaling, keeps the original when re-encoding is not smaller, and retains source DPI by default. `ResamplingMode` defaults to premultiplied-alpha bilinear sampling; choose `Area` for coverage-correct downsampling or `Lanczos3` for a sharper high-quality filter. Filtering remains in encoded sRGB by default. Select `LinearLight` when the workload benefits from physically linear color averaging and accepts its additional cost. Set `OutputDpiX` and `OutputDpiY` to override density. Output can be PNG, JPEG, TIFF, or WebP; use `PngCompression`, the JPEG quality/subsampling/progressive/Huffman settings, or `TiffCompression` and `TiffPredictor` for the selected format.

`MetadataPolicy` can preserve, strip, or selectively copy EXIF, XMP, ICC, orientation, comments, and resolution categories. A JPEG-to-JPEG rewrite preserves selected EXIF, standard single-packet XMP, and ICC bytes, applies embedded orientation to pixels, and neutralizes the copied orientation value. Adobe extended XMP is not copied during re-encoding; when selected XMP includes extension segments, the result reports XMP in `Lost`. The result reports `PolicyApplied`, `Preserved`, `Normalized`, `Stripped`, and `Lost`; unsupported or undecodable input returns the original bytes with `PolicyApplied = false`, and a required strip or selective-copy rewrite is never replaced by the metadata-bearing original merely because it is smaller. Metadata that has no safe output carrier is reported as loss rather than silently claimed as preserved; OfficeIMO does not currently perform ICC color conversion. Animated and multi-page input is rejected so optimization never silently drops frames or pages.

`OfficeRasterExportPlanner` is the shared pre-allocation owner for image export. It combines the caller's `MaximumRasterPixels` with renderer and encoder dimension/pixel limits, then either reduces scale with `IMAGE_RASTER_SCALE_REDUCED` or throws `OfficeImageExportLimitException`, according to `RasterOverflowBehavior`. The returned plan also owns the effective encoding settings: `CreateEncodingOptions()` reduces encoded density with the raster scale so safety limits preserve the document's physical size. Explicit top-level `DpiX`/`DpiY` values apply across formats; when those values are not assigned, format-specific PNG, JPEG, and TIFF density remains authoritative. Drawing's managed PNG and APNG, JPEG, classic unsigned 8/12/16-bit or finite floating 16/24/32/64-bit grayscale/RGB/RGBA/device-CMYK TIFF and 8-bit palette TIFF, uncompressed BMP, composited GIF, and ordinary lossless VP8L and lossy VP8 WebP paths (including raw or lossless-compressed separate alpha planes) enforce encoded-payload and decoded-pixel guards. TIFF accepts both byte orders, chunky or planar strips and tiles with uncompressed, LZW, PackBits, or Deflate payloads and integer or floating-point horizontal prediction; sample and associated-alpha precision is retained until eight-bit RGBA output or ICC conversion. Packed twelve-bit samples require Predictor 1. Mixed component widths and reversed bit order outside CCITT are rejected; arbitrary page selection and bounded multi-page writing use the same page contract. Floating TIFF values are normalized device components, with explicit ICC conversion when supplied and clipping to the SDR range at output; they do not implicitly declare linear scRGB or request automatic scientific-data rescaling. Non-finite color or alpha samples are rejected; unspecified extra channels and tile padding are ignored. TIFF also supports packed grayscale/palette samples, CCITT T.4/T.6 including the optional uncompressed mode, and baseline eight-bit, Huffman/arithmetic extended-sequential eight/twelve-bit or Huffman/arithmetic lossless two-through-sixteen-bit JPEG payloads. Bounded legacy compression-6 interchange images and table-pointer scans retain the eight/twelve/sixteen-bit contract; see the [TIFF support and qualification boundaries](../OfficeIMO.Xps/SUPPORT.md). BigTIFF, animated WebP pixel decoding, and unsupported JPEG-in-TIFF processes remain caller-codec boundaries. `OfficeRasterImageFallbackCodec` can wrap an application codec at the final raster boundary. It reports `IMAGE_SOURCE_DECODED_BY_CALLER_CODEC` when that codec succeeds; if neither Drawing nor the application can decode a source image, it returns a visible placeholder and `IMAGE_SOURCE_DECODE_FALLBACK` instead of allowing the renderer to omit the image silently.

Every format package builds on the same fluent export contract. `FitWithin(width, height)`, `FitWithinWidth(...)`, and `FitWithinHeight(...)` cap both raster and SVG output without enlarging smaller content. `ConfigureOptions(...)` exposes the complete provider-specific option object when no dedicated fluent shortcut exists. Batch limits, cancellation, progress, and `WithRenderTimeout(...)` apply to the complete operation, including streaming saves. Each batch result reports its zero-based `SequenceIndex`; `SequenceCount` is populated when the total is known before streaming or after a fluent builder materializes the complete result list.

For raster complex text, set `TextShapingProvider` and optionally `TextShapingLanguage` on any shared image-export options or use `WithTextShaping(...)` on a fluent builder. `OfficeManagedTextShapingProvider.Instance` is the dependency-light built-in provider for the proven Arabic-script (core and extended Persian/Urdu letters)/TrueType-outline subset. Drawing passes the selected TrueType font bytes, base direction, language, cancellation token, and Unicode source mapping to any provider, then caches the resolved run for measurement and painting. The built-in provider deliberately declines CFF fonts and scripts that require broader GSUB/GPOS behavior. If no provider accepts a complex run, the managed Arabic-script and bounded bidirectional fallback keeps common text visible and adds `IMAGE_TEXT_SHAPING_FALLBACK` as an approximation. Set `Policy.RequireNoLoss = true` when that fallback is not acceptable.

### Deterministic text measurement

```csharp
using OfficeIMO.Drawing;

var measurer = OfficeTextMeasurer.Create();
var style = measurer.CreateStyle(new OfficeFontInfo("Aptos", 11, OfficeFontStyle.Regular));
OfficeTextMetrics metrics = measurer.Measure("Quarterly report", style);

if (metrics.WidthPixels > 240) {
    Console.WriteLine("The label needs wrapping or a smaller font.");
}
```

### Language-pattern hyphenation

`OfficeTextHyphenationPatterns.GetBreakpoints(token, language)` returns optional UTF-16 breaks in
the original word using embedded US English (`en-US`, alias `en`) or reformed German (`de-DE`,
aliases `de`, `de-1996`, `de-DE-1996`) resources. It preserves case, surrounding punctuation and
Unicode source offsets. The resources enforce two letters before a break, and three after it for
US English or two for German. Empty/unsupported tags, internal punctuation, digits and tokens
longer than 512 UTF-16 code units produce no automatic breaks. Other regional English and German
spelling tags are unsupported.

```csharp
using OfficeIMO.Drawing;

IReadOnlyList<int> breaks = OfficeTextHyphenationPatterns.GetBreakpoints("representation", "en-US");
// 3, 5, 8, 10; the caller decides which permitted break fits the line.
```

Renderers can use the same Core result through their existing hyphenation callbacks. The explicit
`OfficeTextHyphenationLexicon` remains available for application-owned dictionaries. Embedded
resources are versioned in `Typography/Hyphenation/manifest.json`; their copyright and permission
notices are retained in [the third-party notices](THIRD-PARTY-NOTICES.md).

### Reusable ink

```csharp
using OfficeIMO.Drawing;

var stroke = new OfficeInkStroke {
    Color = OfficeColor.Crimson,
    Width = 2.2,
    Height = 2.2,
    Bias = OfficeInkBias.Handwriting,
    RecognizedText = "hello"
};
stroke.AddPoint(8, 30, 0.35)
      .AddPoint(60, 12, 1.0)
      .AddPoint(120, 36, 0.55);

var ink = new OfficeInkDocument().Add(stroke);
OfficeDrawing inkDrawing = OfficeInkRenderer.Render(ink, width: 140, height: 60);
```

The model keeps sampled geometry, normalized pressure, pen dimensions and tip shape, transforms, handwriting/drawing bias, language, and recognition alternatives. A format adapter decides how those values map to native storage.

### Structured math

```csharp
using OfficeIMO.Drawing;

OfficeMathExpression expression = OfficeMath.Fraction(
    OfficeMath.Row(
        OfficeMath.Identifier("x"),
        OfficeMath.Operator("+"),
        OfficeMath.Number("1")),
    OfficeMath.Radical(OfficeMath.Identifier("y")));

string mathMl = OfficeMathMarkup.ToMathMl(expression);
string latex = OfficeMathMarkup.ToLatex(expression);
OfficeDrawing mathDrawing = OfficeMathRenderer.Render(expression);
```

Supply a math font when equations need its OpenType MATH spacing, glyph variants,
or stretch assemblies:

```csharp
var mathOptions = new OfficeMathRenderOptions {
    Font = new OfficeFontInfo("Document Math", 24D),
    Dpi = 144D
};
mathOptions.Fonts.Add("Document Math", File.ReadAllBytes("document-math.otf"));
OfficeDrawing equation = OfficeMathRenderer.Render(expression, mathOptions);
```

`UseFontMathMetrics` is enabled by default. Available MATH constants, italic corrections,
accent positions, math kerning, variants and assemblies guide layout; missing data uses
the existing geometric fallback. Set the option to `false` to use the caller's script and
rule settings. Font size is in points and `Dpi` controls drawing density. Optical-size
selection and font tracking use the authored point size for both measurement and paint.
Fonts are caller supplied; OfficeIMO.Core does not ship a math font or require a native
rendering engine.

The same immutable expression tree feeds native OneNote math and Word OMML adapters. The shared model includes right and left scripts, centered upper/lower limits, built-up and slashed fractions, delimiter lists, stacks, matrices, equation arrays, n-ary operators, accents, bars, boxes, and phantoms. OneNote maps all of those structures natively. Word maps the lossless OMML subset; `Stack` and `StretchStack` fail with `NotSupportedException` because OMML has no equivalent, and callers can choose `EquationArray` explicitly when that projection is intended. Drawing owns the AST, portable markup, measurement, and visual layout; each document package owns only its native codec. MathML and LaTeX parsing default to a nesting limit of 128 and expose bounded overloads; excessive nesting fails with `OfficeMathParseException.Code == "DRAWING_MATH_DEPTH"`.

## Find a conversion package

Use `OfficeConversionCapabilityCatalog` when an application needs to discover the package and public API for a source-to-target conversion. Each route includes accepted extensions, its fidelity model, the result type that carries diagnostics, and whether the route is available in the browser converter.

```csharp
using OfficeIMO;

foreach (OfficeConversionCapability route in
         OfficeConversionCapabilityCatalog.FindBySourceExtension(".docx")) {
    Console.WriteLine($"{route.Id}: {route.PackageId} -> {route.TargetExtension}");
}
```

`OfficeIMO.Core` describes these routes but does not execute them. Add the package named by `PackageId`, call the API shown by `Api`, and inspect the returned result or report before accepting the output.

## Examples

### Build a reusable vector scene

```csharp
using OfficeIMO.Drawing;

var drawing = new OfficeDrawing(width: 420, height: 180)
    .AddShape(new OfficeShape {
        Kind = OfficeShapeKind.Rectangle,
        Width = 420,
        Height = 180,
        FillGradient = OfficeLinearGradient.Horizontal(
            OfficeColor.Parse("#F8FBFF"),
            OfficeColor.Parse("#EAF4FF")),
        StrokeColor = OfficeColor.Parse("#B7D7F5"),
        StrokeWidth = 1
    }, x: 0, y: 0)
    .AddText("OfficeIMO.Drawing", 20, 18, 380, 32,
        new OfficeFontInfo("Aptos", 18, OfficeFontStyle.Bold),
        OfficeColor.Parse("#1F2937"),
        OfficeTextAlignment.Left)
    .AddShape(OfficeShape.RoundedRectangle(140, 44, 10), 20, 86)
    .AddText("Shared vector intent", 34, 98, 240, 24);

OfficeDrawingQualityReport report = OfficeDrawingQualityAnalyzer.Analyze(drawing);
// Compare the same scene with a smaller delivery canvas without changing it.
OfficeDrawingQualityReport target = OfficeDrawingQualityAnalyzer.Analyze(drawing, 200, 200);
if (report.HasIssues) {
    foreach (var issue in report.Issues) {
        Console.WriteLine($"{issue.Kind}: {issue.Message}");
    }
}
```

Affine effect groups contribute their transformed child bounds, not the dimensions
of their temporary rendering buffers. Empty groups do not create overflow findings.

### Render a chart snapshot to drawing primitives

```csharp
using OfficeIMO.Drawing;

var snapshot = new OfficeChartSnapshot(
    name: "RevenueChart",
    title: "Revenue by quarter",
    chartKind: OfficeChartKind.ColumnClustered,
    data: new OfficeChartData(
        new[] { "Q1", "Q2", "Q3", "Q4" },
        new[] {
            new OfficeChartSeries("Revenue", new[] { 10d, 18d, 24d, 30d }),
            new OfficeChartSeries("Forecast", new[] { 12d, 19d, 25d, 33d })
        }),
    widthPoints: 420,
    heightPoints: 260);

OfficeChartRenderingResult rendered = OfficeChartDrawingRenderer.RenderWithQuality(snapshot);
OfficeDrawing chartDrawing = rendered.Drawing;

foreach (var issue in rendered.QualityReport.Issues) {
    Console.WriteLine(issue.Message);
}
```

Set a legend frame fill and outline through the shared chart style:

```csharp
var framedStyle = new OfficeChartStyle(
    showBackground: true,
    legendBackgroundColor: OfficeColor.Parse("#FFFFFF"),
    legendBorderColor: OfficeColor.Parse("#375A7F"),
    legendBorderWidth: 1);
var framedSnapshot = new OfficeChartSnapshot("RevenueChart", "Revenue by quarter",
    OfficeChartKind.ColumnClustered, snapshot.Data, 420, 260, style: framedStyle);
OfficeDrawing framedChart = OfficeChartDrawingRenderer.Render(framedSnapshot);
```

The background and border are optional; the border width is a positive value in points.
These settings style the chart legend frame in shared static renders.

Use `layout.WithSecondaryValueAxis(new OfficeChartValueAxisLayout(minimum: 0, maximum: 1,
majorUnit: 0.2, numberFormat: "0%").WithTitle("Completion rate"))` for an independent secondary value-axis scale and title.
Updating the scale without `WithTitle` preserves an existing native title; `WithTitle(null)` removes it explicitly.
Use optional `majorTickMark` and `minorTickMark` settings to select independent tick appearance.
The returned layout preserves the primary axis settings and leaves the original layout unchanged.
Bounds must be finite and ordered; tick units must be positive.
Explicit numeric bounds clip line, area, and scatter series paint to the plot; points outside the range
do not gain false edge markers or labels. Data labels for visible points can extend beyond the plot.

### Style individual chart points

Use `OfficeChartSeries.WithPointStyles` to attach appearances aligned with the series values.
The returned series preserves its data and other settings. A null entry inherits the point
colour; an explicit no-fill style leaves an outlined slice visible without changing its value.

```csharp
var status = new OfficeChartSeries("Status", new[] { 8d, 2d, 1d })
    .WithPointStyles(new OfficeChartPointStyle?[] {
        new(fillColor: OfficeColor.Parse("#168A56")),
        new(noFill: true, outlineColor: OfficeColor.Black, outlineWidth: 2,
            outlineJoin: OfficeStrokeLineJoin.Round),
        new(hatch: OfficeChartHatchPattern.WideForwardDiagonal,
            hatchColor: OfficeColor.Parse("#7300A3"), outlineColor: OfficeColor.Black)
    });
```

Hatches support horizontal, vertical, forward diagonal, backward diagonal, cross,
diagonal cross, and wide forward diagonal strokes. Their default background is white;
set `fillColor` to choose another background. Outline widths use points; joins can be
round, bevel, or miter. `showOutline: false` hides the outline explicitly.
The renderer clips hatches to slice, bar, and marker geometry and styles category legend
swatches with the same appearance. Per-point area styling is not applied by static rendering;
`RenderWithQuality` reports `UnsupportedAppearance` and retains the opaque series fill.
An area series with no native outline renders without a connecting line.

### Pie and doughnut geometry

Pass an `OfficeChartRadialLayout` as the final argument to the style/layout snapshot
constructor to set the first slice boundary and doughnut hole. Rotation is clockwise
from the top, from 0 through 360 degrees; hole size is an inner-to-outer diameter
percentage from 10 through 90, with a default of 50.

```csharp
var series = new OfficeChartSeries("Status", new[] { 7d, 3d })
    .WithPointExplosions(new[] { 25, 0 });
var data = new OfficeChartData(new[] { "Complete", "Pending" }, new[] { series });
var radial = new OfficeChartRadialLayout(firstSliceAngleDegrees: 90, doughnutHolePercent: 70);
var snapshot = new OfficeChartSnapshot("Status", null, OfficeChartKind.Doughnut,
    data, 400, 300, null, new OfficeChartLayout(showLegend: false), radial);
OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(snapshot);
```

Multiple doughnut series form contiguous rings, starting with the first series at the
inside. `WithPointExplosions` offsets selected pie or doughnut slices by a percentage
of their radius while preserving their native editable point records in Word,
PowerPoint, and Excel. Supply one value per point, from 0 through 400; an explicit
zero resets an earlier offset. The renderer keeps exploded slices within the chart
frame and moves their value-label anchors with them. For outside labels, set
`showDataLabels: true`, select the label contents, and use
`dataLabelPosition: OfficeChartDataLabelPosition.OutsideEnd` in `OfficeChartLayout`.
The renderer reserves side gutters, separates labels on each side and connects
them to visible slices. Call `layout.WithDataLabelLeaderLines(false)` to omit the lines.
For multi-series doughnuts, leaders work when `DataLabelSeriesIndexes` selects only the
outer ring. Disable leaders if labeled inner rings would send lines through outer rings.
Imported Word, PowerPoint and Excel charts retain that native line setting in
their static chart snapshots. A frame too narrow or short to fit every outside
label reports an unsupported layout instead of overlapping or dropping labels.
For imported pie and doughnut charts, the first six `varyColors` points use
the document theme colors from a direct, untransformed modern chart color
style, or theme accents for an unstyled/classic style 2 chart. Native point
fills take precedence. Transformed color styles and longer sequences still
need producer-specific proof.

### Load first-party font programs for renderers

```csharp
using System.Collections.Generic;
using System.IO;
using OfficeIMO.Drawing;

OfficeTrueTypeFont? font = OfficeTrueTypeFont.TryLoadDefault(out string? path);
if (font != null) {
    Console.WriteLine($"Loaded {path}");
}

var embedded = new OfficeFontFaceCollection {
    FontVariationResolver = request => request.FamilyName == "Report Variable"
        ? new Dictionary<string, float> { ["wght"] = 720 }
        : null
};
embedded.Add("Report Variable", File.ReadAllBytes("ReportVariable.ttf"));
```

`OfficeFontFaceCollection` accepts TrueType-glyf OpenType, WOFF 1, CFF/CFF2, and TrueType or CFF2 variable fonts. Single-face WOFF 2 decoding is available on .NET 8 and newer; extract and register individual faces from WOFF 2 font collections. The engine is part of `OfficeIMO.Core`; it does not require another font-program package or a license key.

## What it provides

- `DocumentAccessMode`, `DocumentPersistenceMode`, `DocumentCreateOptions`, and `DocumentLoadOptions` for one lifecycle vocabulary across document packages.
- Dependency-free `IOfficeSecurityProvider`, CMS/X.509/XML-signature requests, options, findings, and results under the `OfficeIMO.Security` namespace. The same layer owns bounded OPC, VBA, and ODF/EPUB XML package-signature structure and atomic commit policy. The optional `OfficeIMO.Security` package supplies the concrete cryptographic provider.
- Dependency-free provenance inspection and selective removal contracts under `OfficeIMO.Provenance`, including exact C2PA carriers and IPTC AI-source declarations. Optional cryptographic verification remains in `OfficeIMO.Security`.
- `OfficeColor` immutable RGBA values with named colors and hex parsing.
- `OfficeColorSpaceConverter` for dependency-free CMYK, CIE Lab/XYZ, calibrated gray, and calibrated RGB conversion to sRGB.
- `OfficeImageReader` and `OfficeImageInfo` for dependency-free image inspection where supported.
- `OfficeImageFit` for shared stretch, contain, and cover intent.
- `OfficeFontInfo`, `OfficeFontStyle`, `OfficeTextMeasurer`, and `OfficeTextMetrics` for deterministic layout estimates.
- `OfficeTrueTypeFont` and `OfficeFontFaceCollection` for first-party static and variable font measurement, glyph contours, and bounded renderer integration.
- `OfficeInkDocument`, `OfficeInkStroke`, sampled pressure/style metadata, recognition alternatives, and `OfficeInkRenderer` for reusable ink capture and projection.
- `OfficeMathExpression`, `OfficeMath`, MathML/LaTeX conversion, deterministic measurement, and `OfficeMathRenderer` for reusable structured equations.
- `OfficeShape`, `OfficeDrawing`, gradients, shadows, transforms, clipping, and vector descriptors that format-specific packages can map into their own coordinate systems.
- Linear and radial gradients support encoded sRGB and linear-light RGB interpolation. Use `gradient.WithColorInterpolation(OfficeGradientColorInterpolation.LinearRgb)` to create a detached linear-light gradient; `ColorInterpolation` reports its mode. Cloning, coordinate transforms, SVG import/export and raster output retain the mode. SVG `color-interpolation` inherits through CSS ancestors, including inline styles, independently of gradient template references.
- `OfficeRadialGradient.TransformCoordinates` preserves radial fields under affine rotation, shear and reflection. Its `CoordinateTransform` maps gradient coordinates into normalized shape coordinates; cloning and opacity changes retain that mapping. SVG import expands bounded radial repeat/reflect fields within the shared stop budget and retains other fields with `OfficeRadialGradient.WithSpreadMode(OfficeGradientSpreadMode.Repeat)` or `Reflect`. Cloning, transforms, raster sampling and SVG export preserve the spread mode. PDF export still requires finite expansion. Native boundary/exterior fields export as SVG 2 vector patterns with separate color and alpha composition. The shared SVG importer accepts shrinking-circle Pad, Repeat and Reflect fields, including point ends, and retains transparent regions outside ordinary SVG cones. Ordinary SVG point-focus boundary spreads retain their offset-weighted outside color and alpha; native XPS keeps its distinct endpoint rule. Pattern content retains its tile origin and stroke coverage offsets.
- `OfficeChartSnapshot` and chart rendering primitives shared by PDF and Office exporters.
- `OfficeRasterImage`, `OfficeRasterCanvas`, `OfficeRasterRenderTarget`, and `OfficeDrawingRasterRenderer` for shared dependency-free raster rendering.
- `OfficePngReader`, `OfficePngWriter`, and `OfficeJpegCodec` for PNG/JPEG paths that should not be reimplemented by document packages.
- `OfficeTiffCodec`, `OfficeWebpCodec`, and `OfficeRasterImageEncoder` for shared bounded TIFF, WebP decoding and lossless or lossy WebP encoding, and format-neutral raster output.
- Shared SVG formatting, primitive writing, image projection, text-block rendering, hatch-pattern, data-bar, and sparkline helpers.
- Drawing quality diagnostics for canvas bounds and text overlap checks.

Set `OfficeDrawingRasterRenderOptions.ThrowOnImageDecodeFailure` to `true` when every image must render. An unsupported image, failed optional codec, or decoded raster above `MaximumRasterPixels` then stops rendering with `NotSupportedException`. Successful image decoding happens in the drawing pass, including nested groups and patterns.

High-quality image minification shares the operation's `MaximumRasterPixels` budget with decoded images. The renderer charges its additional output buffer and floating-point sampling scratch before allocation. When that budget prevents an SVG safety comparison, structural findings remain available and the report identifies the unavailable visual inspection.

## Boundaries

- This package owns shared lifecycle contracts, persistence mechanics, drawing intent, ink/math models and renderers, raster buffers, SVG and raster encoding primitives, image projection, text layout helpers, chart drawing, and document-agnostic visual diagnostics.
- Security contracts live here so format packages can expose strongly typed opt-in APIs without depending on the optional cryptographic package. Core includes only the portable AES-CBC compatibility provider; CMS, XML DSig, certificate-chain, and private-key operations remain in `OfficeIMO.Security`.
- Word, Excel, PowerPoint, Visio, and PDF packages own source-document semantics: package parsing, layout policy, coordinate systems, style/theme resolution, and user-facing export APIs.
- Document packages should not add private ink or math ASTs, pixel engines, image encoders/decoders, SVG primitive writers, text wrapping engines, or duplicate image-transform loops when the behavior can reasonably live here.
- PDF keeps PDF-stream and page-writer behavior in `OfficeIMO.Pdf`; when it needs generic image-like drawing, vector descriptors, colors, chart snapshots, PNG helpers, or raster visual QA, it should use `OfficeIMO.Core` and the `OfficeIMO.Drawing` namespace.
- Unsupported or approximate source features belong in stable diagnostics from the adapter, not as silent omissions in a renderer.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: OfficeIMO code is MIT. The incorporated CodeGlyphX VP8 encoder remains Apache-2.0, and its forward transform retains the WebM BSD-3-Clause notice. See [third-party notices](THIRD-PARTY-NOTICES.md), the [CodeGlyphX license](Licenses/CodeGlyphX-LICENSE.txt), and the [WebM license](Licenses/libvpx-LICENSE.txt) and [patent grant](Licenses/libvpx-PATENTS.txt).
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** None.
- **OfficeIMO:** This is the shared foundation; it does not depend on another OfficeIMO runtime package.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.

### Inspected caller raster decoding

`OfficeRasterDecodeOptions.ImageCodec` supplies a trusted decoder for inspected
payloads outside the managed subset. The shared boundary validates resource limits,
frame selection and returned dimensions, preserves input bytes, and observes
cancellation before and after the callback. Static WebP decoding failures do not
fall through to a caller codec. Animated WebP can select frame zero and reports
discarded animation. Drawing exports retain caller provenance and distinguish
visible failure placeholders from decoded source pixels.

### Bounded AVIF still images

AVIF decoding accepts bounded, whole, untransformed 8/10-bit YUV420 color or monochrome grayscale items with reduced or full still-picture headers and optional same-depth full-range monochrome alpha. Full headers use one unlayered operating point, no timing/decoder model, and one shown key frame in a combined frame OBU. Primary grayscale items may use full or limited range; an auxiliary alpha plane must use full range. It produces eight-bit straight-alpha RGBA using the declared CICP range and supported non-constant-luminance matrix. It does not apply ICC, transfer-function or gamut transforms. Image grids, image sequences, twelve-bit coding, crop/rotation properties and applied film grain are outside this decoder contract. The original encoded buffer, reconstruction planes and final pixels share the retained-memory limit; cancellation and work limits apply through reconstruction and composition. Malformed selected items and limit failures cannot invoke a caller codec. A validated item with an unsupported color matrix or vertical/colocated chroma placement may use the explicitly supplied `ImageCodec`; managed YUV420 composition currently qualifies centered/unspecified chroma placement.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Inspect | 2 | 0 | 0 | 0 | 0 | 0 |
| Validate | 2 | 0 | 0 | 0 | 0 | 0 |
| Remove | 2 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Core` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
