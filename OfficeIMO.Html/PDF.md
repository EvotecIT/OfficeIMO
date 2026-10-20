# HTML to PDF and paged images

`OfficeIMO.Html` owns static HTML layout, HTML-to-PDF conversion and PDF-backed image export. The public PDF types retain the `OfficeIMO.Html.Pdf` namespace. Install `OfficeIMO.Html`; use `OfficeIMO.Html.Pdf` when importing PDF as HTML.

## HTML to PDF

```csharp
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;

string html = """
<h1>Quarterly update</h1>
<p>Generated directly by OfficeIMO.</p>
<table>
  <tr><th>Area</th><th>Status</th></tr>
  <tr><td>PDF</td><td>Green</td></tr>
</table>
""";

HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
byte[] pdf = source.ToPdfBytes();
source.SaveAsPdf("quarterly-update.pdf");
```

`ToPdfBytes()` preserves the established print-paged behavior. Use an explicit
render request when the PDF should follow another layout contract:

```csharp
using OfficeIMO.Drawing;

var options = new HtmlToPdfOptions {
    ViewportWidth = 816,
    PageSize = OfficePageSizes.A4,
    Margins = HtmlRenderMargins.All(24)
};

HtmlPdfRenderRequestResult print = source.RenderToPdfResult(
    HtmlRenderRequest.Create(
        HtmlRenderIntentProfile.PrintPaged,
        HtmlRenderEncoder.Pdf,
        options));

HtmlPdfRenderRequestResult screenReflow = source.RenderToPdfResult(
    HtmlRenderRequest.Create(
        HtmlRenderIntentProfile.ScreenMediaPaged,
        HtmlRenderEncoder.Pdf,
        options));

HtmlPdfRenderRequestResult screenSnapshot = source.RenderToPdfResult(
    HtmlRenderRequest.Create(
        HtmlRenderIntentProfile.ScreenSnapshotPaged,
        HtmlRenderEncoder.Pdf,
        options));

byte[] pdf = screenSnapshot.ToBytes();
double firstSliceOffset = screenSnapshot.RenderResult.Surfaces[0].SourceOffsetY;
```

The same request may override CSS media, layout surface, or pagination through
`WithCssMedia(...)`, `WithLayoutSurface(...)`, and `WithPagination(...)`.
Use `WithLayout(...)` or `WithAxes(...)` when coupled geometry axes must change
atomically. Such a custom combination is reported as unqualified unless its effective axes
still exactly match the named profile.

Print paged applies print CSS and normal fragmentation. Screen-media paged keeps
screen CSS while reflowing into page sheets. Screen-snapshot paged preserves one
continuous screen composition and then slices it into fixed canvases. The latter
can split elements at page boundaries. The explicit result retains page order,
dimensions, source offsets, clipping, requested scale and background, provider
identity, and HTML loss diagnostics alongside the combined PDF conversion report.
For stitched output, `SourcePlacements` preserves every contributing source page
or slice and its output offset. Layout, projection, and PDF encoding share one
operation deadline.

For a wide print layout that should fit a fixed paper size, set a CSS layout
width explicitly. The PDF keeps the physical `PageSize`; content and margins are
scaled uniformly after layout. This changes wrapping and page breaks, so choose
the width for the source rather than assuming one width fits every page. CSS
media queries and responsive resources still use the captured `ViewportWidth`.

```csharp
var fitted = new HtmlToPdfOptions {
    PageSize = OfficePageSizes.A4,
    Margins = HtmlRenderMargins.All(24),
    HonorCssPageRules = false,
    PrintLayoutWidthCssPixels = 1200
};
byte[] fittedPdf = source.ToPdfBytes(fitted);
```

Fitting applies to print-paged reflow. When `HonorCssPageRules` is `true`, the
PDF keeps the resolved authored `@page` size and margins, and fits the chosen
layout width inside their printable area. For example, a source with a
`min-width: 1400px` print layout and `@page { size: 330mm 427mm; margin: .75in .5in }`
can use `PrintLayoutWidthCssPixels = 1400` without replacing the authored sheet.
An explicit width must exceed the printable area. By default, PDF conversion also fits
visible layout overflow, including fixed-width descendants, tables and root minimum
widths, within a uniform authored sheet. Hidden or scroll-clipped containers and
paint-only shadows do not expand the fitting width. Root overflow propagated from
`html` or `body` clips the final viewport and permits print fitting of the underlying
layout. Fixed-position boxes use the page area inside its margins and repeat on each page.
CSS media queries retain the
physical print context during reflow. Named or page-specific `@page` rules still need an explicit
`PrintLayoutWidthCssPixels`. Set `AutoFitWidePrintContent = false` to keep the
unscaled print layout. Automatic fitting includes positioned content outside the original
sheet; disable it when that content should remain outside a fixed print surface.
The retained render result uses the wider CSS page dimensions; the PDF is the
scaled paper-size artifact.

A prepared HTML report can be exported directly to a file:

```csharp
var report = HtmlConversionDocument.Load("service-review.html");
report.SaveAsPdf("service-review.pdf").RequireSuccess();
```

The [multi-format report example](../OfficeIMO.Examples/Converters/Html/HtmlMultiFormatReport.cs) also exports editable Word and Excel artifacts from the same prepared source.

## ZIP site bundles

Load a ZIP website through `OfficeIMO.Html.HtmlSiteBundle` and render it with the
same explicit PDF intent used for an HTML document:

```csharp
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;

using Stream source = File.OpenRead("saved-site.zip");
HtmlSiteBundle bundle = await HtmlSiteBundle.LoadAsync(source);
HtmlPdfRenderRequestResult result = await bundle.RenderToPdfResultAsync(
    HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf));
File.WriteAllBytes("saved-site.pdf", result.ToBytes());
```

The result retains the selected HTML surfaces and combined conversion report.
Archive images, CSS and fonts require `ResourcePolicy.AllowEmbeddedPackageResources`.
Local files and remote resources remain governed by their separate PDF permissions
and an explicitly supplied resolver; loading a ZIP does not enable either permission.
The [bundle loader](../OfficeIMO.Html/README.md#load-a-zip-site-bundle) owns entry
selection, input limits, integrity checks and archive resource identities.

## Review every rendered page

The capability gallery saves the input HTML, PDF, and image previews from one resolved
layout. Each artifact has a SHA-256 hash and its own diagnostics. Image evidence records
source page numbers, pixel dimensions, and format validation; PDF evidence runs the
configured readback checks against the exact serialized bytes.

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;

var gallery = new HtmlRenderCapabilityGalleryOptions(new HtmlCapabilityGalleryScenario(
    "quarterly-update", "Quarterly update", "Reports", "Paged report export")) {
    PreviewAllPages = true
};
gallery.PreviewFormats.Add(OfficeImageExportFormat.Webp);
gallery.PdfProofOptions.RequiredTextMarkers.Add("Quarterly update");
HtmlCapabilityGalleryManifest manifest = source.SaveRenderCapabilityGallery("review", gallery);
```

PNG and SVG are included by default; JPEG, TIFF, and WebP can also be selected.
The shared output-count, total-raster-pixel, and total-encoded-byte budgets apply
across all selected page/format previews. The rendering deadline includes preview
production; the cancellation overload also accepts a caller token.
`PreviewAllPages = false` retains the selected-page preview filenames and uses
`PreviewPageIndex`. Free-text `Expectations` are declarations, not test results.
Executed checks appear under each artifact's `Evidence.Checks`; a passed readback or
format check does not establish visual equivalence to a browser or source application.

## Data-driven reports and long tables

Keep data binding outside the renderer. Build finalized, encoded HTML with Razor, Scriban, Handlebars, or the template system already used by the application, then pass that HTML to `HtmlConversionDocument.Parse`. OfficeIMO does not execute JavaScript or template expressions while rendering.

In paged output, table cells measure and wrap through the shared layout engine. A leading `<thead>` and a trailing `<tfoot>` repeat on each table fragment. Use `<tfoot>` for a genuine repeating table footer; put a subtotal or grand-total card that must appear once after the table and give it `break-inside: avoid`. Top-aligned, auto-height tables in one horizontal flex row or one ordinary auto grid row can continue independently, including their repeated groups and following content. See the [parallel-table contract and limits](README.md) before using more complex fragmenting containers.

Ordinary rowspans can cross a safe row boundary without moving their whole row group to the next page. Authored cell, row, row-group and nested block keeps remain effective, and a cut must preserve text lines, atomic content and paragraph widows/orphans. The fragments retain the original cell identity and source alignment. Fragment-specific middle/bottom alignment, internal row breaks through spans and width-changing span reflow remain outside this bounded contract; changing geometry retains source layout with `HtmlRenderPagePseudoGeometryPending`.

Rectangular, fully opaque table backgrounds with automatic height continue through unused space after a safe row cut. Tables with `border-collapse: separate` carry their own solid left and right borders through that space; collapsed grid rules stay with their rows. Table paint stays behind row-group, row and cell colors. Row-group and row backgrounds stay beneath cell fills, including cells spanning later rows. Repeated headers retain that table backdrop beneath their own colors. Footer reservations remain inside the continuing backdrop, and the final fragment ends after its footer. Background continuation uses static, untransformed boxes without background images, shadows or additional clipping. When detached repetition cannot preserve ancestor opacity, transforms or paint clipping, the table retains its original header and footer paint instead of repeating it outside that effect. If pagination crosses that table, `HtmlRenderTableHeaderRepeatSuppressed` and `HtmlRenderTableFooterRepeatSuppressed` report the approximation; `RequireNoLoss` rejects it.

```html
<table>
  <thead><tr><th>Item</th><th>Qty</th><th>Amount</th></tr></thead>
  <tbody><!-- encoded rows from the application's data/template layer --></tbody>
</table>
<section class="totals" style="break-inside:avoid">
  Total: $7,239.46
</section>
```

For large output, prefer `SaveAsPdf(Stream, options)` over `ToPdfBytes()`, and select forward-only PDF object serialization when the destination need not be seekable:

```csharp
using OfficeIMO.Pdf;

var options = new HtmlToPdfOptions {
    PdfOptions = new PdfOptions {
        FileVersion = PdfFileVersion.Pdf17,
        ObjectSerializationMode = PdfObjectSerializationMode.ForwardOnly
    }
};

using Stream destination = File.Create("purchases.pdf");
PdfSaveResult saved = HtmlConversionDocument.Parse(html).SaveAsPdf(destination, options);
Console.WriteLine(saved.Serialization?.PeakRetainedPageContentBytes);
```

This bounds completed PDF page and object payloads and avoids buffering the final artifact inside OfficeIMO. HTML parsing, cascading, and layout still operate on the complete document so CSS page counters, named pages, cross-page fragmentation, and other whole-document rules remain correct. `PdfSerializationReport.IsForwardOnlyLayout` stays `false`; split independently generated documents at the application boundary when a workload requires a hard end-to-end heap ceiling.

## Paged-media and PDF semantics

Tagged output retains source hierarchy for supported text, images and rendered form controls. Inline images stay between their surrounding text, and local-frame content stays within its source-positioned frame scope. Anonymous or generated text and logical `ActualText` replacements retain their source ownership and sequence across paint layers. Flex content keeps that hierarchy when CSS `order`, reverse directions or pagination change where it is painted. Repeated table headers and continued list/table content retain their semantic ancestors across pages. Manual assistive-technology review remains a separate accessibility requirement.

The managed path supports named page rules, margin boxes, page counters, running strings, and running elements. It also maps headings and CSS bookmark controls to PDF outlines, maps supported semantic roles to the tagged structure tree, and marks repeated decorative margin content as PDF artifacts. Presentation MathML keeps its vector paint and searchable logical text inside a tagged `Formula` container. Its description follows the shared accessible-name resolution, then `alttext`, then the logical expression. Tagged formulas carry UTF-8 `application/mathml+xml` associated-file supplements. Exact original standalone MathML survives document cloning and unrelated edits; changed or recovered formulas carry current namespace-aware MathML with a normalization diagnostic. Hidden formulas and untagged output do not emit formula attachments. These associations select PDF 2.0, or PDF 1.7 when configured PDF/A-3 groundwork permits them. Association metadata alone does not establish PDF/A or PDF/UA conformance.

MathML operator spacing uses the [shared rendering contract](README.md#direct-html-rendering)
for both PDF and paged image output. Supported form and side spacing affect the
same retained drawing. Unsupported side syntax or detected authored operator presentation in unqualified
embellished contexts retains approximation diagnostics; source attachments do not make those visual
approximations lossless.

An explicit standalone PDF/UA-1 request cannot retain these Formula attachments: PDF/UA-1 requires PDF 1.7, while native structure-associated files require PDF 2.0. `RequireCompliance(PdfComplianceProfile.PdfUa1)` rejects that combination before writing output. Setting `FileVersion` to `Pdf17` does not override the attachment requirement. The converter never drops MathML or changes the requested profile automatically. Callers needing both PDF/UA-1 and the source supplements can explicitly configure PDF/A-3 groundwork alongside PDF/UA-1, provide `MathMlSourceModificationDate` and embedded fonts, and validate the resulting bytes against both profiles. A PDF 2.0 profile is a separate caller choice and requires its own qualification.

Column notes retain note bodies as `Note` elements with unique IDs. Linked call markers remain in their originating paragraph, and return links remain in the note. The logical body follows column flow; notes follow the body of the page on which they start, with continuations belonging to the same note. These structural guarantees do not replace assistive-technology reading-order review.

Set `HtmlToPdfOptions.MathMlSourceModificationDate` when an archival profile requires a modification date for embedded sources. The renderer does not invent dates.

An emitted tagged Formula MathML supplement can satisfy PDF/A-4f's embedded-file
requirement. For reproducible archival output, supply text and mathematical fonts
explicitly, set `MathMlSourceModificationDate`, request the profile, and validate the
exact saved PDF. See the [PDF compliance contract](../OfficeIMO.Pdf/README.md#generate-and-assess-validator-backed-pdfa)
for saved-artifact checks and external validation requirements.

The [math and archive qualification](../Build/Project/Evidence/2026-09-30/html-master-convergence/math-operator-print-20261020.json) records the selected
ordinary-row operator fixtures and tagged PDF/A-4f archive with explicit fonts,
complete page content, Formula supplement identity, native print comparison,
saved-PDF rendering, exact-artifact external validation and packaged public
consumers. Its declared fixture, font, platform and framework scope applies.

```html
<style>
  @page {
    margin: 18mm;
    @top-center { content: element(report-header, first) }
    @bottom-right { content: "Page " counter(page) " / " counter(pages) }
  }
  .running-header { position: running(report-header) }
  section { break-after: page; -officeimo-pdf-tag-type: Sect }
  section:last-child { break-after: auto }
  h1 { bookmark-level: 1; bookmark-state: open }
  .summary { display: grid; grid-template-rows: 32px 48px }
  .summary-body { display: grid; grid-column: 1; grid-row: 1 / 3; grid-template-rows: subgrid }
  .badge { clip-path: polygon(0 0, 100% 0, 88% 100%, 0 100%) }
</style>
<header class="running-header">Quarterly report</header>
<section>
  <h1>Overview</h1>
  <div class="summary">
    <div class="summary-body"><div>Revenue</div><div>Margin</div></div>
    <div class="badge">Approved</div>
  </div>
</section>
```

`bookmark-level:none` suppresses an automatic heading outline. `bookmark-label:"…"` changes its outline label, and `bookmark-state:open|closed` controls its initial state. `-officeimo-pdf-tag-type` accepts the declared PDF structure roles plus `artifact`/`none` for decorative content. Invalid values keep the normal semantic behavior and produce a typed warning.

Use strict mode when the input must stay inside the declared static contract, and require the PDF-stage report as well:

```csharp
var options = new HtmlToPdfOptions {
    FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss,
    PdfOptions = new OfficeIMO.Pdf.PdfOptions()
        .EnableTaggedPdfCatalogMarkers()
};

var result = HtmlConversionDocument.Parse(html).ToPdfDocumentResult(options);
result.RequireNoLoss();
result.Save("report.pdf").RequireNoLoss();
```

## Accessibility validation

HTML-to-PDF enables tagging by default. Use real headings, table headers with
`scope`, list elements, image descriptions, a document title and a language on
`html`. Accessible image names use the shared resolver, including
`aria-labelledby`, `aria-label`, `alt`, and SVG `title`. Explicitly decorative
images remain paint artifacts, including SVG with `role="presentation"` or
`role="none"`.

Validate the retained render result so source checks use the exact media,
viewport, caller CSS and resolved stylesheets used by the operation:

```csharp
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;

HtmlConversionDocument source = HtmlConversionDocument.Parse("""
    <!doctype html><html lang="en"><head><title>Service review</title></head>
    <body><h1>Service review</h1><p>All scheduled checks are complete.</p></body></html>
    """);
HtmlPdfRenderRequestResult rendered = source.RenderToPdfResult(
    HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf));
HtmlPdfAccessibilityValidationResult validation = HtmlPdfAccessibilityValidator.Validate(rendered);
if (!validation.IsValid) {
    throw new InvalidOperationException(string.Join(Environment.NewLine,
        validation.Issues.Select(issue => issue.Code + ": " + issue.Message)));
}
File.WriteAllBytes("service-review.pdf", rendered.ToBytes());
```

Validation does not load resources again or observe later changes to the caller's
options. Source image-name checks cover CSS-visible source elements before page
selection and clipping. Validate before editing the returned PDF. The older
`Validate(source, pdfBytes)` overload assumes default print CSS without caller or
externally resolved stylesheets; use the retained-result overload for generated
documents. `Validate(pdfBytes)` checks the artifact alone and cannot recover
source omissions that were not represented in its tags.

The [accessible report example](../OfficeIMO.Examples/Converters/Html/HtmlAccessibleReport.cs)
embeds the same regular and bold fonts for layout and PDF encoding, uses
`PdfOptions.UsePdfUa(...)`, and checks the retained result. Its
[HTML source](../OfficeIMO.Examples/Converters/Html/Content/Reports/accessibility-review.html)
contains two pages, ordered and unordered lists, row and column table headers,
a described figure, decorative SVG, an internal link, a labelled checkbox and
two-column reading content. Run it from the repository root:

```sh
dotnet run --project OfficeIMO.Examples/OfficeIMO.Examples.csproj -f net10.0 -- --html-accessible
```

The example writes `HtmlAccessibleReport-PdfUa1.pdf` and
`HtmlAccessibleReport-PdfUa2.pdf` to the examples runner's `Documents` directory.
Validate the exact saved files independently, for example:

```sh
verapdf -f ua1 HtmlAccessibleReport-PdfUa1.pdf
verapdf -f ua2 HtmlAccessibleReport-PdfUa2.pdf
```

The [reproducible report evidence](../Build/Project/Evidence/2026-10-20/html-pdf-accessibility.json)
records the saved-file hashes, fonts, independent profile results and visual
review. Manual assistive-technology acceptance remains separate.

The bounded OfficeIMO validator and independent machine validation have different
roles. Neither establishes the usefulness of descriptions or the complete
assistive-technology experience. Record the PDF hash, PDF reader and screen-reader
versions, then review heading navigation, table-header announcements, list order,
the figure description, the internal link, the checkbox label and the order of
the two reading groups. Check that the decorative rule and repeated page numbers
do not interrupt reading. Record failures and untested tasks explicitly.

## Fonts and text shaping

PDF and image export require glyph coverage in the selected embedded fonts. Missing glyphs produce a structured encoding failure, including under `AllowDiagnosedLoss`. Supply complete font files through `Fonts`, fallback packs or `@font-face`; use `TextShapingProvider` and `TextShapingLanguage` for scripts outside the managed shaping subset. Installed-font fallback depends on the host. Synthetic provider tests establish the API seam, while native script appearance needs separate rendered evidence.

Library-selected default fonts account for visible flow and referenced page-margin strings and running elements before authoritative pagination. Absolute and fixed sources use the same preparation. Unreferenced running-string assignments do not force a different font choice. Explicit font registrations and resource-policy restrictions retain their existing ownership; font discovery shares the layout operation budget and cancellation.

Page-margin font declarations retain authored order and `!important`, including `font` shorthand resets and supported container-unit lengths in `font-size` and `line-height`. In the horizontal margin context, `cqw`, `cqh`, `cqi`, `cqb`, `cqmin` and `cqmax` use the configured print paper as their viewport fallback. This reference is the caller’s `PageSize` before an authored `@page size` changes the output page; body query containers do not supply margin-box dimensions. Calculated font lengths enter the cascade before their nonnegative range is applied in the actual layout context; a zero margin font size does not paint text. Quoted margin content uses the shared CSS escape decoder, including Unicode escapes and `\A` line breaks. Use the same print-paper setting when comparing browser output. This fallback follows the qualified browser behaviour; it does not establish complete paged-media or vertical-writing conformance.

The first-party font engine loads policy-approved TrueType-glyf OpenType, WOFF 1, CFF/CFF2, and TrueType or CFF2 variable fonts. Single-face WOFF 2 decoding is built in on .NET 8 and newer; extract and register individual faces from WOFF 2 font collections. Static faces remain eligible for PDF embedding, including supported OpenType features. Color glyphs and feature-bearing color fonts retain palette-aware outlines. Variable instances and text that requires vector paint use outlines plus logical `ActualText`, preserving extraction and accessibility.

Font resources may use a recognized font MIME type or the generic `application/octet-stream` and `binary/octet-stream` types. The font decoder still validates their bytes; a generic type does not make an invalid font usable. These routes share resource-count, source-byte, decoded-byte, operation and cancellation limits.

PDF outline expansion remains bounded. The selected program must implement
`IOfficeBoundedFontProgram`; conversion carries cancellation into contour expansion
and enforces `MaxOutlinedTextCharactersPerRun` plus the operation-wide
`MaxOutlinedTextPathCommands` budget. The defaults are 16,384 UTF-16 characters per
run and 1,000,000 path commands per conversion. A run above the character limit
still fails. If the path-command budget is exhausted, conversion stops outlining
and uses PDF text for the remaining runs, with a loss-bearing
`HtmlPdfFontOutlineBudgetApproximated` warning because font appearance or shaping
may differ. Raise the outline budget only for trusted inputs.
Outline-only private-use symbols encountered after exhaustion are omitted with
`HtmlPdfUnavailablePrivateUseGlyphOmitted`, rather than passed to an incompatible
PDF text font.

HTML body text and SVG text share font discovery, including nested SVG viewports, effect groups and tiling patterns. Fonts registered through `options.Fonts` are available to layout and PDF encoding in those groups. Automatic installed-font fallback follows `ResourcePolicy` for both HTML and SVG text; `CreatePortableDeterministic()` requires explicit in-memory or allowed document fonts. Already registered style faces in the requested family remain eligible when the nominal face lacks a glyph. If the selected fonts cannot encode painted text, PDF conversion throws `HtmlConversionException` with the encoding diagnostics before returning a generated document. Logical `ActualText` and accessibility labels do not require painted-glyph coverage.

No font-program package or license key is required. Select variable-font axes on the font collection before conversion:

```csharp
using System.Collections.Generic;
using OfficeIMO.Drawing;
using OfficeIMO.Html.Pdf;

var options = new HtmlToPdfOptions();
options.Fonts.FontVariationResolver = request =>
    request.FamilyName == "Report Variable"
        ? new Dictionary<string, float> {
            ["wght"] = 720,
            ["wdth"] = 110
        }
        : null;

options.MaxOutlinedTextCharactersPerRun = 16_384;
options.MaxOutlinedTextPathCommands = 1_000_000;
```

HTML uses the shared Core/PDF owners for bidirectional ordering, line breaking and bounded Arabic per-letter joining. Combining-mark attachment and required lam-alef ligatures need a complete shaping provider. The managed fallback preserves logical searchable text and reports approximation loss; strict mode rejects that fallback.

For complete OpenType GSUB/GPOS shaping, add the optional `OfficeIMO.Drawing.HarfBuzz` adapter and assign its provider to the same options object used by PDF and image output:

```csharp
using OfficeIMO.Drawing.HarfBuzz;

var options = new HtmlToPdfOptions {
    TextShapingProvider = OfficeHarfBuzzTextShapingProvider.Instance,
    TextShapingLanguage = "ar"
};
options.Fonts.Add("Report Arabic", File.ReadAllBytes("ReportArabic.ttf"));
```

Configured providers can retain an ordinary horizontal RTL Arabic word across eligible inline paint boundaries. Each piece must share the selected bounded monochrome font program, collection face and variation coordinates, with matching font size, language, direction and OpenType features. Letter spacing must be zero, source partitions must preserve complete logical clusters, and slice advances must add up to the complete word. Glyph IDs, offsets, advances and continuation glyphs stay with their logical source through measurement and painting. Forced breaks, bidi controls, emergency word wrapping, inline insets and differing shaping settings lie outside this retained-context subset.

Paragraph bidi also supplies the direction for neutral punctuation. With zero letter and word spacing, eligible horizontal RTL slices retain the provider result requested with that resolved direction, so later output does not infer a new direction from punctuation alone. Font, provider, color-glyph or cluster-retention failures in an otherwise eligible context report `HtmlRenderComplexTextShapingUnsupported`.

HTML PDF output paints admitted glyphs as vector outlines with logical `ActualText` for extraction and accessibility, alongside supported decoration paint. A positive-advance run with no visible outline still keeps its logical text and eligible link bounds. Absence of ink alone does not trigger fresh shaping. Provider acceptance and a clean report do not establish native appearance for every font or script.

`HtmlRenderPage.CreateDrawing()` also carries admitted glyphs into Drawing output. Generic PDF export through `PdfPageCanvas.Drawing(...)` accepts retained text as unwrapped positioned monochrome slices; attached text decorations, padding, indentation, outline strokes and specialized metric geometry are outside that route and fail explicitly. Separate decoration shapes already present in the Drawing use normal shape output.

Hosts that ship the OfficeIMO portable browser font set can activate its layout
aliases, fallback order, PDF embedding, and substitutions through one reusable
profile. The profile does not embed or download fonts; the host remains responsible
for supplying pinned and licensed bytes:

```csharp
HtmlPortableBrowserFontProfile fonts = HtmlPortableBrowserFontProfile.Create(
    "my-portable-fonts-v1",
    fileName => File.ReadAllBytes(Path.Combine(fontDirectory, fileName)));

HtmlToPdfOptions options = fonts.CreateHtmlOptions(
    OfficeHarfBuzzTextShapingProvider.Instance);
```

The same `options` instance can drive screen images, browser-print PDF, and
screen-to-page PDF so layout and encoded output use the same fonts.

If no configured provider accepts a run that requires provider-owned complex shaping, OfficeIMO retains logical searchable text and reports `HtmlRenderComplexTextShapingUnsupported`; strict mode rejects that fallback, including outlined-font output.

Install `OfficeIMO.Mhtml.Pdf` when the source is an MHT/MHTML archive. That bridge adds MIME parsing and embedded-resource resolution without putting `OfficeIMO.Email` into ordinary HTML/PDF applications.

Naming is consistent across the direct output APIs:

- `ToPdfBytes()` returns encoded bytes.
- `ToPdfDocument()` returns the first-party PDF model.
- `ToPdfDocumentResult()` returns the PDF model plus diagnostics.
- `ExportImage()` and `ExportImages()` return image output, dimensions, and diagnostics.
- `SaveAsPdf(path)` and `SaveAsPdf(stream)` write to a destination.
- Async counterparts use the same names with `Async` appended.

## One options shape for PDF and all image formats

`HtmlToPdfOptions` derives from `HtmlRenderOptions`, so one configured instance can drive PDF and all five direct image outputs.

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;

var options = new HtmlToPdfOptions {
    PageSize = OfficePageSizes.A4,
    Margins = HtmlRenderMargins.All(32),
    DefaultFontFamily = "Arial",
    BackgroundColor = OfficeColor.White,
    Scale = 1.5,
    PdfOptions = new OfficeIMO.Pdf.PdfOptions()
        .EnableTaggedPdfCatalogMarkers()
};
options.AdditionalStylesheets.Add("@page { margin: 18mm }");

HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
byte[] pdf = source.ToPdfBytes(options);
byte[] png = source.ToPng(options);
byte[] jpeg = source.ToJpeg(options);
byte[] tiff = source.ToTiff(options);
string svg = source.ToSvg(options);
byte[] webp = source.ToWebp(options);
```

`ToPdfBytes()` and default image export use paged print layout. Image export serializes that layout through the PDF owner and renders its page scenes, including multilingual font fallback and margin boxes. Select `Mode = HtmlRenderMode.Continuous`, `.Continuous()`, or `ScreenFullPage` explicitly for an unpaginated screenshot. Image scale uses 96 CSS pixels per inch; `TargetDpi` sets the requested raster resolution.

## Diagnostics and external resources

Options are reusable configuration and are not mutated with operation results. Request a result when diagnostics matter.

```csharp
var options = new HtmlToPdfOptions {
    ResourcePolicy = PdfResourcePolicy.CreateTrustedHost(),
    ResourceResolver = (request, cancellationToken) =>
        Task.FromResult<HtmlResolvedResource?>(null)
};

HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
var result = await source.ToPdfDocumentResultAsync(options);
var pngResult = await source.ExportImageAsync(OfficeImageExportFormat.Png, options);
var svgResult = await source.ExportImageAsync(OfficeImageExportFormat.Svg, options);
await result.SaveAsync("report.pdf");

foreach (var warning in result.Report.Warnings) {
    Console.WriteLine($"{warning.Code}: {warning.Message}");
}

foreach (var diagnostic in pngResult.Diagnostics) {
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
}

foreach (var diagnostic in svgResult.Diagnostics) {
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
}
```

Resource resolution is opt-in. `PdfResourcePolicy` is the host-access gate for local files, remote resolver calls, data URIs, embedded package resources, and installed fonts. `HtmlUrlPolicy` independently validates URL syntax and schemes; timeouts, byte limits, count limits, and stylesheet-depth limits inherited from `HtmlRenderOptions` bound resources after access is granted. HTML defaults to `PdfResourcePolicy.CreateLibraryFontDefault()`: library-selected installed fonts, bounded data URIs and MHTML package parts are allowed; document-selected host fonts and local or remote resolver calls are disabled. Set `AllowDocumentFontEmbedding` explicitly to use installed families named by the document. Portable deterministic mode disables installed-font discovery explicitly.

Under the default policy, MathML can use the renderer's fixed mathematical font
families when a supported installed face is available. CSS cannot use that fallback
list to select another named host font. `AllowDocumentFontEmbedding` still gates
document-selected host families and installed system-UI lists. Disable
`AllowSystemFontEmbedding` or `AllowSystemFontFallback` to prevent installed math
font loading, or supply an in-memory `math` face for portable output. Font loading
retains the shared source-byte, decoded-byte and cancellation limits.

## Command-line conversion

Install `OfficeIMO.Tool` when a script or build pipeline is the desired surface:

```powershell
dotnet tool install --global OfficeIMO.Tool
officeimo html convert report.html --output report.pdf
officeimo html convert archive.mhtml --output archive.pdf
officeimo html capabilities --format json
```

The command uses the same renderer and capability catalog as the .NET API. Host fonts and local and remote resource reads are disabled by default; embedded data and bounded MHTML resources remain available. Supply `--font-family` and `--font-regular` for text outside the standard PDF encoding; `convert` and `render` use the same supplied font family. Standard input/output, caller stylesheets, page limits, atomic file replacement, and explicit embedded fonts for PDF/UA-ready artifacts are supported.

## Explicit document projections

Direct rendering is the normal HTML-to-PDF path. If the desired target is an editable Word document or a Markdown AST, request that target explicitly and then use its PDF converter:

```csharp
using OfficeIMO.Markdown.Html;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Pdf;

HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
byte[] markdownPdf = source.ToMarkdownDocument().ToPdfBytes();

using var word = source.ToWordDocument();
byte[] wordPdf = word.ToPdfBytes();
```

Those adapters remain separate packages and are not dependencies of `OfficeIMO.Html`.

