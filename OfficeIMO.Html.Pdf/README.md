# OfficeIMO.Html.Pdf

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Html.Pdf)](https://www.nuget.org/packages/OfficeIMO.Html.Pdf)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Html.Pdf?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Html.Pdf)

`OfficeIMO.Html.Pdf` converts HTML directly to PDF with the same first-party layout scene used by HTML-to-PNG/JPEG/TIFF/SVG/WebP. It also converts PDF to semantic or positioned-review HTML.

The HTML-to-PDF path has no browser process, Office automation, Markdown bridge, Word bridge, or new external dependency.

## Install

```powershell
dotnet add package OfficeIMO.Html.Pdf
```

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
byte[] pdf = source.ToPdf();
source.SaveAsPdf("quarterly-update.pdf");
```

Install `OfficeIMO.Mhtml.Pdf` when the source is an MHT/MHTML archive. That bridge adds MIME parsing and embedded-resource resolution without putting `OfficeIMO.Email` into ordinary HTML/PDF applications.

Naming is consistent across the direct output APIs:

- `ToPdf()` returns encoded bytes.
- `ToPdfDocument()` returns the first-party PDF model.
- `ToPdfDocumentResult()` returns the PDF model plus diagnostics.
- `ExportImage()` and `ExportImages()` return image output, dimensions, and diagnostics.
- `SaveAsPdf(path)` and `SaveAsPdf(stream)` write to a destination.
- Async counterparts use the same names with `Async` appended.

## One options shape for PDF and all image formats

`HtmlPdfSaveOptions` derives from `HtmlRenderOptions`, so one configured instance can drive PDF and all five direct image outputs.

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;

var options = new HtmlPdfSaveOptions {
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
byte[] pdf = source.ToPdf(options);
byte[] png = source.ToPng(options);
byte[] jpeg = source.ToJpeg(options);
byte[] tiff = source.ToTiff(options);
string svg = source.ToSvg(options);
byte[] webp = source.ToWebp(options);
```

PDF always uses paged layout. Image output honors the selected continuous or paged render mode and page index.

## Diagnostics and external resources

Options are reusable configuration and are not mutated with operation results. Request a result when diagnostics matter.

```csharp
var options = new HtmlPdfSaveOptions {
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

Resource resolution is opt-in. `PdfResourcePolicy` is the host-access gate for local files, remote resolver calls, data URIs, embedded package resources, and installed fonts. `HtmlUrlPolicy` independently validates URL syntax and schemes; timeouts, byte limits, count limits, and stylesheet-depth limits inherited from `HtmlRenderOptions` bound resources after access is granted. The balanced default allows installed fonts plus bounded data URIs and MHTML package parts, but does not call local or remote resolvers. Portable deterministic mode disables installed-font discovery explicitly.

## Command-line conversion

Install `OfficeIMO.Tool` when a script or build pipeline is the desired surface:

```powershell
dotnet tool install --global OfficeIMO.Tool
officeimo html convert report.html --output report.pdf
officeimo html convert archive.mhtml --output archive.pdf
officeimo html capabilities --format json
```

The command uses the same renderer and capability catalog as the .NET API. Local and remote resource reads are disabled by default; embedded data and bounded MHTML resources remain available. Standard input/output, caller stylesheets, page limits, atomic file replacement, and explicit embedded fonts for PDF/UA-ready artifacts are supported.

## Explicit document projections

Direct rendering is the normal HTML-to-PDF path. If the desired target is an editable Word document or a Markdown AST, request that target explicitly and then use its PDF converter:

```csharp
using OfficeIMO.Markdown.Html;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Pdf;

HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
byte[] markdownPdf = source.ToMarkdownDocument().ToPdf();

using var word = source.ToWordDocument();
byte[] wordPdf = word.ToPdf();
```

Those adapters remain separate packages and are not dependencies of `OfficeIMO.Html.Pdf`.

## PDF to HTML

```csharp
using OfficeIMO.Html.Pdf;

string semantic = PdfHtmlConverterExtensions.ToHtml("quarterly-update.pdf", new PdfHtmlSaveOptions {
    Profile = PdfHtmlProfile.Semantic
});

PdfHtmlConverterExtensions.SaveAsHtml("quarterly-update.pdf", "quarterly-review.html", new PdfHtmlSaveOptions {
    Profile = PdfHtmlProfile.PositionedReview,
    IncludeLinkAnnotations = true,
    IncludeFormWidgets = true,
    ImageExportMode = PdfHtmlImageExportMode.EmbeddedDataUri
});
```

PDF-to-HTML profiles describe how an existing PDF is projected to HTML. They are unrelated to HTML-to-PDF, which has one direct rendering path.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** None beyond AngleSharp/AngleSharp.Css already isolated in `OfficeIMO.Html`; no browser process or native HTML renderer.
- **OfficeIMO:** `OfficeIMO.Html`, `OfficeIMO.Pdf`, and `OfficeIMO.Core` own layout, rendering, reverse projection, and reports.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
