# OfficeIMO.Html.Pdf

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Html.Pdf)](https://www.nuget.org/packages/OfficeIMO.Html.Pdf)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Html.Pdf?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Html.Pdf)

`OfficeIMO.Html.Pdf` converts HTML directly to PDF with the same first-party layout scene used by HTML-to-PNG and HTML-to-SVG. It also converts PDF to semantic or positioned-review HTML.

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

byte[] pdf = html.ToPdf();
html.SaveAsPdf("quarterly-update.pdf");
```

Naming is consistent across the direct output APIs:

- `ToPdf()` returns encoded bytes.
- `ToPdfDocument()` returns the first-party PDF model.
- `ToPdfResult()` returns the PDF model plus diagnostics.
- `ToPngResult()` and `ToSvgResult()` return image output, dimensions, and diagnostics; plural forms return every page.
- `SaveAsPdf(path)` and `SaveAsPdf(stream)` write to a destination.
- Async counterparts use the same names with `Async` appended.

## One options shape for PDF, PNG, and SVG

`HtmlPdfSaveOptions` derives from `HtmlRenderOptions`, so one configured instance can drive all three direct outputs.

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;

var options = new HtmlPdfSaveOptions {
    PageSize = OfficePageSizes.A4,
    Margins = HtmlRenderMargins.All(32),
    DefaultFontFamily = "Arial",
    BackgroundColor = OfficeColor.White,
    Scale = 1.5
};

byte[] pdf = html.ToPdf(options);
byte[] png = html.ToPng(options);
string svg = html.ToSvg(options);
```

PDF always uses paged layout. PNG and SVG honor the selected render mode and page index.

## Diagnostics and external resources

Options are reusable configuration and are not mutated with operation results. Request a result when diagnostics matter.

```csharp
var options = new HtmlPdfSaveOptions {
    ResourceResolver = (request, cancellationToken) =>
        Task.FromResult<HtmlResolvedResource?>(null)
};

var result = await html.ToPdfResultAsync(options);
var pngResult = await html.ToPngResultAsync(options);
var svgResult = await html.ToSvgResultAsync(options);
await result.SaveAsync("report.pdf");

foreach (var warning in result.ConversionReport.Warnings) {
    Console.WriteLine($"{warning.Code}: {warning.Message}");
}

foreach (var diagnostic in pngResult.Diagnostics) {
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
}

foreach (var diagnostic in svgResult.Diagnostics) {
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
}
```

Resource resolution is opt-in and controlled by `HtmlUrlPolicy`, timeouts, byte limits, count limits, and stylesheet-depth limits inherited from `HtmlRenderOptions`.

## Explicit document projections

Direct rendering is the normal HTML-to-PDF path. If the desired target is an editable Word document or a Markdown AST, request that target explicitly and then use its PDF converter:

```csharp
using OfficeIMO.Markdown.Html;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Pdf;

byte[] markdownPdf = html.ToMarkdownDocument().ToPdf();

using var word = html.ToWordDocument();
byte[] wordPdf = word.ToPdf();
```

Those adapters remain separate packages and are not dependencies of `OfficeIMO.Html.Pdf`.

## PDF to HTML

```csharp
using OfficeIMO.Html.Pdf;

string semantic = PdfHtmlConverter.ToHtml("quarterly-update.pdf", new PdfHtmlSaveOptions {
    Profile = PdfHtmlProfile.Semantic
});

PdfHtmlConverter.SaveAsHtml("quarterly-update.pdf", "quarterly-review.html", new PdfHtmlSaveOptions {
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
