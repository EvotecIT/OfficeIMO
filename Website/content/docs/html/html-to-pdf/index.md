---
title: Convert HTML to PDF
description: "Convert HTML pages and reports to PDF in .NET without a browser, with page size, margins, extra CSS, headers and footers, and local images and fonts."
order: 43
search_terms:
  - html to pdf
  - htm to pdf
  - html report to pdf
  - css print to pdf
  - pdf from html without browser
packages:
  - OfficeIMO.Html
  - OfficeIMO.Html.Pdf
cmdlets:
  - Convert-OfficeDocument
  - New-OfficeHtmlPdfOptions
commands:
  - officeimo html convert
related:
  - /docs/html/render-and-convert/
  - /docs/html/mhtml-and-site-bundles/
  - /docs/markdown/convert/
limits: The renderer lays out static HTML and CSS; it does not run JavaScript, so pages that build their content in script need a browser capture first.
verified: 2026-11-14
---

`OfficeIMO.Html.Pdf` lays out HTML and CSS with the managed OfficeIMO renderer and writes the result through `OfficeIMO.Pdf`. No browser, Office or native library is involved. Parse or load the page into an `HtmlConversionDocument`, then call `SaveAsPdf` or `ToPdf`. Options live on `HtmlPdfExportOptions`, which extends the `HtmlRenderOptions` used for image export, so one options object can drive both.

Build the HTML with whatever template engine you already use (Razor, Scriban, string templates) and pass the finished markup in; OfficeIMO does not evaluate script or template expressions.

```sh
dotnet add package OfficeIMO.Html.Pdf
```

## Convert an HTML string or file to PDF

### C#

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
source.SaveAsPdf("quarterly-update.pdf");

byte[] bytes = source.ToPdf().ToBytes();
Console.WriteLine($"{bytes.Length} bytes");
```

### PowerShell

```powershell
Convert-OfficeDocument -Path .\Report.html -OutputPath .\Report.pdf

'<h1>Review</h1>', '<p>Approved.</p>' |
    Convert-OfficeDocument -Text { $_ } -From Html -To Pdf -OutputPath .\Review.pdf
```

### CLI

```sh
officeimo html convert report.html --output report.pdf
```

## Set page size, margins and extra CSS

`PageSize` and `Margins` set the default sheet. Authored `@page` rules win unless you set `HonorCssPageRules = false`. Stylesheets in `AdditionalStylesheets` are applied after the page's own CSS, which is the place for house styles.

### C#

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;

var options = new HtmlPdfExportOptions {
    PageSize = OfficePageSizes.A4,
    Margins = HtmlRenderMargins.All(36),
    DefaultFontFamily = "Arial",
    MaxPageCount = 200
};
options.AdditionalStylesheets.Add("body { color: hsl(215 35% 20%) } h1 { border-bottom: 1px solid #ccc }");

HtmlConversionDocument
    .Parse("<h1>Invoice</h1><p>Thank you for your order.</p>")
    .SaveAsPdf("invoice.pdf", options);
```

### PowerShell

```powershell
$options = New-OfficeHtmlPdfOptions -Margin 2cm -Font 'Arial'
Convert-OfficeDocument -Path .\Report.html -To Pdf -OutputPath .\Report.pdf -Options $options
```

### CLI

```sh
officeimo html convert report.html --output report.pdf --stylesheet print.css --max-pages 200
```

## Add page numbers, headers and footers

Paged output supports `@page` margin boxes with `counter(page)` and `counter(pages)`, and running elements with `position: running(name)` and `content: element(name)`. Repeated margin content is tagged as a PDF artifact, so it does not clutter the document structure. A leading `<thead>` repeats on every page a table spans.

### C#

```csharp
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;

string html = """
<style>
  @page {
    margin: 20mm;
    @top-left { content: "Purchase report"; font-size: 9pt; }
    @bottom-right { content: "Page " counter(page) " of " counter(pages); font-size: 9pt; }
  }
</style>
<h1>Purchases</h1>
<table>
  <thead><tr><th>Item</th><th>Qty</th></tr></thead>
  <tbody>
    <tr><td>Paper</td><td>12</td></tr>
    <tr><td>Toner</td><td>3</td></tr>
  </tbody>
</table>
""";

HtmlConversionDocument.Parse(html).SaveAsPdf("purchases.pdf");
```

## Load local images, stylesheets and fonts

By default, conversion reads no local files and makes no network calls; data URIs still work. For your own trusted templates, load the file with a base URI and allow local files through the PDF resource policy.

### C#

```csharp
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;

Directory.CreateDirectory("site");
File.WriteAllText(Path.Combine("site", "site.css"), "h1 { color: navy }");
File.WriteAllText(Path.Combine("site", "index.html"),
    "<link rel=\"stylesheet\" href=\"site.css\"><h1>Team handbook</h1>");

HtmlConversionDocument page = HtmlConversionDocument.Load(Path.Combine("site", "index.html"), new HtmlConversionDocumentOptions {
    Trust = HtmlInputTrust.Trusted,
    BaseUri = new Uri(Path.GetFullPath("site") + Path.DirectorySeparatorChar)
});

page.SaveAsPdf("handbook.pdf", new HtmlPdfExportOptions {
    ResourcePolicy = PdfResourcePolicy.CreateTrustedHost()
});
```

### PowerShell

```powershell
$options = New-OfficeHtmlPdfOptions -IncludeLocalResources
Convert-OfficeDocument -Path .\site\index.html -To Pdf -OutputPath .\handbook.pdf -Options $options
```

### CLI

The CLI reads no local or remote resources. Supply fonts and stylesheets explicitly:

```sh
officeimo html convert index.html --output handbook.pdf --stylesheet site.css --font-family "Noto Sans" --font-regular NotoSans-Regular.ttf --font-bold NotoSans-Bold.ttf
```

## Check what the conversion changed

The PDF result carries a conversion report. Warnings name images that did not load, fonts that were substituted and CSS that was simplified. Set `FidelityPolicy = HtmlRenderFidelityPolicy.FailOnLoss` to stop instead of writing an approximate PDF.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;

HtmlConversionDocument source = HtmlConversionDocument.Parse(
    "<h1>Status</h1><img src=\"https://example.com/chart.png\" alt=\"Chart\">");

OfficeConversionReport report = source.SaveAsPdf("status.pdf", new HtmlPdfExportOptions());

Console.WriteLine($"Loss reported: {report.HasLoss}");
foreach (OfficeDiagnostic diagnostic in report.Diagnostics) {
    Console.WriteLine($"{diagnostic.Severity} {diagnostic.Code}: {diagnostic.Message}");
}
```

## Write large reports to a stream

For long tables, write to a stream with forward-only object serialization so OfficeIMO does not buffer the finished file.

### C#

```csharp
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;

var rows = string.Concat(Enumerable.Range(1, 2000).Select(i => $"<tr><td>Item {i}</td><td>{i * 3}</td></tr>"));
string html = $"<table><thead><tr><th>Item</th><th>Amount</th></tr></thead><tbody>{rows}</tbody></table>";

var options = new HtmlPdfExportOptions {
    PdfOptions = new PdfCreateOptions {
        ObjectSerializationMode = PdfObjectSerializationMode.ForwardOnly
    }
};

using Stream destination = File.Create("large-report.pdf");
HtmlConversionDocument.Parse(html).SaveAsPdf(destination, options);
```
