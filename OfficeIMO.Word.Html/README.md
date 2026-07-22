# OfficeIMO.Word.Html - Word and HTML conversion

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Word.Html)](https://www.nuget.org/packages/OfficeIMO.Word.Html)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Word.Html?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Word.Html)

`OfficeIMO.Word.Html` converts between HTML and `OfficeIMO.Word` documents. It is for document-shaped HTML that should become editable Word content, and for Word documents that should be exported as HTML.

## Install

```powershell
dotnet add package OfficeIMO.Word.Html
```

## Quick start

```csharp
using OfficeIMO.Word;
using OfficeIMO.Html;
using OfficeIMO.Word.Html;

HtmlConversionDocument source = HtmlConversionDocument.Parse("<h1>Hello</h1><p>Body</p>");
using WordDocument document = source.ToWordDocument(new HtmlToWordOptions());

string html = document.ToHtml(new WordToHtmlOptions {
    IncludeDefaultCss = true,
    ExportFootnotes = true,
    ExportEndnotes = true,
    ExportComments = true
});

HtmlTextConversionResult export = document.ToHtmlResult();
Console.WriteLine(export.RequireValue());
```

HTML can also be appended to an existing body, header, or footer with `AddHtmlToBody`, `AddHtmlToHeader`, and `AddHtmlToFooter`.

Use the result API when conversion evidence matters:

```csharp
HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
HtmlToWordResult result = source.ToWordDocumentResult(options);
using WordDocument document = result.RequireValue();

foreach (HtmlDiagnostic diagnostic in result.Report.Diagnostics) {
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.LossKind}");
}
```

`HtmlConversionDocument` is the required source model. It keeps parsing, base-URI handling, resource policy, and source diagnostics in one owner. `HtmlToWordOptions.Limits` uses the shared `HtmlConversionLimits` contract; `MaxHtmlNodes`, `MaxHtmlDepth`, `MaxCssBytes`, and `MaxTotalCssBytes` remain forwarding properties for compatibility. `HtmlToWordOptions.StyleMissingHandler` scopes custom class mapping to one conversion.

## What it maps

- HTML headings, paragraphs, inline formatting, links, images, SVG, lists, tables, captions, form controls, notes, headers, footers, and sections into Word content where supported.
- Word document metadata, paragraph/run styles, lists, tables, images, SVG, footnotes, endnotes, disabled form controls, comments, headers, footers, and optional CSS back to HTML.
- Stylesheets, inline CSS, local/remote resources, image policies, resource limits, and diagnostics through explicit options.
- Document language metadata between HTML `lang` and Word settings.

## Import profiles

```csharp
var safeOptions = HtmlToWordOptions.CreateUntrustedHtmlProfile();
safeOptions.MaxHtmlNodes = 5000;
safeOptions.DiagnosticHandler = diagnostic =>
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Source}");

HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
using WordDocument safeDocument = source.ToWordDocument(safeOptions);
```

Remote images and stylesheets require the async API:

```csharp
HtmlToWordResult remote = await source.ToWordDocumentResultAsync(options, cancellationToken);
using WordDocument remoteDocument = remote.RequireValue();
```

- `CreateOfficeIMOProfile()` keeps the compatibility-oriented defaults.
- `CreateUntrustedHtmlProfile()` keeps external document resources offline by default and enables bounded conversion.
- `CreateTrustedDocumentProfile()` enables document-provided stylesheet links for known-good HTML while keeping resource validation.
- `new HtmlToWordOptions()` embeds data URI images only; use a trusted/compatibility profile or set `ImageProcessing = ImageProcessingMode.Embed` for trusted remote image fetching.
- Local file images are not loaded by default; use a trusted/compatibility profile or add `Uri.UriSchemeFile` to `AllowedImageUriSchemes` for trusted local files.

## Boundaries

- Word document modeling belongs in `OfficeIMO.Word`.
- HTML-to-Markdown ingestion belongs in `OfficeIMO.Markdown.Html`.
- HTML-to-PDF bridge behavior belongs in `OfficeIMO.Html.Pdf`.
- Full support matrices and roadmap detail belong in `Docs/`, not this README.

## Deeper docs

- [Word/HTML support matrix](../Docs/officeimo.word-html-support-matrix.md)
- [Word/HTML roadmap](../Docs/officeimo.word-html-roadmap.md)
- [OfficeIMO.Word](../OfficeIMO.Word/README.md)

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`; `net472` is included when building on Windows.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** Open XML SDK already used by the Word package; HTML DOM/CSS parsing comes through `OfficeIMO.Html`.
- **OfficeIMO:** `OfficeIMO.Word`, `OfficeIMO.Html`, and `OfficeIMO.Drawing`. The bidirectional mapping, resource policy, and diagnostics are first-party.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
