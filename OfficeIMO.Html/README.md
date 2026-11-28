# OfficeIMO.Html

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Html)](https://www.nuget.org/packages/OfficeIMO.Html)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Html?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Html)

Parse, inspect, edit and lay out HTML in .NET, without a browser. `OfficeIMO.Html` is the shared HTML layer behind OfficeIMO's HTML conversions: trust-aware parsing, owned document snapshots you can query and edit, a URL and resource policy, and a browser-free layout engine that the PDF and image exporters paint.

## Install

```sh
dotnet add package OfficeIMO.Html
```

Add [OfficeIMO.Html.Pdf](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Html.Pdf/README.md) for PDF and PNG, JPEG, TIFF, SVG or WebP output.

## Quick start

Parse a page, add table cells, and save the edited HTML:

<!-- powerforge:code source="Build/DocumentationExamples/Html/Program.cs" region="QuickStart" -->
```csharp
using OfficeIMO.Html;
using OfficeIMO.Html.Dom;

HtmlDocument source = HtmlDocumentEngine.Default.ParseDocument(
    "<h1>Service review</h1><table><tbody><tr id='items'><th>Item</th></tr></tbody></table>");
HtmlElement row = source.QuerySelector("#items")!;
HtmlDocumentFragment cells = HtmlDocumentEngine.Default.ParseFragment(
    "<td>Certificates</td><td>Renewed</td>", row);

HtmlDocument edited = source.Edit(document => {
    document.QuerySelector("#items")!.AppendChild(document.ImportNode(cells));
});

HtmlConversionDocument report = HtmlConversionDocument.FromDocument(edited);
File.WriteAllText("Service-review.html", report.ExportSourceHtml());
```
<!-- /powerforge:code -->

`HtmlDocument` snapshots are immutable; `Edit` returns a new one. `HtmlConversionDocument` is the shared input for the Word, Markdown, Excel, PowerPoint, RTF, PDF and image converters.

## What it does

- **Parse safely** with trust profiles: the untrusted default rejects local-file navigation, fetches nothing by itself, and applies shared limits to input size, nodes, CSS and selectors before any converter allocates.
- **Query and edit** owned documents with selectors, fragments and `Edit`, through a replaceable parser provider (AngleSharp by default).
- **Resolve resources** with one URL policy, base URI handling, charset detection, responsive image selection and an operation-scoped resource session with budgets, timeouts and caching.
- **Load ZIP site bundles** with `HtmlSiteBundle`, serving archived stylesheets, fonts and images without extracting files.
- **Lay out** continuous and paged HTML without a browser, with CSS, MathML, embedded fonts and Arabic, Hebrew and Latin shaping.
- **Apply a theme** to plain HTML with `WithTheme`.
- **Report** conversions with stable diagnostics, target preflight and fidelity scoring.
- **Inspect concealed content and provenance** before HTML reaches an AI pipeline.

The [HTML guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.html-guide.md) shows each of these with examples.

## Converting to other formats

| Target | Package |
| --- | --- |
| PDF and images | [OfficeIMO.Html.Pdf](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Html.Pdf/README.md) |
| Word, both directions | [OfficeIMO.Word.Html](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Word.Html/README.md) |
| Markdown | [OfficeIMO.Markdown.Html](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Markdown.Html/README.md) |
| Excel tables | [OfficeIMO.Excel.Html](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Excel.Html/README.md) |
| PowerPoint | [OfficeIMO.PowerPoint.Html](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.PowerPoint.Html/README.md) |
| RTF, both directions | [OfficeIMO.Html.Rtf](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Html.Rtf/README.md) |
| MHTML archives | [OfficeIMO.Mhtml](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Mhtml/README.md) |

## Limits

OfficeIMO.Html doesn't run scripts. The [HTML support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.html-support-matrix.md) lists CSS and layout coverage per target, and the [HTML engine design](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.html-engine-design.md) explains how parsing, layout and the providers fit together.

## Learn more

- [OfficeIMO.Html on officeimo.com](https://officeimo.com/products/html/) and the [HTML docs](https://officeimo.com/docs/html/).
- [HTML guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.html-guide.md) and the [multi-format report example](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Examples/Converters/Html/HtmlMultiFormatReport.cs).
- Command line: `officeimo convert page.html page.pdf` with [OfficeIMO.Tool](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Tool/README.md).

## Targets and dependencies

- Targets: `netstandard2.0`, `net8.0` and `net10.0`, plus `net472` on Windows.
- Dependencies: `OfficeIMO.Core`, `OfficeIMO.Html.Core` and `OfficeIMO.Html.AngleSharp`, with [AngleSharp](https://github.com/AngleSharp/AngleSharp) and AngleSharp.Css for parsing and `System.Text.Encoding.CodePages` for legacy web encodings.
- License: MIT.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 1 | 0 | 0 | 0 | 0 | 0 |
| Read | 1 | 0 | 0 | 0 | 0 | 0 |
| Edit | 1 | 0 | 0 | 0 | 0 | 0 |
| Preserve | 0 | 1 | 0 | 0 | 0 | 0 |
| Inspect | 2 | 0 | 0 | 0 | 0 | 0 |
| Validate | 1 | 0 | 0 | 0 | 0 | 0 |
| Remove | 1 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Html` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
