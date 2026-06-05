# OfficeIMO.Word.Html

Bidirectional HTML conversion for `OfficeIMO.Word`.

## What It Does Today

- Converts HTML into `WordDocument` with headings, paragraphs, inline formatting, links, images, SVG, lists, tables, captions, form controls, notes, headers, footers, and sections.
- Converts `WordDocument` back to HTML with document metadata, optional custom property metadata, optional headers and footers, optional comments, paragraph/run styles, optional list definition CSS, lists, tables, images, SVG, footnotes, endnotes, disabled form controls, and optional default CSS.
- Preserves document-level language metadata between HTML `lang` attributes and `WordDocument.Settings.Language`.
- Supports inline CSS, embedded stylesheet content, external stylesheet paths, remote stylesheet loading through a caller-provided `HttpClient`, stylesheet URI scheme/host policy controls, stylesheet content-type validation, aggregate CSS byte limits, and common CSS page-break declarations.
- Reuses parsed stylesheet rules by CSS content hash, while changed local or remote stylesheet content is parsed as fresh input.
- Supports image embedding, external image links, data URI images, resource timeouts, optional per-image and total image byte limits, declared image content-type validation, and image URI scheme/host policy controls.
- Supports optional HTML node, HTML depth, per-stylesheet CSS byte, aggregate CSS byte, and table-cell limits for safer conversion of large or hostile inputs.
- Provides named import profiles through `HtmlToWordOptions.CreateOfficeIMOProfile()`, `CreateUntrustedHtmlProfile()`, and `CreateTrustedDocumentProfile()`, plus `Clone()` for reusable option templates.
- Reports HTML import diagnostics through `HtmlToWordOptions.Diagnostics` and `DiagnosticHandler` when content is skipped, degraded, or not mapped, such as image load failures, raw HTML comments, disabled or invalid stylesheet links, stylesheet policy/HTTP/transport/timeout failures, and unsupported CSS declarations or values.
- Can opt in to import non-empty raw HTML comments as native Word comments with configurable author and initials metadata.
- Skips unsupported embedded media/widget elements such as `iframe`, `object`, `embed`, `video`, `audio`, and `canvas` with diagnostics instead of importing fallback text as document content.
- Can emit opt-in accessibility diagnostics for missing image alternate text, weak or empty link text, skipped heading levels, and data tables without header cells.
- Lets callers choose unsupported CSS behavior with `HtmlToWordOptions.UnsupportedCssHandling`: ignore, warn by default, or stop conversion with `HtmlUnsupportedCssException`.
- Maps common document-table styling, including table/cell widths, borders, padding, horizontal and vertical alignment, captions, column groups, header/footer row groups, HTML `cellspacing`, and CSS `border-spacing`.
- Imports text and visible value-bearing inputs, dates, text areas, meter/progress values, single-select dropdowns, named radio groups, multi-select values, datalist-backed combo boxes, and checkbox markers into Word content controls where Word has a practical equivalent.
- Targets `netstandard2.0`, `net8.0`, `net10.0`, and `net472` on Windows.

The current supported feature set is tracked in `Docs/officeimo.word-html-support-matrix.md`.

## Entry Points

```csharp
using OfficeIMO.Word.Html;

WordDocument document = html.LoadFromHtml(new HtmlToWordOptions());
string roundTrip = document.ToHtml(new WordToHtmlOptions {
    IncludeDefaultCss = true,
    IncludeRunColorStyles = true,
    IncludeRunHighlightStyles = true,
    ExportEndnotes = true,
    ExportComments = true,
    ExportHeadersAndFooters = true,
    IncludeTableColumnGroups = true
});
```

HTML can also be appended to an existing body, header, or footer with the `AddHtmlToBody`, `AddHtmlToHeader`, and `AddHtmlToFooter` extension methods.

## Import Profiles

```csharp
var safeOptions = HtmlToWordOptions.CreateUntrustedHtmlProfile();
safeOptions.MaxHtmlNodes = 5000;
safeOptions.DiagnosticHandler = diagnostic => Console.WriteLine($"{diagnostic.Code}: {diagnostic.Source}");

WordDocument safeDocument = html.LoadFromHtml(safeOptions);

var trustedOptions = HtmlToWordOptions.CreateTrustedDocumentProfile();
trustedOptions.AllowedStylesheetHosts.Add("cdn.example.com");

WordDocument trustedDocument = trustedHtml.LoadFromHtml(trustedOptions);
```

- `CreateOfficeIMOProfile()` preserves the default compatibility-oriented behavior.
- `CreateUntrustedHtmlProfile()` keeps external document resources offline by default, enables accessibility diagnostics, and applies bounded HTML, CSS, image, and table limits.
- `CreateTrustedDocumentProfile()` enables document-provided stylesheet links for known-good HTML while retaining the current resource validation defaults.
- `Clone()` copies configuration values, allow-lists, configured stylesheets, and callbacks while starting with an empty diagnostics collection.

## Roadmap

The current plan for making this the market-leading HTML and Word conversion surface is tracked in `Docs/officeimo.word-html-roadmap.md`.
