# OfficeIMO.Word.Html

Bidirectional HTML conversion for `OfficeIMO.Word`.

## What It Does Today

- Converts HTML into `WordDocument` with headings, paragraphs, inline formatting, links, images, SVG, lists, tables, captions, form controls, notes, headers, footers, and sections.
- Converts `WordDocument` back to HTML with document metadata, optional custom property metadata, optional headers and footers, optional comments, paragraph/run styles, optional list definition CSS, lists, tables, images, SVG, footnotes, endnotes, and optional default CSS.
- Preserves document-level language metadata between HTML `lang` attributes and `WordDocument.Settings.Language`.
- Supports inline CSS, embedded stylesheet content, external stylesheet paths, remote stylesheet loading through a caller-provided `HttpClient`, and common CSS page-break declarations.
- Supports image embedding, external image links, data URI images, resource timeouts, optional per-image and total image byte limits, declared image content-type validation, and URI scheme/host policy controls.
- Supports optional HTML node, HTML depth, CSS byte, and table-cell limits for safer conversion of large or hostile inputs.
- Reports HTML import diagnostics through `HtmlToWordOptions.Diagnostics` and `DiagnosticHandler` when content is skipped, degraded, or not mapped, such as image load failures and unsupported CSS declarations or values.
- Can emit opt-in accessibility diagnostics for missing image alternate text, weak or empty link text, skipped heading levels, and data tables without header cells.
- Lets callers choose unsupported CSS behavior with `HtmlToWordOptions.UnsupportedCssHandling`: ignore, warn by default, or stop conversion with `HtmlUnsupportedCssException`.
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

## Roadmap

The current plan for making this the market-leading HTML and Word conversion surface is tracked in `Docs/officeimo.word-html-roadmap.md`.
