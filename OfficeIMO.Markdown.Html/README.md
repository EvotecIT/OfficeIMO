# OfficeIMO.Markdown.Html - HTML to Markdown conversion

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Markdown.Html)](https://www.nuget.org/packages/OfficeIMO.Markdown.Html)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Markdown.Html?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Markdown.Html)

`OfficeIMO.Markdown.Html` converts HTML fragments or full HTML documents into Markdown text and `OfficeIMO.Markdown` document models.

## Install

```powershell
dotnet add package OfficeIMO.Markdown.Html
```

## Quick start

```csharp
using OfficeIMO.Markdown;
using OfficeIMO.Html;
using OfficeIMO.Markdown.Html;

HtmlConversionDocument source = HtmlConversionDocument.Parse("<h1>Hello</h1><p>Body</p>");
string markdown = source.ToMarkdown();
MarkdownDoc document = source.ToMarkdownDocument();
```

For multi-target workflows, prepare HTML once and reuse its normalized DOM:

```csharp
using OfficeIMO.Html;

HtmlConversionDocument source = HtmlConversionDocumentBuilder.Build(html);
string markdown = source.ToMarkdown();
MarkdownDoc document = source.ToMarkdownDocument();
```

Create a prepared document with `HtmlConversionDocument.Parse(...)`, `Load(...)`, or `LoadAsync(...)`; all three enter through the same bounded parser. Save Markdown through the prepared document's path, stream, and async `SaveAsMarkdown` methods. Adapter-local filtering works on a clone, so converting to Markdown does not mutate the document consumed by Word, RTF, PDF, or image output.

## Options

```csharp
using OfficeIMO.Markdown.Html;

var options = new HtmlToMarkdownOptions {
    BaseUri = new Uri("https://example.com/docs/"),
    UseBodyContentsOnly = true,
    PreserveUnsupportedBlocks = true,
    PreserveUnsupportedInlineHtml = true,
    SmartHref = true
};

HtmlConversionDocument source = HtmlConversionDocument.Parse("<p><a href=\"guide/start\">Docs</a></p>");
string markdown = source.ToMarkdown(options);
```

### Compatibility controls

```csharp
var options = new HtmlToMarkdownOptions {
    SmartHref = true,
    UnknownBlockHandling = HtmlUnknownTagHandling.Bypass,
    UnknownInlineHandling = HtmlUnknownTagHandling.Bypass
};

options.ExcludeSelectors.Add(".ad, .cookie-banner");
options.TagAliases["highlight"] = "mark";
options.PassThroughTags.Add("custom-widget");

HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
string markdown = source.ToMarkdown(options);
```

- `SmartHref` emits self-describing links such as `<a href="https://example.com">https://example.com</a>` as plain text.
- `ExcludeSelectors` and `ElementFilters` remove matching HTML before conversion.
- `TagAliases` lets custom or legacy tag names reuse built-in converters.
- `PassThroughTags` preserves selected elements as raw HTML.
- `UnknownBlockHandling` and `UnknownInlineHandling` choose whether unknown elements are preserved, bypassed, dropped, or rejected.

## What it maps

- Headings, paragraphs, ordered and unordered lists, block quotes, code blocks, horizontal rules, tables, images, figures, details/summary, definition lists, and raw HTML fallback blocks.
- Emphasis, strong emphasis, strike-through, code spans, links, images, line breaks, and selected inline HTML wrappers.
- Native and ARIA headings, accessible link/image names, ordered-list marker families, reversed lists, item value resets, and code-language hints. Generated Markdown uses decimal CommonMark markers while the typed list retains the original HTML marker family for HTML projection.
- Local EPUB/DPUB-ARIA footnotes and GitHub-style footnote sections as typed `FootnoteRefInline` and `FootnoteDefinitionBlock` nodes, including structured definition bodies.
- Relative links and image targets when a base URI is supplied.
- Shared `data-omd-*` visual host elements back into semantic fenced blocks when host hints are registered.
- Custom block and inline converters for host or plug-in packages.

EPUB footnote references are converted only when their target is an actual footnote, endnote, or rearnote definition in the same HTML document. Generic `epub:type="note"` content stays ordinary content. Cross-document note links remain ordinary links so the converter does not invent a local definition. Footnote backlinks are omitted from Markdown. Accessible names use a deterministic conversion subset: `aria-labelledby`, `aria-label`, host-language alternatives such as image `alt`, optional visible text, and `title` fallback.

## Profiles

- Use the OfficeIMO profile when the downstream consumer is `OfficeIMO.Markdown` or `OfficeIMO.MarkdownRenderer`.
- Use the GitHub Flavored Markdown profile for README and GitHub documentation output.
- Use the CommonMark profile when output should avoid GitHub-only constructs such as pipe tables; HTML tables are emitted as raw HTML.
- Use the portable profile when output should remain friendly to generic Markdown engines.

```csharp
var github = HtmlToMarkdownOptions.CreateGitHubFlavoredMarkdownProfile();
string readme = HtmlConversionDocument
    .Parse("<p><a href=\"https://example.com\">https://example.com</a></p>")
    .ToMarkdown(github);

var commonMark = HtmlToMarkdownOptions.CreateCommonMarkProfile();
string portableTable = HtmlConversionDocument
    .Parse("<table><tr><th>Name</th></tr><tr><td>Value</td></tr></table>")
    .ToMarkdown(commonMark);
```

## Boundaries

- This package owns HTML ingestion into Markdown.
- It does not render Markdown to a host shell; that belongs in `OfficeIMO.MarkdownRenderer`.
- It does not export to PDF; that belongs in `OfficeIMO.Markdown.Pdf` and `OfficeIMO.Pdf`.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** None beyond `OfficeIMO.Html`, which isolates AngleSharp DOM/CSS parsing.
- **OfficeIMO:** `OfficeIMO.Html` and `OfficeIMO.Markdown` own the source models, mapping, plug-in hooks, and diagnostics.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
