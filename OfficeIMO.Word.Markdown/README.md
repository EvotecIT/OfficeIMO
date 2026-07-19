# OfficeIMO.Word.Markdown - Word and Markdown conversion

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Word.Markdown)](https://www.nuget.org/packages/OfficeIMO.Word.Markdown)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Word.Markdown?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Word.Markdown)

`OfficeIMO.Word.Markdown` converts between `OfficeIMO.Word` documents and `OfficeIMO.Markdown` documents.

## Install

```powershell
dotnet add package OfficeIMO.Word.Markdown
```

## Quick start

```csharp
using OfficeIMO.Markdown;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

using var document = WordDocument.Create();
document.AddParagraph("Hello");

string markdown = document.ToMarkdown();
using var fromMarkdown = MarkdownReader.Parse("# Title\n\nBody").ToWordDocument();
```

## AST-first conversion

```csharp
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Word.Markdown;

MarkdownDoc markdownDocument = "<table><tr><td><p>Line 1</p><p>Line 2</p></td></tr></table>"
    .ToMarkdownDocument();

using var wordDocument = markdownDocument.ToWordDocument();
```

Use `MarkdownDoc.ToWordDocument()` when you already have a typed AST and want to avoid flattening back to Markdown text before Word conversion.

## What it maps

- Word to Markdown with headings, paragraphs, lists, task items, tables, images, links, code, footnotes, and GitHub-friendly output.
- Markdown to Word through the typed `OfficeIMO.Markdown` model.
- Markdown image layout options such as local-image allowance and page-content-width fitting.
- Selected AST-preserved inline HTML wrappers such as underline, superscript, and subscript into Word run formatting.

## HTML via Markdown

```csharp
using OfficeIMO.Markdown;
var html = doc.ToHtmlViaMarkdown();           // full HTML document (defaults to HtmlStyle.Word)
var fragment = doc.ToHtmlFragmentViaMarkdown();
doc.SaveAsHtmlViaMarkdown("report.html");     // defaults to HtmlStyle.Word

// Override style if desired:
doc.SaveAsHtmlViaMarkdown("report.html", new HtmlOptions { Style = HtmlStyle.GithubAuto });
```

Use `ToWordDocumentViaMarkdown()` when your source is HTML but you want the AST-first Markdown bridge instead of flattening HTML into Markdown text first.

## Notes

- Default styling: ToHtmlViaMarkdown/SaveAsHtmlViaMarkdown use HtmlStyle.Word for a document‑like look.
- TOC markers: `[TOC]`, `[[TOC]]`, `{:toc}`, and `<!-- TOC -->` are recognized. Markdown -> Word creates a native Word table of contents, and Word -> Markdown exports native TOCs back as `[TOC ...]` markers.
  Parameterized form: `[TOC min=2 max=3 layout=sidebar-right sticky=true scrollspy=true title="On this page"]`.
- Table cells: inline markdown (code/links/emphasis/images) is supported and `<br>` becomes a real line break in HTML.
- AST-preserved inline HTML wrappers such as `<u>`, `<sub>`, and `<sup>` map to real Word run formatting during Markdown -> Word conversion.
- Inline tags without a native Word run equivalent degrade intentionally on the Word leg: `<ins>` is treated as underline semantics and `<q>` roundtrips as literal quoted text.

## Markdown -> Word image layout options

```csharp
var options = new MarkdownToWordOptions {
    AllowLocalImages = true,
    FitImagesToPageContentWidth = true,           // page width minus margins
    MaxImageWidthPercentOfContent = 85            // optional percent-based cap
};

using var doc = MarkdownReader.Parse(markdown).ToWordDocument(options);
```

- Typed contract is available via `MarkdownToWordOptions.ImageLayout`.
- `PreferNarrativeSingleLineDefinitions = true` keeps isolated `Label: value` lines as narrative paragraphs while still allowing grouped definition-list blocks.
- `OnImageLayoutDiagnostic` can be used to inspect final width/height and applied layout constraints.

## Word -> Markdown layout and visual fallback options

```csharp
var options = new WordToMarkdownOptions {
    PageBreakMode = MarkdownPageBreakMode.SemanticBlock,
    UnsupportedContentMode = MarkdownUnsupportedContentMode.Placeholder,
    VisualFallbackMode = MarkdownVisualFallbackMode.SvgFile
};

doc.SaveAsMarkdown("report.md", options);
```

- `PageBreakMode.SemanticBlock` writes page breaks as fenced semantic blocks so Markdown -> Word can restore real Word page breaks.
- Headers, footers, and native Word table-of-contents fields are exported as semantic Markdown blocks/markers and restored on import where possible.
- `VisualFallbackMode.SvgDataUri` embeds supported Word chart snapshots directly in Markdown as SVG images.
- `VisualFallbackMode.SvgFile` writes supported chart snapshots into a sidecar `*.assets` directory next to the Markdown file and links to those files.
- Chart SVG fallbacks preserve cached chart data, chart type, dimensions, series colors, pie/doughnut point colors, common theme/scheme colors, and basic transparency transforms. Unsupported chart constructs remain semantic placeholders instead of being silently dropped when `UnsupportedContentMode.Placeholder` is enabled.

## Explicit IntelligenceX transcript contract

```csharp
var options = MarkdownToWordPresets.CreateIntelligenceXTranscript(
    allowedImageDirectories: new[] { @"C:\Exports\Images" },
    visualMaxWidthPx: 760);

using var doc = MarkdownReader.Parse(markdown).ToWordDocument(options);
```

- `MarkdownToWordPresets.CreateIntelligenceXTranscript(...)` is the explicit DOCX preset for IX transcript export.
- That preset now reuses the shared `MarkdownTranscriptPreparation.CreateIntelligenceXTranscriptReaderOptions(...)` contract, including legacy IX visual JSON upgrades through document transforms.
- `MarkdownToWordCapabilities.PreservesNarrativeSingleLineDefinitionsAsSeparateParagraphs()` exposes the grouped `Label: value` capability probe as a reusable OfficeIMO contract instead of host-local reflection logic.

## Supported features (core)

- Headings 1–6, paragraphs, hard breaks
- Lists (unordered, ordered) and task items
- Tables with per‑column alignment
- Images (alt/title; size hints when provided in Markdown)
- Links and autolinks
- Code spans/blocks
- AST-preserved inline HTML wrappers for underline, superscript, and subscript
- Footnotes
- Front matter passthrough when using OfficeIMO.Markdown model

## Boundaries

- Word document modeling belongs in `OfficeIMO.Word`.
- Markdown parsing and AST behavior belongs in `OfficeIMO.Markdown`.
- HTML ingestion belongs in `OfficeIMO.Markdown.Html` or `OfficeIMO.Word.Html`, depending on the desired source model.
- PDF output belongs in `OfficeIMO.Word.Pdf` and `OfficeIMO.Pdf`.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** None beyond the dependencies of its OfficeIMO format packages.
- **OfficeIMO:** `OfficeIMO.Word`, `OfficeIMO.Markdown`, `OfficeIMO.Markdown.Html`, `OfficeIMO.Word.Html`, and `OfficeIMO.Drawing`.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
