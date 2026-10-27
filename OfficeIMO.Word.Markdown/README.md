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

using MarkdownDocument model = document.ToMarkdown();
string markdown = model.ToMarkdown();
using var fromMarkdown = MarkdownDocument.Parse("# Title\n\nBody").ToWord();
```

## AST-first conversion

```csharp
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Word.Markdown;

using MarkdownDocument markdownDocument = OfficeIMO.Html.HtmlConversionDocument
    .Parse("<table><tr><td><p>Line 1</p><p>Line 2</p></td></tr></table>")
    .ToMarkdown();

using var wordDocument = markdownDocument.ToWord();
```

Use `MarkdownDocument.ToWord()` when you already have a typed AST and want to avoid flattening back to Markdown text before Word conversion.

## What it maps

- Word to Markdown with headings, paragraphs, lists, task items, tables, images, links, code, footnotes, and GitHub-friendly output.
- Markdown to Word through the typed `OfficeIMO.Markdown` model.
- Markdown image layout options such as local-image allowance and page-content-width fitting.
- Selected AST-preserved inline HTML wrappers such as underline, superscript, and subscript into Word run formatting.

## HTML via Markdown

Convert Word into a Markdown model, then render it with Markdown's HTML options:

```csharp
using OfficeIMO.Markdown;
using OfficeIMO.Word.Markdown;

using var source = OfficeIMO.Word.WordDocument.Load("report.docx");
using var markdown = source.ToMarkdown();
string html = markdown.ToHtml(new MarkdownHtmlExportOptions {
    Kind = MarkdownHtmlKind.Document,
    Style = MarkdownHtmlStyle.Word
});
markdown.SaveAsHtml("report.html", new MarkdownHtmlExportOptions {
    Kind = MarkdownHtmlKind.Document,
    Style = MarkdownHtmlStyle.GithubAuto
});
```

For HTML input, add `OfficeIMO.Markdown.Html`, parse with `HtmlConversionDocument.Parse(html)`, then call `ToMarkdown().ToWord()`. `OfficeIMO.Word.Markdown` has no dependency on the HTML engine. The `OfficeIMO` front door wires optional raw-HTML handling; a focused adapter reports unsupported raw blocks through its conversion diagnostics.

## Notes

- Select `MarkdownHtmlStyle.Word` explicitly for document-like HTML styling.
- TOC markers: `[TOC]`, `[[TOC]]`, `{:toc}`, and `<!-- TOC -->` are recognized. Markdown -> Word creates a native Word table of contents, and Word -> Markdown exports native TOCs back as `[TOC ...]` markers.
  Parameterized form: `[TOC min=2 max=3 layout=sidebar-right sticky=true scrollspy=true title="On this page"]`.
- Table cells: inline markdown (code/links/emphasis/images) is supported and `<br>` becomes a real line break in HTML.
- Native Markdown tables cannot express merged Word cells or cell borders. Word-to-Markdown keeps their text in a flat table and reports the formatting loss through `ToMarkdown(options, out report)`.
- AST-preserved inline HTML wrappers such as `<u>`, `<sub>`, and `<sup>` map to real Word run formatting during Markdown -> Word conversion.
- Inline tags without a native Word run equivalent degrade intentionally on the Word leg: `<ins>` is treated as underline semantics and `<q>` roundtrips as literal quoted text.

## Markdown -> Word image layout options

```csharp
using OfficeIMO.Markdown;
var options = new WordMarkdownImportOptions {
    AllowLocalImages = true,
    FitImagesToPageContentWidth = true,           // page width minus margins
    MaxImageWidthPercentOfContent = 85            // optional percent-based cap
};

using var doc = MarkdownDocument.Parse("# Report\n\nBody").ToWord(options);
```

- Typed contract is available via `WordMarkdownImportOptions.ImageLayout`.
- `PreferNarrativeSingleLineDefinitions = true` keeps isolated `Label: value` lines as narrative paragraphs while still allowing grouped definition-list blocks.
- `OnImageLayoutDiagnostic` can be used to inspect final width/height and applied layout constraints.

## Word -> Markdown layout and visual fallback options

```csharp
using OfficeIMO.Markdown;
var options = new WordMarkdownExportOptions {
    PageBreakMode = WordMarkdownPageBreakMode.SemanticBlock,
    UnsupportedContentMode = WordMarkdownUnsupportedContentMode.Placeholder,
    VisualFallbackMode = WordMarkdownVisualFallbackMode.SvgFile
};

doc.SaveAsMarkdown("report.md", options);
```

- `PageBreakMode.SemanticBlock` writes page breaks as fenced semantic blocks so Markdown -> Word can restore real Word page breaks.
- Set `IncludeHeadersAndFootersAsSemanticBlocks = true` to export running content as semantic Markdown blocks. Omitted running content and explicit page breaks remain visible as omission diagnostics, including when the selected policy suppresses warning text. Native Word table-of-contents fields use semantic markers where supported.
- `VisualFallbackMode.SvgDataUri` embeds supported Word chart snapshots directly in Markdown as SVG images.
- `VisualFallbackMode.SvgFile` writes supported chart snapshots into a sidecar `*.assets` directory next to the Markdown file and links to those files. File images and SVG sidecars are staged until loss acceptance; `FailOnLoss` leaves existing Markdown and assets unchanged when the report is rejected.
- Chart SVG fallbacks use the shared Word chart projection. They preserve supported cached values, numeric X coordinates, bubble sizes, category combinations and secondary-axis assignments, dimensions, native plotting order, series and point styles, radial geometry, basic label content, surfaces, axes, and resolved fonts. Secondary scales are automatic; unsupported native settings remain semantic placeholders when `UnsupportedContentMode.Placeholder` is enabled.

## Explicit IntelligenceX transcript contract

```csharp
using OfficeIMO;
using OfficeIMO.Markdown;
using OfficeIMO.Word.Markdown;

var options = WordMarkdownImportOptions.CreateIntelligenceXTranscript(
    allowedImageDirectories: new[] { @"C:\Exports\Images" },
    visualMaxWidth: OfficeLength.Pixels(760));

using var doc = MarkdownDocument.Parse("# Report\n\nBody").ToWord(options);
```

- `WordMarkdownImportOptions.CreateIntelligenceXTranscript(...)` is the explicit DOCX preset for IX transcript export.
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
- **OfficeIMO:** `OfficeIMO.Word`, `OfficeIMO.Markdown`, `OfficeIMO.Markdown.Html`, `OfficeIMO.Word.Html`, and `OfficeIMO.Core`.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 2 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Word.Markdown` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
