# OfficeIMO.Word.Markdown — Markdown ↔ Word Converters

Converters between Word documents and Markdown, built on top of OfficeIMO.Word and OfficeIMO.Markdown.

## What it does

- Convert Word → Markdown with GitHub‑friendly output (headings, lists incl. tasks, tables with alignment, images, links, code, footnotes).
- Convert Markdown → Word using OfficeIMO.Markdown’s typed model (maps to real Word constructs).

## Usage

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

using var doc = WordDocument.Create();
doc.AddParagraph("Hello");
var md = doc.ToMarkdown();            // Word → Markdown

using var doc2 = "# Title\nBody".LoadFromMarkdown(); // Markdown → Word
```

### AST-first Markdown -> Word

```csharp
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Word.Markdown;

var markdownDoc = "<table><tr><td><p>Line 1</p><p>Line 2</p></td></tr></table>".LoadFromHtml();
using var doc3 = markdownDoc.ToWordDocument();
```

- Use `MarkdownDoc.ToWordDocument()` when you already have a typed AST and want to avoid flattening back to markdown text before Word conversion.

### HTML -> Word via Markdown AST

```csharp
using OfficeIMO.Word.Markdown;

using var doc4 = "<table><tr><td><p>Line 1</p><p>Line 2</p></td></tr></table>"
    .LoadFromHtmlViaMarkdown();
```

- Use `LoadFromHtmlViaMarkdown()` when your source is HTML but you want the AST-first markdown bridge instead of flattening HTML into markdown text first.

### HTML via Markdown

```csharp
using OfficeIMO.Markdown;
var html = doc.ToHtmlViaMarkdown();           // full HTML document (defaults to HtmlStyle.Word)
var fragment = doc.ToHtmlFragmentViaMarkdown();
doc.SaveAsHtmlViaMarkdown("report.html");     // defaults to HtmlStyle.Word

// Override style if desired:
doc.SaveAsHtmlViaMarkdown("report.html", new HtmlOptions { Style = HtmlStyle.GithubAuto });
```

### Notes
- Default styling: ToHtmlViaMarkdown/SaveAsHtmlViaMarkdown use HtmlStyle.Word for a document‑like look.
- TOC markers: `[TOC]`, `[[TOC]]`, `{:toc}`, and `<!-- TOC -->` are recognized and rendered.
  Parameterized form: `[TOC min=2 max=3 layout=sidebar-right sticky=true scrollspy=true title="On this page"]`.
- Table cells: inline markdown (code/links/emphasis/images) is supported and `<br>` becomes a real line break in HTML.

### Markdown -> Word image layout options

```csharp
var options = new MarkdownToWordOptions {
    AllowLocalImages = true,
    FitImagesToPageContentWidth = true,           // page width minus margins
    MaxImageWidthPercentOfContent = 85            // optional percent-based cap
};

using var doc = markdown.LoadFromMarkdown(options);
```

- Typed contract is available via `MarkdownToWordOptions.ImageLayout`.
- `PreferNarrativeSingleLineDefinitions = true` keeps isolated `Label: value` lines as narrative paragraphs while still allowing grouped definition-list blocks.
- `OnImageLayoutDiagnostic` can be used to inspect final width/height and applied layout constraints.

### Explicit IntelligenceX transcript contract

```csharp
var options = MarkdownToWordPresets.CreateIntelligenceXTranscript(
    allowedImageDirectories: new[] { @"C:\Exports\Images" },
    visualMaxWidthPx: 760);

using var doc = markdown.LoadFromMarkdown(options);
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
- Footnotes
- Front matter passthrough when using OfficeIMO.Markdown model


