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
- `PreferNarrativeSingleLineDefinitions = true` disables definition-list parsing for `Label: value` narrative lines.
- `OnImageLayoutDiagnostic` can be used to inspect final width/height and applied layout constraints.

## Supported features (core)

- Headings 1–6, paragraphs, hard breaks
- Lists (unordered, ordered) and task items
- Tables with per‑column alignment
- Images (alt/title; size hints when provided in Markdown)
- Links and autolinks
- Code spans/blocks
- Footnotes
- Front matter passthrough when using OfficeIMO.Markdown model


