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
var html = doc.ToHtmlViaMarkdown();           // full HTML document
var fragment = doc.ToHtmlFragmentViaMarkdown();
doc.SaveAsHtmlViaMarkdown("report.html");
```

## Supported features (core)

- Headings 1–6, paragraphs, hard breaks
- Lists (unordered, ordered) and task items
- Tables with per‑column alignment
- Images (alt/title; size hints when provided in Markdown)
- Links and autolinks
- Code spans/blocks
- Footnotes
- Front matter passthrough when using OfficeIMO.Markdown model


