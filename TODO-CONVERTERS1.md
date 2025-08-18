# TODO — Converters (Word ⇄ HTML/Markdown, Word ⇒ PDF)

> Keep current project split: `OfficeIMO.Word.Html`, `OfficeIMO.Word.Markdown`, `OfficeIMO.Word.Pdf`.
> No new facade; each converter exposes its own public API.
> Provide both **async** and **sync** entry points where appropriate.

---

## Goals

- Solid **export** to HTML/Markdown/PDF.
- Practical **import** from HTML and Markdown → native Word structures.
- Round-trip tests (HTML→DOCX→HTML, MD→DOCX→MD) with tolerances.
- Clear options objects; predictable defaults.

---

## Public API (proposed shape)

### HTML

**Export**

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

using var doc = WordDocument.Load("in.docx");
var html = HtmlExporter.ToHtml(doc, new ExportHtmlOptions {
  InlineCss = false,          // emit class names
  DataUriImages = true,       // or extract to files
  PreserveTableWidths = true
});
// write html to file/stream as needed
````

**Import**

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

string html = File.ReadAllText("input.html");
using var doc = HtmlImporter.FromHtml(html, new ImportHtmlOptions {
  BaseUri = null,
  NormalizeWhitespace = true,
  ConvertInlineStylesToNamedStyles = true
});
doc.Save("out.docx");
```

**Options**

* `ExportHtmlOptions`: `InlineCss`, `DataUriImages`, `ImageFolder`, `PreserveTableWidths`.
* `ImportHtmlOptions`: `BaseUri`, `NormalizeWhitespace`, `ConvertInlineStylesToNamedStyles`.

**Coverage targets**

* Blocks: paragraphs, headings, lists (ol/ul), blockquote, code/pre, tables.
* Inline: bold/italic/underline/strikethrough, code, links, images.
* Images: data URIs and external references.
* Unsupported nodes collected into a diagnostics log.

---

### Markdown

**Export**

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

using var doc = WordDocument.Load("in.docx");
var md = MarkdownExporter.ToMarkdown(doc, new ExportMarkdownOptions {
  GfmTables = true,
  EscapeHtml = true
});
```

**Import**

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

string md = File.ReadAllText("README.md");
using var doc = MarkdownImporter.FromMarkdown(md, new ImportMarkdownOptions {
  GitHubFlavored = true,
  HardLineBreaks = false
});
doc.Save("out.docx");
```

**Options**

* `ExportMarkdownOptions`: `GfmTables`, `EscapeHtml`.
* `ImportMarkdownOptions`: `GitHubFlavored`, `HardLineBreaks`.

**Coverage targets**

* Headings, paragraphs, emphasis/strong, inline code, fenced code blocks, lists (ordered/unordered), images, links.
* Tables (pipe syntax); task lists (optional when GFM).

---

### PDF (export)

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var doc = WordDocument.Load("in.docx");
var bytes = PdfExporter.ToPdf(doc, new ExportPdfOptions {
  Dpi = 300,
  DownsampleImages = true,
  EmbedFonts = true
});
File.WriteAllBytes("out.pdf", bytes);
```

**Options**

* `ExportPdfOptions`: `Dpi`, `DownsampleImages`, `EmbedFonts`.

---

## Implementation Tasks

### A. HTML

* [ ] **Export**: map Word blocks/inline to HTML + CSS classes (avoid inline styles by default).
* [ ] **Images**: support data URIs and external image folder; deduplicate.
* [ ] **Lists**: preserve nesting/numbering; map bullet types reasonably.
* [ ] **Tables**: emit `<table>` with `<thead>/<tbody>`; column widths if available.
* [ ] **Import**: HTML DOM → Word structures (paragraphs, runs, links, images, lists, tables, pre/code, blockquote).
* [ ] **Diagnostics**: collect unsupported tags/attributes.
* [ ] **Sync wrappers** for all public async methods (and vice versa if you start sync-first).

### B. Markdown

* [ ] **Export**: headings (#), lists, emphasis, code, links/images, fenced code blocks; tables when enabled.
* [ ] **Import**: parse MD → Word paragraphs/runs/lists/tables; respect hard breaks option; GFM extensions (tables, task lists).
* [ ] **Sync wrappers** mirroring HTML converter style.

### C. PDF

* [ ] Expose both stream and byte\[] outputs.
* [ ] Downsampling controls and font embedding toggles.
* [ ] `WordDocument.SaveAsPdf(path)` convenience overloads (sync + async).

### D. Tests & Examples

* [ ] Round-trip tests: HTML→DOCX→HTML and MD→DOCX→MD (assert structure, allow styling differences).
* [ ] Golden files for typical docs (headings, lists, tables, code, images).
* [ ] Examples: minimal, table-heavy, image-heavy docs.

---

## Acceptance Criteria

* HTML export/import supports common blocks + inline; images work for both data URIs and external files.
* Markdown export/import supports headings/lists/images/links/code and (optionally) tables.
* PDF export configurable for DPI/downsampling; fonts embedded when requested.
* All converters offer both async and sync entry points.
