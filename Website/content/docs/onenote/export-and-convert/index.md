---
title: Convert OneNote to PDF, HTML, Markdown and Images
description: Convert OneNote sections and notebooks to PDF, HTML or Markdown, render pages to PNG, TIFF or SVG images, and extract chunks for search, all offline.
order: 43
meta.seo_title: "Convert OneNote to PDF, image, HTML or Markdown"
search_terms:
  - onenote to pdf
  - onenote to html
  - onenote to markdown
  - one to pdf
  - onenote page to image
  - export onenote
packages:
  - OfficeIMO.OneNote
  - OfficeIMO.OneNote.Markdown
  - OfficeIMO.OneNote.Html
  - OfficeIMO.OneNote.Pdf
  - OfficeIMO.Reader.OneNote
cmdlets:
  - Read-OfficeDocument
commands:
  - officeimo reader read
  - officeimo reader folder
related:
  - /docs/onenote/
  - /docs/onenote/notebooks/
limits: Semantic PDF and HTML flatten the free-form page canvas into reading order; use the visual routes when placement matters. Conflict copies and version history are left out unless you ask for them.
verified: 2026-11-14
---

OneNote pages are free-form canvases, so every destination has two routes. Semantic routes put the content in reading order as ordinary text, which is best for search, editing and accessibility. Visual routes keep each element where it sits on the page, using the same drawing canvas as image export. Each conversion returns an `OfficeConversionReport` that names what was flattened or replaced with a placeholder.

| Goal | Package | Call |
|---|---|---|
| Markdown or plain text | `OfficeIMO.OneNote.Markdown` | `ToMarkdown`, `SaveAsMarkdown` |
| Reflowable HTML | `OfficeIMO.OneNote.Html` | `ToHtml`, `SaveAsHtml` |
| Position-preserving HTML | `OfficeIMO.OneNote.Html` | `ToVisualHtml`, `SaveAsVisualHtml` |
| Selectable-text PDF | `OfficeIMO.OneNote.Pdf` | `ToPdf`, `SaveAsPdf` |
| Position-preserving PDF | `OfficeIMO.OneNote.Pdf` | `ToVisualPdf`, `SaveAsVisualPdf` |
| PNG, JPEG, TIFF, WebP, SVG | `OfficeIMO.OneNote` | `SaveAsImage`, `SaveAsImages` |
| Search chunks and metadata | `OfficeIMO.Reader.OneNote` | `AddOneNoteHandler` |

## Convert a section to Markdown

Current pages are projected by default. Conflict copies and version-history pages are opt-in because they can duplicate or expose superseded content.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Markdown;

using OneNoteSection section = OneNoteSection.Load("Planning.one");

OfficeConversionReport report = section.SaveAsMarkdown("Planning.md", new OneNoteMarkdownExportOptions {
    IncludeConflictPages = false,
    IncludeVersionHistory = false
});

Console.WriteLine(File.ReadAllText("Planning.md"));
```

### PowerShell

```powershell
Read-OfficeDocument -Path .\Planning.one -As Markdown
```

### CLI

```sh
officeimo reader read Planning.one --format markdown --output Planning.md
```

## Convert a section to PDF

### C#

```csharp
using OfficeIMO;
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Pdf;

using OneNoteSection section = OneNoteSection.Load("Planning.one");

OfficeConversionReport semantic = section.SaveAsPdf("Planning.pdf");

OfficeConversionReport visual = section.SaveAsVisualPdf("Planning-visual.pdf",
    new OneNoteVisualPdfExportOptions { Title = "Planning" });

foreach (OfficeDiagnostic diagnostic in visual.Diagnostics) {
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
}
```

## Convert a section to HTML

### C#

```csharp
using OfficeIMO;
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Html;

using OneNoteSection section = OneNoteSection.Load("Planning.one");

string html = section.ToHtml(null, out OfficeConversionReport report);
File.WriteAllText("Planning.html", html);

section.SaveAsVisualHtml("Planning-visual.html", new OneNoteVisualHtmlExportOptions {
    DocumentTitle = "Planning",
    Language = "en",
    IncludeAccessibleText = true
});
```

## Render pages to images

`SaveAsImage` renders one page; `SaveAsImages` renders a section or notebook in table-of-contents order. Rendering covers positioned outlines, styled text, lists and tags, tables, images, printouts, ink and math.

### C#

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.OneNote;

using OneNoteSection section = OneNoteSection.Load("Planning.one");

section.Pages[0].SaveAsImage("page-1.png", new OneNotePageRenderingOptions {
    TargetDpi = 144,
    Format = OfficeImageExportFormat.Png
});

OneNoteNotebook notebook = OneNoteNotebookReader.Read(Path.Combine("Notebook", "Open Notebook.onetoc2"));
notebook.SaveAsImages("Notebook pages", new OneNotePageBatchRenderingOptions {
    TargetDpi = 144,
    MaximumOutputWidth = 1600,
    Format = OfficeImageExportFormat.Tiff
});
```

`MaximumRasterPixels` and `RasterOverflowBehavior` decide whether an oversized page is scaled down (with a diagnostic) or rejected with `OfficeImageExportLimitException`.

## Extract chunks for search

### C#

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.OneNote;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddOneNoteHandler(new ReaderOneNoteOptions {
        IncludeConflictPages = false,
        IncludeVersionHistory = false
    })
    .Build();

foreach (ReaderChunk chunk in reader.Enumerate("Planning.one")) {
    Console.WriteLine(chunk.Markdown ?? chunk.Text);
}
```

### CLI

```sh
officeimo reader folder ./Notebooks --output ./Notebooks-md --format markdown
```
