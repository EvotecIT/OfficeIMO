---
title: Convert EPUB to PDF, Markdown and Images
description: Convert EPUB e-books to PDF or Markdown, render chapters to PNG or SVG images, and save chapter XHTML and text to files for your own converter.
order: 43
search_terms:
  - epub to pdf
  - epub to markdown
  - epub to md
  - epub to image
  - epub to html
packages:
  - OfficeIMO.Epub
  - OfficeIMO.Epub.Image
  - OfficeIMO
cmdlets:
  - Convert-OfficeDocument
  - Read-OfficeDocument
commands:
  - officeimo convert
related:
  - /docs/epub/
  - /docs/epub/create/
  - /docs/html/render-and-convert/
limits: EPUB to PDF is a reading-order text conversion through the Reader, not a page-by-page rendering of the book's CSS; use the image export when appearance matters.
verified: 2026-11-14
---

Pick the route by what the output is for. For a printable or searchable copy, convert the EPUB to PDF or Markdown through the `OfficeIMO` converter or the CLI; both read the chapters in spine order. For a visual preview, `OfficeIMO.Epub.Image` renders chapters with their stylesheets and images through the HTML renderer. For your own pipeline, load the chapter XHTML and text directly from `EpubDocument`.

## Convert an EPUB to PDF

### C#

```csharp
using OfficeIMO;

OfficeConversionReport report = OfficeConverter.Convert("handbook.epub", "handbook.pdf",
    new OfficeConversionOptions { FileConflictPolicy = OfficePublicationPolicy.Replace });

foreach (OfficeDiagnostic diagnostic in report.Diagnostics) {
    Console.WriteLine($"{diagnostic.Severity} {diagnostic.Code}: {diagnostic.Message}");
}
```

### PowerShell

```powershell
Convert-OfficeDocument -Path .\handbook.epub -To Pdf -OutputPath .\handbook.pdf
```

### CLI

```sh
officeimo convert handbook.epub handbook.pdf
```

## Convert an EPUB to Markdown

The Reader writes the chapters as one Markdown document and can extract images into a folder.

### PowerShell

```powershell
Read-OfficeDocument -Path .\handbook.epub -As Markdown
```

### CLI

```sh
officeimo convert handbook.epub handbook.md --assets handbook-images
```

## Render chapters to images

Load the raw chapter HTML and resource bytes so the renderer can use the book's CSS and pictures. The default mode makes one continuous image per chapter; set `Mode = HtmlRenderMode.Paged` for page-sized images.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Epub;
using OfficeIMO.Epub.Image;
using OfficeIMO.Html;

EpubDocument book = EpubDocument.Load("handbook.epub", new EpubLoadOptions {
    IncludeRawHtml = true,
    IncludeResourceData = true
});

var options = new EpubImageExportOptions {
    Format = OfficeImageExportFormat.Png,
    ViewportWidth = 800
};

OfficeConversionReport report = book.SaveAsImages("handbook-pages", options);
Console.WriteLine($"{report.ItemsWritten} images written");
```

## Save chapter XHTML and text

`IncludeRawHtml` keeps each chapter's markup next to its extracted text, which is the starting point for a custom converter.

### C#

```csharp
using OfficeIMO.Epub;

EpubDocument book = EpubDocument.Load("handbook.epub", new EpubLoadOptions {
    IncludeRawHtml = true,
    MaxChapterBytes = 2L * 1024 * 1024
});

Directory.CreateDirectory("chapters");
foreach (EpubChapter chapter in book.Chapters) {
    File.WriteAllText(Path.Combine("chapters", $"chapter-{chapter.Order:000}.txt"), chapter.Text);
    if (chapter.Html != null) {
        File.WriteAllText(Path.Combine("chapters", $"chapter-{chapter.Order:000}.xhtml"), chapter.Html);
    }
}
```

## Convert a CHM help file to EPUB

### CLI

```sh
officeimo convert manual.chm manual.epub
```
