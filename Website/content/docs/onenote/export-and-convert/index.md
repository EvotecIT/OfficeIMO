---
title: "Export and Convert OneNote"
description: "Render native OneNote pages to images and convert notebooks to HTML, Markdown, PDF, or Reader results."
order: 43
meta.seo_title: "Convert OneNote to PDF, image, HTML or Markdown"
---

OfficeIMO reads and writes desktop `.one`, FSSHTTP-encoded `.one`, `.onetoc2` notebook tables of contents, `.onepkg` exports, and notebook directories without requiring Microsoft OneNote.

## Export a page or notebook to images

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.OneNote;

OneNoteSection section = OneNoteSectionReader.Read("Planning.one");
OneNotePage page = section.Pages.Count > 0
    ? section.Pages[0]
    : throw new InvalidOperationException("Planning.one contains no pages.");

OfficeDrawing canvas = page.ToDrawing();

page.ToImage()
    .AtDpi(144)
    .AsPng()
    .Save("page.png");

OneNoteNotebook notebook =
    OneNoteNotebookReader.Read(Path.Combine("Notebook", "Open Notebook.onetoc2"));

notebook.ToImages()
    .AllPages()
    .AtDpi(144)
    .AsTiff()
    .Save("Notebook pages");
```

Load a `.one` section when you need one page, or its notebook's `.onetoc2` table of contents when you need every page. The same Drawing canvas owns SVG and pixel output, so PNG, JPEG, TIFF, and lossless WebP do not rely on separate OneNote renderers. Rendering includes positioned outlines, styled text, lists and tags, tables, images, printouts, ink, and structured math.

## Choose semantic or visual conversion

| Goal | Package and route |
|---|---|
| Editable Markdown or plain text | `OfficeIMO.OneNote.Markdown` |
| Accessible semantic HTML | `OfficeIMO.OneNote.Html` semantic route |
| Position-preserving HTML | `OfficeIMO.OneNote.Html` visual SVG-page route |
| Selectable-text PDF | `OfficeIMO.OneNote.Pdf` semantic route |
| Position-preserving PDF | `OfficeIMO.OneNote.Pdf` visual route |
| Search chunks and rich metadata | `OfficeIMO.Reader.OneNote` |

Current pages are the default conversion surface. Conflict copies and version-history snapshots are opt-in for direct conversions. Reader reports their counts in structured metadata.

## Bound image work

`OneNotePageRenderingOptions` controls page sizing, fonts, included features, source payload limits, and maximum raster pixels. Oversized exports can reduce scale with a diagnostic or throw `OfficeImageExportLimitException`, according to the selected overflow behavior.

For batch jobs, keep page selection and filenames deterministic and retain conversion diagnostics beside the output artifacts.
