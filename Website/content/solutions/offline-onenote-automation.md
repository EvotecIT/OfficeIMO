---
title: "Read and Convert OneNote Files Offline in .NET"
description: "Read native OneNote .one, .onetoc2, and .onepkg files offline in .NET, then export pages to images, HTML, Markdown, PDF, or search chunks."
meta.eyebrow: "Offline OneNote automation"
meta.outcome: "Own OneNote file processing without Graph or desktop OneNote"
meta.primary_label: "Read the OneNote guide"
meta.primary_url: "/docs/onenote/"
---

OfficeIMO.OneNote reads and writes desktop `.one` files, FSSHTTP-encoded sections, `.onetoc2` notebook tables of contents, `.onepkg` exports, and notebook directories without Microsoft OneNote or Microsoft Graph.

Use this route when the input is a native file or offline archive. Graph remains the right API when the source of truth is a live Microsoft 365 notebook and the application needs service authentication, permissions, synchronization, or collaboration state.

## Read and update a native section

```csharp
using OfficeIMO.OneNote;

OneNoteSection section = OneNoteSectionReader.Read("Projects.one");
OneNotePage page = section.Pages[0];

var paragraph = new OneNoteParagraph();
paragraph.Runs.Add(new OneNoteTextRun { Text = "Added offline" });
page.Paragraphs.Add(paragraph);

section.Save("Projects-updated.one");
```

Write to a new artifact during migration. Version history and conflict copies are distinct from the current page surface and should be selected deliberately.

## Choose semantic or visual output

| Required output | Package or route |
| --- | --- |
| Editable Markdown or plain text | `OfficeIMO.OneNote.Markdown` |
| Accessible semantic HTML | `OfficeIMO.OneNote.Html` semantic conversion |
| Position-preserving HTML | `OfficeIMO.OneNote.Html` visual SVG-page conversion |
| Selectable-text PDF | `OfficeIMO.OneNote.Pdf` semantic conversion |
| Position-preserving PDF | `OfficeIMO.OneNote.Pdf` visual conversion |
| PNG, JPEG, TIFF, WebP, or SVG pages | native OneNote image and Drawing APIs |
| Normalized metadata and search chunks | `OfficeIMO.Reader.OneNote` |

Semantic output prioritizes headings, text, lists, tables, links, and editing. Visual output prioritizes page placement. One route cannot maximize both for every notebook, so select the fidelity target instead of presenting conversion as lossless.

## Bound rendering and extraction

OneNote pages can contain positioned outlines, images, printouts, ink, math, and large embedded payloads. Configure page sizing, source payload limits, maximum raster pixels, feature inclusion, and overflow behavior before processing untrusted notebooks.

Continue with [OneNote export and conversion](/docs/onenote/export-and-convert/), the [OneNote product page](/products/onenote/), or the [OneNote API reference](/api/onenote/).
