---
title: "Content Publishing Patterns"
description: "Turn Markdown, HTML, RTF, OneNote, and OpenDocument content into documents and delivery artifacts."
order: 7
meta.seo_title: "Convert Markdown, HTML, RTF and OpenDocument with OfficeIMO"
---

OfficeIMO keeps format engines and adapters separate. Install the source engine for parsing and editing, then add only the adapter for the destination you need.

## Common routes

| Source | Destination | Package route |
|---|---|---|
| Markdown | HTML | `OfficeIMO.Markdown` |
| Markdown | Word | `OfficeIMO.Word.Markdown` |
| Word | Markdown | `OfficeIMO.Word.Markdown` |
| HTML or MHTML | PDF | `OfficeIMO.Html` + `OfficeIMO.Html.Pdf` |
| HTML | Word or Markdown | `OfficeIMO.Word.Html` or `OfficeIMO.Markdown.Html` |
| RTF | Word, Markdown, HTML, PDF | `OfficeIMO.Word.Rtf`, `OfficeIMO.Rtf.Markdown`, `OfficeIMO.Html`, or `OfficeIMO.Rtf.Pdf` |
| OneNote | HTML, Markdown, PDF, images | `OfficeIMO.OneNote.Html`, `.Markdown`, `.Pdf`, or the native image API |
| ODT, ODS, ODP | DOCX, XLSX, PPTX, PDF | the focused OpenDocument adapter for that destination |

## Treat conversion as a result

A useful conversion pipeline records more than the output path:

- Source identity and selected converter route.
- Diagnostics for flattened, omitted, blocked, or unsupported content.
- Output validation status.
- A preview or visual baseline when layout is part of the contract.
- The package versions used to produce the artifact.

RTF and OpenDocument adapters expose conversion reports so a strict workflow can reject loss rather than silently accept it. HTML resource policies control network, file, embedded, and archive content. Reader provides normalized extraction when preserving presentation is not the goal.

## Choose semantic or visual output

Semantic output preserves headings, paragraphs, tables, links, and other meaning for editing and accessibility. Visual output prioritizes page placement and appearance. OneNote HTML/PDF and several OfficeIMO converters expose both kinds of route; choose explicitly instead of assuming one output can maximize both.

Continue with [HTML rendering](/docs/html/render-and-convert/), [RTF](/docs/rtf/), [OneNote export](/docs/onenote/export-and-convert/), or the [conversion capability map](/docs/capabilities/conversions/).
