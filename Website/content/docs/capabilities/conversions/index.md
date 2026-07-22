---
title: "Conversion and Rendering Map"
description: "Choose the owning model, focused adapter, diagnostic policy, and deployment path for OfficeIMO conversions."
layout: docs
---

OfficeIMO keeps document models and conversion adapters separate. That lets an application reference the source engine it needs and add only the destination routes it actually ships.

## Choose by source and destination

| Source | Destination | Focused adapter | Typical use |
|---|---|---|---|
| Word | HTML | `OfficeIMO.Word.Html` | Browser review, publishing, email previews, and content projection |
| Word | Markdown | `OfficeIMO.Word.Markdown` | Repository workflows, text pipelines, and review-friendly exports |
| Word | PDF | `OfficeIMO.Word.Pdf` | Fixed-layout delivery with explicit render diagnostics |
| Word | OpenDocument or RTF | `OfficeIMO.Word.OpenDocument`, `OfficeIMO.Word.Rtf` | Interchange with open-document and rich-text workflows |
| Excel | HTML or PDF | `OfficeIMO.Excel.Html`, `OfficeIMO.Excel.Pdf` | Workbook previews and fixed-layout reports |
| Excel | OpenDocument | `OfficeIMO.Excel.OpenDocument` | XLSX and ODS interoperability |
| PowerPoint | HTML or PDF | `OfficeIMO.PowerPoint.Html`, `OfficeIMO.PowerPoint.Pdf` | Slide review, publishing, and delivery |
| PowerPoint | OpenDocument | `OfficeIMO.PowerPoint.OpenDocument` | PPTX and ODP interoperability |
| Markdown | HTML or PDF | `OfficeIMO.Markdown.Html`, `OfficeIMO.Markdown.Pdf` | Typed text publishing and portable output |
| HTML | PDF | `OfficeIMO.Html.Pdf` | Managed web-to-PDF workflows without a browser process |
| OneNote | HTML, Markdown, or PDF | `OfficeIMO.OneNote.Html`, `OfficeIMO.OneNote.Markdown`, `OfficeIMO.OneNote.Pdf` | Notebook extraction and publishing |
| OpenDocument | PDF | `OfficeIMO.OpenDocument.Pdf` | ODT, ODS, and ODP delivery |
| AsciiDoc, LaTeX, or RTF | Markdown or PDF | format-specific `.Markdown` and `.Pdf` adapters | Bounded text-format bridges with explicit loss handling |

## Preserve the source model

Use the source package for loading, editing, and source-specific validation. The adapter should own projection into the destination. This matters because a table, field, comment, animation, formula, or drawing can have no exact equivalent in the target format.

Do not treat “a file was written” as proof that the conversion preserved everything important. For production routes:

1. load or build the source with its native model;
2. set an explicit resource policy for fonts, images, links, or external content;
3. run the focused converter and capture its structured warnings;
4. inspect or reopen the destination artifact;
5. keep representative fixtures in automated tests.

## Browser-local routes

The [browser converter](/playground/) exposes only routes that can execute safely inside the WebAssembly application. That is intentionally smaller than the managed server-side conversion surface. A missing browser route does not mean that no .NET adapter exists; consult the [complete component index](/docs/capabilities/packages/) and the adapter API.

## Loss policy

Converters should be selected by the content that must survive, not only by file extensions. Decide whether unsupported content should fail the operation, produce a warning, be approximated, or be omitted. Keep the original when legal or audit requirements make the conversion evidence important.

## Next steps

- Use the [Word conversion guide](/docs/word/conversion/) for DOCX-specific resource and review concerns.
- Use the [PDF conversion guide](/docs/pdf/conversion/) for fixed-layout delivery and diagnostic review.
- Use [Reader](/docs/reader/) when the goal is normalized extraction rather than a destination document.
