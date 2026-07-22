---
title: "PDF Conversion and Delivery"
description: "Render Word, Excel, PowerPoint, HTML, Markdown, OpenDocument, and text formats to PDF with explicit diagnostics."
layout: docs
---

OfficeIMO conversion adapters project source models into the first-party PDF engine. Install the adapter that owns the source format; do not add every document package to a service that only needs one route.

| Source | Adapter |
|---|---|
| Word | `OfficeIMO.Word.Pdf` |
| Excel | `OfficeIMO.Excel.Pdf` |
| PowerPoint | `OfficeIMO.PowerPoint.Pdf` |
| HTML | `OfficeIMO.Html.Pdf` |
| Markdown | `OfficeIMO.Markdown.Pdf` |
| OpenDocument | `OfficeIMO.OpenDocument.Pdf` |
| OneNote | `OfficeIMO.OneNote.Pdf` |
| RTF, AsciiDoc, LaTeX | their focused `.Pdf` adapters |

## Fixed layout requires explicit choices

A source document may depend on fonts, page metrics, client-side field updates, formulas, animations, unsupported drawings, or external resources. Configure resource and fallback policies and retain warnings from the conversion result.

For Office sources, test headers and footers, tables, images, links, page breaks, fields, charts, comments or revision views, and content controls used by your templates. For HTML, test CSS, images, resource resolution, and pagination. For international content, exercise real fonts and shaping on every deployment platform.

## Delivery pipeline

1. Validate the source model and required assets.
2. Convert with an explicit adapter and options.
3. Review all structured warnings.
4. Preflight and reopen the produced PDF.
5. Assert page count, key text, links, fields, or signatures.
6. Sign only after all content and metadata mutations are complete.

Use [PDF operations](/docs/pdf/operations/) for post-processing and the [general conversion map](/docs/capabilities/conversions/) for non-PDF destinations.
