---
title: "PDF Authoring and Layout"
description: "Compose reports with typography, tables, images, drawings, headers, footers, forms, annotations, bookmarks, and page-level layout."
layout: docs
---

`OfficeIMO.Pdf` is a first-party managed authoring engine. Its fluent surface covers document metadata, headings and rich text, lists, tables, images, drawings, links, bookmarks, outlines, forms, annotations, headers, footers, page numbering, and text or image watermarks.

## Build semantically

Prefer document blocks—paragraphs, lists, tables, images, and sections—when content should flow. Use canvas and drawing primitives when coordinates are part of the contract. A report can combine both: flowing business content for most pages and positioned elements for labels, stamps, or custom diagrams.

```csharp
using OfficeIMO.Pdf;

PdfDocument.Create()
    .Meta(title: "Service report", author: "Operations")
    .Header(h => h.Text("Quarterly service review").AlignLeft())
    .H1("Service report")
    .Paragraph(p => p.Text("Generated from validated operational data."))
    .Table(new[] {
        new[] { "Service", "Availability", "Status" },
        new[] { "API", "99.97%", "On target" },
        new[] { "Portal", "99.92%", "Review" }
    })
    .Footer(f => f.Text("Confidential").AlignCenter())
    .Save("service-report.pdf");
```

## Fonts and international text

Font resolution is part of deployment. Configure embedded or system font sources explicitly, test fallback, and review text-encoding and shaping diagnostics for the scripts your reports use. Do not assume that a font installed on a developer workstation exists in a container or CI runner.

## Forms and annotations

Author text, choice, checkbox, and signature-oriented fields when the PDF itself is the interaction surface. Text, free-text, and highlight annotations support review workflows. Decide whether annotations remain interactive or are flattened before archival delivery.

## Validate output

Generate to bytes or a stream when a service needs transactional storage. Before publishing, run preflight, inspect diagnostics, reopen the output, and assert page count and critical text or fields.
