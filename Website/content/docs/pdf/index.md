---
title: "PDF"
description: "Create, inspect, transform, render, sign, and validate PDF documents with OfficeIMO.Pdf."
order: 35
meta.seo_title: "PDF authoring and conversion | OfficeIMO"
---

## Install

```shell
dotnet add package OfficeIMO.Pdf
```

`OfficeIMO.Pdf` is the first-party PDF engine used by the OfficeIMO conversion packages. Use it directly for authored PDFs and PDF operations; add a format adapter such as `OfficeIMO.Word.Pdf` when the source is another document model.

## Choose a workflow

| Goal | Guide |
|---|---|
| Build a PDF report, form, or positioned layout | [Authoring and layout](/docs/pdf/authoring/) |
| Inspect, extract, merge, split, stamp, optimize, repair, or redact | [Inspection, extraction, and operations](/docs/pdf/operations/) |
| Encrypt, sign, verify, or preserve existing signed revisions | [Security and digital signatures](/docs/pdf/security/) |
| Render Word, Excel, PowerPoint, HTML, Markdown, or open formats | [Conversion and delivery](/docs/pdf/conversion/) |

The public surface also covers outlines and bookmarks, links, attachments, forms and widgets, annotations, fonts and fallback, text shaping diagnostics, metadata, page extraction, images, signatures, long-term validation evidence, and normalized Reader extraction. The [PDF API reference](/api/pdf/) is generated from the built assembly and is the complete type-level index.

## Create a PDF

```csharp
using OfficeIMO.Pdf;

PdfDocument.Create()
    .Meta(title: "Status report", author: "OfficeIMO")
    .H1("Status report")
    .Paragraph(p => p.Text("Created with ").Bold("OfficeIMO.Pdf"))
    .Table(new[] {
        new[] { "Area", "Status" },
        new[] { "Build", "Ready" }
    })
    .Save("status.pdf");
```

## Pick the correct package

| Starting point | Package |
|---|---|
| Author or edit PDF | `OfficeIMO.Pdf` |
| Word document | `OfficeIMO.Word.Pdf` |
| Excel workbook | `OfficeIMO.Excel.Pdf` |
| PowerPoint presentation | `OfficeIMO.PowerPoint.Pdf` |
| HTML | `OfficeIMO.Html.Pdf` |
| Markdown | `OfficeIMO.Markdown.Pdf` |
| OpenDocument | `OfficeIMO.OpenDocument.Pdf` |

Use the [PDF API reference](/api/pdf/) for type-level detail. For conversion, inspect returned diagnostics before accepting approximated or omitted features.

## Production checks

- Preflight incoming PDFs before extraction or mutation.
- Separate “can read” from “can safely rewrite.”
- Apply explicit limits to untrusted page trees, streams, attachments, and decompression.
- Validate font availability and international-text shaping on the deployment host.
- Verify redaction by re-extracting content; an opaque drawing alone is not redaction.
- Treat signature structure, cryptographic validity, certificate trust, later revisions, and certification permissions as separate decisions.
- Reopen generated output and assert the semantic features your workflow depends on.
