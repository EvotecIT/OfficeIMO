---
title: "PDF"
description: "Create, inspect, transform, render, sign, and validate PDF documents with OfficeIMO.Pdf."
order: 35
meta.seo_title: "PDF authoring and conversion | OfficeIMO"
meta.head_html: '<link rel="alternate" hreflang="en" href="https://officeimo.com/docs/pdf/" /><link rel="alternate" hreflang="x-default" href="https://officeimo.com/docs/pdf/" />'
---

## Install

```shell
dotnet add package OfficeIMO.Pdf
```

`OfficeIMO.Pdf` is the first-party PDF engine used by the OfficeIMO conversion packages. Use it directly for authored PDFs and PDF operations; add a format adapter such as `OfficeIMO.Word.Pdf` when the source is another document model.

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
