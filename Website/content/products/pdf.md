---
title: "OfficeIMO.Pdf"
description: "Create, inspect, edit, merge, split, stamp, sign, validate, and render PDF files with a first-party .NET engine."
layout: product
meta.seo_title: "OfficeIMO.Pdf for .NET applications"
meta.head_html: '<link rel="alternate" hreflang="en" href="https://officeimo.com/products/pdf/" /><link rel="alternate" hreflang="x-default" href="https://officeimo.com/products/pdf/" />'
product_label: "PDF engine"
product_color: "#ef4444"
install: "dotnet add package OfficeIMO.Pdf"
nuget: "OfficeIMO.Pdf"
docs_url: "/docs/pdf/"
api_url: "/api/pdf/"
---

## One PDF model from creation to validation

Use `OfficeIMO.Pdf` when a workflow must own the PDF rather than hand it to a desktop application. The package covers authoring, inspection, page operations, forms, attachments, annotations, rendering, signatures, and validation through the same first-party model.

```csharp
using OfficeIMO.Pdf;

PdfDocument.Create()
    .Meta(title: "Quarterly report", author: "OfficeIMO")
    .H1("Quarterly report")
    .Paragraph("Generated without Office or a browser runtime.")
    .Save("report.pdf");
```

## Choose the workflow you need

| Workflow | Use it for |
|---|---|
| Build | Reports, invoices, forms, labels, portfolios, and page-aware components |
| Inspect | Text, pages, links, images, attachments, outlines, forms, revisions, and active-content diagnostics |
| Transform | Merge, split, extract, reorder, rotate, stamp, watermark, overlay, and metadata changes |
| Secure | CMS-backed signatures, timestamps, certificate validation, and revision-aware inspection through `OfficeIMO.Security` |
| Render | Page images and format adapters used by Word, Excel, PowerPoint, HTML, RTF, and OpenDocument packages |

Complex source formats do not map perfectly to PDF in every case. Conversion results expose diagnostics so applications can decide whether an approximation is acceptable.
