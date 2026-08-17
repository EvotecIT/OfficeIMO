---
title: "OfficeIMO.Pdf"
description: "Create, inspect, edit, merge, split, stamp, sign, validate, and render PDF files with a first-party .NET engine. Compare packages, examples, and limits."
layout: product
meta.seo_title: "OfficeIMO.Pdf for .NET applications"
meta.head_html: '<link rel="alternate" hreflang="en" href="https://officeimo.com/products/pdf/" /><link rel="alternate" hreflang="x-default" href="https://officeimo.com/products/pdf/" />'
product_label: "PDF engine"
product_color: "#ef4444"
install: "dotnet add package OfficeIMO.Pdf"
nuget: "OfficeIMO.Pdf"
docs_url: "/docs/pdf/"
api_url: "/api/pdf/"
meta.software.name: "OfficeIMO.Pdf"
meta.software.application_category: "DeveloperApplication"
meta.software.operating_system: "Windows, Linux, macOS"
meta.software.download_url: "https://www.nuget.org/packages/OfficeIMO.Pdf"
meta.software.price: 0
meta.software.price_currency: "USD"
---

## One PDF model from creation to validation

Use `OfficeIMO.Pdf` when a workflow must own the PDF rather than hand it to a desktop application. The package covers authoring, inspection, page operations, forms, attachments, annotations, rendering, signatures, and validation through the same first-party model.

```csharp
using OfficeIMO.Pdf;

PdfDocument.Create(pdf => pdf.Content(content => content
        .H1("Quarterly report")
        .Paragraph(paragraph => paragraph.Text("Generated without Office or a browser runtime."))))
    .Meta(title: "Quarterly report", author: "OfficeIMO")
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

## Try real PDF workflows in the browser

The [browser document workspace](/apps/officeimo-converter/?workspace=pdf&tool=inspect) runs a focused set of `OfficeIMO.Pdf` operations through WebAssembly. Files stay in the current tab, and every successful operation produces a downloadable artifact plus a JSON report with input and output fingerprints.

| Browser task | Result |
|---|---|
| [Inspect](/pdf/inspect/) | Page, encryption, form, annotation, attachment, active-content, signature-marker, and rewrite-readiness evidence |
| Organize | [Merge PDFs](/pdf/merge/), [split PDFs](/pdf/split/), [extract PDF pages](/pdf/extract-pages/), [delete PDF pages](/pdf/delete-pages/), [reorder PDF pages](/pdf/reorder-pages/), and [rotate PDF pages](/pdf/rotate-pages/) without changing the selected source files |
| [Optimize](/pdf/optimize/) | Deterministic lossless optimization and Fast Web View profiles without rasterizing pages |
| Protect or unlock | [AES-256 Standard password protection](/pdf/protect/), or a [separate unprotected copy](/pdf/unlock/) |
| [Redact](/pdf/redact/) | Literal text removal followed by checks of extracted text, encoded strings, and decoded streams |
| [Compare](/pdf/compare/) | A self-contained visual gallery with expected, actual, and highlighted page differences |

[Browse every PDF workflow](/pdf/), [open PDF tools](/apps/officeimo-converter/?workspace=pdf&tool=inspect), or download the [showcase PDF](/downloads/showcase/pdf/showcase-dashboard.pdf) used by the built-in samples.

## Convert PDF content into editable formats

The workspace also exposes [PDF to DOCX](/convert/pdf-to-word/), [detected PDF tables to XLSX](/convert/pdf-tables-to-excel/), [PDF to PPTX](/convert/pdf-to-powerpoint/), and [PDF to HTML](/convert/pdf-to-html/). These are logical imports, not a promise that every fixed-position page can become the original editable source. Word reconstructs semantic content, Excel imports detected tables, PowerPoint applies its editable page projection, and HTML creates reviewable semantic output. Each download includes diagnostics that identify approximated, visual-only, truncated, or omitted content where the selected adapter reports it.

For application code, use the focused adapter that owns the target:

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Word.Pdf;

PdfDocument pdf = PdfDocument.Open(File.ReadAllBytes("report.pdf"));
PdfWordConversionResult result = pdf.ToWordDocumentResult();
result.Value.Save("report.docx");

foreach (PdfConversionWarning warning in result.Report.Warnings) {
    Console.WriteLine($"{warning.Code}: {warning.Message}");
}
```

## Important boundaries

- Browser files are limited to 25 MiB each, ten PDFs and 75 MiB combined for multi-file tools, 500 parsed pages, 100 split outputs, and 25 pages per visual comparison. Generated artifacts are capped at 96 MiB, with a 64 MiB serialized-PDF ceiling for split archives.
- Optimization is deliberately lossless. Scan-oriented image downsampling and other lossy compression need a separate policy and evidence contract.
- OCR and searchable-PDF generation are provider-bound roadmap work. The core package remains dependency-light and does not pretend that image-only pages contain readable text.
- Password protection is not a digital signature. CMS-backed signing, timestamping, trust validation, and revision inspection remain .NET workflows through `OfficeIMO.Pdf` and `OfficeIMO.Security`.
- Filling a visual signature field is also different from applying a cryptographic signature. Applications should present those actions separately.
