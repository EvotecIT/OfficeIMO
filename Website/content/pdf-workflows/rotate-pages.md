---
title: "Rotate selected PDF pages"
description: "Rotate selected PDF pages by 90, 180, or 270 degrees, create a separate output document, and preserve the original browser-selected file."
meta.workflow_id: "rotate"
meta.eyebrow: "Organize PDF pages"
meta.source_format: "PDF and page selection"
meta.destination_format: "Rotated PDF"
meta.package: "OfficeIMO.Pdf"
meta.package_url: "https://www.nuget.org/packages/OfficeIMO.Pdf"
meta.runtime: "Browser-local WebAssembly or .NET"
meta.primary_url: "/apps/officeimo-converter/?workspace=pdf&tool=rotate"
meta.primary_label: "Rotate pages in the browser"
meta.secondary_url: "/docs/pdf/operations/"
meta.secondary_label: "Read the .NET guide"
meta.summary_title: "Rotation summary"
meta.limit: "Browser rotation accepts 90, 180, or 270 degrees and a one-based page expression."
meta.related_url: "/pdf/"
meta.related_label: "Browse all PDF workflows"
meta.howto.name: "Rotate selected pages in a PDF copy"
meta.howto.description: "Resolve the selected pages, apply a supported clockwise rotation, and write a separate PDF."
meta.howto.steps:
  - name: "Select"
    text: "Choose one PDF and enter pages such as 1,3-5,last."
  - name: "Rotate"
    text: "Choose 90, 180, or 270 degrees."
  - name: "Download"
    text: "Create the rotated PDF and review the resolved selector in the report."
---

Rotation updates the selected pages in a newly written document. It is intended for sideways scans, mixed-orientation packets, and deterministic publishing workflows.

## Rotate from .NET

```csharp
using OfficeIMO.Pdf;

PdfDocument source = PdfDocument.Open("scans.pdf");
PdfPageSelector pages = PdfPageSelector.Parse("1,3-5,last");
PdfDocument rotated = source.Pages.Rotate(90, pages);

File.WriteAllBytes("scans.rotated.pdf", rotated.ToBytes());
```

The browser report records the resolved selector, source and output page counts, and rotation angle. The source file is never overwritten.

## Rotation is not page reflow

This operation changes page orientation metadata and page-space presentation. It does not reinterpret text, crop content, deskew scanned images, or rebuild the document layout. Image deskewing and OCR require different processing contracts.
