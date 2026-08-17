---
title: "Extract selected pages from a PDF"
description: "Extract ranges such as 1-3,5,last into a new PDF, preserve the selected order, and leave the original document unchanged in the browser or .NET."
meta.workflow_id: "extract"
meta.eyebrow: "Organize PDF pages"
meta.source_format: "PDF and page selection"
meta.destination_format: "Extracted PDF"
meta.package: "OfficeIMO.Pdf"
meta.package_url: "https://www.nuget.org/packages/OfficeIMO.Pdf"
meta.runtime: "Browser-local WebAssembly or .NET"
meta.primary_url: "/apps/officeimo-converter/?workspace=pdf&tool=extract"
meta.primary_label: "Extract PDF pages in the browser"
meta.secondary_url: "/docs/pdf/operations/"
meta.secondary_label: "Read the .NET guide"
meta.summary_title: "Extraction summary"
meta.limit: "Page expressions are one-based and resolved against a PDF of at most 500 parsed pages in the browser."
meta.related_url: "/pdf/"
meta.related_label: "Browse all PDF workflows"
meta.howto.name: "Extract a page range into a new PDF"
meta.howto.description: "Resolve a one-based page expression and write only the selected pages to a separate document."
meta.howto.steps:
  - name: "Select"
    text: "Choose one PDF and review its page count."
  - name: "Describe"
    text: "Enter a one-based expression such as 1-3,5,last."
  - name: "Extract"
    text: "Create and download the new PDF plus its operation report."
---

Use extraction when the output should contain only a specific subset of pages. The selector supports individual pages, ranges, comma-separated combinations, and `last`, which is resolved against the current source.

## Extract from .NET

```csharp
using OfficeIMO.Pdf;

PdfDocument source = PdfDocument.Open("case-file.pdf");
PdfPageSelector selection = PdfPageSelector.Parse("1-3,5,last");
PdfDocument extracted = source.Pages.Extract(selection);

File.WriteAllBytes("case-file.extract.pdf", extracted.ToBytes());
```

The result is a new PDF. Page numbers in the output are compacted to the new document order, while the source file remains unchanged.

## Choose the right page operation

Extraction keeps the selected subset. [Delete pages](/pdf/delete-pages/) does the inverse and keeps everything except the selection. [Reorder pages](/pdf/reorder-pages/) creates a document using the supplied sequence. Keeping these operations separate makes automated policy and user confirmation much clearer.
