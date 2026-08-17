---
title: "Delete selected pages from a PDF copy"
description: "Remove selected PDF pages from a newly generated copy after explicit confirmation, while keeping the original browser-selected document unchanged."
meta.workflow_id: "delete"
meta.eyebrow: "Organize PDF pages"
meta.source_format: "PDF and page selection"
meta.destination_format: "PDF without selected pages"
meta.package: "OfficeIMO.Pdf"
meta.package_url: "https://www.nuget.org/packages/OfficeIMO.Pdf"
meta.runtime: "Browser-local WebAssembly or .NET"
meta.primary_url: "/apps/officeimo-converter/?workspace=pdf&tool=delete"
meta.primary_label: "Delete pages in the browser"
meta.secondary_url: "/docs/pdf/operations/"
meta.secondary_label: "Read the .NET guide"
meta.summary_title: "Deletion summary"
meta.limit: "The browser requires explicit confirmation and never overwrites the selected source file."
meta.related_url: "/pdf/"
meta.related_label: "Browse all PDF workflows"
meta.howto.name: "Remove pages from a downloaded PDF copy"
meta.howto.description: "Choose the pages to exclude, confirm the permanent change to the output, and create a separate document."
meta.howto.steps:
  - name: "Select"
    text: "Choose one PDF and enter a one-based page expression."
  - name: "Confirm"
    text: "Acknowledge that the selected pages will not exist in the downloaded copy."
  - name: "Create"
    text: "Generate the new PDF and review the source and output page counts."
---

Deletion is an output transformation, not an in-place edit. The browser requires confirmation because the selected pages are permanently absent from the downloaded artifact, even though the source remains untouched.

## Delete pages from .NET

```csharp
using OfficeIMO.Pdf;

PdfDocument source = PdfDocument.Open("packet.pdf");
PdfPageSelector pagesToRemove = PdfPageSelector.Parse("2,4-6");
PdfDocument result = source.Pages.Delete(pagesToRemove);

File.WriteAllBytes("packet.cleaned.pdf", result.ToBytes());
```

Validate the output page count and required content before replacing or archiving any original outside OfficeIMO. Application code owns storage and overwrite policy; the PDF API returns a new document model.

## When extraction is clearer

If the business rule identifies pages to keep, use [extract pages](/pdf/extract-pages/) instead. Expressing the positive selection often makes retention and disclosure workflows easier to review than maintaining a growing exclusion list.
