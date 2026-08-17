---
title: "Reorder pages in a PDF"
description: "Create a new PDF with every source page exactly once in a supplied one-based sequence, and keep the selected source file unchanged."
meta.workflow_id: "reorder"
meta.eyebrow: "Organize PDF pages"
meta.source_format: "PDF and page sequence"
meta.destination_format: "Reordered PDF"
meta.package: "OfficeIMO.Pdf"
meta.package_url: "https://www.nuget.org/packages/OfficeIMO.Pdf"
meta.runtime: "Browser-local WebAssembly or .NET"
meta.primary_url: "/apps/officeimo-converter/?workspace=pdf&tool=reorder"
meta.primary_label: "Reorder pages in the browser"
meta.secondary_url: "/docs/pdf/operations/"
meta.secondary_label: "Read the .NET guide"
meta.summary_title: "Reorder summary"
meta.limit: "The one-based sequence must include every source page exactly once; subsets, omissions, and duplicates are rejected."
meta.related_url: "/pdf/"
meta.related_label: "Browse all PDF workflows"
meta.howto.name: "Create a PDF with a new page order"
meta.howto.description: "Describe one complete permutation of the source pages and write the resolved order into a separate document."
meta.howto.steps:
  - name: "Select"
    text: "Choose one PDF and review the current page order."
  - name: "Sequence"
    text: "Enter every source page exactly once in the desired order, such as 3,1,2,4-last."
  - name: "Create"
    text: "Generate the reordered PDF and download its operation report."
---

Reordering writes pages in the sequence supplied by the caller. That sequence must be a full permutation: every source page appears exactly once, with no omissions or duplicates. It is useful for corrected scan order, assembled packets, covers placed after generation, and workflows that need a deterministic page tree before signing.

## Reorder from .NET

```csharp
using OfficeIMO.Pdf;

PdfDocument source = PdfDocument.Open("scan.pdf");
PdfPageSelector order = PdfPageSelector.Parse("3,1,2,4-last");
PdfDocument reordered = source.Pages.Reorder(order);

File.WriteAllBytes("scan.reordered.pdf", reordered.ToBytes());
```

The operation creates a new document and preserves the page count. Use [extract pages](/pdf/extract-pages/) or [delete pages](/pdf/delete-pages/) when the output should contain only a subset.

## Separate order from rotation

Reordering changes where pages appear; it does not change their orientation. Apply [rotate pages](/pdf/rotate-pages/) separately so each operation and report says exactly what changed.
