---
title: "Merge PDF files locally"
description: "Merge two to ten PDF files in your selected order, keep the source files unchanged, and download a separate PDF with page and policy evidence."
meta.workflow_id: "merge"
meta.eyebrow: "Organize PDF pages"
meta.source_format: "Two to ten PDFs"
meta.destination_format: "Merged PDF"
meta.package: "OfficeIMO.Pdf"
meta.package_url: "https://www.nuget.org/packages/OfficeIMO.Pdf"
meta.runtime: "Browser-local WebAssembly or .NET"
meta.primary_url: "/apps/officeimo-converter/?workspace=pdf&tool=merge"
meta.primary_label: "Merge PDFs in the browser"
meta.secondary_url: "/docs/pdf/operations/"
meta.secondary_label: "Read the .NET guide"
meta.summary_title: "Merge summary"
meta.limit: "The browser accepts two to ten PDFs, up to 25 MiB each and 75 MiB combined."
meta.related_url: "/pdf/"
meta.related_label: "Browse all PDF workflows"
meta.howto.name: "Merge PDF files in a controlled order"
meta.howto.description: "Combine complete documents through one first-party merge pass and retain the operation report."
meta.howto.steps:
  - name: "Select"
    text: "Choose between two and ten PDF files."
  - name: "Order"
    text: "Arrange the files in the sequence that should appear in the output."
  - name: "Merge"
    text: "Create and download the merged PDF and its JSON operation report."
---

The merge workflow copies complete page trees into a new document. The browser presents the selected order explicitly and returns output page count and policy decisions with the artifact.

## Merge from .NET

```csharp
using OfficeIMO.Pdf;

PdfDocument first = PdfDocument.Load("cover.pdf");
PdfDocument second = PdfDocument.Load("report.pdf");

PdfMergeResult result = PdfDocument.MergeResult(
    new PdfMergeOptions(),
    first,
    second);

File.WriteAllBytes("combined.pdf", result.ToBytes());
Console.WriteLine($"Pages: {result.Report.OutputPageCount}");
```

`PdfMergeOptions` controls how the application handles incoming metadata, destinations, forms, attachments, encryption, and signature-related evidence. Keep the report when those choices affect compliance or downstream review.

## Boundaries

The browser route does not modify any selected file. It creates one new PDF and a companion report. Encrypted inputs still require a supported authentication context, and merging a signed document creates a new revision-independent artifact rather than extending or preserving the original signature guarantee.
