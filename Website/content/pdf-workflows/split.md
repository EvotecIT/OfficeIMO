---
title: "Split a PDF into smaller files"
description: "Split a PDF into consecutive page groups, download the generated documents as a ZIP archive, and keep the original file unchanged in your browser."
meta.workflow_id: "split"
meta.eyebrow: "Organize PDF pages"
meta.source_format: "PDF"
meta.destination_format: "ZIP of PDF parts"
meta.package: "OfficeIMO.Pdf"
meta.package_url: "https://www.nuget.org/packages/OfficeIMO.Pdf"
meta.runtime: "Browser-local WebAssembly or .NET"
meta.primary_url: "/apps/officeimo-converter/?workspace=pdf&tool=split"
meta.primary_label: "Split a PDF in the browser"
meta.secondary_url: "/docs/pdf/operations/"
meta.secondary_label: "Read the .NET guide"
meta.summary_title: "Split summary"
meta.limit: "The browser creates at most 100 parts and caps serialized split output at 64 MiB."
meta.related_url: "/pdf/"
meta.related_label: "Browse all PDF workflows"
meta.howto.name: "Split a PDF into consecutive parts"
meta.howto.description: "Choose the maximum pages per part and package the resulting PDF files into one download."
meta.howto.steps:
  - name: "Select"
    text: "Choose one PDF and confirm its page count."
  - name: "Size"
    text: "Enter the number of consecutive pages to place in each output document."
  - name: "Download"
    text: "Create the parts and download them together as a ZIP archive."
---

Splitting creates consecutive page groups: a 10-page source split at three pages produces ranges 1-3, 4-6, 7-9, and 10. The final part may contain fewer pages.

## Split from .NET

```csharp
using OfficeIMO.Pdf;

PdfDocument source = PdfDocument.Open("book.pdf");
int pageCount = source.Inspect().PageCount;
const int pagesPerPart = 10;

for (int first = 1, part = 1; first <= pageCount; first += pagesPerPart, part++) {
    int last = Math.Min(pageCount, first + pagesPerPart - 1);
    PdfPageSelector pages = PdfPageSelector.Parse($"{first}-{last}");
    File.WriteAllBytes($"book.part-{part:000}.pdf", source.Pages.Extract(pages).ToBytes());
}
```

Application code decides how to name and store individual outputs. The browser packages them into ZIP because browsers handle one bounded download more predictably than a burst of separate files.

## Boundaries

Splitting does not infer chapters, bookmarks, invoices, or logical document boundaries. Use [extract pages](/pdf/extract-pages/) when the required ranges are not consecutive fixed-size groups. Signed and encrypted sources also require deliberate policy because each output is a newly written document.
