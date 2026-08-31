---
title: "Optimize a PDF without rasterizing pages"
description: "Apply deterministic lossless PDF optimization, deduplication, compression, or Fast Web View and retain the original when a candidate is not smaller."
meta.workflow_id: "optimize"
meta.eyebrow: "Publish a PDF"
meta.source_format: "PDF"
meta.destination_format: "Losslessly optimized PDF"
meta.package: "OfficeIMO.Pdf"
meta.package_url: "https://www.nuget.org/packages/OfficeIMO.Pdf"
meta.runtime: "Browser-local WebAssembly or .NET"
meta.primary_url: "/apps/officeimo-converter/?workspace=pdf&tool=optimize"
meta.primary_label: "Optimize a PDF in the browser"
meta.secondary_url: "/docs/pdf/operations/"
meta.secondary_label: "Read the .NET guide"
meta.summary_title: "Optimization summary"
meta.limit: "OfficeIMO optimization is lossless; it does not downsample scans or replace pages with images."
meta.related_url: "/pdf/"
meta.related_label: "Browse all PDF workflows"
meta.howto.name: "Optimize a PDF with an explicit profile"
meta.howto.description: "Choose a lossless policy, inspect the candidate result, and keep the smaller safe artifact."
meta.howto.steps:
  - name: "Select"
    text: "Choose one PDF and note its original size."
  - name: "Profile"
    text: "Choose Balanced, Maximum Compression, Web, or Archival behavior."
  - name: "Review"
    text: "Download the result and inspect saved bytes, actions, skipped opportunities, and linearization evidence."
---

OfficeIMO optimization works on PDF structure and streams without rasterizing pages. Text remains text, and the operation does not deliberately trade fidelity for a smaller scan image.

## Optimize from .NET

```csharp
using OfficeIMO.Pdf;

PdfDocument source = PdfDocument.Load("report.pdf");
PdfOptimizationActionResult result =
    source.Optimization.Apply(PdfOptimizationProfile.Web);

File.WriteAllBytes("report.optimized.pdf", result.Bytes);
Console.WriteLine($"Saved bytes: {result.SavedBytes}");
Console.WriteLine($"Returned original: {result.ReturnedOriginal}");
```

If the optimized candidate is not smaller, the result can retain the original bytes rather than making the file larger merely to claim success. The action report records requested profile, candidate and returned sizes, applied actions, skipped opportunities, and linearization state.

## Compression boundaries

This is not scan downsampling, JPEG recompression, or page rasterization. Those policies can remove selectable text, accessibility information, vector quality, and signature meaning. OfficeIMO keeps that lossy workflow separate until it has explicit quality controls and evidence.
