---
title: "Compare two PDF files visually"
description: "Compare two PDFs locally, review structural findings, and download an HTML gallery containing expected, actual, and highlighted page-difference images."
meta.workflow_id: "compare"
meta.eyebrow: "Understand a PDF"
meta.source_format: "Two PDFs"
meta.destination_format: "HTML comparison gallery"
meta.package: "OfficeIMO.Pdf"
meta.package_url: "https://www.nuget.org/packages/OfficeIMO.Pdf"
meta.runtime: "Browser-local WebAssembly or .NET"
meta.primary_url: "/apps/officeimo-converter/?workspace=pdf&tool=compare"
meta.primary_label: "Compare PDFs in the browser"
meta.secondary_url: "/docs/pdf/operations/"
meta.secondary_label: "Read the .NET guide"
meta.summary_title: "Comparison summary"
meta.limit: "The browser compares at most 25 pages and applies bounded image, pixel, and output budgets."
meta.related_url: "/pdf/"
meta.related_label: "Browse all PDF workflows"
meta.howto.name: "Compare two PDFs locally"
meta.howto.description: "Render both documents under the same bounded policy and inspect exact visual and structural differences."
meta.howto.steps:
  - name: "Choose"
    text: "Select the expected PDF first and the actual PDF second."
  - name: "Compare"
    text: "Render corresponding pages and evaluate structural and pixel-level differences."
  - name: "Review"
    text: "Open or download the self-contained HTML gallery to compare expected, actual, and highlighted images."
---

PDF comparison is useful for regression baselines, generated reports, invoices, and publishing pipelines where a successful file write is not enough. The comparison report keeps structural findings beside rendered evidence so a changed page count is not reduced to a pixel score.

## Compare from .NET

```csharp
using OfficeIMO.Pdf;

byte[] expected = File.ReadAllBytes("approved.pdf");
byte[] actual = File.ReadAllBytes("candidate.pdf");

PdfVisualComparisonReport report = PdfVisualComparer.Compare(expected, actual);
File.WriteAllText("comparison.html", report.ToHtmlGallery("Release comparison"));

Console.WriteLine($"Match: {report.IsMatch}");
```

Application code can configure page selection, channel tolerance, allowed difference ratio, and resource budgets. The browser route deliberately uses exact comparison thresholds and fixed limits so an unexpectedly large document cannot consume unbounded memory.

## Read the evidence correctly

A visual match means the compared pages satisfied the selected rendering threshold. It does not prove semantic equality, identical metadata, equivalent signatures, or byte-for-byte identity. Conversely, a visual difference may be a legitimate font, rendering, or pagination change. Review the gallery and structural findings before accepting or rejecting the candidate.
