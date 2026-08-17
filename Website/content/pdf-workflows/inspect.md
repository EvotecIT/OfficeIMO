---
title: "Inspect a PDF in your browser or .NET"
description: "Inspect PDF pages, encryption, forms, annotations, attachments, active content, signature markers, and rewrite readiness without changing the file."
meta.workflow_id: "inspect"
meta.eyebrow: "Understand a PDF"
meta.source_format: "PDF"
meta.destination_format: "JSON report"
meta.package: "OfficeIMO.Pdf"
meta.package_url: "https://www.nuget.org/packages/OfficeIMO.Pdf"
meta.runtime: "Browser-local WebAssembly or .NET"
meta.primary_url: "/apps/officeimo-converter/?workspace=pdf&tool=inspect"
meta.primary_label: "Inspect a PDF in the browser"
meta.secondary_url: "/docs/pdf/operations/"
meta.secondary_label: "Read the .NET guide"
meta.summary_title: "Inspection summary"
meta.limit: "The browser accepts one PDF up to 25 MiB and applies bounded parsing limits."
meta.related_url: "/pdf/"
meta.related_label: "Browse all PDF workflows"
meta.howto.name: "Inspect a PDF without changing it"
meta.howto.description: "Create a bounded preflight report before deciding whether to extract, rewrite, or reject a document."
meta.howto.steps:
  - name: "Select"
    text: "Choose one PDF or load the built-in sample; the file remains in the current tab."
  - name: "Inspect"
    text: "Run the preflight reader to collect document, security, form, active-content, and rewrite evidence."
  - name: "Review"
    text: "Read the visible readiness messages and download the JSON report when the evidence belongs with a job record."
---

Use inspection before a workflow assumes that a PDF can be read, searched, or rewritten. The report distinguishes basic read access from rewrite readiness and records blockers instead of treating every file as an ordinary unencrypted document.

## Inspect from .NET

```csharp
using OfficeIMO.Pdf;

byte[] input = File.ReadAllBytes("document.pdf");
PdfDocumentPreflight report = PdfDocument.Preflight(input);

Console.WriteLine($"Readable: {report.CanRead}");
Console.WriteLine($"Rewritable: {report.CanRewrite}");
Console.WriteLine($"Pages: {report.DocumentInfo?.PageCount}");
foreach (string diagnostic in report.Diagnostics) {
    Console.WriteLine(diagnostic);
}
```

Inspection reports page count, encryption, forms, annotations, attachments, tagged content, active content, signature markers, and the current capability gates. A signature marker is evidence that signing structures exist; it is not by itself a trust decision or certificate validation result.

## What it does not do

Inspection does not decrypt a file without valid credentials, certify that a signature is trusted, run OCR on image-only pages, or prove that every visual feature can survive a rewrite. Use the reported blockers and diagnostics as policy input rather than a blanket safety label.
