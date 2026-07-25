---
title: "Convert Word Documents to PDF in .NET"
description: "Publish OfficeIMO Word documents as PDF with first-party .NET packages, conversion warnings, explicit resource policy, and no Microsoft Word runtime."
meta.eyebrow: "Document publishing"
meta.source_format: "Word"
meta.destination_format: "PDF"
meta.package: "OfficeIMO.Word.Pdf"
meta.package_url: "https://www.nuget.org/packages/OfficeIMO.Word.Pdf"
meta.runtime: ".NET on Windows, Linux, and macOS"
meta.howto.name: "Save an OfficeIMO Word document as PDF"
meta.howto.description: "Load or create a Word document and publish it through the first-party OfficeIMO PDF engine."
meta.howto.steps:
  - "Install|Add OfficeIMO.Word.Pdf, which brings the focused Word and PDF integration."
  - "Load|Open a supported Word document with WordDocument.Load or create one."
  - "Publish|Call SaveAsPdf with the destination path."
  - "Inspect|Check the returned save result and conversion warnings."
---

`OfficeIMO.Word.Pdf` connects the Word document model to the first-party OfficeIMO PDF engine. It is intended for reports, invoices, letters, contracts, and server-generated documents where the application controls both the source content and publishing policy.

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using WordDocument document = WordDocument.Load("report.docx");
var result = document.SaveAsPdf("report.pdf");
```

The converter handles common document structures such as paragraphs, styled text, tables, images, headers, footers, lists, sections, and page settings. Conversion results expose warnings so a service can distinguish a successful file write from a perfect representation claim.

## Fonts and external resources

PDF output can depend on fonts, images, and other resources. Configure those deliberately for containers and production services. Avoid relying on a developer workstation's installed fonts if the same workload must render predictably on Linux.

## DOC input

For legacy DOC files, load or first convert through `OfficeIMO.Word`. The Word compatibility report remains the right place to inspect legacy import behavior; the PDF result covers the publishing stage. Keeping those stages separate makes troubleshooting much clearer.

For layout-sensitive output, compare representative PDFs against approved baselines. See the [Word guide](/docs/word/), [PDF product](/products/pdf/), and [document publishing workflow](/solutions/document-publishing/).
