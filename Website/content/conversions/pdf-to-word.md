---
title: "Convert PDF to Word in .NET or your browser"
description: "Convert readable PDF content into an editable DOCX, retain conversion warnings, and avoid claiming that fixed page geometry becomes the original Word file."
meta.workflow_id: "pdf-docx"
meta.eyebrow: "PDF logical import"
meta.source_format: "PDF"
meta.destination_format: "DOCX"
meta.package: "OfficeIMO.Word.Pdf"
meta.package_url: "https://www.nuget.org/packages/OfficeIMO.Word.Pdf"
meta.runtime: "Browser-local WebAssembly or .NET"
meta.primary_url: "/apps/officeimo-converter/?workspace=convert&route=pdf-docx"
meta.primary_label: "Convert PDF to Word in the browser"
meta.secondary_url: "/docs/pdf/conversion/"
meta.secondary_label: "Read the conversion guide"
meta.summary_title: "PDF to Word summary"
meta.limit: "This is a logical-content import, not pixel-perfect reconstruction of the application that created the PDF."
meta.related_url: "/pdf/"
meta.related_label: "Browse PDF tools and imports"
meta.howto.name: "Import PDF logical content into a Word document"
meta.howto.description: "Read the PDF, project supported structure into DOCX, and retain every conversion warning."
meta.howto.steps:
  - name: "Select"
    text: "Choose a readable PDF or load the built-in product sample."
  - name: "Convert"
    text: "Project logical PDF content into an editable Word document."
  - name: "Review"
    text: "Download the DOCX and companion report, then inspect warnings before delivery."
---

PDF stores positioned page content, not the original Word paragraphs, styles, sections, and editing history. `OfficeIMO.Word.Pdf` reconstructs supported logical content into a new DOCX and reports approximated, visual-only, truncated, or omitted material.

## Convert from .NET

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Word.Pdf;

PdfDocument pdf = PdfDocument.Open("report.pdf");
PdfWordConversionResult result = pdf.ToWordDocumentResult();
using var document = result.Value;

File.WriteAllBytes("report.docx", document.ToBytes());
foreach (PdfConversionWarning warning in result.Report.Warnings) {
    Console.WriteLine($"{warning.Code}: {warning.Message}");
}
```

The browser route uses the same adapter and downloads both the DOCX and a JSON report with source and output fingerprints, page count, timing, fidelity status, and structured warnings.

## Set the right expectation

Use this route when editable semantic content is more important than recreating every fixed-position detail. Complex columns, unusual font encodings, drawing instructions, and scanned pages may not recover as Word-native structures. Image-only documents need a separate OCR policy before logical import can find text.
