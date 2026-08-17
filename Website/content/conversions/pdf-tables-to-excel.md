---
title: "Import PDF tables into Excel"
description: "Detect logical tables in a PDF and import them into an editable XLSX with explicit warnings when page content is not tabular or rows exceed configured limits."
meta.workflow_id: "pdf-xlsx"
meta.eyebrow: "PDF table import"
meta.source_format: "PDF tables"
meta.destination_format: "XLSX"
meta.package: "OfficeIMO.Excel.Pdf"
meta.package_url: "https://www.nuget.org/packages/OfficeIMO.Excel.Pdf"
meta.runtime: "Browser-local WebAssembly or .NET"
meta.primary_url: "/apps/officeimo-converter/?workspace=convert&route=pdf-xlsx"
meta.primary_label: "Import PDF tables in the browser"
meta.secondary_url: "/docs/pdf/conversion/"
meta.secondary_label: "Read the conversion guide"
meta.summary_title: "PDF tables to Excel summary"
meta.limit: "The route imports detected tables only; prose, drawings, and other page content do not become worksheet cells."
meta.related_url: "/pdf/"
meta.related_label: "Browse PDF tools and imports"
meta.howto.name: "Import detected PDF tables into Excel"
meta.howto.description: "Analyze logical table structures, create an editable workbook, and report omitted or truncated content."
meta.howto.steps:
  - name: "Select"
    text: "Choose a readable PDF containing table-like content."
  - name: "Detect"
    text: "Analyze logical rows and cells under the configured import limits."
  - name: "Review"
    text: "Download the XLSX and inspect warnings for non-tabular page content or truncated rows."
---

This route is intentionally named PDF tables to Excel. It does not pretend that an entire page—paragraphs, images, charts, and drawing commands—maps naturally to worksheet cells.

## Import from .NET

```csharp
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Pdf;

PdfDocument pdf = PdfDocument.Open("statement.pdf");
PdfExcelTableImportResult result = pdf.ImportTablesToExcelDocumentResult();
using var workbook = result.Value;

File.WriteAllBytes("statement.tables.xlsx", workbook.ToBytes());
foreach (var entry in result.Report.Entries) {
    Console.WriteLine(entry);
}
```

The result records detected tables and whether page content remained outside the table-only route. Applications can choose stricter limits and reject a workbook when omitted content or truncation is unacceptable.

## Tables are inferred evidence

Many PDFs draw table-like layouts without storing formal table semantics. Detection depends on readable text, positions, and the logical structures available in the source. Always verify representative rows, headers, merged regions, and numeric values before using the workbook for calculations or decisions.
