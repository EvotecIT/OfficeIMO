---
title: "Convert Excel Workbooks to PDF in .NET"
description: "Publish Excel worksheets and workbooks as PDF with page setup, styles, images, links, and conversion diagnostics using OfficeIMO.Excel.Pdf."
meta.eyebrow: "Spreadsheet publishing"
meta.source_format: "Excel"
meta.destination_format: "PDF"
meta.package: "OfficeIMO.Excel.Pdf"
meta.package_url: "https://www.nuget.org/packages/OfficeIMO.Excel.Pdf"
meta.runtime: ".NET on Windows, Linux, and macOS"
meta.howto.name: "Save an OfficeIMO Excel workbook as PDF"
meta.howto.description: "Load or create an Excel workbook and publish selected worksheet content through the first-party PDF engine."
meta.howto.steps:
  - name: "Install"
    text: "Add the OfficeIMO.Excel.Pdf NuGet package."
  - name: "Load"
    text: "Open a supported workbook with ExcelDocument.Load or create one."
  - name: "Configure"
    text: "Choose worksheets, layout, page setup, and resource policy when needed."
  - name: "Publish"
    text: "Call SaveAsPdf and inspect the result."
---

`OfficeIMO.Excel.Pdf` publishes workbook content through the first-party OfficeIMO PDF engine. It is useful for statements, tabular reports, operational exports, and scheduled jobs where a stable read-only deliverable is more appropriate than an editable spreadsheet.

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;

using ExcelDocument workbook = ExcelDocument.Load("report.xlsx");
var result = workbook.SaveAsPdf("report.pdf");
```

The converter can carry worksheet values, styles, merged cells, images, links, dimensions, page settings, and selected formula results into PDF output. Its save result and report make warnings visible to the calling application.

## Pagination is part of the design

A worksheet is a large canvas; a PDF is a sequence of pages. Decide whether the output should preserve worksheet-like geometry or optimize a selected table or range for readable pages. Test wide sheets, hidden rows and columns, manual page breaks, repeated headings, and printed ranges that matter to the business.

## XLS and XLSB input

Load supported legacy XLS or binary XLSB files through `OfficeIMO.Excel`, then publish the in-memory workbook. When the source format has compatibility findings, retain the Excel-stage report alongside the PDF-stage report.

For more detail, use the [Excel guide](/docs/excel/), [Excel compatibility evidence](/compatibility/#excel), and [document publishing workflow](/solutions/document-publishing/).
