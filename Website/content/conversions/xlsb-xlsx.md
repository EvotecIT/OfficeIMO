---
title: "Convert XLSB and XLSX in .NET"
description: "Read, write, and convert Excel binary XLSB workbooks and modern XLSX files with OfficeIMO.Excel without automating Microsoft Excel."
meta.eyebrow: "Excel binary workbooks"
meta.source_format: "XLSB"
meta.destination_format: "XLSX"
meta.package: "OfficeIMO.Excel"
meta.package_url: "https://www.nuget.org/packages/OfficeIMO.Excel"
meta.runtime: ".NET on Windows, Linux, and macOS"
meta.howto.name: "Convert XLSB to XLSX with OfficeIMO.Excel"
meta.howto.description: "Use the first-party Excel engine to inspect and convert a binary workbook without Microsoft Excel."
meta.howto.steps:
  - "Install|Add the OfficeIMO.Excel NuGet package."
  - "Analyze|Preview XLSB-to-XLSX fidelity findings."
  - "Convert|Call ExcelDocument.Convert for the selected paths."
  - "Verify|Check the formulas, styles, links, and worksheet content important to the workload."
---

XLSB is a binary workbook format used for large or calculation-heavy Excel files. It is not the same container as XLSX, and renaming the extension is not a conversion. `OfficeIMO.Excel` includes a first-party XLSB parser and writer so applications can inspect and transform these workbooks through the normal Excel document model.

## XLSB to XLSX

```csharp
using OfficeIMO.Excel;

ExcelDocumentConversionReport preview =
    ExcelDocument.AnalyzeConversion("model.xlsb", "model.xlsx");

ExcelDocumentConversionResult result =
    ExcelDocument.Convert("model.xlsb", "model.xlsx");
```

## XLSX to XLSB

```csharp
using OfficeIMO.Excel;

ExcelDocumentConversionResult result =
    ExcelDocument.Convert("model.xlsx", "model.xlsb");
```

The format pair can represent many of the same workbook concepts, but their internal records differ. The report matters when the file contains advanced formulas, external links, drawings, charts, macros, or less common workbook records.

For archive conversion, retain the source and compare the business values that matter instead of relying only on “file opened successfully.” For programmatic ingestion, you can load XLSB directly and skip a permanent conversion when a normalized data or document projection is the real goal.

Review the current [Excel feature evidence](/compatibility/#excel) or use [OfficeIMO.Reader](/products/reader/) when the goal is normalized extraction rather than workbook editing.
