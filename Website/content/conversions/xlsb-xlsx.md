---
title: "Convert XLSB and XLSX in .NET"
description: "Read, write, and convert Excel binary XLSB workbooks and modern XLSX files with OfficeIMO.Excel without automating Microsoft Excel."
meta.eyebrow: "Excel binary workbooks"
meta.source_format: "XLSB"
meta.destination_format: "XLSX"
meta.package: "OfficeIMO"
meta.package_url: "https://www.nuget.org/packages/OfficeIMO"
meta.runtime: ".NET on Windows, Linux, and macOS"
meta.primary_label: "Get OfficeIMO"
meta.primary_url: "https://www.nuget.org/packages/OfficeIMO"
meta.secondary_label: "Check Excel compatibility"
meta.secondary_url: "/compatibility/#excel"
meta.summary_title: "XLSB and XLSX summary"
meta.limit: "Binary and Open XML records differ. Review diagnostics for formulas, external links, drawings, charts, macros, and less common workbook records."
meta.related_label: "Explore all conversion guides"
meta.related_url: "/convert/guides/"
meta.howto.name: "Convert XLSB to XLSX with OfficeIMO"
meta.howto.description: "Use the first-party Excel engine to inspect and convert a binary workbook without Microsoft Excel."
meta.howto.steps:
  - name: "Install"
    text: "Add the OfficeIMO NuGet package."
  - name: "Analyze"
    text: "Preview XLSB-to-XLSX fidelity findings."
  - name: "Convert"
    text: "Call OfficeConverter.Convert for the selected paths."
  - name: "Verify"
    text: "Check the formulas, styles, links, and worksheet content important to the workload."
---

XLSB is a binary workbook format used for large or calculation-heavy Excel files. It is not the same container as XLSX, and renaming the extension is not a conversion. `OfficeIMO.Excel` includes a first-party XLSB parser and writer so applications can inspect and transform these workbooks through the normal Excel document model.

File-to-file conversion uses `OfficeConverter` from the `OfficeIMO` package.

## XLSB to XLSX

```csharp
using OfficeIMO;

OfficeConversionReport preview =
    OfficeConverter.Analyze("model.xlsb", "model.xlsx");

OfficeConversionReport result =
    OfficeConverter.Convert("model.xlsb", "model.xlsx");
```

## XLSX to XLSB

```csharp
using OfficeIMO;

OfficeConversionReport result =
    OfficeConverter.Convert("model.xlsx", "model.xlsb");
```

The format pair can represent many of the same workbook concepts, but their internal records differ. The report matters when the file contains advanced formulas, external links, drawings, charts, macros, or less common workbook records.

For archive conversion, retain the source and compare the business values that matter instead of relying only on “file opened successfully.” For programmatic ingestion, you can load XLSB directly and skip a permanent conversion when a normalized data or document projection is the real goal.

Review the current [Excel feature evidence](/compatibility/#excel) or use [OfficeIMO.Reader](/products/reader/) when the goal is normalized extraction rather than workbook editing.
