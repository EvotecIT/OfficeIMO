---
title: "OfficeIMO.IWork"
description: "Read Apple Pages, Numbers, and Keynote files, create native packages, and convert them to Word, Excel, or PowerPoint with explicit loss reports."
layout: product
meta.seo_title: "Apple Pages, Numbers, and Keynote for .NET"
product_color: "#db2777"
product_label: "Apple iWork engine"
install: "dotnet add package OfficeIMO.IWork"
nuget: "OfficeIMO.IWork"
docs_url: "/docs/iwork/"
api_url: "/api/iwork/"
meta.software.name: "OfficeIMO.IWork"
meta.software.application_category: "DeveloperApplication"
meta.software.operating_system: "Windows, Linux, macOS"
meta.software.download_url: "https://www.nuget.org/packages/OfficeIMO.IWork"
meta.software.price: 0
meta.software.price_currency: "USD"
---

## Apple documents without an Apple application

`OfficeIMO.IWork` reads modern Pages, Numbers, and Keynote packages, including ZIP files, directory bundles, and packages with a nested `Index.zip`, plus a read-only iWork ’09 XML profile. It decodes the native IWA archives with first-party readers and depends only on `OfficeIMO.Core`. Conversion to Office formats is opt-in: add `OfficeIMO.Word.IWork` for Pages, `OfficeIMO.Excel.IWork` for Numbers, or `OfficeIMO.PowerPoint.IWork` for Keynote. `OfficeIMO.Reader.IWork` adds the same sources to Reader ingestion. The Numbers example below uses the Excel adapter:

```bash
dotnet add package OfficeIMO.Excel.IWork
```

```csharp
using OfficeIMO.Excel.IWork;
using OfficeIMO.IWork;

IWorkSourceDocument source = IWorkSourceDocument.Load("budget.numbers");
IWorkNumbersProjection workbook = source.ReadNumbers();
Console.WriteLine($"{source.Kind}: {workbook.Sheets.Count} sheets");

using NumbersToExcelResult result = source.ToExcelResult(new IWorkConversionOptions {
    Mode = IWorkConversionMode.EditableOnly,
    AllowPartialEditableReconstruction = true
});
Console.WriteLine($"Partial: {result.Report.IsPartialEditableReconstruction}");
foreach (var diagnostic in result.Report.Diagnostics)
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
result.Value.Save("budget.xlsx");

using var deck = IWorkKeynoteDocument.Create();
deck.AddSlide().AddText("Quarterly review", 60, 100, 840, 120, fontSizePoints: 48);
deck.Save("review.key");
```

Supported content becomes editable DOCX, XLSX, or PPTX. Anything the adapter cannot represent is reported as loss. Partial reconstruction and visual previews are used only when the options allow them; Numbers cell padding, for example, has no XLSX equivalent and needs the partial opt-in shown above.

## Useful for

- Inspecting Pages, Numbers, and Keynote uploads on Windows, Linux, or macOS with bounded package, text, table, image, and formula work.
- Converting Pages to Word, Numbers to Excel, and Keynote to PowerPoint with typed source projections and per-unit loss reports.
- Creating native Keynote slides, Pages body documents, and Numbers tables with text, images, and scalar tables, without a seed file.
- Making bounded edits to a loaded modern file, such as same-length paragraph text, Numbers numeric and Boolean cells, or Keynote image frames, while keeping every other record intact.
- Extracting iWork content for search and retrieval through `OfficeIMO.Reader.IWork`.

This is semantic reconstruction, not a pixel-identical renderer. Charts, animations, vector effects, masks and crops stay in the source records and are reported rather than converted. Native creation and editing cover the narrow profiles listed in the support matrix. Embedded content is never executed and external resources are never fetched.

Start with the [iWork guide](/docs/iwork/) or check the [support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.iwork-support-matrix.md).
