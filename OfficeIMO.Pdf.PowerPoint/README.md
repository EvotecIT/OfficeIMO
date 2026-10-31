# OfficeIMO.Pdf.PowerPoint

Reconstruct PowerPoint content from supported PDF pages.

```sh
dotnet add package OfficeIMO.Pdf.PowerPoint
```

## Import PDF pages

The general PDF route reconstructs supported page content as native slide objects. This is a new semantic projection, not recovery of the original slide deck: unsupported or ambiguous content remains explicit in the conversion report.

```csharp
using OfficeIMO;
using OfficeIMO.Pdf.PowerPoint;
using OfficeIMO.Pdf;

PdfDocument pdf = PdfDocument.Load("handout.pdf");
OfficeConversionReport report = pdf.SaveAsPowerPoint("handout-editable.pptx");

foreach (OfficeDiagnostic finding in report.Diagnostics) {
    Console.WriteLine($"{finding.Code}: {finding.Message}");
}
```

Use the explicit visual profile when a page image is the intended result. Each image is movable and resizable, but text, vectors, charts, and tables inside it are not editable:

```csharp
using OfficeIMO;
var visual = PowerPointPdfImportOptions.CreateVisualPages();
OfficeConversionReport visualReport = pdf.SaveAsPowerPoint(
    "handout-visual.pptx",
    visual);

foreach (OfficeDiagnostic finding in visualReport.Diagnostics) {
    Console.WriteLine($"{finding.Code}: {finding.Message}");
}
```

This is a new semantic projection, not recovery of the original slide deck. Original charts, groups, themes, animations, notes, and authoring intent cannot be recovered reliably from arbitrary PDFs; omissions and simplifications remain explicit warnings. `OfficeConversionReport.HasLoss` and `ThrowIfLoss()` use typed approximation, omission, and failure evidence. Editable-content and visual-page modes therefore report their inherent reconstruction or non-editability instead of claiming lossless conversion.

Use hybrid mode when the original page must remain visible while detected tables stay editable. Row and column caps split a large overlay across duplicate visual-page slides, and each overlay keeps the same centered, aspect-preserving page geometry as its background:

```csharp
using OfficeIMO;
var hybrid = PowerPointPdfImportOptions.CreateHybrid();
hybrid.MaxRowsPerSlide = 18;
hybrid.MaxColumnsPerSlide = 6;

OfficeConversionReport hybridReport = pdf.SaveAsPowerPoint(
    "handout-hybrid.pptx",
    hybrid);

Console.WriteLine($"Editable table segments: {hybridReport.ItemsWritten}");
Console.WriteLine($"Visual-only page content: {hybridReport.HasLoss}");
```

Use editable-table mode when detected data is more important than page appearance:

```csharp
using OfficeIMO;
var options = PowerPointPdfImportOptions.CreateEditableTables();
options.ReadOptions = new PdfLoadOptions {
    Profile = PdfReadProfile.Structured,
    PageSelection = PdfPageSelection.Parse("1-3")
};
options.MaxRows = 400;
options.MaxRowsPerSlide = 18;
options.MaxColumnsPerSlide = 6;
options.MergePageContinuations = true;

OfficeConversionReport report = pdf.SaveAsPowerPoint(
    "financial-statement-tables.pptx",
    options);

foreach (OfficeDiagnostic finding in report.Diagnostics) {
    Console.WriteLine($"{finding.Code}: {finding.Message}");
}

Console.WriteLine($"Non-table page content detected: {report.HasLoss}");
```

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 1 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Pdf.PowerPoint` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
