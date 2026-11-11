# OfficeIMO.OpenDocument.Odp.Pdf

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.OpenDocument.Odp.Pdf)](https://www.nuget.org/packages/OfficeIMO.OpenDocument.Odp.Pdf)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.OpenDocument.Odp.Pdf?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.OpenDocument.Odp.Pdf)

Bidirectional ODP and PDF conversion without Word or Excel dependencies. ODP to PDF converts the presentation through [OfficeIMO.PowerPoint.OpenDocument](../OfficeIMO.PowerPoint.OpenDocument/README.md) and renders it with the [PowerPoint PDF engine](../OfficeIMO.PowerPoint.Pdf/README.md). PDF to ODP reconstructs slides through the shared PowerPoint projection.

## Install

```powershell
dotnet add package OfficeIMO.OpenDocument.Odp.Pdf
```

## Quick start

```csharp
using OfficeIMO;
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Odp.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;

OdpPresentation deck = OdpPresentation.Load("roadmap.odp");
OfficeConversionReport report = deck.SaveAsPdf("roadmap.pdf", new OdpPdfExportOptions {
    PdfOptions = new PowerPointPdfExportOptions { PageLayout = PowerPointPdfPageLayout.NotesPages }
});
foreach (OfficeDiagnostic finding in report.Diagnostics)
    Console.WriteLine($"{finding.Location}: {finding.LossKind} {finding.Message}");

PdfDocument pdf = PdfDocument.Load("handout.pdf");
OfficeConversionReport odpReport = pdf.SaveAsOdp("handout.odp", new PdfOdpExportOptions {
    PdfOptions = PdfPowerPointExportOptions.CreateHybrid()
});
```

Forward conversion takes `OdpPdfExportOptions` (`ConversionOptions` for the OpenDocument import, `PdfOptions` for the PDF engine, `FailOnLoss`); `ToPdf(options, out report)`, `ToPdfAsync`, `SaveAsPdf` and `SaveAsPdfAsync` return the report of both stages. `FailOnLoss` throws `OfficeConversionException` when either stage reports loss. Reverse conversion takes `PdfOdpExportOptions` (`PdfOptions`, `OpenDocumentOptions`, `FailOnLoss`) on a `PdfDocument` or `PdfDocumentReadResult`: `ToOdp(options, out report)`, `ToOdpAsync`, `SaveAsOdp` and `SaveAsOdpAsync`. The report from `ToOdp(options, out report)` and `ToOdpAsync` is a `PdfOdpConversionReport`; cast it to read the separate PDF and OpenDocument stage reports. `SaveAsOdp` and `SaveAsOdpAsync` return a plain `OfficeConversionReport` with the diagnostics of both stages.

## What it exports

- One page per visible slide, sized to the presentation's slide size and in slide order, or notes pages and handouts through `PowerPointPdfExportOptions.PageLayout`; plain speaker notes come across from ODP.
- Text boxes and runs, hyperlinks, images and crops, tables and merges, basic shapes and solid backgrounds, following the [ODP-to-PowerPoint mapping](../OfficeIMO.PowerPoint.OpenDocument/README.md).
- Placeholder roles and inherited solid master backgrounds.

PDF to ODP defaults to reconstructing supported editable content through the shared PowerPoint projection. An already reduced `PdfDocumentReadResult` resolves the default `Auto` profile to detected tables only. Use `PdfPowerPointExportOptions.CreateVisualPages()` for one rendered page image per slide, `CreateHybrid()` for visual pages with editable table overlays, or `CreateEditableTables()` for detected tables only.

## Current limits

- Master and layout drawing content, placeholder geometry, theme and effect appearance, charts, SmartArt, media and animations are omitted or reported by the presentation stage.
- Slide composition follows the PowerPoint PDF engine, not LibreOffice Impress or PowerPoint.
- PDF to ODP rebuilds supported text, tables, basic shapes and images; complex graphics, clipping and reading order remain approximations.

## Related packages

- [OfficeIMO.OpenDocument](../OfficeIMO.OpenDocument/README.md) reads, edits and writes ODP.
- [OfficeIMO.PowerPoint.OpenDocument](../OfficeIMO.PowerPoint.OpenDocument/README.md) converts between ODP and PPTX.
- [OfficeIMO.OpenDocument.Odt.Pdf](../OfficeIMO.OpenDocument.Odt.Pdf/README.md) and [OfficeIMO.OpenDocument.Ods.Pdf](../OfficeIMO.OpenDocument.Ods.Pdf/README.md) cover ODT and ODS.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 2 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.OpenDocument.Odp.Pdf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
