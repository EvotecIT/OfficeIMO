# OfficeIMO.Workflows

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Workflows)](https://www.nuget.org/packages/OfficeIMO.Workflows)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Workflows?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Workflows)

Run document jobs end to end in .NET: convert, print, assemble, protect, sign, redact and publish, with typed requests, limits, cancellation, conflict policies, atomic local output and validation of what was written. Writes through a storage provider can end `Unconfirmed` and can't guarantee atomic replacement or rollback. `OfficeIMO.Workflows` composes OfficeIMO's own engines and adds no document engine of its own, so Studio, the `officeimo` CLI, PSWriteOffice and your own services share the same job contract.

## Install

```sh
dotnet add package OfficeIMO.Workflows
```

## Quick start

```csharp
using OfficeIMO.Workflows;

var runner = new OfficeWorkflowRunner();
OfficeWorkflowResult result = await runner.RunAsync(new OfficeWorkflowRequest {
    Operation = OfficeWorkflowOperation.Convert,
    ConversionRouteId = "docx-pdf",
    InputPath = "report.docx",
    OutputPath = "report.pdf",
    ConflictPolicy = OfficePublicationPolicy.Replace
});

if (!result.Succeeded) throw new InvalidOperationException(result.Summary);
```

Conversions run in process through the OfficeIMO format packages; no office suite or converter is launched.

## What it does

- **Convert** single files and batches, with saved, versioned job definitions that Studio, the CLI and PSWriteOffice run the same way, run reports and optional checkpoints to resume.
- **Print** any convertible document to installed or IPP printers on Windows, macOS and Linux, with sheet layout, preview and job tracking.
- **PDF jobs**: protected or unencrypted copies, PDF/A conversion, page extraction and splitting, certificate-signed copies, form value review and redaction with review.
- **Book publishing**: build EPUB and PDF editions from a book project, delivery bundles, offline KDP delivery and ONIX bibliographic export.
- **Evidence and dossiers** from email conversations, and invoice inspection, conversion and rendering.
- **Project reports** and table exchange for Microsoft Project files.
- **Scans and images**: review OCR before publication, save prepared scan copies, inspect concealed text in raster images and optimize embedded Word images.
- **Providers**: read and write through provider-backed inputs, outputs and folders, such as cloud drives, with membership checks before publication.
- **Provenance**: inspect and remove Content Credentials and other provenance carriers.
- **Apple iWork** routes through the opt-in [OfficeIMO.Workflows.IWork](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Workflows.IWork/README.md).

The [Workflows guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.workflows-guide.md) shows each of these with examples.

## Print a document

```csharp
using OfficeIMO.Workflows;

var printers = new PdfPrinterService();
OfficePrintResult printed = await printers.PrintAsync(new OfficePrintRequest {
    Layout = new PdfPrintPlanRequest { InputPath = "proposal.docx", PagesPerSheet = 2 },
    UsePrinterPaper = true,
    Delivery = new PdfPrintDeliveryOptions { PrinterName = "ipp://printer.local/ipp/print", Duplex = PdfPrintDuplex.LongEdge }
});
```

A PDF prints directly; other formats are converted in memory first. The [print section of the guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.workflows-guide.md#print-a-document) covers composing, previewing and tracking jobs.

## Limits

Each job validates its output before publishing it and reports what a conversion couldn't carry. Route coverage follows the [conversion route catalog](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/conversion-routes.md), and the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md) lists the workflow operations.

## Learn more

- [Workflows docs](https://officeimo.com/docs/workflows/) on officeimo.com and the [Workflows guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.workflows-guide.md).
- [OfficeIMO Studio](https://officeimo.com/studio/), the `officeimo` CLI ([OfficeIMO.Tool](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Tool/README.md)) and [PSWriteOffice](https://github.com/EvotecIT/OfficeIMO/blob/master/PowerShell/PSWriteOffice/README.MD) run these workflows.

## Targets and dependencies

- Targets: `netstandard2.0`, `net8.0` and `net10.0`, plus `net472` on Windows.
- Dependencies: the OfficeIMO format, conversion and PDF packages it orchestrates.
- License: MIT.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 1 | 0 | 0 | 0 | 0 | 0 |
| Read | 1 | 0 | 0 | 0 | 0 | 0 |
| Edit | 1 | 0 | 0 | 0 | 0 | 0 |
| Preserve | 1 | 0 | 0 | 0 | 0 | 0 |
| Validate | 0 | 1 | 0 | 0 | 0 | 0 |
| Convert | 0 | 25 | 0 | 0 | 0 | 0 |
| Export | 1 | 5 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Workflows` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
