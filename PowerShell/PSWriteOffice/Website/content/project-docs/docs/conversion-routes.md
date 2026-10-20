---
title: "Choose a conversion route"
description: "Use PSWriteOffice to inspect supported format pairs, conversion fidelity, owning APIs, known limitations, and batch PDF support in your host."
layout: docs
---

Use `Get-OfficeConversionRoute` to inspect the conversion contracts shipped with the installed OfficeIMO engine. A route describes a source and destination pair, its owning package and API, expected fidelity, support evidence, text formatting, and known limitations.

```powershell
Get-OfficeConversionRoute -SourceExtension docx -TargetExtension pdf |
    Format-List Id, PackageId, Api, Fidelity, SupportLevel,
        SupportEvidence, TextFormattingContract, KnownLimitations,
        PdfBatchSupported, PdfBatchCommand
```

Extension filters accept `docx` or `.docx` and compare without regard to case. Combine them with `-Id` to select an exact route. Unknown identifiers return no rows.

## Choose the PowerShell operation

`PdfBatchSupported` reports whether a route belongs to `Export-OfficeDocumentPdf` batch export in the current host. Batch export requires PowerShell 7.4 or newer. Windows PowerShell 5.1 can inspect the catalog and use supported single-file commands.

```powershell
Get-OfficeConversionRoute -PdfBatch |
    Select-Object Id, SourceExtensions, PackageId, Fidelity, KnownLimitations
```

A false value can mean that the conversion uses a format-specific cmdlet or a .NET API, or that the current host does not expose batch export. A reverse route such as `pdf-docx` does not imply PDF export support. Check the owning `Api` and the command reference for the operation you need.

For example, inspect Word-to-PDF limits, then export one document:

```powershell
$route = Get-OfficeConversionRoute -Id docx-pdf
$route.KnownLimitations
Export-OfficeDocumentPdf -InputPath ./Review.docx -Path ./Review.pdf
```

## Run a general conversion job

`Get-OfficeConversionRoute -Executable` filters to routes registered by the shared runner in this host. `ConvertTo-OfficeDocument` uses that runner for a single destination or a selected-file/directory batch. These jobs require PowerShell 7.4 or newer and include registered first-party iWork routes.

```powershell
Get-OfficeConversionRoute -Executable -SourceExtension pdf |
    Select-Object Id, TargetExtension, Fidelity, KnownLimitations
ConvertTo-OfficeDocument -InputPath ./Review.pdf -OutputPath ./Review.docx
Get-ChildItem ./reviews -Filter *.pdf |
    ConvertTo-OfficeDocument -OutputDirectory ./editable -TargetExtension docx -ItemResults
```

Existing destinations are preserved by default. Set `-ConflictPolicy Rename` or `Replace` explicitly when appropriate. Use `New-OfficeConversionOptions` for renderer settings, `-RequireNoLoss` to reject reported reductions, and `-CheckpointDirectory` with directory jobs that need resumable state. Route registration does not guarantee that an individual file, optional provider or credential will satisfy the operation.

## Check the result

The catalog describes the installed engine's contracts. It does not probe optional providers, fonts, credentials, encryption passwords, or an individual input document. Test representative files and inspect the generated output before relying on a conversion workflow. Review conversion warnings when fidelity matters.

For extraction, use `Get-OfficeDocumentCapability` and the [Reader guide](/docs/pswriteoffice/reader/). Reader support, conversion support, editable document authoring, and batch export are separate operations.

See [PDF automation](/docs/pswriteoffice/pdf/) for export options and [automation patterns](/docs/pswriteoffice/automation-patterns/) for repeatable document jobs.
