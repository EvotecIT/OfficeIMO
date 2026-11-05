---
title: "Choose a conversion route"
description: "Use PSWriteOffice to inspect supported format pairs, conversion fidelity, owning APIs, known limitations, and batch PDF support in your host."
layout: docs
---

Use `Get-OfficeDocumentCapability -Conversion` to inspect the conversion contracts shipped with the installed OfficeIMO engine. A route describes a source and destination pair, its owning package and API, expected fidelity, support evidence, text formatting, and known limitations.

```powershell
Get-OfficeDocumentCapability -Conversion -From docx -To pdf |
    Format-List Id, PackageId, Api, Fidelity, SupportLevel,
        SupportEvidence, TextFormattingContract, KnownLimitations
```

Extension filters accept `docx` or `.docx` and compare without regard to case. Combine them with `-Id` to select an exact route. Unknown identifiers return no rows.

## Choose the PowerShell operation

Use `-Executable` to select routes registered by the conversion runner in the current host. Windows PowerShell 5.1 can inspect the catalog; general conversion jobs require PowerShell 7.4 or newer.

```powershell
Get-OfficeDocumentCapability -Conversion -Executable -To Pdf |
    Select-Object Id, SourceExtensions, PackageId, Fidelity, KnownLimitations

$route = Get-OfficeDocumentCapability -Conversion -Id docx-pdf
$route.KnownLimitations
$report = Convert-OfficeDocument -Path ./Review.docx -To Pdf -OutputPath ./Review.pdf
$report.Diagnostics
```

Catalog support and current-host execution are separate. A listed route does not establish that an optional provider, password, or individual input meets its requirements.

## Run a general conversion job

`Get-OfficeDocumentCapability -Conversion -Executable` filters to routes registered by the shared runner in this host. `Convert-OfficeDocument` uses that runner for one destination or a selected-file/directory batch, including registered first-party iWork routes.

```powershell
Get-OfficeDocumentCapability -Conversion -Executable -From pdf |
    Select-Object Id, TargetExtension, Fidelity, KnownLimitations
Convert-OfficeDocument -Path ./Review.pdf -To Word -OutputPath ./Review.docx
Get-ChildItem ./reviews -Filter *.pdf |
    Convert-OfficeDocument -OutputDirectory ./editable -To Word
```

Each completed file emits its report. Existing output is replaced by default; use `-NoClobber` to preserve destinations or `-AutoRename` to select another filename. Pass native format options through `-Options`, use `-FailOnLoss` to reject reported reductions, and combine `-CheckpointPath` with `-Resume` for resumable directory jobs.

## Check the result

The catalog describes the installed engine's contracts. It does not probe optional providers, fonts, credentials, encryption passwords, or an individual input document. Test representative files and inspect the generated output before relying on a conversion workflow. Review conversion warnings when fidelity matters.

For extraction, use `Get-OfficeDocumentCapability` and the [Reader guide](/docs/pswriteoffice/reader/). Reader support, conversion support, editable document authoring, and batch export are separate operations.

See [PDF automation](/docs/pswriteoffice/pdf/) for export options and [automation patterns](/docs/pswriteoffice/automation-patterns/) for repeatable document jobs.
