---
title: "Reliable Automation Patterns"
description: "Structure document jobs for repeatability, diagnostics, safe mutation, and CI artifacts. Includes examples and cmdlet links."
layout: docs
---

Document automation becomes easier to operate when scripts separate source data, document composition, validation, and delivery.

## Start with the smallest public surface

Use a pipeline for a direct export, a live document object for loop-driven composition, the DSL for a complete authored artifact, and targeted commands for an existing file. Escalating only when the job needs more structure keeps common scripts short without hiding advanced capabilities. See [pipeline, object, and DSL workflows](/docs/pswriteoffice/object-workflows/).

## Keep paths and data explicit

Resolve input and output roots once. Pass plain objects into document builders instead of reading global state from nested DSL blocks. Write validation output beside the artifact or into a CI report folder.

## Prefer canonical commands in reusable scripts

PSWriteOffice publishes aliases for a compact DSL, but canonical commands make long-lived scripts easier to search, document, and compare with generated help. Use aliases when they materially improve a local composition block; keep entry points and operational steps canonical.

## Inspect before mutation

For existing files, collect preflight, capability, signature, compliance, accessibility, or format-detection evidence before changing the document. Write to a new path during development and reopen the result through the same engine.

## Treat conversion as a result

Keep the source, destination, diagnostics, warnings, and any fidelity policy together. A successful file write does not by itself prove that complex layout or unsupported features were preserved.

## Make jobs idempotent

Create a fresh output folder or use stable file names with deliberate overwrite behavior. Avoid appending the same section, sheet, or slide on every rerun unless the script first detects whether it already exists.

For object workflows, leave out the constructor's content block, apply the complete change set to the live document, then use the matching `Save-*` and `Close-*` commands once. Saved DSL constructors are quiet by default; use `-PassThru` only when another command needs the saved file.

## Bound batch work

For Reader and conversion batches, limit concurrency based on document size and native-memory pressure. Capture failures per file so one malformed input does not erase evidence for the rest of the batch.

## Convert one file or a selected batch

`ConvertTo-OfficeDocument` uses the shared OfficeIMO runner. A destination extension selects the single-file target; a batch uses `-TargetExtension`. Existing output is preserved by default. Choose `Rename` or `Replace` explicitly when a job needs a different conflict policy. General conversion jobs require PowerShell 7.4 or newer.

```powershell
Import-Module PSWriteOffice
ConvertTo-OfficeDocument -InputPath ./report.pdf -OutputPath ./report.docx
Get-ChildItem ./reports -Filter *.pdf |
    ConvertTo-OfficeDocument -OutputDirectory ./word -TargetExtension docx -ItemResults
```

Use `-CheckpointDirectory` for resumable directory jobs and retain the per-file results. `Get-OfficeConversionRoute -Executable` lists routes registered by the current host; their loss policies still apply.

## Work with native document families

Project tasks, Access rows, EPUB chapters and invoice lines accept ordinary PowerShell parameters and object pipelines. The [native document example](https://github.com/EvotecIT/OfficeIMO/blob/2fa093a7dc76bfebc53781b5eccc54e715aa3db6/PowerShell/PSWriteOffice/Examples/Workflows/Example-NativeDocumentFamilies.ps1) creates and reads these artifacts without constructing .NET option objects.

- Use `New-OfficeProject`, task/resource/assignment commands and `Invoke-OfficeProjectSchedule` to prepare a schedule. XML preserves the supported scheduling model; native MPP saves assess generation-specific loss. Inspect `-ReportVariable` and opt into `-AllowLoss` only after reviewing the reported reductions.
- Use `New-OfficeAccess` and table/column/row commands for a native MDB or ACCDB table database. The file operations remain inert: they do not start ACE, run SQL, or execute database code.
- Use `New-OfficeEpub`, `Add-OfficeEpubChapter` and `Add-OfficeEpubResource` to compose a publication. Plain chapter text is escaped; XHTML and Markdown are explicit alternatives.
- Use `New-OfficeInvoiceParty`, `New-OfficeInvoice` and `Add-OfficeInvoiceLine` for XML and hybrid PDF output. Editing an imported invoice retains source evidence. `Set-OfficeInvoice -RecalculateTotals` explicitly clears captured amounts so the owner can recompute them; currency conversion is never inferred. File-based authority validation needs caller-supplied pinned rule archives and any required external validator.

## Find exact parameters

The [command reference](/api/powershell/) is the authoritative parameter contract. The [examples](https://github.com/EvotecIT/OfficeIMO/tree/2fa093a7dc76bfebc53781b5eccc54e715aa3db6/PowerShell/PSWriteOffice/Examples) demonstrate complete job shapes. Pin the module version in production automation and review release notes before changing that pin.
