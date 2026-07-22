---
title: "Choose a PSWriteOffice Workflow"
description: "Decide between authoring, editing, inspection, conversion, and normalized Reader extraction."
layout: docs
---

PSWriteOffice exposes several paths that can reach the same file format. Choosing the path by outcome keeps scripts smaller and makes failure handling clearer.

## Create a new artifact

Use the document DSL when the script owns the output from the beginning. `New-OfficeWord`, `New-OfficeExcel`, `New-OfficePowerPoint`, `New-OfficePdf`, and `New-OfficeVisio` accept script blocks that describe the document. The outer command owns creation and final save; nested commands add content to the active document context.

This path works well for scheduled reports, build artifacts, exports, invoices, runbooks, and operational dashboards.

## Load and change an existing file

Use a `Get-*` or import command to open the document, target specific content with `Find-*` or `Get-*`, apply `Set-*`, `Update-*`, `Add-*`, or `Remove-*` operations, and finish with the matching `Save-*` command.

Keep mutation explicit. A useful script shape is:

1. inspect the input and collect diagnostics;
2. decide whether the requested change is safe;
3. apply a bounded set of changes;
4. save to a new path during validation;
5. compare or reopen the result.

## Inspect without changing

Use inspection and preflight commands when the job needs evidence rather than a rewritten file. Examples include `Get-OfficeWordStatistics`, `Get-OfficeExcelPreflight`, `Get-OfficePowerPointInspection`, `Get-OfficePdfDiagnostic`, `Get-OfficeVisioInfo`, and the Reader result commands.

Inspection-first workflows are appropriate for inventory, compliance gates, support diagnostics, and deciding which conversion route to use.

## Normalize many formats

Use `New-OfficeDocumentReader` and the `Get-OfficeDocument*` commands when downstream code should not branch on every source format. Reader exposes normalized documents, chunks, hierarchy, structured data, tables, visuals, assets, detections, and ingest results through one family of commands.

Choose the format-specific family instead when the script must edit native features that the normalized model intentionally does not preserve.

## Convert and review

Use focused converters when the destination format is the outcome. HTML review commands are useful for browser-based review; Markdown is useful for source control and text pipelines; PDF is useful for fixed-layout delivery. Conversion can be loss-aware rather than lossless, so inspect diagnostics for complex documents and retain the source artifact.

## Continue

- [Command families](/docs/pswriteoffice/command-families/)
- [Automation patterns](/docs/pswriteoffice/automation-patterns/)
- [Troubleshooting and diagnostics](/docs/pswriteoffice/troubleshooting/)
