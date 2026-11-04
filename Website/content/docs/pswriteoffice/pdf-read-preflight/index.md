---
title: "Read, Extract, and Preflight PDF Files"
description: "Inspect PDF metadata, pages, text, images, fonts, forms, annotations, signatures, compliance, and rewrite readiness before changing a file."
layout: docs
---

PDF transformation should begin with evidence. Read document information, extract only the content the workflow needs, and preflight the source before deciding whether to merge, redact, flatten, optimize, or rewrite it.

## Inspect the source

```powershell
$info = Get-OfficeDocumentInfo -Path '.\Input.pdf' -Detailed
$info.Details
$preflight = Get-OfficePdfDiagnostic -Path '.\Input.pdf' -Type Preflight
$pages = Read-OfficeDocument -Path '.\Input.pdf' -As Text -ByPage
```

Use `Get-OfficePdfDiagnostic -Type LayoutOverlay` when coordinates are needed for layout review and `Get-OfficePdfRedactionPlan -Text` to locate redaction targets. The read family also exposes images, fonts, attachments, form fields, annotations, signatures, interactions, optimization information, compliance, diagnostics, and rewrite-preservation evidence.

The [inspect-and-preflight recipe](https://github.com/EvotecIT/OfficeIMO/blob/3edfe3bbcf964fd078ee635a8d48562898ea0c87/PowerShell/PSWriteOffice/Examples/Pdf/Recipe-Pdf-InspectAndPreflight.ps1) creates a two-page PDF, reads its metadata and text page by page, and runs the PDF preflight command.

Encrypted files require the correct password. `-IgnorePermissionRestrictions` is an explicit post-authentication policy choice; it does not discover or bypass a missing password.
