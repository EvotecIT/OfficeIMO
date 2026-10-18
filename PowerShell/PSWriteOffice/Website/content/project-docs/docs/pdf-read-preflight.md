---
title: "Read, Extract, and Preflight PDF Files"
description: "Inspect PDF metadata, pages, text, images, fonts, forms, annotations, signatures, compliance, and rewrite readiness before changing a file."
layout: docs
---

PDF transformation should begin with evidence. Read document information, extract only the content the workflow needs, and preflight the source before deciding whether to merge, redact, flatten, optimize, or rewrite it.

## Inspect the source

```powershell
$info = Get-OfficePdfInfo -Path '.\Input.pdf'
$preflight = Get-OfficePdfPreflight -Path '.\Input.pdf'
$pages = Get-OfficePdfText -Path '.\Input.pdf' -ByPage
```

Use `-AsTextBlock` when coordinates are needed for redaction or layout work. The read family also exposes images, fonts, attachments, form fields, annotations, signatures, interactions, optimization information, compliance, diagnostics, and rewrite-preservation evidence.

The [inspect-and-preflight recipe](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Pdf/Recipe-Pdf-InspectAndPreflight.ps1) creates a two-page PDF, reads its metadata and text page by page, and runs the PDF preflight command.

Encrypted files require the correct password. `-IgnorePermissionRestrictions` is an explicit post-authentication policy choice; it does not discover or bypass a missing password.
