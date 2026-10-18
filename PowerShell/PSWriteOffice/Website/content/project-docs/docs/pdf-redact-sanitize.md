---
title: "Redact, Sanitize, and Optimize PDF Files"
description: "Remove sensitive content by detected coordinates, sanitize interactive structures, and optimize a verified delivery copy."
layout: docs
---

Redaction, sanitization, and optimization are separate operations. Redaction removes selected content. Sanitization removes or normalizes risky document structures according to policy. Optimization changes representation to reduce or simplify the file.

## Redact detected text

Use `Get-OfficePdfRedactionPlan -Text` when you want to preview the matching areas, or pass the same literal text directly to `ConvertTo-OfficePdfRedacted`:

```powershell
Get-OfficePdfRedactionPlan -Path '.\Incident.pdf' -Text 'Secret account'

ConvertTo-OfficePdfRedacted `
    -Path '.\Incident.pdf' `
    -OutputPath '.\Incident-Redacted.pdf' `
    -Text 'Secret account'
```

The [redact-detected-text recipe](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Pdf/Recipe-Pdf-RedactDetectedText.ps1) shows the same operation with a self-contained source PDF.

## Prepare a delivery copy

`ConvertTo-OfficePdfSanitized` writes a new sanitized file. `ConvertTo-OfficePdfOptimized -PassThruReport` reports the actions and size outcome. Neither should overwrite the only source copy in an auditable workflow.

The [sanitize-and-optimize recipe](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Pdf/Recipe-Pdf-SanitizeAndOptimize.ps1) keeps the source, sanitized copy, and optimized delivery file separate.

Visual black boxes are not automatically secure redactions. Use the redaction command and verify extracted content rather than relying on a stamp or shape to hide text.
