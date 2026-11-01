---
title: "Export and Review PowerPoint Presentations"
description: "Render slides for lightweight review, inspect speaker notes and metadata, and prepare presentations for delivery."
layout: docs
---

Recipients do not always need an editable PPTX. Export selected slides or the full deck as images when the review channel is chat, email, a ticket, or a web page. Keep the PPTX when the recipient must present or continue editing.

## Review surfaces

`Export-OfficeDocumentImage -Path ./review.pptx -OutputDirectory ./slides` creates visual slide output. Add `-Page 1-3` for a selection, or use `ConvertTo-OfficeDocumentImage` for image results without files. HTML conversion provides another browser-friendly review surface. `Get-OfficePowerPointNotes` separates speaker guidance from visible slide content, while slide summaries and shape inspection support structural checks.

Before delivery, verify:

1. slide count and order;
2. title and required sections;
3. speaker notes that must not contain stale or confidential text;
4. imported slide themes and layouts;
5. representative rendered slides at the intended aspect ratio.

The [quarterly-review recipe](https://github.com/EvotecIT/OfficeIMO/blob/894f02ac68ee967efc3105c566a1d67ba0dbe288/PowerShell/PSWriteOffice/Examples/PowerPoint/Recipe-PowerPoint-QuarterlyReview.ps1) includes charts, a table, and speaker notes suitable for this delivery check.
