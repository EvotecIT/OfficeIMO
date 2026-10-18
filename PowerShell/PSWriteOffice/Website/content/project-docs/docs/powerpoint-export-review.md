---
title: "Export and Review PowerPoint Presentations"
description: "Render slides for lightweight review, inspect speaker notes and metadata, and prepare presentations for delivery."
layout: docs
---

Recipients do not always need an editable PPTX. Export selected slides or the full deck as images when the review channel is chat, email, a ticket, or a web page. Keep the PPTX when the recipient must present or continue editing.

## Review surfaces

`Export-OfficePowerPointImage` creates visual slide output. HTML conversion provides another browser-friendly review surface. `Get-OfficePowerPointNotes` separates speaker guidance from visible slide content, while slide summaries and shape inspection support structural checks.

Before delivery, verify:

1. slide count and order;
2. title and required sections;
3. speaker notes that must not contain stale or confidential text;
4. imported slide themes and layouts;
5. representative rendered slides at the intended aspect ratio.

The [quarterly-review recipe](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/PowerPoint/Recipe-PowerPoint-QuarterlyReview.ps1) includes charts, a table, and speaker notes suitable for this delivery check.
