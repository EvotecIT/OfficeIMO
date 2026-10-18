---
title: "Update Existing PowerPoint Presentations"
description: "Replace text and modify selected slides, shapes, tables, notes, themes, and layouts in an existing PPTX deck."
layout: docs
---

Targeted updates are useful for recurring decks where the design should remain stable while dates, labels, data, or speaker notes change.

## Replace repeated text

Open the presentation, run `Update-OfficePowerPointText`, and close it with `-Save`. Use `-IncludeNotes` when the same term must change in speaker notes; table text is included by default and can be controlled explicitly.

```powershell
$deck = Get-OfficePowerPoint -Path '.\FY24-Review.pptx'
Update-OfficePowerPointText -Presentation $deck `
    -OldValue FY24 -NewValue FY25 -IncludeNotes
Close-OfficePowerPoint -Presentation $deck -Save
```

Use targeted shape, table, notes, background, theme, layout, transition, and section commands when the change is structural. The [update-existing recipe](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/PowerPoint/Recipe-PowerPoint-UpdateExisting.ps1) demonstrates the complete lifecycle.
