---
title: "Update Existing PowerPoint Presentations"
description: "Replace text and modify selected slides, shapes, tables, notes, themes, and layouts in an existing PPTX deck."
layout: docs
---

Targeted updates are useful for recurring decks where the design should remain stable while dates, labels, data, or speaker notes change.

## Replace repeated text

Open the presentation in a content block to save the completed edit atomically. Use `-IncludeNotes` to update speaker notes and `-NoTables` to leave table cells unchanged. Supply `-Slide` with a native slide or a one-based slide number to limit the replacement.

```powershell
Open-OfficePowerPoint -Path '.\FY24-Review.pptx' {
    Update-OfficePowerPointText -OldValue FY24 -NewValue FY25 -IncludeNotes
}
```

Use targeted shape, table, notes, background, theme, layout, transition, and section commands when the change is structural. The [update-existing recipe](https://github.com/EvotecIT/OfficeIMO/blob/caffcd0d02eca2ae6917883df6c3aeb6e3a47664/PowerShell/PSWriteOffice/Examples/PowerPoint/Recipe-PowerPoint-UpdateExisting.ps1) demonstrates the complete lifecycle.
