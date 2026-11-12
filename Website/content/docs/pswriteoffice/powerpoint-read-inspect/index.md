---
title: "Read and Inspect PowerPoint Decks"
description: "Open existing presentations, enumerate slides and shapes, inspect notes and layouts, and summarize a deck before changing it."
layout: docs
---

Use the PowerPoint read surface to inventory a deck before modification, build a presentation catalog, extract speaker notes, or verify that a generated deck has the expected structure.

## Inspect slide by slide

```powershell
$presentation = Open-OfficePowerPoint -Path '.\Briefing.pptx'
try {
    Get-OfficePowerPointSlide -Document $presentation -Summary
    foreach ($slide in (Get-OfficePowerPointSlide -Document $presentation)) {
        Get-OfficePowerPointShape -Slide $slide
        Get-OfficePowerPointNotes -Slide $slide
    }
} finally {
    Close-OfficePowerPoint -Document $presentation
}
```

The inspection family covers slides, shapes, placeholders, notes, themes, layouts, sections, transitions, charts, tables, and presentation metadata. Start with a summary, then query the structures relevant to the job.

The [inspect-deck recipe](https://github.com/EvotecIT/OfficeIMO/blob/b4a2a879a4dad7c2a487e73d7d1cbe149fa32e77/PowerShell/PSWriteOffice/Examples/PowerPoint/Recipe-PowerPoint-InspectDeck.ps1) creates a representative deck and reports title, shape, and note counts per slide.
