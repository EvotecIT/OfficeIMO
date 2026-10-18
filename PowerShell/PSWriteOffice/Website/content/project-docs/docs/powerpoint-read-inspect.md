---
title: "Read and Inspect PowerPoint Decks"
description: "Open existing presentations, enumerate slides and shapes, inspect notes and layouts, and summarize a deck before changing it."
layout: docs
---

Use the PowerPoint read surface to inventory a deck before modification, build a presentation catalog, extract speaker notes, or verify that a generated deck has the expected structure.

## Inspect slide by slide

```powershell
$presentation = Get-OfficePowerPoint -Path '.\Briefing.pptx'
for ($index = 0; $index -lt $presentation.Slides.Count; $index++) {
    $slide = Get-OfficePowerPointSlide -Presentation $presentation -Index $index
    Get-OfficePowerPointSlideSummary -Slide $slide
    Get-OfficePowerPointShape -Slide $slide
    Get-OfficePowerPointNotes -Slide $slide
}
Close-OfficePowerPoint -Presentation $presentation
```

The inspection family covers slides, shapes, placeholders, notes, themes, layouts, sections, transitions, charts, tables, and presentation metadata. Start with a summary, then query the structures relevant to the job.

The [inspect-deck recipe](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/PowerPoint/Recipe-PowerPoint-InspectDeck.ps1) creates a representative deck and reports title, shape, and note counts per slide.
