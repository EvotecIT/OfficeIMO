---
title: "Reuse and Combine PowerPoint Slides"
description: "Copy slides inside a deck, import slides from another presentation, and organize combined material into sections."
layout: docs
---

Slide reuse is the practical PowerPoint equivalent of document merge. It supports reusable architecture slides, standard legal pages, customer-specific inserts, and decks assembled from team-owned sources.

## Copy or import

- `Copy-OfficePowerPointSlide` duplicates a slide inside the current presentation.
- `Import-OfficePowerPointSlide` imports a selected slide from another presentation path or loaded presentation.
- `Add-OfficePowerPointSection`, `Rename-OfficePowerPointSection`, and related commands organize the combined deck.

```powershell
$sourcePath = '.\Reusable-Slides.pptx'
$targetPath = '.\Customer-Briefing.pptx'

PptNew -Path $sourcePath {
    PptSlide {
        PptTitle -Title 'Reusable Architecture'
        PptTextBox -Text 'Shared platform diagram' `
            -X 80 -Y 150 -Width 500 -Height 80
    }
}

PptNew -Path $targetPath {
    PptSlide {
        PptTitle -Title 'Customer Briefing'
        PptTextBox -Text 'Prepared for review' `
            -X 80 -Y 150 -Width 500 -Height 80
    }
}

$presentation = Get-OfficePowerPoint -Path $targetPath

Import-OfficePowerPointSlide -Presentation $presentation `
    -SourcePath $sourcePath -SourceIndex 0 -InsertAt 1

Copy-OfficePowerPointSlide -Presentation $presentation `
    -Index 0 -InsertAt 2

Add-OfficePowerPointSection -Presentation $presentation `
    -Name 'Shared material' -StartSlideIndex 1

Close-OfficePowerPoint -Presentation $presentation -Save
```

Inspect imported theme and layout behavior when the source and target use different design systems. Reopen the saved deck and review slide summaries, placeholders, and notes.

The [reuse-slides recipe](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/PowerPoint/Recipe-PowerPoint-ReuseSlides.ps1) imports one slide, copies another, creates a section, and saves the combined deck.
