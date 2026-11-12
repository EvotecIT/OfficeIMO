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

New-OfficePowerPoint -Path $sourcePath {
    PowerPointSlide -Title 'Reusable Architecture' {
        PowerPointTextBox -Text 'Shared platform diagram' -X 80pt -Y 150pt -Width 500pt -Height 80pt
    }
}

New-OfficePowerPoint -Path $targetPath {
    PowerPointSlide -Title 'Customer Briefing' {
        PowerPointTextBox -Text 'Prepared for review' -X 80pt -Y 150pt -Width 500pt -Height 80pt
    }
    Import-OfficePowerPointSlide -Path $sourcePath -SourceSlide 1 -Position 2
    Copy-OfficePowerPointSlide -Slide 1 -Position 3
    Add-OfficePowerPointSection -Name 'Shared material' -Slide 2
}
```

Inspect imported theme and layout behavior when the source and target use different design systems. Reopen the saved deck and review slide summaries, placeholders, and notes.

The [reuse-slides recipe](https://github.com/EvotecIT/OfficeIMO/blob/a0abd1c0e95b75c0c91ad57b607038477a98462c/PowerShell/PSWriteOffice/Examples/PowerPoint/Recipe-PowerPoint-ReuseSlides.ps1) imports one slide, copies another, creates a section, and saves the combined deck.
