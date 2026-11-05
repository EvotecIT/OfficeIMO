---
title: "Create a PowerPoint deck"
description: "Use PSWriteOffice to generate a PowerPoint deck with slides, sections, notes, and transitions."
layout: docs
---

This pattern is useful when a repeatable status script should produce a deck that can still be edited in PowerPoint.

It is adapted from `Examples/PowerPoint/Example-PowerPointTransitionsAndSizing.ps1` and `Examples/PowerPoint/Example-PowerPointSectionsAndImport.ps1`.

## Example

```powershell
Import-Module PSWriteOffice

$outputDirectory = (New-Item -ItemType Directory -Path (Join-Path $PSScriptRoot 'Output') -Force).FullName
$outputPath = Join-Path $outputDirectory 'ServiceBrief.pptx'
$deck = New-OfficePowerPoint -Path $outputPath
try {
    $slide = Add-OfficePowerPointSlide -Document $deck -Title 'Service Brief' -PassThru
    Add-OfficePowerPointTextBox -Slide $slide -Text 'Generated with PSWriteOffice' -X 80pt -Y 145pt -Width 360pt -Height 50pt
    Add-OfficePowerPointList -Slide $slide -Items 'Current state', 'Risks', 'Next steps' -X 80pt -Y 220pt -Width 520pt -Height 180pt
    Set-OfficePowerPointNotes -Slide $slide -Text 'Use this slide as the spoken summary.'
    Add-OfficePowerPointSection -Document $deck -Name 'Opening' -Slide 1
    Set-OfficePowerPointSlide -Slide $slide -Transition Fade
    Set-OfficePowerPointSlideSize -Document $deck -Preset Screen16x9
    Save-OfficePowerPoint -Document $deck
} finally {
    Close-OfficePowerPoint -Document $deck
}
```

## What this demonstrates

- creating an editable `.pptx` from PowerShell
- adding slide content, speaker notes, and sections
- applying deck-level sizing and slide transitions

## Source

- [Example-PowerPointTransitionsAndSizing.ps1](https://github.com/EvotecIT/OfficeIMO/blob/master/PowerShell/PSWriteOffice/Examples/PowerPoint/Example-PowerPointTransitionsAndSizing.ps1)
- [Example-PowerPointSectionsAndImport.ps1](https://github.com/EvotecIT/OfficeIMO/blob/master/PowerShell/PSWriteOffice/Examples/PowerPoint/Example-PowerPointSectionsAndImport.ps1)
