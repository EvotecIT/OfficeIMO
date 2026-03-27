---
title: PowerPoint Cmdlets
description: PSWriteOffice cmdlets and DSL aliases for building PowerPoint presentations from PowerShell.
order: 63
---

# PowerPoint Cmdlets

PSWriteOffice exposes `OfficeIMO.PowerPoint` through both direct cmdlets and lightweight DSL aliases. The examples below favor the direct cmdlets because they map cleanly to the generated help surface.

## Core workflow

1. Create a presentation with `New-OfficePowerPoint`.
2. Add slides, text boxes, bullets, images, tables, or charts.
3. Save the presentation as part of your script or pipeline output.

## Quick start

```powershell
$ppt = New-OfficePowerPoint -FilePath .\status_update.pptx
$slide = Add-OfficePowerPointSlide -Presentation $ppt

Add-OfficePowerPointTextBox -Slide $slide -Text 'Weekly Status' -X 80 -Y 60 -Width 520 -Height 50
Add-OfficePowerPointBullets -Slide $slide -Bullets @(
    'Shipped the March release',
    'Closed 14 customer issues',
    'Finished security review follow up'
) -X 80 -Y 140 -Width 520 -Height 220

$ppt | Save-OfficePowerPoint
```

## Cmdlet-style slide composition

```powershell
$ppt = New-OfficePowerPoint -FilePath .\deck.pptx

$slide = Add-OfficePowerPointSlide -Presentation $ppt
Add-OfficePowerPointTextBox -Slide $slide -Text 'Quarterly Overview' -X 80 -Y 60 -Width 500 -Height 50
Add-OfficePowerPointBullets -Slide $slide -Bullets @(
    'Revenue grew 18%',
    'New onboarding flow launched',
    'Support backlog dropped below target'
) -X 80 -Y 150 -Width 520 -Height 220

Add-OfficePowerPointImage -Slide $slide -Path .\logo.png -X 560 -Y 40 -Width 120 -Height 72
$ppt | Save-OfficePowerPoint
```

## Data-heavy slides

When your source is tabular data, use the table and chart cmdlets instead of manually building text shapes:

```powershell
$ppt = New-OfficePowerPoint -FilePath .\quarterly_review.pptx
$rows = @(
    [pscustomobject]@{ Product = 'Alpha'; Q1 = 12; Q2 = 15; Q3 = 18; Q4 = 20 }
    [pscustomobject]@{ Product = 'Beta';  Q1 =  9; Q2 = 11; Q3 = 13; Q4 = 14 }
    [pscustomobject]@{ Product = 'Gamma'; Q1 =  6; Q2 =  9; Q3 = 12; Q4 = 16 }
)

$slide = Add-OfficePowerPointSlide -Presentation $ppt
Add-OfficePowerPointTable -Slide $slide -Data $rows -X 60 -Y 140 -Width 420 -Height 200
$ppt | Save-OfficePowerPoint
```

## Useful commands

- `New-OfficePowerPoint` creates a presentation and can host the DSL block.
- `Add-OfficePowerPointSlide` adds new slides when you want explicit control.
- `Add-OfficePowerPointTextBox` and `Add-OfficePowerPointBullets` cover most narrative slide content.
- `Add-OfficePowerPointTable`, `Add-OfficePowerPointChart`, and `Add-OfficePowerPointImage` are the main building blocks for report decks.
- `Add-OfficePowerPointSection` helps organize larger presentations into named sections.
- `Save-OfficePowerPoint` persists the generated deck to disk.

## Related guides

- [PSWriteOffice overview](/docs/pswriteoffice/) -- Module-level installation and command map.
- [PowerPoint overview](/docs/powerpoint/) -- .NET package workflow and layout concepts.
- [PowerPoint product page](/products/powerpoint/) -- Install command and package positioning.
