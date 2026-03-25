---
title: PowerPoint Cmdlets
description: PSWriteOffice cmdlets and DSL aliases for building PowerPoint presentations from PowerShell.
order: 63
---

# PowerPoint Cmdlets

PSWriteOffice exposes `OfficeIMO.PowerPoint` through both direct cmdlets and lightweight DSL aliases. That gives you two ways to work: build a presentation imperatively with explicit objects, or create one inline with `Ppt*` helpers inside `New-OfficePowerPoint`.

## Core workflow

1. Create a presentation with `New-OfficePowerPoint`.
2. Add slides, titles, text boxes, bullets, images, tables, or charts.
3. Save the presentation as part of your script or pipeline output.

## DSL-style quick start

```powershell
New-OfficePowerPoint -Path .\status-update.pptx {
    PptSlide -Layout Title {
        PptTitle -Text "Weekly Status"
        PptSubtitle -Text "Engineering Team"
    }

    PptSlide -Layout TitleAndContent {
        PptTitle -Text "Completed This Week"
        PptContent {
            PptBullet -Text "Shipped the March release"
            PptBullet -Text "Closed 14 customer issues"
            PptBullet -Text "Finished security review follow-up"
        }
    }
}
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
```

## Data-heavy slides

When your source is tabular data, use the table and chart cmdlets instead of manually building text shapes:

```powershell
$rows = @(
    [pscustomobject]@{ Product = 'Alpha'; Q1 = 12; Q2 = 15; Q3 = 18; Q4 = 20 }
    [pscustomobject]@{ Product = 'Beta';  Q1 =  9; Q2 = 11; Q3 = 13; Q4 = 14 }
    [pscustomobject]@{ Product = 'Gamma'; Q1 =  6; Q2 =  9; Q3 = 12; Q4 = 16 }
)

$slide = Add-OfficePowerPointSlide -Presentation $ppt
Add-OfficePowerPointTable -Slide $slide -Data $rows -X 60 -Y 140 -Width 420 -Height 200
```

## Useful commands

- `New-OfficePowerPoint` creates a presentation and can host the DSL block.
- `Add-OfficePowerPointSlide` adds new slides when you want explicit control.
- `Add-OfficePowerPointTextBox` and `Add-OfficePowerPointBullets` cover most narrative slide content.
- `Add-OfficePowerPointTable`, `Add-OfficePowerPointChart`, and `Add-OfficePowerPointImage` are the main building blocks for report decks.
- `Add-OfficePowerPointSection` helps organize larger presentations into named sections.

## Related guides

- [PSWriteOffice overview](/docs/pswriteoffice/) -- Module-level installation and command map.
- [PowerPoint overview](/docs/powerpoint/) -- .NET package workflow and layout concepts.
- [PowerPoint product page](/products/powerpoint/) -- Install command and package positioning.
