---
title: "Review Office documents as HTML"
description: "Export Excel workbooks and PowerPoint decks as semantic or visual HTML review artifacts."
layout: docs
---

Use HTML review output when a workbook or deck needs lightweight inspection without opening the Office application. The semantic profile favors structured content and inventories. The visual profile favors a review snapshot.

## Example

```powershell
Import-Module PSWriteOffice

$outputDirectory = Join-Path $PSScriptRoot 'Output'
$null = New-Item -ItemType Directory -Path $outputDirectory -Force
$workbookPath = Join-Path $outputDirectory 'ServiceReview.xlsx'
$deckPath = Join-Path $outputDirectory 'ServiceReview.pptx'

$rows = @(
    [PSCustomObject]@{ Service = 'Identity'; Status = 'Healthy'; Owner = 'Platform' }
    [PSCustomObject]@{ Service = 'Messaging'; Status = 'Watch'; Owner = 'Collaboration' }
)

New-OfficeExcel -Path $workbookPath {
    Add-OfficeExcelSheet -Name 'Services' -Content {
        Set-OfficeExcelRow -Row 1 -Values 'Service', 'Status', 'Owner'
        Add-OfficeExcelTable -InputObject $rows -Name 'Services' -TableStyle 'TableStyleMedium4'
        foreach ($column in 1..3) {
            Set-OfficeExcelColumn -Column $column -AutoFit
        }
    }
}

$deck = New-OfficePowerPoint -Path $deckPath
$slide = Add-OfficePowerPointSlide -Document $deck -TemplateLayout 1 -PassThru
Set-OfficePowerPointSlide -Slide $slide -Title 'Service Review'
Add-OfficePowerPointTextBox -Slide $slide -Text 'Review the service status before the weekly meeting.' -X 80 -Y 140 -Width 560 -Height 80
$deck | Save-OfficePowerPoint
$deck | Close-OfficePowerPoint

Convert-OfficeDocument -Path $workbookPath -To Html -OutputPath (Join-Path $outputDirectory 'ServiceReview.workbook.html') -Options (New-OfficeExcelHtmlOptions -Title 'Workbook Review')
Convert-OfficeDocument -Path $deckPath -To Html -OutputPath (Join-Path $outputDirectory 'ServiceReview.deck.visual.html') -Options (New-OfficePowerPointHtmlOptions -Profile VisualReview -Title 'Deck Review')
```

## What this demonstrates

- creating real Excel and PowerPoint files first
- exporting workbook content as HTML through `Convert-OfficeDocument -To Html` with `New-OfficeExcelHtmlOptions`
- exporting deck review output through `Convert-OfficeDocument -To Html` with `New-OfficePowerPointHtmlOptions`
- using the visual profile when layout review matters more than raw text

## Source

- [Example-ExcelHtmlReview.ps1](https://github.com/EvotecIT/OfficeIMO/blob/master/PowerShell/PSWriteOffice/Examples/Excel/Example-ExcelHtmlReview.ps1)
- [Example-PowerPointHtmlReview.ps1](https://github.com/EvotecIT/OfficeIMO/blob/master/PowerShell/PSWriteOffice/Examples/PowerPoint/Example-PowerPointHtmlReview.ps1)
