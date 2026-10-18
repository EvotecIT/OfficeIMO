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
        Add-OfficeExcelTable -Data $rows -TableName 'Services' -TableStyle 'TableStyleMedium4'
        Set-OfficeExcelColumn -Column 1, 2, 3 -AutoFit
    }
}

$deck = New-OfficePowerPoint -Path $deckPath -NoSave
$slide = Add-OfficePowerPointSlide -Presentation $deck -Layout 1 -PassThru
Set-OfficePowerPointSlideTitle -Slide $slide -Title 'Service Review'
Add-OfficePowerPointTextBox -Slide $slide -Text 'Review the service status before the weekly meeting.' -X 80 -Y 140 -Width 560 -Height 80
$deck | Save-OfficePowerPoint
$deck | Close-OfficePowerPoint

ConvertTo-OfficeExcelHtml -Path $workbookPath -OutputPath (Join-Path $outputDirectory 'ServiceReview.workbook.html') -Title 'Workbook Review'
ConvertTo-OfficePowerPointHtml -Path $deckPath -Profile VisualReview -OutputPath (Join-Path $outputDirectory 'ServiceReview.deck.visual.html') -Title 'Deck Review'
```

## What this demonstrates

- creating real Excel and PowerPoint files first
- exporting workbook content as HTML through `ConvertTo-OfficeExcelHtml`
- exporting deck review output through `ConvertTo-OfficePowerPointHtml`
- using the visual profile when layout review matters more than raw text

## Source

- [Example-ExcelHtmlReview.ps1](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Excel/Example-ExcelHtmlReview.ps1)
- [Example-PowerPointHtmlReview.ps1](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/PowerPoint/Example-PowerPointHtmlReview.ps1)
