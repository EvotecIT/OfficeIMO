---
title: "Compose a PDF report"
description: "Use PSWriteOffice to compose a PDF with headings, tables, metadata, and follow-up operations."
layout: docs
---

This pattern is useful when a script needs to produce a PDF artifact and then inspect or combine it with other PDF files.

It is adapted from `Examples/Pdf/Example-PdfReportDsl.ps1` and `Examples/Pdf/Example-PdfOperations.ps1`.

## Example

```powershell
Import-Module PSWriteOffice

$outputDirectory = Join-Path $PSScriptRoot 'Output'
$null = New-Item -ItemType Directory -Path $outputDirectory -Force
$coverPath = Join-Path $outputDirectory 'Cover.pdf'
$statusPath = Join-Path $outputDirectory 'Status.pdf'
$finalPath = Join-Path $outputDirectory 'Combined.pdf'

$rows = @(
    [PSCustomObject]@{ Area = 'Word'; Status = 'Ready'; Owner = 'Docs' }
    [PSCustomObject]@{ Area = 'PDF'; Status = 'Review'; Owner = 'Adapters' }
    [PSCustomObject]@{ Area = 'Reader'; Status = 'Ready'; Owner = 'Core' }
)

New-OfficePdf -Path $coverPath {
    PdfHeading -Text 'Documentation Status' -Level 1
    PdfParagraph -Text 'Prepared with PSWriteOffice.'
    PdfMetadata -Title 'Documentation Status' -Author 'PSWriteOffice'
}

New-OfficePdf -Path $statusPath {
    PdfHeading -Text 'Status Detail' -Level 1
    PdfParagraph -Text 'Generated with PSWriteOffice through OfficeIMO.Pdf.'
    PdfTable -InputObject $rows -Property Area,Status,Owner -Header 'Area','Status','Owner'
    PdfBookmark -Name 'Status table'
    PdfMetadata -Title 'Documentation Status' -Author 'PSWriteOffice'
}

Join-OfficePdf -Path $coverPath, $statusPath -OutputPath $finalPath
Get-OfficePdfText -Path $finalPath
```

## What this demonstrates

- composing a PDF through the PowerShell DSL
- adding structured data and metadata
- using the same module for follow-up PDF operations

## Source

- [Example-PdfReportDsl.ps1](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Pdf/Example-PdfReportDsl.ps1)
- [Example-PdfOperations.ps1](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Pdf/Example-PdfOperations.ps1)
