---
title: "Create a Word report"
description: "Use PSWriteOffice to generate a Word document with text, a list, and a table."
layout: docs
---

This pattern is useful when an operational script needs to leave behind a readable Word document.

It is adapted from `Examples/Word/Example-WordBasic.ps1`.

## Example

```powershell
Import-Module PSWriteOffice

$outputPath = Join-Path $PSScriptRoot 'Output\RevenueSnapshot.docx'
$null = New-Item -ItemType Directory -Path (Split-Path $outputPath) -Force
$data = @(
    [PSCustomObject]@{ Region = 'North America'; Revenue = 125000; YoY = '12%' }
    [PSCustomObject]@{ Region = 'EMEA'; Revenue = 98000; YoY = '22%' }
    [PSCustomObject]@{ Region = 'APAC'; Revenue = 143000; YoY = '18%' }
)

New-OfficeWord -Path $outputPath {
    Add-OfficeWordSection {
        Add-OfficeWordParagraph -Text 'Executive Summary' -Style Heading1
        Add-OfficeWordParagraph -Text 'Revenue accelerated in all regions.'
        Add-OfficeWordList -Style Numbered {
            Add-OfficeWordListItem -Text 'North America +12% YoY'
            Add-OfficeWordListItem -Text 'EMEA +22% YoY'
            Add-OfficeWordListItem -Text 'APAC +18% YoY'
        }
        Add-OfficeWordTable -InputObject $data -Style 'GridTable1LightAccent2'
    }
}
```

## What this demonstrates

- creating a Word document from a PowerShell script
- mixing paragraphs, lists, and structured data
- writing output to a predictable script-local folder

## Source

- [Example-WordBasic.ps1](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Word/Example-WordBasic.ps1)
