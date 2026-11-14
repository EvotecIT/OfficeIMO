---
title: "PowerShell"
description: "PSWriteOffice automates Word, Excel, PowerPoint, PDF, Visio, Project, Access, EPUB, YAML, TOML and CSV from PowerShell on Windows, Linux and macOS, without Office."
meta.seo_title: "PSWriteOffice: PowerShell module for Word, Excel, PowerPoint and PDF"
layout: application
meta.application: "powershell"
meta.software.name: "PSWriteOffice"
meta.software.application_category: "DeveloperApplication"
meta.software.operating_system: "Windows, Linux, macOS"
meta.software.download_url: "https://www.powershellgallery.com/packages/PSWriteOffice"
meta.software.price: 0
meta.software.price_currency: "USD"
---

## Quick start

Install the module, then turn the same rows into a Word report, an Excel table, a PowerPoint deck, and a PDF:

```powershell
Install-Module PSWriteOffice

$regions = @(
    [pscustomobject]@{ Region = 'North'; Revenue = 42000; Target = 40000 }
    [pscustomobject]@{ Region = 'South'; Revenue = 38000; Target = 35000 }
    [pscustomobject]@{ Region = 'East';  Revenue = 29000; Target = 30000 }
    [pscustomobject]@{ Region = 'West';  Revenue = 51000; Target = 45000 }
)

New-OfficeWord -Path .\Report.docx {
    Add-OfficeWordSection {
        Add-OfficeWordParagraph -Style Heading1 -Text 'Monthly report'
        Add-OfficeWordParagraph -Text "Generated on $(Get-Date -Format 'MMMM yyyy')"
        Add-OfficeWordTable -InputObject $regions -Style GridTable4Accent1
    }
}

New-OfficeExcel -Path .\Revenue.xlsx {
    Add-OfficeExcelSheet -Name Revenue {
        Add-OfficeExcelTable -InputObject $regions -Name Revenue
    }
}

New-OfficePowerPoint -Path .\Status.pptx {
    PowerPointSlide -Title 'Monthly status' {
        PowerPointList 'Three of four regions above target', 'East is 3% short'
    }
    PowerPointSlide -Title 'Revenue by region' {
        PowerPointTable -InputObject $regions
    }
}

Convert-OfficeDocument -Path .\Report.docx -OutputPath .\Report.pdf
```

The [DSL cookbook](/docs/pswriteoffice/dsl-cookbook/) builds on the same pattern for each format, and `Get-Help <command> -Full` shows every parameter.

## Where it runs

| Platform | PowerShell | Supported |
|----------|------------|-----------|
| Windows  | 5.1        | Yes       |
| Windows  | 7.4+       | Yes       |
| Linux    | 7.4+       | Yes       |
| macOS    | 7.4+       | Yes       |
