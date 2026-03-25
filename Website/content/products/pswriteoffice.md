---
title: "PSWriteOffice"
description: "PowerShell module for creating Office documents with DSL aliases. Word, Excel, PowerPoint, and Markdown from the terminal."
layout: product
product_color: "#d97706"
install: "Install-Module PSWriteOffice"
nuget: ""
docs_url: "/docs/pswriteoffice/"
api_url: "/api/powershell/"
---

## Why PSWriteOffice?

PSWriteOffice brings the OfficeIMO document stack to PowerShell. With 175+ cmdlets and lightweight DSL aliases, you can create Word documents, Excel workbooks, PowerPoint decks, and Markdown output from scripts and automation jobs.

## Features

- **175+ cmdlets** -- comprehensive coverage of document creation, formatting, conversion, and export
- **DSL aliases** -- intuitive shorthand like `WordParagraph`, `ExcelSheet`, and `PptSlide` for concise scripts
- **Word, Excel, PowerPoint & Markdown** -- create and modify all major Office formats from one module
- **Cross-platform** -- runs on Windows, Linux, and macOS wherever PowerShell is available
- **PowerShell 5.1 and 7+** -- compatible with Windows PowerShell and modern PowerShell Core

## Quick start

```powershell
# Install the module
Install-Module PSWriteOffice -Force

# Create a Word document
New-OfficeWord -Path ".\Report.docx" {
    Add-OfficeWordSection {
        Add-OfficeWordParagraph -Style Heading1 -Text "Monthly Report"
        Add-OfficeWordParagraph -Text "Generated on $(Get-Date -Format 'MMMM yyyy')"

        $data = @(
            [pscustomobject]@{ Region = "North"; Revenue = 42000; Target = 40000 }
            [pscustomobject]@{ Region = "South"; Revenue = 38000; Target = 35000 }
            [pscustomobject]@{ Region = "East";  Revenue = 29000; Target = 30000 }
            [pscustomobject]@{ Region = "West";  Revenue = 51000; Target = 45000 }
        )

        Add-OfficeWordTable -InputObject $data -Style GridTable4Accent1
        Add-OfficeWordParagraph {
            Add-OfficeWordText -Text "All regions met or exceeded targets except East."
        }
    }
}

# Create an Excel workbook
New-OfficeExcel -Path ".\Metrics.xlsx" {
    Add-OfficeExcelSheet -Name "Summary" {
        Set-OfficeExcelCell -Address "A1" -Value "Metric"
        Set-OfficeExcelCell -Address "B1" -Value "Value"
        Set-OfficeExcelCell -Address "C1" -Value "Status"
        Set-OfficeExcelCell -Address "A2" -Value "Uptime"
        Set-OfficeExcelCell -Address "B2" -Value "99.97%"
        Set-OfficeExcelCell -Address "C2" -Value "OK"
        Set-OfficeExcelCell -Address "A3" -Value "Response Time"
        Set-OfficeExcelCell -Address "B3" -Value "42ms"
        Set-OfficeExcelCell -Address "C3" -Value "OK"
    }

    Add-OfficeExcelSheet -Name "Details" {
        Set-OfficeExcelCell -Address "A1" -Value "Timestamp"
        Set-OfficeExcelCell -Address "B1" -Value "Endpoint"
        Set-OfficeExcelCell -Address "C1" -Value "Latency"
        $rowIndex = 2
        1..10 | ForEach-Object {
            Set-OfficeExcelCell -Row $rowIndex -Column 1 -Value ((Get-Date).AddMinutes(-$_))
            Set-OfficeExcelCell -Row $rowIndex -Column 2 -Value "/api/data"
            Set-OfficeExcelCell -Row $rowIndex -Column 3 -Value (Get-Random -Minimum 20 -Maximum 80)
            $rowIndex++
        }
    }
}

# Create a PowerPoint presentation
$ppt = New-OfficePowerPoint -FilePath ".\Status.pptx"
$slide = Add-OfficePowerPointSlide -Presentation $ppt
Add-OfficePowerPointTextBox -Slide $slide -Text "Weekly Status" -X 80 -Y 60 -Width 520 -Height 50
Add-OfficePowerPointBullets -Slide $slide -Bullets @(
    "Shipped v3.2 to production",
    "Resolved 14 customer tickets",
    "Completed security audit"
) -X 80 -Y 150 -Width 520 -Height 220
$ppt | Save-OfficePowerPoint
```

## Compatibility

| Platform | PowerShell Version | Supported |
|----------|-------------------|-----------|
| Windows  | 5.1               | Yes       |
| Windows  | 7+                | Yes       |
| Linux    | 7+                | Yes       |
| macOS    | 7+                | Yes       |

PSWriteOffice is available from the [PowerShell Gallery](https://www.powershellgallery.com/packages/PSWriteOffice) and includes generated help for its cmdlet surface.

## Related guides

| Guide | Description |
|-------|-------------|
| [PSWriteOffice overview](/docs/pswriteoffice/) | Start with installation, command families, and module scope. |
| [PowerPoint cmdlets](/docs/pswriteoffice/powerpoint/) | Build generated decks with slides, bullets, tables, and images. |
| [Markdown cmdlets](/docs/pswriteoffice/markdown/) | Generate Markdown reports and repository-friendly docs from scripts. |
| [PowerShell API reference](/api/powershell/) | Browse the full cmdlet surface with parameters and examples. |

## Related packages

| Package | Description |
|---------|-------------|
| [OfficeIMO.Word](/products/word/) | .NET library powering Word document creation |
| [OfficeIMO.Excel](/products/excel/) | .NET library powering Excel workbook creation |
| [OfficeIMO.PowerPoint](/products/powerpoint/) | .NET library powering PowerPoint presentation creation |
