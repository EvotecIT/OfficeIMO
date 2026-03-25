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

PSWriteOffice brings the full power of OfficeIMO to PowerShell. With 150+ cmdlets and intuitive DSL aliases, you can create Word documents, Excel workbooks, and PowerPoint presentations in a few lines of script. It runs cross-platform on Windows, Linux, and macOS with PowerShell 5.1 and 7+.

## Features

- **150+ cmdlets** -- comprehensive coverage of document creation, formatting, and export
- **DSL aliases** -- intuitive shorthand like `WordParagraph`, `ExcelSheet`, and `PptSlide` for concise scripts
- **Word, Excel, PowerPoint & Markdown** -- create and modify all major Office formats from one module
- **Cross-platform** -- runs on Windows, Linux, and macOS wherever PowerShell is available
- **PowerShell 5.1 and 7+** -- compatible with Windows PowerShell and modern PowerShell Core

## Quick start

```powershell
# Install the module
Install-Module PSWriteOffice -Force

# Create a Word document
New-OfficeWord -FilePath "Report.docx" {
    WordParagraph -Text "Monthly Report" -Style Heading1
    WordParagraph -Text "Generated on $(Get-Date -Format 'MMMM yyyy')"
    WordParagraph -Text ""

    $data = @(
        @{ Region = "North"; Revenue = 42000; Target = 40000 }
        @{ Region = "South"; Revenue = 38000; Target = 35000 }
        @{ Region = "East";  Revenue = 29000; Target = 30000 }
        @{ Region = "West";  Revenue = 51000; Target = 45000 }
    )

    WordTable -DataTable $data -Style GridTable4Accent1
    WordParagraph -Text ""
    WordParagraph -Text "All regions met or exceeded targets except East." -Bold $false
}

# Create an Excel workbook
New-OfficeExcel -FilePath "Metrics.xlsx" {
    ExcelSheet -Name "Summary" {
        ExcelRow -Values "Metric", "Value", "Status"
        ExcelRow -Values "Uptime", "99.97%", "OK"
        ExcelRow -Values "Response Time", "42ms", "OK"
        ExcelRow -Values "Error Rate", "0.03%", "OK"
        ExcelRow -Values "Throughput", "12,400 req/s", "OK"
    }
    ExcelSheet -Name "Details" {
        ExcelRow -Values "Timestamp", "Endpoint", "Latency"
        1..100 | ForEach-Object {
            ExcelRow -Values (Get-Date).AddMinutes(-$_), "/api/data", (Get-Random -Min 20 -Max 80)
        }
    }
}

# Create a PowerPoint presentation
New-OfficePowerPoint -FilePath "Status.pptx" {
    PptSlide -Layout Title {
        PptTitle -Text "Weekly Status"
        PptSubtitle -Text "Engineering Team"
    }
    PptSlide -Layout TitleAndContent {
        PptTitle -Text "Completed This Week"
        PptContent {
            PptBullet -Text "Shipped v3.2 to production"
            PptBullet -Text "Resolved 14 customer tickets"
            PptBullet -Text "Completed security audit"
        }
    }
}
```

## Compatibility

| Platform | PowerShell Version | Supported |
|----------|-------------------|-----------|
| Windows  | 5.1               | Yes       |
| Windows  | 7+                | Yes       |
| Linux    | 7+                | Yes       |
| macOS    | 7+                | Yes       |

PSWriteOffice is available from the [PowerShell Gallery](https://www.powershellgallery.com/packages/PSWriteOffice) and can be installed with a single command.

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
