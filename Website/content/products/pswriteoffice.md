---
title: "PSWriteOffice"
description: "Automate Word, Excel, PowerPoint, PDF, Reader, Visio, Markdown, CSV, RTF, OpenDocument, email, AsciiDoc, and LaTeX workflows from PowerShell."
layout: product
product_color: "#d97706"
product_label: "PowerShell document automation"
runtime_label: "PowerShell 5.1 and 7+"
install: "Install-Module PSWriteOffice"
nuget: ""
docs_url: "/docs/pswriteoffice/"
api_url: "/api/powershell/"
---

## Why PSWriteOffice?

PSWriteOffice brings the OfficeIMO document stack to PowerShell. Its authoritative module manifest exports **464 cmdlets and 354 aliases across 15 documentation families**. Create, inspect, update, review, convert, secure, and deliver documents from scripts and automation jobs without maintaining a separate document engine.

## Features

- **464 cmdlets** — manifest-validated coverage across 15 command families
- **DSL aliases** — intuitive shorthand like `WordParagraph`, `ExcelSheet`, and `PptSlide` for concise scripts
- **Broad document platform** — Word, Excel, PowerPoint, PDF, Reader, Visio, Markdown, CSV, RTF, OpenDocument, email, AsciiDoc, LaTeX, and HTML asset workflows
- **Cross-platform** — runs on Windows, Linux, and macOS wherever PowerShell is available
- **PowerShell 5.1 and 7+** — compatible with Windows PowerShell and modern PowerShell Core

## Command families

The counts below are generated from the same manifest used by module packaging. Cross-format commands have one documentation owner so the family totals equal the 464 exported cmdlets exactly.

| Family | Cmdlets | Typical work |
|---|---:|---|
| [Word](/docs/pswriteoffice/word/) | 91 | Authoring, inspection, review, mail merge, protection, and conversion |
| [Excel](/docs/pswriteoffice/excel/) | 155 | Worksheets, ranges, tables, formulas, charts, pivots, validation, repair, and comparison |
| [PowerPoint](/docs/pswriteoffice/powerpoint/) | 57 | Slides, layouts, text, images, shapes, tables, charts, notes, inspection, and rendering |
| [PDF](/docs/pswriteoffice/pdf/) | 74 | Authoring, preflight, extraction, operations, security, signatures, and redaction |
| [Markdown and text formats](/docs/pswriteoffice/open-text-formats/) | 53 | Markdown, RTF, CSV, OpenDocument, email, AsciiDoc, LaTeX, and HTML assets |
| [Visio](/docs/pswriteoffice/visio/) | 20 | Pages, shapes, connectors, stencils, inspection, arrangement, and SVG output |
| [Reader](/docs/pswriteoffice/reader/) | 13 | Format detection, normalized documents, chunks, tables, visuals, assets, and search |
| [Shared authoring primitive](/docs/pswriteoffice/automation-patterns/) | 1 | Reusable typed text runs |

Browse [all 15 generated families](/docs/pswriteoffice/command-families/) or search the [full PowerShell API reference](/api/powershell/) for parameter-level detail.

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
| [Markdown cmdlets](/docs/pswriteoffice/open-text-formats/#markdown) | Generate Markdown reports and repository-friendly docs from scripts. |
| [PowerShell API reference](/api/powershell/) | Browse the full cmdlet surface with parameters and examples. |

## Related packages

| Package | Description |
|---------|-------------|
| [OfficeIMO.Word](/products/word/) | .NET library powering Word document creation |
| [OfficeIMO.Excel](/products/excel/) | .NET library powering Excel workbook creation |
| [OfficeIMO.PowerPoint](/products/powerpoint/) | .NET library powering PowerPoint presentation creation |
