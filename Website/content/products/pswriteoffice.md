---
title: "PSWriteOffice"
description: "Automate Word, Excel, PowerPoint, PDF, Visio, Project, Access, EPUB, YAML, TOML, CSV, Markdown, OpenDocument, and email from PowerShell, on Windows, Linux, and macOS, without Office installed."
layout: product
meta.showcase_id: "pswriteoffice-patch-report"
product_color: "#d97706"
product_label: "PowerShell document automation"
runtime_label: "PowerShell 5.1 and 7+"
install: "Install-Module PSWriteOffice"
nuget: ""
docs_url: "/docs/pswriteoffice/"
api_url: "/api/powershell/"
meta.software.name: "PSWriteOffice"
meta.software.application_category: "DeveloperApplication"
meta.software.operating_system: "Windows, Linux, macOS"
meta.software.download_url: "https://www.powershellgallery.com/packages/PSWriteOffice"
meta.software.price: 0
meta.software.price_currency: "USD"
---

## Why PSWriteOffice?

PSWriteOffice brings the OfficeIMO document engines to PowerShell. Scripts create, read, update, convert, and check documents directly, so a scheduled job, a CI runner, or a Linux container needs neither Office nor LibreOffice. Commands return objects, work in the pipeline, and come with generated help.

## Features

- **Office and PDF** — Word, Excel, PowerPoint, and PDF authoring, inspection, merging, comparison, protection, and conversion
- **Configuration files and data** — edit YAML without losing comments, merge overlays, read and write TOML, and convert between JSON, YAML, TOML, XML, CSV, and Excel with a report of anything the target cannot hold
- **Files that usually need their own app** — Project schedules, Access table databases, EPUB books, Visio diagrams, OpenDocument, and email messages
- **Search and extraction** — one Reader model for text, tables, images, and OCR across mixed document folders
- **Readable scripts** — script-block DSLs and short aliases such as `WordParagraph`, `ExcelSheet`, and `PowerPointSlide`
- **Windows, Linux, and macOS** — PowerShell 7.4 or newer everywhere, and Windows PowerShell 5.1 on Windows

## Command families

{{< pswriteoffice-summary >}}

## Start from a real workflow

| Job | Working path |
|---|---|
| Review patch compliance and maintenance actions | [Patch report walkthrough and native downloads](/showcase/pswriteoffice-patch-report/) |
| Reconcile budget, actual spending and forecast | [Budget review, workbook and CSV input](/showcase/pswriteoffice-budget-review/) |
| Review delivery milestones and overdue actions | [Project review, milestone register and CSV input](/showcase/pswriteoffice-project-delivery/) |
| Export SQL data to Excel and verify the round trip | [DbaClientX database reporting guide](/docs/workflows/database-reporting/) and the {{< pswriteoffice-source kind="blob" path="Examples/Excel/Example-ExcelDbaClientXRoundTrip.ps1" label="runnable Excel example" >}} |
| Move database data through CSV | {{< pswriteoffice-source kind="blob" path="Examples/Csv/Example-CsvDbaClientXRoundTrip.ps1" label="Runnable CSV round-trip example" >}} |
| Create an operational Excel dashboard | {{< pswriteoffice-source kind="blob" path="Examples/Showcase/Showcase-Excel-OperationalDashboard.ps1" label="Excel showcase example" >}} |
| Review quarterly support demand in PowerPoint | {{< pswriteoffice-source kind="blob" path="Examples/Showcase/ServiceBrief.md" label="CSV-driven service brief" >}} with location totals, ticket comparisons and dated actions. |
| Update Helm values or CI settings without losing comments | [Edit and merge YAML](/docs/pswriteoffice/config-data-formats/#edit-yaml-without-rewriting-it) |
| Turn an API export or lock file into CSV, Excel, or YAML | [Convert between data formats](/docs/pswriteoffice/config-data-formats/#convert-between-data-formats) |
| Schedule a project and publish a Gantt PDF | [Project schedules](/docs/pswriteoffice/project-access-epub/#project-schedules) |
| Browse copyable PowerShell patterns | [PSWriteOffice recipe gallery](/docs/workflows/powershell-recipes/) |
| Compare Excel and CSV performance | [PSWriteOffice benchmark guide](/docs/workflows/powershell-benchmarks/) |
| Evaluate mixed search, authenticated PDF, or Confluence publishing | [PSWriteOffice command families](/docs/pswriteoffice/command-families/) |

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

## Compatibility

| Platform | PowerShell Version | Supported |
|----------|-------------------|-----------|
| Windows  | 5.1               | Yes       |
| Windows  | 7.4+              | Yes       |
| Linux    | 7.4+              | Yes       |
| macOS    | 7.4+              | Yes       |

PSWriteOffice is available from the [PowerShell Gallery](https://www.powershellgallery.com/packages/PSWriteOffice) and includes generated help for its cmdlet surface.

## Related guides

| Guide | Description |
|-------|-------------|
| [PSWriteOffice overview](/docs/pswriteoffice/) | Start with installation, command families, and module scope. |
| [PowerPoint cmdlets](/docs/pswriteoffice/powerpoint/) | Build generated decks with slides, bullets, tables, and images. |
| [Configuration files and data](/docs/pswriteoffice/config-data-formats/) | Edit YAML, read and write TOML, and convert between JSON, YAML, TOML, XML, CSV, and Excel. |
| [Project, Access, and EPUB](/docs/pswriteoffice/project-access-epub/) | Schedule projects, build Access tables, and compose EPUB books. |
| [Markdown cmdlets](/docs/pswriteoffice/open-text-formats/#markdown) | Generate Markdown reports and repository-friendly docs from scripts. |
| [Database reporting](/docs/workflows/database-reporting/) | Connect DbaClientX queries and bulk writes to Excel and CSV artifacts. |
| [Performance evidence](/docs/workflows/powershell-benchmarks/) | Reproduce the PowerForge-backed Excel and CSV comparisons. |
| [OfficeIMO roadmap](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/ROADMAP.md) | See open cross-package work without mixing it into current command documentation. |
| [PowerShell API reference](/api/powershell/) | Browse the full cmdlet surface with parameters and examples. |

## Related packages

| Package | Description |
|---------|-------------|
| [OfficeIMO.Word](/products/word/) | .NET library powering Word document creation |
| [OfficeIMO.Excel](/products/excel/) | .NET library powering Excel workbook creation |
| [OfficeIMO.PowerPoint](/products/powerpoint/) | .NET library powering PowerPoint presentation creation |
