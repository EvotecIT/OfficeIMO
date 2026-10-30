# PSWriteOffice

[![PowerShell Gallery Version](https://img.shields.io/powershellgallery/v/PSWriteOffice.svg)](https://www.powershellgallery.com/packages/PSWriteOffice)
[![PowerShell Gallery Downloads](https://img.shields.io/powershellgallery/dt/PSWriteOffice.svg)](https://www.powershellgallery.com/packages/PSWriteOffice)
[![Test PowerShell](https://github.com/EvotecIT/OfficeIMO/actions/workflows/powershell-tests.yml/badge.svg?branch=master)](https://github.com/EvotecIT/OfficeIMO/actions/workflows/powershell-tests.yml)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

PSWriteOffice creates, converts, reads and searches Word, Excel, PowerPoint, PDF and Markdown documents from PowerShell, plus Visio, Project, Access, OneNote, email, EPUB, RTF, OpenDocument, CSV and more. Microsoft Office is not required, the same scripts run on Windows, macOS and Linux, and new documents start from a theme, so the output looks finished without styling every element.

PSWriteOffice is built on [OfficeIMO](https://github.com/EvotecIT/OfficeIMO) and lives in the same repository. Report bugs and request features in [OfficeIMO issues](https://github.com/EvotecIT/OfficeIMO/issues).

## Install

```powershell
Install-Module PSWriteOffice -Scope CurrentUser
```

PSWriteOffice runs on PowerShell 7.4 and later (including 7.6), and on Windows PowerShell 5.1. Upgrading from 3.5 or earlier? The [migration guide](../../MIGRATION.md) maps every renamed and merged command, and calling a removed command prints its replacement.

## One line each

Export objects to a styled Excel table with a frozen header, fitted columns and number formats inferred from the data:

```powershell
Get-Process | Select-Object -First 20 Name, Id, CPU, WorkingSet | Export-OfficeExcel -Path .\Processes.xlsx
```

Export the same objects to a Word table or a deck. `Export-OfficePowerPoint` makes a title slide, then table slides that continue with the header repeated:

```powershell
Get-Service | Select-Object -First 15 Name, Status, StartType | Export-OfficeWord -Path .\Services.docx -Title 'Services'
Get-Service | Where-Object Status -eq Stopped | Export-OfficePowerPoint -Path .\Stopped.pptx -Title 'Stopped services'
```

Convert one file, or a whole folder in parallel:

```powershell
Convert-OfficeDocument -Path .\Services.docx -To Pdf
Get-ChildItem .\Reports -Filter *.docx | Convert-OfficeDocument -To Pdf -OutputDirectory .\Pdf -ThrottleLimit 4
```

Search Word, Excel, PowerPoint, PDF, Markdown and email files at once. `-Pattern` is a regular expression, as in `Select-String`:

```powershell
Search-OfficeDocument -Path .\Contracts -Recurse -Pattern 'termination|renewal'
```

Read a document as text, Markdown or chunks for search and AI ingestion:

```powershell
Read-OfficeDocument -Path .\Handbook.pdf -As Markdown
```

Read rows back from a workbook:

```powershell
Import-OfficeExcel -Path .\Processes.xlsx | Sort-Object CPU -Descending | Select-Object -First 5
```

## Word

Inside a content block, each command adds to the document. The block saves and closes the file when it ends:

```powershell
$milestones = @(
    [pscustomobject]@{ Milestone = 'Discovery'; Owner = 'Product'; Status = 'Done' }
    [pscustomobject]@{ Milestone = 'Pilot'; Owner = 'Operations'; Status = 'At risk' }
    [pscustomobject]@{ Milestone = 'Rollout'; Owner = 'IT'; Status = 'Planned' }
)
$trend = @(
    [pscustomobject]@{ Week = 'W1'; Complete = 20 }
    [pscustomobject]@{ Week = 'W2'; Complete = 45 }
    [pscustomobject]@{ Week = 'W3'; Complete = 70 }
)

New-OfficeWord -Path .\Status.docx -Theme Modern -Force {
    WordHeading 'Northwind migration'
    WordParagraph {
        WordText 'Overall status: ' -Bold
        WordText 'On track' -Bold -Color Success
    }
    WordTable -InputObject $milestones {
        WordTableCondition { $_.Status -eq 'At risk' } -BackgroundColor Danger
    }
    WordChart -Type Line -InputObject $trend -Category Week -Value Complete -Title 'Completion by week'
    WordFooter {
        WordText 'Internal · Page '
        WordPageNumber -IncludeTotal
    }
}
```

`Success`, `Warning`, `Danger` and `Accent1` to `Accent6` are theme colors, so they follow the document's theme.

## Excel

```powershell
$sales = @(
    [pscustomobject]@{ Region = 'North'; Revenue = 182400; Growth = 0.12 }
    [pscustomobject]@{ Region = 'South'; Revenue = 141900; Growth = 0.08 }
    [pscustomobject]@{ Region = 'West'; Revenue = 97300; Growth = -0.03 }
)

New-OfficeExcel -Path .\Sales.xlsx -Theme Modern -Force {
    ExcelSheet 'Sales' {
        ExcelTable $sales -Name Sales -ColumnFormat @{ Growth = 'Percent' }
        ExcelChart -Table Sales -Category Region -Value Revenue -Address F2 -Type Column -Title 'Revenue by region'
    }
}
```

`Export-OfficeExcel` covers the one-table case and streams very large inputs; `New-OfficeExcel` builds workbooks with several sheets, charts, formulas and formatting.

## PowerPoint

Slides are laid out from their content: a title alone makes a section slide, one list, table or chart makes a title-and-content slide, and KPIs make a metrics slide. No coordinates are needed:

```powershell
$quarters = @(
    [pscustomobject]@{ Quarter = 'Q1'; Revenue = 4.8 }
    [pscustomobject]@{ Quarter = 'Q2'; Revenue = 5.1 }
    [pscustomobject]@{ Quarter = 'Q3'; Revenue = 5.7 }
)

New-OfficePowerPoint -Path .\Review.pptx -Theme Boardroom -Force {
    PowerPointSlide -Title 'Quarterly business review' -Subtitle 'Performance, decisions and priorities'
    PowerPointSlide -Title 'Quarter at a glance' {
        PowerPointKpi '5.7m' 'Revenue (USD)' -Change '+12%' -Trend Up
        PowerPointKpi '36%' 'Operating margin' -Change '+1 pt' -Trend Up
        PowerPointKpi '3' 'Decisions needed'
    }
    PowerPointSlide -Title 'Revenue trend' {
        PowerPointChart -InputObject $quarters -Category Quarter -Value Revenue -Type Column -Title 'Revenue (USD m)'
        PowerPointNotes 'Explain the Q2 dip before the Q3 result.'
    }
    PowerPointSlide -Title 'Next steps' {
        PowerPointList 'Approve the budget', 'Confirm owners', 'Review in January'
    }
}
```

## PDF

```powershell
$services = @(
    [pscustomobject]@{ Service = 'Identity'; Availability = 99.98; Status = 'Healthy' }
    [pscustomobject]@{ Service = 'Mail'; Availability = 99.71; Status = 'Watch' }
)

New-OfficePdf -Path .\Status.pdf -Force {
    PdfHeading 'Weekly status'
    PdfParagraph {
        PdfText 'Overall: '
        PdfText 'Healthy' -Bold -Color Success
    }
    PdfTable -InputObject $services
}
```

Combine, split, protect and inspect existing files with the cross-format commands:

```powershell
Merge-OfficeDocument -Path .\Status.pdf, .\Services.pdf -OutputPath .\Pack.pdf
Get-OfficeDocumentInfo -Path .\Pack.pdf
Protect-OfficeDocument -Path .\Pack.pdf -Password (Read-Host -Prompt 'Password' -AsSecureString) -Encrypt
```

## Markdown

```powershell
New-OfficeMarkdown -Path .\Report.md -Force {
    MarkdownHeading 'Stopped services'
    MarkdownTable -InputObject (Get-Service | Where-Object Status -eq Stopped | Select-Object -First 10 Name, StartType)
}
Convert-OfficeDocument -Path .\Report.md -To Word
```

## Working with live documents

Leave out the content block to keep a document open, for example to build it in a loop, then save and close it yourself:

```powershell
$document = New-OfficeWord -Path .\Inventory.docx -Force
foreach ($server in 'SQL01', 'SQL02', 'WEB01') {
    Add-OfficeWordHeading -Document $document -Text $server -Level 2
    Add-OfficeWordParagraph -Document $document -Text "Checked on $(Get-Date -Format d)."
}
Save-OfficeWord -Document $document
Close-OfficeWord -Document $document
```

`Open-OfficeWord`, `Open-OfficeExcel`, `Open-OfficePowerPoint` and `Open-OfficePdf` edit existing files the same way, with or without a block.

## Conventions

Every command follows the same rules:

- `New-` refuses to replace an existing file unless you pass `-Force`. `Export-` and `Convert-` replace their output unless you pass `-NoClobber`. Commands that write files support `-WhatIf`.
- `-Path` is the file a command reads, or creates when it reads none; `-OutputPath` and `-OutputDirectory` name the output when there is also an input. `-LiteralPath` takes names with wildcard characters, and `Get-ChildItem` output pipes into every command that reads files.
- Lengths take a number of points or a unit: `-Width 5cm`, `-SpacingBefore 12pt`, `-Margin 1in`.
- Colors take `#RRGGBB`, a CSS name or a theme role (`Accent1`, `Success`, `Danger`). `-Theme` takes `Modern`, `Classic`, `Minimal`, `Boardroom`, `Dark` or the path of a `.potx`, `.dotx` or `.xltx` template.
- Pages, slides, sheets, rows and columns are numbered from 1. `-Sheet` takes a name or a position.
- On/off options are switches (`-Bold`, `-NoFreeze`); `-Bold:$false` turns formatting off explicitly.
- `-Password` takes a `SecureString`.

## Help

Every command has help with working examples:

```powershell
Get-Help New-OfficePowerPoint -Examples
Get-Command -Module PSWriteOffice -Noun OfficeExcel*
```

- [Command reference](Docs/Readme.md), one generated page per command.
- [Examples](Examples), runnable scripts for every format.
- [Migration guide](../../MIGRATION.md) for scripts written against 3.5 and earlier.

Electronic invoices are no longer part of PSWriteOffice; use the [OfficeIMO.Invoicing](../../OfficeIMO.Invoicing/README.md) packages from .NET.

## Build and test

```powershell
dotnet build .\Sources\PSWriteOffice.sln -c Debug
pwsh -NoLogo -NoProfile -File .\PSWriteOffice.Tests.ps1
pwsh -NoLogo -NoProfile -File .\Build\Build-Module.ps1 -RunMode Build
```

[PSWriteOffice.psm1](PSWriteOffice.psm1) loads the local debug build from `Sources\PSWriteOffice\bin\Debug\` during development.

## License

MIT
