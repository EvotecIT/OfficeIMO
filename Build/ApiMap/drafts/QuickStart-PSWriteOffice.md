# PSWriteOffice

Create, convert, read and search Word, Excel, PowerPoint, PDF and Markdown documents from PowerShell. Microsoft Office is not required, and the same scripts run on Windows, macOS and Linux. New documents start from a theme, so the output looks finished without styling every element.

## Install

```powershell
Install-Module PSWriteOffice -Scope CurrentUser
```

PowerShell 7.6 or later is the primary target. Windows PowerShell 5.1 runs the commands that work there; [Platform support](#platform-support) lists the exceptions.

## Quick start

Export objects to a styled Excel table with a frozen header, fitted columns and inferred number formats:

```powershell
Get-Process | Export-OfficeExcel -Path .\Processes.xlsx -Show
```

Convert a folder of Word documents to PDF, in parallel:

```powershell
Get-ChildItem .\Reports -Filter *.docx | Convert-OfficeDocument -To Pdf -OutputDirectory .\Pdf
```

Search Word, Excel, PowerPoint, PDF and Markdown files at once. `-Pattern` is a regular expression, as in `Select-String`:

```powershell
Search-OfficeDocument -Path .\Contracts -Pattern 'termination|renewal' -Recurse
```

Read a document as page-aware Markdown, ready for search or AI ingestion:

```powershell
Read-OfficeDocument -Path .\Handbook.pdf -As Markdown
```

Turn data into a deck: a title slide, then table slides that continue with the header repeated:

```powershell
Get-Process | Sort-Object CPU -Descending | Select-Object -First 30 Name, Id, CPU |
    Export-OfficePowerPoint -Path .\Processes.pptx -Title 'Busiest processes'
```

Turn a Markdown outline into slides. Headings start slides, lists become bullets and blockquotes become speaker notes:

```powershell
Convert-OfficeDocument -Path .\Review.md -To PowerPoint
```

Write a themed report. Colors such as `Success` and `Danger` come from the theme:

```powershell
$milestones = @(
    [pscustomobject]@{ Milestone = 'Discovery'; Owner = 'Product'; Status = 'Done' }
    [pscustomobject]@{ Milestone = 'Pilot'; Owner = 'Operations'; Status = 'At risk' }
)

New-OfficeWord -Path .\Status.docx -Theme Modern {
    WordHeading 'Northwind migration'
    WordParagraph {
        WordText 'Overall status: ' -Bold
        WordText 'On track' -Bold -Color Success
    }
    WordTable $milestones {
        WordTableCondition { $_.Status -eq 'At risk' } -BackgroundColor Danger
    }
    WordFooter { WordText 'Internal | Page '; WordPageNumber -IncludeTotal }
}
```

Build a deck without coordinates or layout names. Each slide's layout follows its content:

```powershell
New-OfficePowerPoint -Path .\Review.pptx -Theme Boardroom {
    PowerPointSlide -Title 'Quarterly business review' -Subtitle 'Performance, decisions and priorities'
    PowerPointSlide -Title 'Quarter at a glance' {
        PowerPointKpi '5.7m' 'Revenue (USD)' -Change '+12%' -Trend Up
        PowerPointKpi '36%' 'Operating margin' -Change '+1 pt' -Trend Up
        PowerPointKpi '3' 'Decisions needed'
    }
    PowerPointSlide -Title 'Next steps' {
        PowerPointStep 'Approve' 'Confirm owners for each priority.'
        PowerPointStep 'Deliver' 'Report progress monthly.'
        PowerPointStep 'Review' 'Assess results at quarter end.'
        PowerPointNotes 'Close by assigning each decision to a named owner.'
    }
}
```

Every command has help with working examples: `Get-Help New-OfficePowerPoint -Examples`. The `about_PSWriteOffice` topic explains the document model in one page.
