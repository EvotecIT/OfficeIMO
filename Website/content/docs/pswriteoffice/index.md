---
title: PSWriteOffice
description: Overview of the PSWriteOffice PowerShell module for Word, Excel, PowerPoint, and Markdown automation.
order: 60
---

# PSWriteOffice

PSWriteOffice is the PowerShell surface for the OfficeIMO ecosystem. It brings document generation and automation to scripts, build agents, scheduled jobs, and admin tooling without requiring Microsoft Office to be installed.

## What it covers

- Word document creation and editing for reports, runbooks, and generated business documents.
- Excel workbook generation for exports, inventory reports, dashboards, and structured data handoffs.
- PowerPoint presentation automation for status decks, reviews, and generated slides.
- Markdown document generation for READMEs, reports, release notes, and developer-facing content.

## Installation

Install from the PowerShell Gallery:

```powershell
Install-Module -Name PSWriteOffice -Scope CurrentUser
```

For all users (requires Administrator):

```powershell
Install-Module -Name PSWriteOffice -Scope AllUsers
```

Update to the latest version:

```powershell
Update-Module -Name PSWriteOffice
```

## Key Cmdlets

### Word Cmdlets

| Cmdlet | Description |
|--------|-------------|
| `New-OfficeWord` | Create a new Word document |
| `Get-OfficeWord` | Open an existing Word document |
| `Save-OfficeWord` | Save the document to disk |
| `Close-OfficeWord` | Close and dispose the document |
| `Add-OfficeWordSection` | Add a new section |
| `Add-OfficeWordParagraph` | Add a paragraph |
| `Add-OfficeWordText` | Add formatted inline text to a paragraph |
| `Add-OfficeWordTable` | Render object data as a table |
| `Add-OfficeWordImage` | Add an image |
| `Add-OfficeWordHeader` | Add header content |
| `Add-OfficeWordFooter` | Add footer content |

### Excel Cmdlets

| Cmdlet | Description |
|--------|-------------|
| `New-OfficeExcel` | Create a new Excel workbook |
| `Get-OfficeExcel` | Open an existing workbook |
| `Save-OfficeExcel` | Save the workbook |
| `Close-OfficeExcel` | Close and dispose the workbook |
| `Add-OfficeExcelSheet` | Add a worksheet |
| `Set-OfficeExcelCell` | Write values or formulas to a cell |
| `Add-OfficeExcelTable` | Add a table to a worksheet |

### PowerPoint Cmdlets

| Cmdlet | Description |
|--------|-------------|
| `New-OfficePowerPoint` | Create a new presentation |
| `Add-OfficePowerPointSlide` | Add a slide to the presentation |
| `Add-OfficePowerPointTextBox` | Insert positioned text boxes |
| `Add-OfficePowerPointBullets` | Add bulleted content to a slide |
| `Add-OfficePowerPointTable` | Render tabular data in a slide |
| `Add-OfficePowerPointChart` | Add charts from series data |
| `Add-OfficePowerPointImage` | Place images on a slide |
| `Add-OfficePowerPointSection` | Group slides into named sections |
| `Save-OfficePowerPoint` | Persist the generated deck |

### Markdown Cmdlets

| Cmdlet | Description |
|--------|-------------|
| `New-OfficeMarkdown` | Create a new Markdown document |
| `Add-OfficeMarkdownHeading` | Add headings to a document |
| `Add-OfficeMarkdownParagraph` | Add body paragraphs |
| `Add-OfficeMarkdownTable` | Render object data as a Markdown table |
| `Add-OfficeMarkdownCode` | Add fenced code blocks |
| `Add-OfficeMarkdownCallout` | Add note, warning, or tip callouts |
| `Add-OfficeMarkdownTaskList` | Add GitHub-style task lists |
| `Add-OfficeMarkdownTableOfContents` | Generate a TOC from headings |

## Quick Example

```powershell
Import-Module PSWriteOffice

$doc = New-OfficeWord -Path "C:\Reports\Monthly.docx" -PassThru

$doc | Add-OfficeWordParagraph -Text "Monthly Status Report" -Style Heading1
$doc | Add-OfficeWordParagraph -Text "Generated on $(Get-Date -Format 'yyyy-MM-dd')"
$doc | Add-OfficeWordParagraph {
    Add-OfficeWordText -Text "All systems operational." -Bold
}

$doc | Save-OfficeWord
Close-OfficeWord -Document $doc
```

## Pipeline Support

The document object flows through the pipeline, which makes simple composition concise:

```powershell
$doc = New-OfficeWord -Path "pipeline.docx" -PassThru

$doc |
    Add-OfficeWordParagraph -Text "Title" -Style Heading1 |
    Add-OfficeWordParagraph {
        Add-OfficeWordText -Text "Subtitle" -Italic
    } |
    Add-OfficeWordParagraph -Text "Chapter 1 content" |
    Save-OfficeWord

Close-OfficeWord -Document $doc
```

## Discovering Available Commands

List all PSWriteOffice cmdlets:

```powershell
Get-Command -Module PSWriteOffice
```

Get help for a specific cmdlet:

```powershell
Get-Help New-OfficeWord -Full
Get-Help Add-OfficeWordParagraph -Examples
```

## Further Reading

- [Word Cmdlets](/docs/pswriteoffice/word) -- Detailed guide to Word document cmdlets.
- [Excel Cmdlets](/docs/pswriteoffice/excel) -- Detailed guide to Excel workbook cmdlets.
- [PowerPoint Cmdlets](/docs/pswriteoffice/powerpoint) -- Build presentation decks with cmdlets and DSL aliases.
- [Markdown Cmdlets](/docs/pswriteoffice/markdown) -- Generate Markdown reports and docs from PowerShell.
