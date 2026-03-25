---
title: PSWriteOffice
description: Overview of the PSWriteOffice PowerShell module for creating Word and Excel documents.
order: 60
---

# PSWriteOffice

PSWriteOffice is a PowerShell module that wraps the OfficeIMO .NET libraries, bringing Word and Excel document creation to PowerShell scripts and automation workflows. It provides cmdlets for creating, editing, and saving Office documents without requiring Microsoft Office.

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
| `Add-OfficeWordParagraph` | Add a paragraph with text and formatting |
| `Add-OfficeWordTable` | Add a table |
| `Add-OfficeWordImage` | Add an image |
| `Add-OfficeWordPageBreak` | Insert a page break |
| `Add-OfficeWordHeader` | Add header content |
| `Add-OfficeWordFooter` | Add footer content |

### Excel Cmdlets

| Cmdlet | Description |
|--------|-------------|
| `New-OfficeExcel` | Create a new Excel workbook |
| `Get-OfficeExcel` | Open an existing workbook |
| `Save-OfficeExcel` | Save the workbook |
| `Close-OfficeExcel` | Close and dispose the workbook |
| `Add-OfficeExcelWorkSheet` | Add a worksheet |
| `Add-OfficeExcelTable` | Add a table to a worksheet |

## Quick Example

```powershell
Import-Module PSWriteOffice

# Create a Word document
$doc = New-OfficeWord -FilePath "C:\Reports\Monthly.docx"

$doc | Add-OfficeWordParagraph -Text "Monthly Status Report" -Bold -FontSize 24

$doc | Add-OfficeWordParagraph -Text "Generated on $(Get-Date -Format 'yyyy-MM-dd')"

$doc | Add-OfficeWordParagraph -Text "All systems operational." -Color "Green"

$doc | Save-OfficeWord
$doc | Close-OfficeWord
```

## Pipeline Support

All PSWriteOffice cmdlets support the PowerShell pipeline. The document object flows through the pipeline, allowing you to chain operations:

```powershell
$doc = New-OfficeWord -FilePath "pipeline.docx"

$doc |
    Add-OfficeWordParagraph -Text "Title" -Bold -FontSize 20 |
    Add-OfficeWordParagraph -Text "Subtitle" -Italic -FontSize 14 |
    Add-OfficeWordPageBreak |
    Add-OfficeWordParagraph -Text "Chapter 1 content" |
    Save-OfficeWord

$doc | Close-OfficeWord
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
