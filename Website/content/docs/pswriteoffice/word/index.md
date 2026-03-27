---
title: Word Cmdlets
description: PSWriteOffice cmdlets for creating and editing Word documents in PowerShell.
order: 61
---

# Word Cmdlets

PSWriteOffice provides a Word automation surface for creating and editing `.docx` files from scripts. The examples below stay close to the generated help so they reflect the real module surface.

## Creating a Document

```powershell
# Create a new document and return the document object
$doc = New-OfficeWord -Path "C:\Output\report.docx" -PassThru
```

Use `-PassThru` when you want the document object back for further piping. For one-shot DSL usage, you can also call `New-OfficeWord -Path ... { ... }`.

## Opening an Existing Document

```powershell
$doc = Get-OfficeWord -Path "C:\Input\existing.docx"
```

## Adding Paragraphs

```powershell
# Simple text
$doc | Add-OfficeWordParagraph -Text "Hello, World!"

# Styled heading
$doc | Add-OfficeWordParagraph -Text "Report Title" -Style Heading1

# Paragraph with inline formatting
$doc | Add-OfficeWordParagraph {
    Add-OfficeWordText -Text "Important: " -Bold
    Add-OfficeWordText -Text "review this section." -Italic
}

# Alignment
$doc | Add-OfficeWordParagraph -Text "Centered text" -Alignment Center
```

### Common Paragraph Parameters

| Parameter | Type | Description |
|-----------|------|-------------|
| `-Text` | String | The paragraph text |
| `-Alignment` | String | Left, Center, Right, or Both |
| `-Style` | String | Paragraph style such as `Heading1` |
| `-Content` | ScriptBlock | Nested content via `Add-OfficeWordText` and related commands |

## Adding Sections

Sections are the main container when you want to group body content, headers, or footers:

```powershell
$doc | Add-OfficeWordSection {
    Add-OfficeWordParagraph -Text "New section content"
}
```

## Adding Tables

The easiest pattern is to hand the cmdlet object data directly:

```powershell
$services = Get-Service | Select-Object -First 10 -Property Name, Status, StartType
$doc | Add-OfficeWordTable -InputObject $services -Style "GridTable4Accent1"
```

## Adding Images

Add images inside a paragraph block:

```powershell
$doc | Add-OfficeWordParagraph {
    Add-OfficeWordImage -Path "C:\Images\logo.png" -Width 200 -Height 60
}
```

## Headers and Footers

Headers and footers are typically created inside a section:

```powershell
$doc | Add-OfficeWordSection {
    Add-OfficeWordHeader {
        Add-OfficeWordParagraph -Text "Company Name - Confidential" -Style Heading3
    }
    Add-OfficeWordFooter {
        Add-OfficeWordPageNumber -IncludeTotalPages
    }
}
```

## Saving and Closing

Always save and close the document when finished:

```powershell
$doc | Save-OfficeWord
Close-OfficeWord -Document $doc
```

Or use a `try/finally` block for safety:

```powershell
$doc = New-OfficeWord -Path "safe.docx" -PassThru
try {
    $doc | Add-OfficeWordParagraph -Text "Content"
    $doc | Save-OfficeWord
}
finally {
    Close-OfficeWord -Document $doc
}
```

## Complete Example: Generating a Report

```powershell
Import-Module PSWriteOffice

$doc = New-OfficeWord -Path "C:\Reports\ServerReport.docx" -PassThru

$doc | Add-OfficeWordParagraph -Text "Server Health Report" -Style Heading1 -Alignment Center

$date = Get-Date -Format "MMMM dd, yyyy"
$doc | Add-OfficeWordParagraph {
    Add-OfficeWordText -Text "Generated: $date" -Italic
}

$disks = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3"
$diskRows = $disks | ForEach-Object {
    [pscustomobject]@{
        Drive      = $_.DeviceID
        'Size GB'  = [math]::Round($_.Size / 1GB, 1)
        'Free GB'  = [math]::Round($_.FreeSpace / 1GB, 1)
        'Pct Free' = [math]::Round(($_.FreeSpace / $_.Size) * 100, 1)
    }
}

$doc | Add-OfficeWordSection {
    Add-OfficeWordParagraph -Text "Disk Usage" -Style Heading1
    Add-OfficeWordTable -InputObject $diskRows -Style "GridTable4Accent1"
}

$doc | Save-OfficeWord
Close-OfficeWord -Document $doc

Write-Host "Report saved to C:\Reports\ServerReport.docx"
```
