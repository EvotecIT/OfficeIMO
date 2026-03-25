---
title: Word Cmdlets
description: PSWriteOffice cmdlets for creating and editing Word documents in PowerShell.
order: 61
---

# Word Cmdlets

PSWriteOffice provides a comprehensive set of PowerShell cmdlets for creating, editing, and saving Word documents. These cmdlets wrap the OfficeIMO.Word .NET library and expose its functionality through idiomatic PowerShell parameters.

## Creating a Document

```powershell
# Create a new empty document
$doc = New-OfficeWord -FilePath "C:\Output\report.docx"
```

The `-FilePath` parameter specifies where the document will be saved. The file is created immediately but not written until `Save-OfficeWord` is called.

## Opening an Existing Document

```powershell
$doc = Get-OfficeWord -FilePath "C:\Input\existing.docx"
```

## Adding Paragraphs

```powershell
# Simple text
$doc | Add-OfficeWordParagraph -Text "Hello, World!"

# With formatting
$doc | Add-OfficeWordParagraph -Text "Bold Title" -Bold -FontSize 20

$doc | Add-OfficeWordParagraph -Text "Italic subtitle" -Italic -FontSize 14

$doc | Add-OfficeWordParagraph -Text "Colored text" -Color "Blue" -FontFamily "Arial"

# Underline
$doc | Add-OfficeWordParagraph -Text "Important note" -Underline

# Alignment
$doc | Add-OfficeWordParagraph -Text "Centered text" -Alignment Center
```

### Paragraph Formatting Parameters

| Parameter | Type | Description |
|-----------|------|-------------|
| `-Text` | String | The paragraph text |
| `-Bold` | Switch | Bold formatting |
| `-Italic` | Switch | Italic formatting |
| `-Underline` | Switch | Underline formatting |
| `-FontSize` | Int | Font size in half-points |
| `-FontFamily` | String | Font family name |
| `-Color` | String | Text color (name or hex) |
| `-Alignment` | String | Left, Center, Right, or Both |
| `-Style` | String | Paragraph style (Heading1, Heading2, etc.) |

## Adding Sections

Sections allow you to change page layout within a document:

```powershell
$doc | Add-OfficeWordSection

# Add content to the new section
$doc | Add-OfficeWordParagraph -Text "New section content"
```

## Adding Tables

```powershell
# Create a table with specified rows and columns
$doc | Add-OfficeWordTable -Rows 4 -Columns 3 -Style "TableGrid"
```

### Populating Table Cells

After creating a table, access it through the document object:

```powershell
$table = $doc.Tables[-1]  # Last table added

$table.Rows[0].Cells[0].Paragraphs[0].Text = "Name"
$table.Rows[0].Cells[1].Paragraphs[0].Text = "Role"
$table.Rows[0].Cells[2].Paragraphs[0].Text = "Status"

$table.Rows[1].Cells[0].Paragraphs[0].Text = "Alice"
$table.Rows[1].Cells[1].Paragraphs[0].Text = "Developer"
$table.Rows[1].Cells[2].Paragraphs[0].Text = "Active"
```

### Table from PowerShell Objects

A common pattern is generating tables from PowerShell objects:

```powershell
$services = Get-Service | Select-Object -First 10 -Property Name, Status, StartType

# Create the table
$doc | Add-OfficeWordTable -Rows ($services.Count + 1) -Columns 3 -Style "GridTable4Accent1"

$table = $doc.Tables[-1]

# Header row
$table.Rows[0].Cells[0].Paragraphs[0].Text = "Service Name"
$table.Rows[0].Cells[1].Paragraphs[0].Text = "Status"
$table.Rows[0].Cells[2].Paragraphs[0].Text = "Start Type"

# Data rows
for ($i = 0; $i -lt $services.Count; $i++) {
    $table.Rows[$i + 1].Cells[0].Paragraphs[0].Text = $services[$i].Name
    $table.Rows[$i + 1].Cells[1].Paragraphs[0].Text = $services[$i].Status.ToString()
    $table.Rows[$i + 1].Cells[2].Paragraphs[0].Text = $services[$i].StartType.ToString()
}
```

## Adding Images

```powershell
$doc | Add-OfficeWordImage -FilePath "C:\Images\logo.png" -Width 200 -Height 60
```

## Page Breaks

```powershell
$doc | Add-OfficeWordPageBreak
```

## Headers and Footers

```powershell
# Enable headers and footers
$doc | Add-OfficeWordHeader -Text "Company Name - Confidential"
$doc | Add-OfficeWordFooter -Text "Page Footer"
```

## Saving and Closing

Always save and close the document when finished:

```powershell
$doc | Save-OfficeWord
$doc | Close-OfficeWord
```

Or use a `try/finally` block for safety:

```powershell
$doc = New-OfficeWord -FilePath "safe.docx"
try {
    $doc | Add-OfficeWordParagraph -Text "Content"
    $doc | Save-OfficeWord
}
finally {
    $doc | Close-OfficeWord
}
```

## Complete Example: Generating a Report

```powershell
Import-Module PSWriteOffice

$doc = New-OfficeWord -FilePath "C:\Reports\ServerReport.docx"

# Title
$doc | Add-OfficeWordParagraph -Text "Server Health Report" -Bold -FontSize 28 -Alignment Center

# Subtitle
$date = Get-Date -Format "MMMM dd, yyyy"
$doc | Add-OfficeWordParagraph -Text "Generated: $date" -Italic -FontSize 12 -Alignment Center

$doc | Add-OfficeWordPageBreak

# Disk usage section
$doc | Add-OfficeWordParagraph -Text "Disk Usage" -Bold -FontSize 20 -Style "Heading1"

$disks = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3"

$doc | Add-OfficeWordTable -Rows ($disks.Count + 1) -Columns 4 -Style "GridTable4Accent1"
$table = $doc.Tables[-1]

$table.Rows[0].Cells[0].Paragraphs[0].Text = "Drive"
$table.Rows[0].Cells[1].Paragraphs[0].Text = "Size (GB)"
$table.Rows[0].Cells[2].Paragraphs[0].Text = "Free (GB)"
$table.Rows[0].Cells[3].Paragraphs[0].Text = "% Free"

for ($i = 0; $i -lt $disks.Count; $i++) {
    $disk = $disks[$i]
    $table.Rows[$i + 1].Cells[0].Paragraphs[0].Text = $disk.DeviceID
    $table.Rows[$i + 1].Cells[1].Paragraphs[0].Text = [math]::Round($disk.Size / 1GB, 1).ToString()
    $table.Rows[$i + 1].Cells[2].Paragraphs[0].Text = [math]::Round($disk.FreeSpace / 1GB, 1).ToString()
    $table.Rows[$i + 1].Cells[3].Paragraphs[0].Text = [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 1).ToString()
}

$doc | Save-OfficeWord
$doc | Close-OfficeWord

Write-Host "Report saved to C:\Reports\ServerReport.docx"
```
