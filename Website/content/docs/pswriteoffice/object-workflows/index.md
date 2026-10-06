---
title: "Pipeline, Object, and DSL Workflows"
description: "Choose the shortest PSWriteOffice surface for quick exports, incremental composition, complete document authoring, or existing-file updates."
layout: docs
---

PSWriteOffice supports several script shapes because a one-line export and a multi-section report are different jobs. Choose the smallest surface that still makes the document structure clear.

## Quick jobs: use the pipeline

When PowerShell objects already have the rows you want, send them directly to an export command:

```powershell
$services | Export-OfficeExcel -Path '.\Services.xlsx' -WorksheetName 'Services' -TableName 'Services'
```

This is the closest fit for inventory exports, query results, and data passed from modules such as DbaClientX. Start here unless the workbook needs several sheets, formulas, charts, or carefully placed content.

The same object boundary works for operational modules. The [PSEventViewer report recipe](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Integrations/Recipe-PSEventViewer-OfficeReport.ps1) queries typed events once, projects stable report columns, and creates both Excel and Word deliverables without requiring either module to know about the other's internals.

## Incremental jobs: keep the document object

Use an object when normal PowerShell control flow decides what to add. Create the document with `-NoSave`, pass it through composition commands, then save and close it once.

```powershell
$document = New-OfficeWord -Path '.\Access-Review.docx' -NoSave
$heading = $document | Add-OfficeWordParagraph -Text 'Access review' -Style Heading1 -PassThru
$heading | Add-OfficeWordText -Text ' — weekly summary' -Color '#475569'

Add-OfficeWordTable -Document $document -InputObject $findings -Style GridTable4Accent1
$document | Save-OfficeWord
$document | Close-OfficeWord
```

The same model works for Excel:

```powershell
$workbook = New-OfficeExcel -Path '.\Projects.xlsx' -NoSave
$sheet = $workbook | Add-OfficeExcelSheet -Name 'Projects' -PassThru
$sheet | Set-OfficeExcelCell -Address A1 -Value 'Delivery portfolio'
Add-OfficeExcelTable -Worksheet $sheet -InputObject $projects -StartRow 3 -TableName 'Projects' -AutoFit
$workbook | Save-OfficeExcel
$workbook | Close-OfficeExcel
```

PowerPoint and Markdown also expose explicit presentation or document targets. See the [Word](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Word/Recipe-Word-ObjectComposition.ps1), [Excel](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Excel/Recipe-Excel-ObjectComposition.ps1), [PowerPoint](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/PowerPoint/Recipe-PowerPoint-ObjectComposition.ps1), and [Markdown](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Markdown/Recipe-Markdown-ObjectComposition.ps1) object recipes for complete scripts.

## Complete authored documents: use the DSL

Use the DSL when the script owns the whole artifact and its nested structure is easier to read as one composition block:

```powershell
PdfNew -Path '.\Service-Report.pdf' {
    PdfTheme Report
    PdfHeading 'Service report'
    PdfParagraph 'Prepared for the weekly operations review.'
    PdfTable -InputObject $services
}
```

Choose either short aliases or canonical `New-Office*` and `Add-Office*` names for a block. Do not mix both styles in the same example. Saved constructors are quiet; add `-PassThru` only when another command needs the saved file.

PDF flowing content is composed through its DSL. Canvas, stamp, and page-overlay commands handle fixed coordinates after a PDF exists. See [position PDF content](/docs/pswriteoffice/pdf-positioned-content/) for those cases.

## Existing files: open, target, save, close

Do not rebuild a document when the task is a bounded update. Open the existing file, identify the native object to change, write a new result while developing, and close the document:

```powershell
$document = Get-OfficeWord -Path '.\Proposal.docx'
$paragraph = Find-OfficeWordText -Document $document -Text '{{CustomerName}}' | Select-Object -First 1
Set-OfficeWordText -Paragraph $paragraph -Text 'Contoso'
$document | Save-OfficeWord -Path '.\Proposal-Contoso.docx'
$document | Close-OfficeWord
```

The exact target differs by format—paragraph, cell, slide, page, form field, or annotation—but the lifecycle stays recognizable.

## A practical rule

| Job | Start with |
| --- | --- |
| Export rows or query results | Pipeline |
| Add content from loops or conditions | Document object |
| Author the complete artifact in one place | DSL |
| Change a supplied document | Open, target, save, close |
| Read many formats into one downstream model | Reader |

All of these routes use the same OfficeIMO engines. They are complementary public surfaces, not separate feature sets.
