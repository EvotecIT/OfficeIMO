---
title: "Pipeline, Object, and DSL Workflows"
description: "Choose the shortest PSWriteOffice surface for quick exports, incremental composition, complete document authoring, or existing-file updates."
layout: docs
---

PSWriteOffice supports several script shapes because a one-line export and a multi-section report are different jobs. Choose the smallest surface that still makes the document structure clear.

## Quick jobs: use the pipeline

When PowerShell objects already have the rows you want, send them directly to an export command:

```powershell
$services | Export-OfficeExcel -Path '.\Services.xlsx' -Sheet 'Services' -TableName 'Services'
```

This is the closest fit for inventory exports, query results, and data passed from modules such as DbaClientX. Start here unless the workbook needs several sheets, formulas, charts, or carefully placed content.

The same object boundary works for operational modules. The [PSEventViewer report recipe](https://github.com/EvotecIT/OfficeIMO/blob/a0abd1c0e95b75c0c91ad57b607038477a98462c/PowerShell/PSWriteOffice/Examples/Integrations/Recipe-PSEventViewer-OfficeReport.ps1) accepts a supplied snapshot or queries events once, projects stable report columns, and creates Excel, Word and PDF deliverables without requiring either module to know about the other's internals.

## Incremental jobs: keep the document object

Use an object when normal PowerShell control flow decides what to add. Create the document without a content block, pass it through composition commands, then save and close it once.

```powershell
$document = New-OfficeWord -Path '.\Access-Review.docx'
$heading = $document | Add-OfficeWordParagraph -Text 'Access review' -Style Heading1 -PassThru
$heading | Add-OfficeWordText -Text ' — weekly summary' -Color '#475569'

Add-OfficeWordTable -Document $document -InputObject $findings -Style GridTable4Accent1
$document | Save-OfficeWord
$document | Close-OfficeWord
```

The same model works for Excel:

```powershell
$workbook = New-OfficeExcel -Path '.\Projects.xlsx'
try {
    $sheet = $workbook | Add-OfficeExcelSheet -Name 'Projects' -PassThru
    Set-OfficeExcelCell -Sheet $sheet -Address A1 -Value 'Delivery portfolio'
    Add-OfficeExcelTable -Sheet $sheet -InputObject $projects -Address A3 -Name 'Projects'
    Set-OfficeExcelColumn -Sheet $sheet -Column C -Width 48
    Set-OfficeExcelRow -Sheet $sheet -Row 4 -WrapText -AutoFit
    $workbook | Save-OfficeExcel
} finally {
    $workbook | Close-OfficeExcel
}
```

Cells, rows, and columns accept a native worksheet through `-Sheet`, so layout changes fit ordinary PowerShell loops. Numbers, Boolean values, and dates retain their Excel types: `Set-OfficeExcelCell -Sheet $sheet -Address B4 -Value 1250.50 -NumberFormat '#,##0.00'` writes a number; a string such as `'00123'` remains text. Row and column commands are quiet by default; `-PassThru` returns the worksheet when another command needs it. To select a worksheet from a workbook, use `$workbook | Set-OfficeExcelRow -Sheet 'Projects' -Row 4 -Height 36`, or pass the one-based sheet position to `-Sheet`. Use one selector at a time. Apply changes before saving, and close the workbook in `finally` so errors also release the file.

The [readable review recipe](https://github.com/EvotecIT/OfficeIMO/blob/a0abd1c0e95b75c0c91ad57b607038477a98462c/PowerShell/PSWriteOffice/Examples/Excel/Recipe-Excel-ReadableReview.ps1) creates a workbook with fixed column widths, wrapped notes, fitted row heights, and an optional PDF copy using only public commands.

PowerPoint and Markdown also expose explicit presentation or document targets. See the [Word](https://github.com/EvotecIT/OfficeIMO/blob/a0abd1c0e95b75c0c91ad57b607038477a98462c/PowerShell/PSWriteOffice/Examples/Word/Recipe-Word-ObjectComposition.ps1), [Excel](https://github.com/EvotecIT/OfficeIMO/blob/a0abd1c0e95b75c0c91ad57b607038477a98462c/PowerShell/PSWriteOffice/Examples/Excel/Recipe-Excel-ObjectComposition.ps1), [PowerPoint](https://github.com/EvotecIT/OfficeIMO/blob/a0abd1c0e95b75c0c91ad57b607038477a98462c/PowerShell/PSWriteOffice/Examples/PowerPoint/Recipe-PowerPoint-ObjectComposition.ps1), and [Markdown](https://github.com/EvotecIT/OfficeIMO/blob/a0abd1c0e95b75c0c91ad57b607038477a98462c/PowerShell/PSWriteOffice/Examples/Markdown/Recipe-Markdown-ObjectComposition.ps1) object recipes for complete scripts.

## Complete authored documents: use the DSL

Use the DSL when the script owns the whole artifact and its nested structure is easier to read as one composition block:

```powershell
New-OfficePdf -Path '.\Service-Report.pdf' -Theme Modern {
    PdfHeading 'Service report'
    PdfParagraph 'Prepared for the weekly operations review.'
    PdfTable -InputObject $services
}
```

Keep lifecycle commands canonical; choose either short aliases or full command names for the content inside a block. Do not mix both styles in the same example. Saved constructors are quiet; add `-PassThru` only when another command needs the saved file.

PDF flowing content is composed through its DSL. Canvas, stamp, and page-overlay commands handle fixed coordinates after a PDF exists. See [position PDF content](/docs/pswriteoffice/pdf-positioned-content/) for those cases.

## Existing files: open, target, save, close

Do not rebuild a document when the task is a bounded update. Open the existing file, identify the native object to change, write a new result while developing, and close the document:

```powershell
$document = Open-OfficeWord -Path '.\Proposal.docx'
$paragraph = Get-OfficeWordParagraph -Document $document -Pattern '{{CustomerName}}' -SimpleMatch | Select-Object -First 1
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
