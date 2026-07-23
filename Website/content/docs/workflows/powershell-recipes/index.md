---
title: "PowerShell Document Recipes"
description: "Practical PSWriteOffice patterns for reports, presentations, PDF, CSV, Markdown, RTF, and document extraction."
order: 8
meta.seo_title: "PSWriteOffice examples for document automation"
---

PSWriteOffice exposes the OfficeIMO engines as pipeline-friendly commands and authoring blocks. These recipes are starting points for scheduled jobs, CI artifacts, admin tools, and repeatable reporting.

## Turn objects into a finished Excel workbook

```powershell
$rows = Get-Service | Select-Object Name, Status, StartType

$rows | Export-OfficeExcel `
    -Path '.\Service-Inventory.xlsx' `
    -WorksheetName 'Services' `
    -TableName 'ServiceInventory' `
    -BoldTopRow `
    -FreezeTopRow `
    -AutoFit
```

Use [database reporting](/docs/workflows/database-reporting/) when the source is DbaClientX, and the [operational dashboard example](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Showcase/Showcase-Excel-OperationalDashboard.ps1) when the workbook needs a more designed summary.

## Create Word and PDF side by side

```powershell
New-OfficeWord -Path '.\Change-Report.docx' -PdfPath '.\Change-Report.pdf' {
    Add-OfficeWordSection {
        Add-OfficeWordParagraph -Style Heading1 -Text 'Change report'
        Add-OfficeWordParagraph -Text "Generated $(Get-Date -Format u)"
        Add-OfficeWordTable -InputObject $changes -Style GridTable4Accent1
    }
}
```

The [PDF sidecar example](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Pdf/Example-OfficePdfSidecars.ps1) also shows Markdown-to-PDF output. Keep the editable source and delivery PDF together when both are part of the workflow contract.

## Build a PDF report directly

Use `New-OfficePdf` when the PDF is the primary artifact rather than a converted Office document:

```powershell
New-OfficePdf -Path '.\Deployment-Evidence.pdf' {
    Add-OfficePdfHeading -Text 'Deployment evidence' -Level 1
    Add-OfficePdfParagraph -Text 'Generated from the validated deployment run.'
}
```

Continue with the [PDF report DSL example](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Pdf/Example-PdfReportDsl.ps1) for the currently supported authoring components.

## Build repository-friendly Markdown

```powershell
New-OfficeMarkdown -Path '.\release-notes.md' {
    Add-OfficeMarkdownHeading -Text 'Release notes' -Level 1
    Add-OfficeMarkdownParagraph -Text 'Generated from the release pipeline.'
}
```

The [Markdown examples](https://github.com/EvotecIT/PSWriteOffice/tree/main/Examples/Markdown) cover the fluent block, tables, callouts, details, front matter, and conversion-oriented patterns.

## Export delimited data with an explicit contract

```powershell
$rows | Export-OfficeCsv -Path '.\inventory.csv' -Delimiter ';'
$table = Import-OfficeCsv -Path '.\inventory.csv' -AsDataTable -InferSchema
```

Choose the delimiter, culture, compression, quote policy, headers, schema, and error behavior deliberately. Use the [CSV examples](https://github.com/EvotecIT/PSWriteOffice/tree/main/Examples/Csv) and [performance evidence](/docs/workflows/powershell-benchmarks/) for large or database-shaped transfers.

## Extract across document formats

```powershell
$detection = Get-OfficeDocumentDetection -Path '.\incoming\proposal.docx'
$document = Get-OfficeDocument -Path '.\incoming\proposal.docx'
$chunks = Get-OfficeDocumentChunk -Path '.\incoming\proposal.docx'
```

Use the normalized Reader commands for indexing, migration, classification, search preparation, and bulk ingest. Use a native family when the script must modify Word, Excel, PowerPoint, PDF, Visio, RTF, or another format-specific model.

## Create and review a presentation

Start with the [PowerPoint service brief](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Showcase/Showcase-PowerPoint-ServiceBrief.ps1) for a complete deck, or the [PowerPoint examples](https://github.com/EvotecIT/PSWriteOffice/tree/main/Examples/PowerPoint) for focused charts, layouts, themes, transitions, imported slides, and HTML review.

See the [PSWriteOffice command families](/docs/pswriteoffice/command-families/) for the supported surface and [release previews](/docs/workflows/release-previews/) before relying on commands that are still waiting for package publication.
