---
title: "PSWriteOffice DSL Cookbook"
description: "Copy complete PowerShell recipes for Word, Excel, PowerPoint, PDF, Markdown, and multi-format document jobs."
layout: docs
---

The PSWriteOffice DSL is a set of scoped PowerShell commands for composing documents. The outer `New-Office*` command creates and saves the file. Nested blocks establish the active section, sheet, slide, page, or Markdown document so content commands do not need the document object repeated on every line.

If loops or conditions need to keep and pass explicit document objects, use the companion [pipeline, object, and DSL workflow guide](/docs/pswriteoffice/object-workflows/).

## Choose one command style

The scenario recipes use short DSL aliases from the outer constructor through the nested content commands. This keeps composition blocks compact and avoids switching naming styles halfway through a document.

### DSL aliases

```powershell
$rows = @(
    [pscustomobject]@{ Name = 'Alpha'; Status = 'Ready' }
    [pscustomobject]@{ Name = 'Beta'; Status = 'Review' }
)

PdfNew -Path '.\Status.pdf' {
    PdfTheme Report
    PdfHeading -Text 'Status' -Level 1
    PdfTable -InputObject $rows
}
```

### Canonical cmdlets

```powershell
New-OfficePdf -Path '.\Status.pdf' {
    Set-OfficePdfTheme -Theme Report
    Add-OfficePdfHeading -Text 'Status' -Level 1
    Add-OfficePdfTable -InputObject $rows
}
```

These blocks call the same cmdlets and produce the same document. `PdfTheme` maps to `Set-OfficePdfTheme`: applying a theme changes the active PDF composition context, so the canonical verb is `Set`, not `New` or `Add`.

- Word: `WordNew` maps to `New-OfficeWord`; `WordSection` maps to `Add-OfficeWordSection`.
- Excel: `ExcelNew` maps to `New-OfficeExcel`.
- PowerPoint: `PptNew` maps to `New-OfficePowerPoint`.
- PDF: `PdfNew` maps to `New-OfficePdf`; `PdfTheme` maps to `Set-OfficePdfTheme`.
- Markdown: `MarkdownNew` maps to `New-OfficeMarkdown`.

Canonical command names are easier to discover in generated help. Aliases make dense composition blocks easier to scan. Pick one form for a script rather than mixing `New-OfficePdf` with `PdfTheme` and other aliases.

The same plain PowerShell objects can feed an Excel table, PowerPoint chart, PDF table, Word report, or Markdown document. Keep source collection and business calculations outside the DSL. Let the composition block describe the artifact.

## Write a formatted line with one command

Pass several strings to `WordText -Text` when every segment uses the same formatting. The strings are appended to one paragraph:

```powershell
WordNew -Path '.\Formatting.docx' {
    WordSection {
        WordText -Text @(
            'This is a text'
            ' that will show '
            'how WordText joins segments '
            'with the same formatting.'
        ) -FontFamily Tahoma -FontSize 10 -Color Blue
    }
}
```

Use `-Run` when formatting changes within the line. The compact columnar form keeps the text and its formatting arrays together:

```powershell
WordParagraph -Run @{
    Text      = @(
        'Owner: ', $finding.Owner
        '    Due: ', $finding.Due
        '    Severity: ', $finding.Severity
    )
    Bold      = $true, $false, $true, $false, $true, $false
    Underline = 'Single', 'None', $null, $null, $null, $null
    Color     = $null, $null, $null, $null, $null, 'Crimson'
}
```

The equivalent PDF line uses the same run shape:

```powershell
PdfText -Run @{
    Text = @(
        'Owner: ', $finding.Owner
        '    Due: ', $finding.Due
    )
    Bold = $true, $false, $true, $false
}
```

A scalar formatting value applies to every text segment. An array must contain either one value or the same number of values as `Text`, otherwise the command stops with a count error. There is no implicit `-ContinueFormatting`: use a scalar to broadcast intentionally, or put an explicit value or `$null` at each position.

For longer or generated content, `-Run` also accepts one hashtable per segment or objects created with `WordTextRun`, `PdfTextRun`, or the shared `TextRun` helper. Runs support bold, italic, underline and underline style, strike, foreground and background colors, font name and size, superscript or subscript baseline, and links.

## Flowing versus positioned PDF text

A `PdfText` run is inline content in the normal document flow. It does not have its own starting coordinate. Use the page-level positioning surface when text must begin at a fixed `X/Y` location:

```powershell
Add-OfficePdfCanvas -Path '.\Report.pdf' -OutputPath '.\Positioned.pdf' -Content {
    PdfCanvasText -Run @(
        TextRun 'Owner: ' -Bold
        TextRun 'Platform' -Color '#0F766E'
    ) -X 36 -Y 24
}
```

Canvas coordinates use PDF points from the visual top-left. `Add-OfficePdfStamp -X -Y` is the shorter choice for one text or image stamp. `Add-OfficePdfPageOverlay` positions a complete imported PDF page. See [Position text and graphics on PDF pages](/docs/pswriteoffice/pdf-positioned-content/) for the full decision guide and runnable recipe.

## Control pipeline output

Saved DSL constructors are silent by default, so they do not need `Out-Null` or a suppression switch. Add `-PassThru` only when the next command needs the saved file:

```powershell
$file = PdfNew -Path '.\Status.pdf' -PassThru {
    PdfHeading 'Status'
    PdfText 'Ready for review.'
}

$file | Select-Object Name, Length, LastWriteTime
```

When `New-OfficePdf` is used without a path, or with `-NoSave`, it returns the in-memory PDF document because no saved file exists.

## Word recipes

- [Project status report](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Word/Recipe-Word-ProjectStatus.ps1) combines a header and footer, narrative, lists, conditional table rows, a chart, and approval controls.
- [Change approval checklist](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Word/Recipe-Word-ApprovalChecklist.ps1) creates a reusable form with a table of contents, content controls, and a watermark.
- [Executive report](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Showcase/Showcase-Word-ExecutiveReport.ps1) is the larger reference for metadata, bookmarks, footnotes, endnotes, tables, charts, and rich text.
- [Word object composition](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Word/Recipe-Word-ObjectComposition.ps1) builds the same kind of content through a live document and paragraph targets.
- [Inspect, update, merge, and mail merge](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/README.md#read-modify-combine-and-convert) cover existing-document workflows.

Use Word when the reader needs a flowing, editable report with sections, review features, fields, or forms.

## Excel recipes

- [Project tracker](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Excel/Recipe-Excel-ProjectTracker.ps1) adds a structured table, status validation, conditional formatting, a chart, print settings, and an index sheet.
- [Budget dashboard](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Excel/Recipe-Excel-BudgetDashboard.ps1) separates summary formulas and charts from a styled detail table.
- [Operational dashboard](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Showcase/Showcase-Excel-OperationalDashboard.ps1) demonstrates KPI cells, tables, charts, pivots, sparklines, links, threaded comments, print layout, and workbook validation.
- [Quick export](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Excel/Recipe-Excel-QuickExport.ps1) and [Excel object composition](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Excel/Recipe-Excel-ObjectComposition.ps1) show the shorter alternatives to a complete workbook DSL.
- [Read, update, merge, compare, and import](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/README.md#read-modify-combine-and-convert) cover existing-workbook workflows.

Use Excel when the data grid, formula model, filtering, chart interaction, or workbook navigation is part of the deliverable.

## PowerPoint recipes

- [Quarterly business review](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/PowerPoint/Recipe-PowerPoint-QuarterlyReview.ps1) creates a title slide, a data chart, a priority table, bullets, and speaker notes.
- [Training workshop](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/PowerPoint/Recipe-PowerPoint-TrainingWorkshop.ps1) builds learning objectives, an agenda table, a call-to-action slide, and presenter notes.
- [Service brief](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Showcase/Showcase-PowerPoint-ServiceBrief.ps1) combines semantic designer plans with direct slide composition, charts, sections, transitions, and inspection.
- [PowerPoint object composition](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/PowerPoint/Recipe-PowerPoint-ObjectComposition.ps1) keeps a presentation object for loop-driven slide creation.
- [Inspect, update, and reuse slides](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/README.md#read-modify-combine-and-convert) cover existing-presentation workflows.

Use direct slide composition when the script owns placement. Use a deck plan when the content is semantic and the designer should choose layout variants.

## PDF recipes

- [Service invoice](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Pdf/Recipe-Pdf-ServiceInvoice.ps1) composes invoice metadata, line items, totals, payment terms, headers, footers, and a link.
- [Audit report](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Pdf/Recipe-Pdf-AuditReport.ps1) adds findings, bookmarks, page breaks, remediation sections, and interactive form fields.
- [Composed PDF report](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Pdf/Example-PdfReportDsl.ps1) demonstrates themes, backgrounds, borders, rich text, rows, links, bookmarks, tables, and attachments.
- [Form data exchange](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Pdf/Recipe-Pdf-FormDataExchange.ps1), [attachments](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Pdf/Recipe-Pdf-AttachEvidence.ps1), and [page reordering](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Pdf/Recipe-Pdf-ReorderPages.ps1) cover post-composition operations.
- [Inspect, merge, split, position, redact, sanitize, and process forms](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/README.md#read-modify-combine-and-convert) cover completed-PDF workflows.

Use PDF for fixed-layout delivery. Keep an editable source artifact as well when the workflow needs later content changes.

## Markdown recipes

- [Operations runbook](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Markdown/Recipe-Markdown-OperationsRunbook.ps1) creates front matter, a table of contents, warnings, task lists, code, a validation table, and collapsible rollback steps.
- [Release notes](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Markdown/Recipe-Markdown-ReleaseNotes.ps1) creates a release page with metadata, an upgrade callout, a change table, a checklist, code, and known limits.
- [Advanced Markdown](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Markdown/Example-MarkdownAdvanced.ps1) collects the broader typed Markdown surface in one script.
- [Markdown object composition](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Markdown/Recipe-Markdown-ObjectComposition.ps1) uses an explicit Markdown document target.
- [Inspect, publish HTML, and round-trip through Word](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/README.md#read-modify-combine-and-convert) cover Markdown transformation workflows.

Use Markdown when the source should remain diffable, reviewable, and easy to publish into other text or document workflows.

## Create several formats from one data model

The [multi-format status pack](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Workflows/Recipe-MultiFormat-StatusPack.ps1) sends one service-status object array into five thin composition blocks:

```text
PowerShell objects
  |-- Markdown status page
  |-- Word owner report
  |-- Excel analysis workbook
  |-- PowerPoint review deck
  `-- PDF delivery copy
```

This is useful when audiences need different artifacts but the numbers and status labels must remain consistent. Calculate the data once, then keep each document block focused on how that audience consumes it.

## Run and adapt a recipe

```powershell
.\Examples\Word\Recipe-Word-ProjectStatus.ps1
```

Replace the sample objects first. Then adjust visual choices such as styles, colors, layout, and labels. Search the [PowerShell command reference](/api/powershell/) for exact parameters and accepted values.

The [complete example index](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/README.md) also covers Visio, Reader, RTF, CSV, HTML review, DbaClientX, ChartForgeX visuals, and Confluence publishing.
