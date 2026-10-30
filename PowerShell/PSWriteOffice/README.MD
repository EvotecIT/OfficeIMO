# PSWriteOffice

Create, read, edit and convert documents from PowerShell through [OfficeIMO](https://github.com/EvotecIT/OfficeIMO). Microsoft Office is not required.

```powershell
New-OfficePowerPoint -Path ./Service-Review.pptx {
    PowerPointSlide -Title 'Service review' -Subtitle 'October 2026'
    PowerPointSlide -Title 'Current signals' {
        PowerPointCard 'Availability' -Text 'All regions healthy'
        PowerPointCard 'Capacity' -Text 'Storage at 71%'
    }
}
```

The theme supplies fonts, colors, spacing and layouts. Content commands use spelled-out format aliases, such as `PowerPointSlide`, `WordParagraph` and `ExcelTable`.

[![PowerShell Gallery](https://img.shields.io/powershellgallery/v/PSWriteOffice.svg)](https://www.powershellgallery.com/packages/PSWriteOffice)
[![Tests](https://github.com/EvotecIT/OfficeIMO/actions/workflows/powershell-tests.yml/badge.svg?branch=master)](https://github.com/EvotecIT/OfficeIMO/actions/workflows/powershell-tests.yml)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

## Install

```powershell
Install-Module PSWriteOffice -Scope CurrentUser
Import-Module PSWriteOffice
```

Use PowerShell 7.4 or newer on Windows, Linux or macOS, or Windows PowerShell 5.1 with .NET Framework 4.7.2. Some operations have additional requirements; check the owning command's help. OCR can use an optional Tesseract provider, and printer delivery uses the operating system's print services.

PSWriteOffice is maintained in the OfficeIMO repository. Report problems in [OfficeIMO issues](https://github.com/EvotecIT/OfficeIMO/issues). See [MIGRATION.md](../../MIGRATION.md) for upgrade changes and [GitHub Releases](https://github.com/EvotecIT/OfficeIMO/releases) for release history.

## Choose a workflow

| Task | Start with |
| --- | --- |
| Export PowerShell objects to a workbook | `Export-OfficeExcel` |
| Compose a document | `New-OfficeWord`, `New-OfficeExcel`, `New-OfficePowerPoint`, `New-OfficePdf` or `New-OfficeMarkdown` |
| Convert an existing file | `Convert-OfficeDocument` |
| Read content as text, Markdown, tables or structured data | `Read-OfficeDocument` |
| Search files and folders | `Search-OfficeDocument` |
| Inspect native counts, protection and package details | `Get-OfficeDocumentInfo` |
| Render document pages or slides | `Export-OfficeWordImage`, `Export-OfficeExcelImage`, `Export-OfficePowerPointImage` or `Export-OfficePdfImage` |

Use `Get-Help <command> -Full` for parameter sets and examples, or browse the [command reference](Docs/Readme.md). The [example library](Examples/README.md) contains complete scripts grouped by task and format.

## Export rows to Excel

```powershell
$rows = @(
    [pscustomobject]@{ Service = 'Directory'; Incidents = 2; Owner = 'Identity' }
    [pscustomobject]@{ Service = 'Messaging'; Incidents = 1; Owner = 'Collaboration' }
)
$rows | Export-OfficeExcel -Path ./Incidents.xlsx -AutoFit -FreezeTopRow
```

For a larger workflow, see the [readable Excel review](Examples/Excel/Recipe-Excel-ReadableReview.ps1) and [CSV and database examples](Examples/README.md#inspect-convert-and-integrate).

## Compose a PowerPoint deck

Slides choose a layout from their content. Metrics create a KPI layout; cards create a card grid; steps create a process; titled columns create a comparison. Tables and lists use the shared layout engine's continuation rules.

```powershell
New-OfficePowerPoint -Path ./Operations.pptx -Theme Modern {
    PowerPointSlide -Title 'Operations' -Subtitle 'Monthly review'
    PowerPointSection 'Performance'
    PowerPointSlide -Title 'Service health' {
        PowerPointKpi '99.95%' 'Availability' -Change '+0.02 pt'
        PowerPointKpi '182 ms' 'Median response' -Color Success
    }
    PowerPointSection 'Operating model'
    PowerPointSlide -Title 'Operating rhythm' {
        PowerPointStep 'Collect' 'Gather health signals'
        PowerPointStep 'Review' 'Confirm decisions with owners'
        PowerPointStep 'Act' 'Track actions to closure'
    }
    PowerPointSlide -Title 'Incidents' {
        PowerPointTable -InputObject $rows -Property Service, Incidents, Owner
    }
}
```

Choose `-Theme Modern`, `Boardroom` or `Dark`, supply a theme object, or use a PowerPoint template path. Use `-Layout` on a slide when its purpose needs an explicit layout, such as `Section`, `Quote` or `Closing`.

`PowerPointSection` groups the following slides, including their continuations. Use `-Slide` with a native slide or a one-based number to start a section in an existing deck.

`Get-OfficePowerPointSection` returns native section information. Its `SlideIndices` are zero-based; add one when passing an index to a command's `-Slide` parameter.

A PowerPoint content block saves and closes its document when it completes. A failed block leaves an existing destination unchanged. `New-OfficePowerPoint` requires `-Force` to replace an existing file; `-WhatIf` does not run the block. Add `-PassThru` to return the saved `FileInfo`, or `-Show` to open the result in its default application.

Without a content block, creation returns a live document. Save it explicitly and close it when finished:

```powershell
$document = New-OfficePowerPoint -Theme Boardroom
try {
    Add-OfficePowerPointSlide -Document $document -Title 'Customer onboarding' {
        PowerPointList 'Confirm acceptance criteria', 'Assign owners', 'Set the handover date'
    }
    Save-OfficePowerPoint -Document $document -Path ./Onboarding.pptx
} finally {
    Close-OfficePowerPoint -Document $document
}
```

Use `Open-OfficePowerPoint -LiteralPath ./Onboarding.pptx` to edit a saved deck. Closing disposes the live document; it does not save pending edits. [Object composition](Examples/PowerPoint/Recipe-PowerPoint-ObjectComposition.ps1) shows the complete lifecycle with speaker notes.

## Convert files

```powershell
Convert-OfficeDocument -Path ./Report.docx -To Pdf -OutputPath ./Report.pdf

Get-ChildItem ./Reports -File |
    Convert-OfficeDocument -To Pdf -OutputDirectory ./PDF -NoClobber
```

`-Path` supports input discovery; `-LiteralPath` treats a filename literally. `FileInfo` pipelines bind through their full paths. `-OutputPath` names one file; `-OutputDirectory` receives multiple outputs.

Conversions return reports that describe approximations, omissions and errors. Use `-FailOnLoss` when those reductions are unacceptable. `-NoClobber` preserves existing outputs; `-SkipUpToDate` skips an output only when it is newer than its source. A skipped input retains its reason in the result.

For restartable work, use [saved conversion jobs](Examples/Workflows/Recipe-SavedConversionJob.ps1). Jobs retain portable settings and checkpoint evidence; passwords are resolved at execution rather than stored in the job.

## Read and inspect

```powershell
Read-OfficeDocument -Path ./Report.pdf -As Markdown
Get-ChildItem ./Reports -File | Get-OfficeDocumentInfo
Search-OfficeDocument -Path ./Reports -Pattern 'renewal' -SimpleMatch -Recurse
```

For PowerPoint package validation, layout preflight and accessibility checks, use `Test-OfficePowerPoint`. Add feature and review inventories with a typed policy:

```powershell
$options = New-OfficePowerPointInspectionOptions -Features -ReviewComments
$inspection = Test-OfficePowerPoint -LiteralPath ./Operations.pptx -Options $options
$inspection.Accessibility
```

Reader output preserves source locations and provenance where the format provides them. Reconstructing an editable document from PDF or an image cannot restore information absent from its source. Check conversion reports and verify tables, amounts and reading order before using reconstructed output.

## Formats and examples

| Area | Examples |
| --- | --- |
| Word | [Reports and document editing](Examples/Word) |
| Excel and CSV | [Workbooks](Examples/Excel), [CSV](Examples/Csv) |
| PowerPoint | [Decks, templates and editing](Examples/PowerPoint) |
| PDF | [Composition, forms, extraction and operations](Examples/Pdf) |
| Markdown and other text formats | [Markdown](Examples/Markdown), [RTF](Examples/Rtf) |
| Visio | [Diagrams and stencils](Examples/Visio) |
| Mixed files | [Reader](Examples/Reader), [workflows](Examples/Workflows) |

The module also exposes OfficeIMO's Project, Access, EPUB, OpenDocument, email, iWork, OneNote, HTML, AsciiDoc, LaTeX and Confluence workflows. Use the [command index](Docs/Readme.md) to find their current operations and limits.

## Reference and contributions

- [Generated command reference](Docs/Readme.md)
- [Runnable examples](Examples/README.md)
- [OfficeIMO documentation](../../Docs/README.md)
- [Build and validation guide](../../Docs/building-and-releasing.md)
- [License](LICENSE)
