# PSWriteOffice example library

These recipes create, read, update, combine, or convert real files. Start with the workflow you need, then move to the larger showcase scripts when you want to see several features working together.

```powershell
Install-Module PSWriteOffice -Scope CurrentUser
```

Most `Recipe-*` scripts use paths relative to the current folder. Parameterized recipes expose an output directory; inspect their script help and supply it when required. For example, the patch-compliance recipe requires PSWriteOffice 3.0.9 or later and an explicit destination:

```powershell
./Examples/Word/Recipe-Word-PatchCompliance.ps1 -OutputDirectory ./Output/PatchCompliance
```

You can also copy a composition block into your own script and choose its inputs and file paths there.

Creation recipes demonstrate short DSL aliases such as `WordNew`, `ExcelNew`, `PptNew`, `PdfNew`, and `MarkdownNew`, as well as their canonical cmdlet names. Recipes that read, update, merge, split, or convert existing files use operation-specific cmdlets. The [DSL cookbook](https://officeimo.com/docs/pswriteoffice/dsl-cookbook/) shows composition in alias and canonical forms so you can choose one style for your own script.

Saved DSL constructors are quiet by default. Add `-PassThru` only when the next command needs the saved file.

## Choose a script shape

| When you need to | Use | Start here |
| --- | --- | --- |
| Save a conversion job and repeat it through Studio, CLI or PowerShell | Shared saved job | [Plan, run and resume](Workflows/Recipe-SavedConversionJob.ps1) |
| Export rows or query results | Pipeline | [Excel quick export](Excel/Recipe-Excel-QuickExport.ps1) |
| Add content from loops and conditions | Document objects | [Word](Word/Recipe-Word-ObjectComposition.ps1), [Excel](Excel/Recipe-Excel-ObjectComposition.ps1), [PowerPoint](PowerPoint/Recipe-PowerPoint-ObjectComposition.ps1), [Markdown](Markdown/Recipe-Markdown-ObjectComposition.ps1) |
| Author the complete artifact in one block | DSL | [Project status report](Word/Recipe-Word-ProjectStatus.ps1), [service invoice](Pdf/Recipe-Pdf-ServiceInvoice.ps1) |
| Produce a report with a table and chart | DSL with paragraph spacing and cell padding | [Patch compliance report](Word/Recipe-Word-PatchCompliance.ps1) |
| Change a supplied file | Open, target, save, close | [Word update](Word/Recipe-Word-UpdateExisting.ps1), [Excel update](Excel/Recipe-Excel-UpdateExisting.ps1), [PowerPoint update](PowerPoint/Recipe-PowerPoint-UpdateExisting.ps1) |
| Make an Excel review easy to read | Worksheet pipeline | [Column widths, wrapped notes, typed values, and PDF delivery](Excel/Recipe-Excel-ReadableReview.ps1) |

The [workflow guide](https://officeimo.com/docs/pswriteoffice/object-workflows/) explains when these surfaces fit. They use the same document engines and can coexist in a larger automation job.

## Create documents with the DSL

| Format | Practical recipes | Larger examples |
| --- | --- | --- |
| Word | [Project status](Word/Recipe-Word-ProjectStatus.ps1), [approval checklist](Word/Recipe-Word-ApprovalChecklist.ps1), [object composition](Word/Recipe-Word-ObjectComposition.ps1) | [Executive report](Showcase/Showcase-Word-ExecutiveReport.ps1), [advanced Word DSL](Word/Example-WordAdvanced.ps1) |
| Excel | [Quick export](Excel/Recipe-Excel-QuickExport.ps1), [project tracker](Excel/Recipe-Excel-ProjectTracker.ps1), [budget dashboard](Excel/Recipe-Excel-BudgetDashboard.ps1), [object composition](Excel/Recipe-Excel-ObjectComposition.ps1) | [Operational dashboard](Showcase/Showcase-Excel-OperationalDashboard.ps1), [advanced workbook](Excel/Example-ExcelAdvanced.ps1) |
| PowerPoint | [Quarterly review](PowerPoint/Recipe-PowerPoint-QuarterlyReview.ps1), [training workshop](PowerPoint/Recipe-PowerPoint-TrainingWorkshop.ps1), [object composition](PowerPoint/Recipe-PowerPoint-ObjectComposition.ps1), [sections and notes](PowerPoint/Recipe-PowerPoint-SectionsAndNotes.ps1) | [CSV-driven service brief](Showcase/ServiceBrief.md), [themes and layouts](PowerPoint/Example-PowerPointThemeAndLayout.ps1) |
| PDF | [Service invoice](Pdf/Recipe-Pdf-ServiceInvoice.ps1), [audit report](Pdf/Recipe-Pdf-AuditReport.ps1), [attach evidence](Pdf/Recipe-Pdf-AttachEvidence.ps1), [form data exchange](Pdf/Recipe-Pdf-FormDataExchange.ps1) | [Composed PDF report](Pdf/Example-PdfReportDsl.ps1), [PDF operations](Pdf/Example-PdfOperations.ps1) |
| Markdown | [Operations runbook](Markdown/Recipe-Markdown-OperationsRunbook.ps1), [release notes](Markdown/Recipe-Markdown-ReleaseNotes.ps1), [object composition](Markdown/Recipe-Markdown-ObjectComposition.ps1), [definition guide](Markdown/Recipe-Markdown-DefinitionGuide.ps1) | [Advanced Markdown](Markdown/Example-MarkdownAdvanced.ps1), [Markdown DSL](Markdown/Example-MarkdownDsl.ps1) |
| Several formats | [Status pack from shared data](Workflows/Recipe-MultiFormat-StatusPack.ps1), [Markdown to Word and PDF](Workflows/Recipe-Markdown-Word-PdfDelivery.ps1) | [Shared rich-text runs](Showcase/Showcase-RichTextRuns.ps1) |

## Read, modify, combine, and convert

| Format | Read and inspect | Modify or combine | Convert or deliver |
| --- | --- | --- | --- |
| Word | [Inspect an existing document](Word/Recipe-Word-InspectExisting.ps1) | [Update content](Word/Recipe-Word-UpdateExisting.ps1), [merge documents](Word/Recipe-Word-MergeDocuments.ps1), [compare versions](Word/Recipe-Word-CompareDocuments.ps1), [mail-merge letters](Word/Recipe-Word-MailMergeLetters.ps1) | [HTML to Word](Word/Recipe-Word-HtmlToDocument.ps1), [Word and Markdown conversion](Word/Example-WordMarkdownConvert.ps1), [HTML review](Word/Example-WordHtmlConvert.ps1) |
| Excel | [Read and filter rows](Excel/Recipe-Excel-ReadAndFilter.ps1) | [Update a workbook](Excel/Recipe-Excel-UpdateExisting.ps1), [append or replace rows](Excel/Recipe-Excel-AppendAndReplace.ps1), [merge](Excel/Recipe-Excel-MergeWorkbooks.ps1), [compare](Excel/Recipe-Excel-CompareWorkbooks.ps1) | [Template invoice](Excel/Recipe-Excel-TemplateInvoice.ps1), [pivots and sparklines](Excel/Recipe-Excel-PivotAndSparklines.ps1), [import delimited data](Excel/Recipe-Excel-ImportDelimited.ps1) |
| PowerPoint | [Inspect a deck](PowerPoint/Recipe-PowerPoint-InspectDeck.ps1) | [Update content](PowerPoint/Recipe-PowerPoint-UpdateExisting.ps1), [reuse slides](PowerPoint/Recipe-PowerPoint-ReuseSlides.ps1), [copy and remove slides](PowerPoint/Recipe-PowerPoint-CopyAndRemoveSlides.ps1) | [HTML review](PowerPoint/Example-PowerPointHtmlReview.ps1) |
| PDF | [Inspect and preflight](Pdf/Recipe-Pdf-InspectAndPreflight.ps1), [extract text](Pdf/Recipe-Pdf-ExtractText.ps1), [OCR images and scans](Pdf/Recipe-Pdf-Ocr.ps1) | [Merge and split](Pdf/Recipe-Pdf-MergeAndSplit.ps1), [reorder pages](Pdf/Recipe-Pdf-ReorderPages.ps1), [position content](Pdf/Recipe-Pdf-PositionedCanvas.ps1), [redact text](Pdf/Recipe-Pdf-RedactDetectedText.ps1) | [Forms](Pdf/Recipe-Pdf-Forms.ps1), [form data exchange](Pdf/Recipe-Pdf-FormDataExchange.ps1), [sanitize and optimize](Pdf/Recipe-Pdf-SanitizeAndOptimize.ps1) |
| Markdown | [Inspect structured content](Markdown/Recipe-Markdown-InspectContent.ps1) | [Convert to and from Word](Markdown/Recipe-Markdown-WordRoundTrip.ps1) | [Publish HTML](Markdown/Recipe-Markdown-PublishHtml.ps1) |
| Reader | [Search a mixed folder](Reader/Recipe-Reader-SearchFolder.ps1) | [Extract chunks and tables](Reader/Recipe-Reader-ExtractTables.ps1) | [Ingest a bounded folder](Reader/Recipe-Reader-IngestFolder.ps1) |

## Inspect, convert, and integrate

- [Mixed-format conversion review](Workflows/MixedFormat-ConversionReview.md): bounded folder-to-PDF conversion with file outcomes, diagnostics and hash-verified restart; requires PSWriteOffice 3.0.9 and PowerShell 7.4 or newer.
- [Visio service flow](Visio/Recipe-Visio-ServiceFlow.ps1), [architecture map](Visio/Example-Visio-ArchitectureMap.ps1), and [network topology](Visio/Example-Visio-NetworkTopology.ps1)
- [Mixed-document search](Reader/Example-MixedDocumentSearch.ps1), [bounded chunks](Reader/Recipe-Reader-BoundedChunks.ps1), and [table export](Reader/Recipe-Reader-ExportTables.ps1)
- [Update and convert RTF](Rtf/Recipe-Rtf-UpdateAndConvert.ps1) and [RTF/Markdown round trip](Rtf/Example-RtfMarkdownRoundTrip.ps1)
- [Safe CSV export](Csv/Recipe-Csv-SafeExport.ps1), [CSV basics](Csv/Example-CsvBasic.ps1), [advanced CSV options](Csv/Example-CsvAdvanced.ps1), and [DbaClientX round trip](Csv/Example-CsvDbaClientXRoundTrip.ps1)
- [Excel and DbaClientX round trip](Excel/Example-ExcelDbaClientXRoundTrip.ps1)
- [Review supplier offers from HTML](Integrations/Recipe-HtmlSupplierReview.ps1)
- [Prepare a local EML report package or deliver it with Mailozaurr](Integrations/Recipe-Mailozaurr-PdfDelivery.ps1)
- [Turn PSEventViewer results into Word and Excel reports](Integrations/Recipe-PSEventViewer-OfficeReport.ps1)
- [HTML review for Word](Word/Example-WordHtmlConvert.ps1), [Excel](Excel/Example-ExcelHtmlReview.ps1), and [PowerPoint](PowerPoint/Example-PowerPointHtmlReview.ps1)
- [ChartForgeX visuals](Visuals/Example-ChartForgeXVisuals.ps1)
- [Confluence report publishing](Confluence/Example-ConfluenceAzureTableReport.ps1)

The [OfficeIMO website](https://officeimo.com/docs/pswriteoffice/) ingests these scripts for the PowerShell reference and publishes the authored PSWriteOffice guides from this repository. Command pages remain the source of truth for exact parameters; the recipes show how those commands fit into complete jobs.
