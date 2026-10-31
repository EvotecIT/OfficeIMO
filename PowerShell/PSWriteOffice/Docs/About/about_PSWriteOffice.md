---
topic: about_PSWriteOffice
schema: 1.0.0
---
# about_PSWriteOffice

## Short Description

Creates, reads, edits, inspects and converts documents with OfficeIMO.

## Long Description

Use New-OfficeWord, New-OfficeExcel and New-OfficePowerPoint to create
documents. Their script blocks collect content through format-specific
commands such as WordParagraph, ExcelTable and PowerPointSlide.

Short aliases spell out the format. PowerPointSlide and PowerPointList
refer to the same commands as Add-OfficePowerPointSlide and
Add-OfficePowerPointList. Get-Help accepts either name.

Use Open-OfficeWord, Open-OfficeExcel or Open-OfficePowerPoint to edit an
existing document. Save with the corresponding Save command and release
the document with Close. A destination on a New command saves the document
when its content block completes.

Convert-OfficeDocument converts files. OutputPath names one destination;
OutputDirectory names a folder for a batch. To selects the destination
format when its extension does not already select it. LiteralPath keeps
wildcard characters in a file name literal and accepts FileInfo pipelines.

Conversion reports describe each result and any fidelity diagnostics.
Read-OfficeDocument extracts content, Search-OfficeDocument finds text,
and Get-OfficeDocumentInfo inspects file information.

Export-OfficeWord and Export-OfficePowerPoint turn PowerShell objects into
documents. Property selects and orders columns. Use the individual content
commands when a document needs a particular composition.

Use PassThru to retain a native object from commands that provide it. For
example, a PowerPointList returns a text box whose paragraphs and runs can
be edited; a PowerPointTable returns a table. Automatic continuation keeps
the first native object and moves later content to continuation slides.

New Modern Word tables repeat their first row on subsequent pages,
including Word PDF export. Set the table's
RepeatHeaderRowAtTheTopOfEachPage property to false to disable repetition.

## Examples

Create a Word document:

```powershell
New-OfficeWord -Path ./Review.docx {
WordHeading 'Delivery review' -Level 1
WordParagraph 'The service is ready for review.'
}
```

Create a presentation:

```powershell
New-OfficePowerPoint -Path ./Review.pptx {
PowerPointSlide -Title 'Next steps' {
PowerPointList 'Review the results', 'Confirm the schedule'
}
}
```

Convert a document:

```powershell
Convert-OfficeDocument -LiteralPath './Review [final].docx' -OutputPath ./Review.pdf
```

Inspect a document:

```powershell
Get-OfficeDocumentInfo -Path ./Review.docx
```

Discover a command's parameters and examples:

```powershell
Get-Help PowerPointTable -Full
Get-Help Convert-OfficeDocument -Examples
```

## Runtimes

PSWriteOffice runs on PowerShell 7.4 and later, and on Windows PowerShell
5.1. Export-OfficeProjectReport, Get-OfficePrinter and Send-OfficePdfPrinter
require PowerShell 7.4 or later.

OfficeIMO processes documents without Microsoft Office installed.

## See Also

https://github.com/EvotecIT/OfficeIMO/tree/master/PowerShell/PSWriteOffice
https://github.com/EvotecIT/OfficeIMO/blob/master/MIGRATION.md

## Keywords

OfficeIMO, Word, Excel, PowerPoint, PDF, conversion, documents
