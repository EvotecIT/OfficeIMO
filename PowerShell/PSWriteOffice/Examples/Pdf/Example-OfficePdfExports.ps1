$ErrorActionPreference = 'Stop'

Import-Module PSWriteOffice -ErrorAction Stop

$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$wordPath = Join-Path $documents 'Example-OfficePdfExport.docx'
$wordPdf = Join-Path $documents 'Example-OfficePdfExport-Word.pdf'
$excelPath = Join-Path $documents 'Example-OfficePdfExport.xlsx'
$excelPdf = Join-Path $documents 'Example-OfficePdfExport-Excel.pdf'
$markdownPath = Join-Path $documents 'Example-OfficePdfExport.md'
$markdownPdf = Join-Path $documents 'Example-OfficePdfExport-Markdown.pdf'
$powerPointPath = Join-Path $documents 'Example-OfficePdfExport.pptx'
$powerPointPdf = Join-Path $documents 'Example-OfficePdfExport-PowerPoint.pdf'

$rows = @(
    [pscustomobject]@{ Name = 'Alpha'; Status = 'Ready'; Count = 12 }
    [pscustomobject]@{ Name = 'Beta'; Status = 'Review'; Count = 7 }
)

New-OfficeWord -Path $wordPath {
    WordParagraph -Text 'Word PDF export' -Style Heading1
    WordParagraph 'Create the source document, then choose the output format explicitly.'
    WordTable -InputObject $rows -Layout AutoFitToWindow
}
Export-OfficeDocumentPdf -InputPath $wordPath -Path $wordPdf

New-OfficeExcel -Path $excelPath {
    ExcelSheet -Name 'Summary' {
        ExcelTable -Data $rows
        ExcelAutoFit
    }
}
Export-OfficeDocumentPdf -InputPath $excelPath -Path $excelPdf

New-OfficeMarkdown -Path $markdownPath {
    MarkdownHeading -Level 1 -Text 'Markdown PDF export'
    MarkdownParagraph 'Markdown keeps the same authoring mindset with format-appropriate simplification.'
    MarkdownTable -InputObject $rows
}
$markdownOptions = New-OfficeMarkdownPdfOptions `
    -Title 'Markdown PDF export' `
    -Author 'PSWriteOffice' `
    -CreateOutlineFromHeadings
Export-OfficeDocumentPdf `
    -InputPath $markdownPath `
    -Path $markdownPdf `
    -MarkdownOptions $markdownOptions `
    -PdfWarningVariable markdownWarnings `
    -PdfConversionReportVariable markdownReport

New-OfficePowerPoint -Path $powerPointPath {
    PptSlide {
        PptTitle -Title 'PowerPoint PDF export'
        PptBullets -Bullets 'Create the deck', 'Export the PDF', 'Inspect generated output'
    }
}
Export-OfficeDocumentPdf -InputPath $powerPointPath -Path $powerPointPdf

Get-Item -LiteralPath $wordPdf, $excelPdf, $markdownPdf, $powerPointPdf |
    Select-Object FullName, Length |
    Format-Table -AutoSize
