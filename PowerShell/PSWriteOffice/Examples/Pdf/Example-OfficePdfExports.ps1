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
Convert-OfficeDocument -Path $wordPath -To Pdf -OutputPath $wordPdf

New-OfficeExcel -Path $excelPath {
    ExcelSheet -Name 'Summary' {
        ExcelTable -InputObject $rows
        ExcelAutoFit
    }
}
Convert-OfficeDocument -Path $excelPath -To Pdf -OutputPath $excelPdf

New-OfficeMarkdown -Path $markdownPath {
    MarkdownHeading -Level 1 -Text 'Markdown PDF export'
    MarkdownParagraph 'Markdown keeps the same authoring mindset with format-appropriate simplification.'
    MarkdownTable -InputObject $rows
}
$markdownOptions = New-OfficeMarkdownPdfOptions `
    -Title 'Markdown PDF export' `
    -Author 'PSWriteOffice'
$markdownReport = Convert-OfficeDocument `
    -Path $markdownPath -To Pdf `
    -OutputPath $markdownPdf `
    -Options $markdownOptions
$markdownReport.Diagnostics | Format-Table Severity, Code, Message

New-OfficePowerPoint -Path $powerPointPath {
    PowerPointSlide -Title 'PowerPoint PDF export' {
        PowerPointList 'Create the deck', 'Export the PDF', 'Inspect generated output'
    }
}
Convert-OfficeDocument -Path $powerPointPath -To Pdf -OutputPath $powerPointPdf

Get-Item -LiteralPath $wordPdf, $excelPdf, $markdownPdf, $powerPointPdf |
    Select-Object FullName, Length |
    Format-Table -AutoSize
