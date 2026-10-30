[CmdletBinding()]
param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot 'Output')
)

$ErrorActionPreference = 'Stop'

Import-Module PSWriteOffice -ErrorAction Stop

$documents = $OutputDirectory
$null = New-Item -Path $documents -ItemType Directory -Force
$wordPath = Join-Path $documents 'Showcase-RichTextRuns.docx'
$excelPath = Join-Path $documents 'Showcase-RichTextRuns.xlsx'
$pdfPath = Join-Path $documents 'Showcase-RichTextRuns.pdf'
$pptPath = Join-Path $documents 'Showcase-RichTextRuns.pptx'

$runs = @(
    TextRun 'Status: '
    TextRun 'Ready' -Color SeaGreen -Bold
    TextRun ' with '
    TextRun 'named colors' -Color Navy -Italic
    TextRun ' and '
    TextRun 'underline' -UnderlineStyle Dotted -Color DarkSlateBlue
)

$serviceRows = @(
    [pscustomobject]@{ Service = 'Identity Sync'; Status = 'Ready'; Owner = 'Platform' }
    [pscustomobject]@{ Service = 'Backup'; Status = 'Watch'; Owner = 'Operations' }
    [pscustomobject]@{ Service = 'Remote Access'; Status = 'Needs action'; Owner = 'Security' }
)

WordNew -Path $wordPath {
    WordSection {
        WordParagraph -Text 'Rich text runs' -Style Heading1
        WordParagraph -Run $runs
        WordTable -Style TableGrid -InputObject @(
            , @(
                WordTableCellSpec -Run @(
                    WordTextRun 'Service '
                    WordTextRun 'readiness' -Color SeaGreen -Bold
                ) -ColumnSpan 3 -FillColor AliceBlue -Align Center
            )
            , @('Service', 'Status', 'Owner')
            , @((WordTableCellSpec -Run @(WordTextRun 'Identity Sync' -Bold; WordTextRun ' 99.98%' -Color SeaGreen)), 'Ready', 'Platform')
            , @('Backup', (WordTableCellSpec -Run @(WordTextRun 'Watch' -Color DarkOrange -Bold)), 'Operations')
        )
    }
}

ExcelNew -Path $excelPath {
    ExcelSheet -Name 'Summary' -Content {
        ExcelRichText -Address A1 -Run @(
            ExcelTextRun 'Status: '
            ExcelTextRun 'Ready' -Color SeaGreen -Bold
            ExcelTextRun ' for review' -Italic
        )
        ExcelRichText -Address A3 -Run @(
            ExcelTextRun 'Owner: '
            ExcelTextRun 'Platform' -Color Navy -Bold
        )
        ExcelTable -Data $serviceRows -Name 'ServiceReadiness' -StartRow 5 -StartColumn 1
        ExcelAutoFit
    }
}

PdfNew -Path $pdfPath {
    PdfHeading 'Rich text runs'
    PdfText -Run @(
        PdfTextRun 'Status: '
        PdfTextRun 'Ready' -Color SeaGreen -Bold
        PdfTextRun ' with named colors and '
        PdfTextRun 'inline emphasis' -Color Navy -Italic
    )
    PdfTable -HeaderRowCount 1 -InputObject @(
        , @(
            PdfTableCell -Run @(
                PdfTextRun 'Service '
                PdfTextRun 'readiness' -Color SeaGreen -Bold
            ) -ColumnSpan 3 -FillColor AliceBlue -Align Center
        )
        , @('Service', 'Status', 'Owner')
        , @((PdfTableCell -Run @(PdfTextRun 'Identity Sync' -Bold; PdfTextRun ' 99.98%' -Color SeaGreen)), 'Ready', 'Platform')
        , @('Backup', (PdfTableCell -Run @(PdfTextRun 'Watch' -Color DarkOrange -Bold)), 'Operations')
    )
}

PptNew -Path $pptPath {
    PptSlideSize -Preset Screen16x9
    PptSlide {
        PptTitle -Title 'Rich text runs'
        PptTextBox -Run @(
            PptTextRun 'Status: '
            PptTextRun 'Ready' -Color SeaGreen -Bold
            PptTextRun ' with named colors' -Color Navy -Italic
        ) -X 70 -Y 115 -Width 560 -Height 54
        PptTable -InputObject @(
            , @(
                @{
                    Run = @(
                        PptTextRun 'Service '
                        PptTextRun 'readiness' -Color SeaGreen -Bold
                    )
                    ColumnSpan = 3
                    FillColor = 'AliceBlue'
                    Align = 'Center'
                }
            )
            , @('Service', 'Status', 'Owner')
            , @(
                @{ Run = @(PptTextRun 'Identity Sync' -Bold; PptTextRun ' 99.98%' -Color SeaGreen) },
                'Ready',
                'Platform'
            )
            , @('Backup', @{ Run = @(PptTextRun 'Watch' -Color DarkOrange -Bold) }, 'Operations')
        ) -X 70 -Y 190 -Width 560 -Height 160
    }
}

Write-Host "Word document saved to $wordPath"
Write-Host "Excel workbook saved to $excelPath"
Write-Host "PDF saved to $pdfPath"
Write-Host "PowerPoint deck saved to $pptPath"
