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

$serviceRows = @(
    [pscustomobject]@{ Service = 'Identity Sync'; Status = 'Ready'; Owner = 'Platform' }
    [pscustomobject]@{ Service = 'Backup'; Status = 'Watch'; Owner = 'Operations' }
    [pscustomobject]@{ Service = 'Remote Access'; Status = 'Needs action'; Owner = 'Security' }
)

New-OfficeWord -Path $wordPath {
    WordSection {
        WordParagraph -Text 'Rich text runs' -Style Heading1
        WordParagraph {
            WordText 'Status: '
            WordText 'Ready' -Color SeaGreen -Bold
            WordText ' with '
            WordText 'named colors' -Color Navy -Italic
            WordText ' and '
            WordText 'underline' -Underline Dotted -Color DarkSlateBlue
        }
        WordTable -Style TableGrid -InputObject @(
            , @(
                New-OfficeWordTableCell -ColumnSpan 3 -BackgroundColor AliceBlue -Alignment Center -Content {
                    WordText 'Service '
                    WordText 'readiness' -Color SeaGreen -Bold
                }
            )
            , @('Service', 'Status', 'Owner')
            , @((New-OfficeWordTableCell -Content { WordText 'Identity Sync' -Bold; WordText ' 99.98%' -Color SeaGreen }), 'Ready', 'Platform')
            , @('Backup', (New-OfficeWordTableCell -Text 'Watch' -Color DarkOrange -Bold), 'Operations')
        )
    }
}

New-OfficeExcel -Path $excelPath {
    ExcelSheet -Name 'Summary' -Content {
        ExcelCell -Address A1 {
            ExcelText 'Status: '
            ExcelText 'Ready' -Color SeaGreen -Bold
            ExcelText ' for review' -Italic
        }
        ExcelCell -Address A3 {
            ExcelText 'Owner: '
            ExcelText 'Platform' -Color Navy -Bold
        }
        ExcelTable -InputObject $serviceRows -Name 'ServiceReadiness' -Address A5
        ExcelAutoFit
    }
}

New-OfficePdf -Path $pdfPath {
    PdfHeading 'Rich text runs'
    PdfParagraph {
        PdfText 'Status: '
        PdfText 'Ready' -Color SeaGreen -Bold
        PdfText ' with named colors and '
        PdfText 'inline emphasis' -Color Navy -Italic
    }
    PdfTable -HeaderRowCount 1 -InputObject @(
        , @(
            PdfTableCell -ColumnSpan 3 -BackgroundColor AliceBlue -Alignment Center -Content {
                PdfText 'Service '
                PdfText 'readiness' -Color SeaGreen -Bold
            }
        )
        , @('Service', 'Status', 'Owner')
        , @((PdfTableCell -Content { PdfText 'Identity Sync' -Bold; PdfText ' 99.98%' -Color SeaGreen }), 'Ready', 'Platform')
        , @('Backup', (PdfTableCell -Text 'Watch' -Color DarkOrange -Bold), 'Operations')
    )
}

New-OfficePowerPoint -Path $pptPath {
    Set-OfficePowerPointSlideSize -Preset Screen16x9
    PowerPointSlide -Title 'Rich text runs' {
        PowerPointTextBox -Content {
            PowerPointText 'Status: '
            PowerPointText 'Ready' -Color SeaGreen -Bold
            PowerPointText ' with named colors' -Color Navy -Italic
        } -X 70 -Y 115 -Width 560 -Height 54
        PowerPointTable -InputObject @(
            , @(
                @{
                    Run = @(
                        @{ Text = 'Service ' }
                        @{ Text = 'readiness'; Color = 'SeaGreen'; Bold = $true }
                    )
                    ColumnSpan = 3
                    FillColor = 'AliceBlue'
                    Align = 'Center'
                }
            )
            , @('Service', 'Status', 'Owner')
            , @(
                @{ Run = @(@{ Text = 'Identity Sync'; Bold = $true }; @{ Text = ' 99.98%'; Color = 'SeaGreen' }) },
                'Ready',
                'Platform'
            )
            , @('Backup', @{ Run = @(@{ Text = 'Watch'; Color = 'DarkOrange'; Bold = $true }) }, 'Operations')
        ) -X 70 -Y 190 -Width 560 -Height 160
    }
}

Write-Host "Word document saved to $wordPath"
Write-Host "Excel workbook saved to $excelPath"
Write-Host "PDF saved to $pdfPath"
Write-Host "PowerPoint deck saved to $pptPath"
