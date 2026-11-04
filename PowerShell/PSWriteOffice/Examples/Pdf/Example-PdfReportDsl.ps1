$ErrorActionPreference = 'Stop'

Import-Module PSWriteOffice -ErrorAction Stop

$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'Example-PdfReportDsl.pdf'
$attachmentPath = Join-Path $documents 'Example-PdfReportDsl-notes.txt'
Set-Content -LiteralPath $attachmentPath -Value 'Synthetic report notes embedded by PSWriteOffice.' -Encoding UTF8

$services = @(
    [pscustomobject]@{ Service = 'Identity Sync'; Owner = 'Platform'; Status = 'Healthy'; Availability = '99.98%'; Risk = 'Low' }
    [pscustomobject]@{ Service = 'Messaging'; Owner = 'Collaboration'; Status = 'Watch'; Availability = '99.72%'; Risk = 'Medium' }
    [pscustomobject]@{ Service = 'Remote Access'; Owner = 'Security'; Status = 'Needs action'; Availability = '98.84%'; Risk = 'High' }
)

New-OfficePdf -Path $path -Theme Modern {
    PdfMetadata -Title 'PDF Service Review' -Author 'PSWriteOffice' -Subject 'Generated PDF report'
    PdfPageSetup -PageSize A4 -Margin 42
    PdfBackground -Color '#FFFFFF'
    PdfBackgroundShape -Shape TopBand -Height 92 -FillColor '#DBEAFE' -FillOpacity 0.8
    PdfBackgroundShape -Shape Ellipse -X 420 -Y 670 -Width 96 -Height 70 -FillColor '#99F6E4' -FillOpacity 0.35
    PdfPageBorder -Color '#0F766E' -Width 1.2 -Inset 24 -Opacity 0.75
    PdfHeader 'PSWriteOffice PDF report'
    PdfFooter 'Page {page}/{pages}'

    PdfBookmark 'summary'
    PdfHeading 'PDF Service Review'
    PdfParagraph {
        PdfText 'This report uses '
        PdfText 'themes' -Bold -Color '#0F766E'
        PdfText ', '
        PdfText 'rich inline text' -Italic -BackgroundColor '#FEF3C7'
        PdfText ', layout rows, links, and attachments from PSWriteOffice.'
    }
    PdfPanel 'The PDF cmdlets are thin adapters over OfficeIMO.Pdf, so reusable document behavior stays in the engine.'
    PdfHr -Color '#0F766E' -Thickness 1.5 -SpacingBefore 10 -SpacingAfter 12

    PdfRow -Gap 18 -ColumnSeparatorColor '#CBD5E1' {
        PdfColumn -Weight 38 {
            PdfHeading 'Signals' -Level 2 -Color '#0F766E'
            PdfParagraph {
                PdfText 'External reference: '
                PdfText 'Evotec' -Link 'https://evotec.xyz' -Color '#2563EB'
            }
            PdfList -Items 'Healthy services are stable', 'Watch items need owner follow-up', 'High risk gets a focused action' -Numbered
        }
        PdfColumn -Weight 62 {
            PdfHeading 'Service Scorecard' -Level 2 -Color '#1D4ED8'
            PdfTable -InputObject $services
            PdfParagraph {
                PdfText 'Jump to '
                PdfText 'summary' -LinkBookmark 'summary' -Color '#7C3AED'
                PdfText ' when reviewing the scorecard.'
            }
        }
    }

    PdfSpacer 10
    PdfAttachment -Path $attachmentPath -Description 'Generated example notes'
}

$info = (Get-OfficeDocumentInfo -Detailed -Path $path).Details
[pscustomobject]@{
    Path            = $path
    Pages           = $info.PageCount
    HasLinks        = $info.HasLinkAnnotations
    Attachments     = $info.AttachmentCount
    NamedDestinations = ($info.NamedDestinationNames -join ', ')
} | Format-List
