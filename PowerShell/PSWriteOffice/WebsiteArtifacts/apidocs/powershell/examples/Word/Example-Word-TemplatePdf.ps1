param(
    [Parameter(Mandatory)]
    [string] $InputPath,

    [Parameter(Mandatory)]
    [string] $OutputPath
)

Import-Module PSWriteOffice -ErrorAction Stop

# Use this policy for trusted templates whose requested fonts are installed on the host.
$options = New-OfficeWordPdfOptions -AllowSystemFontEmbedding -AllowDocumentFontEmbedding -IncludePageNumbers:$false
Export-OfficeDocumentPdf -InputPath $InputPath -Path $OutputPath -WordOptions $options -PdfConversionReportVariable report
$report.Warnings | Format-Table Code, Message
