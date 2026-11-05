param(
    [Parameter(Mandatory)]
    [string] $InputPath,

    [Parameter(Mandatory)]
    [string] $OutputPath
)

Import-Module PSWriteOffice -ErrorAction Stop

# Use this policy for trusted templates whose requested fonts are installed on the host.
$options = New-OfficeWordPdfOptions -AllowSystemFontEmbedding -AllowDocumentFontEmbedding -IncludePageNumbers:$false
$report = Convert-OfficeDocument -Path $InputPath -To Pdf -OutputPath $OutputPath -Options $options
$report.Diagnostics | Format-Table Severity, Code, Message
