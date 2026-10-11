<#
.SYNOPSIS
Creates a local EML review package with a PDF report or supplied attachments.
.DESCRIPTION
PSWriteOffice generates the sample PDF when no attachments are supplied.
Mailozaurr composes and saves the message. Delivery occurs only with -Send.
.EXAMPLE
./Recipe-Mailozaurr-PdfDelivery.ps1 -OutputDirectory ./Output/Mail
.EXAMPLE
./Recipe-Mailozaurr-PdfDelivery.ps1 -AttachmentPath ./Output/SupplierReview/SupplierReview.pdf,./Output/SupplierReview/SupplierOffers.xlsx -OutputDirectory ./Output/Mail
#>
[CmdletBinding()]
param(
    [string] $SmtpServer,
    [string] $From = 'Reporting <reporting@example.invalid>',
    [string[]] $To = @('Review team <review@example.invalid>'),
    [string] $OutputPath,
    [switch] $Send,
    [string] $OutputDirectory = (Join-Path $PSScriptRoot 'Output/Mail'),
    [string[]] $AttachmentPath,
    [string] $Subject = 'Report review package',
    [ValidateRange(1, 65535)][int] $Port = 587,
    [pscredential] $Credential
)

$ErrorActionPreference = 'Stop'
if ($Send -and [string]::IsNullOrWhiteSpace($SmtpServer)) { throw '-Send requires -SmtpServer.' }
if ($PSBoundParameters.ContainsKey('AttachmentPath') -and -not $AttachmentPath) { throw 'Supply at least one attachment path.' }
Import-Module Mailozaurr -MinimumVersion 3.0.2 -ErrorAction Stop
$null = New-Item -ItemType Directory -Path $OutputDirectory -Force
if (-not $AttachmentPath) {
    Import-Module PSWriteOffice -MinimumVersion 3.0.8 -ErrorAction Stop
    if (-not $OutputPath) { $OutputPath = Join-Path $OutputDirectory 'Daily-Service-Report.pdf' }
    $services = @(
        [pscustomobject]@{ Service = 'Directory'; Status = 'Healthy'; Incidents = 0 }
        [pscustomobject]@{ Service = 'Messaging'; Status = 'Attention'; Incidents = 2 }
        [pscustomobject]@{ Service = 'Database'; Status = 'Healthy'; Incidents = 0 }
    )
    New-OfficePdf -Path $OutputPath {
        PdfHeading -Text 'Daily service report' -Color '#17365D'
        PdfSpacer -Height 12
        PdfParagraph -Text 'Synthetic operations snapshot / review before distribution' -Color '#475569'
        PdfSpacer -Height 14
        PdfTable -InputObject $services
        PdfSpacer -Height 14
        PdfParagraph -Text 'Next action: confirm the messaging incident owner and the agreed follow-up time.'
    }
    $AttachmentPath = @($OutputPath)
}
$attachments = @($AttachmentPath | ForEach-Object {
    $file = Get-Item -LiteralPath $_ -ErrorAction Stop
    if ($file.PSIsContainer) { throw "An attachment must be a file: $_" }
    $file.FullName
})
$text = @'
Hello,

Please review the attached report and any editable registers.
Confirm the findings, decision owner and next action before distribution.

This EML is a local review package. Open it in a mail client to inspect its recipients and attachments.
'@
$messageId = '{0}@example.invalid' -f [guid]::NewGuid().ToString('N')
$message = New-MimeMessage -From $From -To $To -Subject $Subject -AttachmentPath $attachments -TextBody $text -Header @{ 'Message-Id' = $messageId }
$path = Join-Path $OutputDirectory 'ReportReview.eml'
$message | Save-MimeMessage -Path $path
Get-Item -LiteralPath $path | Select-Object Name, Length

if ($Send) {
    $mail = @{
        Server = $SmtpServer; Port = $Port; From = $From; To = $To
        Subject = $Subject; Text = $text; Attachment = $attachments
        SecureSocketOptions = 'StartTls'
    }
    if ($Credential) { $mail.Credential = $Credential }
    Send-EmailMessage @mail
}
