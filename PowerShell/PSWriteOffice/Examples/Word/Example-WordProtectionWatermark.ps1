Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'Word-ProtectedWatermark.docx'

New-OfficeWord -Path $path {
    Add-OfficeWordParagraph -Text 'Confidential report'
    Add-OfficeWordWatermark -Text 'CONFIDENTIAL'
    Protect-OfficeDocument -Restriction ReadOnly -RestrictionPassword 'secret'
}

Write-Host "Document saved to $path"
