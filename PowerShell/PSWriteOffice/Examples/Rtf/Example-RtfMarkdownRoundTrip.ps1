Import-Module PSWriteOffice -ErrorAction Stop

$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$markdownPath = Join-Path $documents 'Rtf-MarkdownRoundTrip.md'
$rtfPath = Join-Path $documents 'Rtf-MarkdownRoundTrip.rtf'
$roundTripMarkdownPath = Join-Path $documents 'Rtf-MarkdownRoundTrip.from-rtf.md'

@'
# Service Note

The weekly service review is ready.

- Identity is healthy.
- Messaging needs follow-up.
- Reporting published the updated dashboard.

| Area | Owner |
| --- | --- |
| Identity | Platform |
| Reporting | Analytics |
'@ | Set-Content -Path $markdownPath -Encoding UTF8

ConvertTo-OfficeRtf -MarkdownPath $markdownPath -OutputPath $rtfPath
ConvertFrom-OfficeRtf -Path $rtfPath -As Markdown -OutputPath $roundTripMarkdownPath

Write-Host "Markdown saved to $markdownPath"
Write-Host "RTF saved to $rtfPath"
Write-Host "Round-tripped Markdown saved to $roundTripMarkdownPath"
