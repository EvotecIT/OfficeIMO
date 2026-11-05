Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$docPath = Join-Path $documents 'Word-HtmlSource.docx'
$htmlPath = Join-Path $documents 'Word-HtmlSource.html'
$roundtripPath = Join-Path $documents 'Word-HtmlRoundtrip.docx'

New-OfficeWord -Path $docPath {
    Add-OfficeWordSection {
        Add-OfficeWordParagraph -Text 'Hello from HTML conversion.' -Style Heading2
        Add-OfficeWordParagraph -Text 'This document will round-trip to HTML.'
    }
}

Convert-OfficeDocument -To Html -Path $docPath -OutputPath $htmlPath
Convert-OfficeDocument -To Word -Path $htmlPath -OutputPath $roundtripPath

Write-Host "HTML saved to $htmlPath"
Write-Host "Round-trip document saved to $roundtripPath"
