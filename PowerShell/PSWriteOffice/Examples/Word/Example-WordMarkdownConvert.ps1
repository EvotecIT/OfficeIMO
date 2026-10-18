Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$docPath = Join-Path $documents 'Word-MarkdownSource.docx'
$markdownPath = Join-Path $documents 'Word-MarkdownSource.md'
$roundtripPath = Join-Path $documents 'Word-MarkdownRoundtrip.docx'

New-OfficeWord -Path $docPath {
    Add-OfficeWordParagraph -Text 'Markdown Bridge' -Style Heading1
    Add-OfficeWordParagraph -Text 'This document will round-trip through Markdown.'
    Add-OfficeWordList -Style Bulleted {
        Add-OfficeWordListItem -Text 'Alpha'
        Add-OfficeWordListItem -Text 'Beta'
    }
}

ConvertTo-OfficeWordMarkdown -Path $docPath -OutputPath $markdownPath
ConvertFrom-OfficeWordMarkdown -Path $markdownPath -OutputPath $roundtripPath

Write-Host "Markdown saved to $markdownPath"
Write-Host "Round-trip document saved to $roundtripPath"
