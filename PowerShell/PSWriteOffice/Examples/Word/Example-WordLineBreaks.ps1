Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'Example-WordLineBreaks.docx'
$document = New-OfficeWord -Path $path -NoSave

# Add-OfficeWordBreak creates a same-paragraph line break similar to Shift+Enter in Word.
$paragraph = Add-OfficeWordParagraph -Target $document -Text 'Line 1 in the same paragraph' -PassThru
Add-OfficeWordBreak -Paragraph $paragraph
Add-OfficeWordText -Paragraph $paragraph -Text 'Line 2 after the line break'
Add-OfficeWordBreak -Paragraph $paragraph
Add-OfficeWordText -Paragraph $paragraph -Text 'Line 3 still in the same paragraph'

# An empty paragraph creates a visible blank line.
Add-OfficeWordParagraph -Target $document
Add-OfficeWordParagraph -Target $document -Text 'This text comes after an empty paragraph break.'

Close-OfficeWord -Document $document -Save

Write-Host "Document saved to $path"
