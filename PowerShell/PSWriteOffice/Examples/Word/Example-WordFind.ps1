Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'Word-Find.docx'

New-OfficeWord -Path $path {
    Add-OfficeWordParagraph -Text 'Hello from PSWriteOffice'
}

$doc = Get-OfficeWord -Path $path
try {
    $paragraph = Add-OfficeWordParagraph -Target $doc -Text 'Page' -PassThru
    Add-OfficeWordBookmark -Paragraph $paragraph -Name 'Bookmark1'
    Add-OfficeWordField -Paragraph $paragraph -Type Page
} finally {
    Close-OfficeWord -Document $doc -Save
}

Find-OfficeWord -Path $path -Text 'Hello'
Get-OfficeWordBookmark -Path $path
Get-OfficeWordField -Path $path -FieldType Page
