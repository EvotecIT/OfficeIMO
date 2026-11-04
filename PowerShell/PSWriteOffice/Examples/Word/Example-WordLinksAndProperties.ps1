Import-Module PSWriteOffice -ErrorAction Stop

$Path = Join-Path $PSScriptRoot 'Example-WordLinksAndProperties.docx'

New-OfficeWord -Path $Path {
    Set-OfficeWordDocumentProperty -Name Title -Value 'Links and Properties Demo'
    Set-OfficeWordDocumentProperty -Name Creator -Value 'PSWriteOffice'
    Set-OfficeWordDocumentProperty -Name ReleaseNumber -Value 7 -Custom

    WordParagraph {
        WordText 'Visit '
        WordText -Text 'Example.org' -Link 'https://example.org'  -Tooltip 'External reference'
        WordText ' or jump to the '
        WordText -Text 'Summary' -Anchor 'Summary'
        WordText '.'
    }

    WordParagraph {
        WordText 'Summary section'
        WordBookmark -Name 'Summary'
    }
}

$links = Get-OfficeWordHyperlink -Path $Path
$properties = Get-OfficeWordDocumentProperty -Path $Path

$links | Format-Table Text, Uri, Anchor, Tooltip
$properties | Sort-Object Scope, Name | Format-Table Name, Scope, Value, CustomPropertyType
