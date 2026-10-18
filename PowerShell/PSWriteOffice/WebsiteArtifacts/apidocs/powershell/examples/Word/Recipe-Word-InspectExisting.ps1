$path = '.\Word-Inspection-Source.docx'
$services = @(
    [pscustomobject]@{ Service = 'Identity'; Owner = 'IAM'; Status = 'Ready' }
    [pscustomobject]@{ Service = 'Messaging'; Owner = 'Collaboration'; Status = 'Review' }
)

WordNew -Path $path {
    WordSection {
        WordParagraph -Text 'Service Readiness' -Style Heading1
        WordParagraph -Text 'Review the Messaging service before publication.'
        WordTable -InputObject $services -Style TableGrid
    }
}

$document = Get-OfficeWord -Path $path -ReadOnly
Get-OfficeWordStatistics -Document $document
$document | Get-OfficeWordParagraph | Select-Object Index, Text
$document | Get-OfficeWordTable | Select-Object Index, RowCount, ColumnCount
Find-OfficeWord -Document $document -Text 'Messaging'
Close-OfficeWord -Document $document
