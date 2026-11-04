$path = '.\Word-Inspection-Source.docx'
$services = @(
    [pscustomobject]@{ Service = 'Identity'; Owner = 'IAM'; Status = 'Ready' }
    [pscustomobject]@{ Service = 'Messaging'; Owner = 'Collaboration'; Status = 'Review' }
)

New-OfficeWord -Path $path {
    WordSection {
        WordParagraph -Text 'Service Readiness' -Style Heading1
        WordParagraph -Text 'Review the Messaging service before publication.'
        WordTable -InputObject $services -Style TableGrid
    }
}

$document = Open-OfficeWord -Path $path -ReadOnly
(Get-OfficeDocumentInfo -Document $document -Detailed).Details.Statistics
$document | Get-OfficeWordParagraph | Select-Object Index, Text
$document | Get-OfficeWordTable | Select-Object Index, RowCount, ColumnCount
Get-OfficeWordParagraph -Document $document -Pattern 'Messaging' -SimpleMatch
Close-OfficeWord -Document $document
