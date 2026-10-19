$source = '.\Reader-Table-Source.docx'
$output = '.\Reader-Table-Exports'
$rows = @(
    [pscustomobject]@{ Service = 'Identity'; Owner = 'Platform' }
    [pscustomobject]@{ Service = 'Messaging'; Owner = 'Collaboration' }
)

WordNew -Path $source {
    WordParagraph -Text 'Service ownership' -Style Heading1
    WordTable -InputObject $rows -Style TableGrid
}

Get-OfficeDocumentTable -Path $source -AsExport -OutputDirectory $output
