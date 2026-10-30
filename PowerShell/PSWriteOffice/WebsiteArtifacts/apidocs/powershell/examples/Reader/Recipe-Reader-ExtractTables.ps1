$path = '.\Reader-Table-Source.md'
$sidecars = '.\Reader-Table-Sidecars'
$scores = @(
    [pscustomobject]@{ Service = 'Identity'; Score = 98 }
    [pscustomobject]@{ Service = 'Messaging'; Score = 94 }
)

MarkdownNew -Path $path {
    MarkdownHeading -Level 1 -Text 'Service scores'
    MarkdownTable -InputObject $scores
}

Get-OfficeDocumentTable -Path $path
Get-OfficeDocumentTable -Path $path -OutputDirectory $sidecars
Get-OfficeDocumentChunk -Path $path
