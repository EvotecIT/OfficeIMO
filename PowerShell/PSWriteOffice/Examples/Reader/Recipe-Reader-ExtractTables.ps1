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

Read-OfficeDocument -As Tables -Path $path
Export-OfficeDocumentAsset -Type Table -Path $path -OutputDirectory $sidecars
Read-OfficeDocument -As Chunks -Path $path
