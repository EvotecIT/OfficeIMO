$path = '.\Markdown-Object-Composition.md'
$services = @(
    [pscustomobject]@{ Service = 'Identity'; Status = 'Healthy' }
    [pscustomobject]@{ Service = 'Messaging'; Status = 'Watch' }
)

$document = New-OfficeMarkdown -Path $path -NoSave
Add-OfficeMarkdownHeading -Document $document -Level 1 -Text 'Service status'
Add-OfficeMarkdownParagraph -Document $document -Text 'This file was composed through an explicit Markdown document object.'
$document | Add-OfficeMarkdownTable -InputObject $services
$document | Save-OfficeMarkdown -Path $path
