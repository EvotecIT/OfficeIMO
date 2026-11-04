Import-Module PSWriteOffice -ErrorAction Stop
$markdown = @(
    [PSCustomObject]@{ Name = 'Alpha'; Value = 1 }
    [PSCustomObject]@{ Name = 'Beta'; Value = 2 }
) | ConvertTo-OfficeMarkdown

$html = Convert-OfficeDocument -To Html -From Markdown -Text $markdown -Options (New-OfficeMarkdownHtmlOptions -DocumentMode) -AsString

Write-Host "Markdown:"
Write-Host $markdown
Write-Host "HTML:"
Write-Host $html
