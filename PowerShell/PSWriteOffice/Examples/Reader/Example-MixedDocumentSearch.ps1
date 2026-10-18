param(
    [Parameter(Position = 0)]
    [string[]] $Path = @('.\Evidence'),

    [Parameter(Position = 1)]
    [string] $Query = 'retention period'
)

# Search-OfficeDocument discovers every registered format when -Extension is omitted.
# The same run can therefore cover Word, Excel, PowerPoint, PDF, PST, OST, EML,
# Markdown, OpenDocument, RTF, Visio, archives, and the other Reader adapters.
$readErrors = @()
$matches = @(Search-OfficeDocument `
        -Path $Path `
        -Recurse `
        -Query $Query `
        -MaxDocuments 5000 `
        -MaxStoreItems 25000 `
        -MaximumResults 100 `
        -MaxDegreeOfParallelism 4 `
        -IncludePageLocations `
        -ErrorVariable +readErrors `
        -ErrorAction SilentlyContinue)

$matches |
    Sort-Object Path, Location, StartIndex |
    Select-Object Path, DocumentType, Match, Location, Pages,
        DocumentLimitReached, SourceLimitReached, SearchLimitReached |
    Format-Table -AutoSize

$matches |
    Group-Object DocumentType |
    Sort-Object Name |
    Select-Object Name, Count |
    Format-Table -AutoSize

if ($readErrors.Count -gt 0) {
    Write-Warning "$($readErrors.Count) input file(s) could not be read. Other files were still searched."
    $readErrors | ForEach-Object {
        [pscustomobject]@{
            Path    = $_.TargetObject
            Message = $_.Exception.Message
        }
    } | Format-Table -AutoSize
}

# For an intentionally unbounded run, replace the three numeric limits above with:
# -NoDocumentLimit -AllStoreItems -AllResults
