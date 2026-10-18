param(
    [Parameter(Mandatory)]
    [string] $AzureTableConnectionString,

    [Parameter(Mandatory)]
    [string] $AzureTableName,

    [Parameter(Mandatory)]
    [string] $ConfluenceSpaceId,

    [Parameter(Mandatory)]
    [string] $ConfluenceTitle,

    [string] $AzureTableFilter,

    [object] $ConfluenceSession,

    [switch] $Publish
)

Import-Module DbaClientX -ErrorAction Stop
Import-Module PSWriteOffice -ErrorAction Stop

if ($null -ne $ConfluenceSession -and $ConfluenceSession -isnot [OfficeIMO.Confluence.ConfluenceSession]) {
    throw '-ConfluenceSession must be an OfficeIMO.Confluence.ConfluenceSession created by New-OfficeConfluenceSession.'
}

$entities = @(
    Get-DbaXAzureTableEntity `
        -ConnectionString $AzureTableConnectionString `
        -TableName $AzureTableName `
        -Filter $AzureTableFilter
)

if ($entities.Count -eq 0) {
    $markdown = "# $ConfluenceTitle`n`n_No Azure Table entities matched the query._"
} else {
    $propertyNames = @(
        $entities |
            ForEach-Object { $_.Properties.Keys } |
            Sort-Object -Unique
    )
    $columns = @('PartitionKey', 'RowKey') + $propertyNames

    function ConvertTo-MarkdownCell {
        param([AllowNull()] $Value)

        if ($null -eq $Value) {
            return ''
        }

        return $Value.ToString().Replace('|', '\|').Replace("`r", ' ').Replace("`n", ' ')
    }

    $lines = [Collections.Generic.List[string]]::new()
    $lines.Add("# $ConfluenceTitle")
    $lines.Add('')
    $lines.Add('| ' + (($columns | ForEach-Object { ConvertTo-MarkdownCell $_ }) -join ' | ') + ' |')
    $lines.Add('| ' + (($columns | ForEach-Object { '---' }) -join ' | ') + ' |')
    foreach ($entity in $entities) {
        $values = foreach ($column in $columns) {
            if ($column -eq 'PartitionKey') {
                $entity.PartitionKey
            } elseif ($column -eq 'RowKey') {
                $entity.RowKey
            } else {
                $entity.Properties[$column]
            }
        }
        $lines.Add('| ' + (($values | ForEach-Object { ConvertTo-MarkdownCell $_ }) -join ' | ') + ' |')
    }

    $markdown = $lines -join [Environment]::NewLine
}

$publishParameters = @{
    SpaceId     = $ConfluenceSpaceId
    Title       = $ConfluenceTitle
    Content     = $markdown
    FailOnLoss  = $true
}

if ($Publish) {
    if ($null -eq $ConfluenceSession) {
        throw 'Provide -ConfluenceSession when -Publish is used.'
    }

    $publishParameters.Session = $ConfluenceSession
} else {
    $publishParameters.PlanOnly = $true
}

Publish-OfficeConfluencePage @publishParameters
