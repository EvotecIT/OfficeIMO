Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$services = @(
    [PSCustomObject]@{
        Name        = 'Directory API'
        Status      = 'Healthy'
        Owner       = 'Team Identity'
        LastPatch   = [datetime]'2026-03-28'
        NodeCount   = 4
    }
    [PSCustomObject]@{
        Name        = 'Billing API'
        Status      = 'Review'
        Owner       = 'Team Finance'
        LastPatch   = [datetime]'2026-03-14'
        NodeCount   = 2
    }
)

$tableData = $services | Select-Object Name, Status,
@{
    Name = 'Owner'
    Expression = { $_.Owner }
},
@{
    Name = 'LastPatch'
    Expression = { $_.LastPatch.ToString('yyyy-MM-dd') }
},
@{
    Name = 'NeedsAttention'
    Expression = { if ($_.Status -eq 'Healthy') { 'No' } else { 'Yes' } }
},
@{
    Name = 'ClusterSize'
    Expression = { "$($_.NodeCount) nodes" }
}

$docPath = Join-Path $documents 'Word-CalculatedColumns.docx'

New-OfficeWord -Path $docPath {
    Add-OfficeWordParagraph -Text 'Calculated and projected columns'
    Add-OfficeWordParagraph -Text 'Shape your objects before Add-OfficeWordTable when you want extra columns or friendlier labels.'
    Add-OfficeWordTable -InputObject $tableData -Style 'GridTable1LightAccent1'
}

Write-Host "Document saved to $docPath"
