Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'Word-ModifyExistingObjects.docx'

$initialRisks = @(
    [PSCustomObject]@{ Item = 'Risk marker'; Owner = 'PMO'; State = 'Open' }
    [PSCustomObject]@{ Item = 'Dependency review'; Owner = 'Architecture'; State = 'Watching' }
)

New-OfficeWord -Path $path {
    WordParagraph -Text 'Operational handover'
    WordParagraph -Text 'Existing risk table follows'
    WordTable -Data $initialRisks -Style TableGrid

    WordParagraph -Text 'Existing approval checklist follows'
    WordList {
        WordListItem -Text 'Initial review'
        WordListItem -Text 'Security approval'
    }
}

# Second pass: treat the file as an existing document that came from a user or template.
$document = Open-OfficeWord -Path $path
try {
    $riskTable = Get-OfficeWordTable -Document $document -Pattern 'Risk marker' -SimpleMatch | Select-Object -First 1

    $riskTable |
        Add-OfficeWordTableRow -InputObject ([ordered]@{
            Item  = 'Mitigation plan'
            Owner = 'Service Desk'
            State = 'Ready'
        }) -PassThru

    $riskTable |
        Add-OfficeWordTableRow -InputObject 'Release communication', 'Operations', 'Draft'

    $riskTable |
        Get-OfficeWordTableCell -Row 2 -Column 2 |
        Set-OfficeWordTableCell -Text 'Investigating' -BackgroundColor '#fff2cc' -ShadingPattern Clear

    Get-OfficeWordList -Document $document -Pattern 'Initial review' -SimpleMatch |
        Add-OfficeWordListItem -Text 'Business sign-off'

    Get-OfficeWordList -Document $document |
        Where-Object { $_.ListItems.Text -contains 'Initial review' } |
        Add-OfficeWordListItem -Text 'Go-live approval'
} finally {
    Close-OfficeWord -Document $document -Save
}

Write-Host "Updated Word document saved to $path"
Write-Host 'Risk table matches:'
Get-OfficeWordTable -Path $path -Pattern 'Mitigation plan' -SimpleMatch |
    Select-Object -First 1 |
    Format-Table
