Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$docPath = Join-Path $documents 'Word-AliasDsl.docx'
if (-not (Test-Path $docPath)) {
    Write-Warning "Expected document '$docPath' not found. Run Example-WordAliasDsl.ps1 first."
    return
}

$document = Get-OfficeWord -Path $docPath -ReadOnly
try {
    $sections = $document | Get-OfficeWordSection
    Write-Host "Sections:" $sections.Count

    $paragraphs = $document | Get-OfficeWordParagraph
    Write-Host 'First paragraphs:'
    $paragraphs | Select-Object -First 3 | ForEach-Object {
        Write-Host " -" $_.Text
    }

    $tables = $document | Get-OfficeWordTable
    Write-Host "Tables:" $tables.Count

    Write-Host 'First text segments:'
    $paragraphs | Select-Object -First 1 | Get-OfficeWordText | Select-Object -First 3 | ForEach-Object {
        Write-Host "  -" $_.Text
    }
} finally {
    Close-OfficeWord -Document $document
}
