Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$data = @(
    [PSCustomObject]@{ Name = 'Alpha'; Score = 92; Owner = 'Ada' }
    [PSCustomObject]@{ Name = 'Beta'; Score = 76; Owner = 'Linus' }
    [PSCustomObject]@{ Name = 'Gamma'; Score = 64; Owner = 'Grace' }
)

$docPath = Join-Path $documents 'Word-TableConditions.docx'
New-OfficeWord -Path $docPath {
    WordSection {
        WordParagraph 'Quality Matrix' -Style Heading1
        WordParagraph 'Rows are highlighted based on score thresholds.'
        WordTable -Data $data -Style TableGrid -Layout Autofit {
            WordTableCondition -FilterScript { $_.Score -ge 90 } -BackgroundColor '#e6fffb'
            WordTableCondition -FilterScript { $_.Score -lt 70 } -BackgroundColor '#ffe6e6'
        }
    }
}

Write-Host "Document saved to $docPath"