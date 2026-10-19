Import-Module PSWriteOffice -ErrorAction Stop
$path = Join-Path $PSScriptRoot 'Example-WordTableCells.docx'
$imagePath = Join-Path $PSScriptRoot 'Example-WordTableCells.png'
$fixturePath = Join-Path $PSScriptRoot 'Example-WordTableCells.fixture.png'
$rows = @(
    [pscustomobject]@{
        Topic   = 'Release readiness'
        Details = 'Pending'
    }
)

$nestedRows = @(
    [pscustomobject]@{ Step = 'Validate'; State = 'Ready' }
    [pscustomobject]@{ Step = 'Release'; State = 'Queued' }
)

Copy-Item -LiteralPath $fixturePath -Destination $imagePath -Force

New-OfficeWord -Path $path {
    WordTable -Data $rows -Style GridTable1LightAccent1 {
        WordTableCell -Row 1 -Column 0 {
            WordParagraph {
                WordText 'Checklist'
            }

            WordImage -Path $imagePath -Width 36 -Height 36 -Description 'Status icon'

            WordList {
                WordListItem 'Confirm issue coverage'
                WordListItem 'Stage release notes'
            }
        }

        WordTableCell -Row 1 -Column 1 {
            WordTable -Data $nestedRows -Style TableGrid
        }
    }
}

Write-Host "Document saved to $path"
Write-Host "Image fixture saved to $imagePath"
