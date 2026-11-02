param([string] $OutputDirectory = (Join-Path $PSScriptRoot 'Documents'))

Import-Module PSWriteOffice -ErrorAction Stop
$null = New-Item -Path $OutputDirectory -ItemType Directory -Force
$path = Join-Path $OutputDirectory 'DslExample.pptx'
$data = @(
    [pscustomobject]@{ Item = 'Alpha'; Qty = 10 }
    [pscustomobject]@{ Item = 'Beta'; Qty = 20 }
)

New-OfficePowerPoint -Path $path -Force {
    PowerPointSlide -Title 'Status Update' -Subtitle 'Generated with PSWriteOffice' {
        PowerPointList 'Wins', 'Risks', 'Next Steps'
        PowerPointNotes 'Keep this under five minutes.'
    }
    PowerPointSlide -Title 'Inventory' {
        PowerPointTable -InputObject $data
    }
}
Write-Host "Presentation saved to $path"
