param([string] $OutputDirectory = (Join-Path $PSScriptRoot 'Documents'))

Import-Module PSWriteOffice -ErrorAction Stop
$null = New-Item -Path $OutputDirectory -ItemType Directory -Force
$path = Join-Path $OutputDirectory 'ExamplePowerPoint1.pptx'

New-OfficePowerPoint -Path $path -Force {
    PowerPointSlide -Title 'Status Update' -Subtitle 'Generated with PSWriteOffice' {
        PowerPointCard 'Current status' -Text 'All service checks passed.'
    }
    PowerPointSlide -Title 'Next Steps' {
        PowerPointList 'Review numbers', 'Plan Q1', 'Ship'
    }
}
Write-Host "Presentation saved to $path"
