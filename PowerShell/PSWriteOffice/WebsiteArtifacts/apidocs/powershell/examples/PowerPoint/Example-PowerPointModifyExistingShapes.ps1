param([string] $OutputDirectory = (Join-Path $PSScriptRoot '../Documents'))

Import-Module PSWriteOffice -ErrorAction Stop
$null = New-Item -Path $OutputDirectory -ItemType Directory -Force
$path = Join-Path $OutputDirectory 'PowerPoint-ModifyExistingShapes.pptx'
$initialRows = @(
    [pscustomobject]@{ Metric = 'Risk'; State = 'Open' }
    [pscustomobject]@{ Metric = 'Quality'; State = 'Watching' }
)

New-OfficePowerPoint -Path $path {
    PowerPointSlide -Title 'Release readiness' {
        PowerPointTextBox 'Status marker: Draft release' -X 70 -Y 110 -Width 420 -Height 45
        PowerPointTable -InputObject $initialRows -X 70 -Y 180 -Width 500 -Height 170
    }
}

Open-OfficePowerPoint -Path $path {
    Get-OfficePowerPointShape -Text 'Status marker' -Kind TextBox |
        Set-OfficePowerPointShapeText -Content {
            PowerPointText 'Status marker: ' -Bold
            PowerPointText 'Ready for launch' -Color Success
        }
    $table = Get-OfficePowerPointShape -Text 'Risk' -Kind Table | Select-Object -First 1
    $table | Add-OfficePowerPointTableRow -Value 'Latency', 'Investigating'
    $table | Set-OfficePowerPointTableCell -Row 2 -Column 2 -Content {
        PowerPointParagraph {
            PowerPointText 'Mitigating' -Color Warning -Link 'https://example.org/risk'
        }
    }
}

$reloaded = Open-OfficePowerPoint -Path $path
try {
    Get-OfficePowerPointShape -Document $reloaded -Text 'Ready' |
        Select-Object SlideIndex, ShapeIndex, Kind, Text
} finally {
    Close-OfficePowerPoint -Document $reloaded
}
