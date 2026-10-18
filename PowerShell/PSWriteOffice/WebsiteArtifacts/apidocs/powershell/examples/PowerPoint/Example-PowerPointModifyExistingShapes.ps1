Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'PowerPoint-ModifyExistingShapes.pptx'

$initialRows = @(
    [PSCustomObject]@{ Metric = 'Risk'; State = 'Open' }
    [PSCustomObject]@{ Metric = 'Quality'; State = 'Watching' }
)

$presentation = New-OfficePowerPoint -Path $path -NoSave
try {
    $slide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 -PassThru
    Set-OfficePowerPointSlideTitle -Slide $slide -Title 'Release readiness'
    Add-OfficePowerPointTextBox -Slide $slide -Text 'Status marker: Draft release' -X 70 -Y 110 -Width 420 -Height 45
    Add-OfficePowerPointTable -Slide $slide -InputObject $initialRows -X 70 -Y 180 -Width 500 -Height 170
} finally {
    Close-OfficePowerPoint -Presentation $presentation -Save
}

# Second pass: find existing shapes, then modify their content directly.
$deck = Get-OfficePowerPoint -Path $path
try {
    Find-OfficePowerPointShape -Presentation $deck -Text 'Status marker' -Kind TextBox |
        Set-OfficePowerPointShapeText -Text 'Status marker: Ready for launch'

    $readinessTable = Find-OfficePowerPointShape -Presentation $deck -Text 'Risk' -Kind Table | Select-Object -First 1

    $readinessTable |
        Add-OfficePowerPointTableRow -Values 'Latency', 'Investigating'

    $readinessTable |
        Add-OfficePowerPointTableRow -Values ([ordered]@{
            Metric = 'Documentation'
            State  = 'Ready'
        })

    $readinessTable |
        Set-OfficePowerPointTableCell -Row 1 -Column 1 -Text 'Mitigating'
} finally {
    Close-OfficePowerPoint -Presentation $deck -Save
}

Write-Host "Updated PowerPoint deck saved to $path"
Write-Host 'Matching shapes:'
$reloaded = Get-OfficePowerPoint -Path $path
try {
    Find-OfficePowerPointShape -Presentation $reloaded -Text 'Ready' |
        Select-Object SlideIndex, ShapeIndex, Kind, Text |
        Format-Table
} finally {
    Close-OfficePowerPoint -Presentation $reloaded
}
