Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot 'Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'LayoutDslExample.pptx'

New-OfficePowerPoint -Path $path {
    $layout = Get-OfficePowerPointLayout | Select-Object -First 1
    if (-not $layout) {
        return
    }

    Set-OfficePowerPointLayoutPlaceholderBounds -Master $layout.MasterIndex -Layout $layout.LayoutIndex -PlaceholderType Title `
        -Left 60 -Top 40 -Width 600 -Height 120 -CreateIfMissing
    Set-OfficePowerPointLayoutPlaceholderTextMargins -Master $layout.MasterIndex -Layout $layout.LayoutIndex -PlaceholderType Title `
        -Left 8 -Top 6 -Right 8 -Bottom 6 -CreateIfMissing
    Set-OfficePowerPointLayoutPlaceholderTextStyle -Master $layout.MasterIndex -Layout $layout.LayoutIndex -PlaceholderType Title `
        -Style Title -FontSize 36 -Bold $true -CreateIfMissing

    PptSlide -LayoutType $layout.Type {
        PptTitle -Title 'Layout Styled'
    }
}

Write-Host "Presentation saved to $path"
