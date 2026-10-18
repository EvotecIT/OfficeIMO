Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot 'Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'ExamplePowerPoint9-Placeholders.pptx'
$presentation = New-OfficePowerPoint -Path $path -NoSave

$layouts = Get-OfficePowerPointLayout -Presentation $presentation
$layout = $layouts | Where-Object { $_.Type } | Select-Object -First 1
if (-not $layout) {
    $layout = $layouts | Select-Object -First 1
}

$slide = if ($layout.Type) {
    Add-OfficePowerPointSlide -Presentation $presentation -LayoutType $layout.Type -Master $layout.MasterIndex -PassThru
} elseif ($layout.Name) {
    Add-OfficePowerPointSlide -Presentation $presentation -LayoutName $layout.Name -Master $layout.MasterIndex -PassThru
} else {
    Add-OfficePowerPointSlide -Presentation $presentation -Layout $layout.LayoutIndex -Master $layout.MasterIndex -PassThru
}

Set-OfficePowerPointPlaceholderText -Slide $slide -PlaceholderType Title -Text 'Status Update'

$layoutPlaceholders = Get-OfficePowerPointLayoutPlaceholder -Slide $slide
$placeholder = $layoutPlaceholders | Where-Object { $_.PlaceholderType } | Select-Object -First 1
if ($placeholder) {
    $placeholderType = $placeholder.PlaceholderType.ToString()
    Set-OfficePowerPointLayoutPlaceholderBounds -Presentation $presentation -Master $layout.MasterIndex -Layout $layout.LayoutIndex `
        -PlaceholderType $placeholderType -Index $placeholder.PlaceholderIndex -Left 60 -Top 140 -Width 520 -Height 240
    Set-OfficePowerPointLayoutPlaceholderTextMargins -Presentation $presentation -Master $layout.MasterIndex -Layout $layout.LayoutIndex `
        -PlaceholderType $placeholderType -Index $placeholder.PlaceholderIndex -Left 12 -Top 8 -Right 12 -Bottom 8
    Set-OfficePowerPointLayoutPlaceholderTextStyle -Presentation $presentation -Master $layout.MasterIndex -Layout $layout.LayoutIndex `
        -PlaceholderType $placeholderType -Index $placeholder.PlaceholderIndex -Style Body -FontSize 18 -Bold $true
}

Save-OfficePowerPoint -Presentation $presentation
Write-Host "Presentation saved to $path"
