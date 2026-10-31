Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot 'Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'LayoutDslExample.pptx'

New-OfficePowerPoint -Path $path {
    $layout = Get-OfficePowerPointLayout | Select-Object -First 1
    if (-not $layout) {
        return
    }

    Set-OfficePowerPointLayoutPlaceholder -TemplateLayout $layout -Type Title `
        -X 60 -Y 40 -Width 600 -Height 120 -CreateIfMissing
    Set-OfficePowerPointLayoutPlaceholder -TemplateLayout $layout -Type Title `
        -TextMargin 6,8,6,8 -CreateIfMissing
    Set-OfficePowerPointLayoutPlaceholder -TemplateLayout $layout -Type Title `
        -Style Title -Size 36 -Bold -CreateIfMissing

    PptSlide -LayoutType $layout.Type {
        PptTitle -Title 'Layout Styled'
    }
}

Write-Host "Presentation saved to $path"
