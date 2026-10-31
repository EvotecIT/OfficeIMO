Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot 'Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'LayoutDslExample.pptx'

New-OfficePowerPoint -Path $path {
    $layout = Get-OfficePowerPointLayout |
        Where-Object Name -eq 'Title and Content' | Select-Object -First 1

    Set-OfficePowerPointLayoutPlaceholder -TemplateLayout $layout -Type Title `
        -X 60pt -Y 40pt -Width 600pt -Height 120pt -TextMargin 6pt,8pt -Style Title -Size 36pt -Bold

    PowerPointSlide -Title 'Layout styled' -TemplateLayout $layout {
        PowerPointParagraph 'One command sets the layout title geometry and formatting.'
    }
}

Write-Host "Presentation saved to $path"
