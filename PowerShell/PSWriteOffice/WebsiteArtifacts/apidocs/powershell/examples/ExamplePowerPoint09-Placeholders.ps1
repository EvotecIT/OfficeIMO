Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot 'Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'ExamplePowerPoint9-Placeholders.pptx'
$presentation = New-OfficePowerPoint -Path $path -Theme Modern

try {
    $layout = Get-OfficePowerPointLayout -Document $presentation |
        Where-Object Name -eq 'Title and Content' | Select-Object -First 1

    Set-OfficePowerPointLayoutPlaceholder -Document $presentation -TemplateLayout $layout -Type Title `
        -X 60pt -Y 40pt -Width 600pt -Height 120pt -TextMargin 6pt,8pt -Style Title -Size 32pt -Bold

    PowerPointSlide -Document $presentation -TemplateLayout $layout -Title 'Status update' {
        PowerPointParagraph 'The title inherits the layout bounds, margins and text style.'
    }
    Save-OfficePowerPoint -Document $presentation
} finally {
    Close-OfficePowerPoint -Document $presentation
}

Write-Host "Presentation saved to $path"
