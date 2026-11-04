param([string] $OutputDirectory = (Join-Path $PSScriptRoot 'Documents'))

Import-Module PSWriteOffice -ErrorAction Stop
$null = New-Item -Path $OutputDirectory -ItemType Directory -Force
$path = Join-Path $OutputDirectory 'DslInspection.pptx'

New-OfficePowerPoint -Path $path -Force {
    $layout = Get-OfficePowerPointLayout | Where-Object Type -eq Text | Select-Object -First 1
    $slide = PowerPointSlide -Title 'Inspect Me' -TemplateLayout $layout -PassThru {
        PowerPointParagraph 'Review the editable template fields before filling the deck.'
    }
    $title = $slide | Get-OfficePowerPointPlaceholder -Type Title
    $layoutPlaceholders = @($slide | Get-OfficePowerPointPlaceholder -FromLayout)
    Write-Host "Title placeholder: $($title.Text)"
    Write-Host "Layout placeholders: $($layoutPlaceholders.Count)"

    PowerPointSlide -Title "Layout: $($layout.Name)" -TemplateLayout $layout
}
Write-Host "Presentation saved to $path"
