param(
    [string] $OutputDirectory = (Join-Path $PWD 'PowerPoint-LayoutPlaceholders')
)

Import-Module PSWriteOffice -ErrorAction Stop
$null = New-Item -Path $OutputDirectory -ItemType Directory -Force
$path = Join-Path $OutputDirectory 'LayoutPlaceholders.pptx'

New-OfficePowerPoint -Path $path -Force -PassThru {
    $layout = Get-OfficePowerPointLayout | Where-Object Type -eq Text | Select-Object -First 1
    $slide = PowerPointSlide -Title 'Template placeholders' -TemplateLayout $layout -PassThru {
        PowerPointParagraph 'Inspect the template before changing its positions or text style.'
    }
    $placeholders = @(Get-OfficePowerPointPlaceholder -Slide $slide -FromLayout)
    Write-Host "Layout placeholders: $($placeholders.Count)"
}
