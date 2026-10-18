Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot 'Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'DslInspection.pptx'

New-OfficePowerPoint -Path $path {
    $layouts = Get-OfficePowerPointLayout
    $layout = $layouts | Select-Object -First 1

    PptSlide {
        PptTitle -Title 'Inspect Me'
        $title = Get-OfficePowerPointPlaceholder -PlaceholderType Title
        $layoutPlaceholders = Get-OfficePowerPointLayoutPlaceholder
        Write-Host "Title placeholder: $($title.Text)"
        Write-Host "Layout placeholders: $($layoutPlaceholders.Count)"
    }

    if ($layout) {
        PptSlide -LayoutType $layout.Type -Content {
            PptTitle -Title "Layout: $($layout.Name)"
        }
    }
}

Write-Host "Presentation saved to $path"
