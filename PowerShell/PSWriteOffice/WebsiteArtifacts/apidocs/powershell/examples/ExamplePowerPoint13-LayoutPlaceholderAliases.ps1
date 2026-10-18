Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot 'Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'LayoutPlaceholderAliases.pptx'

New-OfficePowerPoint -Path $path {
    PptSlide {
        PptTitle -Title 'Alias Demo'
        $placeholders = PptLayoutPlaceholders
        Write-Host "Layout placeholders: $($placeholders.Count)"
    }
}

Write-Host "Presentation saved to $path"
