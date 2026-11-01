[CmdletBinding()]
param([string] $OutputDirectory = (Join-Path $PSScriptRoot '..\Documents'))

Import-Module PSWriteOffice -ErrorAction Stop
$null = New-Item -Path $OutputDirectory -ItemType Directory -Force
$path = Join-Path $OutputDirectory 'PowerPoint-ThemeAndLayout.pptx'
$deck = New-OfficePowerPoint -Path $path
try {
    $slide = Add-OfficePowerPointSlide -Document $deck -TemplateLayout 1 -Title 'Theme Demo' -PassThru
    $layouts = Get-OfficePowerPointLayout -Document $deck
    $targetLayout = $layouts | Where-Object LayoutIndex -ne $slide.LayoutIndex | Select-Object -First 1

    Set-OfficePowerPointTheme -Document $deck -Colors @{
        Accent1 = '#C00000'
        Accent2 = '#00B0F0'
    } -HeadingFont 'Aptos' -BodyFont 'Calibri' -Name 'Contoso Theme' -AllMasters

    $slide | Set-OfficePowerPointSlide -Document $deck -TemplateLayout $targetLayout
    Get-OfficePowerPointTheme -Document $deck | Format-List
    Save-OfficePowerPoint -Document $deck
} finally {
    Close-OfficePowerPoint -Document $deck
}
Write-Host "Presentation saved to $path"
