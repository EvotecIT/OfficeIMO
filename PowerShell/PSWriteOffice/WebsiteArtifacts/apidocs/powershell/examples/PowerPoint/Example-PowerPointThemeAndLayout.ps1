Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'PowerPoint-ThemeAndLayout.pptx'
$ppt = New-OfficePowerPoint -Path $path -NoSave

$slide = Add-OfficePowerPointSlide -Presentation $ppt -Layout 1 -PassThru
Set-OfficePowerPointSlideTitle -Slide $slide -Title 'Theme Demo'

$layouts = Get-OfficePowerPointLayout -Presentation $ppt
$targetLayout = $layouts | Where-Object LayoutIndex -ne $slide.LayoutIndex | Select-Object -First 1

Set-OfficePowerPointThemeColor -Presentation $ppt -Colors @{
    Accent1 = '#C00000'
    Accent2 = '#00B0F0'
} -AllMasters
Set-OfficePowerPointThemeFonts -Presentation $ppt -MajorLatin 'Aptos' -MinorLatin 'Calibri' -AllMasters
Set-OfficePowerPointThemeName -Presentation $ppt -Name 'Contoso Theme' -AllMasters

if ($targetLayout.Type) {
    $slide | Set-OfficePowerPointSlideLayout -LayoutType $targetLayout.Type -Master $targetLayout.MasterIndex
} elseif ($targetLayout.Name) {
    $slide | Set-OfficePowerPointSlideLayout -LayoutName $targetLayout.Name -Master $targetLayout.MasterIndex
} else {
    $slide | Set-OfficePowerPointSlideLayout -Layout $targetLayout.LayoutIndex -Master $targetLayout.MasterIndex
}

$theme = Get-OfficePowerPointTheme -Presentation $ppt
$theme | Format-List

Save-OfficePowerPoint -Presentation $ppt
$ppt | Close-OfficePowerPoint

Write-Host "Presentation saved to $path"
