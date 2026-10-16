#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Writes one translation file per language with every Studio string that language still needs.

.DESCRIPTION
    Collects the strings of the desktop Studio (StudioStrings.resx) and the Apple apps (their String Catalogs) that a
    language doesn't translate yet into <OutputDirectory>/<language>.json. A translator, person or agent, fills each
    "translation" and Import-StudioTranslations.ps1 writes them into both apps. -All exports every string, translated
    or not, to review or redo a language.

.EXAMPLE
    ./Build/Studio/Export-StudioTranslations.ps1 -Language it, es, pt-BR -OutputDirectory Artefacts/Translations
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string[]] $Language,
    [string] $OutputDirectory = 'Artefacts/Translations',
    [switch] $All
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'StudioTranslations.psm1') -Force
# pwsh -File passes a comma-separated list as one value.
$Language = @($Language | ForEach-Object { $_ -split ',' } | ForEach-Object Trim | Where-Object { $_ })

$rules = @(
    'Translate "source" into the language and put the result in "translation"; leave "translation" empty to skip a string.',
    'Keep every placeholder exactly: {0}, {1:N0}, {page}, %@, %lld, %1$@ and ${applicationName}. A positional form such as %2$@ may move a value.',
    'Keep product and format names as they are: OfficeIMO, Studio, PDF, Word, Excel, PowerPoint, Markdown, JSON, PDF/A, Bates.',
    'Keep {page}, {pages}, {date}, {file} and {title} in header and footer text: Studio fills them in.',
    'Keep the length close to English where possible; menus and buttons have limited room. Use the platform''s usual terms (for example the system''s words for Undo, Save and Share).',
    '"context" says where the string appears: a desktop resource key names its screen, and an Apple comment describes its use.'
)

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$english = Get-DesktopStrings
$catalogs = @{}
foreach ($name in $AppleCatalogs) { $catalogs[$name] = Read-AppleCatalog $name }

foreach ($code in $Language) {
    $culture = [System.Globalization.CultureInfo]::GetCultureInfo($code)
    $desktop = Read-ResxValues (Join-Path $DesktopFolder "StudioStrings.$code.resx")
    $strings = [System.Collections.Generic.List[object]]::new()
    foreach ($key in $english.Keys) {
        $current = $desktop[$key]
        if (-not $All -and -not [string]::IsNullOrWhiteSpace($current)) { continue }
        $strings.Add([ordered]@{ id = "desktop:$key"; context = $key; source = $english[$key]; translation = [string] $current })
    }
    foreach ($name in $AppleCatalogs) {
        foreach ($entry in Get-AppleEntries -Catalog $name -Document $catalogs[$name]) {
            $current = Get-AppleTranslation $entry.Entry $code
            if (-not $All -and -not [string]::IsNullOrWhiteSpace($current)) { continue }
            $strings.Add([ordered]@{ id = "apple/${name}:$($entry.Key)"; context = [string] $entry.Comment; source = $entry.Source; translation = [string] $current })
        }
    }
    $file = Join-Path $OutputDirectory "$code.json"
    [ordered]@{
        language = $culture.Name
        englishName = $culture.EnglishName
        nativeName = $culture.NativeName
        rightToLeft = $culture.TextInfo.IsRightToLeft
        rules = $rules
        strings = $strings
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $file -Encoding utf8
    Write-Host "$($culture.Name) ($($culture.EnglishName)): $($strings.Count) strings to translate in $file"
}
