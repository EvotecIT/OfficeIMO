#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Checks that every Studio language is complete in every app.

.DESCRIPTION
    For each language the desktop picker ships (StudioCultureCatalog.ShippedCultures) and each language the Apple apps
    know (the Xcode project's knownRegions), checks that both apps list it, that the desktop StudioStrings.<language>.resx
    and every Apple String Catalog translate every string, and that each translation keeps its placeholders. Prints
    each problem and fails when there is one; -Language limits the check to some languages.
#>
[CmdletBinding()]
param([string[]] $Language)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'StudioTranslations.psm1') -Force
# pwsh -File passes a comma-separated list as one value.
$Language = @($Language | ForEach-Object { $_ -split ',' } | ForEach-Object Trim | Where-Object { $_ })

$shipped = Get-StudioShippedLanguages
$known = Get-AppleKnownRegions
$problems = [System.Collections.Generic.List[string]]::new()
foreach ($missing in $shipped | Where-Object { $_ -notin $known }) { $problems.Add("$missing ships on the desktop but isn't a known region of the Apple project.") }
foreach ($missing in $known | Where-Object { $_ -notin $shipped }) { $problems.Add("$missing is a known region of the Apple project but the desktop doesn't ship it.") }

$languages = if ($Language) { $Language } else { @($shipped) + @($known) | Sort-Object -Unique }
$english = Get-DesktopStrings
$catalogs = @{}
foreach ($name in $AppleCatalogs) { $catalogs[$name] = Read-AppleCatalog $name }

foreach ($code in $languages) {
    $desktop = Read-ResxValues (Join-Path $DesktopFolder "StudioStrings.$code.resx")
    $missingDesktop = 0
    foreach ($key in $english.Keys) {
        $value = $desktop[$key]
        if ([string]::IsNullOrWhiteSpace($value)) { $missingDesktop++; continue }
        if (-not (Test-SameTokens (Get-DesktopTokens $english[$key]) (Get-DesktopTokens $value))) {
            $problems.Add("${code} desktop ${key}: the placeholders differ from English.")
        }
    }
    foreach ($extra in $desktop.Keys | Where-Object { -not $english.Contains($_) }) { $problems.Add("${code} desktop ${extra}: not in the English resources.") }
    if ($missingDesktop) { $problems.Add("${code} desktop: $missingDesktop of $($english.Count) strings untranslated.") }

    foreach ($name in $AppleCatalogs) {
        $missingApple = 0
        foreach ($entry in Get-AppleEntries -Catalog $name -Document $catalogs[$name]) {
            $value = Get-AppleTranslation $entry.Entry $code
            if ([string]::IsNullOrWhiteSpace($value)) { $missingApple++; continue }
            if (-not (Test-SameTokens (Get-AppleTokens $entry.Source) (Get-AppleTokens $value))) {
                $problems.Add("${code} Apple $name `"$($entry.Key)`": the placeholders differ from English.")
            }
        }
        if ($missingApple) { $problems.Add("${code} Apple ${name}: $missingApple strings untranslated.") }
    }
}

if ($problems.Count) {
    $problems | ForEach-Object { Write-Host "Studio translations: $_" }
    throw "The Studio translations have $($problems.Count) problem(s)."
}
Write-Host "Studio translations: $($languages -join ', ') complete on the desktop ($($english.Count) strings) and Apple ($(($AppleCatalogs | ForEach-Object { @(Get-AppleEntries -Catalog $_ -Document $catalogs[$_]).Count } | Measure-Object -Sum).Sum) strings)."
