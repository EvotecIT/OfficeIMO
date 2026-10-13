#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Writes translated strings from Export-StudioTranslations.ps1 files into the desktop and Apple Studio apps.

.DESCRIPTION
    Reads each translation file, refuses it when a translation loses or adds a placeholder or names a string the apps
    don't have, then writes the desktop StudioStrings.<language>.resx and the Apple String Catalogs and registers the
    language in the desktop picker (StudioCultureCatalog.ShippedCultures) and the Xcode project's knownRegions. Empty
    translations are skipped, so a language can arrive in several passes; Test-StudioTranslations.ps1 reports what is
    still missing. -NoRegister writes the strings without offering the language yet.

.EXAMPLE
    ./Build/Studio/Import-StudioTranslations.ps1 -Path Artefacts/Translations/it.json, Artefacts/Translations/es.json
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string[]] $Path,
    [switch] $NoRegister
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'StudioTranslations.psm1') -Force
# pwsh -File passes a comma-separated list as one value.
$Path = @($Path | ForEach-Object { $_ -split ',' } | ForEach-Object Trim | Where-Object { $_ })

$english = Get-DesktopStrings
$catalogs = @{}
foreach ($name in $AppleCatalogs) { $catalogs[$name] = Read-AppleCatalog $name }
# Keys are case-sensitive: the catalogs hold strings such as "Confidential" and "CONFIDENTIAL".
$appleEntries = [System.Collections.Generic.Dictionary[string, object]]::new([StringComparer]::Ordinal)
foreach ($name in $AppleCatalogs) { foreach ($entry in Get-AppleEntries -Catalog $name -Document $catalogs[$name]) { $appleEntries["apple/${name}:$($entry.Key)"] = $entry } }

foreach ($file in $Path) {
    $document = Get-Content -LiteralPath $file -Raw | ConvertFrom-Json
    $code = [System.Globalization.CultureInfo]::GetCultureInfo($document.language).Name
    if ($code -eq 'en' -or $code -like 'en-*') { throw "$file translates into English, the source language." }
    $problems = [System.Collections.Generic.List[string]]::new()
    $desktopValues = Read-ResxValues (Join-Path $DesktopFolder "StudioStrings.$code.resx")
    $desktopCount = 0
    $appleCount = 0
    foreach ($item in $document.strings) {
        $value = [string] $item.translation
        if ([string]::IsNullOrWhiteSpace($value)) { continue }
        if ($item.id -like 'desktop:*') {
            $key = $item.id.Substring(8)
            if (-not $english.Contains($key)) { $problems.Add("$($item.id): no such desktop string."); continue }
            if (-not (Test-SameTokens (Get-DesktopTokens $english[$key]) (Get-DesktopTokens $value))) { $problems.Add("$($item.id): the placeholders differ from English."); continue }
            $desktopValues[$key] = $value
            $desktopCount++
        } elseif ($appleEntries.ContainsKey($item.id)) {
            $entry = $appleEntries[$item.id]
            if (-not (Test-SameTokens (Get-AppleTokens $entry.Source) (Get-AppleTokens $value))) { $problems.Add("$($item.id): the placeholders differ from English."); continue }
            Set-AppleTranslation $entry.Entry $code $value
            $appleCount++
        } else {
            $problems.Add("$($item.id): no such string; export the language again.")
        }
    }
    if ($problems.Count) {
        $problems | ForEach-Object { Write-Host "${file}: $_" }
        throw "$file has $($problems.Count) problem(s); nothing from it was written."
    }
    # A file without desktop strings leaves the language's resources as they are.
    if ($desktopCount) { Write-ResxValues -Language $code -Values $desktopValues }
    foreach ($name in $AppleCatalogs) { Write-AppleCatalog $name $catalogs[$name] }
    if (-not $NoRegister) { Add-StudioLanguageRegistration -Language $code }
    Write-Host "${code}: $desktopCount desktop and $appleCount Apple strings written from $file."
}
Write-Host 'Run ./Build/Studio/Test-StudioTranslations.ps1 to see what each language still needs.'
