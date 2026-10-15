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

# The desktop also looks up keys with an English fallback (GetOrDefault, FormatOrDefault), directly or through a
# helper that adds a prefix, such as T("Saved", "...") for "Book.Saved". A missing key shows English in every language
# without failing the checks above, so each key written in the code must be in the English resources. Keys built from
# values are checked by the desktop tests.
$sourceRoot = Join-Path $PSScriptRoot '..' '..' 'OfficeIMO.Studio'
$sources = Get-ChildItem -Path $sourceRoot -Recurse -Filter '*.cs' | Where-Object FullName -notmatch '[\\/](bin|obj)[\\/]'
$typePattern = [regex] '\b(?:class|struct|interface|record(?:\s+(?:class|struct))?)\s+(\w+)\s*(?=[<:{(\r\n]|where\b)'
$helperPattern = [regex] '(?s)string\s+(\w+)\(\s*string\s+(\w+)\s*,\s*string\s+\w+\s*\)\s*(?:=>|\{[^}]*?return)\s*[\w.?]*(?:GetOrDefault|FormatOrDefault)\(\s*"([^"]*)"\s*\+\s*\2\b'
$literalPattern = [regex] '(?:GetOrDefault|FormatOrDefault)\(\s*"([^"]+)"\s*,'
function Get-EnclosingType([string] $Text, [int] $Position) {
    $name = $null
    foreach ($match in $typePattern.Matches($Text.Substring(0, $Position))) { $name = $match.Groups[1].Value }
    $name
}
$texts = @{}
foreach ($file in $sources) { $texts[$file.FullName] = Get-Content -LiteralPath $file.FullName -Raw }
$helpers = @{}
foreach ($entry in $texts.GetEnumerator()) {
    foreach ($match in $helperPattern.Matches($entry.Value)) {
        $helpers["$(Get-EnclosingType $entry.Value $match.Index)|$($match.Groups[1].Value)"] = $match.Groups[3].Value
    }
}
$helperNames = @($helpers.Keys | ForEach-Object { $_.Split('|')[1] } | Sort-Object -Unique)
$used = [System.Collections.Generic.Dictionary[string, string]]::new([StringComparer]::Ordinal)
foreach ($entry in $texts.GetEnumerator()) {
    $relative = [System.IO.Path]::GetRelativePath((Join-Path $PSScriptRoot '..' '..'), $entry.Key)
    foreach ($match in $literalPattern.Matches($entry.Value)) { $used[$match.Groups[1].Value] = $relative }
    foreach ($name in $helperNames) {
        foreach ($match in [regex]::Matches($entry.Value, "(?<![\w.])$name\(\s*`"([^`"]+)`"\s*,")) {
            $prefix = $helpers["$(Get-EnclosingType $entry.Value $match.Index)|$name"]
            if ($null -eq $prefix) {
                # A helper used outside its own type only resolves when one type defines it.
                $candidates = @($helpers.Keys | Where-Object { $_.EndsWith("|$name") })
                if ($candidates.Count -eq 1) { $prefix = $helpers[$candidates[0]] } else { continue }
            }
            $used[$prefix + $match.Groups[1].Value] = $relative
        }
    }
}
foreach ($key in $used.Keys | Sort-Object) {
    if (-not $english.Contains($key)) { $problems.Add("desktop ${key}: used in $($used[$key]) but not in StudioStrings.resx.") }
}

if ($problems.Count) {
    $problems | ForEach-Object { Write-Host "Studio translations: $_" }
    throw "The Studio translations have $($problems.Count) problem(s)."
}
Write-Host "Studio translations: $($languages -join ', ') complete on the desktop ($($english.Count) strings, $($used.Count) with a fallback in code) and Apple ($(($AppleCatalogs | ForEach-Object { @(Get-AppleEntries -Catalog $_ -Document $catalogs[$_]).Count } | Measure-Object -Sum).Sum) strings)."
