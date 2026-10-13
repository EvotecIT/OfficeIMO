# Shared by Test-, Export- and Import-StudioTranslations.ps1: reads and writes the interface strings of the desktop
# Studio (Avalonia .resx resources) and the native Apple apps (Xcode String Catalogs), so one language is added to
# every app from one translation file.

Set-StrictMode -Version Latest

$script:Root = (Resolve-Path (Join-Path $PSScriptRoot '..' '..')).Path
$script:DesktopFolder = Join-Path $script:Root 'OfficeIMO.Studio/Localization'
$script:DesktopNeutral = Join-Path $script:DesktopFolder 'StudioStrings.resx'
$script:CultureCatalog = Join-Path $script:Root 'OfficeIMO.Studio/Infrastructure/Localization/StudioCultureCatalog.cs'
$script:AppleFolder = Join-Path $script:Root 'Apple/OfficeIMOStudio/Shared'
$script:AppleCatalogs = @('Localizable', 'AppShortcuts', 'InfoPlist')
$script:AppleProject = Join-Path $script:Root 'Apple/OfficeIMOStudio/OfficeIMOStudio.xcodeproj/project.pbxproj'

# ---------------------------------------------------------------------------------------------------------------
# Languages

function Get-StudioShippedLanguages {
    <# The languages the desktop language picker offers, from StudioCultureCatalog.ShippedCultures, without English. #>
    $text = Get-Content -LiteralPath $script:CultureCatalog -Raw
    if ($text -notmatch 'ShippedCultures\s*=\s*\[([^\]]*)\]') { throw "StudioCultureCatalog.cs has no ShippedCultures list." }
    @([regex]::Matches($Matches[1], '"([^"]+)"') | ForEach-Object { $_.Groups[1].Value } | Where-Object { $_ -ne 'en' })
}

function Get-AppleKnownRegions {
    $text = Get-Content -LiteralPath $script:AppleProject -Raw
    if ($text -notmatch 'knownRegions = \(([^)]*)\);') { throw 'The Xcode project lists no knownRegions.' }
    @($Matches[1] -split ',' | ForEach-Object { $_.Trim().Trim('"') } | Where-Object { $_ -and $_ -notin @('Base', 'en') })
}

function Add-StudioLanguageRegistration {
    <# Lists a language in the desktop picker and in the Xcode project's known regions. #>
    param([Parameter(Mandatory)][string] $Language)
    $text = Get-Content -LiteralPath $script:CultureCatalog -Raw
    if ($Language -notin (Get-StudioShippedLanguages)) {
        $text = [regex]::Replace($text, '(ShippedCultures\s*=\s*\[[^\]]*?)(\s*\])', { param($m) $m.Groups[1].Value + ", `"$Language`"" + $m.Groups[2].Value }, 'Singleline')
        Set-Content -LiteralPath $script:CultureCatalog -Value $text -NoNewline -Encoding utf8
    }
    $project = Get-Content -LiteralPath $script:AppleProject -Raw
    if ($Language -notin (Get-AppleKnownRegions)) {
        $entry = if ($Language -match '^[A-Za-z0-9_]+$') { $Language } else { "`"$Language`"" }
        $project = [regex]::Replace($project, '(knownRegions = \([^)]*?)(\n\t*\);)', { param($m) $m.Groups[1].Value + "`n`t`t`t`t$entry," + $m.Groups[2].Value })
        Set-Content -LiteralPath $script:AppleProject -Value $project -NoNewline -Encoding utf8
    }
}

# ---------------------------------------------------------------------------------------------------------------
# Placeholders: what a translation must keep, in any order

function Get-DesktopTokens([string] $Value) {
    # The same rule as StudioTranslationResourceTests: {0}, {0:N0}, {page} and similar, but not escaped braces.
    @([regex]::Matches($Value, '(?<!\{)\{[^{}]+\}(?!\})') | ForEach-Object Value | Sort-Object)
}

function Get-AppleTokens([string] $Value) {
    # printf specifiers with their positions removed, so %1$@ and %@ match, and App Intents ${parameter}s.
    $specifiers = [regex]::Matches($Value, '%(?:\d+\$)?(?:lld|ld|lu|llu|d|i|u|@|f|\.\d+f|e|g|x|c|s)') |
        ForEach-Object { $_.Value -replace '^%\d+\$', '%' }
    $parameters = [regex]::Matches($Value, '\$\{\w+\}') | ForEach-Object Value
    @(@($specifiers) + @($parameters) | Where-Object { $_ } | Sort-Object)
}

function Test-SameTokens([string[]] $Expected, [string[]] $Actual) {
    (@($Expected) -join "`u{1}") -eq (@($Actual) -join "`u{1}")
}

# ---------------------------------------------------------------------------------------------------------------
# Desktop: .resx

function Read-ResxValues([string] $Path) {
    # Resource names are case-sensitive, as ResourceManager treats them.
    $values = [System.Collections.Specialized.OrderedDictionary]::new([StringComparer]::Ordinal)
    if (-not (Test-Path -LiteralPath $Path)) { return $values }
    [xml] $xml = Get-Content -LiteralPath $Path -Raw
    foreach ($data in $xml.root.data) { $values[$data.name] = [string] $data.value }
    $values
}

function Get-DesktopStrings {
    <# The English strings in the neutral resource's order. #>
    Read-ResxValues $script:DesktopNeutral
}

function Write-ResxValues {
    <# Writes a language's resource in the neutral resource's order; keys without a value are left out, so they fall back to English at run time and the checks report them. #>
    param([Parameter(Mandatory)][string] $Language, [Parameter(Mandatory)] $Values)
    $neutral = Get-Content -LiteralPath $script:DesktopNeutral -Raw
    $header = $neutral.Substring(0, $neutral.IndexOf('  <data '))
    $builder = [System.Text.StringBuilder]::new($header)
    foreach ($key in (Get-DesktopStrings).Keys) {
        $value = $Values[$key]
        if ([string]::IsNullOrWhiteSpace($value)) { continue }
        $escaped = [System.Security.SecurityElement]::Escape($value) -replace '&apos;', "'" -replace '&quot;', '"'
        [void] $builder.Append("  <data name=`"$key`" xml:space=`"preserve`">`n    <value>$escaped</value>`n  </data>`n")
    }
    [void] $builder.Append("</root>`n")
    Set-Content -LiteralPath (Join-Path $script:DesktopFolder "StudioStrings.$Language.resx") -Value $builder.ToString() -NoNewline -Encoding utf8
}

# ---------------------------------------------------------------------------------------------------------------
# Apple: String Catalogs

function Read-AppleCatalog([string] $Name) {
    Get-Content -LiteralPath (Join-Path $script:AppleFolder "$Name.xcstrings") -Raw | ConvertFrom-Json -AsHashtable
}

function Get-AppleEntries {
    <# Every string a catalog asks translators for: its key, English text and comment. Strings marked not to translate, and strings stale in the source, are left out. #>
    param([Parameter(Mandatory)][string] $Catalog, [Parameter(Mandatory)] $Document)
    foreach ($key in ($Document.strings.Keys | Sort-Object -CaseSensitive)) {
        $entry = $Document.strings[$key]
        if ($entry.ContainsKey('shouldTranslate') -and -not $entry.shouldTranslate) { continue }
        if ($entry.ContainsKey('extractionState') -and $entry.extractionState -eq 'stale') { continue }
        $source = $key
        if ($entry.ContainsKey('localizations') -and $entry.localizations.ContainsKey('en')) { $source = $entry.localizations.en.stringUnit.value }
        [pscustomobject]@{ Catalog = $Catalog; Key = $key; Source = $source; Comment = $entry['comment']; Entry = $entry }
    }
}

function Get-AppleTranslation($Entry, [string] $Language) {
    if (-not $Entry.ContainsKey('localizations') -or -not $Entry.localizations.ContainsKey($Language)) { return $null }
    $unit = $Entry.localizations[$Language]['stringUnit']
    if ($null -eq $unit -or $unit.state -ne 'translated') { return $null }
    [string] $unit.value
}

function Set-AppleTranslation($Entry, [string] $Language, [string] $Value) {
    if (-not $Entry.ContainsKey('localizations')) { $Entry['localizations'] = @{} }
    $Entry.localizations[$Language] = @{ stringUnit = @{ state = 'translated'; value = $Value } }
}

function Write-AppleCatalog([string] $Name, $Document) {
    # Xcode's own layout: two-space indents, " : " between names and values, names sorted.
    Set-Content -LiteralPath (Join-Path $script:AppleFolder "$Name.xcstrings") -Value ((ConvertTo-CatalogJson $Document 0) + "`n") -NoNewline -Encoding utf8
}

function ConvertTo-CatalogJson($Value, [int] $Depth) {
    $indent = '  ' * ($Depth + 1)
    $close = '  ' * $Depth
    if ($null -eq $Value) { return 'null' }
    if ($Value -is [bool]) { return $Value.ToString().ToLowerInvariant() }
    if ($Value -is [string]) { return ConvertTo-JsonString $Value }
    if ($Value -is [System.Collections.IDictionary]) {
        if ($Value.Count -eq 0) { return '{}' }
        $names = [string[]] @($Value.Keys)
        [Array]::Sort($names, [StringComparer]::Ordinal)
        $members = foreach ($name in $names) { $indent + (ConvertTo-JsonString $name) + ' : ' + (ConvertTo-CatalogJson $Value[$name] ($Depth + 1)) }
        return "{`n" + ($members -join ",`n") + "`n$close}"
    }
    if ($Value -is [System.Collections.IEnumerable]) {
        $items = @($Value)
        if ($items.Count -eq 0) { return '[]' }
        return "[`n" + (($items | ForEach-Object { $indent + (ConvertTo-CatalogJson $_ ($Depth + 1)) }) -join ",`n") + "`n$close]"
    }
    return [System.Convert]::ToString($Value, [cultureinfo]::InvariantCulture)
}

function ConvertTo-JsonString([string] $Text) {
    $escaped = [regex]::Replace($Text, '["\\\x00-\x1f]', {
        param($match)
        switch ($match.Value) {
            '"' { '\"' }
            '\' { '\\' }
            "`n" { '\n' }
            "`r" { '\r' }
            "`t" { '\t' }
            default { '\u{0:x4}' -f [int][char] $match.Value }
        }
    })
    '"' + $escaped + '"'
}

Export-ModuleMember -Function Get-StudioShippedLanguages, Get-AppleKnownRegions, Add-StudioLanguageRegistration, Get-DesktopTokens,
    Get-AppleTokens, Test-SameTokens, Read-ResxValues, Get-DesktopStrings, Write-ResxValues, Read-AppleCatalog, Get-AppleEntries,
    Get-AppleTranslation, Set-AppleTranslation, Write-AppleCatalog -Variable Root, DesktopFolder, AppleCatalogs
