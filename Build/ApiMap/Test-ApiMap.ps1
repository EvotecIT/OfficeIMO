<#
.SYNOPSIS
Validates the PowerShell API map against the current module and the 3.6 design rules.

.DESCRIPTION
Checks that every current PSWriteOffice command and alias is mapped exactly once, that every
target exists, and that target commands, parameters and aliases follow the naming rules in
Docs/officeimo.api-design.md. Run it with the module source built, or pass a metadata file.

.EXAMPLE
./Build/ApiMap/Test-ApiMap.ps1 -ModulePath ./PowerShell/PSWriteOffice/PSWriteOffice.psd1
#>
[CmdletBinding()]
param(
    [string] $ModulePath,
    [string] $MetadataPath,
    [string] $MapDirectory = (Join-Path $PSScriptRoot 'powershell')
)

$ErrorActionPreference = 'Stop'

if ($MetadataPath) {
    $metadata = Get-Content -LiteralPath $MetadataPath -Raw | ConvertFrom-Json
    $currentCommands = @($metadata.Cmdlets.Name)
    $currentAliases = @($metadata.Aliases.Alias)
} elseif ($ModulePath) {
    $module = Import-Module $ModulePath -Force -PassThru
    $currentCommands = @(Get-Command -Module $module.Name -CommandType Cmdlet, Function | ForEach-Object Name)
    $currentAliases = @(Get-Command -Module $module.Name -CommandType Alias | ForEach-Object Name)
} else {
    throw 'Specify -ModulePath or -MetadataPath.'
}

$approvedVerbs = @(Get-Verb | ForEach-Object Verb)
$builtInCommands = @(pwsh -NoProfile -NonInteractive -Command 'Get-Command -CommandType Cmdlet, Function, Alias | ForEach-Object Name')
$aliasPrefixes = 'Word', 'Excel', 'Ppt', 'Pdf', 'Markdown', 'Visio', 'Html', 'Rtf', 'OpenDocument', 'Csv', 'Email', 'Project', 'Access', 'Epub', 'OneNote', 'AsciiDoc', 'Latex', 'Confluence'
$bannedParameters = @{
    'InputPath'  = 'Use -Path for the input file.'
    'SourcePath' = 'Use -Path for the input file.'
    'FilePath'   = 'Use -Path, with FilePath at most as an alias.'
    'SheetIndex' = 'Use -Sheet, which accepts a name, a 1-based index or a sheet.'
    'SheetName'  = 'Use -Sheet.'
    'Presentation' = 'Use -Document.'
    'Workbook'   = 'Use -Document.'
    'NoSave'     = 'Leave out the content block to keep a document open.'
}
$unitSuffix = '(Pixels|Points|Pt|Px|Cm|Centimeters|Inches|Emus|Twips|Mm|Millimeters)$'

$errors = [System.Collections.Generic.List[string]]::new()
$warnings = [System.Collections.Generic.List[string]]::new()
$entries = [System.Collections.Generic.List[object]]::new()
$commands = @{}
$aliasOwners = @{}

$files = @(Get-ChildItem -LiteralPath $MapDirectory -Filter '*.json' | Sort-Object Name)
if ($files.Count -eq 0) { throw "No map files in $MapDirectory." }

foreach ($file in $files) {
    try {
        $data = Get-Content -LiteralPath $file.FullName -Raw | ConvertFrom-Json
    } catch {
        $errors.Add("$($file.Name): invalid JSON. $($_.Exception.Message)")
        continue
    }
    foreach ($command in @($data.commands)) {
        if ($commands.ContainsKey($command.name)) {
            $errors.Add("$($file.Name): command $($command.name) is also defined in $($commands[$command.name].File).")
            continue
        }
        $commands[$command.name] = [pscustomobject]@{ File = $file.Name; Definition = $command }
        if ($command.alias) {
            if ($aliasOwners.ContainsKey($command.alias)) {
                $errors.Add("$($file.Name): alias $($command.alias) is used by $($command.name) and $($aliasOwners[$command.alias]).")
            }
            $aliasOwners[$command.alias] = $command.name
        }
    }
    foreach ($entry in @($data.map)) {
        $entry | Add-Member -NotePropertyName File -NotePropertyValue $file.Name -Force
        $entries.Add($entry)
    }
}

# Coverage: every current command and alias is mapped exactly once.
$byOld = $entries | Group-Object old
foreach ($group in $byOld | Where-Object Count -gt 1) {
    $errors.Add("$($group.Name) is mapped $($group.Count) times ($(@($group.Group.File) -join ', ')).")
}
$mapped = @{}
foreach ($entry in $entries) { $mapped[$entry.old] = $entry }
foreach ($name in $currentCommands) {
    if (-not $mapped.ContainsKey($name)) { $errors.Add("Command $name has no map entry.") }
}
foreach ($name in $currentAliases) {
    if (-not $mapped.ContainsKey($name)) { $errors.Add("Alias $name has no map entry.") }
}
foreach ($entry in $entries) {
    $known = if ($entry.kind -eq 'alias') { $currentAliases } else { $currentCommands }
    if ($entry.old -notin $known) { $warnings.Add("$($entry.File): $($entry.old) is not a current $($entry.kind).") }
}

# Entry rules.
$actions = 'keep', 'rename', 'merge', 'parameter', 'cross-format', 'remove'
foreach ($entry in $entries) {
    $label = "$($entry.File): $($entry.old)"
    if ($entry.action -notin $actions) { $errors.Add("$label has unknown action '$($entry.action)'."); continue }
    if ($entry.action -ne 'remove' -and -not $entry.new) { $errors.Add("$label ($($entry.action)) has no target.") }
    if ($entry.action -ne 'keep' -and -not $entry.reason) { $errors.Add("$label ($($entry.action)) has no reason.") }
    if ($entry.action -in 'rename', 'merge', 'parameter', 'cross-format', 'remove' -and -not $entry.hint) {
        $warnings.Add("$label ($($entry.action)) has no hint.")
    }
    if ($entry.action -eq 'remove') { continue }
    if ($entry.kind -eq 'alias' -and $entry.new -notmatch '-') {
        if (-not $aliasOwners.ContainsKey($entry.new)) { $errors.Add("$label targets alias $($entry.new), which no command declares.") }
    } elseif (-not $commands.ContainsKey($entry.new)) {
        $errors.Add("$label targets command $($entry.new), which no map file defines.")
    } else {
        # The parameter set and renamed parameters must exist on the target command.
        $target = $commands[$entry.new].Definition
        $targetSets = @($target.parameterSets)
        if ($entry.parameterSet) {
            # "Path (-To Pdf)" annotates a set; "A | B" lists alternatives.
            foreach ($setName in ($entry.parameterSet -replace '\s*\(.*$', '') -split '\s*\|\s*') {
                if ($setName -and $setName -notin @($targetSets.name)) {
                    $errors.Add("$label targets parameter set '$setName', which $($entry.new) does not define.")
                }
            }
        }
        $targetParameters = @($targetSets | ForEach-Object { @($_.parameters) } | ForEach-Object name) + 'WhatIf', 'Confirm'
        if ($entry.parameters) {
            foreach ($property in $entry.parameters.PSObject.Properties) {
                foreach ($newName in @($property.Value)) {
                    if (-not $newName) { continue }
                    $first = ($newName -split '[\s:=]')[0].TrimStart('-')
                    if (-not $first) { continue }
                    if ($first -match '^[A-Z][a-z]+-Office') {
                        # The parameter moved to another command.
                        if (-not $commands.ContainsKey($first)) { $errors.Add("$label moves -$($property.Name) to $first, which no map file defines.") }
                        continue
                    }
                    if ($first -match '^Options\.') {
                        # The parameter moved into the options object.
                        if ('Options' -notin $targetParameters) { $errors.Add("$label moves -$($property.Name) into -Options, which $($entry.new) does not define.") }
                        continue
                    }
                    if ($first -notin $targetParameters) {
                        $errors.Add("$label maps -$($property.Name) to -$first, which $($entry.new) does not define.")
                    }
                }
            }
        }
    }
}

# Target command rules.
foreach ($name in $commands.Keys) {
    $definition = $commands[$name].Definition
    $label = "$($commands[$name].File): $name"
    if ($name -notmatch '^(?<verb>[A-Z][a-z]+)-Office[A-Z][A-Za-z0-9]*$') {
        $errors.Add("$label does not follow Verb-Office{Noun}.")
    } elseif ($Matches.verb -notin $approvedVerbs) {
        $errors.Add("$label uses the unapproved verb $($Matches.verb).")
    }
    if ($definition.alias) {
        $alias = $definition.alias
        $prefix = $aliasPrefixes | Where-Object { $alias.StartsWith($_, [StringComparison]::Ordinal) -and $alias.Length -gt $_.Length } | Select-Object -First 1
        if ($alias -match '-') { $errors.Add("$label alias $alias contains a verb.") }
        elseif (-not $prefix) { $errors.Add("$label alias $alias has no format prefix.") }
        if ($alias -in $builtInCommands) { $errors.Add("$label alias $alias collides with a built-in command.") }
        if ($alias -match 's$' -and $aliasOwners.ContainsKey($alias.Substring(0, $alias.Length - 1))) {
            $errors.Add("$label alias $alias is a plural twin of $($alias.Substring(0, $alias.Length - 1)).")
        }
    }
    $sets = @($definition.parameterSets)
    if ($sets.Count -eq 0) { $warnings.Add("$label has no parameter sets.") }
    if ($sets.Count -gt 1 -and @($sets | Where-Object default).Count -ne 1) {
        $errors.Add("$label needs exactly one default parameter set.")
    }
    foreach ($set in $sets) {
        foreach ($parameter in @($set.parameters)) {
            $pLabel = "$label -$($parameter.name)"
            if ($bannedParameters.ContainsKey($parameter.name)) { $errors.Add("$pLabel is not allowed. $($bannedParameters[$parameter.name])") }
            if ($parameter.name -cmatch $unitSuffix) { $errors.Add("$pLabel carries a unit suffix; use an OfficeLength parameter.") }
            if ($parameter.type -in 'bool', 'bool?', 'Boolean', 'Nullable<bool>') { $errors.Add("$pLabel is a boolean; use a switch.") }
        }
    }
}

$summary = [pscustomobject]@{
    MapFiles        = $files.Count
    CurrentCommands = $currentCommands.Count
    CurrentAliases  = $currentAliases.Count
    MapEntries      = $entries.Count
    TargetCommands  = $commands.Count
    TargetAliases   = $aliasOwners.Count
    Actions         = ($entries | Group-Object action | Sort-Object Name | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join ', '
    Errors          = $errors.Count
    Warnings        = $warnings.Count
}
$summary
if ($warnings.Count) { Write-Host "`nWarnings:"; $warnings | Sort-Object | ForEach-Object { Write-Host "  $_" } }
if ($errors.Count) {
    Write-Host "`nErrors:"
    $errors | Sort-Object | ForEach-Object { Write-Host "  $_" }
    exit 1
}
