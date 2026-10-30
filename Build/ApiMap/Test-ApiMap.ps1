<#
.SYNOPSIS
Validates the PowerShell API map against the current module and the 3.6 design rules.

.DESCRIPTION
Checks that every current PSWriteOffice command and alias is mapped exactly once, that every
target exists, and that target commands, parameters and aliases follow the naming rules in
Docs/officeimo.api-design.md. Run it with the module source built, or pass a metadata file.

.EXAMPLE
./Build/ApiMap/Test-ApiMap.ps1 -ModulePath ./PowerShell/PSWriteOffice/PSWriteOffice.psd1

.EXAMPLE
./Build/ApiMap/Test-ApiMap.ps1 -MetadataPath ./metadata.json -ScriptPath ./Build/ApiMap/drafts

Also checks that draft scripts use only target commands and parameters.
#>
[CmdletBinding()]
param(
    [string] $ModulePath,
    [string] $MetadataPath,
    [string] $MapDirectory = (Join-Path $PSScriptRoot 'powershell'),
    [string[]] $ScriptPath
)

$ErrorActionPreference = 'Stop'

if ($MetadataPath) {
    $metadata = Get-Content -LiteralPath $MetadataPath -Raw | ConvertFrom-Json
    if ($metadata.schemaVersion -eq 1 -and $metadata.commands) {
        # The generated website catalog records aliases on their owning commands.
        $currentCommands = @($metadata.commands | Where-Object kind -In 'Cmdlet', 'Function' | ForEach-Object name)
        $currentAliases = @($metadata.commands | ForEach-Object { $_.aliases } | Where-Object { $_ })
    } elseif ($metadata.Cmdlets) {
        $currentCommands = @($metadata.Cmdlets.Name)
        $currentAliases = @($metadata.Aliases.Alias | Where-Object { $_ })
    } else {
        throw 'Metadata must contain the generated command catalog or a Cmdlets/Aliases export.'
    }
    if ($currentCommands.Count -eq 0 -or @($currentCommands | Where-Object { [string]::IsNullOrWhiteSpace($_) }).Count) {
        throw 'Metadata contains no commands or an unnamed command.'
    }
} elseif ($ModulePath) {
    $module = Import-Module $ModulePath -Force -PassThru
    $currentCommands = @(Get-Command -Module $module.Name -CommandType Cmdlet, Function | ForEach-Object Name)
    $currentAliases = @(Get-Command -Module $module.Name -CommandType Alias | ForEach-Object Name)
} else {
    throw 'Specify -ModulePath or -MetadataPath.'
}

$approvedVerbs = @(Get-Verb | ForEach-Object Verb)
$builtInCommands = @(pwsh -NoProfile -NonInteractive -Command 'Get-Command -CommandType Cmdlet, Function, Alias | ForEach-Object Name')
$aliasPrefixes = 'Word', 'Excel', 'PowerPoint', 'Pdf', 'Markdown', 'Visio', 'Html', 'Rtf', 'OpenDocument', 'Csv', 'Email', 'Project', 'Access', 'Epub', 'OneNote', 'AsciiDoc', 'Latex', 'Confluence'
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

# Coverage includes both migration sources and already implemented target commands and aliases.
$byOld = $entries | Group-Object old
foreach ($group in $byOld | Where-Object Count -gt 1) {
    $errors.Add("$($group.Name) is mapped $($group.Count) times ($(@($group.Group.File) -join ', ')).")
}
$mapped = @{}
foreach ($entry in $entries) { $mapped[$entry.old] = $entry }
foreach ($name in $currentCommands) {
    if (-not $mapped.ContainsKey($name) -and -not $commands.ContainsKey($name)) { $errors.Add("Command $name has no map entry.") }
}
foreach ($name in $currentAliases) {
    if (-not $mapped.ContainsKey($name) -and -not $aliasOwners.ContainsKey($name)) { $errors.Add("Alias $name has no map entry.") }
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

# Examples and draft scripts: every OfficeIMO command and parameter they use must exist in the target surface.
$aliasTargets = @{}
foreach ($alias in $aliasOwners.Keys) { $aliasTargets[$alias] = $aliasOwners[$alias] }
$commonParameters = [System.Management.Automation.PSCmdlet]::CommonParameters + [System.Management.Automation.PSCmdlet]::OptionalCommonParameters
$testCode = {
    param([string] $Code, [string] $Label)
    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseInput($Code, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count) { $errors.Add("$Label does not parse: $($parseErrors[0].Message)"); return }
    foreach ($commandAst in $ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.CommandAst] }, $true)) {
        $used = $commandAst.GetCommandName()
        if (-not $used) { continue }
        $resolved = if ($aliasTargets.ContainsKey($used)) { $aliasTargets[$used] } else { $used }
        if (-not $commands.ContainsKey($resolved)) {
            if ($used -match '-Office' -or $used -cmatch '^(Word|Excel|Ppt|Pdf|Markdown|Visio|Rtf|OpenDocument|Epub)[A-Z]') {
                $errors.Add("$Label uses $used, which is not a target command or alias.")
            }
            continue
        }
        $known = @($commands[$resolved].Definition.parameterSets | ForEach-Object { @($_.parameters) } | ForEach-Object { $_.name; @($_.aliases) }) + $commonParameters + 'WhatIf', 'Confirm'
        foreach ($element in $commandAst.CommandElements) {
            if ($element -isnot [System.Management.Automation.Language.CommandParameterAst]) { continue }
            $matchesParameter = @($known | Where-Object { $_ -and $_.StartsWith($element.ParameterName, [StringComparison]::OrdinalIgnoreCase) })
            if ($matchesParameter.Count -eq 0) {
                $errors.Add("$Label passes -$($element.ParameterName) to $resolved, which does not define it.")
            }
        }

        # Parameter set resolution, as PowerShell does it: some set must accept every named parameter,
        # the positional arguments and all of its mandatory parameters. Splatted calls can't be checked.
        if (@($commandAst.CommandElements | Where-Object { $_ -is [System.Management.Automation.Language.VariableExpressionAst] -and $_.Splatted }).Count) { continue }
        $definition = $commands[$resolved].Definition
        $byName = @{}
        foreach ($set in @($definition.parameterSets)) {
            foreach ($parameter in @($set.parameters)) {
                $byName[$parameter.name] = $parameter
                foreach ($alias in @($parameter.aliases | Where-Object { $_ })) { $byName[$alias] = $parameter }
            }
        }
        $pipeline = $commandAst.Parent
        $receivesPipeline = $pipeline -is [System.Management.Automation.Language.PipelineAst] -and $pipeline.PipelineElements.IndexOf($commandAst) -gt 0
        $named = @{}
        $positional = [System.Collections.Generic.List[object]]::new()
        $elements = $commandAst.CommandElements
        for ($i = 1; $i -lt $elements.Count; $i++) {
            $element = $elements[$i]
            if ($element -isnot [System.Management.Automation.Language.CommandParameterAst]) { $positional.Add($element); continue }
            $candidates = @($byName.Keys | Where-Object { $_.Equals($element.ParameterName, [StringComparison]::OrdinalIgnoreCase) })
            if (-not $candidates.Count) { $candidates = @($byName.Keys | Where-Object { $_.StartsWith($element.ParameterName, [StringComparison]::OrdinalIgnoreCase) }) }
            $targets = @($candidates | ForEach-Object { $byName[$_].name } | Sort-Object -Unique)
            if ($targets.Count -gt 1) { $errors.Add("$Label passes -$($element.ParameterName) to $resolved, which matches several parameters: $($targets -join ', ').") }
            if ($targets.Count -ne 1) {
                if (-not $element.Argument -and $i + 1 -lt $elements.Count -and $elements[$i + 1] -isnot [System.Management.Automation.Language.CommandParameterAst] -and $element.ParameterName -notin 'Verbose', 'Debug', 'WhatIf', 'Confirm') { $i++ }
                continue
            }
            $parameter = $byName[$targets[0]]
            $isSwitch = $parameter.type -in 'switch', 'SwitchParameter'
            if (-not $element.Argument -and -not $isSwitch -and $i + 1 -lt $elements.Count -and $elements[$i + 1] -isnot [System.Management.Automation.Language.CommandParameterAst]) { $i++ }
            $named[$parameter.name] = $true
        }
        $viable = @()
        foreach ($set in @($definition.parameterSets)) {
            $setParameters = @($set.parameters)
            $setNames = @($setParameters | ForEach-Object { $_.name })
            if (@($named.Keys | Where-Object { $_ -notin $setNames }).Count) { continue }
            $free = @($setParameters | Where-Object { $null -ne $_.position -and -not $named.ContainsKey($_.name) } | Sort-Object -Property position)
            $blocks = [System.Collections.Generic.List[object]]::new()
            foreach ($candidate in @($setParameters | Where-Object { $_.type -eq 'scriptblock' -and $null -eq $_.position -and -not $named.ContainsKey($_.name) })) { $blocks.Add($candidate) }
            $bound = @{}
            $next = 0
            $tooMany = $false
            foreach ($argument in $positional) {
                if ($next -lt $free.Count) { $bound[$free[$next].name] = $true; $next++ }
                elseif ($argument -is [System.Management.Automation.Language.ScriptBlockExpressionAst] -and $blocks.Count) { $bound[$blocks[0].name] = $true; $blocks.RemoveAt(0) }
                else { $tooMany = $true }
            }
            if ($tooMany) { continue }
            $missing = @($setParameters | Where-Object {
                $_.mandatory -and -not $named.ContainsKey($_.name) -and -not $bound.ContainsKey($_.name) -and
                -not ($receivesPipeline -and $_.pipeline -and $_.pipeline -ne 'none')
            })
            if ($missing.Count) { continue }
            $viable += $set
        }
        if (-not $viable.Count) {
            $errors.Add("$Label calls $resolved with arguments no parameter set accepts (check mandatory parameters and parameters from different sets).")
        } elseif ($viable.Count -gt 1 -and -not $receivesPipeline -and -not @($viable | Where-Object { $_.default }).Count) {
            $errors.Add("$Label calls $resolved ambiguously: parameter sets $(@($viable | ForEach-Object { $_.name }) -join ', ') all accept it and none is the default.")
        }
    }
}

$exampleCount = 0
$commandsWithTooFewExamples = 0
foreach ($name in $commands.Keys) {
    $examples = @($commands[$name].Definition.examples | Where-Object { $_ })
    if ($examples.Count -lt 2) { $commandsWithTooFewExamples++ }
    foreach ($example in $examples) {
        $exampleCount++
        . $testCode -Code ([string]$example.code) -Label "$($commands[$name].File): $name example '$($example.title)'"
    }
}
foreach ($script in @($ScriptPath | Where-Object { $_ } | ForEach-Object { Get-ChildItem -Path $_ -Recurse -Include '*.ps1', '*.psm1' })) {
    . $testCode -Code (Get-Content -LiteralPath $script.FullName -Raw) -Label $script.Name
}

$summary = [pscustomobject]@{
    MapFiles        = $files.Count
    CurrentCommands = $currentCommands.Count
    CurrentAliases  = $currentAliases.Count
    MapEntries      = $entries.Count
    TargetCommands  = $commands.Count
    TargetAliases   = $aliasOwners.Count
    Actions         = ($entries | Group-Object action | Sort-Object Name | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join ', '
    Examples        = $exampleCount
    CommandsUnder2Examples = $commandsWithTooFewExamples
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
