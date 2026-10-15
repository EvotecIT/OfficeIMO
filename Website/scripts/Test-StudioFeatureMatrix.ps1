#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Verifies that the Studio feature matrix is complete and that its roadmap links resolve.

.DESCRIPTION
    data/studio_features.json is the public record of what OfficeIMO Studio does on each system. Every feature states
    its support on every platform; partial support explains its limit; planned support points to an existing section
    of Docs/ROADMAP.md, which remains the single backlog.
#>

[CmdletBinding()]
param(
    [string] $SiteRoot = (Split-Path -Parent $PSScriptRoot),
    [string] $RepositoryRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = 'Stop'
$matrixPath = Join-Path (Resolve-Path -LiteralPath $SiteRoot).Path 'data/studio_features.json'
$roadmapPath = Join-Path (Resolve-Path -LiteralPath $RepositoryRoot).Path 'Docs/ROADMAP.md'
$matrix = Get-Content -LiteralPath $matrixPath -Raw | ConvertFrom-Json
$problems = [System.Collections.Generic.List[string]]::new()

# GitHub heading anchors: lower case, punctuation other than hyphens removed, spaces as hyphens.
$anchors = [System.Collections.Generic.HashSet[string]]::new()
foreach ($line in Get-Content -LiteralPath $roadmapPath) {
    if ($line -match '^#{1,6}\s+(.+?)\s*$') {
        $anchor = ($Matches[1].ToLowerInvariant() -replace '[^\p{L}\p{Nd}\s_-]', '') -replace '\s', '-'
        [void] $anchors.Add($anchor)
    }
}

$appIds = @($matrix.apps.id)
$platformIds = @($matrix.platforms.id)
$stateIds = @($matrix.states.id)
foreach ($platform in $matrix.platforms) {
    if ($platform.app -notin $appIds) { $problems.Add("Platform '$($platform.id)' names unknown app '$($platform.app)'.") }
}

$featureIds = [System.Collections.Generic.HashSet[string]]::new()
foreach ($group in $matrix.groups) {
    foreach ($feature in $group.features) {
        if (-not $featureIds.Add([string] $feature.id)) { $problems.Add("Feature id '$($feature.id)' is used twice.") }
        $states = @{}
        foreach ($property in $feature.support.PSObject.Properties) { $states[$property.Name] = [string] $property.Value }
        foreach ($platformId in $states.Keys) {
            if ($platformId -notin $platformIds) { $problems.Add("Feature '$($feature.id)' names unknown platform '$platformId'.") }
        }
        foreach ($platformId in $platformIds) {
            $state = $states[$platformId]
            if (-not $state) { $problems.Add("Feature '$($feature.id)' has no support state for '$platformId'."); continue }
            if ($state -notin $stateIds) { $problems.Add("Feature '$($feature.id)' uses unknown state '$state' on '$platformId'.") }
            $note = if ($feature.notes) { [string] $feature.notes.$platformId } else { '' }
            if ($state -eq 'partial' -and -not $note) { $problems.Add("Feature '$($feature.id)' is partial on '$platformId' without a note describing the limit.") }
            if ($state -eq 'planned' -and -not $feature.roadmap) { $problems.Add("Feature '$($feature.id)' is planned on '$platformId' without a roadmap section.") }
        }
        if ($feature.roadmap -and -not $anchors.Contains([string] $feature.roadmap)) {
            $problems.Add("Feature '$($feature.id)' links to roadmap section '#$($feature.roadmap)', which Docs/ROADMAP.md does not have.")
        }
    }
}

foreach ($milestone in $matrix.roadmap) {
    if (-not $anchors.Contains([string] $milestone.anchor)) {
        $problems.Add("Roadmap milestone '$($milestone.title)' links to '#$($milestone.anchor)', which Docs/ROADMAP.md does not have.")
    }
    foreach ($platformId in $milestone.platforms) {
        if ($platformId -notin $platformIds) { $problems.Add("Roadmap milestone '$($milestone.title)' names unknown platform '$platformId'.") }
    }
}

if ($problems.Count -gt 0) {
    $problems | ForEach-Object { Write-Host "Studio feature matrix: $_" }
    throw "The Studio feature matrix has $($problems.Count) problem(s)."
}
Write-Host "Studio feature matrix: $($featureIds.Count) features on $($platformIds.Count) platforms verified."
