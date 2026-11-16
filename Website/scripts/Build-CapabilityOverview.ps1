#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Builds the capability overview: what each document job can do on each OfficeIMO surface, and where to read more.

.DESCRIPTION
    data/capability_jobs.json names, for each job, the packages, PowerShell commands, MCP tools and Studio features
    that do it. This script resolves every name against its owning machine-readable source and writes
    data/capability_overview.json, which the capabilities page renders:

    - .NET packages from data/documentation_catalog.json (with each package's guide URL)
    - PowerShell commands from the PSWriteOffice command metadata that the pipeline syncs before this step, the same
      inventory the PowerShell reference pages are built from
    - MCP tools and agent skills from data/agent_integrations.json
    - Studio support on each system from data/studio_features.json
    - conversion route counts per surface from data/format_map.json

    A name that its source does not know fails the build, so the overview cannot claim a capability that was renamed
    or removed. Run with -Check to verify that the committed overview is current without rewriting it.
#>

[CmdletBinding()]
param(
    [string] $SiteRoot = (Split-Path -Parent $PSScriptRoot),
    [string] $RepositoryRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)),
    [switch] $Check
)

$ErrorActionPreference = 'Stop'
$siteRoot = (Resolve-Path -LiteralPath $SiteRoot).Path
$repositoryRoot = (Resolve-Path -LiteralPath $RepositoryRoot).Path

function Read-JsonFile([string] $Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Required JSON file was not found: $Path" }
    Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
}

$sources = [ordered]@{
    jobs       = 'Website/data/capability_jobs.json'
    packages   = 'Website/data/documentation_catalog.json'
    powershell = 'Website/data/apidocs/powershell/command-metadata.json'
    agents     = 'Website/data/agent_integrations.json'
    studio     = 'Website/data/studio_features.json'
    formats    = 'Website/data/format_map.json'
}
$data = @{}
foreach ($key in $sources.Keys) { $data[$key] = Read-JsonFile (Join-Path $repositoryRoot $sources[$key]) }

$problems = [System.Collections.Generic.List[string]]::new()

$packages = @{}
foreach ($component in $data.packages.components) {
    if ($component.kind -eq 'package' -and $component.packageId) { $packages[[string] $component.packageId] = $component }
}
$commands = @{}
foreach ($command in $data.powershell.commands) {
    if ($command.kind -ne 'Alias') { $commands[[string] $command.name] = $command }
}
$tools = @{}
foreach ($tool in $data.agents.tools) { $tools[[string] $tool.name] = $tool }
$features = @{}
foreach ($group in $data.studio.groups) {
    foreach ($feature in $group.features) { $features[[string] $feature.id] = $feature }
}
$stateRank = @{ 'available' = 3; 'partial' = 2; 'planned' = 1; 'not-applicable' = 0 }
$stateLabels = @{}
foreach ($state in $data.studio.states) { $stateLabels[[string] $state.id] = [string] $state.label }

$routeCounts = @{}
foreach ($route in $data.formats.routes) {
    foreach ($surface in $route.surfaces) { $routeCounts[[string] $surface] = 1 + [int] $routeCounts[[string] $surface] }
}

function New-Cell([string] $Guide, [object[]] $Items, [string] $Extra) {
    [ordered]@{ guide = $Guide; extra = $Extra; items = @($Items) }
}

$jobs = foreach ($job in $data.jobs.jobs) {
    $cells = [ordered]@{}

    $items = foreach ($id in $job.dotnet.packages) {
        if (-not $packages.ContainsKey([string] $id)) { $problems.Add("Job '$($job.id)' names package '$id', which the documentation catalog does not list."); continue }
        [ordered]@{ label = [string] $id; url = [string] $packages[$id].docsUrl }
    }
    $guide = if ($job.dotnet.guide) { [string] $job.dotnet.guide } else { '/libraries/' }
    $extra = if ($job.routes) { "$($routeCounts['dotnet']) conversion routes" } else { '' }
    $cells.dotnet = New-Cell $guide $items $extra

    $items = foreach ($name in $job.powershell.commands) {
        if (-not $commands.ContainsKey([string] $name)) { $problems.Add("Job '$($job.id)' names command '$name', which the website's PSWriteOffice snapshot does not list. Run scripts/Sync-PSWriteOfficeApiDocs.ps1 if the module exports it."); continue }
        [ordered]@{ label = [string] $name; url = "/api/powershell/$(([string] $name).ToLowerInvariant())/" }
    }
    $extra = if ($job.routes) { "$($routeCounts['powershell']) conversion routes" } else { '' }
    $cells.powershell = New-Cell ([string] $job.powershell.guide) $items $extra

    $items = @(foreach ($name in $job.agents.tools) {
        if (-not $tools.ContainsKey([string] $name)) { $problems.Add("Job '$($job.id)' names MCP tool '$name', which the agents page does not list."); continue }
        [ordered]@{ label = [string] $name; url = "/agents/#tool-$name" }
    })
    if ($job.agents.skills) {
        foreach ($skill in $data.agents.skills) { $items += [ordered]@{ label = "$($skill.name) skill"; url = '/agents/#tools' } }
    }
    $extra = if ($job.routes) { "$($routeCounts['cli']) routes in the CLI" } else { '' }
    $cells.agents = New-Cell ([string] $job.agents.guide) $items $extra

    $studioFeatures = foreach ($id in $job.studio.features) {
        if (-not $features.ContainsKey([string] $id)) { $problems.Add("Job '$($job.id)' names Studio feature '$id', which the Studio feature matrix does not have."); continue }
        $features[$id]
    }
    $platforms = foreach ($platform in $data.studio.platforms) {
        $best = 'not-applicable'
        foreach ($feature in $studioFeatures) {
            $state = [string] $feature.support.($platform.id)
            if ($stateRank[$state] -gt $stateRank[$best]) { $best = $state }
        }
        $label = [string] $platform.label
        $short = if ($label.Length -gt 6) { $label.Substring(0, 3) } else { $label }
        [ordered]@{ id = [string] $platform.id; label = $label; short = $short; state = $best; stateLabel = $stateLabels[$best] }
    }
    $items = foreach ($feature in $studioFeatures) {
        [ordered]@{ label = [string] $feature.title; url = "/studio/#feature-$($feature.id)" }
    }
    $guide = if ($items) { [string] @($items)[0].url } else { '/studio/#features' }
    $cell = New-Cell $guide $items ''
    $cell.platforms = @($platforms)
    $cells.studio = $cell

    foreach ($surface in 'dotnet', 'powershell', 'agents', 'studio') {
        if (-not $cells[$surface].guide) { $problems.Add("Job '$($job.id)' has no guide for '$surface'.") }
    }

    [ordered]@{ id = [string] $job.id; title = [string] $job.title; summary = [string] $job.summary; cells = $cells }
}

if ($problems.Count -gt 0) {
    $problems | ForEach-Object { Write-Host "Capability overview: $_" }
    throw "The capability overview has $($problems.Count) problem(s)."
}

$overview = [ordered]@{
    schemaVersion = 1
    sources       = @($sources.Values)
    summary       = [ordered]@{
        packageCount  = $packages.Count
        commandCount  = $commands.Count
        toolCount     = $tools.Count
        skillCount    = @($data.agents.skills).Count
        studioFeatures = $features.Count
        routeCount    = @($data.formats.routes).Count
    }
    surfaces      = @($data.jobs.surfaces)
    jobs          = @($jobs)
}

$json = ($overview | ConvertTo-Json -Depth 12) + [Environment]::NewLine
$outputPath = Join-Path $siteRoot 'data/capability_overview.json'
if ($Check) {
    $current = if (Test-Path -LiteralPath $outputPath) { [IO.File]::ReadAllText($outputPath) } else { '' }
    if (($current -replace "`r`n", "`n") -ne ($json -replace "`r`n", "`n")) {
        throw "data/capability_overview.json is out of date. Run Website/scripts/Build-CapabilityOverview.ps1."
    }
    Write-Host "Capability overview: $(@($jobs).Count) jobs are current."
    return
}
[IO.File]::WriteAllText($outputPath, $json, [Text.UTF8Encoding]::new($false))
Write-Host "Capability overview: $(@($jobs).Count) jobs written to data/capability_overview.json."
