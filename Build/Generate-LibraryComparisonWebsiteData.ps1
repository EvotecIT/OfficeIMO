#requires -Version 7.4
<#
.SYNOPSIS
Validates and copies the published comparison catalog for build-time templates,
and writes a compact summary of every published result for the area pages.
.DESCRIPTION
The catalog is the only evidence input. This does not measure workloads, promote
quick results, fill missing environments, or update the legacy Excel matrix.
The summary keeps every library, case, platform and run mode from each result;
it drops only the raw samples and the BenchmarkDotNet job settings, which are the
same for every row and are still in the linked result JSON.
.PARAMETER CatalogPath
Merged PowerForge v4 index.json to validate and copy.
.PARAMETER WebsiteIndexPath
Build-time template data file receiving the exact published catalog bytes.
.PARAMETER SummaryPath
Summary of all published results, read by the benchmark area pages.
.PARAMETER DryRun
Validates the catalog and returns output paths without writing files.
.EXAMPLE
./Build/Generate-LibraryComparisonWebsiteData.ps1 -CatalogPath ./Website/static/data/benchmarks/library-comparisons/index.json
.OUTPUTS
An output plan in DryRun mode. Otherwise writes an identical catalog copy.
#>
[CmdletBinding()]
param(
    [string] $CatalogPath = (Join-Path $PSScriptRoot '../Website/static/data/benchmarks/library-comparisons/index.json'),
    [string] $WebsiteIndexPath = (Join-Path $PSScriptRoot '../Website/data/benchmarks_index.json'),
    [string] $SummaryPath = (Join-Path $PSScriptRoot '../Website/static/data/benchmarks/library-comparisons/summary.json'),
    [switch] $DryRun
)

$ErrorActionPreference = 'Stop'
$CatalogPath = (Resolve-Path -LiteralPath $CatalogPath -ErrorAction Stop).ProviderPath
$WebsiteIndexPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($WebsiteIndexPath)
$SummaryPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($SummaryPath)
$catalogBytes = [System.IO.File]::ReadAllBytes($CatalogPath)
$json = [System.Text.Encoding]::UTF8.GetString($catalogBytes).TrimStart([char] 0xFEFF)
if (-not (Test-Json -Json $json -SchemaFile (Join-Path $PSScriptRoot 'library-comparisons.schema.json'))) {
    throw "Invalid v4 comparison catalog: '$CatalogPath'."
}
$catalog = $json | ConvertFrom-Json -Depth 100
if ($DryRun) {
    [pscustomobject]@{
        DryRun = $true
        CatalogPath = $CatalogPath
        OutputPath = @($WebsiteIndexPath, $SummaryPath)
        ComparisonCount = @($catalog.comparisons).Count
        EntryCount = @($catalog.entries).Count
    }
    return
}

New-Item -ItemType Directory -Path (Split-Path -Parent $WebsiteIndexPath) -Force | Out-Null
# Retain the source timestamp strings and numeric representation on every locale.
[System.IO.File]::WriteAllBytes($WebsiteIndexPath, $catalogBytes)
Write-Host "Copied comparison catalog to '$WebsiteIndexPath'."

# BenchmarkDotNet job settings and statistic columns. They describe how a run was configured, not which case
# a row measures, so keeping them would split one comparison into one table per row.
$jobColumns = [System.Collections.Generic.HashSet[string]]::new([string[]]@(
    'Namespace', 'Type', 'FullName', 'BenchmarkDotNetReport', 'Method', 'Job',
    'AnalyzeLaunchVariance', 'EvaluateOverhead', 'MaxAbsoluteError', 'MaxRelativeError', 'MinInvokeCount',
    'MinIterationTime', 'OutlierMode', 'Affinity', 'EnvironmentVariables', 'Jit', 'LargeAddressAware', 'Platform',
    'PowerPlanMode', 'Runtime', 'AllowVeryLargeObjects', 'Concurrent', 'CpuGroups', 'Force', 'HeapAffinitizeMask',
    'HeapCount', 'NoAffinitize', 'RetainVm', 'Server', 'Arguments', 'BuildConfiguration', 'Clock', 'EngineFactory',
    'NuGetReferences', 'Toolchain', 'IsMutator', 'InvocationCount', 'IterationCount', 'IterationTime', 'LaunchCount',
    'MaxIterationCount', 'MaxWarmupIterationCount', 'MemoryRandomization', 'MinIterationCount',
    'MinWarmupIterationCount', 'RunStrategy', 'UnrollFactor', 'WarmupCount',
    'Error', 'StdDev', 'StdErr', 'Median', 'Mean', 'Ratio', 'RatioSD', 'Alloc Ratio', 'Allocated', 'Gen0', 'Gen1',
    'Gen2', 'Code Size', 'Completed Work Items', 'Lock Contentions'
), [System.StringComparer]::OrdinalIgnoreCase)

function Get-CaseParameters {
    param($Variables)
    $parameters = [ordered]@{}
    if (-not $Variables) { return $parameters }
    foreach ($property in $Variables.PSObject.Properties) {
        if ($jobColumns.Contains($property.Name)) { continue }
        # BenchmarkDotNet flattens several parameters into one value: "Easy&Producer=IText&Workflow=Bundles".
        $parts = ([string] $property.Value) -split '&'
        $parameters[$property.Name] = $parts[0]
        foreach ($part in $parts | Select-Object -Skip 1) {
            $pair = $part -split '=', 2
            if ($pair.Count -eq 2 -and -not $jobColumns.Contains($pair[0])) { $parameters[$pair[0]] = $pair[1] }
        }
    }
    $parameters
}

function Get-Number {
    param($Value)
    if ($Value -is [double] -or $Value -is [long] -or $Value -is [int] -or $Value -is [decimal]) { return [double] $Value }
    $null
}

$comparisons = @{}
foreach ($comparison in @($catalog.comparisons)) { $comparisons[$comparison.comparisonId] = $comparison }
# Published results sit beside the catalog.
$catalogDirectory = Split-Path -Parent $CatalogPath
$runs = foreach ($entry in @($catalog.entries)) {
    $resultFile = Join-Path $catalogDirectory (Split-Path -Leaf $entry.resultPath)
    if (-not (Test-Path -LiteralPath $resultFile)) {
        throw "Published result '$($entry.resultPath)' is missing from '$resultFile'."
    }
    $result = Get-Content -LiteralPath $resultFile -Raw -Encoding UTF8 | ConvertFrom-Json -Depth 100
    $metadata = $result.metadata
    $versions = [ordered]@{}
    if ($metadata) {
        foreach ($property in $metadata.PSObject.Properties) {
            if ($property.Name -match '^benchmark\.dependency\.([^.]+)\.version$') { $versions[$Matches[1]] = [string] $property.Value }
        }
    }
    $environment = $result.environment
    $rows = foreach ($row in @($result.summary)) {
        $metrics = $row.metrics
        $allocated = $null
        if ($metrics) {
            $allocated = Get-Number $metrics.BytesAllocatedPerOperation
            if ($null -eq $allocated) { $allocated = Get-Number $metrics.Allocated }
        }
        $time = Get-Number $row.medianMs
        $statistic = 'median'
        if ($null -eq $time) { $time = Get-Number $row.meanMs; $statistic = 'mean' }
        [ordered]@{
            library = [string] $row.scenario
            case = Get-CaseParameters $row.variables
            timeMs = $time
            statistic = $statistic
            allocatedBytes = $allocated
            status = [string] $row.status
        }
    }
    [ordered]@{
        comparisonId = $entry.comparisonId
        title = [string] $comparisons[$entry.comparisonId].title
        libraries = @($comparisons[$entry.comparisonId].libraries)
        area = $entry.area
        platform = $entry.platform
        runMode = $entry.runMode
        publish = [bool] $entry.publish
        generatedUtc = $entry.generatedUtc
        resultPath = $entry.resultPath
        sourceCommit = if ($metadata) { [string] ($metadata.gitSha ?? $metadata.'benchmark.workload.sourceCommit') } else { '' }
        environment = [ordered]@{
            processor = [string] $environment.processorName
            os = ([string] $environment.osDescription) -replace '\s*\[.*$', ''
            runtime = ([string] $environment.runtimeVersion) -replace '\s*\(.*$', ''
            cores = $environment.logicalCoreCount
            affinity = if ($metadata) { [string] $metadata.'benchmark.workload.affinityMask' } else { '' }
        }
        versions = $versions
        rows = @($rows)
    }
}
$summary = [ordered]@{ schemaVersion = 1; runs = @($runs) }
New-Item -ItemType Directory -Path (Split-Path -Parent $SummaryPath) -Force | Out-Null
$summaryJson = $summary | ConvertTo-Json -Depth 20 -Compress
[System.IO.File]::WriteAllText($SummaryPath, $summaryJson + "`n", [System.Text.UTF8Encoding]::new($false))
Write-Host "Wrote $(@($runs).Count) result summaries to '$SummaryPath'."
