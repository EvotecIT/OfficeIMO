param(
    [Parameter(Mandatory)][string] $Binary,
    [Parameter(Mandatory)][string] $Assets,
    [Parameter(Mandatory)][string] $OutputRoot,
    [int[]] $Rows = @(10000,100000),
    [int[]] $Columns = @(20),
    [string[]] $Browsers = @('Chromium','Firefox','WebKit'),
    [string[]] $Stacks = @('bundled','current'),
    [string[]] $Formats = @('xlsx','csv'),
    [ValidateSet('native','compatibility','batched')][string[]] $Lanes = @('native','compatibility','batched'),
    [switch] $Unique,
    [int] $WarmupCount = 1,
    [int] $IterationCount = 3,
    [UInt64] $ProcessorAffinityMask = 0,
    [switch] $Plan
)
$ErrorActionPreference = 'Stop'
Import-Module PSPublishModule -ErrorAction Stop
$repository = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
Add-Type -Path (Join-Path $PSScriptRoot 'DataTablesBenchmarkClient.cs')
$processorPolicy = @{}
if ($ProcessorAffinityMask -ne 0) { $processorPolicy.ProcessorAffinityMask = $ProcessorAffinityMask }
try {
    $result = Invoke-BenchmarkSuite -Path (Join-Path $PSScriptRoot 'datatables.benchmark.ps1') -OutputRoot $OutputRoot `
        -Variable @{ Binary = (Resolve-Path -LiteralPath $Binary).Path; Repository = $repository; Assets = (Resolve-Path -LiteralPath $Assets).Path; Evidence = [IO.Path]::GetFullPath($OutputRoot); Rows = $Rows; Columns = $Columns; Browsers = ($Browsers -join ','); Stacks = ($Stacks -join ','); Formats = ($Formats -join ','); Unique = [bool]$Unique } `
        -WarmupCount $WarmupCount -IterationCount $IterationCount -Engine $Lanes @processorPolicy -Plan:$Plan
    $result
    if (-not $Plan -and @($result.Summary | Where-Object Status -ne 'Succeeded').Count -gt 0) { throw 'An export comparison failed. Inspect the retained summary and validation evidence.' }
} finally {
    if ($null -ne [DataTablesBenchmarkClient]::Current) { [DataTablesBenchmarkClient]::Current.Dispose(); [DataTablesBenchmarkClient]::Current = $null }
}
