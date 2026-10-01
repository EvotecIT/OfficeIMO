param(
    [string] $BinaryRoot = (Join-Path $PSScriptRoot '../../OfficeIMO.CSV/bin/Release/net10.0'),
    [string] $OutputRoot = (Join-Path $PSScriptRoot '../../Ignore/Benchmarks/CsvSustainedRead'),
    [string] $ModulePath = 'PSPublishModule',
    [int[]] $Rows = @(100000, 1000000),
    [ValidateRange(1, 1024)] [int] $Degree = 4,
    [ValidateRange(1, 65536)] [int] $Batch = 1024,
    [int] $WarmupCount = 2,
    [int] $IterationCount = 7,
    [ValidateSet('FirstRow', 'AllRowsAsync', 'TypedSequential', 'TypedParallel')]
    [string[]] $Operation = @('FirstRow', 'AllRowsAsync', 'TypedSequential', 'TypedParallel'),
    [ValidateSet('Snapshot', 'Incremental')]
    [string[]] $Engine = @('Snapshot', 'Incremental'),
    [ValidateSet('Plain', 'Multiline')]
    [string[]] $Shape = @('Plain', 'Multiline'),
    [string] $AffinityMask,
    [switch] $SampleMemory,
    [switch] $Plan
)
$ErrorActionPreference = 'Stop'
if ([Environment]::Version.Major -lt 10) { throw 'Use a fresh PowerShell process on .NET 10 or newer.' }
if (@($Rows | Where-Object { $_ -lt 1 }).Count -gt 0) { throw 'Rows must contain positive counts.' }
$BinaryRoot = (Resolve-Path -LiteralPath $BinaryRoot).Path
$OutputRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputRoot)
# Load the exact library before importing a tool that might carry another OfficeIMO version.
[void] [Reflection.Assembly]::LoadFrom((Join-Path $BinaryRoot 'OfficeIMO.Core.dll'))
Import-Module $ModulePath -ErrorAction Stop
if (-not ('PowerForge.BenchmarkMemoryProbe' -as [type])) {
    throw 'This lane requires PowerForge BenchmarkMemoryProbe. Use a build of its owning source until a containing package is published.'
}
$process = [Diagnostics.Process]::GetCurrentProcess()
$previousPriority = $process.PriorityClass
$previousAffinity = $null
$fixtureRoot = Join-Path $OutputRoot ('fixtures-' + [Guid]::NewGuid().ToString('N'))
try {
    if ($AffinityMask) {
        if (-not ([OperatingSystem]::IsWindows() -or [OperatingSystem]::IsLinux())) {
            throw 'Processor-affinity qualification requires Windows or Linux.'
        }
        $previousAffinity = $process.ProcessorAffinity
        $mask = if ($AffinityMask.StartsWith('0x')) { [Convert]::ToInt64($AffinityMask.Substring(2), 16) } else { [long] $AffinityMask }
        if ($mask -eq 0) { throw 'AffinityMask must be nonzero.' }
        $process.ProcessorAffinity = [IntPtr] $mask
        if ($process.ProcessorAffinity.ToInt64() -ne $mask) { throw 'Processor affinity was not applied exactly.' }
    }
    $process.PriorityClass = [Diagnostics.ProcessPriorityClass]::Normal
    if (-not $Plan) { [void] (New-Item -ItemType Directory -Path $fixtureRoot) }
    $caseNames = @(foreach ($count in $Rows) { foreach ($notes in $Shape) { "Rows-$count-$notes" } })
    $result = Invoke-BenchmarkSuite -Path (Join-Path $PSScriptRoot 'csv-sustained-read.benchmark.ps1') `
        -OutputRoot $OutputRoot -Variable @{ BinaryRoot = $BinaryRoot; FixtureRoot = $fixtureRoot; Rows = $Rows; Degree = $Degree; Batch = $Batch; SampleMemory = [bool] $SampleMemory; SelectedEngineSetup = ($Engine.Count -eq 1) } `
        -WarmupCount $WarmupCount -IterationCount $IterationCount -Operation $Operation -Engine $Engine -Case $caseNames -Plan:$Plan
    $result
    if (-not $Plan -and @($result.Samples | Where-Object Status -ne 'Succeeded').Count -gt 0) {
        throw 'CSV sustained-read qualification failed. Inspect the retained result artifacts.'
    }
}
finally {
    $process.PriorityClass = $previousPriority
    if ($null -ne $previousAffinity) { $process.ProcessorAffinity = $previousAffinity }
    $process.Dispose()
    if (-not $Plan) { Write-Host "Generated fixtures retained for scoped cleanup: $fixtureRoot" }
}
