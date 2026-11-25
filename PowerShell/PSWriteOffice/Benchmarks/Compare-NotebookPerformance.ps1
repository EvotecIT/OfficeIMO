[CmdletBinding()]
param(
    [ValidateSet('Warm', 'Cold', 'Batch')]
    [string] $Mode = 'Warm',

    [ValidateSet(12, 120, 1200)]
    [int[]] $CellCount = @(12, 120, 1200),

    [ValidateSet('ReadCells', 'Markdown', 'Html', 'StartupMarkdown', 'BatchMarkdown')]
    [string[]] $Operation,

    [ValidateRange(1, 100)]
    [int] $IterationCount = 5,

    [ValidateRange(0, 100)]
    [int] $WarmupCount = 2,

    [ValidateRange(1, 64)]
    [int] $BatchSize = 8,

    [ValidateRange(1, 16)]
    [int] $ThrottleLimit = 4,

    [ValidateRange(0, 1000)]
    [int] $MemorySamplingIntervalMilliseconds = 0,

    [string] $ModuleManifest = (Join-Path $PSScriptRoot '../PSWriteOffice.psd1'),

    [string] $PowerShellNotebookManifest,

    [string] $PowerForge = 'PSPublishModule',

    [string] $OutputRoot = (Join-Path $PSScriptRoot '../Ignore/Benchmarks/Notebook'),

    [switch] $Plan
)

$ErrorActionPreference = 'Stop'
$operations = @{
    Warm  = @('ReadCells', 'Markdown', 'Html')
    Cold  = @('StartupMarkdown')
    Batch = @('BatchMarkdown')
}
if (-not $Operation) {
    $Operation = $operations[$Mode]
}
foreach ($name in $Operation) {
    if ($name -notin $operations[$Mode]) {
        throw "Operation '$name' does not belong to the $Mode run."
    }
}
if ($Mode -eq 'Cold') {
    $WarmupCount = 0
}
if ($Mode -eq 'Batch' -and $PowerShellNotebookManifest) {
    throw 'Batch measures the PSWriteOffice command. Omit PowerShellNotebookManifest.'
}

$engines = @('PSWriteOffice')
if ($PowerShellNotebookManifest) {
    if ($PSVersionTable.PSVersion -lt [version]'6.2') {
        throw 'PowerShellNotebook 3.0.0 requires PowerShell 6.2 or later.'
    }
    $PowerShellNotebookManifest = (Resolve-Path -LiteralPath $PowerShellNotebookManifest).ProviderPath
    $manifest = Import-PowerShellDataFile -LiteralPath $PowerShellNotebookManifest
    $license = Join-Path (Split-Path $PowerShellNotebookManifest) 'LICENSE'
    $pinnedLicense = '32504C072766A8A3EBF5FAE1B23A97C1DA41D3D6F384ACAB3A6A924FC19BCFF9'
    if ($manifest.ModuleVersion -ne '3.0.0' -or (Get-FileHash -LiteralPath $license).Hash -ne $pinnedLicense) {
        throw 'Use the Gallery PowerShellNotebook 3.0.0 package with its original MIT license.'
    }
    $engines += 'PowerShellNotebook'
}

$ModuleManifest = (Resolve-Path -LiteralPath $ModuleManifest).ProviderPath
$sourceManifest = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../PSWriteOffice.psd1'))
$previousDevelopmentMode = $env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES
$previousBuildConfiguration = $env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION
$fixtureRoot = Join-Path ([IO.Path]::GetFullPath($OutputRoot)) ("input-" + [Guid]::NewGuid().ToString('N'))

try {
    if ($ModuleManifest -eq $sourceManifest) {
        $env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES = 'true'
        $env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION = 'Release'
    }
    if (-not $Plan -and $Mode -ne 'Cold') {
        Import-Module -Name $ModuleManifest
        if ($PowerShellNotebookManifest) {
            Import-Module -Name $PowerShellNotebookManifest
        }
    }

    Import-Module -Name $PowerForge -MinimumVersion 3.0.194
    $notebookBenchmarks = Join-Path $PSScriptRoot 'Notebook'
    . (Join-Path $notebookBenchmarks 'notebook-performance.fixtures.ps1')
    $filesPerCase = if ($Mode -eq 'Batch') { $BatchSize } else { 1 }
    if (-not $Plan) {
        [void][IO.Directory]::CreateDirectory($fixtureRoot)
        foreach ($size in $CellCount) {
            New-NotebookPerformanceFixture $fixtureRoot $size $filesPerCase -Batch:($Mode -eq 'Batch')
        }
    }
    $runMode = $Mode.ToLowerInvariant()
    if ($MemorySamplingIntervalMilliseconds -gt 0) {
        $runMode += "-memory-${MemorySamplingIntervalMilliseconds}ms"
    }
    $parameters = @{
        Path = Join-Path $notebookBenchmarks 'notebook-performance.benchmark.ps1'
        OutputRoot = $OutputRoot
        RunMode = $runMode
        Engine = $engines
        Operation = $Operation
        WarmupCount = $WarmupCount
        IterationCount = $IterationCount
        MemorySamplingIntervalMilliseconds = $MemorySamplingIntervalMilliseconds
        Plan = $Plan
        Variable = @{
            Mode = $Mode
            Cells = $CellCount -join ','
            FilesPerCase = $filesPerCase
            ThrottleLimit = $ThrottleLimit
            ModuleManifest = $ModuleManifest
            PowerShellNotebookManifest = $PowerShellNotebookManifest
            FixtureRoot = $fixtureRoot
            PowerShellPath = (Get-Process -Id $PID).Path
            ObserveMemory = $MemorySamplingIntervalMilliseconds -gt 0
        }
    }
    $result = Invoke-BenchmarkSuite @parameters
    $result
    if (-not $Plan) {
        $failed = @($result.Samples | Where-Object Status -ne 'Succeeded')
        if ($failed.Count -gt 0) {
            $first = $failed[0]
            throw "Notebook benchmark failed for $($first.Engine)/$($first.Operation)/$($first.Scenario). $($first.Reason)"
        }
        if (@($result.Samples).Count -eq 0) {
            throw 'The notebook benchmark completed without measurements.'
        }
    }
} finally {
    $env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES = $previousDevelopmentMode
    $env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION = $previousBuildConfiguration
    if (Test-Path -LiteralPath $fixtureRoot) {
        Remove-Item -LiteralPath $fixtureRoot -Recurse
    }
}
