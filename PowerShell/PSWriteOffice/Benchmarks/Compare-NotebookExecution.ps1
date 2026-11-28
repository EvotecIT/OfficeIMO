[CmdletBinding()]
param(
    [ValidateSet('Cold', 'Warm')]
    [string] $Mode = 'Cold',

    [ValidateSet('State12', 'State120', 'LargeText')]
    [string[]] $Scenario = @('State12', 'State120', 'LargeText'),

    [ValidateSet('OfficeIMO', 'OfficeIMOProcess', 'PowerShellNotebook', 'DotNetInteractive')]
    [string[]] $Engine = @('OfficeIMO', 'OfficeIMOProcess', 'PowerShellNotebook', 'DotNetInteractive'),

    [ValidateRange(1, 100)]
    [int] $IterationCount = 5,

    [string] $ModuleManifest = (Join-Path $PSScriptRoot '../PSWriteOffice.psd1'),
    [string] $PowerShellNotebookManifest,
    [string] $DotnetReplPath,
    [string] $InteractiveBenchmarkAssembly,
    [string] $PowerForge = 'PSPublishModule',
    [string] $OutputRoot = (Join-Path $PSScriptRoot '../Ignore/Benchmarks/NotebookExecution'),
    [switch] $Plan
)

$ErrorActionPreference = 'Stop'
$benchmarkPath = Join-Path $PSScriptRoot 'NotebookExecution'
. (Join-Path $benchmarkPath 'notebook-execution.fixtures.ps1')
Import-Module -Name $PowerForge -MinimumVersion 3.0.194

$previousDevelopment = $env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES
$previousConfiguration = $env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION
$previousTelemetry = $env:DOTNET_INTERACTIVE_CLI_TELEMETRY_OPTOUT
$fixtureRoot = Join-Path ([IO.Path]::GetFullPath($OutputRoot)) ('inputs-' + [Guid]::NewGuid().ToString('N'))

try {
    $toolVersion = ''
    $officeimoWarmModuleLoaded = $false
    if (-not $Plan) {
        if ($Mode -eq 'Warm' -and [Environment]::Version.Major -ne 10) {
            throw 'Warm notebook benchmarks require PowerShell running on .NET 10.'
        }
        if ('PowerShellNotebook' -in $Engine) {
            $PowerShellNotebookManifest = (Resolve-Path -LiteralPath $PowerShellNotebookManifest).ProviderPath
            $manifest = Import-PowerShellDataFile -LiteralPath $PowerShellNotebookManifest
            $peerDirectory = Split-Path $PowerShellNotebookManifest
            $peerFiles = @{
                'LICENSE' = '32504C072766A8A3EBF5FAE1B23A97C1DA41D3D6F384ACAB3A6A924FC19BCFF9'
                'PowerShellNotebook.psm1' = 'CDBF608FDC28C7F3512619B90ACB3AAD82EBD6A6E5E54DC90C71435D7A0225CC'
                'InvokePowerShellNotebook.ps1' = '4609BD3496384F4CB667385A41CDBA60EF476862E5517AE8B84B543277CF6408'
                'PowerShellNotebookDSL.ps1' = '23360CAE585F3DD0CB67A4A74DB5A5888DF9598FE96B3945A27A9CE341515C6E'
            }
            if ($manifest.ModuleVersion -ne '3.0.0') {
                throw 'Use the Gallery PowerShellNotebook 3.0.0 package.'
            }
            foreach ($entry in $peerFiles.GetEnumerator()) {
                if ((Get-FileHash -LiteralPath (Join-Path $peerDirectory $entry.Key)).Hash -ne $entry.Value) {
                    throw "Use the Gallery PowerShellNotebook 3.0.0 package. '$($entry.Key)' differs from the package."
                }
            }
            if ($Mode -eq 'Warm') {
                Import-Module -Name $PowerShellNotebookManifest -WarningAction SilentlyContinue
            }
        }
        if ('OfficeIMO' -in $Engine -or 'OfficeIMOProcess' -in $Engine) {
            $ModuleManifest = (Resolve-Path -LiteralPath $ModuleManifest).ProviderPath
            $sourceManifest = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../PSWriteOffice.psd1'))
            if ($ModuleManifest -eq $sourceManifest) {
                $env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES = 'true'
                $env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION = 'Release'
            }
            if ($Mode -eq 'Warm') {
                Import-Module -Name $ModuleManifest -WarningAction SilentlyContinue
                $officeimoWarmModuleLoaded = $true
            }
        }
        if ('DotNetInteractive' -in $Engine) {
            $env:DOTNET_INTERACTIVE_CLI_TELEMETRY_OPTOUT = '1'
            if ($Mode -eq 'Warm') {
                $InteractiveBenchmarkAssembly = (Resolve-Path -LiteralPath $InteractiveBenchmarkAssembly).ProviderPath
                Add-Type -LiteralPath $InteractiveBenchmarkAssembly
            } else {
                $DotnetReplPath = (Resolve-Path -LiteralPath $DotnetReplPath).ProviderPath
                $toolVersion = (& $DotnetReplPath --version) -join ' '
                if ($LASTEXITCODE -ne 0 -or $toolVersion -notmatch '^0\.3\.259(\+|$)') {
                    throw "Use dotnet-repl 0.3.259. Reported version '$toolVersion'."
                }
            }
        }
        [void][IO.Directory]::CreateDirectory($fixtureRoot)
        Get-NotebookExecutionCases $Scenario $fixtureRoot | ForEach-Object { Write-NotebookExecutionFixture $_ }
    }

    $parameters = @{
        Path = Join-Path $benchmarkPath 'notebook-execution.benchmark.ps1'
        OutputRoot = $OutputRoot
        RunMode = $Mode.ToLowerInvariant()
        Engine = $Engine
        Operation = if ($Mode -eq 'Cold') { 'RunAndSave' } else { 'ExecuteCells' }
        IterationCount = $IterationCount
        WarmupCount = 0
        Plan = $Plan
        Variable = @{
            Mode = $Mode
            Scenarios = $Scenario -join ','
            FixtureRoot = $fixtureRoot
            ModuleManifest = $ModuleManifest
            PowerShellNotebookManifest = $PowerShellNotebookManifest
            DotnetReplPath = $DotnetReplPath
            DotnetReplVersion = $toolVersion
            PowerShellPath = (Get-Process -Id $PID).Path
            OfficeimoWarmModuleLoaded = $officeimoWarmModuleLoaded
            InteractiveBenchmarkAssembly = $InteractiveBenchmarkAssembly
        }
    }
    $result = Invoke-BenchmarkSuite @parameters
    $result
    if (-not $Plan) {
        $failed = @($result.Samples | Where-Object Status -ne 'Succeeded')
        if ($failed.Count -gt 0) {
            $first = $failed[0]
            throw "Notebook execution benchmark failed for $($first.Engine)/$($first.Scenario). $($first.Reason)"
        }
        $expected = $Scenario.Count * $Engine.Count * $IterationCount
        if (@($result.Samples).Count -ne $expected) {
            throw "Expected $expected notebook execution samples."
        }
    }
} finally {
    $env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES = $previousDevelopment
    $env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION = $previousConfiguration
    $env:DOTNET_INTERACTIVE_CLI_TELEMETRY_OPTOUT = $previousTelemetry
    if (Test-Path -LiteralPath $fixtureRoot) {
        Remove-Item -LiteralPath $fixtureRoot -Recurse
    }
}
