# Each project keeps its own restore, test process and TRX directory. A failed
# contract must not prevent the remaining contracts in the shard from running.
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $ProjectsJson,
    [Parameter(Mandatory)][string] $EvidenceDirectory,
    [Parameter(Mandatory)][string] $DiagnosticsDirectory
)

$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $false
$projects = @($ProjectsJson | ConvertFrom-Json)
if ($projects.Count -eq 0) { throw 'A Windows shard must contain at least one contract.' }
$configuration = if ($env:BUILD_CONFIGURATION) { $env:BUILD_CONFIGURATION } else { 'Release' }
$verbosity = if ($env:TEST_VERBOSITY) { $env:TEST_VERBOSITY } else { 'minimal' }
$filter = 'Category!=ExcelImageVisualGate&Category!=Performance'
New-Item -ItemType Directory -Force -Path $EvidenceDirectory, $DiagnosticsDirectory | Out-Null
$records = [System.Collections.Generic.List[object]]::new()
$failed = $false

function Invoke-DotnetStep([string] $Name, [string[]] $Arguments) {
    Write-Host "::group::$Name"
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $code = 1
    try {
        & dotnet @Arguments | Out-Host
        $code = $LASTEXITCODE
    } catch {
        Write-Warning "${Name}: $_"
    } finally {
        $watch.Stop()
        Write-Host '::endgroup::'
    }
    return [ordered]@{ name = $Name; seconds = $watch.Elapsed.TotalSeconds; exit_code = $code }
}

foreach ($project in $projects) {
    $name = [string]$project.name
    $path = [string]$project.project
    if ($name -notmatch '^[A-Za-z0-9-]+$' -or $path -notmatch '^OfficeIMO\.[A-Za-z0-9./-]+\.csproj$') {
        throw "Invalid Windows contract: $name / $path"
    }
    $steps = [System.Collections.Generic.List[object]]::new()
    $restore = Invoke-DotnetStep "$name / restore" @('restore', $path, "-bl:$DiagnosticsDirectory/$name-restore-{}.binlog")
    $steps.Add($restore)
    if ($restore.exit_code -eq 0) {
        # Exact image pixels belong to the dedicated modern-runtime visual gate.
        $test = Invoke-DotnetStep "$name / net472" @(
            'test', $path, '--configuration', $configuration, '--framework', 'net472',
            '--no-restore', '--verbosity', $verbosity, '--filter', $filter,
            '--logger', "console;verbosity=$verbosity", '--logger', 'trx',
            "-bl:$DiagnosticsDirectory/$name-test-{}.binlog"
        )
        $steps.Add($test)
        if ($test.exit_code -eq 0 -and $name -eq 'Drawing-CodeGlyphX') {
            foreach ($adapter in @('Drawing', 'Word', 'Excel', 'PowerPoint', 'Pdf')) {
                $steps.Add((Invoke-DotnetStep "$name / pack $adapter" @(
                    'pack', "OfficeIMO.$adapter.CodeGlyphX/OfficeIMO.$adapter.CodeGlyphX.csproj",
                    '--configuration', $configuration, '--no-restore', '--output', 'artifacts/codeglyphx-packages',
                    "-bl:$DiagnosticsDirectory/$name-pack-$adapter-{}.binlog"
                )))
            }
        }
    }
    if (@($steps | Where-Object exit_code -NE 0).Count) { $failed = $true }
    $records.Add([ordered]@{ name = $name; project = $path; steps = @($steps.ToArray()) })
    # Preserve completed results even if a subsequent project or job fails.
    ConvertTo-Json -InputObject @($records.ToArray()) -Depth 8 |
        Set-Content -LiteralPath (Join-Path $EvidenceDirectory 'execution.json') -Encoding utf8
}
if ($failed) { exit 1 }
