param([Parameter(Mandatory)][string] $CompareDirectory, [long] $Mask = 65535, [int] $Iterations = 24, [int] $Warmup = 32, [int] $Operations = 4,
    [Parameter(Mandatory)][string] $ModulePath, [string] $Cases = 'cases.json', [string] $Label = 'three-way-diagnostic')
$ErrorActionPreference = 'Stop'
$controlModule = Import-Module -Name $ModulePath -PassThru
$controlModuleRoot = Split-Path $ModulePath -Parent
$controlOwnerSha = (Get-FileHash -LiteralPath (Join-Path $controlModuleRoot 'Lib/Core/PowerForge.dll')).Hash
$controlCmdletSha = (Get-FileHash -LiteralPath (Join-Path $controlModuleRoot 'Lib/Core/PSPublishModule.dll')).Hash
$env:OFFICEIMO_PERFORMANCE_SNAPSHOTS = Join-Path $PSScriptRoot 'net10.0'
$env:OFFICEIMO_COMPARISON_CASES = Join-Path $PSScriptRoot $Cases
$env:OFFICEIMO_PERFORMANCE_AA = '0'
$env:OFFICEIMO_BENCHMARK_OUTPUT = Join-Path $PSScriptRoot 'files'
$env:OFFICEIMO_BENCHMARK_DATA = Join-Path $PSScriptRoot 'files'
if ($IsWindows -or $IsLinux) { [Diagnostics.Process]::GetCurrentProcess().ProcessorAffinity = [IntPtr]::new($Mask) }
[Diagnostics.Process]::GetCurrentProcess().PriorityClass = 'Normal'
[Reflection.Assembly]::LoadFrom((Join-Path $CompareDirectory 'bin/Release/net10.0/Compare.dll')) | Out-Null
$manifest = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'manifest-net10.0.json') -Raw | ConvertFrom-Json
function Assert-FrozenDimensionlessSnapshot {
    foreach ($file in $manifest.Manifest) {
        foreach ($side in 'Before','After') {
            $directory = if ($side -eq 'Before') { 'baseline-excel' } else { 'candidate-excel' }
            if ((Get-FileHash -LiteralPath (Join-Path $PSScriptRoot "net10.0/$directory/$($file.Name)")).Hash -ne $file.$side) {
                throw 'Dimensionless control snapshot changed'
            }
        }
    }
}
Assert-FrozenDimensionlessSnapshot
$global:BufferProbes = @{}
$global:BufferExpected = @{}
$global:BufferOperations = $Operations
$fixtures = [ordered]@{}
function Get-WorksheetFixtureHashes($probe) {
    $flags = [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic
    $instance = [SnapshotProbe].GetField('_instance', $flags).GetValue($probe)
    $type = $instance.GetType()
    $pathField = $type.GetField('_path', $flags)
    $stream = if ($null -ne $pathField) {
        [IO.File]::OpenRead([string]$pathField.GetValue($instance))
    } else {
        [IO.MemoryStream]::new([byte[]]$type.GetField('_workbookBytes', $flags).GetValue($instance), $false)
    }
    $archive = [IO.Compression.ZipArchive]::new($stream, [IO.Compression.ZipArchiveMode]::Read, $false)
    $hashes = [ordered]@{}
    try {
        foreach ($entry in $archive.Entries | Where-Object {$_.FullName -match '^xl/(worksheets/|styles.xml$|sharedStrings.xml$)'} | Sort-Object FullName) {
            $part = $entry.Open()
            $hash = [Security.Cryptography.SHA256]::Create()
            try { $hashes[$entry.FullName] = [Convert]::ToHexString($hash.ComputeHash($part)) }
            finally { $hash.Dispose(); $part.Dispose() }
        }
    } finally { $archive.Dispose() }
    return $hashes
}
foreach ($key in [SnapshotProbe]::Keys()) {
    $before = [SnapshotProbe]::Load($key, 'baseline')
    $after = [SnapshotProbe]::Load($key, 'candidate')
    $control = [SnapshotProbe]::Load($key, 'baseline')
    [SnapshotProbe]::ValidateEquivalent($before, $after)
    [SnapshotProbe]::ValidateEquivalent($before, $control)
    $global:BufferProbes[$key] = @{ Before = $before; Control = $control; After = $after }
    $global:BufferExpected[$key] = @{ Before=$before.RunSynchronously().ToString(); Control=$control.RunSynchronously().ToString(); After=$after.RunSynchronously().ToString() }
    $beforeHashes = Get-WorksheetFixtureHashes $before
    $afterHashes = Get-WorksheetFixtureHashes $after
    foreach ($part in $beforeHashes.Keys) {
        if ($beforeHashes[$part] -ne $afterHashes[$part]) { throw "Decoded fixture part differs: $key / $part" }
    }
    if ($beforeHashes.Count -ne $afterHashes.Count) { throw "Fixture part inventory differs: $key" }
    $fixtures[$key] = @{Before=$beforeHashes;After=$afterHashes;Observation=$global:BufferExpected[$key]}
}
$fixtures | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $PSScriptRoot "$Label-$Mask-fixtures.json") -Encoding utf8
try {
    Invoke-BenchmarkSuite -OutputRoot (Join-Path $PSScriptRoot "$Label-$Mask") -WarmupCount $Warmup -IterationCount $Iterations -RunOrder Rotated -OutlierMode None -Settings {
        New-BenchmarkSuite 'Large worksheet budget comparison' {
            Add-BenchmarkMetadata AffinityMask $(if ($IsWindows -or $IsLinux) { $Mask } else { 'OS scheduled' })
            Add-BenchmarkMetadata SourceContract 'Qualified undimensioned index baseline; two private used-range budget paths differ; bounded fragments can decline optional indexing and existing streaming validation remains authoritative'
            Add-BenchmarkMetadata Normalization 'Complete workbook reads per sample; OperationsPerSample records batch size' 
            Add-BenchmarkMetadata OperationsPerSample $Operations
            Add-BenchmarkMetadata Diagnostic 'Before and Control use identical baseline source and bytes; After uses candidate; all three engines rotate within every retained iteration'
            Add-BenchmarkMetadata ControlModuleVersion ($controlModule.Version.ToString())
            Add-BenchmarkMetadata ControlOwnerSha256 $controlOwnerSha
            Add-BenchmarkMetadata ControlCmdletsSha256 $controlCmdletSha
            Add-BenchmarkMetadata Runtime ([System.Runtime.InteropServices.RuntimeInformation]::FrameworkDescription)
            Add-BenchmarkMetadata Priority ([Diagnostics.Process]::GetCurrentProcess().PriorityClass.ToString())
            Add-BenchmarkCaseSource @($global:BufferProbes.Keys | Sort-Object | ForEach-Object { [pscustomobject]@{Name=$_} })
            Set-BenchmarkSetup {
                param($case, $run)
                $run.Probe = $global:BufferProbes[$case.Scenario][$case.Engine]
                $run.Expected = $global:BufferExpected[$case.Scenario][$case.Engine]
            }
            foreach ($engine in 'Before','Control','After') {
                Add-BenchmarkEngine $engine {
                    Add-BenchmarkOperation ReadAllFields {
                        param($case, $run)
                        for ($i=0; $i -lt $global:BufferOperations; $i++) { $run.Result = $run.Probe.RunSynchronously() }
                    }
                }
            }
            Add-BenchmarkValidation {
                param($case, $run)
                if ($run.Result.ToString() -ne $run.Expected) { throw 'Returned count differs from this build''s independently validated worksheet.' }
            }
            Add-BenchmarkComparison -Baseline Before -Metric MeanMs,MedianMs -TieTolerance 0.05
        }
    }
} finally {
    foreach ($pair in $global:BufferProbes.Values) { $pair.Before.Dispose(); $pair.Control.Dispose(); $pair.After.Dispose() }
}







Assert-FrozenDimensionlessSnapshot
$paths = @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot "$Label-$Mask") -Filter run-report.json -Recurse)
if ($paths.Count -ne 1) { throw 'Wrong dimensionless control report inventory' }
$report = Get-Content -LiteralPath $paths[0].FullName -Raw | ConvertFrom-Json
$expectedCount = $global:BufferProbes.Count * 3
if ($report.summary.Count -ne $expectedCount -or @($report.summary | Where-Object { $_.sampleCount -ne $Iterations -or $_.failureCount -ne 0 }).Count) { throw 'Incomplete dimensionless rotated controls' }
[ordered]@{ Manifest = $manifest; Report = $report; Sha256 = (Get-FileHash -LiteralPath $paths[0].FullName).Hash; RunnerSha256 = (Get-FileHash -LiteralPath $PSCommandPath).Hash; Policy = "Three rotated engines; Before and Control use identical bytes; warmups$Warmup;retained$Iterations;operations$Operations;no outlier removal;each decoded fixture and complete typed result qualified" } | ConvertTo-Json -Depth 26 | Set-Content -LiteralPath (Join-Path $PSScriptRoot "retained-$Label-$Mask.json") -Encoding utf8
Write-Output "Dimensionless rotated controls qualified: $expectedCount observations"
