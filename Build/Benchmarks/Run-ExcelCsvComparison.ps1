param(
    # Lane names; a comma-separated list also works from pwsh -File.
    [string[]] $Lane = @('Xlsx65K', 'Xlsb65K', 'ExcelGenerated', 'ExcelWriters', 'XLBenchReadAll', 'XLBenchSheets',
        'CsvPackageAssets', 'CsvSpacesQuotes', 'Csv65K', 'CsvGenerated', 'CsvWrite', 'CsvWriteModern', 'CsvPackageAssetsModern', 'Csv65KModern', 'Csv65KInterned'),
    [string] $OutputRoot,
    [string] $FixtureRoot,
    [string] $XLBenchRoot = $env:OFFICEIMO_XLBENCH_ROOT,
    [string] $Framework = 'net10.0',
    [ValidateRange(1, 100)] [int] $WarmupCount = 3,
    [ValidateRange(3, 200)] [int] $IterationCount = 12,
    [string] $AffinityMask,
    [switch] $LatestCompetitors,
    [switch] $SkipBuild,
    [switch] $Plan
)
# Opt-in comparison of the OfficeIMO Excel and CSV readers and writers with the libraries their
# benchmark projects already compare against. Every lane validates its output inside the benchmark
# classes; this script only selects lanes, runs them in-process and merges the results.
$ErrorActionPreference = 'Stop'
$repositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path

# Scratch and fixtures follow EVOTEC_SCRATCH_ROOT on every system; otherwise they stay in the ignored Ignore/ folder.
$scratchRoot = if ($env:EVOTEC_SCRATCH_ROOT) { Join-Path $env:EVOTEC_SCRATCH_ROOT 'officeimo-comparison' }
    else { Join-Path $repositoryRoot 'Ignore/Benchmarks/ExcelCsvComparison' }
if (-not $OutputRoot) { $OutputRoot = Join-Path $scratchRoot ('run-' + (Get-Date -Format 'yyyyMMdd-HHmmss')) }
if (-not $FixtureRoot) { $FixtureRoot = Join-Path $scratchRoot 'fixtures' }
$OutputRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputRoot)
$FixtureRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($FixtureRoot)

$excelProject = Join-Path $repositoryRoot 'OfficeIMO.Excel.Benchmarks/OfficeIMO.Excel.Benchmarks.csproj'
$csvProject = Join-Path $repositoryRoot 'OfficeIMO.CSV.Benchmarks/OfficeIMO.CSV.Benchmarks.csproj'
# CsvHelper 34 replaced the APIs of the main CSV project, so its lanes run in the Modern project.
$csvModernProject = Join-Path $repositoryRoot 'OfficeIMO.CSV.Benchmarks.Modern/OfficeIMO.CSV.Benchmarks.Modern.csproj'
$lanes = [ordered]@{
    # 65K sales workbook: OfficeIMO default and single-pass reading against each library's primary reader.
    Xlsx65K = @{ Project = 'Excel'; Filters = @('OfficeIMO', 'OfficeIMO_SinglePass', 'ExcelReaderNet', 'Sylvan', 'ExcelDataReader',
            'ClosedXML', 'EPPlus', 'MiniExcel', 'HypeLab', 'ExcelPrime', 'XLibur') | ForEach-Object { "*MarkPflug65KXlsxBenchmarks.$_" } }
    Xlsb65K = @{ Project = 'Excel'; Filters = @('OfficeIMO', 'OfficeIMO_SinglePass', 'ExcelReaderNet', 'Sylvan', 'ExcelDataReader',
            'ExcelPrime') | ForEach-Object { "*MarkPflug65KXlsbBenchmarks.$_" } }
    ExcelGenerated = @{ Project = 'Excel'; Filters = @('*ExcelReaderGeneratedBenchmarks.*') }
    ExcelWriters = @{ Project = 'Excel'; Filters = @('*ExcelReaderGeneratedWriteBenchmarks.OfficeIMO', '*ExcelReaderGeneratedWriteBenchmarks.ExcelReaderNet',
            '*ExcelReaderGeneratedXlsxWriteBenchmarks.*') }
    XLBenchReadAll = @{ Project = 'XLBench'; Filters = @('XLBench.Benchmarks.Read.*.OpenAndReadAll') }
    XLBenchSheets = @{ Project = 'XLBench'; Filters = @('XLBench.Benchmarks.MultiSheet.*.OpenAndReadAllSheets') }
    CsvPackageAssets = @{ Project = 'CSV'; Filters = @('*.PackageAssetsCsvBenchmarks.*') }
    CsvSpacesQuotes = @{ Project = 'CSV'; Filters = @('*.PackageAssetsSpacesQuotesCsvBenchmarks.*') }
    Csv65K = @{ Project = 'CSV'; Filters = @('*.MarkPflug65KCsvBenchmarks.*') }
    CsvGenerated = @{ Project = 'CSV'; Filters = @('*.CsvExcelReaderGeneratedBenchmarks.*') }
    # Typed objects, DataTable and IDataReader to UTF-8 streams and TextWriters, byte-identical across libraries.
    CsvWrite = @{ Project = 'CSV'; Filters = @('*.CsvWriteComparisonBenchmarks.*') }
    CsvWriteModern = @{ Project = 'CSV.Modern'; Filters = @('*.CsvWriteModernBenchmarks.*') }
    CsvPackageAssetsModern = @{ Project = 'CSV.Modern'; Filters = @('*.PackageAssetsModernCsvBenchmarks.*') }
    Csv65KModern = @{ Project = 'CSV.Modern'; Filters = @('*.MarkPflug65KModernCsvBenchmarks.*') }
    # OfficeIMO's string reuse option runs in its own process, apart from the library comparison.
    Csv65KInterned = @{ Project = 'CSV.Modern'; Filters = @('*.MarkPflug65KInternedCsvBenchmarks.*') }
}

# EPPlus 8 refuses to run without a license declaration; the run includes it only when the caller made one.
$laneNotes = @{}
if (-not $env:EPPlusLicense) {
    $lanes.Xlsx65K.Filters = @($lanes.Xlsx65K.Filters | Where-Object { $_ -notlike '*.EPPlus' })
    $laneNotes.Xlsx65K = 'EPPlus excluded: set EPPlusLicense to its license declaration to include it.'
}

function Get-HostDescription {
    $cpu = $null
    $memory = $null
    if ($IsWindows) {
        $cpu = (Get-CimInstance Win32_Processor | Select-Object -First 1).Name
        $memory = (Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory
    } elseif ($IsMacOS) {
        $cpu = (& sysctl -n machdep.cpu.brand_string 2>$null)
        $memory = [long](& sysctl -n hw.memsize 2>$null)
    } elseif (Test-Path /proc/cpuinfo) {
        $cpu = ((Get-Content /proc/cpuinfo | Where-Object { $_ -like 'model name*' } | Select-Object -First 1) -split ':', 2)[1]
        $memory = 1KB * [long](((Get-Content /proc/meminfo | Select-Object -First 1) -split '\s+')[1])
    }
    [ordered]@{
        OperatingSystem = [Runtime.InteropServices.RuntimeInformation]::OSDescription
        Architecture = [Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString()
        Processor = "$cpu".Trim()
        LogicalProcessors = [Environment]::ProcessorCount
        MemoryGB = if ($memory) { [Math]::Round($memory / 1GB, 1) } else { $null }
        DotNetSdk = (& dotnet --version)
        Framework = $Framework
        HardwareIntrinsicsDisabled = $env:DOTNET_EnableHWIntrinsic -eq '0'
        # A child process sees the same DOTNET_ settings as the benchmark processes.
        VectorSupport = (& ([Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) -NoProfile -Command {
                @(
                    if ([Numerics.Vector]::IsHardwareAccelerated) { "Vector<T> $([Numerics.Vector[byte]]::Count * 8)-bit" }
                    if ([Runtime.Intrinsics.Vector128]::IsHardwareAccelerated) { 'Vector128' }
                    if ([Runtime.Intrinsics.Vector256]::IsHardwareAccelerated) { 'Vector256' }
                    if ([Runtime.Intrinsics.Vector512]::IsHardwareAccelerated) { 'Vector512' }
                    if ([Runtime.Intrinsics.X86.Avx2]::IsSupported) { 'AVX2' }
                    if ([Runtime.Intrinsics.X86.Avx512BW]::IsSupported) { 'AVX-512BW' }
                    if ([Runtime.Intrinsics.Arm.AdvSimd]::IsSupported) { 'AdvSIMD' }
                ) -join ', '
            }) | ForEach-Object { if ($_) { $_ } else { 'none' } }
        AffinityMask = if ($IsMacOS) { 'unpinned (macOS)' }
            else { '0x{0:X}' -f [Diagnostics.Process]::GetCurrentProcess().ProcessorAffinity.ToInt64() }
        Commit = (& git -C $repositoryRoot rev-parse HEAD)
        WorkingTreeClean = -not (& git -C $repositoryRoot status --porcelain --untracked-files=no)
        WarmupCount = $WarmupCount
        IterationCount = $IterationCount
        CompetitorMode = if ($LatestCompetitors) { 'latest' } else { 'pinned' }
        CompetitorVersions = $competitorVersions
    }
}

$Lane = @($Lane | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
$unknown = @($Lane | Where-Object { -not $lanes.Contains($_) })
if ($unknown.Count -gt 0) { throw "Unknown lane(s): $($unknown -join ', '). Lanes: $($lanes.Keys -join ', ')." }
$selected = @($Lane)
$xlbenchProject = if ($XLBenchRoot) { Join-Path $XLBenchRoot 'src/XLBench/XLBench.csproj' }
$xlbenchAvailable = $xlbenchProject -and (Test-Path $xlbenchProject)
if ($Plan) {
    foreach ($name in $selected) {
        $entry = $lanes[$name]
        $note = if ($entry.Project -eq 'XLBench' -and -not $xlbenchAvailable) { ' (skipped: no XLBench checkout)' } else { '' }
        "{0,-17} {1,-8} {2}{3}" -f $name, $entry.Project, ($entry.Filters -join ' '), $note
    }
    "Output:   $OutputRoot"
    "Fixtures: $FixtureRoot"
    return
}

if ($AffinityMask) {
    # Children inherit the affinity. macOS has no processor-affinity API, so runs there are unpinned.
    if ($IsMacOS) { throw 'Processor affinity is not available on macOS; run without -AffinityMask.' }
    $mask = if ($AffinityMask.StartsWith('0x')) { [Convert]::ToInt64($AffinityMask.Substring(2), 16) } else { [long] $AffinityMask }
    [Diagnostics.Process]::GetCurrentProcess().ProcessorAffinity = [IntPtr] $mask
}

[void] (New-Item -ItemType Directory -Force -Path $OutputRoot, $FixtureRoot)
$temporaryRoot = Join-Path $OutputRoot 'tmp'
[void] (New-Item -ItemType Directory -Force -Path $temporaryRoot)
$env:OFFICEIMO_BENCHMARK_DATA = $FixtureRoot
$env:TEMP = $temporaryRoot; $env:TMP = $temporaryRoot; $env:TMPDIR = $temporaryRoot

$binaries = @{}
$assetFiles = @{}
$projects = @{ Excel = $excelProject; CSV = $csvProject; 'CSV.Modern' = $csvModernProject }
# The competitor watch restores the newest stable competitor packages into its own package and build folders,
# so the shared NuGet cache and the ordinary pinned build stay as they are.
$latestRoot = Join-Path $scratchRoot 'latest-competitors'
$latestArguments = @('-p:OfficeImoLatestCompetitors=true', "-p:RestorePackagesPath=$(Join-Path $latestRoot 'packages')",
    '--artifacts-path', (Join-Path $latestRoot 'artifacts'))
foreach ($kind in @($selected | ForEach-Object { $lanes[$_].Project } | Sort-Object -Unique)) {
    if ($kind -eq 'XLBench') {
        if (-not $xlbenchAvailable) { continue }
        # The checkout must build against this repository's OfficeIMO.Excel through its OfficeImoSource property.
        $buildArguments = @('build', $xlbenchProject, '-c', 'Release', "-p:OfficeImoSource=$(Join-Path $repositoryRoot 'OfficeIMO.Excel/OfficeIMO.Excel.csproj')")
        $binaryDirectory = Join-Path $XLBenchRoot "src/XLBench/bin/Release/$Framework"
        $binaries[$kind] = Join-Path $binaryDirectory 'XLBench.dll'
    } elseif ($LatestCompetitors) {
        $projectName = [IO.Path]::GetFileNameWithoutExtension($projects[$kind])
        if (-not $SkipBuild) {
            # Floating versions resolve against the feed, not the HTTP cache, on every watch run.
            & dotnet restore $projects[$kind] -p:TargetFrameworks=$Framework --force --no-http-cache @latestArguments -nologo -v q | Out-Host
            if ($LASTEXITCODE -ne 0) { throw "Restoring the latest competitors for the $kind benchmarks failed." }
        }
        $buildArguments = @('build', $projects[$kind], '-c', 'Release', '-f', $Framework, '--no-restore', "-p:TargetFrameworks=$Framework") + $latestArguments
        $binaries[$kind] = Join-Path $latestRoot "artifacts/bin/$projectName/release_$Framework/$projectName.dll"
        $assetFiles[$kind] = Join-Path $latestRoot "artifacts/obj/$projectName/project.assets.json"
    } else {
        $buildArguments = @('build', $projects[$kind], '-c', 'Release', '-f', $Framework)
        $binaries[$kind] = Join-Path (Split-Path $projects[$kind]) "bin/Release/$Framework/$([IO.Path]::GetFileNameWithoutExtension($projects[$kind])).dll"
        $assetFiles[$kind] = Join-Path (Split-Path $projects[$kind]) 'obj/project.assets.json'
    }
    if (-not $SkipBuild) {
        & dotnet @buildArguments -nologo -v q | Out-Host
        if ($LASTEXITCODE -ne 0) { throw "Building the $kind benchmarks failed." }
    }
    if (-not (Test-Path $binaries[$kind])) { throw "Missing benchmark binary: $($binaries[$kind])" }
}

# Resolved versions of the watched competitors, read from the restore that built each project.
$watched = @('ExcelReader.NET', 'Sep', 'Sylvan.Data.Csv', 'Sylvan.Data.Excel', 'CsvHelper')
$competitorVersions = [ordered]@{}
foreach ($kind in $assetFiles.Keys) {
    if (-not (Test-Path $assetFiles[$kind])) { continue }
    $libraries = (Get-Content $assetFiles[$kind] -Raw | ConvertFrom-Json).libraries.PSObject.Properties.Name
    foreach ($id in $watched) {
        $match = $libraries | Where-Object { $_ -like "$id/*" } | Select-Object -First 1
        if ($match) {
            if (-not $competitorVersions.Contains($kind)) { $competitorVersions[$kind] = [ordered]@{} }
            $competitorVersions[$kind][$id] = $match.Split('/')[1]
        }
    }
}

$environment = Get-HostDescription
$rows = [Collections.Generic.List[object]]::new()
$laneResults = [Collections.Generic.List[object]]::new()
foreach ($name in $selected) {
    $entry = $lanes[$name]
    if ($entry.Project -eq 'XLBench' -and -not $xlbenchAvailable) {
        $laneResults.Add([ordered]@{ Lane = $name; Status = 'Skipped'; Note = 'No XLBench checkout: set -XLBenchRoot or OFFICEIMO_XLBENCH_ROOT.' })
        Write-Host "[$name] skipped: no XLBench checkout"
        continue
    }

    $laneRoot = Join-Path $OutputRoot $name
    $binary = $binaries[$entry.Project]
    $arguments = @($binary, '--filter') + $entry.Filters + @('--inProcess', '--join', '--exporters', 'json',
        '--warmupCount', "$WarmupCount", '--iterationCount', "$IterationCount", '--artifacts', $laneRoot)
    Write-Host "[$name] $(Get-Date -Format HH:mm:ss) running $($entry.Filters.Count) filter(s)"
    $log = Join-Path $OutputRoot "$name.log"
    Push-Location (Split-Path $binary)
    try { & dotnet @arguments *> $log } finally { Pop-Location }
    $exitCode = $LASTEXITCODE
    $report = Get-ChildItem $laneRoot -Recurse -Filter '*-report*.json' -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $report) {
        $laneResults.Add([ordered]@{ Lane = $name; Status = 'Failed'; Note = "No results (exit $exitCode); see $log" })
        Write-Host "[$name] failed: no results (exit $exitCode)"
        continue
    }

    $json = Get-Content $report.FullName -Raw | ConvertFrom-Json
    if (-not $environment.Contains('BenchmarkHost')) {
        $hostInfo = $json.HostEnvironmentInfo
        $environment.BenchmarkHost = [ordered]@{
            Processor = $hostInfo.ProcessorName; Runtime = $hostInfo.RuntimeVersion; Architecture = $hostInfo.Architecture
            BenchmarkDotNet = $hostInfo.BenchmarkDotNetVersion
        }
    }
    $missing = 0
    foreach ($benchmark in $json.Benchmarks) {
        $statistics = $benchmark.Statistics
        if (-not $statistics) { $missing++ }
        $rows.Add([pscustomobject]@{
            Lane = $name
            Type = $benchmark.Type
            Namespace = $benchmark.Namespace
            Method = $benchmark.Method
            Parameters = "$($benchmark.Parameters)"
            MeanMs = if ($statistics) { $statistics.Mean / 1e6 } else { $null }
            MedianMs = if ($statistics) { $statistics.Median / 1e6 } else { $null }
            StdErrMs = if ($statistics) { $statistics.StandardError / 1e6 } else { $null }
            AllocatedMB = if ($benchmark.Memory) { $benchmark.Memory.BytesAllocatedPerOperation / 1MB } else { $null }
        })
    }
    # A lane whose validation throws leaves its benchmark without statistics; keep that visible.
    $status = if ($missing -gt 0 -or $exitCode -ne 0) { 'Partial' } else { 'Succeeded' }
    $notes = @($laneNotes[$name]; if ($status -ne 'Succeeded') { "$missing case(s) without results, exit $exitCode; see $log" }) | Where-Object { $_ }
    $laneResults.Add([ordered]@{ Lane = $name; Status = $status; Note = $notes -join ' ' })
    Write-Host "[$name] $status ($($json.Benchmarks.Count) cases)"
}

# Ratios compare each library with OfficeIMO's matching lane: OfficeIMO_<Operation> when the method names an
# operation, otherwise OfficeIMO; XLBench classes are one per library, so they compare within a scenario.
foreach ($row in $rows) {
    $group = $rows | Where-Object { $_.Lane -eq $row.Lane -and $_.Parameters -eq $row.Parameters }
    $reference = if ($row.Lane -like 'XLBench*') {
        $group | Where-Object { $_.Type -like 'OfficeImo*' -and $_.Method -eq $row.Method } | Select-Object -First 1
    } else {
        $operation = if ($row.Method -match '^[^_]+_(.+)$') { $Matches[1] }
        ($group | Where-Object { $operation -and $_.Method -eq "OfficeIMO_$operation" } | Select-Object -First 1) ??
            ($group | Where-Object { $_.Method -eq 'OfficeIMO' } | Select-Object -First 1)
    }
    $ratio = if ($reference -and $reference.MeanMs -and $row.MeanMs) { $row.MeanMs / $reference.MeanMs } else { $null }
    $row | Add-Member -NotePropertyName VsOfficeIMO -NotePropertyValue $ratio
}

$summary = [ordered]@{ Environment = $environment; Lanes = $laneResults; Results = $rows }
$summary | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $OutputRoot 'results.json') -Encoding utf8

$markdown = [Text.StringBuilder]::new()
[void] $markdown.AppendLine('# OfficeIMO Excel and CSV comparison').AppendLine()
foreach ($key in $environment.Keys) {
    $value = $environment[$key]
    if ($value -is [Collections.IDictionary]) {
        $value = ($value.GetEnumerator() | ForEach-Object {
                $item = $_.Value
                if ($item -is [Collections.IDictionary]) { $item = ($item.GetEnumerator() | ForEach-Object { "$($_.Key) $($_.Value)" }) -join ', ' }
                "$($_.Key): $item"
            }) -join '; '
    }
    [void] $markdown.AppendLine("- ${key}: $value")
}
[void] $markdown.AppendLine().AppendLine('| Lane | Status | Note |').AppendLine('|---|---|---|')
foreach ($result in $laneResults) { [void] $markdown.AppendLine("| $($result.Lane) | $($result.Status) | $($result.Note) |") }
[void] $markdown.AppendLine().AppendLine('Mean and median in ms; x OfficeIMO is mean divided by the matching OfficeIMO lane (above 1 is slower).').AppendLine()
[void] $markdown.AppendLine('| Lane | Benchmark | Parameters | Mean | Median | StdErr | Allocated MB | x OfficeIMO |')
[void] $markdown.AppendLine('|---|---|---|---:|---:|---:|---:|---:|')
foreach ($row in $rows) {
    $label = if ($row.Lane -like 'XLBench*') { $row.Type -replace 'Benchmarks$', '' } else { "$($row.Type).$($row.Method)" }
    $format = { param($value, $digits) if ($null -eq $value) { 'n/a' } else { $value.ToString("F$digits", [Globalization.CultureInfo]::InvariantCulture) } }
    [void] $markdown.AppendLine("| $($row.Lane) | $label | $($row.Parameters) | $(& $format $row.MeanMs 2) | $(& $format $row.MedianMs 2) | $(& $format $row.StdErrMs 2) | $(& $format $row.AllocatedMB 1) | $(& $format $row.VsOfficeIMO 2) |")
}
$markdown.ToString() | Set-Content (Join-Path $OutputRoot 'results.md') -Encoding utf8

# Competitor ledger: one entry per watched library case and run, kept under the scratch root across runs.
$packageOf = @{ ExcelReaderNet = 'ExcelReader.NET'; Sep = 'Sep'; CsvHelper = 'CsvHelper' }
$ledgerPath = Join-Path $scratchRoot 'competitor-ledger.json'
$ledger = [Collections.Generic.List[object]]::new()
if (Test-Path $ledgerPath) { foreach ($entry in @(Get-Content $ledgerPath -Raw | ConvertFrom-Json)) { $ledger.Add($entry) } }
$date = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture)
foreach ($row in $rows | Where-Object { $_.Lane -notlike 'XLBench*' -and $null -ne $_.VsOfficeIMO }) {
    $library = ($row.Method -split '_', 2)[0]
    $project = $lanes[$row.Lane].Project
    $package = if ($library -eq 'Sylvan') { if ($project -like 'CSV*') { 'Sylvan.Data.Csv' } else { 'Sylvan.Data.Excel' } } else { $packageOf[$library] }
    if (-not $package) { continue }
    $ledger.Add([pscustomobject]@{
        Date = $date; Commit = $environment.Commit.Substring(0, 10); Mode = $environment.CompetitorMode; Host = $environment.Processor
        Lane = $row.Lane; Benchmark = "$($row.Type).$($row.Method)"; Parameters = $row.Parameters
        Package = $package; Version = $(if ($competitorVersions.Contains($project)) { $competitorVersions[$project][$package] }); MeanMs = [Math]::Round($row.MeanMs, 3); Ratio = [Math]::Round($row.VsOfficeIMO, 3)
    })
}
ConvertTo-Json @($ledger) -Depth 3 | Set-Content $ledgerPath -Encoding utf8
$ledgerMarkdown = [Text.StringBuilder]::new()
[void] $ledgerMarkdown.AppendLine('# Competitor ledger').AppendLine()
[void] $ledgerMarkdown.AppendLine('Ratio is the competitor''s mean divided by OfficeIMO''s matching lane; above 1 means OfficeIMO is faster.').AppendLine()
[void] $ledgerMarkdown.AppendLine('| Date (UTC) | Commit | Mode | Lane | Benchmark | Parameters | Package | Version | Mean ms | Ratio |')
[void] $ledgerMarkdown.AppendLine('|---|---|---|---|---|---|---|---|---:|---:|')
foreach ($entry in $ledger | Sort-Object Date -Descending -Stable) {
    [void] $ledgerMarkdown.AppendLine("| $($entry.Date) | $($entry.Commit) | $($entry.Mode) | $($entry.Lane) | $($entry.Benchmark) | $($entry.Parameters) | $($entry.Package) | $($entry.Version) | $($entry.MeanMs) | $($entry.Ratio) |")
}
$ledgerMarkdown.ToString() | Set-Content (Join-Path $scratchRoot 'competitor-ledger.md') -Encoding utf8
Write-Host "Ledger: $(Join-Path $scratchRoot 'competitor-ledger.md')"
Remove-Item $temporaryRoot -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "Results: $(Join-Path $OutputRoot 'results.md')"
if (@($laneResults | Where-Object Status -in 'Failed', 'Partial').Count -gt 0) { exit 1 }
