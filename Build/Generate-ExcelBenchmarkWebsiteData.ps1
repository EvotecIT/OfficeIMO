param(
    [string] $SummaryPath = ".\Docs\benchmarks\comparison-current\officeimo.excel.comparison-summary.json",
    [string] $ManifestPath = ".\Docs\benchmarks\comparison-current\officeimo.excel.comparison-suite-manifest.json",
    [string[]] $PackageProfilePath = @(),
    [string] $WebsiteDataPath = ".\Website\data\benchmarks-excel.json",
    [string] $WebsiteSummaryPath = ".\Website\data\benchmarks-excel-summary.json",
    [string] $WebsiteIndexPath = ".\Website\data\benchmarks-excel-index.json",
    [string] $StaticDataDirectory = ".\Website\static\data",
    [string] $MarkdownPath = ".\Docs\benchmarks\officeimo.excel.comparison-report.md",
    [ValidateSet("quick", "full")]
    [string] $RunMode = "quick",
    [switch] $Publish,
    [switch] $NoPublish
)

$ErrorActionPreference = "Stop"

function Write-TextUtf8NoBom([string] $Path, [string] $Value) {
    $directory = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path $directory)) {
        New-Item -ItemType Directory -Force -Path $directory | Out-Null
    }

    $encoding = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::WriteAllText((Resolve-OrCreatePath $Path), $Value, $encoding)
}

function Resolve-OrCreatePath([string] $Path) {
    $full = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $Path))
    return $full
}

function Read-JsonFile([string] $Path) {
    if (-not (Test-Path $Path)) {
        throw "JSON file not found: $Path"
    }

    return Get-Content -Path $Path -Raw -Encoding UTF8 | ConvertFrom-Json
}

function Format-Milliseconds([Nullable[double]] $Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -ge 1000.0) {
        return [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:N2} s", ($Value / 1000.0))
    }

    return [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:N2} ms", $Value)
}

function Format-Bytes([Nullable[double]] $Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -ge 1GB) { return [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:N2} GB", ($Value / 1GB)) }
    if ($Value -ge 1MB) { return [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:N1} MB", ($Value / 1MB)) }
    if ($Value -ge 1KB) { return [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:N1} KB", ($Value / 1KB)) }
    return [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:N0} B", $Value)
}

function Format-Ratio([Nullable[double]] $Value) {
    if ($null -eq $Value) { return $null }
    return [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:N2}x", $Value)
}

function Get-ScenarioCategory([string] $Scenario, [string] $ArtifactKind) {
    if ($ArtifactKind -like "*package*") { return "Package size" }
    if ($Scenario -like "*blog-2023*") { return "Plain string export" }
    if ($Scenario -like "*autofit*") { return "AutoFit and mutation" }
    if ($Scenario -like "read-*" -or $Scenario -like "enumerate-*") {
        if ($Scenario -like "*objects*") { return "Typed object read" }
        if ($Scenario -like "*stream*") { return "Streaming read" }
        if ($Scenario -like "*sparse*") { return "Sparse read" }
        if ($Scenario -like "*shared-string*" -or $Scenario -like "*helloworld*") { return "Dense string read" }
        return "Range and table read"
    }
    if ($Scenario -like "build-object-datatable*") { return "Object projection" }
    if ($Scenario -like "*shared-strings*" -or $Scenario -like "*cellvalue-strings*") { return "Shared string write" }
    if ($Scenario -like "*formula*") { return "Formula write/read" }
    if ($Scenario -eq "write-bulk-report") { return "Formatted report write" }
    if ($Scenario -like "report-workbook*") { return "Report workbook" }
    if ($Scenario -like "realworld-report-*") {
        if ($Scenario -eq "realworld-report-all-in-one" -or $Scenario -eq "realworld-report-core") { return "Real-world report" }
        return "Anti-cheat report variants"
    }
    if ($Scenario -like "realworld-*") { return "Real-world feature mix" }
    if ($Scenario -like "*dataset*") {
        if ($Scenario -like "*direct-export*") { return "Plain streaming export" }
        return "DataSet table export"
    }
    if ($Scenario -like "*datatable*" -or $Scenario -like "*datareader*") {
        if ($Scenario -like "*plain*") { return "Plain streaming export" }
        return "DataTable table export"
    }
    if ($Scenario -like "*insertobjects*" -or $Scenario -like "*rowsfrom*") { return "Typed object export" }
    if ($Scenario -like "*cellvalues*" -or $Scenario -eq "append-plain-rows") { return "Plain cell export" }
    if ($Scenario -like "*cellvalue-*") { return "Cell writer" }
    return "Other"
}

function Get-Meta {
    $dotnetSdk = $null
    if (Get-Command dotnet -ErrorAction SilentlyContinue) {
        try { $dotnetSdk = (& dotnet --version).Trim() } catch { $dotnetSdk = $null }
    }

    $commit = $null
    $branch = $null
    if (Get-Command git -ErrorAction SilentlyContinue) {
        try { $commit = (& git rev-parse HEAD).Trim() } catch { $commit = $null }
        try { $branch = (& git rev-parse --abbrev-ref HEAD).Trim() } catch { $branch = $null }
    }

    return [ordered]@{
        commit = $commit
        branch = $branch
        dotnetSdk = $dotnetSdk
        runtime = [System.Runtime.InteropServices.RuntimeInformation]::FrameworkDescription
        osDescription = [System.Runtime.InteropServices.RuntimeInformation]::OSDescription
        osArchitecture = [System.Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString()
        processArchitecture = [System.Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture.ToString()
        machineName = [Environment]::MachineName
        processorCount = [Environment]::ProcessorCount
    }
}

function Convert-SummaryRow([object] $Row) {
    $packageBytes = if ($null -ne $Row.PackageBytes) { [double] $Row.PackageBytes } else { $null }
    $allocatedBytes = if ($null -ne $Row.MeanAllocatedBytes) { [double] $Row.MeanAllocatedBytes } else { $null }
    $meanMs = if ($null -ne $Row.MeanMilliseconds) { [double] $Row.MeanMilliseconds } else { $null }
    $medianMs = if ($null -ne $Row.MedianMilliseconds) { [double] $Row.MedianMilliseconds } else { $null }

    return [ordered]@{
        artifactKind = $Row.ArtifactKind
        rowCount = [int] $Row.RowCount
        workload = $Row.Workload
        category = Get-ScenarioCategory $Row.Scenario $Row.ArtifactKind
        scenario = $Row.Scenario
        library = $Row.Library
        notes = $Row.Notes
        meanMilliseconds = $meanMs
        meanText = Format-Milliseconds $meanMs
        medianMilliseconds = $medianMs
        medianText = Format-Milliseconds $medianMs
        standardDeviationMilliseconds = $Row.StandardDeviationMilliseconds
        standardErrorMilliseconds = $Row.StandardErrorMilliseconds
        meanAllocatedBytes = $allocatedBytes
        meanAllocatedText = Format-Bytes $allocatedBytes
        packageBytes = $packageBytes
        packageText = Format-Bytes $packageBytes
        bestLibrary = $Row.BestLibrary
        bestMeanMilliseconds = $Row.BestMeanMilliseconds
        ratioToOfficeImo = $Row.RatioToOfficeImo
        ratioToOfficeImoText = Format-Ratio $Row.RatioToOfficeImo
        ratioToBest = $Row.RatioToBest
        ratioToBestText = Format-Ratio $Row.RatioToBest
        allocatedRatioToOfficeImo = $Row.AllocatedRatioToOfficeImo
        allocatedRatioToOfficeImoText = Format-Ratio $Row.AllocatedRatioToOfficeImo
        packageRatioToOfficeImo = $Row.PackageRatioToOfficeImo
        packageRatioToOfficeImoText = Format-Ratio $Row.PackageRatioToOfficeImo
        outcome = $Row.Outcome
        artifactPath = $Row.ArtifactPath
    }
}

function Convert-PackageProfileScenario([object] $Scenario, [int] $RowCount, [string] $SourcePath) {
    $meanMs = if ($null -ne $Scenario.AverageMilliseconds) { [double] $Scenario.AverageMilliseconds } else { $null }
    $medianMs = if ($null -ne $Scenario.MedianMilliseconds) { [double] $Scenario.MedianMilliseconds } else { $null }
    $allocatedBytes = if ($null -ne $Scenario.AverageAllocatedBytes) { [double] $Scenario.AverageAllocatedBytes } else { $null }
    $packageBytes = if ($Scenario.Package -and $null -ne $Scenario.Package.FileSizeBytes) { [double] $Scenario.Package.FileSizeBytes } else { $null }

    return [ordered]@{
        artifactKind = "focused-package-profile"
        rowCount = $RowCount
        workload = "package"
        category = "Package size"
        scenario = $Scenario.Scenario
        library = $Scenario.Library
        notes = $Scenario.Notes
        meanMilliseconds = $meanMs
        meanText = Format-Milliseconds $meanMs
        medianMilliseconds = $medianMs
        medianText = Format-Milliseconds $medianMs
        standardDeviationMilliseconds = $Scenario.StandardDeviationMilliseconds
        standardErrorMilliseconds = $Scenario.StandardErrorMilliseconds
        meanAllocatedBytes = $allocatedBytes
        meanAllocatedText = Format-Bytes $allocatedBytes
        packageBytes = $packageBytes
        packageText = Format-Bytes $packageBytes
        worksheetRowCount = if ($Scenario.Package) { $Scenario.Package.WorksheetRowCount } else { $null }
        worksheetCellCount = if ($Scenario.Package) { $Scenario.Package.WorksheetCellCount } else { $null }
        sharedStringCount = if ($Scenario.Package) { $Scenario.Package.SharedStringCount } else { $null }
        uniqueSharedStringCount = if ($Scenario.Package) { $Scenario.Package.UniqueSharedStringCount } else { $null }
        sourcePath = $SourcePath
    }
}

function Add-PackageProfileRatios([object[]] $Rows) {
    $groups = @{}
    foreach ($row in $Rows) {
        $key = "$($row["rowCount"])|$($row["scenario"])"
        if (-not $groups.ContainsKey($key)) { $groups[$key] = New-Object System.Collections.Generic.List[object] }
        $groups[$key].Add($row)
    }

    foreach ($group in $groups.Values) {
        $office = $group | Where-Object { $_["library"] -eq "OfficeIMO.Excel" } | Select-Object -First 1
        $best = $group | Sort-Object { $_["meanMilliseconds"] } | Select-Object -First 1
        foreach ($row in $group) {
            if ($office -and $office["meanMilliseconds"] -gt 0) {
                $row["ratioToOfficeImo"] = [math]::Round($row["meanMilliseconds"] / $office["meanMilliseconds"], 4)
                $row["ratioToOfficeImoText"] = Format-Ratio $row["ratioToOfficeImo"]
                if ($row["meanAllocatedBytes"] -and $office["meanAllocatedBytes"]) {
                    $row["allocatedRatioToOfficeImo"] = [math]::Round($row["meanAllocatedBytes"] / $office["meanAllocatedBytes"], 4)
                    $row["allocatedRatioToOfficeImoText"] = Format-Ratio $row["allocatedRatioToOfficeImo"]
                }
                if ($row["packageBytes"] -and $office["packageBytes"]) {
                    $row["packageRatioToOfficeImo"] = [math]::Round($row["packageBytes"] / $office["packageBytes"], 4)
                    $row["packageRatioToOfficeImoText"] = Format-Ratio $row["packageRatioToOfficeImo"]
                }
            }
            if ($best -and $best["meanMilliseconds"] -gt 0) {
                $row["bestLibrary"] = $best["library"]
                $row["bestMeanMilliseconds"] = $best["meanMilliseconds"]
                $row["ratioToBest"] = [math]::Round($row["meanMilliseconds"] / $best["meanMilliseconds"], 4)
                $row["ratioToBestText"] = Format-Ratio $row["ratioToBest"]
                $row["outcome"] = if ($row["library"] -eq $best["library"]) { "Win" } else { $row["ratioToBestText"] + " vs best" }
            }
        }
    }
}

function Build-Summary([object[]] $Rows) {
    $summary = @()
    $groups = @{}
    foreach ($row in ($Rows | Where-Object { $_["library"] -eq "OfficeIMO.Excel" })) {
        $key = "$($row["rowCount"])|$($row["artifactKind"])|$($row["workload"])|$($row["category"])"
        if (-not $groups.ContainsKey($key)) { $groups[$key] = New-Object System.Collections.Generic.List[object] }
        $groups[$key].Add($row)
    }

    foreach ($group in $groups.Values) {
        $first = $group[0]
        $losses = @($group | Where-Object { $_["outcome"] -and $_["outcome"] -ne "Win" })
        $wins = @($group | Where-Object { $_["outcome"] -eq "Win" })
        $biggestLoss = $losses | Sort-Object { $_["ratioToBest"] } -Descending | Select-Object -First 1
        $summary += [ordered]@{
            rowCount = $first["rowCount"]
            artifactKind = $first["artifactKind"]
            workload = $first["workload"]
            category = $first["category"]
            officeImoWins = $wins.Count
            officeImoLosses = $losses.Count
            biggestLossScenario = if ($biggestLoss) { $biggestLoss["scenario"] } else { $null }
            biggestLossBestLibrary = if ($biggestLoss) { $biggestLoss["bestLibrary"] } else { $null }
            biggestLossRatioToBest = if ($biggestLoss) { $biggestLoss["ratioToBest"] } else { $null }
            biggestLossRatioToBestText = if ($biggestLoss) { $biggestLoss["ratioToBestText"] } else { $null }
        }
    }
    return @($summary | Sort-Object { $_["rowCount"] }, { $_["artifactKind"] }, { $_["workload"] }, { $_["category"] })
}

function Write-MarkdownReport([object] $Document, [string] $Path) {
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("# OfficeIMO.Excel Benchmark Report")
    $lines.Add("")
    $lines.Add("Generated: $($Document.generatedUtc)")
    $lines.Add("Run mode: $($Document.runMode)")
    $lines.Add("Publish: $($Document.publish)")
    $lines.Add("Machine: $($Document.meta.machineName) ($($Document.meta.processorCount) processors)")
    $lines.Add("")
    $lines.Add("## How to Read")
    foreach ($note in $Document.howToRead) {
        $lines.Add("- $note")
    }
    $lines.Add("")
    $lines.Add("## At a Glance")
    $lines.Add("")
    $lines.Add("| Rows | Artifact | Workload | Category | OfficeIMO wins | OfficeIMO losses | Biggest loss |")
    $lines.Add("| ---: | --- | --- | --- | ---: | ---: | --- |")
    foreach ($row in $Document.summary) {
        $loss = if ($row["biggestLossScenario"]) { "$($row["biggestLossScenario"]) vs $($row["biggestLossBestLibrary"]) ($($row["biggestLossRatioToBestText"]))" } else { "" }
        $lines.Add("| $($row["rowCount"]) | $($row["artifactKind"]) | $($row["workload"]) | $($row["category"]) | $($row["officeImoWins"]) | $($row["officeImoLosses"]) | $loss |")
    }
    $lines.Add("")
    $lines.Add("## Rows")
    $lines.Add("")
    $lines.Add("| Rows | Artifact | Scenario | Library | Mean | Allocated | Package | Best | Outcome |")
    $lines.Add("| ---: | --- | --- | --- | ---: | ---: | ---: | --- | --- |")
    foreach ($row in ($Document.rows | Sort-Object { $_["rowCount"] }, { $_["artifactKind"] }, { $_["scenario"] }, { $_["meanMilliseconds"] })) {
        $lines.Add("| $($row["rowCount"]) | $($row["artifactKind"]) | $($row["scenario"]) | $($row["library"]) | $($row["meanText"]) | $($row["meanAllocatedText"]) | $($row["packageText"]) | $($row["bestLibrary"]) | $($row["outcome"]) |")
    }

    Write-TextUtf8NoBom $Path (($lines -join "`n") + "`n")
}

$summaryInput = Read-JsonFile $SummaryPath
$manifest = if (Test-Path $ManifestPath) { Read-JsonFile $ManifestPath } else { $null }
$publishValue = if ($Publish) { $true } elseif ($NoPublish) { $false } else { $RunMode -eq "full" }
$packageProfilePaths = @(
    foreach ($path in $PackageProfilePath) {
        if ([string]::IsNullOrWhiteSpace($path)) { continue }
        foreach ($expandedPath in ($path -split "[,;]")) {
            if (-not [string]::IsNullOrWhiteSpace($expandedPath)) {
                $expandedPath.Trim()
            }
        }
    }
)

$rows = New-Object System.Collections.Generic.List[object]
foreach ($row in $summaryInput.Rows) {
    $rows.Add((Convert-SummaryRow $row))
}

foreach ($path in $packageProfilePaths) {
    if ([string]::IsNullOrWhiteSpace($path)) { continue }
    $profile = Read-JsonFile $path
    foreach ($scenario in $profile.Scenarios) {
        $rows.Add((Convert-PackageProfileScenario $scenario ([int] $profile.RowCount) $path))
    }
}

$allRows = [object[]] $rows.ToArray()
$focusedPackageRows = [object[]] @($allRows | Where-Object { $_["artifactKind"] -eq "focused-package-profile" })
Add-PackageProfileRatios -Rows $focusedPackageRows

$document = [ordered]@{
    schemaVersion = 1
    generatedUtc = (Get-Date).ToUniversalTime().ToString("o")
    runMode = $RunMode
    publish = $publishValue
    framework = if ($manifest) { $manifest.Framework } else { $summaryInput.Framework }
    configuration = if ($manifest) { "Release" } else { $null }
    source = [ordered]@{
        summaryPath = $SummaryPath
        manifestPath = if (Test-Path $ManifestPath) { $ManifestPath } else { $null }
        packageProfilePaths = $packageProfilePaths
    }
    meta = Get-Meta
    howToRead = @(
        "Mean: average elapsed time for the measured operation. Lower is better.",
        "Allocated: managed memory allocated by the measured operation. Lower is better.",
        "Package: generated XLSX file size when package profiling is available.",
        "Ratio to OfficeIMO compares each library against OfficeIMO for the same scenario and row count.",
        "Quick runs are useful for engineering direction; full runs should be used for public claims.",
        "Benchmarks are machine-specific and should be treated as reproducible evidence, not universal guarantees."
    )
    notes = @(
        "Generated from OfficeIMO.Excel comparison-suite artifacts.",
        "Additional focused package-profile artifacts can be included without changing the benchmark harness.",
        "Website consumers should read this generated JSON instead of hand-maintained benchmark rows."
    )
    manifest = $manifest
    summary = Build-Summary -Rows $allRows
    rows = @($allRows | Sort-Object { $_["rowCount"] }, { $_["artifactKind"] }, { $_["scenario"] }, { $_["library"] })
}

$summaryDocument = [ordered]@{
    schemaVersion = $document.schemaVersion
    generatedUtc = $document.generatedUtc
    runMode = $document.runMode
    publish = $document.publish
    framework = $document.framework
    meta = $document.meta
    source = $document.source
    howToRead = $document.howToRead
    notes = $document.notes
    summary = $document.summary
}

$indexDocument = [ordered]@{
    schemaVersion = 1
    entries = @(
        [ordered]@{
            generatedUtc = $document.generatedUtc
            runMode = $document.runMode
            publish = $document.publish
            framework = $document.framework
            summaryPath = $WebsiteSummaryPath
            dataPath = $WebsiteDataPath
            markdownPath = $MarkdownPath
            meta = $document.meta
        }
    )
}

Write-TextUtf8NoBom $WebsiteDataPath (($document | ConvertTo-Json -Depth 20) + "`n")
Write-TextUtf8NoBom $WebsiteSummaryPath (($summaryDocument | ConvertTo-Json -Depth 16) + "`n")
Write-TextUtf8NoBom $WebsiteIndexPath (($indexDocument | ConvertTo-Json -Depth 10) + "`n")
Write-MarkdownReport $document $MarkdownPath

if (-not [string]::IsNullOrWhiteSpace($StaticDataDirectory)) {
    if (-not (Test-Path $StaticDataDirectory)) {
        New-Item -ItemType Directory -Force -Path $StaticDataDirectory | Out-Null
    }

    Copy-Item -Path $WebsiteDataPath -Destination (Join-Path $StaticDataDirectory "benchmarks-excel.json") -Force
    Copy-Item -Path $WebsiteSummaryPath -Destination (Join-Path $StaticDataDirectory "benchmarks-excel-summary.json") -Force
    Copy-Item -Path $WebsiteIndexPath -Destination (Join-Path $StaticDataDirectory "benchmarks-excel-index.json") -Force
}

Write-Host "Excel benchmark website data written to '$WebsiteDataPath'."
Write-Host "Excel benchmark website summary written to '$WebsiteSummaryPath'."
Write-Host "Excel benchmark index written to '$WebsiteIndexPath'."
if (-not [string]::IsNullOrWhiteSpace($StaticDataDirectory)) {
    Write-Host "Excel benchmark static data copied to '$StaticDataDirectory'."
}
Write-Host "Excel benchmark markdown report written to '$MarkdownPath'."
