param(
    [string] $SummaryPath = ".\Docs\benchmarks\comparison-current\officeimo.excel.comparison-summary.json",
    [string] $ManifestPath = ".\Docs\benchmarks\comparison-current\officeimo.excel.comparison-suite-manifest.json",
    [string[]] $PackageProfilePath = @(),
    [string] $WebsiteDataPath = ".\Website\data\benchmarks-excel.json",
    [string] $WebsiteSummaryPath = ".\Website\data\benchmarks-excel-summary.json",
    [string] $WebsiteIndexPath = ".\Website\data\benchmarks-excel-index.json",
    [string] $StaticDataDirectory = ".\Website\static\data",
    [string] $MarkdownPath = ".\Docs\benchmarks\officeimo.excel.comparison-report.md",
    [string] $MatrixPartialPath = ".\Website\themes\officeimo\partials\generated\benchmarks-excel.html",
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

function Encode-Html([object] $Value) {
    if ($null -eq $Value) { return "" }
    return [System.Net.WebUtility]::HtmlEncode([string] $Value)
}

function Format-MatrixRatio([Nullable[double]] $Value) {
    if ($null -eq $Value) { return $null }
    if ([math]::Abs($Value - 1.0) -lt 0.005) { return "1x" }
    if ($Value -ge 100.0) {
        return [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:N0}x", $Value)
    }
    if ($Value -ge 10.0) {
        return [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:N1}x", $Value)
    }

    return [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:N2}x", $Value)
}

function Format-SortNumber([Nullable[double]] $Value) {
    if ($null -eq $Value) { return "" }
    return [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:R}", [double] $Value)
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

function Get-EvidenceMeta([object] $Manifest, [object] $Summary, [object[]] $Profiles) {
    $profile = @($Profiles) | Select-Object -First 1
    $machineName = if ($Manifest -and $Manifest.MachineName) { $Manifest.MachineName } elseif ($profile -and $profile.MachineName) { $profile.MachineName } else { $null }
    $runtime = if ($Manifest -and $Manifest.Framework) { $Manifest.Framework } elseif ($Summary -and $Summary.Framework) { $Summary.Framework } elseif ($profile -and $profile.Framework) { $profile.Framework } else { $null }

    return [ordered]@{
        commit = $null
        branch = $null
        dotnetSdk = $null
        runtime = $runtime
        osDescription = $null
        osArchitecture = $null
        processArchitecture = $null
        machineName = $machineName
        processorCount = $null
    }
}

function Get-LatestEvidenceTimestamp([object] $Manifest, [object] $Summary, [object[]] $Profiles) {
    $timestamps = New-Object System.Collections.Generic.List[DateTimeOffset]
    foreach ($value in @($Manifest.GeneratedAtUtc, $Summary.GeneratedAtUtc) + @($Profiles | ForEach-Object { $_.GeneratedAtUtc })) {
        if ([string]::IsNullOrWhiteSpace([string] $value)) { continue }
        if ($value -is [DateTime]) {
            $timestamps.Add([DateTimeOffset]::new(([DateTime] $value).ToUniversalTime()))
            continue
        }
        if ($value -is [DateTimeOffset]) {
            $timestamps.Add(([DateTimeOffset] $value).ToUniversalTime())
            continue
        }
        $parsed = [DateTimeOffset]::MinValue
        if ([DateTimeOffset]::TryParse([string] $value, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind, [ref] $parsed)) {
            $timestamps.Add($parsed.ToUniversalTime())
        }
    }

    if ($timestamps.Count -eq 0) {
        throw "Benchmark evidence does not contain a GeneratedAtUtc timestamp."
    }

    return ($timestamps | Sort-Object -Descending | Select-Object -First 1).ToString("o")
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

function Build-Matrix([object[]] $Rows, [string[]] $LibraryNames) {
    $matrixRows = @()
    $groups = @{}
    foreach ($row in ($Rows | Where-Object { $null -ne $_["meanMilliseconds"] })) {
        $key = "$($row["rowCount"])|$($row["workload"])|$($row["category"])|$($row["scenario"])"
        if (-not $groups.ContainsKey($key)) { $groups[$key] = New-Object System.Collections.Generic.List[object] }
        $groups[$key].Add($row)
    }

    foreach ($group in $groups.Values) {
        $orderedRows = @($group | Sort-Object { $_["meanMilliseconds"] })
        if ($orderedRows.Count -eq 0) { continue }

        $first = $orderedRows[0]
        $fastest = $orderedRows[0]
        $artifactKinds = @($group | ForEach-Object { $_["artifactKind"] } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
        $cells = @()

        foreach ($library in $LibraryNames) {
            $libraryRow = @($group | Where-Object { $_["library"] -eq $library } | Sort-Object { $_["meanMilliseconds"] } | Select-Object -First 1)
            if ($libraryRow.Count -eq 0 -or $null -eq $libraryRow[0]) {
                $cells += [ordered]@{
                    library = $library
                    measured = $false
                style = "missing"
                meanMilliseconds = $null
                meanText = $null
                allocatedText = $null
                packageText = $null
                    ratioToFastest = $null
                    ratioToFastestText = $null
                    allocatedRatioToFastestText = $null
                    packageRatioToFastestText = $null
                }
                continue
            }

            $row = $libraryRow[0]
            $ratio = if ($fastest["meanMilliseconds"] -gt 0) { [double] $row["meanMilliseconds"] / [double] $fastest["meanMilliseconds"] } else { $null }
            $allocatedRatio = if ($row["meanAllocatedBytes"] -and $fastest["meanAllocatedBytes"]) { [double] $row["meanAllocatedBytes"] / [double] $fastest["meanAllocatedBytes"] } else { $null }
            $packageRatio = if ($row["packageBytes"] -and $fastest["packageBytes"]) { [double] $row["packageBytes"] / [double] $fastest["packageBytes"] } else { $null }
            $style = if ($row["library"] -eq $fastest["library"]) { "fastest" } elseif ($ratio -le 1.10) { "near" } elseif ($ratio -le 2.0) { "mid" } else { "slow" }

            $cells += [ordered]@{
                library = $library
                measured = $true
                style = $style
                meanMilliseconds = $row["meanMilliseconds"]
                meanText = $row["meanText"]
                allocatedText = $row["meanAllocatedText"]
                packageText = $row["packageText"]
                ratioToFastest = if ($null -ne $ratio) { [math]::Round($ratio, 4) } else { $null }
                ratioToFastestText = Format-MatrixRatio $ratio
                allocatedRatioToFastestText = Format-MatrixRatio $allocatedRatio
                packageRatioToFastestText = Format-MatrixRatio $packageRatio
            }
        }

        $matrixRows += [ordered]@{
            rowCount = $first["rowCount"]
            workload = $first["workload"]
            category = $first["category"]
            scenario = $first["scenario"]
            artifactKind = ($artifactKinds -join ", ")
            fastestLibrary = $fastest["library"]
            fastestMeanMilliseconds = $fastest["meanMilliseconds"]
            fastestMeanText = $fastest["meanText"]
            fastestAllocatedText = $fastest["meanAllocatedText"]
            fastestPackageText = $fastest["packageText"]
            cells = $cells
        }
    }

    return [ordered]@{
        libraries = $LibraryNames
        rows = @($matrixRows | Sort-Object { $_["rowCount"] }, { $_["workload"] }, { $_["category"] }, { $_["scenario"] })
    }
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

function Format-MatrixHeaderLabel([string] $Value) {
    $encoded = Encode-Html $Value
    $encoded = [regex]::Replace($encoded, '(?<=[a-z])(?=[A-Z])', '<wbr>')
    $encoded = $encoded -replace '\.', '.<wbr>'
    $encoded = $encoded -replace ' ', ' <wbr>'
    return '<span class="imo-benchmark-column-label">' + $encoded + '</span>'
}

function Write-MatrixPartial([object] $Document, [string] $Path) {
    $lines = New-Object System.Collections.Generic.List[string]
    $metrics = $Document.metrics
    $matrix = $Document.matrix
    $meta = $Document.meta
    $source = $Document.source

    $lines.Add('<section class="imo-benchmark-dashboard" data-excel-benchmarks>')
    $lines.Add('<div class="imo-benchmark-hero">')
    $lines.Add('<div class="imo-benchmark-hero__copy">')
    $lines.Add('<p class="imo-benchmark-eyebrow">Excel engineering detail</p>')
    $lines.Add('<h2>Full comparison matrix</h2>')
    $lines.Add('<p>Each row is a scenario group. Library cells show mean time, relative time versus the fastest library in that row, and allocation or package size when available.</p>')
    $lines.Add('</div>')
    $lines.Add('</div>')
    $lines.Add('<div class="imo-benchmark-meta" data-benchmark-meta>')
    $lines.Add('<span>Generated ' + (Encode-Html $Document.generatedUtc) + '</span>')
    $lines.Add('<span>' + (Encode-Html $Document.runMode) + '</span>')
    $publishText = if ($Document.publish) { 'publishable' } else { 'local quick' }
    $lines.Add('<span>' + (Encode-Html $publishText) + '</span>')
    $lines.Add('<span>' + (Encode-Html $Document.framework) + '</span>')
    if ($meta -and $meta.dotnetSdk) {
        $lines.Add('<span>SDK ' + (Encode-Html $meta.dotnetSdk) + '</span>')
    }
    if ($source -and $source.summaryPath) {
        $lines.Add('<span>Evidence: committed Excel comparison artifacts</span>')
    }
    $lines.Add('</div>')
    $lines.Add('<div class="imo-benchmark-kpis" data-benchmark-kpis>')
    $lines.Add('<article><strong>' + (Encode-Html $metrics.matrixRows) + '</strong><span>scenario rows</span></article>')
    $lines.Add('<article><strong>' + (Encode-Html $metrics.measurementRows) + '</strong><span>measurements</span></article>')
    $lines.Add('<article><strong>' + (Encode-Html $metrics.libraryCount) + '</strong><span>libraries</span></article>')
    $lines.Add('<article><strong>' + (Encode-Html @($metrics.rowCounts).Count) + '</strong><span>row tiers: ' + (Encode-Html (@($metrics.rowCounts) -join ', ')) + '</span></article>')
    $lines.Add('<article><strong>' + (Encode-Html $metrics.artifactKindCount) + '</strong><span>artifact types</span></article>')
    $lines.Add('<article><strong>' + (Encode-Html $metrics.focusedProfileCount) + '</strong><span>focused profiles</span></article>')
    $lines.Add('</div>')
    $lines.Add('<section class="imo-benchmark-matrix">')
    $lines.Add('<div class="imo-benchmark-matrix__header">')
    $lines.Add('<div>')
    $lines.Add('<h3>Comparison Matrix</h3>')
    $lines.Add('<p>Ratios are relative to the fastest measured library in the same scenario row. Lower is better.</p>')
    $lines.Add('</div>')
    $lines.Add('</div>')
    $lines.Add('<div class="imo-benchmark-tools" data-benchmark-controls>')
    $lines.Add('<label><span>Search</span><input type="search" data-benchmark-filter="search" placeholder="Scenario, category, library"></label>')
    $lines.Add('<label><span>Rows</span><select data-benchmark-filter="rowCount"><option value="">All rows</option>')
    foreach ($rowCount in (@($matrix.rows | ForEach-Object { $_.rowCount }) | Sort-Object -Unique)) {
        $lines.Add('<option value="' + (Encode-Html $rowCount) + '">' + (Encode-Html $rowCount) + '</option>')
    }
    $lines.Add('</select></label>')
    $lines.Add('<label><span>Workload</span><select data-benchmark-filter="workload"><option value="">All workloads</option>')
    foreach ($workload in (@($matrix.rows | ForEach-Object { $_.workload } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) | Sort-Object -Unique)) {
        $lines.Add('<option value="' + (Encode-Html $workload) + '">' + (Encode-Html $workload) + '</option>')
    }
    $lines.Add('</select></label>')
    $lines.Add('<label><span>Category</span><select data-benchmark-filter="category"><option value="">All categories</option>')
    foreach ($category in (@($matrix.rows | ForEach-Object { $_.category } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) | Sort-Object -Unique)) {
        $lines.Add('<option value="' + (Encode-Html $category) + '">' + (Encode-Html $category) + '</option>')
    }
    $lines.Add('</select></label>')
    $lines.Add('<label><span>Library</span><select data-benchmark-filter="library"><option value="">Any library</option>')
    foreach ($library in $matrix.libraries) {
        $lines.Add('<option value="' + (Encode-Html $library) + '">' + (Encode-Html $library) + '</option>')
    }
    $lines.Add('</select></label>')
    $lines.Add('<label class="imo-benchmark-sort-mode"><span>Sort metric</span><select data-benchmark-sort-mode><option value="time">Time, then relative</option><option value="ratio">Relative, then time</option></select></label>')
    $lines.Add('<button type="button" class="imo-benchmark-reset" data-benchmark-reset>Reset</button>')
    $lines.Add('</div>')
    $lines.Add('<div class="imo-benchmark-table-wrap">')
    $lines.Add('<table class="imo-benchmark-table imo-benchmark-table--matrix">')
    $lines.Add('<thead>')
    $lines.Add('<tr>')
    $lines.Add('<th aria-sort="none"><button type="button" data-benchmark-sort="scenario" data-sort-type="text">Scenario</button></th>')
    $lines.Add('<th aria-sort="none"><button type="button" data-benchmark-sort="fastest" data-sort-type="number">Fastest</button></th>')
    foreach ($library in $matrix.libraries) {
        $lines.Add('<th aria-sort="none"><button type="button" data-benchmark-sort="library:' + (Encode-Html $library) + '" data-sort-type="number">' + (Format-MatrixHeaderLabel $library) + '</button></th>')
    }
    $lines.Add('</tr>')
    $lines.Add('</thead>')
    $lines.Add('<tbody data-benchmark-matrix>')
    $rowIndex = 0
    foreach ($row in $matrix.rows) {
        $lines.Add('<tr data-benchmark-row data-original-index="' + (Encode-Html $rowIndex) + '" data-row-count="' + (Encode-Html $row.rowCount) + '" data-workload="' + (Encode-Html $row.workload) + '" data-category="' + (Encode-Html $row.category) + '" data-scenario="' + (Encode-Html $row.scenario) + '" data-fastest-library="' + (Encode-Html $row.fastestLibrary) + '" data-fastest-ms="' + (Encode-Html (Format-SortNumber $row.fastestMeanMilliseconds)) + '">')
        $scenarioMeta = ([string] $row.rowCount) + ' rows - ' + ([string] $row.workload) + ' - ' + ([string] $row.category) + ' - ' + ([string] $row.artifactKind)
        $lines.Add('<td class="imo-benchmark-scenario" data-label="Scenario"><strong>' + (Encode-Html $row.scenario) + '</strong><small>' + (Encode-Html $scenarioMeta) + '</small></td>')
        $lines.Add('<td class="imo-benchmark-fastest" data-label="Fastest"><strong>' + (Encode-Html $row.fastestLibrary) + '</strong><small>' + (Encode-Html $row.fastestMeanText) + '</small></td>')
        foreach ($cell in $row.cells) {
            $cellClasses = 'imo-benchmark-value imo-benchmark-value--' + ([string] $cell.style)
            if ($cell.measured) {
                $cellAttributes = ' data-mean-ms="' + (Encode-Html (Format-SortNumber $cell.meanMilliseconds)) + '"'
                if ($null -ne $cell.ratioToFastest) {
                    $cellAttributes += ' data-ratio-to-fastest="' + (Encode-Html (Format-SortNumber $cell.ratioToFastest)) + '"'
                }
            } else {
                $cellClasses += ' imo-benchmark-value--missing'
                $cellAttributes = ''
            }
            $lines.Add('<td class="' + (Encode-Html $cellClasses) + '" data-library="' + (Encode-Html $cell.library) + '" data-label="' + (Encode-Html $cell.library) + '"' + $cellAttributes + '>')
            if ($cell.measured) {
                $lines.Add('<strong>' + (Encode-Html $cell.meanText) + '</strong>')
                $lines.Add('<span>' + (Encode-Html $cell.ratioToFastestText) + '</span>')
                if ($cell.packageText) {
                    $lines.Add('<small>' + (Encode-Html $cell.packageText) + '</small>')
                } elseif ($cell.allocatedText) {
                    $lines.Add('<small>' + (Encode-Html $cell.allocatedText) + '</small>')
                }
            } else {
                $lines.Add('<span class="imo-benchmark-missing">-</span>')
            }
            $lines.Add('</td>')
        }
        $lines.Add('</tr>')
        $rowIndex++
    }
    $lines.Add('</tbody>')
    $lines.Add('</table>')
    $lines.Add('</div>')
    $lines.Add('<p class="imo-benchmark-count" data-benchmark-count>Showing ' + (Encode-Html @($matrix.rows).Count) + ' of ' + (Encode-Html @($matrix.rows).Count) + ' rows</p>')
    $lines.Add('</section>')
    $lines.Add('</section>')

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
$packageProfiles = New-Object System.Collections.Generic.List[object]
foreach ($row in $summaryInput.Rows) {
    $rows.Add((Convert-SummaryRow $row))
}

foreach ($path in $packageProfilePaths) {
    if ([string]::IsNullOrWhiteSpace($path)) { continue }
    $profile = Read-JsonFile $path
    $packageProfiles.Add($profile)
    foreach ($scenario in $profile.Scenarios) {
        $rows.Add((Convert-PackageProfileScenario $scenario ([int] $profile.RowCount) $path))
    }
}

$allRows = [object[]] $rows.ToArray()
$allPackageProfiles = [object[]] $packageProfiles.ToArray()
$focusedPackageRows = [object[]] @($allRows | Where-Object { $_["artifactKind"] -eq "focused-package-profile" })
Add-PackageProfileRatios -Rows $focusedPackageRows
$scenarioGroupCount = @($allRows | Group-Object { "$($_["artifactKind"])|$($_["rowCount"])|$($_["scenario"])" }).Count
$libraryNames = @($allRows | ForEach-Object { $_["library"] } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
$rowCounts = @($allRows | ForEach-Object { $_["rowCount"] } | Where-Object { $null -ne $_ } | Sort-Object -Unique)
$artifactKinds = @($allRows | ForEach-Object { $_["artifactKind"] } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
$matrix = Build-Matrix -Rows $allRows -LibraryNames $libraryNames
$metrics = [ordered]@{
    scenarioGroups = $scenarioGroupCount
    matrixRows = @($matrix.rows).Count
    measurementRows = $allRows.Count
    libraries = $libraryNames
    libraryCount = $libraryNames.Count
    rowCounts = $rowCounts
    artifactKinds = $artifactKinds
    artifactKindCount = $artifactKinds.Count
    focusedProfileCount = $packageProfilePaths.Count
}

$document = [ordered]@{
    schemaVersion = 1
    generatedUtc = Get-LatestEvidenceTimestamp -Manifest $manifest -Summary $summaryInput -Profiles $allPackageProfiles
    runMode = $RunMode
    publish = $publishValue
    framework = if ($manifest) { $manifest.Framework } else { $summaryInput.Framework }
    configuration = if ($manifest) { "Release" } else { $null }
    source = [ordered]@{
        summaryPath = $SummaryPath
        manifestPath = if (Test-Path $ManifestPath) { $ManifestPath } else { $null }
        packageProfilePaths = $packageProfilePaths
    }
    meta = Get-EvidenceMeta -Manifest $manifest -Summary $summaryInput -Profiles $allPackageProfiles
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
        "Generated provenance comes from the committed benchmark evidence, not the website build environment.",
        "Website consumers should read this generated JSON instead of hand-maintained benchmark rows."
    )
    manifest = $manifest
    metrics = $metrics
    matrix = $matrix
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
    metrics = $document.metrics
    matrix = $document.matrix
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
Write-MatrixPartial $document $MatrixPartialPath

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
Write-Host "Excel benchmark matrix partial written to '$MatrixPartialPath'."
