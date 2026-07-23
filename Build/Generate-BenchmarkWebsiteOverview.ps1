param(
    [string] $ExcelHighlightsPath = ".\Docs\benchmarks\readme-current\officeimo.excel.comparison.json",
    [string] $CsvHighlightsPath = ".\Docs\benchmarks\readme-current\officeimo.csv.comparison.json",
    [string] $OutputPath = ".\Website\themes\officeimo\partials\generated\benchmarks-overview.html"
)

$ErrorActionPreference = 'Stop'

function Resolve-RequiredPath([string] $Path, [string] $Label) {
    $resolved = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $Path))
    if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) {
        throw "$Label was not found at '$resolved'."
    }
    return $resolved
}

function Encode-Html([object] $Value) {
    return [System.Net.WebUtility]::HtmlEncode([string] $Value)
}

function Format-Milliseconds([object] $Value) {
    $number = [double] $Value
    return [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, '{0:0.00} ms', $number)
}

function Format-Ratio([object] $Value, [bool] $IsOfficeImo) {
    if ($IsOfficeImo) {
        return 'OfficeIMO baseline'
    }

    $number = [double] $Value
    return [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, '{0:0.00}x OfficeIMO', $number)
}

function Get-SnapshotLabel([object] $Document) {
    $first = @($Document.comparison)[0]
    if ($first -and $first.variables -and $first.variables.Snapshot) {
        return [string] $first.variables.Snapshot
    }

    if ($Document.metadata.generatedAtUtc) {
        return ([DateTimeOffset]::Parse([string] $Document.metadata.generatedAtUtc)).ToString('yyyy-MM-dd')
    }

    return 'current committed snapshot'
}

function Add-ComparisonFamily(
    [System.Collections.Generic.List[string]] $Lines,
    [object] $Document,
    [string] $Family,
    [string] $Title,
    [string] $Description,
    [string] $MetricLabel,
    [string] $EvidenceLabel,
    [string] $EvidenceUrl
) {
    $scenarioOrder = [System.Collections.Generic.List[string]]::new()
    $scenarioRows = [ordered]@{}
    foreach ($row in @($Document.comparison)) {
        $scenario = [string] $row.scenario
        if (-not $scenarioRows.Contains($scenario)) {
            $scenarioRows[$scenario] = [System.Collections.Generic.List[object]]::new()
            $scenarioOrder.Add($scenario)
        }
        $scenarioRows[$scenario].Add($row)
    }

    $snapshot = Get-SnapshotLabel $Document
    $Lines.Add('<section class="imo-benchmark-family" id="' + (Encode-Html $Family) + '-evidence" data-benchmark-family="' + (Encode-Html $Family) + '">')
    $Lines.Add('<header class="imo-benchmark-family__header">')
    $Lines.Add('<div><p class="imo-benchmark-eyebrow">Published comparison</p><h2>' + (Encode-Html $Title) + '</h2></div>')
    $Lines.Add('<p>' + (Encode-Html $Description) + '</p>')
    $Lines.Add('</header>')
    $Lines.Add('<div class="imo-benchmark-family__meta"><span>Snapshot ' + (Encode-Html $snapshot) + '</span><span>25,000 rows</span><span>.NET 8</span><a href="' + (Encode-Html $EvidenceUrl) + '" target="_blank" rel="noopener">' + (Encode-Html $EvidenceLabel) + '</a></div>')
    $Lines.Add('<div class="imo-benchmark-highlight-grid">')

    foreach ($scenario in $scenarioOrder) {
        $rows = @($scenarioRows[$scenario])
        $operation = [string] $rows[0].operation
        $contract = if ($rows[0].variables -and $rows[0].variables.Contract) { [string] $rows[0].variables.Contract } else { $operation }
        $Lines.Add('<article class="imo-benchmark-highlight">')
        $Lines.Add('<div class="imo-benchmark-highlight__heading"><div><p>' + (Encode-Html $operation) + '</p><h3>' + (Encode-Html $scenario) + '</h3></div><span>' + (Encode-Html $contract) + '</span></div>')
        $Lines.Add('<div class="imo-benchmark-highlight__table-wrap"><table><thead><tr><th scope="col">Library</th><th scope="col">' + (Encode-Html $MetricLabel) + '</th><th scope="col">Relative</th></tr></thead><tbody>')
        foreach ($row in $rows) {
            $isOfficeImo = ([string] $row.engine).StartsWith('OfficeIMO', [System.StringComparison]::OrdinalIgnoreCase)
            $rowClass = if ($isOfficeImo) { ' class="is-officeimo"' } else { '' }
            $Lines.Add('<tr' + $rowClass + '><th scope="row">' + (Encode-Html $row.engine) + '</th><td>' + (Encode-Html (Format-Milliseconds $row.actual)) + '</td><td>' + (Encode-Html (Format-Ratio $row.ratio $isOfficeImo)) + '</td></tr>')
        }
        $Lines.Add('</tbody></table></div>')
        $Lines.Add('</article>')
    }

    $Lines.Add('</div>')
    $Lines.Add('</section>')
}

$excelPath = Resolve-RequiredPath $ExcelHighlightsPath 'Excel benchmark highlights'
$csvPath = Resolve-RequiredPath $CsvHighlightsPath 'CSV benchmark highlights'
$excel = Get-Content -LiteralPath $excelPath -Raw -Encoding UTF8 | ConvertFrom-Json -Depth 100
$csv = Get-Content -LiteralPath $csvPath -Raw -Encoding UTF8 | ConvertFrom-Json -Depth 100

$lines = [System.Collections.Generic.List[string]]::new()
$lines.Add('<section class="imo-benchmark-evidence" aria-labelledby="benchmark-evidence-title">')
$lines.Add('<div class="imo-benchmark-evidence__intro"><p class="imo-benchmark-eyebrow">Measured comparisons</p><h2 id="benchmark-evidence-title">Current Excel and CSV evidence</h2><p>These compact views are generated from committed benchmark artifacts. They compare equivalent, validated work on one recorded machine; they are reproducible evidence, not a promise for every workload or environment.</p></div>')

Add-ComparisonFamily -Lines $lines -Document $excel -Family 'excel' -Title 'Excel report and data pipelines' -Description 'Median timings from rotated local runs with 20 warmups and 9 measured iterations. The scenarios cover feature-rich output, styled IDataReader writes, typed reads, and compact streaming writes.' -MetricLabel 'Median' -EvidenceLabel 'Inspect Excel benchmark evidence' -EvidenceUrl 'https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/benchmarks/readme-current/officeimo.excel.comparison.json'
Add-ComparisonFamily -Lines $lines -Document $csv -Family 'csv' -Title 'CSV read and write pipelines' -Description 'BenchmarkDotNet means for wide CSV workloads. Read lanes traverse every field; write lanes validate every emitted value so faster output cannot hide incomplete work.' -MetricLabel 'Mean' -EvidenceLabel 'Inspect CSV benchmark evidence' -EvidenceUrl 'https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/benchmarks/readme-current/officeimo.csv.comparison.json'

$lines.Add('</section>')

$resolvedOutput = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $OutputPath))
$outputDirectory = Split-Path -Parent $resolvedOutput
if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
    New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
}

$content = ($lines -join "`n") + "`n"
[System.IO.File]::WriteAllText($resolvedOutput, $content, [System.Text.UTF8Encoding]::new($false))
Write-Host "Benchmark website overview written to '$resolvedOutput'."
