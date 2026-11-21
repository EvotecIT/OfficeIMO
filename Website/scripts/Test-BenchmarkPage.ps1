param(
    [string] $SiteRoot = (Join-Path $PSScriptRoot '..\_site')
)

$ErrorActionPreference = 'Stop'

$resolvedSiteRoot = if ([System.IO.Path]::IsPathRooted($SiteRoot)) {
    [System.IO.Path]::GetFullPath($SiteRoot)
} else {
    [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $SiteRoot))
}
$declarationPath = Join-Path $PSScriptRoot '..\..\Build\benchmark-declarations.json'
$areaCatalogPath = Join-Path $PSScriptRoot '..\data\benchmark_areas.json'
$catalogPath = Join-Path $resolvedSiteRoot 'data\benchmarks\evidence\index.json'
$indexPath = Join-Path $resolvedSiteRoot 'data\benchmarks\index.json'
$hubPath = Join-Path $resolvedSiteRoot 'benchmarks\index.html'

foreach ($path in @($catalogPath, $indexPath, $hubPath)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "The benchmark site output is missing '$path'."
    }
}
if (Test-Path -LiteralPath (Join-Path $resolvedSiteRoot 'benchmarks\archive\index.html') -PathType Leaf) {
    throw 'The retired benchmark archive page is still generated.'
}

# The published catalog carries exactly the declared comparisons and the standard publish policy.
$declarations = Get-Content -LiteralPath $declarationPath -Raw -Encoding UTF8 | ConvertFrom-Json
$catalog = Get-Content -LiteralPath $catalogPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ($catalog.schemaVersion -ne 5) {
    throw "The published benchmark catalog uses schema $($catalog.schemaVersion); only schema 5 is published."
}
$declaredIds = @($declarations.comparisons.comparisonId | Sort-Object)
$catalogIds = @($catalog.comparisons.comparisonId | Sort-Object)
if (($declaredIds -join ',') -cne ($catalogIds -join ',')) {
    throw 'The published catalog does not declare the same comparisons as Build/benchmark-declarations.json; merge the catalog again.'
}
if ($catalog.policy.minimumWarmups -lt 3 -or $catalog.policy.minimumMeasuredSamples -lt 10 -or
    $catalog.policy.confidenceLevel -lt 0.95 -or -not $catalog.policy.requireOutputValidation -or
    @($catalog.policy.publishableRestrictions | Where-Object { $_ -ne 'open' }).Count -gt 0) {
    throw 'The published catalog does not carry the standard publish policy.'
}

# Every recorded result is published beside the catalog with its recorded hash.
foreach ($entry in @($catalog.entries)) {
    $published = Join-Path (Split-Path -Parent $catalogPath) ([string] $entry.artifactRelativePath)
    if (-not (Test-Path -LiteralPath $published -PathType Leaf)) {
        throw "Benchmark result '$($entry.resultPath)' of lane '$($entry.laneKey)' was not published."
    }
    $hash = (Get-FileHash -LiteralPath $published -Algorithm SHA256).Hash
    if (-not [string]::Equals($hash, [string] $entry.resultSha256, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Benchmark result '$($entry.resultPath)' does not match its catalog SHA-256."
    }
}

# The hub lists every area with its comparisons and links to the area's page and on-demand results.
$areaTitles = @{}
foreach ($area in (Get-Content -LiteralPath $areaCatalogPath -Raw -Encoding UTF8 | ConvertFrom-Json).areas) {
    $areaTitles[$area.id] = $area.title
}
$index = Get-Content -LiteralPath $indexPath -Raw -Encoding UTF8 | ConvertFrom-Json
$hubHtml = Get-Content -LiteralPath $hubPath -Raw -Encoding UTF8
$fragmentRoot = @(Get-ChildItem -LiteralPath (Join-Path $resolvedSiteRoot '_powerforge\fragments\benchmarks') -Directory)
if ($fragmentRoot.Count -ne 1) {
    throw 'The benchmark area results were not written as page fragments.'
}
if ($hubHtml -notmatch 'assets/powerforge/benchmarks\.v1\.js' -or $hubHtml -notmatch 'assets/powerforge/benchmarks\.v1\.css') {
    throw 'The benchmark hub does not load the benchmark component assets.'
}
$comparisonCount = 0
foreach ($area in @($index.areas)) {
    if (-not $areaTitles.ContainsKey([string] $area.id)) {
        throw "Benchmark area '$($area.id)' has no title in data/benchmark_areas.json."
    }
    if ($hubHtml -notmatch [regex]::Escape("id=""benchmarks-$($area.id)""") -or
        $hubHtml -notmatch [regex]::Escape("href=""/benchmarks/$($area.id)/""")) {
        throw "The benchmark hub does not list the '$($area.id)' area with a link to its page."
    }
    $areaFile = Join-Path $resolvedSiteRoot "data\benchmarks\$($area.file)"
    $fragment = Join-Path $fragmentRoot[0].FullName ([System.IO.Path]::ChangeExtension([string] $area.file, '.html'))
    $areaPage = Join-Path $resolvedSiteRoot "benchmarks\$($area.id)\index.html"
    foreach ($path in @($areaFile, $fragment, $areaPage)) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            throw "Benchmark area '$($area.id)' is missing '$path'."
        }
    }

    # Each area page renders every comparison of the area; one without a published run says so.
    $areaHtml = Get-Content -LiteralPath $areaPage -Raw -Encoding UTF8
    $areaData = Get-Content -LiteralPath $areaFile -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($comparison in @($areaData.comparisons)) {
        $comparisonCount++
        $match = [regex]::Match($areaHtml, '<section class="pf-bench" id="benchmark-' + [regex]::Escape([string] $comparison.id) + '"(?s:.*?)</section>')
        if (-not $match.Success) {
            throw "Benchmark area page '/benchmarks/$($area.id)/' does not render comparison '$($comparison.id)'."
        }
        $publishedRuns = @($comparison.runs | Where-Object published)
        if ($publishedRuns.Count -eq 0 -and $match.Value -notmatch 'No published run yet\.') {
            throw "Comparison '$($comparison.id)' has no published run but its section does not say so."
        }
        foreach ($run in $publishedRuns) {
            if ($match.Value -notmatch [regex]::Escape("data-pf-bench-run=""$($run.id)""")) {
                throw "Comparison '$($comparison.id)' does not render its published run '$($run.id)'."
            }
        }
        foreach ($run in @($comparison.runs | Where-Object { -not $_.published })) {
            if ($match.Value -match [regex]::Escape("data-pf-bench-run=""$($run.id)""")) {
                throw "Comparison '$($comparison.id)' renders the unpublished run '$($run.id)'."
            }
        }
    }
}
if ($comparisonCount -ne $declaredIds.Count) {
    throw "The area pages render $comparisonCount comparisons; $($declaredIds.Count) are declared."
}

& (Join-Path $PSScriptRoot 'Test-BenchmarkPage.Browser.ps1') -SiteRoot $resolvedSiteRoot

Write-Host "Benchmark pages verified: $comparisonCount declared comparisons across $(@($index.areas).Count) areas, $(@($catalog.entries | Where-Object publish).Count) published lanes."
