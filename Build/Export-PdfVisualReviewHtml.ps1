param(
    [Parameter(Mandatory = $true)]
    [string] $OutputDirectory,

    [Parameter(Mandatory = $true)]
    [string] $ProofSummaryPath
)

$ErrorActionPreference = 'Stop'

function ConvertTo-HtmlText {
    param([AllowNull()][object] $Value)
    [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Format-Metric {
    param(
        [AllowNull()][object] $Value,
        [string] $Format = '0.0000'
    )

    if ($null -eq $Value) {
        return 'n/a'
    }

    ([double]$Value).ToString($Format, [Globalization.CultureInfo]::InvariantCulture)
}

function Add-List {
    param(
        [Parameter(Mandatory = $true)]
        [System.Text.StringBuilder] $Builder,

        [AllowNull()][object[]] $Items,

        [string] $EmptyText
    )

    $values = @($Items | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    if ($values.Count -eq 0) {
        [void]$Builder.Append('<p class="quiet">').Append((ConvertTo-HtmlText $EmptyText)).AppendLine('</p>')
        return
    }

    [void]$Builder.AppendLine('<ul>')
    foreach ($item in $values) {
        [void]$Builder.Append('<li>').Append((ConvertTo-HtmlText $item)).AppendLine('</li>')
    }
    [void]$Builder.AppendLine('</ul>')
}

$resolvedOutput = (Resolve-Path -LiteralPath $OutputDirectory).Path
$resolvedProof = (Resolve-Path -LiteralPath $ProofSummaryPath).Path
$proof = Get-Content -LiteralPath $resolvedProof -Raw | ConvertFrom-Json
$htmlPath = Join-Path $resolvedOutput 'index.html'
$builder = [System.Text.StringBuilder]::new()

[void]$builder.AppendLine('<!doctype html>')
[void]$builder.AppendLine('<html lang="en"><head><meta charset="utf-8">')
[void]$builder.AppendLine('<meta name="viewport" content="width=device-width,initial-scale=1">')
[void]$builder.AppendLine('<title>OfficeIMO PDF visual review</title>')
[void]$builder.AppendLine(@'
<style>
:root{color-scheme:light dark;font-family:Inter,Segoe UI,Arial,sans-serif;background:#0b1020;color:#e8edf7}
*{box-sizing:border-box}body{margin:0;background:#0b1020;color:#e8edf7;line-height:1.5}
main{max-width:1680px;margin:0 auto;padding:28px}.hero,.card{background:#131a2c;border:1px solid #2a3653;border-radius:16px;box-shadow:0 12px 34px #0005}
.hero{padding:28px;margin-bottom:24px;background:linear-gradient(135deg,#17213c,#101729)}
h1,h2,h3,h4{line-height:1.2;margin-top:0}h1{font-size:clamp(2rem,5vw,3.5rem);margin-bottom:8px}
h2{margin-top:38px;font-size:1.7rem}h3{margin-bottom:6px}.quiet{color:#aeb9cf}.status{display:inline-flex;padding:4px 10px;border-radius:999px;font-weight:700}
.pass{background:#143d2b;color:#a9f3c9}.fail{background:#511f2a;color:#ffc2ca}.pending{background:#493a12;color:#ffe49a}
.summary{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:14px;margin-top:22px}
.summary .card{padding:16px}.metric{font-size:1.8rem;font-weight:800}.comparison,.scenario{padding:22px;margin:18px 0}
.page{margin-top:22px;padding-top:18px;border-top:1px solid #2a3653}.visual-grid{display:grid;grid-template-columns:repeat(4,minmax(220px,1fr));gap:12px;overflow:auto}
figure{margin:0;background:#090d18;border:1px solid #26324b;border-radius:12px;padding:10px;min-width:220px}
figure img{display:block;width:100%;height:auto;background:white;border-radius:6px}figcaption{padding-top:8px;color:#c5cee0;font-weight:650}
table{width:100%;border-collapse:collapse;margin:14px 0;font-variant-numeric:tabular-nums}th,td{padding:8px 10px;border-bottom:1px solid #2a3653;text-align:right}
th:first-child,td:first-child{text-align:left}a{color:#8cc8ff}.artifacts{display:flex;flex-wrap:wrap;gap:8px}.artifacts a{padding:5px 9px;border:1px solid #344463;border-radius:8px;text-decoration:none}
details{margin-top:12px}code{color:#b9d8ff}@media(max-width:1050px){.visual-grid{grid-template-columns:repeat(2,minmax(240px,1fr))}}@media(max-width:620px){main{padding:14px}.visual-grid{grid-template-columns:1fr}}
</style>
'@)
[void]$builder.AppendLine('</head><body><main>')
[void]$builder.AppendLine('<section class="hero">')
[void]$builder.AppendLine('<p class="quiet">OfficeIMO managed conversion evidence</p>')
[void]$builder.AppendLine('<h1>PDF visual review gallery</h1>')
[void]$builder.Append('<p>').Append((ConvertTo-HtmlText $proof.qualityContract.goal)).AppendLine('</p>')
[void]$builder.Append('<p class="quiet">Commit <code>').Append((ConvertTo-HtmlText $proof.commit)).Append('</code> · generated ').Append((ConvertTo-HtmlText $proof.generatedAt)).AppendLine('</p>')

$comparisons = @($proof.externalReferenceComparisons)
$passedComparisons = @($comparisons | Where-Object { $_.passed }).Count
$scenarioCount = @($proof.scenarios).Count
$artifactCount = @($proof.scenarios | ForEach-Object { @($_.artifacts).Count } | Measure-Object -Sum).Sum
[void]$builder.AppendLine('<div class="summary">')
foreach ($item in @(
    @{ Label = 'Conversion scenarios'; Value = $scenarioCount },
    @{ Label = 'Scenario artifacts'; Value = $artifactCount },
    @{ Label = 'External references passed'; Value = "$passedComparisons / $($comparisons.Count)" },
    @{ Label = 'Renderer commit'; Value = $proof.commit }
)) {
    [void]$builder.AppendLine('<div class="card">')
    [void]$builder.Append('<div class="metric">').Append((ConvertTo-HtmlText $item.Value)).AppendLine('</div>')
    [void]$builder.Append('<div class="quiet">').Append((ConvertTo-HtmlText $item.Label)).AppendLine('</div></div>')
}
[void]$builder.AppendLine('</div></section>')

[void]$builder.AppendLine('<h2>Microsoft Office reference comparisons</h2>')
[void]$builder.AppendLine('<p class="quiet">Reference and OfficeIMO pages are shown side by side. The overlay uses red for reference-only ink, blue for OfficeIMO-only ink, black for overlap, and white for shared background. Diff images are the heat-map artifacts used by the regression gate.</p>')
if ($comparisons.Count -eq 0) {
    [void]$builder.AppendLine('<p class="quiet">No external raster comparison artifacts were generated for this run.</p>')
}

foreach ($comparison in $comparisons) {
    $statusClass = if ($comparison.passed) { 'pass' } elseif (-not $comparison.rasterizerAvailable) { 'pending' } else { 'fail' }
    $statusText = if ($comparison.passed) { 'passed' } elseif (-not $comparison.rasterizerAvailable) { 'rasterizer not run' } else { 'failed' }
    [void]$builder.AppendLine('<article class="card comparison">')
    [void]$builder.Append('<span class="status ').Append($statusClass).Append('">').Append((ConvertTo-HtmlText $statusText)).AppendLine('</span>')
    [void]$builder.Append('<h3>').Append((ConvertTo-HtmlText $comparison.scenarioId)).AppendLine('</h3>')
    [void]$builder.Append('<p class="quiet">').Append((ConvertTo-HtmlText $comparison.producer)).Append(' ').Append((ConvertTo-HtmlText $comparison.producerVersion)).AppendLine('</p>')

    foreach ($page in @($comparison.pages)) {
        $stem = "external-reference-$($comparison.scenarioId).page$($page.page)"
        [void]$builder.AppendLine('<section class="page">')
        [void]$builder.Append('<h4>Page ').Append((ConvertTo-HtmlText $page.page)).AppendLine('</h4>')
        [void]$builder.AppendLine('<table><thead><tr><th>Metric</th><th>Current</th><th>Pinned</th><th>Delta</th><th>Budget</th></tr></thead><tbody>')
        foreach ($metric in @(
            @{ Label = 'Changed-pixel ratio'; Property = 'differentPixelRatio'; Budget = 'maximumDifferentPixelRatio'; Format = '0.0000' },
            @{ Label = 'Mean absolute error'; Property = 'meanAbsoluteError'; Budget = 'maximumMeanAbsoluteError'; Format = '0.000' },
            @{ Label = 'Root mean square error'; Property = 'rootMeanSquareError'; Budget = 'maximumRootMeanSquareError'; Format = '0.000' },
            @{ Label = 'Luminance MAE'; Property = 'meanLuminanceError'; Budget = 'maximumMeanLuminanceError'; Format = '0.000' }
        )) {
            [void]$builder.Append('<tr><td>').Append((ConvertTo-HtmlText $metric.Label)).Append('</td>')
            [void]$builder.Append('<td>').Append((Format-Metric $page.current.($metric.Property) $metric.Format)).Append('</td>')
            [void]$builder.Append('<td>').Append((Format-Metric $page.pinnedBaseline.($metric.Property) $metric.Format)).Append('</td>')
            [void]$builder.Append('<td>').Append((Format-Metric $page.delta.($metric.Property) $metric.Format)).Append('</td>')
            [void]$builder.Append('<td>').Append((Format-Metric $comparison.thresholds.($metric.Budget) $metric.Format)).AppendLine('</td></tr>')
        }
        [void]$builder.AppendLine('</tbody></table>')
        [void]$builder.AppendLine('<div class="visual-grid">')
        foreach ($panel in @(
            @{ Suffix = 'microsoft-office.png'; Label = 'Microsoft Office reference' },
            @{ Suffix = 'officeimo.png'; Label = 'OfficeIMO candidate' },
            @{ Suffix = 'overlay.png'; Label = 'Alignment overlay' },
            @{ Suffix = 'diff.png'; Label = 'Difference heat map' }
        )) {
            $fileName = "$stem.$($panel.Suffix)"
            [void]$builder.Append('<figure><img loading="lazy" src="').Append((ConvertTo-HtmlText $fileName)).Append('" alt="').Append((ConvertTo-HtmlText "$($comparison.scenarioId) page $($page.page) $($panel.Label)")).Append('"><figcaption>').Append((ConvertTo-HtmlText $panel.Label)).AppendLine('</figcaption></figure>')
        }
        [void]$builder.AppendLine('</div></section>')
    }

    [void]$builder.Append('<p><a href="external-reference-').Append((ConvertTo-HtmlText $comparison.scenarioId)).AppendLine('.comparison.json">Download exact comparison metrics</a></p>')
    [void]$builder.AppendLine('</article>')
}

[void]$builder.AppendLine('<h2>Conversion scenario evidence and warnings</h2>')
foreach ($scenario in @($proof.scenarios)) {
    [void]$builder.AppendLine('<article class="card scenario">')
    [void]$builder.Append('<h3>').Append((ConvertTo-HtmlText $scenario.id)).AppendLine('</h3>')
    [void]$builder.Append('<p class="quiet">').Append((ConvertTo-HtmlText $scenario.converter)).Append(' · ').Append((ConvertTo-HtmlText $scenario.sourceFormat)).Append(' → ').Append((ConvertTo-HtmlText $scenario.targetFormat)).Append(' · ').Append((ConvertTo-HtmlText $scenario.status)).AppendLine('</p>')
    [void]$builder.AppendLine('<h4>Expected warnings</h4>')
    Add-List -Builder $builder -Items @($scenario.expectedWarnings) -EmptyText 'No warnings are accepted for this scenario.'
    [void]$builder.AppendLine('<details><summary>Declared simplifications</summary>')
    Add-List -Builder $builder -Items @($scenario.expectedSimplifications) -EmptyText 'No simplifications are declared.'
    [void]$builder.AppendLine('</details><h4>Artifacts</h4><div class="artifacts">')
    foreach ($artifact in @($scenario.artifacts)) {
        [void]$builder.Append('<a href="').Append((ConvertTo-HtmlText $artifact.file)).Append('">').Append((ConvertTo-HtmlText $artifact.file)).AppendLine('</a>')
    }
    [void]$builder.AppendLine('</div></article>')
}

[void]$builder.AppendLine('<h2>Supporting evidence</h2>')
[void]$builder.AppendLine('<div class="artifacts"><a href="conversion-proof-summary.json">Proof summary JSON</a><a href="conversion-scenarios.json">Scenario manifest</a><a href="reference-corpus.json">External reference metadata</a><a href="pdf-conversion-support-matrix.md">Support matrix</a><a href="index.md">Markdown index</a></div>')
[void]$builder.AppendLine('</main></body></html>')

[System.IO.File]::WriteAllText($htmlPath, $builder.ToString(), [System.Text.UTF8Encoding]::new($false))
Write-Host "HTML visual review gallery written to $htmlPath"
