[CmdletBinding(DefaultParameterSetName = 'SingleBaseline')]
param(
    [Parameter(Mandatory, ParameterSetName = 'SingleBaseline')][string] $BaselineBundle,
    [Parameter(Mandatory, ParameterSetName = 'NamedBaselines')][System.Collections.IDictionary] $BaselineBundles,
    [Parameter(Mandatory)][string] $OutputRoot,
    [string[]] $Cases,
    [string[]] $EquivalentBaselineCases,
    [string[]] $Engine,
    [bool] $Compare = $true,
    [int] $WarmupCount = 1,
    [int] $IterationCount = 3,
    [switch] $Plan
)
$ErrorActionPreference = 'Stop'
Import-Module PSPublishModule -ErrorAction Stop
$repository = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$helper = Join-Path $PSScriptRoot 'report-workloads.mjs'
$node = (Get-Command node -ErrorAction Stop).Source
$evidence = [IO.Path]::GetFullPath($OutputRoot)
$manifest = Join-Path $evidence 'workloads.json'
$selection = @()
if ($Cases) { $selection = @('--cases', ($Cases -join ',')) }
if ($EquivalentBaselineCases) { $selection += @('--equivalent-baseline-cases', ($EquivalentBaselineCases -join ',')) }
$baselines = @()
if ($PSCmdlet.ParameterSetName -eq 'NamedBaselines') {
    if ($BaselineBundles.Count -eq 0) { throw 'BaselineBundles must declare at least one named baseline.' }
    $resolved = [ordered]@{}
    foreach ($name in @($BaselineBundles.Keys | Sort-Object)) {
        if ($name -cnotmatch '^[a-z][a-z0-9-]*$' -or $name -eq 'current') { throw 'Baseline names must be lowercase labels other than current.' }
        if ($BaselineBundles[$name] -isnot [string] -or [string]::IsNullOrWhiteSpace($BaselineBundles[$name])) { throw "Baseline $name must provide a bundle path." }
        $resolved[$name] = (Resolve-Path -LiteralPath $BaselineBundles[$name]).Path
    }
    [IO.Directory]::CreateDirectory($evidence) | Out-Null
    $baselineManifest = Join-Path $evidence 'baseline-bundles.json'
    [IO.File]::WriteAllText($baselineManifest, ($resolved | ConvertTo-Json -Compress), [Text.UTF8Encoding]::new($false))
    $baselines = @('--baselines', $baselineManifest)
} else { $baselines = @('--baseline', (Resolve-Path -LiteralPath $BaselineBundle).Path) }
& $node $helper prepare @baselines --current (Join-Path $repository 'OfficeIMO.JavaScript/bundles/officeimo.mjs') --evidence $evidence @selection
if ($LASTEXITCODE -ne 0) { throw 'Workload preparation failed.' }
$prepared = Get-Content -LiteralPath $manifest -Raw | ConvertFrom-Json
$declaredEngines = @($prepared.engines.PSObject.Properties.Name)
if (-not $Engine) { $Engine = $declaredEngines }
if (@($Engine | Select-Object -Unique).Count -ne $Engine.Count -or @($Engine | Where-Object { $declaredEngines -cnotcontains $_ }).Count -gt 0) {
    throw 'Engine must select unique names declared by the prepared workload manifest.'
}
$comparisonBaselines = @($Engine | Where-Object { $_ -cne 'current' })
if (@($prepared.cases | Where-Object currentOnly).Count -gt 0 -and ($Compare -or $comparisonBaselines.Count -gt 0)) {
    throw 'Analytical-report diagnostics require -Engine current -Compare:$false; the old snapshot report is not equivalent work.'
}
if ($Compare -and ($Engine -cnotcontains 'current' -or $comparisonBaselines.Count -eq 0)) { throw 'Comparison requires current and at least one declared baseline engine.' }
$variables = @{ Node = $node; Helper = $helper; Manifest = $manifest; Evidence = $evidence
    Qpdf = (Get-Command qpdf -ErrorAction Stop).Source; Pdftotext = (Get-Command pdftotext -ErrorAction Stop).Source
    Python = (Get-Command python3 -ErrorAction Stop).Source; Compare = $Compare; ComparisonBaselines = ($comparisonBaselines -join ',') }
$result = Invoke-BenchmarkSuite -Path (Join-Path $PSScriptRoot 'report-workloads.benchmark.ps1') -OutputRoot $evidence -Variable $variables -Engine $Engine -WarmupCount $WarmupCount -IterationCount $IterationCount -Plan:$Plan
if ($Plan) { $result | Select-Object Scenario,Engine,Operation,IsSkipped; return }
# Shared-runner warmups skip validation; qualify those actual outputs outside timing too.
foreach ($file in Get-ChildItem -LiteralPath $evidence -Filter 'sample.*' -Recurse -File) {
    $iteration = 0
    if (-not [int]::TryParse($file.Directory.Name, [ref] $iteration) -or $iteration -ge 0) { continue }
    $measurementPath = Join-Path $file.DirectoryName 'measurement.json'
    # A failed warmup has no measurement: preserve its partial output and the runner's actual error.
    if (-not (Test-Path -LiteralPath $measurementPath -PathType Leaf)) { continue }
    $measurement = Get-Content -LiteralPath $measurementPath -Raw | ConvertFrom-Json
    $proof = & $node $helper validate --manifest $manifest --case $measurement.case --output $file.FullName --qpdf $variables.Qpdf --pdftotext $variables.Pdftotext --python $variables.Python
    if ($LASTEXITCODE -ne 0 -or ($proof | ConvertFrom-Json).sha256 -ne $measurement.outputSha256) { throw 'Warmup output validation failed; report retained.' }
    $proof | Set-Content -LiteralPath (Join-Path $file.DirectoryName 'warmup-validation.json') -Encoding utf8
    Remove-Item -LiteralPath $file.FullName -ErrorAction Stop
}
$runtime = & $node $helper verify --manifest $manifest
if ($LASTEXITCODE -ne 0) { throw 'Measured workload runtime changed during the cohort.' }
$runtime | Set-Content -LiteralPath (Join-Path $evidence 'runtime-final.json') -Encoding utf8
$result
if (@($result.Summary | Where-Object Status -ne 'Succeeded').Count -gt 0) { throw 'Report comparison has failed samples; inspect retained evidence.' }
