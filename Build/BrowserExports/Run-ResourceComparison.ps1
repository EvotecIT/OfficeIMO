param(
    [Parameter(Mandatory)][string] $BaselineDist,
    [Parameter(Mandatory)][string] $OutputRoot,
    [string] $ReuseInputManifest,
    [int[]] $Dimensions = @(2048,4096),
    [bool] $IncludeFonts = $true,
    [int] $WarmupCount = 1,
    [int] $IterationCount = 3,
    [switch] $Plan
)
$ErrorActionPreference = 'Stop'
Import-Module PSPublishModule -ErrorAction Stop
$repository = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$helper = Join-Path $PSScriptRoot 'engine-resources.mjs'
$evidence = [IO.Path]::GetFullPath($OutputRoot)
$node = (Get-Command node -ErrorAction Stop).Source
$qpdf = (Get-Command qpdf -ErrorAction Stop).Source
$pdftotext = (Get-Command pdftotext -ErrorAction Stop).Source
$manifest = Join-Path $evidence 'inputs.json'
if (-not $Plan) {
    $reuse = @()
    if ($ReuseInputManifest) { $reuse = @('--reuse', (Resolve-Path -LiteralPath $ReuseInputManifest).Path) }
    & $node $helper prepare --repository $repository --baseline (Resolve-Path -LiteralPath $BaselineDist).Path --current (Join-Path $repository 'OfficeIMO.JavaScript/dist') `
        --evidence $evidence --dimensions ($Dimensions -join ',') --qpdf $qpdf --pdftotext $pdftotext @reuse
    if ($LASTEXITCODE -ne 0) { throw 'Resource fixture preparation failed.' }
    Get-ChildItem -LiteralPath (Join-Path $evidence 'inputs') -Filter '*-expected.pdf' -File | Remove-Item -ErrorAction Stop
}
$result = Invoke-BenchmarkSuite -Path (Join-Path $PSScriptRoot 'resources.benchmark.ps1') -OutputRoot $evidence `
    -Variable @{ Node = $node; Helper = $helper; Manifest = $manifest; Qpdf = $qpdf; Pdftotext = $pdftotext; Evidence = $evidence; Dimensions = $Dimensions; IncludeFonts = $IncludeFonts } `
    -WarmupCount $WarmupCount -IterationCount $IterationCount -Plan:$Plan
if ($Plan) { $result | Select-Object Scenario,Engine,Operation,IsSkipped; return }
# PowerForge excludes warmups from its validation stage. Qualify their retained PDFs outside all timings too.
foreach ($pdf in Get-ChildItem -LiteralPath $evidence -Filter sample.pdf -Recurse -File) {
    $index = 0
    if (-not [int]::TryParse($pdf.Directory.Name, [ref] $index) -or $index -ge 0) { continue }
    $measurement = Get-Content -LiteralPath (Join-Path $pdf.DirectoryName 'measurement.json') -Raw | ConvertFrom-Json
    $proof = & $node $helper validate --manifest $manifest --case $measurement.case --output $pdf.FullName --qpdf $qpdf --pdftotext $pdftotext
    if ($LASTEXITCODE -ne 0) { throw 'Warmup resource output validation failed; PDF retained.' }
    $proof | Set-Content -LiteralPath (Join-Path $pdf.DirectoryName 'warmup-validation.json') -Encoding utf8
    if (($proof | ConvertFrom-Json).sha256 -ne $measurement.outputSha256) { throw 'Warmup validation does not match measured bytes.' }
    Remove-Item -LiteralPath $pdf.FullName -ErrorAction Stop
}
$graph = & $node $helper verify --manifest $manifest
if ($LASTEXITCODE -ne 0) { throw 'Measured PDF graph changed during the cohort.' }
$graph | Set-Content -LiteralPath (Join-Path $evidence 'graph-final.json') -Encoding utf8
$result
if (-not $Plan -and @($result.Summary | Where-Object Status -ne 'Succeeded').Count -gt 0) { throw 'A resource comparison failed; retain and inspect every failed sample.' }
