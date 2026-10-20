param(
    [Parameter(Mandatory)][string] $BaselineBundle,
    [Parameter(Mandatory)][string] $OutputRoot,
    [string[]] $Cases,
    [ValidateSet('baseline','current')][string[]] $Engine = @('baseline','current'),
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
& $node $helper prepare --baseline (Resolve-Path -LiteralPath $BaselineBundle).Path --current (Join-Path $repository 'OfficeIMO.JavaScript/bundles/officeimo.mjs') --evidence $evidence @selection
if ($LASTEXITCODE -ne 0) { throw 'Workload preparation failed.' }
if (@((Get-Content -LiteralPath $manifest -Raw | ConvertFrom-Json).cases | Where-Object currentOnly).Count -gt 0 -and ($Compare -or $Engine -contains 'baseline')) {
    throw 'Editable-analysis diagnostics require -Engine current -Compare:$false; the old snapshot report is not equivalent work.'
}
$variables = @{ Node = $node; Helper = $helper; Manifest = $manifest; Evidence = $evidence
    Qpdf = (Get-Command qpdf -ErrorAction Stop).Source; Pdftotext = (Get-Command pdftotext -ErrorAction Stop).Source
    Python = (Get-Command python3 -ErrorAction Stop).Source; Compare = $Compare }
$result = Invoke-BenchmarkSuite -Path (Join-Path $PSScriptRoot 'report-workloads.benchmark.ps1') -OutputRoot $evidence -Variable $variables -Engine $Engine -WarmupCount $WarmupCount -IterationCount $IterationCount -Plan:$Plan
if ($Plan) { $result | Select-Object Scenario,Engine,Operation,IsSkipped; return }
# Shared-runner warmups skip validation; qualify those actual outputs outside timing too.
foreach ($file in Get-ChildItem -LiteralPath $evidence -Filter 'sample.*' -Recurse -File) {
    $iteration = 0
    if (-not [int]::TryParse($file.Directory.Name, [ref] $iteration) -or $iteration -ge 0) { continue }
    $measurement = Get-Content -LiteralPath (Join-Path $file.DirectoryName 'measurement.json') -Raw | ConvertFrom-Json
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
