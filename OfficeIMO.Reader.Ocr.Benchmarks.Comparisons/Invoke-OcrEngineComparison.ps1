[CmdletBinding()]
param(
    [ValidateRange(0, 20)]
    [int] $WarmupCount = 0,

    [ValidateRange(1, 100)]
    [int] $IterationCount = 3,

    [string] $EnvironmentRoot,

    [string] $OutputRoot,

    [string] $PSPublishModulePath,

    [string[]] $Case,

    [ValidateSet('Tesseract', 'RapidOCR')]
    [string[]] $Engine,

    [switch] $Plan,

    [switch] $RefreshFixtures
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
if ([string]::IsNullOrWhiteSpace($EnvironmentRoot)) {
    $EnvironmentRoot = Join-Path $repositoryRoot 'Ignore\Benchmarks\OcrEngineComparison'
}
if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    $OutputRoot = Join-Path $EnvironmentRoot 'runs'
}

$requiredPaths = @{
    TesseractExecutable = Join-Path $EnvironmentRoot 'tesseract-root\usr\bin\tesseract'
    TesseractRoot = Join-Path $EnvironmentRoot 'tesseract-root'
    RapidPackages = Join-Path $EnvironmentRoot 'rapid-packages'
    RapidModels = Join-Path $EnvironmentRoot 'rapid-models'
}
foreach ($entry in $requiredPaths.GetEnumerator()) {
    if (-not (Test-Path -LiteralPath $entry.Value)) {
        throw "OCR comparison prerequisite '$($entry.Key)' was not found at '$($entry.Value)'. See README.md for isolated setup."
    }
}

$lockPath = Join-Path $PSScriptRoot 'environment.lock.json'
$environmentLock = Get-Content -LiteralPath $lockPath -Raw | ConvertFrom-Json
if ($environmentLock.schemaVersion -ne 1) {
    throw "Unsupported OCR comparison environment lock schema '$($environmentLock.schemaVersion)'."
}

$lockedModels = @($environmentLock.rapidOcr.models.PSObject.Properties)
$actualModels = @(Get-ChildItem -LiteralPath $requiredPaths.RapidModels -File -Filter '*.onnx')
if ($actualModels.Count -ne $lockedModels.Count) {
    throw "RapidOCR model set does not match '$lockPath'. Expected $($lockedModels.Count) ONNX files and found $($actualModels.Count)."
}
foreach ($model in $lockedModels) {
    $modelPath = Join-Path $requiredPaths.RapidModels $model.Name
    if (-not (Test-Path -LiteralPath $modelPath -PathType Leaf)) {
        throw "Locked RapidOCR model '$($model.Name)' is missing from '$($requiredPaths.RapidModels)'."
    }
    $actualHash = (Get-FileHash -LiteralPath $modelPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $expectedHash = ([string] $model.Value).ToLowerInvariant()
    if ($actualHash -ne $expectedHash) {
        throw "RapidOCR model '$($model.Name)' failed the SHA-256 lock check."
    }
}

$fixtures = Join-Path $EnvironmentRoot 'fixtures'
$fixtureManifest = Join-Path $fixtures 'cases.json'
$prepareScript = Join-Path $PSScriptRoot 'tools\prepare_fixtures.py'
$runnerScript = Join-Path $PSScriptRoot 'tools\run_ocr.py'

function ConvertTo-WslPath {
    param([Parameter(Mandatory)][string] $Path)
    $resolved = (Resolve-Path -LiteralPath $Path).Path
    if ($resolved.Length -lt 3 -or $resolved[1] -ne ':') {
        throw "The OCR comparison currently requires a Windows drive path, but received '$resolved'."
    }
    '/mnt/{0}{1}' -f $resolved.Substring(0, 1).ToLowerInvariant(), $resolved.Substring(2).Replace('\', '/')
}

if ($RefreshFixtures.IsPresent -or -not (Test-Path -LiteralPath $fixtureManifest -PathType Leaf)) {
    [void] (New-Item -ItemType Directory -Force -Path $fixtures)
    $pythonPath = ConvertTo-WslPath $requiredPaths.RapidPackages
    $preparePath = ConvertTo-WslPath $prepareScript
    $fixturePath = ConvertTo-WslPath $fixtures
    & wsl.exe -- env "PYTHONPATH=$pythonPath" python3 $preparePath $fixturePath
    if ($LASTEXITCODE -ne 0) {
        throw 'OCR comparison fixture generation failed.'
    }
}

if ([string]::IsNullOrWhiteSpace($PSPublishModulePath)) {
    Import-Module PSPublishModule -MinimumVersion 3.0.128 -Force -ErrorAction Stop
} else {
    Import-Module (Resolve-Path -LiteralPath $PSPublishModulePath).Path -Force -ErrorAction Stop
}

$tesseractRootWsl = ConvertTo-WslPath $requiredPaths.TesseractRoot
$rapidPackagesWsl = ConvertTo-WslPath $requiredPaths.RapidPackages
$rapidModelsWsl = ConvertTo-WslPath $requiredPaths.RapidModels
$runnerWsl = ConvertTo-WslPath $runnerScript
$fixturesWsl = ConvertTo-WslPath $fixtures

$tesseractVersion = (& wsl.exe -- env `
    "LD_LIBRARY_PATH=$tesseractRootWsl/usr/lib/x86_64-linux-gnu" `
    "TESSDATA_PREFIX=$tesseractRootWsl/usr/share/tesseract-ocr/5/tessdata" `
    "$tesseractRootWsl/usr/bin/tesseract" --version | Select-Object -First 1).Trim()
$pythonVersions = & wsl.exe -- env "PYTHONPATH=$rapidPackagesWsl" python3 -c `
    "import importlib.metadata; print(importlib.metadata.version('rapidocr')); print(importlib.metadata.version('onnxruntime'))"
if ($LASTEXITCODE -ne 0 -or @($pythonVersions).Count -ne 2) {
    throw 'Could not resolve RapidOCR and ONNX Runtime versions.'
}
if ($tesseractVersion -notmatch ('^tesseract\s+' + [regex]::Escape([string] $environmentLock.tesseract.version) + '(?:\s|$)')) {
    throw "Tesseract version '$tesseractVersion' does not match the locked version '$($environmentLock.tesseract.version)'."
}
if ([string] $pythonVersions[0] -ne [string] $environmentLock.rapidOcr.version) {
    throw "RapidOCR version '$($pythonVersions[0])' does not match the locked version '$($environmentLock.rapidOcr.version)'."
}
if ([string] $pythonVersions[1] -ne [string] $environmentLock.rapidOcr.onnxRuntimeVersion) {
    throw "ONNX Runtime version '$($pythonVersions[1])' does not match the locked version '$($environmentLock.rapidOcr.onnxRuntimeVersion)'."
}

function Get-DirectoryBytes {
    param([Parameter(Mandatory)][string[]] $Path)
    $bytes = 0L
    foreach ($item in $Path) {
        $bytes += (Get-ChildItem -LiteralPath $item -Recurse -File | Measure-Object Length -Sum).Sum
    }
    $bytes
}

$invoke = @{
    Path = Join-Path $PSScriptRoot 'ocr-engine-comparison.benchmark.ps1'
    OutputRoot = $OutputRoot
    WarmupCount = $WarmupCount
    IterationCount = $IterationCount
    RunMode = 'local-wsl-cpu'
    Variable = @{
        WslExecutable = 'wsl.exe'
        RunnerPath = $runnerWsl
        FixtureRoot = $fixturesWsl
        TesseractRoot = $tesseractRootWsl
        RapidPackages = $rapidPackagesWsl
        RapidModels = $rapidModelsWsl
        ManifestPath = $fixtureManifest
        TesseractVersion = $tesseractVersion
        RapidOcrVersion = $pythonVersions[0]
        OnnxRuntimeVersion = $pythonVersions[1]
        TesseractFootprintBytes = Get-DirectoryBytes $requiredPaths.TesseractRoot
        RapidFootprintBytes = Get-DirectoryBytes @($requiredPaths.RapidPackages, $requiredPaths.RapidModels)
    }
}
if ($Plan.IsPresent) {
    $invoke.Plan = $true
}
if ($Case) {
    $invoke.Case = $Case
}
if ($Engine) {
    $invoke.Engine = $Engine
}

$result = Invoke-BenchmarkSuite @invoke
if (-not $Plan.IsPresent) {
    $failed = @($result.Summary | Where-Object { $_.FailureCount -gt 0 -or $_.Status -eq 'Failed' })
    if ($failed.Count -gt 0) {
        throw "OCR comparison run $($result.RunId) contains failed lanes. Inspect '$($result.Artifacts['run-report.json'])'."
    }
}
$result
