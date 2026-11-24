[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidateSet('Pdf', 'WidgetHtml')][string] $Mode,
    [string] $PythonPath = 'python3',
    [string] $ModulePath = 'PSPublishModule',
    [string] $OutputRoot = (Join-Path $PSScriptRoot '../../../Ignore/Benchmarks/AdvancedNotebook'),
    [string] $OfficeDriver = (Join-Path $PSScriptRoot 'OfficeDriver/bin/Release/net10.0/OfficeNotebookExport.dll'),
    [string] $Chrome,
    [string] $WidgetManagerDist,
    [string] $RequireJs,
    [string] $WidgetValidator = (Join-Path $PSScriptRoot 'WidgetVerifier/bin/Release/net10.0/WidgetVerifier.dll'),
    [ValidateRange(1, 100)][int] $IterationCount = 5,
    [ValidateRange(0, 100)][int] $WarmupCount = 0,
    [switch] $QualifyOnly
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'AdvancedHelpers.ps1')
$python = (Get-Command $PythonPath -CommandType Application | Select-Object -First 1).Source
$version = & $python -c 'import nbconvert; print(nbconvert.__version__)'
if ($LASTEXITCODE -ne 0 -or $version -ne '7.17.2') { throw 'Select the pinned Python environment from requirements.txt.' }
$officeDriver = (Resolve-Path -LiteralPath $OfficeDriver).ProviderPath
$fixtureRoot = Join-Path $PSScriptRoot '../../../OfficeIMO.TestAssets/Notebook/AdvancedOutputs'
$OutputRoot = [IO.Path]::GetFullPath($OutputRoot)
$qualification = Join-Path $OutputRoot ('qualification-' + [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmssfff'))
$derived = Join-Path $qualification 'fixtures'
[void][IO.Directory]::CreateDirectory($derived)

$cases = if ($Mode -eq 'Pdf') {
    if (-not $Chrome) { throw 'Specify -Chrome with the local Chromium executable for Kaleido.' }
    $Chrome = (Resolve-Path -LiteralPath $Chrome).ProviderPath
    Get-Command -Name 'pdftotext', 'pdftoppm' -CommandType Application | Out-Null
    $null = & $python (Join-Path $PSScriptRoot 'BuildPlotlySubset.py') (Join-Path $fixtureRoot 'Charts/plotly-advanced.ipynb') $derived
    if ($LASTEXITCODE -ne 0) { throw 'Could not derive the PDF benchmark fixtures.' }
    foreach ($plot in @(
        @{ Name = 'Scatter3d'; File = 'upstream-helix.ipynb'; Witness = 'Plotly evidence HELIX' },
        @{ Name = 'Surface'; File = 'surface.ipynb'; Witness = 'Plotly evidence SURFACE' },
        @{ Name = 'ScatterGl'; File = 'scattergl.ipynb'; Witness = 'Plotly evidence SCATTERGL' }
    )) {
        @{ Name = $plot.Name; InputPath = (Join-Path $derived $plot.File); Witness = $plot.Witness }
    }
} else {
    if (-not $WidgetManagerDist -or -not $RequireJs) {
        throw 'Specify -WidgetManagerDist and -RequireJs with the local widget runtime assets.'
    }
    $WidgetManagerDist = (Resolve-Path -LiteralPath $WidgetManagerDist).ProviderPath
    $RequireJs = (Resolve-Path -LiteralPath $RequireJs).ProviderPath
    $WidgetValidator = (Resolve-Path -LiteralPath $WidgetValidator).ProviderPath
    $widgetPath = Join-Path $derived 'play-slider.ipynb'
    & $python (Join-Path $PSScriptRoot 'BuildWidgetSubset.py') (Join-Path $fixtureRoot 'Widgets/widget-list.ipynb') $widgetPath
    if ($LASTEXITCODE -ne 0) { throw 'Could not derive the linked widget fixture.' }
    @{ Name = 'LinkedPlayAndSlider'; InputPath = $widgetPath; Witness = '' }
}
$settings = [ordered]@{
    Mode = $Mode
    Cases = @($cases)
    Python = $python
    Chrome = $Chrome
    OfficeDriver = $officeDriver
    NbconvertDriver = Join-Path $PSScriptRoot 'NbconvertAdvance.py'
    PdfValidator = Join-Path $PSScriptRoot 'ValidateChartPdf.py'
    WidgetValidator = $WidgetValidator
    WidgetManagerDist = $WidgetManagerDist
    RequireJs = $RequireJs
}
foreach ($case in $settings.Cases) {
    foreach ($engine in 'Office', 'Nbconvert') {
        $out = New-AdvancedCaseOutput $settings (Join-Path $qualification "$($case.Name)/$engine") $engine
        $null = Invoke-AdvancedFileSave $settings $case $engine $out
        Assert-AdvancedFileArtifact $settings $case $out -Preview
    }
}
$settingsPath = Join-Path $qualification 'settings.json'
$settings | ConvertTo-Json -Depth 5 | Set-Content -Encoding utf8 -LiteralPath $settingsPath
if ($QualifyOnly) { Write-Output $qualification; return }

Import-Module -Name $ModulePath -MinimumVersion 3.0.194
$run = @{
    Path = Join-Path $PSScriptRoot 'advanced-saved-outputs.benchmark.ps1'
    OutputRoot = $OutputRoot
    WarmupCount = $WarmupCount
    IterationCount = $IterationCount
    Variable = @{ SettingsPath = $settingsPath }
}
$result = Invoke-BenchmarkSuite @run
$result
$failed = @($result.Samples | Where-Object Status -ne 'Succeeded')
if ($failed.Count -or -not $result.Samples.Count) {
    $failed | Format-Table Engine, Operation, Scenario, Reason -Wrap | Out-String | Write-Host
    throw 'Advanced notebook comparisons require validated artifacts from both engines.'
}
