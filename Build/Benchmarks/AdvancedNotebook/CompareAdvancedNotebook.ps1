[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidateSet('Pdf', 'WidgetHtml', 'ChartHtml')][string] $Mode,
    [ValidateSet('Original', 'Expanded')][string] $ChartSet = 'Original',
    [string] $PythonPath = 'python3',
    [string] $ModulePath = 'PSPublishModule',
    [string] $OutputRoot = (Join-Path $PSScriptRoot '../../../Ignore/Benchmarks/AdvancedNotebook'),
    [string] $OfficeDriver = (Join-Path $PSScriptRoot 'OfficeDriver/bin/Release/net10.0/OfficeNotebookExport.dll'),
    [string] $Chrome,
    [string] $TopojsonDirectory,
    [string] $WidgetManagerDist,
    [string] $RequireJs,
    [string] $WidgetValidator = (Join-Path $PSScriptRoot 'WidgetVerifier/bin/Release/net10.0/WidgetVerifier.dll'),
    [ValidateRange(1, 100)][int] $IterationCount = 5,
    [ValidateRange(0, 100)][int] $WarmupCount = 0,
    [switch] $QualifyOnly,
    [switch] $Plan
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'AdvancedHelpers.ps1')
$python = (Get-Command $PythonPath -CommandType Application | Select-Object -First 1).Source
$version = & $python -c 'import nbconvert; print(nbconvert.__version__)'
if ($LASTEXITCODE -ne 0 -or $version -ne '7.17.2') { throw 'Select the pinned Python environment from requirements.txt.' }
if ($Plan -and $QualifyOnly) { throw 'Choose -Plan or -QualifyOnly.' }
$officeDriver = (Resolve-Path -LiteralPath $OfficeDriver).ProviderPath
$fixtureRoot = Join-Path $PSScriptRoot '../../../OfficeIMO.TestAssets/Notebook/AdvancedOutputs'
$OutputRoot = [IO.Path]::GetFullPath($OutputRoot)
$qualification = Join-Path $OutputRoot ('qualification-' + [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmssfff'))
$derived = Join-Path $qualification 'fixtures'
[void][IO.Directory]::CreateDirectory($derived)

$plotlyJs = $null
$topojsonSha256 = $null
$cases = if ($Mode -ne 'WidgetHtml') {
    $plots = if ($Mode -eq 'ChartHtml') {
        $ChartSet = 'Expanded'
        @{ Name = 'HeatmapInteraction'; Id = 'station-heatmap'; Witness = 'Plotly evidence HEATMAP' }
    } elseif ($ChartSet -eq 'Expanded') {
        @{ Name = 'Mesh3d'; Id = 'survey-mesh'; Witness = 'Plotly evidence MESH' }
        @{ Name = 'ScatterGeo'; Id = 'survey-route'; Witness = 'Plotly evidence ROUTE' }
        @{ Name = 'Choropleth'; Id = 'survey-regions'; Witness = 'Plotly evidence REGIONS' }
        @{ Name = 'Heatmap'; Id = 'station-heatmap'; Witness = 'Plotly evidence HEATMAP' }
        @{ Name = 'Subplots'; Id = 'station-subplots'; Witness = 'Plotly evidence SUBPLOTS' }
    } else {
        @{ Name = 'Scatter3d'; Id = 'upstream-helix'; Witness = 'Plotly evidence HELIX' }
        @{ Name = 'Surface'; Id = 'surface'; Witness = 'Plotly evidence SURFACE' }
        @{ Name = 'ScatterGl'; Id = 'scattergl'; Witness = 'Plotly evidence SCATTERGL' }
    }
    if ($Mode -eq 'Pdf') {
        if (-not $Chrome) { throw 'Specify -Chrome with the local Chromium executable for Kaleido.' }
        $Chrome = (Resolve-Path -LiteralPath $Chrome).ProviderPath
        Get-Command -Name 'pdftotext', 'pdftoppm', 'pdfimages' -CommandType Application | Out-Null
        if ($ChartSet -eq 'Expanded') {
            if (-not $TopojsonDirectory) { throw 'Specify -TopojsonDirectory with the local world_110m.json oracle asset.' }
            $TopojsonDirectory = (Resolve-Path -LiteralPath $TopojsonDirectory).ProviderPath
            $topojsonSha256 = (Get-FileHash -LiteralPath (Join-Path $TopojsonDirectory 'world_110m.json')).Hash
            if ($topojsonSha256 -ne 'D75915EAA31C870DF6B972C9E5BB86910197825F33DCFEF740F3B2F68CFFE843') {
                throw 'world_110m.json does not match the pinned oracle asset in README.md.'
            }
        }
    } else {
        $WidgetValidator = (Resolve-Path -LiteralPath $WidgetValidator).ProviderPath
        $plotlyJs = & $python -c 'from pathlib import Path; import plotly; assert plotly.__version__ == "6.9.0"; print(Path(plotly.__file__).parent / "package_data" / "plotly.min.js")'
        if ($LASTEXITCODE -ne 0) { throw 'Select the pinned Plotly.py environment from requirements.txt.' }
        $plotlyJs = (Resolve-Path -LiteralPath $plotlyJs).ProviderPath
    }
    $source = if ($ChartSet -eq 'Expanded') { 'Charts/plotly-interactive.ipynb' } else { 'Charts/plotly-advanced.ipynb' }
    $arguments = @((Join-Path $PSScriptRoot 'BuildPlotlySubset.py'), (Join-Path $fixtureRoot $source), $derived)
    foreach ($plot in $plots) { $arguments += @('--case', $plot.Id, $plot.Witness) }
    $null = & $python @arguments
    if ($LASTEXITCODE -ne 0) { throw 'Could not derive the chart benchmark fixtures.' }
    foreach ($plot in $plots) {
        @{ Name = $plot.Name; InputPath = (Join-Path $derived ($plot.Id + '.ipynb')); Witness = $plot.Witness }
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
    TopojsonDirectory = $TopojsonDirectory
    TopojsonSha256 = $topojsonSha256
    PlotlyJs = $plotlyJs
    OfficeDriver = $officeDriver
    NbconvertDriver = Join-Path $PSScriptRoot 'NbconvertAdvance.py'
    PdfValidator = Join-Path $PSScriptRoot 'ValidateChartPdf.py'
    WidgetValidator = $WidgetValidator
    WidgetManagerDist = $WidgetManagerDist
    RequireJs = $RequireJs
}
if (-not $Plan) {
    foreach ($case in $settings.Cases) {
        foreach ($engine in 'Office', 'Nbconvert') {
            $out = New-AdvancedCaseOutput $settings (Join-Path $qualification "$($case.Name)/$engine") $engine
            $null = Invoke-AdvancedFileSave $settings $case $engine $out
            Assert-AdvancedFileArtifact $settings $case $out $engine -Preview
        }
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
if ($Plan) { Invoke-BenchmarkSuite @run -Plan; return }
$result = Invoke-BenchmarkSuite @run
$result
$failed = @($result.Samples | Where-Object Status -ne 'Succeeded')
if ($failed.Count -or -not $result.Samples.Count) {
    $failed | Format-Table Engine, Operation, Scenario, Reason -Wrap | Out-String | Write-Host
    throw 'Advanced notebook comparisons require validated artifacts from both engines.'
}
