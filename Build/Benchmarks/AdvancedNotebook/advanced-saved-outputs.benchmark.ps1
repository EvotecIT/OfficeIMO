$settings = [IO.File]::ReadAllText((Get-BenchmarkInput SettingsPath -Required)) | ConvertFrom-Json
. (Join-Path $PSScriptRoot 'AdvancedHelpers.ps1')
New-BenchmarkSuite 'notebook-advanced-saved-outputs' {
    Set-BenchmarkPolicy -Warmup 1 -Iteration 5 -Order Rotated -OutlierMode None
    Add-BenchmarkMetadata Boundary 'Fresh process startup, library load, saved notebook read, conversion, file write and exit. Plotly WebPDF includes Kaleido snapshots of saved vendor MIME figures. Browser interaction and PDF raster/content validation are outside the timer.'
    Add-BenchmarkMetadata OfficeDriverSha256 (Get-FileHash -LiteralPath $settings.OfficeDriver).Hash
    $assemblyDirectory = Split-Path -Parent $settings.OfficeDriver
    foreach ($assembly in 'Core', 'Notebook.Markdown', 'Notebook.Pdf', 'Notebook.Html') {
        Add-BenchmarkMetadata ("OfficeIMO.$assembly.Sha256") (Get-FileHash -LiteralPath (Join-Path $assemblyDirectory "OfficeIMO.$assembly.dll")).Hash
    }
    Add-BenchmarkMetadata Nbconvert '7.17.2'
    if ($settings.Mode -ne 'WidgetHtml') { Add-BenchmarkMetadata Plotly '6.9.0 (Plotly.js 3.7.0)' }
    if ($settings.Mode -eq 'Pdf') { Add-BenchmarkMetadata Kaleido '1.5.2' }
    if ($settings.TopojsonSha256) { Add-BenchmarkMetadata OracleTopojsonSha256 $settings.TopojsonSha256 }
    Add-BenchmarkCaseSource {
        foreach ($item in $settings.Cases) {
            [pscustomobject]@{
                Name = $item.Name
                InputPath = $item.InputPath
                InputSha256 = (Get-FileHash -LiteralPath $item.InputPath).Hash
                Witness = $item.Witness
            }
        }
    }
    Set-BenchmarkSetup {
        param($case, $run)
        $run.OutputPath = New-AdvancedCaseOutput $settings $run.OutputDirectory $case.Engine
    }
    foreach ($engine in 'Office', 'Nbconvert') {
        Add-BenchmarkEngine $engine {
            Add-BenchmarkOperation $settings.Mode {
                param($case, $run)
                $null = Invoke-AdvancedFileSave $settings $case $case.Engine $run.OutputPath
            }
        }
    }
    Add-BenchmarkValidation {
        param($case, $run)
        Assert-AdvancedFileArtifact $settings $case $run.OutputPath $case.Engine
    }
    Add-BenchmarkMetric ArtifactBytes { param($case, $run) (Get-Item -LiteralPath $run.OutputPath).Length }
    Add-BenchmarkComparison Engine -Baseline Nbconvert -Metric MedianMs
    Set-BenchmarkArtifacts Json, Csv, Markdown
}
