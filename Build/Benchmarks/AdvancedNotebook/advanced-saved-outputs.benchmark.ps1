$settings = [IO.File]::ReadAllText((Get-BenchmarkInput SettingsPath -Required)) | ConvertFrom-Json
. (Join-Path $PSScriptRoot 'AdvancedHelpers.ps1')
New-BenchmarkSuite 'notebook-advanced-saved-outputs' {
    Set-BenchmarkPolicy -Warmup 1 -Iteration 5 -Order Rotated -OutlierMode None
    Add-BenchmarkMetadata Boundary 'Fresh process startup, module/library load, saved notebook read, conversion, file write and exit. Plotly WebPDF includes Kaleido snapshot of the saved vendor MIME figures. Browser widget validation and PDF raster/content validation are outside the timer.'
    Add-BenchmarkMetadata OfficeDriverSha256 (Get-FileHash -LiteralPath $settings.OfficeDriver).Hash
    Add-BenchmarkMetadata Nbconvert '7.17.2'
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
        Assert-AdvancedFileArtifact $settings $case $run.OutputPath
    }
    Add-BenchmarkMetric ArtifactBytes { param($case, $run) (Get-Item -LiteralPath $run.OutputPath).Length }
    Set-BenchmarkArtifacts Json, Csv, Markdown
}
