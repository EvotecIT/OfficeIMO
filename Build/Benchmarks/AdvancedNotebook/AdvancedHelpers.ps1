function Invoke-AdvancedFileSave {
    param($Settings, $Case, [string] $Engine, [string] $OutputPath)
    if ($Engine -eq 'Office') {
        $format = switch ($Settings.Mode) {
            'Pdf' { 'pdf' }
            'WidgetHtml' { 'widget-html' }
            'ChartHtml' { 'chart-html' }
        }
        & dotnet $Settings.OfficeDriver $format $Case.InputPath $OutputPath
    } else {
        $arguments = @($Settings.NbconvertDriver, $Settings.Mode, $Case.InputPath, $OutputPath)
        if ($Settings.Mode -eq 'Pdf') {
            $arguments += @('--chrome', $Settings.Chrome)
            if ($Settings.TopojsonDirectory) { $arguments += @('--topojson', $Settings.TopojsonDirectory) }
        }
        & $Settings.Python @arguments
    }
    if ($LASTEXITCODE -ne 0) { throw "$Engine save failed with exit code $LASTEXITCODE." }
}
function New-AdvancedCaseOutput {
    param($Settings, [string] $Directory, [string] $Engine)
    [void][IO.Directory]::CreateDirectory($Directory)
    if ($Settings.Mode -eq 'WidgetHtml' -and $Engine -eq 'Nbconvert') {
        $assets = Join-Path $Directory 'assets'
        [void][IO.Directory]::CreateDirectory($assets)
        Copy-Item -LiteralPath $Settings.RequireJs -Destination (Join-Path $assets 'require.js') -Force
        Copy-Item -LiteralPath (Join-Path $Settings.WidgetManagerDist 'embed-amd.js') -Destination $assets -Force
        foreach ($relativeFontDirectory in '@jupyter-widgets', '@jupyter-widgets/html-manager/dist') {
            $fontDirectory = Join-Path $Directory $relativeFontDirectory
            [void][IO.Directory]::CreateDirectory($fontDirectory)
            foreach ($name in '1551f4f60c37af51121f.woff2', 'eeccf4f66002c6f2ba24.woff', 'be9ee23c0c6390141475.ttf') {
                Copy-Item -LiteralPath (Join-Path $Settings.WidgetManagerDist $name) -Destination (Join-Path $fontDirectory $name) -Force
            }
        }
    }
    if ($Settings.Mode -eq 'ChartHtml' -and $Engine -eq 'Nbconvert') {
        $assets = Join-Path $Directory 'assets'
        [void][IO.Directory]::CreateDirectory($assets)
        Copy-Item -LiteralPath $Settings.PlotlyJs -Destination (Join-Path $assets 'plotly.min.js') -Force
    }
    $extension = if ($Settings.Mode -eq 'Pdf') { '.pdf' } else { '.html' }
    $path = Join-Path $Directory ("saved" + $extension)
    [IO.File]::Delete($path)
    $path
}
function Assert-AdvancedFileArtifact {
    param($Settings, $Case, [string] $Path, [string] $Engine, [switch] $Preview)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Expected conversion artifact '$Path'." }
    $previewPath = [IO.Path]::ChangeExtension($Path, '.png')
    if ($Settings.Mode -eq 'Pdf') {
        $arguments = @($Settings.PdfValidator, $Path, $Case.Witness)
        if ($Engine -eq 'Office') { $arguments += '--require-vectors' }
        if ($Preview) { $arguments += @('--preview', $previewPath) }
        $detail = & $Settings.Python @arguments
    } else {
        $arguments = @($Settings.WidgetValidator)
        if ($Settings.Mode -eq 'ChartHtml') { $arguments += @('--chart', $Path, $Engine) }
        else { $arguments += $Path }
        if ($Preview) { $arguments += $previewPath }
        $detail = & dotnet @arguments
    }
    if ($LASTEXITCODE -ne 0) { throw "Artifact validation failed for '$Path'. $detail" }
    if ($Preview) { $detail | Write-Host }
}
