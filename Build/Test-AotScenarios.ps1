[CmdletBinding()]
param(
    [string] $RepositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path,
    [string] $RuntimeIdentifier = '',
    [ValidateSet('Debug', 'Release')]
    [string] $Configuration = 'Release',
    [string] $JsonOutputPath = ''
)

$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $false

if ([string]::IsNullOrWhiteSpace($RuntimeIdentifier)) {
    $architecture = switch ([System.Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture) {
        'Arm64' { 'arm64' }
        default { 'x64' }
    }
    $RuntimeIdentifier = if ($IsWindows) {
        "win-$architecture"
    } elseif ($IsMacOS) {
        "osx-$architecture"
    } else {
        "linux-$architecture"
    }
}

$scenarios = @(
    [ordered]@{ id = 'word'; title = 'Word create, save, and reload'; project = 'OfficeIMO.Word.AotSmoke/OfficeIMO.Word.AotSmoke.csproj' },
    [ordered]@{ id = 'excel'; title = 'Excel typed table create, save, and reload'; project = 'OfficeIMO.Excel.AotSmoke/OfficeIMO.Excel.AotSmoke.csproj' },
    [ordered]@{ id = 'powerpoint'; title = 'PowerPoint chart create, duplicate, save, and reload'; project = 'OfficeIMO.PowerPoint.AotSmoke/OfficeIMO.PowerPoint.AotSmoke.csproj' },
    [ordered]@{ id = 'markdown'; title = 'Markdown fluent composition and rendering'; project = 'OfficeIMO.Markdown.AotSmoke/OfficeIMO.Markdown.AotSmoke.csproj' },
    [ordered]@{ id = 'csv'; title = 'CSV parse and schema inspection'; project = 'OfficeIMO.CSV.AotSmoke/OfficeIMO.CSV.AotSmoke.csproj' },
    [ordered]@{ id = 'reader-csv'; title = 'Reader CSV normalized extraction'; project = 'OfficeIMO.Reader.Csv.AotSmoke/OfficeIMO.Reader.Csv.AotSmoke.csproj' },
    [ordered]@{ id = 'reader-all'; title = 'Reader all-formats registration and representative extraction'; project = 'OfficeIMO.Reader.All.AotSmoke/OfficeIMO.Reader.All.AotSmoke.csproj' },
    [ordered]@{ id = 'html-pdf-image'; title = 'HTML to SVG, PNG, and searchable PDF'; project = 'OfficeIMO.Html.AotSmoke/OfficeIMO.Html.AotSmoke.csproj' }
)

$artifactRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("OfficeIMO-AotValidation-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $artifactRoot -Force | Out-Null
$results = [System.Collections.Generic.List[object]]::new()
$failures = [System.Collections.Generic.List[string]]::new()

try {
    foreach ($scenario in $scenarios) {
        $projectPath = Join-Path $RepositoryRoot $scenario.project
        $publishPath = Join-Path $artifactRoot $scenario.id
        $sdkArtifactsPath = Join-Path $artifactRoot ($scenario.id + '-sdk')

        Write-Host "[$($scenario.id)] restore and publish NativeAOT in isolated SDK state" -ForegroundColor Cyan
        $publishOutput = @(& dotnet publish $projectPath --configuration $Configuration --runtime $RuntimeIdentifier --artifacts-path $sdkArtifactsPath --output $publishPath 2>&1)
        $publishExit = $LASTEXITCODE
        $diagnosticCodes = @([regex]::Matches(($publishOutput -join "`n"), '\bIL\d{4}\b') |
            ForEach-Object Value | Sort-Object -Unique)

        if ($publishExit -eq 0) {
            $projectName = [System.IO.Path]::GetFileNameWithoutExtension($projectPath)
            $executablePath = Join-Path $publishPath ($projectName + $(if ($RuntimeIdentifier -like 'win-*') { '.exe' } else { '' }))
            Write-Host "[$($scenario.id)] execute native binary" -ForegroundColor Cyan
            $runOutput = @(& $executablePath 2>&1)
            $runExit = $LASTEXITCODE
            $status = if ($runExit -eq 0) { 'passed' } else { 'runtime-failed' }
            $runOutput | ForEach-Object { Write-Host $_ }
        } else {
            $status = 'publish-blocked'
            $publishOutput | Where-Object { $_ -match '\bIL\d{4}\b' } | Select-Object -Unique | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
        }

        if ($status -ne 'passed') {
            $failures.Add("$($scenario.id): expected a passing native publish and execution, got $status")
        }

        $results.Add([pscustomobject] [ordered]@{
            id = $scenario.id
            title = $scenario.title
            status = $status
            diagnosticCodes = $diagnosticCodes
        })
    }

    $results | Format-Table id, status, @{ Name = 'diagnostics'; Expression = { $_.diagnosticCodes -join ', ' } } -AutoSize

    if (-not [string]::IsNullOrWhiteSpace($JsonOutputPath)) {
        $resolvedOutputPath = [System.IO.Path]::GetFullPath($JsonOutputPath)
        New-Item -ItemType Directory -Path (Split-Path -Parent $resolvedOutputPath) -Force | Out-Null
        [ordered]@{
            schemaVersion = 1
            runtimeIdentifier = $RuntimeIdentifier
            configuration = $Configuration
            scenarios = @($results)
        } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $resolvedOutputPath -Encoding utf8
        Write-Host "Wrote $resolvedOutputPath" -ForegroundColor Green
    }

    if ($failures.Count -gt 0) {
        throw "NativeAOT validation did not match the documented contract:`n- $($failures -join "`n- ")"
    }
} finally {
    if (Test-Path -LiteralPath $artifactRoot) {
        Remove-Item -LiteralPath $artifactRoot -Recurse -Force
    }
}
