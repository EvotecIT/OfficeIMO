[CmdletBinding()]
param(
    [string] $RepositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path,
    [string] $RuntimeIdentifier = '',
    [ValidateSet('Debug', 'Release')]
    [string] $Configuration = 'Release',
    [switch] $IncludeKnownBlocked,
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
    [ordered]@{ id = 'word'; title = 'Word create, save, and reload'; project = 'OfficeIMO.Word.AotSmoke/OfficeIMO.Word.AotSmoke.csproj'; expected = 'pass'; diagnostics = @() },
    [ordered]@{ id = 'markdown'; title = 'Markdown fluent composition and rendering'; project = 'OfficeIMO.Markdown.AotSmoke/OfficeIMO.Markdown.AotSmoke.csproj'; expected = 'pass'; diagnostics = @() },
    [ordered]@{ id = 'csv'; title = 'CSV parse and schema inspection'; project = 'OfficeIMO.CSV.AotSmoke/OfficeIMO.CSV.AotSmoke.csproj'; expected = 'pass'; diagnostics = @() },
    [ordered]@{ id = 'reader-csv'; title = 'Reader CSV normalized extraction'; project = 'OfficeIMO.Reader.Csv.AotSmoke/OfficeIMO.Reader.Csv.AotSmoke.csproj'; expected = 'pass'; diagnostics = @() },
    [ordered]@{ id = 'html-pdf-image'; title = 'HTML to SVG, PNG, and searchable PDF'; project = 'OfficeIMO.Html.AotSmoke/OfficeIMO.Html.AotSmoke.csproj'; expected = 'pass'; diagnostics = @() },
    [ordered]@{ id = 'excel'; title = 'Excel create, save, and reload'; project = 'OfficeIMO.Excel.AotSmoke/OfficeIMO.Excel.AotSmoke.csproj'; expected = 'blocked'; diagnostics = @('IL2072') },
    [ordered]@{ id = 'powerpoint'; title = 'PowerPoint create, save, and reload'; project = 'OfficeIMO.PowerPoint.AotSmoke/OfficeIMO.PowerPoint.AotSmoke.csproj'; expected = 'blocked'; diagnostics = @('IL2060', 'IL2075', 'IL2087', 'IL3050') }
)

if (-not $IncludeKnownBlocked) {
    $scenarios = @($scenarios | Where-Object expected -EQ 'pass')
}

$artifactRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("OfficeIMO-AotValidation-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $artifactRoot -Force | Out-Null
$results = [System.Collections.Generic.List[object]]::new()
$failures = [System.Collections.Generic.List[string]]::new()

try {
    foreach ($scenario in $scenarios) {
        $projectPath = Join-Path $RepositoryRoot $scenario.project
        $publishPath = Join-Path $artifactRoot $scenario.id
        Write-Host "[$($scenario.id)] clean prior native state" -ForegroundColor Cyan
        $cleanOutput = @(& dotnet clean $projectPath --configuration $Configuration --runtime $RuntimeIdentifier --verbosity quiet 2>&1)
        $cleanExit = $LASTEXITCODE
        if ($cleanExit -ne 0) {
            $failures.Add("$($scenario.id): clean failed")
            $results.Add([pscustomobject] [ordered]@{ id = $scenario.id; title = $scenario.title; expected = $scenario.expected; status = 'clean-failed'; diagnosticCodes = @() })
            continue
        }

        Write-Host "[$($scenario.id)] restore $RuntimeIdentifier" -ForegroundColor Cyan
        $restoreOutput = @(& dotnet restore $projectPath --runtime $RuntimeIdentifier 2>&1)
        $restoreExit = $LASTEXITCODE
        if ($restoreExit -ne 0) {
            $failures.Add("$($scenario.id): restore failed")
            $results.Add([pscustomobject] [ordered]@{ id = $scenario.id; title = $scenario.title; expected = $scenario.expected; status = 'restore-failed'; diagnosticCodes = @() })
            continue
        }

        Write-Host "[$($scenario.id)] publish NativeAOT" -ForegroundColor Cyan
        $publishOutput = @(& dotnet publish $projectPath --configuration $Configuration --runtime $RuntimeIdentifier --no-restore --output $publishPath 2>&1)
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

        if ($scenario.expected -eq 'pass' -and $status -ne 'passed') {
            $failures.Add("$($scenario.id): expected a passing publish and execution, got $status")
        }
        if ($scenario.expected -eq 'blocked') {
            $missingDiagnostics = @($scenario.diagnostics | Where-Object { $_ -notin $diagnosticCodes })
            if ($status -ne 'publish-blocked') {
                $failures.Add("$($scenario.id): the documented blocker changed; update the compatibility matrix")
            } elseif ($missingDiagnostics.Count -gt 0) {
                $failures.Add("$($scenario.id): expected diagnostics were not reproduced: $($missingDiagnostics -join ', ')")
            }
        }

        $results.Add([pscustomobject] [ordered]@{
            id = $scenario.id
            title = $scenario.title
            expected = $scenario.expected
            status = $status
            diagnosticCodes = $diagnosticCodes
        })
    }

    $results | Format-Table id, expected, status, @{ Name = 'diagnostics'; Expression = { $_.diagnosticCodes -join ', ' } } -AutoSize

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
