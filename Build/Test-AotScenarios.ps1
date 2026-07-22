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
    [ordered]@{ id = 'production-libraries'; title = 'Production library coverage: 85 fully rooted plus bounded Google adapter'; project = 'OfficeIMO.All.AotSmoke/OfficeIMO.All.AotSmoke.csproj'; targetFramework = 'net10.0' },
    [ordered]@{ id = 'word'; title = 'Word create, save, and reload'; project = 'OfficeIMO.Word.AotSmoke/OfficeIMO.Word.AotSmoke.csproj' },
    [ordered]@{ id = 'excel'; title = 'Excel typed table create, save, and reload'; project = 'OfficeIMO.Excel.AotSmoke/OfficeIMO.Excel.AotSmoke.csproj' },
    [ordered]@{ id = 'powerpoint'; title = 'PowerPoint chart create, duplicate, save, and reload'; project = 'OfficeIMO.PowerPoint.AotSmoke/OfficeIMO.PowerPoint.AotSmoke.csproj' },
    [ordered]@{ id = 'markdown'; title = 'Markdown fluent composition and rendering'; project = 'OfficeIMO.Markdown.AotSmoke/OfficeIMO.Markdown.AotSmoke.csproj' },
    [ordered]@{ id = 'csv'; title = 'CSV parse and schema inspection'; project = 'OfficeIMO.CSV.AotSmoke/OfficeIMO.CSV.AotSmoke.csproj' },
    [ordered]@{ id = 'reader-csv'; title = 'Reader CSV normalized extraction'; project = 'OfficeIMO.Reader.Csv.AotSmoke/OfficeIMO.Reader.Csv.AotSmoke.csproj' },
    [ordered]@{ id = 'reader-all'; title = 'Reader all-formats registration and representative extraction'; project = 'OfficeIMO.Reader.All.AotSmoke/OfficeIMO.Reader.All.AotSmoke.csproj'; targetFramework = 'net10.0' },
    [ordered]@{ id = 'html-pdf-image'; title = 'HTML to SVG, PNG, and searchable PDF'; project = 'OfficeIMO.Html.AotSmoke/OfficeIMO.Html.AotSmoke.csproj' },
    [ordered]@{ id = 'markup-cli'; title = 'Markup production CLI startup and command discovery'; project = 'OfficeIMO.Markup.Cli/OfficeIMO.Markup.Cli.csproj'; targetFramework = 'net10.0'; aotValidation = $true; runArguments = @('--help') },
    [ordered]@{ id = 'reader-tool'; title = 'Reader production CLI startup and command discovery'; project = 'OfficeIMO.Reader.Tool/OfficeIMO.Reader.Tool.csproj'; targetFramework = 'net10.0'; aotValidation = $true; runArguments = @('--help') }
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
        $publishArguments = [System.Collections.Generic.List[string]]::new()
        @('publish', $projectPath, '--configuration', $Configuration, '--runtime', $RuntimeIdentifier, '--artifacts-path', $sdkArtifactsPath, '--output', $publishPath) |
            ForEach-Object { $publishArguments.Add($_) }
        if ($scenario.targetFramework) {
            $publishArguments.Add('--framework')
            $publishArguments.Add([string] $scenario.targetFramework)
        }
        if ($scenario.aotValidation) {
            $publishArguments.Add('-p:AotValidation=true')
        }

        $publishOutput = @(& dotnet @publishArguments 2>&1)
        $publishExit = $LASTEXITCODE
        $diagnosticCodes = @([regex]::Matches(($publishOutput -join "`n"), '\b(?:IL|NETSDK|CS)\d{4}\b') |
            ForEach-Object Value | Sort-Object -Unique)

        if ($publishExit -eq 0) {
            $projectName = [System.IO.Path]::GetFileNameWithoutExtension($projectPath)
            $executablePath = Join-Path $publishPath ($projectName + $(if ($RuntimeIdentifier -like 'win-*') { '.exe' } else { '' }))
            Write-Host "[$($scenario.id)] execute native binary" -ForegroundColor Cyan
            $runArguments = @($scenario.runArguments)
            $runOutput = @(& $executablePath @runArguments 2>&1)
            $runExit = $LASTEXITCODE
            $status = if ($runExit -eq 0) { 'passed' } else { 'runtime-failed' }
            $runOutput | ForEach-Object { Write-Host $_ }
        } else {
            $status = 'publish-blocked'
            $publishOutput | Where-Object { $_ -match '\b(?:IL|NETSDK|CS)\d{4}\b|:\s+error\s+' } | Select-Object -Unique | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
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

        # NativeAOT SDK state can be several gigabytes for the broad Reader and
        # rendering graphs. Each result is complete once its executable exits,
        # so release that scenario before starting the next one instead of
        # making the final linker compete with every previous publish.
        foreach ($scenarioArtifactPath in @($publishPath, $sdkArtifactsPath)) {
            if ([System.IO.Directory]::Exists($scenarioArtifactPath)) {
                [System.IO.Directory]::Delete($scenarioArtifactPath, $true)
            }
        }
    }

    $results | Format-Table id, status, @{ Name = 'diagnostics'; Expression = { $_.diagnosticCodes -join ', ' } } -AutoSize

    if (-not [string]::IsNullOrWhiteSpace($JsonOutputPath)) {
        $resolvedOutputPath = [System.IO.Path]::GetFullPath($JsonOutputPath)
        New-Item -ItemType Directory -Path (Split-Path -Parent $resolvedOutputPath) -Force | Out-Null
        [ordered]@{
            schemaVersion = 1
            runtimeIdentifier = $RuntimeIdentifier
            configuration = $Configuration
            scenarioCount = $results.Count
            scenarios = @($results)
        } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $resolvedOutputPath -Encoding utf8
        Write-Host "Wrote $resolvedOutputPath" -ForegroundColor Green
    }

    if ($failures.Count -gt 0) {
        throw "NativeAOT validation did not match the documented contract:`n- $($failures -join "`n- ")"
    }
} finally {
    if ([System.IO.Directory]::Exists($artifactRoot)) {
        [System.IO.Directory]::Delete($artifactRoot, $true)
    }
}
