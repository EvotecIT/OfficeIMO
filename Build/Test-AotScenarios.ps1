[CmdletBinding()]
param(
    [string] $RepositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path,
    [string] $RuntimeIdentifier = '',
    [ValidateSet('Debug', 'Release')]
    [string] $Configuration = 'Release',
    [string] $JsonOutputPath = '',
    [string[]] $Scenario = @(),
    [ValidateSet('all', 'libraries', 'iwork', 'surfaces')]
    [string] $ScenarioGroup = 'all'
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
    [ordered]@{ id = 'production-libraries'; group = 'libraries'; title = 'Ordinary production library coverage: fully rooted libraries plus bounded Google adapter'; project = 'OfficeIMO.All.AotSmoke/OfficeIMO.All.AotSmoke.csproj'; targetFramework = 'net10.0' },
    [ordered]@{ id = 'word'; group = 'libraries'; title = 'Word dictionary template binding, save, and reload'; project = 'OfficeIMO.Word.AotSmoke/OfficeIMO.Word.AotSmoke.csproj' },
    [ordered]@{ id = 'excel'; group = 'libraries'; title = 'Excel DataReader, typed, and generic-only dictionary table create, save, and reload'; project = 'OfficeIMO.Excel.AotSmoke/OfficeIMO.Excel.AotSmoke.csproj' },
    [ordered]@{ id = 'powerpoint'; group = 'libraries'; title = 'PowerPoint chart and generic-only dictionary table create, duplicate, save, and reload'; project = 'OfficeIMO.PowerPoint.AotSmoke/OfficeIMO.PowerPoint.AotSmoke.csproj' },
    [ordered]@{ id = 'security'; group = 'libraries'; title = 'Optional security provider CMS and XML DSig signing and verification'; project = 'OfficeIMO.Security.AotSmoke/OfficeIMO.Security.AotSmoke.csproj' },
    [ordered]@{ id = 'provenance-c2pa'; group = 'libraries'; title = 'Optional C2PA adapter provider-unavailable process boundary'; project = 'OfficeIMO.Provenance.C2pa.AotSmoke/OfficeIMO.Provenance.C2pa.AotSmoke.csproj' },
    [ordered]@{ id = 'markdown'; group = 'libraries'; title = 'Markdown fluent composition and rendering'; project = 'OfficeIMO.Markdown.AotSmoke/OfficeIMO.Markdown.AotSmoke.csproj' },
    [ordered]@{ id = 'csv'; group = 'libraries'; title = 'CSV parse and schema inspection'; project = 'OfficeIMO.CSV.AotSmoke/OfficeIMO.CSV.AotSmoke.csproj' },
    [ordered]@{ id = 'reader-csv'; group = 'libraries'; title = 'Reader CSV normalized extraction'; project = 'OfficeIMO.Reader.Csv.AotSmoke/OfficeIMO.Reader.Csv.AotSmoke.csproj' },
    [ordered]@{ id = 'reader-iwork-net8'; group = 'iwork'; title = 'iWork source and Reader native-fixture extraction on .NET 8'; project = 'OfficeIMO.Reader.IWork.AotSmoke/OfficeIMO.Reader.IWork.AotSmoke.csproj'; targetFramework = 'net8.0' },
    [ordered]@{ id = 'reader-iwork'; group = 'iwork'; title = 'iWork source and Reader native-fixture extraction, formulas, diagnostics, JSON and limits'; project = 'OfficeIMO.Reader.IWork.AotSmoke/OfficeIMO.Reader.IWork.AotSmoke.csproj'; targetFramework = 'net10.0' },
    [ordered]@{ id = 'iwork-conversion-net8'; group = 'iwork'; title = 'iWork editable conversion and native Keynote creation on .NET 8'; project = 'OfficeIMO.IWork.AotSmoke/OfficeIMO.IWork.AotSmoke.csproj'; targetFramework = 'net8.0' },
    [ordered]@{ id = 'iwork-conversion'; group = 'iwork'; title = 'iWork editable conversion and native Keynote creation'; project = 'OfficeIMO.IWork.AotSmoke/OfficeIMO.IWork.AotSmoke.csproj'; targetFramework = 'net10.0' },
    [ordered]@{ id = 'iwork-pdf-net8'; group = 'iwork'; title = 'Composed iWork PDF workflows, portable fonts and atomic rejection on .NET 8'; project = 'OfficeIMO.Workflows.IWork.AotSmoke/OfficeIMO.Workflows.IWork.AotSmoke.csproj'; targetFramework = 'net8.0' },
    [ordered]@{ id = 'iwork-pdf'; group = 'iwork'; title = 'Composed iWork PDF workflows, portable fonts and atomic rejection'; project = 'OfficeIMO.Workflows.IWork.AotSmoke/OfficeIMO.Workflows.IWork.AotSmoke.csproj'; targetFramework = 'net10.0' },
    [ordered]@{ id = 'reader-all'; group = 'surfaces'; title = 'Reader all-formats registration and representative extraction'; project = 'OfficeIMO.Reader.All.AotSmoke/OfficeIMO.Reader.All.AotSmoke.csproj'; targetFramework = 'net10.0' },
    [ordered]@{ id = 'html-pdf-image'; group = 'surfaces'; title = 'HTML to SVG, PNG, and searchable PDF'; project = 'OfficeIMO.Html.AotSmoke/OfficeIMO.Html.AotSmoke.csproj' },
    [ordered]@{ id = 'officeimo-tool'; group = 'surfaces'; title = 'Unified production CLI startup and command discovery'; project = 'OfficeIMO.Tool/OfficeIMO.Tool.csproj'; targetFramework = 'net10.0'; aotValidation = $true; runArguments = @('--help') },
    [ordered]@{ id = 'studio-engine-library'; group = 'surfaces'; title = 'Studio engine shared library: open, render, rotate, undo, redo and save through its C interface'; project = 'OfficeIMO.Studio.Native/OfficeIMO.Studio.Native.csproj'; targetFramework = 'net10.0'; nativeLibrary = 'OfficeIMOStudioEngine'; journey = 'OfficeIMO.Studio.Native/Test-NativeEngine.ps1'; fixtures = @(
        'OfficeIMO.Pdf.Tests/Pdf/Fixtures/Interoperability/openpreserve-high-resolution-image.pdf',
        'OfficeIMO.TestAssets/MultilingualLayout/latin-0-native.pdf',
        'OfficeIMO.Pdf.Tests/Pdf/Fixtures/Interoperability/openpreserve-pdfa1b-text.pdf'
    ) }
)

if (@($scenarios | Where-Object { $_.group -notin @('libraries', 'iwork', 'surfaces') }).Count -gt 0) {
    throw 'Every NativeAOT scenario must belong to a supported CI group.'
}
if ($Scenario.Count -gt 0) {
    $unknown = @($Scenario | Where-Object { $_ -notin $scenarios.id })
    if ($unknown.Count -gt 0) { throw "Unknown NativeAOT scenarios: $($unknown -join ', ')" }
    $scenarios = @($scenarios | Where-Object { $_.id -in $Scenario })
}
if ($ScenarioGroup -ne 'all') {
    $scenarios = @($scenarios | Where-Object { $_.group -eq $ScenarioGroup })
}
if ($scenarios.Count -eq 0) {
    throw 'The requested NativeAOT scenario selection is empty.'
}

$artifactRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("OfficeIMO-AotValidation-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $artifactRoot -Force | Out-Null
$results = [System.Collections.Generic.List[object]]::new()
$failures = [System.Collections.Generic.List[string]]::new()

function Remove-AotArtifactDirectory {
    param(
        [Parameter(Mandatory)]
        [string] $Path
    )

    if (-not [System.IO.Directory]::Exists($Path)) {
        return
    }

    $maximumAttempts = 10
    for ($attempt = 1; $attempt -le $maximumAttempts; $attempt++) {
        try {
            [System.IO.Directory]::Delete($Path, $true)
            return
        } catch [System.IO.IOException] {
            if ($attempt -eq $maximumAttempts) {
                throw
            }

            # Windows can retain the just-executed native image briefly while
            # antivirus and process teardown release their file handles.
            Start-Sleep -Milliseconds (100 * $attempt)
        } catch [System.UnauthorizedAccessException] {
            if ($attempt -eq $maximumAttempts) {
                throw
            }

            Start-Sleep -Milliseconds (100 * $attempt)
        }
    }
}

try {
    foreach ($scenarioDefinition in $scenarios) {
        $projectPath = Join-Path $RepositoryRoot $scenarioDefinition.project
        $executableSha256 = $null
        $runOutput = @()
        $runExit = $null
        $publishPath = Join-Path $artifactRoot $scenarioDefinition.id
        $sdkArtifactsPath = Join-Path $artifactRoot ($scenarioDefinition.id + '-sdk')

        Write-Host "[$($scenarioDefinition.id)] restore and publish NativeAOT in isolated SDK state" -ForegroundColor Cyan
        $publishArguments = [System.Collections.Generic.List[string]]::new()
        # Keep each isolated graph in foreground compiler processes so completed
        # scenarios do not leave build servers competing with the next publisher.
        @('publish', $projectPath, '--configuration', $Configuration, '--runtime', $RuntimeIdentifier, '--artifacts-path', $sdkArtifactsPath, '--output', $publishPath,
            '--disable-build-servers', '-m:1', '-nodeReuse:false', '-p:UseSharedCompilation=false') |
            ForEach-Object { $publishArguments.Add($_) }
        if ($scenarioDefinition.targetFramework) {
            $publishArguments.Add('--framework')
            $publishArguments.Add([string] $scenarioDefinition.targetFramework)
        }
        if ($scenarioDefinition.aotValidation) {
            $publishArguments.Add('-p:AotValidation=true')
        }
        if ($RuntimeIdentifier -like 'linux-*' -and $scenarioDefinition.targetFramework -eq 'net8.0') {
            # Bound the .NET 8 native compiler as well as MSBuild on hosted Linux.
            $publishArguments.Add('-p:IlcSingleThreaded=true')
        }

        $publishOutput = [System.Collections.Generic.List[string]]::new()
        $publishEnvironment = @{}
        try {
            if ($RuntimeIdentifier -like 'linux-*' -and $scenarioDefinition.targetFramework -eq 'net8.0') {
                # Reserve room for native code generation and the hosted runner.
                # These defaults affect publisher children only; explicit caller
                # settings are preserved and native execution uses the original environment.
                # .NET 8 NativeAOT accepts bare hexadecimal digits, not a 0x prefix.
                $defaults = @{ DOTNET_GCHeapHardLimit = '200000000'; DOTNET_GCConserveMemory = '5' }
                foreach ($name in $defaults.Keys) {
                    $previous = [Environment]::GetEnvironmentVariable($name, 'Process')
                    if ([string]::IsNullOrEmpty($previous)) {
                        $publishEnvironment[$name] = $previous
                        [Environment]::SetEnvironmentVariable($name, $defaults[$name], 'Process')
                    }
                }
            }
            & dotnet @publishArguments 2>&1 | ForEach-Object {
                $publishLine = [string] $_
                $publishOutput.Add($publishLine)
                # Preserve compiler progress and diagnostics even if the runner exits
                # before this publish can produce its final JSON receipt.
                Write-Host $publishLine
            }
            $publishExit = $LASTEXITCODE
        } finally {
            foreach ($name in $publishEnvironment.Keys) {
                [Environment]::SetEnvironmentVariable($name, $publishEnvironment[$name], 'Process')
            }
        }
        $diagnosticCodes = @([regex]::Matches(($publishOutput -join "`n"), '\b(?:IL|NETSDK|CS)\d{4}\b') |
            ForEach-Object Value | Sort-Object -Unique)

        if ($publishExit -eq 0) {
            $projectName = [System.IO.Path]::GetFileNameWithoutExtension($projectPath)
            if ($scenarioDefinition.nativeLibrary) {
                # A shared library has no entry point: a host drives its exported C interface instead.
                $extension = if ($RuntimeIdentifier -like 'win-*') { '.dll' } elseif ($RuntimeIdentifier -like 'osx-*') { '.dylib' } else { '.so' }
                $executablePath = Join-Path $publishPath ($scenarioDefinition.nativeLibrary + $extension)
                $executableSha256 = (Get-FileHash -LiteralPath $executablePath -Algorithm SHA256).Hash
                Write-Host "[$($scenarioDefinition.id)] drive the native library through its C interface" -ForegroundColor Cyan
                $runOutput = @()
                $runExit = 0
                foreach ($fixture in $scenarioDefinition.fixtures) {
                    $fixtureName = [System.IO.Path]::GetFileNameWithoutExtension($fixture)
                    $journeyArguments = @('-NoLogo', '-NoProfile', '-File', (Join-Path $RepositoryRoot $scenarioDefinition.journey),
                        '-Library', $executablePath, '-Pdf', (Join-Path $RepositoryRoot $fixture),
                        '-WorkDirectory', (Join-Path $artifactRoot ($scenarioDefinition.id + '-journey-' + $fixtureName)))
                    $runOutput += "Fixture: $fixture"
                    $runOutput += @(& ([Environment]::ProcessPath) @journeyArguments 2>&1)
                    if ($LASTEXITCODE -ne 0) { $runExit = $LASTEXITCODE }
                }
            } else {
                $executablePath = Join-Path $publishPath ($projectName + $(if ($RuntimeIdentifier -like 'win-*') { '.exe' } else { '' }))
                $executableSha256 = (Get-FileHash -LiteralPath $executablePath -Algorithm SHA256).Hash
                Write-Host "[$($scenarioDefinition.id)] execute native binary" -ForegroundColor Cyan
                $runArguments = @($scenarioDefinition.runArguments)
                $runOutput = @(& $executablePath @runArguments 2>&1)
                $runExit = $LASTEXITCODE
            }
            $status = if ($runExit -eq 0) { 'passed' } else { 'runtime-failed' }
            $runOutput | ForEach-Object { Write-Host $_ }
        } else {
            $status = 'publish-blocked'
            $publishOutput | Where-Object { $_ -match '\b(?:IL|NETSDK|CS)\d{4}\b|:\s+error\s+' } | Select-Object -Unique | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
        }

        if ($status -ne 'passed') {
            $failures.Add("$($scenarioDefinition.id): expected a passing native publish and execution, got $status")
        }

        $results.Add([pscustomobject] [ordered]@{
            id = $scenarioDefinition.id
            title = $scenarioDefinition.title
            status = $status
            diagnosticCodes = $diagnosticCodes
            targetFramework = [string] $scenarioDefinition.targetFramework
            publishExitCode = $publishExit
            runExitCode = $runExit
            executableSha256 = $executableSha256
            runOutput = @($runOutput | ForEach-Object { [string] $_ })
        })

        # NativeAOT SDK state can be several gigabytes for the broad Reader and
        # rendering graphs. Each result is complete once its executable exits,
        # so release that scenario before starting the next one instead of
        # making the final linker compete with every previous publish.
        foreach ($scenarioArtifactPath in @($publishPath, $sdkArtifactsPath, (Join-Path $artifactRoot ($scenarioDefinition.id + '-journey')))) {
            Remove-AotArtifactDirectory -Path $scenarioArtifactPath
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
            scenarioGroup = $ScenarioGroup
            scenarioCount = $results.Count
            scenarios = @($results)
        } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $resolvedOutputPath -Encoding utf8
        Write-Host "Wrote $resolvedOutputPath" -ForegroundColor Green
    }

    if ($failures.Count -gt 0) {
        throw "NativeAOT validation did not match the documented contract:`n- $($failures -join "`n- ")"
    }
} finally {
    Remove-AotArtifactDirectory -Path $artifactRoot
}
