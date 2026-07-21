param(
    [string] $ManifestPath,
    [Parameter(Mandatory)]
    [string] $OutputDirectory,
    [ValidateSet('doc', 'xls', 'xlsb', 'ppt')]
    [string[]] $Format = @('doc', 'xls', 'xlsb', 'ppt'),
    [switch] $SkipOfficeImoGeneratedOutputs
)

$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $false

$repoRoot = Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')
if ([string]::IsNullOrWhiteSpace($ManifestPath)) {
    $ManifestPath = Join-Path $repoRoot 'OfficeIMO.TestAssets/Documents/OfficeInteroperabilityCorpus/corpus-manifest.json'
}
$manifestFile = Resolve-Path -LiteralPath $ManifestPath
$manifest = Get-Content -LiteralPath $manifestFile -Raw | ConvertFrom-Json
if ($manifest.schemaVersion -ne 2) {
    throw "Unsupported Office interoperability corpus schema '$($manifest.schemaVersion)'; expected 2."
}

$soffice = Get-Command soffice -ErrorAction SilentlyContinue
if ($null -eq $soffice) {
    $soffice = Get-Command libreoffice -ErrorAction SilentlyContinue
}
if ($null -eq $soffice) {
    throw 'LibreOffice was not found. Install soffice/libreoffice before running the external corpus oracle.'
}

$outputRoot = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $outputRoot) {
    if (@(Get-ChildItem -LiteralPath $outputRoot -Force).Count -ne 0) {
        throw "LibreOffice oracle output directory must be empty: $outputRoot"
    }
} else {
    [void] (New-Item -ItemType Directory -Path $outputRoot)
}

$documentsRoot = Split-Path -Parent (Split-Path -Parent $manifestFile)
$started = [DateTime]::UtcNow
$versionOutput = (& $soffice.Source --version 2>&1 | Out-String).Trim()
$results = [Collections.Generic.List[object]]::new()
$failures = [Collections.Generic.List[string]]::new()

function Get-ConversionTarget {
    param([Parameter(Mandatory)][string] $SourceFormat)

    switch ($SourceFormat) {
        'doc'  { return @{ Extension = '.docx'; Filter = 'docx:Office Open XML Text' } }
        'xls'  { return @{ Extension = '.xlsx'; Filter = 'xlsx:Calc MS Excel 2007 XML' } }
        'xlsb' { return @{ Extension = '.xlsx'; Filter = 'xlsx:Calc MS Excel 2007 XML' } }
        'ppt'  { return @{ Extension = '.pptx'; Filter = 'pptx:Impress MS PowerPoint 2007 XML' } }
        default { throw "Unsupported LibreOffice oracle source format '$SourceFormat'." }
    }
}

function Invoke-LibreOfficeConversion {
    param(
        [Parameter(Mandatory)][string] $SourcePath,
        [Parameter(Mandatory)][string] $TargetDirectory,
        [Parameter(Mandatory)][string] $FilterName,
        [Parameter(Mandatory)][string] $ProfileDirectory
    )

    [void] (New-Item -ItemType Directory -Path $TargetDirectory -Force)
    [void] (New-Item -ItemType Directory -Path $ProfileDirectory -Force)
    $profileUri = ([Uri]::new(
        [IO.Path]::GetFullPath($ProfileDirectory),
        [UriKind]::Absolute)).AbsoluteUri
    $output = & $soffice.Source `
        "-env:UserInstallation=$profileUri" `
        --headless --nologo --nolockcheck --nodefault --nofirststartwizard `
        --convert-to $FilterName --outdir $TargetDirectory $SourcePath 2>&1
    return @{
        ExitCode = $LASTEXITCODE
        Output = ($output | Out-String).Trim()
    }
}

function Test-LibreOfficeArtifact {
    param(
        [Parameter(Mandatory)][string] $CollectionId,
        [Parameter(Mandatory)][int] $Index,
        [Parameter(Mandatory)][string] $FormatId,
        [Parameter(Mandatory)][string] $SourceFormat,
        [Parameter(Mandatory)][string] $SourcePath,
        [Parameter(Mandatory)][string] $ArtifactName,
        [Parameter(Mandatory)][string] $ExpectedSha256,
        [Parameter(Mandatory)][string[]] $Oracles
    )

    $caseName = '{0}-{1:d3}' -f $CollectionId, $Index
    $caseRoot = Join-Path $outputRoot $caseName
    $convertedDirectory = Join-Path $caseRoot 'converted'
    $profileDirectory = Join-Path $caseRoot 'profile-open'
    $target = Get-ConversionTarget -SourceFormat $SourceFormat
    $expectedOutput = Join-Path $convertedDirectory ([IO.Path]::GetFileNameWithoutExtension($ArtifactName) + $target.Extension)
    $status = 'passed'
    $message = ''
    $renderOutput = $null

    try {
        if (-not (Test-Path -LiteralPath $SourcePath -PathType Leaf)) {
            throw "Source artifact does not exist: $SourcePath"
        }
        $actualSha256 = (Get-FileHash -LiteralPath $SourcePath -Algorithm SHA256).Hash.ToLowerInvariant()
        if (-not [string]::Equals(
                $actualSha256,
                $ExpectedSha256,
                [StringComparison]::OrdinalIgnoreCase)) {
            throw "Source artifact SHA-256 mismatch. Expected $ExpectedSha256, got $actualSha256."
        }

        $conversion = Invoke-LibreOfficeConversion `
            -SourcePath $SourcePath `
            -TargetDirectory $convertedDirectory `
            -FilterName $target.Filter `
            -ProfileDirectory $profileDirectory
        if ($conversion.ExitCode -ne 0) {
            throw "LibreOffice exited with code $($conversion.ExitCode): $($conversion.Output)"
        }
        if ((-not (Test-Path -LiteralPath $expectedOutput -PathType Leaf)) -or
            (Get-Item -LiteralPath $expectedOutput).Length -eq 0) {
            throw "LibreOffice did not create a non-empty '$($target.Extension)' artifact. Output: $($conversion.Output)"
        }

        if ($Oracles -contains 'libreoffice-render') {
            $renderDirectory = Join-Path $caseRoot 'rendered'
            $renderProfile = Join-Path $caseRoot 'profile-render'
            $render = Invoke-LibreOfficeConversion `
                -SourcePath $SourcePath `
                -TargetDirectory $renderDirectory `
                -FilterName 'pdf' `
                -ProfileDirectory $renderProfile
            $expectedPdf = Join-Path $renderDirectory ([IO.Path]::GetFileNameWithoutExtension($ArtifactName) + '.pdf')
            if ($render.ExitCode -ne 0) {
                throw "LibreOffice render exited with code $($render.ExitCode): $($render.Output)"
            }
            if ((-not (Test-Path -LiteralPath $expectedPdf -PathType Leaf)) -or
                (Get-Item -LiteralPath $expectedPdf).Length -eq 0) {
                throw "LibreOffice did not render a non-empty PDF. Output: $($render.Output)"
            }
            $renderOutput = [IO.Path]::GetRelativePath($outputRoot, $expectedPdf).Replace('\', '/')
        }
    } catch {
        $status = 'failed'
        $message = $_.Exception.Message
        $failures.Add("$CollectionId/$ArtifactName`: $message")
    }

    $results.Add([ordered]@{
        collectionId = $CollectionId
        formatId = $FormatId
        artifact = $ArtifactName
        sha256 = $ExpectedSha256
        status = $status
        convertedOutput = if (Test-Path -LiteralPath $expectedOutput -PathType Leaf) {
            [IO.Path]::GetRelativePath($outputRoot, $expectedOutput).Replace('\', '/')
        } else { $null }
        renderedOutput = $renderOutput
        message = $message
    })
}

$selectedCollections = @($manifest.collections | Where-Object {
    $Format -contains $_.format -and $_.oracles -contains 'libreoffice-open'
})
foreach ($collection in $selectedCollections) {
    $index = 0
    foreach ($artifact in $collection.artifacts) {
        $index++
        $relativeSource = Join-Path $collection.root $artifact.file
        $sourcePath = Join-Path $documentsRoot $relativeSource
        Test-LibreOfficeArtifact `
            -CollectionId $collection.id `
            -Index $index `
            -FormatId $collection.formatId `
            -SourceFormat $collection.format `
            -SourcePath $sourcePath `
            -ArtifactName $artifact.file `
            -ExpectedSha256 $artifact.sha256 `
            -Oracles $collection.oracles
    }
}

if (-not $SkipOfficeImoGeneratedOutputs) {
    $generatedRoot = Join-Path $outputRoot 'officeimo-generated'
    $generatorProject = Join-Path $repoRoot 'Build/OfficeInteroperabilityArtifacts/OfficeIMO.OfficeInteroperabilityArtifacts.Tool.csproj'
    Push-Location $repoRoot
    try {
        & dotnet run --project $generatorProject --configuration Release --framework net8.0 -- --output $generatedRoot
        if ($LASTEXITCODE -ne 0) {
            throw "OfficeIMO interoperability artifact generation failed with exit code $LASTEXITCODE."
        }
    } finally {
        Pop-Location
    }

    $generatedManifestPath = Join-Path $generatedRoot 'officeimo-artifacts.json'
    $generatedManifest = Get-Content -LiteralPath $generatedManifestPath -Raw | ConvertFrom-Json
    if ($generatedManifest.schemaVersion -ne 1) {
        throw "Unsupported OfficeIMO artifact manifest schema '$($generatedManifest.schemaVersion)'; expected 1."
    }
    $index = 0
    foreach ($artifact in @($generatedManifest.artifacts | Where-Object { $Format -contains $_.format })) {
        $index++
        $oracles = if ($artifact.format -eq 'ppt') {
            @('libreoffice-open', 'libreoffice-render')
        } else {
            @('libreoffice-open')
        }
        Test-LibreOfficeArtifact `
            -CollectionId 'officeimo-generated-binary' `
            -Index $index `
            -FormatId $artifact.formatId `
            -SourceFormat $artifact.format `
            -SourcePath (Join-Path $generatedRoot $artifact.file) `
            -ArtifactName $artifact.file `
            -ExpectedSha256 $artifact.sha256 `
            -Oracles $oracles
    }
}

$report = [ordered]@{
    schemaVersion = 1
    oracle = 'LibreOffice'
    oracleVersion = $versionOutput
    corpusSchemaVersion = $manifest.schemaVersion
    startedUtc = $started.ToString('O')
    completedUtc = [DateTime]::UtcNow.ToString('O')
    summary = [ordered]@{
        total = $results.Count
        passed = @($results | Where-Object status -eq 'passed').Count
        failed = @($results | Where-Object status -eq 'failed').Count
    }
    results = $results
}
$reportPath = Join-Path $outputRoot 'oracle-report.json'
$report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $reportPath -Encoding utf8NoBOM

Write-Host "LibreOffice corpus oracle: $($report.summary.passed)/$($report.summary.total) passed."
Write-Host "Report: $reportPath"
if ($failures.Count -ne 0) {
    throw "LibreOffice corpus oracle failed:`n$($failures -join [Environment]::NewLine)"
}
