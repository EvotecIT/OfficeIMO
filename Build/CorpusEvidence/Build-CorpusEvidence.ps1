param(
    [string] $CatalogPath = (Join-Path $PSScriptRoot 'corpus-evidence.json'),
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Docs\Compatibility\generated'),
    [switch] $Verify
)

$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..')).Path
$verifiedArtifactCount = 0

function Resolve-RepositoryPath {
    param([Parameter(Mandatory)][string] $RelativePath)
    if ([IO.Path]::IsPathRooted($RelativePath) -or $RelativePath -match '(^|[/\\])\.\.([/\\]|$)') {
        throw "Corpus evidence path must be repository-relative: $RelativePath"
    }
    $fullPath = [IO.Path]::GetFullPath((Join-Path $repoRoot $RelativePath))
    if (-not $fullPath.StartsWith($repoRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Corpus evidence path escapes the repository: $RelativePath"
    }
    return $fullPath
}

function Assert-Text {
    param([object] $Value, [string] $Message)
    if ([string]::IsNullOrWhiteSpace([string] $Value)) { throw $Message }
}

function Assert-Sha256 {
    param([object] $Value, [string] $Message)
    if ([string] $Value -notmatch '^[a-fA-F0-9]{64}$') { throw $Message }
}

function Assert-ExecutableSourceTest {
    param(
        [Parameter(Mandatory)][object] $Record,
        [Parameter(Mandatory)][string] $Label
    )
    Assert-Text $Record.sourceProject "$Label is missing sourceProject."
    Assert-Text $Record.sourceFile "$Label is missing sourceFile."
    Assert-Text $Record.sourceExecution "$Label is missing sourceExecution."
    Assert-Text $Record.sourceTest "$Label is missing sourceTest."
    $execution = [string] $Record.sourceExecution
    if ($execution -notin @('xunit', 'microsoft-office-xunit', 'benchmark-validation')) {
        throw "$Label selected unsupported sourceExecution '$execution'."
    }
    $projectPath = Resolve-RepositoryPath ([string] $Record.sourceProject)
    if (-not (Test-Path -LiteralPath $projectPath -PathType Leaf) -or
        -not $projectPath.EndsWith('.csproj', [StringComparison]::OrdinalIgnoreCase)) {
        throw "$Label source project is missing or is not a .csproj: $($Record.sourceProject)"
    }
    $projectDirectory = [IO.Path]::GetDirectoryName($projectPath)
    $sourcePath = Resolve-ContainedArtifactPath $projectDirectory ([string] $Record.sourceFile) $Label
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf) -or
        -not $sourcePath.EndsWith('.cs', [StringComparison]::OrdinalIgnoreCase)) {
        throw "$Label source file is missing or is not a .cs file: $($Record.sourceFile)"
    }
    $methodPattern = '\b(?:public|private|internal|protected)\s+(?:static\s+)?(?:async\s+)?[A-Za-z0-9_<>,?\[\].]+\s+' +
        [regex]::Escape([string] $Record.sourceTest) + '\s*\('
    $matches = @(Select-String -LiteralPath $sourcePath -Pattern $methodPattern)
    if ($matches.Count -eq 0) {
        throw "$Label sourceTest '$($Record.sourceTest)' does not resolve to a C# method in '$($Record.sourceProject)'."
    }
    if ($matches.Count -gt 1) {
        $locations = $matches | ForEach-Object { "$($_.Path):$($_.LineNumber)" }
        throw "$Label sourceTest '$($Record.sourceTest)' is ambiguous in '$($Record.sourceProject)': $($locations -join ', ')"
    }
    $sourceLines = @(Get-Content -LiteralPath $sourcePath)
    $attributeStart = [Math]::Max(0, $matches[0].LineNumber - 12)
    $attributeCount = $matches[0].LineNumber - $attributeStart
    $declarationContext = @($sourceLines[$attributeStart..($attributeStart + $attributeCount - 1)]) -join "`n"
    if ($execution -in @('xunit', 'microsoft-office-xunit') -and
        $declarationContext -notmatch '\[[A-Za-z0-9_.]*(Fact|Theory)(Attribute)?(?:\(|\])') {
        throw "$Label sourceTest '$($Record.sourceTest)' is not declared as an xUnit Fact or Theory."
    }
    if ($execution -eq 'benchmark-validation' -and
        $declarationContext -notmatch '\[GlobalSetup(?:Attribute)?(?:\(|\])') {
        throw "$Label sourceTest '$($Record.sourceTest)' is not declared as a BenchmarkDotNet GlobalSetup."
    }
}

function Resolve-ContainedArtifactPath {
    param(
        [Parameter(Mandatory)][string] $BasePath,
        [Parameter(Mandatory)][string] $RelativePath,
        [Parameter(Mandatory)][string] $Label
    )
    if ([IO.Path]::IsPathRooted($RelativePath) -or $RelativePath -match '(^|[/\\])\.\.([/\\]|$)') {
        throw "$Label artifact path must stay below its corpus root: $RelativePath"
    }
    $baseFullPath = [IO.Path]::GetFullPath($BasePath).TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
    $fullPath = [IO.Path]::GetFullPath((Join-Path $baseFullPath $RelativePath))
    if (-not $fullPath.StartsWith($baseFullPath + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw "$Label artifact path escapes its corpus root: $RelativePath"
    }
    return $fullPath
}

function Get-CanonicalTextSha256 {
    param([Parameter(Mandatory)][string] $Path)
    $text = (Get-Content -LiteralPath $Path -Raw).Replace("`r`n", "`n").Replace("`r", "`n")
    $bytes = [Text.UTF8Encoding]::new($false).GetBytes($text)
    return [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes)).ToLowerInvariant()
}

function Assert-ArtifactIdentity {
    param(
        [Parameter(Mandatory)][string] $Path,
        [Parameter(Mandatory)][string] $ExpectedSha256,
        [object] $ExpectedBytes,
        [string] $HashMode = 'raw',
        [Parameter(Mandatory)][string] $Label
    )
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "$Label artifact is missing: $Path" }
    $actualSha256 = if ($HashMode -eq 'canonical-text') {
        Get-CanonicalTextSha256 $Path
    } elseif ($HashMode -eq 'raw') {
        (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
    } else {
        throw "$Label artifact selected unsupported hashMode '$HashMode'."
    }
    if (-not $actualSha256.Equals($ExpectedSha256, [StringComparison]::OrdinalIgnoreCase)) {
        throw "$Label artifact hash mismatch for '$Path'. Expected $ExpectedSha256, actual $actualSha256."
    }
    if ($null -ne $ExpectedBytes -and [long] $ExpectedBytes -ne (Get-Item -LiteralPath $Path).Length) {
        $actualBytes = (Get-Item -LiteralPath $Path).Length
        throw "$Label artifact byte length mismatch for '$Path'. Expected $ExpectedBytes, actual $actualBytes."
    }
    $script:verifiedArtifactCount++
}

function Get-ManifestRecordCount {
    param([object] $Manifest, [string] $ManifestPath, [string] $ValidationKind, [string] $Label)
    $manifestDirectory = [IO.Path]::GetDirectoryName($ManifestPath)
    switch ($ValidationKind) {
        'office-collections' {
            $documentsRoot = [IO.Path]::GetDirectoryName($manifestDirectory)
            $records = @($Manifest.collections)
            foreach ($record in $records) {
                Assert-Text $record.producer "$Label collection '$($record.id)' is missing producer provenance."
                Assert-Text $record.producerVersion "$Label collection '$($record.id)' is missing producerVersion provenance."
                foreach ($artifact in @($record.artifacts)) {
                    Assert-Sha256 $artifact.sha256 "$Label artifact '$($record.id)/$($artifact.file)' is missing a stable SHA-256."
                    $relativePath = Join-Path ([string] $record.root) ([string] $artifact.file)
                    $artifactPath = Resolve-ContainedArtifactPath $documentsRoot $relativePath $Label
                    Assert-ArtifactIdentity $artifactPath ([string] $artifact.sha256) $null 'raw' "$Label artifact '$($record.id)/$($artifact.file)'"
                    if ($artifact.approvedReport) {
                        $reportPath = Resolve-ContainedArtifactPath ([IO.Path]::GetDirectoryName($artifactPath)) ([string] $artifact.approvedReport) $Label
                        if (-not (Test-Path -LiteralPath $reportPath -PathType Leaf)) {
                            throw "$Label approved report is missing for '$($record.id)/$($artifact.file)': $reportPath"
                        }
                    }
                }
            }
            return $records.Count
        }
        'word-artifacts' {
            $documentsRoot = [IO.Path]::GetDirectoryName([IO.Path]::GetDirectoryName($manifestDirectory))
            $records = @($Manifest.artifacts)
            foreach ($record in $records) {
                Assert-Text $record.producer "$Label artifact '$($record.id)' is missing producer provenance."
                Assert-Text $record.producerVersion "$Label artifact '$($record.id)' is missing producerVersion provenance."
                Assert-Text $record.lossPolicy "$Label artifact '$($record.id)' is missing its semantic/package loss policy."
                Assert-ExecutableSourceTest $record "$Label artifact '$($record.id)'"
                if ($record.path) {
                    Assert-Sha256 $record.sha256 "$Label artifact '$($record.id)' is missing a stable SHA-256."
                    $artifactPath = Resolve-ContainedArtifactPath $documentsRoot ([string] $record.path) $Label
                    $hashMode = if ($record.hashMode) { [string] $record.hashMode } else { 'raw' }
                    Assert-ArtifactIdentity $artifactPath ([string] $record.sha256) $null $hashMode "$Label artifact '$($record.id)'"
                }
            }
            return $records.Count
        }
        'open-document-fixtures' {
            $records = @($Manifest.fixtures)
            foreach ($record in $records) {
                Assert-Text $record.producer "$Label fixture '$($record.file)' is missing producer provenance."
                Assert-Text $record.producerVersion "$Label fixture '$($record.file)' is missing producerVersion provenance."
                Assert-Sha256 $record.sha256 "$Label fixture '$($record.file)' is missing a stable SHA-256."
                $artifactPath = Resolve-ContainedArtifactPath $manifestDirectory ([string] $record.file) $Label
                Assert-ArtifactIdentity $artifactPath ([string] $record.sha256) $record.bytes 'raw' "$Label fixture '$($record.file)'"
            }
            foreach ($record in @($Manifest.externalArtifacts)) {
                Assert-Text $record.producer "$Label external artifact '$($record.id)' is missing producer provenance."
                Assert-Text $record.producerVersion "$Label external artifact '$($record.id)' is missing producerVersion provenance."
                Assert-Sha256 $record.semanticTextSha256 "$Label external artifact '$($record.id)' is missing a stable semantic hash."
                Assert-Text $record.sourceUrl "$Label external artifact '$($record.id)' is missing its verification URL."
                if ([long] $record.minBytes -lt 1 -or [long] $record.maxBytes -lt [long] $record.minBytes) {
                    throw "$Label external artifact '$($record.id)' is missing a valid package-size oracle."
                }
                if ([long] $record.paragraphCount -lt 1) { throw "$Label external artifact '$($record.id)' is missing its paragraph-count oracle." }
            }
            return $records.Count + @($Manifest.externalArtifacts).Count
        }
        'rtf-fixtures' {
            $fixtures = @($Manifest.fixtures)
            foreach ($record in $fixtures) {
                Assert-Text $record.producer "$Label record '$($record.id)' is missing producer provenance."
                Assert-Text $record.producerVersion "$Label record '$($record.id)' is missing producerVersion provenance."
                Assert-Sha256 $record.sha256 "$Label record '$($record.id)' is missing a stable SHA-256."
                $artifactPath = Resolve-ContainedArtifactPath $manifestDirectory ([string] $record.file) $Label
                Assert-ArtifactIdentity $artifactPath ([string] $record.sha256) $record.bytes 'raw' "$Label fixture '$($record.id)'"
            }
            $external = @($Manifest.externalArtifacts)
            foreach ($record in $external) {
                Assert-Text $record.producer "$Label record '$($record.id)' is missing producer provenance."
                Assert-Text $record.producerVersion "$Label record '$($record.id)' is missing producerVersion provenance."
                Assert-Sha256 $record.sha256 "$Label record '$($record.id)' is missing a stable SHA-256."
                Assert-Text $record.sourceUrl "$Label external artifact '$($record.id)' is missing its verification URL."
                if ([long] $record.bytes -lt 1) { throw "$Label external artifact '$($record.id)' is missing its expected byte length." }
                if (@($record.requiredHeaderFragments).Count -eq 0) { throw "$Label external artifact '$($record.id)' is missing its header oracle." }
            }
            return $fixtures.Count + $external.Count
        }
        'pdf-source-cases' {
            $sources = @{}
            foreach ($source in @($Manifest.sources)) {
                Assert-Text $source.repository "$Label source '$($source.id)' is missing repository provenance."
                Assert-Text $source.commit "$Label source '$($source.id)' is missing an immutable producer version."
                $sources[$source.id] = $source
            }
            $records = @($Manifest.cases)
            foreach ($record in $records) {
                if (-not $sources.ContainsKey([string] $record.source)) { throw "$Label case '$($record.id)' references an unknown producer source." }
                Assert-Sha256 $record.sha256 "$Label case '$($record.id)' is missing a stable SHA-256."
                $artifactPath = Resolve-ContainedArtifactPath $manifestDirectory ([string] $record.file) $Label
                Assert-ArtifactIdentity $artifactPath ([string] $record.sha256) $record.byteLength 'raw' "$Label case '$($record.id)'"
            }
            return $records.Count
        }
        default { throw "Unsupported corpus validation kind '$ValidationKind'." }
    }
}

$catalog = Get-Content -LiteralPath $CatalogPath -Raw | ConvertFrom-Json
if ($catalog.schemaVersion -ne 1) { throw "Unsupported corpus evidence schemaVersion '$($catalog.schemaVersion)'." }
$ids = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$reports = [Collections.Generic.List[object]]::new()

foreach ($corpus in @($catalog.corpora)) {
    Assert-Text $corpus.id 'Every corpus requires an id.'
    if (-not $ids.Add([string] $corpus.id)) { throw "Duplicate corpus evidence id '$($corpus.id)'." }
    Assert-Text $corpus.owner "Corpus '$($corpus.id)' is missing its owner."
    Assert-Text $corpus.diffPolicy.kind "Corpus '$($corpus.id)' is missing a diff-policy kind."
    Assert-Text $corpus.diffPolicy.acceptance "Corpus '$($corpus.id)' is missing an acceptance rule."
    if (@($corpus.diffPolicy.stableFields).Count -eq 0) { throw "Corpus '$($corpus.id)' has no stable comparison fields." }

    $manifestPath = Resolve-RepositoryPath ([string] $corpus.manifest)
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { throw "Corpus manifest is missing: $($corpus.manifest)" }
    $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
    $recordCount = Get-ManifestRecordCount $manifest $manifestPath ([string] $corpus.validationKind) ([string] $corpus.id)
    $reports.Add([ordered]@{
        id = [string] $corpus.id
        owner = [string] $corpus.owner
        manifest = ([string] $corpus.manifest).Replace('\\', '/')
        manifestSha256 = Get-CanonicalTextSha256 $manifestPath
        recordCount = $recordCount
        diffPolicy = $corpus.diffPolicy
    })
}

$report = [ordered]@{
    schemaVersion = 1
    generatedFrom = 'Build/CorpusEvidence/corpus-evidence.json'
    summary = [ordered]@{
        corpusCount = $reports.Count
        recordCount = [Linq.Enumerable]::Sum([int[]] @($reports | ForEach-Object recordCount))
        verifiedArtifactCount = $verifiedArtifactCount
    }
    corpora = $reports
}
$json = ($report | ConvertTo-Json -Depth 10) + "`n"
$markdown = [Text.StringBuilder]::new()
[void] $markdown.AppendLine('# Cross-producer corpus evidence')
[void] $markdown.AppendLine()
[void] $markdown.AppendLine('Generated from `Build/CorpusEvidence/corpus-evidence.json`. Every listed corpus has producer/version provenance and an executable stable package or semantic diff policy.')
[void] $markdown.AppendLine()
[void] $markdown.AppendLine('| Corpus | Owner | Records | Diff policy | Stable comparisons |')
[void] $markdown.AppendLine('| --- | --- | ---: | --- | --- |')
foreach ($item in $reports) {
    [void] $markdown.AppendLine("| $($item.id) | $($item.owner) | $($item.recordCount) | $($item.diffPolicy.kind) | $(@($item.diffPolicy.stableFields) -join ', ') |")
}
$markdownText = $markdown.ToString().Replace("`r`n", "`n")

$outputs = [ordered]@{
    'corpus-evidence.json' = $json.Replace("`r`n", "`n")
    'corpus-evidence.md' = $markdownText
}
if ($Verify) {
    foreach ($name in $outputs.Keys) {
        $path = Join-Path $OutputDirectory $name
        if (-not (Test-Path -LiteralPath $path -PathType Leaf) -or
            (Get-Content -LiteralPath $path -Raw).Replace("`r`n", "`n") -ne $outputs[$name]) {
            throw "Generated corpus evidence is missing or stale: $path"
        }
    }
    Write-Host "Verified $($reports.Count) corpus contracts with $($report.summary.recordCount) producer records and $verifiedArtifactCount checked-in artifacts."
    return
}

[IO.Directory]::CreateDirectory($OutputDirectory) | Out-Null
$utf8 = [Text.UTF8Encoding]::new($false)
foreach ($name in $outputs.Keys) {
    [IO.File]::WriteAllText((Join-Path $OutputDirectory $name), $outputs[$name], $utf8)
}
Write-Host "Generated $($reports.Count) corpus contracts with $($report.summary.recordCount) producer records and $verifiedArtifactCount checked-in artifacts."
