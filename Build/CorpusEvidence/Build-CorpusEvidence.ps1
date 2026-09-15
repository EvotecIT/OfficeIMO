param(
    [string] $CatalogPath = (Join-Path $PSScriptRoot 'corpus-evidence.json'),
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Docs\Compatibility\generated'),
    [switch] $Verify
)

$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..')).Path

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

function Get-ManifestRecordCount {
    param([object] $Manifest, [string] $ValidationKind, [string] $Label)
    switch ($ValidationKind) {
        'office-collections' {
            $records = @($Manifest.collections)
            foreach ($record in $records) {
                Assert-Text $record.producer "$Label collection '$($record.id)' is missing producer provenance."
                Assert-Text $record.producerVersion "$Label collection '$($record.id)' is missing producerVersion provenance."
                foreach ($artifact in @($record.artifacts)) {
                    Assert-Sha256 $artifact.sha256 "$Label artifact '$($record.id)/$($artifact.file)' is missing a stable SHA-256."
                }
            }
            return $records.Count
        }
        'word-artifacts' {
            $records = @($Manifest.artifacts)
            foreach ($record in $records) {
                Assert-Text $record.producer "$Label artifact '$($record.id)' is missing producer provenance."
                Assert-Text $record.producerVersion "$Label artifact '$($record.id)' is missing producerVersion provenance."
                Assert-Text $record.lossPolicy "$Label artifact '$($record.id)' is missing its semantic/package loss policy."
                if ($record.path) { Assert-Sha256 $record.sha256 "$Label artifact '$($record.id)' is missing a stable SHA-256." }
            }
            return $records.Count
        }
        'open-document-fixtures' {
            $records = @($Manifest.fixtures)
            foreach ($record in $records) {
                Assert-Text $record.producer "$Label fixture '$($record.file)' is missing producer provenance."
                Assert-Text $record.producerVersion "$Label fixture '$($record.file)' is missing producerVersion provenance."
                Assert-Sha256 $record.sha256 "$Label fixture '$($record.file)' is missing a stable SHA-256."
            }
            foreach ($record in @($Manifest.externalArtifacts)) {
                Assert-Text $record.producer "$Label external artifact '$($record.id)' is missing producer provenance."
                Assert-Text $record.producerVersion "$Label external artifact '$($record.id)' is missing producerVersion provenance."
                Assert-Sha256 $record.semanticTextSha256 "$Label external artifact '$($record.id)' is missing a stable semantic hash."
            }
            return $records.Count + @($Manifest.externalArtifacts).Count
        }
        'rtf-fixtures' {
            $records = @($Manifest.fixtures) + @($Manifest.externalArtifacts)
            foreach ($record in $records) {
                Assert-Text $record.producer "$Label record '$($record.id)' is missing producer provenance."
                Assert-Text $record.producerVersion "$Label record '$($record.id)' is missing producerVersion provenance."
                Assert-Sha256 $record.sha256 "$Label record '$($record.id)' is missing a stable SHA-256."
            }
            return $records.Count
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
    $recordCount = Get-ManifestRecordCount $manifest ([string] $corpus.validationKind) ([string] $corpus.id)
    $reports.Add([ordered]@{
        id = [string] $corpus.id
        owner = [string] $corpus.owner
        manifest = ([string] $corpus.manifest).Replace('\\', '/')
        manifestSha256 = (Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash.ToLowerInvariant()
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
    Write-Host "Verified $($reports.Count) corpus contracts with $($report.summary.recordCount) producer records."
    return
}

[IO.Directory]::CreateDirectory($OutputDirectory) | Out-Null
$utf8 = [Text.UTF8Encoding]::new($false)
foreach ($name in $outputs.Keys) {
    [IO.File]::WriteAllText((Join-Path $OutputDirectory $name), $outputs[$name], $utf8)
}
Write-Host "Generated $($reports.Count) corpus contracts with $($report.summary.recordCount) producer records."
