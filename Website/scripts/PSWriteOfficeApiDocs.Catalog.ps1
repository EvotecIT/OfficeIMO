function Test-DocumentationCatalog {
    param(
        [Parameter(Mandatory)][string] $Path,
        [Parameter(Mandatory)][version] $ExpectedVersion,
        [Parameter(Mandatory)][int] $ExpectedCommandCount
    )

    try {
        $catalog = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json
        $familyTotal = @($catalog.families | Measure-Object commandCount -Sum).Sum
        return [version]$catalog.module.version -eq $ExpectedVersion -and
            $catalog.module.commandCount -eq $ExpectedCommandCount -and
            $familyTotal -eq $ExpectedCommandCount
    } catch { return $false }
}

function Set-PSWriteOfficeDocumentationSnapshot {
    param(
        [Parameter(Mandatory)][string] $SiteRootPath,
        [Parameter(Mandatory)] $Source
    )

    $repository = [regex]::Escape($Source.Repo)
    $module = Import-PowerShellDataFile -LiteralPath (Join-Path $SiteRootPath 'data/apidocs/powershell/PSWriteOffice.psd1')
    $sourceRef = [Uri]::EscapeDataString('v' + [string] $module.ModuleVersion)
    # Only floating repository refs change. Explicit historical tags, other
    # repositories, URL suffixes and branches such as main-next retain their meaning.
    $pattern = '(?<prefix>https://(?:github\.com/' + $repository +
        '/(?:blob|tree)/|raw\.githubusercontent\.com/' + $repository +
        '/))(?:main|master)(?=[/?#\s"''<>)]|$)'
    $replacement = [System.Text.RegularExpressions.MatchEvaluator] {
        param($match)
        $match.Groups['prefix'].Value + $sourceRef
    }

    $catalogPath = Join-Path $SiteRootPath 'data/pswriteoffice_command_catalog.json'
    $catalog = Get-Content -LiteralPath $catalogPath -Raw | ConvertFrom-Json -Depth 20
    # A family's examples belong to this module version even if its catalog was
    # previously pinned. Historical links in guide prose are handled separately.
    $examplesPattern = '^(?<prefix>https://github\.com/' + $repository +
        '/tree/)[^/]+(?=/Examples(?:[/#?]|$))'
    $catalogChanged = $false
    foreach ($family in @($catalog.families)) {
        $examplesUrl = [regex]::Replace([string] $family.examplesUrl, $examplesPattern, $replacement)
        if ($examplesUrl -cne $family.examplesUrl) {
            $family.examplesUrl = $examplesUrl
            $catalogChanged = $true
        }
    }
    if ($catalogChanged) {
        [IO.File]::WriteAllText($catalogPath, ($catalog | ConvertTo-Json -Depth 20 -Compress) + [Environment]::NewLine)
    }

    $pagesChanged = $false
    $documentationRoot = Join-Path $SiteRootPath 'content/docs/pswriteoffice'
    foreach ($page in Get-ChildItem -LiteralPath $documentationRoot -Filter '*.md' -File -Recurse) {
        $content = [IO.File]::ReadAllText($page.FullName)
        $pinnedContent = [regex]::Replace($content, $pattern, $replacement)
        $family = @($catalog.families | Where-Object id -CEQ $page.Directory.Name)
        if ($page.Name -ceq 'index.md' -and $family.Count -eq 1) {
            # Imported and fallback summaries follow the same catalog totals.
            $summaryCount = [regex]::new('\b\d+(?= exported commands\b)|(?<=\bexports )\d+(?= commands\b)')
            $pinnedContent = $summaryCount.Replace($pinnedContent, [string] $family[0].commandCount, 1)
        }
        if ($content -cne $pinnedContent) {
            [IO.File]::WriteAllText($page.FullName, $pinnedContent, [Text.UTF8Encoding]::new($false))
            $pagesChanged = $true
        }
    }
    [PSCustomObject]@{ CatalogUpdated = $catalogChanged; DocumentationUpdated = $pagesChanged }
}

function Get-DocumentationCatalogFingerprint {
    param([Parameter(Mandatory)][string] $Path)

    # Git checkout line endings and UTF-8 BOMs do not change catalog content.
    $text = [IO.File]::ReadAllText($Path).Replace("`r`n", "`n")
    $bytes = [Text.UTF8Encoding]::new($false).GetBytes($text)
    $hash = [Security.Cryptography.SHA256]::Create()
    try {
        [ordered]@{
            bytes = $bytes.Length
            sha256 = [Convert]::ToHexString($hash.ComputeHash($bytes)).ToLowerInvariant()
        }
    } finally { $hash.Dispose() }
}

function Assert-DocumentationCatalogSnapshot {
    param(
        [Parameter(Mandatory)][string] $CatalogPath,
        [Parameter(Mandatory)][string] $EvidencePath,
        [Parameter(Mandatory)][string] $ExpectedSourceRef,
        [Parameter(Mandatory)][version] $ExpectedVersion,
        [Parameter(Mandatory)][int] $ExpectedCommandCount
    )

    try {
        $evidence = Get-Content -LiteralPath $EvidencePath -Raw -ErrorAction Stop | ConvertFrom-Json
        $actual = Get-DocumentationCatalogFingerprint -Path $CatalogPath
        $valid = $evidence.schemaVersion -eq 1 -and $evidence.moduleName -ceq 'PSWriteOffice' -and
            $evidence.sourceRef -ceq $ExpectedSourceRef -and
            $evidence.sourceCommit -ceq $ExpectedSourceRef -and
            [version]$evidence.moduleVersion -eq $ExpectedVersion -and
            $evidence.normalization -ceq 'utf8-no-bom-lf' -and
            $evidence.producer.path -ceq 'Build/Export-WebsiteDocumentationCatalog.ps1' -and
            $evidence.producer.fingerprint.bytes -gt 0 -and
            $evidence.producer.fingerprint.sha256 -cmatch '^[0-9a-f]{64}$' -and
            $evidence.sourceCatalog.bytes -gt 0 -and
            $evidence.sourceCatalog.sha256 -cmatch '^[0-9a-f]{64}$' -and
            $evidence.derivedCatalog.bytes -eq $actual.bytes -and
            $evidence.derivedCatalog.sha256 -ceq $actual.sha256 -and
            (Test-DocumentationCatalog -Path $CatalogPath -ExpectedVersion $ExpectedVersion -ExpectedCommandCount $ExpectedCommandCount)
        if ($valid) { return $evidence }
    } catch { }
    throw 'The checked-in PSWriteOffice documentation catalog does not match its pinned source/derived evidence.'
}

function New-DocumentationCatalogSnapshot {
    param(
        [Parameter(Mandatory)][string] $RepositoryRoot,
        [Parameter(Mandatory)][string] $SourceCommit,
        [Parameter(Mandatory)][string] $SourceRef,
        [Parameter(Mandatory)][version] $ExpectedVersion,
        [Parameter(Mandatory)][int] $ExpectedCommandCount
    )

    # Run the owning producer from the already verified clean, pinned source.
    # Caller-supplied API exports cannot override this documentation authority.
    $producerPath = 'Build/Export-WebsiteDocumentationCatalog.ps1'
    $producer = Join-Path $RepositoryRoot $producerPath
    $temporaryCatalog = [IO.Path]::GetTempFileName()
    try {
        & $producer -RepositoryRoot $RepositoryRoot -OutputPath $temporaryCatalog | Out-Null
        if (-not (Test-DocumentationCatalog -Path $temporaryCatalog -ExpectedVersion $ExpectedVersion -ExpectedCommandCount $ExpectedCommandCount)) {
            throw 'The pinned PSWriteOffice producer generated an incomplete documentation catalog.'
        }
        [PSCustomObject]@{
            Content = [IO.File]::ReadAllText($temporaryCatalog)
            Evidence = [ordered]@{
                schemaVersion = 1
                moduleName = 'PSWriteOffice'
                moduleVersion = [string] $ExpectedVersion
                sourceRef = $SourceRef
                sourceCommit = $SourceCommit
                normalization = 'utf8-no-bom-lf'
                producer = [ordered]@{ path = $producerPath; fingerprint = Get-DocumentationCatalogFingerprint -Path $producer }
                sourceCatalog = Get-DocumentationCatalogFingerprint -Path $temporaryCatalog
                derivedCatalog = $null
            }
        }
    } finally { Remove-Item -LiteralPath $temporaryCatalog }
}

function Save-DocumentationCatalogEvidence {
    param(
        [Parameter(Mandatory)][string] $CatalogPath,
        [Parameter(Mandatory)][string] $EvidencePath,
        [Parameter(Mandatory)] $Evidence
    )

    # Consumer link rewriting is recorded separately from the source output.
    $Evidence.derivedCatalog = Get-DocumentationCatalogFingerprint -Path $CatalogPath
    [IO.File]::WriteAllText($EvidencePath, ($Evidence | ConvertTo-Json -Depth 8) + "`n", [Text.UTF8Encoding]::new($false))
}
