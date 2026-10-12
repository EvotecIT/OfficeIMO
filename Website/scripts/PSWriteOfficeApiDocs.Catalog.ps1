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
        [Parameter(Mandatory)][string] $SiteRootPath,
        [Parameter(Mandatory)][version] $ExpectedVersion,
        [Parameter(Mandatory)][int] $ExpectedCommandCount
    )

    try {
        $evidence = Get-Content -LiteralPath $EvidencePath -Raw -ErrorAction Stop | ConvertFrom-Json
        $actual = Get-DocumentationCatalogFingerprint -Path $CatalogPath
        $valid = $evidence.schemaVersion -eq 2 -and $evidence.moduleName -ceq 'PSWriteOffice' -and
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
        if ($valid) {
            Assert-DocumentationPageSnapshot -SiteRootPath $SiteRootPath -Evidence $evidence
            return $evidence
        }
    } catch { }
    throw 'The checked-in PSWriteOffice documentation catalog does not match its pinned source/derived evidence.'
}

function Test-CompleteSourceDocumentation {
    param(
        [Parameter(Mandatory)][string] $RepositoryRoot,
        [Parameter(Mandatory)][string] $SourceCommit
    )

    # A clean sparse checkout can omit tracked pages without appearing dirty.
    if (-not (Test-Path -LiteralPath (Join-Path $RepositoryRoot 'Build/Export-WebsiteDocumentationCatalog.ps1') -PathType Leaf)) { return $false }
    $trackedFiles = @(& git -C $RepositoryRoot ls-tree -r --name-only $SourceCommit -- 'Website/content/project-docs/docs')
    if ($LASTEXITCODE -ne 0) { return $false }
    $trackedPages = @($trackedFiles | Where-Object { $_ -cmatch '\.md$' })
    if ($trackedPages.Count -eq 0) { return $false }
    foreach ($page in $trackedPages) {
        if (-not (Test-Path -LiteralPath (Join-Path $RepositoryRoot $page) -PathType Leaf)) { return $false }
    }
    return $true
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
                schemaVersion = 2
                moduleName = 'PSWriteOffice'
                moduleVersion = [string] $ExpectedVersion
                sourceRef = $SourceRef
                sourceCommit = $SourceCommit
                normalization = 'utf8-no-bom-lf'
                producer = [ordered]@{ path = $producerPath; fingerprint = Get-DocumentationCatalogFingerprint -Path $producer }
                sourceCatalog = Get-DocumentationCatalogFingerprint -Path $temporaryCatalog
                derivedCatalog = $null
                documentation = @(New-DocumentationPageEvidence -RepositoryRoot $RepositoryRoot)
                navigation = $null
            }
        }
    } finally { Remove-Item -LiteralPath $temporaryCatalog }
}

function Save-DocumentationCatalogEvidence {
    param(
        [Parameter(Mandatory)][string] $CatalogPath,
        [Parameter(Mandatory)][string] $EvidencePath,
        [Parameter(Mandatory)][string] $SiteRootPath,
        [Parameter(Mandatory)] $Evidence
    )

    # Consumer link rewriting is recorded separately from the source output.
    $Evidence.derivedCatalog = Get-DocumentationCatalogFingerprint -Path $CatalogPath
    foreach ($page in $Evidence.documentation) {
        $page.derived = Get-DocumentationCatalogFingerprint -Path (Join-Path $SiteRootPath $page.sitePath)
    }
    $Evidence.navigation = Get-PSWriteOfficeNavigationFingerprint -SiteRootPath $SiteRootPath
    [IO.File]::WriteAllText($EvidencePath, ($Evidence | ConvertTo-Json -Depth 8) + "`n", [Text.UTF8Encoding]::new($false))
}
