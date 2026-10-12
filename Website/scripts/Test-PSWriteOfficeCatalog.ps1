param(
    [string] $SiteRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'PSWriteOfficeApiDocs.Catalog.ps1')
$apiRoot = Join-Path $SiteRoot 'data/apidocs/powershell'
$catalogPath = Join-Path $SiteRoot 'data/pswriteoffice_command_catalog.json'
$evidencePath = Join-Path $apiRoot 'documentation-manifest.json'
$manifest = Import-PowerShellDataFile -LiteralPath (Join-Path $apiRoot 'PSWriteOffice.psd1')
$configuration = Get-Content -LiteralPath (Join-Path $SiteRoot 'site.json') -Raw | ConvertFrom-Json
$source = @($configuration.Sources | Where-Object Slug -CEQ 'pswriteoffice')
if ($source.Count -ne 1) { throw 'Exactly one PSWriteOffice source is required.' }
$expected = @{
    ExpectedSourceRef = $source[0].Ref
    ExpectedVersion = [version] $manifest.ModuleVersion
    ExpectedCommandCount = @($manifest.CmdletsToExport | Sort-Object -Unique).Count
}
$null = Assert-DocumentationCatalogSnapshot -CatalogPath $catalogPath -EvidencePath $evidencePath @expected

$temporaryRoot = Join-Path ([IO.Path]::GetTempPath()) ('officeimo-catalog-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temporaryRoot | Out-Null
$temporaryCatalog = Join-Path $temporaryRoot 'catalog.json'
$temporaryEvidence = Join-Path $temporaryRoot 'documentation-manifest.json'
$cases = [Collections.Generic.List[string]]::new()
try {
    Copy-Item -LiteralPath $evidencePath -Destination $temporaryEvidence
    # The original regression preserved the version and both aggregate totals.
    foreach ($case in 'description', 'family-allocation', 'url', 'featured-command') {
        $catalog = Get-Content -LiteralPath $catalogPath -Raw | ConvertFrom-Json -Depth 20
        switch ($case) {
            'description' { $catalog.families[0].description = 'Stale description' }
            'family-allocation' { $catalog.families[0].commandCount--; $catalog.families[1].commandCount++ }
            'url' { $catalog.families[0].examplesUrl = 'https://example.invalid/stale' }
            'featured-command' { $catalog.families[0].featuredCommands = @('Get-StaleCommand') }
        }
        $catalog | ConvertTo-Json -Depth 20 -Compress | Set-Content -LiteralPath $temporaryCatalog -Encoding utf8
        if (-not (Test-DocumentationCatalog -Path $temporaryCatalog -ExpectedVersion $expected.ExpectedVersion -ExpectedCommandCount $expected.ExpectedCommandCount)) {
            throw "The '$case' fixture must retain the formerly accepted aggregate totals."
        }
        $rejected = $false
        try { $null = Assert-DocumentationCatalogSnapshot -CatalogPath $temporaryCatalog -EvidencePath $temporaryEvidence @expected }
        catch {
            if ($_.Exception.Message -notlike '*does not match its pinned source/derived evidence*') { throw }
            $rejected = $true
        }
        if (-not $rejected) { throw "An altered '$case' catalog was accepted." }
        $cases.Add($case)
    }

    # The same checked-in content remains valid in Windows and Linux checkouts.
    $content = [IO.File]::ReadAllText($catalogPath).Replace("`r`n", "`n")
    [IO.File]::WriteAllText($temporaryCatalog, $content.Replace("`n", "`r`n"), [Text.UTF8Encoding]::new($true))
    $null = Assert-DocumentationCatalogSnapshot -CatalogPath $temporaryCatalog -EvidencePath $temporaryEvidence @expected
    $cases.Add('bom-crlf-checkout')

    foreach ($case in 'wrong-source', 'missing-source-fingerprint') {
        $evidence = Get-Content -LiteralPath $evidencePath -Raw | ConvertFrom-Json -Depth 10
        if ($case -eq 'wrong-source') { $evidence.sourceCommit = '0000000000000000000000000000000000000000' }
        else { $evidence.sourceCatalog = $null }
        $evidence | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $temporaryEvidence
        $rejected = $false
        try { $null = Assert-DocumentationCatalogSnapshot -CatalogPath $temporaryCatalog -EvidencePath $temporaryEvidence @expected }
        catch { $rejected = $true }
        if (-not $rejected) { throw "Invalid '$case' evidence was accepted." }
        $cases.Add($case)
    }

    # Completeness follows the pinned Git inventory, including sparse checkouts.
    $sourceFixture = Join-Path $temporaryRoot 'source'
    $sourceDocs = Join-Path $sourceFixture 'Website/content/project-docs/docs'
    New-Item -ItemType Directory -Path $sourceDocs -Force | Out-Null
    foreach ($page in '_index.md', 'word.md') {
        Set-Content -LiteralPath (Join-Path $sourceDocs $page) -Value '# Documentation fixture'
    }
    & git -C $sourceFixture init --quiet
    & git -C $sourceFixture add -- 'Website/content/project-docs/docs'
    & git -C $sourceFixture -c user.name='Documentation test' -c user.email='docs@example.invalid' commit --quiet -m 'Documentation fixture'
    if ($LASTEXITCODE -ne 0) { throw 'Could not commit the documentation fixture.' }
    $sourceCommit = & git -C $sourceFixture rev-parse HEAD
    if (-not (Test-CompleteSourceDocumentation -RepositoryRoot $sourceFixture -SourceCommit $sourceCommit)) {
        throw 'The complete tracked documentation fixture was rejected.'
    }
    & git -C $sourceFixture sparse-checkout set --no-cone '/Website/content/project-docs/docs/_index.md'
    if ($LASTEXITCODE -ne 0 -or @(git -C $sourceFixture status --porcelain).Count -ne 0) {
        throw 'The partial documentation fixture must be a clean sparse checkout.'
    }
    if (Test-CompleteSourceDocumentation -RepositoryRoot $sourceFixture -SourceCommit $sourceCommit) {
        throw 'An incomplete sparse documentation checkout was accepted.'
    }
    $cases.Add('tracked-documentation-completeness')
} finally {
    # Git creates hidden metadata and read-only loose objects on Windows.
    # This unique test directory contains only the fixture created above.
    foreach ($item in Get-ChildItem -LiteralPath $temporaryRoot -Force -Recurse) {
        $item.Attributes = $item.Attributes -band (-bnot ([IO.FileAttributes]::Hidden -bor [IO.FileAttributes]::ReadOnly))
    }
    Remove-Item -LiteralPath $temporaryRoot -Recurse
}

[PSCustomObject]@{ Status = 'passed'; CatalogCases = @($cases); CommandCount = $expected.ExpectedCommandCount }
