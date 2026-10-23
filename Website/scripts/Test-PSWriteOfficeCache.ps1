param()
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'PSWriteOfficeApiDocs.Bundle.ps1')
. (Join-Path $PSScriptRoot 'PSWriteOfficeApiDocs.Cache.ps1')
$temporary = Join-Path ([IO.Path]::GetTempPath()) ('officeimo-api-cache-' + [guid]::NewGuid().ToString('N'))
$cases = [Collections.Generic.List[string]]::new()
try {
    $module = Join-Path $temporary 'PowerShell/PSWriteOffice'
    $site = Join-Path $temporary 'Website'
    $api = Join-Path $site 'data/apidocs/powershell'
    $null = New-Item -ItemType Directory -Force -Path $module, $api, (Join-Path $temporary 'OfficeIMO.Word')
    $null = New-Item -ItemType Directory -Path (Join-Path $site 'scripts')
    Set-Content -LiteralPath (Join-Path $site 'scripts/PSWriteOfficeApiDocs.Catalog.ps1') -Value '# catalog generator'
    $manifest = "@{ ModuleVersion = '1.0.0'; CmdletsToExport = @('Get-OfficeDemo'); FunctionsToExport = @() }"
    Set-Content -LiteralPath (Join-Path $module 'PSWriteOffice.psd1') -Value $manifest
    Set-Content -LiteralPath (Join-Path $temporary 'OfficeIMO.Word/Document.cs') -Value 'class Document {}'
    & git -C $temporary init --quiet
    & git -C $temporary config user.name 'OfficeIMO cache fixture'
    & git -C $temporary config user.email 'fixture@example.invalid'
    & git -C $temporary add .
    & git -C $temporary commit --quiet -m 'Producer inputs'
    if ($LASTEXITCODE -ne 0) { throw 'Fixture commit failed.' }
    $producer = (& git -C $temporary rev-parse HEAD).Trim()
    Set-Content -LiteralPath (Join-Path $site 'style.css') -Value 'body {}'
    & git -C $temporary add .
    & git -C $temporary commit --quiet -m 'Unrelated website edit'
    $head = (& git -C $temporary rev-parse HEAD).Trim()
    if (-not (Test-PSWriteOfficeSnapshotInputs -RepositoryRoot $module -SourceCommit $head -SnapshotCommit $producer)) {
        throw 'An unchanged module/engine graph was invalidated by website content.'
    }
    $cases.Add('website-edit-reuses-identical-inputs')

    foreach ($path in @('OfficeIMO.Word/Document.cs', 'PowerShell/PSWriteOffice/PSWriteOffice.psd1', 'Website/scripts/PSWriteOfficeApiDocs.Catalog.ps1')) {
        $input = Join-Path $temporary $path
        $original = [IO.File]::ReadAllBytes($input)
        Add-Content -LiteralPath $input -Value '# changed input'
        if (Test-PSWriteOfficeSnapshotInputs -RepositoryRoot $module -SourceCommit $head -SnapshotCommit $producer) {
            throw "Changed input accepted: $path"
        }
        [IO.File]::WriteAllBytes($input, $original)
        $cases.Add('dirty-' + $path)
    }
    $newSource = Join-Path $temporary 'OfficeIMO.Word/New.cs'
    Set-Content -LiteralPath $newSource -Value 'class New {}'
    if (Test-PSWriteOfficeSnapshotInputs -RepositoryRoot $module -SourceCommit $head -SnapshotCommit $producer) {
        throw 'An untracked engine input was accepted.'
    }
    Remove-Item -LiteralPath $newSource
    $cases.Add('untracked-engine-input-rebuilds')

    Copy-Item -LiteralPath (Join-Path $module 'PSWriteOffice.psd1') -Destination $api
    [IO.File]::WriteAllText((Join-Path $api 'PSWriteOffice-Help.xml'), '<helpItems><command><details><name>Get-OfficeDemo</name></details></command></helpItems>')
    [IO.File]::WriteAllText((Join-Path $api 'command-metadata.json'), '{"commands":[{"name":"Get-OfficeDemo"}]}')
    $components = [ordered]@{}
    foreach ($component in @(@{Key='manifest';Name='PSWriteOffice.psd1'}, @{Key='help';Name='PSWriteOffice-Help.xml'}, @{Key='metadata';Name='command-metadata.json'})) {
        $path = Join-Path $api $component.Name
        $components[$component.Key] = @{bytes=(Get-Item -LiteralPath $path).Length;sha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash}
    }
    @{schemaVersion=1;moduleName='PSWriteOffice';moduleVersion='1.0.0';sourceCommit=$producer;components=$components} |
        ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $api 'bundle-manifest.json')
    if (-not (Test-CurrentPSWriteOfficeSnapshot -RepositoryRoot $module -SiteRoot $site -SourceCommit $head -SkipDocumentation)) {
        throw 'A complete, unchanged API snapshot was rejected.'
    }
    $cases.Add('verified-bundle-reused')
    # The cached route must preserve checked-in examples when a source checkout
    # has no Examples directory, just like the fresh-export route.
    $examples = Join-Path $api 'examples'
    $null = New-Item -ItemType Directory -Path $examples
    $fallback = Join-Path $examples 'fallback.ps1'
    Set-Content -LiteralPath $fallback -Value '# retained example'
    $exampleHash = (Get-FileHash -LiteralPath $fallback).Hash
    $reused = & (Join-Path $PSScriptRoot 'Sync-PSWriteOfficeApiDocs.ps1') -SiteRoot $site -PSWriteOfficeRoot $module -SkipDocumentation
    if ($reused.apiSnapshotSource -ne 'verified-unchanged-inputs' -or $reused.examplesUpdated -or
        (Get-FileHash -LiteralPath $fallback).Hash -cne $exampleHash) {
        throw 'Snapshot reuse removed fallback examples or claimed a missing source was copied.'
    }
    $cases.Add('missing-examples-preserves-fallback')
    $exportRoot = Join-Path $temporary 'exports'
    $exportApi = Join-Path $exportRoot 'apidocs/powershell'
    $null = New-Item -ItemType Directory -Path $exportApi
    foreach ($name in @('PSWriteOffice.psd1', 'command-metadata.json', 'bundle-manifest.json')) {
        Copy-Item -LiteralPath (Join-Path $api $name) -Destination $exportApi
    }
    Copy-Item -LiteralPath (Join-Path $api 'PSWriteOffice-Help.xml') -Destination (Join-Path $exportApi 'PSWriteOffice-help.xml')
    $bundlePath = Join-Path $exportApi 'bundle-manifest.json'
    $bundle = Get-Content -LiteralPath $bundlePath -Raw | ConvertFrom-Json
    $bundle.sourceCommit = $head
    $bundle | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $bundlePath
    foreach ($skip in @($false, $true)) {
        $imported = & (Join-Path $PSScriptRoot 'Sync-PSWriteOfficeApiDocs.ps1') -SiteRoot $site -PSWriteOfficeRoot $module -ExportedArtifactsRoot $exportRoot -SkipDocumentation -SkipExamples:$skip
        if ($imported.apiSnapshotSource -ne 'current-source' -or $imported.examplesUpdated -or
            (Get-FileHash -LiteralPath $fallback).Hash -cne $exampleHash) {
            throw 'API bundle import removed fallback examples.'
        }
    }
    $cases.Add('bundle-import-preserves-fallback-and-skipped-examples')
    $bundleExamples = Join-Path $exportApi 'examples'
    $null = New-Item -ItemType Directory -Path $bundleExamples
    Set-Content -LiteralPath (Join-Path $bundleExamples 'replacement.ps1') -Value '# current bundle example'
    $skipped = & (Join-Path $PSScriptRoot 'Sync-PSWriteOfficeApiDocs.ps1') -SiteRoot $site -PSWriteOfficeRoot $module -ExportedArtifactsRoot $exportRoot -SkipDocumentation -SkipExamples
    if ($skipped.examplesUpdated -or (Get-FileHash -LiteralPath $fallback).Hash -cne $exampleHash) {
        throw 'SkipExamples did not preserve examples during API import.'
    }
    $replaced = & (Join-Path $PSScriptRoot 'Sync-PSWriteOfficeApiDocs.ps1') -SiteRoot $site -PSWriteOfficeRoot $module -ExportedArtifactsRoot $exportRoot -SkipDocumentation
    if (-not $replaced.examplesUpdated -or (Test-Path -LiteralPath $fallback) -or
        -not (Test-Path -LiteralPath (Join-Path $examples 'replacement.ps1'))) {
        throw 'A current bundle example did not replace the fallback.'
    }
    $cases.Add('bundle-examples-replace-fallback-only-when-requested')
    Add-Content -LiteralPath (Join-Path $api 'PSWriteOffice-Help.xml') -Value '<stale/>'
    if (Test-CurrentPSWriteOfficeSnapshot -RepositoryRoot $module -SiteRoot $site -SourceCommit $head -SkipDocumentation) {
        throw 'A damaged cached help component was accepted.'
    }
    $cases.Add('damaged-help-rebuilds')
    if (Test-PSWriteOfficeSnapshotInputs -RepositoryRoot $module -SourceCommit $head -SnapshotCommit 'not-a-commit') {
        throw 'Invalid snapshot provenance was accepted.'
    }
    $cases.Add('invalid-provenance-rebuilds')

    # A committed engine change also invalidates a formerly valid producer.
    Add-Content -LiteralPath (Join-Path $temporary 'OfficeIMO.Word/Document.cs') -Value '// changed'
    & git -C $temporary add OfficeIMO.Word/Document.cs
    & git -C $temporary commit --quiet -m 'Engine change'
    $changed = (& git -C $temporary rev-parse HEAD).Trim()
    if (Test-PSWriteOfficeSnapshotInputs -RepositoryRoot $module -SourceCommit $changed -SnapshotCommit $producer) {
        throw 'Committed engine changes were accepted.'
    }
    $cases.Add('committed-engine-input-rebuilds')

    Set-Content -LiteralPath (Join-Path $temporary 'OfficeIMO.Word/Document.cs') -Value 'class Document {}'
    & git -C $temporary add OfficeIMO.Word/Document.cs
    & git -C $temporary commit --quiet -m 'Restore equivalent engine inputs'
    $equivalent = (& git -C $temporary rev-parse HEAD).Trim()
    $shallow = Join-Path $temporary 'shallow'
    & git clone --quiet --depth=1 ([Uri]::new($temporary).AbsoluteUri) $shallow
    if ($LASTEXITCODE -ne 0) { throw 'Shallow fixture clone failed.' }
    if (-not (Test-PSWriteOfficeSnapshotInputs -RepositoryRoot (Join-Path $shallow 'PowerShell/PSWriteOffice') `
            -SourceCommit $equivalent -SnapshotCommit $producer)) {
        throw 'A shallow checkout could not qualify identical producer inputs.'
    }
    $cases.Add('shallow-checkout-fetches-producer-tree')
} finally {
    foreach ($item in Get-ChildItem -LiteralPath $temporary -Force -Recurse) {
        $item.Attributes = $item.Attributes -band (-bnot ([IO.FileAttributes]::Hidden -bor [IO.FileAttributes]::ReadOnly))
    }
    # Git dot directories stay hidden by name on Unix; use the same fixture
    # teardown as Test-PSWriteOfficeCatalog.ps1 after clearing fixture attributes.
    [IO.Directory]::Delete($temporary, $true)
}
[pscustomobject]@{Status='passed';Cases=@($cases)}
