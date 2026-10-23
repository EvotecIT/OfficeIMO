function Test-PSWriteOfficeSnapshotInputs {
    <# Checks the producer revision against all module/engine/build inputs.
       A shallow checkout may fetch the recorded commit's tree; unavailable
       provenance or changed/uncommitted inputs require the normal fresh export. #>
    param(
        [Parameter(Mandatory)][string] $RepositoryRoot,
        [Parameter(Mandatory)][string] $SourceCommit,
        [Parameter(Mandatory)][string] $SnapshotCommit
    )

    if ($SnapshotCommit -cnotmatch '^[0-9a-f]{40}$') { return $false }
    $top = (& git -C $RepositoryRoot rev-parse --show-toplevel 2>$null).Trim()
    if ($LASTEXITCODE -ne 0) { return $false }
    & git -C $top cat-file -e "${SnapshotCommit}^{commit}" 2>$null
    if ($LASTEXITCODE -ne 0) {
        # Fetch only trees, not another checkout or a full repository history.
        $start = [Diagnostics.ProcessStartInfo]::new('git')
        $start.WorkingDirectory = $top
        $start.UseShellExecute = $false
        $start.RedirectStandardOutput = $true
        $start.RedirectStandardError = $true
        foreach ($argument in @('fetch', '--no-tags', '--depth=1', '--filter=blob:none', 'origin', $SnapshotCommit)) {
            $start.ArgumentList.Add($argument)
        }
        if ($env:GITHUB_TOKEN) {
            $authorization = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('x-access-token:' + $env:GITHUB_TOKEN))
            $start.Environment['GIT_CONFIG_COUNT'] = '1'
            $start.Environment['GIT_CONFIG_KEY_0'] = 'http.https://github.com/.extraheader'
            $start.Environment['GIT_CONFIG_VALUE_0'] = 'AUTHORIZATION: basic ' + $authorization
        }
        $process = [Diagnostics.Process]::new()
        $process.StartInfo = $start
        try {
            $null = $process.Start()
            $stdout = $process.StandardOutput.ReadToEndAsync()
            $stderr = $process.StandardError.ReadToEndAsync()
            if (-not $process.WaitForExit(60000)) {
                $process.Kill($true)
                $process.WaitForExit()
                return $false
            }
            $null = $stdout.GetAwaiter().GetResult()
            $null = $stderr.GetAwaiter().GetResult()
            if ($process.ExitCode -ne 0) { return $false }
        } finally { $process.Dispose() }
    }

    # Match the producer's engine/build boundary. Generated exports are outputs,
    # not inputs; including them would invalidate their own producer revision.
    # Git magic keeps wildcard pathspecs intact through native Unix invocation.
    $inputs = @('PowerShell/PSWriteOffice', ':(top)OfficeIMO.*', ':(top)Directory.Build.*',
        ':(top)Directory.Packages.*', 'global.json', 'NuGet.Config', 'nuget.config', 'Build',
        'Website/scripts/Sync-PSWriteOfficeApiDocs.ps1',
        ':(top)Website/scripts/PSWriteOfficeApiDocs.*.ps1',
        ':(exclude)PowerShell/PSWriteOffice/WebsiteArtifacts/**',
        ':(exclude)PowerShell/PSWriteOffice/Docs/Generated/**',
        ':(exclude)PowerShell/PSWriteOffice/Lib/**',
        ':(exclude)PowerShell/PSWriteOffice/Artefacts/**')
    foreach ($arguments in @(
        @('diff', '--quiet', $SnapshotCommit, $SourceCommit, '--'),
        @('diff', '--quiet', 'HEAD', '--'),
        @('diff', '--cached', '--quiet', 'HEAD', '--')
    )) {
        & git -C $top @arguments @inputs
        if ($LASTEXITCODE -ne 0) { return $false }
    }
    $untracked = @(& git -C $top ls-files --others --exclude-standard -- @inputs)
    return $LASTEXITCODE -eq 0 -and $untracked.Count -eq 0
}

function Test-CurrentPSWriteOfficeSnapshot {
    <# A cache hit still validates the API component hashes, command inventory,
       catalog, rendered documentation and navigation. Damaged outputs rebuild. #>
    param(
        [Parameter(Mandatory)][string] $RepositoryRoot,
        [Parameter(Mandatory)][string] $SiteRoot,
        [Parameter(Mandatory)][string] $SourceCommit,
        [switch] $SkipDocumentation
    )
    try {
        $api = Join-Path $SiteRoot 'data/apidocs/powershell'
        $bundle = Get-Content -LiteralPath (Join-Path $api 'bundle-manifest.json') -Raw | ConvertFrom-Json
        if (-not (Test-PSWriteOfficeSnapshotInputs -RepositoryRoot $RepositoryRoot -SourceCommit $SourceCommit -SnapshotCommit $bundle.sourceCommit)) {
            return $false
        }
        $manifest = Import-PowerShellDataFile -LiteralPath (Join-Path $RepositoryRoot 'PSWriteOffice.psd1')
        $commands = @(Get-ManifestCommandNames -Path (Join-Path $RepositoryRoot 'PSWriteOffice.psd1'))
        $valid = Test-ApiBundle -ManifestPath (Join-Path $api 'PSWriteOffice.psd1') `
            -HelpPath (Join-Path $api 'PSWriteOffice-Help.xml') -MetadataPath (Join-Path $api 'command-metadata.json') `
            -ExpectedCommands $commands -ExpectedVersion ([version]$manifest.ModuleVersion) -ExpectedSourceCommit $bundle.sourceCommit
        if (-not $valid.Valid) { return $false }
        if (-not $SkipDocumentation) {
            $null = Assert-DocumentationCatalogSnapshot -CatalogPath (Join-Path $SiteRoot 'data/pswriteoffice_command_catalog.json') `
                -EvidencePath (Join-Path $api 'documentation-manifest.json') -SiteRootPath $SiteRoot `
                -ExpectedSourceRef $bundle.sourceCommit -ExpectedVersion ([version]$manifest.ModuleVersion) -ExpectedCommandCount $commands.Count
        }
        return $true
    } catch {
        Write-Verbose "The PSWriteOffice snapshot requires a fresh export: $($_.Exception.Message)"
        return $false
    }
}
