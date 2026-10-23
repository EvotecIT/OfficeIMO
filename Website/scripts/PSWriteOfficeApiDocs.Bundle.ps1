function Get-ManifestCommandNames {
    param([Parameter(Mandatory)][string] $Path)

    $manifest = Import-PowerShellDataFile -LiteralPath $Path
    @($manifest.FunctionsToExport) + @($manifest.CmdletsToExport) |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and $_ -ne '*' } |
        Sort-Object -Unique
}

function Get-HelpCommandNames {
    param([Parameter(Mandatory)][string] $Path)

    [xml] $help = Get-Content -LiteralPath $Path -Raw
    @($help.helpItems.command) |
        ForEach-Object { [string] $_.details.name } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Sort-Object -Unique
}

function Get-MetadataCommandNames {
    param([Parameter(Mandatory)][string] $Path)

    $metadata = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
    @($metadata.commands) |
        ForEach-Object { [string] $_.name } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Sort-Object -Unique
}

function Test-ApiBundle {
    param(
        [Parameter(Mandatory)][string] $ManifestPath,
        [Parameter(Mandatory)][string] $HelpPath,
        [Parameter(Mandatory)][string] $MetadataPath,
        [Parameter(Mandatory)][string[]] $ExpectedCommands,
        [Parameter(Mandatory)][version] $ExpectedVersion,
        [Parameter(Mandatory)][string] $ExpectedSourceCommit
    )

    $bundleManifestPath = Join-Path (Split-Path -Parent $ManifestPath) 'bundle-manifest.json'
    $paths = @($ManifestPath, $HelpPath, $MetadataPath, $bundleManifestPath)
    $missing = @($paths | Where-Object { -not (Test-Path -LiteralPath $_ -PathType Leaf) })
    if ($missing.Count -gt 0) {
        return [PSCustomObject]@{ Valid = $false; Reason = "missing: $($missing -join ', ')"; CommandCount = 0 }
    }

    try {
        $provenance = Get-Content -LiteralPath $bundleManifestPath -Raw | ConvertFrom-Json
        if ($provenance.schemaVersion -ne 1 -or $provenance.moduleName -cne 'PSWriteOffice' -or
            [version]$provenance.moduleVersion -ne $ExpectedVersion -or $provenance.sourceCommit -cne $ExpectedSourceCommit) {
            return [PSCustomObject]@{ Valid = $false; Reason = 'bundle provenance does not match the pinned release'; CommandCount = 0 }
        }
        foreach ($component in @(
            @{ Key = 'manifest'; Path = $ManifestPath },
            @{ Key = 'help'; Path = $HelpPath },
            @{ Key = 'metadata'; Path = $MetadataPath }
        )) {
            $evidence = $provenance.components.($component.Key)
            if ($null -eq $evidence -or (Get-Item -LiteralPath $component.Path).Length -ne $evidence.bytes -or
                (Get-FileHash -LiteralPath $component.Path -Algorithm SHA256).Hash -ine $evidence.sha256) {
                return [PSCustomObject]@{ Valid = $false; Reason = "bundle $($component.Key) does not match its release fingerprint"; CommandCount = 0 }
            }
        }
        $bundleManifest = Import-PowerShellDataFile -LiteralPath $ManifestPath
        if ([version]$bundleManifest.ModuleVersion -ne $ExpectedVersion) {
            return [PSCustomObject]@{ Valid = $false; Reason = "bundle version $($bundleManifest.ModuleVersion) does not match source version $ExpectedVersion"; CommandCount = 0 }
        }
        $sets = [ordered]@{
            manifest = @(Get-ManifestCommandNames -Path $ManifestPath)
            help = @(Get-HelpCommandNames -Path $HelpPath)
            metadata = @(Get-MetadataCommandNames -Path $MetadataPath)
        }
    } catch {
        return [PSCustomObject]@{ Valid = $false; Reason = "could not parse bundle: $($_.Exception.Message)"; CommandCount = 0 }
    }

    foreach ($entry in $sets.GetEnumerator()) {
        $difference = @(Compare-Object -ReferenceObject $ExpectedCommands -DifferenceObject $entry.Value)
        if ($difference.Count -gt 0) {
            return [PSCustomObject]@{
                Valid = $false
                Reason = "$($entry.Key) covers $($entry.Value.Count) of $($ExpectedCommands.Count) authoritative commands"
                CommandCount = $entry.Value.Count
            }
        }
    }

    [PSCustomObject]@{ Valid = $true; Reason = 'complete'; CommandCount = $ExpectedCommands.Count }
}
