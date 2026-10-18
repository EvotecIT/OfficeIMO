BeforeAll {
    $exportScript = Join-Path $PSScriptRoot '../Build/Export-WebsiteCommandMetadata.ps1'
    $sourceRoot = Join-Path $TestDrive 'module source #1'
    $cmdletRoot = Join-Path $sourceRoot 'Sources/PSWriteOffice/Cmdlets'
    $null = New-Item -ItemType Directory -Path $cmdletRoot
    $fixtureSource = @'
using System.Management.Automation;
namespace WebsiteMetadataFixture {
    [Cmdlet(VerbsCommon.Get, "ExportMetadata")]
    [Alias("Read-ExportMetadata")]
    public sealed class GetExportMetadata : PSCmdlet { }
}
'@
    $fixtureSource | Set-Content -LiteralPath (Join-Path $cmdletRoot 'GetExportMetadata.cs')
    Add-Type -TypeDefinition $fixtureSource -OutputAssembly (Join-Path $sourceRoot 'MetadataFixture.dll')
    $fixtureManifest = Join-Path $sourceRoot 'PSWriteOffice.psd1'
    New-ModuleManifest -Path $fixtureManifest -ModuleVersion '1.0.0' -RootModule 'MetadataFixture.dll' `
        -CmdletsToExport 'Get-ExportMetadata' -FunctionsToExport @() -AliasesToExport @()
    $metadataPath = Join-Path $TestDrive 'metadata.json'
}

Describe 'Website command metadata export' {
    It 'exports relative source links and declared aliases from an absolute filesystem path' {
        & (Get-Process -Id $PID).Path -NoProfile -NonInteractive -File $exportScript `
            -RepositoryRoot $sourceRoot -ManifestPath $fixtureManifest -OutputPath $metadataPath
        $LASTEXITCODE | Should -Be 0
        $metadata = @(Get-Content -LiteralPath $metadataPath -Raw | ConvertFrom-Json)
        $metadata.Count | Should -Be 1
        $metadata[0].name | Should -Be 'Get-ExportMetadata'
        $metadata[0].sourcePath | Should -Be 'Sources/PSWriteOffice/Cmdlets/GetExportMetadata.cs'
        $metadata[0].sourceLine | Should -Be 5
        $metadata[0].aliases | Should -Contain 'Read-ExportMetadata'
        $metadata[0].sourceUrl | Should -Be 'https://github.com/EvotecIT/OfficeIMO/blob/master/PowerShell/PSWriteOffice/Sources/PSWriteOffice/Cmdlets/GetExportMetadata.cs#L5'
    }
}
