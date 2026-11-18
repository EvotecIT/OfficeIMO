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
    $helpRoot = New-Item -ItemType Directory -Path (Join-Path $sourceRoot 'en-US')
    @'
<?xml version="1.0" encoding="utf-8"?>
<helpItems schema="maml" xmlns="http://msh">
  <command:command xmlns:maml="http://schemas.microsoft.com/maml/2004/10" xmlns:command="http://schemas.microsoft.com/maml/dev/command/2004/10">
    <command:details>
      <command:name>Get-ExportMetadata</command:name>
      <maml:description>
        <maml:para>Reads export metadata</maml:para>
        <maml:para>for the website.</maml:para>
      </maml:description>
    </command:details>
  </command:command>
</helpItems>
'@ | Set-Content -LiteralPath (Join-Path $helpRoot 'PSWriteOffice-help.xml') -Encoding utf8
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
        $metadata[0].synopsis | Should -Be 'Reads export metadata for the website.'
        $metadata[0].sourceUrl | Should -Be 'https://github.com/EvotecIT/OfficeIMO/blob/master/PowerShell/PSWriteOffice/Sources/PSWriteOffice/Cmdlets/GetExportMetadata.cs#L5'
    }

    It 'fails instead of exporting commands without a synopsis when the module has no help' {
        $withoutHelp = Join-Path $TestDrive 'module without help'
        $null = New-Item -ItemType Directory -Path $withoutHelp
        Copy-Item -LiteralPath $fixtureManifest, (Join-Path $sourceRoot 'MetadataFixture.dll') -Destination $withoutHelp
        $outputPath = Join-Path $TestDrive 'metadata-without-help.json'
        # Windows PowerShell 5.1 turns the child's error output into a terminating error under Stop; the exit code is what this test checks.
        $ErrorActionPreference = 'Continue'
        & (Get-Process -Id $PID).Path -NoProfile -NonInteractive -File $exportScript `
            -RepositoryRoot $sourceRoot -ManifestPath (Join-Path $withoutHelp 'PSWriteOffice.psd1') -OutputPath $outputPath 2>$null
        $LASTEXITCODE | Should -Not -Be 0
        Test-Path -LiteralPath $outputPath | Should -BeFalse
    }
}
