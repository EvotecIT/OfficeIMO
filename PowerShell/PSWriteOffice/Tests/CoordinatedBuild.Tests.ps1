Describe 'OfficeIMO and PSWriteOffice release configuration' -Skip:($PSVersionTable.PSVersion -lt [version]'7.4' -or -not (Get-Module -ListAvailable PSPublishModule | Where-Object Version -ge ([version]'3.0.170'))) {
    It 'uses one package-derived version and ordered publication without enabling signing in Build mode' {
        $repositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot '../../..')).Path
        $jsonPath = Join-Path $TestDrive 'coordinated.build.json'
        & (Join-Path $repositoryRoot 'Build/Build-Release.ps1') -RunMode Build -SignModule:$false -JsonOnly -JsonPath $jsonPath -PassThru | Out-Null
        $configuration = Get-Content -LiteralPath $jsonPath -Raw | ConvertFrom-Json
        $project = @($configuration.Segments | Where-Object Type -eq ProjectBuild)
        $release = @($configuration.Segments | Where-Object Type -eq Release)
        $gate = @($configuration.Segments | Where-Object Type -eq Gate)

        $project.Count | Should -Be 1
        $release.Count | Should -Be 1
        $project[0].Configuration.BuildBeforeModule | Should -BeTrue
        $project[0].Configuration.UseAsReleaseVersionSource | Should -BeTrue
        $project[0].Configuration.SignAssemblies | Should -BeFalse
        $project[0].Configuration.SignPackages | Should -BeFalse
        $release[0].Configuration.VersionSource | Should -Be ProjectBuild
        $release[0].Configuration.PrimaryProject | Should -Be OfficeIMO.Word
        $release[0].Configuration.SynchronizeModuleVersion | Should -BeTrue
        $release[0].Configuration.BuildOrder | Should -Be @('Packages', 'Module')
        $release[0].Configuration.PublishOrder | Should -Be @('NuGet', 'PowerShellGallery', 'GitHub')
        $gate[0].Configuration.Mode | Should -Be Build
    }

    It 'keeps documentation generation in the native non-publishing gate' {
        $repositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot '../../..')).Path
        $jsonPath = Join-Path $TestDrive 'documentation.build.json'
        & (Join-Path $repositoryRoot 'Build/Build-Module.ps1') -RunMode Documentation -IncludePackages:$false -JsonOnly -JsonPath $jsonPath -PassThru | Out-Null
        $configuration = Get-Content -LiteralPath $jsonPath -Raw | ConvertFrom-Json
        @($configuration.Segments | Where-Object Type -eq Gate)[0].Configuration.Mode | Should -Be Documentation
        @($configuration.Segments | Where-Object Type -eq Release).Count | Should -Be 0
        @($configuration.Segments | Where-Object Type -eq ProjectBuild).Count | Should -Be 0
    }
}
