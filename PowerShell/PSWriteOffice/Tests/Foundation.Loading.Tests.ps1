BeforeAll {
    $moduleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) { $env:PSWRITEOFFICE_MODULE_MANIFEST } else { Join-Path $PSScriptRoot '../PSWriteOffice.psd1' }
    $conflictPath = $env:PSWRITEOFFICE_CONFLICT_OPENXML_PATH
    if (!$conflictPath -and $PSEdition -eq 'Core') {
        $cache = ((& dotnet nuget locals global-packages --list) -replace '^global-packages:\s*','').Trim()
        $conflictPath = Join-Path $cache 'documentformat.openxml/3.1.1/lib/net8.0/DocumentFormat.OpenXml.dll'
    }
}
Describe 'Foundation assembly isolation and module lifetime' {
    It 'isolates a real conflicting Open XML version, exposes OfficeIMO type literals and removes its lookup handler' {
        if ($PSEdition -ne 'Core') { Set-ItResult -Skipped -Because 'ALC qualification requires PowerShell 7'; return }
        Test-Path -LiteralPath $conflictPath | Should -BeTrue -Because 'the test project downloads the isolated Open XML 3.1.1 fixture'
        $output = & pwsh -NoProfile -File (Join-Path $PSScriptRoot 'Foundation.Loading.Smoke.ps1') -ModulePath $moduleManifest -ConflictPath $conflictPath 2>&1
        $LASTEXITCODE | Should -Be 0 -Because ($output -join [Environment]::NewLine)
        $result = ($output | Where-Object { $_ -is [string] -and $_.StartsWith('{') } | Select-Object -Last 1) | ConvertFrom-Json
        $result.WordType | Should -Be 'OfficeIMO.Word.WordDocument'
        $result.WordIsolated | Should -BeTrue
        $result.ModuleIsDefault | Should -BeFalse
        $result.ConflictIsDefault | Should -BeTrue
        $result.ModuleOpenXmlVersion | Should -Not -Be $result.ConflictVersion
        ($result.Hint -join ' ') | Should -Match 'ExcelNew.*New-OfficeExcel'
        $result.PriorWasCalled | Should -BeTrue
        $result.HandlerRestored | Should -BeTrue
    }
    It 'registers and removes hints when the loader does not supply its Core-only assembly variable' {
        if ($PSEdition -ne 'Core') { Set-ItResult -Skipped -Because 'This fixture uses the net8 test payload'; return }
        $output = & pwsh -NoProfile -File (Join-Path $PSScriptRoot 'Foundation.Loading.Smoke.ps1') -ModulePath $moduleManifest -ConflictPath $conflictPath -WithoutLoaderAssembly 2>&1
        $LASTEXITCODE | Should -Be 0 -Because ($output -join [Environment]::NewLine)
        $result = ($output | Where-Object { $_ -is [string] -and $_.StartsWith('{') } | Select-Object -Last 1) | ConvertFrom-Json
        ($result.Hint -join ' ') | Should -Match 'ExcelNew.*New-OfficeExcel'
        $result.HandlerRestored | Should -BeTrue
    }

}
