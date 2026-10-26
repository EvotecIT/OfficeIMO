BeforeAll {
    $moduleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) { $env:PSWRITEOFFICE_MODULE_MANIFEST } else { Join-Path $PSScriptRoot '../PSWriteOffice.psd1' }
    $conflictPath = $env:PSWRITEOFFICE_CONFLICT_OPENXML_PATH
    if (!$conflictPath) {
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
        ($result.Hint -join ' ') | Should -BeLike '*ExcelNew was replaced by New-OfficeExcel*'
        $result.PriorWasCalled | Should -BeTrue
        $result.HandlerRestored | Should -BeTrue
    }
}
