BeforeAll {
    $manifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) { $env:PSWRITEOFFICE_MODULE_MANIFEST } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $manifest -Global -ErrorAction Stop
}

Describe 'Cross-format document search' {
    It 'searches multiple regular expressions and emits flat matches for pipeline documents' {
        $path = Join-Path $TestDrive 'patterns.md'
        Set-Content -LiteralPath $path -Value 'Invoice INV-42 and INV-730. Total USD 1250.50.' -Encoding UTF8
        $matches = @(Read-OfficeDocument -Path $path | Search-OfficeDocument -Pattern 'INV-\d+', 'USD \d+\.\d+')
        $matches | Should -HaveCount 3
        $matches.Match | Should -Contain 'INV-42'
        $matches.Match | Should -Contain 'INV-730'
        $matches.Match | Should -Contain 'USD 1250.50'
        $matches.Path | Select-Object -Unique | Should -Be $path
    }

    It 'keeps literal metacharacters and case sensitivity distinct from regular expressions' {
        $path = Join-Path $TestDrive 'literal.md'
        Set-Content -LiteralPath $path -Value 'a.b axb A.B' -Encoding UTF8
        $literal = @(Search-OfficeDocument -Path $path -Pattern 'a.b' -SimpleMatch -CaseSensitive)
        $literal | Should -HaveCount 1
        $literal[0].Match | Should -Be 'a.b'
        @(Search-OfficeDocument -Path $path -Pattern 'a.b').Count | Should -Be 3
    }

    It 'resolves literal bracketed filenames through FileInfo and reports unreadable inputs independently' {
        $path = Join-Path $TestDrive 'report[1].md'
        Set-Content -LiteralPath $path -Value 'INV-42' -Encoding UTF8
        $matches = @(Get-Item -LiteralPath $path | Search-OfficeDocument -Pattern 'INV-\d+')
        $matches | Should -HaveCount 1
        $matches[0].Path | Should -Be $path
        $failures = @()
        $mixed = @(Search-OfficeDocument -LiteralPath @((Join-Path $TestDrive 'missing.md'), $path) `
            -Pattern 'INV-\d+' -ErrorAction SilentlyContinue -ErrorVariable +failures)
        $mixed | Should -HaveCount 1
        $failures | Should -HaveCount 1
        $failures[0].TargetObject | Should -Be (Join-Path $TestDrive 'missing.md')
    }

    It 'shares a per-document result ceiling across patterns and accepts zero for all results' {
        $path = Join-Path $TestDrive 'limits.md'
        Set-Content -LiteralPath $path -Value 'INV-1 INV-2 USD-3' -Encoding UTF8
        $warnings = @()
        $limited = @(Search-OfficeDocument -Path $path -Pattern 'INV-\d+', 'USD-\d+' -MaximumResults 2 `
            -WarningAction SilentlyContinue -WarningVariable +warnings)
        $limited | Should -HaveCount 2
        $limited.SearchLimitReached | Should -Not -Contain $false
        $warnings | Should -Not -BeNullOrEmpty
        @(Search-OfficeDocument -Path $path -Pattern 'INV-\d+', 'USD-\d+' -MaximumResults 0).Count | Should -Be 3
    }

    It 'rejects invalid patterns before attempting input discovery' {
        { Search-OfficeDocument -LiteralPath (Join-Path $TestDrive 'missing.md') -Pattern '[' -ErrorAction Stop } |
            Should -Throw -ExpectedMessage '*Invalid regular expression*'
    }
}
