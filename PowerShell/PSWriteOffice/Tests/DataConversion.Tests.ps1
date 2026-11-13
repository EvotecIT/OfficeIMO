BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        $sourceRoot = Join-Path (Join-Path (Join-Path $PSScriptRoot '..') 'Sources') 'PSWriteOffice'

        if (-not $env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES) {
            $env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES = 'true'
        }

        if (-not $env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION) {
            $releasePath = Join-Path (Join-Path $sourceRoot 'bin') 'Release'
            $env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION = if (Test-Path $releasePath) { 'Release' } else { 'Debug' }
        }

        Join-Path (Join-Path $PSScriptRoot '..') 'PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
}

Describe 'Structured-data routes in Convert-OfficeDocument' {
    It 'converts JSON to CSV and TOML by the output extension and reports what TOML cannot hold' {
        $json = Join-Path $TestDrive 'customers.json'
        Set-Content -LiteralPath $json -Value '[{"id":1,"name":"Ann","address":{"city":"Oslo"}},{"id":2,"name":"Bob","address":{"city":null}}]' -NoNewline

        $csv = Join-Path $TestDrive 'customers.csv'
        $report = Convert-OfficeDocument -Path $json -OutputPath $csv
        $report.Target | Should -Be 'CSV'
        Get-Content -LiteralPath $csv | Should -Be @('id,name,address.city', '1,Ann,Oslo', '2,Bob,')

        $toml = Join-Path $TestDrive 'customers.toml'
        $report = Convert-OfficeDocument -Path $json -OutputPath $toml -Options @{ RootName = 'customer' }
        $report.Diagnostics.Code | Should -Contain 'null-dropped'
        $report.Diagnostics.Code | Should -Contain 'root-wrapped'
        $data = Import-OfficeToml -Path $toml
        $data.customer.name | Should -Be @('Ann', 'Bob')
        $data.customer[0].address.city | Should -Be 'Oslo'
    }

    It 'selects the rows of a TOML array of tables and stops on loss with FailOnLoss' {
        $toml = Join-Path $TestDrive 'Cargo.lock'
        Set-Content -LiteralPath $toml -Value "version = 4`n`n[[package]]`nname = `"serde`"`nversion = `"1.0.210`"`n`n[[package]]`nname = `"toml`"`nversion = `"0.8.19`"`n" -NoNewline
        $source = Join-Path $TestDrive 'lock.toml'
        Copy-Item -LiteralPath $toml -Destination $source

        $csv = Join-Path $TestDrive 'packages.csv'
        Convert-OfficeDocument -Path $source -OutputPath $csv -Options @{ Select = 'package' } | Out-Null
        Get-Content -LiteralPath $csv | Should -Be @('name,version', 'serde,1.0.210', 'toml,0.8.19')

        $json = Join-Path $TestDrive 'nulls.json'
        Set-Content -LiteralPath $json -Value '{"a":null}' -NoNewline
        { Convert-OfficeDocument -Path $json -OutputPath (Join-Path $TestDrive 'nulls.toml') -FailOnLoss -ErrorAction Stop } | Should -Throw
        Test-Path -LiteralPath (Join-Path $TestDrive 'nulls.toml') | Should -BeFalse
    }
}

Describe 'TOML commands' {
    It 'reads tables, arrays of tables and typed values from text and files' {
        $text = @(
            'title = "Report"'
            'retries = 3'
            'ratio = 0.5'
            'enabled = true'
            'published = 2026-11-11T09:30:00Z'
            'day = 2026-11-11'
            'at = 07:15:00'
            ''
            '[owner]'
            'name = "Ops"'
            ''
            '[[server]]'
            'host = "alpha"'
            ''
            '[[server]]'
            'host = "beta"'
        ) -join "`n"

        $data = $text | ConvertFrom-OfficeToml
        $data.title | Should -Be 'Report'
        $data.retries | Should -BeOfType [long]
        $data.ratio | Should -BeOfType [double]
        $data.enabled | Should -BeTrue
        $data.published | Should -BeOfType [DateTimeOffset]
        $data.day | Should -BeOfType [DateTime]
        $data.at | Should -Be ([TimeSpan]'07:15:00')
        $data.owner.name | Should -Be 'Ops'
        $data.server.host | Should -Be @('alpha', 'beta')
        @($data.PSObject.Properties.Name) | Should -Be @('title', 'retries', 'ratio', 'enabled', 'published', 'day', 'at', 'owner', 'server')

        $path = Join-Path $TestDrive 'report.toml'
        Set-Content -LiteralPath $path -Value $text -NoNewline
        $table = Import-OfficeToml -Path $path -AsHashtable
        $table | Should -BeOfType [System.Collections.Specialized.OrderedDictionary]
        $table['server'][1]['host'] | Should -Be 'beta'
    }

    It 'reports a parse error with its code and continues with the next file' {
        $bad = Join-Path $TestDrive 'bad.toml'
        $good = Join-Path $TestDrive 'good.toml'
        Set-Content -LiteralPath $bad -Value 'a = ' -NoNewline
        Set-Content -LiteralPath $good -Value 'a = 1' -NoNewline

        $result = @(Import-OfficeToml -Path $bad, $good -ErrorAction SilentlyContinue -ErrorVariable failures)
        $result.a | Should -Be 1
        $failures | Should -HaveCount 1
        $failures[0].FullyQualifiedErrorId | Should -Match '^toml-'
        { 'a = ' | ConvertFrom-OfficeToml -ErrorAction Stop } | Should -Throw
    }

    It 'writes values before sections, lists of objects as arrays of tables, and round-trips' {
        $value = [ordered]@{
            servers = @([ordered]@{ host = 'alpha'; port = 80 }, [ordered]@{ host = 'beta'; port = 443 })
            title = 'Report'
            owner = [ordered]@{ name = 'Ops'; updated = [DateTimeOffset]'2026-11-11T09:30:00Z' }
            tags = @('a', 'b')
        }

        $text = ConvertTo-OfficeToml -InputObject $value -Depth 3
        $expected = @(
            'title = "Report"'
            'tags = ["a", "b"]'
            ''
            '[[servers]]'
            'host = "alpha"'
            'port = 80'
            ''
            '[[servers]]'
            'host = "beta"'
            'port = 443'
            ''
            '[owner]'
            'name = "Ops"'
            'updated = 2026-11-11T09:30:00Z'
        )
        $text | Should -BeExactly (($expected -join "`n") + "`n")

        $back = $text | ConvertFrom-OfficeToml
        $back.servers.port | Should -Be @(80, 443)
        $back.owner.updated | Should -Be ([DateTimeOffset]'2026-11-11T09:30:00Z')
    }

    It 'wraps piped objects under RootName, omits nulls with warnings and exports atomically' {
        $text = @([pscustomobject]@{ name = 'a'; note = $null }, [pscustomobject]@{ name = 'b'; note = 'x' }) |
            ConvertTo-OfficeToml -RootName item -WarningVariable warnings -WarningAction SilentlyContinue
        ($text | ConvertFrom-OfficeToml).item.name | Should -Be @('a', 'b')
        $warnings.Message | Should -Contain "TOML documents are tables, so the input was written under the key 'item'."
        $warnings.Message | Should -Contain 'TOML has no null, so 1 null value was omitted.'

        $path = Join-Path $TestDrive 'settings.toml'
        $file = [ordered]@{ level = 'Information' } | Export-OfficeToml -Path $path -PassThru
        $file.FullName | Should -Be $path
        [IO.File]::ReadAllText($path) | Should -BeExactly "level = `"Information`"`n"
        { [ordered]@{ level = 'Debug' } | Export-OfficeToml -Path $path -NoClobber -ErrorAction Stop } | Should -Throw
        [IO.File]::ReadAllText($path) | Should -BeExactly "level = `"Information`"`n"
    }
}
