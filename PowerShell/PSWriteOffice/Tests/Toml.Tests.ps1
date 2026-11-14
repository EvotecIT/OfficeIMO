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

    function Join-TomlLines([string] $NewLine, [string[]] $Lines) { ($Lines -join $NewLine) + $NewLine }
}

Describe 'TOML document commands' {
    It 'saves an edit inside an Open block and keeps every other byte, including CRLF, comments and spelling' {
        $path = Join-Path $TestDrive 'Cargo.toml'
        $lines = @(
            '# Crate manifest'
            '[package]'
            'name    = "core"   # crate name'
            'version = "0.3.0"'
            'flags   = 0xFF'
            ''
            '[dependencies]'
            "serde = { version = '1', features = ['derive'] }"
        )
        [IO.File]::WriteAllText($path, (Join-TomlLines "`r`n" $lines))

        Open-OfficeToml -Path $path {
            Set-OfficeTomlNode -Select 'package.version' -Value '0.4.0'
            Set-OfficeTomlNode -Select 'package.flags' -Value 255
        }

        $lines[3] = 'version = "0.4.0"'
        [IO.File]::ReadAllText($path) | Should -BeExactly (Join-TomlLines "`r`n" $lines)
    }

    It 'removes a key with its comment line, adds a table section and reports missing paths and nulls' {
        $text = Join-TomlLines "`n" @(
            'title = "Site"'
            '# temporary override'
            'debug = true'
            ''
            '[params]'
            'theme = "dark"'
        )
        $document = Open-OfficeToml -Text $text
        try {
            Remove-OfficeTomlNode -Document $document -Select 'debug'
            $node = Set-OfficeTomlNode -Document $document -Select 'menu.weight' -Value 10 -PassThru
            $node.Path | Should -Be 'menu.weight'
            $node.Value | Should -Be 10

            $document.ToString() | Should -BeExactly (Join-TomlLines "`n" @(
                    'title = "Site"'
                    ''
                    '[params]'
                    'theme = "dark"'
                    ''
                    '[menu]'
                    'weight = 10'
                ))

            Remove-OfficeTomlNode -Document $document -Select 'params.missing' -ErrorAction SilentlyContinue -ErrorVariable missing
            $missing[0].FullyQualifiedErrorId | Should -Match '^toml-node-not-found'
            Set-OfficeTomlNode -Document $document -Select 'title' -Value $null -ErrorAction SilentlyContinue -ErrorVariable nullError
            $nullError[0].Exception.Message | Should -Match 'TOML has no null'
        } finally {
            Close-OfficeToml -Document $document
        }
    }

    It 'reports node values with 1-based positions and returns the root table without Select' {
        $document = Open-OfficeToml -Text (Join-TomlLines "`n" @('[server]', 'port = 8080', 'hosts = ["a", "b"]'))
        try {
            $port = Get-OfficeTomlNode -Document $document -Select 'server.port'
            $port.Value | Should -Be 8080
            $port.Value | Should -BeOfType [long]
            $port.Line | Should -Be 2
            $port.Column | Should -Be 8

            (Get-OfficeTomlNode -Document $document -Select 'server.hosts[1]').Value | Should -Be 'b'
            $root = Get-OfficeTomlNode -Document $document -AsHashtable
            $root.Path | Should -Be '$'
            $root.Value.server.port | Should -Be 8080
        } finally {
            Close-OfficeToml -Document $document
        }
    }

    It 'creates a live document, saves it to the path it names later and saves it there again' {
        $path = Join-Path $TestDrive 'live.toml'
        $document = New-OfficeToml
        Set-OfficeTomlNode -Document $document -Select 'name' -Value 'web'
        (Save-OfficeToml -Document $document -Path $path -PassThru).FullName | Should -Be $path
        Set-OfficeTomlNode -Document $document -Select 'replicas' -Value 3
        Save-OfficeToml -Document $document
        Close-OfficeToml -Document $document

        [IO.File]::ReadAllText($path) | Should -BeExactly "name = `"web`"`nreplicas = 3`n"
    }

    It 'writes a new file from a content block and refuses to replace it without Force' {
        $path = Join-Path $TestDrive 'pyproject.toml'
        New-OfficeToml -Path $path {
            Set-OfficeTomlNode -Select 'project.name' -Value 'reporting'
        }
        (Import-OfficeToml -Path $path).project.name | Should -Be 'reporting'
        { New-OfficeToml -Path $path { Set-OfficeTomlNode -Select 'a' -Value 1 } -ErrorAction Stop } | Should -Throw
    }

    It 'fails on invalid TOML with the engine code in the error ID' {
        { Open-OfficeToml -Text "a = 1`na = 2`n" -ErrorAction Stop } | Should -Throw -ErrorId 'toml-Toml.Load.DuplicateKey*'
    }

    It 'merges overlays through Merge-OfficeDocument, keeping the base comments, and reports what changed' {
        $base = Join-Path $TestDrive 'pyproject.toml'
        $overlay = Join-Path $TestDrive 'pyproject.local.toml'
        $output = Join-Path $TestDrive 'pyproject.merged.toml'
        [IO.File]::WriteAllText($base, (Join-TomlLines "`n" @('# shared settings', '[tool.ruff]', 'line-length = 88 # team default', 'exclude = ["build"]')))
        [IO.File]::WriteAllText($overlay, (Join-TomlLines "`n" @('[tool.ruff]', 'line-length = 100', '', '[tool.ruff.lint]', 'select = ["E"]')))

        $report = Merge-OfficeDocument -Path $base, $overlay -OutputPath $output -PassThruReport

        @($report).Count | Should -Be 1
        @($report.Replaced) | Should -Be @('tool.ruff.line-length')
        @($report.Added) | Should -Be @('tool.ruff.lint')
        [IO.File]::ReadAllText($output) | Should -BeExactly (Join-TomlLines "`n" @('# shared settings', '[tool.ruff]', 'line-length = 100 # team default', 'exclude = ["build"]', '', '[tool.ruff.lint]', 'select = ["E"]'))
    }

    It 'merges arrays of tables by key into a live document through the Arrays option' {
        $overlay = Join-Path $TestDrive 'plugins.local.toml'
        [IO.File]::WriteAllText($overlay, (Join-TomlLines "`n" @('[[plugin]]', 'name = "lint"', 'version = "2.0"', '', '[[plugin]]', 'name = "docs"', 'version = "1.0"')))
        $document = Open-OfficeToml -Text (Join-TomlLines "`n" @('[[plugin]]', 'name = "lint"', 'version = "1.0"'))
        try {
            $report = Merge-OfficeDocument -Document $document -Path $overlay -Options @{ Arrays = @{ MergeByKey = 'name' } } -PassThruReport
            @($report.Replaced) | Should -Be @('plugin[0].version')
            @($report.Added) | Should -Be @('plugin[1]')
            $document.ToString() | Should -BeExactly (Join-TomlLines "`n" @('[[plugin]]', 'name = "lint"', 'version = "2.0"', '', '[[plugin]]', 'name = "docs"', 'version = "1.0"'))
        } finally {
            Close-OfficeToml -Document $document
        }
    }

    It 'reports kind conflicts in the merge report and rejects passwords' {
        $base = Join-Path $TestDrive 'conflict.toml'
        $overlay = Join-Path $TestDrive 'conflict.local.toml'
        $output = Join-Path $TestDrive 'conflict.merged.toml'
        [IO.File]::WriteAllText($base, "log = `"info`"`n")
        [IO.File]::WriteAllText($overlay, "[log]`nlevel = `"debug`"`n")

        $report = Merge-OfficeDocument -Path $base, $overlay -OutputPath $output -PassThruReport
        @($report.Diagnostics.Code) | Should -Be @('toml-merge-kind-conflict')
        (Import-OfficeToml -Path $output).log.level | Should -Be 'debug'
        { Merge-OfficeDocument -Path $base, $overlay -OutputPath (Join-Path $TestDrive 'other.toml') -Password 'x' } | Should -Throw '*no password*'
    }
}

Describe 'Test-OfficeToml' {
    It 'returns a Boolean per file and the error code and position with Detailed' {
        $good = Join-Path $TestDrive 'good.toml'
        $bad = Join-Path $TestDrive 'bad.toml'
        [IO.File]::WriteAllText($good, "[server]`nport = 8080`n")
        [IO.File]::WriteAllText($bad, "[server]`nport = 8080`nport = 9090`n")

        Test-OfficeToml -Path $bad, $good | Should -Be @($false, $true)
        @(Get-ChildItem -LiteralPath $TestDrive -Filter '*.toml' | Where-Object Name -in 'good.toml', 'bad.toml' | Sort-Object Name | Test-OfficeToml) | Should -Be @($false, $true)

        $result = Test-OfficeToml -LiteralPath $bad -Detailed
        $result.IsValid | Should -BeFalse
        $result.Path | Should -Be $bad
        $result.Code | Should -Be 'toml-load-duplicate-key'
        $result.Line | Should -Be 3
        $result.Column | Should -Be 1
        $result.Message | Should -Not -BeLike '*line 3*'

        $valid = Test-OfficeToml -Text 'a = 1' -Detailed
        $valid.IsValid | Should -BeTrue
        $valid.Path | Should -BeNullOrEmpty
        $valid.Code | Should -BeNullOrEmpty
    }

    It 'applies the untrusted limits with Untrusted' {
        $deep = ('a' + ('.a' * 70)) + " = 1`n"
        Test-OfficeToml -Text $deep | Should -BeTrue
        Test-OfficeToml -Text $deep -Untrusted | Should -BeFalse
        (Test-OfficeToml -Text $deep -Untrusted -Detailed).Code | Should -BeLike 'toml-limit*'
    }
}
