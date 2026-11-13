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

    It 'creates a live document for a path and saves it there twice' {
        $path = Join-Path $TestDrive 'live.toml'
        $document = New-OfficeToml -Path $path
        Set-OfficeTomlNode -Document $document -Select 'name' -Value 'web'
        (Save-OfficeToml -Document $document -PassThru).FullName | Should -Be $path
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
}
