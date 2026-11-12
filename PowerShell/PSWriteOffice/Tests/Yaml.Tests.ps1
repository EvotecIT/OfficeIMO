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

    function Join-YamlLines([string] $NewLine, [string[]] $Lines) { ($Lines -join $NewLine) + $NewLine }
}

Describe 'YAML document commands' {
    It 'saves an edit inside an Open block and keeps every other byte, including CRLF, comments and quoting' {
        $path = Join-Path $TestDrive 'values.yaml'
        $lines = @(
            '# Defaults for the web chart'
            'image:'
            '  repository: nginx   # pinned by CI'
            "  tag: '1.25'"
            'replicas: 1'
            ''
            '# Ports the service exposes'
            'ports: [80, 443]'
            'enabled: TRUE'
        )
        [IO.File]::WriteAllText($path, (Join-YamlLines "`r`n" $lines))

        Open-OfficeYaml -Path $path {
            Set-OfficeYamlNode -Select 'image.tag' -Value '1.27'
            Set-OfficeYamlNode -Select 'enabled' -Value $true
        }

        $lines[3] = "  tag: '1.27'"
        [IO.File]::ReadAllText($path) | Should -BeExactly (Join-YamlLines "`r`n" $lines)
    }

    It 'removes an entry with its comment line and adds a nested mapping in the neighbors layout' {
        $text = Join-YamlLines "`n" @(
            'metadata:'
            '    name: web'
            '    # temporary override'
            '    debug: true'
            'spec:'
            '    replicas: 2'
        )
        $document = Open-OfficeYaml -Text $text
        try {
            Remove-OfficeYamlNode -Document $document -Select 'metadata.debug'
            $node = Set-OfficeYamlNode -Document $document -Select 'spec.resources' -Value ([ordered]@{ cpu = '500m'; memory = '256Mi' }) -PassThru
            $node.Path | Should -Be 'spec.resources'

            $document.ToString() | Should -BeExactly (Join-YamlLines "`n" @(
                    'metadata:'
                    '    name: web'
                    'spec:'
                    '    replicas: 2'
                    '    resources:'
                    '        cpu: 500m'
                    '        memory: 256Mi'
                ))
            Remove-OfficeYamlNode -Document $document -Select 'metadata.missing' -ErrorAction SilentlyContinue -ErrorVariable missing
            $missing[0].FullyQualifiedErrorId | Should -Match '^yaml-node-not-found'
        } finally {
            Close-OfficeYaml -Document $document
        }
    }

    It 'reports node values with 1-based positions and addresses documents of a stream' {
        $document = Open-OfficeYaml -Text (Join-YamlLines "`n" @('kind: Service', 'spec:', '  port: 80', '---', 'kind: Deployment'))
        try {
            $port = Get-OfficeYamlNode -Document $document -Select 'spec.port'
            $port.Value | Should -Be 80
            $port.Value | Should -BeOfType [long]
            $port.Line | Should -Be 3
            $port.Column | Should -Be 9

            (Get-OfficeYamlNode -Document $document -Select '$[1].kind').Value | Should -Be 'Deployment'
            @(Get-OfficeYamlNode -Document $document).Path | Should -Be @('$[0]', '$[1]')
        } finally {
            Close-OfficeYaml -Document $document
        }
    }

    It 'creates a live document, saves it to the path it names later and saves it there again' {
        $path = Join-Path $TestDrive 'live.yaml'
        $document = New-OfficeYaml
        Set-OfficeYamlNode -Document $document -Select 'name' -Value 'web'
        (Save-OfficeYaml -Document $document -Path $path -PassThru).FullName | Should -Be $path
        Set-OfficeYamlNode -Document $document -Select 'replicas' -Value 3
        Save-OfficeYaml -Document $document
        Close-OfficeYaml -Document $document

        [IO.File]::ReadAllText($path) | Should -BeExactly "name: web`nreplicas: 3`n"
    }

    It 'merges overlays through Merge-OfficeDocument and reports what changed' {
        $base = Join-Path $TestDrive 'values.yaml'
        $overlay = Join-Path $TestDrive 'values.prod.yml'
        $output = Join-Path $TestDrive 'values.merged.yaml'
        [IO.File]::WriteAllText($base, (Join-YamlLines "`n" @('# defaults', 'image:', '  tag: dev   # overridden per environment', 'debug: true', 'ports: [80]')))
        [IO.File]::WriteAllText($overlay, (Join-YamlLines "`n" @('image:', '  tag: "1.27"', 'debug: null', 'ports: [443]')))

        $report = Merge-OfficeDocument -Path $base, $overlay -OutputPath $output -Options @{ NullRemovesKey = $true } -PassThruReport

        @($report).Count | Should -Be 1
        @($report.Replaced) | Should -Be @('image.tag', 'ports')
        @($report.Removed) | Should -Be @('debug')
        [IO.File]::ReadAllText($output) | Should -BeExactly (Join-YamlLines "`n" @('# defaults', 'image:', "  tag: '1.27'   # overridden per environment", 'ports: [443]'))
    }
}

Describe 'YAML data conversion' {
    It 'reads JSON text as ConvertFrom-Json does, with or without AsHashtable' {
        $json = '{"name":"web","replicas":3,"ratio":0.5,"enabled":false,"owner":null,"tags":["a","b"],"limits":{"cpu":"500m"},"big":123456789012345678901234567890}'

        # Windows PowerShell 5.1 reads integers beyond Int64 as lossy doubles, so compare the rest there.
        $parityJson = if ($PSVersionTable.PSEdition -eq 'Desktop') { $json -replace ',"big":\d+', '' } else { $json }
        $expected = $parityJson | ConvertFrom-Json | ConvertTo-Json -Depth 5 -Compress
        $parityJson | ConvertFrom-OfficeYaml | ConvertTo-Json -Depth 5 -Compress | Should -BeExactly $expected

        $table = $json | ConvertFrom-OfficeYaml -AsHashtable
        $table | Should -BeOfType [System.Collections.Specialized.OrderedDictionary]
        @($table.Keys) | Should -Be @('name', 'replicas', 'ratio', 'enabled', 'owner', 'tags', 'limits', 'big')
        $table.replicas | Should -BeOfType [long]
        $table.ratio | Should -BeOfType [double]
        $table.big | Should -BeOfType [System.Numerics.BigInteger]
        $table.limits | Should -BeOfType [System.Collections.Specialized.OrderedDictionary]
    }

    It 'resolves scalars through the selected schema' {
        $yaml = Join-YamlLines "`n" @('port: 0x1F', 'answer: yes', 'version: 1.20', 'empty: ~')

        $core = $yaml | ConvertFrom-OfficeYaml
        $core.port | Should -Be 31
        $core.answer | Should -BeExactly 'yes'
        $core.version | Should -Be 1.2
        $core.empty | Should -BeNullOrEmpty

        ($yaml | ConvertFrom-OfficeYaml -Schema Yaml11).answer | Should -BeTrue
        $strings = $yaml | ConvertFrom-OfficeYaml -Schema Failsafe
        foreach ($name in 'port', 'answer', 'version', 'empty') { $strings.$name | Should -BeOfType [string] }
        $strings.version | Should -BeExactly '1.20'
    }

    It 'limits nesting with Depth exactly as ConvertFrom-Json does' {
        $json = '{"a":{"b":{"c":1}}}'
        { $json | ConvertFrom-Json -Depth 2 } | Should -Throw
        { $json | ConvertFrom-OfficeYaml -Depth 2 } | Should -Throw -ErrorId 'yaml-limit-exceeded*'
        ($json | ConvertFrom-OfficeYaml -Depth 3).a.b.c | Should -Be 1
        $block = "a:`n  b:`n    c: 1"
        { $block | ConvertFrom-OfficeYaml -Depth 2 } | Should -Throw -ErrorId 'yaml-limit-exceeded*'
        ($block | ConvertFrom-OfficeYaml -Depth 3).a.b.c | Should -Be 1
    }

    It 'enumerates a root sequence unless NoEnumerate and reads multi-document streams only with AllDocuments' {
        @("- 1`n- 2" | ConvertFrom-OfficeYaml).Count | Should -Be 2
        $array = "- 1`n- 2" | ConvertFrom-OfficeYaml -NoEnumerate
        , $array | Should -BeOfType [object[]]
        @(, $array).Count | Should -Be 1

        $stream = "name: a`n---`nname: b`n"
        { $stream | ConvertFrom-OfficeYaml } | Should -Throw -ErrorId 'yaml-multiple-documents*'
        @($stream | ConvertFrom-OfficeYaml -AllDocuments).name | Should -Be @('a', 'b')

        # Get-Content emits lines; they are joined before parsing.
        $path = Join-Path $TestDrive 'stream.yaml'
        [IO.File]::WriteAllText($path, $stream)
        @(Get-Content -LiteralPath $path | ConvertFrom-OfficeYaml -AllDocuments).name | Should -Be @('a', 'b')
        @(Import-OfficeYaml -LiteralPath $path -AllDocuments).name | Should -Be @('a', 'b')
    }

    It 'writes objects that read back to the same data and quotes strings YAML 1.1 would retype' {
        $data = [ordered]@{
            name     = 'web'
            answer   = 'yes'
            version  = '1.29.2'
            octal    = '012'
            replicas = 3
            ratio    = 0.25
            enabled  = $true
            owner    = $null
            ports    = @(80, 443)
            mode     = [System.IO.FileAttributes]::ReadOnly
        }

        $yaml = $data | ConvertTo-OfficeYaml
        $yaml | Should -BeExactly (Join-YamlLines "`n" @(
                'name: web'
                "answer: 'yes'"
                'version: 1.29.2'
                "octal: '012'"
                'replicas: 3'
                'ratio: 0.25'
                'enabled: true'
                'owner: null'
                'ports:'
                '- 80'
                '- 443'
                'mode: ReadOnly'
            ))
        ($yaml | ConvertFrom-OfficeYaml -AsHashtable | ConvertTo-Json -Compress) |
            Should -BeExactly ([ordered]@{ name = 'web'; answer = 'yes'; version = '1.29.2'; octal = '012'; replicas = 3; ratio = 0.25; enabled = $true; owner = $null; ports = @(80, 443); mode = 'ReadOnly' } | ConvertTo-Json -Compress)
    }

    It 'truncates past Depth with the ConvertTo-Json warning and string form' {
        $value = @{ a = @{ b = @{ c = @{ d = 1 } } } }
        $null = $value | ConvertTo-Json -WarningVariable jsonWarning -WarningAction SilentlyContinue
        $yaml = $value | ConvertTo-OfficeYaml -WarningVariable yamlWarning -WarningAction SilentlyContinue

        $yaml | Should -BeExactly "a:`n  b:`n    c: System.Collections.Hashtable`n"
        # Windows PowerShell 5.1's ConvertTo-Json truncates without a warning; PowerShell 7 words it this way.
        $yamlWarning[0].Message | Should -BeExactly 'Resulting YAML is truncated as serialization has exceeded the set depth of 2.'
        if ($jsonWarning) { $yamlWarning[0].Message | Should -BeExactly ($jsonWarning[0].Message -replace 'JSON', 'YAML') }
        ($value | ConvertTo-OfficeYaml -Depth 3 -WarningVariable deeper -WarningAction SilentlyContinue) | Should -BeExactly "a:`n  b:`n    c:`n      d: 1`n"
        $deeper | Should -BeNullOrEmpty
    }

    It 'writes several piped objects as one sequence and honors writer options' {
        $yaml = [pscustomobject]@{ id = 1 }, [pscustomobject]@{ id = 2 } | ConvertTo-OfficeYaml -WriteOptions @{ SequenceIndent = 4; SequenceDashOffset = 2 }
        $yaml | Should -BeExactly "- id: 1`n- id: 2`n"
        ([ordered]@{ list = 1, 2 } | ConvertTo-OfficeYaml -WriteOptions @{ SequenceIndent = 4; SequenceDashOffset = 2 }) |
            Should -BeExactly "list:`n  - 1`n  - 2`n"
    }

    It 'exports and imports files, refusing to replace a file with NoClobber' {
        $path = Join-Path $TestDrive 'settings.yaml'
        [ordered]@{ logging = [ordered]@{ level = 'Information' }; retries = 3 } | Export-OfficeYaml -Path $path
        [IO.File]::ReadAllBytes($path)[0] | Should -Be ([byte][char]'l')
        $settings = Import-OfficeYaml -Path $path
        $settings.logging.level | Should -Be 'Information'
        $settings.retries | Should -Be 3

        { @{ a = 1 } | Export-OfficeYaml -Path $path -NoClobber -ErrorAction Stop } | Should -Throw
        (Import-OfficeYaml -Path $path).retries | Should -Be 3
    }
}

Describe 'YAML validation' {
    It 'returns Booleans and, with Detailed, the diagnostic code and 1-based position' {
        $bad = Join-Path $TestDrive 'bad.yaml'
        $good = Join-Path $TestDrive 'good.yaml'
        [IO.File]::WriteAllText($bad, "name: web`nports: [80, 443`n")
        [IO.File]::WriteAllText($good, "name: web`n")

        Test-OfficeYaml -Path $bad, $good | Should -Be @($false, $true)

        $result = Test-OfficeYaml -LiteralPath $bad -Detailed
        $result.IsValid | Should -BeFalse
        $result.Path | Should -Be $bad
        $result.Code | Should -BeLike 'yaml-syntax-*'
        $result.Line | Should -Be 3
        $result.Column | Should -Be 1
        @($result.Diagnostics.Code) | Should -Contain $result.Code

        { Open-OfficeYaml -Text "ports: [80, 443`n" } | Should -Throw -ErrorId 'yaml-syntax-*'
    }

    It 'applies Strict to input the default mode accepts with a warning' {
        $lenient = "steps: [build,`ntest]`n"
        $default = Test-OfficeYaml -Text $lenient -Detailed
        $default.IsValid | Should -BeTrue
        @($default.Diagnostics.Code) | Should -Contain 'yaml-lenient-indentation'
        Test-OfficeYaml -Text $lenient -Strict | Should -BeFalse
    }
}
