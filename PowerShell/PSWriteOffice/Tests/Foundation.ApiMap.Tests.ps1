BeforeAll {
    $repo = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../..'))
    $moduleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) { $env:PSWRITEOFFICE_MODULE_MANIFEST } else { Join-Path $PSScriptRoot '../PSWriteOffice.psd1' }
    $module = Import-Module $moduleManifest -Global -PassThru -ErrorAction Stop
    $maps = @(Get-ChildItem -LiteralPath (Join-Path $repo 'Build/ApiMap/powershell') -Filter '*.json' | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw -Encoding UTF8 | ConvertFrom-Json })
}
Describe 'Command map foundations' {
    It 'validates map coverage and integrity against the built module metadata' {
        $metadata = [ordered]@{
            Cmdlets = @($module.ExportedCmdlets.Values | ForEach-Object { @{Name=$_.Name} })
            Aliases = @($module.ExportedAliases.Values | ForEach-Object { @{Alias=$_.Name} })
        }
        $metadataPath = Join-Path $TestDrive 'module.json'
        $metadata | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $metadataPath
        $output = & pwsh -NoProfile -File (Join-Path $repo 'Build/ApiMap/Test-ApiMap.ps1') -MetadataPath $metadataPath 2>&1
        $LASTEXITCODE | Should -Be 0 -Because ($output -join [Environment]::NewLine)
    }
    It 'matches the compiled hint resource to every retired entry hint' {
        $assembly = (Get-Command New-OfficeWord).ImplementingType.Assembly
        $reader = [IO.StreamReader]::new($assembly.GetManifestResourceStream('PSWriteOffice.CommandHints.json'))
        try { $data = $reader.ReadToEnd() | ConvertFrom-Json } finally { $reader.Dispose() }
        $hints = @{}
        foreach ($property in $data.PSObject.Properties) { $hints[$property.Name] = $property.Value }
        $expected = @($maps.map | Where-Object { $_.action -ne 'keep' -and $_.hint })
        $hints.Count | Should -Be $expected.Count
        foreach ($entry in $expected) { $hints[$entry.old] | Should -Be $entry.hint }
    }
    It 'exports the exact target command and parameter-set surface' -Skip {
        # M5 enables this after migrating the shipped commands. Integrity above remains active in M3.
        $actual = @($module.ExportedCmdlets.Keys | Sort-Object)
        $expected = @($maps.commands.name | Sort-Object)
        Compare-Object $actual $expected | Should -BeNullOrEmpty
        foreach ($target in $maps.commands) {
            $command = Get-Command $target.name -Module $module.Name -ErrorAction Stop
            if ($target.parameterSets.Count -gt 1) { @($command.ParameterSets | Where-Object IsDefault).Count | Should -Be 1 }
            foreach ($set in $target.parameterSets) {
                $actualSet = @($command.ParameterSets | Where-Object Name -eq $set.name)
                $actualSet.Count | Should -Be 1
                foreach ($parameter in $set.parameters) { $actualSet[0].Parameters.Name | Should -Contain $parameter.name }
            }
        }
    }
    It 'exports one format-prefixed singular DSL alias per command without built-in collisions' -Skip {
        # M5 removes today's compatibility aliases and enables this assertion.
        $builtIns = @(& pwsh -NoProfile -Command 'Get-Command -CommandType Cmdlet,Function,Alias | ForEach-Object Name')
        $aliases = @($module.ExportedAliases.Values)
        @($aliases | Group-Object ResolvedCommandName | Where-Object Count -gt 1).Count | Should -Be 0
        $prefixes = 'Word|Excel|Ppt|Pdf|Markdown|Visio|Html|Rtf|OpenDocument|Csv|Email|Project|Access|Epub|OneNote|IWork|AsciiDoc|Latex|Confluence'
        foreach ($alias in $aliases) {
            $alias.Name | Should -Match "^($prefixes)[A-Z][A-Za-z0-9]*$"
            $alias.Name | Should -Not -Match '-'
            $builtIns | Should -Not -Contain $alias.Name
            # The map records the singular domain noun and legitimate Office terms such as Notes/Contents.
            $target = @($maps.commands | Where-Object name -eq $alias.ResolvedCommandName)
            $target.Count | Should -Be 1
            $alias.Name | Should -BeExactly $target[0].alias
            if ($alias.Name -match 's$' -and $alias.Name -notmatch '(Notes|Contents|Series)$') {
                $alias.Name | Should -Not -Match 's$' -Because 'DSL aliases use a singular noun'
            }
        }
        foreach ($target in $maps.commands | Where-Object alias) {
            @($aliases | Where-Object ResolvedCommandName -eq $target.name).Count | Should -Be 1
        }
    }
}
