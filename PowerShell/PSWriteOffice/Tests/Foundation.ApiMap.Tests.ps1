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
    It 'exports the exact target command and parameter-set surface' {
        $actual = @($module.ExportedCmdlets.Keys | Sort-Object)
        $expected = @($maps.commands.name | Sort-Object)
        Compare-Object $actual $expected | Should -BeNullOrEmpty
        $commonParameters = @(
            'Verbose', 'Debug', 'ErrorAction', 'WarningAction', 'InformationAction', 'ProgressAction'
            'ErrorVariable', 'WarningVariable', 'InformationVariable', 'OutVariable', 'OutBuffer', 'PipelineVariable'
            'WhatIf', 'Confirm', 'UseTransaction'
        )
        foreach ($target in $maps.commands) {
            $command = Get-Command $target.name -Module $module.Name -ErrorAction Stop
            $metadata = [System.Management.Automation.CommandMetadata]::new($command)
            $supportsShouldProcess = $target.PSObject.Properties['shouldProcess'] -and $target.shouldProcess
            $metadata.SupportsShouldProcess | Should -Be ([bool] $supportsShouldProcess) -Because $target.name
            $command.Parameters.ContainsKey('WhatIf') | Should -Be ([bool] $supportsShouldProcess) -Because $target.name
            $command.Parameters.ContainsKey('Confirm') | Should -Be ([bool] $supportsShouldProcess) -Because $target.name
            foreach ($property in $command.ImplementingType.GetProperties()) {
                if ($property.Name -in $commonParameters) {
                    @($property.GetCustomAttributes([System.Management.Automation.ParameterAttribute], $true)).Count |
                        Should -Be 0 -Because "$($target.name) uses native common parameters"
                }
            }
            Compare-Object @($command.ParameterSets.Name | Sort-Object) @($target.parameterSets.name | Sort-Object) |
                Should -BeNullOrEmpty -Because $target.name
            if ($target.parameterSets.Count -gt 1) {
                @($command.ParameterSets | Where-Object IsDefault).Count | Should -Be 1 -Because $target.name
            }
            foreach ($set in $target.parameterSets) {
                $actualSet = @($command.ParameterSets | Where-Object Name -eq $set.name)
                $actualSet.Count | Should -Be 1 -Because "$($target.name) $($set.name)"
                $expectedDefault = $set.PSObject.Properties['default'] -and $set.default
                if ($target.parameterSets.Count -gt 1 -or $expectedDefault) {
                    $actualSet[0].IsDefault | Should -Be ([bool] $expectedDefault) -Because "$($target.name) $($set.name)"
                }
                $actualNames = @($actualSet[0].Parameters.Name | Where-Object { $_ -notin $commonParameters } | Sort-Object)
                Compare-Object $actualNames @($set.parameters.name | Sort-Object) |
                    Should -BeNullOrEmpty -Because "$($target.name) $($set.name)"
                foreach ($parameter in $set.parameters) {
                    $actualParameter = @($actualSet[0].Parameters | Where-Object Name -eq $parameter.name)[0]
                    $label = "$($target.name) $($set.name) -$($parameter.name)"
                    $mandatory = $parameter.PSObject.Properties['mandatory'] -and $parameter.mandatory
                    $actualParameter.IsMandatory | Should -Be ([bool] $mandatory) -Because $label
                    $position = [int]::MinValue
                    if ($parameter.PSObject.Properties['position']) { $position = [int] $parameter.position }
                    $actualParameter.Position | Should -Be $position -Because $label
                    $pipeline = 'none'
                    if ($parameter.PSObject.Properties['pipeline']) { $pipeline = $parameter.pipeline }
                    $actualParameter.ValueFromPipeline | Should -Be ($pipeline -in 'value', 'byValue', 'byValueAndPropertyName') -Because $label
                    $actualParameter.ValueFromPipelineByPropertyName |
                        Should -Be ($pipeline -in 'property', 'propertyName', 'byPropertyName', 'byValueAndPropertyName') -Because $label
                    $actualType = $actualParameter.ParameterType
                    $underlyingType = [Nullable]::GetUnderlyingType($actualType)
                    if ($underlyingType) { $actualType = $underlyingType }
                    $expectedType = [System.Management.Automation.PSTypeName]::new($parameter.type).Type
                    if (-not $expectedType) { $expectedType = $actualType.Assembly.GetType($parameter.type, $false) }
                    if ($expectedType) {
                        $actualType.FullName | Should -Be $expectedType.FullName -Because $label
                    } else {
                        $actualType.Name | Should -Be ($parameter.type -replace '^.*\.', '') -Because $label
                    }
                }
            }
        }
    }
    It 'exports one format-prefixed singular DSL alias per command without built-in collisions' {
        $builtIns = @(& pwsh -NoProfile -Command 'Get-Command -CommandType Cmdlet,Function,Alias | ForEach-Object Name')
        $aliases = @($module.ExportedAliases.Values)
        @($aliases | Group-Object ResolvedCommandName | Where-Object Count -gt 1).Count | Should -Be 0
        $prefixes = 'Word|Excel|PowerPoint|Pdf|Markdown|Visio|Html|Rtf|OpenDocument|Csv|Email|Project|Access|Epub|OneNote|IWork|AsciiDoc|Latex|Confluence'
        foreach ($alias in $aliases) {
            $alias.Name | Should -Match "^($prefixes)[A-Z][A-Za-z0-9]*$"
            $alias.Name | Should -Not -Match '-'
            $builtIns | Should -Not -Contain $alias.Name
            # The reviewed map owns domain spelling, including the agreed Markdown shortcuts.
            $target = @($maps.commands | Where-Object name -eq $alias.ResolvedCommandName)
            $target.Count | Should -Be 1
            $alias.Name | Should -BeExactly $target[0].alias
        }
        foreach ($target in $maps.commands | Where-Object alias) {
            @($aliases | Where-Object ResolvedCommandName -eq $target.name).Count | Should -Be 1
        }
    }
}

Describe 'Migration generator document preservation' {
    It 'preserves surrounding text or rejects ambiguous regions: <Region>' -ForEach @(
        @{ Region = 'unmarked'; ExtraBegin = ''; End = ''; Valid = $false }
        @{ Region = 'paired'; ExtraBegin = ''; End = '<!-- officeimo-api-map:end -->'; Valid = $true }
        @{ Region = 'missing-end'; ExtraBegin = ''; End = ''; Valid = $false }
        @{ Region = 'duplicate-begin'; ExtraBegin = '<!-- officeimo-api-map:begin -->'; End = '<!-- officeimo-api-map:end -->'; Valid = $false }
    ) {
        $path = Join-Path $TestDrive "$Region.md"
        $prefix = "# Maintained introduction`r`n`r`n"
        $suffix = "`r`n## Maintained history`r`n"
        $begin = if ($Region -eq 'unmarked') { '' } else { '<!-- officeimo-api-map:begin -->' }
        $inputText = $prefix + "$begin`n$ExtraBegin`nOld generated content`n$End" + $suffix
        [IO.File]::WriteAllText($path, $inputText)
        $generate = {
            $ErrorActionPreference = 'Continue'
            & pwsh -NoProfile -File (Join-Path $repo 'Build/ApiMap/New-MigrationGuide.ps1') -OutputPath $path -FailOnDataWarning 2>&1
        }
        if ($Valid) {
            $output = & $generate
            $LASTEXITCODE | Should -Be 0 -Because ($output -join [Environment]::NewLine)
            $first = [IO.File]::ReadAllText($path)
            $first.StartsWith($prefix + '<!-- officeimo-api-map:begin -->') | Should -BeTrue
            $first.EndsWith('<!-- officeimo-api-map:end -->' + $suffix) | Should -BeTrue
            $output = & $generate
            $LASTEXITCODE | Should -Be 0 -Because ($output -join [Environment]::NewLine)
            [IO.File]::ReadAllText($path) | Should -BeExactly $first
        } else {
            $output = & $generate
            $LASTEXITCODE | Should -Not -Be 0
            ($output -join [Environment]::NewLine) | Should -Match 'one ordered pair'
            [IO.File]::ReadAllText($path) | Should -BeExactly $inputText
        }
    }
}
