Describe 'Packaged AssemblyLoadContext conflict isolation' {
    It 'loads after a default-context Open XML assembly without sharing the default ALC' {
        $packagedModuleRoot = Join-Path $PSScriptRoot '..\Artefacts\Unpacked\Modules'
        $packagedModule = Join-Path $packagedModuleRoot 'PSWriteOffice'
        $packagedLoader = Join-Path $packagedModule 'Lib\Core\PSWriteOffice.ModuleLoadContext.dll'
        $conflictOpenXmlPath = $env:PSWRITEOFFICE_CONFLICT_OPENXML_PATH
        $requirePackagedSmoke = $env:PSWRITEOFFICE_REQUIRE_ALC_SMOKE -eq 'true'

        if ($PSVersionTable.PSEdition -ne 'Core') {
            Set-ItResult -Skipped -Because 'AssemblyLoadContext isolation is only available in PowerShell Core'
            return
        }

        if (-not $requirePackagedSmoke -and (
                -not (Test-Path -LiteralPath $packagedModule) -or
                -not (Test-Path -LiteralPath $packagedLoader) -or
                [string]::IsNullOrWhiteSpace($conflictOpenXmlPath) -or
                -not (Test-Path -LiteralPath $conflictOpenXmlPath))) {
            Set-ItResult -Skipped -Because 'packaged ALC smoke prerequisites are supplied by the dedicated packaged-alc-conflict workflow'
            return
        }

        Test-Path -LiteralPath $packagedModule | Should -BeTrue -Because 'the packaged module must exist before the ALC smoke test runs'
        Test-Path -LiteralPath $packagedLoader | Should -BeTrue -Because 'the packaged module must include the module-scoped ALC loader'
        [string]::IsNullOrWhiteSpace($conflictOpenXmlPath) | Should -BeFalse -Because 'the conflict Open XML assembly path must be provided by the workflow'
        Test-Path -LiteralPath $conflictOpenXmlPath | Should -BeTrue -Because 'the conflict Open XML assembly must exist before the smoke test runs'

        $moduleRootLiteral = $packagedModuleRoot.Replace("'", "''")
        $conflictOpenXmlLiteral = $conflictOpenXmlPath.Replace("'", "''")
        $script = @"
`$ErrorActionPreference = 'Stop'
`$WarningPreference = 'SilentlyContinue'
`$moduleRoot = '$moduleRootLiteral'
`$conflictOpenXmlPath = '$conflictOpenXmlLiteral'
`$env:PSModulePath = `$moduleRoot + [IO.Path]::PathSeparator + `$env:PSModulePath

Add-Type -Path `$conflictOpenXmlPath -ErrorAction Stop
`$defaultOpenXmlAssembly = [AppDomain]::CurrentDomain.GetAssemblies() |
    Where-Object { `$_.GetName().Name -eq 'DocumentFormat.OpenXml' } |
    Select-Object -First 1
`$defaultOpenXmlAlc = [System.Runtime.Loader.AssemblyLoadContext]::GetLoadContext(`$defaultOpenXmlAssembly)

Import-Module PSWriteOffice -Force
`$command = Get-Command New-OfficeWord -ErrorAction Stop
`$excelSheetAlias = Get-Alias ExcelSheet -ErrorAction Stop
`$commandAssembly = `$command.ImplementingType.Assembly
`$commandAlc = [System.Runtime.Loader.AssemblyLoadContext]::GetLoadContext(`$commandAssembly)
`$loadedAssemblies = [System.Runtime.Loader.AssemblyLoadContext]::All |
    ForEach-Object {
        `$alc = `$_
        foreach (`$assembly in `$alc.Assemblies) {
            if (`$assembly.GetName().Name -in @('PSWriteOffice', 'OfficeIMO.Word', 'DocumentFormat.OpenXml')) {
                [pscustomobject]@{
                    Assembly = `$assembly.GetName().Name
                    Version = `$assembly.GetName().Version.ToString()
                    ALC = `$alc.Name
                    IsDefault = [object]::ReferenceEquals(`$alc, [System.Runtime.Loader.AssemblyLoadContext]::Default)
                    Location = `$assembly.Location
                }
            }
        }
    }

[pscustomobject]@{
    DefaultOpenXmlAssembly = `$defaultOpenXmlAssembly.Location
    DefaultOpenXmlVersion = `$defaultOpenXmlAssembly.GetName().Version.ToString()
    DefaultOpenXmlALC = `$defaultOpenXmlAlc.Name
    DefaultOpenXmlALCIsDefault = [object]::ReferenceEquals(`$defaultOpenXmlAlc, [System.Runtime.Loader.AssemblyLoadContext]::Default)
    NewOfficeWordAssembly = `$commandAssembly.Location
    ExcelSheetAliasTarget = `$excelSheetAlias.ResolvedCommandName
    NewOfficeWordALC = `$commandAlc.Name
    NewOfficeWordALCIsDefault = [object]::ReferenceEquals(`$commandAlc, [System.Runtime.Loader.AssemblyLoadContext]::Default)
    LoadedAssemblies = @(`$loadedAssemblies)
} | ConvertTo-Json -Depth 6 -Compress
"@
        $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($script))
        $output = pwsh -NoProfile -ExecutionPolicy Bypass -EncodedCommand $encoded 2>&1
        $LASTEXITCODE | Should -Be 0 -Because ($output -join [Environment]::NewLine)

        $json = $output | Where-Object { $_ -is [string] -and $_.TrimStart().StartsWith('{') } | Select-Object -Last 1
        $json | Should -Not -BeNullOrEmpty -Because ($output -join [Environment]::NewLine)
        $result = $json | ConvertFrom-Json

        $result.DefaultOpenXmlAssembly | Should -Be $conflictOpenXmlPath
        $result.DefaultOpenXmlALCIsDefault | Should -BeTrue
        $result.NewOfficeWordAssembly | Should -BeLike '*\Artefacts\Unpacked\Modules\PSWriteOffice\Lib\Core\PSWriteOffice.dll'
        $result.ExcelSheetAliasTarget | Should -Be 'Add-OfficeExcelSheet'
        $result.NewOfficeWordALC | Should -Be 'PSWriteOffice'
        $result.NewOfficeWordALCIsDefault | Should -BeFalse

        $loadedAssemblies = @($result.LoadedAssemblies)
        $psWriteOfficeAssembly = $loadedAssemblies | Where-Object { $_.Assembly -eq 'PSWriteOffice' -and $_.ALC -eq 'PSWriteOffice' } | Select-Object -First 1
        $officeImoAssembly = $loadedAssemblies | Where-Object { $_.Assembly -eq 'OfficeIMO.Word' -and $_.ALC -eq 'PSWriteOffice' } | Select-Object -First 1
        $moduleOpenXmlAssembly = $loadedAssemblies | Where-Object { $_.Assembly -eq 'DocumentFormat.OpenXml' -and $_.ALC -eq 'PSWriteOffice' } | Select-Object -First 1
        $defaultOpenXmlAssembly = $loadedAssemblies | Where-Object { $_.Assembly -eq 'DocumentFormat.OpenXml' -and $_.IsDefault } | Select-Object -First 1

        $psWriteOfficeAssembly.IsDefault | Should -BeFalse
        $officeImoAssembly.IsDefault | Should -BeFalse
        $moduleOpenXmlAssembly.IsDefault | Should -BeFalse
        $defaultOpenXmlAssembly.IsDefault | Should -BeTrue
    }
}
