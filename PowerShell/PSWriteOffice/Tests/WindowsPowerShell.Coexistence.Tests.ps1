BeforeDiscovery {
    # dbatools is not a PSWriteOffice dependency; the test runs only where it is installed. It may be installed for
    # PowerShell 7 only, so look in both editions' module folders.
    $windowsPowerShell = if ($env:SystemRoot) { Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe' }
    $hasWindowsPowerShell = $windowsPowerShell -and (Test-Path -LiteralPath $windowsPowerShell)
    $documents = [Environment]::GetFolderPath('MyDocuments')
    $dbatoolsRoot = @(
        if ($documents) { Join-Path $documents 'WindowsPowerShell\Modules'; Join-Path $documents 'PowerShell\Modules' }
        if ($env:ProgramFiles) { Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules'; Join-Path $env:ProgramFiles 'PowerShell\Modules' }
    ) | Where-Object { $_ -and (Test-Path -LiteralPath (Join-Path $_ 'dbatools')) } | Select-Object -First 1
    $orders = @(
        @{ First = 'dbatools'; Second = 'PSWriteOffice'; DbatoolsRoot = $dbatoolsRoot }
        @{ First = 'PSWriteOffice'; Second = 'dbatools'; DbatoolsRoot = $dbatoolsRoot }
    )
}

Describe 'Windows PowerShell coexistence with dbatools' -Skip:(-not $hasWindowsPowerShell -or -not $dbatoolsRoot) {
    BeforeAll {
        $windowsPowerShell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
        $moduleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) { $env:PSWRITEOFFICE_MODULE_MANIFEST } else { Join-Path $PSScriptRoot '../PSWriteOffice.psd1' }
        $moduleManifest = (Resolve-Path -LiteralPath $moduleManifest).Path
    }

    It 'imports <First> then <Second> and both keep working' -TestCases $orders {
        param($First, $Second, $DbatoolsRoot)
        $workRoot = Join-Path $TestDrive "$First-$Second"
        New-Item -ItemType Directory -Path $workRoot -Force | Out-Null
        $script = @'
param($ModuleManifest, $DbatoolsRoot, $Order, $WorkRoot)
$ErrorActionPreference = 'Stop'
$WarningPreference = 'SilentlyContinue'
$env:PSModulePath = $DbatoolsRoot + [IO.Path]::PathSeparator + $env:PSModulePath
foreach ($name in $Order -split ',') {
    if ($name -eq 'PSWriteOffice') { Import-Module $ModuleManifest -DisableNameChecking } else { Import-Module dbatools -DisableNameChecking }
}

$markdown = Join-Path $WorkRoot 'input.md'
Set-Content -LiteralPath $markdown -Value "# Title`n`n| A | B |`n|---|---|`n| 1 | 2 |"
foreach ($extension in 'docx', 'html', 'pdf') {
    Convert-OfficeDocument -Path $markdown -OutputPath (Join-Path $WorkRoot "output.$extension") | Out-Null
}
$csv = @([pscustomobject]@{ Name = 'a'; Value = 1 }) | ConvertTo-OfficeCsv
$builder = New-DbaConnectionStringBuilder -SqlInstance 'localhost' -Database 'master'

$moduleLibrary = [IO.Path]::GetDirectoryName((Get-Command Convert-OfficeDocument).ImplementingType.Assembly.Location)
$json = [AppDomain]::CurrentDomain.GetAssemblies() | Where-Object {
    $_.GetName().Name -eq 'System.Text.Json' -and $_.Location -and [IO.Path]::GetDirectoryName($_.Location) -eq $moduleLibrary
} | Select-Object -First 1
$asyncDisposable = $json.GetType('System.Text.Json.Utf8JsonWriter', $true).GetInterface('System.IAsyncDisposable')
[pscustomobject]@{
    Outputs = @(Get-ChildItem -LiteralPath $WorkRoot -Filter 'output.*' | Where-Object Length -gt 0).Count
    Csv = [bool]$csv
    Builder = [string]$builder.DataSource
    JsonAsyncInterfacesLocation = $asyncDisposable.Assembly.Location
    ModuleLibrary = $moduleLibrary
} | ConvertTo-Json -Compress
'@
        $scriptPath = Join-Path $workRoot 'coexistence.ps1'
        Set-Content -LiteralPath $scriptPath -Value $script -Encoding UTF8
        $output = & $windowsPowerShell -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $scriptPath `
            -ModuleManifest $moduleManifest -DbatoolsRoot $dbatoolsRoot -Order "$First,$Second" -WorkRoot $workRoot 2>&1
        $LASTEXITCODE | Should -Be 0 -Because ($output -join [Environment]::NewLine)

        $json = $output | Where-Object { $_ -is [string] -and $_.StartsWith('{') } | Select-Object -Last 1
        $json | Should -Not -BeNullOrEmpty -Because ($output -join [Environment]::NewLine)
        $result = $json | ConvertFrom-Json
        $result.Outputs | Should -Be 3
        $result.Csv | Should -BeTrue
        $result.Builder | Should -Be 'localhost'
        [IO.Path]::GetDirectoryName($result.JsonAsyncInterfacesLocation) | Should -Be $result.ModuleLibrary -Because 'System.Text.Json must bind to the module''s Microsoft.Bcl.AsyncInterfaces, not an older copy another module loaded first'
    }
}
