BeforeAll {
    $manifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $manifest -Global -ErrorAction Stop
}

Describe 'Retired PowerShell command families' {
    It 'removes electronic invoice commands and points callers to the native packages' {
        $commands = @(Get-Command -Module PSWriteOffice -CommandType Cmdlet, Function, Alias |
            Where-Object { $_.Name -match 'OfficeInvoice|OfficePdfElectronicInvoice' })
        $commands.Count | Should -Be 0
        $records = @(& {
            $PSModuleAutoLoadingPreference = 'None'
            try { New-OfficeInvoice -ErrorAction Stop } catch { $_ }
        } 3>&1)
        $errors = @($records | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] })
        $errors.Count | Should -Be 1
        $errors[0].FullyQualifiedErrorId | Should -Be 'CommandNotFoundException'
        $warnings = @($records | Where-Object { $_ -is [System.Management.Automation.WarningRecord] })
        ($warnings -join "`n") | Should -Match 'OfficeIMO\.Invoicing'
    }
    It 'uses slide summaries through Get-OfficePowerPointSlide and explains the retired name' {
        Get-Command Get-OfficePowerPointSlideSummary -ErrorAction SilentlyContinue | Should -BeNullOrEmpty
        $records = @(& {
            $PSModuleAutoLoadingPreference = 'None'
            try { Get-OfficePowerPointSlideSummary -ErrorAction Stop } catch { $_ }
        } 3>&1)
        $errors = @($records | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] })
        $errors.Count | Should -Be 1
        $errors[0].FullyQualifiedErrorId | Should -Be 'CommandNotFoundException'
        $warnings = @($records | Where-Object { $_ -is [System.Management.Automation.WarningRecord] })
        ($warnings -join "`n") | Should -Match 'Get-OfficePowerPointSlide.*-Summary'
    }
    It 'merges <OldCommand> into the slide editing command' -TestCases @(
        @{ OldCommand = 'Set-OfficePowerPointSlideTitle' },
        @{ OldCommand = 'Set-OfficePowerPointSlideLayout' },
        @{ OldCommand = 'Set-OfficePowerPointSlideTransition' },
        @{ OldCommand = 'Set-OfficePowerPointBackground' }
    ) {
        param($OldCommand)
        Get-Command $OldCommand -ErrorAction SilentlyContinue | Should -BeNullOrEmpty
        $records = @(& {
            $PSModuleAutoLoadingPreference = 'None'
            try { & $OldCommand -ErrorAction Stop } catch { $_ }
        } 3>&1)
        $errors = @($records | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] })
        $errors.Count | Should -Be 1
        $errors[0].FullyQualifiedErrorId | Should -Be 'CommandNotFoundException'
        ($records | Where-Object { $_ -is [System.Management.Automation.WarningRecord] }) -join "`n" |
            Should -Match 'Set-OfficePowerPointSlide'
    }
}
