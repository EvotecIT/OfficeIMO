# The test-only cmdlet assembly targets net8.0; native Desktop qualification uses its own supported payload.
Describe 'PowerShell 7 foundation fixture' -Skip:($PSEdition -ne 'Core') {
BeforeAll {
    $moduleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) { $env:PSWRITEOFFICE_MODULE_MANIFEST } else { Join-Path $PSScriptRoot '../PSWriteOffice.psd1' }
    Import-Module $moduleManifest -Global -ErrorAction Stop
    $binary = (Get-Command New-OfficeWord -ErrorAction Stop).ImplementingType.Assembly
    $testAssemblyPath = if ($env:PSWRITEOFFICE_FOUNDATION_TEST_ASSEMBLY) { $env:PSWRITEOFFICE_FOUNDATION_TEST_ASSEMBLY } else {
        Join-Path $PSScriptRoot '../Sources/PSWriteOffice.Tests/bin/Release/net8.0/PSWriteOffice.Tests.dll'
    }
    if (!(Test-Path -LiteralPath $testAssemblyPath)) { throw 'Build PSWriteOffice.Tests before running the foundation Pester tests.' }
    if ($PSEdition -eq 'Core') {
        $context = [System.Runtime.Loader.AssemblyLoadContext]::GetLoadContext($binary)
        $testAssembly = $context.LoadFromAssemblyPath([IO.Path]::GetFullPath($testAssemblyPath))
    } else { $testAssembly = [Reflection.Assembly]::LoadFrom([IO.Path]::GetFullPath($testAssemblyPath)) }
    Import-Module -Assembly $testAssembly -Global -ErrorAction Stop
}

Describe 'Shared parameter transformations' {
    It 'binds unquoted lengths with invariant units and numeric points' {
        (Test-OfficeFoundationBinding -Width 5cm).Width.ToEmus() | Should -Be 1800000
        (Test-OfficeFoundationBinding -Width 12pt).Width.ToEmus() | Should -Be 152400
        (Test-OfficeFoundationBinding -Width 1in).Width.ToEmus() | Should -Be 914400
        (Test-OfficeFoundationBinding -Width 10mm).Width.ToEmus() | Should -Be 360000
        (Test-OfficeFoundationBinding -Width 96px).Width.ToEmus() | Should -Be 914400
        (Test-OfficeFoundationBinding -Width 914400emu).Width.ToEmus() | Should -Be 914400
        (Test-OfficeFoundationBinding -Width 12).Width.ToEmus() | Should -Be 152400
        { Test-OfficeFoundationBinding -Width '2,5cm' } | Should -Throw
    }
    It 'accepts hex, CSS colors, roles and typed options, rejecting unknown keys' {
        (Test-OfficeFoundationBinding -Color '#12345678').Color.ToString() | Should -Be '#12345678'
        (Test-OfficeFoundationBinding -Color RebeccaPurple).Color.ToString() | Should -Be '#663399'
        (Test-OfficeFoundationBinding -Color Accent1).Color.Role.ToString() | Should -Be Accent1
        (Test-OfficeFoundationBinding -Options @{Title='Report';Width='5cm'}).Options.Title | Should -Be Report
        { Test-OfficeFoundationBinding -Options @{Typo=4} } | Should -Throw
    }
    It 'accepts a theme object, a name, and a literal template path' {
        (Test-OfficeFoundationBinding -Theme Modern).Theme.Theme.Name | Should -Be Modern
        (Test-OfficeFoundationBinding -Theme ([OfficeIMO.OfficeTheme]::Dark)).Theme.Theme.Name | Should -Be Dark
        $template = Join-Path $TestDrive 'theme[1].dotx'; Set-Content -LiteralPath $template -Value fixture
        (Test-OfficeFoundationBinding -Theme $template).Theme.TemplatePath | Should -Be $template
    }
    It 'warns for a plain password and accepts SecureString and PSCredential without warnings' {
        $records = @(& { Test-OfficeFoundationBinding -Password 'test-only' } 3>&1)
        $warnings = @($records | Where-Object { $_ -is [System.Management.Automation.WarningRecord] })
        $result = @($records | Where-Object { $_ -isnot [System.Management.Automation.WarningRecord] })
        $result[0].PasswordLength | Should -Be 9
        $warnings.Count | Should -Be 1
        [string]$warnings[0] | Should -Not -Match 'test-only'
        $password = ConvertTo-SecureString 'test-only' -AsPlainText -Force
        try {
            $warnings = @()
            (Test-OfficeFoundationBinding -Password $password -WarningVariable warnings).PasswordLength | Should -Be 9
            (Test-OfficeFoundationBinding -Password ([pscredential]::new('user', $password)) -WarningVariable warnings).PasswordLength | Should -Be 9
            $warnings.Count | Should -Be 0
        } finally { $password.Dispose() }
    }
    It 'preserves sheet names, one-based positions, format enums and tri-state switches' {
        (Test-OfficeFoundationBinding -Sheet '1').Sheet.Name | Should -Be '1'
        (Test-OfficeFoundationBinding -Sheet 1).Sheet.Position | Should -Be 1
        { Test-OfficeFoundationBinding -Sheet 0 } | Should -Throw
        (Test-OfficeFoundationBinding).Bold | Should -BeNullOrEmpty
        (Test-OfficeFoundationBinding -Bold:$false).Bold | Should -BeFalse
        (Test-OfficeFoundationBinding -Bold).Bold | Should -BeTrue
        (Test-OfficeFoundationBinding -NoFreeze).Freeze | Should -BeFalse
        (Test-OfficeFoundationBinding -NoFreeze:$false).Freeze | Should -BeTrue
        (Test-OfficeFoundationBinding -To Pdf).To.ToString() | Should -Be Pdf
        { Test-OfficeFoundationBinding -To Unknown } | Should -Throw
    }
    It 'completes shared length, color, theme and format values' {
        foreach ($case in @(@('Width','5','5cm'), @('Color','Acc','Accent1'), @('Theme','Bo','Boardroom'), @('To','Pd','Pdf'))) {
            $line = "Test-OfficeFoundationBinding -$($case[0]) $($case[1])"
            $completion = TabExpansion2 $line $line.Length
            $completion.CompletionMatches.CompletionText | Should -Contain $case[2]
        }
    }
}

Describe 'Provider paths, output policy and errors' {
    It 'resolves wildcard input and binds FileInfo FullName literally, including brackets' {
        $path = Join-Path $TestDrive 'file[1].txt'; Set-Content -LiteralPath $path -Value fixture
        @(Test-OfficeFoundationPath -Path (Join-Path $TestDrive '*.txt')) | Should -Contain $path
        @(Get-Item -LiteralPath $path | Test-OfficeFoundationPath) | Should -Be @($path)
        @([pscustomobject]@{PSPath="FileSystem::$path"} | Test-OfficeFoundationPath) | Should -Be @($path)
        { Test-OfficeFoundationPath -Path $path -LiteralPath $path } | Should -Throw
        @(Test-OfficeFoundationPath -LiteralPath $path,$TestDrive -IncludeDirectory) | Should -Be @($path,$TestDrive)
    }
    It 'keeps valid wildcard matches when another matching item is a directory' {
        $folder = Join-Path $TestDrive 'mixed'; $null = New-Item -ItemType Directory -Path $folder
        $valid = Join-Path $folder 'valid.txt'; Set-Content -LiteralPath $valid -Value fixture
        $invalid = Join-Path $folder 'nested'; $null = New-Item -ItemType Directory -Path $invalid
        $errors = @()
        @(Test-OfficeFoundationPath -Path (Join-Path $folder '*') -ErrorAction SilentlyContinue -ErrorVariable errors) | Should -Be @($valid)
        $errors.Count | Should -Be 1
        $errors[0].TargetObject | Should -Be $invalid
        $errors[0].FullyQualifiedErrorId | Should -Match '^input-unreadable,'
    }
    It 'rejects a directory OutputPath and mutually exclusive output inputs' {
        { Export-OfficeFoundation -OutputPath $TestDrive -ErrorAction Stop } | Should -Throw
        { Export-OfficeFoundation -OutputPath a.txt -OutputDirectory $TestDrive } | Should -Throw
        { Export-OfficeFoundation -OutputPath 'Env:HOME' -ErrorAction Stop } | Should -Throw
    }
    It 'supports output directory, NoClobber and WhatIf without effects' {
        $file = Export-OfficeFoundation -OutputDirectory $TestDrive
        $file | Should -BeOfType ([IO.FileInfo])
        Set-Content -LiteralPath $file.FullName -Value original
        Export-OfficeFoundation -OutputPath $file.FullName -NoClobber -ErrorAction SilentlyContinue
        (Get-Content -LiteralPath $file.FullName -Raw).Trim() | Should -Be original
        $target = Join-Path $TestDrive 'whatif.txt'
        Export-OfficeFoundation -OutputPath $target -WhatIf
        Test-Path -LiteralPath $target | Should -BeFalse
    }
    It 'continues failed input items with a stable diagnostic code, category and target' {
        $missing = Join-Path $TestDrive 'missing.txt'
        $present = Join-Path $TestDrive 'present.txt'; Set-Content -LiteralPath $present -Value fixture
        $errors = @()
        $result = @(Test-OfficeFoundationPath -LiteralPath $missing,$present -ErrorAction SilentlyContinue -ErrorVariable errors)
        $result | Should -Be @($present)
        $errors.Count | Should -Be 1
        $errors[0].FullyQualifiedErrorId | Should -Match '^input-unreadable,'
        $errors[0].TargetObject | Should -Be $missing
        $errors[0].CategoryInfo.Category | Should -Be ObjectNotFound
    }
}

Describe 'Content lifecycle and implicit document binding' {
    It 'saves and closes a successful block, returning FileInfo only with PassThru' {
        $path = Join-Path $TestDrive 'new.txt'
        $output = @(New-OfficeFoundationDocument -Path $path { Add-OfficeFoundationText 'first'; Add-OfficeFoundationText 'second' })
        $output.Count | Should -Be 0
        Get-Content -LiteralPath $path -Raw | Should -Be 'firstsecond'
        $file = Open-OfficeFoundationDocument -LiteralPath $path -PassThru { Add-OfficeFoundationText 'third' }
        $file | Should -BeOfType ([IO.FileInfo])
        Get-Content -LiteralPath $path -Raw | Should -Be 'firstsecondthird'
    }
    It 'returns an unsaved live document with no block and accepts Document from the pipeline' {
        $document = New-OfficeFoundationDocument
        try {
            $document | Add-OfficeFoundationText 'explicit'
            $document.Text | Should -Be explicit
            $document.Disposed | Should -BeFalse
        } finally { $document.Dispose() }
        { Add-OfficeFoundationText 'outside' -ErrorAction Stop } | Should -Throw '*Supply -Document*'
    }
    It 'removes partial new files and preserves forced replacement and open inputs when blocks throw' {
        $new = Join-Path $TestDrive 'failed.txt'
        { New-OfficeFoundationDocument -Path $new { throw 'block failure' } } | Should -Throw '*block failure*'
        Test-Path -LiteralPath $new | Should -BeFalse
        $existing = Join-Path $TestDrive 'existing.txt'; Set-Content -LiteralPath $existing -Value original -NoNewline
        { New-OfficeFoundationDocument -Path $existing -Force { throw 'replacement failure' } } | Should -Throw
        Get-Content -LiteralPath $existing -Raw | Should -Be original
        { Open-OfficeFoundationDocument -LiteralPath $existing { Add-OfficeFoundationText 'changed'; throw 'open failure' } } | Should -Throw
        Get-Content -LiteralPath $existing -Raw | Should -Be original
        @(Get-ChildItem -LiteralPath $TestDrive -Force -File | Where-Object Name -Like '.*').Count | Should -Be 0
    }
    It 'refuses New overwrite without Force and fails native loss unless AllowLoss' {
        $path = Join-Path $TestDrive 'loss.txt'; Set-Content -LiteralPath $path -Value original -NoNewline
        { New-OfficeFoundationDocument -Path $path {} } | Should -Throw '*already exists*'
        { New-OfficeFoundationDocument -Path $path -Force { Add-OfficeFoundationText 'loss' -Loss } } | Should -Throw '*fidelity loss*'
        Get-Content -LiteralPath $path -Raw | Should -Be original
        New-OfficeFoundationDocument -Path $path -Force -AllowLoss { Add-OfficeFoundationText 'accepted' -Loss }
        Get-Content -LiteralPath $path -Raw | Should -Be accepted
    }
    It 'reports native loss by diagnostic code while preserving block failures' {
        $path = Join-Path $TestDrive 'coded-loss.txt'
        try {
            New-OfficeFoundationDocument -Path $path -ErrorAction Stop { Add-OfficeFoundationText 'loss' -Loss }
            throw 'Expected a native-save loss failure'
        } catch {
            $_.FullyQualifiedErrorId | Should -Match '^chart-flattened,'
            $_.CategoryInfo.Category | Should -Be InvalidData
            $_.TargetObject | Should -Be $path
        }
        Test-Path -LiteralPath $path | Should -BeFalse
        { New-OfficeFoundationDocument -Path $path { throw 'original authoring failure' } } | Should -Throw '*original authoring failure*'
    }
    It 'does not invoke content under WhatIf and does not prompt per content item' {
        $path = Join-Path $TestDrive 'skip.txt'
        New-OfficeFoundationDocument -Path $path -WhatIf { throw 'must not run' }
        Test-Path -LiteralPath $path | Should -BeFalse
        (Get-Command Add-OfficeFoundationText).Parameters.ContainsKey('WhatIf') | Should -BeFalse
        (Get-Command New-OfficeFoundationDocument).Parameters.ContainsKey('Confirm') | Should -BeTrue
    }
}

Describe 'Bounded background execution' {
    It 'streams completion order, continues per-item errors, and writes progress on the pipeline thread' {
        $slow = Join-Path $TestDrive 'slow.txt'; Set-Content -LiteralPath $slow -Value slow -NoNewline
        $fast = Join-Path $TestDrive 'fast.txt'; Set-Content -LiteralPath $fast -Value fast -NoNewline
        $bad = Join-Path $TestDrive 'bad.txt'; Set-Content -LiteralPath $bad -Value fail -NoNewline
        $errors = @()
        $results = @(Test-OfficeFoundationBatch -LiteralPath $slow,$fast,$bad -ThrottleLimit 2 -ErrorVariable errors -ErrorAction SilentlyContinue)
        $results | Should -Be @('fast','slow')
        $errors.Count | Should -Be 1
        $errors[0].TargetObject | Should -Be $bad
        $errors[0].FullyQualifiedErrorId | Should -Match '^input-invalid,'
    }
}

}
