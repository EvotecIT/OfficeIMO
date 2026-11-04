BeforeAll {
    $manifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $manifest -Global -ErrorAction Stop
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
}

Describe 'PowerPoint document lifecycle through the module' {
    It 'keeps a new document in memory until explicitly saved' {
        $path = Join-Path (Join-Path $TestDrive 'created-on-save') 'Review.pptx'
        $document = New-OfficePowerPoint -Path $path
        try {
            PowerPointSlide -Document $document -Title 'Saved explicitly'
            Test-Path -LiteralPath $path | Should -BeFalse
            $file = $document | Save-OfficePowerPoint -PassThru
            $file | Should -BeOfType ([System.IO.FileInfo])
            $file.FullName | Should -Be $path
            PowerPointSlide -Document $document -Title 'Unsaved edit'
        } finally {
            $output = @($document | Close-OfficePowerPoint)
            $output | Should -HaveCount 0
        }
        $saved = @(Get-OfficePowerPointSlide -LiteralPath $path -Summary)
        $saved | Should -HaveCount 1
        $saved[0].Title | Should -Be 'Saved explicitly'
    }

    It 'saves an externally created native document and keeps it open for further editing' {
        $path = Join-Path (Join-Path $TestDrive 'external-copy') 'External.pptx'
        $seed = New-OfficePowerPoint
        try {
            $documentType = $seed.GetType()
            $document = $documentType::Create()
        } finally {
            Close-OfficePowerPoint -Document $seed
        }
        try {
            PowerPointSlide -Document $document -Title 'External document'
            $file = Save-OfficePowerPoint -Document $document -Path $path -PassThru
            $file.FullName | Should -Be $path
            PowerPointSlide -Document $document -Title 'Still editable'
            Save-OfficePowerPoint -Document $document
        } finally {
            Close-OfficePowerPoint -Document $document
        }
        @(Get-OfficePowerPointSlide -LiteralPath $path -Summary) | Should -HaveCount 2
    }

    It 'supports literal file paths and FileInfo pipeline input when opening a document' {
        $path = Join-Path $TestDrive 'Review[final].pptx'
        $file = New-OfficePowerPoint -Path $path -PassThru {
            PowerPointSlide -Title 'Literal path'
        }
        $file | Should -BeOfType ([System.IO.FileInfo])
        $document = $file | Open-OfficePowerPoint
        try {
            $document.Slides.Count | Should -Be 1
            (Get-OfficePowerPointSlide -Document $document -Slide 1 -Summary).Title | Should -Be 'Literal path'
        } finally {
            Close-OfficePowerPoint -Document $document
        }
        $stream = [System.IO.File]::Open($path, 'Open', 'ReadWrite', 'None')
        $stream.Dispose()
    }

    It 'does not execute authoring or write a file under WhatIf' {
        $path = Join-Path $TestDrive 'WhatIf.pptx'
        $script:entered = $false
        New-OfficePowerPoint -Path $path -WhatIf {
            $script:entered = $true
            PowerPointSlide -Title 'Not created'
        }
        $script:entered | Should -BeFalse
        Test-Path -LiteralPath $path | Should -BeFalse
    }

    It 'keeps the original destination when a forced authoring block fails' {
        $path = Join-Path $TestDrive 'Existing.pptx'
        New-OfficePowerPoint -Path $path { PowerPointSlide -Title 'Original' }
        $original = [System.IO.File]::ReadAllBytes($path)
        {
            New-OfficePowerPoint -Path $path -Force {
                PowerPointSlide -Title 'Incomplete edit'
                throw 'authoring-failed'
            }
        } | Should -Throw '*authoring-failed*'
        [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($path)) |
            Should -Be ([Convert]::ToBase64String($original))
    }

    It 'preserves loaded encryption when saved without a password override' {
        $path = Join-Path $TestDrive 'Encrypted.pptx'
        $copy = Join-Path $TestDrive 'Encrypted-copy.pptx'
        $password = ConvertTo-SecureString 'synthetic-test-password' -AsPlainText -Force
        New-OfficePowerPoint -Path $path -Password $password {
            PowerPointSlide -Title 'Protected source'
        }
        $document = Open-OfficePowerPoint -LiteralPath $path -Password $password
        try {
            PowerPointSlide -Document $document -Title 'Protected addition'
            Save-OfficePowerPoint -Document $document -Path $copy
        } finally {
            Close-OfficePowerPoint -Document $document
        }
        { Get-ZipEntriesLocal -Path $copy } | Should -Throw
        $saved = Open-OfficePowerPoint -LiteralPath $copy -Password $password
        try {
            $saved.Slides.Count | Should -Be 2
            (Get-OfficePowerPointSlide -Document $saved -Slide 2 -Summary).Title | Should -Be 'Protected addition'
        } finally {
            Close-OfficePowerPoint -Document $saved
            $password.Dispose()
        }
    }

    It 'accepts a credential password override when saving an encrypted document' {
        $path = Join-Path $TestDrive 'Old-password.pptx'
        $copy = Join-Path $TestDrive 'New-password.pptx'
        $oldPassword = ConvertTo-SecureString 'synthetic-old-password' -AsPlainText -Force
        $newPassword = ConvertTo-SecureString 'synthetic-new-password' -AsPlainText -Force
        $credential = [pscredential]::new('unused-test-identity', $newPassword)
        New-OfficePowerPoint -Path $path -Password $oldPassword {
            PowerPointSlide -Title 'Password change'
        }
        $document = Open-OfficePowerPoint -LiteralPath $path -Password $oldPassword
        try {
            Save-OfficePowerPoint -Document $document -Path $copy -Password $credential
        } finally {
            Close-OfficePowerPoint -Document $document
        }
        { Open-OfficePowerPoint -LiteralPath $copy -Password $oldPassword -ErrorAction Stop } | Should -Throw
        $saved = Open-OfficePowerPoint -LiteralPath $copy -Password $newPassword
        try {
            (Get-OfficePowerPointSlide -Document $saved -Slide 1 -Summary).Title | Should -Be 'Password change'
        } finally {
            Close-OfficePowerPoint -Document $saved
            $oldPassword.Dispose()
            $newPassword.Dispose()
        }
    }

    It 'preserves encrypted source files that differ only by case on a case-sensitive filesystem' {
        $path = Join-Path $TestDrive 'CaseProbe'
        [System.IO.File]::WriteAllText($path, 'probe')
        $caseSensitive = -not (Test-Path -LiteralPath (Join-Path $TestDrive 'caseprobe'))
        Remove-Item -LiteralPath $path
        if (-not $caseSensitive) {
            Set-ItResult -Skipped -Because 'The test filesystem is case insensitive.'
            return
        }
        $source = Join-Path $TestDrive 'Deck.pptx'
        $copy = Join-Path $TestDrive 'deck.pptx'
        $password = ConvertTo-SecureString 'synthetic-case-password' -AsPlainText -Force
        New-OfficePowerPoint -Path $source -Password $password {
            PowerPointSlide -Title 'Case-sensitive source'
        }
        $original = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($source))
        $document = Open-OfficePowerPoint -LiteralPath $source -Password $password
        try {
            Save-OfficePowerPoint -Document $document -Path $copy
        } finally {
            Close-OfficePowerPoint -Document $document
        }
        [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($source)) | Should -Be $original
        $saved = Open-OfficePowerPoint -LiteralPath $copy -Password $password
        try {
            $saved.Slides.Count | Should -Be 1
        } finally {
            Close-OfficePowerPoint -Document $saved
            $password.Dispose()
        }
    }
}
