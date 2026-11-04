BeforeAll {
    $modulePath = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '../PSWriteOffice.psd1'
    }
    Import-Module $modulePath -Global -ErrorAction Stop
}

Describe 'Native nonExcel mutation decisions' {
    It 'keeps PDF extraction read-only under WhatIf while preserving object inspection' {
        $source = Join-Path $TestDrive 'extract.pdf'
        $attachment = Join-Path $TestDrive 'note.txt'
        [IO.File]::WriteAllText($attachment, 'Attached evidence')
        New-OfficePdf -Path $source {
            PdfImage -Path (Join-Path $PSScriptRoot 'Assets/CellImage.png') -Width 20 -Height 20
            Add-OfficePdfAttachment -Path $attachment
        }
        foreach ($command in 'Get-OfficePdfImage', 'Get-OfficePdfAttachment') {
            $directory = Join-Path $TestDrive $command
            @(& $command -Path $source).Count | Should -BeGreaterThan 0
            @(& $command -Path $source -OutputDirectory $directory -WhatIf).Count | Should -Be 0
            Test-Path -LiteralPath $directory | Should -BeFalse
            @(& $command -Path $source -OutputDirectory $directory -Confirm:$false).Count | Should -BeGreaterThan 0
            @(Get-ChildItem -LiteralPath $directory -File).Count | Should -BeGreaterThan 0
        }
    }

    It 'does not create the Visio gallery output directory under native WhatIf' {
        $directory = Join-Path $TestDrive 'gallery-declined'
        @(New-OfficeVisioGallery -OutputDirectory $directory -WhatIf).Count | Should -Be 0
        Test-Path -LiteralPath $directory | Should -BeFalse
    }

    It 'keeps a no-save Word document live and tracked when WhatIf refuses closure' {
        $document = New-OfficeWord
        try {
            $document | Add-OfficeWordParagraph -Text 'Unsaved work'
            Close-OfficeWord -Document $document -WhatIf
            $document.IsDisposed | Should -BeFalse
            Close-OfficeWord -WhatIf
            $document.IsDisposed | Should -BeFalse
            $document | Add-OfficeWordParagraph -Text 'Still editable'
            Close-OfficeWord -Confirm:$false
            $document.IsDisposed | Should -BeTrue
        } finally {
            if (-not $document.IsDisposed) { Close-OfficeWord -Document $document -Confirm:$false }
        }
    }
}

Describe 'NonExcel source and rich content contracts' {
    It 'returns each newly appended Word run and edits that run through the pipeline' {
        $document = New-OfficeWord
        try {
            $paragraph = Add-OfficeWordParagraph -Document $document -Text 'Existing' -PassThru
            $runs = @(Add-OfficeWordText -Paragraph $paragraph -Text 'New one', 'New two' -Bold -Italic -PassThru)
            $runs.Count | Should -Be 2
            $runs[0].Text | Should -Be 'New one'
            $runs[1].Text | Should -Be 'New two'
            $runs[0] | Set-OfficeWordText -Text 'Replacement' -Bold:$false -Italic:$false
            $paragraph.Text | Should -Be 'Existing'
            $runs[0].Text | Should -Be 'Replacement'
            $runs[0].Bold | Should -BeFalse
            $runs[0].Italic | Should -BeFalse
            $runs[1].Bold | Should -BeTrue
            $runs[1].Italic | Should -BeTrue
        } finally { Close-OfficeWord -Document $document -Confirm:$false }
    }

    It 'rejects conflicting rich inputs through native parameter binding' {
        $document = New-OfficeWord
        try {
            { Add-OfficeWordParagraph -Document $document -Text 'Plain' -Content { WordText 'Rich' } -ErrorAction Stop } |
                Should -Throw -ErrorId 'AmbiguousParameterSet,*'
            { Set-OfficePdfHeader -Text 'Plain' -Content { PdfText 'Rich' } -ErrorAction Stop } |
                Should -Throw -ErrorId 'AmbiguousParameterSet,*'
            { Set-OfficePdfFooter -Document $null -Text 'Dead route' -ErrorAction Stop } |
                Should -Throw -ErrorId 'NamedParameterNotFound,*'
            $empty = Add-OfficeWordParagraph -Document $document -PassThru
            $empty | Should -Not -BeNullOrEmpty
        } finally { Close-OfficeWord -Document $document -Confirm:$false }
    }

    It 'reads literal bracket PDF filenames and preserves a caller-owned live document' {
        $source = Join-Path $TestDrive 'request[1].pdf'
        New-OfficePdf -Path $source {
            PdfParagraph 'Request'
            PdfFormField -Name Requester -Value Ada
        }
        $document = Open-OfficePdf -LiteralPath $source
        @(Get-OfficePdfFormField -LiteralPath $source -Name Requester).Count | Should -Be 1
        @(Get-OfficePdfFormField -Document $document -Name Requester).Count | Should -Be 1
        @(Get-OfficePdfAnnotation -Document $document).Count | Should -BeGreaterThan 0
        @(Get-OfficePdfAnnotation -LiteralPath $source).Count | Should -BeGreaterThan 0
        $copy = Join-Path $TestDrive 'still-owned.pdf'
        Save-OfficePdf -Document $document -Path $copy
        Test-Path -LiteralPath $copy | Should -BeTrue
        Close-OfficePdf -Document $document
    }

    It 'parses text into each CSV output shape and gives reader ownership to its caller' {
        $text = "Name||Count`nAlpha||2`nBeta||3"
        $rows = @(Import-OfficeCsv -Text $text -Delimiter '||')
        $rows.Count | Should -Be 2
        $rows[0].Name | Should -Be Alpha
        (Import-OfficeCsv -Text $text -Delimiter '||' -As Hashtable | Select-Object -First 1)['Name'] | Should -Be Alpha
        $document = Import-OfficeCsv -Text $text -Delimiter '||' -As Document
        @(Import-OfficeCsv -Document $document).Count | Should -Be 2
        $table = Import-OfficeCsv -Text $text -Delimiter '||' -As DataTable
        $table.Rows.Count | Should -Be 2
        $reader = Import-OfficeCsv -Text $text -Delimiter '||' -As DataReader
        try {
            $reader.Read() | Should -BeTrue
            $reader.GetValue(0) | Should -Be Alpha
            $reader.Read() | Should -BeTrue
            $reader.GetValue(0) | Should -Be Beta
        } finally { $reader.Dispose() }
        { Import-OfficeCsv -Text $text -Document $document -ErrorAction Stop } | Should -Throw -ErrorId 'AmbiguousParameterSet,*'
    }

    It 'keeps all three ODF loss thresholds across Word Excel and PowerPoint' {
        foreach ($policy in 'ReportOnly', 'ThrowOnSkippedOrUnsupported', 'ThrowOnAnyLoss') {
            foreach ($command in 'New-OfficeWordOpenDocumentOptions', 'New-OfficeExcelOpenDocumentOptions', 'New-OfficePowerPointOpenDocumentOptions') {
                (& $command -LossPolicy $policy).LossPolicy.ToString() | Should -Be $policy
            }
        }
    }
}

Describe 'Native Google service-account credential validation' {
    It 'rejects an incomplete synthetic credential before token acquisition' {
        $key = Join-Path $TestDrive 'incomplete-service-account.json'
        [IO.File]::WriteAllText($key, '{"type":"service_account","client_email":"synthetic@example.invalid"}')
        { New-OfficeGoogleWorkspaceSession -KeyPath $key -Scope Documents -ErrorAction Stop } |
            Should -Throw '*missing private_key*'
    }
}

Describe 'Native PDF stamp permission and page selection' {
    It 'selects a page without demanding content-copy permission for an authorized stamp' {
        $source = Join-Path $TestDrive 'modify-without-copy.pdf'
        $output = Join-Path $TestDrive 'reviewed.pdf'
        $user = ConvertTo-SecureString 'synthetic-reader' -AsPlainText -Force
        $owner = ConvertTo-SecureString 'synthetic-owner' -AsPlainText -Force
        New-OfficePdf -Path $source -Password $user -OwnerPassword $owner -Permission ModifyContents {
            PdfParagraph 'First page'
            PdfPageBreak
            PdfParagraph 'Last page'
        }
        Add-OfficePdfStamp -LiteralPath $source -OutputPath $output -Text 'Reviewed' -Page last -Password $user -Opacity 0.5 -Rotation 15
        $inspection = Open-OfficePdf -LiteralPath $output -Password $owner
        try { $pages = @($inspection.Read().Pages | ForEach-Object { $_.TextBlocks.Text -join ' ' }) }
        finally { $inspection.Dispose() }
        $pages.Count | Should -Be 2
        $pages[0] | Should -Not -Match 'Reviewed'
        $pages[1] | Should -Match 'Reviewed'
    }
}

Describe 'Native PDF annotation selection and XFDF option capture' {
    It 'selects URI identities without copy permission and rejects unsafe residual updates' {
        $source = Join-Path $TestDrive 'annotation-source.pdf'
        $output = Join-Path $TestDrive 'annotation-reviewed.pdf'
        $user = ConvertTo-SecureString 'synthetic-reviewer' -AsPlainText -Force
        $owner = ConvertTo-SecureString 'synthetic-annotation-owner' -AsPlainText -Force
        New-OfficePdf -Path $source -Password $user -OwnerPassword $owner -Permission 'ModifyAnnotations' {
            PdfParagraph {
                PdfText 'Review link' -Link 'https://example.invalid/review'
                PdfText ' Keep link' -Link 'https://example.invalid/keep'
            }
        }
        { Set-OfficePdfAnnotation -LiteralPath $source -OutputPath $output -Uri 'https://example.invalid/rev*' -Contents Reviewed -Password $user -ErrorAction Stop } |
            Should -Throw '*Append-only annotation updates retain replaced annotation data*'
        Test-Path -LiteralPath $output | Should -BeFalse
        Set-OfficePdfAnnotation -LiteralPath $source -OutputPath $output -Uri 'https://example.invalid/rev*' -Contents Reviewed -Password $owner
        $annotations = @(Get-OfficePdfAnnotation -LiteralPath $output -Password $owner)
        ($annotations | Where-Object Uri -eq 'https://example.invalid/review').Contents | Should -Be Reviewed
        ($annotations | Where-Object Uri -eq 'https://example.invalid/keep').Contents | Should -Not -Be Reviewed
    }

    It 'captures typed XFDF settings before applying later hashtable overrides' {
        $source = Join-Path $TestDrive 'form[1].pdf'
        $directory = Join-Path $TestDrive 'completed'
        New-OfficePdf -Path $source { PdfFormField -Name Reviewer -Value Pending }
        $options = [OfficeIMO.Pdf.PdfFormFillerOptions]::new()
        $options.KeepNeedAppearances = $true
        $xml = '<?xml version="1.0"?><xfdf xmlns="http://ns.adobe.com/xfdf/"><fields><field name="Reviewer"><value>Ada</value></field></fields></xfdf>'
        Import-OfficePdfXfdf -LiteralPath $source -OutputDirectory $directory -Xfdf $xml -Options $options, @{ KeepNeedAppearances = $false }
        $options.KeepNeedAppearances | Should -BeTrue
        (Get-OfficePdfFormField -LiteralPath (Join-Path $directory 'form[1].pdf') -Name Reviewer).Value | Should -Be Ada
    }
}
