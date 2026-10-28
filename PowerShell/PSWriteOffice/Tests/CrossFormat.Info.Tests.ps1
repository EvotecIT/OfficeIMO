BeforeAll {
    $manifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) { $env:PSWRITEOFFICE_MODULE_MANIFEST } else { Join-Path $PSScriptRoot '../PSWriteOffice.psd1' }
    Import-Module $manifest -Global -ErrorAction Stop
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
}

Describe 'Cross-format document information' {
    It 'inspects literal paths and FileInfo pipelines without losing empty sheets' {
        $path = Join-Path $TestDrive 'inventory[1].xlsx'
        New-OfficeExcel -Path $path {
            ExcelSheet 'Data' { ExcelCell -Address A1 -Value 'Inventory' }
            ExcelSheet 'Empty' { }
        } | Out-Null
        $before = (Get-FileHash -LiteralPath $path).Hash
        $info = Get-Item -LiteralPath $path | Get-OfficeDocumentInfo -Detailed
        $info.Format.ToString() | Should -Be Excel
        $info.SheetCount | Should -Be 2
        $info.Details.Worksheets.Count | Should -Be 2
        $info.IsEncrypted | Should -BeFalse
        $info.Modified | Should -Not -BeNullOrEmpty
        (Get-OfficeDocumentInfo -LiteralPath $path).SheetCount | Should -Be 2
        (Get-FileHash -LiteralPath $path).Hash | Should -Be $before
    }

    It 'inspects unsaved native edits after the associated file is removed' {
        $path = Join-Path $TestDrive 'live.xlsx'
        New-OfficeExcel -Path $path { ExcelSheet 'Saved' { } } | Out-Null
        $workbook = Get-OfficeExcel -Path $path
        try {
            $workbook | Add-OfficeExcelSheet -Name Unsaved
            Remove-Item -LiteralPath $path
            $info = $workbook | Get-OfficeDocumentInfo -Detailed
            $info.SheetCount | Should -Be 2
            $info.Details.Worksheets.Name | Should -Be @('Saved', 'Unsaved')
            $info.Modified | Should -BeNullOrEmpty
            $workbook | Add-OfficeExcelSheet -Name 'Still open'
            $workbook.Sheets.Count | Should -Be 3
        } finally { $workbook | Close-OfficeExcel }
    }

    It 'identifies a package from its contents when its extension is wrong' {
        $source = Join-Path $TestDrive 'source.docx'
        $attachment = Join-Path $TestDrive 'attachment.bin'
        New-OfficeWord -Path $source { WordSection { WordParagraph -Text 'Native document' } } | Out-Null
        Copy-Item -LiteralPath $source -Destination $attachment
        $info = Get-OfficeDocumentInfo $attachment -Detailed
        $info.Format.ToString() | Should -Be Word
        $info.DetectedFormat.ToString() | Should -Be Word
        $info.Confidence.ToString() | Should -Be High
        $info.Details.GetType().FullName | Should -Be 'OfficeIMO.Word.WordDocumentInfo'
    }

    It 'inventories locked Office content and opens native details only with the password' {
        $path = Join-Path $TestDrive 'encrypted.docx'
        New-OfficeWord -Path $path -Password 'Synthetic-inspection-password' {
            WordSection { WordParagraph -Text 'Protected content' }
        } | Out-Null
        $before = (Get-FileHash -LiteralPath $path).Hash
        $locked = Get-OfficeDocumentInfo $path -Detailed
        $locked.IsEncrypted | Should -BeTrue
        $locked.PageCount | Should -BeNullOrEmpty
        $locked.Details | Should -BeNullOrEmpty
        $password = ConvertTo-SecureString 'Synthetic-inspection-password' -AsPlainText -Force
        $opened = Get-OfficeDocumentInfo $path -Password $password -Detailed
        $opened.IsEncrypted | Should -BeTrue
        $opened.Details.GetType().FullName | Should -Be 'OfficeIMO.Word.WordDocumentInfo'
        (Get-FileHash -LiteralPath $path).Hash | Should -Be $before
    }

    It 'reports an unreadable source independently and returns the valid inventory row' {
        $valid = Join-Path $TestDrive 'valid.xlsx'
        $broken = Join-Path $TestDrive 'broken.docx'
        New-OfficeExcel -Path $valid { ExcelSheet 'Data' { } } | Out-Null
        Set-Content -LiteralPath $broken -Value 'not an Office package'
        $failures = @()
        $rows = @(Get-OfficeDocumentInfo -Path @($valid, $broken) -ErrorVariable +failures -ErrorAction SilentlyContinue)
        $rows | Should -HaveCount 1
        $rows[0].SheetCount | Should -Be 1
        $failures | Should -HaveCount 1
    }

    It 'honors inspection hashtables and skips container probing when requested' {
        $path = Join-Path $TestDrive 'bounded.xlsx'
        New-OfficeExcel -Path $path { ExcelSheet 'Data' { } } | Out-Null
        { Get-OfficeDocumentInfo $path -Options @{ MaximumInputBytes = 1 } -ErrorAction Stop } | Should -Throw
        $info = Get-OfficeDocumentInfo $path -NoContainerInspection
        $info.Detection.ContainerInspected | Should -BeFalse
        $info.Detection.Evidence | Should -Not -Contain 'container:xl/workbook.xml'
        { Get-OfficeDocumentInfo $path -Options @{ UnknownInspectionSetting = 1 } -ErrorAction Stop } | Should -Throw '*Unknown*'
    }

    It 'preserves Reader detection defaults and applies only explicit detection overrides' {
        $source = Join-Path $TestDrive 'policy-source.docx'
        $path = Join-Path $TestDrive 'policy-source.txt'
        New-OfficeWord -Path $source { WordSection { WordParagraph -Text 'Detection profile' } } | Out-Null
        Copy-Item -LiteralPath $source -Destination $path
        $reader = New-OfficeDocumentReader -Options @{
            DetectionMode = 'ExtensionOnly'; DetectionInspectContainers = $false
            DetectionMaxProbeBytes = 17; DetectionMaxContainerEntries = 2
        }
        $inherited = Get-OfficeDocumentInfo $path -Reader $reader
        $inherited.Detection.ContentInspected | Should -BeFalse
        $inherited.Detection.InspectedBytes | Should -Be 0
        $overridden = Get-OfficeDocumentInfo $path -Reader $reader -DetectionMode PreferContent -MaximumProbeBytes 64
        $overridden.Detection.ContentInspected | Should -BeTrue
        $overridden.Detection.InspectedBytes | Should -BeLessOrEqual 64
        $overridden.Detection.ContainerInspected | Should -BeFalse
        $reader.DefaultOptions.DetectionMode.ToString() | Should -Be ExtensionOnly
        $reader.DefaultOptions.DetectionMaxProbeBytes | Should -Be 17
    }

    It 'uses one batch password for encrypted and ordinary Office documents' {
        $plain = Join-Path $TestDrive 'plain-inventory.xlsx'
        $locked = Join-Path $TestDrive 'locked-inventory.docx'
        New-OfficeExcel -Path $plain { ExcelSheet 'Data' { } } | Out-Null
        New-OfficeWord -Path $locked -Password 'Synthetic-batch-password' {
            WordSection { WordParagraph -Text 'Protected inventory' }
        } | Out-Null
        $password = ConvertTo-SecureString 'Synthetic-batch-password' -AsPlainText -Force
        $rows = @(Get-OfficeDocumentInfo -Path @($plain, $locked) -Password $password -Detailed -ErrorAction Stop)
        $rows | Should -HaveCount 2
        ($rows | Where-Object Format -EQ Excel).SheetCount | Should -Be 1
        ($rows | Where-Object Format -EQ Excel).IsEncrypted | Should -BeFalse
        ($rows | Where-Object Format -EQ Word).IsEncrypted | Should -BeTrue
        ($rows | Where-Object Format -EQ Word).Details | Should -Not -BeNullOrEmpty
    }
}
