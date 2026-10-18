BeforeAll {
    $moduleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) { $env:PSWRITEOFFICE_MODULE_MANIFEST } else { Join-Path $PSScriptRoot '../PSWriteOffice.psd1' }
    Import-Module $moduleManifest -Global -ErrorAction Stop
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
}

Describe 'Excel worksheet editing' {
    It 'edits worksheet pipelines quietly and persists row and column layout' {
        $path = Join-Path $TestDrive 'WorksheetTargets.xlsx'
        $workbook = New-OfficeExcel -Path $path -NoSave
        try {
            $first = $workbook | Add-OfficeExcelSheet -Name 'First' -PassThru
            $second = $workbook | Add-OfficeExcelSheet -Name 'Second' -PassThru
            $output = @(@($first, $second) | Set-OfficeExcelRow -Row 1 -Values 'Name', 'Notes' -Height 30)
            $output.Count | Should -Be 0
            $output = @(@($first, $second) | Set-OfficeExcelColumn -ColumnName B -Width 42 -Hidden $false -PassThru)
            $output.Count | Should -Be 2
            [object]::ReferenceEquals($output[0], $first) | Should -BeTrue
            [object]::ReferenceEquals($output[1], $second) | Should -BeTrue
            $first | Set-OfficeExcelRow -Row 2 -Values 'Identity', 'Ready'
            $workbook | Save-OfficeExcel
        } finally {
            $workbook | Close-OfficeExcel
        }

        $rows = @(Import-OfficeExcel -Path $path -WorksheetName 'First')
        $rows[0].Name | Should -Be 'Identity'
        $rows[0].Notes | Should -Be 'Ready'
        foreach ($entry in 'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml') {
            $xml = Get-ZipXmlDocumentLocal -Path $path -Entry $entry
            $xml.SelectSingleNode("//*[local-name()='row' and @r='1']").GetAttribute('ht') | Should -Be '30'
            $xml.SelectSingleNode("//*[local-name()='col' and @min='2']").GetAttribute('width') | Should -Be '42'
        }
    }

    It 'selects named and indexed worksheets through workbook pipelines' {
        $path = Join-Path $TestDrive 'WorkbookTargets.xlsx'
        $workbook = New-OfficeExcel -Path $path -NoSave
        try {
            $workbook | Add-OfficeExcelSheet -Name 'Unchanged'
            $workbook | Add-OfficeExcelSheet -Name 'Review'
            $workbook | Set-OfficeExcelRow -Sheet 'Review' -Row 1 -Values 'Service', 'Status'
            $workbook | Set-OfficeExcelColumn -SheetIndex 1 -ColumnName A -StartRow 2 -Values 'Identity', 'Messaging'
            $workbook | Set-OfficeExcelCell -Sheet 'Review' -Address B2 -Value 'Ready'
            $workbook | Set-OfficeExcelCell -SheetIndex 1 -Address B3 -Value 'Watch'
            $workbook | Save-OfficeExcel
        } finally {
            $workbook | Close-OfficeExcel
        }

        $rows = @(Import-OfficeExcel -Path $path -WorksheetName 'Review')
        $rows.Count | Should -Be 2
        $rows[0].Service | Should -Be 'Identity'
        $rows[0].Status | Should -Be 'Ready'
        $rows[1].Service | Should -Be 'Messaging'
        $rows[1].Status | Should -Be 'Watch'
        $xml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/worksheets/sheet1.xml'
        @($xml.SelectNodes("//*[local-name()='c']")).Count | Should -Be 0
    }

    It 'rejects ambiguous or orphaned worksheet targets for <Command>' -ForEach @(
        @{ Command = 'Set-OfficeExcelCell'; Options = @{ Address = 'A1'; Value = 'Changed' } }
        @{ Command = 'Set-OfficeExcelRow'; Options = @{ Row = 1; Values = @('Changed') } }
        @{ Command = 'Set-OfficeExcelColumn'; Options = @{ Column = 1; Values = @('Changed') } }
    ) {
        $workbook = New-OfficeExcel -Path (Join-Path $TestDrive 'Rejected.xlsx') -NoSave
        try {
            $sheet = $workbook | Add-OfficeExcelSheet -Name 'Data' -PassThru
            $usedRange = $sheet.UsedRangeA1
            { & $Command -Worksheet $sheet -Document $workbook @Options } | Should -Throw '*either -Worksheet or -Document*'
            { & $Command -Document $workbook -Sheet Data -SheetIndex 0 @Options } | Should -Throw '*either -Sheet or -SheetIndex*'
            { & $Command -Sheet Data @Options } | Should -Throw '*together with -Document*'
            { & $Command -Document $workbook -SheetIndex 5 @Options } | Should -Throw '*SheetIndex is out of range*'
            $sheet.UsedRangeA1 | Should -Be $usedRange
            $text = $null
            $sheet.TryGetCellText(1, 1, [ref] $text) | Should -BeFalse
        } finally {
            $workbook | Close-OfficeExcel
        }
    }

    It 'rejects blank worksheet names without editing the first sheet for <Command>' -ForEach @(
        @{ Command = 'Set-OfficeExcelCell'; Options = @{ Address = 'A1'; Value = 'Changed' } }
        @{ Command = 'Set-OfficeExcelRow'; Options = @{ Row = 1; Values = @('Changed') } }
        @{ Command = 'Set-OfficeExcelColumn'; Options = @{ Column = 1; Values = @('Changed') } }
    ) {
        $workbook = New-OfficeExcel -Path (Join-Path $TestDrive 'BlankSheet.xlsx') -NoSave
        try {
            $sheet = $workbook | Add-OfficeExcelSheet -Name 'Data' -PassThru
            $workbook | Add-OfficeExcelSheet -Name 'Review'
            foreach ($name in '', '   ') {
                { & $Command -Document $workbook -Sheet $name @Options } | Should -Throw '*Sheet cannot be empty or whitespace*'
            }
            $text = $null
            $sheet.TryGetCellText(1, 1, [ref] $text) | Should -BeFalse

            # Omitting the selector still intentionally chooses the first worksheet.
            & $Command -Document $workbook @Options
            $sheet.TryGetCellText(1, 1, [ref] $text) | Should -BeTrue
            $text | Should -Be 'Changed'
        } finally {
            $workbook | Close-OfficeExcel
        }
    }

    It 'preserves scalar Excel types through the <Target> command' -ForEach @(
        @{ Target = 'Cell' }
        @{ Target = 'Row' }
        @{ Target = 'Column' }
    ) {
        $path = Join-Path $TestDrive "$Target-Types.xlsx"
        $values = @(123.45, $true, [datetime]'2026-10-16', '00123')
        $workbook = New-OfficeExcel -Path $path -NoSave
        try {
            $sheet = $workbook | Add-OfficeExcelSheet -Name 'Data' -PassThru
            if ($Target -eq 'Cell') {
                for ($index = 0; $index -lt $values.Count; $index++) {
                    $sheet | Set-OfficeExcelCell -Row 2 -Column ($index + 1) -Value $values[$index]
                }
            } elseif ($Target -eq 'Row') {
                $sheet | Set-OfficeExcelRow -Row 2 -Values $values
            } else {
                $sheet | Set-OfficeExcelColumn -Column 1 -StartRow 2 -Values $values
            }
            $workbook | Save-OfficeExcel
        } finally {
            $workbook | Close-OfficeExcel
        }

        $xml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $cells = @($xml.SelectNodes("//*[local-name()='c']"))
        $cells.Count | Should -Be 4
        $cells[0].GetAttribute('t') | Should -Not -Be 's'
        $cells[0].SelectSingleNode("*[local-name()='v']").InnerText | Should -Be '123.45'
        $cells[1].GetAttribute('t') | Should -Be 'b'
        $cells[1].SelectSingleNode("*[local-name()='v']").InnerText | Should -Be '1'
        $cells[2].GetAttribute('t') | Should -Not -Be 's'
        [double]::Parse($cells[2].SelectSingleNode("*[local-name()='v']").InnerText, [cultureinfo]::InvariantCulture) | Should -Be $values[2].ToOADate()
        $cells[3].GetAttribute('t') | Should -Be 's'
        $strings = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/sharedStrings.xml'
        $strings.SelectSingleNode("//*[local-name()='si']/*[local-name()='t']").InnerText | Should -Be '00123'
    }
}
