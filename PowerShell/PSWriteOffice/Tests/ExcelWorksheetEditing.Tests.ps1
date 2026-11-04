BeforeAll {
    $moduleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) { $env:PSWRITEOFFICE_MODULE_MANIFEST } else { Join-Path $PSScriptRoot '../PSWriteOffice.psd1' }
    Import-Module $moduleManifest -Global -ErrorAction Stop
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
}

Describe 'Excel worksheet editing' {
    It 'edits typed worksheet targets quietly and persists row and column layout' {
        $path = Join-Path $TestDrive 'WorksheetTargets.xlsx'
        $workbook = New-OfficeExcel -Path $path
        try {
            $first = $workbook | Add-OfficeExcelSheet -Name 'First' -PassThru
            $second = $workbook | Add-OfficeExcelSheet -Name 'Second' -PassThru
            $output = @(foreach ($target in $first, $second) { Set-OfficeExcelRow -Sheet $target -Row 1 -Values 'Name', 'Notes' -Height 30 })
            $output.Count | Should -Be 0
            $output = @(foreach ($target in $first, $second) { Set-OfficeExcelColumn -Sheet $target -Column B -Width 42 -Hidden:$false -PassThru })
            $output.Count | Should -Be 2
            $output[0].Name | Should -Be First
            $output[1].Name | Should -Be Second
            Set-OfficeExcelRow -Sheet $first -Row 2 -Values 'Identity', 'Ready'
            $workbook | Save-OfficeExcel
        } finally {
            $workbook | Close-OfficeExcel
        }

        $rows = @(Import-OfficeExcel -Path $path -Sheet 'First')
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
        $workbook = New-OfficeExcel -Path $path
        try {
            $workbook | Add-OfficeExcelSheet -Name 'Unchanged'
            $workbook | Add-OfficeExcelSheet -Name 'Review'
            $workbook | Set-OfficeExcelRow -Sheet 'Review' -Row 1 -Values 'Service', 'Status'
            $workbook | Set-OfficeExcelColumn -Sheet 2 -Column A -StartRow 2 -Values 'Identity', 'Messaging'
            $workbook | Set-OfficeExcelCell -Sheet 'Review' -Address B2 -Value 'Ready'
            $workbook | Set-OfficeExcelCell -Sheet 2 -Address B3 -Value 'Watch'
            $workbook | Save-OfficeExcel
        } finally {
            $workbook | Close-OfficeExcel
        }

        $rows = @(Import-OfficeExcel -Path $path -Sheet 'Review')
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
        $workbook = New-OfficeExcel -Path (Join-Path $TestDrive 'Rejected.xlsx')
        try {
            $sheet = $workbook | Add-OfficeExcelSheet -Name 'Data' -PassThru
            $usedRange = $sheet.UsedRange.Address
            $foreignBook = New-OfficeExcel
            try {
                $foreign = Add-OfficeExcelSheet -Document $foreignBook -Name Data -PassThru
                { & $Command -Document $workbook -Sheet $foreign @Options } | Should -Throw
                { & $Command -Document $workbook -Sheet 0 @Options } | Should -Throw
                { & $Command -Sheet Data @Options } | Should -Throw
                { & $Command -Document $workbook -Sheet 6 @Options } | Should -Throw
            } finally {
                Close-OfficeExcel -Document $foreignBook
            }
            $sheet.UsedRange.Address | Should -Be $usedRange
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
        $workbook = New-OfficeExcel -Path (Join-Path $TestDrive 'BlankSheet.xlsx')
        try {
            $sheet = $workbook | Add-OfficeExcelSheet -Name 'Data' -PassThru
            $workbook | Add-OfficeExcelSheet -Name 'Review'
            foreach ($name in '', '   ') {
                { & $Command -Document $workbook -Sheet $name @Options } | Should -Throw
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
        $workbook = New-OfficeExcel -Path $path
        try {
            $sheet = $workbook | Add-OfficeExcelSheet -Name 'Data' -PassThru
            if ($Target -eq 'Cell') {
                for ($index = 0; $index -lt $values.Count; $index++) {
                    Set-OfficeExcelCell -Sheet $sheet -Row 2 -Column ($index + 1) -Value $values[$index]
                }
            } elseif ($Target -eq 'Row') {
                Set-OfficeExcelRow -Sheet $sheet -Row 2 -Values $values
            } else {
                Set-OfficeExcelColumn -Sheet $sheet -Column 1 -StartRow 2 -Values $values
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
