BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        $sourceRoot = Join-Path (Join-Path (Join-Path $PSScriptRoot '..') 'Sources') 'PSWriteOffice'

        if (-not $env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES) {
            $env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES = 'true'
        }

        if (-not $env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION) {
            $releasePath = Join-Path (Join-Path $sourceRoot 'bin') 'Release'
            $env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION = if (Test-Path $releasePath) { 'Release' } else { 'Debug' }
        }

        Join-Path (Join-Path $PSScriptRoot '..') 'PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
}

Describe 'CSV and Excel mutation contracts' {
    It 'does not create a CSV file when Export-OfficeCsv is invoked with WhatIf' {
        $path = Join-Path $TestDrive 'whatif.csv'

        1..3 |
            ForEach-Object { [pscustomobject]@{ Name = "Row$_"; Value = $_ } } |
            Export-OfficeCsv -Path $path -WhatIf

        Test-Path -LiteralPath $path | Should -BeFalse
    }

    It 'does not copy an Excel template when New-OfficeExcel NoSave is invoked with WhatIf' {
        $template = Join-Path $TestDrive 'template.xlsx'
        $target = Join-Path $TestDrive 'target.xlsx'
        New-OfficeExcel -Path $template {
            ExcelSheet -Name Data {
                ExcelCell -Address A1 -Value 'Template'
            }
        }

        New-OfficeExcel -TemplatePath $template -Path $target -NoSave -WhatIf | Out-Null

        Test-Path -LiteralPath $target | Should -BeFalse
    }

    It 'does not create an Excel package when New-OfficeExcel NoSave is invoked with WhatIf' {
        $target = Join-Path $TestDrive 'nosave-whatif.xlsx'

        New-OfficeExcel -Path $target -NoSave -WhatIf {
            ExcelSheet -Name Data
        } | Out-Null

        Test-Path -LiteralPath $target | Should -BeFalse
    }

    It 'returns a NoSave Excel workbook that can be saved later' {
        $target = Join-Path $TestDrive 'nosave-later-save.xlsx'

        $document = New-OfficeExcel -Path $target -NoSave {
            ExcelSheet -Name Data {
                ExcelCell -Address A1 -Value 'Ready'
            }
        }
        try {
            $document.Save($target)
        } finally {
            $document.Dispose()
        }

        Test-Path -LiteralPath $target | Should -BeTrue
    }

    It 'returns a live NoSave Excel workbook for emitted method calls' {
        $target = Join-Path $TestDrive 'nosave-live-workbook.xlsx'

        $workbook = New-OfficeExcel -Path $target -NoSave
        try {
            $sheet = $workbook.AddWorksheet('Data')
            $sheet.Cell(1, 1, 'Ready')
            Save-OfficeExcel -Document $workbook -Path $target
        } finally {
            $workbook.Dispose()
        }

        Test-Path -LiteralPath $target | Should -BeTrue
    }

    It 'does not copy a Word template when New-OfficeWord is invoked with WhatIf' {
        $template = Join-Path $TestDrive 'template.docx'
        $target = Join-Path $TestDrive 'target.docx'
        New-OfficeWord -Path $template {
            WordParagraph -Text 'Template'
        }

        New-OfficeWord -TemplatePath $template -Path $target -WhatIf | Out-Null

        Test-Path -LiteralPath $target | Should -BeFalse
    }

    It 'does not create a Word package when New-OfficeWord NoSave is invoked with WhatIf' {
        $target = Join-Path $TestDrive 'nosave-whatif.docx'

        New-OfficeWord -Path $target -NoSave -WhatIf {
            WordParagraph -Text 'Draft'
        } | Out-Null

        Test-Path -LiteralPath $target | Should -BeFalse
    }

    It 'imports a CSV source into a new Excel workbook' {
        $csv = Join-Path $TestDrive 'source.csv'
        $xlsx = Join-Path $TestDrive 'created.xlsx'
        1..3 |
            ForEach-Object { [pscustomobject]@{ Name = "Row$_"; Value = $_ } } |
            Export-Csv -Path $csv -NoTypeInformation

        $result = Import-OfficeExcelDelimitedText -Path $xlsx -SourcePath $csv -PassThru

        Test-Path -LiteralPath $xlsx | Should -BeTrue
        $result.RowCount | Should -Be 3
        $rows = Import-OfficeExcel -Path $xlsx -WorksheetName Import
        $rows.Count | Should -Be 3
        $rows[0].Name | Should -Be 'Row1'
    }

    It 'reports the detected delimiter for a delimited-text import' {
        $csv = Join-Path $TestDrive 'source-semicolon.csv'
        $xlsx = Join-Path $TestDrive 'created-semicolon.xlsx'
        Set-Content -LiteralPath $csv -Value "Name;Value`r`nAlpha;1" -Encoding UTF8

        $result = Import-OfficeExcelDelimitedText -Path $xlsx -SourcePath $csv -PassThru

        $result.Delimiter | Should -Be ';'
        $result.RowCount | Should -Be 1
    }

    It 'reports the actual sanitized and unique table name for downstream commands' {
        $csv = Join-Path $TestDrive 'source-table-name.csv'
        $xlsx = Join-Path $TestDrive 'created-table-name.xlsx'
        Set-Content -LiteralPath $csv -Value "Name,Value`r`nAlpha,1" -Encoding UTF8
        [pscustomobject]@{ Name = 'Existing'; Value = 0 } |
            Export-OfficeExcel -Path $xlsx -WorksheetName Seed -TableName Sales_Data

        $result = Import-OfficeExcelDelimitedText -Path $xlsx -SourcePath $csv -SheetName 'Sales Data' -PassThru

        $result.TableName | Should -Be 'Sales_Data2'
        $table = Get-OfficeExcelTable -Path $xlsx | Where-Object Name -eq $result.TableName
        $table.Name | Should -Be $result.TableName
        [pscustomobject]@{ Name = 'Beta'; Value = 2 } |
            Add-OfficeExcelTableRow -Path $xlsx -Sheet 'Sales Data' -TableName $result.TableName
        @(Import-OfficeExcel -Path $xlsx -WorksheetName 'Sales Data').Count | Should -Be 2
    }

    It 'preserves the first data row when importing headerless CSV into Excel' {
        $csv = Join-Path $TestDrive 'source-no-header.csv'
        $xlsx = Join-Path $TestDrive 'created-no-header.xlsx'
        Set-Content -LiteralPath $csv -Value "Alpha,1`r`nBeta,2" -Encoding UTF8

        $result = Import-OfficeExcelDelimitedText -Path $xlsx -SourcePath $csv -NoHeader -PassThru

        $result.RowCount | Should -Be 2
        $rows = @(Import-OfficeExcel -Path $xlsx -WorksheetName Import -NoHeader)
        $rows.Count | Should -Be 2
        $rows[0].Column1 | Should -Be 'Alpha'
        $rows[0].Column2 | Should -Be 1
    }

    It 'infers DataTable column types when importing Excel data' {
        $xlsx = Join-Path $TestDrive 'datatable-schema.xlsx'
        New-OfficeExcel -Path $xlsx {
            ExcelSheet -Name Data {
                ExcelCell -Address A1 -Value 'Name'
                ExcelCell -Address B1 -Value 'Quantity'
                ExcelCell -Address C1 -Value 'When'
                ExcelCell -Address A2 -Value 'Alpha'
                ExcelCell -Address B2 -Value 42
                ExcelCell -Address C2 -Value ([datetime]'2026-08-05')
            }
        }

        $table = Import-OfficeExcel -Path $xlsx -WorksheetName Data -AsDataTable

        $table.Columns['Name'].DataType | Should -Be ([string])
        $table.Columns['Quantity'].DataType | Should -Not -Be ([object])
        $table.Columns['When'].DataType | Should -Be ([datetime])
    }

    It 'projects scalar Excel table rows into the existing Value column' {
        $xlsx = Join-Path $TestDrive 'scalar-table-row.xlsx'
        $rows = @(
            [pscustomobject]@{ Name = 'Alpha'; Value = 1 }
            'Beta'
        )

        New-OfficeExcel -Path $xlsx {
            ExcelSheet -Name Data {
                Add-OfficeExcelTable -InputObject $rows -TableName Rows
            }
        }

        $data = Import-OfficeExcel -Path $xlsx -WorksheetName Data
        $data.Count | Should -Be 2
        ($null -eq $data[1].Name) | Should -BeTrue
        $data[1].Value | Should -Be 'Beta'
    }

    It 'imports a CSV source into Excel after skipping preamble rows' {
        $csv = Join-Path $TestDrive 'source-preamble.csv'
        $xlsx = Join-Path $TestDrive 'created-preamble.xlsx'
        Set-Content -LiteralPath $csv -Value "generated by vendor`r`nName,Value`r`nRow1,1`r`nRow2,2" -Encoding UTF8

        $result = Import-OfficeExcelDelimitedText -Path $xlsx -SourcePath $csv -SkipRows 1 -PassThru

        $result.RowCount | Should -Be 2
        $rows = @(Import-OfficeExcel -Path $xlsx -WorksheetName Import)
        $rows.Count | Should -Be 2
        $rows[0].Name | Should -Be 'Row1'
    }

    It 'exports the union of PowerShell object columns and blanks missing values' {
        $xlsx = Join-Path $TestDrive 'projected.xlsx'

        @(
            [pscustomobject]@{ Name = 'Row1'; Value = 1 }
            [pscustomobject]@{ Name = 'Row2'; Extra = 'Ignored' }
        ) | Export-OfficeExcel -Path $xlsx -WorksheetName Data -TableName Data

        $rows = @(Import-OfficeExcel -Path $xlsx -WorksheetName Data)
        $rows.Count | Should -Be 2
        $rows[0].Name | Should -Be 'Row1'
        $rows[0].Value | Should -Be 1
        $rows[1].Name | Should -Be 'Row2'
        ($null -eq $rows[1].Value) | Should -BeTrue
        $rows[1].Extra | Should -Be 'Ignored'
    }

    It 'reads Excel data rows as real custom objects' {
        $xlsx = Join-Path $TestDrive 'excel-data-rows.xlsx'

        [pscustomobject]@{ Name = 'Row1'; Value = 1 } |
            Export-OfficeExcel -Path $xlsx -WorksheetName Data -TableName Data

        $rows = @(Get-OfficeExcelData -Path $xlsx -Sheet Data)

        $rows.Count | Should -Be 1
        $rows[0].GetType().FullName | Should -Be 'System.Management.Automation.PSCustomObject'
        $rows[0].Name | Should -Be 'Row1'
        $rows[0].Value | Should -Be 1
    }

    It 'does not append Excel table rows when Add-OfficeExcelTableRow is invoked with WhatIf' {
        $xlsx = Join-Path $TestDrive 'table.xlsx'
        [pscustomobject]@{ Name = 'Row1'; Value = 1 } |
            Export-OfficeExcel -Path $xlsx -WorksheetName Data -TableName Data

        [pscustomobject]@{ Name = 'Row2'; Value = 2 } |
            Add-OfficeExcelTableRow -Path $xlsx -Sheet Data -TableName Data -WhatIf

        $rows = @(Import-OfficeExcel -Path $xlsx -WorksheetName Data)
        $rows.Count | Should -Be 1
        $rows[0].Name | Should -Be 'Row1'
    }

    It 'does not add threaded comments to live Excel documents when WhatIf is used' {
        $xlsx = Join-Path $TestDrive 'threaded-comment-whatif.xlsx'
        New-OfficeExcel -Path $xlsx {
            Add-OfficeExcelSheet -Name Data
        }

        $document = Get-OfficeExcel -Path $xlsx
        try {
            $document | Add-OfficeExcelThreadedComment -Sheet Data -Address A1 -Text 'Nope' -Author Reviewer -WhatIf
            $audit = $document | Get-OfficeExcelCommentAudit -IncludeComments
            $audit.ThreadedCommentCount | Should -Be 0
        } finally {
            $document | Close-OfficeExcel
        }
    }

    It 'does not join workbooks into live Excel documents when WhatIf is used' {
        $target = Join-Path $TestDrive 'join-target-whatif.xlsx'
        $source = Join-Path $TestDrive 'join-source-whatif.xlsx'
        New-OfficeExcel -Path $target {
            Add-OfficeExcelSheet -Name Target
        }
        New-OfficeExcel -Path $source {
            Add-OfficeExcelSheet -Name Source
        }

        $document = Get-OfficeExcel -Path $target
        try {
            $document | Join-OfficeExcelWorkbook -SourcePath $source -WhatIf | Should -BeNullOrEmpty
            $summary = $document | Get-OfficeExcelSummary -IncludeSheets
            $summary.SheetCount | Should -Be 1
            $summary.Sheets.Name | Should -Contain 'Target'
            $summary.Sheets.Name | Should -Not -Contain 'Source'
        } finally {
            $document | Close-OfficeExcel
        }
    }

    It 'does not repair live Excel documents when WhatIf is used' {
        $xlsx = Join-Path $TestDrive 'repair-whatif.xlsx'
        New-OfficeExcel -Path $xlsx {
            Add-OfficeExcelSheet -Name Data
        }

        $document = Get-OfficeExcel -Path $xlsx
        try {
            $document | Repair-OfficeExcelWorkbook -PassThru -WhatIf | Should -BeNullOrEmpty
        } finally {
            $document | Close-OfficeExcel
        }
    }
}
