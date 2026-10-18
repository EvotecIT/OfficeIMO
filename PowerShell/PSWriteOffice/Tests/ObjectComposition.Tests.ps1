BeforeAll {
    $moduleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $moduleManifest -Global -ErrorAction Stop
}

Describe 'Object-style document composition' {
    It 'exports the concise PowerPoint section alias used by DSL recipes' {
        (Get-Alias PptSection).ResolvedCommandName | Should -Be 'Add-OfficePowerPointSection'
    }

    It 'composes and saves a Word document without a DSL scriptblock' {
        $path = Join-Path $TestDrive 'ObjectFlow.docx'
        $rows = @(
            [pscustomobject]@{ Service = 'Identity'; Status = 'Ready' }
            [pscustomobject]@{ Service = 'Certificates'; Status = 'Watch' }
        )

        $document = New-OfficeWord -Path $path -NoSave
        try {
            $heading = $document | Add-OfficeWordParagraph -Text 'Service status' -Style Heading1 -PassThru
            $heading | Add-OfficeWordText -Text ' — reviewed' -Italic

            $details = $document | Add-OfficeWordParagraph -PassThru
            $details | Add-OfficeWordText -Run @{
                Text  = 'Owner: ', 'Platform', '  Status: ', 'Ready'
                Bold  = $true, $false, $true, $true
                Color = $null, $null, $null, 'SeaGreen'
            }

            Add-OfficeWordTable -Document $document -InputObject $rows -Style GridTable4Accent1 -Layout AutoFitToWindow
            $document | Save-OfficeWord
        } finally {
            $document | Close-OfficeWord
        }

        $path | Should -Exist
        $readBack = Get-OfficeWord -Path $path -ReadOnly
        try {
            ($readBack.Paragraphs.Text -join '') | Should -Match 'Service status — reviewed'
            $readBack.Tables.Count | Should -Be 1
        } finally {
            $readBack | Close-OfficeWord
        }
    }

    It 'adds a Word section through the document pipeline' {
        $path = Join-Path $TestDrive 'ObjectSections.docx'
        $document = New-OfficeWord -Path $path -NoSave
        try {
            $section = $document | Add-OfficeWordSection -BreakType NextPage -PassThru
            $section | Add-OfficeWordParagraph -Text 'Second section' -PassThru | Should -Not -BeNullOrEmpty
            $document | Save-OfficeWord
        } finally {
            $document | Close-OfficeWord
        }

        $readBack = Get-OfficeWord -Path $path -ReadOnly
        try {
            $readBack.Sections.Count | Should -BeGreaterThan 1
            ($readBack.Paragraphs.Text -join "`n") | Should -Match 'Second section'
        } finally {
            $readBack | Close-OfficeWord
        }
    }

    It 'composes and saves an Excel workbook without a DSL scriptblock' {
        $path = Join-Path $TestDrive 'ObjectFlow.xlsx'
        $rows = @(
            [pscustomobject]@{ Project = 'Atlas'; Owner = 'Operations'; Progress = 0.80 }
            [pscustomobject]@{ Project = 'Beacon'; Owner = 'Security'; Progress = 0.55 }
        )

        $workbook = New-OfficeExcel -Path $path -NoSave
        try {
            $sheet = $workbook | Add-OfficeExcelSheet -Name 'Projects' -PassThru
            $sheet | Set-OfficeExcelCell -Address A1 -Value 'Portfolio' -BackgroundColor '#D9EAF7'
            Add-OfficeExcelTable -Worksheet $sheet -InputObject $rows -StartRow 3 -TableName 'Projects' -AutoFit
            Set-OfficeExcelCell -Document $workbook -Sheet 'Projects' -Address C4 -NumberFormat '0%'
            $workbook | Save-OfficeExcel
        } finally {
            $workbook | Close-OfficeExcel
        }

        $path | Should -Exist
        $data = @(Import-OfficeExcel -Path $path -WorksheetName 'Projects' -Range 'A3:C5')
        $data.Count | Should -Be 2
        $data[0].Project | Should -Be 'Atlas'
    }

    It 'reuses the active Excel scope for an explicit matching workbook target' {
        $path = Join-Path $TestDrive 'NestedObjectFlow.xlsx'
        $workbook = New-OfficeExcel -Path $path -NoSave
        try {
            $workbook | Add-OfficeExcelSheet -Name 'Outer' -Content {
                Add-OfficeExcelSheet -Document $workbook -Name 'Inner' -Content {
                    Set-OfficeExcelCell -Address A1 -Value 'Message'
                    Set-OfficeExcelCell -Address A2 -Value 'Nested object target'
                }
            }
            $workbook | Save-OfficeExcel
        } finally {
            $workbook | Close-OfficeExcel
        }

        $rows = @(Import-OfficeExcel -Path $path -WorksheetName 'Inner' -Range A1:A2)
        $rows.Count | Should -Be 1
        $rows[0].Message | Should -Be 'Nested object target'
    }

    It 'rejects mismatched nested Excel targets before creating a worksheet' {
        $target = New-OfficeExcel -Path (Join-Path $TestDrive 'Target.xlsx') -NoSave
        $active = $null
        try {
            $sheetCount = $target.Sheets.Count
            $active = New-OfficeExcel -Path (Join-Path $TestDrive 'Active.xlsx') -NoSave {
                {
                    Add-OfficeExcelSheet -Document $target -Name 'Rejected' -Content {
                        Set-OfficeExcelCell -Address A1 -Value 'Must not be written'
                    }
                } | Should -Throw '*does not match the active Excel composition scope*'
            }

            $target.Sheets.Count | Should -Be $sheetCount
        } finally {
            if ($active) { $active | Close-OfficeExcel }
            $target | Close-OfficeExcel
        }
    }

    It 'rejects mismatched nested Word targets before changing the document' {
        $target = New-OfficeWord -Path (Join-Path $TestDrive 'Target.docx') -NoSave
        $active = $null
        try {
            $sectionCount = $target.Sections.Count
            $paragraphCount = $target.Paragraphs.Count
            $targetSection = $target.Sections[0]

            $active = New-OfficeWord -Path (Join-Path $TestDrive 'Active.docx') -NoSave {
                {
                    Add-OfficeWordSection -Document $target -Content {
                        Add-OfficeWordParagraph -Text 'Must not be written'
                    }
                } | Should -Throw '*different document*'

                {
                    Add-OfficeWordParagraph -Target $target -Text 'Must not be written' -Content {
                        Add-OfficeWordText -Text 'Nested text'
                    }
                } | Should -Throw '*different document*'

                {
                    Add-OfficeWordParagraph -Target $targetSection -Content {
                        Add-OfficeWordText -Text 'Must not be written'
                    }
                } | Should -Throw '*different document*'
            }

            $target.Sections.Count | Should -Be $sectionCount
            $target.Paragraphs.Count | Should -Be $paragraphCount
        } finally {
            if ($active) { $active | Close-OfficeWord }
            $target | Close-OfficeWord
        }
    }

    It 'removes saved-path associations when a NoSave DSL block fails' {
        $assembly = (Get-Command New-OfficeExcel).ImplementingType.Assembly
        $flags = [System.Reflection.BindingFlags]::NonPublic -bor [System.Reflection.BindingFlags]::Static
        $excelField = $assembly.GetType('PSWriteOffice.Services.Excel.ExcelDocumentService').GetField('AssociatedPaths', $flags)
        $wordField = $assembly.GetType('PSWriteOffice.Services.Word.WordDocumentService').GetField('AssociatedPaths', $flags)
        $excelCount = $excelField.GetValue($null).Count
        $wordCount = $wordField.GetValue($null).Count

        {
            New-OfficeExcel -Path (Join-Path $TestDrive 'Failed.xlsx') -NoSave {
                throw 'Expected Excel DSL failure'
            }
        } | Should -Throw '*Expected Excel DSL failure*'

        {
            New-OfficeWord -Path (Join-Path $TestDrive 'Failed.docx') -NoSave {
                throw 'Expected Word DSL failure'
            }
        } | Should -Throw '*Expected Word DSL failure*'

        $excelField.GetValue($null).Count | Should -Be $excelCount
        $wordField.GetValue($null).Count | Should -Be $wordCount
    }
}
