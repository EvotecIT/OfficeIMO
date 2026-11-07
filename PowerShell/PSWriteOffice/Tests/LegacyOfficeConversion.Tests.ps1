BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    $documentAssets = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../../OfficeIMO.TestAssets/Documents')).ProviderPath
    $dbxAssets = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../../OfficeIMO.Email.Tests/Corpora/dbx-metadata-unknown-producer')).ProviderPath
}

Describe 'Legacy Office conversion cmdlets' {
    It 'converts Word documents between docx and doc' {
        $docxPath = Join-Path $TestDrive 'source.docx'
        $docPath = Join-Path $TestDrive 'converted.doc'
        $roundTripPath = Join-Path $TestDrive 'roundtrip.docx'

        New-OfficeWord -Path $docxPath {
            Add-OfficeWordParagraph -Text 'Legacy Word conversion'
        } | Out-Null

        $docFile = Convert-OfficeDocument -Path $docxPath -OutputPath $docPath
        $docFile | Should -BeOfType OfficeIMO.OfficeConversionReport
        Test-Path -LiteralPath $docPath | Should -BeTrue
        [BitConverter]::ToString([IO.File]::ReadAllBytes($docPath)[0..3]) | Should -Be 'D0-CF-11-E0'

        Convert-OfficeDocument -Path $docPath -OutputPath $roundTripPath
        $paragraphs = @(Get-OfficeWordParagraph -Path $roundTripPath)
        $paragraphs.Text | Should -Contain 'Legacy Word conversion'
    }

    It 'converts Excel workbooks between xlsx and xls' {
        $xlsxPath = Join-Path $TestDrive 'source.xlsx'
        $xlsPath = Join-Path $TestDrive 'converted.xls'
        $roundTripPath = Join-Path $TestDrive 'roundtrip.xlsx'

        New-OfficeExcel -Path $xlsxPath {
            ExcelSheet 'Data' {
                Set-OfficeExcelRow -Row 1 -Values 'Name', 'Value'
                Set-OfficeExcelRow -Row 2 -Values 'Ada', 42
            }
        } | Out-Null

        $xlsFile = Convert-OfficeDocument -Path $xlsxPath -OutputPath $xlsPath
        $xlsFile | Should -BeOfType OfficeIMO.OfficeConversionReport
        Test-Path -LiteralPath $xlsPath | Should -BeTrue
        [BitConverter]::ToString([IO.File]::ReadAllBytes($xlsPath)[0..3]) | Should -Be 'D0-CF-11-E0'

        Convert-OfficeDocument -Path $xlsPath -OutputPath $roundTripPath
        $rows = @(Import-OfficeExcel -Path $roundTripPath -Sheet 'Data')
        $rows | Should -HaveCount 1
        $rows[0].Name | Should -Be 'Ada'
        [int] $rows[0].Value | Should -Be 42
    }

    It 'honors NoClobber before explicitly replacing an existing conversion output' {
        $sourcePath = Join-Path $TestDrive 'overwrite-source.docx'
        $existingPath = Join-Path $TestDrive 'existing-target.docx'
        $targetPath = Join-Path $TestDrive 'overwrite-target.doc'

        New-OfficeWord -Path $sourcePath {
            Add-OfficeWordParagraph -Text 'Overwrite source'
        } | Out-Null
        New-OfficeWord -Path $existingPath {
            Add-OfficeWordParagraph -Text 'Keep me'
        } | Out-Null

        Convert-OfficeDocument -Path $existingPath -OutputPath $targetPath | Out-Null
        $originalBytes = [IO.File]::ReadAllBytes($targetPath)

        $skipped = Convert-OfficeDocument -Path $sourcePath -OutputPath $targetPath -NoClobber -ErrorAction Stop
        $skipped.Skipped | Should -BeTrue
        $skipped.SkippedReason | Should -Be 'OutputExists'
        [IO.File]::ReadAllBytes($targetPath) | Should -Be $originalBytes

        Convert-OfficeDocument -Path $sourcePath -OutputPath $targetPath
        $paragraphs = @(Get-OfficeWordParagraph -Path $targetPath)
        $paragraphs.Text | Should -Contain 'Overwrite source'
    }

    It 'imports independently produced Word 6/95 text through the existing Word conversion route' {
        $source = Join-Path $documentAssets 'LegacyDocCorpus/apache-poi-word69/Bug60936.doc'
        $sourceHash = (Get-FileHash -LiteralPath $source).Hash
        $destination = Join-Path $TestDrive 'word69.docx'
        $expected = '4 sk{0}re a p{1}ed 7 lety' -f [char] 0x00F3, [char] 0x0159

        $report = Get-Item -LiteralPath $source |
            Convert-OfficeDocument -OutputPath $destination -AllowLegacyImportLoss -ErrorAction Stop
        $report.OutputPath | Should -Be $destination
        $report.Diagnostics.Code | Should -Not -Contain 'doc-word69-text-codepage-unsupported'
        $report.Diagnostics.Code | Should -Contain 'doc-word69-formatting-unsupported'
        $report.Diagnostics.Code | Should -Contain 'doc-word69-bookmarks'
        (Get-OfficeWordParagraph -Path $destination).Text -join "`n" | Should -Match ([regex]::Escape($expected))
        Open-OfficeWord -Path $destination {
            Add-OfficeWordParagraph -Text 'Reviewed Word 6/95 migration'
        } | Out-Null
        $paragraphs = @(Get-OfficeWordParagraph -Path $destination)
        $paragraphs.Text | Should -Contain 'Reviewed Word 6/95 migration'
        $paragraphs.Text -join "`n" | Should -Match ([regex]::Escape($expected))
        (Get-FileHash -LiteralPath $source).Hash | Should -Be $sourceHash
    }

    It 'converts a native BIFF4 workspace with editable array formulas and retained caches' {
        $source = Join-Path $documentAssets 'LegacyXlsCorpus/apache-openoffice-testdocs/formula_import_biff4.xlw'
        $sourceHash = (Get-FileHash -LiteralPath $source).Hash
        $destination = Join-Path $TestDrive 'biff4.xlsx'
        $saveOptions = New-OfficeExcelSaveOptions
        $saveOptions.AllowLoss = $true

        $report = Convert-OfficeDocument -LiteralPath $source -OutputPath $destination -Options $saveOptions -ErrorAction Stop
        $report.OutputPath | Should -Be $destination
        @($report.Diagnostics | Where-Object {
            $_.Code -eq 'early-biff-formula-unsupported' -and $_.Location -eq "'Array'!B13"
        }).Count | Should -Be 1
        $sheet = Get-ZipXmlDocumentLocal -Path $destination -Entry 'xl/worksheets/sheet4.xml'
        $ns = [System.Xml.XmlNamespaceManager]::new($sheet.NameTable)
        $ns.AddNamespace('s', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
        $sheet.SelectNodes('//s:f[@t="array"]', $ns).Count | Should -Be 5
        $formula = $sheet.SelectSingleNode('//s:c[@r="B9"]/s:f', $ns)
        $formula.InnerText | Should -Be 'H9:I10'
        $formula.GetAttribute('ref') | Should -Be 'B9:C10'
        @('B9', 'C9', 'B10', 'C10') | ForEach-Object {
            $sheet.SelectSingleNode(('//s:c[@r="{0}"]/s:v' -f $_), $ns).InnerText
        } | Should -Be @('1', '2', '3', '4')

        Open-OfficeExcel -Path $destination {
            Set-OfficeExcelCell -Sheet Array -Address A40 -Value 'Reviewed BIFF4 migration'
        } | Out-Null
        Read-OfficeDocument -Path $destination -Sheet Array -Range A40:A40 -NoHeaderRow -As Text |
            Should -Match 'Reviewed BIFF4 migration'
        (Get-FileHash -LiteralPath $source).Hash | Should -Be $sourceHash
    }

    It 'converts Pages 09 running stories through typed workflow settings' {
        $source = Join-Path $documentAssets 'IWorkCorpus/legacy-pages-running-stories/header+footer.pages'
        $sourceHash = (Get-FileHash -LiteralPath $source).Hash
        $destination = Join-Path $TestDrive 'pages09.docx'
        $rejected = Join-Path $TestDrive 'pages09-rejected.docx'
        $conversionType = Get-TestPSWriteOfficeType -AssemblyName OfficeIMO -TypeName OfficeIMO.OfficeConversionOptions -CommandName Convert-OfficeDocument
        $settingsType = Get-TestPSWriteOfficeType -AssemblyName OfficeIMO.Workflows.IWork -TypeName OfficeIMO.Workflows.IWork.IWorkWorkflowSettings -CommandName Convert-OfficeDocument
        $options = [Activator]::CreateInstance($conversionType)
        $settings = [Activator]::CreateInstance($settingsType)
        $settings.ConversionOptions.Mode = 'EditableOnly'
        $options.RegisteredSettings = $settings

        { Convert-OfficeDocument -LiteralPath $source -OutputPath $rejected -Options $options -ErrorAction Stop } |
            Should -Throw
        Test-Path -LiteralPath $rejected | Should -BeFalse
        $settings.ConversionOptions.AllowPartialEditableReconstruction = $true
        $projection = Get-Item -LiteralPath $source | Open-OfficeIWork -ErrorAction Stop
        try {
            $projection.ReadPages().Sections[0].DefaultPageHeaderContents[0].PlainText | Should -Be 'Header.'
            $report = Convert-OfficeDocument -LiteralPath $source -OutputPath $destination -Options $options -ErrorAction Stop
            $report.OutputPath | Should -Be $destination
        } finally {
            $projection.Dispose()
        }
        $word = Open-OfficeWord -Path $destination
        try {
            $word.PageSetup.Header.Paragraphs.Text | Should -Contain 'Header.'
            $word.PageSetup.Footer.Paragraphs.Text | Should -Contain 'Footer.'
            $word.PageSetup.Header.AddParagraph('Reviewed Pages 09 header') | Out-Null
            Save-OfficeWord -Document $word -Path $destination -ErrorAction Stop
        } finally {
            Close-OfficeWord -Document $word
        }
        $reopened = Open-OfficeWord -Path $destination -ReadOnly
        try {
            $reopened.PageSetup.Header.Paragraphs.Text | Should -Contain 'Reviewed Pages 09 header'
            $reopened.PageSetup.Footer.Paragraphs.Text | Should -Contain 'Footer.'
        } finally {
            Close-OfficeWord -Document $reopened
        }
        (Get-FileHash -LiteralPath $source).Hash | Should -Be $sourceHash
    }

    It 'converts native Word XML lists through the existing XML route' {
        $source = Join-Path $documentAssets 'WordXmlMigration/word-lists-native-2003.xml'
        $sourceHash = (Get-FileHash -LiteralPath $source).Hash
        $destination = Join-Path $TestDrive 'wordxml.docx'

        Convert-OfficeDocument -LiteralPath $source -OutputPath $destination -Route word-xml-docx -AllowLegacyImportLoss -ErrorAction Stop | Out-Null
        Open-OfficeWord -Path $destination {
            Add-OfficeWordParagraph -Text 'Reviewed Word XML migration'
        } | Out-Null
        $word = Open-OfficeWord -Path $destination -ReadOnly
        try {
            $paragraphs = @(Get-OfficeWordParagraph -Document $word)
            $listParagraphs = @($paragraphs | Where-Object IsListItem)
            $listParagraphs | Should -HaveCount 11
            $listParagraphs.Text | Should -Contain 'Decimal continued'
            $listParagraphs.Text | Should -Contain 'Decimal starts seven'
            @($listParagraphs | Where-Object ListItemLevel -EQ 1).Text |
                Should -Be @('Outline child one', 'Outline child two')
            $paragraphs.Text | Should -Contain 'Reviewed Word XML migration'
        } finally {
            Close-OfficeWord -Document $word
        }
        (Get-FileHash -LiteralPath $source).Hash | Should -Be $sourceHash
    }

    It 'converts native SpreadsheetML through the existing XML route' {
        $source = Join-Path $documentAssets 'SpreadsheetXmlCorpus/excel-preservation.xml'
        $sourceHash = (Get-FileHash -LiteralPath $source).Hash
        $destination = Join-Path $TestDrive 'spreadsheetxml.xlsx'

        $report = Convert-OfficeDocument -LiteralPath $source -OutputPath $destination -Route spreadsheet-xml-xlsx -ErrorAction Stop
        $report.OutputPath | Should -Be $destination
        $sheet = Get-ZipXmlDocumentLocal -Path $destination -Entry 'xl/worksheets/sheet1.xml'
        $ns = [System.Xml.XmlNamespaceManager]::new($sheet.NameTable)
        $ns.AddNamespace('s', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
        $sheet.SelectSingleNode('//s:mergeCell', $ns).GetAttribute('ref') | Should -Be 'A3:B4'
        Open-OfficeExcel -Path $destination {
            Set-OfficeExcelCell -Sheet Data -Address C6 -Value 'Reviewed SpreadsheetML migration'
        } | Out-Null
        Read-OfficeDocument -Path $destination -Sheet Data -Range C6:C6 -NoHeaderRow -As Text |
            Should -Match 'Reviewed SpreadsheetML migration'
        (Get-FileHash -LiteralPath $source).Hash | Should -Be $sourceHash
    }

    It 'reads and searches a native DBX file through the existing Reader surface' {
        $source = Join-Path $dbxAssets 'Inbox.dbx'
        $sourceHash = (Get-FileHash -LiteralPath $source).Hash

        $document = Get-Item -LiteralPath $source | Read-OfficeDocument -MaximumStoreItems 10 -ErrorAction Stop
        $document.Kind.ToString() | Should -Be Email
        ($document.Metadata | Where-Object Name -EQ StoreFormat).Value | Should -Be Dbx
        ($document.Metadata | Where-Object Name -EQ ItemCount).Value | Should -Be '1'
        ($document.Metadata | Where-Object Name -EQ Subject).Value | Should -Be 'Welcome to Outlook Express 6'
        $metadata = @($document.Metadata | Where-Object { $_.Name -in 'NormalizedSubject', 'Importance', 'DeclaredSize' })
        $metadata.Name | Should -Be @('NormalizedSubject', 'Importance', 'DeclaredSize')
        $metadata.Value | Should -Be @('Welcome to Outlook Express 6', 'Normal', '10139')
        $metadata.ValueType | Should -Be @('string', 'string', 'number')
        $metadata.SourceObjectId | Select-Object -Unique | Should -Be 'message[0]'
        $json = [pscustomobject]@{ Metadata = $document.Metadata } | ConvertTo-Json -Depth 5 | ConvertFrom-Json
        $jsonMetadata = $json.Metadata
        ($jsonMetadata | Where-Object Name -EQ NormalizedSubject).Value | Should -Be 'Welcome to Outlook Express 6'
        ($jsonMetadata | Where-Object Name -EQ Importance).Value | Should -Be 'Normal'
        ($jsonMetadata | Where-Object Name -EQ DeclaredSize).ValueType | Should -Be 'number'
        ($jsonMetadata | Where-Object Name -EQ DeclaredSize).Value | Should -Be '10139'
        $document.Chunks[0].Location.SourceBlockKind | Should -Be 'email-message'
        $bodyText = ($document.Chunks.Text -join ' ') -replace '\s+', ' '
        $firstBodyPhrase = $bodyText.IndexOf('The solution for all your messaging needs', [StringComparison]::Ordinal)
        $lastBodyPhrase = $bodyText.IndexOf('Thank you for choosing Internet Explorer and Outlook Express 6.', [StringComparison]::Ordinal)
        $firstBodyPhrase | Should -BeGreaterThan -1
        $lastBodyPhrase | Should -BeGreaterThan $firstBodyPhrase
        $matches = @($document | Search-OfficeDocument -Pattern 'Welcome to Outlook Express 6' -SimpleMatch -ErrorAction Stop)
        $matches.Count | Should -BeGreaterThan 0
        $matches.Path | Select-Object -Unique | Should -Be $source
        (Get-FileHash -LiteralPath $source).Hash | Should -Be $sourceHash
    }

    It 'applies the source byte limit before replacing a legacy conversion destination' {
        $source = Join-Path $documentAssets 'LegacyDocCorpus/apache-poi-word69/Bug60936.doc'
        $sourceHash = (Get-FileHash -LiteralPath $source).Hash
        $destination = Join-Path $TestDrive 'bounded-word69.docx'
        [IO.File]::WriteAllText($destination, 'Keep the existing destination')
        $destinationHash = (Get-FileHash -LiteralPath $destination).Hash

        { Convert-OfficeDocument -LiteralPath $source -OutputPath $destination -AllowLegacyImportLoss -MaximumInputBytes ((Get-Item -LiteralPath $source).Length - 1) -ErrorAction Stop } |
            Should -Throw
        (Get-FileHash -LiteralPath $destination).Hash | Should -Be $destinationHash
        (Get-FileHash -LiteralPath $source).Hash | Should -Be $sourceHash
    }
}
