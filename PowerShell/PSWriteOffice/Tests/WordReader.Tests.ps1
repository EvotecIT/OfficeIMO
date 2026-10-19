BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop

    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
}

Describe 'Word reader helpers' {
    It 'finds text and exposes bookmarks/fields' {
        $path = Join-Path $TestDrive 'WordReader.docx'

        New-OfficeWord -Path $path {
            Add-OfficeWordParagraph -Text 'Hello world'
        }

        $doc = Get-OfficeWord -Path $path
        try {
            $null = $doc.AddBookmark('Bookmark1')
            $paragraph = $doc.AddParagraph('Page')
            $pageFieldType = Get-TestPSWriteOfficeEnumValue -AssemblyName 'OfficeIMO.Word' -TypeName 'OfficeIMO.Word.WordFieldType' -Name 'Page'
            $null = $paragraph.AddField($pageFieldType)
        } finally {
            Close-OfficeWord -Document $doc -Save
        }

        $matches = Find-OfficeWord -Path $path -Text 'Hello'
        $matches.Count | Should -BeGreaterThan 0

        $bookmarks = Get-OfficeWordBookmark -Path $path
        $bookmarks.Name | Should -Contain 'Bookmark1'

        $fields = Get-OfficeWordField -Path $path -FieldType Page
        $fields.Count | Should -BeGreaterThan 0
    }

    It 'reads content controls and table of contents' {
        $path = Join-Path $TestDrive 'WordReaderControls.docx'
        $imagePath = New-TestOfficeImageFile -Directory $TestDrive

        New-OfficeWord -Path $path {
            Add-OfficeWordParagraph { Add-OfficeWordContentControl -Text 'Client' -Alias 'ClientName' -Tag 'ClientTag' }
            Add-OfficeWordParagraph { Add-OfficeWordCheckBox -Checked -Alias 'Approved' -Tag 'ApprovedTag' }
            Add-OfficeWordParagraph { Add-OfficeWordCheckBox -Alias 'Rejected' -Tag 'RejectedTag' }
            Add-OfficeWordParagraph { Add-OfficeWordDatePicker -Date (Get-Date) -Alias 'DueDate' -Tag 'DueTag' }
            Add-OfficeWordParagraph { Add-OfficeWordDropDownList -Items 'Low', 'Medium', 'High' -Alias 'Priority' -Tag 'PriorityTag' }
            Add-OfficeWordParagraph { Add-OfficeWordComboBox -Items 'Red', 'Green' -Alias 'Color' -Tag 'ColorTag' }
            Add-OfficeWordParagraph { Add-OfficeWordPictureControl -Path $imagePath -Alias 'Logo' -Tag 'LogoTag' }
            Add-OfficeWordParagraph { Add-OfficeWordRepeatingSection -SectionTitle 'Items' -Alias 'LineItems' -Tag 'LineItemsTag' }
            $toc = Add-OfficeWordTableOfContents -PassThru
            Set-OfficeWordTableOfContents -TableOfContents $toc -Text 'Contents' -TextNoContent 'No entries'
            Update-OfficeWordTableOfContents -TableOfContents $toc
        }

        $controls = Get-OfficeWordContentControl -Path $path -Alias 'Client*'
        $controls.Count | Should -Be 1

        $checkBoxes = Get-OfficeWordCheckBox -Path $path -Alias 'Approved'
        $checkBoxes.Count | Should -Be 1
        $checkBoxes[0].IsChecked | Should -Be $true

        $unchecked = Get-OfficeWordCheckBox -Path $path -Unchecked
        $unchecked.Count | Should -Be 1

        $datePickers = Get-OfficeWordDatePicker -Path $path -Alias 'DueDate'
        $datePickers.Count | Should -Be 1
        $datePickers[0].Date | Should -Not -BeNullOrEmpty

        $dropDowns = Get-OfficeWordDropDownList -Path $path -Tag 'PriorityTag'
        $dropDowns.Count | Should -Be 1
        $dropDowns[0].Items | Should -Contain 'Medium'

        $comboBoxes = Get-OfficeWordComboBox -Path $path -Alias 'Color'
        $comboBoxes.Count | Should -Be 1
        $comboBoxes[0].Items | Should -Contain 'Red'

        $pictures = Get-OfficeWordPictureControl -Path $path -Alias 'Logo'
        $pictures.Count | Should -Be 1

        $repeating = Get-OfficeWordRepeatingSection -Path $path -Alias 'LineItems'
        $repeating.Count | Should -Be 1

        $toc = Get-OfficeWordTableOfContents -Path $path
        $toc | Should -Not -BeNullOrEmpty
        $toc.Text | Should -Be 'Contents'
        $toc.TextNoContent | Should -Be 'No entries'
    }

    It 'reads hyperlinks and document properties' {
        $path = Join-Path $TestDrive 'WordReaderLinksAndProperties.docx'

        New-OfficeWord -Path $path {
            Set-OfficeWordDocumentProperty -Name Title -Value 'Reader document'
            Set-OfficeWordDocumentProperty -Name Subject -Value 'Reader smoke test'
            Set-OfficeWordDocumentProperty -Name Ticket -Value 42 -Custom

            Add-OfficeWordParagraph {
                Add-OfficeWordText -Text 'See '
                Add-OfficeWordHyperlink -Text 'Example' -Url 'https://example.org/docs' -Styled
                Add-OfficeWordText -Text ' and '
                Add-OfficeWordHyperlink -Text 'Summary' -Anchor 'Summary'
            }

            Add-OfficeWordParagraph {
                Add-OfficeWordText -Text 'Summary destination'
                Add-OfficeWordBookmark -Name 'Summary'
            }
        } | Out-Null

        $links = Get-OfficeWordHyperlink -Path $path
        $links.Count | Should -Be 2

        $external = Get-OfficeWordHyperlink -Path $path -Url 'https://example.org/*'
        $external.Count | Should -Be 1
        $external[0].Text | Should -Be 'Example'

        $anchor = Get-OfficeWordHyperlink -Path $path -Anchor 'Summary'
        $anchor.Count | Should -Be 1
        $anchor[0].Anchor | Should -Be 'Summary'

        $properties = Get-OfficeWordDocumentProperty -Path $path
        ($properties | Where-Object { $_.Scope -eq 'BuiltIn' -and $_.Name -eq 'Title' } | Select-Object -First 1).Value | Should -Be 'Reader document'
        ($properties | Where-Object { $_.Scope -eq 'BuiltIn' -and $_.Name -eq 'Subject' } | Select-Object -First 1).Value | Should -Be 'Reader smoke test'
        ($properties | Where-Object { $_.Scope -eq 'Custom' -and $_.Name -eq 'Ticket' } | Select-Object -First 1).Value | Should -Be 42

        $customOnly = Get-OfficeWordDocumentProperty -Path $path -Custom
        $customOnly.Count | Should -Be 1
        $customOnly[0].CustomPropertyType | Should -Be 'NumberInteger'
    }

    It 'supports background images and preserved mail merge fields' {
        $path = Join-Path $TestDrive 'WordReaderBackgroundMerge.docx'
        $imagePath = New-TestOfficeImageFile -Directory $TestDrive

        New-OfficeWord -Path $path {
            Set-OfficeWordBackground -ImagePath $imagePath

            Add-OfficeWordParagraph {
                Add-OfficeWordText -Text 'Hello '
                Add-OfficeWordField -Type MergeField -Parameters '"FirstName"'
            }

            Invoke-OfficeWordMailMerge -InputObject ([pscustomobject]@{
                FirstName = 'Morgan'
            }) -PreserveFields
        } | Out-Null

        $doc = Get-OfficeWord -Path $path -ReadOnly
        try {
            $mergeFieldType = Get-TestPSWriteOfficeEnumValue -AssemblyName 'OfficeIMO.Word' -TypeName 'OfficeIMO.Word.WordFieldType' -Name 'MergeField'
            $doc.Fields.Where({ $_.FieldType -eq $mergeFieldType }).Count | Should -Be 1
            $doc.Fields.Where({ $_.FieldType -eq $mergeFieldType })[0].Text | Should -Be 'Morgan'
        } finally {
            $doc.Dispose()
        }

        $documentXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'word/document.xml'
        $namespaceManager = New-Object System.Xml.XmlNamespaceManager($documentXml.NameTable)
        $namespaceManager.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        $documentXml.SelectSingleNode('//w:background', $namespaceManager) | Should -Not -BeNullOrEmpty
        $documentXml.SelectSingleNode('//w:background/w:drawing', $namespaceManager) | Should -Not -BeNullOrEmpty
    }
}
