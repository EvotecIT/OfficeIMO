BeforeAll {
    $manifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) { $env:PSWRITEOFFICE_MODULE_MANIFEST } else { Join-Path $PSScriptRoot '..\PSWriteOffice.psd1' }
    Import-Module $manifest -Global -ErrorAction Stop
}

Describe 'Document asset export' {
    It 'writes all table formats as flat native materialization results' {
        $path = Join-Path $TestDrive 'table[1].md'
        Set-Content -LiteralPath $path -Value "| Item | Value |`n| --- | --- |`n| Delivery | 42 |" -Encoding UTF8
        $directory = Join-Path $TestDrive 'tables'
        $results = @(Get-Item -LiteralPath $path | Export-OfficeDocumentAsset -OutputDirectory $directory -Type Table)
        $results | Should -HaveCount 3
        @($results | Where-Object Written) | Should -HaveCount 3
        $results[0].GetType().Name | Should -Be 'ReaderTableMaterializedExport'
        @($results | ForEach-Object { $_.Format.ToString() } | Sort-Object) -join ',' | Should -Be 'Csv,Json,Markdown'
        foreach ($result in $results) { Test-Path -LiteralPath $result.Path | Should -BeTrue }
        Get-Content -LiteralPath ($results | Where-Object { $_.Format -eq 'Csv' }).Path -Raw | Should -Match 'Delivery'
    }

    It 'honors WhatIf without parsing the source or creating the destination' {
        $path = Join-Path $TestDrive 'unparseable.docx'
        Set-Content -LiteralPath $path -Value 'This is not a document package.'
        $directory = Join-Path $TestDrive 'what-if'
        @(Export-OfficeDocumentAsset $path $directory -WhatIf -ErrorAction Stop) | Should -HaveCount 0
        Test-Path -LiteralPath $directory | Should -BeFalse
    }

    It 'filters embedded media and preserves existing payloads with NoClobber' {
        $path = Join-Path $TestDrive 'media.docx'
        $image = Join-Path $PSScriptRoot 'Assets\CellImage.png'
        New-OfficeWord -Path $path {
            WordSection { WordParagraph { WordImage -Path $image -Width 24 -Height 24 } }
        } | Out-Null
        $directory = Join-Path $TestDrive 'media'
        $first = @(Export-OfficeDocumentAsset $path $directory -Kind image -Extension png -MediaType image/png -ValidatePayloadHash)
        $first | Should -HaveCount 1
        $first[0].Written | Should -BeTrue
        $first[0].GetType().Name | Should -Be 'OfficeDocumentMaterializedAsset'
        [IO.File]::WriteAllText($first[0].Path, 'Existing payload')
        $again = @(Export-OfficeDocumentAsset $path $directory -NoClobber)
        $again | Should -HaveCount 1
        $again[0].Written | Should -BeFalse
        $again[0].SkippedReason | Should -Not -BeNullOrEmpty
        Get-Content -LiteralPath $first[0].Path -Raw | Should -BeExactly 'Existing payload'
        @(Export-OfficeDocumentAsset $path $directory -ValidatePayloadHash)[0].Written | Should -BeTrue
        (Get-FileHash -LiteralPath $first[0].Path).Hash | Should -Be (Get-FileHash -LiteralPath $image).Hash
    }

    It 'writes visual payloads and JSON independently without rendering or executing the visual' {
        $path = Join-Path $TestDrive 'diagram.md'
        Set-Content -LiteralPath $path -Value '# Diagram', '', '```mermaid', 'graph TD', 'A --> B', '```' -Encoding UTF8
        $all = @(Export-OfficeDocumentAsset $path (Join-Path $TestDrive 'visual-all') -Type Visual -Indented)
        $all | Should -HaveCount 2
        @($all | Where-Object Written) | Should -HaveCount 2
        $all[0].GetType().Name | Should -Be 'ReaderVisualMaterializedExport'
        $metadata = @(Export-OfficeDocumentAsset $path (Join-Path $TestDrive 'visual-metadata') -Type Visual -NoPayload)
        $metadata | Should -HaveCount 1
        $metadata[0].Format.ToString() | Should -Be 'Json'
        (Get-Content -LiteralPath $metadata[0].Path -Raw | ConvertFrom-Json).kind | Should -Be 'mermaid'
        $payload = @(Export-OfficeDocumentAsset $path (Join-Path $TestDrive 'visual-payload') -Type Visual -NoMetadata)
        $payload | Should -HaveCount 1
        $payload[0].Format.ToString() | Should -Be 'Payload'
        Get-Content -LiteralPath $payload[0].Path -Raw | Should -Match 'A --> B'
    }

    It 'keeps valid FileInfo inputs when another source is missing' {
        $path = Join-Path $TestDrive 'resilient.md'
        Set-Content -LiteralPath $path -Value "| Item |`n| --- |`n| Delivery |" -Encoding UTF8
        $errors = @()
        $results = @(Export-OfficeDocumentAsset -LiteralPath @((Join-Path $TestDrive 'missing.md'), $path) -OutputDirectory (Join-Path $TestDrive 'resilient') -Type Table -TableFormat Csv -ErrorAction SilentlyContinue -ErrorVariable +errors)
        $results | Should -HaveCount 1
        $results[0].Written | Should -BeTrue
        $errors | Should -HaveCount 1
    }

    It 'selects a worksheet by position and routes the A1 range to the native reader' {
        $path = Join-Path $TestDrive 'sheets.xlsx'
        $workbook = [OfficeIMO.Excel.ExcelDocument]::Create()
        try {
            $first = $workbook.AddWorksheet('2')
            $null = $first.CellValue(1, 1, 'Header')
            $null = $first.CellValue(2, 1, 'Excluded')
            $second = $workbook.AddWorksheet('Evidence')
            $null = $second.CellValue(1, 1, 'Header')
            $null = $second.CellValue(2, 1, 'Selected')
            $null = $second.CellValue(3, 1, 'Outside range')
            $workbook.Save($path)
        } finally { $workbook.Dispose() }
        $results = @(Export-OfficeDocumentAsset $path (Join-Path $TestDrive 'selected-sheet') -Type Table -TableFormat Csv -Sheet 2 -Range A1:A2)
        $results | Should -HaveCount 1
        $text = Get-Content -LiteralPath $results[0].Path -Raw
        $text | Should -Match 'Selected'
        $text | Should -Not -Match 'Excluded|Outside range'
    }

    It 'enforces source admission without mutating caller options or creating output files' {
        $path = Join-Path $TestDrive 'bounded.md'
        Set-Content -LiteralPath $path -Value "| Item |`n| --- |`n| Delivery |" -Encoding UTF8
        $directory = Join-Path $TestDrive 'bounded'
        $options = New-OfficeDocumentReaderOptions
        $options.MaxInputBytes = 1000
        { Export-OfficeDocumentAsset $path $directory -Type Table -Options $options -MaximumInputBytes 1 -ErrorAction Stop } | Should -Throw
        $options.MaxInputBytes | Should -Be 1000
        Test-Path -LiteralPath $directory | Should -BeFalse
    }

    It 'leaves one complete output when two sources share a native filename and NoClobber is set' {
        $paths = foreach ($name in 'first', 'second') {
            $folder = Join-Path $TestDrive $name
            $null = New-Item -ItemType Directory -Path $folder
            $path = Join-Path $folder 'same.md'
            Set-Content -LiteralPath $path -Value "| Item |`n| --- |`n| $name |" -Encoding UTF8
            $path
        }
        $results = @(Export-OfficeDocumentAsset $paths (Join-Path $TestDrive 'colliding') -Type Table -TableFormat Csv -NoClobber -ThrottleLimit 2)
        $results | Should -HaveCount 2
        @($results | Where-Object Written) | Should -HaveCount 1
        @($results | Where-Object { -not $_.Written }) | Should -HaveCount 1
        $text = Get-Content -LiteralPath $results[0].Path -Raw
        $text | Should -Match 'first|second'
    }

    It 'rejects switches that do not apply to the selected content before creating the destination' {
        $path = Join-Path $TestDrive 'invalid.md'
        Set-Content -LiteralPath $path -Value '# Evidence'
        $directory = Join-Path $TestDrive 'invalid'
        { Export-OfficeDocumentAsset $path $directory -Type Table -NoPayload -ErrorAction Stop } | Should -Throw '*Visual*'
        { Export-OfficeDocumentAsset $path $directory -Type Visual -TableFormat Csv -ErrorAction Stop } | Should -Throw '*Table*'
        { Export-OfficeDocumentAsset $path $directory -Type Visual -NoPayload -NoMetadata -ErrorAction Stop } | Should -Throw '*both*'
        Test-Path -LiteralPath $directory | Should -BeFalse
    }
}
