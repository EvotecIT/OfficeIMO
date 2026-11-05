BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')

}

Describe 'PSWriteOffice public API consistency' {
    It 'binds a password on the piped file route without writing under WhatIf' {
        $sourcePath = Join-Path $TestDrive 'encrypted.docx'
        $outputPath = Join-Path $TestDrive 'encrypted.pdf'

        Set-Content -LiteralPath $sourcePath -Value 'WhatIf must not load this source'
        { $sourcePath | Convert-OfficeDocument -OutputPath $outputPath -Password 'secret' -WhatIf -ErrorAction Stop } |
            Should -Not -Throw
        Test-Path -LiteralPath $outputPath | Should -BeFalse
    }

    It 'normalizes filesystem directories used as HTML base URIs' {
        $assetsPath = Join-Path $TestDrive 'Assets'
        New-Item -ItemType Directory -Path $assetsPath | Out-Null

        foreach ($options in @(
            New-OfficeHtmlConversionOptions -BaseUri $assetsPath
            New-OfficeHtmlRenderOptions -BaseUri $assetsPath
        )) {
            $options.BaseUri.IsAbsoluteUri | Should -BeTrue
            $options.BaseUri.IsFile | Should -BeTrue
            $options.BaseUri.AbsoluteUri | Should -Match '/$'
        }

        (New-OfficeHtmlConversionOptions -BaseUri 'https://example.test/assets/').BaseUri.AbsoluteUri |
            Should -Be 'https://example.test/assets/'
    }

    It 'converts a saved FileInfo pipeline and retains the written-file report' {
        $wordPath = Join-Path $TestDrive 'pipeline.docx'
        $pdfPath = Join-Path $TestDrive 'pipeline.pdf'

        $report = New-OfficeWord -Path $wordPath -PassThru -Content {
            Add-OfficeWordParagraph -Text 'Pipeline contract'
        } | Convert-OfficeDocument -OutputPath $pdfPath

        $report | Should -BeOfType OfficeIMO.OfficeConversionReport
        $report.OutputPath | Should -Be $pdfPath
        (Get-Item -LiteralPath $pdfPath).Length | Should -BeGreaterThan 0
    }

    It 'keeps saved New commands quiet unless PassThru is requested' {
        $cases = @(
            @{ Name = 'Word'; Extension = 'docx' }
            @{ Name = 'Excel'; Extension = 'xlsx' }
            @{ Name = 'PowerPoint'; Extension = 'pptx' }
            @{ Name = 'Markdown'; Extension = 'md' }
            @{ Name = 'PDF'; Extension = 'pdf' }
            @{ Name = 'Visio'; Extension = 'vsdx' }
            @{ Name = 'OpenDocument'; Extension = 'odt' }
        )

        foreach ($case in $cases) {
            $quietPath = Join-Path $TestDrive ("quiet-{0}.{1}" -f $case.Name, $case.Extension)
            $quietOutput = @(switch ($case.Name) {
                    Word { New-OfficeWord -Path $quietPath {} }
                    Excel { New-OfficeExcel -Path $quietPath {} }
                    PowerPoint { New-OfficePowerPoint -Path $quietPath {} }
                    Markdown { New-OfficeMarkdown -Path $quietPath -Content {} }
                    PDF { New-OfficePdf -Path $quietPath { PdfParagraph 'API contract' } }
                    Visio { New-OfficeVisio -Path $quietPath {} }
                    OpenDocument { New-OfficeOpenDocument -Kind Text -Path $quietPath {} }
                })
            $quietOutput | Should -HaveCount 0 -Because "$($case.Name) should be quiet by default"
            Test-Path -LiteralPath $quietPath | Should -BeTrue

            $passThruPath = Join-Path $TestDrive ("passthru-{0}.{1}" -f $case.Name, $case.Extension)
            $passThruOutput = @(switch ($case.Name) {
                    Word { New-OfficeWord -Path $passThruPath -PassThru {} }
                    Excel { New-OfficeExcel -Path $passThruPath -PassThru {} }
                    PowerPoint { New-OfficePowerPoint -Path $passThruPath -PassThru {} }
                    Markdown { New-OfficeMarkdown -Path $passThruPath -Content {} -PassThru }
                    PDF { New-OfficePdf -Path $passThruPath -PassThru { PdfParagraph 'API contract' } }
                    Visio { New-OfficeVisio -Path $passThruPath -PassThru {} }
                    OpenDocument { New-OfficeOpenDocument -Kind Text -Path $passThruPath -PassThru {} }
                })
            $passThruOutput | Should -HaveCount 1 -Because "$($case.Name) PassThru should emit the saved file"
            $passThruOutput[0] | Should -BeOfType System.IO.FileInfo
            $passThruOutput[0].FullName | Should -Be $passThruPath
        }
    }

    It 'binds named and hexadecimal colors on native Word runs' {
        $document = New-OfficeWord
        try {
            $paragraph = $document | Add-OfficeWordParagraph -Text 'Retained ' -PassThru
            $red = $paragraph | Add-OfficeWordText -Text 'red' -Color Red -PassThru
            $short = $paragraph | Add-OfficeWordText -Text ' shorthand' -Color '#abc' -PassThru
            $red.Color.ToRgbHex() | Should -Be 'FF0000'
            $short.Color.ToRgbHex() | Should -Be 'AABBCC'
            $chartRows = @([pscustomobject]@{ Label = 'North'; Sales = 10; Returns = 2 })
            $chartRows | Add-OfficeWordChart -Document $document -Type Bar -Category Label -Value Sales,Returns -SeriesColor Red,'#0f0'
            $path = Join-Path $TestDrive 'NativeColorBinding.docx'
            Save-OfficeWord -Document $document -Path $path
            $entry = @(Get-ZipEntriesLocal -Path $path | Where-Object { $_ -match '^word/charts/chart\d+\.xml$' })
            $entry.Count | Should -Be 1
            $chart = Get-ZipXmlDocumentLocal -Path $path -Entry $entry[0]
            @($chart.SelectNodes("//*[local-name()='ser']/*[local-name()='spPr']/*[local-name()='solidFill']/*[local-name()='srgbClr']") |
                ForEach-Object val) | Should -Be @('FF0000', '00FF00')
            $before = $paragraph.Text
            { $paragraph | Add-OfficeWordText -Text 'invalid' -Color 'not-a-color' -ErrorAction Stop } |
                Should -Throw '*Expected an OfficeColor*'
            $paragraph.Text | Should -Be $before

            $input = 'Add-OfficeWordText -Color Re'
            $completion = [System.Management.Automation.CommandCompletion]::CompleteInput($input, $input.Length, $null)
            $completion.CompletionMatches.CompletionText | Should -Contain 'RebeccaPurple'
            $completion.CompletionMatches.CompletionText | Should -Contain 'Red'
        } finally {
            Close-OfficeWord -Document $document -Confirm:$false
        }
    }

    It 'guides open string domains with completion instead of false enums' {
        $cases = @(
            @{ Input = 'Set-OfficePdfPageSetup -PageSize A'; Expected = 'A4' }
            @{ Input = 'Add-OfficePdfTable -TableStyle Tech'; Expected = 'TechnicalDocument' }
            @{ Input = 'Add-OfficeMarkdownCallout -Kind W'; Expected = 'Warning' }
            @{ Input = 'Add-OfficeExcelPivotTable -Style PivotStyleMedium9'; Expected = 'PivotStyleMedium9' }
            @{ Input = 'Export-OfficeDocumentAsset -Kind Im'; Expected = 'Image' }
        )

        foreach ($case in $cases) {
            $completion = [System.Management.Automation.CommandCompletion]::CompleteInput($case.Input, $case.Input.Length, $null)
            $completion.CompletionMatches.CompletionText | Should -Contain $case.Expected
        }
    }

    It 'bounds native Visio index domains without representing galleries as enums' {
        $contracts = @(
            @{ Command = 'Add-OfficeVisioConnector'; Parameter = 'LinePattern'; Min = 0; Max = 23 }
            @{ Command = 'Add-OfficeVisioRectangle'; Parameter = 'LinePattern'; Min = 0; Max = 23 }
            @{ Command = 'Add-OfficeVisioRectangle'; Parameter = 'FillPattern'; Min = 0; Max = 40 }
            @{ Command = 'Add-OfficeVisioStencilShape'; Parameter = 'LinePattern'; Min = 0; Max = 23 }
            @{ Command = 'Add-OfficeVisioStencilShape'; Parameter = 'FillPattern'; Min = 0; Max = 40 }
        )

        foreach ($contract in $contracts) {
            $parameter = (Get-Command $contract.Command).Parameters[$contract.Parameter]
            $parameter.ParameterType | Should -Be ([Nullable[int]])
            $range = $parameter.Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateRangeAttribute] }
            $range.MinRange | Should -Be $contract.Min
            $range.MaxRange | Should -Be $contract.Max
        }

        (Get-Command Add-OfficeVisioContainer).Parameters['ContainerStyle'].ParameterType | Should -Be ([Nullable[int]])
        (Get-Command Add-OfficeVisioContainer).Parameters['HeadingStyle'].ParameterType | Should -Be ([Nullable[int]])
    }
}
