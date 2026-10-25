BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop

    $script:OfficeCommands = @(Get-Command -Module PSWriteOffice -CommandType Cmdlet)
    $script:OfficeCommandsByName = (Get-Module PSWriteOffice).ExportedCommands
}

Describe 'PSWriteOffice public API consistency' {
    It 'uses Path as the canonical source-file parameter name' {
        $filePathCommands = @($script:OfficeCommands | Where-Object { $_.Parameters.ContainsKey('FilePath') })
        $inputPathCommands = @($script:OfficeCommands | Where-Object { $_.Parameters.ContainsKey('InputPath') })

        $filePathCommands.Name | Should -BeNullOrEmpty
        @($inputPathCommands.Name | Sort-Object) | Should -Be @('ConvertTo-OfficeDocument', 'Export-OfficeConversionJob', 'Export-OfficeDocumentPdf')
    }

    It 'reserves Show for semantic visibility choices' {
        $allowed = @(
            'Set-OfficeExcelGridlines'
            'Set-OfficeExcelSheetVisibility'
        )
        $commands = @($script:OfficeCommands | Where-Object { $_.Parameters.ContainsKey('Show') })

        @($commands.Name | Sort-Object) | Should -Be @($allowed | Sort-Object)
    }

    It 'keeps format conversion out of New and Save lifecycle parameters' {
        $lifecycle = @($script:OfficeCommands | Where-Object { $_.Verb -in 'New', 'Save' })

        @($lifecycle | Where-Object { $_.Parameters.ContainsKey('PdfPath') }).Name | Should -BeNullOrEmpty
        @($lifecycle | Where-Object { $_.Verb -eq 'New' -and $_.Parameters.ContainsKey('AutoSave') }).Name | Should -BeNullOrEmpty
        (Get-Command Get-OfficeWord).Parameters.Keys | Should -Not -Contain 'AutoSave'
        (Get-Command Get-OfficeExcel).Parameters.Keys | Should -Not -Contain 'AutoSave'
    }

    It 'provides discoverable builders for common advanced option families' {
        $contracts = @(
            @{ Command = 'Compare-OfficeWordDocument'; Parameter = 'Options'; Builder = 'New-OfficeWordComparisonOptions' }
            @{ Command = 'Resolve-OfficeWordRevision'; Parameter = 'Filter'; Builder = 'New-OfficeWordRevisionFilter' }
            @{ Command = 'Get-OfficeDocumentHierarchy'; Parameter = 'ChunkingOptions'; Builder = 'New-OfficeReaderHierarchyOptions' }
            @{ Command = 'Compare-OfficePdfVisual'; Parameter = 'Options'; Builder = 'New-OfficePdfVisualComparisonOptions' }
            @{ Command = 'ConvertTo-OfficePdfWord'; Parameter = 'Options'; Builder = 'New-OfficePdfToWordOptions' }
            @{ Command = 'ConvertTo-OfficePdfExcel'; Parameter = 'Options'; Builder = 'New-OfficePdfExcelImportOptions' }
            @{ Command = 'ConvertTo-OfficePdfPowerPoint'; Parameter = 'Options'; Builder = 'New-OfficePdfToPowerPointOptions' }
            @{ Command = 'ConvertTo-OfficeOpenDocument'; Parameter = 'WordOptions'; Builder = 'New-OfficeWordOpenDocumentOptions' }
            @{ Command = 'ConvertTo-OfficeOpenDocument'; Parameter = 'ExcelOptions'; Builder = 'New-OfficeExcelOpenDocumentOptions' }
            @{ Command = 'ConvertTo-OfficeOpenDocument'; Parameter = 'PowerPointOptions'; Builder = 'New-OfficePowerPointOpenDocumentOptions' }
            @{ Command = 'Export-OfficeWordImage'; Parameter = 'Options'; Builder = 'New-OfficeWordImageOptions' }
            @{ Command = 'Export-OfficeExcelImage'; Parameter = 'Options'; Builder = 'New-OfficeExcelWorkbookImageOptions' }
            @{ Command = 'Export-OfficeExcelChartImage'; Parameter = 'Options'; Builder = 'New-OfficeExcelImageOptions' }
            @{ Command = 'Export-OfficePowerPointImage'; Parameter = 'Options'; Builder = 'New-OfficePowerPointImageOptions' }
            @{ Command = 'Export-OfficePdfImage'; Parameter = 'Options'; Builder = 'New-OfficePdfImageOptions' }
            @{ Command = 'Export-OfficeVisioImage'; Parameter = 'Options'; Builder = 'New-OfficeVisioImageOptions' }
            @{ Command = 'Export-OfficeHtmlImage'; Parameter = 'DocumentOptions'; Builder = 'New-OfficeHtmlConversionOptions' }
            @{ Command = 'Export-OfficeHtmlImage'; Parameter = 'RenderOptions'; Builder = 'New-OfficeHtmlRenderOptions' }
            @{ Command = 'Get-OfficeEmail'; Parameter = 'Options'; Builder = 'New-OfficeEmailReaderOptions' }
            @{ Command = 'Get-OfficeEmail'; Parameter = 'StoreOptions'; Builder = 'New-OfficeEmailStoreReaderOptions' }
            @{ Command = 'Save-OfficeEmail'; Parameter = 'Options'; Builder = 'New-OfficeEmailWriterOptions' }
            @{ Command = 'Get-OfficeEmailMailbox'; Parameter = 'Options'; Builder = 'New-OfficeEmailMailboxReaderOptions' }
            @{ Command = 'Save-OfficeEmailMailbox'; Parameter = 'Options'; Builder = 'New-OfficeEmailMailboxWriterOptions' }
        )

        foreach ($contract in $contracts) {
            $consumer = Get-Command $contract.Command
            $builder = Get-Command $contract.Builder
            $builder.OutputType.Type.Name | Should -Contain $consumer.Parameters[$contract.Parameter].ParameterType.Name
            @($builder.Parameters.Values | Where-Object ParameterType -eq ([System.Nullable[bool]])) |
                Should -HaveCount 0 -Because "$($contract.Builder) should expose PowerShell switches rather than nullable booleans"
            $helpPath = Join-Path (Join-Path $PSScriptRoot '..\Docs') "$($contract.Builder).md"
            Test-Path -LiteralPath $helpPath | Should -BeTrue -Because "$($contract.Builder) should have generated command help"
            $exampleText = Get-Content -LiteralPath $helpPath -Raw
            $exampleText | Should -Match ([regex]::Escape($contract.Command)) -Because "$($contract.Builder) should show how its result reaches the consuming command"
            $exampleText | Should -Not -Match "(?m)-[A-Za-z]+\s+'Value'" -Because "$($contract.Builder) should not publish generated placeholder examples"
        }
    }

    It 'offers PassThru on commands that save artifacts' {
        $missing = @($script:OfficeCommands | Where-Object { $_.Verb -eq 'Save' -and -not $_.Parameters.ContainsKey('PassThru') })

        $missing.Name | Should -BeNullOrEmpty
    }

    It 'offers PassThru on commands that perform mutations' {
        $mutationVerbs = @('Add', 'Clear', 'Copy', 'Edit', 'Move', 'Protect', 'Remove', 'Rename', 'Set', 'Unprotect', 'Update')
        $intentionalTransforms = @('Set-OfficeConfluenceManagedSection', 'Set-OfficePdfCompliance', 'Set-OfficePdfElectronicInvoice')
        $missing = @(
            $script:OfficeCommands |
                Where-Object {
                    $_.Verb -in $mutationVerbs -and
                    $_.Name -notin $intentionalTransforms -and
                    -not $_.Parameters.ContainsKey('PassThru')
                }
        )

        $missing.Name | Should -BeNullOrEmpty
    }

    It 'exposes one format-neutral PDF export command' {
        $command = Get-Command Export-OfficeDocumentPdf

        $command.Parameters.Keys | Should -Contain 'InputPath'
        $command.Parameters.Keys | Should -Contain 'Path'
        $command.Parameters.Keys | Should -Contain 'Document'
        $command.Parameters.Keys | Should -Contain 'PassThru'
        $command.Parameters.Keys | Should -Contain 'Open'
        $command.ParameterSets.Name | Should -Contain 'Document'
        $command.ParameterSets.Name | Should -Contain 'Path'
    }

    It 'accepts a password when a source path string is piped to PDF export' {
        $sourcePath = Join-Path $TestDrive 'encrypted.docx'
        $outputPath = Join-Path $TestDrive 'encrypted.pdf'

        { $sourcePath | Export-OfficeDocumentPdf -Path $outputPath -Password 'secret' -WhatIf -ErrorAction Stop } |
            Should -Not -Throw
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

    It 'provides a discoverable PowerShell builder for every PDF export options type' {
        $command = Get-Command Export-OfficeDocumentPdf
        $builders = [ordered]@{
            WordOptions       = 'New-OfficeWordPdfOptions'
            ExcelOptions      = 'New-OfficeExcelPdfOptions'
            PowerPointOptions = 'New-OfficePowerPointPdfOptions'
            MarkdownOptions   = 'New-OfficeMarkdownPdfOptions'
            RtfOptions        = 'New-OfficeRtfPdfOptions'
        }

        foreach ($entry in $builders.GetEnumerator()) {
            $builder = Get-Command $entry.Value
            $builder.OutputType.Type.Name | Should -Contain $command.Parameters[$entry.Key].ParameterType.Name
        }

        foreach ($builderName in $builders.Values) {
            $builder = Get-Command $builderName
            @($builder.Parameters.Values | Where-Object ParameterType -eq ([System.Nullable[bool]])) |
                Should -HaveCount 0 -Because "$builderName should expose idiomatic switches rather than nullable .NET booleans"
        }
    }

    It 'does not require raw writer or OpenDocument load option objects for common controls' {
        $openDocument = Get-Command Get-OfficeOpenDocument
        foreach ($parameter in @(
            'Password',
            'MaxPackageBytes',
            'MaxEntries',
            'MaxEntryUncompressedBytes',
            'MaxTotalUncompressedBytes',
            'MaxTotalKdfIterations',
            'MaxCompressionRatio',
            'MaxDepth',
            'MaxXmlCharacters',
            'MaxXmlDepth'
        )) {
            $openDocument.Parameters.Keys | Should -Contain $parameter
        }

        foreach ($name in 'Save-OfficeAsciiDoc', 'Save-OfficeLatex') {
            $command = Get-Command $name
            $command.Parameters.Keys | Should -Contain 'Mode'
            $command.Parameters.Keys | Should -Contain 'LineEnding'
        }
    }

    It 'uses Path and DestinationPath for file copies' {
        $command = Get-Command Copy-OfficeExcelWorkbook

        $command.Parameters.Keys | Should -Contain 'Path'
        $command.Parameters.Keys | Should -Contain 'DestinationPath'
        $command.Parameters.Keys | Should -Not -Contain 'InputPath'
        $command.Parameters['Path'].Aliases | Should -Contain 'InputPath'
        $command.Parameters['DestinationPath'].Aliases | Should -Contain 'OutputPath'
    }

    It 'uses one explicit persistence contract when closing Office documents' {
        foreach ($name in 'Close-OfficeWord', 'Close-OfficeExcel', 'Close-OfficePowerPoint') {
            $command = Get-Command $name
            $command.Parameters.Keys | Should -Contain 'Save' -Because "$name should make persistence explicit"
            $command.Parameters.Keys | Should -Contain 'Path' -Because "$name should support save-as while closing"
            $command.Parameters.Keys | Should -Contain 'Open' -Because "$name should use the common viewer switch"
            $command.Parameters.Keys | Should -Contain 'WhatIf' -Because "$name can save or dispose a live document"
            $command.Parameters.Keys | Should -Contain 'Confirm' -Because "$name can save or dispose a live document"
        }

        $word = New-OfficeWord -Path (Join-Path $TestDrive 'close-word.docx') -NoSave
        $excel = New-OfficeExcel -Path (Join-Path $TestDrive 'close-excel.xlsx') -NoSave
        $powerPoint = New-OfficePowerPoint -Path (Join-Path $TestDrive 'close-powerpoint.pptx') -NoSave
        try {
            { $word | Close-OfficeWord -Open -ErrorAction Stop } | Should -Throw '*Use -Save or -Path with -Open*'
            { $excel | Close-OfficeExcel -Open -ErrorAction Stop } | Should -Throw '*Use -Save or -Path with -Open*'
            { $powerPoint | Close-OfficePowerPoint -Open -ErrorAction Stop } | Should -Throw '*Use -Save or -Path with -Open*'
        } finally {
            if ($word) { $word | Close-OfficeWord -Confirm:$false }
            if ($excel) { $excel | Close-OfficeExcel -Confirm:$false }
            if ($powerPoint) { $powerPoint | Close-OfficePowerPoint -Confirm:$false }
        }
    }

    It 'rejects Open when NoSave prevents a file from being written' {
        $cases = @(
            { New-OfficeWord -Path (Join-Path $TestDrive 'word.docx') -NoSave -Open -ErrorAction Stop }
            { New-OfficeExcel -Path (Join-Path $TestDrive 'excel.xlsx') -NoSave -Open -ErrorAction Stop }
            { New-OfficePowerPoint -Path (Join-Path $TestDrive 'powerpoint.pptx') -NoSave -Open -ErrorAction Stop }
            { New-OfficePdf -Path (Join-Path $TestDrive 'document.pdf') -NoSave -Open -ErrorAction Stop }
            { New-OfficeVisio -Path (Join-Path $TestDrive 'visio.vsdx') -NoSave -Open -ErrorAction Stop }
        )

        foreach ($case in $cases) {
            $case | Should -Throw '*-Open cannot be used with -NoSave*'
        }
    }

    It 'pipes a saved FileInfo directly into format-neutral PDF export' {
        $wordPath = Join-Path $TestDrive 'pipeline.docx'
        $pdfPath = Join-Path $TestDrive 'pipeline.pdf'

        $pdf = New-OfficeWord -Path $wordPath -PassThru -Content {
            Add-OfficeWordParagraph -Text 'Pipeline contract'
        } | Export-OfficeDocumentPdf -Path $pdfPath -PassThru

        $pdf | Should -BeOfType System.IO.FileInfo
        $pdf.FullName | Should -Be $pdfPath
        $pdf.Length | Should -BeGreaterThan 0
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
                    Word { New-OfficeWord -Path $quietPath }
                    Excel { New-OfficeExcel -Path $quietPath }
                    PowerPoint { New-OfficePowerPoint -Path $quietPath }
                    Markdown { New-OfficeMarkdown -Path $quietPath }
                    PDF { New-OfficePdf -Path $quietPath { PdfParagraph 'API contract' } }
                    Visio { New-OfficeVisio -Path $quietPath }
                    OpenDocument { New-OfficeOpenDocument -Kind Text -Path $quietPath }
                })
            $quietOutput | Should -HaveCount 0 -Because "$($case.Name) should be quiet by default"
            Test-Path -LiteralPath $quietPath | Should -BeTrue

            $passThruPath = Join-Path $TestDrive ("passthru-{0}.{1}" -f $case.Name, $case.Extension)
            $passThruOutput = @(switch ($case.Name) {
                    Word { New-OfficeWord -Path $passThruPath -PassThru }
                    Excel { New-OfficeExcel -Path $passThruPath -PassThru }
                    PowerPoint { New-OfficePowerPoint -Path $passThruPath -PassThru }
                    Markdown { New-OfficeMarkdown -Path $passThruPath -PassThru }
                    PDF { New-OfficePdf -Path $passThruPath -PassThru { PdfParagraph 'API contract' } }
                    Visio { New-OfficeVisio -Path $passThruPath -PassThru }
                    OpenDocument { New-OfficeOpenDocument -Kind Text -Path $passThruPath -PassThru }
                })
            $passThruOutput | Should -HaveCount 1 -Because "$($case.Name) PassThru should emit the saved file"
            $passThruOutput[0] | Should -BeOfType System.IO.FileInfo
            $passThruOutput[0].FullName | Should -Be $passThruPath
        }
    }

    It 'offers matching Office-prefixed DSL roots without replacing canonical New commands' {
        $roots = [ordered]@{
            OfficeWord         = 'New-OfficeWord'
            OfficeExcel        = 'New-OfficeExcel'
            OfficePowerPoint   = 'New-OfficePowerPoint'
            OfficePdf          = 'New-OfficePdf'
            OfficeMarkdown     = 'New-OfficeMarkdown'
            OfficeRtf          = 'New-OfficeRtf'
            OfficeVisio        = 'New-OfficeVisio'
            OfficeOpenDocument = 'New-OfficeOpenDocument'
        }

        foreach ($entry in $roots.GetEnumerator()) {
            $alias = Get-Command $entry.Key
            $alias | Should -BeOfType System.Management.Automation.AliasInfo
            $alias.ResolvedCommandName | Should -Be $entry.Value
        }

        (Get-Command VisioStencilImport).ResolvedCommandName | Should -Be 'Import-OfficeVisioStencil'
    }

    It 'preserves published PDF import option builders as compatibility aliases' {
        $aliases = [ordered]@{
            'New-OfficePdfWordImportOptions' = 'New-OfficePdfToWordOptions'
            'New-OfficePdfPowerPointImportOptions' = 'New-OfficePdfToPowerPointOptions'
        }

        foreach ($entry in $aliases.GetEnumerator()) {
            $alias = Get-Command $entry.Key
            $alias | Should -BeOfType System.Management.Automation.AliasInfo
            $alias.ResolvedCommandName | Should -Be $entry.Value
        }
    }

    It 'uses enum parameter types for closed value domains' {
        $contracts = @(
            @{ Command = 'Add-OfficeExcelTable'; Parameter = 'TableStyle'; Type = 'ExcelTableStyle' }
            @{ Command = 'Add-OfficeExcelConditionalRule'; Parameter = 'RuleType'; Type = 'OfficeExcelConditionalRuleType' }
            @{ Command = 'Add-OfficeExcelConditionalRule'; Parameter = 'Operator'; Type = 'ExcelConditionalFormattingOperator' }
            @{ Command = 'Add-OfficeExcelPivotTable'; Parameter = 'DataFunction'; Type = 'ExcelPivotDataFunction[]' }
            @{ Command = 'Add-OfficeExcelPivotTable'; Parameter = 'Layout'; Type = 'ExcelPivotLayout' }
            @{ Command = 'Import-OfficeExcel'; Parameter = 'FormulaMode'; Type = 'OfficeExcelFormulaMode' }
            @{ Command = 'Add-OfficeWordTable'; Parameter = 'Layout'; Type = 'OfficeWordTableLayout' }
            @{ Command = 'Set-OfficeWordParagraphStyle'; Parameter = 'Alignment'; Type = 'WordParagraphAlignment' }
            @{ Command = 'Set-OfficeWordTextStyle'; Parameter = 'Underline'; Type = 'WordUnderlineStyle' }
            @{ Command = 'Add-OfficePowerPointShape'; Parameter = 'ShapeType'; Type = 'OfficePresetShapeType' }
            @{ Command = 'Find-OfficePowerPointShape'; Parameter = 'Kind'; Type = 'OfficePowerPointShapeKind[]' }
            @{ Command = 'Set-OfficePdfPage'; Parameter = 'BoxName'; Type = 'PdfPageBoundaryBox' }
            @{ Command = 'Save-OfficeAsciiDoc'; Parameter = 'LineEnding'; Type = 'OfficeLineEnding' }
        )

        foreach ($contract in $contracts) {
            $type = (Get-Command $contract.Command).Parameters[$contract.Parameter].ParameterType
            $nullableType = [Nullable]::GetUnderlyingType($type)
            $actual = if ($nullableType) { $nullableType.Name } else { $type.Name }
            $actual | Should -Be $contract.Type -Because "$($contract.Command) -$($contract.Parameter) is a closed value domain"
        }
    }

    It 'normalizes and completes open color values consistently' {
        $colorParameters = @(
            foreach ($command in $script:OfficeCommands) {
                foreach ($parameter in $command.Parameters.Values) {
                    if ($parameter.ParameterType -in [string], [string[]] -and $parameter.Name -match 'Color') {
                        [pscustomobject]@{ Command = $command.Name; Parameter = $parameter }
                    }
                }
            }
            foreach ($contract in @(
                    @{ Command = 'Add-OfficePdfTable'; Parameter = 'HeaderFill' }
                    @{ Command = 'Add-OfficePdfTable'; Parameter = 'RowStripeFill' }
                    @{ Command = 'Set-OfficeExcelCell'; Parameter = 'GradientFrom' }
                    @{ Command = 'Set-OfficeExcelCell'; Parameter = 'GradientTo' }
                    @{ Command = 'Set-OfficePowerPointThemeColor'; Parameter = 'Value' }
                )) {
                [pscustomobject]@{
                    Command = $contract.Command
                    Parameter = (Get-Command $contract.Command).Parameters[$contract.Parameter]
                }
            }
        )

        $colorParameters | Should -Not -BeNullOrEmpty
        foreach ($entry in $colorParameters) {
            $transformations = @($entry.Parameter.Attributes | Where-Object { $_.TypeId.Name -eq 'OfficeColorArgumentTransformationAttribute' })
            $transformations | Should -HaveCount 1 -Because "$($entry.Command) -$($entry.Parameter.Name) should accept the same named and hexadecimal color language"
            if ($entry.Command -like '*OfficeExcel*' -and $entry.Command -notlike '*ImageOptions') {
                $transformations[0].UseExcelArgb | Should -BeTrue -Because "$($entry.Command) -$($entry.Parameter.Name) should normalize shorthand without reordering legacy Excel ARGB"
            }
            @($entry.Parameter.Attributes | Where-Object { $_.TypeId.Name -eq 'ArgumentCompleterAttribute' }) |
                Should -HaveCount 1 -Because "$($entry.Command) -$($entry.Parameter.Name) should offer named color completion"
        }

        (New-OfficeTextRun -Text 'status' -Color Red).Color | Should -Be '#FF0000'
        (New-OfficeTextRun -Text 'status' -Color '#abc').Color | Should -Be '#AABBCC'
        (New-OfficeTextRun -Text 'status' -Color '#abcd').Color | Should -Be '#abcd'
        (New-OfficeTextRun -Text 'status' -Color '#FF4F81BD').Color | Should -Be '#FF4F81BD'
        (New-OfficeTextRun -Text 'status' -BackgroundColor None).BackgroundColor | Should -Be 'None'
        { New-OfficeTextRun -Text 'status' -Color 'not-a-color' -ErrorAction Stop } | Should -Throw '*known Office color name*'

        $excelColorTransform = @((Get-Command Set-OfficeExcelCell).Parameters.BackgroundColor.Attributes |
                Where-Object { $_.TypeId.Name -eq 'OfficeColorArgumentTransformationAttribute' })[0]
        $excelColorTransform.UseExcelArgb | Should -BeTrue
        $excelColorTransform.Transform($ExecutionContext, '#abc') | Should -Be '#AABBCC'
        $excelColorTransform.Transform($ExecutionContext, '#abcd') | Should -Be '#DDAABBCC'
        $excelColorTransform.Transform($ExecutionContext, '#FF4F81BD') | Should -Be '#FF4F81BD'

        $input = 'New-OfficeTextRun -Color Re'
        $completion = [System.Management.Automation.CommandCompletion]::CompleteInput($input, $input.Length, $null)
        $completion.CompletionMatches.CompletionText | Should -Contain 'RebeccaPurple'
        $completion.CompletionMatches.CompletionText | Should -Contain 'Red'

        $input = 'New-OfficeTextRun -BackgroundColor N'
        $completion = [System.Management.Automation.CommandCompletion]::CompleteInput($input, $input.Length, $null)
        $completion.CompletionMatches.CompletionText | Should -Contain 'None'
    }

    It 'guides open string domains with completion instead of false enums' {
        $cases = @(
            @{ Input = 'Set-OfficePdfPageSetup -PageSize A'; Expected = 'A4' }
            @{ Input = 'Add-OfficePdfTable -TableStyle Tech'; Expected = 'TechnicalDocument' }
            @{ Input = 'Add-OfficeMarkdownCallout -Kind W'; Expected = 'Warning' }
            @{ Input = 'New-OfficeTextRun -UnderlineStyle D'; Expected = 'Double' }
            @{ Input = 'New-OfficeTextRun -Baseline N'; Expected = 'Normal' }
            @{ Input = 'New-OfficeTextRun -TabLeader H'; Expected = 'Hyphens' }
            @{ Input = 'New-OfficeTextRun -TabLeader U'; Expected = 'Underscores' }
            @{ Input = 'New-OfficeTextRun -TabAlignment D'; Expected = 'DecimalSeparator' }
            @{ Input = 'Add-OfficeExcelPivotTable -PivotStyle PivotStyleMedium9'; Expected = 'PivotStyleMedium9' }
            @{ Input = 'Get-OfficeDocumentAsset -Kind Im'; Expected = 'Image' }
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
