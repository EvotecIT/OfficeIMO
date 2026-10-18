BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
}

Describe 'Authenticated PDF automation' {
    It 'exposes an explicit authenticated permission override on PDF readers' {
        foreach ($name in @(
            'Add-OfficePdfCanvas',
            'Add-OfficePdfPageOverlay',
            'Add-OfficePdfStamp',
            'ConvertTo-OfficePdfFlatAnnotation',
            'ConvertTo-OfficePdfFlatForm',
            'ConvertTo-OfficePdfHtml',
            'ConvertTo-OfficePdfMarkdown',
            'ConvertTo-OfficePdfOptimized',
            'ConvertTo-OfficePdfRedacted',
            'ConvertTo-OfficePdfSanitized',
            'ConvertTo-OfficePdfSearchable',
            'Copy-OfficePdfPage',
            'Export-OfficePdfImage',
            'Export-OfficePdfLayoutOverlay',
            'Export-OfficePdfXfdf',
            'Get-OfficePdf',
            'Get-OfficePdfAnnotation',
            'Get-OfficePdfAppendOnlyMutation',
            'Get-OfficePdfAttachment',
            'Get-OfficePdfCompliance',
            'Get-OfficePdfDiagnostic',
            'Get-OfficePdfFont',
            'Get-OfficePdfFormField',
            'Get-OfficePdfImage',
            'Get-OfficePdfInfo',
            'Get-OfficePdfInteractionMap',
            'Get-OfficePdfOptimization',
            'Get-OfficePdfPreflight',
            'Get-OfficePdfRedactionPlan',
            'Get-OfficePdfSignature',
            'Get-OfficePdfText',
            'Import-OfficePdfXfdf',
            'Invoke-OfficePdfOcrMerge',
            'Join-OfficePdf',
            'Move-OfficePdfPage',
            'New-OfficePdfSignature',
            'Remove-OfficePdfAnnotation',
            'Remove-OfficePdfPage',
            'Set-OfficePdfAnnotation',
            'Set-OfficePdfForm',
            'Set-OfficePdfMetadata',
            'Set-OfficePdfPage',
            'Set-OfficePdfSignature',
            'Split-OfficePdf'
        )) {
            $parameters = (Get-Command $name).Parameters.Keys
            $parameters | Should -Contain 'Password'
            $parameters | Should -Contain 'IgnorePermissionRestrictions'
        }

        $overlayParameters = (Get-Command Add-OfficePdfPageOverlay).Parameters.Keys
        $overlayParameters | Should -Contain 'ReadOptions'
        $overlayParameters | Should -Contain 'SourceReadOptions'
        (Get-Command ConvertTo-OfficePdfSearchable).Parameters.Keys | Should -Contain 'ReadOptions'

        foreach ($name in @('Compare-OfficePdfVisual', 'Test-OfficePdfRewrite')) {
            $parameters = (Get-Command $name).Parameters.Keys
            $parameters | Should -Contain 'ReferencePassword'
            $parameters | Should -Contain 'IgnoreReferencePermissionRestrictions'
            $parameters | Should -Contain 'DifferencePassword'
            $parameters | Should -Contain 'IgnoreDifferencePermissionRestrictions'
        }
    }

    It 'ignores usage restrictions only after valid password authentication' {
        $path = Join-Path $TestDrive 'restricted.pdf'
        New-OfficePdf -Path $path -Password 'open' -OwnerPassword 'owner' -Permission -3904 {
            PdfParagraph 'Authenticated restricted text'
        } | Out-Null

        { Get-OfficePdfText -Path $path -Password 'open' } | Should -Throw
        { Get-OfficePdfText -Path $path -Password 'wrong' -IgnorePermissionRestrictions } | Should -Throw
        Get-OfficePdfText -Path $path -Password 'open' -IgnorePermissionRestrictions |
            Should -Match 'Authenticated restricted text'
        Get-OfficePdfText -Path $path -Password 'owner' |
            Should -Match 'Authenticated restricted text'
    }

    It 'merges independently authenticated restricted sources and reports what happened' {
        $first = Join-Path $TestDrive 'restricted-one.pdf'
        $second = Join-Path $TestDrive 'restricted-two.pdf'
        $output = Join-Path $TestDrive 'restricted-merged.pdf'
        New-OfficePdf -Path $first -Password 'open-one' -OwnerPassword 'owner-one' -Permission -3904 {
            PdfParagraph 'Restricted source one'
        } | Out-Null
        New-OfficePdf -Path $second -Password 'open-two' -OwnerPassword 'owner-two' -Permission -3904 {
            PdfParagraph 'Restricted source two'
        } | Out-Null

        { Join-OfficePdf -Path $first, $second -Password 'open-one', 'open-two' -OutputPath $output } |
            Should -Throw

        $report = Join-OfficePdf -Path $first, $second -Password 'open-one', 'open-two' `
            -IgnorePermissionRestrictions -OutputPath $output -PassThruReport

        $report.GetType().FullName | Should -Be 'OfficeIMO.Pdf.PdfMergeReport'
        $report.Sources.Count | Should -Be 2
        $report.Sources[0].PermissionRestrictionsIgnored | Should -BeTrue
        $report.Sources[1].PermissionRestrictionsIgnored | Should -BeTrue
        $report.Sources[0].PasswordAuthenticationRole.ToString() | Should -Be 'User'
        $report.OutputPageCount | Should -Be 2
        $report.OutputHasEncryption | Should -BeTrue
        { Get-OfficePdfText -Path $output } | Should -Throw '*requires a password*'
        { Get-OfficePdfText -Path $output -Password 'open-one' } | Should -Throw
        { Get-OfficePdfText -Path $output -Password 'owner-two' } | Should -Throw
        Get-OfficePdfText -Path $output -Password 'owner-one' | Should -Match 'Restricted source one'
        Get-OfficePdfText -Path $output -Password 'open-one' -IgnorePermissionRestrictions |
            Should -Match 'Restricted source two'
    }

    It 'executes page mutations through the authenticated document contract' {
        $source = Join-Path $TestDrive 'restricted-pages.pdf'
        $moved = Join-Path $TestDrive 'restricted-pages-moved.pdf'
        $removed = Join-Path $TestDrive 'restricted-pages-removed.pdf'
        $rotated = Join-Path $TestDrive 'restricted-pages-rotated.pdf'
        $boxed = Join-Path $TestDrive 'restricted-pages-boxed.pdf'
        New-OfficePdf -Path $source -Password 'open' -OwnerPassword 'owner' -Permission -3904 {
            PdfParagraph 'Restricted page one'
            PdfPageBreak
            PdfParagraph 'Restricted page two'
            PdfPageBreak
            PdfParagraph 'Restricted page three'
        } | Out-Null

        Move-OfficePdfPage -Path $source -PageRange 1 -BeforePage 4 -OutputPath $moved `
            -Password 'open' -IgnorePermissionRestrictions -PassThru | Should -BeOfType System.IO.FileInfo
        Remove-OfficePdfPage -Path $source -PageRange 2 -OutputPath $removed `
            -Password 'open' -IgnorePermissionRestrictions -PassThru | Should -BeOfType System.IO.FileInfo
        Set-OfficePdfPage -Path $source -PageRange 1 -Rotation 90 -OutputPath $rotated `
            -Password 'open' -IgnorePermissionRestrictions -PassThru | Should -BeOfType System.IO.FileInfo
        Set-OfficePdfPage -Path $source -PageRange 1 -BoxName CropBox `
            -Left 10 -Bottom 10 -Right 300 -Top 500 -OutputPath $boxed `
            -Password 'open' -IgnorePermissionRestrictions -PassThru | Should -BeOfType System.IO.FileInfo

        foreach ($output in @($moved, $removed, $rotated, $boxed)) {
            { Get-OfficePdfText -Path $output } | Should -Throw '*requires a password*'
            (Get-OfficePdfInfo -Path $output -Password 'owner').Security.HasEncryption | Should -BeTrue
        }
        $movedPages = @(Get-OfficePdfText -Path $moved -ByPage -Password 'owner')
        $movedPages[2].Text | Should -Match 'Restricted page one'
        (Get-OfficePdfInfo -Path $removed -Password 'owner').PageCount | Should -Be 2
        (Get-OfficePdfInfo -Path $rotated -Password 'owner').PageCount | Should -Be 3
        (Get-OfficePdfInfo -Path $boxed -Password 'owner').Pages[0].Geometry.CropBox.Width | Should -Be 290
    }

    It 'fills authenticated restricted forms through the shared read contract' {
        $source = Join-Path $TestDrive 'restricted-form.pdf'
        $output = Join-Path $TestDrive 'restricted-form-filled.pdf'
        $flat = Join-Path $TestDrive 'restricted-form-flat.pdf'
        $filledAndFlat = Join-Path $TestDrive 'restricted-form-filled-flat.pdf'
        $incremental = Join-Path $TestDrive 'restricted-form-incremental.pdf'
        $imported = Join-Path $TestDrive 'restricted-form-imported.pdf'
        New-OfficePdf -Path $source -Password 'open' -OwnerPassword 'owner' -Permission -3904 {
            PdfParagraph 'Restricted form'
            PdfFormField -Name Name -Type Text -Value Ada
        } | Out-Null

        Set-OfficePdfForm -Path $source -OutputPath $output -Field @{ Name = 'Grace' } `
            -Password 'open' -IgnorePermissionRestrictions -PassThru | Should -BeOfType System.IO.FileInfo
        ConvertTo-OfficePdfFlatForm -Path $source -OutputPath $flat `
            -Password 'open' -IgnorePermissionRestrictions | Should -BeOfType System.IO.FileInfo
        Set-OfficePdfForm -Path $source -OutputPath $filledAndFlat -Field @{ Name = 'Flattened' } -Flatten `
            -Password 'open' -IgnorePermissionRestrictions -PassThru | Should -BeOfType System.IO.FileInfo
        Set-OfficePdfForm -Path $source -OutputPath $incremental -Field @{ Name = 'Incremental' } -Incremental `
            -Password 'owner' -PassThru | Should -BeOfType System.IO.FileInfo
        $xfdf = Export-OfficePdfXfdf -Path $source -Password 'open' -IgnorePermissionRestrictions
        $xfdf = $xfdf -replace '>Ada<', '>Imported<'
        Import-OfficePdfXfdf -Path $source -Xfdf $xfdf -OutputPath $imported `
            -Password 'open' -IgnorePermissionRestrictions

        foreach ($resultPath in @($output, $flat, $filledAndFlat, $incremental, $imported)) {
            { Get-OfficePdfInfo -Path $resultPath } | Should -Throw '*requires a password*'
            (Get-OfficePdfInfo -Path $resultPath -Password 'owner').Security.HasEncryption | Should -BeTrue
        }
        (Get-OfficePdfFormField -Path $output -Name Name -Password 'owner').Value | Should -Be 'Grace'
        (Get-OfficePdfInfo -Path $flat -Password 'owner').FormFieldCount | Should -Be 0
        (Get-OfficePdfInfo -Path $filledAndFlat -Password 'owner').FormFieldCount | Should -Be 0
        (Get-OfficePdfFormField -Path $incremental -Name Name -Password 'owner').Value | Should -Be 'Incremental'
        (Get-OfficePdfFormField -Path $imported -Name Name -Password 'owner').Value | Should -Be 'Imported'
    }

    It 'sanitizes and updates metadata on authenticated restricted PDFs' {
        $source = Join-Path $TestDrive 'restricted-rewrite.pdf'
        $metadata = Join-Path $TestDrive 'restricted-metadata.pdf'
        $sanitized = Join-Path $TestDrive 'restricted-sanitized.pdf'
        New-OfficePdf -Path $source -Password 'open' -OwnerPassword 'owner' -Permission -3904 {
            PdfParagraph 'Restricted rewrite source'
        } | Out-Null

        Set-OfficePdfMetadata -Path $source -OutputPath $metadata -Title 'Authenticated metadata' `
            -Password 'open' -IgnorePermissionRestrictions -PassThru | Should -BeOfType System.IO.FileInfo
        $sanitization = ConvertTo-OfficePdfSanitized -Path $source -OutputPath $sanitized `
            -Password 'open' -IgnorePermissionRestrictions

        foreach ($output in @($metadata, $sanitized)) {
            { Get-OfficePdfInfo -Path $output } | Should -Throw '*requires a password*'
            (Get-OfficePdfInfo -Path $output -Password 'owner').Security.HasEncryption | Should -BeTrue
        }
        (Get-OfficePdfInfo -Path $metadata -Password 'owner').Metadata.Title | Should -Be 'Authenticated metadata'
        Get-OfficePdfText -Path $sanitized -Password 'owner' | Should -Match 'Restricted rewrite source'
        $sanitization.MutationPlan.Warnings | Should -Contain 'Input.PermissionRestrictionsIgnored'
    }

    It 'proves rewrites between independently authenticated document instances' {
        $reference = Join-Path $TestDrive 'restricted-proof-reference.pdf'
        $difference = Join-Path $TestDrive 'restricted-proof-difference.pdf'
        New-OfficePdf -Path $reference -Password 'reference-open' -OwnerPassword 'reference-owner' -Permission -3904 {
            PdfParagraph 'Restricted rewrite proof'
        } | Out-Null
        New-OfficePdf -Path $difference -Password 'difference-open' -OwnerPassword 'difference-owner' -Permission -3904 {
            PdfParagraph 'Restricted rewrite proof'
        } | Out-Null

        $readOptionsType = (Get-Command Add-OfficePdfPageOverlay).Parameters['ReadOptions'].ParameterType
        $referenceReadOptions = [Activator]::CreateInstance($readOptionsType)
        $differenceReadOptions = [Activator]::CreateInstance($readOptionsType)
        $referenceReadOptions.Limits.MaxInputBytes = 10MB
        $differenceReadOptions.Limits.MaxInputBytes = 10MB
        $rewriteOptionsType = (Get-Command Test-OfficePdfRewrite).Parameters['Options'].ParameterType
        $rewriteOptions = [Activator]::CreateInstance($rewriteOptionsType)
        $rewriteOptions.OriginalReadOptions = $referenceReadOptions
        $rewriteOptions.RewrittenReadOptions = $differenceReadOptions

        $report = Test-OfficePdfRewrite -ReferencePath $reference -DifferencePath $difference `
            -Options $rewriteOptions `
            -ReferencePassword 'reference-open' -IgnoreReferencePermissionRestrictions `
            -DifferencePassword 'difference-open' -IgnoreDifferencePermissionRestrictions

        $report.IsPreserved | Should -BeTrue
        [object]::ReferenceEquals($rewriteOptions.OriginalReadOptions, $referenceReadOptions) | Should -BeTrue
        [object]::ReferenceEquals($rewriteOptions.RewrittenReadOptions, $differenceReadOptions) | Should -BeTrue
    }
}

Describe 'General existing-page visual stamping' {
    It 'draws arbitrary page-aware canvas content on selected existing pages' {
        $source = Join-Path $TestDrive 'canvas-source.pdf'
        $output = Join-Path $TestDrive 'canvas-output.pdf'
        New-OfficePdf -Path $source {
            PdfParagraph 'Canvas page one'
            PdfPageBreak
            PdfParagraph 'Canvas page two'
        } | Out-Null

        (Get-Command Add-OfficePdfCanvas).Parameters.Keys | Should -Contain 'ConfigureRendering'
        $configured = [System.Collections.Generic.List[bool]]::new()
        Add-OfficePdfCanvas -Path $source -OutputPath $output -PageRange 1 -ConfigureRendering {
            param($options)
            $configured.Add($true)
            $options.CompressContentStreams = $false
        } -Content {
            param($canvas, $page)
            $null = $canvas.Text("Canvas overlay $($page.PageNumber)/$($page.PageCount)", 36, 36, $page.Width - 72, 24, 11)
        } -PassThru | Should -BeOfType System.IO.FileInfo

        $configured | Should -HaveCount 1
        $pages = @(Get-OfficePdfText -Path $output -ByPage)
        $pages[0].Text | Should -Match 'Canvas overlay 1/2'
        $pages[1].Text | Should -Not -Match 'Canvas overlay'
    }

    It 'accepts PowerShell rich text runs in the canvas DSL without native array conversion' {
        $source = Join-Path $TestDrive 'canvas-dsl-source.pdf'
        $output = Join-Path $TestDrive 'canvas-dsl-output.pdf'
        New-OfficePdf -Path $source {
            PdfParagraph 'Original page content'
        } | Out-Null

        Add-OfficePdfCanvas -Path $source -OutputPath $output -Content {
            PdfCanvasText -Run @(
                TextRun 'Owner: ' -Bold
                @{ Text = 'Platform'; Color = '#0F766E' }
            ) -X 36 -Y 24 -FontSize 10

            PdfCanvasText 'Review copy' -X 36 -Y 60 -Italic
        } -PassThru | Should -BeOfType System.IO.FileInfo

        $text = Get-OfficePdfText -Path $output
        $text | Should -Match 'Owner: Platform'
        $text | Should -Match 'Review copy'
        $text | Should -Match 'Original page content'
    }

    It 'reports the fixed-position link limitation in canvas terms' {
        $source = Join-Path $TestDrive 'canvas-link-source.pdf'
        $output = Join-Path $TestDrive 'canvas-link-output.pdf'
        New-OfficePdf -Path $source {
            PdfParagraph 'Original page content'
        } | Out-Null

        {
            Add-OfficePdfCanvas -Path $source -OutputPath $output -Content {
                PdfCanvasText -Run @(
                    TextRun 'Open portal' -LinkUri 'https://example.com'
                ) -X 36 -Y 24
            } -ErrorAction Stop
        } | Should -Throw '*PDF canvas text runs do not support links*'
    }

    It 'imports a selected source PDF page as an overlay and an underlay' {
        $target = Join-Path $TestDrive 'overlay-target.pdf'
        $source = Join-Path $TestDrive 'overlay-source.pdf'
        $overlay = Join-Path $TestDrive 'overlay-output.pdf'
        $underlay = Join-Path $TestDrive 'underlay-output.pdf'
        New-OfficePdf -Path $target {
            PdfParagraph 'Target page one'
            PdfPageBreak
            PdfParagraph 'Target page two'
        } | Out-Null
        New-OfficePdf -Path $source {
            PdfParagraph 'Source page one'
            PdfPageBreak
            PdfParagraph 'Source page two'
        } | Out-Null

        Add-OfficePdfPageOverlay -Path $target -SourcePath $source -SourcePageNumber 2 `
            -PageRange 2 -OutputPath $overlay -PassThru | Should -BeOfType System.IO.FileInfo
        Add-OfficePdfPageOverlay -Path $target -SourcePath $source -SourcePageNumber 1 `
            -PageRange 1 -Underlay -OutputPath $underlay -PassThru | Should -BeOfType System.IO.FileInfo

        $overlayPages = @(Get-OfficePdfText -Path $overlay -ByPage)
        $overlayPages[0].Text | Should -Not -Match 'Source page two'
        $overlayPages[1].Text | Should -Match 'Source.*page.*two'
        (Get-OfficePdfText -Path $underlay) | Should -Match 'Source.*page.*one'
    }

    It 'rejects an oversized overlay source before the shared stream overload reads it' {
        $target = Join-Path $TestDrive 'bounded-overlay-target.pdf'
        $source = Join-Path $TestDrive 'bounded-overlay-source.pdf'
        $output = Join-Path $TestDrive 'bounded-overlay-output.pdf'
        New-OfficePdf -Path $target { PdfParagraph 'Target' } | Out-Null
        New-OfficePdf -Path $source { PdfParagraph 'Source' } | Out-Null
        $readOptionsType = (Get-Command Add-OfficePdfPageOverlay).Parameters['SourceReadOptions'].ParameterType
        $sourceReadOptions = [Activator]::CreateInstance($readOptionsType)
        $sourceReadOptions.Limits.MaxInputBytes = 1

        { Add-OfficePdfPageOverlay -Path $target -SourcePath $source -OutputPath $output `
                -SourceReadOptions $sourceReadOptions -ErrorAction Stop } |
            Should -Throw '*configured limit*'
        Test-Path -LiteralPath $output | Should -BeFalse
    }
}

Describe 'PDF table value and typed-cell contracts' {
    It 'normalizes collection properties with the requested separator' {
        $path = Join-Path $TestDrive 'collection-table.pdf'
        $rows = @(
            [pscustomobject]@{ Name = 'Directory'; Tags = @('Identity', 'Critical') }
        )

        New-OfficePdf -Path $path {
            PdfTable -InputObject $rows -Property Name, Tags -CollectionSeparator ' | '
        } | Out-Null

        $text = Get-OfficePdfText -Path $path
        $text | Should -Match 'Identity \| Critical'
        $text | Should -Not -Match 'System\.Object\[\]'
    }

    It 'renders typed check boxes, form fields, images, and links in table cells' {
        $path = Join-Path $TestDrive 'typed-table-cells.pdf'
        $checkBox = New-OfficePdfTableCellCheckBox -Name Approved -Checked
        $field = New-OfficePdfTableCellField -Name Reviewer -Value 'Alice'
        $image = New-OfficePdfTableCellImage -Path (Join-Path $PSScriptRoot 'Assets\CellImage.png') -Width 18 -Height 18

        New-OfficePdf -Path $path {
            PdfTable -InputObject @(
                @(
                    (New-OfficePdfTableCell -Text 'Approved' -CheckBox $checkBox),
                    (New-OfficePdfTableCell -Text 'Reviewer' -FormField $field)
                ),
                @(
                    (New-OfficePdfTableCell -Text 'Logo' -Image $image),
                    (New-OfficePdfTableCell -Text 'Portal' -LinkUri 'https://example.com' -LinkContents 'Open portal')
                )
            )
        } | Out-Null

        $fields = @(Get-OfficePdfFormField -Path $path)
        $fields.Name | Should -Contain 'Approved'
        $fields.Name | Should -Contain 'Reviewer'
        @(Get-OfficePdfImage -Path $path).Count | Should -BeGreaterThan 0
        (Get-OfficePdfInfo -Path $path).LinkUris | Should -Contain 'https://example.com'
    }
}

Describe 'Composed generated-PDF headers and footers' {
    It 'captures caller-local values for deferred header and footer composers' {
        $path = Join-Path $TestDrive 'header-footer-closures.pdf'
        New-OfficePdf -Path $path {
            $headerLabel = 'Scoped header'
            $footerLabel = 'Scoped footer'
            PdfHeader -Compose {
                param($header)
                $null = $header.Text($headerLabel)
            }
            PdfFooter -Compose {
                param($footer)
                $null = $footer.Text($footerLabel)
            }
            PdfParagraph 'Scoped body'
        } | Out-Null

        $text = Get-OfficePdfText -Path $path
        $text | Should -Match 'Scoped header'
        $text | Should -Match 'Scoped footer'
    }

    It 'supports default, first-page, and even-page variants through the native composers' {
        $path = Join-Path $TestDrive 'header-footer-variants.pdf'
        New-OfficePdf -Path $path {
            PdfHeader -Compose {
                param($header)
                $null = $header.Zones('Default header', 'Default center', 'Page {page}/{pages}')
                $null = $header.FirstPageZones('First header', 'First center', 'First {page}')
                $null = $header.EvenPagesZones('Even header', 'Even center', 'Even {page}/{pages}')
            }
            PdfFooter -Compose {
                param($footer)
                $null = $footer.Zones('Default footer', 'Default center', '{page}/{pages}')
                $null = $footer.FirstPageZones('First footer', 'First center', 'First {page}')
                $null = $footer.EvenPagesZones('Even footer', 'Even center', 'Even {page}/{pages}')
            }
            PdfParagraph 'Body one'
            PdfPageBreak
            PdfParagraph 'Body two'
            PdfPageBreak
            PdfParagraph 'Body three'
        } | Out-Null

        $pages = @(Get-OfficePdfText -Path $path -ByPage)
        $pages[0].Text | Should -Match 'First header'
        $pages[0].Text | Should -Match 'First footer'
        $pages[1].Text | Should -Match 'Even header'
        $pages[1].Text | Should -Match 'Even footer'
        $pages[2].Text | Should -Match 'Default header'
        $pages[2].Text | Should -Match 'Default footer'
    }

    It 'preserves styled header and footer runs with styled page tokens' {
        $path = Join-Path $TestDrive 'header-footer-rich-text.pdf'
        New-OfficePdf -Path $path {
            PdfHeader -Compose {
                param($header)
                $label = New-OfficeTextRun -Text 'Rich header ' -Bold | ConvertTo-OfficePdfTextRun
                $pageStyle = New-OfficeTextRun -Italic | ConvertTo-OfficePdfTextRun
                $null = $header.Text({
                    param($text)
                    $null = $text.Run($label).CurrentPage($pageStyle)
                })
            }
            PdfFooter -Compose {
                param($footer)
                $label = New-OfficeTextRun -Text 'Rich footer ' -Underline | ConvertTo-OfficePdfTextRun
                $pageStyle = New-OfficeTextRun -Kind Superscript | ConvertTo-OfficePdfTextRun
                $null = $footer.Text({
                    param($text)
                    $null = $text.Run($label).CurrentPage($pageStyle).Text('/').TotalPages($pageStyle)
                })
            }
            PdfParagraph 'Styled body'
        } | Out-Null

        $text = Get-OfficePdfText -Path $path
        $text | Should -Match 'Rich header\s*1'
        $text | Should -Match 'Rich footer'
        $text | Should -Match '1\s+1'
        $text | Should -Match '/'
    }
}
