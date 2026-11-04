BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
}

Describe 'Authenticated PDF automation' {
    It 'imports XFDF text lines and one piped PDF but rejects multiple source PDFs before writing' {
        $source = Join-Path $TestDrive 'xfdf-source.pdf'
        $second = Join-Path $TestDrive 'xfdf-second.pdf'
        $output = Join-Path $TestDrive 'xfdf-imported.pdf'
        $xfdfPath = Join-Path $TestDrive 'fields.xfdf'
        New-OfficePdf -Path $source {
            PdfParagraph 'Pipeline form'
            PdfFormField -Name Name -Type Text -Value Ada
        } | Out-Null
        Copy-Item -LiteralPath $source -Destination $second
        $xfdf = (Export-OfficePdfXfdf -Path $source) -replace '>Ada<', '>Grace<'
        [IO.File]::WriteAllText($xfdfPath, $xfdf)
        $tooSmall = New-OfficePdfFormFillOptions -MaxFields 1 -MaxValueCharacters 4
        { Import-OfficePdfXfdf -Path $source -XfdfPath $xfdfPath -OutputPath $output -Options $tooSmall -ErrorAction Stop } |
            Should -Throw '*value character limit*'
        Test-Path -LiteralPath $output | Should -BeFalse
        $fillOptions = New-OfficePdfFormFillOptions -MaxFields 1 -MaxValueCharacters 5 -KeepNeedAppearances
        Get-Item -LiteralPath $source | Import-OfficePdfXfdf -XfdfPath $xfdfPath -OutputPath $output -Options $fillOptions
        $filled = [OfficeIMO.Pdf.PdfDocument]::Load($output)
        try { $filled.Inspect().AcroFormNeedAppearances | Should -BeTrue } finally { $filled.Dispose() }
        (Get-OfficePdfFormField -Path $output | Where-Object Name -eq Name).Value | Should -Be 'Grace'
        $original = [Convert]::ToBase64String([IO.File]::ReadAllBytes($output))
        { Get-Item -LiteralPath $source, $second | Import-OfficePdfXfdf -XfdfPath $xfdfPath -OutputPath $output -ErrorAction Stop } |
            Should -Throw '*exactly one input*'
        [Convert]::ToBase64String([IO.File]::ReadAllBytes($output)) | Should -Be $original
        $xfdf -split '\r?\n' | Import-OfficePdfXfdf -Path $source -OutputPath $output
        (Get-OfficePdfFormField -Path $output | Where-Object Name -eq Name).Value | Should -Be 'Grace'
    }

    It 'uses factory interaction limits when recovering native selection regions' {
        $source = Join-Path $TestDrive 'interaction-limits.pdf'
        New-OfficePdf -Path $source { PdfParagraph 'More than one selectable glyph' } | Out-Null
        $small = New-OfficePdfInteractionOptions -MaxTextRegions 1
        { Get-OfficePdfInteractionMap -Path $source -Options $small -ErrorAction Stop } | Should -Throw '*limit*'
        $options = New-OfficePdfInteractionOptions -MaxTextRegions 1000 -MaxImageRegions 10
        $map = Get-OfficePdfInteractionMap -Path $source -Options $options
        $map.PageNumber | Should -Be 1
        $map.TextRegions.Count | Should -BeGreaterThan 1
    }

    It 'ignores usage restrictions only after valid password authentication' {
        $path = Join-Path $TestDrive 'restricted.pdf'
        New-OfficePdf -Path $path -Password 'open' -OwnerPassword 'owner' -Permission None {
            PdfParagraph 'Authenticated restricted text'
        } | Out-Null

        { Read-OfficeDocument -As Text -Path $path -Password 'open' -ErrorAction Stop } | Should -Throw
        { Read-OfficeDocument -As Text -Path $path -Password 'wrong' -Options (New-OfficePdfReadOptions -IgnorePermissionRestrictions) -ErrorAction Stop } | Should -Throw
        Read-OfficeDocument -As Text -Path $path -Password 'open' -Options (New-OfficePdfReadOptions -IgnorePermissionRestrictions) |
            Should -Match 'Authenticated restricted text'
        Read-OfficeDocument -As Text -Path $path -Password 'owner' |
            Should -Match 'Authenticated restricted text'

        { ConvertTo-OfficeDocumentImage -LiteralPath $path -Format Svg -Password open -ErrorAction Stop } |
            Should -Throw
        { ConvertTo-OfficeDocumentImage -LiteralPath $path -Format Svg -Password wrong -Options @{ PermissionPolicy = 'IgnoreRestrictions' } -ErrorAction Stop } |
            Should -Throw
        $image = ConvertTo-OfficeDocumentImage -LiteralPath $path -Format Svg -Password open -Options @{ PermissionPolicy = 'IgnoreRestrictions' } -ErrorAction Stop
        $image.GetType().FullName | Should -Be 'OfficeIMO.Drawing.OfficeImageExportResult'
        $image.Bytes.Length | Should -BeGreaterThan 0
    }

    It 'merges independently authenticated restricted sources and reports what happened' {
        $first = Join-Path $TestDrive 'restricted-one.pdf'
        $second = Join-Path $TestDrive 'restricted-two.pdf'
        $output = Join-Path $TestDrive 'restricted-merged.pdf'
        New-OfficePdf -Path $first -Password 'open-one' -OwnerPassword 'owner-one' -Permission None {
            PdfParagraph 'Restricted source one'
        } | Out-Null
        New-OfficePdf -Path $second -Password 'open-two' -OwnerPassword 'owner-two' -Permission None {
            PdfParagraph 'Restricted source two'
        } | Out-Null

        { Merge-OfficeDocument -Path $first, $second -Password 'open-one', 'open-two' -OutputPath $output } |
            Should -Throw

        $report = Merge-OfficeDocument -Path $first, $second -Password 'open-one', 'open-two' `
            -Options (New-OfficePdfReadOptions -IgnorePermissionRestrictions) -OutputPath $output -PassThruReport

        $report.GetType().FullName | Should -Be 'OfficeIMO.Pdf.PdfMergeReport'
        $report.Sources.Count | Should -Be 2
        $report.Sources[0].PermissionRestrictionsIgnored | Should -BeTrue
        $report.Sources[1].PermissionRestrictionsIgnored | Should -BeTrue
        $report.Sources[0].PasswordAuthenticationRole.ToString() | Should -Be 'User'
        $report.OutputPageCount | Should -Be 2
        $report.OutputHasEncryption | Should -BeTrue
        { Read-OfficeDocument -As Text -Path $output -ErrorAction Stop } | Should -Throw '*requires a password*'
        { Read-OfficeDocument -As Text -Path $output -Password 'open-one' -ErrorAction Stop } | Should -Throw
        { Read-OfficeDocument -As Text -Path $output -Password 'owner-two' -ErrorAction Stop } | Should -Throw
        Read-OfficeDocument -As Text -Path $output -Password 'owner-one' | Should -Match 'Restricted source one'
        Read-OfficeDocument -As Text -Path $output -Password 'open-one' -Options (New-OfficePdfReadOptions -IgnorePermissionRestrictions) |
            Should -Match 'Restricted source two'
    }

    It 'executes page mutations through the authenticated document contract' {
        $source = Join-Path $TestDrive 'restricted-pages.pdf'
        $moved = Join-Path $TestDrive 'restricted-pages-moved.pdf'
        $removed = Join-Path $TestDrive 'restricted-pages-removed.pdf'
        $rotated = Join-Path $TestDrive 'restricted-pages-rotated.pdf'
        $boxed = Join-Path $TestDrive 'restricted-pages-boxed.pdf'
        New-OfficePdf -Path $source -Password 'open' -OwnerPassword 'owner' -Permission None {
            PdfParagraph 'Restricted page one'
            PdfPageBreak
            PdfParagraph 'Restricted page two'
            PdfPageBreak
            PdfParagraph 'Restricted page three'
        } | Out-Null

        Move-OfficePdfPage -Path $source -Page 1 -BeforePage 4 -OutputPath $moved `
            -Password 'open' -IgnorePermissionRestrictions -PassThru | Should -BeOfType System.IO.FileInfo
        Remove-OfficePdfPage -Path $source -Page 2 -OutputPath $removed `
            -Password 'open' -IgnorePermissionRestrictions -PassThru | Should -BeOfType System.IO.FileInfo
        Set-OfficePdfPage -Path $source -Page 1 -Rotation 90 -OutputPath $rotated `
            -Password 'open' -IgnorePermissionRestrictions -PassThru | Should -BeOfType System.IO.FileInfo
        Set-OfficePdfPage -Path $source -Page 1 -BoxName CropBox `
            -BoxLeft 10 -BoxBottom 10 -BoxRight 300 -BoxTop 500 -OutputPath $boxed `
            -Password 'open' -IgnorePermissionRestrictions -PassThru | Should -BeOfType System.IO.FileInfo

        foreach ($output in @($moved, $removed, $rotated, $boxed)) {
            { Read-OfficeDocument -As Text -Path $output -ErrorAction Stop } | Should -Throw '*requires a password*'
            ((Get-OfficeDocumentInfo -Detailed -Path $output -Password 'owner').Details).Security.HasEncryption | Should -BeTrue
        }
        $movedPages = @(Read-OfficeDocument -As Text -Path $moved -ByPage -Password 'owner')
        $movedPages[2].Text | Should -Match 'Restricted page one'
        ((Get-OfficeDocumentInfo -Detailed -Path $removed -Password 'owner').Details).PageCount | Should -Be 2
        ((Get-OfficeDocumentInfo -Detailed -Path $rotated -Password 'owner').Details).PageCount | Should -Be 3
        ((Get-OfficeDocumentInfo -Detailed -Path $boxed -Password 'owner').Details).Pages[0].Geometry.CropBox.Width | Should -Be 290
    }

    It 'fills authenticated restricted forms through the shared read contract' {
        $source = Join-Path $TestDrive 'restricted-form.pdf'
        $output = Join-Path $TestDrive 'restricted-form-filled.pdf'
        $flat = Join-Path $TestDrive 'restricted-form-flat.pdf'
        $filledAndFlat = Join-Path $TestDrive 'restricted-form-filled-flat.pdf'
        $incremental = Join-Path $TestDrive 'restricted-form-incremental.pdf'
        $imported = Join-Path $TestDrive 'restricted-form-imported.pdf'
        New-OfficePdf -Path $source -Password 'open' -OwnerPassword 'owner' -Permission None {
            PdfParagraph 'Restricted form'
            PdfFormField -Name Name -Type Text -Value Ada
        } | Out-Null

        Set-OfficePdfFormField -Path $source -OutputPath $output -Value @{ Name = 'Grace' } `
            -Password 'open' -IgnorePermissionRestrictions -PassThru | Should -BeOfType System.IO.FileInfo
        Update-OfficePdf -FlattenForm -Path $source -OutputPath $flat `
            -Password 'open' -IgnorePermissionRestrictions | Should -BeOfType System.IO.FileInfo
        Set-OfficePdfFormField -Path $source -OutputPath $filledAndFlat -Value @{ Name = 'Flattened' } -Flatten `
            -Password 'open' -IgnorePermissionRestrictions -PassThru | Should -BeOfType System.IO.FileInfo
        Set-OfficePdfFormField -Path $source -OutputPath $incremental -Value @{ Name = 'Incremental' } -Incremental `
            -Password 'owner' -PassThru | Should -BeOfType System.IO.FileInfo
        $xfdf = Export-OfficePdfXfdf -Path $source -Password 'open' -IgnorePermissionRestrictions
        $xfdf = $xfdf -replace '>Ada<', '>Imported<'
        Import-OfficePdfXfdf -Path $source -Xfdf $xfdf -OutputPath $imported `
            -Password 'open' -IgnorePermissionRestrictions

        foreach ($resultPath in @($output, $flat, $filledAndFlat, $incremental, $imported)) {
            (Get-OfficeDocumentInfo -Detailed -Path $resultPath).PdfPreflight.CanRead | Should -BeFalse
            (Get-OfficeDocumentInfo -Detailed -Path $resultPath).PdfPreflight.Probe.Security.HasEncryption | Should -BeTrue
            ((Get-OfficeDocumentInfo -Detailed -Path $resultPath -Password 'owner').Details).Security.HasEncryption | Should -BeTrue
        }
        (Get-OfficePdfFormField -Path $output -Name Name -Password 'owner').Value | Should -Be 'Grace'
        ((Get-OfficeDocumentInfo -Detailed -Path $flat -Password 'owner').Details).FormFieldCount | Should -Be 0
        ((Get-OfficeDocumentInfo -Detailed -Path $filledAndFlat -Password 'owner').Details).FormFieldCount | Should -Be 0
        (Get-OfficePdfFormField -Path $incremental -Name Name -Password 'owner').Value | Should -Be 'Incremental'
        (Get-OfficePdfFormField -Path $imported -Name Name -Password 'owner').Value | Should -Be 'Imported'
    }

    It 'sanitizes and updates metadata on authenticated restricted PDFs' {
        $source = Join-Path $TestDrive 'restricted-rewrite.pdf'
        $metadata = Join-Path $TestDrive 'restricted-metadata.pdf'
        $sanitized = Join-Path $TestDrive 'restricted-sanitized.pdf'
        New-OfficePdf -Path $source -Password 'open' -OwnerPassword 'owner' -Permission None {
            PdfParagraph 'Restricted rewrite source'
        } | Out-Null

        Set-OfficePdfMetadata -Path $source -OutputPath $metadata -Title 'Authenticated metadata' `
            -Password 'open' -IgnorePermissionRestrictions -PassThru | Should -BeOfType System.IO.FileInfo
        $sanitization = Optimize-OfficePdf -Sanitize -PassThru -Path $source -OutputPath $sanitized `
            -Password 'open' -IgnorePermissionRestrictions

        foreach ($output in @($metadata, $sanitized)) {
            (Get-OfficeDocumentInfo -Detailed -Path $output).PdfPreflight.CanRead | Should -BeFalse
            (Get-OfficeDocumentInfo -Detailed -Path $output).PdfPreflight.Probe.Security.HasEncryption | Should -BeTrue
            ((Get-OfficeDocumentInfo -Detailed -Path $output -Password 'owner').Details).Security.HasEncryption | Should -BeTrue
        }
        ((Get-OfficeDocumentInfo -Detailed -Path $metadata -Password 'owner').Details).Metadata.Title | Should -Be 'Authenticated metadata'
        Read-OfficeDocument -As Text -Path $sanitized -Password 'owner' | Should -Match 'Restricted rewrite source'
        $sanitization.MutationPlan.Warnings | Should -Contain 'Input.PermissionRestrictionsIgnored'
    }

    It 'proves rewrites between independently authenticated document instances' {
        $reference = Join-Path $TestDrive 'restricted-proof-reference.pdf'
        $difference = Join-Path $TestDrive 'restricted-proof-difference.pdf'
        New-OfficePdf -Path $reference -Password 'reference-open' -OwnerPassword 'reference-owner' -Permission None {
            PdfParagraph 'Restricted rewrite proof'
        } | Out-Null
        New-OfficePdf -Path $difference -Password 'difference-open' -OwnerPassword 'difference-owner' -Permission None {
            PdfParagraph 'Restricted rewrite proof'
        } | Out-Null

        $readOptionsType = (Get-Command Add-OfficePdfStamp).Parameters['LoadOptions'].ParameterType
        $referenceReadOptions = [Activator]::CreateInstance($readOptionsType)
        $differenceReadOptions = [Activator]::CreateInstance($readOptionsType)
        $referenceReadOptions.Limits.MaxInputBytes = 10MB
        $differenceReadOptions.Limits.MaxInputBytes = 10MB
        $rewriteOptions = [OfficeIMO.Pdf.PdfRewritePreservationOptions]::new()
        $rewriteOptions.OriginalReadOptions = $referenceReadOptions
        $rewriteOptions.RewrittenReadOptions = $differenceReadOptions

        $report = Test-OfficePdfRewrite -ReferencePath $reference -DifferencePath $difference `
            -Options $rewriteOptions `
            -Password 'reference-open' -IgnorePermissionRestrictions `
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

        $configured = [System.Collections.Generic.List[bool]]::new()
        Add-OfficePdfStamp -Path $source -OutputPath $output -Page 1 -ConfigureRendering {
            param($options)
            $configured.Add($true)
            $options.CompressContentStreams = $false
        } -Content {
            param($canvas, $page)
            $null = $canvas.Text("Canvas overlay $($page.PageNumber)/$($page.PageCount)", [OfficeIMO.OfficeLength]::Points(36), [OfficeIMO.OfficeLength]::Points(36), ($page.Width - [OfficeIMO.OfficeLength]::Points(72)), [OfficeIMO.OfficeLength]::Points(24), [OfficeIMO.OfficeLength]::Points(11))
        } -PassThru | Should -BeOfType System.IO.FileInfo

        $configured | Should -HaveCount 1
        $pages = @(Read-OfficeDocument -As Text -Path $output -ByPage)
        $pages[0].Text | Should -Match 'Canvas overlay 1/2'
        $pages[1].Text | Should -Not -Match 'Canvas overlay'
    }

    It 'accepts PowerShell rich text runs in the canvas DSL without native array conversion' {
        $source = Join-Path $TestDrive 'canvas-dsl-source.pdf'
        $output = Join-Path $TestDrive 'canvas-dsl-output.pdf'
        New-OfficePdf -Path $source {
            PdfParagraph 'Original page content'
        } | Out-Null

        Add-OfficePdfStamp -Path $source -OutputPath $output -Content {
            PdfCanvasText -Content {
                PdfText 'Owner: ' -Bold
                PdfText 'Platform' -Color '#0F766E'
            } -X 36 -Y 24 -Size 10

            PdfCanvasText 'Review copy' -X 36 -Y 60 -Italic
        } -PassThru | Should -BeOfType System.IO.FileInfo

        $text = Read-OfficeDocument -As Text -Path $output
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
            Add-OfficePdfStamp -Path $source -OutputPath $output -Content {
                PdfCanvasText -Content { PdfText 'Open portal' -Link 'https://example.com' } -X 36 -Y 24
            } -ErrorAction Stop
        } | Should -Throw '*Canvas text*not support links*'
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

        Add-OfficePdfStamp -Path $target -OverlayPath $source -OverlayPage 2 `
            -Page 2 -OutputPath $overlay -PassThru | Should -BeOfType System.IO.FileInfo
        Add-OfficePdfStamp -Path $target -OverlayPath $source -OverlayPage 1 `
            -Page 1 -Behind -OutputPath $underlay -PassThru | Should -BeOfType System.IO.FileInfo

        $overlayPages = @(Read-OfficeDocument -As Text -Path $overlay -ByPage)
        $overlayPages[0].Text | Should -Not -Match 'Source page two'
        $overlayPages[1].Text | Should -Match 'Source.*page.*two'
        (Read-OfficeDocument -As Text -Path $underlay) | Should -Match 'Source.*page.*one'
    }

    It 'rejects an oversized overlay source before the shared stream overload reads it' {
        $target = Join-Path $TestDrive 'bounded-overlay-target.pdf'
        $source = Join-Path $TestDrive 'bounded-overlay-source.pdf'
        $output = Join-Path $TestDrive 'bounded-overlay-output.pdf'
        New-OfficePdf -Path $target { PdfParagraph 'Target' } | Out-Null
        New-OfficePdf -Path $source { PdfParagraph 'Source' } | Out-Null
        $readOptionsType = (Get-Command Add-OfficePdfStamp).Parameters['OverlayLoadOptions'].ParameterType
        $sourceReadOptions = [Activator]::CreateInstance($readOptionsType)
        $sourceReadOptions = $sourceReadOptions.Clone(1, $null)

        { Add-OfficePdfStamp -Path $target -OverlayPath $source -OutputPath $output `
                -OverlayLoadOptions $sourceReadOptions -ErrorAction Stop } |
            Should -Throw '*InputBytes*maximum 1*'
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

        $text = Read-OfficeDocument -As Text -Path $path
        $text | Should -Match 'Identity \| Critical'
        $text | Should -Not -Match 'System\.Object\[\]'
    }

    It 'renders typed check boxes, form fields, images, and links in table cells' {
        $path = Join-Path $TestDrive 'typed-table-cells.pdf'

        New-OfficePdf -Path $path {
            PdfTable -InputObject @(
                @(
                    (New-OfficePdfTableCell { PdfText 'Approved'; PdfFormField -Name Approved -Type CheckBox -Checked }),
                    (New-OfficePdfTableCell { PdfText 'Reviewer'; PdfFormField -Name Reviewer -Value 'Alice'; PdfFormField -Name Status -Type Choice -Option 'Ready', 'Pending' -Value Ready -ListBox })
                ),
                @(
                    (New-OfficePdfTableCell { PdfText 'Logo'; PdfImage -Path (Join-Path $PSScriptRoot 'Assets/CellImage.png') -Width 18 -Height 18 }),
                    (New-OfficePdfTableCell -Text 'Portal' -Link 'https://example.com' -LinkContents 'Open portal')
                )
            )
        } | Out-Null

        $fields = @(Get-OfficePdfFormField -Path $path)
        $fields.Name | Should -Contain 'Approved'
        $fields.Name | Should -Contain 'Reviewer'
        ($fields | Where-Object Name -eq Status).IsCombo | Should -BeFalse
        ($fields | Where-Object Name -eq Status).OptionCount | Should -Be 2
        @(Get-OfficePdfImage -Path $path).Count | Should -BeGreaterThan 0
        ((Get-OfficeDocumentInfo -Detailed -Path $path).Details).LinkUris | Should -Contain 'https://example.com'
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

        $text = Read-OfficeDocument -As Text -Path $path
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

        $pages = @(Read-OfficeDocument -As Text -Path $path -ByPage)
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
            PdfHeader {
                PdfText 'Rich header ' -Bold
                PdfPageNumber -Italic
            }
            PdfFooter {
                PdfText 'Rich footer ' -Underline Single
                PdfPageNumber -IncludeTotal -Baseline Superscript
            }
            PdfParagraph 'Styled body'
        } | Out-Null

        $text = Read-OfficeDocument -As Text -Path $path
        $text | Should -Match 'Rich header\s*1'
        $text | Should -Match 'Rich footer'
        $text | Should -Match '1\s+1'
        $text | Should -Match '/'
    }
}
