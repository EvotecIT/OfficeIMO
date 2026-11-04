BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    $oneNoteSectionType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.OneNote' -TypeName 'OfficeIMO.OneNote.OneNoteSection'
    $oneNotePageType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.OneNote' -TypeName 'OfficeIMO.OneNote.OneNotePage'
    $oneNoteMarkdownOptionsType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.OneNote.Markdown' -TypeName 'OfficeIMO.OneNote.Markdown.OneNoteMarkdownExportOptions'
    $oneNotePdfOptionsType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.OneNote.Pdf' -TypeName 'OfficeIMO.OneNote.Pdf.OneNotePdfExportOptions'
}

Describe 'OfficeIMO 3.4 PowerShell surfaces' {
    It 'rejects invalid native iWork package bytes' {
        Get-Command Open-OfficeIWork -ErrorAction Stop | Out-Null
        $invalid = Join-Path $TestDrive 'not-iwork.pages'
        [IO.File]::WriteAllText($invalid, 'not an iWork package')
        { Open-OfficeIWork -Path $invalid -ErrorAction Stop } | Should -Throw
    }

    It 'reads native OneNote and converts it with explicit fidelity evidence' {
        $oneNoteAssembly = $oneNoteSectionType.Assembly
        $newOneNoteObject = {
            param([string] $TypeName)
            [Activator]::CreateInstance($oneNoteAssembly.GetType($TypeName, $true))
        }
        $section = & $newOneNoteObject 'OfficeIMO.OneNote.OneNoteSection'
        $section.Name = 'Operations'
        $page = & $newOneNoteObject 'OfficeIMO.OneNote.OneNotePage'
        $page.Title = 'Daily checks'
        $outline = & $newOneNoteObject 'OfficeIMO.OneNote.OneNoteOutline'
        $paragraph = & $newOneNoteObject 'OfficeIMO.OneNote.OneNoteParagraph'
        $run = & $newOneNoteObject 'OfficeIMO.OneNote.OneNoteTextRun'
        $run.Text = 'Review backup status'
        $paragraph.Runs.Add($run)
        $outline.Children.Add($paragraph)
        $page.Outlines.Add($outline)
        $section.Pages.Add($page)

        $onePath = Join-Path $TestDrive 'operations.one'
        $section.Save($onePath)
        $loaded = Open-OfficeOneNote -Path $onePath
        $loaded.GetType().FullName | Should -Be 'OfficeIMO.OneNote.OneNoteSection'
        $loaded.Pages[0].Title | Should -Be 'Daily checks'

        $markdownPath = Join-Path $TestDrive 'operations.md'
        $markdownReport = Convert-OfficeDocument -Path $onePath -OutputPath $markdownPath
        $markdownReport | Should -BeOfType OfficeIMO.OfficeConversionReport
        Get-Content -LiteralPath $markdownPath -Raw | Should -Match 'Review backup status'

        [IO.File]::WriteAllText($markdownPath, 'caller-owned')
        { Convert-OfficeDocument -Path $onePath -OutputPath $markdownPath -NoClobber -ErrorAction Stop } | Should -Throw
        [IO.File]::ReadAllText($markdownPath) | Should -Be 'caller-owned'
        Convert-OfficeDocument -Path $onePath -OutputPath $markdownPath | Should -BeOfType OfficeIMO.OfficeConversionReport
        Get-Content -LiteralPath $markdownPath -Raw | Should -Match 'Review backup status'

        $htmlPath = Join-Path $TestDrive 'operations.html'
        Convert-OfficeDocument -Path $onePath -OutputPath $htmlPath | Should -BeOfType OfficeIMO.OfficeConversionReport
        Get-Content -LiteralPath $htmlPath -Raw | Should -Match 'Review backup status'

        $pdfPath = Join-Path $TestDrive 'operations.pdf'
        $pdfEvidence = Convert-OfficeDocument -Path $onePath -OutputPath $pdfPath
        $pdfEvidence.GetType().FullName | Should -Be 'OfficeIMO.OfficeConversionReport'
        Read-OfficeDocument -As Text -Path $pdfPath | Should -Match 'Review backup status'

        # Semantic PDF projection flattens an explicitly sized OneNote canvas.
        $page.Width = 1200
        $section.Save($onePath)
        $lossPath = Join-Path $TestDrive 'operations-loss.pdf'
        [IO.File]::WriteAllText($lossPath, 'caller-owned')
        {
            Convert-OfficeDocument -Path $onePath -OutputPath $lossPath -FailOnLoss -ErrorAction Stop
        } | Should -Throw
        [IO.File]::ReadAllText($lossPath) | Should -Be 'caller-owned'
    }

    It 'does not silently ignore OneNote projection options supplied beside PDF options' {
        $output = Join-Path $TestDrive 'conflicting-options.pdf'
        $source = Join-Path $TestDrive 'options-source.one'
        $section = $oneNoteSectionType::Create()
        $page = $oneNotePageType::new()
        $page.Title = 'Projection options'
        $section.Pages.Add($page)
        $section.Save($source)
        $projection = $oneNoteMarkdownOptionsType::new()
        $pdf = $oneNotePdfOptionsType::new()
        { Convert-OfficeDocument -Path $source -OutputPath $output -Options $projection, $pdf -ErrorAction Stop } |
            Should -Throw '*projection*'
        Test-Path -LiteralPath $output | Should -BeFalse
    }

    It 'rejects OneNote renderer options for a different output format' -TestCases @(
        @{ Type = 'OfficeIMO.Markdown.MarkdownHtmlExportOptions'; Extension = '.md' }
        @{ Type = 'OfficeIMO.Markdown.MarkdownHtmlExportOptions'; Extension = '.pdf' }
        @{ Type = 'OfficeIMO.OneNote.Pdf.OneNotePdfExportOptions'; Extension = '.md' }
        @{ Type = 'OfficeIMO.OneNote.Pdf.OneNotePdfExportOptions'; Extension = '.html' }
    ) {
        param($Type, $Extension)
        $source = Join-Path $TestDrive 'renderer-options.one'
        $section = $oneNoteSectionType::Create()
        $page = $oneNotePageType::new()
        $page.Title = 'Renderer options'
        $section.Pages.Add($page)
        $section.Save($source)
        $optionsType = Get-TestPSWriteOfficeType -AssemblyName ($Type -replace '\.[^.]+$', '') -TypeName $Type
        $options = [Activator]::CreateInstance($optionsType)
        $output = Join-Path $TestDrive ("mismatched-options$Extension")
        { Convert-OfficeDocument -Path $source -OutputPath $output -Options $options -ErrorAction Stop } |
            Should -Throw '*option*'
        Test-Path -LiteralPath $output | Should -BeFalse
    }

    It 'keeps structural package inspection separate from active-content policy' {
        $path = Join-Path $TestDrive 'security.docx'
        New-OfficeWord -Path $path {
            WordParagraph 'Package security fixture'
        } | Out-Null

        $securityReportType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.Core' -TypeName 'OfficeIMO.OfficePackageSecurityReport'
        $securityOptionsType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.Core' -TypeName 'OfficeIMO.OfficePackageSecurityOptions'
        $inventoryReport = Test-OfficeDocument -Path $path -Check Security
        $nativeInventory = @($inventoryReport.NativeReports | Where-Object { $securityReportType.IsInstanceOfType($_) })
        $nativeInventory | Should -HaveCount 1
        $inventory = $nativeInventory[0]
        $inventory.ContainerKind.ToString() | Should -Be 'OpenXml'
        $inventory.IsValid | Should -BeTrue
        $inventory.MacroPartCount | Should -Be 0

        $untrusted = Test-OfficeDocument -Path $path -Check Security -Untrusted
        $untrusted.IsValid | Should -BeTrue

        # The security inventory recognizes package-part identity without executing its bytes.
        $activePath = Join-Path $TestDrive 'active-content.docx'
        Copy-Item -LiteralPath $path -Destination $activePath
        $archive = [IO.Compression.ZipFile]::Open($activePath, [IO.Compression.ZipArchiveMode]::Update)
        try {
            $part = $archive.CreateEntry('word/vbaProject.bin')
            $writer = [IO.StreamWriter]::new($part.Open())
            try { $writer.Write('Inert package security inventory fixture') }
            finally { $writer.Dispose() }
        } finally { $archive.Dispose() }

        $activeHash = (Get-FileHash -LiteralPath $activePath -Algorithm SHA256).Hash
        $secureDefaults = $securityOptionsType::SecureDefaults
        $secureDefaults.MaxPartCount = 512
        $allowed = Test-OfficeDocument -Path $activePath -Check Security -Options $secureDefaults
        $allowed.IsValid | Should -BeTrue
        $combined = Test-OfficeDocument -Path $activePath -Check Security -Options $secureDefaults -Untrusted
        $combined.IsValid | Should -BeFalse
        $combinedEvidence = @($combined.NativeReports | Where-Object { $securityReportType.IsInstanceOfType($_) })
        $combinedEvidence | Should -HaveCount 1
        $combinedEvidence[0].ContainerKind.ToString() | Should -Be 'OpenXml'
        $combinedEvidence[0].IsValid | Should -BeFalse
        $combinedEvidence[0].MacroPartCount | Should -Be 1
        @($combinedEvidence[0].Findings | Where-Object { $_.Rule.ToString() -eq 'Macros' -and $_.Severity.ToString() -eq 'Error' }).Count | Should -Be 1
        $secureDefaults.MaxPartCount | Should -Be 512
        foreach ($policy in 'Macros', 'ActiveX', 'EmbeddedPayloads', 'ExternalRelationships') {
            $secureDefaults.$policy.ToString() | Should -Be 'Allow'
        }

        $secureDefaults.MaxPackageBytes = 1
        $limited = Test-OfficeDocument -Path $path -Check Security -Options $secureDefaults -Untrusted
        $limited.IsValid | Should -BeFalse
        $limitEvidence = @($limited.NativeReports | Where-Object { $_.GetType().FullName -eq 'OfficeIMO.OfficePackageSecurityException' })
        $limitEvidence | Should -HaveCount 1
        $limitEvidence[0].Rule.ToString() | Should -Be PackageSize
        $limitEvidence[0].Limit | Should -Be 1
        $secureDefaults.MaxPackageBytes | Should -Be 1
        $secureDefaults.Macros.ToString() | Should -Be Allow
        (Get-FileHash -LiteralPath $activePath -Algorithm SHA256).Hash | Should -Be $activeHash
    }

    It 'reports provenance evidence without claiming authorship' {
        $path = Join-Path $TestDrive 'provenance.txt'
        [IO.File]::WriteAllText($path, "Human-readable text`n")

        $assessment = Get-OfficeDocumentProvenance -Path $path
        $assessment.Structural.Format.ToString() | Should -Be 'StructuredText'
        $assessment.Verification | Should -BeNullOrEmpty
        $assessment.HasVerifiedContentCredential | Should -BeFalse

        $unavailable = Get-OfficeDocumentProvenance -Path $path -C2paToolPath 'definitely-missing-c2patool-for-pswriteoffice'
        $unavailable.Verification.Status.ToString() | Should -Be 'ProviderUnavailable'
        $unavailable.HasVerifiedContentCredential | Should -BeFalse
    }
}
