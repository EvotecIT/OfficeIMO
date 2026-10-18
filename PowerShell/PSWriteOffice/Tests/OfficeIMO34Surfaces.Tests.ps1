BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
}

Describe 'OfficeIMO 3.4 PowerShell surfaces' {
    It 'exposes bounded iWork inspection and loss-aware conversion contracts' {
        $get = Get-Command Get-OfficeIWork
        $convert = Get-Command ConvertFrom-OfficeIWork

        $get.Parameters['Options'].ParameterType.FullName | Should -Be 'OfficeIMO.IWork.IWorkReadOptions'
        [Nullable]::GetUnderlyingType($get.Parameters['Kind'].ParameterType).FullName | Should -Be 'OfficeIMO.IWork.IWorkDocumentKind'
        $convert.Parameters['ReadOptions'].ParameterType.FullName | Should -Be 'OfficeIMO.IWork.IWorkReadOptions'
        $convert.Parameters['ConversionOptions'].ParameterType.FullName | Should -Be 'OfficeIMO.IWork.IWorkConversionOptions'
        $convert.Parameters.Keys | Should -Contain 'FailOnLoss'
        $convert.Parameters.Keys | Should -Contain 'PassThruReport'

        $invalid = Join-Path $TestDrive 'not-iwork.pages'
        [IO.File]::WriteAllText($invalid, 'not an iWork package')
        { Get-OfficeIWork -Path $invalid -ErrorAction Stop } | Should -Throw
    }

    It 'reads native OneNote and converts it with explicit fidelity evidence' {
        $oneNoteAssembly = (Get-Command Get-OfficeOneNote).Parameters['Options'].ParameterType.Assembly
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
        $loaded = Get-OfficeOneNote -Path $onePath
        $loaded.GetType().FullName | Should -Be 'OfficeIMO.OneNote.OneNoteSection'
        $loaded.Pages[0].Title | Should -Be 'Daily checks'

        $markdownPath = Join-Path $TestDrive 'operations.md'
        $markdownReport = ConvertFrom-OfficeOneNote -Path $onePath -OutputPath $markdownPath -PassThruReport
        $markdownReport.GetType().FullName | Should -Be 'OfficeIMO.OneNote.Markdown.OneNoteMarkdownConversionReport'
        Get-Content -LiteralPath $markdownPath -Raw | Should -Match 'Review backup status'

        [IO.File]::WriteAllText($markdownPath, 'caller-owned')
        { ConvertFrom-OfficeOneNote -Path $onePath -OutputPath $markdownPath -ErrorAction Stop } | Should -Throw
        [IO.File]::ReadAllText($markdownPath) | Should -Be 'caller-owned'
        ConvertFrom-OfficeOneNote -Path $onePath -OutputPath $markdownPath -Force | Should -BeOfType System.IO.FileInfo
        Get-Content -LiteralPath $markdownPath -Raw | Should -Match 'Review backup status'

        $htmlPath = Join-Path $TestDrive 'operations.html'
        ConvertFrom-OfficeOneNote -Path $onePath -OutputPath $htmlPath | Should -BeOfType System.IO.FileInfo
        Get-Content -LiteralPath $htmlPath -Raw | Should -Match 'Review backup status'

        $pdfPath = Join-Path $TestDrive 'operations.pdf'
        $pdfEvidence = ConvertFrom-OfficeOneNote -Path $onePath -OutputPath $pdfPath -PassThruReport
        $pdfEvidence.GetType().FullName | Should -Be 'OfficeIMO.Pdf.PdfDocumentConversionResult'
        Get-OfficePdfText -Path $pdfPath | Should -Match 'Review backup status'

        $run.Text = 'Arabic layout evidence: ' + (-join @(0x0645, 0x0631, 0x062D, 0x0628, 0x0627 | ForEach-Object { [char] $_ }))
        $section.Save($onePath)
        $lossPath = Join-Path $TestDrive 'operations-loss.pdf'
        [IO.File]::WriteAllText($lossPath, 'caller-owned')
        {
            ConvertFrom-OfficeOneNote `
                -Path $onePath `
                -OutputPath $lossPath `
                -FailOnLoss `
                -Force `
                -ErrorAction Stop
        } | Should -Throw
        [IO.File]::ReadAllText($lossPath) | Should -Be 'caller-owned'
    }

    It 'does not silently ignore OneNote projection options supplied beside PDF options' {
        $output = Join-Path $TestDrive 'conflicting-options.pdf'
        $command = Get-Command ConvertFrom-OfficeOneNote
        $projection = [Activator]::CreateInstance($command.Parameters['ProjectionOptions'].ParameterType)
        $pdf = [Activator]::CreateInstance($command.Parameters['PdfOptions'].ParameterType)

        { ConvertFrom-OfficeOneNote -Path (Join-Path $TestDrive 'unused.one') -OutputPath $output -ProjectionOptions $projection -PdfOptions $pdf -ErrorAction Stop } |
            Should -Throw '*either through -ProjectionOptions or through -PdfOptions*'
    }

    It 'rejects OneNote renderer options for a different output format' -TestCases @(
        @{ Parameter = 'HtmlOptions'; Extension = '.md'; Message = '*HtmlOptions can only be used with .html output*' }
        @{ Parameter = 'HtmlOptions'; Extension = '.pdf'; Message = '*HtmlOptions can only be used with .html output*' }
        @{ Parameter = 'PdfOptions'; Extension = '.md'; Message = '*PdfOptions can only be used with .pdf output*' }
        @{ Parameter = 'PdfOptions'; Extension = '.html'; Message = '*PdfOptions can only be used with .pdf output*' }
    ) {
        param($Parameter, $Extension, $Message)

        $command = Get-Command ConvertFrom-OfficeOneNote
        $options = [Activator]::CreateInstance($command.Parameters[$Parameter].ParameterType)
        $arguments = @{
            Path = Join-Path $TestDrive 'unused.one'
            OutputPath = Join-Path $TestDrive ("mismatched-options$Extension")
            ErrorAction = 'Stop'
        }
        $arguments[$Parameter] = $options

        { ConvertFrom-OfficeOneNote @arguments } | Should -Throw $Message
    }

    It 'keeps structural package inspection separate from active-content policy' {
        $path = Join-Path $TestDrive 'security.docx'
        New-OfficeWord -Path $path {
            WordParagraph 'Package security fixture'
        } | Out-Null

        $inventory = Get-OfficePackageSecurity -Path $path
        $inventory.ContainerKind.ToString() | Should -Be 'OpenXml'
        $inventory.IsValid | Should -BeTrue
        $inventory.MacroPartCount | Should -Be 0

        $untrusted = Get-OfficePackageSecurity -Path $path -Untrusted
        $untrusted.IsValid | Should -BeTrue
        $securityOptionsType = (Get-Command Get-OfficePackageSecurity).Parameters['Options'].ParameterType
        $secureDefaults = $securityOptionsType.GetProperty('SecureDefaults').GetValue($null)
        { Get-OfficePackageSecurity -Path $path -Options $secureDefaults -Untrusted -ErrorAction Stop } | Should -Throw
    }

    It 'reports provenance evidence without claiming authorship' {
        $path = Join-Path $TestDrive 'provenance.txt'
        [IO.File]::WriteAllText($path, "Human-readable text`n")

        $assessment = Get-OfficeProvenance -Path $path
        $assessment.Structural.Format.ToString() | Should -Be 'StructuredText'
        $assessment.Verification | Should -BeNullOrEmpty
        $assessment.HasVerifiedContentCredential | Should -BeFalse

        $unavailable = Get-OfficeProvenance -Path $path -C2paToolPath 'definitely-missing-c2patool-for-pswriteoffice'
        $unavailable.Verification.Status.ToString() | Should -Be 'ProviderUnavailable'
        $unavailable.HasVerifiedContentCredential | Should -BeFalse
    }
}
