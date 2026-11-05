BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop

    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
}

Describe 'Easy local OCR commands' {
    It 'recognizes image text with runtime and word evidence through the canonical reader' {
        $runtime = New-TestTesseractExecutable -Directory $TestDrive
        $options = New-OfficeOcrOptions -TesseractPath $runtime -Language English, Polish -NoLanguageDownload
        $review = Read-OfficeDocument -Path (Join-Path $PSScriptRoot 'Assets\CellImage.png') -As Ocr -OcrOptions $options
        $review.Ocr.Text | Should -Be 'OfficeIMO OCR'
        $review.Ocr.Pages[0].Provider | Should -Be 'tesseract-cli'
        $review.Ocr.Pages[0].Model | Should -Be 'tessdata:eng+pol'
        @($review.Ocr.Pages[0].Words).Count | Should -BeGreaterThan 0
    }

    It 'rejects unsupported OCR review input before starting an unavailable provider for <Extension>' -TestCases @(
        @{ Extension = '.txt'; Content = 'Unsupported OCR review source' }
        @{ Extension = '.svg'; Content = "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'><rect width='10' height='10'/></svg>" }
    ) {
        param($Extension, $Content)
        $unsupportedPath = Join-Path $TestDrive ('unsupported-review' + $Extension)
        Set-Content -LiteralPath $unsupportedPath -Value $Content
        $options = New-OfficeOcrOptions -TesseractPath 'definitely-missing-tesseract-for-pswriteoffice' -NoLanguageDownload
        { Read-OfficeDocument -Path $unsupportedPath -As Ocr -OcrOptions $options -ErrorAction Stop } |
            Should -Throw '*Image content is not a supported image*'
        $output = Join-Path $TestDrive ('unsupported-' + $Extension.TrimStart('.') + '.txt')
        { Convert-OfficeDocument -Ocr -Path $unsupportedPath -To Text -OutputPath $output -OcrOptions $options -ErrorAction Stop } |
            Should -Throw '*Image content is not a supported image*'
        Test-Path -LiteralPath $output | Should -BeFalse
    }

    It 'validates OCR session options before opening input files' {
        $missingPdf = Join-Path $TestDrive 'missing-input.pdf'
        $outputPdf = Join-Path $TestDrive 'missing-output.pdf'
        { New-OfficeOcrOptions -Language English -LanguageExpression eng -NoLanguageDownload -ErrorAction Stop } |
            Should -Throw '*Language*LanguageExpression*not both*'
        { Update-OfficePdf -Searchable -Path $missingPdf -OutputPath $outputPdf -OcrOptions (
                New-OfficeOcrOptions -Language English -LanguageExpression eng -NoLanguageDownload -ErrorAction Stop
            ) -ErrorAction Stop } | Should -Throw '*Language*LanguageExpression*not both*'
        Test-Path -LiteralPath $outputPdf | Should -BeFalse
    }

    It 'preserves captured session options and applies explicit language overrides independently' {
        $runtime = New-TestTesseractExecutable -Directory $TestDrive
        $session = (New-OfficeOcrOptions).SessionOptions
        $session.Engine.Language = 'pol'
        $baseline = New-OfficeOcrOptions -SessionOptions $session -TesseractPath $runtime -NoLanguageDownload
        $friendly = New-OfficeOcrOptions -SessionOptions $session -TesseractPath $runtime -Language English, Polish -NoLanguageDownload
        $raw = New-OfficeOcrOptions -SessionOptions $session -TesseractPath $runtime -LanguageExpression 'eng+pol' -NoLanguageDownload
        $image = Join-Path $PSScriptRoot 'Assets\CellImage.png'
        (Read-OfficeDocument -Path $image -As Ocr -OcrOptions $baseline).Ocr.Pages[0].Model | Should -Be 'tessdata:pol'
        (Read-OfficeDocument -Path $image -As Ocr -OcrOptions $friendly).Ocr.Pages[0].Model | Should -Be 'tessdata:eng+pol'
        (Read-OfficeDocument -Path $image -As Ocr -OcrOptions $raw).Ocr.Pages[0].Model | Should -Be 'tessdata:eng+pol'
        $session.Engine.Language | Should -Be 'pol'
        $baseline.SessionOptions.Engine.Language | Should -Be 'pol'
    }

    It 'writes a searchable PDF whose OCR text is readable' {
        $runtime = New-TestTesseractExecutable -Directory $TestDrive
        $inputPath = Join-Path $TestDrive 'scan.pdf'
        $outputPath = Join-Path $TestDrive 'scan-searchable.pdf'
        $imagePath = Join-Path $PSScriptRoot 'Assets\CellImage.png'

        New-OfficePdf -Path $inputPath {
            PdfImage -Path $imagePath -Width 180 -Height 120
        }

        $result = Update-OfficePdf -Searchable `
            -Path $inputPath `
            -OutputPath $outputPath `
            -PassThru `
                -OcrOptions (New-OfficeOcrOptions -TesseractPath $runtime -NoLanguageDownload)

        $result.WasModified | Should -BeTrue
        $result.AddedWordCount | Should -Be 2
        Test-Path -LiteralPath $outputPath | Should -BeTrue
        Read-OfficeDocument -As Text -Path $outputPath | Should -Match 'OfficeIMO OCR'
    }

    It 'loads password-encrypted PDFs for searchable OCR' {
        $runtime = New-TestTesseractExecutable -Directory $TestDrive
        $inputPath = Join-Path $TestDrive 'encrypted-scan.pdf'
        $outputPath = Join-Path $TestDrive 'encrypted-searchable.pdf'
        $imagePath = Join-Path $PSScriptRoot 'Assets\CellImage.png'

        New-OfficePdf -Path $inputPath -Password 'open' {
            PdfImage -Path $imagePath -Width 180 -Height 120
        }

        {
            Update-OfficePdf -Searchable `
                -Path $inputPath `
                -OutputPath $outputPath `
                -Password 'wrong' `
                -OcrOptions (New-OfficeOcrOptions -TesseractPath $runtime -NoLanguageDownload) `
                -ErrorAction Stop
        } | Should -Throw

        $result = Update-OfficePdf -Searchable `
            -Path $inputPath `
            -OutputPath $outputPath `
            -Password 'open' `
            -PassThru `
                -OcrOptions (New-OfficeOcrOptions -TesseractPath $runtime -NoLanguageDownload)

        $result.WasModified | Should -BeTrue
        Read-OfficeDocument -As Text -Path $outputPath -Password 'open' | Should -Match 'OfficeIMO OCR'
    }

    It 'fails instead of reporting output when the destination appears during OCR' {
        $inputPath = Join-Path $TestDrive 'late-destination-input.pdf'
        $outputPath = Join-Path $TestDrive 'late-destination-output.pdf'
        $runtime = New-TestTesseractExecutable -Directory $TestDrive -CreateFilePath $outputPath
        $imagePath = Join-Path $PSScriptRoot 'Assets\CellImage.png'

        New-OfficePdf -Path $inputPath {
            PdfImage -Path $imagePath -Width 180 -Height 120
        }

        {
            Update-OfficePdf -Searchable -NoClobber `
                -Path $inputPath `
                -OutputPath $outputPath `
                -OcrOptions (New-OfficeOcrOptions -TesseractPath $runtime -NoLanguageDownload) `
                -ErrorAction Stop
        } | Should -Throw

        [IO.File]::ReadAllText($outputPath).Trim() | Should -Be 'late destination'
    }

    It 'preserves an existing searchable PDF destination with NoClobber and otherwise replaces it' {
        $runtime = New-TestTesseractExecutable -Directory $TestDrive
        $inputPath = Join-Path $TestDrive 'force-input.pdf'
        $outputPath = Join-Path $TestDrive 'force-output.pdf'
        $imagePath = Join-Path $PSScriptRoot 'Assets\CellImage.png'

        New-OfficePdf -Path $inputPath {
            PdfImage -Path $imagePath -Width 180 -Height 120
        }
        [IO.File]::WriteAllText($outputPath, 'caller-owned')

        {
            Update-OfficePdf -Searchable -NoClobber `
                -Path $inputPath `
                -OutputPath $outputPath `
                -OcrOptions (New-OfficeOcrOptions -TesseractPath $runtime -NoLanguageDownload) `
                -ErrorAction Stop
        } | Should -Throw
        [IO.File]::ReadAllText($outputPath) | Should -Be 'caller-owned'

        Update-OfficePdf -Searchable `
            -Path $inputPath `
            -OutputPath $outputPath `
                -OcrOptions (New-OfficeOcrOptions -TesseractPath $runtime -NoLanguageDownload)
        Read-OfficeDocument -As Text -Path $outputPath | Should -Match 'OfficeIMO OCR'
    }
}
