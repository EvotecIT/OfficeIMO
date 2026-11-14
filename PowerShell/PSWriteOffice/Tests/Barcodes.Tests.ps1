BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
}

Describe 'QR codes and barcodes' {
    It 'adds a QR code and a sized Code 128 barcode to Word paragraphs' {
        $path = Join-Path $TestDrive 'codes.docx'
        New-OfficeWord -Path $path {
            WordParagraph 'Scan for the agenda:'
            WordBarcode 'https://evotec.xyz/events'
            WordParagraph { WordBarcode 'ORD-2026-0042' -Type Code128 -Width 6cm -AlternativeText 'Order number' }
        } | Out-Null

        $document = [OfficeIMO.Word.WordDocument]::Load($path)
        try {
            $document.Images.Count | Should -Be 2
            $document.Images[0].AltText | Should -Be 'QR Code: https://evotec.xyz/events'
            [math]::Round($document.Images[0].Width.Value.ToPoints(), 1) | Should -Be 90
            $document.Images[1].AltText | Should -Be 'Order number'
            [math]::Round($document.Images[1].Width.Value.ToCentimeters(), 2) | Should -Be 6
        } finally {
            $document.Dispose()
        }
    }

    It 'anchors an EAN-13 barcode at a worksheet cell' {
        $path = Join-Path $TestDrive 'codes.xlsx'
        New-OfficeExcel -Path $path {
            ExcelSheet 'Labels' { ExcelBarcode '590123412345' -Type Ean -Address C4 }
        } | Out-Null

        $workbook = [OfficeIMO.Excel.ExcelDocument]::Load($path)
        try {
            $image = @($workbook.Sheets[0].Images)
            $image.Count | Should -Be 1
            $image[0].RowIndex | Should -Be 4
            $image[0].ColumnIndex | Should -Be 3
            $image[0].Description | Should -Be 'EAN: 590123412345'
        } finally {
            $workbook.Dispose()
        }
    }

    It 'keeps a flowing slide code square and places a pinned code exactly' {
        $path = Join-Path $TestDrive 'codes.pptx'
        New-OfficePowerPoint -Path $path {
            PowerPointSlide -Title 'Slides' { PowerPointBarcode 'https://evotec.xyz/talks' }
            PowerPointSlide -Title 'Contact' { PowerPointBarcode 'https://evotec.xyz' -X 20cm -Y 12cm -Width 3cm }
        } | Out-Null

        $presentation = [OfficeIMO.PowerPoint.PowerPointDocument]::Load($path)
        try {
            $flowing = @($presentation.Slides[0].Images)[0]
            $flowing.Width.ToEmus() | Should -Be $flowing.Height.ToEmus()
            $flowing.AltText | Should -Be 'QR Code: https://evotec.xyz/talks'
            $pinned = @($presentation.Slides[1].Images)[0]
            [math]::Round($pinned.Left.ToCentimeters(), 2) | Should -Be 20
            [math]::Round($pinned.Top.ToCentimeters(), 2) | Should -Be 12
            [math]::Round($pinned.Width.ToCentimeters(), 2) | Should -Be 3
        } finally {
            $presentation.Dispose()
        }
    }

    It 'draws PDF codes as vector content in the DSL and on an existing document' {
        $dslPath = Join-Path $TestDrive 'codes.pdf'
        New-OfficePdf -Path $dslPath {
            PdfParagraph 'Pay online:'
            PdfBarcode 'https://pay.example.com/inv/2026-0042' -Alignment Right
        } | Out-Null
        $document = New-OfficePdf
        $updated = PdfBarcode 'SN-0042-PL' -Type DataMatrix -Document $document -PassThru
        [object]::ReferenceEquals($updated, $document) | Should -BeTrue

        @(Get-OfficePdfImage -Path $dslPath).Count | Should -Be 0
        (Get-Item $dslPath).Length | Should -BeGreaterThan 0
    }

    It 'rejects options that do not apply to the symbology' {
        $path = Join-Path $TestDrive 'invalid.docx'
        { New-OfficeWord -Path $path { WordBarcode 'A1' -Type Code128 -ErrorCorrection H } } | Should -Throw '*QR codes only*'
        Test-Path $path | Should -BeFalse
    }
}
