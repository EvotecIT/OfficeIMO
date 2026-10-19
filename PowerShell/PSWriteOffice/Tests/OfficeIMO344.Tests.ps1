BeforeAll {
    $moduleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $moduleManifest -Global -ErrorAction Stop
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
}

Describe 'OfficeIMO 3.4.4 integration' {
    It 'preserves carriage returns and CRLF through Excel save and readback' {
        $path = Join-Path $TestDrive 'line-endings.xlsx'
        $values = @("First`rSecond", "First`r`nSecond", "First`nSecond")
        New-OfficeExcel -Path $path {
            ExcelSheet 'Text' {
                ExcelCell -Address A1 -Value 'Text'
                ExcelCell -Address A2 -Value $values[0]
                ExcelCell -Address A3 -Value $values[1]
                ExcelCell -Address A4 -Value $values[2]
            }
        } | Out-Null

        $rows = @(Get-OfficeExcelData -Path $path -Sheet Text)
        $rows | Should -HaveCount 3
        for ($index = 0; $index -lt $values.Count; $index++) {
            [System.String]::Equals($rows[$index].Text, $values[$index], [System.StringComparison]::Ordinal) |
                Should -BeTrue
        }
    }

    It 'stamps images and watermarks from files and releases their handles' {
        $source = Join-Path $TestDrive 'source.pdf'
        $image = New-TestOfficeImageFile -Directory $TestDrive
        New-OfficePdf -Path $source { PdfParagraph 'Original body' } | Out-Null
        foreach ($watermark in @($false, $true)) {
            $output = Join-Path $TestDrive "stamped-$watermark.pdf"
            Add-OfficePdfStamp -Path $source -OutputPath $output -Image $image -X 20 -Y 20 -Width 30 -Height 30 -Watermark:$watermark
            (Get-OfficePdfPreflight -Path $output).CanRead | Should -BeTrue
            (Read-OfficePdf -Path $output).Images.Count | Should -BeGreaterThan 0
        }
        $handle = [System.IO.File]::Open($image, 'Open', 'ReadWrite', 'None')
        $handle.Dispose()
    }

    It 'rejects oversized image files across PDF composition and stamping' {
        $image = Join-Path $TestDrive 'oversized.png'
        $stream = [System.IO.File]::Create($image)
        try { $stream.SetLength(128MB + 1) } finally { $stream.Dispose() }
        $source = Join-Path $TestDrive 'bounded-source.pdf'
        New-OfficePdf -Path $source { PdfParagraph 'Original' } | Out-Null

        { New-OfficePdf { PdfImage -Path $image -Width 20 -Height 20 } -ErrorAction Stop } |
            Should -Throw '*Encoded image input*exceeds*byte limit*'
        { New-OfficePdf { Set-OfficePdfBackgroundImage -Path $image } -ErrorAction Stop } |
            Should -Throw '*Encoded image input*exceeds*byte limit*'
        { New-OfficePdfTableCellImage -Path $image -Width 20 -Height 20 -ErrorAction Stop } |
            Should -Throw '*Encoded image input*exceeds*byte limit*'
        foreach ($watermark in @($false, $true)) {
            $output = Join-Path $TestDrive "blocked-$watermark.pdf"
            { Add-OfficePdfStamp -Path $source -OutputPath $output -Image $image -Watermark:$watermark -ErrorAction Stop } |
                Should -Throw '*exceeds the configured maximum size*'
            Test-Path -LiteralPath $output | Should -BeFalse
        }
        $handle = [System.IO.File]::Open($image, 'Open', 'ReadWrite', 'None')
        $handle.Dispose()
    }

    It 'rejects oversized invoice XML before writing a PDF' {
        $xml = Join-Path $TestDrive 'oversized.xml'
        $stream = [System.IO.File]::Create($xml)
        try { $stream.SetLength(16MB + 1) } finally { $stream.Dispose() }
        $output = Join-Path $TestDrive 'blocked-invoice.pdf'
        { New-OfficePdf -Path $output {
                PdfElectronicInvoice -Path $xml
            } -ErrorAction Stop } | Should -Throw '*CII XML exceeds the maximum byte length*'
        Test-Path -LiteralPath $output | Should -BeFalse
        $handle = [System.IO.File]::Open($xml, 'Open', 'ReadWrite', 'None')
        $handle.Dispose()
    }
}
