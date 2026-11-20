BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop

    function New-ChartData {
        [OfficeIMO.Drawing.OfficeChartData]::new([string[]] @('Q1', 'Q2'),
            [OfficeIMO.OfficeSeries[]] @([OfficeIMO.OfficeSeries]::new('Revenue', [double[]] @(12, 18))))
    }
}

Describe 'Word and PowerPoint chart image export' {
    It 'exports a named chart from a saved Word document' {
        $path = Join-Path $TestDrive 'report.docx'
        $document = [OfficeIMO.Word.WordDocument]::Create($path)
        try {
            $chart = $document.AddChart([OfficeIMO.OfficeChartType]::ColumnClustered, (New-ChartData))
            $chart.Name = 'Revenue chart'
            $document.Save()
        } finally { $document.Dispose() }

        $result = Export-OfficeDocumentImage -Path $path -Chart 'Revenue chart' -OutputDirectory (Join-Path $TestDrive 'word')

        $result.Name | Should -Be 'Revenue chart'
        Test-Path -LiteralPath $result.SavedPath | Should -BeTrue
        [System.IO.Path]::GetFileName($result.SavedPath) | Should -Be 'report-Revenue chart.png'
    }

    It 'keeps a numbered file for each Word chart that shares a long name' {
        $path = Join-Path $TestDrive 'long.docx'
        $name = 'Revenue ' + ('x' * 130)
        $document = [OfficeIMO.Word.WordDocument]::Create($path)
        try {
            foreach ($index in 1..2) {
                $chart = $document.AddChart([OfficeIMO.OfficeChartType]::ColumnClustered, (New-ChartData))
                $chart.Name = $name
            }
            $document.Save()
        } finally { $document.Dispose() }

        $results = @(Export-OfficeDocumentImage -Path $path -Chart $name -OutputDirectory (Join-Path $TestDrive 'long'))

        $results.Count | Should -Be 2
        @($results.SavedPath | Select-Object -Unique).Count | Should -Be 2
        [System.IO.Path]::GetFileNameWithoutExtension($results[1].SavedPath) | Should -BeLike '*-2'
    }

    It 'returns a named PowerPoint chart as an SVG image and rejects an unknown chart name' {
        $path = Join-Path $TestDrive 'deck.pptx'
        $deck = [OfficeIMO.PowerPoint.PowerPointDocument]::Create($path, $null)
        try {
            $chart = $deck.AddSlide().AddChart([OfficeIMO.OfficeChartType]::ColumnClustered, (New-ChartData))
            $chart.Name = 'Revenue chart'
            $deck.Save()
        } finally { $deck.Dispose() }

        $result = ConvertTo-OfficeDocumentImage -Path $path -Chart 'Revenue chart' -Format Svg

        $result.Format | Should -Be 'Svg'
        [System.Text.Encoding]::UTF8.GetString($result.Bytes) | Should -BeLike '*<svg*'
        { ConvertTo-OfficeDocumentImage -Path $path -Chart 'Missing' -ErrorAction Stop } | Should -Throw "*Chart 'Missing' does not exist*"
    }
}
