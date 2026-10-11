BeforeAll {
    $manifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) { $env:PSWRITEOFFICE_MODULE_MANIFEST } else { Join-Path $PSScriptRoot '../PSWriteOffice.psd1' }
    Import-Module $manifest -Global -ErrorAction Stop
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')

    function Read-SavedChart {
        param([string] $Path)
        $entry = @(Get-ZipEntriesLocal -Path $Path | Where-Object { $_ -match '^ppt/charts/chart[^/]+\.xml$' })
        $entry.Count | Should -Be 1
        Get-ZipXmlDocumentLocal -Path $Path -Entry $entry[0]
    }
}

Describe 'PowerPoint chart calculated properties' {
    It 'saves categorical labels and numeric series projected by Select-Object' {
        $rows = @(
            [pscustomobject]@{ Site = 'Łódź'; Tickets = 20 }
            [pscustomobject]@{ Site = 'Remote'; Tickets = 30 }
        ) | Select-Object @{Name='Location';Expression={$_.Site}}, @{Name='Reviewed tickets';Expression={$_.Tickets * 2}}
        $path = Join-Path $TestDrive 'categorical.pptx'
        $deck = New-OfficePowerPoint -Path $path -NoSave
        try {
            $slide = Add-OfficePowerPointSlide -Presentation $deck -Layout 1 -PassThru
            Add-OfficePowerPointChart -Slide $slide -InputObject $rows -CategoryProperty Location -SeriesProperty 'Reviewed tickets' -ErrorAction Stop
            Save-OfficePowerPoint -Presentation $deck
        } finally { $deck.Dispose() }
        $chart = Read-SavedChart $path
        ($chart.SelectNodes("//*[local-name()='cat']//*[local-name()='pt']/*[local-name()='v']") | ForEach-Object InnerText) -join ',' | Should -Be 'Łódź,Remote'
        ($chart.SelectNodes("//*[local-name()='val']//*[local-name()='pt']/*[local-name()='v']") | ForEach-Object InnerText) -join ',' | Should -Be '40,60'
    }

    It 'saves calculated numeric properties on both scatter axes' {
        $rows = @(
            [pscustomobject]@{ Period = 1; Tickets = 20 }
            [pscustomobject]@{ Period = 2; Tickets = 30 }
        ) | Select-Object @{Name='Quarter';Expression={$_.Period}}, @{Name='Demand';Expression={$_.Tickets / 2}}
        $path = Join-Path $TestDrive 'scatter.pptx'
        $deck = New-OfficePowerPoint -Path $path -NoSave
        try {
            $slide = Add-OfficePowerPointSlide -Presentation $deck -Layout 1 -PassThru
            Add-OfficePowerPointChart -Slide $slide -Type Scatter -InputObject $rows -XProperty Quarter -YProperty Demand -ErrorAction Stop
            Save-OfficePowerPoint -Presentation $deck
        } finally { $deck.Dispose() }
        $chart = Read-SavedChart $path
        ($chart.SelectNodes("//*[local-name()='xVal']//*[local-name()='pt']/*[local-name()='v']") | ForEach-Object InnerText) -join ',' | Should -Be '1,2'
        ($chart.SelectNodes("//*[local-name()='yVal']//*[local-name()='pt']/*[local-name()='v']") | ForEach-Object InnerText) -join ',' | Should -Be '10,15'
    }

    It 'still rejects null and nonnumeric projected series values' {
        $deck = New-OfficePowerPoint -Path (Join-Path $TestDrive 'invalid.pptx') -NoSave
        try {
            $slide = Add-OfficePowerPointSlide -Presentation $deck -Layout 1 -PassThru
            $bad = [pscustomobject]@{ Site = 'Depot'; Value = 'pending' } | Select-Object Site, @{Name='Tickets';Expression={$_.Value}}
            { Add-OfficePowerPointChart -Slide $slide -InputObject $bad -CategoryProperty Site -SeriesProperty Tickets -ErrorAction Stop } | Should -Throw '*must be numeric*'
            $nullValue = [pscustomobject]@{ Site = 'Depot' } | Select-Object Site, @{Name='Tickets';Expression={$null}}
            { Add-OfficePowerPointChart -Slide $slide -InputObject $nullValue -CategoryProperty Site -SeriesProperty Tickets -ErrorAction Stop } | Should -Throw '*cannot be null*'
            @($slide.Charts).Count | Should -Be 0
        } finally { $deck.Dispose() }
    }
}
