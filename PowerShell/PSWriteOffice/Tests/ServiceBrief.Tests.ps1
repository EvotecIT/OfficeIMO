BeforeAll {
    $manifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) { $env:PSWRITEOFFICE_MODULE_MANIFEST } else { Join-Path $PSScriptRoot '../PSWriteOffice.psd1' }
    Import-Module $manifest -Global -ErrorAction Stop
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    $script = Join-Path $PSScriptRoot '../Examples/Showcase/Showcase-PowerPoint-ServiceBrief.ps1'
    function Read-ServiceBriefText {
        param([string] $Path)
        $presentation = Get-ZipXmlDocumentLocal -Path $Path -Entry 'ppt/presentation.xml'
        $count = $presentation.SelectNodes("/*[local-name()='presentation']/*[local-name()='sldIdLst']/*[local-name()='sldId']").Count
        $text = foreach ($index in 0..($count - 1)) {
            $entry = Get-TestPowerPointSlideEntry -Path $Path -Index $index
            $slide = Get-ZipXmlDocumentLocal -Path $Path -Entry $entry
            foreach ($node in $slide.SelectNodes("//*[local-name()='t']")) { $node.InnerText }
        }
        [pscustomobject]@{ Count=$count; Text=($text -join "`n") }
    }
}

Describe 'Service demand brief' {
    It 'derives the baseline, priority review and like-for-like chart from the CSV' {
        $output = Join-Path $TestDrive 'default'
        & $script -OutputDirectory $output -ModulePath $manifest | Out-Null
        $path = Join-Path $output 'Showcase-PowerPoint-ServiceBrief.pptx'
        $brief = Read-ServiceBriefText $path
        $brief.Count | Should -Be 8
        $brief.Text | Should -Match '600 people'
        $brief.Text | Should -Match '1080 quarterly tickets'
        $brief.Text | Should -Match '-10.0 % from 1200'
        $brief.Text | Should -Match 'Remote staff increased by 16 tickets'
        $brief.Text | Should -Match '15 Oct'
        $charts = @(Get-ZipEntriesLocal -Path $path | Where-Object { $_ -match '/charts/chart[^/]+\.xml$' })
        $charts.Count | Should -Be 1
        $chart = Get-ZipXmlDocumentLocal -Path $path -Entry $charts[0]
        $series = $chart.SelectNodes("//*[local-name()='ser']")
        $series.Count | Should -Be 2
        $names = @($series | ForEach-Object { $_.SelectSingleNode("*[local-name()='tx']//*[local-name()='v']").InnerText })
        ($names -join ',') | Should -Be 'Previous quarter,Current quarter'
        $sums = @($series | ForEach-Object {
            ($_.SelectNodes("*[local-name()='val']//*[local-name()='pt']/*[local-name()='v']") |
                ForEach-Object { [long]$_.InnerText } | Measure-Object -Sum).Sum
        })
        ($sums -join ',') | Should -Be '1200,1080'
    }

    It 'recomputes changed input and handles a zero baseline without inventing a percentage' {
        $input = Join-Path $TestDrive 'supplied.csv'
        @('Site,Owner,SupportedPeople,PreviousQuarterTickets,CurrentQuarterTickets',
            'Łódź support,Service team,75,0,23') | Set-Content -LiteralPath $input -Encoding UTF8
        $output = Join-Path $TestDrive 'supplied'
        & $script -InputCsv $input -ReviewDate '2026-12-31' -OutputDirectory $output -ModulePath $manifest | Out-Null
        $brief = Read-ServiceBriefText (Join-Path $output 'Showcase-PowerPoint-ServiceBrief.pptx')
        $brief.Text | Should -Match '75 people'
        $brief.Text | Should -Match '23 quarterly tickets'
        $brief.Text | Should -Match 'no prior-quarter baseline'
        $brief.Text | Should -Match 'Łódź support increased by 23 tickets'
        $brief.Text | Should -Match 'Supplied CSV'
        $brief.Text | Should -Match '15 Jan'
    }

    It 'rejects invalid counts and duplicate locations before writing a document' {
        $cases = @(
            @('Site,Owner,SupportedPeople,PreviousQuarterTickets,CurrentQuarterTickets', 'Depot,Team,10,-1,20'),
            @('Site,Owner,SupportedPeople,PreviousQuarterTickets,CurrentQuarterTickets', 'Depot,Team,0,10,20'),
            @('Site,Owner,SupportedPeople,PreviousQuarterTickets,CurrentQuarterTickets', 'Depot,Team,10,10,20', 'Depot,Other,20,30,40')
        )
        $number = 0
        foreach ($case in $cases) {
            $input = Join-Path $TestDrive "invalid-$number.csv"
            $output = Join-Path $TestDrive "invalid-$number"
            $case | Set-Content -LiteralPath $input -Encoding UTF8
            { & $script -InputCsv $input -OutputDirectory $output -ModulePath $manifest } | Should -Throw
            Test-Path -LiteralPath $output | Should -BeFalse
            $number++
        }
    }
}
