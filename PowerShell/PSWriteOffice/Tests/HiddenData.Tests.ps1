BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop

    function New-ReviewedProposal([string] $Path) {
        $document = New-OfficeWord -Path $Path
        $document.AddParagraph('Proposal for the client.').AddComment('Jane Reviewer', 'JR', 'Check the price.') | Out-Null
        Set-OfficeWordDocumentProperty -Document $document -Name Creator -Value 'Jane Reviewer'
        Save-OfficeWord -Document $document
    }
}

Describe 'Hidden-data commands' {
    It 'reports findings for each supported file without changing it and continues past unsupported files' {
        $proposal = Join-Path $TestDrive 'inspect.docx'
        $notes = Join-Path $TestDrive 'notes.txt'
        New-ReviewedProposal $proposal
        Set-Content -LiteralPath $notes -Value 'Plain text'
        $hash = (Get-FileHash -LiteralPath $proposal).Hash

        $reports = @(Get-ChildItem -LiteralPath $proposal, $notes | Get-OfficeHiddenData -ErrorAction SilentlyContinue -ErrorVariable failures)

        $reports.Count | Should -Be 1
        $reports[0].Path | Should -Be $proposal
        $reports[0].OutputPath | Should -BeNullOrEmpty
        $reports[0].Changed | Should -BeFalse
        @($reports[0].Findings | Where-Object Kind -EQ 'Comments').Count | Should -Be 1
        @($reports[0].Findings | Where-Object Outcome -NE 'Found').Count | Should -Be 0
        @($failures).Count | Should -Be 1
        $failures[0].FullyQualifiedErrorId | Should -BeLike 'hidden-data-unsupported-format*'
        (Get-FileHash -LiteralPath $proposal).Hash | Should -Be $hash
    }

    It 'writes a cleaned copy beside the input, keeps it with NoClobber and leaves the input alone' {
        $proposal = Join-Path $TestDrive 'remove.docx'
        $cleaned = Join-Path $TestDrive 'remove.cleaned.docx'
        New-ReviewedProposal $proposal
        $hash = (Get-FileHash -LiteralPath $proposal).Hash

        Remove-OfficeHiddenData -Path $proposal -WhatIf
        Test-Path -LiteralPath $cleaned | Should -BeFalse

        $report = Remove-OfficeHiddenData -Path $proposal
        $report.OutputPath | Should -Be $cleaned
        $report.Changed | Should -BeTrue
        (Get-OfficeHiddenData -LiteralPath $cleaned).Findings.Count | Should -Be 0
        (Get-FileHash -LiteralPath $proposal).Hash | Should -Be $hash

        $written = (Get-Item -LiteralPath $cleaned).LastWriteTimeUtc
        Remove-OfficeHiddenData -Path $proposal -NoClobber -ErrorAction SilentlyContinue -ErrorVariable refused | Should -BeNullOrEmpty
        $refused[0].FullyQualifiedErrorId | Should -BeLike 'hidden-data-output-rejected*'
        (Get-Item -LiteralPath $cleaned).LastWriteTimeUtc | Should -Be $written
        (Remove-OfficeHiddenData -Path $proposal).OutputPath | Should -Be $cleaned
    }

    It 'removes only the categories New-OfficeHiddenDataOptions selects' {
        $proposal = Join-Path $TestDrive 'selected.docx'
        $copy = Join-Path $TestDrive 'selected-shared.docx'
        New-ReviewedProposal $proposal

        $report = Remove-OfficeHiddenData -Path $proposal -OutputPath $copy -Options (New-OfficeHiddenDataOptions -Kind Comments)

        @($report.Findings | Where-Object Kind -EQ 'Comments').Outcome | Should -Be 'Removed'
        $author = @($report.Findings | Where-Object { $_.Kind -eq 'DocumentProperties' -and $_.Value -eq 'Jane Reviewer' })
        $author.Outcome | Should -Be 'Retained'
        $author.Reason | Should -Be 'NotSelected'
        $remaining = (Get-OfficeHiddenData -LiteralPath $copy).Kinds.ToString() -split ', '
        $remaining | Should -Not -Contain 'Comments'
        $remaining | Should -Contain 'DocumentProperties'
    }
}
