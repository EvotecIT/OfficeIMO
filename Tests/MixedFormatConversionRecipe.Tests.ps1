BeforeAll {
    $repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).ProviderPath
    $recipe = Join-Path $repoRoot 'Examples/Workflows/Example-MixedFormat-ConversionReview.ps1'
    $sampleRoot = Join-Path $repoRoot 'Examples/Samples/ConversionReview'
    $moduleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) { $env:PSWRITEOFFICE_MODULE_MANIFEST } else { Join-Path $repoRoot 'PSWriteOffice.psd1' }
    Import-Module $moduleManifest -Global -ErrorAction Stop
}

Describe 'Mixed-format conversion review recipe' -Skip:($PSVersionTable.PSEdition -ne 'Core') {
    It 'records native outputs and reuses verified PDFs without replacing the previous review' {
        $destination = Join-Path $TestDrive 'normal'
        $first = & $recipe -OutputDirectory $destination -ModulePath $moduleManifest
        $first.Selected | Should -Be 3
        $first.Completed | Should -Be 3
        $first.Reused | Should -Be 0
        $first.ReviewRequired | Should -Be 0
        $hashes = @(Get-ChildItem $first.PdfDirectory -File | Sort-Object Name | Get-FileHash | Select-Object -ExpandProperty Hash)
        $firstReview = Join-Path $first.ReviewDirectory 'review.json'
        $reviewHash = (Get-FileHash $firstReview).Hash
        $second = & $recipe -OutputDirectory $destination -ModulePath $moduleManifest
        $second.Completed | Should -Be 3
        $second.Reused | Should -Be 3
        $second.ReviewDirectory | Should -Not -Be $first.ReviewDirectory
        @(Get-ChildItem $second.PdfDirectory -File | Sort-Object Name | Get-FileHash | Select-Object -ExpandProperty Hash) | Should -Be $hashes
        (Get-FileHash $firstReview).Hash | Should -Be $reviewHash
        $review = Get-Content (Join-Path $second.ReviewDirectory 'review.json') -Raw | ConvertFrom-Json
        $review.ModuleVersion | Should -Be '3.0.9'
        @($review.Outcomes).Count | Should -Be 3
        @($review.Outcomes | Where-Object { -not $_.Reused }).Count | Should -Be 0
        @($review.AvailablePdfRoutes | Where-Object Id -EQ 'markdown-pdf').Count | Should -Be 1
    }

    It 'retains failed and skipped outcomes and retries a corrected source without replacing completed PDFs' {
        $source = Join-Path $TestDrive 'mixed-input'
        $destination = Join-Path $TestDrive 'mixed-output'
        $null = New-Item -ItemType Directory -Path $source
        Copy-Item (Join-Path $sampleRoot '*') $source
        $broken = Join-Path $source 'Correctable.docx'
        'Not an Open XML package' | Set-Content $broken
        'Unsupported example file' | Set-Content (Join-Path $source 'Unsupported.dat')
        { & $recipe -InputDirectory $source -OutputDirectory $destination -ModulePath $moduleManifest } | Should -Throw '*Conversion needs attention*'
        $firstReview = Get-ChildItem (Join-Path $destination 'Reports') -Recurse -Filter review.json | Select-Object -First 1
        $firstReviewHash = (Get-FileHash $firstReview.FullName).Hash
        $first = Get-Content $firstReview.FullName -Raw | ConvertFrom-Json
        $first.Summary.Completed | Should -Be 3
        $first.Summary.Failed | Should -Be 1
        $first.Summary.Skipped | Should -Be 1
        @($first.Outcomes | Where-Object { $_.Status -eq 'Failed' -and $_.ReviewRequired }).Count | Should -Be 1
        @($first.Outcomes | Where-Object { $_.Status -eq 'Skipped' -and $_.ReviewRequired }).Count | Should -Be 1
        WordNew -Path $broken { WordParagraph -Text 'Corrected synthetic handover attachment.' }
        $retry = & $recipe -InputDirectory $source -OutputDirectory $destination -ModulePath $moduleManifest -RetryFailed
        $retry.Completed | Should -Be 4
        $retry.Reused | Should -Be 3
        $retry.Failed | Should -Be 0
        $retry.Skipped | Should -Be 1
        $retry.ReviewRequired | Should -BeGreaterThan 0
        $retryReview = Get-Content (Join-Path $retry.ReviewDirectory 'review.json') -Raw | ConvertFrom-Json
        $warningOutcomes = @($retryReview.Outcomes | Where-Object { @($_.Diagnostics | Where-Object Severity -EQ 'Warning').Count })
        $warningOutcomes.Count | Should -BeGreaterThan 0
        @($warningOutcomes | Where-Object { -not $_.ReviewRequired }).Count | Should -Be 0
        $pdf = Join-Path $retry.PdfDirectory 'Correctable.docx.pdf'
        Add-Content $pdf 'Changed output'
        $changedHash = (Get-FileHash $pdf).Hash
        { & $recipe -InputDirectory $source -OutputDirectory $destination -ModulePath $moduleManifest -RetryFailed } | Should -Throw '*Conversion needs attention*'
        (Get-FileHash $pdf).Hash | Should -Be $changedHash
        (Get-FileHash $firstReview.FullName).Hash | Should -Be $firstReviewHash
    }

    It 'writes a review for an empty folder and reports that no PDFs were completed' {
        $source = Join-Path $TestDrive 'empty-input'
        $destination = Join-Path $TestDrive 'empty-output'
        $null = New-Item -ItemType Directory -Path $source
        { & $recipe -InputDirectory $source -OutputDirectory $destination -ModulePath $moduleManifest } | Should -Throw '*Conversion needs attention*'
        $reviewPath = Get-ChildItem (Join-Path $destination 'Reports') -Recurse -Filter review.json | Select-Object -First 1
        $review = Get-Content $reviewPath.FullName -Raw | ConvertFrom-Json
        $review.Summary.Selected | Should -Be 0
        $review.Summary.Completed | Should -Be 0
        @($review.Outcomes).Count | Should -Be 0
        (Get-Content (Join-Path $reviewPath.DirectoryName 'outcomes.csv') -TotalCount 1) | Should -Match 'InputPath.*OutputPath.*Status'
    }

    It 'passes the source byte limit to the engine and retains oversized-file failures' {
        $destination = Join-Path $TestDrive 'limited-output'
        { & $recipe -OutputDirectory $destination -ModulePath $moduleManifest -MaximumInputBytes 16 } | Should -Throw '*Conversion needs attention*'
        $reviewPath = Get-ChildItem (Join-Path $destination 'Reports') -Recurse -Filter review.json | Select-Object -First 1
        $review = Get-Content $reviewPath.FullName -Raw | ConvertFrom-Json
        $review.Summary.Completed | Should -Be 0
        $review.Summary.Failed | Should -Be 3
        $review.Limits.InputBytes | Should -Be 16
        @($review.Outcomes | Where-Object { $_.Status -eq 'Failed' -and $_.ReviewRequired }).Count | Should -Be 3
        @(Get-ChildItem (Join-Path $destination 'Pdf') -Recurse -File -Filter '*.pdf' -ErrorAction SilentlyContinue).Count | Should -Be 0
    }
}
