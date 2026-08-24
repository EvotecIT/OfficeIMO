[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path -Parent $PSScriptRoot
$project = Join-Path $repositoryRoot 'Build/RealWorldCorpus/OfficeIMO.RealWorldCorpus.Tool.csproj'
$scratch = Join-Path ([System.IO.Path]::GetTempPath()) ("officeimo-real-corpus-contract-" + [guid]::NewGuid().ToString('N'))
$inputDirectory = Join-Path $scratch 'input'
$formatInputDirectory = Join-Path $scratch 'all-formats-input'
$policyInputDirectory = Join-Path $scratch 'policy-input'
$packagePolicyInputDirectory = Join-Path $scratch 'package-policy-input'
$traversalInputDirectory = Join-Path $scratch 'traversal-input'
$firstJson = Join-Path $scratch 'first/report.json'
$firstMarkdown = Join-Path $scratch 'first/report.md'
$secondJson = Join-Path $scratch 'second/report.json'
$secondMarkdown = Join-Path $scratch 'second/report.md'

try {
    [void] (New-Item -ItemType Directory -Path $inputDirectory -Force)
    [void] (New-Item -ItemType Directory -Path $formatInputDirectory -Force)
    [void] (New-Item -ItemType Directory -Path $policyInputDirectory -Force)
    [void] (New-Item -ItemType Directory -Path $packagePolicyInputDirectory -Force)
    [void] (New-Item -ItemType Directory -Path (Join-Path $traversalInputDirectory 'a/b/c') -Force)
    [System.IO.File]::WriteAllText((Join-Path $inputDirectory 'renamed-html.blob'), '<!doctype html><html><body><h1>Content detected</h1></body></html>')
    [System.IO.File]::WriteAllText((Join-Path $inputDirectory 'ordinary.html'), '<!doctype html><html><body><p>Second document</p></body></html>')
    [System.IO.File]::WriteAllText((Join-Path $inputDirectory 'third.html'), '<!doctype html><html><body><p>Third document</p></body></html>')
    [System.IO.File]::WriteAllText((Join-Path $inputDirectory 'ordinary-copy.dat'), '<!doctype html><html><body><p>Second document</p></body></html>')
    [System.IO.File]::WriteAllText((Join-Path $inputDirectory 'sample.rtf'), '{\rtf1\ansi OfficeIMO corpus contract}')
    [System.IO.File]::WriteAllText((Join-Path $inputDirectory 'second.rtf'), '{\rtf1\ansi A second RTF document}')
    [System.IO.File]::WriteAllBytes((Join-Path $inputDirectory 'mislabeled.pdf'), [byte[]](1, 2, 3, 4, 5, 6, 7, 8))
    [System.IO.File]::WriteAllBytes((Join-Path $inputDirectory 'oversize.bin'), [byte[]]::new(4097))

    $formatFixtures = [ordered]@{
        'document.docx' = 'Website/Apps/OfficeIMO.Web.Converter/wwwroot/samples/basic.docx'
        'workbook.xlsx' = 'Website/Apps/OfficeIMO.Web.Converter/wwwroot/samples/basic.xlsx'
        'presentation.pptx' = 'Website/Apps/OfficeIMO.Web.Converter/wwwroot/samples/basic.pptx'
        'document.pdf' = 'Website/static/downloads/showcase/pdf/showcase-dashboard.pdf'
    }
    foreach ($fixture in $formatFixtures.GetEnumerator()) {
        [System.IO.File]::Copy(
            (Join-Path $repositoryRoot $fixture.Value),
            (Join-Path $formatInputDirectory $fixture.Key))
    }
    [System.IO.File]::WriteAllText((Join-Path $formatInputDirectory 'document.html'), '<!doctype html><html><body><p>HTML contract</p></body></html>')
    [System.IO.File]::WriteAllText((Join-Path $formatInputDirectory 'document.rtf'), '{\rtf1\ansi RTF contract}')

    $deepRtf = '{\rtf1\ansi ' + ('{}' * 250001) + '}'
    $deepHtml = '<!doctype html><html><body>' + ('<b>' * 300) + 'limit' + ('</b>' * 300) + '</body></html>'
    [System.IO.File]::WriteAllText((Join-Path $policyInputDirectory 'deep.rtf'), $deepRtf)
    [System.IO.File]::WriteAllText((Join-Path $policyInputDirectory '![image](target).html'), $deepHtml)

    $compressedPackage = Join-Path $packagePolicyInputDirectory 'compressed.docx'
    $packageStream = [System.IO.File]::Create($compressedPackage)
    try {
        $archive = [System.IO.Compression.ZipArchive]::new(
            $packageStream,
            [System.IO.Compression.ZipArchiveMode]::Create,
            $false)
        try {
            $contentTypes = $archive.CreateEntry('[Content_Types].xml')
            $writer = [System.IO.StreamWriter]::new($contentTypes.Open())
            try {
                $writer.Write('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
            } finally {
                $writer.Dispose()
            }
            $document = $archive.CreateEntry('word/document.xml', [System.IO.Compression.CompressionLevel]::SmallestSize)
            $writer = [System.IO.StreamWriter]::new($document.Open())
            try {
                $writer.Write('<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>')
                $writer.Write('A' * (2 * 1024 * 1024))
                $writer.Write('</w:t></w:r></w:p></w:body></w:document>')
            } finally {
                $writer.Dispose()
            }
        } finally {
            $archive.Dispose()
        }
    } finally {
        $packageStream.Dispose()
    }

    $common = @(
        'run', '--input', $inputDirectory,
        '--corpus-id', '![corpus](target)',
        '--source-uri', '![source](target)',
        '--archive-sha256', ('a' * 64),
        '--formats', 'Html,Rtf',
        '--max-per-format', '2',
        '--max-total', '3',
        '--max-file-bytes', '4096',
        '--max-traversal-entries', '20',
        '--timeout-seconds', '30',
        '--parallelism', '2'
    )

    & dotnet run --project $project --framework net10.0 --configuration Release -- @common --json $firstJson --markdown $firstMarkdown
    if ($LASTEXITCODE -ne 0) { throw "First corpus contract run failed with exit code $LASTEXITCODE." }
    & dotnet run --project $project --framework net10.0 --configuration Release --no-build -- @common --json $secondJson --markdown $secondMarkdown
    if ($LASTEXITCODE -ne 0) { throw "Second corpus contract run failed with exit code $LASTEXITCODE." }

    $firstJsonText = Get-Content -LiteralPath $firstJson -Raw
    $first = $firstJsonText | ConvertFrom-Json -Depth 100
    $second = Get-Content -LiteralPath $secondJson -Raw | ConvertFrom-Json -Depth 100
    if ($first.schemaVersion -ne 1 -or $first.measurementStatus -ne 'measured') { throw 'The report schema or measurement state is invalid.' }
    if ($first.configuration.packagePolicy.maxPackageBytes -ne 4096 -or
        $first.configuration.packagePolicy.maxPartCount -ne 4096 -or
        $first.configuration.packagePolicy.maxPartUncompressedBytes -ne 67108864 -or
        $first.configuration.packagePolicy.maxXmlCharactersInPart -ne 33554432 -or
        $first.configuration.packagePolicy.maxTotalUncompressedBytes -ne 268435456 -or
        $first.configuration.packagePolicy.maxCompressionRatio -ne 500) {
        throw 'The JSON evidence did not record the package policy applied by the worker.'
    }
    if ($first.totals.discovered -ne 8 -or $first.totals.oversize -ne 1) { throw 'The report did not account for every synthetic input.' }
    if ($first.totals.duplicateContent -ne 1) { throw 'Exact duplicate content was not identified.' }
    if ($first.totals.selected -ne 3) { throw 'The expected unique HTML and RTF sample was not selected.' }
    if (@($first.files | Where-Object { $_.outcome -eq 'not-selected' }).Count -ne 2) { throw 'The per-format and total sample caps were not exercised.' }
    if (@($first.files | Where-Object { $_.extension -eq '.blob' -and $_.detectedKind -eq 'html' }).Count -ne 1) { throw 'Content detection did not override the unknown extension.' }
    if (@($first.files | Where-Object { $_.extension -eq '.pdf' -and $_.outcome -eq 'not-eligible' -and $_.contentKind -eq 'unknown' }).Count -ne 1) { throw 'Extension-only evidence entered a content-detected stratum.' }
    if (@($first.files | Where-Object { $null -ne $_.sourceName }).Count -ne 0) { throw 'Source names leaked despite the default privacy mode.' }
    if (@($first.files | Where-Object { $_.PSObject.Properties.Name -contains 'fullPath' }).Count -ne 0 -or $firstJsonText.Contains($scratch, [StringComparison]::OrdinalIgnoreCase)) { throw 'A full input path leaked into the JSON evidence.' }
    if (@($first.files | Where-Object { $_.selected -and $_.outcome -notin @('completed', 'completed-with-warnings', 'completed-with-errors') }).Count -ne 0) { throw 'A selected synthetic document did not complete the normalized reader contract.' }

    $firstSelection = @($first.files | Where-Object selected | ForEach-Object sha256 | Sort-Object)
    $secondSelection = @($second.files | Where-Object selected | ForEach-Object sha256 | Sort-Object)
    if (Compare-Object $firstSelection $secondSelection) { throw 'Hash-stratified selection changed between identical runs.' }
    $markdown = Get-Content -LiteralPath $firstMarkdown -Raw
    if ($markdown.Contains($scratch, [StringComparison]::OrdinalIgnoreCase) -or $markdown.Contains($inputDirectory, [StringComparison]::OrdinalIgnoreCase)) { throw 'A full input path leaked into the Markdown evidence.' }
    if ($markdown -notmatch 'not a reliability guarantee' -or $markdown -notmatch 'does not assess visual fidelity') { throw 'The generated evidence omitted its interpretation boundaries.' }
    if ($markdown -notmatch '\| 4096 \| 20 \| 30 \| 2 \| 4096 \| 67108864 \| 33554432 \| 268435456 \| 500 \|') {
        throw 'The Markdown evidence did not record the package policy applied by the worker.'
    }
    if ($markdown.Contains('![corpus](target)', [StringComparison]::Ordinal) -or $markdown.Contains('![source](target)', [StringComparison]::Ordinal)) {
        throw 'Provenance values were able to inject Markdown content into the evidence report.'
    }
    if (-not $markdown.Contains('&#33;&#91;corpus&#93;&#40;target&#41;', [StringComparison]::Ordinal) -or
        -not $markdown.Contains('&#33;&#91;source&#93;&#40;target&#41;', [StringComparison]::Ordinal) -or
        -not $markdown.Contains(('a' * 64), [StringComparison]::Ordinal)) {
        throw 'Inert provenance values were not retained in the Markdown evidence.'
    }

    $formatJson = Join-Path $scratch 'all-formats/report.json'
    $formatMarkdown = Join-Path $scratch 'all-formats/report.md'
    & dotnet run --project $project --framework net10.0 --configuration Release --no-build -- run `
        --input $formatInputDirectory --json $formatJson --markdown $formatMarkdown `
        --corpus-id synthetic-all-formats --formats Word,Excel,PowerPoint,Pdf,Html,Rtf `
        --max-per-format 1 --max-total 6 --max-file-bytes 1048576 `
        --max-traversal-entries 10 --timeout-seconds 30 --parallelism 3
    if ($LASTEXITCODE -ne 0) { throw "All-format corpus contract run failed with exit code $LASTEXITCODE." }
    $formatReport = Get-Content -LiteralPath $formatJson -Raw | ConvertFrom-Json -Depth 100
    if ($formatReport.totals.discovered -ne 6 -or $formatReport.totals.eligibleUnique -ne 6 -or $formatReport.totals.selected -ne 6) {
        throw 'The all-format contract did not content-detect and select exactly one file per format.'
    }
    foreach ($kind in @('word', 'excel', 'powerPoint', 'pdf', 'html', 'rtf')) {
        $records = @($formatReport.files | Where-Object { $_.contentKind -eq $kind })
        if ($records.Count -ne 1 -or -not $records[0].selected -or
            $records[0].contentConfidence -notin @('medium', 'high') -or
            $records[0].outcome -notin @('completed', 'completed-with-warnings') -or
            $records[0].errorDiagnostics -ne 0) {
            $observed = $records | Select-Object contentKind, contentConfidence, selected, outcome, errorDiagnostics, exceptionType | ConvertTo-Json -Compress
            throw "The $kind normalized-reader contract was not reached successfully: $observed"
        }
    }

    $policyJson = Join-Path $scratch 'policy/report.json'
    $policyMarkdown = Join-Path $scratch 'policy/report.md'
    & dotnet run --project $project --framework net10.0 --configuration Release --no-build -- run `
        --input $policyInputDirectory --json $policyJson --markdown $policyMarkdown `
        --corpus-id synthetic-policy-contract --formats Html,Rtf --max-per-format 2 --max-total 2 `
        --max-file-bytes 1048576 --max-traversal-entries 10 --timeout-seconds 30 --parallelism 2 --include-source-names
    if ($LASTEXITCODE -ne 0) { throw "Policy corpus contract run failed with exit code $LASTEXITCODE." }
    $policy = Get-Content -LiteralPath $policyJson -Raw | ConvertFrom-Json -Depth 100
    if ($policy.totals.rejectedByPolicy -ne 2 -or $policy.totals.failed -ne 0) {
        $observed = @($policy.files | Where-Object selected | Select-Object contentKind, outcome, exceptionType) | ConvertTo-Json -Compress
        throw "HTML and RTF resource limits were not classified as policy rejections: $observed"
    }
    $policyMarkdownText = Get-Content -LiteralPath $policyMarkdown -Raw
    if ($policyMarkdownText.Contains('![image](target).html', [StringComparison]::Ordinal)) {
        throw 'A source name was able to inject Markdown content into the evidence report.'
    }
    if (-not $policyMarkdownText.Contains('&#33;&#91;image&#93;&#40;target&#41;&#46;html', [StringComparison]::Ordinal)) {
        throw 'The inert source name was not retained as readable Markdown evidence.'
    }

    $packagePolicyJson = Join-Path $scratch 'package-policy/report.json'
    $packagePolicyMarkdown = Join-Path $scratch 'package-policy/report.md'
    & dotnet run --project $project --framework net10.0 --configuration Release --no-build -- run `
        --input $packagePolicyInputDirectory --json $packagePolicyJson --markdown $packagePolicyMarkdown `
        --corpus-id synthetic-package-policy-contract --formats Word --max-per-format 1 --max-total 1 `
        --max-file-bytes 4194304 --max-traversal-entries 5 --timeout-seconds 30 --parallelism 1
    if ($LASTEXITCODE -ne 0) { throw "Package policy corpus contract run failed with exit code $LASTEXITCODE." }
    $packagePolicy = Get-Content -LiteralPath $packagePolicyJson -Raw | ConvertFrom-Json -Depth 100
    $packageRecord = @($packagePolicy.files | Where-Object selected)
    if ($packagePolicy.totals.rejectedByPolicy -ne 1 -or $packagePolicy.totals.failed -ne 0 -or
        $packageRecord.Count -ne 1 -or $packageRecord[0].exceptionType -ne 'OfficeIMO.OfficePackageSecurityException') {
        $observed = $packageRecord | Select-Object contentKind, contentConfidence, outcome, exceptionType | ConvertTo-Json -Compress
        throw "Compressed package expansion was not rejected by the shared package policy: $observed"
    }

    & dotnet run --project $project --framework net10.0 --configuration Release --no-build -- verify-markdown-contract
    if ($LASTEXITCODE -ne 0) { throw "Dynamic Markdown contract failed with exit code $LASTEXITCODE." }

    & dotnet run --project $project --framework net10.0 --configuration Release --no-build -- run `
        --input $inputDirectory --json (Join-Path $scratch 'unknown/report.json') --markdown (Join-Path $scratch 'unknown/report.md') `
        --max-totalx 1 *> $null
    if ($LASTEXITCODE -eq 0) { throw 'An unknown corpus option was silently accepted.' }

    & dotnet run --project $project --framework net10.0 --configuration Release --no-build -- run `
        --input $inputDirectory --json (Join-Path $scratch 'duplicate/report.json') --markdown (Join-Path $scratch 'duplicate/report.md') `
        --max-total 2 --max-total 3 *> $null
    if ($LASTEXITCODE -eq 0) { throw 'A duplicate corpus option was silently accepted.' }

    & dotnet run --project $project --framework net10.0 --configuration Release --no-build -- run `
        --input $inputDirectory --json (Join-Path $inputDirectory 'prior-report.json') --markdown (Join-Path $scratch 'inside-input/report.md') *> $null
    if ($LASTEXITCODE -eq 0) { throw 'A report path inside the corpus input was silently accepted.' }

    $hardLinkedReport = Join-Path $scratch 'hard-linked-report.json'
    [void] (New-Item -ItemType HardLink -Path $hardLinkedReport -Target (Join-Path $inputDirectory 'ordinary.html'))
    & dotnet run --project $project --framework net10.0 --configuration Release --no-build -- run `
        --input $inputDirectory --json $hardLinkedReport --markdown (Join-Path $scratch 'hard-linked-report.md') *> $null
    if ($LASTEXITCODE -eq 0) { throw 'A hard-linked report path bypassed input containment.' }

    $sharedReportTarget = Join-Path $scratch 'shared-report-target.txt'
    [System.IO.File]::WriteAllText($sharedReportTarget, 'existing report target')
    $sharedJsonAlias = Join-Path $scratch 'shared-report.json'
    $sharedMarkdownAlias = Join-Path $scratch 'shared-report.md'
    [void] (New-Item -ItemType HardLink -Path $sharedJsonAlias -Target $sharedReportTarget)
    [void] (New-Item -ItemType HardLink -Path $sharedMarkdownAlias -Target $sharedReportTarget)
    & dotnet run --project $project --framework net10.0 --configuration Release --no-build -- run `
        --input $inputDirectory --json $sharedJsonAlias --markdown $sharedMarkdownAlias *> $null
    if ($LASTEXITCODE -eq 0) { throw 'Hard-linked report aliases were accepted as distinct output files.' }

    $mutableInput = Join-Path $scratch 'mutable.html'
    [System.IO.File]::WriteAllText($mutableInput, '<!doctype html><html><body>first</body></html>')
    $classificationText = & dotnet run --project $project --framework net10.0 --configuration Release --no-build -- `
        classify-file --stage classification --input $mutableInput --max-file-bytes 4096
    if ($LASTEXITCODE -ne 0) { throw 'Direct classification worker contract failed.' }
    $classification = $classificationText | ConvertFrom-Json -Depth 100
    if (-not $classification.succeeded -or $classification.sha256.Length -ne 64) {
        throw 'Direct classification did not return a snapshot hash.'
    }
    [System.IO.File]::WriteAllText($mutableInput, '<!doctype html><html><body>second</body></html>')
    $probeText = & dotnet run --project $project --framework net10.0 --configuration Release --no-build -- `
        probe-file --stage probe --input $mutableInput --max-file-bytes 4096 --expected-sha256 $classification.sha256
    if ($LASTEXITCODE -ne 0) { throw 'Direct probe worker contract failed.' }
    $probe = $probeText | ConvertFrom-Json -Depth 100
    if ($probe.succeeded -or $probe.exceptionType -ne 'OfficeIMO.RealWorldCorpus.CorpusInputChangedException') {
        throw 'A changed input was parsed under the hash recorded during classification.'
    }

    $inputAlias = Join-Path $scratch 'input-alias'
    $insideReportDirectory = Join-Path $inputDirectory 'linked-reports'
    [void] (New-Item -ItemType Directory -Path $insideReportDirectory -Force)
    if ($IsWindows) {
        [void] (New-Item -ItemType Junction -Path $inputAlias -Target $inputDirectory)
    } else {
        [void] [System.IO.Directory]::CreateSymbolicLink($inputAlias, $inputDirectory)
    }
    & dotnet run --project $project --framework net10.0 --configuration Release --no-build -- run `
        --input $inputAlias --json (Join-Path $inputDirectory 'linked-input-report.json') `
        --markdown (Join-Path $scratch 'linked-input-report.md') *> $null
    if ($LASTEXITCODE -eq 0) { throw 'A linked input path bypassed report containment.' }

    $reportAlias = Join-Path $scratch 'report-alias'
    if ($IsWindows) {
        [void] (New-Item -ItemType Junction -Path $reportAlias -Target $insideReportDirectory)
    } else {
        [void] [System.IO.Directory]::CreateSymbolicLink($reportAlias, $insideReportDirectory)
    }
    & dotnet run --project $project --framework net10.0 --configuration Release --no-build -- run `
        --input $inputDirectory --json (Join-Path $reportAlias 'linked-report.json') `
        --markdown (Join-Path $scratch 'linked-report.md') *> $null
    if ($LASTEXITCODE -eq 0) { throw 'A linked report parent bypassed report containment.' }

    if ($IsWindows) {
        $insideInputReport = Join-Path $inputDirectory 'extended-report.json'
        $extendedInputReport = '\\?\' + $insideInputReport
        & dotnet run --project $project --framework net10.0 --configuration Release --no-build -- run `
            --input $inputDirectory --json $extendedInputReport `
            --markdown (Join-Path $scratch 'extended-report.md') *> $null
        if ($LASTEXITCODE -eq 0) { throw 'A Windows extended path bypassed report containment.' }

        $sameReport = Join-Path $scratch 'same-physical-report.json'
        $extendedSameReport = '\\?\' + $sameReport
        & dotnet run --project $project --framework net10.0 --configuration Release --no-build -- run `
            --input $inputDirectory --json $sameReport --markdown $extendedSameReport *> $null
        if ($LASTEXITCODE -eq 0) { throw 'Windows path aliases were accepted as distinct report files.' }

        $drive = $inputDirectory.Substring(0, 1).ToLowerInvariant()
        $uncInput = "\\localhost\$drive`$" + $inputDirectory.Substring(2)
        if (Test-Path -LiteralPath $uncInput) {
            & dotnet run --project $project --framework net10.0 --configuration Release --no-build -- run `
                --input $inputDirectory --json (Join-Path $uncInput 'unc-report.json') `
                --markdown (Join-Path $scratch 'unc-report.md') *> $null
            if ($LASTEXITCODE -eq 0) { throw 'A Windows local-to-UNC alias bypassed report containment.' }

            $uncSameReport = "\\localhost\$drive`$" + $sameReport.Substring(2)
            & dotnet run --project $project --framework net10.0 --configuration Release --no-build -- run `
                --input $inputDirectory --json $sameReport --markdown $uncSameReport *> $null
            if ($LASTEXITCODE -eq 0) { throw 'Windows local-to-UNC aliases were accepted as distinct report files.' }
        }
    }

    if (-not $IsWindows) {
        $danglingReportAlias = Join-Path $scratch 'dangling-report.json'
        [void] [System.IO.File]::CreateSymbolicLink(
            $danglingReportAlias,
            (Join-Path $inputDirectory 'not-yet-created-report.json'))
        & dotnet run --project $project --framework net10.0 --configuration Release --no-build -- run `
            --input $inputDirectory --json $danglingReportAlias `
            --markdown (Join-Path $scratch 'dangling-report.md') *> $null
        if ($LASTEXITCODE -eq 0) { throw 'A dangling report link bypassed report containment.' }
    }

    & dotnet run --project $project --framework net10.0 --configuration Release --no-build -- run `
        --input $traversalInputDirectory --json (Join-Path $scratch 'traversal/report.json') --markdown (Join-Path $scratch 'traversal/report.md') `
        --max-traversal-entries 2 *> $null
    if ($LASTEXITCODE -eq 0) { throw 'Empty directories were not counted against the traversal limit.' }

    $global:LASTEXITCODE = 0
    Write-Host 'Real-world corpus runner contract passed.'
} finally {
    if (Test-Path -LiteralPath $scratch) {
        Remove-Item -LiteralPath $scratch -Recurse -Force
    }
}
