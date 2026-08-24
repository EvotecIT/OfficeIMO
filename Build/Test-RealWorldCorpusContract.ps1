[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path -Parent $PSScriptRoot
$project = Join-Path $repositoryRoot 'Build/RealWorldCorpus/OfficeIMO.RealWorldCorpus.Tool.csproj'
$scratch = Join-Path ([System.IO.Path]::GetTempPath()) ("officeimo-real-corpus-contract-" + [guid]::NewGuid().ToString('N'))
$inputDirectory = Join-Path $scratch 'input'
$policyInputDirectory = Join-Path $scratch 'policy-input'
$traversalInputDirectory = Join-Path $scratch 'traversal-input'
$firstJson = Join-Path $scratch 'first/report.json'
$firstMarkdown = Join-Path $scratch 'first/report.md'
$secondJson = Join-Path $scratch 'second/report.json'
$secondMarkdown = Join-Path $scratch 'second/report.md'

try {
    [void] (New-Item -ItemType Directory -Path $inputDirectory -Force)
    [void] (New-Item -ItemType Directory -Path $policyInputDirectory -Force)
    [void] (New-Item -ItemType Directory -Path (Join-Path $traversalInputDirectory 'a/b/c') -Force)
    [System.IO.File]::WriteAllText((Join-Path $inputDirectory 'renamed-html.blob'), '<!doctype html><html><body><h1>Content detected</h1></body></html>')
    [System.IO.File]::WriteAllText((Join-Path $inputDirectory 'ordinary.html'), '<!doctype html><html><body><p>Second document</p></body></html>')
    [System.IO.File]::WriteAllText((Join-Path $inputDirectory 'third.html'), '<!doctype html><html><body><p>Third document</p></body></html>')
    [System.IO.File]::WriteAllText((Join-Path $inputDirectory 'ordinary-copy.dat'), '<!doctype html><html><body><p>Second document</p></body></html>')
    [System.IO.File]::WriteAllText((Join-Path $inputDirectory 'sample.rtf'), '{\rtf1\ansi OfficeIMO corpus contract}')
    [System.IO.File]::WriteAllText((Join-Path $inputDirectory 'second.rtf'), '{\rtf1\ansi A second RTF document}')
    [System.IO.File]::WriteAllBytes((Join-Path $inputDirectory 'mislabeled.pdf'), [byte[]](1, 2, 3, 4, 5, 6, 7, 8))
    [System.IO.File]::WriteAllBytes((Join-Path $inputDirectory 'oversize.bin'), [byte[]]::new(4097))

    $deepRtf = '{\rtf1\ansi ' + ('{}' * 250001) + '}'
    $deepHtml = '<!doctype html><html><body>' + ('<b>' * 300) + 'limit' + ('</b>' * 300) + '</body></html>'
    [System.IO.File]::WriteAllText((Join-Path $policyInputDirectory 'deep.rtf'), $deepRtf)
    [System.IO.File]::WriteAllText((Join-Path $policyInputDirectory '![image](target).html'), $deepHtml)

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
    if ($markdown.Contains('![corpus](target)', [StringComparison]::Ordinal) -or $markdown.Contains('![source](target)', [StringComparison]::Ordinal)) {
        throw 'Provenance values were able to inject Markdown content into the evidence report.'
    }
    if (-not $markdown.Contains('&#33;&#91;corpus&#93;&#40;target&#41;', [StringComparison]::Ordinal) -or
        -not $markdown.Contains('&#33;&#91;source&#93;&#40;target&#41;', [StringComparison]::Ordinal) -or
        -not $markdown.Contains(('a' * 64), [StringComparison]::Ordinal)) {
        throw 'Inert provenance values were not retained in the Markdown evidence.'
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
