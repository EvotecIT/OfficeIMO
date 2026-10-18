param(
    [string] $ModulePath = (Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'),
    [Parameter(Mandatory)]
    [string] $FixtureDirectory,
    [Parameter(Mandatory)]
    [string] $OutputDirectory,
    [string] $TesseractPath,
    [string] $TessdataDirectory
)

$ErrorActionPreference = 'Stop'
Import-Module $ModulePath -Force -ErrorAction Stop
$null = New-Item -ItemType Directory -Path $OutputDirectory -Force
$manifest = Get-Content -LiteralPath (Join-Path $FixtureDirectory 'manifest.json') -Raw | ConvertFrom-Json
$session = @{ NoLanguageDownload = $true }
if ($TesseractPath) { $providerVersion = (& $TesseractPath --version 2>&1 | Select-Object -First 1).ToString() }
if ($TesseractPath) { $session.TesseractPath = $TesseractPath }
if ($TessdataDirectory) { $session.TessdataDirectory = $TessdataDirectory }

function Assert-Equal($Expected, $Actual, [string] $Context) {
    if ($Expected -cne $Actual) { throw "$Context expected '$Expected', actual '$Actual'." }
}
function Normalize-Text([string] $Text) { [regex]::Replace($Text.Normalize(), '\s+', ' ').Trim() }

$results = foreach ($case in $manifest.cases) {
    $inputPath = Join-Path $FixtureDirectory ($case.id + '-scan.png')
    $sourceHash = (Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash.ToLowerInvariant()
    Assert-Equal $case.files.'scan.png' $sourceHash ($case.id + ' source hash')
    $review = Get-OfficeImageDocument -Path $inputPath @session
    Assert-Equal (Normalize-Text ($case.readingOrder -join ' ')) (Normalize-Text $review.Ocr.Text) ($case.id + ' recognized reading order')
    Assert-Equal ([int](@($case.table).Count -gt 0)) $review.Ocr.Document.Tables.Count ($case.id + ' table count')
    if (@($case.table).Count -gt 0) {
        $table = $review.Ocr.Document.Tables[0]
        Assert-Equal @($case.table).Count $table.Rows.Count ($case.id + ' table rows')
        for ($row = 0; $row -lt $table.Rows.Count; $row++) {
            Assert-Equal ($case.table[$row] -join '|') ($table.Rows[$row] -join '|') ($case.id + ' row ' + $row)
        }
    }
    $formats = @('docx', 'pdf', 'txt', 'md', 'html', 'json')
    if (@($case.table).Count -gt 0) { $formats += @('xlsx', 'csv') }
    foreach ($format in $formats) {
        $path = Join-Path $OutputDirectory ($case.id + '.' + $format)
        $settings = @{ InputObject = $review; OutputPath = $path; Force = $true; PassThru = $true }
        if ($format -eq 'xlsx') { $settings.FirstRowIsHeader = $true }
        $export = ConvertFrom-OfficeImage @settings
        if (-not $export.File.Exists -or $export.File.Length -eq 0) { throw "$path was not committed." }
        switch ($format) {
            'xlsx' {
                $excel = Get-OfficeExcel -Path $path -ReadOnly
                try {
                    Assert-Equal 1 $excel.Sheets.Count 'Excel sheet count'
                    $sheet = $excel.Sheets[0]
                    for ($row = 0; $row -lt @($case.table).Count; $row++) {
                        for ($column = 0; $column -lt @($case.table[$row]).Count; $column++) {
                            $cell = $sheet.CellAt($row + 1, $column + 1)
                            # Read the underlying typed cell value so string formatting cannot conceal a text-only workbook.
                            $value = $cell.GetValue().Value
                            if ($row -gt 0 -and $column -gt 0) {
                                $expected = [double]::Parse($case.table[$row][$column], [Globalization.CultureInfo]::InvariantCulture)
                                Assert-Equal $expected ([double] $value) "Excel row $row column $column"
                                if ($value -is [string]) { throw 'Numeric ledger cells were saved as text.' }
                            } else { Assert-Equal $case.table[$row][$column] ([string] $value) "Excel row $row column $column" }
                        }
                    }
                } finally { $excel.Dispose() }
            }
            'docx' {
                $word = Get-OfficeWord -Path $path -ReadOnly
                try {
                    Assert-Equal ([int](@($case.table).Count -gt 0)) $word.Tables.Count 'Word table count'
                    Assert-Equal 0 $word.Images.Count 'Word exports editable content without duplicating the source scan'
                    if ($word.Tables.Count -gt 0) {
                        Assert-Equal @($case.table).Count $word.Tables[0].Rows.Count 'Word table rows'
                        for ($row = 0; $row -lt @($case.table).Count; $row++) {
                            for ($column = 0; $column -lt @($case.table[$row]).Count; $column++) {
                                $text = ($word.Tables[0].Rows[$row].Cells[$column].Paragraphs | ForEach-Object Text) -join ' '
                                Assert-Equal $case.table[$row][$column] (Normalize-Text $text) "Word row $row column $column"
                            }
                        }
                    }
                } finally { $word.Dispose() }
            }
            'pdf' {
                $pdf = Get-OfficePdf -Path $path
                Assert-Equal (Normalize-Text $review.Ocr.Text) (Normalize-Text $pdf.Read().Text) 'Searchable PDF text'
            }
            'txt' { Assert-Equal (Normalize-Text $review.Ocr.Text) (Normalize-Text (Get-Content -LiteralPath $path -Raw -Encoding UTF8)) 'Text export' }
            'md' {
                $markdown = Get-Content -LiteralPath $path -Raw -Encoding UTF8
                if ($markdown -match '\[Image:') { throw 'Markdown duplicated the source image.' }
                if (@($case.table).Count -gt 0) {
                    Assert-Equal 1 ([regex]::Matches($markdown, 'Return credit').Count) 'Markdown table has no duplicate cell paragraphs'
                    if ($markdown -notmatch '\| Return credit \| 3 \|') { throw 'Markdown table was lost.' }
                } else {
                    $last = -1
                    foreach ($line in $case.readingOrder) {
                        $position = $markdown.IndexOf($line, [StringComparison]::Ordinal)
                        if ($position -le $last) { throw "Markdown reading order lost '$line'." }
                        $last = $position
                    }
                }
            }
            'html' {
                $html = Get-Content -LiteralPath $path -Raw -Encoding UTF8
                Assert-Equal ([int](@($case.table).Count -gt 0)) ([regex]::Matches($html, '<table>').Count) 'HTML table count'
                if ($html -match '<img\b') { throw 'HTML duplicated the source image.' }
                foreach ($line in $case.readingOrder) {
                    if (@($case.table).Count -eq 0 -and $html.IndexOf($line, [StringComparison]::Ordinal) -lt 0) { throw "HTML lost '$line'." }
                }
                foreach ($row in $case.table) {
                    $cells = ($row | ForEach-Object { '<td>' + [Net.WebUtility]::HtmlEncode($_) + '</td>' }) -join ''
                    if (-not $html.Contains($cells)) { throw "HTML lost table row '$cells'." }
                }
            }
            'json' {
                $json = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
                Assert-Equal (Normalize-Text $review.Ocr.Text) (Normalize-Text $json.pages[0].text) 'JSON reading order'
                Assert-Equal $review.Ocr.Document.Tables.Count @($json.pages[0].tables).Count 'JSON tables'
                if (@($case.table).Count -gt 0) {
                    for ($row = 0; $row -lt @($case.table).Count; $row++) {
                        Assert-Equal ($case.table[$row] -join '|') ($json.pages[0].tables[0].rows[$row] -join '|') "JSON row $row"
                    }
                }
            }
            'csv' {
                $csv = @(Import-Csv -LiteralPath $path)
                Assert-Equal (@($case.table).Count - 1) $csv.Count 'CSV body rows'
                Assert-Equal "'-13.50" $csv[1].Amount 'CSV negative source text is escaped'
            }
        }
    }
    Assert-Equal $sourceHash ((Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash.ToLowerInvariant()) 'Unchanged source image'
    [pscustomobject]@{ Case = $case.id; SourceSha256 = $sourceHash; Provider = $review.Ocr.Pages[0].Provider; Model = $review.Ocr.Pages[0].Model; Language = $review.Ocr.Pages[0].Language; AcceptedWords = $review.Ocr.AcceptedWordCount; Tables = $review.Ocr.Document.Tables.Count; Formats = $formats; Passed = $true }
}
[pscustomobject]@{ PowerShell = $PSVersionTable.PSVersion.ToString(); Edition = $PSVersionTable.PSEdition; ProviderVersion = $providerVersion; Cases = @($results) } |
    ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'qualification.json') -Encoding UTF8
$results | Format-Table Case, AcceptedWords, Tables, Passed
