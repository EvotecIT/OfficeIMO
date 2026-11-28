function New-NotebookPerformanceFixture {
    param([string] $Root, [int] $Cells, [int] $Files, [switch] $Batch)

    $directory = Join-Path $Root "Notebook$Cells"
    [void][IO.Directory]::CreateDirectory($directory)
    $notebook = [ordered]@{
        nbformat = 4
        nbformat_minor = 5
        metadata = [ordered]@{
            kernelspec = [ordered]@{ name = 'python3'; display_name = 'Python 3'; language = 'python' }
        }
        cells = @(for ($index = 1; $index -le $Cells / 2; $index++) {
            $suffix = $index.ToString('D4', [CultureInfo]::InvariantCulture)
            [ordered]@{
                cell_type = 'markdown'
                id = "markdown-$suffix"
                metadata = @{}
                source = "# Notebook heading $suffix`n`n| Column |`n| --- |`n| Table witness $suffix |`n"
            }
            [ordered]@{
                cell_type = 'code'
                id = "code-$suffix"
                execution_count = $null
                metadata = @{}
                source = "print('Notebook source $suffix')`n"
                outputs = @([ordered]@{ output_type = 'stream'; name = 'stdout'; text = "Notebook result $suffix`n" })
            }
        })
    }

    $firstSource = $notebook.cells[0].source
    for ($index = 1; $index -le $Files; $index++) {
        if ($Batch) {
            $notebook.cells[0].source = "$firstSource`nBatch input witness $index.`n"
        }
        $json = $notebook | ConvertTo-Json -Depth 8 -Compress
        [IO.File]::WriteAllText((Join-Path $directory "notebook-$index.ipynb"), $json)
    }
}

function Assert-NotebookPerformanceCells {
    param($BenchmarkCase, $Cells)

    if ($Cells.Count -ne $BenchmarkCase.Cells) {
        throw "Expected $($BenchmarkCase.Cells) ordered cells and read $($Cells.Count)."
    }
    for ($index = 1; $index -le $BenchmarkCase.Cells / 2; $index++) {
        $suffix = $index.ToString('D4', [CultureInfo]::InvariantCulture)
        $prose = $Cells[($index - 1) * 2]
        $code = $Cells[($index - 1) * 2 + 1]
        if ($prose.Kind -ne 'markdown' -or $prose.Source -cne "# Notebook heading $suffix`n`n| Column |`n| --- |`n| Table witness $suffix |`n" -or
            $code.Kind -ne 'code' -or $code.Source -cne "print('Notebook source $suffix')`n") {
            throw "Incorrect cell kind, text, or order at pair $index."
        }
    }
}

function Assert-NotebookPerformanceArtifact {
    param([string] $Path, [int] $CellCount, [switch] $Html, [int] $BatchIndex)

    $text = [IO.File]::ReadAllText($Path)
    $tablePattern = if ($Html) { '<table(?:>|\s)' } else {
        '(?m)^\|[ \t]*Column[ \t]*\|[ \t]*\r?\n\|[ \t]*-{3,}[ \t]*\|[ \t]*\r?\n\|[ \t]*Table witness \d{4}[ \t]*\|'
    }
    if ([regex]::Matches($text, $tablePattern).Count -ne $CellCount / 2) {
        throw "Expected $($CellCount / 2) tables in '$Path'."
    }
    if ($Html) {
        $text = [Net.WebUtility]::HtmlDecode([regex]::Replace($text, '<[^>]+>', ''))
    }
    if ($BatchIndex -gt 0 -and -not $text.Contains("Batch input witness $BatchIndex.")) {
        throw "Missing the original batch input witness in '$Path'."
    }
    $offset = 0
    for ($index = 1; $index -le $CellCount / 2; $index++) {
        $suffix = $index.ToString('D4', [CultureInfo]::InvariantCulture)
        foreach ($witness in "Notebook heading $suffix", "Table witness $suffix", "print('Notebook source $suffix')", "Notebook result $suffix") {
            $position = $text.IndexOf($witness, $offset, [StringComparison]::Ordinal)
            if ($position -lt 0) {
                throw "Missing ordered witness '$witness' in '$Path'."
            }
            $offset = $position + $witness.Length
        }
    }
}
