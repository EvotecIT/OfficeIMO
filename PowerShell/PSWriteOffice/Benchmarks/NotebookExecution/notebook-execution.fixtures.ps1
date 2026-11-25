function Get-NotebookExecutionCases {
    param([string[]] $Scenario, [string] $FixtureRoot)

    foreach ($name in $Scenario) {
        $count = if ($name -eq 'State120') { 120 } else { 12 }
        $sources = @(for ($index = 1; $index -le $count; $index++) {
            if ($index -eq 1) {
                @'
$step = 0
function Get-NotebookProof {
    $script:step++
    'notebook-proof:{0:D3}:{1}' -f $script:step, ($script:step * 7)
}
Get-NotebookProof
'@
            } elseif ($index -eq $count -and $name -eq 'LargeText') {
                @'
$marker = Get-NotebookProof
$body = (1..512 | ForEach-Object { 'row:{0:D4}:{1}' -f $_, ('x' * 160) }) -join [char]10
$marker + [char]10 + $body
'@
            } else {
                'Get-NotebookProof'
            }
        })
        $expected = @(for ($index = 1; $index -le $count; $index++) {
            $marker = 'notebook-proof:{0:D3}:{1}' -f $index, ($index * 7)
            if ($index -eq $count -and $name -eq 'LargeText') {
                $rows = @(for ($row = 1; $row -le 512; $row++) {
                    'row:{0:D4}:{1}' -f $row, ('x' * 160)
                })
                (@($marker) + $rows) -join [char]10
            } else {
                $marker
            }
        })
        $expectedText = $expected -join [char]10
        $hash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes($expectedText)))
        [pscustomobject]@{
            Name = $name
            Count = $count
            InputPath = Join-Path $FixtureRoot "$name.ipynb"
            Sources = [string[]] $sources
            Expected = [string[]] $expected
            ExpectedSha256 = $hash
            ExpectedBytes = [Text.Encoding]::UTF8.GetByteCount($expectedText)
        }
    }
}

function Write-NotebookExecutionFixture {
    param($Fixture)

    $cells = @(
        [ordered]@{
            cell_type = 'markdown'
            id = 'introduction'
            metadata = @{}
            source = '# Persistent PowerShell session'
        }
        for ($index = 0; $index -lt $Fixture.Count; $index++) {
            [ordered]@{
                cell_type = 'code'
                id = "code-$index"
                metadata = @{
                    polyglot_notebook = @{ kernelName = 'pwsh' }
                }
                source = $Fixture.Sources[$index]
                execution_count = $null
                outputs = @()
            }
        }
    )
    $notebook = [ordered]@{
        nbformat = 4
        nbformat_minor = 5
        metadata = @{
            kernelspec = @{ name = '.net-powershell'; display_name = '.NET (PowerShell)'; language = 'PowerShell' }
            language_info = @{ name = 'PowerShell' }
        }
        cells = $cells
    }
    [IO.File]::WriteAllText($Fixture.InputPath, ($notebook | ConvertTo-Json -Depth 10), [Text.UTF8Encoding]::new($false))
}

function Read-NotebookExecutionOutput {
    param([string] $Path, [int] $CellCount)

    $document = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json -AsHashtable -Depth 32
    if ($document.nbformat -ne 4 -or $document.cells.Count -ne ($CellCount + 1) -or
        $document.cells[0].cell_type -ne 'markdown' -or
        (-join $document.cells[0].source).Trim() -ne '# Persistent PowerShell session') {
        throw 'The saved notebook structure or its introduction changed.'
    }
    for ($index = 1; $index -le $CellCount; $index++) {
        $cell = $document.cells[$index]
        if ($cell.cell_type -ne 'code') {
            throw "Expected a saved code cell at position $index."
        }
        $text = foreach ($output in $cell.outputs) {
            if ($output.output_type -eq 'stream' -and $output.name -eq 'stdout') {
                -join $output.text
            } elseif ($output.output_type -in 'display_data', 'execute_result' -and $output.data.Contains('text/plain')) {
                -join $output.data['text/plain']
            } else {
                throw "Unexpected output '$($output.output_type)/$($output.name)' in saved cell $index."
            }
        }
        -join $text
    }
}

function Assert-NotebookExecutionText {
    param($Case, [string[]] $Text)

    if ($Text.Count -ne $Case.Count) {
        throw "Expected $($Case.Count) text results and received $($Text.Count)."
    }
    $normalized = @(for ($index = 0; $index -lt $Case.Count; $index++) {
        $actual = $Text[$index].Replace([string] [char] 13 + [char] 10, [string] [char] 10).Trim([char] 10)
        if ($actual -cne $Case.Expected[$index]) {
            throw "Cell $index text did not match the expected state marker or payload."
        }
        $actual
    })
    $bytes = [Text.Encoding]::UTF8.GetBytes($normalized -join [char] 10)
    $hash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))
    if ($hash -ne $Case.ExpectedSha256) {
        throw 'The canonical notebook output checksum did not match.'
    }
}
