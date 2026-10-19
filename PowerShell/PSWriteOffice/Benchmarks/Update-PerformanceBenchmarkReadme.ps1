param(
    [Parameter(Mandatory)]
    [string] $SummaryPath,

    [Parameter(Mandatory)]
    [string] $ReadmePath,

    [Parameter(Mandatory)]
    [string] $BlockId,

    [string] $BaselineEngine = 'PSWriteOffice'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function ConvertTo-BenchmarkDouble {
    param([object] $Value)

    [double]::Parse([string]$Value, [Globalization.CultureInfo]::InvariantCulture)
}

function Format-BenchmarkDuration {
    param([double] $Milliseconds)

    if ($Milliseconds -ge 1000) {
        return ([string]::Format([Globalization.CultureInfo]::InvariantCulture, '{0:n2} s', ($Milliseconds / 1000)))
    }

    [string]::Format([Globalization.CultureInfo]::InvariantCulture, '{0:n1} ms', $Milliseconds)
}

function Format-BenchmarkResult {
    param(
        [hashtable] $ByEngine,
        [string] $Baseline
    )

    $successful = @(
        foreach ($entry in $ByEngine.GetEnumerator()) {
            if ($entry.Value.Status -eq 'Succeeded') {
                [pscustomobject]@{ Engine = $entry.Key; MedianMs = ConvertTo-BenchmarkDouble -Value $entry.Value.MedianMs }
            }
        }
    ) | Sort-Object MedianMs, Engine

    if ($successful.Count -eq 0) {
        return 'No successful rows'
    }

    $winner = $successful[0]
    if ($successful.Count -gt 1) {
        $runnerUp = $successful[1]
        if (([double]$runnerUp.MedianMs / [double]$winner.MedianMs) -le 1.05) {
            return "$($winner.Engine) and $($runnerUp.Engine) within 5%"
        }
    }

    if ($winner.Engine -eq $Baseline) {
        return "$Baseline fastest"
    }

    if ($ByEngine.ContainsKey($Baseline) -and $ByEngine[$Baseline].Status -eq 'Succeeded') {
        $ratio = (ConvertTo-BenchmarkDouble -Value $ByEngine[$Baseline].MedianMs) / [double]$winner.MedianMs
        return ([string]::Format([Globalization.CultureInfo]::InvariantCulture, '{0} fastest; {1} {2:n2}x slower', $winner.Engine, $Baseline, $ratio))
    }

    "$($winner.Engine) fastest"
}

function Format-BenchmarkRatio {
    param(
        [double] $Milliseconds,
        [double] $BaselineMilliseconds,
        [string] $Engine,
        [string] $Baseline
    )

    if ($BaselineMilliseconds -le 0) {
        return ''
    }

    if ($Engine -eq $Baseline) {
        return '1.00x'
    }

    $ratio = $Milliseconds / $BaselineMilliseconds
    if ($ratio -ge 1) {
        return ([string]::Format([Globalization.CultureInfo]::InvariantCulture, '{0:n2}x slower', $ratio))
    }

    $speedup = $BaselineMilliseconds / $Milliseconds
    return ([string]::Format([Globalization.CultureInfo]::InvariantCulture, '{0:n2}x faster', $speedup))
}

function Format-BenchmarkCell {
    param(
        [object] $Row,
        [object] $BaselineRow,
        [string] $Engine,
        [string] $Baseline
    )

    if ($null -eq $Row) {
        return 'n/a'
    }

    if ($Row.Status -eq 'Skipped') {
        return 'Skipped'
    }

    if ($Row.Status -ne 'Succeeded') {
        return 'Failed'
    }

    $duration = Format-BenchmarkDuration -Milliseconds (ConvertTo-BenchmarkDouble -Value $Row.MedianMs)
    if ($null -eq $BaselineRow -or $BaselineRow.Status -ne 'Succeeded') {
        return $duration
    }

    $ratio = Format-BenchmarkRatio -Milliseconds (ConvertTo-BenchmarkDouble -Value $Row.MedianMs) -BaselineMilliseconds (ConvertTo-BenchmarkDouble -Value $BaselineRow.MedianMs) -Engine $Engine -Baseline $Baseline
    if ([string]::IsNullOrWhiteSpace($ratio)) {
        return $duration
    }

    "$duration ($ratio)"
}

function Update-MarkdownBlock {
    param(
        [string] $Path,
        [string] $Id,
        [string] $Content
    )

    $start = "<!-- BENCHMARK:$Id`:START -->"
    $end = "<!-- BENCHMARK:$Id`:END -->"
    $text = Get-Content -Raw -LiteralPath $Path
    $pattern = [regex]::Escape($start) + '(?s).*?' + [regex]::Escape($end)
    $replacement = $start + [Environment]::NewLine + $Content.TrimEnd() + [Environment]::NewLine + $end
    if ($text -notmatch $pattern) {
        throw "Benchmark block '$Id' was not found in '$Path'."
    }

    $updated = [regex]::Replace($text, $pattern, [System.Text.RegularExpressions.MatchEvaluator]{ param($match) $replacement })
    Set-Content -LiteralPath $Path -Value $updated -NoNewline
}

$summary = Import-Csv -LiteralPath $SummaryPath
$engines = @(
    $BaselineEngine
    $summary.Engine |
        Where-Object { $_ -ne $BaselineEngine } |
        Sort-Object -Unique
) | Select-Object -Unique

$markdown = [Text.StringBuilder]::new()
[void]$markdown.Append('| Scenario | Rows |')
foreach ($engine in $engines) {
    [void]$markdown.Append(' ').Append($engine).Append(' |')
}
[void]$markdown.AppendLine(' Result |')
[void]$markdown.Append('| --- | ---: |')
foreach ($engine in $engines) {
    [void]$markdown.Append(' ---: |')
}
[void]$markdown.AppendLine(' --- |')

$groups = $summary |
    Group-Object Scenario, RowCount |
    Sort-Object @{ Expression = { $_.Group[0].Scenario } }, @{ Expression = { [int]$_.Group[0].RowCount } }

foreach ($group in $groups) {
    $first = $group.Group[0]
    $byEngine = @{}
    foreach ($row in $group.Group) {
        $byEngine[$row.Engine] = $row
    }
    $baselineRow = if ($byEngine.ContainsKey($BaselineEngine)) { $byEngine[$BaselineEngine] } else { $null }

    [void]$markdown.Append('| ').Append($first.Scenario).Append(' | ').Append($first.RowCount).Append(' |')
    foreach ($engine in $engines) {
        $row = if ($byEngine.ContainsKey($engine)) { $byEngine[$engine] } else { $null }
        [void]$markdown.Append(' ').Append((Format-BenchmarkCell -Row $row -BaselineRow $baselineRow -Engine $engine -Baseline $BaselineEngine)).Append(' |')
    }
    [void]$markdown.Append(' ').Append((Format-BenchmarkResult -ByEngine $byEngine -Baseline $BaselineEngine)).AppendLine(' |')
}

Update-MarkdownBlock -Path $ReadmePath -Id $BlockId -Content $markdown.ToString()
