function Resolve-BenchmarkRequestedRowCount {
    [CmdletBinding()]
    param(
        [object[]] $RowCount,

        [Parameter(Mandatory)]
        [string] $Suite
    )

    $values = if ($RowCount -and $RowCount.Count -gt 0) {
        $RowCount
    } else {
        Get-ExcelBenchmarkDefaultRowCount -Suite $Suite
    }

    $resolved = foreach ($value in @($values)) {
        foreach ($part in ([string]$value -split ',')) {
            $text = $part.Trim()
            $parsed = 0
            if ([string]::IsNullOrWhiteSpace($text) -or
                -not [int]::TryParse(
                    $text,
                    [Globalization.NumberStyles]::Integer,
                    [Globalization.CultureInfo]::InvariantCulture,
                    [ref]$parsed)) {
                throw "RowCount must be an integer. Value: '$text'."
            }
            if ($parsed -le 0) {
                throw "RowCount must be greater than zero. Value: $parsed"
            }
            $parsed
        }
    }

    @($resolved | Select-Object -Unique)
}

function Assert-BenchmarkRequestedLaneCompletion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]] $Summary,

        [Parameter(Mandatory)]
        [object[]] $ExpectedRun,

        [Parameter(Mandatory)]
        [int[]] $RowCount
    )

    foreach ($expected in $ExpectedRun) {
        $scenario = [string]$expected.Case.Name
        foreach ($expectedRowCount in $RowCount) {
            $actualRows = @($Summary | Where-Object {
                $actualRowCount = if ($_.Variables -and $_.Variables.ContainsKey('RowCount')) {
                    [string]$_.Variables['RowCount']
                } else {
                    $null
                }

                $_.Engine -eq $expected.Engine -and
                $_.Scenario -eq $scenario -and
                $actualRowCount -eq [string]$expectedRowCount
            })

            if ($actualRows.Count -eq 0 -or
                @($actualRows | Where-Object Status -NE 'Succeeded').Count -gt 0) {
                throw "Requested benchmark lane '$($expected.Engine) / $scenario / $expectedRowCount rows' did not complete successfully."
            }
        }
    }
}
