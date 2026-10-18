BeforeAll {
    $validationHelperPath = Join-Path $PSScriptRoot '..\Benchmarks\Benchmark.ResultValidation.ps1'

    . $validationHelperPath
}

Describe 'Benchmark result validation' {
    It 'normalizes every requested row count' {
        @(Resolve-BenchmarkRequestedRowCount -RowCount @('1000,5000', 10000) -Suite Smoke) |
            Should -Be @(1000, 5000, 10000)
    }

    It 'rejects a requested row count that is absent from the summary' {
        $expected = [pscustomobject]@{
            Engine = 'PSWriteOffice'
            Case = [pscustomobject]@{ Name = 'text-objects-default' }
        }
        $summary = [pscustomobject]@{
            Engine = 'PSWriteOffice'
            Scenario = 'text-objects-default'
            Variables = @{ RowCount = '1000' }
            Status = 'Succeeded'
        }

        {
            Assert-BenchmarkRequestedLaneCompletion `
                -Summary @($summary) `
                -ExpectedRun @($expected) `
                -RowCount @(1000, 5000)
        } | Should -Throw '*5000 rows*'
    }

    It 'accepts only when every requested row count succeeds' {
        $expected = [pscustomobject]@{
            Engine = 'PSWriteOffice'
            Case = [pscustomobject]@{ Name = 'text-objects-default' }
        }
        $summary = 1000, 5000 | ForEach-Object {
            [pscustomobject]@{
                Engine = 'PSWriteOffice'
                Scenario = 'text-objects-default'
                Variables = @{ RowCount = [string]$_ }
                Status = 'Succeeded'
            }
        }

        {
            Assert-BenchmarkRequestedLaneCompletion `
                -Summary @($summary) `
                -ExpectedRun @($expected) `
                -RowCount @(1000, 5000)
        } | Should -Not -Throw
    }
}
