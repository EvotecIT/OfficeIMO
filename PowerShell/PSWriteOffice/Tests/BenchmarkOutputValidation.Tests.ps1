BeforeAll {
    . (Join-Path $PSScriptRoot '..\Benchmarks\Excel\excel-performance.validation.ps1')
}

Describe 'Benchmark output validation' {
    It 'accepts invariant and current-culture numeric CSV representations' {
        Test-CsvBenchmarkValueEquivalent -Expected 1.137 -Actual '1.137' | Should -BeTrue

        $current = [Globalization.CultureInfo]::CurrentCulture
        $localized = ([double]1.137).ToString($null, $current)
        Test-CsvBenchmarkValueEquivalent -Expected 1.137 -Actual $localized | Should -BeTrue
    }

    It 'accepts typed booleans and dates without weakening string equality' {
        Test-CsvBenchmarkValueEquivalent -Expected $true -Actual 'True' | Should -BeTrue
        Test-CsvBenchmarkValueEquivalent -Expected ([datetime]'2024-01-01T00:01:00') -Actual '01/01/2024 00:01:00' | Should -BeTrue
        Test-CsvBenchmarkValueEquivalent -Expected "Line 1`r`nLine 2" -Actual "Line 1`r`nLine 2" | Should -BeTrue
        Test-CsvBenchmarkValueEquivalent -Expected 'Alpha' -Actual 'alpha' | Should -BeFalse
    }

    It 'rejects a parseable but incorrect numeric value' {
        Test-CsvBenchmarkValueEquivalent -Expected 1.137 -Actual '2.137' | Should -BeFalse
    }

    It 'compares scalar values carried through PowerShell object wrappers' {
        $wrapped = [Management.Automation.PSObject]::AsPSObject([int]1)

        Test-CsvBenchmarkValueEquivalent -Expected $wrapped -Actual '1' | Should -BeTrue
    }
}
