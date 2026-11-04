BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        $sourceRoot = Join-Path (Join-Path (Join-Path $PSScriptRoot '..') 'Sources') 'PSWriteOffice'

        if (-not $env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES) {
            $env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES = 'true'
        }

        if (-not $env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION) {
            $releasePath = Join-Path (Join-Path $sourceRoot 'bin') 'Release'
            $env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION = if (Test-Path $releasePath) { 'Release' } else { 'Debug' }
        }

        Join-Path (Join-Path $PSScriptRoot '..') 'PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
}

Describe 'Excel IDataReader import' {
    It 'returns an IDataReader that remains usable after the cmdlet returns' {
        $xlsx = Join-Path $TestDrive 'excel-reader.xlsx'

        @(
            [pscustomobject]@{ Id = 1; Name = 'Alpha' }
            [pscustomobject]@{ Id = 2; Name = 'Beta' }
            [pscustomobject]@{ Id = 3; Name = 'Gamma' }
        ) | Export-OfficeExcel -Path $xlsx -Sheet Data -TableName Data

        $reader = Import-OfficeExcel -Path $xlsx -Sheet Data -As DataReader -ChunkRows 1 -SchemaSampleRows 2
        try {
            $reader | Should -BeOfType ([System.Data.IDataReader])
            $reader.FieldCount | Should -Be 2
            $reader.GetName(0) | Should -Be 'Id'
            $reader.GetName(1) | Should -Be 'Name'
            $reader.GetFieldType(0) | Should -Be ([double])
            $reader.GetFieldType(1) | Should -Be ([string])

            $reader.Read() | Should -BeTrue
            [double] $reader.GetValue(0) | Should -Be 1
            $reader.GetString(1) | Should -Be 'Alpha'

            $reader.Read() | Should -BeTrue
            [double] $reader.GetValue(0) | Should -Be 2
            $reader.GetString(1) | Should -Be 'Beta'

            $reader.Read() | Should -BeTrue
            [double] $reader.GetValue(0) | Should -Be 3
            $reader.GetString(1) | Should -Be 'Gamma'
            $reader.Read() | Should -BeFalse
        } finally {
            if ($reader -is [System.IDisposable]) {
                $reader.Dispose()
            }
        }
    }

    It 'closes the owned workbook when the IDataReader reaches EOF' {
        $xlsx = Join-Path $TestDrive 'excel-reader-eof-close.xlsx'

        [pscustomobject]@{ Id = 1; Name = 'Alpha' } |
            Export-OfficeExcel -Path $xlsx -Sheet Data -TableName Data

        $reader = Import-OfficeExcel -Path $xlsx -Sheet Data -As DataReader

        $rowCount = 0
        while ($reader.Read()) {
            $rowCount++
        }

        $rowCount | Should -BeGreaterOrEqual 1
        $reader.IsClosed | Should -BeTrue
        { Remove-Item -LiteralPath $xlsx -Force -ErrorAction Stop } | Should -Not -Throw
    }

    It 'rejects selecting a range and a named table together' {
        $xlsx = Join-Path $TestDrive 'excel-reader-conflict.xlsx'

        [pscustomobject]@{ Id = 1; Name = 'Alpha' } |
            Export-OfficeExcel -Path $xlsx -Sheet Data -TableName Data

        { Import-OfficeExcel -Path $xlsx -Sheet Data -Range A1:B2 -Table Data -ErrorAction Stop } |
            Should -Throw
    }
}
