function Get-ExcelBenchmarkRegionGroups {
    param([object[]] $Rows)

    foreach ($region in @('NA', 'EU', 'APAC', 'LATAM')) {
        [pscustomobject]@{
            Name = $region
            TableName = 'Data_' + $region
            Data = @($Rows | Where-Object { $_.Region -eq $region })
        }
    }
}

function Get-ExcelBenchmarkSummaryRows {
    param([int] $Rows)

    @(
        [pscustomobject]@{ Metric = 'Rows'; Formula = 'COUNTA(Data!A2:A{0})' -f ($Rows + 1); NumberFormat = '#,##0' }
        [pscustomobject]@{ Metric = 'Average score'; Formula = 'AVERAGE(Data!G2:G{0})' -f ($Rows + 1); NumberFormat = '#,##0.00' }
        [pscustomobject]@{ Metric = 'Tickets'; Formula = 'SUM(Data!I2:I{0})' -f ($Rows + 1); NumberFormat = '#,##0' }
        [pscustomobject]@{ Metric = 'Enabled'; Formula = 'COUNTIF(Data!E2:E{0},TRUE)' -f ($Rows + 1); NumberFormat = '#,##0' }
    )
}

function Get-ExcelBenchmarkAppendSplit {
    param([object[]] $Rows)

    $initialCount = [math]::Max(1, [math]::Floor($Rows.Count * 0.8))
    [pscustomobject]@{
        Initial = @($Rows | Select-Object -First $initialCount)
        Append = @($Rows | Select-Object -Skip $initialCount)
    }
}

function Get-ExcelBenchmarkSmallSheetGroups {
    param([object[]] $Rows, [int] $SheetCount = 20)

    $safeSheetCount = [math]::Max(1, [math]::Min($SheetCount, [math]::Max(1, $Rows.Count)))
    $rowsPerSheet = [math]::Max(1, [math]::Ceiling($Rows.Count / $safeSheetCount))
    for ($sheetIndex = 0; $sheetIndex -lt $safeSheetCount; $sheetIndex++) {
        $sheetRows = @($Rows | Select-Object -Skip ($sheetIndex * $rowsPerSheet) -First $rowsPerSheet)
        if ($sheetRows.Count -eq 0) { continue }
        [pscustomobject]@{
            Name = 'Sheet{0:00}' -f ($sheetIndex + 1)
            TableName = 'Data{0:00}' -f ($sheetIndex + 1)
            Data = $sheetRows
        }
    }
}

function Get-ExcelBenchmarkWorkbookMergeInput {
    param([object[]] $Rows, [string] $BasePath)

    $firstCount = [math]::Max(1, [math]::Floor($Rows.Count / 2))
    [pscustomobject]@{
        SourceA = [IO.Path]::Combine([IO.Path]::GetDirectoryName($BasePath), ([IO.Path]::GetFileNameWithoutExtension($BasePath) + '.source-a.xlsx'))
        SourceB = [IO.Path]::Combine([IO.Path]::GetDirectoryName($BasePath), ([IO.Path]::GetFileNameWithoutExtension($BasePath) + '.source-b.xlsx'))
        RowsA = @($Rows | Select-Object -First $firstCount)
        RowsB = @($Rows | Select-Object -Skip $firstCount)
    }
}

function Test-ExcelBenchmarkOutput {
    param([object] $Case, [object] $Run)

    if ($Case.OperationKey -in @('WriteCsv', 'WriteCsvGZip', 'CsvToExcel', 'WriteWorkbook')) {
        Assert-BenchmarkPath $Run.Path
    }

    $tabularReadOperations = @(
        'ReadFullSheet',
        'ReadRange',
        'ReadFullSheetDataTable',
        'ReadRangeDataTable',
        'ReadFullSheetDataReader',
        'ReadRangeDataReader',
        'ReadNoHeaderRange'
    )
    if ($Case.OperationKey -in $tabularReadOperations) {
        $expectedRows = if ($Case.OperationKey -eq 'ReadNoHeaderRange') {
            [int]$Run.ExpectedRows + 1
        } else {
            [int]$Run.ExpectedRows
        }
        Assert-BenchmarkValue ([int]$Run.ActualRows) $expectedRows -Message "Expected $expectedRows rows returned by '$($Case.OperationKey)'."
        if ($Case.OperationKey -ne 'ReadNoHeaderRange') {
            Assert-BenchmarkValue ([int]$Run.ActualFields) ([int]$Run.ColumnCount) -Message "Expected every field returned by '$($Case.OperationKey)'."
            Assert-BenchmarkValue ([int]$Run.FirstId) 1 -Message "Expected '$($Case.OperationKey)' to preserve the first row identifier."
            Assert-BenchmarkValue ([int]$Run.LastId) ([int]$Run.ExpectedRows) -Message "Expected '$($Case.OperationKey)' to preserve the last row identifier."
        }
    }

    if ($Case.OperationKey -eq 'ReadUsedRangeDataTable') {
        $expectedRows = [int]$Run.ExpectedRows
        Assert-BenchmarkValue ([int]$Run.ActualRows) $expectedRows -Message "Expected $expectedRows rows returned by '$($Case.OperationKey)'."
    }

    if ($Case.OperationKey -eq 'ReadTableMetadata') {
        Assert-BenchmarkValue ([int]$Run.ActualTableCount) 1 -Message 'Expected one workbook table metadata result.'
        Assert-BenchmarkValue (@($Run.ActualTableNames) -contains 'Data') $true -Message "Expected table metadata to include 'Data'."
    }

    if ($Case.OperationKey -eq 'ReadNamedRangeMetadata') {
        Assert-BenchmarkValue ([int]$Run.ActualNamedRangeCount) 1 -Message 'Expected one workbook named range metadata result.'
        Assert-BenchmarkValue (@($Run.ActualNamedRangeNames) -contains 'SalesData') $true -Message "Expected named range metadata to include 'SalesData'."
    }

    if ([bool]$Run.SkipWorkbookValidation) {
        return
    }

    if (-not [bool]$Case.ValidateWorkbook -or [string]$Case.FileExtension -ne '.xlsx') {
        return
    }

    $document = Get-OfficeExcel -Path $Run.Path -ReadOnly
    if ($document) {
        Close-OfficeExcel -Document $document
    }
    Test-ExcelBenchmarkOpenXml -Path $Run.Path

    if ([string]$Case.Scenario -in @('objects-default', 'text-objects-default', 'wide-objects-default', 'datatable-default', 'datareader-table')) {
        Test-ExcelBenchmarkTabularValues -Case $Case -Run $Run
    }
}

function Test-ExcelBenchmarkTabularValues {
    param([object] $Case, [object] $Run)

    $actualRows = @(Import-OfficeExcel -Path $Run.Path -WorksheetName $Run.WorksheetName)
    $isDataTable = $Run.Payload -is [Data.DataTable]
    $expectedRows = if ($isDataTable) { @($Run.Payload.Rows) } else { @($Run.Payload) }
    Assert-BenchmarkValue $actualRows.Count $expectedRows.Count -Message "Expected '$($Case.Scenario)' to preserve every data row."
    if ($expectedRows.Count -eq 0) {
        return
    }

    $lastIndex = $expectedRows.Count - 1
    $middleIndex = [int] [Math]::Floor($lastIndex / 2)
    $indexes = @(0, $middleIndex, $lastIndex) | Select-Object -Unique
    foreach ($index in $indexes) {
        $expected = $expectedRows[$index]
        $actual = $actualRows[$index]
        $expectedValues = if ($isDataTable) {
            foreach ($column in $Run.Payload.Columns) {
                [pscustomobject]@{ Name = $column.ColumnName; Value = $expected[$column] }
            }
        } else {
            $expected.PSObject.Properties
        }
        foreach ($property in $expectedValues) {
            $expectedValue = ConvertTo-ExcelBenchmarkComparableValue -Value $property.Value
            $actualProperty = $actual.PSObject.Properties[$property.Name]
            if ($null -eq $actualProperty) {
                throw "Expected '$($Case.Scenario)' row $index to contain column '$($property.Name)'."
            }

            $actualValue = ConvertTo-ExcelBenchmarkComparableValue -Value $actualProperty.Value
            Assert-BenchmarkValue $actualValue $expectedValue -Message "Expected '$($Case.Scenario)' row $index column '$($property.Name)' to preserve its value."
        }
    }
}

function ConvertTo-ExcelBenchmarkComparableValue {
    param([AllowNull()][object] $Value)

    if ($null -eq $Value -or $Value -is [DBNull]) {
        return '<null>'
    }
    if ($Value -is [datetime]) {
        return $Value.ToString('O', [Globalization.CultureInfo]::InvariantCulture)
    }
    if ($Value -is [bool]) {
        return $Value.ToString().ToLowerInvariant()
    }
    if ($Value -is [decimal] -or $Value -is [double] -or $Value -is [single]) {
        return ([decimal]::Round([Convert]::ToDecimal($Value, [Globalization.CultureInfo]::InvariantCulture), 10)).ToString('G29', [Globalization.CultureInfo]::InvariantCulture)
    }
    if ($Value -is [IFormattable]) {
        return $Value.ToString($null, [Globalization.CultureInfo]::InvariantCulture)
    }

    return [string] $Value
}

function Test-CsvBenchmarkOutput {
    param([object] $Case, [object] $Run)

    $expectedRows = [int]$Run.ExpectedRows
    if ($Case.OperationKey -in @('ReadCsvSource', 'ReadCsvDataTable', 'ReadCsvGZipDataTable', 'ReadCsvQuickSingleColumn', 'ReadCsvQuickAllColumns')) {
        Assert-BenchmarkValue ([int]$Run.ActualRows) $expectedRows -Message "Expected $expectedRows rows returned by '$($Case.OperationKey)'."
        if ($Case.OperationKey -eq 'ReadCsvQuickSingleColumn') {
            Assert-BenchmarkValue ([int]$Run.AccessedFields) $expectedRows -Message "Expected $expectedRows first-column values accessed by '$($Case.OperationKey)'."
            Assert-BenchmarkValue ([string]$Run.LastValue) ([string]($expectedRows - 1)) -Message "Expected '$($Case.OperationKey)' to access the last Column0 value."
        }
        if ($Case.OperationKey -eq 'ReadCsvQuickAllColumns') {
            $expectedFields = $expectedRows * [int]$Run.ColumnCount
            Assert-BenchmarkValue ([int]$Run.AccessedFields) $expectedFields -Message "Expected $expectedFields values accessed by '$($Case.OperationKey)'."
            Assert-BenchmarkValue ([string]$Run.LastValue) ('Value{0}_{1}' -f ($expectedRows - 1), ([int]$Run.ColumnCount - 1)) -Message "Expected '$($Case.OperationKey)' to access the last field value."
        }
        $Run.RowsProcessed = [int]$Run.ActualRows
        return
    }

    $path = $Run.Path
    Assert-BenchmarkPath $path
    $actualRows = if ($Case.OperationKey -eq 'WriteCsvGZip') {
        $table = ConvertFrom-NativeGZipCsvToDataTable -Path $path
        if ($table -and $table.Rows) { @($table.Select()) } else { @() }
    } else {
        @(Import-Csv -Path $path)
    }
    Assert-BenchmarkValue $actualRows.Count $expectedRows -Message "Expected $expectedRows rows in '$path'."
    Test-CsvBenchmarkTabularValues -Case $Case -Run $Run -ActualRows $actualRows
    $Run.RowsProcessed = [int]$actualRows.Count
}

function Test-CsvBenchmarkTabularValues {
    param([object] $Case, [object] $Run, [object[]] $ActualRows)

    $expectedRows = if ($Run.Payload -is [Data.DataTable]) {
        @($Run.Payload.Select())
    } else {
        @($Run.Payload)
    }

    Assert-BenchmarkValue $ActualRows.Count $expectedRows.Count -Message "Expected '$($Case.Scenario)' to preserve every CSV data row."
    if ($expectedRows.Count -eq 0) {
        return
    }

    $expectedColumns = @(Get-CsvBenchmarkColumnNames -Row $expectedRows[0])
    $actualColumns = @(Get-CsvBenchmarkColumnNames -Row $ActualRows[0])
    Assert-BenchmarkValue ($actualColumns -join [char]31) ($expectedColumns -join [char]31) -Message "Expected '$($Case.Scenario)' to preserve the CSV header and column order."

    for ($rowIndex = 0; $rowIndex -lt $expectedRows.Count; $rowIndex++) {
        $expectedRow = $expectedRows[$rowIndex]
        $actualRow = $ActualRows[$rowIndex]
        foreach ($column in $expectedColumns) {
            $expectedValue = Get-CsvBenchmarkCellValue -Row $expectedRow -Column $column
            $actualValue = Get-CsvBenchmarkCellValue -Row $actualRow -Column $column
            if (-not (Test-CsvBenchmarkValueEquivalent -Expected $expectedValue -Actual $actualValue)) {
                throw "Expected '$($Case.Scenario)' row $rowIndex column '$column' to preserve value '$expectedValue'; actual value was '$actualValue'."
            }
        }
    }
}

function Get-CsvBenchmarkColumnNames {
    param([object] $Row)

    if ($Row -is [Data.DataRow]) {
        return @($Row.Table.Columns | ForEach-Object { [string]$_.ColumnName })
    }

    return @($Row.PSObject.Properties | ForEach-Object { [string]$_.Name })
}

function Get-CsvBenchmarkCellValue {
    param([object] $Row, [string] $Column)

    if ($Row -is [Data.DataRow]) {
        return $Row[$Column]
    }

    return $Row.PSObject.Properties[$Column].Value
}

function Test-CsvBenchmarkValueEquivalent {
    param([AllowNull()][object] $Expected, [AllowNull()][object] $Actual)

    if ($Expected -is [Management.Automation.PSObject]) {
        $Expected = $Expected.PSObject.BaseObject
    }
    if ($Actual -is [Management.Automation.PSObject]) {
        $Actual = $Actual.PSObject.BaseObject
    }

    if ($null -eq $Expected -or $Expected -is [DBNull]) {
        return $null -eq $Actual -or $Actual -is [DBNull] -or [string]::IsNullOrEmpty([string]$Actual)
    }

    if ($Expected -is [string]) {
        return ([string]$Actual) -ceq $Expected
    }

    $actualText = [string]$Actual
    if ($Expected -is [bool]) {
        $parsed = $false
        return [bool]::TryParse($actualText, [ref]$parsed) -and $parsed -eq $Expected
    }

    if ($Expected -is [guid]) {
        $parsed = [guid]::Empty
        return [guid]::TryParse($actualText, [ref]$parsed) -and $parsed -eq $Expected
    }

    $currentCulture = [Globalization.CultureInfo]::CurrentCulture
    $invariantCulture = [Globalization.CultureInfo]::InvariantCulture
    $cultures = if ($currentCulture.Name -eq $invariantCulture.Name) {
        @($invariantCulture)
    } else {
        @($currentCulture, $invariantCulture)
    }

    if ($Expected -is [datetime]) {
        foreach ($culture in $cultures) {
            $parsed = [datetime]::MinValue
            if ([datetime]::TryParse($actualText, $culture, [Globalization.DateTimeStyles]::AllowWhiteSpaces, [ref]$parsed) -and
                $parsed.Ticks -eq $Expected.Ticks) {
                return $true
            }
        }
        return $false
    }

    if ($Expected -is [datetimeoffset]) {
        foreach ($culture in $cultures) {
            $parsed = [datetimeoffset]::MinValue
            if ([datetimeoffset]::TryParse($actualText, $culture, [Globalization.DateTimeStyles]::AllowWhiteSpaces, [ref]$parsed) -and
                $parsed.UtcTicks -eq $Expected.UtcTicks) {
                return $true
            }
        }
        return $false
    }

    if ($Expected -is [timespan]) {
        foreach ($culture in $cultures) {
            $parsed = [timespan]::Zero
            if ([timespan]::TryParse($actualText, $culture, [ref]$parsed) -and $parsed -eq $Expected) {
                return $true
            }
        }
        return $false
    }

    if ($Expected -is [System.Byte] -or $Expected -is [System.SByte] -or
        $Expected -is [System.Int16] -or $Expected -is [System.UInt16] -or
        $Expected -is [System.Int32] -or $Expected -is [System.UInt32] -or
        $Expected -is [System.Int64] -or $Expected -is [System.UInt64]) {
        foreach ($culture in $cultures) {
            $parsed = [decimal]::Zero
            if ([decimal]::TryParse($actualText, [Globalization.NumberStyles]::Integer, $culture, [ref]$parsed) -and
                $parsed -eq [decimal]$Expected) {
                return $true
            }
        }
        return $false
    }

    if ($Expected -is [decimal]) {
        foreach ($culture in $cultures) {
            $parsed = [decimal]::Zero
            if ([decimal]::TryParse($actualText, [Globalization.NumberStyles]::Number, $culture, [ref]$parsed) -and
                $parsed -eq $Expected) {
                return $true
            }
        }
        return $false
    }

    if ($Expected -is [double] -or $Expected -is [single]) {
        $expectedNumber = [double]$Expected
        $tolerance = [math]::Max(1e-10, [math]::Abs($expectedNumber) * 1e-12)
        foreach ($culture in $cultures) {
            $parsed = 0.0
            $style = [Globalization.NumberStyles]::Float -bor [Globalization.NumberStyles]::AllowThousands
            if ([double]::TryParse($actualText, $style, $culture, [ref]$parsed) -and
                [math]::Abs($parsed - $expectedNumber) -le $tolerance) {
                return $true
            }
        }
        return $false
    }

    $expectedText = [string]$Expected
    return $actualText -ceq $expectedText
}

function Test-ExcelBenchmarkOpenXml {
    param([string] $Path)

    $validatorType = 'DocumentFormat.OpenXml.Validation.OpenXmlValidator' -as [type]
    $spreadsheetType = 'DocumentFormat.OpenXml.Packaging.SpreadsheetDocument' -as [type]
    if ($null -eq $validatorType -or $null -eq $spreadsheetType) {
        return
    }

    $document = $spreadsheetType::Open($Path, $false)
    try {
        $validator = [Activator]::CreateInstance($validatorType)
        $errors = @($validator.Validate($document))
        if ($errors.Count -gt 0) {
            $details = @(
                $errors |
                    Select-Object -First 5 |
                    ForEach-Object {
                        $part = if ($_.Part -and $_.Part.Uri) { [string]$_.Part.Uri } else { 'unknown part' }
                        '{0}: {1}' -f $part, $_.Description
                    }
            ) -join '; '

            throw "Expected '$Path' to pass OpenXML validation. First errors: $details"
        }
    } finally {
        if ($document) {
            $document.Dispose()
        }
    }
}
