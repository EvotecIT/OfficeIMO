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
    if ($Case.OperationKey -eq 'ConvertToCsvText') {
        $lines = @($Run.CsvLines)
        Assert-BenchmarkValue $lines.Count ($expectedRows + 1) -Message "Expected one CSV string per record from '$($Case.OperationKey)'."
        $actualRows = @($lines | ConvertFrom-Csv)
        Test-CsvBenchmarkTabularValues -Case $Case -Run $Run -ActualRows $actualRows
        $Run.RowsProcessed = $actualRows.Count
        return
    }
    if ($Case.OperationKey -eq 'ConvertFromCsvText') {
        Assert-BenchmarkValue ([int]$Run.ActualRows) $expectedRows -Message "Expected $expectedRows rows returned by '$($Case.OperationKey)'."
        Test-CsvBenchmarkTabularValues -Case $Case -Run $Run -ActualRows @($Run.ConvertedRows)
        $Run.RowsProcessed = [int]$Run.ActualRows
        return
    }
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

    # Every cell is still compared on every iteration; the comparison runs compiled because a scripted loop
    # over 10,000-row outputs took seconds per iteration.
    $mismatch = [PSWriteOfficeBenchmarks.CsvValueComparer]::FindMismatch([object[]]$expectedRows, [object[]]$ActualRows, [string[]]$expectedColumns)
    if ($null -ne $mismatch) {
        throw "Expected '$($Case.Scenario)' $mismatch"
    }
}

function Get-CsvBenchmarkColumnNames {
    param([object] $Row)

    if ($Row -is [Data.DataRow]) {
        return @($Row.Table.Columns | ForEach-Object { [string]$_.ColumnName })
    }

    return @($Row.PSObject.Properties | ForEach-Object { [string]$_.Name })
}


function Test-CsvBenchmarkValueEquivalent {
    param([AllowNull()][object] $Expected, [AllowNull()][object] $Actual)

    [PSWriteOfficeBenchmarks.CsvValueComparer]::AreEquivalent($Expected, $Actual)
}

if (-not ('PSWriteOfficeBenchmarks.CsvValueComparer' -as [type])) {
    # C# 5 syntax so Windows PowerShell 5.1 can compile it.
    # PowerShell 7 references its runtime by default; Windows PowerShell needs System.Data added.
    $comparerReferences = @{}
    if ($PSVersionTable.PSEdition -eq 'Desktop') {
        $comparerReferences.ReferencedAssemblies = @('System.Data', 'System.Xml')
    }
    Add-Type @comparerReferences -TypeDefinition @'
using System;
using System.Data;
using System.Globalization;
using System.Management.Automation;

namespace PSWriteOfficeBenchmarks
{
    public static class CsvValueComparer
    {
        public static string FindMismatch(object[] expectedRows, object[] actualRows, string[] columns)
        {
            for (int rowIndex = 0; rowIndex < expectedRows.Length; rowIndex++)
            {
                object expectedRow = expectedRows[rowIndex];
                object actualRow = actualRows[rowIndex];
                foreach (string column in columns)
                {
                    object expected = GetCell(expectedRow, column);
                    object actual = GetCell(actualRow, column);
                    if (!AreEquivalent(expected, actual))
                    {
                        return string.Format(CultureInfo.InvariantCulture, "row {0} column '{1}' to preserve value '{2}'; actual value was '{3}'.", rowIndex, column, ToText(expected), ToText(actual));
                    }
                }
            }

            return null;
        }

        private static object GetCell(object row, string column)
        {
            DataRow dataRow = Unwrap(row) as DataRow;
            if (dataRow != null)
            {
                return dataRow[column];
            }

            PSPropertyInfo property = PSObject.AsPSObject(row).Properties[column];
            if (property == null)
            {
                throw new InvalidOperationException("Row has no column '" + column + "'.");
            }

            return property.Value;
        }

        private static object Unwrap(object value)
        {
            PSObject wrapped = value as PSObject;
            return wrapped != null ? wrapped.BaseObject : value;
        }

        private static string ToText(object value)
        {
            return value == null ? string.Empty : LanguagePrimitives.ConvertTo<string>(value);
        }

        public static bool AreEquivalent(object expected, object actual)
        {
            expected = Unwrap(expected);
            actual = Unwrap(actual);

            if (expected == null || expected is DBNull)
            {
                return actual == null || actual is DBNull || string.IsNullOrEmpty(ToText(actual));
            }

            if (expected is string)
            {
                return string.Equals(ToText(actual), (string)expected, StringComparison.Ordinal);
            }

            string actualText = ToText(actual);
            if (expected is bool)
            {
                bool parsedBool;
                return bool.TryParse(actualText, out parsedBool) && parsedBool == (bool)expected;
            }

            if (expected is Guid)
            {
                Guid parsedGuid;
                return Guid.TryParse(actualText, out parsedGuid) && parsedGuid == (Guid)expected;
            }

            CultureInfo current = CultureInfo.CurrentCulture;
            CultureInfo invariant = CultureInfo.InvariantCulture;
            CultureInfo[] cultures = current.Name == invariant.Name ? new[] { invariant } : new[] { current, invariant };

            if (expected is DateTime)
            {
                foreach (CultureInfo culture in cultures)
                {
                    DateTime parsed;
                    if (DateTime.TryParse(actualText, culture, DateTimeStyles.AllowWhiteSpaces, out parsed) && parsed.Ticks == ((DateTime)expected).Ticks)
                    {
                        return true;
                    }
                }

                return false;
            }

            if (expected is DateTimeOffset)
            {
                foreach (CultureInfo culture in cultures)
                {
                    DateTimeOffset parsed;
                    if (DateTimeOffset.TryParse(actualText, culture, DateTimeStyles.AllowWhiteSpaces, out parsed) && parsed.UtcTicks == ((DateTimeOffset)expected).UtcTicks)
                    {
                        return true;
                    }
                }

                return false;
            }

            if (expected is TimeSpan)
            {
                foreach (CultureInfo culture in cultures)
                {
                    TimeSpan parsed;
                    if (TimeSpan.TryParse(actualText, culture, out parsed) && parsed == (TimeSpan)expected)
                    {
                        return true;
                    }
                }

                return false;
            }

            if (expected is byte || expected is sbyte || expected is short || expected is ushort ||
                expected is int || expected is uint || expected is long || expected is ulong)
            {
                decimal expectedInteger = Convert.ToDecimal(expected, CultureInfo.InvariantCulture);
                foreach (CultureInfo culture in cultures)
                {
                    decimal parsed;
                    if (decimal.TryParse(actualText, NumberStyles.Integer, culture, out parsed) && parsed == expectedInteger)
                    {
                        return true;
                    }
                }

                return false;
            }

            if (expected is decimal)
            {
                foreach (CultureInfo culture in cultures)
                {
                    decimal parsed;
                    if (decimal.TryParse(actualText, NumberStyles.Number, culture, out parsed) && parsed == (decimal)expected)
                    {
                        return true;
                    }
                }

                return false;
            }

            if (expected is double || expected is float)
            {
                double expectedNumber = Convert.ToDouble(expected, CultureInfo.InvariantCulture);
                double tolerance = Math.Max(1e-10, Math.Abs(expectedNumber) * 1e-12);
                foreach (CultureInfo culture in cultures)
                {
                    double parsed;
                    if (double.TryParse(actualText, NumberStyles.Float | NumberStyles.AllowThousands, culture, out parsed) &&
                        Math.Abs(parsed - expectedNumber) <= tolerance)
                    {
                        return true;
                    }
                }

                return false;
            }

            return string.Equals(actualText, ToText(expected), StringComparison.Ordinal);
        }
    }
}
'@
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
