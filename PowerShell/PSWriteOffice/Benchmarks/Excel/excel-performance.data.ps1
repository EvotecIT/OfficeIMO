function Get-ExcelBenchmarkData {
    param(
        [string] $Profile,
        [int] $Count
    )

    switch ($Profile) {
        MixedObjects {
            [pscustomobject]@{ Data = @(New-ExcelBenchmarkRows -Count $Count); ColumnCount = 10; WorksheetName = 'Data' }
        }
        TextObjects {
            [pscustomobject]@{ Data = @(New-ExcelBenchmarkTextRows -Count $Count); ColumnCount = 10; WorksheetName = 'Data' }
        }
        WideObjects {
            [pscustomobject]@{ Data = @(New-ExcelBenchmarkWideRows -Count $Count); ColumnCount = 40; WorksheetName = 'Data' }
        }
        CsvQuotedObjects {
            [pscustomobject]@{ Data = @(New-ExcelBenchmarkCsvQuotedRows -Count $Count); ColumnCount = 10; WorksheetName = 'Data' }
        }
        CsvMultilineObjects {
            [pscustomobject]@{ Data = @(New-ExcelBenchmarkCsvMultilineRows -Count $Count); ColumnCount = 10; WorksheetName = 'Data' }
        }
        CsvWideObjects {
            [pscustomobject]@{ Data = @(New-ExcelBenchmarkWideRows -Count $Count); ColumnCount = 40; WorksheetName = 'Data' }
        }
        DbatoolsQuickCsv {
            [pscustomobject]@{ Data = $null; ColumnCount = 10; WorksheetName = 'Data' }
        }
        DbatoolsQuotedCsv {
            [pscustomobject]@{ Data = $null; ColumnCount = 10; WorksheetName = 'Data' }
        }
        DbatoolsWideCsv {
            [pscustomobject]@{ Data = $null; ColumnCount = 50; WorksheetName = 'Data' }
        }
        DataTable {
            [pscustomobject]@{ Data = New-ExcelBenchmarkDataTable -Count $Count; ColumnCount = 6; WorksheetName = 'Data' }
        }
        DataSet {
            [pscustomobject]@{ Data = New-ExcelBenchmarkDataSet -Count $Count; ColumnCount = 6; WorksheetName = 'Sales' }
        }
        default { throw "Unknown benchmark profile '$Profile'." }
    }
}

function Get-ExcelBenchmarkPayloadCount {
    param([object] $Value)

    if ($Value -is [Data.DataSet]) {
        $count = 0
        foreach ($table in $Value.Tables) {
            $count += $table.Rows.Count
        }
        return $count
    }

    if ($Value -is [Data.DataTable]) {
        return $Value.Rows.Count
    }

    return @($Value).Count
}

function New-ExcelBenchmarkRows {
    param([int] $Count)

    for ($i = 1; $i -le $Count; $i++) {
        [pscustomobject]@{
            Id = $i
            Name = 'Server-{0:000000}' -f $i
            Department = 'Department-{0}' -f ($i % 25)
            Region = @('NA', 'EU', 'APAC', 'LATAM')[$i % 4]
            IsEnabled = ($i % 3) -ne 0
            Created = ([datetime]'2024-01-01').AddMinutes($i)
            Score = [math]::Round(($i * 1.137) % 1000, 3)
            Owner = 'owner{0}@example.test' -f ($i % 250)
            TicketCount = $i % 17
            Notes = 'Benchmark row {0}' -f $i
        }
    }
}

function New-ExcelBenchmarkTextRows {
    param([int] $Count)

    for ($i = 1; $i -le $Count; $i++) {
        [pscustomobject]@{
            Id = $i.ToString([Globalization.CultureInfo]::InvariantCulture)
            Name = 'Server-{0:000000}' -f $i
            Department = 'Department-{0}' -f ($i % 25)
            Region = @('NA', 'EU', 'APAC', 'LATAM')[$i % 4]
            IsEnabled = if (($i % 3) -ne 0) { 'true' } else { 'false' }
            Created = ([datetime]'2024-01-01').AddMinutes($i).ToString('O', [Globalization.CultureInfo]::InvariantCulture)
            Score = ([math]::Round(($i * 1.137) % 1000, 3)).ToString([Globalization.CultureInfo]::InvariantCulture)
            Owner = 'owner{0}@example.test' -f ($i % 250)
            TicketCount = ($i % 17).ToString([Globalization.CultureInfo]::InvariantCulture)
            Notes = 'Benchmark row {0}' -f $i
        }
    }
}

function New-ExcelBenchmarkWideRows {
    param([int] $Count)

    for ($i = 1; $i -le $Count; $i++) {
        $row = [ordered]@{
            Id = $i
            Name = 'Wide-{0:000000}' -f $i
            Created = ([datetime]'2024-01-01').AddSeconds($i)
            Enabled = ($i % 2) -eq 0
        }
        for ($column = 1; $column -le 36; $column++) {
            $row["Metric$column"] = [math]::Round((($i + $column) * 1.017) % 10000, 4)
        }
        [pscustomobject]$row
    }
}

function New-ExcelBenchmarkCsvQuotedRows {
    param([int] $Count)

    for ($i = 1; $i -le $Count; $i++) {
        [pscustomobject]@{
            Id = $i
            Name = 'Server, "{0:000000}"' -f $i
            Department = 'Department "{0}"' -f ($i % 25)
            Region = @('NA, East', 'EU "West"', 'APAC', 'LATAM')[$i % 4]
            IsEnabled = ($i % 3) -ne 0
            Created = ([datetime]'2024-01-01').AddMinutes($i)
            Score = [math]::Round(($i * 1.137) % 1000, 3)
            Owner = 'owner{0}@example.test' -f ($i % 250)
            TicketCount = $i % 17
            Notes = 'Quoted, comma, and ""escape"" row {0}' -f $i
        }
    }
}

function New-ExcelBenchmarkCsvMultilineRows {
    param([int] $Count)

    for ($i = 1; $i -le $Count; $i++) {
        [pscustomobject]@{
            Id = $i
            Name = 'Server-{0:000000}' -f $i
            Department = 'Department-{0}' -f ($i % 25)
            Region = @('NA', 'EU', 'APAC', 'LATAM')[$i % 4]
            IsEnabled = ($i % 3) -ne 0
            Created = ([datetime]'2024-01-01').AddMinutes($i)
            Score = [math]::Round(($i * 1.137) % 1000, 3)
            Owner = 'owner{0}@example.test' -f ($i % 250)
            TicketCount = $i % 17
            Notes = "Line 1 for row $i`r`nLine 2 with comma, quote "" and row $i"
        }
    }
}

function New-ExcelBenchmarkDbatoolsCsvSource {
    param(
        [Parameter(Mandatory)]
        [string] $Path,

        [Parameter(Mandatory)]
        [int] $Count,

        [int] $ColumnCount = 10,

        [switch] $QuoteAll
    )

    $directory = [IO.Path]::GetDirectoryName($Path)
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not [IO.Directory]::Exists($directory)) {
        [IO.Directory]::CreateDirectory($directory) | Out-Null
    }

    $encoding = [Text.UTF8Encoding]::new($false)
    $writer = [IO.StreamWriter]::new($Path, $false, $encoding)
    try {
        $headers = @(
            for ($column = 0; $column -lt $ColumnCount; $column++) {
                'Column{0}' -f $column
            }
        )
        $writer.WriteLine(($headers -join ','))

        $random = [Random]::new(42)
        $builder = [Text.StringBuilder]::new()
        for ($row = 0; $row -lt $Count; $row++) {
            $null = $builder.Clear()
            for ($column = 0; $column -lt $ColumnCount; $column++) {
                if ($column -gt 0) {
                    $null = $builder.Append(',')
                }

                $value = switch ($column) {
                    0 { $row.ToString([Globalization.CultureInfo]::InvariantCulture); break }
                    1 { 'Name{0}' -f $row; break }
                    2 { $random.Next(1, 100).ToString([Globalization.CultureInfo]::InvariantCulture); break }
                    3 { $random.NextDouble().ToString('F4', [Globalization.CultureInfo]::InvariantCulture); break }
                    4 { [datetime]::Now.AddDays(-$random.Next(365)).ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture); break }
                    5 { if ($random.Next(0, 2) -eq 0) { 'true' } else { 'false' }; break }
                    default { 'Value{0}_{1}' -f $row, $column; break }
                }

                if ($QuoteAll.IsPresent) {
                    $null = $builder.Append('"')
                    $null = $builder.Append($value)
                    $null = $builder.Append('"')
                } else {
                    $null = $builder.Append($value)
                }
            }

            $writer.WriteLine($builder.ToString())
        }
    } finally {
        $writer.Dispose()
    }
}

function New-ExcelBenchmarkDataTable {
    param([int] $Count)

    $table = [Data.DataTable]::new('Data')
    $null = $table.Columns.Add('Id', [int])
    $null = $table.Columns.Add('Name', [string])
    $null = $table.Columns.Add('Created', [datetime])
    $null = $table.Columns.Add('Amount', [decimal])
    $null = $table.Columns.Add('Enabled', [bool])
    $null = $table.Columns.Add('Notes', [string])
    for ($i = 1; $i -le $Count; $i++) {
        $row = $table.NewRow()
        $row.Id = $i
        $row.Name = 'Account-{0:000000}' -f $i
        $row.Created = ([datetime]'2024-01-01').AddMinutes($i)
        $row.Amount = [decimal]([math]::Round(($i * 11.317) % 100000, 2))
        $row.Enabled = ($i % 4) -ne 0
        $row.Notes = 'DataTable row {0}' -f $i
        $table.Rows.Add($row)
    }
    , $table
}

function New-ExcelBenchmarkDataSet {
    param([int] $Count)

    $dataSet = [Data.DataSet]::new('Report')
    $sales = New-ExcelBenchmarkDataTable -Count $Count
    $sales.TableName = 'Sales'
    $inventory = [Data.DataTable]::new('Inventory')
    $null = $inventory.Columns.Add('Sku', [string])
    $null = $inventory.Columns.Add('Quantity', [int])
    $null = $inventory.Columns.Add('Updated', [datetime])
    for ($i = 1; $i -le ([math]::Max(1, [math]::Floor($Count / 4))); $i++) {
        $row = $inventory.NewRow()
        $row.Sku = 'SKU-{0:000000}' -f $i
        $row.Quantity = $i % 500
        $row.Updated = ([datetime]'2024-01-01').AddHours($i)
        $inventory.Rows.Add($row)
    }
    $dataSet.Tables.Add($sales)
    $dataSet.Tables.Add($inventory)
    , $dataSet
}

function Get-ExcelBenchmarkColumnName {
    param([int] $ColumnNumber)

    $name = ''
    $value = $ColumnNumber
    while ($value -gt 0) {
        $value--
        $name = [char][int](65 + ($value % 26)) + $name
        $value = [int][math]::Floor($value / 26)
    }
    $name
}

function Get-ExcelBenchmarkColumnCount {
    param([string] $Profile)

    switch ($Profile) {
        WideObjects { 40 }
        CsvWideObjects { 40 }
        DbatoolsQuickCsv { 10 }
        DbatoolsQuotedCsv { 10 }
        DbatoolsWideCsv { 50 }
        DataTable { 6 }
        DataSet { 6 }
        default { 10 }
    }
}

function Get-ExcelBenchmarkRange {
    param([int] $ColumnCount, [int] $Rows)

    'A1:{0}{1}' -f (Get-ExcelBenchmarkColumnName -ColumnNumber $ColumnCount), ($Rows + 1)
}

function Get-ExcelBenchmarkRangeEndCell {
    param([int] $ColumnCount, [int] $Rows)

    '{0}{1}' -f (Get-ExcelBenchmarkColumnName -ColumnNumber $ColumnCount), ($Rows + 1)
}

function Get-ExcelBenchmarkExtension {
    param([object] $Case)

    $extension = [string]$Case.FileExtension
    if ([string]::IsNullOrWhiteSpace($extension)) { return '.xlsx' }
    if ($extension.StartsWith('.')) { return $extension }
    ".$extension"
}

function Initialize-ExcelBenchmarkInput {
    param([object] $Case, [object] $Run)

    foreach ($path in @($Run.Path, $Run.SourcePath)) {
        if ($path -and (Test-Path -LiteralPath $path)) {
            Remove-Item -LiteralPath $path -Force
        }
    }

    switch ([string]$Case.OperationKey) {
        ReadCsvSource {
            $Run.Payload | Export-Csv -Path $Run.SourcePath -NoTypeInformation -Encoding utf8 -UseQuotes AsNeeded
        }
        ReadCsvDataTable {
            $Run.Payload | Export-Csv -Path $Run.SourcePath -NoTypeInformation -Encoding utf8 -UseQuotes AsNeeded
        }
        ReadCsvGZipDataTable {
            Write-NativeGZipCsv -InputObject $Run.Payload -Path $Run.SourcePath
        }
        ReadCsvQuickSingleColumn {
            New-ExcelBenchmarkDbatoolsCsvSource -Path $Run.SourcePath -Count ([int]$Case.RowCount) -ColumnCount $Run.ColumnCount -QuoteAll:([string]$Case.DataProfile -eq 'DbatoolsQuotedCsv')
        }
        ReadCsvQuickAllColumns {
            New-ExcelBenchmarkDbatoolsCsvSource -Path $Run.SourcePath -Count ([int]$Case.RowCount) -ColumnCount $Run.ColumnCount -QuoteAll:([string]$Case.DataProfile -eq 'DbatoolsQuotedCsv')
        }
        CsvToExcel {
            $Run.Payload | Export-Csv -Path $Run.SourcePath -NoTypeInformation -Encoding utf8 -UseQuotes AsNeeded
        }
        ReadFullSheet { New-ExcelBenchmarkDefaultWorkbook -Engine PSWriteOffice -Run $Run }
        ReadRange { New-ExcelBenchmarkDefaultWorkbook -Engine PSWriteOffice -Run $Run }
        ReadFullSheetDataTable { New-ExcelBenchmarkDefaultWorkbook -Engine PSWriteOffice -Run $Run }
        ReadRangeDataTable { New-ExcelBenchmarkDefaultWorkbook -Engine PSWriteOffice -Run $Run }
        ReadFullSheetDataReader { New-ExcelBenchmarkDefaultWorkbook -Engine PSWriteOffice -Run $Run }
        ReadRangeDataReader { New-ExcelBenchmarkDefaultWorkbook -Engine PSWriteOffice -Run $Run }
        ReadNoHeaderRange { New-ExcelBenchmarkDefaultWorkbook -Engine PSWriteOffice -Run $Run }
        ReadUsedRangeDataTable { New-ExcelBenchmarkDefaultWorkbook -Engine PSWriteOffice -Run $Run }
        ReadTableMetadata { New-ExcelBenchmarkDefaultWorkbook -Engine PSWriteOffice -Run $Run -TableName Data }
        ReadNamedRangeMetadata { New-ExcelBenchmarkNamedRangeWorkbook -Engine PSWriteOffice -Run $Run }
    }
}

function New-ExcelBenchmarkDefaultWorkbook {
    param(
        [string] $Engine,
        [object] $Run,
        [string] $TableName
    )

    switch ($Engine) {
        PSWriteOffice {
            if ($TableName) {
                $Run.Payload | Export-OfficeExcel -Path $Run.Path -WorksheetName $Run.WorksheetName -TableName $TableName
            } else {
                $Run.Payload | Export-OfficeExcel -Path $Run.Path -WorksheetName $Run.WorksheetName
            }
        }
        ImportExcel {
            if ($TableName) {
                $Run.Payload | Export-Excel -Path $Run.Path -WorksheetName $Run.WorksheetName -TableName $TableName -AutoFilter
            } else {
                $Run.Payload | Export-Excel -Path $Run.Path -WorksheetName $Run.WorksheetName
            }
        }
        ExcelFast {
            Export-Workbook -Destination $Run.Path -InputObject $Run.Payload -SheetName $Run.WorksheetName -Force
        }
        default {
            throw "Engine '$Engine' cannot create workbook input."
        }
    }
}

function New-ExcelBenchmarkNamedRangeWorkbook {
    param([string] $Engine, [object] $Run)

    switch ($Engine) {
        PSWriteOffice {
            New-OfficeExcel -Path $Run.Path {
                Add-OfficeExcelSheet -Name $Run.WorksheetName -Content {
                    Add-OfficeExcelTable -Data $Run.Payload -TableName Data
                    Set-OfficeExcelNamedRange -Name SalesData -Range $Run.Range
                }
            } | Out-Null
        }
        ImportExcel {
            $Run.Payload | Export-Excel -Path $Run.Path -WorksheetName $Run.WorksheetName -TableName Data -AutoFilter -RangeName SalesData
        }
    }
}
