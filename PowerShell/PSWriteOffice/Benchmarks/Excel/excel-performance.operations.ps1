function Invoke-ExcelBenchmarkOperation {
    param(
        [string] $Engine,
        [object] $Case,
        [object] $Run
    )

    switch ([string]$Case.OperationKey) {
        WriteCsv { Invoke-ExcelBenchmarkWriteCsv -Engine $Engine -Run $Run }
        WriteCsvGZip { Invoke-ExcelBenchmarkWriteCsvGZip -Engine $Engine -Run $Run }
        WriteCsvDataTable { Invoke-ExcelBenchmarkWriteCsvDataTable -Engine $Engine -Run $Run }
        ReadCsvSource { Invoke-ExcelBenchmarkReadCsv -Engine $Engine -Run $Run }
        ReadCsvDataTable { Invoke-ExcelBenchmarkReadCsvDataTable -Engine $Engine -Run $Run }
        ReadCsvGZipDataTable { Invoke-ExcelBenchmarkReadCsvGZipDataTable -Engine $Engine -Run $Run }
        ReadCsvQuickSingleColumn { Invoke-ExcelBenchmarkReadCsvQuickSingleColumn -Engine $Engine -Run $Run }
        ReadCsvQuickAllColumns { Invoke-ExcelBenchmarkReadCsvQuickAllColumns -Engine $Engine -Run $Run }
        CsvToExcel { Invoke-ExcelBenchmarkCsvToExcel -Engine $Engine -Run $Run }
        WriteWorkbook { Invoke-ExcelBenchmarkWriteWorkbook -Engine $Engine -Case $Case -Run $Run }
        ReadFullSheet { Invoke-ExcelBenchmarkReadWorkbook -Engine $Engine -Case $Case -Run $Run -Mode Full }
        ReadRange { Invoke-ExcelBenchmarkReadWorkbook -Engine $Engine -Case $Case -Run $Run -Mode Range }
        ReadFullSheetDataTable { Invoke-ExcelBenchmarkReadDataTable -Run $Run -Mode Full }
        ReadRangeDataTable { Invoke-ExcelBenchmarkReadDataTable -Run $Run -Mode Range }
        ReadFullSheetDataReader { Invoke-ExcelBenchmarkReadDataReader -Run $Run -Mode Full }
        ReadRangeDataReader { Invoke-ExcelBenchmarkReadDataReader -Run $Run -Mode Range }
        ReadNoHeaderRange { Invoke-ExcelBenchmarkReadWorkbook -Engine $Engine -Case $Case -Run $Run -Mode NoHeader }
        ReadUsedRangeDataTable {
            $dataTable = Get-OfficeExcelUsedRange -Path $Run.Path -Sheet $Run.WorksheetName -AsDataTable
            $Run.ActualRows = if ($dataTable -and $dataTable.Rows) { [int]$dataTable.Rows.Count } else { 0 }
        }
        ReadTableMetadata {
            $tables = @(Get-OfficeExcelTable -Path $Run.Path -Sheet $Run.WorksheetName)
            $Run.ActualTableCount = $tables.Count
            $Run.ActualTableNames = @($tables | ForEach-Object { $_.Name })
        }
        ReadNamedRangeMetadata {
            $ranges = @(Get-OfficeExcelNamedRange -Path $Run.Path -Sheet $Run.WorksheetName)
            $Run.ActualNamedRangeCount = $ranges.Count
            $Run.ActualNamedRangeNames = @($ranges | ForEach-Object { $_.Name })
        }
        default { throw "Unknown benchmark operation '$($Case.OperationKey)'." }
    }
}

function Invoke-ExcelBenchmarkWriteCsv {
    param([string] $Engine, [object] $Run)

    switch ($Engine) {
        PSWriteOffice { $Run.Payload | Export-OfficeCsv -Path $Run.Path }
        NativeCsv { $Run.Payload | Export-Csv -Path $Run.Path -NoTypeInformation -Encoding utf8 -UseQuotes AsNeeded }
        default { throw "Engine '$Engine' does not support CSV write." }
    }
}

function Invoke-ExcelBenchmarkWriteCsvGZip {
    param([string] $Engine, [object] $Run)

    switch ($Engine) {
        PSWriteOffice { $Run.Payload | Export-OfficeCsv -Path $Run.Path -CompressionType GZip }
        NativeCsv { Write-NativeGZipCsv -InputObject $Run.Payload -Path $Run.Path }
        default { throw "Engine '$Engine' does not support GZip CSV write." }
    }
}

function Invoke-ExcelBenchmarkWriteCsvDataTable {
    param([string] $Engine, [object] $Run)

    switch ($Engine) {
        PSWriteOffice { Export-OfficeCsv -InputObject $Run.Payload -Path $Run.Path }
        NativeCsv { $Run.Payload.Rows | Export-Csv -Path $Run.Path -NoTypeInformation -Encoding utf8 -UseQuotes AsNeeded }
        default { throw "Engine '$Engine' does not support DataTable CSV write." }
    }
}

function Invoke-ExcelBenchmarkReadCsv {
    param([string] $Engine, [object] $Run)

    $rows = switch ($Engine) {
        PSWriteOffice { @(Import-OfficeCsv -Path $Run.SourcePath) }
        NativeCsv { @(Import-Csv -Path $Run.SourcePath) }
        default { throw "Engine '$Engine' does not support CSV read." }
    }
    $Run.ActualRows = @($rows).Count
}

function Invoke-ExcelBenchmarkReadCsvDataTable {
    param([string] $Engine, [object] $Run)

    switch ($Engine) {
        PSWriteOffice {
            $table = Import-OfficeCsv -Path $Run.SourcePath -AsDataTable
            $Run.ActualRows = if ($table -and $table.Rows) { [int]$table.Rows.Count } else { 0 }
        }
        NativeCsv {
            $table = ConvertFrom-NativeCsvToDataTable -Path $Run.SourcePath
            $Run.ActualRows = if ($table -and $table.Rows) { [int]$table.Rows.Count } else { 0 }
        }
        default { throw "Engine '$Engine' does not support CSV DataTable read." }
    }
}

function Invoke-ExcelBenchmarkReadCsvGZipDataTable {
    param([string] $Engine, [object] $Run)

    switch ($Engine) {
        PSWriteOffice {
            $table = Import-OfficeCsv -Path $Run.SourcePath -CompressionType GZip -AsDataTable
            $Run.ActualRows = if ($table -and $table.Rows) { [int]$table.Rows.Count } else { 0 }
        }
        NativeCsv {
            $table = ConvertFrom-NativeGZipCsvToDataTable -Path $Run.SourcePath
            $Run.ActualRows = if ($table -and $table.Rows) { [int]$table.Rows.Count } else { 0 }
        }
        default { throw "Engine '$Engine' does not support GZip CSV DataTable read." }
    }
}

function ConvertFrom-NativeCsvToDataTable {
    param([Parameter(Mandatory)][string] $Path)

    ConvertFrom-NativeCsvRowsToDataTable -Rows (Import-Csv -Path $Path)
}

function ConvertFrom-NativeGZipCsvToDataTable {
    param([Parameter(Mandatory)][string] $Path)

    $stream = [IO.File]::OpenRead($Path)
    try {
        $gzip = [IO.Compression.GZipStream]::new($stream, [IO.Compression.CompressionMode]::Decompress, $false)
        try {
            $reader = [IO.StreamReader]::new($gzip, [Text.UTF8Encoding]::new($false))
            try {
                $csvText = $reader.ReadToEnd()
            } finally {
                $reader.Dispose()
            }
        } finally {
            $gzip.Dispose()
        }
    } finally {
        $stream.Dispose()
    }

    ConvertFrom-NativeCsvRowsToDataTable -Rows ($csvText | ConvertFrom-Csv)
}

function ConvertFrom-NativeCsvRowsToDataTable {
    param([Parameter(Mandatory)][object[]] $Rows)

    $table = [Data.DataTable]::new('CsvData')
    $columnNames = $null

    $table.BeginLoadData()
    try {
        foreach ($row in $Rows) {
            if ($null -eq $columnNames) {
                $columnNames = @($row.PSObject.Properties.Name)
                foreach ($columnName in $columnNames) {
                    $null = $table.Columns.Add([string]$columnName, [string])
                }
            }

            $dataRow = $table.NewRow()
            foreach ($columnName in $columnNames) {
                $value = $row.$columnName
                $dataRow[$columnName] = if ($null -eq $value) { [DBNull]::Value } else { $value }
            }

            $table.Rows.Add($dataRow)
        }
    } finally {
        $table.EndLoadData()
    }

    , $table
}

function Write-NativeGZipCsv {
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [object[]] $InputObject,

        [Parameter(Mandatory)]
        [string] $Path
    )

    $directory = [IO.Path]::GetDirectoryName($Path)
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not [IO.Directory]::Exists($directory)) {
        [IO.Directory]::CreateDirectory($directory) | Out-Null
    }

    $stream = [IO.File]::Create($Path)
    try {
        $gzip = [IO.Compression.GZipStream]::new($stream, [IO.Compression.CompressionLevel]::Optimal, $false)
        try {
            $writer = [IO.StreamWriter]::new($gzip, [Text.UTF8Encoding]::new($false))
            try {
                foreach ($line in ($InputObject | ConvertTo-Csv -NoTypeInformation -UseQuotes AsNeeded)) {
                    $writer.WriteLine($line)
                }
            } finally {
                $writer.Dispose()
            }
        } finally {
            $gzip.Dispose()
        }
    } finally {
        $stream.Dispose()
    }
}

function Invoke-ExcelBenchmarkReadCsvQuickSingleColumn {
    param([string] $Engine, [object] $Run)

    $count = 0
    $lastValue = $null
    switch ($Engine) {
        PSWriteOffice {
            Import-OfficeCsv -Path $Run.SourcePath | ForEach-Object {
                $lastValue = $_.Column0
                $count++
            }
        }
        NativeCsv {
            Import-Csv -Path $Run.SourcePath | ForEach-Object {
                $lastValue = $_.Column0
                $count++
            }
        }
        default { throw "Engine '$Engine' does not support CSV QuickTest single-column read." }
    }

    $Run.ActualRows = $count
    $Run.AccessedFields = $count
    $Run.LastValue = $lastValue
}

function Invoke-ExcelBenchmarkReadCsvQuickAllColumns {
    param([string] $Engine, [object] $Run)

    $count = 0
    $accessedFields = 0
    $lastValue = $null
    $columns = for ($column = 0; $column -lt [int]$Run.ColumnCount; $column++) {
        'Column{0}' -f $column
    }

    switch ($Engine) {
        PSWriteOffice {
            Import-OfficeCsv -Path $Run.SourcePath | ForEach-Object {
                $count++
                foreach ($column in $columns) {
                    $lastValue = $_.$column
                    $accessedFields++
                }
            }
        }
        NativeCsv {
            Import-Csv -Path $Run.SourcePath | ForEach-Object {
                $count++
                foreach ($column in $columns) {
                    $lastValue = $_.$column
                    $accessedFields++
                }
            }
        }
        default { throw "Engine '$Engine' does not support CSV QuickTest all-columns read." }
    }

    $Run.ActualRows = $count
    $Run.AccessedFields = $accessedFields
    $Run.LastValue = $lastValue
}

function Invoke-ExcelBenchmarkCsvToExcel {
    param([string] $Engine, [object] $Run)

    switch ($Engine) {
        PSWriteOffice { Import-OfficeExcelDelimitedText -Path $Run.Path -SourcePath $Run.SourcePath -SheetName $Run.WorksheetName | Out-Null }
        ImportExcel { Import-Csv -Path $Run.SourcePath | Export-Excel -Path $Run.Path -WorksheetName $Run.WorksheetName }
        default { throw "Engine '$Engine' does not support CSV-to-Excel conversion." }
    }
}

function Invoke-ExcelBenchmarkReadWorkbook {
    param([string] $Engine, [object] $Case, [object] $Run, [string] $Mode)

    switch ($Engine) {
        PSWriteOffice {
            $rows = switch ($Mode) {
                Full { @(Import-OfficeExcel -Path $Run.Path -WorksheetName $Run.WorksheetName) }
                Range { @(Import-OfficeExcel -Path $Run.Path -WorksheetName $Run.WorksheetName -Range $Run.Range) }
                NoHeader { @(Import-OfficeExcel -Path $Run.Path -WorksheetName $Run.WorksheetName -Range $Run.Range -NoHeader) }
            }
        }
        ImportExcel {
            $rows = switch ($Mode) {
                Full { @(Import-Excel -Path $Run.Path -WorksheetName $Run.WorksheetName) }
                Range { @(Import-Excel -Path $Run.Path -WorksheetName $Run.WorksheetName -StartRow 1 -EndRow ([int]$Case.RowCount + 1) -StartColumn 1 -EndColumn $Run.ColumnCount) }
                NoHeader { @(Import-Excel -Path $Run.Path -WorksheetName $Run.WorksheetName -StartRow 1 -EndRow ([int]$Case.RowCount + 1) -StartColumn 1 -EndColumn $Run.ColumnCount -NoHeader) }
            }
        }
        ExcelFast {
            $rows = switch ($Mode) {
                Full { @(Import-Workbook -Path $Run.Path -SheetName $Run.WorksheetName) }
                Range { @(Import-Workbook -Path $Run.Path -SheetName $Run.WorksheetName -StartCell 'A1' -EndCell $Run.RangeEndCell) }
                NoHeader { @(Import-Workbook -Path $Run.Path -SheetName $Run.WorksheetName -StartCell 'A1' -EndCell $Run.RangeEndCell -NoHeaders) }
            }
        }
    }
    $Run.ActualRows = @($rows).Count
    if ($Mode -ne 'NoHeader' -and $Run.ActualRows -gt 0) {
        $Run.ActualFields = @($rows[0].PSObject.Properties).Count
        $Run.FirstId = $rows[0].Id
        $Run.LastId = $rows[$Run.ActualRows - 1].Id
    }
}
function Invoke-ExcelBenchmarkWriteWorkbook {
    param([string] $Engine, [object] $Case, [object] $Run)

    switch ([string]$Case.Scenario) {
        objects-table { Invoke-ExcelBenchmarkObjectsTable -Engine $Engine -Run $Run }
        objects-default { Invoke-ExcelBenchmarkObjectsDefault -Engine $Engine -Run $Run }
        text-objects-default { Invoke-ExcelBenchmarkObjectsDefault -Engine $Engine -Run $Run }
        objects-no-table { Invoke-ExcelBenchmarkObjectsNoTable -Engine $Engine -Run $Run }
        objects-table-autofit { Invoke-ExcelBenchmarkObjectsTableAutofit -Engine $Engine -Run $Run }
        objects-title-freeze { Invoke-ExcelBenchmarkObjectsTitleFreeze -Engine $Engine -Run $Run }
        wide-objects-default { Invoke-ExcelBenchmarkObjectsDefault -Engine $Engine -Run $Run }
        datatable-default { Invoke-ExcelBenchmarkDataTableDefault -Engine $Engine -Run $Run }
        datareader-table { Invoke-ExcelBenchmarkDataReaderTable -Engine $Engine -Run $Run }
        multi-sheet-regions { Invoke-ExcelBenchmarkMultiSheetRegions -Engine $Engine -Run $Run }
        summary-formulas { Invoke-ExcelBenchmarkSummaryFormulas -Engine $Engine -Run $Run }
        append-existing-table { Invoke-ExcelBenchmarkAppendExistingTable -Engine $Engine -Run $Run }
        update-existing-workbook { Invoke-ExcelBenchmarkUpdateExistingWorkbook -Engine $Engine -Run $Run }
        many-small-sheets { Invoke-ExcelBenchmarkManySmallSheets -Engine $Engine -Run $Run }
        workbook-package-merge { Invoke-ExcelBenchmarkWorkbookPackageMerge -Engine $Engine -Run $Run }
        named-range-workbook { New-ExcelBenchmarkNamedRangeWorkbook -Engine $Engine -Run $Run }
        chart-only-workbook { Invoke-ExcelBenchmarkChartOnlyWorkbook -Engine $Engine -Run $Run }
        pivot-only-workbook { Invoke-ExcelBenchmarkPivotOnlyWorkbook -Engine $Engine -Run $Run }
        report-workbook { Invoke-ExcelBenchmarkReportWorkbook -Engine $Engine -Run $Run }
        dataset-worksheets { Export-OfficeExcel -Path $Run.Path -InputObject $Run.Payload }
        default { throw "Unknown workbook scenario '$($Case.Scenario)'." }
    }
}

function Invoke-ExcelBenchmarkObjectsTable {
    param([string] $Engine, [object] $Run)
    switch ($Engine) {
        PSWriteOffice { $Run.Payload | Export-OfficeExcel -Path $Run.Path -WorksheetName $Run.WorksheetName -TableName Data }
        ImportExcel { $Run.Payload | Export-Excel -Path $Run.Path -WorksheetName $Run.WorksheetName -TableName Data -AutoFilter }
    }
}

function Invoke-ExcelBenchmarkObjectsDefault {
    param([string] $Engine, [object] $Run)
    switch ($Engine) {
        PSWriteOffice { $Run.Payload | Export-OfficeExcel -Path $Run.Path -WorksheetName $Run.WorksheetName }
        ImportExcel { $Run.Payload | Export-Excel -Path $Run.Path -WorksheetName $Run.WorksheetName }
        ExcelFast {
            Export-Workbook -Destination $Run.Path -InputObject $Run.Payload -SheetName $Run.WorksheetName -Force

            # ExcelFast 0.0.1-alpha16 writes an empty drawing element inside sheetData,
            # which is not valid Open XML. Keep its own normalization cost in the timed
            # operation so the comparison only accepts a standards-valid workbook.
            $workbook = Get-Workbook -Path $Run.Path
            try {
                Save-Workbook -Workbook $workbook -Force
            } finally {
                if ($workbook -is [IDisposable]) {
                    $workbook.Dispose()
                }
            }
        }
    }
}

function Invoke-ExcelBenchmarkObjectsNoTable {
    param([string] $Engine, [object] $Run)
    switch ($Engine) {
        PSWriteOffice { $Run.Payload | Export-OfficeExcel -Path $Run.Path -WorksheetName $Run.WorksheetName -NoTable }
        ImportExcel { $Run.Payload | Export-Excel -Path $Run.Path -WorksheetName $Run.WorksheetName }
    }
}

function Invoke-ExcelBenchmarkObjectsTableAutofit {
    param([string] $Engine, [object] $Run)
    switch ($Engine) {
        PSWriteOffice { $Run.Payload | Export-OfficeExcel -Path $Run.Path -WorksheetName $Run.WorksheetName -TableName Data -AutoFit }
        ImportExcel { $Run.Payload | Export-Excel -Path $Run.Path -WorksheetName $Run.WorksheetName -TableName Data -AutoFilter -AutoSize }
    }
}

function Invoke-ExcelBenchmarkObjectsTitleFreeze {
    param([string] $Engine, [object] $Run)
    switch ($Engine) {
        PSWriteOffice { $Run.Payload | Export-OfficeExcel -Path $Run.Path -WorksheetName $Run.WorksheetName -TableName Data -Title 'Operational export' -StartRow 3 -FreezeTopRow -BoldTopRow }
        ImportExcel { $Run.Payload | Export-Excel -Path $Run.Path -WorksheetName $Run.WorksheetName -TableName Data -AutoFilter -Title 'Operational export' -StartRow 3 -FreezeTopRow -BoldTopRow }
    }
}

function Invoke-ExcelBenchmarkDataTableDefault {
    param([string] $Engine, [object] $Run)
    switch ($Engine) {
        PSWriteOffice { Export-OfficeExcel -Path $Run.Path -InputObject $Run.Payload -WorksheetName $Run.WorksheetName -TableName Data }
        ImportExcel { $Run.Payload | Export-Excel -Path $Run.Path -WorksheetName $Run.WorksheetName }
    }
}

function Invoke-ExcelBenchmarkDataReaderTable {
    param([string] $Engine, [object] $Run)
    if ($Engine -ne 'PSWriteOffice') {
        throw "Engine '$Engine' does not expose an equivalent IDataReader workbook export."
    }

    $reader = $Run.Payload.CreateDataReader()
    try {
        Export-OfficeExcel -Path $Run.Path -InputObject $reader -WorksheetName $Run.WorksheetName -TableName Data
    } finally {
        $reader.Dispose()
    }
}

function Invoke-ExcelBenchmarkMultiSheetRegions {
    param([string] $Engine, [object] $Run)
    switch ($Engine) {
        PSWriteOffice {
            New-OfficeExcel -Path $Run.Path {
                foreach ($group in (Get-ExcelBenchmarkRegionGroups -Rows $Run.Payload)) {
                    Add-OfficeExcelSheet -Name $group.Name -Content {
                        Add-OfficeExcelTable -Data $group.Data -TableName $group.TableName
                        Set-OfficeExcelFreeze -TopRows 1
                    }
                }
            } | Out-Null
        }
        ImportExcel {
            $excel = $null
            try {
                foreach ($group in (Get-ExcelBenchmarkRegionGroups -Rows $Run.Payload)) {
                    $excel = if ($excel) {
                        $group.Data | Export-Excel -ExcelPackage $excel -WorksheetName $group.Name -TableName $group.TableName -AutoFilter -FreezeTopRow -BoldTopRow -PassThru
                    } else {
                        $group.Data | Export-Excel -Path $Run.Path -WorksheetName $group.Name -TableName $group.TableName -AutoFilter -FreezeTopRow -BoldTopRow -PassThru
                    }
                }
            } finally {
                if ($excel) { Close-ExcelPackage -ExcelPackage $excel }
            }
        }
    }
}

function Invoke-ExcelBenchmarkSummaryFormulas {
    param([string] $Engine, [object] $Run)
    $summaryRows = Get-ExcelBenchmarkSummaryRows -Rows ([int]$Run.Payload.Count)
    switch ($Engine) {
        PSWriteOffice {
            New-OfficeExcel -Path $Run.Path {
                Add-OfficeExcelSheet -Name $Run.WorksheetName -Content {
                    Add-OfficeExcelTable -Data $Run.Payload -TableName Data
                    Set-OfficeExcelFreeze -TopRows 1
                }
                Add-OfficeExcelSheet -Name Summary -Content {
                    Set-OfficeExcelCell -Address A1 -Value Metric
                    Set-OfficeExcelCell -Address B1 -Value Value
                    $row = 2
                    foreach ($summaryRow in $summaryRows) {
                        Set-OfficeExcelCell -Row $row -Column 1 -Value $summaryRow.Metric
                        Set-OfficeExcelFormula -Address ('B{0}' -f $row) -Formula $summaryRow.Formula
                        Set-OfficeExcelCell -Row $row -Column 2 -NumberFormat $summaryRow.NumberFormat
                        $row++
                    }
                }
            } | Out-Null
        }
        ImportExcel {
            $excel = $Run.Payload | Export-Excel -Path $Run.Path -WorksheetName $Run.WorksheetName -TableName Data -AutoFilter -FreezeTopRow -BoldTopRow -PassThru
            try {
                $summary = $excel.Workbook.Worksheets.Add('Summary')
                $summary.Cells['A1'].Value = 'Metric'
                $summary.Cells['B1'].Value = 'Value'
                $row = 2
                foreach ($summaryRow in $summaryRows) {
                    $summary.Cells[$row, 1].Value = $summaryRow.Metric
                    $summary.Cells[$row, 2].Formula = $summaryRow.Formula
                    $summary.Cells[$row, 2].Style.Numberformat.Format = $summaryRow.NumberFormat
                    $row++
                }
            } finally {
                Close-ExcelPackage -ExcelPackage $excel
            }
        }
    }
}

function Invoke-ExcelBenchmarkAppendExistingTable {
    param([string] $Engine, [object] $Run)
    $split = Get-ExcelBenchmarkAppendSplit -Rows $Run.Payload
    switch ($Engine) {
        PSWriteOffice {
            $split.Initial | Export-OfficeExcel -Path $Run.Path -WorksheetName $Run.WorksheetName -TableName Data
            if ($split.Append.Count -gt 0) { Add-OfficeExcelTableRow -InputPath $Run.Path -Sheet $Run.WorksheetName -TableName Data -InputObject $split.Append | Out-Null }
        }
        ImportExcel {
            $split.Initial | Export-Excel -Path $Run.Path -WorksheetName $Run.WorksheetName -TableName Data -AutoFilter
            if ($split.Append.Count -gt 0) { $split.Append | Export-Excel -Path $Run.Path -WorksheetName $Run.WorksheetName -Append }
        }
    }
}

function Invoke-ExcelBenchmarkUpdateExistingWorkbook {
    param([string] $Engine, [object] $Run)
    switch ($Engine) {
        PSWriteOffice {
            $Run.Payload | Export-OfficeExcel -Path $Run.Path -WorksheetName $Run.WorksheetName -TableName Data
            Edit-OfficeExcelRow -InputPath $Run.Path -Sheet $Run.WorksheetName -ScriptBlock {
                param($row)
                $row.Set('TicketCount', ([int]$row.Get[int]('TicketCount') + 1))
                if ($row.Get[bool]('IsEnabled')) { $row.Set('Notes', 'Reviewed') }
            } | Out-Null
            $document = Get-OfficeExcel -Path $Run.Path
            try {
                $sheet = $document.Sheets | Where-Object { $_.Name -eq $Run.WorksheetName } | Select-Object -First 1
                $formulaColumn = $Run.ColumnCount + 1
                $sheet.Cell(1, $formulaColumn, 'ScoreDouble', $null, $null)
                $sheet.Cell(2, $formulaColumn, $null, 'G2*2', '#,##0.00')
                $document | Save-OfficeExcel
            } finally {
                $document | Close-OfficeExcel
            }
        }
        ImportExcel {
            $Run.Payload | Export-Excel -Path $Run.Path -WorksheetName $Run.WorksheetName -TableName Data -AutoFilter
            $excel = Open-ExcelPackage -Path $Run.Path
            try {
                $sheet = $excel.Workbook.Worksheets[$Run.WorksheetName]
                for ($row = 2; $row -le ([int]$Run.Payload.Count + 1); $row++) {
                    $sheet.Cells[$row, 9].Value = [int]$sheet.Cells[$row, 9].Value + 1
                    if ([bool]$sheet.Cells[$row, 5].Value) { $sheet.Cells[$row, 10].Value = 'Reviewed' }
                }
                $formulaColumn = $Run.ColumnCount + 1
                $sheet.Cells[1, $formulaColumn].Value = 'ScoreDouble'
                $sheet.Cells[2, $formulaColumn].Formula = 'G2*2'
                $sheet.Cells[2, $formulaColumn].Style.Numberformat.Format = '#,##0.00'
            } finally {
                Close-ExcelPackage -ExcelPackage $excel
            }
        }
    }
}

function Invoke-ExcelBenchmarkManySmallSheets {
    param([string] $Engine, [object] $Run)
    switch ($Engine) {
        PSWriteOffice {
            New-OfficeExcel -Path $Run.Path {
                foreach ($group in (Get-ExcelBenchmarkSmallSheetGroups -Rows $Run.Payload -SheetCount 20)) {
                    Add-OfficeExcelSheet -Name $group.Name -Content {
                        Add-OfficeExcelTable -Data $group.Data -TableName $group.TableName
                        Set-OfficeExcelFreeze -TopRows 1
                    }
                }
            } | Out-Null
        }
        ImportExcel {
            $excel = $null
            try {
                foreach ($group in (Get-ExcelBenchmarkSmallSheetGroups -Rows $Run.Payload -SheetCount 20)) {
                    $excel = if ($excel) {
                        $group.Data | Export-Excel -ExcelPackage $excel -WorksheetName $group.Name -TableName $group.TableName -AutoFilter -FreezeTopRow -BoldTopRow -PassThru
                    } else {
                        $group.Data | Export-Excel -Path $Run.Path -WorksheetName $group.Name -TableName $group.TableName -AutoFilter -FreezeTopRow -BoldTopRow -PassThru
                    }
                }
            } finally {
                if ($excel) { Close-ExcelPackage -ExcelPackage $excel }
            }
        }
    }
}

function Invoke-ExcelBenchmarkWorkbookPackageMerge {
    param([string] $Engine, [object] $Run)
    $input = Get-ExcelBenchmarkWorkbookMergeInput -Rows $Run.Payload -BasePath $Run.Path
    switch ($Engine) {
        PSWriteOffice {
            $input.RowsA | Export-OfficeExcel -Path $input.SourceA -WorksheetName Data -TableName DataA
            $input.RowsB | Export-OfficeExcel -Path $input.SourceB -WorksheetName Data -TableName DataB
            Join-OfficeExcelWorkbook -Path $Run.Path -SourcePath @($input.SourceA, $input.SourceB) -CopyMode Package -SheetNamePrefix Merged | Out-Null
        }
        ImportExcel {
            $input.RowsA | Export-Excel -Path $input.SourceA -WorksheetName Data -TableName DataA -AutoFilter
            $input.RowsB | Export-Excel -Path $input.SourceB -WorksheetName Data -TableName DataB -AutoFilter
            $targetPackage = [OfficeOpenXml.ExcelPackage]::new([IO.FileInfo]$Run.Path)
            try {
                foreach ($item in @(
                    [pscustomobject]@{ Path = $input.SourceA; Name = 'MergedDataA' }
                    [pscustomobject]@{ Path = $input.SourceB; Name = 'MergedDataB' }
                )) {
                    $sourcePackage = [OfficeOpenXml.ExcelPackage]::new([IO.FileInfo]$item.Path)
                    try {
                        $null = $targetPackage.Workbook.Worksheets.Add($item.Name, $sourcePackage.Workbook.Worksheets['Data'])
                    } finally {
                        $sourcePackage.Dispose()
                    }
                }
                $targetPackage.Save()
            } finally {
                $targetPackage.Dispose()
            }
        }
    }
}

function Invoke-ExcelBenchmarkChartOnlyWorkbook {
    param([string] $Engine, [object] $Run)
    switch ($Engine) {
        PSWriteOffice {
            New-OfficeExcel -Path $Run.Path {
                Add-OfficeExcelSheet -Name $Run.WorksheetName -Content {
                    Add-OfficeExcelTable -Data $Run.Payload -TableName Data
                    Add-OfficeExcelChart -TableName Data -Row 2 -Column 12 -Type ColumnClustered -Title 'Score by region'
                }
            } | Out-Null
        }
        ImportExcel {
            $lastRow = [int]$Run.Payload.Count + 1
            $excel = $Run.Payload | Export-Excel -Path $Run.Path -WorksheetName $Run.WorksheetName -TableName Data -AutoFilter -PassThru
            try {
                Add-ExcelChart -Worksheet $excel.Workbook.Worksheets[$Run.WorksheetName] -ChartType ColumnClustered -Title 'Score by region' -XRange "D2:D$lastRow" -YRange "G2:G$lastRow" -Row 2 -Column 12 -Width 640 -Height 360
            } finally {
                Close-ExcelPackage -ExcelPackage $excel
            }
        }
    }
}

function Invoke-ExcelBenchmarkPivotOnlyWorkbook {
    param([string] $Engine, [object] $Run)
    switch ($Engine) {
        PSWriteOffice {
            New-OfficeExcel -Path $Run.Path {
                Add-OfficeExcelSheet -Name $Run.WorksheetName -Content {
                    Add-OfficeExcelTable -Data $Run.Payload -TableName Data
                    Add-OfficeExcelPivotTable -SourceRange $Run.Range -DestinationCell L4 -Name SummaryPivot -RowField Region -ColumnField Department -DataField Score, TicketCount -DataFunction Average, Sum -DataDisplayName 'Average Score', Tickets -DataNumberFormat '#,##0.00', '#,##0'
                }
            } | Out-Null
        }
        ImportExcel {
            $excel = $Run.Payload | Export-Excel -Path $Run.Path -WorksheetName $Run.WorksheetName -TableName Data -AutoFilter -PassThru
            try {
                $worksheet = $excel.Workbook.Worksheets[$Run.WorksheetName]
                Add-PivotTable -ExcelPackage $excel -Address $worksheet.Cells['L4'] -SourceWorksheet $worksheet -SourceRange $worksheet.Tables[0].Address -PivotTableName SummaryPivot -PivotRows Region -PivotColumns Department -PivotData @{ Score = 'Average'; TicketCount = 'Sum' } -PivotNumberFormat '#,##0.00'
            } finally {
                Close-ExcelPackage -ExcelPackage $excel
            }
        }
    }
}

function Invoke-ExcelBenchmarkReportWorkbook {
    param([string] $Engine, [object] $Run)
    switch ($Engine) {
        PSWriteOffice {
            $lastRow = [int]$Run.Payload.Count + 1
            New-OfficeExcel -Path $Run.Path {
                Add-OfficeExcelSheet -Name $Run.WorksheetName -Content {
                    Add-OfficeExcelTable -Data $Run.Payload -TableName Data -AutoFit
                    Set-OfficeExcelFreeze -TopRows 1
                    Add-OfficeExcelConditionalRule -Range "G2:G$lastRow" -Operator GreaterThan -Formula1 '750'
                    Add-OfficeExcelConditionalDataBar -Range "G2:G$lastRow" -Color '#70AD47'
                    Add-OfficeExcelConditionalColorScale -Range "I2:I$lastRow" -StartColor '#F4CCCC' -EndColor '#D9EAD3'
                    Add-OfficeExcelConditionalIconSet -Range "I2:I$lastRow"
                    Add-OfficeExcelValidationList -Range "D2:D$lastRow" -Values NA, EU, APAC, LATAM
                    Set-OfficeExcelColumnStyleByHeader -Header Score -NumberFormat '#,##0.000'
                    Set-OfficeExcelColumnStyleByHeader -Header Created -NumberFormat 'yyyy-mm-dd hh:mm'
                    Add-OfficeExcelChart -TableName Data -Row 2 -Column 12 -Type ColumnClustered -Title 'Score by region'
                    Add-OfficeExcelPivotTable -SourceRange $Run.Range -DestinationCell L24 -Name SummaryPivot -RowField Region -ColumnField Department -DataField Score, TicketCount -DataFunction Average, Sum -DataDisplayName 'Average Score', Tickets -DataNumberFormat '#,##0.00', '#,##0'
                }
            } | Out-Null
        }
        ImportExcel {
            $lastRow = [int]$Run.Payload.Count + 1
            $excel = $Run.Payload | Export-Excel -Path $Run.Path -WorksheetName $Run.WorksheetName -TableName Data -AutoFilter -AutoSize -FreezeTopRow -BoldTopRow -PassThru
            try {
                $worksheet = $excel.Workbook.Worksheets[$Run.WorksheetName]
                Add-ConditionalFormatting -Worksheet $worksheet -Address "G2:G$lastRow" -RuleType GreaterThan -ConditionValue 750 -BackgroundColor LightPink
                Add-ConditionalFormatting -Worksheet $worksheet -Address "G2:G$lastRow" -DataBarColor Green
                Add-ConditionalFormatting -Worksheet $worksheet -Address "I2:I$lastRow" -RuleType ThreeColorScale
                Add-ConditionalFormatting -Worksheet $worksheet -Address "I2:I$lastRow" -ThreeIconsSet TrafficLights1
                Add-ExcelDataValidationRule -Worksheet $worksheet -Range "D2:D$lastRow" -ValidationType List -ValueSet @('NA', 'EU', 'APAC', 'LATAM')
                $worksheet.Cells["G2:G$lastRow"].Style.Numberformat.Format = '#,##0.000'
                $worksheet.Cells["F2:F$lastRow"].Style.Numberformat.Format = 'yyyy-mm-dd hh:mm'
                Add-ExcelChart -Worksheet $worksheet -ChartType ColumnClustered -Title 'Score by region' -XRange "D2:D$lastRow" -YRange "G2:G$lastRow" -Row 2 -Column 12 -Width 640 -Height 360
                Add-PivotTable -ExcelPackage $excel -Address $worksheet.Cells['L24'] -SourceWorksheet $worksheet -SourceRange $worksheet.Tables[0].Address -PivotTableName SummaryPivot -PivotRows Region -PivotColumns Department -PivotData @{ Score = 'Average'; TicketCount = 'Sum' } -PivotNumberFormat '#,##0.00'
            } finally {
                Close-ExcelPackage -ExcelPackage $excel
            }
        }
    }
}
