function Invoke-ExcelBenchmarkReadDataTable {
    param([object] $Run, [string] $Mode)

    $parameters = @{
        Path = $Run.Path
        WorksheetName = $Run.WorksheetName
        AsDataTable = $true
    }
    if ($Mode -eq 'Range') {
        $parameters.Range = $Run.Range
    }

    $table = Import-OfficeExcel @parameters
    $Run.ActualRows = if ($table -and $table.Rows) { [int]$table.Rows.Count } else { 0 }
    $Run.ActualFields = if ($table -and $table.Columns) { [int]$table.Columns.Count } else { 0 }
    if ($Run.ActualRows -gt 0 -and $Run.ActualFields -gt 0) {
        $Run.FirstId = $table.Rows[0][0]
        $Run.LastId = $table.Rows[$Run.ActualRows - 1][0]
    }
}

function Invoke-ExcelBenchmarkReadDataReader {
    param([object] $Run, [string] $Mode)

    $parameters = @{
        Path = $Run.Path
        WorksheetName = $Run.WorksheetName
        AsDataReader = $true
    }
    if ($Mode -eq 'Range') {
        $parameters.Range = $Run.Range
    }

    $reader = Import-OfficeExcel @parameters
    try {
        $Run.ActualRows = 0
        $Run.ActualFields = [int]$reader.FieldCount
        while ($reader.Read()) {
            $id = $reader.GetValue(0)
            if ($Run.ActualRows -eq 0) {
                $Run.FirstId = $id
            }
            $Run.LastId = $id
            $Run.ActualRows++
        }
    } finally {
        if ($reader -is [IDisposable]) {
            $reader.Dispose()
        }
    }
}
