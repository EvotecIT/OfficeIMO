# Documented Row.Height and Column.Width getters return points. Keep native
# observations separate from file decoding and managed table projection.
function Get-NativeTableGeometry {
    param($Shape)
    $table = $Shape.Table
    $objects.Add($table)
    $rows = $table.Rows
    $columns = $table.Columns
    $objects.Add($rows)
    $objects.Add($columns)
    $heights = [Collections.Generic.List[double]]::new()
    $widths = [Collections.Generic.List[double]]::new()
    for ($index = 1; $index -le $rows.Count; $index++) {
        $row = $rows.Item($index)
        $objects.Add($row)
        $heights.Add([double]$row.Height)
    }
    for ($index = 1; $index -le $columns.Count; $index++) {
        $column = $columns.Item($index)
        $objects.Add($column)
        $widths.Add([double]$column.Width)
    }
    return [pscustomobject]@{
        rowHeightsPoints = $heights.ToArray()
        columnWidthsPoints = $widths.ToArray()
    }
}
