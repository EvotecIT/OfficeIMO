function Get-NativeTableCellInventory {
    <#
    .SYNOPSIS
    Reads native cell text at every table grid coordinate without changing content.
    .PARAMETER Shape
    A native table shape. Merged cells can occupy multiple observed coordinates.
    .PARAMETER Objects
    Retains observed COM wrappers for the isolated worker's lifetime.
    .OUTPUTS
    One-based grid coordinates, raw native text including paragraph CRs, and row/column spans.
    #>
    param($Shape, [Collections.Generic.List[object]] $Objects)

    $table = $Shape.Table
    $Objects.Add($table)
    $rows = $table.Rows
    $columns = $table.Columns
    $Objects.Add($rows)
    $Objects.Add($columns)
    for ($row = 1; $row -le $rows.Count; $row++) {
        for ($column = 1; $column -le $columns.Count; $column++) {
            $range = $table.Cells($row, $column, $row, $column)
            $cell = $range.Item(1)
            $text = $cell.TextRange
            $Objects.Add($range)
            $Objects.Add($cell)
            $Objects.Add($text)
            [pscustomobject]@{
                row = $row
                column = $column
                text = [string] $text.Text
                columnSpan = [int] $cell.Width
                rowSpan = [int] $cell.Height
                style = Get-NativeTableCellStyle -Cell $cell -Objects $Objects
            }
        }
    }
}
