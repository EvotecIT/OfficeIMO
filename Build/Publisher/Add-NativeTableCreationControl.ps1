function Add-NativeTableCreationControl {
    <#
    .SYNOPSIS
    Authors borderless table controls with unequal tracks and an independent label.
    .PARAMETER FirstPage
    Task-owned new publication page.
    .PARAMETER Edited
    Changes cell paragraphs. The merged profile clears its horizontal cell and grows its vertical and block cells.
    .PARAMETER Merged
    Authors horizontal, vertical and rectangular merged cells in a three-by-four grid.
    .OUTPUTS
    The physical page count.
    #>
    param($FirstPage, [switch]$Edited, [switch]$Merged)

    $label = Add-TextFrame $FirstPage 'TableLabel' 'Independent table label' 72 72 300 54 12
    $label.TextFrame.ColumnSpacing = 0
    $label.TextFrame.TextRange.ParagraphFormat.SetLineSpacing(4, 18)
    $rowCount = if ($Merged) { 3 } else { 2 }
    $columnCount = if ($Merged) { 4 } else { 2 }
    $width = if ($Merged) { 480 } else { 300 }
    $height = if ($Merged) { 162 } else { 96 }
    $shape = $FirstPage.Shapes.AddTable($rowCount, $columnCount, 72, 144, $width, $height)
    $shape.Name = 'AuthoredTable'
    $objects.Add($shape)
    $table = $shape.Table
    $objects.Add($table)
    $table.ApplyAutoFormat(-2, $false, $false, $true, $true) # pbTableAutoFormatNone
    $columnWidths = if ($Merged) { @(90, 110, 130, 150) } else { @(120, 180) }
    for ($column = 1; $column -le $columnCount; $column++) {
        $table.Columns.Item($column).Width = $columnWidths[$column - 1]
    }
    for ($row = 1; $row -le $rowCount; $row++) {
        $table.Rows.Item($row).Height = 30 + 12 * $row
    }
    if ($Merged) {
        foreach ($coordinates in @(@(1, 1, 1, 2), @(2, 1, 3, 1), @(2, 3, 3, 4))) {
            $range = $table.Cells($coordinates[0], $coordinates[1], $coordinates[2], $coordinates[3])
            $objects.Add($range)
            $range.Merge()
        }
    }
    $values = @('Item', 'Count', 'Alpha', '')
    if ($Edited) {
        $values[0] = "Updated`rItem"
        $values[3] = '7'
    }
    for ($row = 1; $row -le $rowCount; $row++) {
        for ($column = 1; $column -le $columnCount; $column++) {
            if ($Merged -and (($row -eq 1 -and $column -eq 2) -or ($row -eq 3 -and $column -ne 2) -or ($row -eq 2 -and $column -eq 4))) {
                continue
            }
            $index = ($row - 1) * 2 + $column - 1
            $value = if ($Merged) { 'R{0}C{1}' -f $row, $column } else { $values[$index] }
            if ($Merged -and $row -eq 1 -and $column -eq 1) { $value = 'Horizontal' }
            if ($Merged -and $row -eq 2 -and $column -eq 1) { $value = 'Vertical' }
            if ($Merged -and $row -eq 2 -and $column -eq 3) { $value = "Block`rSecond line" }
            if ($Merged -and $Edited) {
                if ($row -eq 1 -and $column -eq 1) { $value = '' }
                if ($row -eq 2 -and $column -eq 1) { $value = "Vertical`rRetained" }
                if ($row -eq 2 -and $column -eq 3) { $value = "Updated`rBlock`rEnd" }
            }
            $cell = $table.Cells($row, $column, $row, $column).Item(1)
            $objects.Add($cell)
            $cell.MarginLeft = 0
            $cell.MarginRight = 0
            $cell.MarginTop = 0
            $cell.MarginBottom = 0
            $cell.Fill.Visible = 0
            $cell.TextRange.Text = $value
            $range = $cell.TextRange
            $objects.Add($range)
            $range.Font.Name = 'Arial'
            $range.Font.Size = 12
            $range.Font.Color.RGB = 0
            $range.Font.Bold = if ($row -eq 1) { -1 } else { 0 }
            $paragraph = $range.ParagraphFormat
            $objects.Add($paragraph)
            $paragraph.SpaceBefore = 0
            $paragraph.SpaceAfter = 0
            $paragraph.SetLineSpacing(4, 18)
            if ($value.Length) { $expectedText.Add($value) }
        }
    }
    $expectedShapes.Add([pscustomobject]@{ name = 'AuthoredTable'; kind = 'Table'; left = 72; top = 144; width = $width; height = $height; rows = $rowCount; columns = $columnCount })
    return 1
}
