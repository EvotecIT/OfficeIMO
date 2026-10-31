function Add-NativeTableCreationControl {
    <#
    .SYNOPSIS
    Authors a borderless two-by-two table with unequal tracks and an independent label.
    .PARAMETER FirstPage
    Task-owned new publication page.
    .PARAMETER Edited
    Uses two paragraphs in the first cell and refills the final empty cell.
    .OUTPUTS
    The physical page count.
    #>
    param($FirstPage, [switch]$Edited)

    $label = Add-TextFrame $FirstPage 'TableLabel' 'Independent table label' 72 72 300 54 12
    $label.TextFrame.ColumnSpacing = 0
    $label.TextFrame.TextRange.ParagraphFormat.SetLineSpacing(4, 18)
    $shape = $FirstPage.Shapes.AddTable(2, 2, 72, 144, 300, 96)
    $shape.Name = 'AuthoredTable'
    $objects.Add($shape)
    $table = $shape.Table
    $objects.Add($table)
    $table.ApplyAutoFormat(-2, $false, $false, $true, $true) # pbTableAutoFormatNone
    $table.Columns.Item(1).Width = 120
    $table.Columns.Item(2).Width = 180
    $table.Rows.Item(1).Height = 42
    $table.Rows.Item(2).Height = 54
    $values = @('Item', 'Count', 'Alpha', '')
    if ($Edited) {
        $values[0] = "Updated`rItem"
        $values[3] = '7'
    }
    for ($row = 1; $row -le 2; $row++) {
        for ($column = 1; $column -le 2; $column++) {
            $index = ($row - 1) * 2 + $column - 1
            $cell = $table.Cells($row, $column, $row, $column).Item(1)
            $objects.Add($cell)
            $cell.MarginLeft = 0
            $cell.MarginRight = 0
            $cell.MarginTop = 0
            $cell.MarginBottom = 0
            $cell.Fill.Visible = 0
            $cell.TextRange.Text = $values[$index]
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
            if ($values[$index].Length) { $expectedText.Add($values[$index]) }
        }
    }
    $expectedShapes.Add([pscustomobject]@{ name = 'AuthoredTable'; kind = 'Table'; left = 72; top = 144; width = 300; height = 96; rows = 2; columns = 2 })
    return 1
}
