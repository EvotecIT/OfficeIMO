function Add-NativeTableDecorationControl {
    <#
    .SYNOPSIS
    Decorates the existing two-by-two control without changing its text or tracks.
    .PARAMETER FirstPage
    Task-owned new publication page.
    .PARAMETER Edited
    Changes the first cell's fill and selected border colors and widths.
    .PARAMETER TextEdited
    Grows the first cell to two paragraphs while keeping its initial decoration.
    .OUTPUTS
    The physical page count.
    #>
    param($FirstPage, [switch] $Edited, [switch] $TextEdited)

    $null = Add-NativeTableCreationControl -FirstPage $FirstPage
    $shape = $FirstPage.Shapes.Item('AuthoredTable')
    $table = $shape.Table
    $objects.Add($shape)
    $objects.Add($table)
    $fills = @(
        @{ row = 1; column = 1; rgb = $(if ($Edited) { 0x00C0E0C0 } else { 0x00A0E0FF }) },
        @{ row = 1; column = 2; rgb = 0x00FFE0A0 },
        @{ row = 2; column = 2; rgb = 0x00B4F0B4 }
    )
    foreach ($fill in $fills) {
        $range = $table.Cells($fill.row, $fill.column, $fill.row, $fill.column)
        $cell = $range.Item(1)
        $format = $cell.Fill
        $color = $format.ForeColor
        $objects.Add($range)
        $objects.Add($cell)
        $objects.Add($format)
        $objects.Add($color)
        $format.Solid()
        $format.Visible = -1
        $format.Transparency = 0
        $color.RGB = $fill.rgb
    }
    $borders = @(
        @{ row = 1; column = 1; side = 'BorderLeft'; rgb = $(if ($Edited) { 0x00008000 } else { 0x000000FF }); width = $(if ($Edited) { 1 } else { 2 }) },
        @{ row = 1; column = 1; side = 'BorderTop'; rgb = 0x00FF0000; width = 3 },
        @{ row = 1; column = 1; side = 'BorderRight'; rgb = 0x00008000; width = 1 },
        @{ row = 1; column = 1; side = 'BorderBottom'; rgb = 0x00800080; width = $(if ($Edited) { 3 } else { 1 }) },
        @{ row = 2; column = 2; side = 'BorderRight'; rgb = 0x00FF00FF; width = 4 },
        @{ row = 2; column = 2; side = 'BorderBottom'; rgb = 0x00FFFF00; width = 2 }
    )
    foreach ($border in $borders) {
        $range = $table.Cells($border.row, $border.column, $border.row, $border.column)
        $cell = $range.Item(1)
        $format = $cell.($border.side)
        $color = $format.Color
        $objects.Add($range)
        $objects.Add($cell)
        $objects.Add($format)
        $objects.Add($color)
        $color.RGB = $border.rgb
        $format.Weight = $border.width
    }
    if ($TextEdited) {
        $range = $table.Cells(1, 1, 1, 1)
        $cell = $range.Item(1)
        $text = $cell.TextRange
        $objects.Add($range)
        $objects.Add($cell)
        $objects.Add($text)
        $text.Text = "Updated`rItem"
        $expectedText[0] = "Updated`rItem"
    }
    return 1
}
