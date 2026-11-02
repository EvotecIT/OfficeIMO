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
    .PARAMETER FillVisibility
    Hides the first cell fill and activates a lavender fill in the initially unfilled cell.
    .PARAMETER RestoreFillVisibility
    Applies the visibility changes, then restores their visible states while retaining the latent colors.
    .PARAMETER BorderTopology
    Clears the first left edge, paints both top cell edges green, and adds an orange bottom-left edge.
    .PARAMETER ClearBorders
    Clears all six original edges while keeping cell fills and text.
    .PARAMETER DefaultCellFill
    Retains the native white-fill activation caused by adding an edge to the previously unfilled cell.
    .OUTPUTS
    The physical page count.
    #>
    param($FirstPage, [switch] $Edited, [switch] $TextEdited, [switch] $FillVisibility, [switch] $RestoreFillVisibility,
        [switch] $BorderTopology, [switch] $ClearBorders, [switch] $DefaultCellFill)

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
    $edgeChanges = @()
    if ($ClearBorders) {
        $edgeChanges = @($borders | ForEach-Object { @{ row = $_.row; column = $_.column; side = $_.side; width = 0; rgb = $null } })
    } elseif ($BorderTopology) {
        $edgeChanges = @(
            @{ row = 1; column = 1; side = 'BorderLeft'; width = 0; rgb = $null },
            @{ row = 1; column = 1; side = 'BorderTop'; width = 2; rgb = 0x00008000 },
            @{ row = 1; column = 2; side = 'BorderTop'; width = 2; rgb = 0x00008000 },
            @{ row = 2; column = 1; side = 'BorderBottom'; width = 2; rgb = 0x000080FF }
        )
    }
    foreach ($change in $edgeChanges) {
        $range = $table.Cells($change.row, $change.column, $change.row, $change.column)
        $cell = $range.Item(1)
        $format = $cell.($change.side)
        $objects.Add($range)
        $objects.Add($cell)
        $objects.Add($format)
        if ($null -ne $change.rgb) {
            $color = $format.Color
            $objects.Add($color)
            $color.RGB = $change.rgb
        }
        $format.Weight = $change.width
    }
    if ($BorderTopology -and -not $DefaultCellFill) {
        $range = $table.Cells(2, 1, 2, 1)
        $cell = $range.Item(1)
        $format = $cell.Fill
        $objects.Add($range)
        $objects.Add($cell)
        $objects.Add($format)
        $format.Visible = 0
    }
    if ($FillVisibility -or $RestoreFillVisibility) {
        foreach ($change in @(
            @{ row = 1; column = 1; visible = 0; rgb = $null },
            @{ row = 2; column = 1; visible = -1; rgb = 0x00FFC0E0 }
        )) {
            $range = $table.Cells($change.row, $change.column, $change.row, $change.column)
            $cell = $range.Item(1)
            $format = $cell.Fill
            $objects.Add($range)
            $objects.Add($cell)
            $objects.Add($format)
            if ($null -ne $change.rgb) {
                $format.Solid()
                $format.Transparency = 0
                $color = $format.ForeColor
                $objects.Add($color)
                $color.RGB = $change.rgb
            }
            $format.Visible = $change.visible
        }
        if ($RestoreFillVisibility) {
            foreach ($change in @(@{ row = 1; column = 1; visible = -1 }, @{ row = 2; column = 1; visible = 0 })) {
                $range = $table.Cells($change.row, $change.column, $change.row, $change.column)
                $cell = $range.Item(1)
                $format = $cell.Fill
                $objects.Add($range)
                $objects.Add($cell)
                $objects.Add($format)
                $format.Visible = $change.visible
            }
        }
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
