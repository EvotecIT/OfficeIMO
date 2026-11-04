function Add-NativeTableUtilityPairControl {
    <#
    .SYNOPSIS
    Authors two independent tables to isolate utility allocation and peer preservation.
    .PARAMETER FirstPage
    The task-owned new publication page.
    .PARAMETER Kind
    The second table is bare, filled, bordered or decorated; Composed also changes
    the first table's border topology and grows the second table's first cell.
    .OUTPUTS
    The physical page count.
    #>
    param($FirstPage, [ValidateSet('Bare', 'Styled', 'Borders', 'Fills', 'Composed', 'SingleFill', 'TargetComposed')] [string] $Kind)

    $null = Add-NativeTableDecorationControl -FirstPage $FirstPage -BorderTopology:($Kind -eq 'Composed')
    $first = $FirstPage.Shapes.Item('AuthoredTable')
    $label = $FirstPage.Shapes.Item('TableLabel')
    $objects.Add($first)
    $objects.Add($label)
    $first.Name = 'FirstDecoratedTable'
    $label.Name = 'FirstTableLabel'
    @($expectedShapes | Where-Object name -eq 'AuthoredTable')[-1].name = 'FirstDecoratedTable'
    @($expectedShapes | Where-Object name -eq 'TableLabel')[-1].name = 'FirstTableLabel'

    if ($Kind -in @('Bare', 'SingleFill')) {
        $null = Add-NativeTableCreationControl -FirstPage $FirstPage
    } else {
        $null = Add-NativeTableDecorationControl -FirstPage $FirstPage -BorderOnly:($Kind -eq 'Borders') -FillOnly:($Kind -eq 'Fills') -TextEdited:($Kind -in @('Composed', 'TargetComposed')) -BorderTopology:($Kind -eq 'TargetComposed')
    }
    $second = $FirstPage.Shapes.Item('AuthoredTable')
    $secondLabel = $FirstPage.Shapes.Item('TableLabel')
    $objects.Add($second)
    $objects.Add($secondLabel)
    $second.Name = 'SecondTable'
    $second.Top = 360
    $secondLabel.Name = 'SecondTableLabel'
    $secondLabel.Top = 300
    if ($Kind -eq 'SingleFill') {
        $table = $second.Table
        $range = $table.Cells(1, 1, 1, 1)
        $cell = $range.Item(1)
        $fill = $cell.Fill
        $color = $fill.ForeColor
        foreach ($value in @($table, $range, $cell, $fill, $color)) { $objects.Add($value) }
        $fill.Solid()
        $fill.Visible = -1
        $fill.Transparency = 0
        $color.RGB = 0x00A0E0FF
    }
    $secondExpected = @($expectedShapes | Where-Object name -eq 'AuthoredTable')[-1]
    $secondExpected.name = 'SecondTable'
    $secondExpected.top = 360
    $labelExpected = @($expectedShapes | Where-Object name -eq 'TableLabel')[-1]
    $labelExpected.name = 'SecondTableLabel'
    $labelExpected.top = 300
    return 1
}
