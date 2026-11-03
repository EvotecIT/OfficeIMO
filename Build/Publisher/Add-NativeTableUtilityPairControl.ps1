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
    param($FirstPage, [ValidateSet('Bare', 'Styled', 'Borders', 'Fills', 'Composed')] [string] $Kind)

    $null = Add-NativeTableDecorationControl -FirstPage $FirstPage -BorderTopology:($Kind -eq 'Composed')
    $first = $FirstPage.Shapes.Item('AuthoredTable')
    $label = $FirstPage.Shapes.Item('TableLabel')
    $objects.Add($first)
    $objects.Add($label)
    $first.Name = 'FirstDecoratedTable'
    $label.Name = 'FirstTableLabel'
    @($expectedShapes | Where-Object name -eq 'AuthoredTable')[-1].name = 'FirstDecoratedTable'
    @($expectedShapes | Where-Object name -eq 'TableLabel')[-1].name = 'FirstTableLabel'

    if ($Kind -eq 'Bare') {
        $null = Add-NativeTableCreationControl -FirstPage $FirstPage
    } else {
        $null = Add-NativeTableDecorationControl -FirstPage $FirstPage -BorderOnly:($Kind -eq 'Borders') -FillOnly:($Kind -eq 'Fills') -TextEdited:($Kind -eq 'Composed')
    }
    $second = $FirstPage.Shapes.Item('AuthoredTable')
    $secondLabel = $FirstPage.Shapes.Item('TableLabel')
    $objects.Add($second)
    $objects.Add($secondLabel)
    $second.Name = 'SecondTable'
    $second.Top = 360
    $secondLabel.Name = 'SecondTableLabel'
    $secondLabel.Top = 300
    $secondExpected = @($expectedShapes | Where-Object name -eq 'AuthoredTable')[-1]
    $secondExpected.name = 'SecondTable'
    $secondExpected.top = 360
    $labelExpected = @($expectedShapes | Where-Object name -eq 'TableLabel')[-1]
    $labelExpected.name = 'SecondTableLabel'
    $labelExpected.top = 300
    return 1
}
