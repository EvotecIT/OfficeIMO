function Add-NativeMergedTableDecorationControl {
    <#
    .SYNOPSIS
    Decorates the seven logical cells in the existing merged-table control.
    .DESCRIPTION
    Reuses the native merged-table creator and applies final fills after borders,
    preserving the unequal tracks, merged anchors and independent label.
    .PARAMETER FirstPage
    The task-owned new publication page.
    .PARAMETER Variant
    Base applies fills and borders. Edited changes their paint and topology.
    Text changes only the creator's cell text; Composed combines both changes.
    Fills omits borders, while Borders leaves all cell fills invisible.
    .EXAMPLE
    Add-NativeMergedTableDecorationControl -FirstPage $page -Variant Composed
    .OUTPUTS
    The physical page count.
    #>
    param(
        $FirstPage,
        [ValidateSet('Base', 'Edited', 'Text', 'Composed', 'Fills', 'Borders')]
        [string] $Variant = 'Base'
    )

    Set-NativePhase 'AuthorMergedDecorationTable'
    $textEdited = $Variant -in @('Text', 'Composed')
    $paintEdited = $Variant -in @('Edited', 'Composed')
    $null = Add-NativeTableCreationControl -FirstPage $FirstPage -Merged -Edited:$textEdited

    Set-NativePhase 'FindMergedDecorationTable'
    $shapes = $FirstPage.Shapes
    $shape = $shapes.Item('AuthoredTable')
    $table = $shape.Table
    $objects.Add($shapes)
    $objects.Add($shape)
    $objects.Add($table)
    $anchors = @(
        @{
            row = 1
            column = 1
            rgb = 0x00A0E0FF
            visible = -1
        },
        @{
            row = 1
            column = 3
            rgb = 0x00FFC0E0
            visible = -1
        },
        @{
            row = 1
            column = 4
            rgb = $null
            visible = 0
        },
        @{
            row = 2
            column = 1
            rgb = 0x00FFE0A0
            visible = -1
        },
        @{
            row = 2
            column = 2
            rgb = 0x0080C0FF
            visible = -1
        },
        @{
            row = 2
            column = 3
            rgb = 0x00B4F0B4
            visible = -1
        },
        @{
            row = 3
            column = 2
            rgb = $null
            visible = 0
        }
    )
    if ($paintEdited) {
        $anchors[0].rgb = 0x00C0E0C0
        $anchors[1].visible = 0
        $anchors[2].rgb = 0x00B4F0B4
        $anchors[2].visible = -1
        $anchors[3].visible = 0
        $anchors[5].rgb = 0x00FFC0E0
    }

    if ($Variant -ne 'Fills') {
        for ($row = 1; $row -le 3; $row++) {
            Set-NativePhase ('AuthorMergedDecorationBorderRow-' + $row)
            foreach ($anchor in $anchors | Where-Object { $_.row -eq $row }) {
                $range = $table.Cells($anchor.row, $anchor.column, $anchor.row, $anchor.column)
                $cell = $range.Item(1)
                $objects.Add($range)
                $objects.Add($cell)
                foreach ($side in @('Left', 'Top', 'Right', 'Bottom')) {
                    $border = $cell.('Border' + $side)
                    $color = $border.Color
                    $objects.Add($border)
                    $objects.Add($color)
                    $color.RGB = 0x00008000
                    $border.Weight = 0.5
                }
            }
        }
        $edgeChanges = @(
            @{
                row = 3
                column = 2
                side = 'Bottom'
                width = 0
                rgb = $null
            }
        )
        if ($paintEdited) {
            $edgeChanges = @(
                @{
                    row = 1
                    column = 1
                    side = 'Left'
                    width = 0
                    rgb = $null
                },
                @{
                    row = 1
                    column = 1
                    side = 'Top'
                    width = 2
                    rgb = 0x00FF0000
                },
                @{
                    row = 3
                    column = 2
                    side = 'Bottom'
                    width = 1
                    rgb = 0x000080FF
                }
            )
        }
        foreach ($change in $edgeChanges) {
            Set-NativePhase ('ChangeMergedDecorationBorder-{0}-{1}-{2}' -f $change.row, $change.column, $change.side)
            $range = $table.Cells($change.row, $change.column, $change.row, $change.column)
            $cell = $range.Item(1)
            $border = $cell.('Border' + $change.side)
            $objects.Add($range)
            $objects.Add($cell)
            $objects.Add($border)
            if ($null -ne $change.rgb) {
                $color = $border.Color
                $objects.Add($color)
                $color.RGB = $change.rgb
            }
            $border.Weight = $change.width
        }
    }

    # Native border authoring can activate white fills, so establish final paint last.
    for ($row = 1; $row -le 3; $row++) {
        Set-NativePhase ('AuthorMergedDecorationFillRow-' + $row)
        foreach ($anchor in $anchors | Where-Object { $_.row -eq $row }) {
            $range = $table.Cells($anchor.row, $anchor.column, $anchor.row, $anchor.column)
            $cell = $range.Item(1)
            $fill = $cell.Fill
            $objects.Add($range)
            $objects.Add($cell)
            $objects.Add($fill)
            if ($Variant -eq 'Borders') {
                $fill.Visible = 0
                continue
            }
            if ($null -ne $anchor.rgb) {
                $fill.Solid()
                $fill.Transparency = 0
                $color = $fill.ForeColor
                $objects.Add($color)
                $color.RGB = $anchor.rgb
            }
            $fill.Visible = $anchor.visible
        }
    }
    return 1
}
