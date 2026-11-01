function Add-NativeWrapOrderControl {
    <#
    .SYNOPSIS
    Authors overlapping native text and square-wrapped primitive/picture obstacles.
    .DESCRIPTION
    Retains the worker's object references and uses no additional native getters.
    The paired order variant changes stacking only; NoWrap supplies a flow control.
    #>
    param($FirstPage, [string]$ImagePath,
        [ValidateSet('Original','Reverse','RectangleOnly','PictureOnly','Swap')][string]$Order='Original',
        [switch]$NoWrap, [ValidateRange(0,3)][int]$Side=0, [switch]$Outside,
        [switch]$FrameOutside, [switch]$Layout, [switch]$UseDefaults)

    $text = ((1..8 | ForEach-Object {
        'Paragraph {0}. Alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma tau.' -f $_
    }) -join "`r")
    $frameLeft = if ($FrameOutside) { 480 } else { 72 }
    $frameWidth = if ($FrameOutside) { 100 } else { 360 }
    $frame = Add-TextFrame $FirstPage 'WrappedStory' $text $frameLeft 72 $frameWidth 504 14
    if ($Layout) {
        $format = $frame.TextFrame; $objects.Add($format)
        $format.Columns = 2; $format.ColumnSpacing = 6
        $format.MarginLeft = 4; $format.MarginTop = 5; $format.MarginRight = 6; $format.MarginBottom = 7
        $format.VerticalTextAlignment = 1
    }
    $rectangleLeft = if ($Outside) { 480 } else { 180 }
    $pictureLeft = if ($Outside) { 480 } else { 144 }
    $rectangle = $FirstPage.Shapes.AddShape(1, $rectangleLeft, 132, 72, 72)
    $objects.Add($rectangle)
    $rectangle.Name = 'WrapRectangle'
    $rectangle.Fill.ForeColor.RGB = 255
    $rectangle.Line.Visible = 0
    $expectedShapes.Add([pscustomobject]@{name='WrapRectangle';kind='Rectangle';left=$rectangleLeft;top=132;width=72;height=72})
    New-ReferencePicture $ImagePath
    $picture = $FirstPage.Shapes.AddPicture($ImagePath, 0, -1, $pictureLeft, 264, 96, 96)
    $objects.Add($picture)
    $picture.Name = 'WrapPicture'
    $expectedShapes.Add([pscustomobject]@{name='WrapPicture';kind='Image';left=$pictureLeft;top=264;width=96;height=96})
    foreach ($shape in @($frame, $rectangle, $picture)) {
        $wrap = $shape.TextWrap
        $objects.Add($wrap)
        if ($UseDefaults -and -not [object]::ReferenceEquals($shape, $frame)) {
            $expectedShapes | Where-Object name -eq $shape.Name | Add-Member -NotePropertyName observedWrapType -NotePropertyValue ([int]$wrap.Type)
            continue
        }
        $wrap.Type = if ($NoWrap -or [object]::ReferenceEquals($shape, $frame)) { 0 } else { 1 }
        if ($NoWrap -or [object]::ReferenceEquals($shape, $frame)) { continue }
        $wrap.Side = $Side # Both=0, Left=1, Right=2, Larger=3
        $wrap.DistanceAuto = 0 # msoFalse
        $wrap.DistanceLeft = 6
        $wrap.DistanceRight = 6
        $wrap.DistanceTop = 6
        $wrap.DistanceBottom = 6
    }
    $desired = @($frame, $rectangle, $picture)
    switch ($Order) {
        'Reverse' { $desired = @($picture, $rectangle, $frame) }
        'RectangleOnly' { $desired = @($picture, $frame, $rectangle) }
        'PictureOnly' { $desired = @($rectangle, $frame, $picture) }
        'Swap' { $desired = @($frame, $picture, $rectangle) }
    }
    if ($Order -ne 'Original') {
        foreach ($shape in $desired) { $shape.ZOrder(0) }
    }
    return 1
}
