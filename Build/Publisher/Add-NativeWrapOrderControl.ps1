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
        [switch]$NoWrap)

    $text = ((1..8 | ForEach-Object {
        'Paragraph {0}. Alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma tau.' -f $_
    }) -join "`r")
    $frame = Add-TextFrame $FirstPage 'WrappedStory' $text 72 72 360 504 14
    $rectangle = $FirstPage.Shapes.AddShape(1, 180, 132, 72, 72)
    $objects.Add($rectangle)
    $rectangle.Name = 'WrapRectangle'
    $rectangle.Fill.ForeColor.RGB = 255
    $rectangle.Line.Visible = 0
    $expectedShapes.Add([pscustomobject]@{name='WrapRectangle';kind='Rectangle';left=180;top=132;width=72;height=72})
    New-ReferencePicture $ImagePath
    $picture = $FirstPage.Shapes.AddPicture($ImagePath, 0, -1, 144, 264, 96, 96)
    $objects.Add($picture)
    $picture.Name = 'WrapPicture'
    $expectedShapes.Add([pscustomobject]@{name='WrapPicture';kind='Image';left=144;top=264;width=96;height=96})
    foreach ($shape in @($frame, $rectangle, $picture)) {
        $wrap = $shape.TextWrap
        $objects.Add($wrap)
        $wrap.Type = if ($NoWrap -or [object]::ReferenceEquals($shape, $frame)) { 0 } else { 1 }
        if ($NoWrap -or [object]::ReferenceEquals($shape, $frame)) { continue }
        $wrap.Side = 0 # pbWrapSideBoth
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
