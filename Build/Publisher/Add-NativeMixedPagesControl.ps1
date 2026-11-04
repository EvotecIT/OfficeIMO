function Add-NativeMixedPagesControl {
    <#
    .SYNOPSIS
    Authors overlapping text, rectangles and shared pictures in explicit append order.
    .DESCRIPTION
    Uses the isolated corpus worker's object tracking, text-frame helper and PNG
    generator. The first and last pages use different object orders; the middle
    page remains empty. All dimensions are integer points. Wrapping is explicitly
    disabled to match public model creation without text-wrap exclusions.
    .PARAMETER Document
    Task-owned new publication.
    .PARAMETER FirstPage
    Its first physical page.
    .PARAMETER ImagePath
    Unused path for the synthetic shared PNG.
    .PARAMETER Edited
    Authors changed first-page text, last-band paint/bounds and top-picture bounds.
    .PARAMETER ReverseOrder
    Reverses existing first/last-page shapes through native ZOrder without adding objects.
    .OUTPUTS
    The physical page count.
    #>
    param($Document, $FirstPage, [string]$ImagePath, [switch]$Edited, [switch]$ReverseOrder)

    $middle = $Document.Pages.Add(1, 1)
    $objects.Add($middle)
    $last = $Document.Pages.Add(1, 2)
    $objects.Add($last)
    New-ReferencePicture $ImagePath

    $background = $FirstPage.Shapes.AddShape(1, 72, 144, 240, 120)
    $background.Name = 'FirstBackground'
    $background.Fill.ForeColor.RGB = 255
    $background.Line.Visible = 0
    $objects.Add($background)
    $expectedShapes.Add([pscustomobject]@{ name = 'FirstBackground'; kind = 'Rectangle'; left = 72; top = 144; width = 240; height = 120 })
    $firstImage = $FirstPage.Shapes.AddPicture($ImagePath, 0, -1, 108, 144, 96, 96)
    $firstImage.Name = 'FirstImage'
    $objects.Add($firstImage)
    $expectedShapes.Add([pscustomobject]@{ name = 'FirstImage'; kind = 'Image'; left = 108; top = 144; width = 96; height = 96 })
    $firstTextValue = if ($Edited) { 'Edited first: overlap order' } else { 'Mixed first: overlap order' }
    $firstText = Add-TextFrame $FirstPage 'FirstText' $firstTextValue 72 144 288 108 18
    $firstBand = $FirstPage.Shapes.AddShape(1, 72, 168, 240, 12)
    $firstBand.Name = 'FirstBand'
    $firstBand.Fill.ForeColor.RGB = 16711680
    $firstBand.Line.Visible = 0
    $objects.Add($firstBand)
    $expectedShapes.Add([pscustomobject]@{ name = 'FirstBand'; kind = 'Rectangle'; left = 72; top = 168; width = 240; height = 12 })

    $lastText = Add-TextFrame $last 'LastText' 'Mixed last: overlap order' 72 144 288 108 18
    $lastImage = $last.Shapes.AddPicture($ImagePath, 0, -1, 108, 144, 96, 96)
    $lastImage.Name = 'LastImage'
    $objects.Add($lastImage)
    $expectedShapes.Add([pscustomobject]@{ name = 'LastImage'; kind = 'Image'; left = 108; top = 144; width = 96; height = 96 })
    $bandTop = if ($Edited) { 180 } else { 168 }
    $lastBand = $last.Shapes.AddShape(1, 72, $bandTop, 240, 12)
    $lastBand.Name = 'LastBand'
    $lastBand.Fill.ForeColor.RGB = if ($Edited) { 16711680 } else { 32768 }
    $lastBand.Line.Visible = 0
    $objects.Add($lastBand)
    $expectedShapes.Add([pscustomobject]@{ name = 'LastBand'; kind = 'Rectangle'; left = 72; top = $bandTop; width = 240; height = 12 })
    $imageLeft = if ($Edited) { 228 } else { 216 }
    $imageWidth = if ($Edited) { 84 } else { 96 }
    $topImage = $last.Shapes.AddPicture($ImagePath, 0, -1, $imageLeft, 144, $imageWidth, 96)
    $topImage.Name = 'TopImage'
    $objects.Add($topImage)
    $expectedShapes.Add([pscustomobject]@{ name = 'TopImage'; kind = 'Image'; left = $imageLeft; top = 144; width = $imageWidth; height = 96 })
    foreach ($shape in @($background, $firstImage, $firstText, $firstBand, $lastText, $lastImage, $lastBand, $topImage)) {
        $wrap = $shape.TextWrap
        $objects.Add($wrap)
        $wrap.Type = 0 # pbWrapTypeNone
        if ($wrap.Type -ne 0) { throw 'The mixed-order control requires wrapping disabled.' }
    }
    if ($ReverseOrder) {
        # msoBringToFront (0) applied in the desired back-to-front order.
        foreach ($shape in @($firstBand, $firstText, $firstImage, $background, $topImage, $lastBand, $lastImage, $lastText)) {
            $shape.ZOrder(0)
        }
    }
    return 3
}
