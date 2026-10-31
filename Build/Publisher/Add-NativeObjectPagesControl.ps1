function Add-NativeObjectPagesControl {
    <#
    .SYNOPSIS
    Authors independent shape/picture page states for native writer qualification.
    .DESCRIPTION
    Runs inside the existing isolated corpus worker, using its owned-object and
    expected-shape collections and synthetic image helper. The middle page stays
    empty; the edited state changes only rectangle bounds/fill and picture bounds.
    .PARAMETER Document
    Task-owned new publication.
    .PARAMETER FirstPage
    Its first physical page.
    .PARAMETER ImagePath
    Unused path for the synthetic PNG retained with this control.
    .PARAMETER Edited
    Produces the independently authored post-edit state.
    .OUTPUTS
    The physical page count.
    #>
    param($Document, $FirstPage, [string]$ImagePath, [switch]$Edited)

    $middle = $Document.Pages.Add(1, 1)
    $objects.Add($middle)
    $last = $Document.Pages.Add(1, 2)
    $objects.Add($last)
    $rectangleBounds = if ($Edited) { @(90, 180, 180, 84) } else { @(72, 216, 144, 72) }
    $pictureBounds = if ($Edited) { @(300, 240, 120, 72) } else { @(330, 216, 96, 96) }
    $rectangle = $FirstPage.Shapes.AddShape(1, $rectangleBounds[0], $rectangleBounds[1], $rectangleBounds[2], $rectangleBounds[3])
    $rectangle.Name = 'FirstPageRectangle'
    $rectangle.Fill.ForeColor.RGB = if ($Edited) { 16711680 } else { 255 }
    $rectangle.Line.Visible = 0
    $objects.Add($rectangle)
    $expectedShapes.Add([pscustomobject]@{ name = 'FirstPageRectangle'; kind = 'Rectangle'; left = $rectangleBounds[0]; top = $rectangleBounds[1]; width = $rectangleBounds[2]; height = $rectangleBounds[3] })
    New-ReferencePicture $ImagePath
    $image = $last.Shapes.AddPicture($ImagePath, 0, -1, $pictureBounds[0], $pictureBounds[1], $pictureBounds[2], $pictureBounds[3])
    $image.Name = 'LastPageImage'
    $objects.Add($image)
    $expectedShapes.Add([pscustomobject]@{ name = 'LastPageImage'; kind = 'Image'; left = $pictureBounds[0]; top = $pictureBounds[1]; width = $pictureBounds[2]; height = $pictureBounds[3] })
    return 3
}
