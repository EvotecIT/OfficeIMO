function Add-NativeImagePagesControl {
    <#
    .SYNOPSIS
    Authors two distinct embedded images on separated native physical pages.
    .DESCRIPTION
    Runs in the isolated corpus worker and uses its owned-object collection and
    existing synthetic PNG generators. The middle page remains empty.
    .PARAMETER Document
    Task-owned new publication.
    .PARAMETER FirstPage
    Its first physical page.
    .PARAMETER ImagePath
    Unused output path for the first synthetic PNG.
    .PARAMETER SecondImagePath
    Unused output path for the distinct wide synthetic PNG.
    .PARAMETER Edited
    Independently authors the state after moving and resizing the last image.
    .OUTPUTS
    The physical page count.
    #>
    param($Document, $FirstPage, [string]$ImagePath, [string]$SecondImagePath, [switch]$Edited)

    $middle = $Document.Pages.Add(1, 1)
    $objects.Add($middle)
    $last = $Document.Pages.Add(1, 2)
    $objects.Add($last)
    New-ReferencePicture $ImagePath
    New-NativeReplacementImage $SecondImagePath
    $first = $FirstPage.Shapes.AddPicture($ImagePath, 0, -1, 72, 144, 96, 96)
    $first.Name = 'FirstPageImage'
    $objects.Add($first)
    $expectedShapes.Add([pscustomobject]@{ name = 'FirstPageImage'; kind = 'Image'; left = 72; top = 144; width = 96; height = 96 })
    $bounds = if ($Edited) { @(252, 180, 144, 72) } else { @(216, 144, 192, 96) }
    $second = $last.Shapes.AddPicture($SecondImagePath, 0, -1, $bounds[0], $bounds[1], $bounds[2], $bounds[3])
    $second.Name = 'LastPageImage'
    $objects.Add($second)
    $expectedShapes.Add([pscustomobject]@{ name = 'LastPageImage'; kind = 'Image'; left = $bounds[0]; top = $bounds[1]; width = $bounds[2]; height = $bounds[3] })
    return 3
}
