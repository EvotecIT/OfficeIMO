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
    .PARAMETER ReplaceLast
    Replaces the last picture with the first PNG while retaining both bounds.
    .PARAMETER ReplacementJpegPath
    Unused output path for a JPEG that replaces the last picture after its PNG replacement.
    .PARAMETER Cropped
    Applies the independently specified four-edge source crop to the first picture.
    .PARAMETER CropLast
    Applies that crop to the final JPEG placement after native replacement.
    .PARAMETER SharedImages
    Adds another placement of the first image to the last physical page.
    .PARAMETER ReplaceSharedFirst
    Replaces the first placement with the wide PNG, retaining the shared last-page placement.
    .OUTPUTS
    The physical page count.
    #>
    param($Document, $FirstPage, [string]$ImagePath, [string]$SecondImagePath, [switch]$Edited, [switch]$ReplaceLast,
        [string]$ReplacementJpegPath, [switch]$Cropped, [switch]$CropLast, [switch]$SharedImages,
        [switch]$ReplaceSharedFirst)

    $middle = $Document.Pages.Add(1, 1)
    $objects.Add($middle)
    $last = $Document.Pages.Add(1, 2)
    $objects.Add($last)
    New-ReferencePicture $ImagePath
    New-NativeReplacementImage $SecondImagePath
    $first = $FirstPage.Shapes.AddPicture($ImagePath, 0, -1, 72, 144, 96, 96)
    $first.Name = 'FirstPageImage'
    $objects.Add($first)
    if ($Cropped) {
        $crop = Get-NativePictureCropBasis $ImagePath
        $format = $first.PictureFormat
        $objects.Add($format)
        $first.LockAspectRatio = 0
        $format.CropLeft = $crop.points[0]; $format.CropTop = $crop.points[1]
        $format.CropRight = $crop.points[2]; $format.CropBottom = $crop.points[3]
        $first.Left = 72; $first.Top = 144; $first.Width = 96; $first.Height = 96
    }
    $expectedShapes.Add([pscustomobject]@{ name = 'FirstPageImage'; kind = 'Image'; left = 72; top = 144; width = 96; height = 96 })
    if ($SharedImages) {
        $shared = $last.Shapes.AddPicture($ImagePath, 0, -1, 72, 144, 96, 96)
        $shared.Name = 'SharedLastPageImage'
        $objects.Add($shared)
        $expectedShapes.Add([pscustomobject]@{ name = 'SharedLastPageImage'; kind = 'Image'; left = 72; top = 144; width = 96; height = 96 })
    }
    if ($ReplaceSharedFirst) {
        $format = $first.PictureFormat
        $objects.Add($format)
        $format.Replace($SecondImagePath, 1)
        $first.LockAspectRatio = 0
        $format.CropLeft = 0; $format.CropTop = 0; $format.CropRight = 0; $format.CropBottom = 0
        $first.Left = 72; $first.Top = 144; $first.Width = 96; $first.Height = 96
    }
    $bounds = if ($Edited) { @(252, 180, 144, 72) } else { @(216, 144, 192, 96) }
    $second = $last.Shapes.AddPicture($SecondImagePath, 0, -1, $bounds[0], $bounds[1], $bounds[2], $bounds[3])
    $second.Name = 'LastPageImage'
    $objects.Add($second)
    if ($ReplaceLast -or $ReplacementJpegPath) {
        $format = $second.PictureFormat
        $objects.Add($format)
        $format.Replace($ImagePath, 1) # pbPictureInsertAsEmbedded
        if ($ReplacementJpegPath) {
            New-NativeReplacementJpeg $ReplacementJpegPath
            $format.Replace($ReplacementJpegPath, 1)
        }
        $second.LockAspectRatio = 0
        $points = if ($CropLast) { (Get-NativePictureCropBasis $ReplacementJpegPath).points } else { @(0, 0, 0, 0) }
        $format.CropLeft = $points[0]
        $format.CropTop = $points[1]
        $format.CropRight = $points[2]
        $format.CropBottom = $points[3]
        $second.Left = $bounds[0]
        $second.Top = $bounds[1]
        $second.Width = $bounds[2]
        $second.Height = $bounds[3]
    }
    $expectedShapes.Add([pscustomobject]@{ name = 'LastPageImage'; kind = 'Image'; left = $bounds[0]; top = $bounds[1]; width = $bounds[2]; height = $bounds[3] })
    return 3
}
