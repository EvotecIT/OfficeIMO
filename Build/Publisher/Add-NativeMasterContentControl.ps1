function Add-NativeMasterContentControl {
    <#
    .SYNOPSIS
    Authors one populated master applied to three pages with overlapping foreground.
    .DESCRIPTION
    Uses the isolated worker's text/image helpers and object tracking. One master
    rectangle, image and text frame remain owned once. The last page reuses the
    image and adds foreground artwork. Wrapping is disabled explicitly.
    .PARAMETER Document
    Task-owned new publication.
    .PARAMETER FirstPage
    Its first physical page.
    .PARAMETER ImagePath
    Unused path for the synthetic shared PNG.
    .PARAMETER ReplacementPath
    Unused wide PNG path for the edited master-image replacement.
    .PARAMETER Edited
    Authors changed master text, rectangle fill/bounds and image content.
    .OUTPUTS
    The physical page count.
    #>
    param($Document, $FirstPage, [string]$ImagePath, [string]$ReplacementPath, [switch]$Edited)

    $second = $Document.Pages.Add(1, 1)
    $third = $Document.Pages.Add(1, 2)
    $objects.Add($second)
    $objects.Add($third)
    $master = $Document.MasterPages.Item(1)
    $objects.Add($master)
    if ($Document.MasterPages.Count -ne 1 -or $master.IsTwoPageMaster) {
        throw 'Master content requires one single-page master.'
    }
    foreach ($page in @($FirstPage, $second, $third)) { $page.Master = $master }
    New-ReferencePicture $ImagePath
    $backgroundTop = if ($Edited) { 60 } else { 72 }
    $background = $master.Shapes.AddShape(1, 36, $backgroundTop, 240, 96)
    $background.Name = 'MasterBackground'
    $background.Fill.ForeColor.RGB = if ($Edited) { 16711680 } else { 255 }
    $background.Line.Visible = 0
    $objects.Add($background)
    $expectedShapes.Add([pscustomobject]@{ name = 'MasterBackground'; kind = 'Rectangle'; left = 36; top = $backgroundTop; width = 240; height = 96 })
    $logo = $master.Shapes.AddPicture($ImagePath, 0, -1, 72, 72, 96, 96)
    $logo.Name = 'MasterImage'
    $objects.Add($logo)
    $expectedShapes.Add([pscustomobject]@{ name = 'MasterImage'; kind = 'Image'; left = 72; top = 72; width = 96; height = 96 })
    $masterTextValue = if ($Edited) { 'Updated master content' } else { 'Master: shared content' }
    $masterText = Add-TextFrame $master 'MasterText' $masterTextValue 36 72 288 108 18
    $firstBand = $FirstPage.Shapes.AddShape(1, 36, 84, 240, 12)
    $firstBand.Name = 'FirstForegroundBand'
    $firstBand.Fill.ForeColor.RGB = 16711680
    $firstBand.Line.Visible = 0
    $objects.Add($firstBand)
    $expectedShapes.Add([pscustomobject]@{ name = 'FirstForegroundBand'; kind = 'Rectangle'; left = 36; top = 84; width = 240; height = 12 })
    $pageImage = $third.Shapes.AddPicture($ImagePath, 0, -1, 180, 72, 96, 96)
    $pageImage.Name = 'LastForegroundImage'
    $objects.Add($pageImage)
    $expectedShapes.Add([pscustomobject]@{ name = 'LastForegroundImage'; kind = 'Image'; left = 180; top = 72; width = 96; height = 96 })
    $lastBand = $third.Shapes.AddShape(1, 36, 108, 240, 12)
    $lastBand.Name = 'LastForegroundBand'
    $lastBand.Fill.ForeColor.RGB = 32768
    $lastBand.Line.Visible = 0
    $objects.Add($lastBand)
    $expectedShapes.Add([pscustomobject]@{ name = 'LastForegroundBand'; kind = 'Rectangle'; left = 36; top = 108; width = 240; height = 12 })
    $lastText = Add-TextFrame $third 'LastForegroundText' 'Page foreground' 36 180 288 108 12
    if ($Edited) {
        New-NativeReplacementImage $ReplacementPath
        $format = $logo.PictureFormat
        $objects.Add($format)
        $format.Replace($ReplacementPath, 1)
        $logo.LockAspectRatio = 0
        $format.CropLeft = 0; $format.CropTop = 0; $format.CropRight = 0; $format.CropBottom = 0
        $logo.Left = 72; $logo.Top = 72; $logo.Width = 96; $logo.Height = 96
    }
    foreach ($shape in @($background, $logo, $masterText, $firstBand, $pageImage, $lastBand, $lastText)) {
        $wrap = $shape.TextWrap
        $objects.Add($wrap)
        $wrap.Type = 0 # pbWrapTypeNone
        if ($wrap.Type -ne 0) { throw 'Master content requires wrapping disabled.' }
    }
    return 3
}
