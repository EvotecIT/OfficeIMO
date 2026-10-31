function Add-NativeComparisonPictures {
    <#
    .SYNOPSIS
    Authors a picture-ownership comparison in a newly created publication.
    .DESCRIPTION
    Called by the isolated native worker after its Publisher process is identified.
    Applies one single-page master to two physical pages, adds one master picture,
    and duplicates an embedded picture on the first physical page. The duplicate
    uses the same input image; shared native image-store identity must be checked
    from the saved PUB rather than inferred from the Duplicate API.
    .PARAMETER Document
    The task-owned, newly created Microsoft Publisher publication.
    .PARAMETER FirstPage
    Its first physical page.
    .PARAMETER PicturePath
    The synthetic PNG produced by the worker's New-ReferencePicture helper.
    .PARAMETER ReplacementPath
    Optional synthetic PNG to embed in the original physical placement after
    duplication. The other physical placement and master retain the first image.
    .PARAMETER Cropped
    Applies explicit nonzero source-image crop to the original physical placement.
    .OUTPUTS
    Authored page/master assignments and observed crop/link values for each picture.
    #>
    param(
        $Document,
        $FirstPage,
        [string] $PicturePath,
        [string] $ReplacementPath,
        [switch] $Cropped
    )

    Set-NativePhase 'PictureOwnership-AddSecondPage'
    $second = $Document.Pages.Add(1, 1)
    $objects.Add($second)
    Set-NativePhase 'PictureOwnership-ObserveMaster'
    $master = $Document.MasterPages.Item(1)
    $objects.Add($master)
    if ($Document.MasterPages.Count -ne 1 -or $master.IsTwoPageMaster) {
        throw 'The picture control requires one single-page default master.'
    }

    $assignments = [Collections.Generic.List[object]]::new()
    $pages = @($FirstPage, $second)
    for ($index = 0; $index -lt $pages.Count; $index++) {
        Set-NativePhase ('PictureOwnership-AssignMaster-' + ($index + 1))
        $pages[$index].Master = $master
        $assignments.Add([pscustomobject]@{
            pageIndex = $index + 1
            masterPageId = [int]$master.PageID
        })
    }

    Set-NativePhase 'PictureOwnership-AddMasterPicture'
    $masterPicture = $master.Shapes.AddPicture($PicturePath, 0, -1, 24, 24, 48, 48)
    $objects.Add($masterPicture)
    Set-NativePhase 'PictureOwnership-AddPhysicalPicture'
    $original = $FirstPage.Shapes.AddPicture($PicturePath, 0, -1, 72, 144, 96, 96)
    $objects.Add($original)
    Set-NativePhase 'PictureOwnership-DuplicatePhysicalPicture'
    # Publisher documents that Duplicate returns the new Shape immediately after
    # the source in Shapes. It does not promise native BLIP deduplication.
    $duplicate = $original.Duplicate()
    $objects.Add($duplicate)

    $placements = @(
        [pscustomobject]@{
            shape = $masterPicture
            name = 'MasterPicture'
            pageRole = 'Master'
            pageId = [int]$master.PageID
            pageIndex = 1
            left = 24
            top = 24
            width = 48
            height = 48
        },
        [pscustomobject]@{
            shape = $original
            name = 'PhysicalPictureOriginal'
            pageRole = 'Physical'
            pageId = [int]$FirstPage.PageID
            pageIndex = 1
            left = 72
            top = 144
            width = 96
            height = 96
        },
        [pscustomobject]@{
            shape = $duplicate
            name = 'PhysicalPictureDuplicate'
            pageRole = 'Physical'
            pageId = [int]$FirstPage.PageID
            pageIndex = 1
            left = 216
            top = 144
            width = 96
            height = 96
        }
    )
    $pictureObservations = [Collections.Generic.List[object]]::new()
    foreach ($placement in $placements) {
        Set-NativePhase ('PictureOwnership-Place-' + $placement.name)
        $shape = $placement.shape
        $shape.Name = $placement.name
        $shape.Line.Visible = 0
        $format = $shape.PictureFormat
        $objects.Add($format)
        if ($ReplacementPath -and $placement.name -eq 'PhysicalPictureOriginal') {
            Set-NativePhase 'PictureOwnership-ReplacePhysicalPicture'
            $format.Replace($ReplacementPath, 1) # pbPictureInsertAsEmbedded
        }
        $croppedTarget = $Cropped -and $placement.name -eq 'PhysicalPictureOriginal'
        if (($ReplacementPath -or $Cropped) -and $placement.name -eq 'PhysicalPictureOriginal') {
            $shape.LockAspectRatio = 0
        }
        # Numeric crop values are points relative to the original image size.
        # Zero deliberately qualifies the uncropped rectangular edit profile.
        $wantedCrop = if ($croppedTarget) { @(6, 3, 0, 3) } else { @(0, 0, 0, 0) }
        $format.CropLeft = [double]$wantedCrop[0]
        $format.CropTop = [double]$wantedCrop[1]
        $format.CropRight = [double]$wantedCrop[2]
        $format.CropBottom = [double]$wantedCrop[3]
        $shape.Left = $placement.left
        $shape.Top = $placement.top
        $shape.Width = $placement.width
        $shape.Height = $placement.height
        foreach ($property in @('Left', 'Top', 'Width', 'Height')) {
            if ([Math]::Abs([double]$shape.$property - $placement.($property.ToLowerInvariant())) -gt 0.001) {
                throw ('Publisher did not retain the authored picture bounds: ' + $placement.name + '/' + $property)
            }
        }
        $expectedShapes.Add([pscustomobject]@{
            name = $placement.name
            kind = 'Image'
            pageRole = $placement.pageRole
            pageIndex = $placement.pageIndex
            left = $placement.left
            top = $placement.top
            width = $placement.width
            height = $placement.height
            cropPoints = [pscustomobject]@{
                left = $wantedCrop[0]
                top = $wantedCrop[1]
                right = $wantedCrop[2]
                bottom = $wantedCrop[3]
            }
        })
        Set-NativePhase ('PictureOwnership-ReadCropLeft-' + $placement.name)
        $cropLeft = [double]$format.CropLeft
        Set-NativePhase ('PictureOwnership-ReadCropTop-' + $placement.name)
        $cropTop = [double]$format.CropTop
        Set-NativePhase ('PictureOwnership-ReadCropRight-' + $placement.name)
        $cropRight = [double]$format.CropRight
        Set-NativePhase ('PictureOwnership-ReadCropBottom-' + $placement.name)
        $cropBottom = [double]$format.CropBottom
        Set-NativePhase ('PictureOwnership-ReadIsLinked-' + $placement.name)
        $isLinked = $format.IsLinked
        $crop = [pscustomobject]@{
            left = $cropLeft
            top = $cropTop
            right = $cropRight
            bottom = $cropBottom
        }
        $cropValues = @($crop.left, $crop.top, $crop.right, $crop.bottom)
        $wrongCrop = @(0..3 | Where-Object { [Math]::Abs($cropValues[$_] - $wantedCrop[$_]) -gt 0.001 })
        if ($isLinked -ne 0 -or $wrongCrop.Count -gt 0) {
            throw ('Publisher did not retain the embedded picture/crop contract: ' + $placement.name)
        }
        $pictureObservations.Add([pscustomobject]@{
            name = $placement.name
            pageRole = $placement.pageRole
            pageId = $placement.pageId
            pageIndex = $placement.pageIndex
            isLinked = [bool]$isLinked
            cropPoints = $crop
        })
    }
    Set-NativePhase 'PictureOwnership-PicturesObserved'
    return [pscustomobject]@{
        pages = 2
        panelWidthPoints = $null
        masterAssignments = $assignments.ToArray()
        pictureObservations = $pictureObservations.ToArray()
    }
}
