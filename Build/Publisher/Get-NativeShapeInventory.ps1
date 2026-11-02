function Get-NativeShapeInventory {
    <#
    .SYNOPSIS
    Reads native shape geometry, text, picture and table observations without editing them.
    .PARAMETER Page
    A physical or master page in the already opened native publication.
    .PARAMETER Objects
    Retains observed COM wrappers for the isolated worker's lifetime.
    .OUTPUTS
    One inventory object per shape, including table text at one-based grid coordinates.
    #>
    param($Page, [Collections.Generic.List[object]] $Objects)
    $nativeShapes = [Collections.Generic.List[object]]::new()
    for ($shapeIndex = 1; $shapeIndex -le $Page.Shapes.Count; $shapeIndex++) {
        $shape = $Page.Shapes.Item($shapeIndex)
        $objects.Add($shape)
        $text = $null
        $font = $null
        $languageId = $null
        $pairKerningThreshold = $null
        $manualKerning = $null
        $fontSizePoints = $null
        $nextLink = $false
        $previousLink = $false
        $frameLayout = $null
        $picture = $null
        $tableGeometry = $null
        $tableCells = $null
        if ($shape.HasTextFrame -eq -1) {
            $frame = $shape.TextFrame
            $range = $frame.TextRange
            $objects.Add($frame)
            $objects.Add($range)
            $text = [string]$range.Text
            $font = [string]$range.Font.Name
            $languageId = [int]$range.LanguageID
            $fontProperties = $range.Font
            $objects.Add($fontProperties)
            $pairKerningThreshold = [double]$fontProperties.AutomaticPairKerningThreshold
            $manualKerning = [double]$fontProperties.Kerning
            $fontSizePoints = [double]$fontProperties.Size
            $nextLink = $frame.HasNextLink -eq -1
            $previousLink = $frame.HasPreviousLink -eq -1
            $frameLayout = [pscustomobject]@{
                columns = [int]$frame.Columns; columnSpacing = [double]$frame.ColumnSpacing
                left = [double]$frame.MarginLeft; top = [double]$frame.MarginTop
                right = [double]$frame.MarginRight; bottom = [double]$frame.MarginBottom
                verticalAlignment = [int]$frame.VerticalTextAlignment
            }
        }
        if ($shape.Type -eq 13) {
            $format = $shape.PictureFormat
            $objects.Add($format)
            if ($format.IsEmpty -eq 0) {
                $picture = [pscustomobject]@{
                    fileName = [string]$format.FileName
                    isLinked = $format.IsLinked -ne 0
                    cropPoints = [pscustomobject]@{
                        left = [double]$format.CropLeft
                        top = [double]$format.CropTop
                        right = [double]$format.CropRight
                        bottom = [double]$format.CropBottom
                    }
                }
            }
        }
        if ($shape.Type -eq 18) {
            $tableGeometry = Get-NativeTableGeometry $shape
            $tableCells = @(Get-NativeTableCellInventory -Shape $shape -Objects $Objects)
        }
        $nativeShapes.Add([pscustomobject]@{
            name = [string]$shape.Name
            type = [int]$shape.Type
            left = [double]$shape.Left
            top = [double]$shape.Top
            width = [double]$shape.Width
            height = [double]$shape.Height
            text = $text
            font = $font
            languageId = $languageId
            automaticPairKerningThresholdPoints = $pairKerningThreshold
            manualKerningPoints = $manualKerning
            fontSizePoints = $fontSizePoints
            nextLink = $nextLink
            previousLink = $previousLink
            frameLayout = $frameLayout
            picture = $picture
            tableGeometry = $tableGeometry
            tableCells = $tableCells
        })
    }
    return $nativeShapes.ToArray()
}
