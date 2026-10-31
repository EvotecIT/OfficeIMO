function Get-NativePictureCropBasis {
    <#
    .SYNOPSIS
    Maps the four-edge crop control's fractions to native automation points.
    .DESCRIPTION
    Reads source dimensions and encoded DPI through the existing Windows image
    provider. The independent native control uses these points; saved OfficeArt
    fractions must still be checked rather than inferred from the point setters.
    .PARAMETER Path
    Synthetic source image created by the isolated corpus worker.
    .OUTPUTS
    Source metrics, desired fractions and point arguments for left/top/right/bottom.
    #>
    param([string]$Path)
    $image = [Drawing.Image]::FromFile($Path)
    try {
        $width = $image.Width * 72.0 / $image.HorizontalResolution
        $height = $image.Height * 72.0 / $image.VerticalResolution
        return [pscustomobject]@{
            widthPixels = $image.Width; heightPixels = $image.Height
            dpiX = $image.HorizontalResolution; dpiY = $image.VerticalResolution
            fractions = @(0.25, 0.125, 0.0625, 0.1875)
            points = @(($width * 0.25), ($height * 0.125), ($width * 0.0625), ($height * 0.1875))
        }
    } finally { $image.Dispose() }
}
