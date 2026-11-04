function New-NativeReplacementImage {
    <#
    .SYNOPSIS
    Creates the synthetic wide PNG used by native picture-replacement controls.
    .DESCRIPTION
    Writes a new 64-by-32 image at 96 DPI with contrasting halves and a black
    marker. The different aspect ratio distinguishes image metadata from frame
    geometry. Called only by the isolated corpus worker, after System.Drawing
    has been loaded for its original synthetic image.
    .PARAMETER Path
    New task-owned PNG path. Existing content is never overwritten.
    .OUTPUTS
    None. The PNG is written to Path.
    #>
    param([string] $Path)
    $bitmap = [Drawing.Bitmap]::new(64, 32)
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    try {
        $bitmap.SetResolution(96, 96)
        $graphics.Clear([Drawing.Color]::Magenta)
        $graphics.FillRectangle([Drawing.Brushes]::Cyan, 32, 0, 32, 32)
        $graphics.FillRectangle([Drawing.Brushes]::Black, 24, 8, 16, 16)
        $stream = [IO.FileStream]::new($Path, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
        try {
            $bitmap.Save($stream, [Drawing.Imaging.ImageFormat]::Png)
        } finally {
            $stream.Dispose()
        }
    } finally {
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}
