function New-NativeReplacementJpeg {
    <#
    .SYNOPSIS
    Creates the synthetic JPEG used by repeated picture-replacement controls.
    .DESCRIPTION
    Writes a new 48-by-24 image at 96 DPI with contrasting halves and a black
    marker. Runs only in the native corpus worker after System.Drawing is loaded.
    .PARAMETER Path
    New task-owned JPEG path. Existing content is never overwritten.
    .OUTPUTS
    None. The JPEG is written to Path.
    #>
    param([string]$Path)
    $bitmap = [Drawing.Bitmap]::new(48, 24)
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    try {
        $bitmap.SetResolution(96, 96)
        $graphics.Clear([Drawing.Color]::Yellow)
        $graphics.FillRectangle([Drawing.Brushes]::Red, 24, 0, 24, 24)
        $graphics.FillRectangle([Drawing.Brushes]::Black, 18, 6, 12, 12)
        $stream = [IO.FileStream]::new($Path, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
        try { $bitmap.Save($stream, [Drawing.Imaging.ImageFormat]::Jpeg) }
        finally { $stream.Dispose() }
    } finally {
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}
