namespace OfficeIMO.Drawing;

/// <summary>
/// JPEG chroma subsampling selection.
/// </summary>
public enum OfficeJpegSubsampling {
    /// <summary>
    /// 4:4:4 (no subsampling).
    /// </summary>
    Y444,
    /// <summary>
    /// 4:2:2 (half horizontal chroma resolution).
    /// </summary>
    Y422,
    /// <summary>
    /// 4:2:0 (half horizontal + vertical chroma resolution).
    /// </summary>
    Y420
}
