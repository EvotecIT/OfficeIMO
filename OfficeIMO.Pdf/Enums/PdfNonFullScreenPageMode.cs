namespace OfficeIMO.Pdf;

/// <summary>
/// Viewer page mode requested when leaving full-screen display.
/// </summary>
public enum PdfNonFullScreenPageMode {
    /// <summary>
    /// Shows neither outlines nor thumbnails when full-screen mode exits.
    /// </summary>
    UseNone,

    /// <summary>
    /// Shows the document outline panel when full-screen mode exits.
    /// </summary>
    UseOutlines,

    /// <summary>
    /// Shows the page thumbnails panel when full-screen mode exits.
    /// </summary>
    UseThumbs,

    /// <summary>
    /// Shows the optional-content group panel when full-screen mode exits.
    /// </summary>
    UseOC
}
