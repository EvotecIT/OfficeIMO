namespace OfficeIMO.Pdf;

/// <summary>
/// Catalog page mode requested for the initial viewer chrome of a generated PDF.
/// </summary>
public enum PdfCatalogPageMode {
    /// <summary>
    /// Shows neither outlines nor thumbnails when the document is opened.
    /// </summary>
    UseNone,

    /// <summary>
    /// Shows the document outline panel when the document is opened.
    /// </summary>
    UseOutlines,

    /// <summary>
    /// Shows the page thumbnails panel when the document is opened.
    /// </summary>
    UseThumbs,

    /// <summary>
    /// Opens the document in full-screen mode when supported by the viewer.
    /// </summary>
    FullScreen,

    /// <summary>
    /// Shows the optional-content group panel when the document is opened.
    /// </summary>
    UseOC,

    /// <summary>
    /// Shows the attachments panel when the document is opened.
    /// </summary>
    UseAttachments
}
